"""
Robinhood Option Profit/Loss Analysis
Generate Profit/Loss Statement in XLS
Author: Rahul Manocha
Date : 02/26/2021
"""
import robin_stocks.robinhood as rh
import pyotp
import os
from datetime import datetime
from pytz import timezone
import pprint
import xlsxwriter
import xlsxwriter.utility as xlsutils
import argparse

class AccessRH:
  def __init__(self,rh_filepath = './login.txt', mfa_method = 'sms', expiresIn = '2000', store_session = False):
    self.rh_filepath = rh_filepath
    self.mfa_method = mfa_method
    loginfile = open(self.rh_filepath,'r')
    self.rh_user = loginfile.readline().rstrip()
    self.rh_pass = loginfile.readline().rstrip()
    self.rh_mfa = loginfile.readline().rstrip()
    self.expiresIn = expiresIn
    self.store_session = store_session
    loginfile.close()
    if mfa_method != 'sms':
      print('Enter this 2FA in Robinhood App:')
      self.totp = pyotp.TOTP(self.rh_mfa).now()
      print(self.totp)

  def attempt_login(self):
    if self.mfa_method == 'sms':
      try:
        self.login = rh.login(self.rh_user,self.rh_pass,expiresIn=self.expiresIn,store_session=self.store_session,by_sms=True)
      except:
        print("ERROR : Login by SMS Failed")
      else:
        print("INFO: MFA Login with SMS Successful")
    else:
      try:
        self.login = rh.login(self.rh_user,self.rh_pass,expiresIn=self.expiresIn,store_session=self.store_session,mfa_code=self.totp)
      except:
        print("ERROR : Login by MFA OTP Failed")
      else:
        print("INFO: MFA Login with OTP Successful")
  def attempt_logout(self):
    try:
      rh.logout()
    except:
      print("ERROR: Logout Not Successul")

  def test_login(self):
    try:
     my_stocks = rh.build_holdings()
     tickers = my_stocks.keys()
     print("Holdings:")
     tick = ""
     for key in tickers:
        tick += " " + key
     print(tick)
    except:
      print("ERROR: Cannot Access Rohinhood Account")


class ParseRHOptions:
  def __init__(self,current_date=datetime.now(timezone('America/Los_Angeles')).strftime("%Y%m%d")):
    self.now =  current_date
    self.filled_options = {}
    self.options_profit = {}
    self.open_contracts = {}
    self.sellyear = set()

  def date_format(self,date):
    # convert from yyyy-mm-dd to mmddyyyy
    datestr = datetime.strptime(date[0:10],'%Y-%m-%d').strftime("%Y%m%d")
    return datestr

  def date_delta(self,date1,date2):
    #days between 2 dates
    delta = datetime.strptime(date1,'%Y%m%d') - datetime.strptime(date2,'%Y%m%d')
    return delta.days

  def get_year(self,date):
    return datetime.strptime(date,'%Y%m%d').strftime("%Y")

  def gen_key(self,option):
    strike = option['strike_price']
    expiry = option['expiration_date']
    open_strategy = option['opening_strategy']
    close_strategy = option['closing_strategy']
    if open_strategy is None:
       strategy = close_strategy
    else:
       strategy = open_strategy
    key = strike + expiry + strategy
    return key

  # Read Filled Option Contracts from RH
  # Parse options into a dictionary of lists
  # Only Single Legged Calls/Puts are supported currently
  def parse_option_orders(self):
    all_orders = rh.orders.get_all_option_orders()
    for order in all_orders:
      if order['state'] == 'filled':
         for leg in order['legs']:
           opening_strategy = order['opening_strategy']
           closing_strategy = order['closing_strategy']
           # Ignore strategies other than single legged calls and puts
           if opening_strategy is not None:
             if 'spread' in opening_strategy or 'iron_condor' in opening_strategy:
                  continue
           if closing_strategy is not None:
             if 'spread' in closing_strategy or 'iron_condor' in closing_strategy:
                  continue
           instrument_data = rh.helper.request_get(leg['option'])
           temp = {}
           symbol = order['chain_symbol']
           temp['expiration_date']  =  self.date_format(instrument_data['expiration_date'])
           temp['strike_price'] = instrument_data['strike_price']
           temp['type'] = instrument_data['type']
           temp['side'] = leg['side']
           temp['created_at'] = self.date_format(order['created_at'])
           temp['direction']  = order['direction']
           temp['quantity']   = order['quantity']
           temp['type']       = order['type']
           temp['opening_strategy'] = opening_strategy
           temp['closing_strategy'] = closing_strategy
           temp['price'] = leg['executions'][0]['price']
           temp['processed_quantity'] = order['processed_quantity']

           if symbol not in self.filled_options.keys():
             self.filled_options[symbol] = {}
             self.filled_options[symbol]['open'] = []
             self.filled_options[symbol]['close']  = []
           if opening_strategy is not None :
              self.filled_options[symbol]['open'].append(temp)
           if closing_strategy is not None:
              self.filled_options[symbol]['close'].append(temp)

  # Step through and Open and Closed contracts to find profit/loss
  def find_profit_loss(self):
    for symbol in self.filled_options.keys(): # for each symbol match open and close contracts
      self.options_profit[symbol] = {}
      self.options_profit[symbol]['profit'] = []
      self.options_profit[symbol]['cost']   = []
      self.options_profit[symbol]['duration'] = []
      self.options_profit[symbol]['strategy'] = []
      self.options_profit[symbol]['year'] = []
      self.open_contracts[symbol] = []
      chain = self.filled_options[symbol]['open']
      open_list = sorted(chain, key = lambda i: i['created_at']) # sort in order of creation
      chain = self.filled_options[symbol]['close']
      close_list = sorted(chain,key = lambda i: i['created_at'])
      for contract in open_list:
         self.find_close(contract, close_list,self.options_profit[symbol],self.open_contracts[symbol])
      if not self.options_profit[symbol]['profit']:
        del self.options_profit[symbol]

  # match open and closed contracts for a given stock to find profit/loss
  def find_close(self,contract,close_list, balancesheet, open_contracts):
      now = self.now
      # Step 1: Find contract in close_list
      openkey = self.gen_key(contract)
      for closed in close_list:
        if openkey ==  self.gen_key(closed): # match found
            #Step 2: Check strategy, and find profit loss
            openquant = float(contract['processed_quantity'])
            closequant = float(closed['processed_quantity'])
            openprice = float(contract['price'])
            closeprice = float(closed['price'])
            strike = float(contract['strike_price'])
            duration = self.date_delta(closed['created_at'] , contract['created_at'])
            sellyear = self.get_year(closed['created_at'])
            opendirection = contract['direction']
            closedirection = closed['direction']
            strategy = contract['opening_strategy']
            quant = min(openquant,closequant)
            if 'long' in strategy: # longcall and long put
              profit = (closeprice - openprice) * quant
              cost = openprice * quant
            elif 'short' in strategy: #short call and short put
              profit = (openprice - closeprice) * quant
              cost = strike * quant # cost is the collateral held
            contract['processed_quantity'] = str(abs(openquant - quant))
            closed['processed_quantity']   = str(abs(closequant - quant))
            balancesheet['profit'].append(profit)
            balancesheet['cost'].append(cost)
            balancesheet['duration'].append(duration)
            balancesheet['strategy'].append(strategy)
            balancesheet['year'].append(sellyear)
            self.sellyear.add(sellyear)

      if float(contract['processed_quantity']) > 0 :
           # Step 3 : Either contract is still open or expired
           if contract['expiration_date'] < now: # contract has expired
           #find profit/loss and update
             profit = float(contract['price'])  * float(contract['processed_quantity'])
             duration = self.date_delta(contract['expiration_date'] , contract['created_at'])
             sellyear = self.get_year(contract['expiration_date'])
             collateral = float(contract['strike_price']) * float(contract['processed_quantity'])
             balancesheet['duration'].append(duration)
             balancesheet['year'].append(sellyear)
             self.sellyear.add(sellyear)
             balancesheet['strategy'].append(contract['opening_strategy'])
             if contract['direction'] == 'credit' : #profit
                balancesheet['profit'].append(profit)
                balancesheet['cost'].append(collateral) #cost is collateral held for sell call or buy call
             else: #loss if debit
                balancesheet['profit'].append(-1 * profit)
                balancesheet['cost'].append(profit) # for buy call or buy put expiry, cost is same as loss
           else: # contract is still open
             open_contracts.append(contract)

class GenXlsx:
  def __init__(self,options_profit,open_contracts,sellyear,workbookname='./OptionsProfit',current_date=datetime.now(timezone('America/Los_Angeles')).strftime("%Y%m%d")):
    self.now = current_date
    self.options_profit = options_profit
    self.open_contracts = open_contracts
    self.sellyear = sorted(sellyear)
    self.workbookname = workbookname + '_' + self.now + '.xlsx'
    print("Options Profit/Loss Statement Generated in File: %s" % self.workbookname)
    self.workbook = xlsxwriter.Workbook(self.workbookname)
    # Add a bold format to use to highlight cells.
    self.bold = self.workbook.add_format({'bold': True})
    # Add a number format for cells with money.
    self.money = self.workbook.add_format({'num_format': '$#,####.00'})
    self.profit = self.workbook.add_format({'num_format' : '$#,####.00'})
    self.profit.set_bg_color('#00FF00') # lime
    self.loss = self.workbook.add_format({'num_format' : '$#,####.00'})
    self.loss.set_bg_color('#FF00FF') # pink


  def accumulate_sum_by_year(self,profit,year):
        cumprofit = {}
        for key,val in zip(year,profit):
           cumprofit[key] = cumprofit.get(key,0) + val
        return cumprofit

  def accumulated_profit_worksheet(self):
    # Create Accumulated Summary Worksheet
    # Formatting :
    # Stock Year1 Year2 Year3 Total
    # AAPL  123    456    789  1368
    # Total 123    456    789  1368
    self.accum_worksheet = self.workbook.add_worksheet('Summary')
    self.accum_worksheet.write('A1', 'Stock Ticker',self.bold)
    row = 0
    col = 1
    for year in self.sellyear:
        self.accum_worksheet.write(row,col,year,self.bold)
        col+=1
    self.accum_worksheet.write(row,col,'Stock Total',self.bold)
    total_col = col
    row = 1
    col = 0
    for ticker in self.options_profit.keys():
        self.accum_worksheet.write(row,col,ticker,self.bold)
        col+=1
        profit_list = self.options_profit[ticker]['profit']
        year = self.options_profit[ticker]['year']
        cumprofit = self.accumulate_sum_by_year(profit_list,year)
        for year in self.sellyear:
            if year in cumprofit.keys():
                self.accum_worksheet.write(row,col,cumprofit.get(year)*100,self.money)
            else:
                self.accum_worksheet.write(row,col,0,self.money)
            col+=1
        row_total = sum(cumprofit.values())*100
        if row_total > 0: # profit
            self.accum_worksheet.write(row,col,row_total,self.profit)
        else:
            self.accum_worksheet.write(row,col,row_total,self.loss)
        row+=1
        col=0
    self.accum_worksheet.write(row,col,'Year Total',self.bold)
    total_row = row
    total_format = self.workbook.add_format({'num_format': '$#,####.00'})
    total_format.set_bg_color('#FFFF00')
    for col in range(1,total_col+1):
      cell_range = xlsutils.xl_range(1,col,total_row-1,col)
      sum_str = '=SUM(' + cell_range + ')'
      self.accum_worksheet.write(total_row,col,sum_str,total_format)
    #self.workbook.close()

  def  itemized_profit_worksheet(self):
    # Create Itemized Worksheet of Closed contracts
    # Formatting:
    # Ticker Cost Duration(days) Strategy    Year    Profit
    # AAPL   123   30            long_call    2019    456
    self.item_worksheet = self.workbook.add_worksheet('Itemized')
    self.item_worksheet.write('A1', 'Stock Ticker',self.bold)
    self.item_worksheet.write('B1', 'Cost', self.bold)
    self.item_worksheet.write('C1', 'Duration(days)', self.bold)
    self.item_worksheet.write('D1', 'Strategy', self.bold)
    self.item_worksheet.write('E1', 'Year', self.bold)
    self.item_worksheet.write('F1', 'Profit', self.bold)
    row = 1
    col = 0
    for ticker in self.options_profit.keys():
      cost = [c * 100 for c in self.options_profit[ticker]['cost']]
      duration = self.options_profit[ticker]['duration']
      profit = [p * 100 for p in self.options_profit[ticker]['profit']]
      strategy = self.options_profit[ticker]['strategy']
      year = self.options_profit[ticker]['year']
      for iter in range(0, len(cost)):
          self.item_worksheet.write(row,col,ticker,self.bold)
          self.item_worksheet.write(row,col+1,cost[iter],self.money)
          self.item_worksheet.write(row,col+2,duration[iter])
          self.item_worksheet.write(row,col+3,strategy[iter])
          self.item_worksheet.write(row,col+4,year[iter])
          self.item_worksheet.write(row,col+5,profit[iter])
          row += 1
          col = 0
  def close(self):
    self.workbook.close()


if __name__ == '__main__':

    print("Robinhood Options Profit/Loss Analysis v0.1")

    # Parse Input Arguments
    text = 'Robinhood Options Profit/Loss Analysis v0.1'
    parser = argparse.ArgumentParser(description=text)
    parser.add_argument("-a", "--auth", help="Input Authentication method: mfa/sms",default='sms')
    # Read arguments from the command line
    args = parser.parse_args()
    if args.auth:
        print('Authentication Method : %s' % args.auth)
        mfa_method = args.auth

    # Login into RH account using mfa_method
    rh_obj = AccessRH(mfa_method=mfa_method,expiresIn='5000' , store_session=True)
    rh_obj.attempt_login()
    rh_obj.test_login()

    tz = timezone('America/Los_Angeles')
    print("TimeZone: %s" % tz)
    current_date = datetime.now(tz)
    print("Date: %s" % current_date)

    # Parse Filled Options to Find Profit/Loss
    option_obj = ParseRHOptions(current_date=current_date.strftime("%Y%m%d"))
    option_obj.parse_option_orders()
    option_obj.find_profit_loss()

    # Generate XLS File with Summary of Profit/Loss and Itemized Contracts
    # Summary Worksheet Formatting:
        # Stock Year1 Year2 Year3 Total
        # AAPL  123    456    789  1368
        # Total 123    456    789  1368
    # Itemized Worksheet Formatting:
        # Ticker Cost Duration(days) Strategy    Year    Profit
        # AAPL   123   30            long_call    2019    456
    xlsobj = GenXlsx(option_obj.options_profit,option_obj.open_contracts,option_obj.sellyear,current_date=current_date.strftime("%Y%m%d"))
    xlsobj.accumulated_profit_worksheet()
    xlsobj.itemized_profit_worksheet()
    xlsobj.close()


