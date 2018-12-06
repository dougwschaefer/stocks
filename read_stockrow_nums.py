#!/usr/bin/env python3
# lucky.py - Opens several Google search results.

"""
Created on August 11, 2018

@author: Doug Schaefer
"""

import pandas as pd
import openpyxl
import datetime
import requests
import io
#import pandas_datareader as web
from yahoofinancials import YahooFinancials
import date_converter

start = datetime.date(2008, 1, 1)
start_str = str(start.year)+"-"+str(start.month)+"-"+str(start.day)
now = datetime.datetime.now()
now_str = str(now.year)+"-"+str(now.month)+"-"+str(now.day)
years = [2008, 2009]

i = 2010
while i <= now.year:
    years.append(i)
    i = i+1

rev = pd.DataFrame(columns = years)
shares = pd.DataFrame(columns = years)
divs = pd.DataFrame(columns = years)
goodwill = pd.DataFrame(columns = years)
LTdebt = pd.DataFrame(columns = years)
TotDebt = pd.DataFrame(columns = years)
liabilities = pd.DataFrame(columns = years)
equity = pd.DataFrame(columns = years)
price = pd.DataFrame(columns = years)
adjPrice = pd.DataFrame(columns = years)

tickers = pd.read_csv("/Users/Doug/Stocks/Data/zacks/zacks_custom_screen.csv")

for company in tickers["Ticker"][:]:
#for company in ['AAPL', 'ABMD', 'GSK', 'IBM']:

    try:
        income_statement = pd.read_excel("https://stockrow.com/api/companies/" + company + "/financials.xlsx?dimension=MRY&section=Income%20Statement&sort=desc")  

        rev_row = []
        shares_row = []
        divs_row = []

        len_income = len(income_statement.columns)
        end_inc_stat = income_statement.columns[0].year
        start_inc_stat = income_statement.columns[len_income-1].year
    
        for data_year in range (2008, now.year+1) :
            if (data_year >= start_inc_stat) and (data_year <= end_inc_stat):
                if "Revenue" in income_statement.index:
                    rev_row = rev_row + [income_statement.loc["Revenue"].iloc[end_inc_stat-data_year]]
                else:
                    rev_row = rev_row + [""]
                if "Weighted Average Shs Out" in income_statement.index:
                    shares_row = shares_row + [income_statement.loc["Weighted Average Shs Out"].iloc[end_inc_stat-data_year]]
                else:
                    shares_row = shares_row + [""]
                if "Dividend per Share" in income_statement.index:
                    divs_row = divs_row + [income_statement.loc["Dividend per Share"].iloc[end_inc_stat-data_year]]
                else:
                    divs_row = divs_row + [""]
            else:
                rev_row = rev_row + [""]
                shares_row = shares_row + [""]
                divs_row = divs_row + [""]
            
        rev.loc[company] = rev_row
        shares.loc[company] = shares_row
        divs.loc[company] = divs_row
        
    except:
        print("unable to read income statement from " + company)
    
    try:
        balance_sheet = pd.read_excel("https://stockrow.com/api/companies/" + company + "/financials.xlsx?dimension=MRY&section=Balance%20Sheet&sort=desc")

        goodwill_row = []
        LTdebt_row = []
        TotDebt_row = []
        equity_row = []
        liabilities_row = []
    
        len_balance = len(balance_sheet.columns)
        end_bal_sht = balance_sheet.columns[0].year
        start_bal_sht = balance_sheet.columns[len_balance-1].year
    
        for data_year in range (2008, now.year+1) :
            if (data_year >= start_bal_sht) and (data_year <= end_bal_sht):
                if "Goodwill and Intangible Assets" in balance_sheet.index:
                    goodwill_row = goodwill_row + [balance_sheet.loc["Goodwill and Intangible Assets"].iloc[end_bal_sht-data_year]]
                else:
                    goodwill_row = goodwill_row + [""]
                if "Long-term debt" in balance_sheet.index:
                    LTdebt_row = LTdebt_row + [balance_sheet.loc["Long-term debt"].iloc[end_bal_sht-data_year]]
                else:
                    LTdebt_row = LTdebt_row + [""]
                if "Total debt" in balance_sheet.index:
                    TotDebt_row = TotDebt_row + [balance_sheet.loc["Total debt"].iloc[end_bal_sht-data_year]]
                else:
                    TotDebt_row = TotDebt_row + [""]
                if "Total liabilities" in balance_sheet.index:
                    liabilities_row = liabilities_row + [balance_sheet.loc["Total liabilities"].iloc[end_bal_sht-data_year]]
                else:
                    liabilities_row = liabilities_row + [""]
                if "Shareholders Equity" in balance_sheet.index:    
                    equity_row = equity_row + [balance_sheet.loc["Shareholders Equity"].iloc[end_bal_sht-data_year]]
                else:
                    equity_row = equity_row + [""]
            else:
                goodwill_row = goodwill_row + [""]
                LTdebt_row = LTdebt_row + [""]
                TotDebt_row = TotDebt_row + [""]
                liabilities_row = liabilities_row + [""]
                equity_row = equity_row + [""]
            
        goodwill.loc[company] = goodwill_row
        LTdebt.loc[company] = LTdebt_row
        TotDebt.loc[company] = TotDebt_row
        liabilities.loc[company] = liabilities_row
        equity.loc[company] = equity_row
        
    except:
        print("unable to read balance sheet from " + company)
        
    try:
        yahoo_financials = YahooFinancials(company)
        price_history = yahoo_financials.get_historical_price_data(start_str, now_str, 'monthly')
        price_history_list = price_history[company]["prices"]
        price_len = len(price_history_list)
        
        tickers_index = tickers.index[tickers["Ticker"] == company]
#        fiscal_month = tickers.iloc[tickers_index].loc["Month of Foscal Yr End"]
        fiscal_month = tickers.iloc[tickers_index]["Month of Fiscal Yr End"].iloc[0]
        
        adj_price_row = []
        price_row = []
        
        for data_year in range (2008, now.year+1):
            for i in range (0, price_len):
                price_date = date_converter.string_to_date(price_history_list[i]['formatted_date'], '%Y-%m-%d')
                price_year = price_date.year
                if price_year > data_year:
                    adj_price_row = adj_price_row + [""]
                    price_row = price_row + [""]
                    break
                elif price_year == data_year:
                    price_month = price_date.month
                    if price_month > fiscal_month:
                        adj_price_row = adj_price_row + [""]
                        price_row = price_row + [""]
                        break
                    elif price_month == fiscal_month:
                        adj_price_row = adj_price_row + [price_history_list[i]['adjclose']]
                        price_row = price_row + [price_history_list[i]['close']]
                        break
                    elif i == (price_len-1):
                        adj_price_row = adj_price_row + [""]
                        price_row = price_row + [""]
                        break
                    
        adjPrice.loc[company] = adj_price_row
        price.loc[company] = price_row

    except:
        print("unable to read price history from " + company)

writer = pd.ExcelWriter('/Users/Doug/Stocks/Data/stockrow/stockrows.xlsx')
rev.to_excel(writer,'rev')
shares.to_excel(writer,'shares')
divs.to_excel(writer,'divs')
goodwill.to_excel(writer,'goodwill')
LTdebt.to_excel(writer,'LTdebt')
TotDebt.to_excel(writer,'TotDebt')
liabilities.to_excel(writer,'liabilities')
equity.to_excel(writer,'equity')
adjPrice.to_excel(writer,'adjPrice')
price.to_excel(writer,'price')

writer.save()




