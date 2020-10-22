import pandas as pd
import finnhub
import json
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import time
import datetime



class FinnhubConnector:
    finnhub_client = finnhub.Client(api_key='btakssf48v6vivh8r4ag')

    def __init__(self, ticker):
        self.ticker = ticker

    def get_company_financials(self, statement_type, frequency):
        company_financial = self.finnhub_client.financials(symbol=self.ticker, statement=statement_type, freq=frequency)

        company_financial = pd.DataFrame(company_financial['financials'])
        company_financial = company_financial[:5]
        company_financial = company_financial.dropna(axis=1, how='all')
        company_financial = company_financial.set_index('period')
        company_financial = company_financial.sort_index()
        return company_financial


    def get_economic_data(self, codename):
        codedata = self.finnhub_client.economic_data(codename)
        df = pd.DataFrame(codedata['data'])
        return df

    # ---------------------------------------------
#   COME BACK TO THIS
    def get_stock_candles(self):
        d = datetime.date(2000, 1, 1)
        d = time.mktime(d.timetuple())
        stock = self.finnhub_client.stock_candles(symbol=self.ticker, resolution='D', _from=d,
                                             to=time.mktime(datetime.date.today().timetuple()))
        stock = pd.DataFrame(stock)
        return stock

    def get_stock_quote(self):
        codedata = self.finnhub_client.quote(self.ticker)
        return codedata
    #   COME BACK TO THIS
#---------------------------------------------
    def metrics(self):
        metrics = self.finnhub_client.company_basic_financials(symbol=self.ticker, metric="all")
        metrics = metrics['series']
        stats = pd.read_csv("financialstatementaccounts.csv")
        stats = stats[stats['statement'] == 'bf']
        for name in stats["accountName"]:
            try:
                y = pd.DataFrame(metrics['annual'][name])
            except:
                pass
        return y

