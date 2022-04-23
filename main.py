import json
from datetime import datetime, timedelta
from string import ascii_uppercase

import openpyxl
from openpyxl.styles import Font
from openpyxl.utils.dataframe import dataframe_to_rows

import pandas as pd

import requests

SITE = 'https://bank.gov.ua/'
CMD = 'NBUStatService/v1/statdirectory/exchange?valcode={0}&date={1}&json'
API_CALL = f'{SITE}{CMD}'


class RateForPeriod:
    def __init__(self, currencies, start_date: str, end_date: str):
        self.df = None
        if isinstance(currencies, str):
            currencies = currencies.replace(' ', '')
            if ',' in currencies:
                currencies = currencies.split(',')
            else:
                currencies = [currencies]
        elif not isinstance(currencies, (list, tuple,)):
            print('[currencies parameter]: allowed only str or list')
            raise
        if isinstance(currencies, tuple):
            currencies = list(currencies)

        self.currencies = currencies
        self.start_date = datetime.strptime(start_date, '%Y-%m-%d')
        self.end_date = datetime.strptime(end_date, '%Y-%m-%d')
        if self.start_date > self.end_date:
            self.start_date, self.end_date = self.end_date, self.start_date
            print('[Dates swapped] start_date > end_date')

    def get_rates(self):
        the_date = self.start_date
        data = []
        while the_date <= self.end_date:
            fmt_date = the_date.strftime("%Y%m%d")
            cur = [the_date.strftime('%Y-%m-%d')]
            for currency in self.currencies:
                one = self._get_rate_per_date(currency.lower(), fmt_date)
                cur.append(one)
            the_date = the_date + timedelta(days=1)
            data.append(cur)

        df = pd.DataFrame(data, columns=['Date']+self.currencies)
        self.df = df
        return self.df

    def save_xlsx(self, filename: str):
        headers = list(self.df)  # get the headers
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'Bank rates'  # set the name of sheet

        # pandas dataFrame to openpyxl worksheet:
        for r in dataframe_to_rows(self.df, index=False, header=True):
            ws.append(r)

        # set the bold font for header:
        for index, column_name in enumerate(headers):
            cell = ws.cell(row=1, column=index+1)
            cell.font = Font(bold=True)

        # set the width of 1st column:
        ws.column_dimensions[ascii_uppercase[0]].width = 11
        try:
            wb.save(filename)
            print(f'File {filename} saved OK...')
        except Exception as err:
            print(f'Error during saving file {filename}: {err}')

    def _get_rate_per_date(self, currency: str, yyyymmdd: str):
        api = API_CALL.format(currency, yyyymmdd)
        # print(api)
        result = ''
        response = requests.get(api, headers=self._headers())
        if response.status_code == 200:
            response = json.loads(response.content.decode())
            if response:
                response = response[0]
                rate = response.get('rate')
                if rate:
                    result = rate
                else:
                    print(response.get('message'))
        else:
            print(response.status_code, response.content.decode())
        return result

    def _headers(self, user_agent=None):
        def_user_agent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'
        user_agent = user_agent if not None else def_user_agent
        headers = {
            'User-Agent': user_agent
        }
        return headers


if __name__ == '__main__':
    # TODO: Need to add support of command line
    cur = ('EUR', 'USD')
    start_date = '2022-02-01'
    end_date = '2022-02-10'
    c = RateForPeriod(cur, start_date, end_date)
    rates = c.get_rates()
    c.save_xlsx('rate.xlsx')
