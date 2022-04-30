import argparse
import json
import logging
from datetime import datetime, timedelta
from string import ascii_uppercase

import openpyxl
from openpyxl.styles import Font
from openpyxl.utils.dataframe import dataframe_to_rows

import pandas as pd

import requests


__version__ = '1.0.1'
logging.basicConfig(filename='bankurs.log', level=logging.INFO,
                    format='%(asctime)s %(filename)s %(funcName)s %(message)s')
logger = logging.getLogger(__name__)

SITE = 'https://bank.gov.ua/'
CMD = 'NBUStatService/v1/statdirectory/exchange?valcode={0}&date={1}&json'
API_CALL = f'{SITE}{CMD}'
# based on API manual https://bank.gov.ua/ua/open-data/api-dev

class RateForPeriod:
    def __init__(self, currencies, start_date: str, end_date: str):
        if not currencies:
            logger.warning(
                f'[currencies parameter] "{currencies}" value is not allowed')
            raise
        if isinstance(currencies, str):
            currencies = currencies.replace(' ', '')
            if ',' in currencies:
                currencies = currencies.split(',')
            else:
                currencies = [currencies]
        elif not isinstance(currencies, (list, tuple,)):
            logger.warning('[currencies parameter]: allowed only str or list')
            raise
        if isinstance(currencies, tuple):
            currencies = list(currencies)

        self.currencies = currencies
        d_fmt = '%Y-%m-%d'  # agreed date format for dates passed
        self.start_date = datetime.strptime(start_date, d_fmt)
        self.end_date = datetime.strptime(end_date, d_fmt)
        if self.start_date > self.end_date:
            self.start_date, self.end_date = self.end_date, self.start_date
            logger.info('[Dates swapped] start_date > end_date')
        curs = '_'.join(currencies)
        self.filename = f'rates_{curs}_{self.start_date.strftime(d_fmt)}_' \
                        f'{self.end_date.strftime(d_fmt)}'
        self.df = None  # will be dataFrame with rates

    def get_rates(self):
        logger.info(f'Getting {self.currencies} rates for '
                    f'{self.start_date}-{self.end_date}...')
        the_date = self.start_date
        data = []
        while the_date <= self.end_date:
            fmt_date = the_date.strftime("%Y%m%d")
            row = [the_date.strftime('%Y-%m-%d')]
            for currency in self.currencies:
                one = self._get_rate_per_date(currency.lower(), fmt_date)
                row.append(one)
            the_date = the_date + timedelta(days=1)
            data.append(row)

        df = pd.DataFrame(data, columns=['Date']+self.currencies)
        self.df = df
        return self.df

    def save_xlsx(self, filename: str):
        if self.df.empty:
            logger.warning('[Save error] DataFrame is empty')
            raise
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
            logger.info(f'File {filename} saved OK...')
            return filename
        except Exception as err:
            logger.warning(f'Error during saving file {filename}: {err}')
            return

    def _get_rate_per_date(self, currency: str, yyyymmdd: str):
        api = API_CALL.format(currency, yyyymmdd)
        # logger.debug(api)
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
                    logger.debug(response.get('message'))
        else:
            logger.warning(
                f'{response.status_code}: {response.content.decode()}')
        return result

    def _headers(self, user_agent=None):
        def_user_agent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'
        if user_agent is None:
            user_agent = def_user_agent
        headers = {
            'User-Agent': user_agent
        }
        return headers


if __name__ == '__main__':
    txt = f'Extraction of official FX rates of National Bank of Ukraine'\
          f' (c) 2022 Andy Reo, version {__version__}'
    parser = argparse.ArgumentParser(description=txt)
    parser.add_argument("currencies", type=str,
                        help="code(s) of currency - USD, EUR, ...")
    parser.add_argument("start_date", type=str,
                        help="start date in format yyyy-mm-dd")
    parser.add_argument("end_date", type=str,
                        help="end date in format yyyy-mm-dd")
    args = parser.parse_args()

    c = RateForPeriod(args.currencies,
                      start_date=args.start_date,
                      end_date=args.end_date)
    rates = c.get_rates()
    c.save_xlsx(f'{c.filename}.xlsx')
