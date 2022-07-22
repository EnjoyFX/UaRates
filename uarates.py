import argparse
import json
import logging
from datetime import datetime, timedelta
from string import ascii_uppercase
from os.path import basename

import openpyxl
from openpyxl.styles import Font
from openpyxl.utils.dataframe import dataframe_to_rows

import pandas as pd

import requests


__version__ = '1.0.2'
log_name = basename(__file__).split('.')[0]
logging.basicConfig(filename=f'{log_name}.log', level=logging.INFO,
                    format='%(asctime)s %(filename)s %(funcName)s %(message)s')
logger = logging.getLogger(log_name)

SITE = 'https://bank.gov.ua/'
CMD = 'NBU_Exchange/exchange_site?valcode={0}&start={1}&end={2}&sort=exchangedate&order=desc&json'
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
        self.dates = None

    def get_rates(self):
        logger.info(f'Getting {self.currencies} rates for '
                    f'{self.start_date}-{self.end_date}...')
        data = {}
        fmt_date_start = self.start_date.strftime("%Y%m%d")
        fmt_date_end = self.end_date.strftime("%Y%m%d")

        for currency in self.currencies:
            one = self._get_rates_for_daterange(currency.lower(), fmt_date_start, fmt_date_end)
            if not data:
                data = {'Date': self.dates}
            data[currency] = one

        self.df = pd.DataFrame(data)
        return self  # for chain of methods ability

    def save_xlsx(self, filename: str):
        if self.df.empty:
            e = '[Save error] DataFrame is empty'
            logger.warning(e)
            raise Exception(e)
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

    def _get_rates_for_daterange(self, currency: str, yyyymmdd1: str, yyyymmdd2: str):
        api = API_CALL.format(currency, yyyymmdd1, yyyymmdd2)
        # logger.debug(api)
        result = ''
        response = requests.get(api, headers=self._headers())
        if response.status_code == 200:
            response = json.loads(response.content.decode())
            if response:
                rates = [r.get('rate') for r in response]
                dates = [r.get('exchangedate') for r in response]
                if self.dates is None:
                    self.dates = dates
                else:
                    if self.dates!=dates:
                        logger.critical('Different date ranges for currencies!')
        else:
            logger.warning(
                f'{response.status_code}: {response.content.decode()}')
        return rates

    def _headers(self, user_agent=None):
        def_user_agent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'
        if user_agent is None:
            user_agent = def_user_agent
        headers = {
            'User-Agent': user_agent
        }
        return headers


def main():
    txt = f'Extraction of official FX rates of National Bank of Ukraine'\
          f' (c) 2022 Andy Reo, version {__version__}'
    today = datetime.now().strftime('%Y-%m-%d')
    parser = argparse.ArgumentParser(description=txt)
    parser.add_argument("currencies", type=str,
                        help="code(s) of currency - USD, EUR, ...")
    parser.add_argument("start_date", type=str, nargs='?',
                        help="start date in format yyyy-mm-dd", default=today)
    parser.add_argument("end_date", type=str, nargs='?',
                        help="end date in format yyyy-mm-dd", default=today)
    args = parser.parse_args()

    c = RateForPeriod(args.currencies.upper(),
                      start_date=args.start_date,
                      end_date=args.end_date)
    c.get_rates().save_xlsx(f'{c.filename}.xlsx')


if __name__ == '__main__':
    main()
