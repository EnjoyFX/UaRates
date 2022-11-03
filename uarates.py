import argparse
import json
import logging
from datetime import datetime, timedelta
from string import ascii_uppercase
from os.path import basename

import openpyxl
from openpyxl.styles import Font

import requests


__version__ = '1.0.2'
log_name = basename(__file__).split('.')[0]
logging.basicConfig(filename=f'{log_name}.log', level=logging.INFO,
                    format='%(asctime)s %(filename)s %(funcName)s %(message)s')
logger = logging.getLogger(log_name)

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
        self.headers = ['Date'] + self.currencies
        self.df = None  # will be data with rates

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

        self.df = data
        return self  # for chain of methods ability

    def save_xlsx(self, filename: str):
        if not self.df:
            e = '[Save error] DataFrame is empty'
            logger.warning(e)
            raise Exception(e)
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'Bank rates'  # set the name of sheet

        ws.append(self.headers)
        for r in self.df:
            ws.append(r)

        # set the bold font for header:
        for index, column_name in enumerate(self.headers):
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

    @staticmethod
    def _headers(user_agent=None):
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
