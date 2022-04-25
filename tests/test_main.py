import os
from unittest import TestCase

import pandas as pd

from main import RateForPeriod


class TestRateForPeriod(TestCase):
    def test_get_rates(self):
        t = RateForPeriod('EUR', '2022-01-01', '2022-01-02')
        expected = pd.DataFrame([['2022-01-01', 30.9226],
                                 ['2022-01-02', 30.9226]],
                                columns=['Date', 'EUR'])
        received = t.get_rates()
        assert received.equals(expected)

    def test_save_xlsx(self):
        t = RateForPeriod('EUR', '2022-01-01', '2022-01-02')
        expected = 'rates_EUR_2022-01-01_2022-01-02'
        if os.path.exists(expected):
            os.remove(expected)  # remove if file with expected name exists
        received = t.filename
        assert expected == received, "Filenames are not the same"

        t = RateForPeriod('EUR', '2022-01-02', '2022-01-01')  # reversed dates
        received = t.filename
        assert expected == received, "Dates are not reversed!"
        t.get_rates()
        assert expected == t.save_xlsx(received), "Error during save file"
        assert os.path.exists(expected), "File not saved!"
        os.remove(expected)

    def test__get_rate_per_date(self):
        t = RateForPeriod('EUR', '2022-01-01', '2022-01-01')
        received = t._get_rate_per_date('EUR', '20220101')
        self.assertEqual(str(received), '30.9226')

    def test__headers(self):
        t = RateForPeriod('EUR', '2022-01-01', '2021-12-31')
        test_agent = 'test agent'
        value = t._headers(user_agent=test_agent)
        self.assertEqual(value, {'User-Agent': test_agent})
        value = t._headers(user_agent=None)
        self.assertEqual(value, {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
