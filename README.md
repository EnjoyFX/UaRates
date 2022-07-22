# UArates

Extraction of official FX rates of National Bank of Ukraine
(based on API manual https://bank.gov.ua/ua/open-data/api-dev) with ability to save it as Excel-file

### Positional arguments:
* _**currencies**_  code(s) of currency - USD, EUR, ... (lower case **allowed**)

* _**start_date**_  start date in format yyyy-mm-dd

* _**end_date**_    end date in format yyyy-mm-dd

#### Options:
* _**-h, --help**_  show this help message and exit


#### Usage:
* get rates for EUR for today and save to Excel file:
```
uarates.py EUR
```
* get rates for USD from June 1, 2022 to June 30, 2022:
```
uarates.py USD 2022-06-01 2022-06-30
```
* get rates for USD and EUR from April 1, 2022 to June 30, 2022:
```
uarates.py USD,EUR 2022-04-01 2022-06-30
```