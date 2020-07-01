import pandas as pd
import time
import datetime

# Initializations
start = time.time()
today = datetime.datetime.now()
day = today.day
month = today.month
year = today.year
prev_month = datetime.date(2020, month - 1, 13)


def choose_month():
    return today.strftime("%B") if day > 8 else prev_month.strftime("%B")


CAC_url = "http://www.cleanenergyregulator.gov.au/DocumentAssets/Documents/Carbon%20Abatement%20Contract%20table.csv"
CAC = pd.read_csv(CAC_url, encoding='utf-8', thousands=',')


# # Doesn't work!!!
# try:
#     CAC['Actual contract end date'] = pd.to_datetime(str(CAC['Actual contract end date']))
# except:
#     "-"

# CAC['Actual contract end date'] = pd.to_datetime(CAC['Actual contract end date'])
CAC.to_excel('CAC - ' + str(choose_month()) + '(CAC).xlsx')
