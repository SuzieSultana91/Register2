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


ERF_url = "http://www.cleanenergyregulator.gov.au/DocumentAssets/Documents/Emissions%20Reduction%20Fund%20Register.csv"
ERF = pd.read_csv(ERF_url, encoding='utf-8', thousands=',')


# Revoked status
ERF['Revoked Status'] = ["Bad" if "Revoked" in ERF['Project Name'][i] else "Good" for i in range(len(ERF))]

# Final status
ERF['Final ACCUs issued'] = [int(str(ERF['ACCUs Total units issued'][i]).replace(',', '')) -
                             int(str(ERF['Total Number of KACCUs units relinquished'][i]).replace(',', '')) -
                             int(str(ERF['Total Number of NKACCUs units relinquished'][i]).replace(',', ''))
                             for i in range(len(ERF))]

# Brutto ma funziona
Dates = []
for i in range(len(ERF)):
    dates = ERF['Date Project Registered'][i].split('/')
    day = int(dates[0])
    month = int(dates[1])
    year = int(dates[2])
    Dates.append(datetime.datetime(year, month, day))
ERF['Dates'] = Dates
# NON funziona: scambia alcuni giorni con alcuni mesi
# ERF['Date Project Registered'] = pd.to_datetime(ERF['Date Project Registered'])


# Create workbook
ERF.to_excel('ERF - ' + str(choose_month()) + '(ERF).xlsx', sheet_name='ERF')

