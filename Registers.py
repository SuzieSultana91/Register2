import pandas as pd
import time
import datetime

today = datetime.datetime.now()
month = today.strftime("%B")



# IDEA: DataFrame.to_excel(excel_writer, sheet_name='Sheet1', na_rep='',
# float_format=None, columns=None, header=True, index=True, index_label=None,
# startrow=0, startcol=0, engine=None, merge_cells=True, encoding=None, inf_rep='inf',
# verbose=True, freeze_panes=None)

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
# ERF['Date Project Registered'] = pd.to_datetime(ERF['Date Project Registered'])

ERF.to_excel('ERF Register(Registers).xlsx')

CAC_url = "http://www.cleanenergyregulator.gov.au/DocumentAssets/Documents/Carbon%20Abatement%20Contract%20table.csv"
CAC = pd.read_csv(CAC_url, encoding='utf-8', thousands=',')

# No NEEDED anymore because of thousands=','
# CAC['Volume of abatement committed under contract'] = \
#     [int(str(CAC['Volume of abatement committed under contract'][i]).replace(',', '')) for i in range(len(CAC))]
# CAC['Volume of abatement sold to the Commonwealth under contract'] = \
#     [int(str(CAC['Volume of abatement sold to the Commonwealth under contract'][i]).replace(',', ''))
#      for i in range(len(CAC))]

# CAC['Actual contract end date'] = pd.to_datetime(CAC['Actual contract end date'])
CAC.to_excel('CAC Register(Registers).xlsx')

# Merged database
# PERFETTO!!!!! MA ------ ho ancora delle ripetizioni, non le ho eliminate!!!
BOTH = pd.merge(CAC, ERF, how='left')
merged = pd.merge(ERF, BOTH, how='outer')
new_Final = [merged['Final ACCUs issued'][0]]
# Alcuni ERF Project ID hanno due CAC ID civersi, per questo motivo vengono ripetuti (il che va bene, almeno non
# si perdono informazioni)
# Pero' devo stare attenta a non contarli due volte (i.e. quando sommo i Total ACCUs!!)
Repeated = ['']
for i in range(len(merged) - 1):
    if merged['Project ID'][i + 1] == merged['Project ID'][i]:
        new_Final.append(0)
        Repeated.append('Repeated')
    else:
        new_Final.append(merged['Final ACCUs issued'][i + 1])
        Repeated.append('')
merged['new_Final ACCUs issued'] = new_Final
merged['Repeated'] = Repeated



# Il problema era che non si riuscivano a raggruppare per anni/quarters/mesi le date in excel.
# NON funziona : scambia alcuni giorni con mesi
# merged['Date Project Registered'] = pd.to_datetime(merged['Date Project Registered'], format='%m-%d-%Y')

# merged['Dates'] = [ERF['Dates'][i] if i in range(len(ERF)) else 0 for i in range(len(merged))]
# print(merged['Dates'])

name = "Combined registered - " + str(today.strftime("%B")) +"(Registers).xlsx"
writer = pd.ExcelWriter(name, engine='xlsxwriter')
merged.to_excel(writer, sheet_name='Sheet1')
workbook = writer.book
worksheet = writer.sheets['Sheet1']
format1 = workbook.add_format({'num_format': '#,##0.00'})
# Setting the format but not the column width
# worksheet.set_column('J:J', None, format1)



Vol_url = "http://www.cleanenergyregulator.gov.au/DocumentAssets/Documents/Voluntary%20Cancellations.csv"
Vol = pd.read_csv(Vol_url, encoding='utf-8')
Vol.drop(['Unnamed: 4', 'Unnamed: 5'], axis=1, inplace=True)

Vol['Number of units'] = [int(str(Vol['Number of units'][i]).replace(',', '')) for i in range(len(Vol))]

Vol.to_excel('Voluntary Surrenders(Registers).xlsx')

writer.save()