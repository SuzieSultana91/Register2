import pandas as pd
import xlsxwriter
import time
import datetime
from openpyxl import Workbook
import xlsxwriter
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
import plotly.graph_objects as go

# Initializations
start = time.time()
today = datetime.datetime.now()
day = today.day
month = today.month
year = today.year
prev_month = datetime.date(2020, month - 1, 13)


def choose_month():
    return today.strftime("%B") if day > 8 else prev_month.strftime("%B")


# Registers
ERF_url = "http://www.cleanenergyregulator.gov.au/DocumentAssets/Documents/Emissions%20Reduction%20Fund%20Register.csv"
ERF = pd.read_csv(ERF_url, encoding='utf-8')
# encoding="ISO-8859-1"

# Revoked Status
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
# Non funziona : inverte alcuni giorni con mesi e viceversa
# ERF['Date Project Registered'] = pd.to_datetime(ERF['Date Project Registered'])

# ERF.to_excel('ERF Register(MonthlyReport).xlsx')

CAC_url = "http://www.cleanenergyregulator.gov.au/DocumentAssets/Documents/Carbon%20Abatement%20Contract%20table.csv"
CAC = pd.read_csv(CAC_url, encoding='utf-8')

# Numbers are strings
CAC['Volume of abatement committed under contract'] = \
    [int(str(CAC['Volume of abatement committed under contract'][i]).replace(',', '')) for i in range(len(CAC))]
CAC['Volume of abatement sold to the Commonwealth under contract'] = \
    [int(str(CAC['Volume of abatement sold to the Commonwealth under contract'][i]).replace(',', ''))
     for i in range(len(CAC))]

# CAC.to_excel('CAC Register(MonthlyReport).xlsx')

# CAC Register merged with ERF Register

# NON funziona
# BOTH = pd.merge(ERF, CAC, left_on='Contract ID', right_on='Carbon Abatement Contract ID')
# BOTH.to_excel("Both Registers.xlsx")

# FUNZIONA: la somma dei Volume committed and delivered is correct, BUT total ACCUs too little.
# BOTH = pd.merge(CAC, ERF, how='left')
# BOTH.to_excel('Combined Registers.xlsx')
# BOTH.join(ERF).to_excel('both2.xlsx')
# Carino: giusti i volume leggermente troppo grande total ACCU
# pd.merge(ERF, BOTH, how='outer').to_excel('Combined3.xlsx')

# Carino: giusti i volume leggermente troppo grande total ACCU
# merged = pd.merge(ERF, BOTH, how='outer')
# merged.loc[merged['Project ID'].isnull(), 'Project ID'].unique()

# OTHER TRY
# merged = pd.merge(ERF, CAC, how='outer', left_on='Contract ID', right_on='Carbon Abatement Contract ID')
# merged = merged.drop('Contract ID', 1) # drop duplicate info
# merged.to_excel('merged.xlsx')


# PERFETTO!!!!! MA ------ ho ancora delle ripetizioni, non le ho eliminate!!!
BOTH = pd.merge(CAC, ERF, how='left')
merged = pd.merge(ERF, BOTH, how='outer')
new_Final = [merged['Final ACCUs issued'][0]]
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
merged.to_excel('Combined Registers(MonthlyReport).xlsx')

# Nop ....
# CAC_merged_ERF = CAC.merge(ERF, left_on='Project ID', right_on='Project ID')
# CAC_merged_ERF.drop(["Scheme Participant", "Project Name", "Method Type", "Project Description",
#                      "Date Project Registered", "Project location (postcode)",
#                      "Project Area(s), where the project is an area based offsets project",
#                      "If the area-based project is covered by a regional natural resource management plan, "
#                      "is it consistent with that plan?", "Joint Implementation project",
#                      "Is the project area or project areas subject to a Carbon Maintenance Obligation (CMO)?",
#                      "Conditional upon all regulatory approvals being obtained",
#                      "Conditional upon the written consent of relevant interest holders",
#                      "Nominated Permanence Period, if applicable", "Finish date of permanence period, if applicable",
#                      "Contract ID", "KACCUs Total units issued", "KACCUs Units Issued in Financial Year 2012/13",
#                      "KACCUs Units Issued in Financial Year 2013/14", "KACCUs Units Issued in Financial Year 2014/15",
#                      "KACCUs Units Issued in Financial Year 2015/16", "KACCUs Units Issued in Financial Year 2016/17",
#                      "KACCUs Units Issued in Financial Year 2017/18", "KACCUs Units Issued in Financial Year 2018/19",
#                      "KACCUs Units Issued in Financial Year 2019/20", "Name of person/s to whom the KACCUs issued",
#                      "Total Number of KACCUs units relinquished", "NKACCUs Total units issued",
#                      "NKACCUs Units Issued in Financial Year 2012/13", "NKACCUs Units Issued in Financial Year 2013/14",
#                      "NKACCUs Units Issued in Financial Year 2014/15", "NKACCUs Units Issued in Financial Year 2015/16",
#                      "NKACCUs Units Issued in Financial Year 2016/17", "NKACCUs Units Issued in Financial Year 2017/18",
#                      "NKACCUs Units Issued in Financial Year 2018/19", "NKACCUs Units Issued in Financial Year 2019/20",
#                      "Name of person/s to whom the NKACCUs issued", "Total Number of NKACCUs units relinquished",
#                      "Notes"], axis=1, inplace=True)


# Voluntary Surrenders Register
Vol_url = "http://www.cleanenergyregulator.gov.au/DocumentAssets/Documents/Voluntary%20Cancellations.csv"
Vol = pd.read_csv(Vol_url, encoding='utf-8')
Vol.drop(['Unnamed: 4', 'Unnamed: 5'], axis=1, inplace=True)

Vol['Number of units'] = [int(str(Vol['Number of units'][i]).replace(',', '')) for i in range(len(Vol))]

# Create Excel workbook
workbook = xlsxwriter.Workbook(choose_month() + ' report(MonthlyReport).xlsx')
worksheet = workbook.add_worksheet('ERF Summary of projects')

Volume_committed_under_active_contract = sum([CAC['Volume of abatement committed under contract'][i]
                                              for i in range(len(CAC))
                                              if CAC['Status'][i] == "Active"])
Total_ACCUs_issued = sum([ERF['Final ACCUs issued'][i] for i in range(len(ERF)) if ERF['Revoked Status'][i] == 'Good'])
Total_volume_delivered_to_Commonwealth = sum([CAC['Volume of abatement sold to the Commonwealth under contract'][i]
                                              for i in range(len(CAC))
                                              if CAC['Status'][i] in ["Active", "Completed"]])
Total_number_of_projects = len([i for i in range(len(ERF)) if ERF['Revoked Status'][i] == "Good"])
Total_number_of_revoked_projects = len([i for i in range(len(ERF)) if ERF['Revoked Status'][i] == "Bad"])
Number_of_active_contracted_projects = len([i for i in range(len(CAC)) if CAC['Status'][i] == "Active"])

worksheet.set_column(0, 20, 35)
cell_format = workbook.add_format({'bold': True})
# cell_format = workbook.add_format({'bold': True, 'font_color': 'red'})
# cell_format.set_font_size(16)
# cell_format.set_align('center')
cell_format.set_underline()
cell_format.set_align('center')
worksheet.write('A1', 'Metric', cell_format)
worksheet.write('B1', 'Number of ACCUs', cell_format)
row = 1
col = 0

# worksheet.write('A1', 'Metric', bold)
# worksheet.write('B1', 'Number of ACCUs', bold)
# worksheet.write('A2', 'Volume committed under active contract')
# worksheet.write('B2', Volume_committed_under_active_contract)
# worksheet.write('A3', 'Total ACCUs issued')
# worksheet.write('B3', Total_ACCUs_issued)
# worksheet.write('A4', 'Total volume delivered to Commonwealth')
# worksheet.write('B4', Total_volume_delivered_to_Commonwealth)
# worksheet.write('A5', 'Total number of projects')
# worksheet.write('B5', Total_number_of_projects)
# worksheet.write('A6', 'Total number of revoked projects')
# worksheet.write('B6', Total_number_of_revoked_projects)
# worksheet.write('A6', 'Number of active contracted projects')
# worksheet.write('B6', Number_of_active_contracted_projects)

data = (
    # ['Volume committed under active contract', "{:,.0f}".format(Volume_committed_under_active_contract)],
    ['Volume committed under active contract', Volume_committed_under_active_contract],
    ['Total ACCUs issued', Total_ACCUs_issued],
    ['Total volume delivered to Commonwealth', Total_volume_delivered_to_Commonwealth],
    ['Total number of projects', Total_number_of_projects],
    ['Total number of revoked projects', Total_number_of_revoked_projects],
    ['Number of active contracted projects', Number_of_active_contracted_projects]
)
worksheet.set_column('B:B', 60)
worksheet.set_column('B:B', 60)

number_format = workbook.add_format({'num_format': '#,##0'})
number_format.set_align('center')
for name, number in data:
    worksheet.write(row, col, name, number_format)
    worksheet.write(row, col + 1, number, number_format)
    row += 1

# First attempt for Waterfall Graph

# Data to plot. Do not include a total, it will be calculated
index = ['Issued Accus', 'Delivered under CAC', 'Carbon pricing mechanism', 'Voluntary surrenders',
         'Safeguard', 'ACCU relinquishments']
data = {'amount': [350000, -30000, -7500, -25000, 95000, -7000]}

# Store data and create a blank series to use for the waterfall
trans = pd.DataFrame(data=data, index=index)
blank = trans.amount.cumsum().shift(1).fillna(0)

# Get the net total number for the final element in the waterfall
total = trans.sum().amount
trans.loc["Surplus"] = total
blank.loc["Surplus"] = total

# The steps graphically show the levels as well as used for label placement
step = blank.reset_index(drop=True).repeat(3).shift(-1)
step[1::3] = np.nan

# When plotting the last element, we want to show the full bar, set the blank to 0
blank.loc["Surplus"] = 0

# Plot and label
my_plot = trans.plot(kind='bar', stacked=True, bottom=blank, legend=None, figsize=(10, 5), title="2014 Sales Waterfall")
my_plot.plot(step.index, step.values, 'k')
my_plot.set_xlabel("Estimated ACCU surplus")

# Get the y-axis position for the labels
y_height = trans.amount.cumsum().shift(1).fillna(0)

# Get an offset so labels don't sit right on top of the bar
max = trans.max()
neg_offset = max / 25
pos_offset = max / 50
plot_offset = int(max / 15)

# Start label loop
loop = 0
for index, row in trans.iterrows():
    # For the last item in the list, we don't want to double count
    if row['amount'] == total:
        y = y_height[loop]
    else:
        y = y_height[loop] + row['amount']
    # Determine if we want a neg or pos offset
    if row['amount'] > 0:
        y += pos_offset
    else:
        y -= neg_offset
    my_plot.annotate("{:,.0f}".format(row['amount']), (loop, y), ha="center")
    loop += 1

# Scale up the y axis so there is room for the labels
my_plot.set_ylim(0, blank.max() + int(plot_offset))
# Rotate the labels
my_plot.set_xticklabels(trans.index, rotation=0)
my_plot.get_figure().savefig("waterfall.png", dpi=200, bbox_inches='tight')

# Second attempt for Waterfall Graph

Voluntary_surrender = sum([Vol['Number of units'][i] for i in range(len(Vol))
                           if Vol['Unit type'][i] in ['KACCU', 'NKACCU']])

net = (Total_ACCUs_issued / 1000000 - Total_volume_delivered_to_Commonwealth / 1000000 -
       14458807 / 1000000 - Voluntary_surrender / 1000000 - 531099 / 1000000 - 113617 / 1000000)
fig = go.Figure(go.Waterfall(
    name="20", orientation="v",
    measure=["relative", "relative", "relative", "relative", "relative", "relative", "total"],
    x=['Issued ACCUs', 'Delivered under CAC', 'Carbon pricing mechanism', 'Voluntary surrenders',
       'Safeguard', 'ACCU relinquishments', 'Surplus'],
    textposition="outside",
    text=[round(Total_ACCUs_issued / 1000000, 1), round(-Total_volume_delivered_to_Commonwealth / 1000000, 1),
          round(-14458807 / 1000000, 1), round(-Voluntary_surrender / 1000000, 1), round(-531099 / 1000000, 1),
          round(-113617 / 1000000, 1), round(net, 1)],
    y=[round(Total_ACCUs_issued / 1000000, 1), round(-Total_volume_delivered_to_Commonwealth / 1000000, 1),
       round(-14458807 / 1000000, 1), round(-Voluntary_surrender / 1000000, 1), round(-531099 / 1000000, 1),
       round(-113617 / 1000000, 1), round(0, 1)],
    connector={"line": {"color": "rgb(63, 63, 63)"}},
))

fig.update_layout(
    title="Estimated ACCUs Surplus",
    showlegend=True
)

# Plot
fig.show()

worksheet = workbook.add_worksheet('Surplus')
worksheet.set_column(0, 20, 25)
cell_format = workbook.add_format({'bold': True})
cell_format.set_underline()
cell_format.set_align('center')
worksheet.write('A1', 'Category', cell_format)
worksheet.write('B1', 'Amount', cell_format)
row = 1
col = 0

number_format = workbook.add_format({'num_format': '#,##0'})
number_format.set_align('center')
data = (
    ['Issued ACCUs', Total_ACCUs_issued],
    ['Delivered under CAC', Total_volume_delivered_to_Commonwealth],
    ['Carbon pricing mechanism', int(14458807)],
    ['Voluntary surrenders', Voluntary_surrender],
    ['Safeguard', int(531099)],
    ['ACCU relinquishments', int(113617)],
    ['Surplus', net * 1000000]
)
worksheet.set_column('B:B', 60)
worksheet.set_column('B:B', 60)

for name, number in data:
    worksheet.write(row, col, name, number_format)
    worksheet.write(row, col + 1, number, number_format)
    row += 1


# Add a format for Surplus. Light red fill with dark red text.
red_format = workbook.add_format({'bg_color': '#FFC7CE',
                                  'font_color': '#9C0006'})
# Apply conditional formats to the cell range.
worksheet.conditional_format('B2:B6', {'type':     'text',
                                       'criteria': 'containing',
                                       'value':    'Surplus',
                                       'format':   red_format})





workbook.close()


# To Excel
# ERF.to_excel('ERF Register.xlsx', engine='xlsxwriter')

end = time.time()

print(end - start)
