import pandas as pd
import xlsxwriter
import time


# Initializations
start = time.time()

# Registers
ERF_url = "http://www.cleanenergyregulator.gov.au/DocumentAssets/Documents/Emissions%20Reduction%20Fund%20Register.csv"
ERF = pd.read_csv(ERF_url, encoding='utf-8')
# encoding="ISO-8859-1"
# print(ERF.columns)

# Revoked Status
ERF['Revoked Status'] = ["Bad" if "Revoked" in ERF['Project Name'][i] else "Good" for i in range(len(ERF))]
# print(ERF['Revoked Status'])


# To Excel
# ERF.to_excel('ERF Register.xlsx', engine='xlsxwriter')

# x = 1945268493
# xr = "{:,.2f}".format(x)
# print(xr)
# print("Total cost is: {:,.2f}".format(x))

# Cosi' funziona!!!!
# number_format = workbook.add_format({'num_format': '#,##0.00'})
# worksheet.write('A1', 1234.56, number_format)




end = time.time()

print(end - start)



