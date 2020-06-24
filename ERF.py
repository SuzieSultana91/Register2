import pandas as pd

ERF_url = "http://www.cleanenergyregulator.gov.au/DocumentAssets/Documents/Emissions%20Reduction%20Fund%20Register.csv"
ERF = pd.read_csv(ERF_url, encoding='utf-8', thousands=',')


# Revoked status
ERF['Revoked Status'] = ["Bad" if "Revoked" in ERF['Project Name'][i] else "Good" for i in range(len(ERF))]

# Final status
ERF['Final ACCUs issued'] = [int(str(ERF['ACCUs Total units issued'][i]).replace(',', '')) -
                             int(str(ERF['Total Number of KACCUs units relinquished'][i]).replace(',', '')) -
                             int(str(ERF['Total Number of NKACCUs units relinquished'][i]).replace(',', ''))
                             for i in range(len(ERF))]

ERF['Date Project Registered'] = pd.to_datetime(ERF['Date Project Registered'])

ERF.to_excel('ERF.xlsx')