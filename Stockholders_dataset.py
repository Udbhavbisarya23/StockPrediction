import yfinance as yf
import pandas as pd
import xlwt
from xlwt import Workbook


"""
Getting list of S&P 500 companies
"""
payload = pd.read_html(
    'https://en.wikipedia.org/wiki/List_of_S%26P_500_companies')
first_table = payload[0]
companies = first_table['Symbol'].values.tolist()
print(companies)


"""
Getting institutional holders for each company
"""
dic = {}
empty=set()
for company in companies:
    holders = yf.Ticker(company).institutional_holders
    dic[company] = holders
    print(len(dic),company)
    print(holders)

for company in dic:
    try:
        wb = Workbook()
        holders = dic[company]
        sheet = wb.add_sheet(company)
        sheet.write(0, 1, 'Holder')
        sheet.write(0, 2, 'Shares')
        sheet.write(0, 3, 'Date Reported')
        sheet.write(0, 4, '% Out')
        sheet.write(0, 5, 'Value')
        for i in range(10):
            sheet.write(i+1, 1, holders.iloc[i]['Holder'])
            sheet.write(i+1, 2, str(holders.iloc[i]['Shares']))
            sheet.write(i+1, 3, holders.iloc[i]['Date Reported'])
            sheet.write(i+1, 4, holders.iloc[i]['% Out'])
            sheet.write(i+1, 5, str(holders.iloc[i]['Value']))
        wb.save(company+".xls")
    except:
        empty.add(company)

print(empty)

