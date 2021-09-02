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


"""
Getting Trade data for each company and storing it
"""
wb = Workbook()
dic = {}
for company in companies:
    data = yf.download(company, start="2021-01-01", end="2021-08-15")
    dic[company] = data
    # print(data.iloc[0].Name)
    # break

"""
Writing to the excel sheet
"""
for company in dic:
    data = dic[company]
    sheet = wb.add_sheet(company)
    sheet.write(0, 0, 'Date')
    sheet.write(0, 1, 'Open')
    sheet.write(0, 2, 'High')
    sheet.write(0, 3, 'Low')
    sheet.write(0, 4, 'Close')
    sheet.write(0, 5, 'Adj Close')
    sheet.write(0, 6, 'Volume')
    for i in range(data.shape[0]):
        # sheet.write(i+1, 0, data.iloc[i]['Date'])
        sheet.write(i+1, 1, data.iloc[i]['Open'])
        sheet.write(i+1, 2, data.iloc[i]['High'])
        sheet.write(i+1, 3, data.iloc[i]['Low'])
        sheet.write(i+1, 4, data.iloc[i]['Close'])
        sheet.write(i+1, 5, data.iloc[i]['Adj Close'])
        sheet.write(i+1, 6, data.iloc[i]['Volume'])
    wb.save("Trade_data.xls")

wb.save("Trade_data.xls")
