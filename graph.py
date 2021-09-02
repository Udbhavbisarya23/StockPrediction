import os
import yfinance as yf
from xlrd import open_workbook, sheet
import xlwt
from xlwt import Workbook
import xlsxwriter

directory = os.path.join(os.path.dirname(__file__), 'Dataset/Stockholders')


# Get the list of all the stockholders and sort them in alphabetical order
all_stockholders = set(())
for filename in os.listdir(directory):
    company = open_workbook(os.path.join(directory, filename))
    sheet = company.sheet_by_index(0)
    for ind in range(1, 11, 1):
        all_stockholders.add(str(sheet.cell_value(ind, 1)))

stockholder_arr = []
for stockholder in all_stockholders:
    stockholder_arr.append(stockholder)
stockholder_arr.sort()


# Creating a map for stockholders to indices
stockholder_map = {}
for i in range(len(stockholder_arr)):
    stockholder_map[stockholder_arr[i]] = i


# Creating a map for companies to indices
companies_map = {}
ind = 0
for filename in os.listdir(directory):
    company = filename
    company = company[:-4]
    companies_map[company] = ind
    ind += 1

# Creating a map for companies total shares
'''Uncomment the following lines to create the excel sheet to get total shares sheet
companys_shares = {}
wb = Workbook()
sheet = wb.add_sheet("Total Shares")
count=0

for filename in os.listdir(directory):
    company = filename
    company = company[:-4]
    companys_shares[company] = yf.Ticker(company).info['sharesOutstanding']
    print(len(companys_shares),company,companys_shares[company])
    sheet.write(count,1,company)
    sheet.write(count,2,companys_shares[company])
    count+=1

wb.save("Total_Shares.xls")
'''

# Creating the map of company to total stocks
company_shares={}
wb=open_workbook("Total_Shares.xls")
sheet = wb.sheet_by_index(0)
for ind in range(0,503):
    company = str(sheet.cell_value(ind, 1))
    shares = sheet.cell_value(ind, 2)
    company_shares[company]=shares

'''Manually adding the missing shares'''
company_shares["BBWI"]=264750000
company_shares["OGN"]=253540000

# Create the adjaceny matrix with companies as rows and stockholders as columns
adj_matr = [[0]*len(stockholder_arr)
            for i in range(len(os.listdir(directory)))]


# Fill adjaceny matrix
for filename in os.listdir(directory):
    wb = open_workbook(os.path.join(directory, filename))
    sheet = wb.sheet_by_index(0)
    company = filename[:-4]
    for ind in range(1, 11, 1):
        stockholder = str(sheet.cell_value(ind, 1))
        shares = int(sheet.cell_value(ind, 2))
        adj_matr[companies_map[company]][stockholder_map[stockholder]] = shares/int(company_shares[company])

#Creating an Excel Sheet for savind corporation network
wb = xlsxwriter.Workbook('Corporation_network.xlsx')
sheet = wb.add_worksheet()
for i in range(len(stockholder_arr)):
    sheet.write(0,i+1,stockholder_arr[i])
count=0
for filename in os.listdir(directory):
    company = filename[:-4]
    sheet.write(count+1,0,company)
    count+=1
for i in range(len(companies_map)):
    for j in range(len(stockholder_map)):
        sheet.write(i+1,j+1,adj_matr[i][j])
wb.close()

print("Corporation network ready")