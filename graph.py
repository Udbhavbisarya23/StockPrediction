import os
from xlrd import open_workbook, sheet

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
        adj_matr[companies_map[company]][stockholder_map[stockholder]] = 1

for row in adj_matr:
    print(row)
