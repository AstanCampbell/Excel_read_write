
# Reading an excel file using Python
import xlrd
import xlwt

# Give the location of the file
loc = ("C:Documents\\Hackathon\\textAnalysis.xls")
# To open Workbook
wb = xlrd.open_workbook('Planning_Data_Final2.xls', on_demand = True)
sheet = wb.sheet_by_name('Planning_Data_Final')

row_num = 13

# For row 0 and column 0 the cells and columns are indexed just like arrays.

#analyses the cell which is indexed below
str = sheet.cell(row_num,2).value
#print("-- Number of potential units  :")
units = [int(s) for s in str.split() if s.isdigit()]
#print(units)

total_units = sum(units)
#print(total_units)

for row_index in range(1, sheet.nrows):
    str = sheet.cell(row_index,2).value
    units = [int(s) for s in str.split() if s.isdigit()]
    total_units = sum(units)
    print("-- Reference : \n" + sheet.cell(row_index,1).value)
    print("-- Proposal  : " + sheet.cell(row_index,2).value)
    print("-- Number of potential units  : ", total_units)
