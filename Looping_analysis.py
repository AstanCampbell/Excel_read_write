
# Reading an excel file using Python
import xlrd
import xlwt
import pandas as pd

# Give the location of the file
loc = ("C:Documents\\Hackathon\\textAnalysis.xls")
# To open Workbook
wb = xlrd.open_workbook('Planning_Data_Final2.xls', on_demand = True)
sheet = wb.sheet_by_name('Planning_Data_Final')


# For row 0 and column 0 the cells and columns are indexed just like arrays.
#analyses the cell which is indexed below

for row_index in range(1, sheet.nrows):
    str = sheet.cell(row_index,2).value
    units = [int(s) for s in str.split() if s.isdigit()]
    total_units = sum(units)
    print("-- Reference : \n" + sheet.cell(row_index,1).value)
    print("-- Proposal  : " + sheet.cell(row_index,2).value)
    print("-- Number of potential units  : ", total_units)
