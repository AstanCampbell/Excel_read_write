# Reading an excel file using Python
import xlrd
import xlwt

# Give the location of the file
loc = ("C:Documents\\Hackathon\\textAnalysis.xls")
# To open Workbook
wb = xlrd.open_workbook('Planning_Data_Final2.xls', on_demand = True)
sheet = wb.sheet_by_name('Planning_Data_Final')
workbook = xlwt.Workbook()
sheet = workbook.add_sheet('Planning_Data_Final')

cell_num = 5
#sheet.write(cell_num,0,[int(s) for s in str.split() if s.isdigit()] )
#workbook.save('Planning_Data_Final2.xls')

# For row 0 and column 0 the cells and columns are indexed just like arrays.

print("-- Reference : \n" + sheet.cell(1,1).value)
print("-- Proposal  :\n" + sheet.cell(cell_num,2).value)
#analyses the cell which is indexed below
str = sheet.cell(cell_num,2).value
print("-- Number of potential units  :")
units = [int(s) for s in str.split() if s.isdigit()]
print(units)

#sheet = wb.add_sheet()
sheet.write(cell_num,0,'units')
workbook.save('Planning_Data_Final2.xls')

for row_index in range(0, sheet.nrows):
    Number_of_units = sheet.cell(row_index, 0).value
    Proposal = sheet.cell(row_index, 1).value

 #   print(Number_of_units + Proposal)
