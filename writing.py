import xlwt
import xlrd


from xlwt import Workbook
wb = Workbook()
sheet1 = wb.add_sheet('total')
sheet1.write(3, 0, 7)
wb.save('stackoverflow.xls')
