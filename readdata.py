import xlrd
import XLUtils
path = r"C://Users//SauravSharma//Pytest//classes//TestData.xlsx"

workbook = xlrd.open_workbook(path)
sheet=workbook.sheet_by_name("login")

rowCount = sheet.nrows
cols = sheet.ncols

print(rowCount)
print(cols)

for curr_row in range(1, rowCount):
    username = sheet.cell_value(curr_row, 0)
    password = sheet.cell_value(curr_row, 1)

print(username + " " + password)