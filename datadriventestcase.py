import time
import xlrd
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager

driver = webdriver.Chrome(ChromeDriverManager().install())

driver.get("http://orangehrm.qedgetech.com/")
driver.maximize_window()

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

    driver.find_element_by_xpath("//input[@id='txtUsername']").send_keys(username)
    driver.find_element_by_xpath("//input[@id='txtPassword']").send_keys(username)
    time.sleep(5)
    driver.find_element_by_id("btnLogin").click()

    print(username + " " + password)