import xlrd

xlrd.xlsx.ensure_elementtree_imported(False, None)
xlrd.xlsx.Element_has_iter = True

import time
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.webdriver.common.by import By

driver = webdriver.Chrome(ChromeDriverManager().install())
driver.implicitly_wait(10)

driver.get("https://yopmail.com/en/")

path = r"C://Users//SauravSharma//Pytest//classes//yopmailTestData.xlsx"
workbook = xlrd.open_workbook(path)
sheet=workbook.sheet_by_name("Sheet1")
rowCount = sheet.nrows
colCount = sheet.ncols
print(rowCount)
print(colCount)

email = driver.find_element(By.XPATH, "//*[@id='login']")
user_emails = driver.find_element(By.XPATH, "(//div[@class='inputsend']/input)[1]")
sub = driver.find_element(By.XPATH, "(//div[@class='inputsend']/input)[2]")
msg = driver.find_element(By.XPATH, "//main[@class='yscrollbar']/div")

for curr_row in range(1, rowCount):
    EmailF = sheet.cell_value(curr_row, 0)
    userEmails = sheet.cell_value(curr_row, 1)
    Subject = sheet.cell_value(curr_row, 2)
    Message = sheet.cell_value(curr_row, 3)
    print(userEmails , Subject , Message)
    email.send_keys(EmailF)
    driver.find_element(By.XPATH, "//*[@class='material-icons-outlined f36']").click()
    driver.find_element(By.XPATH, "//button[@id='newmail']").click()
    driver.switch_to.frame("ifmail")
    user_emails.send_keys(userEmails)
    sub.send_keys(Subject)
    msg.send_keys(Message)
    driver.find_element(By.XPATH, "//*[text()='Send the message']").click()
    print(userEmails, '+' , Subject ,'+', Message)
    time.sleep(1)
    successfullySentMessage =(driver.find_element(By.XPATH, "//div[contains(text(),'Your message has')]").text)
    message_expected = "Your message has been sent"
    assert successfullySentMessage == message_expected
    print(successfullySentMessage)