import xlrd
xlrd.xlsx.ensure_elementtree_imported(False, None)
xlrd.xlsx.Element_has_iter = True

import time
import pytest
from selenium.webdriver.common.by import By

@pytest.mark.usefixtures("init_driver") 
class BaseTest():
    pass

class TestHubSpot(BaseTest):

    def test_sendEmail(self):
        path = r"C://Users//SauravSharma//Pytest//classes//yopmailTestData.xlsx"
        workbook = xlrd.open_workbook(path)
        sheet=workbook.sheet_by_name("Sheet1")
        rowCount = sheet.nrows
        colCount = sheet.ncols
        print(rowCount) 
        print(colCount)
        for curr_row in range(1, rowCount):
            EmailF = sheet.cell_value(curr_row, 0)
            userEmails = sheet.cell_value(curr_row, 1)
            Subject = sheet.cell_value(1, 2)
            Message = sheet.cell_value(1, 3)
            self.driver.get("https://yopmail.com/en/")
            self.driver.find_element(By.XPATH, "//*[@id='login']").send_keys(EmailF)
            self.driver.find_element(By.XPATH, "//*[@class='material-icons-outlined f36']").click()
            self.driver.find_element(By.XPATH, "//button[@id='newmail']").click()
            time.sleep(2)
            self.driver.switch_to.frame("ifmail")
            self.driver.find_element(By.XPATH, "(//div[@class='inputsend']/input)[1]").send_keys(userEmails)
            self.driver.find_element(By.XPATH, "(//div[@class='inputsend']/input)[2]").send_keys(Subject)
            self.driver.find_element(By.XPATH, "//main[@class='yscrollbar']/div").send_keys(Message)
            self.driver.find_element(By.XPATH, "//*[text()='Send the message']").click()
            time.sleep(1)
            successfullySentMessage =self.driver.find_element(By.XPATH, 
                    "//div[contains(text(),'Your message has')]").text
            message_expected = "Your message has been sent"
            assert successfullySentMessage == message_expected
            print(successfullySentMessage)

    def test_receiveEmail(self):
        path = r"C://Users//SauravSharma//Pytest//classes//yopmailTestData.xlsx"
        workbook = xlrd.open_workbook(path)
        sheet=workbook.sheet_by_name("Sheet2")
        rowCount = sheet.nrows
        colCount = sheet.ncols
        print(rowCount)
        print(colCount)
        for curr_row in range(1, rowCount-1):
            userEmails = sheet.cell_value(curr_row, 0)
            EmailF = sheet.cell_value(1, 1)
            Subject = sheet.cell_value(1, 2)
            Message = sheet.cell_value(1, 3)
            self.driver.get("https://yopmail.com/en/")
            self.driver.find_element(By.XPATH, "//*[@id='login']").send_keys(userEmails)
            self.driver.find_element(By.XPATH, "//*[@class='material-icons-outlined f36']").click()
            time.sleep(3)
            self.driver.switch_to.frame("ifmail")
            time.sleep(1)
            receivedEmailFrom = self.driver.find_element(By.XPATH, "//span[@class='ellipsis b']").text
            time.sleep(1)
            receivedSubject = self.driver.find_element(By.XPATH, "//div[@class='ellipsis nw b f18']").text
            time.sleep(1)
            receivedMessage = self.driver.find_element(By.XPATH, "//*[text()='For testing purpose']").text
            assert receivedEmailFrom == EmailF
            assert receivedSubject == Subject
            assert receivedMessage == Message
            print('emailfrom :' + EmailF) 
            print('subject :' + Subject)
            print('message :' + Message)