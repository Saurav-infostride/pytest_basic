import time
import pytest
from selenium.webdriver.common.by import By

@pytest.mark.usefixtures("init_driver") 
class BaseTest():
    pass

class TestHubSpot(BaseTest):
    @pytest.mark.parametrize(  
        "emailFrom, emailTo, emailSubject, emailMessage",
        [
            ("jhonny@yopmail.com", "saurav@yopmail.com", "Hello, their", "testing on Yopmail"),
        ]
    )
    def test_sendEmail(self, emailFrom, emailTo, emailSubject, emailMessage):
        self.driver.get("https://yopmail.com/en/")
        self.driver.find_element(By.XPATH, "//*[@id='login']").send_keys(emailFrom)
        self.driver.find_element(By.XPATH, "//*[@class='material-icons-outlined f36']").click()
        self.driver.find_element(By.XPATH, "//button[@id='newmail']").click()
        # time.sleep(5)
        self.driver.switch_to.frame("ifmail")
        self.driver.find_element(By.XPATH, "(//div[@class='inputsend']/input)[1]").send_keys(emailTo)
        self.driver.find_element(By.XPATH, "(//div[@class='inputsend']/input)[2]").send_keys(emailSubject)
        self.driver.find_element(By.XPATH, "//main[@class='yscrollbar']/div").send_keys(emailMessage)
        self.driver.find_element(By.XPATH, "//*[text()='Send the message']").click()
        time.sleep(5)
        message =(self.driver.find_element(By.XPATH, "//div[contains(text(),'Your message has')]").text)
        message_expected = "Your message has been sent"
        assert message == message_expected

    @pytest.mark.parametrize(  
        "enterEmail,",
        [
            ("saurav@yopmail.com"),
        ]
    )

    def test_receiveEmail(self,enterEmail):
        expected_subject = "Hello, their"
        expectedEmail = "<jhonny@yopmail.com>"
        self.driver.get("https://yopmail.com/en/")
        self.driver.find_element(By.XPATH, "//*[@id='login']").send_keys(enterEmail)
        self.driver.find_element(By.XPATH, "//*[@class='material-icons-outlined f36']").click()
        time.sleep(5)
        self.driver.switch_to.frame("ifmail")
        receivedEmailFrom = (self.driver.find_element(By.XPATH,"//span[@class='ellipsis b']").text)
        receivedSubject = (self.driver.find_element(By.XPATH, "//div[@class='ellipsis nw b f18']").text)
        assert receivedEmailFrom == expectedEmail
        assert receivedSubject == expected_subject


