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
            ("jhonny@yopmail.com", "Saurav@yopmail.com", "Ignore it, only for testing purpose", "For testing purpose"),
            ("jhonny@yopmail.com", "Bob@yopmail.com", "Ignore it, only for testing purpose", "For testing purpose"),
            ("jhonny@yopmail.com", "Jake@yopmail.com", "Ignore it, only for testing purpose", "For testing purpose"),
            ("jhonny@yopmail.com", "Jack@yopmail.com", "Ignore it, only for testing purpose", "For testing purpose"),
            ("jhonny@yopmail.com", "Tommy@yopmail.com", "Ignore it, only for testing purpose", "For testing purpose"),
        ]
    )
    def test_Yop_sendEmail(self, emailFrom, emailTo, emailSubject, emailMessage):
        self.driver.get("https://yopmail.com/en/")
        self.driver.find_element(By.XPATH, "//*[@id='login']").send_keys(emailFrom)
        self.driver.find_element(By.XPATH, "//*[@class='material-icons-outlined f36']").click()
        self.driver.find_element(By.XPATH, "//button[@id='newmail']").click()
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
            ("Bob@yopmail.com"),
            ("Jake@yopmail.com"),
            ("Jack@yopmail.com"),
            ("Tommy@yopmail.com"),
        ]
    )

    def test_receiveEmail(self,enterEmail):
        expected_subject = "Ignore it, only for testing purpose"
        expectedEmail = "<jhonny@yopmail.com>"
        self.driver.get("https://yopmail.com/en/")
        self.driver.find_element(By.XPATH, "//*[@id='login']").send_keys(enterEmail)
        self.driver.find_element(By.XPATH, "//*[@class='material-icons-outlined f36']").click()
        self.driver.switch_to.frame("ifmail")
        receivedEmailFrom = (self.driver.find_element(By.XPATH,"//span[@class='ellipsis b']").text)
        receivedSubject = (self.driver.find_element(By.XPATH, "//div[@class='ellipsis nw b f18']").text)
        assert receivedEmailFrom == expectedEmail
        print(receivedEmailFrom)
        assert receivedSubject == expected_subject
        print(receivedSubject)