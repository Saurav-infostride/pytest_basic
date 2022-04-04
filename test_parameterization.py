import time
import pytest
from selenium.webdriver.common.by import By

@pytest.mark.usefixtures("init_driver") 
class BaseTest():
    pass

class TestHubSpot(BaseTest):
    @pytest.mark.parametrize(  
        "username, password",
        [
            ("standard_user", "secret_sauce"),
        ]
    )
    def test_login(self, username, password):
        self.driver.get("https://www.saucedemo.com/")
        time.sleep(2)
        self.driver.find_element(By.ID, 'user-name').send_keys(username)
        time.sleep(2)
        self.driver.find_element(By.ID, 'password').send_keys(password)
        time.sleep(2)
        self.driver.find_element(By.ID, 'login-button').click()

    def test_homepage(self):
        time.sleep(2)
        self.driver.find_element(By.XPATH, "//button[@id='add-to-cart-sauce-labs-backpack']").click()
        time.sleep(2)
        self.driver.find_element(By.XPATH, "//a[@class='shopping_cart_link']").click()

    def test_cartpage(self):
        time.sleep(2)
        self.driver.find_element(By.XPATH, "//button[text()='Checkout']").click()

    def test_checkoutYourInfoPage(self):
        time.sleep(2)
        self.driver.find_element(By.XPATH, "(//*[@class='input_error form_input'])[1]").send_keys("Saurav")
        time.sleep(2)
        self.driver.find_element(By.XPATH, "(//*[@class='input_error form_input'])[2]").send_keys("Sharma")
        time.sleep(2)
        self.driver.find_element(By.XPATH, "(//*[@class='input_error form_input'])[3]").send_keys(123445)
        time.sleep(2)
        self.driver.find_element(By.ID, "continue").click()

    def test_checkoutOverviewPage(self):
        time.sleep(2)
        self.driver.find_element(By.XPATH, "//button[@id='finish']").click()

    def test_logout(self):
        self.driver.find_element(By.XPATH, '//div[@class="bm-burger-button"]').click()
        time.sleep(2)
        self.driver.find_element(By.XPATH, '//a[@id="logout_sidebar_link"]').click()
        time.sleep(2)
