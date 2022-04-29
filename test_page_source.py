from selenium import webdriver
from selenium.webdriver.chrome.options import Options

import time
import pytest
from selenium.webdriver.common.by import By

@pytest.mark.usefixtures("init_driver") 
class BaseTest():
    pass

class TestHeadLess(BaseTest):

    def test_getPageSource(self):
        options = Options()
        options.headless = True
        options.add_argument("--window-size=1920,1200")
        self.driver.get("https://giftshop.redcross.org.uk/collections/new-in")
        print(self.driver.page_source)
        self.driver.quit()