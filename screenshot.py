import pytest
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By

driver = webdriver.Chrome(ChromeDriverManager().install())
driver.implicitly_wait(10)
driver.delete_all_cookies()
driver.fullscreen_window()
driver.get("http://www.google.com")
driver.save_screenshot(r"C://Users//SauravSharma//Pytest//classes//screenshot//image.png")
