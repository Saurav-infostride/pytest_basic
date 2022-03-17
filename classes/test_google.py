import pytest
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By


"""  for report generation
      py.test -s test_fixture_google.py -v -s --html=google_test_report.html"""


driver = None

@pytest.fixture(scope='module')
def setup_module(module):
    global driver
    driver = webdriver.Chrome(ChromeDriverManager().install())
    driver.implicitly_wait(10)
    driver.delete_all_cookies()
    driver.fullscreen_window()
    driver.get("http://www.google.com")

def teardown_module(module):
    driver.quit()

def test_google_title():
    assert driver.title == "Google"

def test_google_url():
    assert driver.current_url == "https://www.google.com/?gws_rd=ssl"

