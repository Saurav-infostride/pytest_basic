import pytest

# @py..fixture(params=["chrome", "firefox"],scope='class')
# def init_driver(request):
#     if request.param == "chrome":
#         web_driver = webdriver.Chrome(ChromeDriverManager().install())
#     if request.param == "firefox":
#         web_driver = webdriver.Firefox(executable_path=GeckoDriverManager().install())
#     request.cls.driver = web_driver
#     yield
#     web_driver.close()

'''In code below we have called the fixtures from conftest.py file....'''

@pytest.mark.usefixtures("init_driver") 
class BaseTest():
    pass

class Test_Google(BaseTest):
    def test_rediff_title(self):
        self.driver.get('http://www.rediff.com')
        assert self.driver.title == "Rediff.com: News | Rediffmail | Stock Quotes | Shopping"

    
    def test_rediff_url(self):
        self.driver.get('http://www.rediff.com')
        assert self.driver.current_url == "https://www.rediff.com/"