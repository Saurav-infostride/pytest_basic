import pytest
import selenium


def test_1():
    x = 20
    y = 10
    assert x == y 

def test_2():
    x = 31
    y = 31
    assert x == y

@pytest.mark.webtest   #marker
def test_3():
    orig_title = "Hello"
    expec_title = "Hello"
    assert orig_title == expec_title

def test_4():
    name = "selenium"
    title = "Selenium is used for web automation"
    assert name in title

def test_5():
    name = "jenkins"
    title = "Jenkins is CI server"
    assert name in title, "Title does not match"

@pytest.mark.webtest
def test_login():
    assert "admin" == "admin"