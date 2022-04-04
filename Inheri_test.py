import pytest
import selenium


def test_ba():
    a = 3
    b = 4
    assert a == b

""" @pytest.mark.webtest -> this basically helps us to choose the test which we want to run
                just make a pytest.ini file and write the marker syntax init
            use py.test -m webtest  command in terminal to run the test cases"""

# @pytest.mark.webtest    
def test_b1():
    c = 3
    d = 4
    assert c == d

def test_b2():
    a = 4
    b = 4
    assert a == b
   
def test_b3():
    a = 3
    b = 4
    assert a == b, "Incorrect"

# @pytest.mark.webtest
def test_b4():
    assert "Saurav" == "SAURAV", "Should be equal"

# @pytest.mark.webtest
def test_login():
    assert "admin" == "admin"

