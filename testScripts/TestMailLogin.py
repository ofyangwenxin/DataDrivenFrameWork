# encoding=utf-8

from selenium import webdriver
from pageObjects.LoginPage import LoginPage
from appModules.LoginAction import LoginAction
import time


def testMailLogin():
    try:
        driver = webdriver.Chrome(executable_path="E:\\driver\\chromedriver")
        driver.get("http://mail.126.com")
        driver.implicitly_wait(5)
        driver.maximize_window()
        LoginAction.login(driver, username="y284438799", password="15090027384")
        time.sleep(5)
        assert u"未读邮件" in driver.page_source
    except Exception, e:
        raise e
    finally:
        driver.quit()


if __name__ == '__main__':
    testMailLogin()
    print u"登录126邮箱成功"
