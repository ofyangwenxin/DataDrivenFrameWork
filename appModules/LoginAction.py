# encoding=utf-8

from pageObjects.LoginPage import LoginPage

class LoginAction(object):
    def __init__(self):
        print "login..."

    @staticmethod
    def login(driver, username, password):
        try:
            login = LoginPage(driver)
            login.switchAccountLogin().click()

            login.switchToFrame()

            login.userNameObj().send_keys(username)
            login.passwordObj().send_keys(password)
            login.loginButton().click()
            login.switchToDefaultFrame()
        except Exception, e:
            raise e

if __name__ == '__main__':
    from selenium import webdriver
    import time

    driver = webdriver.Chrome(executable_path="E:\\driver\\chromedriver")
    driver.get("http://mail.126.com")
    time.sleep(5)

    LoginAction.login(driver, username="y284438799", password="15090027384")
    time.sleep(5)
    driver.quit()
