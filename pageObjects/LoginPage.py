# encoding=utf-8
from util.ObjectMap import *
from util.ParseConfigurationFile import ParseConfigFile


class LoginPage(object):

    def __init__(self, driver):
        self.driver = driver
        self.parseCF = ParseConfigFile()
        self.loginOptions = self.parseCF.getItemsSection("126mail_login")
        print self.loginOptions

    def switchToFrame(self):
        locateType, locatorExpression = self.loginOptions["loginPage.frame".lower()].split(">")
        self.driver.switch_to_frame(getElement(self.driver, locateType, locatorExpression))

    def switchToDefaultFrame(self):
        self.driver.switch_to_default_content()

    def switchAccountLogin(self):
        try:
            elementObj = getElement(self.driver, "id", "switchAccountLogin")
            return elementObj
        except Exception, e:
            raise e

    def userNameObj(self):
        try:
            locateType, locatorExpression = self.loginOptions["loginPage.username".lower()].split(">")
            elementObj = getElement(self.driver, locateType, locatorExpression)
            return elementObj
        except Exception, e:
            raise e

    def passwordObj(self):
        try:
            locateType, locatorExpression = self.loginOptions["loginPage.password".lower()].split(">")
            elementObj = getElement(self.driver, locateType, locatorExpression)
            return elementObj
        except Exception, e:
            raise e

    def loginButton(self):
        try:
            locateType, locatorExpression = self.loginOptions["loginPage.loginbutton".lower()].split(">")
            elementObj = getElement(self.driver, locateType, locatorExpression)
            return elementObj
        except Exception, e:
            raise e

if __name__ == '__main__':
    from selenium import webdriver
    driver = webdriver.Chrome(executable_path="E:\\driver\\chromedriver")
    driver.get("http://mail.126.com")
    import time
    time.sleep(5)

    login = LoginPage(driver)
    getElement(driver, "id", "switchAccountLogin").click()
    time.sleep(3)

    login.switchToFrame()

    login.userNameObj().send_keys("y284438799")
    login.passwordObj().send_keys("15090027384")
    login.loginButton().click()

    login.switchToDefaultFrame()
    time.sleep(5)
    driver.quit()