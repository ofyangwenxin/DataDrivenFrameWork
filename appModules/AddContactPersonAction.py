# encoding=utf-8

from pageObjects.HomePage import HomePage
from pageObjects.AddressBookPage import AddressBookPage
import time

class AddContactPerson(object):

    def __init__(self):
        print "add contact person."

    @staticmethod
    def add(driver, contactName, contactEmail, isStar, contactPhone, contactComment):
        try:
            # 创建主页实例对象
            hp = HomePage(driver)
            hp.addressLink().click()
            time.sleep(3)
            # 创建添加联系人页实例对象
            apb = AddressBookPage(driver)
            apb.createContactPersonButton().click()
            time.sleep(3)
            if contactName:
                apb.contactPersonName().send_keys(contactName)
            apb.contactPersonEmail().send_keys(contactEmail)
            if isStar == u"是":
                apb.starContacts().click()
            if contactPhone:
                apb.contactPersonMobile().send_keys(contactPhone)
            if contactComment:
                apb.contactPersonComment().send_keys(contactComment)
            apb.saveContactPerson().click()
        except Exception, e:
            raise e

if __name__ == '__main__':
    from LoginAction import LoginAction
    from selenium import webdriver

    driver = webdriver.Chrome(executable_path="E:\\driver\\chromedriver")
    driver.get("http://mail.126.com")
    time.sleep(5)

    LoginAction.login(driver, username="y284438799", password="15090027384")
    time.sleep(5)

    AddContactPerson.add(driver, u"张三", "zs@qq.com", u"是", "", "")
    time.sleep(3)
    assert u"张三" in driver.page_source
    driver.quit()