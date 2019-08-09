# encoding=utf-8

from util.ObjectMap import *
from util.ParseConfigurationFile import ParseConfigFile

class HomePage(object):

    def __init__(self, driver):
        self.driver = driver
        self.parseCF = ParseConfigFile()

    def addressLink(self):
        try:
            locateType, locatorExpression = self.parseCF.getOptionValue("126mail_homePage", "homePage.addressbook").split(">")
            elementObj = getElement(self.driver, locateType, locatorExpression)
            return elementObj
        except Exception, e:
            raise e