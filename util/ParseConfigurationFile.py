# encoding=utf-8
from ConfigParser import ConfigParser
from config.VarConfig import pageElementLocatorPath

class ParseConfigFile(object):

    def __init__(self):
        self.cf = ConfigParser()
        self.cf.read(pageElementLocatorPath)

    def getItemsSection(self, sectionName):
        # 获取配置文件中指定section下的所有option键值对, 并以字典类型返回给调用者
        # 使用self.cf.items(sectionName)这种方法获取到的配置文件中的options内容均被转换成小写
        optionsDict = dict(self.cf.items(sectionName))
        return optionsDict

    def getOptionValue(self, sectionName, optionName):
        # 获取指定section下的指定option值
        value = self.cf.get(sectionName, optionName)
        return value

if __name__ == '__main__':
    pc = ParseConfigFile()
    print pc.getItemsSection("126mail_login")
    print pc.getOptionValue("126mail_login", "loginPage.frame")