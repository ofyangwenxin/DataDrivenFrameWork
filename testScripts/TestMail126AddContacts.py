# encoding=utf-8

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from util.ParseExcel import ParseExcel
from config.VarConfig import *
from appModules.LoginAction import LoginAction
from appModules.AddContactPersonAction import AddContactPerson
import traceback
from time import sleep

# 设置此次测试的环境编码为utf-8
import sys
reload(sys)
sys.setdefaultencoding("utf-8")

# 将excel数据文件加载到内存
excelObj = ParseExcel()
excelObj.loadWorkBook(dataFilePath)

def LaunchBrowser():
    # 创建Chrome浏览器的一个Options实例对象
    chrome_options = Options()
    # 向Options实例中添加禁用扩展插件的设置参数项
    chrome_options.add_argument("--disable-extensions")
    # 添加屏蔽 --ignore-certificate-errors提示信息的设置参数项
    chrome_options.add_experimental_option("excludeSwitches", ["ignore-certificate-errors"])
    # 添加浏览器最大化的设置参数项，已启动最大化
    chrome_options.add_argument("--start-maximized")
    # 启动带有自定义设置的chrome浏览器
    driver = webdriver.Chrome(executable_path="E:\\driver\\chromedriver", chrome_options=chrome_options)
    driver.get("http://mail.126.com")
    sleep(3)
    return driver

def test126MailAddContacts():
    try:
        userSheet = excelObj.getSheetByName(u"126账号")
        isExecuteUser = excelObj.getColumn(userSheet, account_isExecute)
        dataBookColumn = excelObj.getColumn(userSheet, account_dataBook)
        print u"测试为126邮箱添加联系人执行开始..."
        for idx, i in enumerate(isExecuteUser[1:]):
            if i.value == 'y':
                userRow = excelObj.getRow(userSheet, idx+2)
                username = userRow[account_username - 1].value
                password = str(userRow[account_password - 1].value)
                print username, password

                # 创建浏览器实例对象
                driver = LaunchBrowser()

                # 登录
                LoginAction.login(driver, username, password)
                sleep(3)
                dataBookName = dataBookColumn[idx + 1].value
                dataSheet = excelObj.getSheetByName(dataBookName)
                isExecuteData = excelObj.getColumn(dataSheet, contacts_isExecute)

                contactNum = 0 # 记录添加成功联系人人数
                isExecuteNum = 0 # 记录需要执行联系人人数
                for id, data in enumerate(isExecuteData[1:]):
                    if data.value == "y":
                        isExecuteNum += 1
                        rowContent = excelObj.getRow(dataSheet, id+2)
                        contactPersonName = rowContent[contacts_contactPersonName - 1].value
                        contactPersonEmail = rowContent[contacts_contactPersonEmail - 1].value
                        isStar = rowContent[contacts_isStar - 1].value
                        contactPersonPhone = rowContent[contacts_contactPersonMobile - 1].value
                        contactPersonComment = rowContent[contacts_contactPersonComment - 1].value

                        # 添加联系人成功后，断言的关键字
                        assertKeyWord = rowContent[contacts_assertKeyWords - 1].value
                        print contactPersonName, contactPersonEmail, assertKeyWord
                        print contactPersonPhone, contactPersonComment, isStar

                        # 执行联系人操作
                        AddContactPerson.add(driver,
                                             contactPersonName,
                                             contactPersonEmail,
                                             isStar,
                                             contactPersonPhone,
                                             contactPersonComment)
                        sleep(1)

                        # 在联系人工作表中写入添加联系人执行时间
                        excelObj.writeCellCurrentTime(dataSheet,
                                                      rowNo=id + 2, colsNo=contacts_runTime)
                        try:
                            assert assertKeyWord in driver.page_source
                        except AssertionError, e:
                            # 断言失败，在联系人工作表中写入添加联系人测试失败信息
                            excelObj.writeCell(dataSheet, "faild", rowNo=id + 2,
                                               colsNo=contacts_textResult, style="red")
                        else:
                            # 断言成功，在联系人工作表中写入成功信息
                            excelObj.writeCell(dataSheet, "pass", rowNo=id+2,
                                               colsNo=contacts_textResult, style="green")
                            contactNum += 1
                print "contactNum=%s, isExecuteNum=%s" %(contactNum, isExecuteNum)
                if contactNum == isExecuteNum:
                    print u"为用户%s添加%d个联系人，测试通过！" %(username, contactNum)
                else:
                    excelObj.writeCell(userSheet, "faild", rowNo=idx+2,
                                       colsNo=account_testResult, style="red")
            else:
                print u"用户%s被设置为忽略执行！" %excelObj.getCellOfValue(userSheet, rowNo=idx+2,
                                                                colsNo=account_username)
        driver.quit()
    except Exception, e:
        print u"数据驱动框架主程序发生异常，异常信息为："
        print traceback.print_exc()


if __name__ == '__main__':
    test126MailAddContacts()
    print u"登录126邮箱成功!"


