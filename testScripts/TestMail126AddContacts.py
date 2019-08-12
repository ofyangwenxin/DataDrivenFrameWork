# encoding=utf-8

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from util.ParseExcel import ParseExcel
from config.VarConfig import *
from appModules.LoginAction import LoginAction
from appModules.AddContactPersonAction import AddContactPerson
import traceback
from time import sleep
from util.Log import *

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
    logging.info(u"126邮箱添加联系人数据驱动测试开始...")
    try:
        userSheet = excelObj.getSheetByName(u"126账号")
        # 获取126账号sheet表中是否执行列
        isExecuteUser = excelObj.getColumn(userSheet, account_isExecute)
        # 获取126账号sheet表中数据表列
        dataBookColumn = excelObj.getColumn(userSheet, account_dataBook)
        print u"测试为126邮箱添加联系人执行开始..."
        for idx, i in enumerate(isExecuteUser[1:]):
            if i.value == 'y':  # 要执行
                userRow = excelObj.getRow(userSheet, idx+2)
                username = userRow[account_username - 1].value
                password = str(userRow[account_password - 1].value)
                print username, password

                # 创建浏览器实例对象
                driver = LaunchBrowser()
                logging.info(u"启动浏览器，访问126邮箱主页")

                # 登录
                LoginAction.login(driver, username, password)
                sleep(3)
                try:
                    assert u"收 信" in driver.page_source
                    logging.info(u"用户%s登录后，断言页面关键字'收信'成功" % username)
                except AssertionError, e:
                    logging.debug(u"用户%s登录后， 断言页面关键字'收信'失败，",
                                  u"异常信息：%s" % (username, str(traceback.format_exc())))
                dataBookName = dataBookColumn[idx + 1].value
                # 获取账号对应的联系人sheet页数据
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
                        logging.info(u"添加联系人%s成功" % contactPersonEmail)

                        # 在联系人工作表 中写入添加联系人执行时间
                        excelObj.writeCellCurrentTime(dataSheet,
                                                      rowNo=id + 2, colsNo=contacts_runTime)
                        try:
                            assert assertKeyWord in driver.page_source
                        except AssertionError, e:
                            # 断言失败，在联系人工作表中写入添加联系人测试失败信息
                            excelObj.writeCell(dataSheet, "faild", rowNo=id + 2,
                                               colsNo=contacts_textResult, style="red")
                            logging.info(u"断言关键字'%s'失败" % assertKeyWord)
                        else:
                            # 断言成功，在联系人工作表中写入成功信息
                            excelObj.writeCell(dataSheet, "pass", rowNo=id+2,
                                               colsNo=contacts_textResult, style="green")
                            contactNum += 1
                            logging.info(u"断言关键字'%s'成功" % assertKeyWord)
                    else:
                        logging.info(u"联系人%s被忽略执行" % contactPersonEmail)
                if contactNum == isExecuteNum:
                    excelObj.writeCell(userSheet, "pass", rowNo=idx + 2,
                                       colsNo=account_testResult, style="green")
                else:
                    excelObj.writeCell(userSheet, "faild", rowNo=idx+2,
                                       colsNo=account_testResult, style="red")
                logging.info(u"为用户%s添加%d个联系人，%d个成功\n" % (username, isExecuteNum, contactNum))
            else:
                ignoreUsername = excelObj.getCellOfValue(userSheet,
                                                         rowNo=idx + 2, colsNo=account_username)
                logging.info(u"用户%s被忽略执行\n" % ignoreUsername)
        driver.quit()
    except Exception, e:
        logging.debug(u"数据驱动框架主程序执行过程发生异常，异常信息: %s" % str(traceback.format_exc()))


if __name__ == '__main__':
    test126MailAddContacts()
    print u"登录126邮箱成功!"


