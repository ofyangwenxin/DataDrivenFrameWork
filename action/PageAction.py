# encoding=utf-8

from selenium import webdriver
from config.VarConfig import chromeDriverFilePath
from util.ObjectMap import getElement
from util.ClipboardUtil import Clipboard
from util.KeyBoardUtil import KeyboardKeys
from util.DirAndTime import *
from util.WaitUtil import WaitUtil
from selenium.webdriver.chrome.options import Options
import time

# 定义全局driver变量
driver = None
# 定义全局的等待类实例对象
waitUtil = None


def open_browser(browserName, *arg):
    global driver, waitUtil
    try:
        if (browserName.lower() == 'ie'):
            # driver = webdriver.Ie(executable_path=ieDriverFilePath)
            pass
        elif browserName.lower() == 'chrome':
            chrome_options = Options()
            # 添加屏蔽 --ignore-certificate-errors提示信息的设置参数项
            chrome_options.add_experimental_option("excludeSwitches", ["ignore-certificate-errors"])
            driver = webdriver.Chrome(executable_path=chromeDriverFilePath, chrome_options=chrome_options)
        else:
            # driver = webdriver.Firefox(executable_path=filefoxDriverFilePath)
            pass
        waitUtil = WaitUtil(driver)
    except Exception, e:
        raise e


def visit_url(url, *arg):
    # 访问某个网址
    global driver
    try:
        driver.get(url)
    except Exception, e:
        raise e


def close_browser(*arg):
    # 关闭浏览器
    global driver
    try:
        driver.quit()
    except Exception, e:
        raise e


def sleep(sleepSeconds, *arg):
    # 强制等待
    try:
        time.sleep(int(sleepSeconds))
    except Exception, e:
        raise e


def clear(locationType, locatorExpression, *arg):
    # 清除输入框默认内容
    global driver
    try:
        getElement(driver, locationType, locatorExpression)
    except Exception, e:
        raise e


def input_string(locationType, locatorExpressin, inputContent):
    # 在页面输入框中输入内容
    global driver
    try:
        getElement(driver, locationType, locatorExpressin).send_keys(inputContent)
    except Exception, e:
        raise e


def click(locationType, locatorExpression, *arg):
    # 单击页面元素
    global driver
    try:
        getElement(driver, locationType, locatorExpression).click()
    except Exception, e:
        raise e


def assert_string_in_pagesource(assertString, *arg):
    # 断言页面源码是否存在某关键字或关键字符串
    global driver
    try:
        assert assertString in driver.page_source, u"%s not found in page source!" % assertString
    except AssertionError, e:
        raise AssertionError(e)


def assert_title(titleStr, *arg):
    # 断言页面标题是否存在给定的关键字符串
    global driver
    try:
        assert titleStr in driver.title, u"%s not found in title!" % titleStr
    except Exception, e:
        raise e


def getTitle(*arg):
    # 获取页面标题
    global driver
    try:
        return driver.title
    except Exception, e:
        raise e


def getPageSource(*arg):
    # 获取页面源码
    global driver
    try:
        return driver.page_source
    except Exception, e:
        raise e


def switch_to_frame(locationType, frameLocatorExpression, *arg):
    # 切换进入frame
    global driver
    try:
        driver.switch_to.frame(getElement(driver, locationType, frameLocatorExpression))
    except Exception, e:
        raise e


def switch_to_default_content(*arg):
    # 切出 frame
    global driver
    try:
        driver.switch_to.default_content()
    except Exception, e:
        raise e


def paste_string(pasteString, *arg):
    # 模拟Ctrl + V操作
    try:
        Clipboard.setText(pasteString)
        time.sleep(2)
        KeyboardKeys.twoKeys("ctrl", "v")
    except Exception, e:
        raise e


def press_tab_key(*arg):
    # 模拟Tab键
    try:
        KeyboardKeys.oneKey("tab")
    except Exception, e:
        raise e


def press_enter_key(*arg):
    # 模拟Enter键
    try:
        KeyboardKeys.oneKey("enter")
    except Exception, e:
        raise e


def maximize_browser():
    # 窗口最大化
    global driver
    try:
        driver.maximize_window()
    except Exception, e:
        raise e


def capture_screen(*arg):
    # 截取屏幕图片
    global driver
    currTime = getCurrentTime()
    picNameAndPath = str(createCurrentDateDir() + "\\" + str(currTime) + ".png")
    try:
        driver.get_screenshot_as_file(picNameAndPath.replace('\\', r'\\'))
    except Exception, e:
        raise e
    else:
        return picNameAndPath


def waitPresenceOfElementLocated(locationType, locatorExpression, *arg):
    # 显示等待页面元素出现在DOM中，但不一定可见， 存在返回该页面元素对象
    global waitUtil
    try:
        waitUtil.presenceOfElementLocated(locationType, locatorExpression)
    except Exception, e:
        raise e


def waitFrame_available_and_switch_to_it(locationType, locatorExpression, *arg):
    # 检查frame是否存在，存在则切换进frame控件中
    global waitUtil
    try:
        waitUtil.frame_available_and_switch_to_it(locationType, locatorExpression)
    except Exception, e:
        raise e


def waitVisiblityOfElementLocated(locationType, locatorExpression, *arg):
    # 显示等待页面元素的出现，并且可见，存在返回该页面元素对象
    global waitUtil
    try:
        waitUtil.visibility_element_located(locationType, locatorExpression)
    except Exception, e:
        raise e
