# encoding=utf-8

from action.PageAction import *
import time


def TestSendMailWithAttachment():
    open_browser("chrome")
    maximize_browser()
    visit_url("http://mail.126.com")
    time.sleep(3)
    assert_string_in_pagesource(u"126网易免费邮--你的专业电子邮")

    click("id", "switchAccountLogin")
    time.sleep(3)

    waitFrame_available_and_switch_to_it("tag_name", "iframe")

    input_string("xpath", "//input[@name='email']", "y284438799")
    input_string("xpath", "//input[@name='password']", "15090027384")

    click("id", "dologin")
    time.sleep(5)
    assert_title(u"网易邮箱")
    print u"登录成功"

    waitVisiblityOfElementLocated("xpath", "//span[text()='写 信']")
    click("xpath", "//span[text()='写 信']")
    input_string("xpath", "//div[contains(@id, '_mail_emailinput')]/input", "1376789881@qq.com")
    input_string("xpath", "//div[@aria-label='邮件主题输入框，请输入邮件主题']/input", u"新邮件")
    click("xpath", "//div[contains(@title, '600首MP3')]")
    time.sleep(3)
    paste_string(u"e:\\a.txt")
    press_enter_key()
    # 切换进邮件正文的frame
    waitFrame_available_and_switch_to_it("xpath", "//iframe[@tabindex=1]")
    input_string("xpath", "/html/body", u"发给光荣之光的一封信")
    switch_to_default_content()
    click("xpath", "//header//span[text()='发送']")
    time.sleep(3)
    assert_string_in_pagesource(u"发送成功")
    print u"邮件发送成功"
    close_browser()


if __name__ == '__main__':
    TestSendMailWithAttachment()
