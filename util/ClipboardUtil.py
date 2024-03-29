# encoding=utf-8
import win32clipboard as w
import win32con


class Clipboard(object):

    @staticmethod
    def getText():
        # 读取剪切板
        w.OpenClipboard()
        d = w.GetClipboardData(win32con.CF_TEXT)
        w.CloseClipboard()
        return d

    @staticmethod
    def setText(aString):
        # 设置剪切板内容
        w.OpenClipboard()
        w.EmptyClipboard()
        w.SetClipboardData(win32con.CF_UNICODETEXT, aString)
        w.CloseClipboard()
