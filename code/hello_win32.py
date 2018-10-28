# import win32api
# import win32con
# win32api.MessageBox(win32con.NULL, 'Python 你好！', '你好', win32con.MB_OK)
# import win32api
# import win32con
# keyname='Software\Microsoft\Internet Explorer\Main'
# page = 'www.sina.com.cn'
# title = 'I love sina web site!'
# search_page = 'http://www.baidu.com'
# key = win32api.RegOpenKey(win32con.HKEY_CURRENT_USER, keyname, 0, win32con.KEY_ALL_ACCESS)
# win32api.RegSetValueEx(key, 'Start Page', 0, win32con.REG_SZ, page)
# win32api.RegSetValueEx(key, 'Window Title', 0, win32con.REG_SZ, title)
# win32api.RegSetValueEx(key, 'Search Page', 0, win32con.REG_SZ, search_page)


# import win32com
# from win32com.client import Dispatch, constants,DispatchEx
#
# w = win32com.client.Dispatch('Wps.Application')
# w = win32com.client.DispatchEx('Word.Application')
# doc = w.Documents.Open( FileName =r'C:\Users\Administrator\Desktop\Sublime Text3快捷键.md')

#模块引用
import win32com
from win32com.client import Dispatch,DispatchEx
#打开word文档
word= Dispatch('Word.Application',userName = "Administrator")
# worddoc = word.Documents.Add()
# print(word.Visiable)
# word.Visiable=1
# path=r"C:\Users\Administrator\Desktop\Sublime Te
worddoc = word.Documents.Add()
# #中文路径乱码问题处理
# path="c:/文档.docx"
# FileName=path.decode("utf8")
# #读取表格
# table = doc.Tables[0]
# #表格插入行
# table.Cell(0,0).Select()
# word.Selection.InsertRowsBelow(1) #当前行的下面插入一行
# #向表格中填写内容
# table.Cell(0,0).Range.Text='abc'
# str = "你好"
# #中文写入乱码处理
# table.Cell(0,1).Range.Text=str.decode("utf8")
# table.Cell(0,2).Range.Text=(u'%s' % str)
# #文档另存为
# path="c:/result.docx"
# doc.SaveAs(path)
