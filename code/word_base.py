
import os
from win32com.client import DispatchEx
import  config

class WordWrap:

    def __init__(self, templatefile=None):
        self.filename = templatefile
        self.wordApp = DispatchEx('Word.Application')
        if templatefile == None:
            self.wordDoc = self.wordApp.Documents.Add()
        else:
            self.wordDoc = self.wordApp.Documents.open(templatefile)
        #set up the selection
        self.wordDoc.Range(0,0) .Select()
        self.wordSel = self.wordApp.Selection
        # #fetch the styles in the document - see below
        # self.wordDoc.getStyleDictionary()
    def show(self):
        #自动化操作的界面显示出来
        self.wordApp.Visible = 1
    def saveAs(self, filename):
        #另存为
        self.wordDoc.SaveAs(filename)
    def printout(self):
        #打印出来
        self.wordDoc.PrintOut()
    def selectEnd(self):
        # 选中整个文档末尾
        self.wordSel.Collapse(0)

    def addText(self, text):
        self.wordSel.InsertAfter(text)
        self.selectEnd()
    def save(self):
        # self.wordDoc.SaveAs(self.filename,12,False,True,False,False,False,False,False,0,False,False,0,False)
        self.wordDoc.Save()

    def close(self):
        self.wordDoc.Close(0)
        self.wordApp.Quit()
    def replace_word(self,old_str,new_str):
        self.wordSel.Find.ClearFormatting()
        self.wordSel.Find.Replacement.ClearFormatting()
        self.wordSel.Find.Execute(old_str,False,False,False,False,False,True,1,True,new_str,2)


    def replace_text(self,old_str,new_str):
        #替换正文所有匹配内容.
        self.wordApp.ActiveWindow.ActivePane.View.SeekView = 0
        self.replace_word(old_str,new_str)
    def replace_headers_text(self,old_str,new_str):
        #替换页眉所有匹配内容
        self.wordApp.ActiveWindow.ActivePane.View.SeekView = 1
        self.wordApp.ActiveDocument.Sections[0].Headers[0].Range.Find.ClearFormatting()
        self.wordApp.ActiveDocument.Sections[0].Headers[0].Range.Find.Replacement.ClearFormatting()
        self.wordApp.ActiveDocument.Sections[0].Headers[0].Range.Find.Execute(old_str, False, False, False, False, False, True, 1, False, new_str, 2)
        self.wordApp.ActiveWindow.ActivePane.View.SeekView = 0

    def replace_footers_text(self,old_str,new_str):
        #替换页眉所有匹配内容
        self.wordApp.ActiveWindow.ActivePane.View.SeekView = 10
        self.wordApp.ActiveDocument.Sections[0].Footers[0].Range.Find.ClearFormatting()
        self.wordApp.ActiveDocument.Sections[0].Footers[0].Range.Find.Replacement.ClearFormatting()
        self.wordApp.ActiveDocument.Sections[0].Footers[0].Range.Find.Execute(old_str, False, False, False, False, False, True, 1, False, new_str, 2)
        self.wordApp.ActiveWindow.ActivePane.View.SeekView = 0

    def update_catalog(self):
        '''更新目录'''
        self.wordApp.ActiveWindow.ActivePane.View.SeekView = 0
        self.wordApp.ActiveDocument.Fields(1).Update()

    def all_copy(self):
        '''选中全文并进行拷贝'''

        self.wordSel.WholeStory()
        self.wordSel.Copy()

    def paste_end(self):
        '''在word末尾粘贴剪贴板的内容'''
        self.wordSel.WholeStory()
        self.wordSel.setRange(self.wordSel.end,self.wordSel.end)
        self.wordSel.PasteAndFormat(16)

    def all_cut(self):
        '''选中全文并剪切'''
        self.wordSel.WholeStory()
        self.wordSel.Cut()

    def word_before(self,search_word):

        '''找到匹配单词的前面'''
        self.wordSel.SetRange(0,0)
        self.wordSel.Find.Wrap = 0
        self.wordSel.Find.Text=search_word
        self.wordSel.Find.MatchCase = False
        self.wordSel.Find.MatchByte = True
        self.wordSel.Find.MatchWildcards = False
        self.wordSel.Find.MatchWholeWord = False
        self.wordSel.Find.MatchFuzzy = False
        self.wordSel.Find.Replacement.Text = ''
        self.wordSel.Find.Execute()
        self.wordSel.SetRange(self.wordSel.start,self.wordSel.start)

    def paste_origin(self):
        '''根据原有格式粘贴,不改变格式'''
        self.wordSel.PasteAndFormat(16)



def test_demo():
    '''测试案列'''
    w_app = WordWrap(r'C:\Users\Administrator\Desktop\2015.docx')

    w_app.replace_word('质量','zhiliang')
    w_app.replace_headers_text('质量','zhiliang')
    w_app.replace_footers_text('质量','zhiliang')
    w_app.replace_footers_text('27','31')
    w_app.update_catalog()
    w_app.saveAs(r'G:\development\myproject\stu_win32\hahahaha')
    w_app.close()

def handler_file(file_path):
    '''对一个文件进行处理和保存'''
    w_app = WordWrap(file_path)

    for text_word in config.text_list:
        w_app.replace_word(text_word[0],text_word[1])
    for page_header in config.page_header_list:
        w_app.replace_headers_text(page_header[0],page_header[1])
    for page_footer in config.page_footer_list:
        w_app.replace_footers_text(page_footer[0],page_footer[1])

    w_app.update_catalog()

    w_app.saveAs(file_path)
    w_app.close()

def handle_dir_below(dir):
    '''处理一个目录下的所有word'''
    if not os.path.exists(dir):
        print('不存在该目录')
        return

    for root_dir,dir_list,file_list in os.walk(dir):
        for filename in file_list:
            if os.path.splitext(filename)[1] in ['.docx','.doc','.wps','.dot','.dotx','.dotm','.docm']:
                handler_file(os.path.join(root_dir,filename))
                new_file_name=filename.replace(config.file_name[0],config.file_name[1])
                os.rename(os.path.join(root_dir,filename),os.path.join(root_dir,new_file_name))


def main():
    for handler_dir in config.dir_list:
        handle_dir_below(handler_dir)

if __name__ == '__main__':
    main()

