import win32com.client as win32
from win32com.client import Dispatch

class ReplaceToMaker:
    def __init__(self) :
        self.word = win32.gencache.EnsureDispatch("Word.application")
        self.excel = Dispatch("Excel.Application") #?óë??? ?îÑÎ°úÍ∑∏?û® ?ã§?ñâ
        self.excel.Visible = False
        self.word.Visible = False
        

    def ActiveVisible(self) :
        self.excel.Visible = True #?ïû?úºÎ°? ?ã§?ñâÍ≥ºÏ†ï?ùÑ Î≥¥Ïù¥Í≤?
        self.word.Visible = True
    def DeactiveVisible(self):
        self.excel.Visible = False
        self.word.Visible = False

    def ReplaceToText(self, path, mark, text, target_path=""):
        if(path[-4:] == "xlsx"):
            self.__search_replace_all_excel(path,mark,text,target_path)
            return 0
        elif(path[-4:] == "docx"):
            self.__search_replace_all_word(path,mark,text,target_path)
            return 0
        else:
            return -1

    def __search_replace_all_word(self, path, find_str, replace_str,target_path=""):
        doc = self.word.Documents.Open(path)
        maintext = self.word.ActiveDocument.StoryRanges(1)
        a = maintext.Text.count(find_str) # Î≥∏Î¨∏ Ï§? Ï∞æÏùÑ ?ã®?ñ¥?ùò ?àòÎ•? ?Ñº?ã§.

        for i in range(0, a):
            self.word.Selection.GoTo(What=win32.constants.wdGoToSection, Which=win32.constants.wdGoToFirst)
            self.word.Selection.Find.Text = find_str # Ï∞æÏùÑ ?ã®?ñ¥Î•? Ï∞æÎäî?ã§.
            self.word.Selection.Find.Replacement.Text = "" # Ï∞æÏùÑ ?ã®?ñ¥Î•? Ïß??ö¥?ã§.
            self.word.Selection.Find.Execute(Replace=1, Forward=True)
            self.word.Selection.InsertAfter(replace_str) # ?ï¥?ãπ ?úÑÏπòÏóê ?ÇΩ?ûÖ?ïòÍ≥†Ïûê ?ïò?äî ?ã®?ñ¥Î•? ?ûÖ?†•?ïú?ã§.

        if(target_path == ""):
            doc.Save()
        else:
            doc.SaveAs(target_path)

    def __search_replace_all_excel(self, path, find_str, replace_str, target_path=""):
        workbook = self.excel.Workbooks.Open(path)

        for worksheet in workbook.Worksheets:
            cell = worksheet.UsedRange.Find(find_str)
            if cell:
                rg = worksheet.Range(worksheet.usedRange.Address)
                rg.Replace(find_str, replace_str)
        
        if(target_path == ""):
            workbook.Save()
        else:
            workbook.SaveAs(target_path)

class ReadWriteExecl:
    def __init__(self) :
        self.excel = Dispatch("Excel.Application") #?óë??? ?îÑÎ°úÍ∑∏?û® ?ã§?ñâ
    
    def read_mark(self,path):
        workbook = self.excel.Workbooks.Open(path)
        ws = workbook.Worksheets(1)
        count = ws.Cells(1,2).Value
        result_dict = {}
        for i in range(int(count)) :
            result_dict[ws.Cells(i+2,1).Value] = (ws.Cells(i+2,2).Value, ws.Cells(i+2,3).Value)
    
        return result_dict
        

    def write_mark(self,dict_mark,path):
        wb = self.excel.Workbooks.Add() #?óë??? ?îÑÎ°úÍ∑∏?û®?óê Workbook Ï∂îÍ??(Í∞ùÏ≤¥ ?Ñ§?†ï)
        ws = wb.Worksheets("sheet1")
        count = 2
        for key in dict_mark.keys():
            ws.Cells(count,1).Value = key
            ws.Cells(count,2).Value = dict_mark[key][0]
            ws.Cells(count,3).Value = dict_mark[key][1]
            count = count+1
        ws.Cells(1,1).Value = "total"
        ws.Cells(1,2).Value = count - 2
        print(path)
        wb.SaveAs(path)



class AutoMarkerChanger:
    def __init__(self):
        self.rtmark = ReplaceToMaker()
        self.rwexcel = ReadWriteExecl()
        
    def import_mark(self, path):
        # dict(key: str mark_number, value: tuple(str name, str value)
        return self.rwexcel.read_mark(path)

    def export_mark(self, dict_mark, path):
        # dict(key: str mark_number, value: tuple(str name, str value), str ?†à???Í≤ΩÎ°ú?è¥?çî
        self.rwexcel.write_mark(dict_mark, path)

    def run(self, origin_file_path_list, dict_mark, target_path):
        for path in origin_file_path_list:
            for mark in dict_mark.keys():
                self.rtmark.ReplaceToText(path, mark,dict_mark[mark][0], target_path)






if __name__ == "__main__":
    # rf = ReplaceToMaker()
    # rf.ActiveVisible()
    # patht = r"D:\expert_project\auto_invoice\NEW ?ãú?óò?ó∞Íµ¨Í≥Ñ?öç?Ñú 102?ò∏ (?òà?ãú?ñë?ãù?è¨?ï®,Î™ÖÌåêÎ≥?Í≤?).docx"
    # print(rf.ReplaceToText(patht,"{@#mark@6}","test"))

    # re = ReadWriteExecl()
    # dict = re.read_mark(r"D:\expert_project\auto_invoice\export.xlsx")
    # print(dict)
    # re.write_mark(dict,r"D:\expert_project\auto_invoice\export2.xlsx")

    amc = AutoMarkerChanger()
    mark_dict = amc.import_mark(r"D:\expert_project\auto_invoice\export.xlsx")
    amc.export_mark(mark_dict, r"D:\expert_project\auto_invoice\export2.xlsx")

    file_list = [
        r"D:\expert_project\auto_invoice\NEW ?ãú?óò?ó∞Íµ¨Í≥Ñ?öç?Ñú 102?ò∏ (?òà?ãú?ñë?ãù?è¨?ï®,Î™ÖÌåêÎ≥?Í≤?).docx",
        r"D:\expert_project\auto_invoice\?ö©?èÑ?Ñ§Î™ÖÏÑú ?ñë?ãù(?ûê?û¨,Í≤¨Ìíà).doc",
        r"D:\expert_project\auto_invoice\?†ÑÍ∏∞Ïù∏Ï¶? Î©¥Ï†ú?ôï?ù∏?Ñú-Í±¥Ï°∞Í∏?.xlsx"
    ]
    amc.run(file_list, mark_dict, "")
    # (filename, code)
    '''
    if ~~~~~ :
        error!!!
        print(!!!!)
        return 3
    dict = {3 : !!!!}
    '''

