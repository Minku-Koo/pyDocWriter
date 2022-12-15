import win32com.client as win32
from win32com.client import Dispatch
import os.path
import time
class ReplaceToMaker:
    def __init__(self) :
        self.word = win32.gencache.EnsureDispatch("Word.application")
        self.excel = Dispatch("Excel.Application") #엑셀 프로그램 실행
        self.excel.Visible = True
        self.word.Visible = True
    # def __del__(self):
    #     self.word.Quit()
    #     self.excel.Quit()
    def ActiveVisible(self) :
        self.excel.Visible = True #앞으로 실행과정을 보이게
        self.word.Visible = True
    def DeactiveVisible(self):
        self.excel.Visible = False
        self.word.Visible = False

    def ReplaceToText(self, path, dict_mark, target_path=""):
        if(path[-4:] == "xlsx" or path[-3:] == "xls"):
            ret =  self.__search_replace_all_excel(path,dict_mark,target_path)
            return ret
        elif(path[-4:] == "docx" or path[-3:] == "doc"):
            return self.__search_replace_all_word(path,dict_mark,target_path)
        else:
            return -1

    def __search_replace_all_word(self, path, dict_mark, target_path):
        try:
            doc = self.word.Documents.Open(path)
            maintext = self.word.ActiveDocument.StoryRanges(1)
            for mark in dict_mark.keys():
                find_str = "{@#mark@"+str(mark)+"}"
                replace_str = dict_mark[mark][0]
                a = maintext.Text.count(find_str) # 본문 중 찾을 단어의 수를 센다.

                for i in range(0, a):
                    self.word.Selection.GoTo(What=win32.constants.wdGoToSection, Which=win32.constants.wdGoToFirst)
                    self.word.Selection.Find.Text = find_str # 찾을 단어를 찾는다.
                    self.word.Selection.Find.Replacement.Text = "" # 찾을 단어를 지운다.
                    self.word.Selection.Find.Execute(Replace=1, Forward=True)
                    self.word.Selection.InsertAfter(replace_str) # 해당 위치에 삽입하고자 하는 단어를 입력한다.

            if(target_path == ""):
                doc.Save
            else:
                print(target_path)
                doc.SaveAs(target_path)
            doc.Close()
            return 0
        except Exception as e :
            print(e)
            return -1

    def __search_replace_all_excel(self, path, dict_mark, target_path):
        try:
            workbook = self.excel.Workbooks.Open(path)
            for mark in dict_mark.keys():
                find_str = "{@#mark@"+str(mark)+"}"
                replace_str = dict_mark[mark][0]
                for worksheet in workbook.Worksheets:
                    cell = worksheet.UsedRange.Find(find_str)
                    if cell:
                        rg = worksheet.Range(worksheet.usedRange.Address)
                        rg.Replace(find_str, replace_str)
            
            if(target_path == ""):
                workbook.Save
            else:
                workbook.SaveAs(target_path)
            workbook.Close()
            return 0
        except:
            return -1


class ReadWriteExecl:
    def __init__(self) :
        self.excel = Dispatch("Excel.Application") #엑셀 프로그램 실행
        self.excel.Visible = True

    # def __del__(self) :
    #     self.excel.Quit()
    def read_mark(self,path):
        try:
            workbook = self.excel.Workbooks.Open(path)
            ws = workbook.Worksheets(1)
            count = ws.Cells(1,2).Value
            result_dict = {}
            for i in range(int(count)) :
                result_dict[ws.Cells(i+2,1).Value] = (ws.Cells(i+2,2).Value, ws.Cells(i+2,3).Value)
            workbook.Close()
            return result_dict
        except:
            return None
        

    def write_mark(self,dict_mark,path):
        try:
            wb = self.excel.Workbooks.Add() #엑셀 프로그램에 Workbook 추가(객체 설정)
            ws = wb.Worksheets("sheet1")
            count = 2
            for key in dict_mark.keys():
                if not key == None:
                    ws.Cells(count,1).Value = '\''+str(key)
                if not dict_mark[key][0] == None:
                    ws.Cells(count,2).Value = '\''+str(dict_mark[key][0])
                if not dict_mark[key][1] == None:
                    ws.Cells(count,3).Value = '\''+str(dict_mark[key][1])
                count = count+1
            ws.Cells(1,1).Value = "total"
            ws.Cells(1,2).Value = '\''+str(count - 2)
            wb.SaveAs(path)
            wb.Close()
            return 0
        except:
            return -1


##################################public func###########################################

class AutoMarkerChanger:
        
    def import_mark(self, path):
        rwexcel = ReadWriteExecl()
        path = path.replace('/','\\') 
        # dict(key: str mark_number, value: tuple(str name, str value)
        return rwexcel.read_mark(path) # good : dict / fail : None

    def export_mark(self, dict_mark, path):
        rwexcel = ReadWriteExecl()
        # dict(key: str mark_number, value: tuple(str name, str value), str 절대경로폴더
        path = path.replace('/','\\') 
        if os.path.exists(path):
            datetime = time.strftime('%Y%m%d%H%M%S')
            file_type = path.split('.')[-1]
            path = path[:-len(file_type)-1] + '_' + datetime +'.'+file_type
        return rwexcel.write_mark(dict_mark, path) # good : 0 / fail : -1

    def run(self, origin_file_path_list, dict_mark, target_path):
        rtmark = ReplaceToMaker()
        log = []
        target_path = target_path.replace('/','\\') 

        if target_path[-1] != '\\':
            target_path = target_path+'\\'


        for path in origin_file_path_list:
            path = path.replace('/','\\') 
            filename = path.split('\\')[-1]
            targetfilepath = target_path+filename
            if os.path.exists(targetfilepath):
                datetime = time.strftime('%Y%m%d%H%M%S')
                file_type = targetfilepath.split('.')[-1]
                targetfilepath = targetfilepath[:-len(file_type)-1] + '_' +datetime +'.'+ file_type

            path = path.replace('/','\\')
            ret=rtmark.ReplaceToText(path, dict_mark,targetfilepath)
            log.append((targetfilepath,ret))
        return log


if __name__ == "__main__":
    # rf = ReplaceToMaker()
    # rf.ActiveVisible()
    # patht = r"D:\expert_project\auto_invoice\NEW 시험연구계획서 102호 (예시양식포함,명판변경).docx"
    # print(rf.ReplaceToText(patht,"{@#mark@6}","test"))

    # re = ReadWriteExecl()
    # dict = re.read_mark(r"D:\expert_project\auto_invoice\export.xlsx")
    # print(dict)
    # re.write_mark(dict,r"D:\expert_project\auto_invoice\export2.xlsx")

    amc = AutoMarkerChanger()
    mark_dict = amc.import_mark(r"D:\expert_project\auto_invoice\export.xlsx")
    # print(mark_dict)
    print(amc.export_mark(mark_dict,r"D:\expert_project\auto_invoice\export2.xlsx"))
    target_path = "D:\\expert_project\\auto_invoice\\export"
    file_list = [
        r"D:\expert_project\auto_invoice\NEW 시험연구계획서 102호 (예시양식포함,명판변경).docx",
        r"D:\expert_project\auto_invoice\용도설명서 양식(자재,견품).doc",
        r"D:\expert_project\auto_invoice\전기인증 면제확인서-건조기.xlsx"
    ]
    print(amc.run(file_list,mark_dict,target_path))

