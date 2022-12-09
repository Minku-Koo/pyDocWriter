# -*- coding: utf-8 -*-
# pip install pyqt5

import sys
from PyQt5.QtWidgets import *
from PyQt5.QtGui import * 
from PyQt5.QtCore import * 
from PyQt5.QtGui import QIcon, QPixmap
from PyQt5 import QtCore
from PyQt5.QtWidgets import QWidget, QDesktopWidget, QApplication, QGridLayout, QLabel, QLineEdit, QTextEdit
import os
import time
import webbrowser
import datetime 
from PyQt5 import uic
from utils import msAuto

class Worker(QThread):
    def __init__(self, parents):
        super().__init__(parents)
        self.parents = parents
        self.__running = True

    def run(self):
        while self.__running:
            time.sleep(0.05)
            if self.parents.scroll_flag:
                self.__scroll()
            else:
                self.stop()

    def __scroll(self):
        self.parents.group_scroll_area.verticalScrollBar().setValue(
            self.parents.group_scroll_area.verticalScrollBar().maximum()
        )
        self.parents.scroll_flag = False
        return 
    
    def stop(self):
        self.__running = False
        self.quit()
        return 


class ShowWe(QThread):
    def __init__(self, parents):
        super().__init__(parents)
        self.parents = parents
        self.__running = True
        self.num = 0
        self.our_init_logo_view = '''
■■■■■■    ■■■■■■    ■■■■■■    ■■■■■■
■■■■■■    ■■■■■■    ■■■■■■    ■■■■■■
■          ■    ■                 ■                ■          ■
■          ■    ■                 ■                ■          ■
■          ■    ■                 ■                ■          ■
■          ■    ■                 ■                ■          ■
■          ■    ■                 ■                ■          ■
■          ■    ■                 ■                ■          ■
■          ■    ■                 ■                ■          ■
■          ■    ■                 ■                ■          ■
■          ■    ■                 ■                ■          ■
■■■■■■    ■■■■■■    ■■■■■■    ■          ■
■■■■■■    ■■■■■■    ■■■■■■    ■          ■
          ■■              ■■              ■■    ■          ■
          ■■              ■■              ■■    ■          ■
          ■■              ■■              ■■    ■          ■
          ■■              ■■              ■■    ■          ■
          ■■              ■■              ■■    ■          ■
          ■■              ■■              ■■    ■          ■
          ■■              ■■              ■■    ■          ■
■■■■■■    ■■■■■■    ■■■■■■    ■■■■■■ 
■■■■■■    ■■■■■■    ■■■■■■    ■■■■■■ 
'''
        self.show_log = self.our_init_logo_view.split("\n")
        
    def run(self):
        while self.__running:
            if self.parents.init_run_flog:
                time.sleep(0.08)
                self.__show()
            else:
                self.stop()

    def __show(self):
        self.parents.add_log(self.show_log[self.num])
        self.num += 1
        if len(self.show_log) == self.num:
            self.parents.init_run_flog = False
        return 

    def stop(self):
        self.__running = False
        self.parents.start_logo_view_over = True
        self.quit()
        return 



class DocWriter(QWidget):
    def __init__(self):
        super().__init__()
        
        self.HEIGHT = 700
        self.WIDTH = 600 * 2
        self.mark_num = 0
        self.logo_img_size = 80
        self.mark_input_height = 30
        self.run_btn_height = 60
        self.file_list_height = 17
        self.file_list_icon_width = 20
        self.resize(self.WIDTH, self.HEIGHT)

        self.title = "MS OFFICE MERGER (MOM)"
        self.font_path = "./font/"
        self.title_font_name = "Samsung Sharp Sans Bold"
        self.samsung_one_font = "SamsungOne 400"
        self.img_path = "./img/"
        self.logo_file = "logo.png"
        self.excel_icon_filename = "excel_icon.png"
        self.docs_icon_filename = "docs_icon.png"
        self.gui_background_color = "white"
        self.reset_btn_name = 'RESET'
        self.help_btn_name = 'HELP'
        self.doc_load_name = "File Load"
        self.target_path_btn_name = '저장 폴더 지정'
        self.excel_import_name = 'Import'
        self.excel_export_name = 'Export'

        self.run_done = '''Work Done!'''
        self.run_text = '실 행'
        self.mark_rem_btn_text = '삭 제'
        self.mark_add_btn_text = '추 가'
        
        # self.ms_loaded_file_list = []
        self.ms_loaded_file_index = 0
        self.ms_loaded_file_label = {}  # filename : (rem_btn_label, fileLabel)

        self.mark_obj_dict = {}             # markindex : (mark_line_num, mark_line_name, mark_line_value, rem_btn)
        # self.mark_value_dict = {}           # markindex : 

        self.output_target_path = './'

        self.help_link_url = "https://www.naver.com/"
        self.our_logo_link_url = 'https://html-color-codes.info/Korean/'

        self.target_active_style = 'border-style:solid;border-color:#000000;border-width:1px;'
        self.target_unactive_style = 'border-style:solid;border-color:#c7c7c7;border-width:1px;'
        self.no_border = 'border-width:0px;'

        self.scroll_flag = True
        self.init_run_flog = True
        self.start_logo_view_over = False

        self.amc = msAuto.AutoMarkerChanger()
        self.excel_export_filename = '/Output_MsOfficeMerger.xlsx'

        self.title_font = QFont(self.title_font_name, 35)
        self.general_font = QFont(self.samsung_one_font)
        
        # self.title_font_name

        ##################################################
        ############# Log ################################
        self.log_comment = """Log View"""
        self.file_loaded_log = '''파일 업로드 완료'''
        self.mark_value_not_enough = '''모든 칸을 채워주세요'''
        self.no_mark_input = '''하나 이상의 mark가 입력되어야 합니다'''
        self.target_not_exist = '''결과 저장 폴더를 지정해주세요'''
        self.import_success = '''엑셀 파일에서 데이터 불러오기 성공'''
        self.export_success = '''데이터 엑셀로 내보내기 성공'''

        


        self.initGUI()

        self.__wonderful()
        return 

    def initGUI(self): # main user interface 
        self.setWindowTitle(self.title) #GUI Title
        self.setWindowIcon(QIcon(self.img_path + self.logo_file)) #set Icon File, 16x16, PNG file
        self.setStyleSheet(f"background-color:{self.gui_background_color};") 

        self.grid = QGridLayout()
        self.setLayout(self.grid)

        self.guiHeaderFileBox()
        self.guiControlPannel()
        self.guiLogView()

        self.show() # show GUI
        return 

    def guiHeaderFileBox(self):
        title = QLabel(self.title, self)
        title.setFont(self.title_font)
        title.setAlignment(Qt.AlignCenter)
        # title.setStyleSheet('border-style:solid;border-color:black;border-width:1px;')
        self.grid.addWidget(title, 0, 2, 1, 3)
        
        doc_load_btn = QPushButton(self.doc_load_name, self)
        doc_load_btn.setText(self.doc_load_name)
        doc_load_btn.clicked.connect(self.load_file_list)
        self.grid.addWidget(doc_load_btn, 1, 0, 1, 1)


        self.file_workspace = QGroupBox(self)
        self.file_workspace.setStyleSheet('border-style:solid;border-color:black;border-width:1px;')
        

        self.file_scroll_area = QScrollArea(self)
        self.file_scroll_area.setWidgetResizable(True)
        self.grid.addWidget(self.file_scroll_area, 2, 0, 2, 5)

        self.file_vbox = QVBoxLayout()

        self.file_workspace.setLayout(self.file_vbox)
        self.file_scroll_area.setWidget(self.file_workspace)

        # self.select_box = QComboBox(self)
        # self.select_box.addItem(self.select_options_name[0])
        # self.select_box.addItem(self.select_options_name[1])
        # self.select_box.currentIndexChanged.connect(self.change_combo)
        
        # self.grid.addWidget(self.select_box, 4, 0, 1, 1)

        return 

    def guiControlPannel(self):
        self.groupbox = QGroupBox(self)
        self.group_scroll_area = QScrollArea(self)
        self.group_scroll_area.setWidgetResizable(True)
        
        self.grid.addWidget(self.group_scroll_area, 4, 0, 5, 5)
        self.mark_vbox = QVBoxLayout()

        import_btn = QPushButton(self)
        import_btn.setText(self.excel_import_name)
        import_btn.clicked.connect(self.mark_import_excel)
        self.grid.addWidget(import_btn, 9, 0, 1, 1)

        export_btn = QPushButton(self)
        export_btn.setText(self.excel_export_name)
        export_btn.clicked.connect(self.mark_export_excel)
        self.grid.addWidget(export_btn, 9, 1, 1, 1)

        mark_rem_btn = QPushButton(self)
        mark_rem_btn.setText(self.mark_rem_btn_text)
        mark_rem_btn.clicked.connect(self.remove_mark)
        self.grid.addWidget(mark_rem_btn, 9, 3, 1, 1)

        mark_plus_btn = QPushButton(self)
        mark_plus_btn.setText(self.mark_add_btn_text)
        mark_plus_btn.clicked.connect(self.generate_mark)
        self.grid.addWidget(mark_plus_btn, 9, 4, 1, 1)

        self.target_btn = QPushButton(self)
        self.target_btn.setEnabled(True)
        self.target_btn.setText(self.target_path_btn_name)
        self.target_btn.clicked.connect(self.set_output_target_path)
        self.grid.addWidget(self.target_btn, 10, 0, 1, 1)

        self.target_path_box = QLabel(self)
        self.target_path_box.setStyleSheet(self.target_active_style)
        self.target_path_box.setText('')
        self.target_path_box.setFixedHeight(25)
        self.grid.addWidget(self.target_path_box, 10, 1, 1, 4)

        
        self.groupbox.setLayout(self.mark_vbox)
        self.group_scroll_area.setWidget(self.groupbox)

        # self.generate_mark()
        return 

    def guiLogView(self):
        help_btn = QPushButton(self)
        help_btn.setText(self.help_btn_name)
        help_btn.clicked.connect(lambda: self.open_webbrowser(self.help_link_url))
        self.grid.addWidget(help_btn, 1, 8, 1, 1)

        init_btn = QPushButton(self)
        init_btn.setText(self.reset_btn_name)
        init_btn.clicked.connect(self.__reset)
        self.grid.addWidget(init_btn, 1, 9, 1, 1)

        self.log_view = QTextBrowser(self)
        self.grid.addWidget(self.log_view, 2, 7, 7, 3)

        run_btn = QPushButton(self)
        run_btn.setText(self.run_text)
        run_btn.clicked.connect(self.__run)
        run_btn.setFixedHeight(self.run_btn_height)
        self.grid.addWidget(run_btn, 9, 7, 2, 2)

        logo_img = QPixmap(self.img_path + self.logo_file).scaled(self.logo_img_size, self.logo_img_size)
        logo_img_box = QLabel()
        logo_img_box.setPixmap(logo_img)
        self.grid.addWidget(logo_img_box, 10, 9, 1, 1, alignment=Qt.AlignRight)


        return 

    # + click event
    def generate_mark(self):
        self.mark_num += 1
        self.create_mark()

        self.scroll_flag = True
        self.worker = Worker(self)
        self.worker.start()

        return 

    def create_mark(self, mark_num = 0, mark_name = '', mark_val = ''):
        self.mark_box = QHBoxLayout()

        mark_line_num = QPushButton(self)
        if mark_num == 0:
            mark_line_num.setText(f"mark{self.mark_num}")
        else:
            mark_line_num.setText(f"mark{mark_num}")
        mark_line_num.setEnabled(False)
        mark_line_num.setFixedHeight(self.mark_input_height)
        mark_line_num.setFixedWidth(100)

        mark_line_name = QLineEdit(self)
        if mark_name:
            mark_line_name.setText(mark_name)
        mark_line_name.setFixedWidth(160)
        mark_line_name.setFixedHeight(self.mark_input_height)

        mark_line_value = QPlainTextEdit(self)
        # mark_line_value = QLineEdit(self)
        mark_line_value.setFixedHeight(self.mark_input_height)
        if mark_val:
            mark_line_value.setPlainText(mark_val)

        self.mark_box.addWidget(mark_line_num)
        self.mark_box.addWidget(mark_line_name)
        self.mark_box.addWidget(mark_line_value)

        # self.mark_box.setAlignment(Qt.AlignTop)
        self.mark_vbox.addLayout(self.mark_box)
        # self.mark_vbox.setAlignment(Qt.AlignTop)
        value_tp = (mark_line_num, mark_line_name, mark_line_value)
        if mark_num == 0:
            self.mark_obj_dict[self.mark_num] = value_tp
        else:
            self.mark_obj_dict[mark_num] = value_tp
        return 

    def remove_mark(self):
        if self.mark_num == 0:
            return 
        
        self.mark_vbox.removeWidget(self.mark_obj_dict[self.mark_num][0])
        self.mark_vbox.removeWidget(self.mark_obj_dict[self.mark_num][1])
        self.mark_vbox.removeWidget(self.mark_obj_dict[self.mark_num][2])
        
        del  self.mark_obj_dict[self.mark_num]
        self.mark_num -= 1
        self.mark_vbox.setAlignment(Qt.AlignTop)
        return 

    def load_file_list(self):
        flist = QFileDialog.getOpenFileNames(self, 'Open file', './', 'ms file(*.xlsx *.xls *.docx)')
        for filename in flist[0]:
            if not self.ms_loaded_file_label.get(filename):
                self.create_file_list_label(filename)
        self.add_log(self.file_loaded_log)
        return 

    def create_file_list_label(self, filename):
        file_group_box = QHBoxLayout()
        
        if "xls" in filename.split(".")[-1]:
            file_icon_img = QPixmap(self.img_path + self.excel_icon_filename).scaled(self.file_list_icon_width, self.file_list_height)
        elif "doc" in filename.split(".")[-1]:
            file_icon_img = QPixmap(self.img_path + self.docs_icon_filename).scaled(self.file_list_icon_width, self.file_list_height)
        file_icon_img_box = QLabel()
        file_icon_img_box.setFixedWidth(self.file_list_icon_width)
        file_icon_img_box.setFixedHeight(self.file_list_height)
        file_icon_img_box.setStyleSheet(self.no_border)
        file_icon_img_box.setPixmap(file_icon_img)

        rem_btn = QPushButton(self)
        rem_btn.setText('X')
        rem_btn.clicked.connect(lambda: self.remove_file_list(filename))
        rem_btn.setFixedWidth(20)
        rem_btn.setFixedHeight(20)

        file_label = QLabel(self)
        file_label.setText(filename)
        file_label.setStyleSheet(self.no_border)
        
        file_group_box.addWidget(rem_btn)
        file_group_box.addWidget(file_icon_img_box)
        file_group_box.addWidget(file_label)

        self.file_vbox.addLayout(file_group_box)
        self.file_vbox.setAlignment(Qt.AlignTop)

        self.ms_loaded_file_label[filename] = (rem_btn, file_icon_img_box, file_label)
        return 

    def remove_file_list(self, filename):
        self.file_vbox.removeWidget(self.ms_loaded_file_label[filename][0])
        self.file_vbox.removeWidget(self.ms_loaded_file_label[filename][1])
        self.file_vbox.removeWidget(self.ms_loaded_file_label[filename][2])

        del self.ms_loaded_file_label[filename]
        
        return 

    def set_output_target_path(self):
        self.output_target_path = QFileDialog.getExistingDirectory(self, 'Select Directory', './')
        self.target_path_box.setText(self.output_target_path)
        return 

    def open_webbrowser(self, url):
        webbrowser.open(url)
        return 

    def add_log(self, text):
        if self.start_logo_view_over:
            self.log_view.clear()
            self.start_logo_view_over = False
        self.log_view.append(text)
        return 


    def mark_export_excel(self):
        export_path = QFileDialog.getExistingDirectory(self, 'Select Directory', './')
        if not export_path:
            return 
        export_path += self.excel_export_filename
        export_path = export_path.replace("/", "\\")
        export_dict = {}
        if self.mark_num == 0:
            self.add_log(self.no_mark_input)
            return 
        for mark_number in range(1, self.mark_num + 1):
            print(self.mark_obj_dict[mark_number])
            lb_name_text = self.mark_obj_dict[mark_number][1].text()
            lb_value_text = self.mark_obj_dict[mark_number][2].toPlainText()
            export_dict[mark_number] = (lb_name_text, lb_value_text)
        self.amc.export_mark(export_dict, export_path)
        self.add_log(self.export_success)
        return 

    def mark_import_excel(self):
        imported_file = QFileDialog.getOpenFileNames(self, 'Select Directory', './')[0]
        if not imported_file:
            return 
        imported_file = imported_file[0].replace("/", "\\")
        mark_dict = self.amc.import_mark(imported_file)
        print(mark_dict)

        for w in self.file_workspace.findChildren(QLabel):
            w.deleteLater()
        for w in self.file_workspace.findChildren(QPushButton):
            w.deleteLater()

        self.mark_num = 0
        for mark_num in mark_dict:
            self.create_mark(int(mark_num), mark_dict[mark_num][0], mark_dict[mark_num][1])
        self.mark_num = int(mark_num)
        self.add_log(self.import_success)
        return 

    def __reset(self):
        self.log_view.clear()
        
        for w in self.groupbox.findChildren(QPushButton):
            w.deleteLater()
        for w in self.groupbox.findChildren(QLineEdit):
            w.deleteLater()
        for w in self.groupbox.findChildren(QPlainTextEdit):
            w.deleteLater()
        for w in self.file_workspace.findChildren(QLabel):
            w.deleteLater()
        for w in self.file_workspace.findChildren(QPushButton):
            w.deleteLater()

        self.mark_num = 0
        self.ms_loaded_file_label = {}
        self.ms_loaded_file_index = 0
        self.mark_obj_dict = {} 
        self.output_target_path = './'
        self.target_path_box.setText('')
        return 

    def __run(self):
        if self.mark_num == 0:
            return 
        for mark_number in range(1, self.mark_num + 1):
            lb_name = self.mark_obj_dict[mark_number][1]
            lb_value = self.mark_obj_dict[mark_number][2]
            # print(lb_name.text(), lb_value.text())
            print(lb_name.text(), lb_value.toPlainText())
        self.add_log(self.run_done)
        return 

    def __wonderful(self):
        self.init_run_flog = True
        show_we = ShowWe(self)
        show_we.start()
        return 

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = DocWriter()
    sys.exit(app.exec_())

'''
### 에러 사양
- import 시, 형식이 갖추어지지 않은 엑셀일 경우 (중요!!)
    - 1열 숫자 오름차순 맞는지

### 수정 사항
- import/export 파일 열지 않도록
- class 호출하는 순간, 켜져있던 ms 문서들 다 종료됨 ;;

'''