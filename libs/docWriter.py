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
# import datetime 
# from PyQt5 import uic
from methods import msAuto

def resource_path(relative_path):
    try:
        # PyInstaller에 의해 임시폴더에서 실행될 경우 임시폴더로 접근하는 함수
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath("./")
    return os.path.join(base_path, relative_path)

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
        self.parents.add_log(self.show_log[self.num], 'black')
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

        self.version = 1.0
        
        ##################################################
        ############# Size Space #########################
        self.HEIGHT = 700
        self.WIDTH = 600 * 2
        self.mark_num = 0
        self.logo_img_size = 80
        self.mark_input_height = 30
        self.run_btn_height = 60
        self.file_list_height = 20
        self.file_list_icon_width = 26
        self.mark_control_btn_height = 60
        self.our_logo_btn_size = 70
        self.setFixedSize(self.WIDTH, self.HEIGHT)

        ##################################################
        ############# Text Space #########################
        self.title = "MS OFFICE MERGER (MOM)"   # Documents Changer
        self.font_path = "./font/"
        self.img_path = "./img/"
        self.logo_file = "logo.png"
        self.excel_icon_filename = "excel_icon.png"
        self.docs_icon_filename = "docs_icon.png"
        self.reset_btn_name = '초기화'
        self.help_btn_name = '도움말'
        self.doc_load_name = "파일 불러오기"
        self.target_path_btn_name = '저장 폴더 지정'
        self.excel_import_name = '형식 불러오기'
        self.excel_export_name = '형식 내보내기'
        self.run_text = '실   행'
        self.mark_rem_btn_text = '-'
        self.mark_add_btn_text = '+'
        self.init_alert_msg = '''초기화를 진행하시겠습니까?''' 
        self.init_alert_title = '''Reset Confirm'''
        
        ##################################################
        ############# variable Space #####################
        self.ms_loaded_file_index = 0
        self.ms_loaded_file_label = {}  # filename : (rem_btn_label, fileLabel)

        self.mark_obj_dict = {}             # markindex : (mark_line_num, mark_line_name, mark_line_value, rem_btn)

        self.output_target_path = ''

        self.help_link_url = "https://lndhub.samsung.com/lndhub/blog/techBlogDetail/AYULHBmkGmZgAcLt?type=Blog"
        self.our_logo_link_url = 'https://html-color-codes.info/Korean/'
        
        self.scroll_flag = True
        self.init_run_flog = True
        self.start_logo_view_over = False

        ##################################################
        ############# Style ################################
        self.gui_background_color = 'white'#"#FAFAFA"
        self.title_font_name = "Samsung Sharp Sans Bold"
        self.samsung_one_font = "SamsungOne 400"

        self.simple_border_style = 'border-style:solid;border-color:#000000;border-width:1px;'

        self.no_border = 'border-radius: 1px;border-width:0px; '
        
        self.btn_default_style = '''
                                '''
        self.run_btn_style = '''font-size : 35px;
                                font-weight : 900;
                                font-family: Samsung Sharp Sans Bold;
                            '''
        self.sub_btn_style = '''
                            '''

        self.mark_header_style = '''
                            '''
        self.header_style_sheet = '''
                                    border-radius: 1px;
                                    color : #2E2E2E;
                                    background-color: #E6E6E6;
                                    '''
        ##################################################
        ############# Log ################################
        self.mark_naming = 'MOM'
        self.log_comment = """이곳에 작업 로그가 보여집니다."""
        self.mark_value_not_enough = '''모든 칸을 채워주세요'''
        self.no_mark_input = f'''하나 이상의 {self.mark_naming}이 입력되어야 합니다'''
        self.target_not_exist = '''결과 저장 폴더를 지정해주세요'''
        self.import_success = '''엑셀 파일에서 Mark 데이터 불러오기 성공'''
        self.export_success = f'''{self.mark_naming} 데이터 엑셀로 내보내기 성공'''
        self.file_loaded_log = '''파일이 정상적으로 업로드 되었습니다'''
        self.no_file_input = '''1개 이상의 파일을 업로드해주세요.'''
        self.not_enough_mark_value = f'''입력되지 않은 {self.mark_naming}가 있습니다.'''
        self.log_color = {'red':"FF0000", 'blue':'2E2EFE', 'black':'000000'}

        ##################################################
        ############# Call Class #########################
        self.amc = msAuto.AutoMarkerChanger()
        self.excel_export_filename = 'Output_MsOfficeMerger.xlsx'

        self.title_font = QFont(self.title_font_name, 35)
        self.general_font = QFont(self.title_font_name)
        
        self.initGUI()
        # self.__show_our_logo_dynamic()

        return 

    def initGUI(self): # main user interface 
        self.setWindowTitle(self.title + f"  ver {self.version}") #GUI Title
        self.setWindowIcon(QIcon(resource_path(self.img_path + self.logo_file))) #set Icon File, 16x16, PNG file
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
        

        self.file_workspace = QGroupBox(self)
        # self.file_workspace.setStyleSheet(self.simple_border_style)
        self.file_scroll_area = QScrollArea(self)
        self.file_scroll_area.setWidgetResizable(True)
        self.grid.addWidget(self.file_scroll_area, 1, 0, 2, 5)

        self.file_vbox = QVBoxLayout()
        self.file_workspace.setLayout(self.file_vbox)

        self.file_scroll_area.setWidget(self.file_workspace)

        doc_load_btn = QPushButton(self.doc_load_name, self)
        doc_load_btn.setText(self.doc_load_name)
        doc_load_btn.setStyleSheet(self.btn_default_style)
        doc_load_btn.clicked.connect(self.load_file_list)
        self.grid.addWidget(doc_load_btn, 3, 0, 1, 1)

        self.create_file_list_label('', header = True)

        return 

    def guiControlPannel(self):
        self.groupbox = QGroupBox(self)
        self.group_scroll_area = QScrollArea(self)
        self.group_scroll_area.setWidgetResizable(True)
        
        self.grid.addWidget(self.group_scroll_area, 4, 0, 5, 5)
        self.mark_vbox = QVBoxLayout()

        import_btn = QPushButton(self)
        import_btn.setText(self.excel_import_name)
        import_btn.setStyleSheet(self.btn_default_style)
        import_btn.clicked.connect(self.mark_import_excel)
        self.grid.addWidget(import_btn, 9, 0, 1, 1)

        export_btn = QPushButton(self)
        export_btn.setText(self.excel_export_name)
        export_btn.setStyleSheet(self.btn_default_style)
        export_btn.clicked.connect(self.mark_export_excel)
        self.grid.addWidget(export_btn, 9, 1, 1, 1)

        mark_rem_btn = QPushButton(self)
        mark_rem_btn.setText(self.mark_rem_btn_text)
        mark_rem_btn.setStyleSheet(self.btn_default_style)
        mark_rem_btn.setFixedWidth(self.mark_control_btn_height)
        mark_rem_btn.clicked.connect(self.remove_mark)
        self.grid.addWidget(mark_rem_btn, 9, 3, 1, 1)

        mark_plus_btn = QPushButton(self)
        mark_plus_btn.setText(self.mark_add_btn_text)
        mark_plus_btn.setStyleSheet(self.btn_default_style)
        mark_plus_btn.setFixedWidth(self.mark_control_btn_height)
        mark_plus_btn.clicked.connect(self.generate_mark)
        self.grid.addWidget(mark_plus_btn, 9, 4, 1, 1)

        self.target_btn = QPushButton(self)
        self.target_btn.setEnabled(True)
        self.target_btn.setText(self.target_path_btn_name)
        self.target_btn.setStyleSheet(self.btn_default_style)
        self.target_btn.clicked.connect(self.set_output_target_path)
        self.grid.addWidget(self.target_btn, 10, 0, 1, 1, alignment=Qt.AlignTop)

        self.target_path_box = QLabel(self)
        self.target_path_box.setStyleSheet(self.simple_border_style)
        self.target_path_box.setText('')
        self.target_path_box.setFixedHeight(25)
        self.grid.addWidget(self.target_path_box, 10, 1, 1, 4, alignment=Qt.AlignTop)

        # self.version_box = QLabel(self)
        # self.version_box.setText(f"Ver {self.version}")
        # self.grid.addWidget(self.version_box, 11, 1, 1, 1, alignment=Qt.AlignLeft)

        self.groupbox.setLayout(self.mark_vbox)
        self.group_scroll_area.setWidget(self.groupbox)

        self.generate_mark(header = True)
        self.generate_mark()
        return 

    def guiLogView(self):
        help_btn = QPushButton(self)
        help_btn.setText(self.help_btn_name)
        help_btn.setStyleSheet(self.sub_btn_style)
        help_btn.clicked.connect(lambda: self.open_webbrowser(self.help_link_url))
        self.grid.addWidget(help_btn, 0, 8, 1, 1, alignment=Qt.AlignBottom)

        init_btn = QPushButton(self)
        init_btn.setText(self.reset_btn_name)
        init_btn.setStyleSheet(self.sub_btn_style)
        init_btn.clicked.connect(self.__reset)
        self.grid.addWidget(init_btn, 0, 9, 1, 1, alignment=Qt.AlignBottom)

        self.log_view = QTextBrowser(self)
        self.log_view.append(self.log_comment)
        self.grid.addWidget(self.log_view, 1, 7, 8, 3)

        run_btn = QPushButton(self)
        run_btn.setText(self.run_text)
        # run_btn.setFont(QFont(self.title_font_name, 35))
        run_btn.setStyleSheet(self.run_btn_style)
        run_btn.clicked.connect(self.__run)
        run_btn.setFixedHeight(self.run_btn_height)
        self.grid.addWidget(run_btn, 9, 7, 2, 3, alignment=Qt.AlignTop)

        '''
        logo_img = QIcon(self.img_path + self.logo_file)
        logo_img_box = QPushButton()
        logo_img_box.setIcon(logo_img)
        logo_img_box.clicked.connect(lambda: self.open_webbrowser(self.help_link_url)) 
        logo_img_box.setStyleSheet(self.no_border)
        logo_img_box.setIconSize(QSize(self.our_logo_btn_size, self.our_logo_btn_size)) 
        self.grid.addWidget(logo_img_box, 10, 9, 1, 1, alignment=Qt.AlignRight)
        '''

        return 

    # + click event
    def generate_mark(self, header = False):
        if not header:
            self.mark_num += 1
        self.create_mark(header = header)

        self.scroll_flag = True
        self.worker = Worker(self)
        self.worker.start()

        return 

    def create_mark(self, mark_num = 0, mark_name = '', mark_val = '', header = False):
        self.mark_box = QHBoxLayout()

        if header:
            mark_line_num = QLabel(self)
            mark_line_num.setText('번호')
            mark_line_num.setStyleSheet(self.header_style_sheet)
            mark_line_num.setAlignment(Qt.AlignCenter)
            # mark_line_num.setStyleSheet(self.mark_header_style)
        else:
            mark_line_num = QPushButton(self)
            if mark_num == 0:
                mark_line_num.setText(f"{self.mark_naming}{self.mark_num}")
            else:
                mark_line_num.setText(f"{self.mark_naming}{mark_num}")
            
            mark_line_num.setStyleSheet(self.mark_header_style)
        mark_line_num.setFixedHeight(self.mark_input_height)
        mark_line_num.setFixedWidth(100)
        mark_line_num.setEnabled(False)

        mark_line_name = QLineEdit()
        # mark_line_name = QLabel()
        if header:
            mark_line_name = QLabel(self)
            mark_line_name.setText('설명')
            mark_line_name.setAlignment(Qt.AlignCenter)
            mark_line_name.setStyleSheet(self.header_style_sheet)
            # mark_line_name.setFixedWidth(160)
            # mark_line_name.setFixedHeight(self.mark_input_height)
        else:
            mark_line_name = QLineEdit(self)
            if mark_name:
                mark_line_name.setText(mark_name)
        mark_line_name.setFixedWidth(250)
        mark_line_name.setFixedHeight(self.mark_input_height)
        

        if header:
            mark_line_value = QLabel(self)
            mark_line_value.setText('데이터')
            mark_line_value.setAlignment(Qt.AlignCenter)
            mark_line_value.setStyleSheet(self.header_style_sheet)
        else:
            mark_line_value = QPlainTextEdit(self)
            # mark_line_value = QLineEdit(self)
            if mark_val:
                mark_line_value.setPlainText(mark_val)

        mark_line_value.setFixedWidth(440)
        mark_line_value.setFixedHeight(self.mark_input_height)

        self.mark_box.addWidget(mark_line_num)
        self.mark_box.addWidget(mark_line_value)
        self.mark_box.addWidget(mark_line_name)

        self.mark_vbox.addLayout(self.mark_box)
        self.mark_box.setAlignment(Qt.AlignTop)
        if header:
            # self.mark_box.setAlignment(Qt.AlignTop)
            self.mark_vbox.setAlignment(Qt.AlignTop)
        else:
            # self.mark_box.setAlignment(Qt.AlignTop)
            self.mark_vbox.setAlignment(Qt.AlignTop)
            # self.groupbox.setAlignment(Qt.AlignTop)
            pass

        if not header:
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
        return 

    def load_file_list(self):
        flist = QFileDialog.getOpenFileNames(self, 'Open file', './', 'ms file(*.xlsx *.xls *.docx *.doc)')
        
        for filename in flist[0]:
            if not self.ms_loaded_file_label.get(filename):
                self.create_file_list_label(filename)

        if flist[0]:
            self.add_log(self.file_loaded_log, 'black')
        return 

    def create_file_list_label(self, filename, header = False):
        file_group_box = QHBoxLayout()
        if not header:
            if "xls" in filename.split(".")[-1]:
                # file_icon_img = QPixmap(self.img_path + self.excel_icon_filename).scaled(self.file_list_icon_width, self.file_list_height)
                file_icon_img = QPixmap(resource_path(self.img_path + self.excel_icon_filename)).scaled(self.file_list_icon_width, self.file_list_height)
                
            elif "doc" in filename.split(".")[-1]:
                file_icon_img = QPixmap(resource_path(self.img_path + self.docs_icon_filename)).scaled(self.file_list_icon_width, self.file_list_height)
        
        file_icon_img_box = QLabel()
        
        if header:
            file_icon_img_box = QLabel()
            file_icon_img_box.setStyleSheet(self.header_style_sheet)
            file_icon_img_box.setText("종류")
            file_icon_img_box.setFixedWidth(self.file_list_icon_width)
        else:
            file_icon_img_box.setFixedWidth(self.file_list_icon_width)
            file_icon_img_box.setFixedHeight(self.file_list_height)
            file_icon_img_box.setStyleSheet(self.no_border)
            file_icon_img_box.setPixmap(file_icon_img)
        

        rem_btn = QPushButton(self)
        if header:
            rem_btn.setText("삭제")
            rem_btn.setStyleSheet(self.header_style_sheet)
        else:
            rem_btn.setText('X')
            rem_btn.clicked.connect(lambda: self.remove_file_list(filename))
        rem_btn.setFixedWidth(34)
        rem_btn.setFixedHeight(20)

        file_label = QLabel(self)
        if header:
            file_label.setText("파일명")
            file_label.setAlignment(Qt.AlignCenter)
            file_label.setStyleSheet(self.header_style_sheet)
        else:
            file_label.setText(filename)
            file_label.setStyleSheet(self.no_border)

        file_label.setFixedWidth(720)
        
        file_group_box.addWidget(rem_btn)
        file_group_box.addWidget(file_icon_img_box)
        file_group_box.addWidget(file_label)

        self.file_vbox.addLayout(file_group_box)
        self.file_vbox.setAlignment(Qt.AlignTop)
        if not header:
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

    def add_log(self, text, color = 'black'):
        if self.start_logo_view_over:
            self.log_view.clear()
        #     self.start_logo_view_over = False

        if color != 'black':
            # if self.init_run_flog and not self.start_logo_view_over:
            #     self.log_view.clear()
            #     self.start_logo_view_over = False
            #     self.init_run_flog = False
            pass

        self.log_view.append(f'<span style=\"color:#{self.log_color[color]};\">{text}</span>')
        return 


    def mark_export_excel(self):
        export_path, _ = QFileDialog.getSaveFileName(self, 'Select Directory', 
                                                f'./{self.excel_export_filename}', 
                                                'Excel (*.xlsx)')
        if not export_path:
            return 
        # export_path += self.excel_export_filename
        export_path = export_path.replace("/", "\\")
        export_dict = {}
        if self.mark_num == 0:
            self.add_log(self.no_mark_input, 'red')
            return 
        for mark_number in range(1, self.mark_num + 1):
            print(self.mark_obj_dict[mark_number])
            lb_name_text = self.mark_obj_dict[mark_number][1].text()
            lb_value_text = self.mark_obj_dict[mark_number][2].toPlainText()
            export_dict[mark_number] = (lb_name_text, lb_value_text)
        self.amc.export_mark(export_dict, export_path)
        self.add_log(self.export_success, 'black')
        return 

    def mark_import_excel(self):
        imported_file = QFileDialog.getOpenFileNames(self, 'Select Directory', './')[0]
        if not imported_file:
            return 
        self.mark_obj_dict  = {}
        imported_file = imported_file[0].replace("/", "\\")
        mark_dict = self.amc.import_mark(imported_file)
        print(mark_dict)
        for w in self.groupbox.findChildren(QPushButton):
            w.deleteLater()
        for w in self.groupbox.findChildren(QLineEdit):
            w.deleteLater()
        for w in self.groupbox.findChildren(QPlainTextEdit):
            w.deleteLater()

        self.mark_num = 0
        for mark_num in mark_dict:
            self.create_mark(mark_num= int(mark_num), mark_val= str(mark_dict[mark_num][1]), mark_name= str(mark_dict[mark_num][0]))
        self.mark_num = int(mark_num)
        self.add_log(self.import_success, 'black')
        return 

    def __reset(self):
        buttonReply = QMessageBox.information(
                                            self, self.init_alert_title, self.init_alert_msg, 
                                            QMessageBox.Yes | QMessageBox.Cancel
                                            )
        if buttonReply == QMessageBox.Cancel:
            return 
        self.log_view.clear()
        self.add_log(self.log_comment, 'black')

        for w in self.groupbox.findChildren(QPushButton):
            w.deleteLater()
        for w in self.groupbox.findChildren(QLineEdit):
            w.deleteLater()
        for w in self.groupbox.findChildren(QPlainTextEdit):
            w.deleteLater()

        # for w in self.file_workspace.findChildren(QLabel):
        #     w.deleteLater()
        # for w in self.file_workspace.findChildren(QPushButton):
        #     w.deleteLater()
        
        for fname in self.ms_loaded_file_label.keys():
            for k in self.ms_loaded_file_label[fname]:
                self.file_vbox.removeWidget(k)

        self.mark_num = 0
        self.ms_loaded_file_label = {}
        self.ms_loaded_file_index = 0
        self.mark_obj_dict = {} 
        self.output_target_path = ''
        self.target_path_box.setText('')
        return 

    def file_work_result(self, filename, done = False):
        if not done:
            return f'''{filename} -> 작업 실패'''
        return f'''{filename} -> 작업 성공'''

    def __run(self):
        print(self.ms_loaded_file_label.keys())
        if not self.ms_loaded_file_label.keys():
            self.add_log(self.no_file_input, 'red')
            return 
        if self.mark_num == 0:
            self.add_log(self.no_mark_input, 'red')
            return 
        if not self.output_target_path:
            self.add_log(self.target_not_exist, 'red')
            return 

        mark_dict = {}
        for mark_number in range(1, self.mark_num + 1):
            print(self.mark_obj_dict[mark_number][1].text(), self.mark_obj_dict[mark_number][2].toPlainText())
            if not self.mark_obj_dict[mark_number][1].text() or not self.mark_obj_dict[mark_number][2].toPlainText():
                self.add_log(self.not_enough_mark_value, 'red')
                return 
            lb_name = self.mark_obj_dict[mark_number][1].text()
            lb_value = self.mark_obj_dict[mark_number][2].toPlainText()
            mark_dict[mark_number] = (lb_value, lb_name)
            print(mark_number)
            # print(lb_name.text(), lb_value.toPlainText())
        
        ret = self.amc.run(list(self.ms_loaded_file_label.keys()), mark_dict, self.output_target_path)
        
        for log_info in ret:
            if log_info[1] == 0: # success
                self.add_log(self.file_work_result(log_info[0], True), 'blue')
            else:
                self.add_log(self.file_work_result(log_info[0], False), 'red')
        return 

    def __show_our_logo_dynamic(self):
        self.init_run_flog = True
        show_we = ShowWe(self)
        show_we.start()
        return 

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = DocWriter()
    sys.exit(app.exec_())

