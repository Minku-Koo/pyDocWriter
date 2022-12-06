# -*- conding: utf-8 -*-

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
# import shutil
# from PyQt5.Qt import Qt
import datetime 
from PyQt5 import uic


class Worker(QThread):
    def __init__(self, parents):
        super().__init__(parents)
        self.parents = parents
        self.running = True

    def run(self):
        while self.running:
            time.sleep(0.05)
            if self.parents.scroll_flag:
                self.scroll()
            else:
                self.stop()

    def scroll(self):
        self.parents.group_scroll_area.verticalScrollBar().setValue(
            self.parents.group_scroll_area.verticalScrollBar().maximum()
        )
        self.parents.scroll_flag = False
        return 

    def stop(self):
        self.running = False
        self.quit()
        return 


class DocWriter(QWidget):
    def __init__(self):
        super().__init__()
        
        self.HEIGHT = 700
        self.WIDTH = 600 * 2
        self.mark_num = 0
        self.logo_img_size = 80
        self.resize(self.WIDTH, self.HEIGHT)

        self.title = "Doc Writer"
        self.description = """Hi, This is Document Writer what you want"""
        self.logo_file = "logo.png"
        self.gui_background_color = "white"
        self.reset_btn_name = 'RESET'
        self.doc_load_name = "File Load"
        self.select_options_name = ('option1', 'option2')
        self.target_path_btn_name = 'Target'
        self.excel_import_name = 'Import'
        self.excel_export_name = 'Export'
        self.log_init_comment = """This is Log View"""
        # self.makr_name = 'mark' 
        self.run_text = 'RUN'

        self.scroll_flag = True

        self.title_font = QFont()
        self.title_font.setPointSize(30)

        self.initGUI()
        return 

    

    def initGUI(self): # main user interface 
        self.setWindowTitle(self.title) #GUI Title
        self.setWindowIcon(QIcon(self.logo_file)) #set Icon File, 16x16, PNG file
        self.setStyleSheet(f"background-color:{self.gui_background_color};") #배경색 설정

        self.grid = QGridLayout()
        self.setLayout(self.grid)

        self.guiHeader()
        self.guiControlPannel()
        self.guiLogView()

        self.show() # show GUI
        return 

    def guiHeader(self):
        title = QLabel(self.title, self)
        title.setFont(self.title_font)
        title.setStyleSheet('border-style:solid;border-color:black;border-width:1px;')
        self.grid.addWidget(title, 0, 2, 1, 2)
        
        simple_desc = QLabel(self.description, self)
        simple_desc.setStyleSheet('border-style:solid;border-color:black;border-width:1px;')
        self.grid.addWidget(simple_desc, 1, 0, 1, 5)

        doc_load_btn = QPushButton(self)
        doc_load_btn.setText(self.doc_load_name)
        self.grid.addWidget(doc_load_btn, 2, 0, 1, 1)

        select_box = QComboBox(self)
        select_box.addItem(self.select_options_name[0])
        select_box.addItem(self.select_options_name[1])
        self.grid.addWidget(select_box, 2, 3, 1, 1)

        target_btn = QPushButton(self)
        target_btn.setText(self.target_path_btn_name)
        self.grid.addWidget(target_btn, 2, 4, 1, 1)

        return 

    def guiControlPannel(self):
        self.groupbox = QGroupBox(self)
        self.group_scroll_area = QScrollArea(self)
        self.group_scroll_area.setWidgetResizable(True)
        
        self.grid.addWidget(self.group_scroll_area, 3, 0, 5, 5)
        self.mark_vbox = QVBoxLayout()

        import_btn = QPushButton(self)
        import_btn.setText(self.excel_import_name)
        self.grid.addWidget(import_btn, 8, 0, 1, 1)

        export_btn = QPushButton(self)
        export_btn.setText(self.excel_export_name)
        self.grid.addWidget(export_btn, 8, 1, 1, 1)

        mark_plus_btn = QPushButton(self)
        mark_plus_btn.setText('+')
        mark_plus_btn.clicked.connect(self.generate_mark)
        self.grid.addWidget(mark_plus_btn, 8, 4, 1, 1)

        run_btn = QPushButton(self)
        run_btn.setText(self.run_text)
        self.grid.addWidget(run_btn, 9, 1, 1, 3)

        self.groupbox.setLayout(self.mark_vbox)
        self.group_scroll_area.setWidget(self.groupbox)

        self.generate_mark()
        return 

    def guiLogView(self):
        init_btn = QPushButton(self)
        init_btn.setText(self.reset_btn_name)
        self.grid.addWidget(init_btn, 0, 9, 1, 1)

        log_view = QTextBrowser()
        log_view.append(self.log_init_comment)
        # self.tb.setAcceptRichText(True)
        # self.tb.setOpenExternalLinks(True)
        self.grid.addWidget(log_view, 1, 6, 7, 4)

        logo_img = QPixmap(self.logo_file).scaled(self.logo_img_size, self.logo_img_size)
        logo_img_box = QLabel()
        logo_img_box.setPixmap(logo_img)
        self.grid.addWidget(logo_img_box, 9, 9, 1, 1)

        return 

    # + click event
    def generate_mark(self):
        self.create_mark()

        self.scroll_flag = True
        self.worker = Worker(self)
        self.worker.start()

        # self.set_scroll_to_down()
        
        return 


    def set_scroll_to_down(self):
        self.group_scroll_area.verticalScrollBar().setValue(
            self.group_scroll_area.verticalScrollBar().maximum()
        )
        return 

    def create_mark(self):
        self.mark_num += 1
        mark_box = QHBoxLayout()

        mark_line_num = QLineEdit(self)
        mark_line_num.setText(f"{self.mark_num}")
        mark_line_name = QLineEdit(self)
        mark_line_value = QTextEdit(self)
        mark_line_value.setFixedWidth(380)
        mark_line_value.setFixedHeight(40)
        rem_btn = QPushButton(self)
        rem_btn.setText('X')
        rem_btn.clicked.connect(self.remove_mark)
        rem_btn.setFixedWidth(30)
        rem_btn.setFixedHeight(30)

        mark_box.addWidget(mark_line_num)
        mark_box.addWidget(mark_line_name)
        mark_box.addWidget(mark_line_value)
        mark_box.addWidget(rem_btn)

        self.mark_vbox.addLayout(mark_box)
        self.mark_vbox.setAlignment(Qt.AlignTop)
        
        return 

    def remove_mark(self):
        return 

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = DocWriter()
    sys.exit(app.exec_())