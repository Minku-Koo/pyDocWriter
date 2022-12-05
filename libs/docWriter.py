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
# import shutil
# from PyQt5.Qt import Qt
import datetime 

# addWidget (self, QWidget, row, column, rowSpan, columnSpan, Qt.Alignment alignment = 0)


class DocWriter(QWidget):
    def __init__(self):
        super().__init__()
        
        self.HEIGHT = 600
        self.WIDTH = 600 * 2
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
        return 

    def guiLogView(self):
        init_btn = QPushButton(self)
        init_btn.setText(self.reset_btn_name)
        self.grid.addWidget(init_btn, 0, 10, 1, 1)

        log_view = QTextBrowser()
        log_view.append(self.log_init_comment)
        # self.tb.setAcceptRichText(True)
        # self.tb.setOpenExternalLinks(True)
        self.grid.addWidget(log_view, 1, 6, 8, 5)

        logo_img = QPixmap(self.logo_file).scaled(self.logo_img_size, self.logo_img_size)
        logo_img_box = QLabel()
        logo_img_box.setPixmap(logo_img)
        self.grid.addWidget(logo_img_box, 9, 10, 1, 1)

        return 


if __name__ == '__main__':
    app = QApplication(sys.argv)
    # screen = app.primaryScreen()
    ex = DocWriter()
    sys.exit(app.exec_())
    pass