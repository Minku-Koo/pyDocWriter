# -*- conding: utf-8 -*-

# pip install pyqt5

import sys
from PyQt5 import QtCore
from PyQt5.QtWidgets import *
from PyQt5.QtGui import * 
from PyQt5.QtCore import * 
from PyQt5 import QtCore
import PyQt5
from PyQt5.QtWidgets import QWidget, QDesktopWidget, QApplication
import os
import shutil
from PyQt5.Qt import Qt
import datetime 

class DocWriter(QWidget):
    def __init__(self):
        super().__init__()
        self.HEIGHT = 600
        self.WIDTH = 600 * 2
        self.resize(self.WIDTH, self.HEIGHT)
        self.title = "Doc Writer"
        self.logo_file = "logo.png"
        self.gui_background_color = "background-color:white;"

        self.initGUI()
        return 

    def initGUI(self): # main user interface 
        self.setWindowTitle(self.title) #GUI Title
        self.setWindowIcon(QIcon(self.logo_file)) #set Icon File, 16x16, PNG file
        self.setStyleSheet(self.gui_background_color) #배경색 설정

        uic.loadUi(option_ui, self)
        self.show()
        return 


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = DocWriter()
    sys.exit(app.exec_())
    pass