# -*- conding: utf-8 -*-
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QLabel, QVBoxLayout, QWidget
from PyQt5.QtWidgets import * 
from PyQt5.QtGui import * 
from PyQt5.QtCore import * 
import sys
from libs import docWriter

class BaseWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.HEIGHT = 300
        self.WIDTH = 300
        self.resize(self.WIDTH, self.HEIGHT)
        self.w = None  # No external window yet.

        self.button = QPushButton(self)
        self.button.setText("func1")
        self.button.clicked.connect(self.show_new_window)
        

    def show_new_window(self, checked):
        if self.w is None:
            self.w = docWriter.DocWriter()
        self.w.show()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = BaseWindow()
    w.show()
    app.exec()