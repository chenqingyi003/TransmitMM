#!/usr/bin/env python 
# -*- coding:utf-8 -*-

# __author__ = 'Hugo Chen'
# __time__   = '2018/12/4'
# God helps who help themselves!

from PyQt5.QtCore import *
from PyQt5.QtWidgets import QLineEdit


class InputLineEdit(QLineEdit):
    dbClickLineSignal = pyqtSignal(str)  # 定义clicked信号,带str类型参数的

    def __init__(self,parent=None):
        super(InputLineEdit,self).__init__(parent)
        self.setPlaceholderText("Please Input your name by English!")

    #重写mouseDoubleClickEvent,双击选择当前文本:
    def mouseDoubleClickEvent(self, QMouseEvent):
        if QMouseEvent.button() == Qt.LeftButton:
            self.selectAll()
            self.dbClickLineSignal.emit(self.text())  # 发送clicked信号

def printTest():
    print("This is a demo print!")

if __name__ == '__main__':
    print("This is InputLineEdit!")