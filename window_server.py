#!/usr/bin/env python 
# -*- coding:utf-8 -*-

# __author__ = 'Hugo Chen'
# __time__   = '2019/2/1'
# God helps who help themselves!


import sys,os
from PyQt5 import QtCore,QtWidgets,QtGui
from PyQt5.QtWidgets import *

#导入本地自定义模块
from InputLineEdit import *
from file_server_thread import *

class MainFrame(QtWidgets.QWidget):
    #1. 构造函数和初始化函数放在最前面:
    def __init__(self,type,parent=None):    #init函数是传参数时使用的
        super(MainFrame,self).__init__(parent)    #调用父类的构造函数
        self.type = type    #主界面类型

    #2. 事件处理函数:函数重载
    def closeEvent(self, event):
        reply = QMessageBox.question(self, 'Warning', '\nAre you sure to quit?\t\n',\
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()
        pass

    #3. 界面窗口相关设置
    def setupWindow(self):
        # 设置软件所有字体
        font = QtGui.QFont()
        font.setPixelSize(12)    # 字体大小
        font.setFamily(("Arial"))    # 字体类型:Arial,Verdana,New Courier
        self.setFont(font)
        # 设置工具窗口的标题
        self.setWindowTitle(self.title)
        # 设置窗口的图标
        self.setWindowIcon(QtGui.QIcon("images/app_icon.png"))
        # 设置初始化窗口的大小
        self.resize(self.size)

    #4. 初始化变量
    def initialValue(self):
        if self.type == 0:    #默认类型的话
            self.title = "File Transmission Server"
            self.size = QtCore.QSize(800, 200)
        else:
            self.title = "File Transmission Client"
            self.size = QtCore.QSize(800, 200)

    #5 生成界面
    def setupUi(self):
        mainVlyt = QVBoxLayout()

        # 1.第一行是Server的红色标题:
        titleHlyt = QHBoxLayout()
        titleLbl = QLabel("File Save Server",self)
        titltLblStyle = "QLabel{color: #EE9E9E;\
                    font-size: 16px;\
                    min-height: 26px;\
                    font-weight: bold;\
                    padding-left: 10px;\
                    border: 0 px outset #9A989B;\
                    background-color: #F0F0F0;}"
        titleLbl.setStyleSheet(titltLblStyle)
        titleHlyt.addStretch()
        titleHlyt.addWidget(titleLbl)
        titleHlyt.addStretch()

        # 2.第二行是Server的IP和PORT的设置:
        ipLabel = QLabel("Server IP:",self)
        self.ipLine = QLineEdit(self)
        portLabel = QLabel("PORT:",self)
        self.portLine = QLineEdit(self)
        self.ipLine.setText('192.168.1.124')  # 此处服务端绑定的IP必须是本机的IP
        self.portLine.setText('8080')
        self.setBtn = QPushButton("Set",self)
        self.setBtn.clicked.connect(self.setIpPortSlot) #建立槽函数
        ipPortHlyt = QHBoxLayout()
        ipPortHlyt.addWidget(ipLabel)
        ipPortHlyt.addWidget(self.ipLine)
        ipPortHlyt.addWidget(portLabel)
        ipPortHlyt.addWidget(self.portLine)
        ipPortHlyt.addWidget(self.setBtn)

        # 3.第三行是文件路径和选择按钮:
        inputLyt = QHBoxLayout()
        nameLbl = QLabel("File Path:",self)    # 将窗口传给控件作为parent
        self.nameLine = InputLineEdit(self)
        self.nameLine.dbClickLineSignal.connect(self.updateLbl)  # 建立槽函数
        self.chooseBtn = QPushButton("Choose..", self)
        self.chooseBtn.clicked.connect(self.chooseLocalFile)  # 建立槽函数

        inputLyt.addWidget(nameLbl)
        inputLyt.addWidget(self.nameLine)
        inputLyt.addWidget(self.chooseBtn)

        # 4.第四行是Save按钮:
        self.sendBtn = QPushButton("Save", self)
        self.sendBtn.clicked.connect(self.save2File)    #建立槽函数

        # 5.第五行是显示对话框:
        self.outputText = QTextBrowser(self)

        mainVlyt.addLayout(titleHlyt)
        mainVlyt.addLayout(ipPortHlyt)
        mainVlyt.addLayout(inputLyt)
        mainVlyt.addWidget(self.sendBtn)
        mainVlyt.addWidget(self.outputText)
        mainVlyt.setAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter )
        self.setLayout(mainVlyt)
        self.setupWindow()

    #6. 启动相关进程:
    def prepare2Work(self):
        self.listenThd = ListenThread()
        self.listenThd.recvSignal.connect(self.setReceiveSlot)

    #7. 界面使用的相关函数，槽函数+普通函数
    def updateLbl(self,nameStr):
        if nameStr!="":
            self.outputText.setText("Get File Path: "+nameStr)
        pass

    def chooseLocalFile(self): # 选择本地文件
        filePath = QFileDialog.getSaveFileName(self,'Save file to','','')[0]
        print(filePath)
        if not filePath.strip():
            pass
        else:
            self.nameLine.setText(filePath)

    def setIpPortSlot(self):
        ipStr = self.ipLine.text().strip()
        portStr = self.portLine.text().strip()
        self.listenThd.setServerIpPort(ipStr,portStr)    #设置监听进程IP和端口
        self.listenThd.start()    #启动监听进程
        pass

    def save2File(self):
        file_path = self.nameLine.text().strip()
        if file_path=='':
            QMessageBox.warning(self, 'Warning', '\nPlease choose a valid path first!\t\n',QMessageBox.Yes)
        else:
            allData = self.outputText.toPlainText()
            with open(file_path, 'w', encoding='utf-8')as f:
                f.write(allData)
            QMessageBox.information(self, 'Information', '\nFile save finished!\t\n', QMessageBox.Yes)
        pass

    def setReceiveSlot(self,allData):
        self.outputText.setText(allData)
        pass

if __name__=='__main__':
    #创建应用程序和对象
    app = QtWidgets.QApplication(sys.argv)
    printTest()
    ui = MainFrame(0)    #创建主界面实例
    ui.initialValue()    #初始化界面变量数值
    ui.setupUi()    #创建主界面UI
    ui.prepare2Work()    #创建监听进程,未启动
    ui.show()    #显示界面
    sys.exit(app.exec_())