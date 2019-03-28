#!/usr/bin/env python 
# -*- coding:utf-8 -*-

# __author__ = 'Hugo Chen'
# __time__   = 2019/3/28
# God helps who help themselves!

from PyQt5.QtCore import *
import socket, os, json, struct

class SendThread(QThread):
    # 最先定义信号类型:
    sendFileSignal = pyqtSignal(str)  # 定义clicked信号,带str类型参数的

    def __init__(self, parent=None):  # init函数是传参数时使用的
        super(SendThread, self).__init__(parent)  # 调用父类的构造函数

    # 先设置好服务端IP端口等信息,再等待里连接:
    def setSend2IpPort(self, ip, port):
        # 此处服务端绑定的IP是服务端(发送目的地)的IP
        self.ADD = (ip, int(port))  # 端口的格式是int型
        self.sendFileSignal.emit("Waiting for sending file...")  # 设置完成后再发送等待连接的信号;

    # 设置好要发送的文件,然后再启动进程:
    def setSend2FilePath(self, filePath):
        self.filePath = filePath

    def uploadFile(self):
        # 上传文件
        #ADD = (IP, PORT)
        client = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        try:
            client.connect(self.ADD)
        except Exception as e:
            client.close()
            self.sendFileSignal.emit('Exception: ', e)
        else:
            while True:
                file_path = self.filePath.strip()
                if not os.path.exists(file_path):
                    self.sendFileSignal.emit('文件不存在')
                    continue
                file_name = (os.path.split(file_path))[1]
                self.sendFileSignal.emit('Uploading filename:', file_name)
                with open(file_path, 'r', encoding='utf-8')as f:
                    data = f.read()
                    size = len(data)

                hander = {
                    'file_name': file_name,
                    'length': size
                }
                # 报头序列化
                hander_json = json.dumps(hander)
                # 报头bytes转换
                hander_bytes = hander_json.encode('utf-8')
                # 报头长度固定
                s_hander = struct.pack('i', len(hander_bytes))
                # 传输报头长度
                client.send(s_hander)
                # 传输报头数据s
                client.send(hander_bytes)
                # 传输文件数据
                client.send(data.encode('utf-8'))
                break
            client.close()
            self.sendFileSignal.emit("Uploading finished!")  # 发送完成后打印;

    # 事件处理函数:函数重载
    def run(self):
        self.uploadFile()

if __name__ == '__main__':
    print("This is main function!")
    #uploadFile('D:/data.txt')
    print("Application finished!")
    exit(0)