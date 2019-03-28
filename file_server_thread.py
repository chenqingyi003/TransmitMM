#!/usr/bin/env python 
# -*- coding:utf-8 -*-

# __author__ = 'Hugo Chen'
# __time__   = '2019/2/1'
# God helps who help themselves!

from PyQt5.QtCore import *
import struct, json, socket

'''
IP = '192.168.2.68'  # 此处服务端绑定的IP必须是本机的IP
PORT = 8080
ADD = (IP, PORT)
sever = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
sever.bind(ADD)
sever.listen(5)
'''

class ListenThread(QThread):
    # 最先定义信号类型:
    recvSignal = pyqtSignal(str)  # 定义clicked信号,带str类型参数的

    def __init__(self,parent=None):    #init函数是传参数时使用的
        super(ListenThread,self).__init__(parent)    #调用父类的构造函数

    # 先设置好服务端IP端口等信息,再等待里连接:
    def setServerIpPort(self,ip,port):
        # IP = '192.168.2.68'  # 此处服务端绑定的IP必须是本机的IP
        # PORT = 8080
        ADD = (ip, int(port))    # 端口的格式是int型
        self.sever = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        try:
            self.sever.bind(ADD)
        except Exception as e:
            self.sever.close()
            print('Exception: ', e)
        else:
            self.sever.listen(5)
            self.recvSignal.emit("waiting for connect...")  # 设置完成后再发送等待连接的信号;

    # 服务端一直开启着,等待客户端的连接,每连接传输一次文件，需要点一下回车，按q退出程序.
    def listen2Client(self):
        while True:
            conn, addr = self.sever.accept()
            '''获取报头长度bytes'''
            s_hander = conn.recv(4)

            '''解包bytes-》tuple-》int，获得报头长度'''
            s_hander = struct.unpack('i', s_hander)[0]
            print('s_hander: '+str(s_hander))

            '''获取报头数据，bytes'''  #此处是二进制
            b_hander = conn.recv(s_hander)
            print(b_hander)

            '''报头数据解码 bytes-》str'''
            json_hander = b_hander.decode('utf-8')
            print('json_hander: '+json_hander)

            '''报头数据反序列化 str-》dict'''
            hander = json.loads(json_hander)
            print(hander)

            '''获取报头字典，取的文件长度，取出文件内容'''
            file_size = hander['length']
            res = b''
            size = 0
            while size < file_size:
                data = conn.recv(1024)
                size += len(data)
                res += data
            self.recvSignal.emit(res.decode('utf-8'))

    # 事件处理函数:函数重载
    def run(self):
        self.listen2Client()

if __name__ == '__main__':
    print("This is main function!")

    print("Application finished!")
