#!/usr/bin/env python 
# -*- coding:utf-8 -*-

# __author__ = 'Hugo Chen'
# __time__   = '2019/2/1'
# God helps who help themselves!

import struct, json, socket

IP = '192.168.2.68'
PORT = 8080
ADD = (IP, PORT)
sever = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
sever.bind(ADD)
sever.listen(5)

# 服务端一直开启着,等待客户端的连接,每连接传输一次文件，需要点一下回车，按q退出程序.
def listen2Client():
    n=0
    while True:
        n = n+1
        status = input("Pause, click any key to go on!")
        if status=='q':
            print('HAHA, Byebye!')
            break
        print("waiting for connect...", n)

        conn, addr = sever.accept()
        '''获取报头长度bytes'''
        s_hander = conn.recv(4)

        '''解包bytes-》tuple-》int，获得报头长度'''
        s_hander = struct.unpack('i', s_hander)[0]
        print(s_hander)

        '''获取报头数据，bytes'''
        b_hander = conn.recv(s_hander)
        print(b_hander)

        '''报头数据解码 bytes-》str'''
        json_hander = b_hander.decode('utf-8')
        print(json_hander)

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
        print(res.decode('utf-8'))

        # date = (conn.recv(file_size)).decode('utf-8')
        # print(date)

if __name__ == '__main__':
    print("Waiting for connect...")
    listen2Client()
    print("Application finished!")
