#!/usr/bin/env python 
# -*- coding:utf-8 -*-

# __author__ = 'Hugo Chen'
# __time__   = '2019/1/25'
# God helps who help themselves!

import socket

skt = socket.socket()
host = socket.gethostname()
port = 12345
skt.bind((host,port))

skt.listen(5)   # 等待客户端连接,一般最小为1,最大为5

n=0

while True:
    n = n+1
    status = input("Pause, click any key to go on!")
    if status=='q':
        print('HAHA, Byebye!')
        break
    print("waiting for connect...", n)
    conn, addr = skt.accept()
    print('连接地址:', addr)
    # msg = '欢迎访问菜鸟教程！' + "\r\n"
    # conn.send(msg.encode('utf-8'))
    conn.send("Welcome 2 Hugo Chen".encode('utf-8'))
    conn.send("这个厉害了,这次是中文！".encode('utf-8')) # utf-8编码一个汉字占3个字节
    conn.close()