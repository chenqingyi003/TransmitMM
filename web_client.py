#!/usr/bin/env python 
# -*- coding:utf-8 -*-

# __author__ = 'Hugo Chen'
# __time__   = '2019/1/25'
# God helps who help themselves!

# import re
# 
# regex = '[\+\-]?[\d]+([\.][\d]*)?([Ee][+-]?[\d]+)?'
# 
# print( re.match(regex,'102909') )

import socket

skt = socket.socket()
host = '192.168.2.68'
port = 12345

print("host: ",host," port: ",port)

skt.connect((host,port))

# 接收数据:
buffer = []
nn = 0
while True:
    # 每次最多接收1k字节:
    nn+=1
    data = skt.recv(1024)
    # print("get data %d: %s", (nn,data) )
    if data:
        buffer.append(data)
    else:
        print('finish receive data:',nn)
        break
dataAll = b''.join(buffer)
print("get all data:",dataAll.decode('utf-8'))



#print('1: ' + skt.recv(19).decode('utf-8'))   # 此处的16是指接受了16个字符
#print('2: ' + skt.recv(19).decode('utf-8'))   # 此处的16是指接受了16个字符
#print('3: ' + skt.recv(16).decode('utf-8'))   # 此处的16是指接受了16个字符
#print('4: ' + skt.recv(16).decode('utf-8'))   # 此处的16是指接受了16个字符
skt.close()

print("Connect finished!")

'''
if __name__ == '__main__':
    print("This is main function!")

    print("Application finished!")
'''