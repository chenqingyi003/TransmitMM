#!/usr/bin/env python 
# -*- coding:utf-8 -*-

# __author__ = 'Hugo Chen'
# __time__   = '2019/2/1'
# God helps who help themselves!
import socket, os, json, struct

IP = '192.168.2.68'
PORT = 8080

def uploadFile(filePath):
    # 上传文件
    ADD = (IP, PORT)
    client = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    client.connect(ADD)
    while True:
        file_path = filePath.strip()
        if not os.path.exists(file_path):
            print('文件不存在')
            continue
        file_name = (os.path.split(file_path))[1]
        print('uploading filename:',file_name)
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
    print("Uploading finished!")

if __name__ == '__main__':
    print("This is main function!")
    uploadFile('D:/data.txt')
    print("Application finished!")
    exit(0)
