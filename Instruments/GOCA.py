# -*- coding: utf-8 -*-
# this module is to write the control functions of GOCA
# date:4/14/2022
# author:Jiawen Zhou
import time
import socket

# s=socket.socket(socket.AF_INET,socket.SOCK_STREAM)

def open_GOCA(obj,goca_port):
    '''
    Open ZVA connection
    :param obj: GUI Class obj
    :param goca_port:port number
    :return: True/False
    '''
    try:
        obj.GOCA =socket.socket(socket.AF_INET,socket.SOCK_STREAM)
        (host,port)=goca_port.split('::')
        obj.GOCA.connect((host,int(port)))
        if True:
            return True
        else:
            return False
    except Exception as e:
        print(e)

def close_GOCA(obj):
    '''
    Close GOCA connection
    :param obj: GUI Class obj
    :return: True/False
    '''
    try:
        obj.GOCA.close()
        print('GOCA端口关闭成功!')
    except Exception as e:
        print(e)

def init_GOCA(obj):
    '''
    initiate GOCA
    :param obj: GUI Class obj
    :return: True/False
    '''
    try:
        obj.GOCA.sendall(b'goca')
        time.sleep(0.3)
        obj.GOCA.sendall(b'LD1310off')
        time.sleep(0.3)
        for i in range(3):
            obj.GOCA.sendall(b'LD1550off')
            time.sleep(3)
            data=obj.GOCA.recv(1024).decode('utf-8')
            if not 'OK'in data:
                i+=1
                if i==3:
                    print('光座初始化失败！(LDoff)')
                    return False
                continue
            else:
                print('光座初始化(LDoff)')
                break
        for i in range(3):
            obj.GOCA.sendall(b'EOmode')
            time.sleep(3)
            data1=obj.GOCA.recv(1024).decode('utf-8')
            if not 'OK'in data:
                i+=1
                if i==3:
                    print('光座初始化失败！(EO mode)')
                    return False
                continue
            else:
                print('光座初始化成功！')
                return True

        # if not 'OK'in data:
        #     print('光座初始化失败！(LDoff)')
        #     return False
        # else:
        #     obj.GOCA.sendall(b'EOmode')
        #     time.sleep(3)
        #     data1=obj.GOCA.recv(1024).decode('utf-8')
        #     if not 'OK'in data1:
        #         obj.GOCA.sendall(b'EOmode')
        #         time.sleep(3)
        #         data1=obj.GOCA.recv(1024).decode('utf-8')
        #         if not 'OK'in data1:
        #             print('光座初始化失败！(EO mode)')
        #             return False
        # else:
        #     print('光座初始化成功！')
        #     return True
    except Exception as e:
        print(e)
