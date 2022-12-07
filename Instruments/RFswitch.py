# -*- coding: utf-8 -*-
# this module is to write the control functions of RF switch
# date:4/14/2022
# author:Jiawen Zhou
import time
import serial

def open_RFSW(obj,rfsw_port):
    '''
    Open RFSW connection
    :param obj: GUI Class obj
    :param rfsw_port: port number
    :return: True/False
    '''
    try:
        obj.RFSW=serial.Serial(rfsw_port,9600,timeout=5)
        return True
    except Exception as e:
        print(e)
        return False

def close_RFSW(obj):
    '''
    Close RF switch connection
    :param obj: GUI Class obj
    :return: True/False
    '''
    try:
        if serial.Serial.isOpen(obj.RFSW):
            obj.RFSW.close()
            print('RF开关串口关闭成功!')
    except Exception as e:
        print(e)
        return

def switch_RFchannel(obj,ch):
    '''
    Switch the RF channel
    :param obj: GUI Class obj
    :param ch:
    :return: True/False
    '''
    ch=str(ch)
    if 0<=int(ch)<=4:
        cmd1=('SetSW:01,K01,'+ch+'\n').encode('utf-8')
        cmd2=('SetSW:01,K02,'+ch+'\n').encode('utf-8')
        obj.RFSW.write(b'ReadSW:01\n')
        time.sleep(0.1)
        data=obj.RFSW.read_all().decode('utf-8')
        data=data[data.index('SW'):data.index('END')-1].split(',0')
        if not data[1]==data[2]==ch:
            obj.RFSW.write(cmd1)
            time.sleep(0.1)
            obj.RFSW.write(cmd2)
            time.sleep(0.1)
            obj.RFSW.write(b'ReadSW:01\n')
            time.sleep(0.1)
            data=obj.RFSW.read_all().decode('utf-8')
            data=data[data.index('SW'):data.index('END')-1].split(',0')
            if data[1]==data[2]==ch:
                return True
            else:
                return False
                print("channel set error!, please check the RF switch")
        else:
            return True
    else:
        return False
        print("channel input error")