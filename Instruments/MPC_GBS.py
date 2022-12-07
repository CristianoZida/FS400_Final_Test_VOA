# -*- coding: utf-8 -*-
#this module is to write the control functions of power meter
# date:3/17/2022
# author:Jiawen Zhou
import serial

def open_MPC(obj,mpc_port):
    '''
    Open MPC connection
    :param obj: GUI Class obj
    :param mpc_port: MPC port number
    :return: True/False
    '''
    try:
        obj.MPC = serial.Serial(mpc_port,115200,timeout=3)
        return True
    except Exception as e:
        print(e)
        return False

def close_MPC(obj):
    '''
    Close MPC connection
    :param obj: GUI Class obj:
    :return: True/False
    '''
    try:
        if serial.Serial.isOpen(obj.MPC):
            obj.MPC.close()
            print('MPC串口关闭成功!')
    except Exception as e:
        print(e)
        return

def set_vol(obj,ch,v):
    '''
    set the voltage of channel1
    :param obj:
    :param ch:1-4
    :param v: voltage to set (0-140)
    :return: NA
    '''
    if not 0<=v<=140:print('voltage set error!');return
    obj.MPC.write('SOUR{}:VOLT {}\n'.format(ch,v).encode('utf-8'))


def get_ch(obj,ch):
    '''
    get the voltage set
    :param obj:
    :param ch:channel to query,1-4
    :return: voltage set
    '''
    obj.MPC.write('SOUR{}:VOLT?\n'.format(ch).encode('utf-8'))
    s=obj.MPC.read_all().decode('utf-8').split('.')[0]
    return int(s)



