# -*- coding: utf-8 -*-
#this module is to write the control functions of power meter
# date:3/17/2022
# author:Jiawen Zhou
import serial
import math
import time

def open_PM(obj,pm_port):
    '''
    Open PM connection
    :param obj: GUI Class obj
    :param pm_port: port number
    :return: True/False
    '''
    try:
        obj.PM = serial.Serial(pm_port,9600,timeout=3)
        return True
    except Exception as e:
        print(e)
        return False

def close_PM(obj):
    '''
    Close PM connection
    :param obj: GUI Class obj:
    :return: True/False
    '''
    try:
        if serial.Serial.isOpen(obj.PM):
            obj.PM.close()
            print('功率计串口关闭成功!')
    except Exception as e:
        print(e)
        return

def read_PM(obj,ch,read_times=10):
    '''
    read PM power, read 10 times default  to keep light power monitor accurate
    :param obj: GUI Class obj
    :param ch: Channel of PM
    :param read_times: read times
    :return: The power--round(pwr,2)
    '''
    sum=0
    for i in range(read_times):
        pwr=-75
        if str(ch)=='1':
            obj.PM.write(b'\xEE\xAA\x03\x01\x00\x00')
        elif str(ch)=='2':
            obj.PM.write(b'\xEE\xAA\x03\x02\x00\x00')
        else:
            pwr=0
            print("Error Power meter channel or CH input error")
            return 0 #"Error Power meter channel"
        time.sleep(0.1)
        p=obj.PM.read_all()
        if len(p)!=0:
            a,b=[i for i in p]
            pwr=((a<<8)+b)/100-100
        else:
            return 0
        sum+=10**(pwr/10)
    pwr=10*math.log10(sum/read_times)
    return round(pwr,2)

def get_wavelength(obj,ch):
    '''
    Get PM wavelength
    :param obj: GUI Class obj
    :param ch: channel
    :return: The setting wavelength
    '''
    wave='1550'
    if str(ch)=='1':
        obj.PM.write(b'\xEE\xAA\x01\x01\x00\x00')
    elif str(ch)=='2':
        obj.PM.write(b'\xEE\xAA\x01\x02\x00\x00')
    else:
        wave=''
        return 0
    if len(wave)!=0:
        time.sleep(0.3)
        p=obj.PM.read_all()
        if len(p)!=0:
            #need to check the output value here
            wave=[str(i) for i in p]
            wave=wave[0]+wave[1]
        else:
            return 0
    return wave

def int2byte(a):
    '''
    convert the integer to bytearray
    :param a: integer
    :return: bytearray
    '''
    b=hex(a).replace('0x','')
    if len(b)==1:
        b="0"+b
    c=bytearray.fromhex(b)
    return c

def set_wavelength(obj,ch,wave=1550):
    '''
    Set PM wavelength
    :param obj: GUI Class obj
    :param ch: channel
    :param wave: wavelength
    :return: 1 and 0(success or failed)
    '''
    #wave=int(wave)
    if not 800<=wave<=1700:
        return 0
    wave_1=Get_wavelength(obj,ch)
    if str(wave_1)==str(wave):
        return 1
    else:
        a=int2byte(wave//100)+int2byte(int(round(wave%100,0)))
        if ch==1:
            obj.PM.write(b'\xEE\xAA\x04'+a+b'\x00')
        elif ch==2:
            obj.PM.write(b'\xEE\xAA\x42'+a+b'\x00')
        else:
            return 0

        time.sleep(0.3)
        p=obj.PM.read_all()
        if len(p)!=0:
            return 1
        else:
            return 0