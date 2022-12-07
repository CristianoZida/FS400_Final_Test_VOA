# -*- coding: utf-8 -*-
# this module is to write the control functions of ZVA
# date:4/14/2022
# author:Jiawen Zhou
import time
import pyvisa
#import os
#from pyvisa import constants

rm=pyvisa.ResourceManager()
rm.list_resources()

def open_ZVA(obj,zva_port):
    '''
    Open ZVA connection
    :param obj: GUI Class obj
    :param zva_port: port number
    :return: True/False
    '''
    try:
        obj.ZVA = rm.open_resource(zva_port)
        obj.ZVA.timeout=5
        # rm.visalib.set_buffer(obj.ZVA.session,constants.VI_IO_IN_BUF,1048576)
        # rm.visalib.set_buffer(obj.ZVA.session,constants.VI_IO_OUT_BUF,1048576)
        ins=obj.ZVA.query('*IDN?')
        if "Rohde&Schwarz" in ins:
            return True
        else:
            return False
    except Exception as e:
        print(e)
        return False

def close_ZVA(obj):
    '''
    Close PM connection
    :param obj: GUI Class obj
    :return: True/False
    '''
    try:
        if obj.ZVA in rm.list_opened_resources():
            obj.ZVA.close()
            print('ZVA端口关闭成功!')
    except Exception as e:
        print(e)
        return

def init_ZVA(obj):
    '''
    initiate ZVA
    :param obj: GUI Class obj
    :return: query str
    '''
    obj.ZVA.write('SYST:DISP:UPD On')
    time.sleep(0.1)
    obj.ZVA.write('*rst')
    time.sleep(0.1)
    return obj.ZVA.query('*opc?') #if return'1\n' then initiation sucessed

def recall_cal(obj,calfile,ch):
    '''
    load calibration file
    :param obj: GUI Class obj
    :param calfile: calibration file path
    :param ch: channel
    :return: True/False
    '''
    obj.ZVA.write('*rst')
    time.sleep(0.1)
    obj.ZVA.write('MMEM:CDIR \'C:\\Rohde&Schwarz\\Nwa\\RecallSets\'')
    #obj.ZVA.write('MMEM:CDIR \'C:\\Users\\Instrument\\Desktop\\AutoData\'')
    time.sleep(0.1)
    calfile=str(calfile)+'cal-SW'+str(ch)+'.zvx'
    obj.ZVA.write('MMEM:load:stat 1, \''+calfile+'\'')
    #time.sleep(3)

def singleTest(obj):
    '''
    perform single test
    :param obj: GUI Class obj
    :return: query str
    '''
    obj.ZVA.write('INIT:CONT off')
    time.sleep(0.1)
    # obj.ZVA.write('INIT:IMM')
    # time.sleep(5)
    # return obj.ZVA.query('*opc?')
    obj.ZVA.write('INIT:IMMediate *OPC')
    time.sleep(0.1)
    return obj.ZVA.query('*opc?')

def savedata(obj,sn,fre_ch,temp,timestr,ch):
    '''
    store test data to ZVA local folder
    :param obj: GUI Class obj
    :param sn:
    :param fre_ch:
    :param temp:
    :param timestr:
    :param ch:
    :return: query str
    '''
    ch=str(ch)
    if ch=='1':
        biasch='XI'
    elif ch=='2':
        biasch='XQ'
    elif ch=='3':
        biasch='YI'
    elif ch=='4':
        biasch='YQ'
    else:
        print('Channel error')
        return
    filename=sn+'_400GCH'+str(fre_ch)+'_'+str(temp)+'oC_'+biasch+'_'+timestr+'.csv'
    obj.ZVA.write('MMEM:CDIR \'C:\\Users\\Instrument\\Desktop\\AutoData\'')
    time.sleep(0.1)
    obj.ZVA.write(':MMEM:STOR:TRAC:chan 1, \''+filename+'\',form')
    time.sleep(0.1)
    return [obj.ZVA.query('*opc?'),filename]

def read_data(obj,filename):
    '''
    Read test data from local ZVA folder and read to storage ignore coding errors
    :param obj: GUI Class obj
    :param filename:
    :return: the data queryed from the ZVA
    '''
    obj.ZVA.write('MMEM:DATA? \''+filename+'\'')#test_400GCH13_25oC_XI_0420A.csv\'')
    time.sleep(0.1)
    data=obj.ZVA._read_raw().decode('utf-8',errors='ignore')
    return data