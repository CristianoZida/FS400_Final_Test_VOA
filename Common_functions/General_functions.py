# -*- coding: utf-8 -*-
# date:3/17/2022
##update on 07192022:add function to read MCU and FPGA version
# author:Jiawen Zhou
#this module is to write the general functions of DC test##3/17/2022
import os.path
import time

import pandas as pd
import serial
import sys
import numpy as np
import math

# ------------------------------------------------------------------------------------------------

# log output to console and file
# http://stackoverflow.com/questions/14906764/how-to-redirect-stdout-to-both-file-and-console-with-scripting
class Logger_old(object):
    def __init__(self):
        self.terminal = sys.stdout
        self.log = open("logfile.log", "a")

    def write(self, message):
        self.terminal.write(message)
        self.log.write(message)

    def flush(self):
        # this flush method is needed for python 3 compatibility.
        # this handles the flush command by doing nothing.
        # you might want to specify some extra behavior here.
        pass

# ------------------------------------------------------------------------------------------------

class Logger(object):
    def __init__(self, path=r"C:\Test_result"):
        self.terminal = sys.stdout
        self.path = path
        if not os.path.exists(self.path):
            os.mkdir(self.path)
        #distutils.dir_util.mkpath(self.path)
        self.name = self.path + "\\logfile.log"
        self.log = open(self.name, "a")
        self.log_enabled = True

    def write(self, message):
        self.terminal.write(message)
        if self.log_enabled:
            self.log.write(message)

    def WriteLogOnly(self, message):
        self.log.write(message + '\n')

    def WriteTerminalOnly(self, message):
        self.terminal.write(message)

    def ChangeLogFile(self, ANewName):
        # finish the old one
        self.log.close()
        self.name = ANewName
        # the new one
        self.log = open(self.name, "a")

    def CloseLogFile(self):
        # finish the old one
        self.log.close()

    def OpenLogFile(self):
        self.log = open(self.name, "a")

    def flush(self):
        # this flush method is needed for python 3 compatibility.
        # this handles the flush command by doing nothing.
        # you might want to specify some extra behavior here.
        # self.log.flush() does not work
        pass
        '''
        # leads to delay on cmd-line
        self.CloseLogFile()
        time.sleep(0.5)
        self.OpenLogFile()
        '''
def SN_check(sn):
    '''
    Check SN
    :param sn: Serial Number#For example:"FS2022010190"
    :return:True/False
    '''
    #return True
    if sn!='':
        return True
    else:
        print('SN未输入,请检查SN！')
        return False
    # if sn[0:4]!='FS20' or len(sn)!=12:
    #     print('SN输入有误，请检查SN！')
    #     return False
    # else:
    #     return True

def plot_data(obj,data_show):
    '''
    Plot data into the groupbox
    :param obj:
    :param data_show:Dataframe
    :return: Curve in the main window
    '''
    pass

def printout_log(obj,log=''):
    '''
    Print out log into the plain text edit
    :param obj:
    :param log: string plus the sys time
    :return: NA
    '''
    #obj.test_status.setStyleSheet("background-color: rgb(255, 0, 0);")
    #timestamp=time.strftime("%Y/%m/%d")+ "_" + time.strftime("%I:%M:%S")+' '
    #log=timestamp+log
    obj.plainTextEdit.setPlainText(log)#'2212341\nasfasfsf\nasfasfasf\nasfjaljsfk\nset color done!\n2212341\nasfasfsf\nasfasfasf\nasfjaljsfk\nset color done!\n2212341\nasfasfsf\nasfasfasf\nasfjaljsfk\nset color done!\n2212341\nasfasfsf\nasfasfasf\nasfjaljsfk\nset color done!\n2212341\nasfasfsf\nasfasfasf\nasfjaljsfk\nset color done!\n2212341\nasfasfsf\nasfasfasf\nasfjaljsfk\nset color done!\n2212341\nasfasfsf\nasfasfasf\nasfjaljsfk\nset color done!\n2212341\nasfasfsf\nasfasfasf\nasfjaljsfk\nset color done!\n2212341\nasfasfsf\nasfasfasf\nasfjaljsfk\nset color done!\n')

def get_timestamp(s=0):
    '''
    get current time in different format
    :param s: input:format style##
     1:for report generate
     0:for log record
    :return: formated current time
    '''
    if s==0:
        timestamp=time.strftime("%Y/%m/%d") + "_" + time.strftime("%H:%M:%S")
    elif s==1:
        timestamp=time.strftime("%Y%m%d") + "_" + time.strftime("%H%M%S")
    else:
        print('error input,style can only be 0 or 1!')
        return
    return timestamp

def read_MCUandFPGA(obj):
    '''
    Read the MCU and FPGA information from the Board
    :param obj:
    :return: (testBrdClp,ctlBrdClp,ctlBrdModVer,ctlBrdFPGAver,MCUver)
    '''
    try:
        if serial.Serial.isOpen(obj.CtrlB):
            #Read test board cpld update date
            obj.CtrlB.write(b'cpld_spi_rd 0x007 3\n')
            time.sleep(0.1)
            s=obj.CtrlB.read_all().decode('utf-8')
            s1=s.split('\r\n')[1].strip().split(' ')
            clp_date=''
            for i in s1:
                clp_date+=i[2:]+'_'
            testBrdClp=clp_date[:-1]

            #Read ontrol board cpld update date
            obj.CtrlB.write(b'cpld_rd 0x0010 3\n')
            time.sleep(0.1)
            s=obj.CtrlB.read_all().decode('utf-8')
            s1=s.split('\n')[1].strip().split('  ')
            clp_date=''
            for i in s1:
                clp_date+=i+'_'
            ctlBrdClp=clp_date[:-1]

            #Read control board module version
            obj.CtrlB.write(b'cpld_rd 0x0000 16\n')
            time.sleep(0.1)
            s=obj.CtrlB.read_all().decode('utf-8')
            s1=s[s.index('ASCII:\r\n')+8:].strip().split('\r\n')
            s2=(s1[0]+s1[1]).split('\t')
            clp_date=''
            for i in s2:
                clp_date+=i.strip()
            ctlBrdModVer=clp_date

            #Read control board FPGA version
            obj.CtrlB.write(b'fpga_spi_rd 0x0002\n')
            time.sleep(0.1)
            s=obj.CtrlB.read_all().decode('utf-8')
            s1=s.split('\r\n')[1].split('=')[1][2:-1]
            ctlBrdFPGAver=s1

            #Read MCU version
            obj.CtrlB.write(b'ver\n')
            time.sleep(0.1)
            s = obj.CtrlB.read_all().decode('utf-8')
            s1=s.split('\r\n')
            s2=s1[0].split('\n')[1]+s1[1]
            MCUver = s2 #For example:'Data:Jun 15 2022.Time:14:55:42.'

            print('testBrdClp,ctlBrdClp,ctlBrdModVer,ctlBrdFPGAver,MCUver:\n',(testBrdClp,ctlBrdClp,ctlBrdModVer,ctlBrdFPGAver,MCUver))
            return (testBrdClp,ctlBrdClp,ctlBrdModVer,ctlBrdFPGAver,MCUver)
        else:
            print('Error, Control board port is not open')
    except Exception as e:
        print('MCU or FPGA version reading error:',e)

def str_to_ascii(s):
    '''
    Turn the string to ASCII hex code
    :param s: String
    :return: ASCII hex code list[]
    '''
    return [hex(ord(i)) for i in s]

def ascii_to_str(a):
    '''
    Turn the ASCII hex code to string
    :param a: ASCII hex code lis[]
    :return: String
    '''
    s=''
    for i in a:
        s+=chr(int(i,16))
    return s

def write_SN(obj,sn):
    '''
    Write the SN into the eeprom
    :param obj,sn
    :return: NA
    '''
    try:
        sn=str(sn)
        if len(sn)!=10:
            print('SN input error lenth is {}, not 10!!!'.format(len(sn)))
            obj.CtrlB.close()
            return
        ascii_sn=str_to_ascii(sn)
        if serial.Serial.isOpen(obj.CtrlB):
            obj.CtrlB.write(b'slave_on\n')
            time.sleep(0.1)
            obj.CtrlB.write(b'cpld_spi_wr 0x31 1\n')
            time.sleep(0.1)
            obj.CtrlB.write(b'fs400_eep_arg_read 1 0\n')
            time.sleep(0.1)
            a = obj.CtrlB.read_until(b'check ok')
            time.sleep(0.1)
            b=str_to_ascii(sn)
            cod1 = b[0] + b[1][2:] + b[2][2:] + b[3][2:]
            cod2 = b[4] + b[5][2:] + b[6][2:] + b[7][2:]
            cod3 = b[8] + b[9][2:] + '20' + '20'
            cod4 = '0x20202020'
            obj.CtrlB.write('fs400_set_eep_sn {} {} {} {}\n'.format(cod1,cod2,cod3,cod4).encode('utf-8'))
            time.sleep(0.1)
            obj.CtrlB.write(b'fs400_eep_arg_save 1 0\n');time.sleep(0.2)
            print(obj.CtrlB.read_until(b'Save ok'))
            obj.CtrlB.write(b'fs400_eep_arg_read 1 0\n')
            time.sleep(0.1)
            a = obj.CtrlB.read_until(b'check ok').decode('utf-8')
            time.sleep(0.1)
            print(a)
            sn_read=read_eeprom(obj)
            if sn_read==sn:
                print('已写入SN：{}'.format(sn))
            else:
                print('写入SN失败，读取的SN为{},而非{}'.format(sn_read,sn))
            # obj.CtrlB.write(b'slave_off\n')
            # obj.CtrlB.close()
    except Exception as e:
        print(e,'Please check whether board or eeprom is up and configed!!!')

def read_eeprom(obj,s=1):
    '''
    Read the EEPROM information from the device
    :param obj:
    :param s: query code
    :return: output:if s=0 then return all the information in list
        if s= others then return the SN in str
        example:['fs400_eep_arg_read 1 0\nadd:  00 | 01 | 02 | 03 | 04 | 05 | 06 | 07  ',
 '000: 0x30 0x39 0x01 0x30 0x31 0x32 0x33 0x34 ',
 '008: 0x35 0x36 0x37 0x38 0x39 0x30 0x32 0x34 ',
 '016: 0x36 0x38 0x30 0x0e 0x00 0xcf 0xe0 0x4a ',
 '024: 0x19 0x31 0x44 0x00 0x58 0x47 0x00 0x00 ',
 '032: 0x00 0x00 0xff 0xff 0xff 0xff 0xff 0xff ',
 '040: 0xff 0xff 0xff 0xff 0xff 0xff 0xff 0xff ',
 '048: 0xff 0xff 0xff 0xff 0xff 0xff 0xff 0xff ',
 '056: 0xff 0xff 0xff 0xff 0x7a 0x6b 0x69 0x11 ',
 'page_0 crc check ok']
    '''
    if serial.Serial.isOpen(obj.CtrlB):
        obj.CtrlB.write(b'slave_on\n')
        time.sleep(0.1)
        obj.CtrlB.write(b'cpld_spi_wr 0x31 1\n')
        time.sleep(0.1)
        obj.CtrlB.write(b'fs400_eep_arg_read 1 0\n')
        time.sleep(0.1)
        a=obj.CtrlB.read_until(b'check ok')
        time.sleep(0.1)
        obj.CtrlB.write(b'slave_off\n')
        obj.CtrlB.close()
        if s==0:
            return a.decode('utf-8').split('\r\n')
        else:
            b = a.decode('utf-8')
            b=b[b.index('\r\n000:'):]
            b = b.split('\r\n')
            sn= b[1].split(' ')[4:9]+b[2].split(' ')[1:6]
            return ascii_to_str(sn)
    else:
        print('Error, Control board port is not open')

def read_eeprom_PeVoa(obj):
    '''
    Read the EEPROM information from the device
    :param obj:
    :param s: query code
    :return: output:if s=0 then return all the information in list
        if s= others then return the SN in str
        example:['fs400_eep_arg_read 1 0\nadd:  00 | 01 | 02 | 03 | 04 | 05 | 06 | 07  ',
 '000: 0x30 0x39 0x01 0x30 0x31 0x32 0x33 0x34 ',
 '008: 0x35 0x36 0x37 0x38 0x39 0x30 0x32 0x34 ',
 '016: 0x36 0x38 0x30 0x0e 0x00 0xcf 0xe0 0x4a ',
 '024: 0x19 0x31 0x44 0x00 0x58 0x47 0x00 0x00 ',
 '032: 0x00 0x00 0xff 0xff 0xff 0xff 0xff 0xff ',
 '040: 0xff 0xff 0xff 0xff 0xff 0xff 0xff 0xff ',
 '048: 0xff 0xff 0xff 0xff 0xff 0xff 0xff 0xff ',
 '056: 0xff 0xff 0xff 0xff 0x7a 0x6b 0x69 0x11 ',
 'page_0 crc check ok']
    '''
    if serial.Serial.isOpen(obj.CtrlB):
        obj.CtrlB.timeout = 3
        sn = ''
        cmd=b'fs400_eep_arg_read 0\r\n'
        obj.CtrlB.write(cmd)
        try:
            s = obj.CtrlB.read_until(b'shell>').decode('utf-8').split('\r\n')
            obj.CtrlB.read_all()
            sn = [i.split(':')[1].strip() for i in s if 'sn:' in i][0]
        except Exception as e:
            print('SN can not get...\n', e)
            sn = ''
        return sn
    else:
        print('Error, Control board port is not open')

#update on 0604_2022
def write_eeprom(obj,df):#vbias,vpi ):
    '''
    Write the test data into the eeprom
    :param obj,df: to get vbias:[1,2,3,4,5,6] list,vpi:float
    :return: NA
    '''
    try:
        #Firstly get vbias and vpi from the dataframe result
        vbias=[]
        vpi_all=[]
        for t in range(8,14):
            vbias.append(df.iloc[:,t].median())
            vpi_all.append(df.iloc[:,t+12].median())
        vpi=min(vpi_all)
        if serial.Serial.isOpen(obj.CtrlB):
            #page 0
            # SN,Driver and TIA version
            #Driver config
            obj.CtrlB.flushInput()
            obj.CtrlB.flushOutput()
            obj.CtrlB.write(b'fs400_drv_read 0x0401\n');time.sleep(0.2)
            s=obj.CtrlB.read_all().decode('utf-8').split(' ')
            c1=s[3][-2:]+s[4][-2:]
            #judge whether the return value is 'e0cf' to avoid bug
            count=0
            while count<3:
                if not c1=='e0cf':
                    obj.CtrlB.write(b'fs400_drv_read 0x0401\n');time.sleep(0.2)
                    s=obj.CtrlB.read_all().decode('utf-8').split(' ')
                    c1=s[3][-2:]+s[4][-2:]
                    count+=1
                    time.sleep(0.2)
                    print('The {} time try:\n'.format(str(count)),s)
                else:
                    break
            obj.CtrlB.write(b'fs400_drv_read 0x0405\n');time.sleep(0.2)
            s=obj.CtrlB.read_all().decode('utf-8').split(' ')
            c2=s[3][-2:]+s[4][-2:]
            obj.CtrlB.write(b'fs400_drv_read 0x0409\n');time.sleep(0.2)
            s=obj.CtrlB.read_all().decode('utf-8').split(' ')
            c3=s[3][-2:]+s[4][-2:]
            #write in the drv information
            m='fs400_set_eep_drv_ver 0x0 0x{} 0x{} 0x{}\n'.format(c1,c2,c3).encode('utf-8')
            print('Driver information write:',m)
            obj.CtrlB.write(m);time.sleep(0.2)#b'fs400_set_eep_drv_ver 0x0 0xe0cf 0x194a 0x4431\n'
            #TIA config
            obj.CtrlB.write(b'fs400_tia_read 0x0001\n');time.sleep(0.2)
            s=obj.CtrlB.read_all().decode('utf-8').split(' ')
            a1=s[3][-2:]+s[4][-2:]
            #judge whether the return value is 'e0cf' to avoid bug
            count=0
            while count<3:
                if not a1=='e0cf':
                    obj.CtrlB.write(b'fs400_tia_read 0x0001\n');time.sleep(0.2)
                    s=obj.CtrlB.read_all().decode('utf-8').split(' ')
                    a1=s[3][-2:]+s[4][-2:]
                    count+=1
                    time.sleep(0.2)
                    print('The {} time try:\n'.format(str(count)),s)
                else:
                    break
            obj.CtrlB.write(b'fs400_tia_read 0x0005\n');time.sleep(0.2)
            s=obj.CtrlB.read_all().decode('utf-8').split(' ')
            a2=s[3][-2:]+s[4][-2:]
            obj.CtrlB.write(b'fs400_tia_read 0x0009\n');time.sleep(0.2)
            s=obj.CtrlB.read_all().decode('utf-8').split(' ')
            a3=s[3][-2:]+s[4][-2:]
            #Drv config
            m1='fs400_set_eep_tia_ver 0x0 0x{} 0x{} 0x{}\n'.format(a1,a2,a3).encode('utf-8')
            print('TIA information write:',m1)
            obj.CtrlB.write(m1);time.sleep(0.2)#b'fs400_set_eep_tia_ver 0x0 0xe0cf 0x4014 0x4303\n'
            #save config
            obj.CtrlB.write(b'fs400_eep_arg_save 1 0\n');time.sleep(0.2)
            print(obj.CtrlB.read_until(b'Save ok'))
            #page 1
            #write vbias and vpi,read pkd and write
            #vbias
            vcode=''
            for i in range(6):
                vcode+=str((round(vbias[i]/4095*3.3*10000)))+' '
            obj.CtrlB.write(b'fs400_set_eep_vbias '+vcode[:-1].encode('utf-8')+b'\n');time.sleep(0.5)
            print(obj.CtrlB.read_all())
            #vpi
            cyc=str(round(vpi*100))
            obj.CtrlB.write(b'fs400_set_eep_limit_cycle 25500 '+cyc.encode('utf-8')+b'\n');time.sleep(0.5)
            print(obj.CtrlB.read_all())
            #abc speed limit
            obj.CtrlB.write(b'fs400_set_eep_abc_th_value 650 300\n');time.sleep(0.5)
            print(obj.CtrlB.read_all())
            #save once
            obj.CtrlB.write(b'fs400_eep_arg_save 1 1\n');time.sleep(0.2)
            print(obj.CtrlB.read_until(b'Save ok'))
            #pkd read
            obj.CtrlB.write(b'fs400_drv_read 0x0427\n');time.sleep(0.2)
            print(obj.CtrlB.read_all())
            obj.CtrlB.write(b'fs400_drv_read 0x0427\n');time.sleep(0.2)
            s=obj.CtrlB.read_all().decode('utf-8').split(' ')
            b1=s[3][-2:]+s[4][-2:]
            obj.CtrlB.write(b'fs400_drv_read 0x0527\n');time.sleep(0.2)
            print(obj.CtrlB.read_all())
            obj.CtrlB.write(b'fs400_drv_read 0x0527\n');time.sleep(0.2)
            s=obj.CtrlB.read_all().decode('utf-8').split(' ')
            b2=s[3][-2:]+s[4][-2:]
            obj.CtrlB.write(b'fs400_drv_read 0x0627\n');time.sleep(0.2)
            print(obj.CtrlB.read_all())
            obj.CtrlB.write(b'fs400_drv_read 0x0627\n');time.sleep(0.2)
            s=obj.CtrlB.read_all().decode('utf-8').split(' ')
            b3=s[3][-2:]+s[4][-2:]
            obj.CtrlB.write(b'fs400_drv_read 0x0727\n');time.sleep(0.2)
            print(obj.CtrlB.read_all())
            obj.CtrlB.write(b'fs400_drv_read 0x0727\n');time.sleep(0.2)
            s=obj.CtrlB.read_all().decode('utf-8').split(' ')
            b4=s[3][-2:]+s[4][-2:]
            m2='fs400_set_eep_drv_pdk 0x{} 0x{} 0x{} 0x{}\n'.format(b1,b2,b3,b4).encode('utf-8')
            print('PKD information write:',m2)
            obj.CtrlB.write(m2);time.sleep(0.2)#b'fs400_set_eep_drv_pdk 0x1122 0x1122 0x1122 0x1122\n'
            #save config
            obj.CtrlB.write(b'fs400_eep_arg_save 1 1\n');time.sleep(0.2)
            print(obj.CtrlB.read_until(b'Save ok'))
            #page3
            obj.CtrlB.write(b'fs400_set_eep_400g_pre 0 127 -12 -6 1 4\n');time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_400g_pre 5 4 2 0 -1 -1\n');time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_400g_pre 10 0 0 0 0 0\n');time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_400g_pre 15 0 0 0 0 0\n');time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_400g_pre 20 0 0 0 0 0\n');time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_400g_pre 25 0 0 0 0 0\n');time.sleep(0.2)
            obj.CtrlB.write(b'fs400_eep_arg_save 1 3\n');time.sleep(0.2)
            print(obj.CtrlB.read_until(b'Save ok'))
            #print out the result of pages
            obj.CtrlB.write(b'fs400_eep_arg_read 1 0\n');time.sleep(0.2)
            print('Page 0:\n',obj.CtrlB.read_until(b'check ok'))
            obj.CtrlB.write(b'fs400_eep_arg_read 1 1\n');time.sleep(0.2)
            print('Page 1:\n',obj.CtrlB.read_until(b'check ok'))
            obj.CtrlB.write(b'fs400_eep_arg_read 1 3\n');time.sleep(0.2)
            print('Page 3:\n',obj.CtrlB.read_until(b'check ok'))
            print('Eeprom write OK')
        else:
            print('Error, Control board port is not open')
    except Exception as e:
        print(e,'Please check whether board or driver is up and configed!!!')

#for alu eeprom write in
#update on 0604_2022
def write_eeprom_alu(obj,df,flag='IDT'):#vbias,vpi ):
    '''
    Write the test data into the eeprom
    :param obj,df: to get vbias:[1,2,3,4,5,6] list,vpi:float
    :return: NA
    '''
    try:
        #Firstly get vbias and vpi from the dataframe result
        vbias=[]
        vpi_all=[]
        for t in range(8,14):
            vbias.append(df.iloc[:,t].median())
            vpi_all.append(df.iloc[:,t+12].median())
        vpi=min(vpi_all)
        if serial.Serial.isOpen(obj.CtrlB):
            #page 0
            # SN,Driver and TIA version
            #Driver config
            obj.CtrlB.flushInput()
            obj.CtrlB.flushOutput()
            #Read first to avoid erase the SN
            obj.CtrlB.write(b'fs400_eep_arg_read 1 0\n');
            time.sleep(0.2)
            print('Page 0:\n', obj.CtrlB.read_until(b'check ok'))

            if flag=='IDT':
                obj.CtrlB.write(b'fs400_drv_read 0x0401\n');time.sleep(0.2)
                s=obj.CtrlB.read_all().decode('utf-8').split(' ')
                c1=s[3][-2:]+s[4][-2:]
                #judge whether the return value is 'e0cf' to avoid bug
                count=0
                while count<3:
                    if not c1=='e0cf':
                        obj.CtrlB.write(b'fs400_drv_read 0x0401\n');time.sleep(0.2)
                        s=obj.CtrlB.read_all().decode('utf-8').split(' ')
                        c1=s[3][-2:]+s[4][-2:]
                        count+=1
                        time.sleep(0.2)
                        print('The {} time try:\n'.format(str(count)),s)
                    else:
                        break
                obj.CtrlB.write(b'fs400_drv_read 0x0405\n');time.sleep(0.2)
                s=obj.CtrlB.read_all().decode('utf-8').split(' ')
                c2=s[3][-2:]+s[4][-2:]
                obj.CtrlB.write(b'fs400_drv_read 0x0409\n');time.sleep(0.2)
                s=obj.CtrlB.read_all().decode('utf-8').split(' ')
                c3=s[3][-2:]+s[4][-2:]
                #write in the drv information
                m='fs400_set_eep_drv_ver 0x0 0x{} 0x{} 0x{}\n'.format(c1,c2,c3).encode('utf-8')
                print('Driver information write:',m)
                obj.CtrlB.write(m);time.sleep(0.2)#b'fs400_set_eep_drv_ver 0x0 0xe0cf 0x194a 0x4431\n'
                #TIA config
                obj.CtrlB.write(b'fs400_tia_read 0x0001\n');time.sleep(0.2)
                s=obj.CtrlB.read_all().decode('utf-8').split(' ')
                a1=s[3][-2:]+s[4][-2:]
                #judge whether the return value is 'e0cf' to avoid bug
                count=0
                while count<3:
                    if not a1=='e0cf':
                        obj.CtrlB.write(b'fs400_tia_read 0x0001\n');time.sleep(0.2)
                        s=obj.CtrlB.read_all().decode('utf-8').split(' ')
                        a1=s[3][-2:]+s[4][-2:]
                        count+=1
                        time.sleep(0.2)
                        print('The {} time try:\n'.format(str(count)),s)
                    else:
                        break
                obj.CtrlB.write(b'fs400_tia_read 0x0005\n');time.sleep(0.2)
                s=obj.CtrlB.read_all().decode('utf-8').split(' ')
                a2=s[3][-2:]+s[4][-2:]
                obj.CtrlB.write(b'fs400_tia_read 0x0009\n');time.sleep(0.2)
                s=obj.CtrlB.read_all().decode('utf-8').split(' ')
                a3=s[3][-2:]+s[4][-2:]
                #Drv config
                m1='fs400_set_eep_tia_ver 0x0 0x{} 0x{} 0x{}\n'.format(a1,a2,a3).encode('utf-8')
                print('TIA information write:',m1)
                obj.CtrlB.write(m1);time.sleep(0.2)#b'fs400_set_eep_tia_ver 0x0 0xe0cf 0x4014 0x4303\n'
                #save config
                obj.CtrlB.write(b'fs400_eep_arg_save 1 0\n');time.sleep(0.2)
                print(obj.CtrlB.read_until(b'Save ok'))
            elif flag=='ALU':
                print('ALU device information write in')
                obj.CtrlB.write(b'fs400_set_eep_drv_ver 0x1 0x0 0x0 0x0\n');time.sleep(0.5)
                obj.CtrlB.write(b'fs400_set_eep_tia_ver 0x1 0x0 0x0 0x0\n');time.sleep(0.5)
                obj.CtrlB.write(b'fs400_eep_arg_save 1 0\n');time.sleep(0.2)
                print(obj.CtrlB.read_until(b'Save ok'))
            else:
                print('Flag error')
                return

            #page 1
            #write vbias and vpi,read pkd and write
            #vbias
            vcode=''
            for i in range(6):
                vcode+=str((round(vbias[i]/4095*3.3*10000)))+' '
            obj.CtrlB.write(b'fs400_set_eep_vbias '+vcode[:-1].encode('utf-8')+b'\n');time.sleep(0.5)
            print(obj.CtrlB.read_all())
            #vpi
            cyc=str(round(vpi*100))
            obj.CtrlB.write(b'fs400_set_eep_limit_cycle 25500 '+cyc.encode('utf-8')+b'\n');time.sleep(0.5)
            print(obj.CtrlB.read_all())
            #abc speed limit
            obj.CtrlB.write(b'fs400_set_eep_abc_th_value 650 300\n');time.sleep(0.5)
            print(obj.CtrlB.read_all())
            #save once
            obj.CtrlB.write(b'fs400_eep_arg_save 1 1\n');time.sleep(0.2)
            print(obj.CtrlB.read_until(b'Save ok'))
            print(obj.CtrlB.read_all())

            if flag=='IDT':
                #pkd read
                obj.CtrlB.write(b'fs400_drv_read 0x0427\n');time.sleep(0.2)
                if not b'fs400_drv_read' in obj.CtrlB.read_all():
                    print(obj.CtrlB.read_all())
                obj.CtrlB.write(b'fs400_drv_read 0x0427\n');time.sleep(0.2)
                s=obj.CtrlB.read_all().decode('utf-8').split(' ')
                b1=s[3][-2:]+s[4][-2:]
                obj.CtrlB.write(b'fs400_drv_read 0x0527\n');time.sleep(0.2)
                print(obj.CtrlB.read_all())
                obj.CtrlB.write(b'fs400_drv_read 0x0527\n');time.sleep(0.2)
                s=obj.CtrlB.read_all().decode('utf-8').split(' ')
                b2=s[3][-2:]+s[4][-2:]
                obj.CtrlB.write(b'fs400_drv_read 0x0627\n');time.sleep(0.2)
                print(obj.CtrlB.read_all())
                obj.CtrlB.write(b'fs400_drv_read 0x0627\n');time.sleep(0.2)
                s=obj.CtrlB.read_all().decode('utf-8').split(' ')
                b3=s[3][-2:]+s[4][-2:]
                obj.CtrlB.write(b'fs400_drv_read 0x0727\n');time.sleep(0.2)
                print(obj.CtrlB.read_all())
                obj.CtrlB.write(b'fs400_drv_read 0x0727\n');time.sleep(0.2)
                s=obj.CtrlB.read_all().decode('utf-8').split(' ')
                b4=s[3][-2:]+s[4][-2:]
                m2='fs400_set_eep_drv_pdk 0x{} 0x{} 0x{} 0x{}\n'.format(b1,b2,b3,b4).encode('utf-8')
                print('PKD information write:',m2)
                obj.CtrlB.write(m2);time.sleep(0.2)#b'fs400_set_eep_drv_pdk 0x1122 0x1122 0x1122 0x1122\n'
                #save config
                obj.CtrlB.write(b'fs400_eep_arg_save 1 1\n');time.sleep(0.2)
                print(obj.CtrlB.read_until(b'Save ok'))

            #page2
            obj.CtrlB.write(b'fs400_set_eep_voainfo 0x3b 0x0 -170 4310\n');time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_voax 0x0 0x0 0x0 0x0\n');time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_voay 0x0 0x0 0x0 0x0\n');time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_sensor_k_b 0x0 0x0\n');time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_tia_k_b 0x0 0x0\n');time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_drv_k_b 0x0 0x0\n');time.sleep(0.2)
            obj.CtrlB.write(b'fs400_eep_arg_save 1 2\n');time.sleep(0.2)
            print(obj.CtrlB.read_until(b'Save ok'))

            #page3
            obj.CtrlB.write(b'fs400_set_eep_400g_pre 0 127 -12 -6 1 4\n');time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_400g_pre 5 4 2 0 -1 -1\n');time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_400g_pre 10 0 0 0 0 0\n');time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_400g_pre 15 0 0 0 0 0\n');time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_400g_pre 20 0 0 0 0 0\n');time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_400g_pre 25 0 0 0 0 0\n');time.sleep(0.2)
            obj.CtrlB.write(b'fs400_eep_arg_save 1 3\n');time.sleep(0.2)
            print(obj.CtrlB.read_until(b'Save ok'))
            #page7
            if flag=='IDT':
                obj.CtrlB.write(b'fs400_set_eep_dev_drv1 0 0x0 0x01 0x0 0x50 0x50 0x50 0x50\n');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_drv2 0 0x0601 0xac 0x0e44 0x7771 0x3727\n');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_drv3 0 0x5111 0xff 0x1770 0x12c0\n');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_tia1 0 0x1e 0x03 0x000a 0xaf88\n');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_tia2 0 0xffff 0x60 0x18 2500\n');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_eep_arg_save 1 7\n');time.sleep(0.2)
                print(obj.CtrlB.read_until(b'Save ok'))
            elif flag=='ALU':
                obj.CtrlB.write(b'fs400_set_eep_dev_drv1 0 0x0 0x0 800 2000 2000 2000 2000\n');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_drv2 0 0xffff 0xff 0xffff 0xffff 0xffff\n');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_drv3 0 0xffff 0x01 0x1770 0x12c0\n');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_tia1 0 0x1e 0x01 0xffff 0xffff\n');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_tia2 0 0xffff 2000 1500 2500\n');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_eep_arg_save 1 7\n');time.sleep(0.2)
                print(obj.CtrlB.read_until(b'Save ok'))
            else:
                print('Flag error')
                return

            #print out the result of pages
            obj.CtrlB.write(b'fs400_eep_arg_read 1 0\n');time.sleep(0.2)
            print('Page 0:\n',obj.CtrlB.read_until(b'check ok'))
            obj.CtrlB.write(b'fs400_eep_arg_read 1 1\n');time.sleep(0.2)
            print('Page 1:\n',obj.CtrlB.read_until(b'check ok'))
            obj.CtrlB.write(b'fs400_eep_arg_read 1 2\n');time.sleep(0.2)
            print('Page 2:\n',obj.CtrlB.read_until(b'check ok'))
            obj.CtrlB.write(b'fs400_eep_arg_read 1 3\n');time.sleep(0.2)
            print('Page 3:\n',obj.CtrlB.read_until(b'check ok'))
            obj.CtrlB.write(b'fs400_eep_arg_read 1 7\n');time.sleep(0.2)
            print('Page 7:\n',obj.CtrlB.read_until(b'check ok'))
            print('Eeprom write OK')
        else:
            print('Error, Control board port is not open')
    except Exception as e:
        print(e,'Please check whether board or driver is up and configed!!!')

def write_eeprom_final(obj,sn,df,flag='IDT',Drv3='0x12fc'):
    '''
    Write the test data into the eeprom
    Update on 20220908:Write all the pages including all VOA calibration data, and add read and check,
    Attention:Always write all we need before save, otherwise don`t save, or will erase the data.
    :param obj,df: to get vbias:[1,2,3,4,5,6] list,vpi:float
    :return: NA
    '''
    try:
        #Firstly get vbias and vpi from the dataframe result
        vbias=[]
        vpi_all=[]
        for t in range(8,14):
            vbias.append(df.iloc[:,t].median())
            vpi_all.append(df.iloc[:,t+12].median())
        vpi=min(vpi_all)
        if serial.Serial.isOpen(obj.CtrlB):
            #***----------page 0----------***
            # SN,Driver and TIA version
            #Driver config
            obj.CtrlB.flushInput()
            obj.CtrlB.flushOutput()
            print('Start write Page 0:')
            #Read first to avoid erase the SN
            obj.CtrlB.write(b'fs400_eep_arg_read 1 0\n\r');
            time.sleep(0.2)
            print(obj.CtrlB.read_until(b'check ok'))
            #Need to power on Driver and TIA first
            obj.CtrlB.write(b'fs400_set_eep_device 0x46 0x53\n\r');
            time.sleep(0.2)
            s = obj.CtrlB.read_all().decode('utf-8')
            print(s)
            obj.CtrlB.write(b'fs400_set_eep_ver 0x03\n\r');
            time.sleep(0.2)
            s = obj.CtrlB.read_all().decode('utf-8')
            print(s)
            #write SN
            print('Start write SN in page 0:')
            sn = str(sn)
            if len(sn) != 10:
                print('SN input error lenth is {}, not 10!!!'.format(len(sn)))
                obj.CtrlB.close()
                return
            ascii_sn = str_to_ascii(sn)
            b = str_to_ascii(sn)
            cod1 = b[0] + b[1][2:] + b[2][2:] + b[3][2:]
            cod2 = b[4] + b[5][2:] + b[6][2:] + b[7][2:]
            cod3 = b[8] + b[9][2:] + '20' + '20'
            cod4 = '0x20202020'
            obj.CtrlB.write('fs400_set_eep_sn {} {} {} {}\n\r'.format(cod1, cod2, cod3, cod4).encode('utf-8'))
            time.sleep(0.1)
            #read and write TIA Driver information
            if flag=='IDT':
                obj.CtrlB.write(b'fs400_drv_read 0x0401\n\r');time.sleep(0.2)
                s=obj.CtrlB.read_all().decode('utf-8').split(' ')
                c1=s[3][-2:]+s[4][-2:]
                #judge whether the return value is 'e0cf' to avoid bug
                count=0
                while count<3:
                    if not c1=='e0cf':
                        obj.CtrlB.write(b'fs400_drv_read 0x0401\n\r');time.sleep(0.2)
                        s=obj.CtrlB.read_all().decode('utf-8').split(' ')
                        c1=s[3][-2:]+s[4][-2:]
                        count+=1
                        time.sleep(0.2)
                        print('The {} time try:\n\r'.format(str(count)),s)
                    else:
                        break
                obj.CtrlB.write(b'fs400_drv_read 0x0405\n\r');time.sleep(0.2)
                s=obj.CtrlB.read_all().decode('utf-8').split(' ')
                c2=s[3][-2:]+s[4][-2:]
                obj.CtrlB.write(b'fs400_drv_read 0x0409\n\r');time.sleep(0.2)
                s=obj.CtrlB.read_all().decode('utf-8').split(' ')
                c3=s[3][-2:]+s[4][-2:]
                #write in the drv information
                m='fs400_set_eep_drv_ver 0x0 0x{} 0x{} 0x{}\n\r'.format(c1,c2,c3).encode('utf-8')
                print('Driver information write:',m)
                obj.CtrlB.write(m);time.sleep(0.2)#b'fs400_set_eep_drv_ver 0x0 0xe0cf 0x194a 0x4431\n'
                #TIA config
                obj.CtrlB.write(b'fs400_tia_read 0x0001\n\r');time.sleep(0.2)
                s=obj.CtrlB.read_all().decode('utf-8').split(' ')
                a1=s[3][-2:]+s[4][-2:]
                #judge whether the return value is 'e0cf' to avoid bug
                count=0
                while count<3:
                    if not a1=='e0cf':
                        obj.CtrlB.write(b'fs400_tia_read 0x0001\n\r');time.sleep(0.2)
                        s=obj.CtrlB.read_all().decode('utf-8').split(' ')
                        a1=s[3][-2:]+s[4][-2:]
                        count+=1
                        time.sleep(0.2)
                        print('The {} time try:\n\r'.format(str(count)),s)
                    else:
                        break
                obj.CtrlB.write(b'fs400_tia_read 0x0005\n\r');time.sleep(0.2)
                s=obj.CtrlB.read_all().decode('utf-8').split(' ')
                a2=s[3][-2:]+s[4][-2:]
                obj.CtrlB.write(b'fs400_tia_read 0x0009\n\r');time.sleep(0.2)
                s=obj.CtrlB.read_all().decode('utf-8').split(' ')
                a3=s[3][-2:]+s[4][-2:]
                #Drv config
                m1='fs400_set_eep_tia_ver 0x0 0x{} 0x{} 0x{}\n\r'.format(a1,a2,a3).encode('utf-8')
                print('TIA information write:',m1)
                obj.CtrlB.write(m1);time.sleep(0.2)#b'fs400_set_eep_tia_ver 0x0 0xe0cf 0x4014 0x4303\n'
            elif flag=='ALU':
                print('ALU device information write in')
                obj.CtrlB.write(b'fs400_set_eep_drv_ver 0x1 0x0 0x0 0x0\n\r');time.sleep(0.5)
                obj.CtrlB.write(b'fs400_set_eep_tia_ver 0x1 0x0 0x0 0x0\n\r');time.sleep(0.5)
            else:
                print('Flag error')
                return
            # save config
            obj.CtrlB.write(b'fs400_eep_arg_save 1 0\n\r');
            time.sleep(0.2)
            print(obj.CtrlB.read_until(b'Save ok'))
            time.sleep(1)

            #***----------page 1----------***
            #write vbias and vpi,read pkd and write
            #vbias
            print('Start write Page 1:')
            vcode=''
            for i in range(6):
                vcode+=str((round(vbias[i]/4095*3.3*10000)))+' '
            obj.CtrlB.write(b'fs400_set_eep_vbias '+vcode[:-1].encode('utf-8')+b'\n\r');time.sleep(0.5)
            print(obj.CtrlB.read_all())
            #vpi
            cyc=str(round(vpi*100))
            obj.CtrlB.write(b'fs400_set_eep_limit_cycle 25500 '+cyc.encode('utf-8')+b'\n\r');time.sleep(0.5)
            print(obj.CtrlB.read_all())
            #abc speed limit
            obj.CtrlB.write(b'fs400_set_eep_abc_th_value 650 300\n\r');time.sleep(0.5)
            print(obj.CtrlB.read_all())

            if flag=='IDT':
                #pkd read
                obj.CtrlB.write(b'fs400_drv_read 0x0427\n\r');time.sleep(0.2)
                if not b'fs400_drv_read' in obj.CtrlB.read_all():
                    print(obj.CtrlB.read_all())
                obj.CtrlB.write(b'fs400_drv_read 0x0427\n\r');time.sleep(0.2)
                s=obj.CtrlB.read_all().decode('utf-8').split(' ')
                b1=s[3][-2:]+s[4][-2:]
                obj.CtrlB.write(b'fs400_drv_read 0x0527\n\r');time.sleep(0.2)
                print(obj.CtrlB.read_all())
                obj.CtrlB.write(b'fs400_drv_read 0x0527\n\r');time.sleep(0.2)
                s=obj.CtrlB.read_all().decode('utf-8').split(' ')
                b2=s[3][-2:]+s[4][-2:]
                obj.CtrlB.write(b'fs400_drv_read 0x0627\n\r');time.sleep(0.2)
                print(obj.CtrlB.read_all())
                obj.CtrlB.write(b'fs400_drv_read 0x0627\n\r');time.sleep(0.2)
                s=obj.CtrlB.read_all().decode('utf-8').split(' ')
                b3=s[3][-2:]+s[4][-2:]
                obj.CtrlB.write(b'fs400_drv_read 0x0727\n\r');time.sleep(0.2)
                print(obj.CtrlB.read_all())
                obj.CtrlB.write(b'fs400_drv_read 0x0727\n\r');time.sleep(0.2)
                s=obj.CtrlB.read_all().decode('utf-8').split(' ')
                b4=s[3][-2:]+s[4][-2:]
                m2='fs400_set_eep_drv_pdk 0x{} 0x{} 0x{} 0x{}\n\r'.format(b1,b2,b3,b4).encode('utf-8')
                print('PKD information write:\n\r',m2)
                obj.CtrlB.write(m2);time.sleep(0.2)#b'fs400_set_eep_drv_pdk 0x1122 0x1122 0x1122 0x1122\n'
                obj.CtrlB.write(b'fs400_set_eep_tia_pdk 0 0 0 0\n\r');time.sleep(0.2)
            else:
                obj.CtrlB.write(b'fs400_set_eep_drv_pdk 0 0 0 0\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_tia_pdk 0 0 0 0\n\r');time.sleep(0.2)
            #save config
            obj.CtrlB.write(b'fs400_eep_arg_save 1 1\n\r');time.sleep(0.2)
            print(obj.CtrlB.read_until(b'Save ok'))
            print(obj.CtrlB.read_all(), 'page 1 saved!')
            time.sleep(1)

            #***----------page 2----------***
            print('Start write Page 2:')
            obj.CtrlB.write(b'fs400_set_eep_voainfo 0x3b 0x1 -170 4230\n\r');time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_voax 0x0 0x0 0x0 0x0\n\r');time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_voay 0x0 0x0 0x0 0x0\n\r');time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_sensor_k_b 0x0 0x0\n\r');time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_tia_k_b 0x0 0x0\n\r');time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_drv_k_b 0x0 0x0\n\r');time.sleep(0.2)
            obj.CtrlB.write(b'fs400_eep_arg_save 1 2\n\r');time.sleep(0.2)
            print(obj.CtrlB.read_until(b'Save ok'))
            print(obj.CtrlB.read_all(), 'page 2 saved!')
            time.sleep(1)

            #***----------page 3----------***
            print('Start write Page 3:')
            obj.CtrlB.write(b'fs400_set_eep_400g_pre 0 127 9 1 -5 -7\n\r');time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_400g_pre 5 -4 1 3 2 -1\n\r');time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_400g_pre 10 -3 -2 1 2 1\n\r');time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_400g_pre 15 0 -1 -1 0 1\n\r');time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_400g_pre 20 1 0 -1 -1 0\n\r');time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_400g_pre 25 0 0 0 0 0\n\r');time.sleep(0.2)
            obj.CtrlB.write(b'fs400_eep_arg_save 1 3\n\r');time.sleep(0.2)
            print(obj.CtrlB.read_until(b'Save ok'))
            print(obj.CtrlB.read_all(), 'page 3 saved!')
            time.sleep(1)

            #***----------page 4----------***
            print('Start write Page 4:')
            obj.CtrlB.write(b'fs400_set_eep_200g_pre 0 127 -5 -15 -4 1\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_200g_pre 5 3 3 1 -1 -1\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_200g_pre 10 -1 0 1 1 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_200g_pre 15 0 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_200g_pre 20 0 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_200g_pre 25 0 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_eep_arg_save 1 4\n\r');
            time.sleep(0.2)
            print(obj.CtrlB.read_until(b'Save ok'))
            print(obj.CtrlB.read_all(), 'page 4 saved!')
            time.sleep(1)

            #***----------page 5----------***
            print('Start write Page 5:')
            obj.CtrlB.write(b'fs400_set_eep_200gps_pre 0 127 -10 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_200gps_pre 5 0 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_200gps_pre 10 0 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_200gps_pre 15 0 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_200gps_pre 20 0 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_200gps_pre 25 0 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_eep_arg_save 1 5\n\r');
            time.sleep(0.2)
            print(obj.CtrlB.read_until(b'Save ok'))
            print(obj.CtrlB.read_all(), 'page 5 saved!')
            time.sleep(1)

            #***----------page 7----------***
            print('Start write Page 7:')
            if flag=='IDT':
                obj.CtrlB.write(b'fs400_set_eep_dev_drv1 0 0x0 0x01 0x0 0x740074 0x740074\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_drv2 0 0x0601 0xac 0x0e44 0x7771 0x3727\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_drv3 0 0x5111 0x01 0x1770 '+(Drv3).encode('utf-8')+b'\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_tia1 0 0x03 0x000a 0xaf88\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_tia2 0 0x60 0xffff 0x18 2500\n\r');time.sleep(0.2)
            elif flag=='ALU':
                obj.CtrlB.write(b'fs400_set_eep_dev_drv1 0 0x0 0x0 0 0x7d007d0 0x7d007d0\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_drv2 0 0xffff 0xff 0xffff 0xffff 0xffff\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_drv3 0 0xffff 0x01 0x1f40 0x12fc\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_tia1 0 0x01 0xffff 0xffff\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_tia2 0 2000 0xffff 1000 2500\n\r');time.sleep(0.2)
            else:
                print('Device Flag error')
                return
            obj.CtrlB.write(b'fs400_eep_arg_save 1 7\n\r');time.sleep(0.2)
            print(obj.CtrlB.read_until(b'Save ok'))
            print(obj.CtrlB.read_all(), 'page 7 saved!')
            time.sleep(1)

            # ***----------page 8----------***
            print('Start write Page 8:')
            if flag=='IDT':
                obj.CtrlB.write(b'fs400_set_eep_dev_drv1 1 0x0 0x01 0x0 0x580058 0x580058\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_drv2 1 0x0601 0xac 0x0e44 0x7771 0x3727\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_drv3 1 0x5111 0x01 0x1770 '+(Drv3).encode('utf-8')+b'\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_tia1 1 0x03 0x000a 0xaf88\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_tia2 1 0x60 0xffff 0x18 2500\n\r');time.sleep(0.2)
            elif flag=='ALU':
                obj.CtrlB.write(b'fs400_set_eep_dev_drv1 1 0x0 0x0 0 0x7d007d0 0x7d007d0\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_drv2 1 0xffff 0xff 0xffff 0xffff 0xffff\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_drv3 1 0xffff 0x01 0x1f40 0x12fc\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_tia1 1 0x01 0xffff 0xffff\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_tia2 1 2000 0xffff 1500 2500\n\r');time.sleep(0.2)
            else:
                print('Device Flag error')
                return
            obj.CtrlB.write(b'fs400_eep_arg_save 1 8\n\r');
            time.sleep(0.2)
            print(obj.CtrlB.read_until(b'Save ok'))
            print(obj.CtrlB.read_all(), 'page 8 saved!')
            time.sleep(1)

            # ***----------page 9----------***
            print('Start write Page 9:')
            if flag=='IDT':
                obj.CtrlB.write(b'fs400_set_eep_dev_drv1 2 0x0 0x01 0x0 0xc000c0 0xc000c0\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_drv2 2 0x0601 0xac 0x0e44 0x7771 0x3727\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_drv3 2 0x5111 0x01 0x1770 '+(Drv3).encode('utf-8')+b'\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_tia1 2 0x03 0x000a 0xaf88\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_tia2 2 0x60 0xffff 0x08 2500\n\r');time.sleep(0.2)
            elif flag=='ALU':
                obj.CtrlB.write(b'fs400_set_eep_dev_drv1 2 0x0 0x0 0x3e8 0x7d007d0 0x7d007d0\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_drv2 2 0xffff 0xff 0xffff 0xffff 0xffff\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_drv3 2 0xffff 0x01 0x1f40 0x12fc\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_tia1 2 0x01 0xffff 0xffff\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_tia2 2 2000 0xffff 1500 2500\n\r');time.sleep(0.2)
            else:
                print('Device Flag error')
                return
            obj.CtrlB.write(b'fs400_eep_arg_save 1 9\n\r');
            time.sleep(0.2)
            print(obj.CtrlB.read_until(b'Save ok'))
            print(obj.CtrlB.read_all(), 'page 9 saved!')
            time.sleep(1)

            # ***----------page 11----------***
            print('Start write Page 11:')
            if flag=='IDT':
                obj.CtrlB.write(b'fs400_set_eep_dev_drv1 4 0x0 0x01 0x0 0x900090 0x900090\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_drv2 4 0x0601 0xac 0x0e44 0x7771 0x3727\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_drv3 4 0x5111 0x01 0x1770 '+(Drv3).encode('utf-8')+b'\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_tia1 4 0x03 0x000a 0xaf88\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_tia2 4 0x60 0xffff 0x00 2500\n\r');time.sleep(0.2)
            elif flag=='ALU':
                pass
                # obj.CtrlB.write(b'fs400_set_eep_dev_drv1 4 0x0 0x0 0x320 0x7d007d0 0x7d007d0\n\r');time.sleep(0.2)
                # obj.CtrlB.write(b'fs400_set_eep_dev_drv2 4 0xffff 0xff 0xffff 0xffff 0xffff\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_drv3 4 0xffff 0x01 0x0 0x0\n\r');time.sleep(0.2)
                # obj.CtrlB.write(b'fs400_set_eep_dev_tia1 4 0x1e 0x01 0xffff 0xffff\n\r');time.sleep(0.2)
                # obj.CtrlB.write(b'fs400_set_eep_dev_tia2 4 2000 0xffff 1500 2500\n\r');time.sleep(0.2)
            else:
                print('Device Flag error')
                return
            obj.CtrlB.write(b'fs400_eep_arg_save 1 11\n\r');
            time.sleep(0.2)
            print(obj.CtrlB.read_until(b'Save ok'))
            print(obj.CtrlB.read_all(), 'page 11 saved!')
            time.sleep(1)

            # ***----------page 12----------***
            print('Start write Page 12:')
            obj.CtrlB.write(b'fs400_set_eep_pre_12 0 127 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_pre_12 5 0 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_pre_12 10 0 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_pre_12 15 0 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_pre_12 20 0 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_pre_12 25 0 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_eep_arg_save 1 12\n\r');
            time.sleep(0.2)
            print(obj.CtrlB.read_until(b'Save ok'))
            print(obj.CtrlB.read_all(), 'page 12 saved!')
            time.sleep(1)

            # ***----------page 13----------***
            print('Start write Page 13:')
            if flag=='IDT':
                obj.CtrlB.write(b'fs400_set_eep_dev_drv1 5 0x0 0x01 0x0 0x900090 0x900090\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_drv2 5 0x0601 0xac 0x0e44 0x7771 0x3727\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_drv3 5 0x5111 0x01 0x1770 '+(Drv3).encode('utf-8')+b'\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_tia1 5 0x03 0x000a 0xaf88\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_tia2 5 0x60 0xffff 0x00 2500\n\r');time.sleep(0.2)
            elif flag=='ALU':
                pass
                # obj.CtrlB.write(b'fs400_set_eep_dev_drv1 5 0x0 0x0 0x320 0x7d007d0 0x7d007d0\n\r');time.sleep(0.2)
                # obj.CtrlB.write(b'fs400_set_eep_dev_drv2 5 0xffff 0xff 0xffff 0xffff 0xffff\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_drv3 5 0xffff 0x01 0x0 0x0\n\r');time.sleep(0.2)
                # obj.CtrlB.write(b'fs400_set_eep_dev_tia1 5 0x1e 0x01 0xffff 0xffff\n\r');time.sleep(0.2)
                # obj.CtrlB.write(b'fs400_set_eep_dev_tia2 5 2000 0xffff 1500 2500\n\r');time.sleep(0.2)
            else:
                print('Device Flag error')
                return
            obj.CtrlB.write(b'fs400_eep_arg_save 1 13\n\r');
            time.sleep(0.2)
            print(obj.CtrlB.read_until(b'Save ok'))
            print(obj.CtrlB.read_all(), 'page 13 saved!')
            time.sleep(1)

            # ***----------page 14----------***
            print('Start write Page 14:')
            obj.CtrlB.write(b'fs400_set_eep_pre_14 0 99 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_pre_14 5 0 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_pre_14 10 0 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_pre_14 15 0 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_pre_14 20 0 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_pre_14 25 0 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_eep_arg_save 1 14\n\r');
            time.sleep(0.2)
            print(obj.CtrlB.read_until(b'Save ok'))
            print(obj.CtrlB.read_all(), 'page 14 saved!')
            time.sleep(1)

            # ***----------page 15----------***
            print('Start write Page 15:')
            if flag=='IDT':
                obj.CtrlB.write(b'fs400_set_eep_dev_drv1 6 0x0 0x01 0x0 0x700070 0x700070\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_drv2 6 0x0601 0xac 0x0e44 0x7771 0x3727\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_drv3 6 0x5111 0x01 0x1770 '+(Drv3).encode('utf-8')+b'\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_tia1 6 0x03 0x000a 0xaf88\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_tia2 6 0x60 0xffff 0x00 2500\n\r');time.sleep(0.2)
            elif flag=='ALU':
                pass
                # obj.CtrlB.write(b'fs400_set_eep_dev_drv1 6 0x0 0x0 0x320 0x7d007d0 0x7d007d0\n\r');time.sleep(0.2)
                # obj.CtrlB.write(b'fs400_set_eep_dev_drv2 6 0xffff 0xff 0xffff 0xffff 0xffff\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_drv3 6 0xffff 0x01 0x0 0x0\n\r');time.sleep(0.2)
                # obj.CtrlB.write(b'fs400_set_eep_dev_tia1 6 0x1e 0x01 0xffff 0xffff\n\r');time.sleep(0.2)
                # obj.CtrlB.write(b'fs400_set_eep_dev_tia2 6 2000 0xffff 1500 2500\n\r');time.sleep(0.2)
            else:
                print('Device Flag error')
                return
            obj.CtrlB.write(b'fs400_eep_arg_save 1 15\n\r');
            time.sleep(0.2)
            print(obj.CtrlB.read_until(b'Save ok'))
            print(obj.CtrlB.read_all(), 'page 15 saved!')
            time.sleep(1)

            # ***----------page 16----------***
            print('Start write Page 16:')
            obj.CtrlB.write(b'fs400_set_eep_pre_16 0 127 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_pre_16 5 0 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_pre_16 10 0 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_pre_16 15 0 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_pre_16 20 0 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_pre_16 25 0 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_eep_arg_save 1 16\n\r');
            time.sleep(0.2)
            print(obj.CtrlB.read_until(b'Save ok'))
            print(obj.CtrlB.read_all(), 'page 16 saved!')

            # #print out the result of pages
            # obj.CtrlB.write(b'fs400_eep_arg_read 1 0\n');time.sleep(0.2)
            # print('Page 0:\n',obj.CtrlB.read_until(b'check ok'))
            # obj.CtrlB.write(b'fs400_eep_arg_read 1 1\n');time.sleep(0.2)
            # print('Page 1:\n',obj.CtrlB.read_until(b'check ok'))
            # obj.CtrlB.write(b'fs400_eep_arg_read 1 2\n');time.sleep(0.2)
            # print('Page 2:\n',obj.CtrlB.read_until(b'check ok'))
            # obj.CtrlB.write(b'fs400_eep_arg_read 1 3\n');time.sleep(0.2)
            # print('Page 3:\n',obj.CtrlB.read_until(b'check ok'))
            # obj.CtrlB.write(b'fs400_eep_arg_read 1 7\n');time.sleep(0.2)
            # print('Page 7:\n',obj.CtrlB.read_until(b'check ok'))
            print('Eeprom write OK')
        else:
            print('Error, Control board port is not open')
    except Exception as e:
        print(e,'Please check whether board or driver is up and configed!!!')



def write_eeprom_final_old(obj,sn,df,flag='IDT',Drv3='0x12fc'):
    '''
    Write the test data into the eeprom
    Update on 20220908:Write all the pages including all VOA calibration data, and add read and check,
    Attention:Always write all we need before save, otherwise don`t save, or will erase the data.
    :param obj,df: to get vbias:[1,2,3,4,5,6] list,vpi:float
    :return: NA
    '''
    try:
        #Firstly get vbias and vpi from the dataframe result
        vbias=[]
        vpi_all=[]
        for t in range(8,14):
            vbias.append(df.iloc[:,t].median())
            vpi_all.append(df.iloc[:,t+12].median())
        vpi=min(vpi_all)
        if serial.Serial.isOpen(obj.CtrlB):
            #***----------page 0----------***
            # SN,Driver and TIA version
            #Driver config
            obj.CtrlB.flushInput()
            obj.CtrlB.flushOutput()
            print('Start write Page 0:')
            #Read first to avoid erase the SN
            obj.CtrlB.write(b'fs400_eep_arg_read 1 0\n\r');
            time.sleep(0.2)
            print(obj.CtrlB.read_until(b'check ok'))
            #Need to power on Driver and TIA first
            obj.CtrlB.write(b'fs400_set_eep_device 0x46 0x53\n\r');
            time.sleep(0.2)
            s = obj.CtrlB.read_all().decode('utf-8')
            print(s)
            obj.CtrlB.write(b'fs400_set_eep_ver 0x03\n\r');
            time.sleep(0.2)
            s = obj.CtrlB.read_all().decode('utf-8')
            print(s)
            #write SN
            print('Start write SN in page 0:')
            sn = str(sn)
            if len(sn) != 10:
                print('SN input error lenth is {}, not 10!!!'.format(len(sn)))
                obj.CtrlB.close()
                return
            ascii_sn = str_to_ascii(sn)
            b = str_to_ascii(sn)
            cod1 = b[0] + b[1][2:] + b[2][2:] + b[3][2:]
            cod2 = b[4] + b[5][2:] + b[6][2:] + b[7][2:]
            cod3 = b[8] + b[9][2:] + '20' + '20'
            cod4 = '0x20202020'
            obj.CtrlB.write('fs400_set_eep_sn {} {} {} {}\n\r'.format(cod1, cod2, cod3, cod4).encode('utf-8'))
            time.sleep(0.1)
            #read and write TIA Driver information
            if flag=='IDT':
                obj.CtrlB.write(b'fs400_drv_read 0x0401\n\r');time.sleep(0.2)
                s=obj.CtrlB.read_all().decode('utf-8').split(' ')
                c1=s[3][-2:]+s[4][-2:]
                #judge whether the return value is 'e0cf' to avoid bug
                count=0
                while count<3:
                    if not c1=='e0cf':
                        obj.CtrlB.write(b'fs400_drv_read 0x0401\n\r');time.sleep(0.2)
                        s=obj.CtrlB.read_all().decode('utf-8').split(' ')
                        c1=s[3][-2:]+s[4][-2:]
                        count+=1
                        time.sleep(0.2)
                        print('The {} time try:\n\r'.format(str(count)),s)
                    else:
                        break
                obj.CtrlB.write(b'fs400_drv_read 0x0405\n\r');time.sleep(0.2)
                s=obj.CtrlB.read_all().decode('utf-8').split(' ')
                c2=s[3][-2:]+s[4][-2:]
                obj.CtrlB.write(b'fs400_drv_read 0x0409\n\r');time.sleep(0.2)
                s=obj.CtrlB.read_all().decode('utf-8').split(' ')
                c3=s[3][-2:]+s[4][-2:]
                #write in the drv information
                m='fs400_set_eep_drv_ver 0x0 0x{} 0x{} 0x{}\n\r'.format(c1,c2,c3).encode('utf-8')
                print('Driver information write:',m)
                obj.CtrlB.write(m);time.sleep(0.2)#b'fs400_set_eep_drv_ver 0x0 0xe0cf 0x194a 0x4431\n'
                #TIA config
                obj.CtrlB.write(b'fs400_tia_read 0x0001\n\r');time.sleep(0.2)
                s=obj.CtrlB.read_all().decode('utf-8').split(' ')
                a1=s[3][-2:]+s[4][-2:]
                #judge whether the return value is 'e0cf' to avoid bug
                count=0
                while count<3:
                    if not a1=='e0cf':
                        obj.CtrlB.write(b'fs400_tia_read 0x0001\n\r');time.sleep(0.2)
                        s=obj.CtrlB.read_all().decode('utf-8').split(' ')
                        a1=s[3][-2:]+s[4][-2:]
                        count+=1
                        time.sleep(0.2)
                        print('The {} time try:\n\r'.format(str(count)),s)
                    else:
                        break
                obj.CtrlB.write(b'fs400_tia_read 0x0005\n\r');time.sleep(0.2)
                s=obj.CtrlB.read_all().decode('utf-8').split(' ')
                a2=s[3][-2:]+s[4][-2:]
                obj.CtrlB.write(b'fs400_tia_read 0x0009\n\r');time.sleep(0.2)
                s=obj.CtrlB.read_all().decode('utf-8').split(' ')
                a3=s[3][-2:]+s[4][-2:]
                #Drv config
                m1='fs400_set_eep_tia_ver 0x0 0x{} 0x{} 0x{}\n\r'.format(a1,a2,a3).encode('utf-8')
                print('TIA information write:',m1)
                obj.CtrlB.write(m1);time.sleep(0.2)#b'fs400_set_eep_tia_ver 0x0 0xe0cf 0x4014 0x4303\n'
            elif flag=='ALU':
                print('ALU device information write in')
                obj.CtrlB.write(b'fs400_set_eep_drv_ver 0x1 0x0 0x0 0x0\n\r');time.sleep(0.5)
                obj.CtrlB.write(b'fs400_set_eep_tia_ver 0x1 0x0 0x0 0x0\n\r');time.sleep(0.5)
            else:
                print('Flag error')
                return
            # save config
            obj.CtrlB.write(b'fs400_eep_arg_save 1 0\n\r');
            time.sleep(0.2)
            print(obj.CtrlB.read_until(b'Save ok'))

            #***----------page 1----------***
            #write vbias and vpi,read pkd and write
            #vbias
            print('Start write Page 1:')
            vcode=''
            for i in range(6):
                vcode+=str((round(vbias[i]/4095*3.3*10000)))+' '
            obj.CtrlB.write(b'fs400_set_eep_vbias '+vcode[:-1].encode('utf-8')+b'\n\r');time.sleep(0.5)
            print(obj.CtrlB.read_all())
            #vpi
            cyc=str(round(vpi*100))
            obj.CtrlB.write(b'fs400_set_eep_limit_cycle 25500 '+cyc.encode('utf-8')+b'\n\r');time.sleep(0.5)
            print(obj.CtrlB.read_all())
            #abc speed limit
            obj.CtrlB.write(b'fs400_set_eep_abc_th_value 650 300\n\r');time.sleep(0.5)
            print(obj.CtrlB.read_all())
            #save once
            # obj.CtrlB.write(b'fs400_eep_arg_save 1 1\n');time.sleep(0.2)
            # print(obj.CtrlB.read_until(b'Save ok'))

            if flag=='IDT':
                #pkd read
                obj.CtrlB.write(b'fs400_drv_read 0x0427\n\r');time.sleep(0.2)
                if not b'fs400_drv_read' in obj.CtrlB.read_all():
                    print(obj.CtrlB.read_all())
                obj.CtrlB.write(b'fs400_drv_read 0x0427\n\r');time.sleep(0.2)
                s=obj.CtrlB.read_all().decode('utf-8').split(' ')
                b1=s[3][-2:]+s[4][-2:]
                obj.CtrlB.write(b'fs400_drv_read 0x0527\n\r');time.sleep(0.2)
                print(obj.CtrlB.read_all())
                obj.CtrlB.write(b'fs400_drv_read 0x0527\n\r');time.sleep(0.2)
                s=obj.CtrlB.read_all().decode('utf-8').split(' ')
                b2=s[3][-2:]+s[4][-2:]
                obj.CtrlB.write(b'fs400_drv_read 0x0627\n\r');time.sleep(0.2)
                print(obj.CtrlB.read_all())
                obj.CtrlB.write(b'fs400_drv_read 0x0627\n\r');time.sleep(0.2)
                s=obj.CtrlB.read_all().decode('utf-8').split(' ')
                b3=s[3][-2:]+s[4][-2:]
                obj.CtrlB.write(b'fs400_drv_read 0x0727\n\r');time.sleep(0.2)
                print(obj.CtrlB.read_all())
                obj.CtrlB.write(b'fs400_drv_read 0x0727\n\r');time.sleep(0.2)
                s=obj.CtrlB.read_all().decode('utf-8').split(' ')
                b4=s[3][-2:]+s[4][-2:]
                m2='fs400_set_eep_drv_pdk 0x{} 0x{} 0x{} 0x{}\n\r'.format(b1,b2,b3,b4).encode('utf-8')
                print('PKD information write:\n\r',m2)
                obj.CtrlB.write(m2);time.sleep(0.2)#b'fs400_set_eep_drv_pdk 0x1122 0x1122 0x1122 0x1122\n'
            else:
                pass
            #save config
            obj.CtrlB.write(b'fs400_eep_arg_save 1 1\n\r');time.sleep(0.2)
            print(obj.CtrlB.read_until(b'Save ok'))
            print(obj.CtrlB.read_all(), 'page 1 saved!')

            #***----------page 2----------***
            print('Start write Page 2:')
            obj.CtrlB.write(b'fs400_set_eep_voainfo 0x3b 0x1 -170 4310\n\r');time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_voax 0x0 0x0 0x0 0x0\n\r');time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_voay 0x0 0x0 0x0 0x0\n\r');time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_sensor_k_b 0x0 0x0\n\r');time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_tia_k_b 0x0 0x0\n\r');time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_drv_k_b 0x0 0x0\n\r');time.sleep(0.2)
            obj.CtrlB.write(b'fs400_eep_arg_save 1 2\n\r');time.sleep(0.2)
            print(obj.CtrlB.read_until(b'Save ok'))
            print(obj.CtrlB.read_all(), 'page 2 saved!')

            #***----------page 3----------***
            print('Start write Page 3:')
            obj.CtrlB.write(b'fs400_set_eep_400g_pre 0 127 -12 -6 1 4\n\r');time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_400g_pre 5 4 2 0 0 0\n\r');time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_400g_pre 10 0 0 0 0 0\n\r');time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_400g_pre 15 0 0 0 0 0\n\r');time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_400g_pre 20 0 0 0 0 0\n\r');time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_400g_pre 25 0 0 0 0 0\n\r');time.sleep(0.2)
            obj.CtrlB.write(b'fs400_eep_arg_save 1 3\n\r');time.sleep(0.2)
            print(obj.CtrlB.read_until(b'Save ok'))
            print(obj.CtrlB.read_all(), 'page 3 saved!')

            #***----------page 4----------***
            print('Start write Page 4:')
            obj.CtrlB.write(b'fs400_set_eep_200g_pre 0 127 -5 -3 -1 1\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_200g_pre 5 3 3 1 -1 -1\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_200g_pre 10 -1 0 1 1 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_200g_pre 15 0 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_200g_pre 20 0 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_200g_pre 25 0 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_eep_arg_save 1 4\n\r');
            time.sleep(0.2)
            print(obj.CtrlB.read_until(b'Save ok'))
            print(obj.CtrlB.read_all(), 'page 4 saved!')

            #***----------page 5----------***
            print('Start write Page 5:')
            obj.CtrlB.write(b'fs400_set_eep_200gps_pre 0 99 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_200gps_pre 5 0 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_200gps_pre 10 0 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_200gps_pre 15 0 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_200gps_pre 20 0 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_200gps_pre 25 0 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_eep_arg_save 1 5\n\r');
            time.sleep(0.2)
            print(obj.CtrlB.read_until(b'Save ok'))
            print(obj.CtrlB.read_all(), 'page 5 saved!')

            #***----------page 7----------***
            print('Start write Page 7:')
            if flag=='IDT':
                obj.CtrlB.write(b'fs400_set_eep_dev_drv1 0 0x0 0x01 0x0 0x740074 0x740074\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_drv2 0 0x0601 0xac 0x0e44 0x7771 0x3727\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_drv3 0 0x5111 0x01 0x1770 '+(Drv3).encode('utf-8')+b'\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_tia1 0 0x03 0x000a 0xaf88\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_tia2 0 0x60 0xffff 0x18 2500\n\r');time.sleep(0.2)
            elif flag=='ALU':
                obj.CtrlB.write(b'fs400_set_eep_dev_drv1 0 0x0 0x0 0x320 0x7d007d0 0x7d007d0\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_drv2 0 0xffff 0xff 0xffff 0xffff 0xffff\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_drv3 0 0xffff 0x01 0x1770 0x12c0\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_tia1 0 0x1e 0x01 0xffff 0xffff\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_tia2 0 2000 0xffff 1500 2500\n\r');time.sleep(0.2)
            else:
                print('Device Flag error')
                return
            obj.CtrlB.write(b'fs400_eep_arg_save 1 7\n\r');time.sleep(0.2)
            print(obj.CtrlB.read_until(b'Save ok'))
            print(obj.CtrlB.read_all(), 'page 7 saved!')

            # ***----------page 8----------***
            print('Start write Page 8:')
            if flag=='IDT':
                obj.CtrlB.write(b'fs400_set_eep_dev_drv1 1 0x0 0x01 0x0 0x580058 0x580058\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_drv2 1 0x0601 0xac 0x0e44 0x7771 0x3727\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_drv3 1 0x5111 0x01 0x1770 '+(Drv3).encode('utf-8')+b'\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_tia1 1 0x03 0x000a 0xaf88\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_tia2 1 0x60 0xffff 0x18 2500\n\r');time.sleep(0.2)
            elif flag=='ALU':
                obj.CtrlB.write(b'fs400_set_eep_dev_drv1 1 0x0 0x0 0x320 0x7d007d0 0x7d007d0\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_drv2 1 0xffff 0xff 0xffff 0xffff 0xffff\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_drv3 1 0xffff 0x01 0x1770 0x12c0\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_tia1 1 0x1e 0x01 0xffff 0xffff\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_tia2 1 2000 0xffff 1500 2500\n\r');time.sleep(0.2)
            else:
                print('Device Flag error')
                return
            obj.CtrlB.write(b'fs400_eep_arg_save 1 8\n\r');
            time.sleep(0.2)
            print(obj.CtrlB.read_until(b'Save ok'))
            print(obj.CtrlB.read_all(), 'page 8 saved!')

            # ***----------page 9----------***
            print('Start write Page 9:')
            if flag=='IDT':
                obj.CtrlB.write(b'fs400_set_eep_dev_drv1 2 0x0 0x01 0x0 0xc000c0 0xc000c0\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_drv2 2 0x0601 0xac 0x0e44 0x7771 0x3727\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_drv3 2 0x5111 0x01 0x1770 '+(Drv3).encode('utf-8')+b'\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_tia1 2 0x03 0x000a 0xaf88\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_tia2 2 0x60 0xffff 0x08 2500\n\r');time.sleep(0.2)
            elif flag=='ALU':
                obj.CtrlB.write(b'fs400_set_eep_dev_drv1 2 0x0 0x0 0x320 0x7d007d0 0x7d007d0\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_drv2 2 0xffff 0xff 0xffff 0xffff 0xffff\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_drv3 2 0xffff 0x01 0x1770 0x12c0\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_tia1 2 0x1e 0x01 0xffff 0xffff\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_tia2 2 2000 0xffff 1500 2500\n\r');time.sleep(0.2)
            else:
                print('Device Flag error')
                return
            obj.CtrlB.write(b'fs400_eep_arg_save 1 9\n\r');
            time.sleep(0.2)
            print(obj.CtrlB.read_until(b'Save ok'))
            print(obj.CtrlB.read_all(), 'page 9 saved!')

            # ***----------page 11----------***
            print('Start write Page 11:')
            if flag=='IDT':
                obj.CtrlB.write(b'fs400_set_eep_dev_drv1 4 0x0 0x01 0x0 0x900090 0x900090\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_drv2 4 0x0601 0xac 0x0e44 0x7771 0x3727\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_drv3 4 0x5111 0x01 0x1770 '+(Drv3).encode('utf-8')+b'\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_tia1 4 0x03 0x000a 0xaf88\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_tia2 4 0x60 0xffff 0x00 2500\n\r');time.sleep(0.2)
            elif flag=='ALU':
                obj.CtrlB.write(b'fs400_set_eep_dev_drv1 4 0x0 0x0 0x320 0x7d007d0 0x7d007d0\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_drv2 4 0xffff 0xff 0xffff 0xffff 0xffff\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_drv3 4 0xffff 0x01 0x1770 0x12c0\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_tia1 4 0x1e 0x01 0xffff 0xffff\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_tia2 4 2000 0xffff 1500 2500\n\r');time.sleep(0.2)
            else:
                print('Device Flag error')
                return
            obj.CtrlB.write(b'fs400_eep_arg_save 1 11\n\r');
            time.sleep(0.2)
            print(obj.CtrlB.read_until(b'Save ok'))
            print(obj.CtrlB.read_all(), 'page 11 saved!')

            # ***----------page 12----------***
            print('Start write Page 12:')
            obj.CtrlB.write(b'fs400_set_eep_pre_12 0 127 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_pre_12 5 0 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_pre_12 10 0 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_pre_12 15 0 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_pre_12 20 0 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_pre_12 25 0 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_eep_arg_save 1 12\n\r');
            time.sleep(0.2)
            print(obj.CtrlB.read_until(b'Save ok'))
            print(obj.CtrlB.read_all(), 'page 12 saved!')

            # ***----------page 13----------***
            print('Start write Page 13:')
            if flag=='IDT':
                obj.CtrlB.write(b'fs400_set_eep_dev_drv1 5 0x0 0x01 0x0 0x900090 0x900090\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_drv2 5 0x0601 0xac 0x0e44 0x7771 0x3727\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_drv3 5 0x5111 0x01 0x1770 '+(Drv3).encode('utf-8')+b'\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_tia1 5 0x03 0x000a 0xaf88\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_tia2 5 0x60 0xffff 0x00 2500\n\r');time.sleep(0.2)
            elif flag=='ALU':
                obj.CtrlB.write(b'fs400_set_eep_dev_drv1 5 0x0 0x0 0x320 0x7d007d0 0x7d007d0\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_drv2 5 0xffff 0xff 0xffff 0xffff 0xffff\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_drv3 5 0xffff 0x01 0x1770 0x12c0\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_tia1 5 0x1e 0x01 0xffff 0xffff\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_tia2 5 2000 0xffff 1500 2500\n\r');time.sleep(0.2)
            else:
                print('Device Flag error')
                return
            obj.CtrlB.write(b'fs400_eep_arg_save 1 13\n\r');
            time.sleep(0.2)
            print(obj.CtrlB.read_until(b'Save ok'))
            print(obj.CtrlB.read_all(), 'page 13 saved!')

            # ***----------page 14----------***
            print('Start write Page 14:')
            obj.CtrlB.write(b'fs400_set_eep_pre_14 0 99 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_pre_14 5 0 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_pre_14 10 0 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_pre_14 15 0 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_pre_14 20 0 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_pre_14 25 0 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_eep_arg_save 1 14\n\r');
            time.sleep(0.2)
            print(obj.CtrlB.read_until(b'Save ok'))
            print(obj.CtrlB.read_all(), 'page 14 saved!')

            # ***----------page 15----------***
            print('Start write Page 15:')
            if flag=='IDT':
                obj.CtrlB.write(b'fs400_set_eep_dev_drv1 6 0x0 0x01 0x0 0x700070 0x700070\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_drv2 6 0x0601 0xac 0x0e44 0x7771 0x3727\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_drv3 6 0x5111 0x01 0x1770 '+(Drv3).encode('utf-8')+b'\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_tia1 6 0x03 0x000a 0xaf88\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_tia2 6 0x60 0xffff 0x00 2500\n\r');time.sleep(0.2)
            elif flag=='ALU':
                obj.CtrlB.write(b'fs400_set_eep_dev_drv1 6 0x0 0x0 0x320 0x7d007d0 0x7d007d0\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_drv2 6 0xffff 0xff 0xffff 0xffff 0xffff\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_drv3 6 0xffff 0x01 0x1770 0x12c0\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_tia1 6 0x1e 0x01 0xffff 0xffff\n\r');time.sleep(0.2)
                obj.CtrlB.write(b'fs400_set_eep_dev_tia2 6 2000 0xffff 1500 2500\n\r');time.sleep(0.2)
            else:
                print('Device Flag error')
                return
            obj.CtrlB.write(b'fs400_eep_arg_save 1 15\n\r');
            time.sleep(0.2)
            print(obj.CtrlB.read_until(b'Save ok'))
            print(obj.CtrlB.read_all(), 'page 15 saved!')

            # ***----------page 16----------***
            print('Start write Page 16:')
            obj.CtrlB.write(b'fs400_set_eep_pre_16 0 127 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_pre_16 5 0 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_pre_16 10 0 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_pre_16 15 0 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_pre_16 20 0 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_set_eep_pre_16 25 0 0 0 0 0\n\r');
            time.sleep(0.2)
            obj.CtrlB.write(b'fs400_eep_arg_save 1 16\n\r');
            time.sleep(0.2)
            print(obj.CtrlB.read_until(b'Save ok'))
            print(obj.CtrlB.read_all(), 'page 16 saved!')

            # #print out the result of pages
            # obj.CtrlB.write(b'fs400_eep_arg_read 1 0\n');time.sleep(0.2)
            # print('Page 0:\n',obj.CtrlB.read_until(b'check ok'))
            # obj.CtrlB.write(b'fs400_eep_arg_read 1 1\n');time.sleep(0.2)
            # print('Page 1:\n',obj.CtrlB.read_until(b'check ok'))
            # obj.CtrlB.write(b'fs400_eep_arg_read 1 2\n');time.sleep(0.2)
            # print('Page 2:\n',obj.CtrlB.read_until(b'check ok'))
            # obj.CtrlB.write(b'fs400_eep_arg_read 1 3\n');time.sleep(0.2)
            # print('Page 3:\n',obj.CtrlB.read_until(b'check ok'))
            # obj.CtrlB.write(b'fs400_eep_arg_read 1 7\n');time.sleep(0.2)
            # print('Page 7:\n',obj.CtrlB.read_until(b'check ok'))
            print('Eeprom write OK')
        else:
            print('Error, Control board port is not open')
    except Exception as e:
        print(e,'Please check whether board or driver is up and configed!!!')

def write_VOAcal2EEP(obj,voaCal):
    '''
    write the result of VOA calibration to eeprom
    :param voaCal:dataframe of VOA calibration data
    example:
    SN	RX_BW_DESK	TEMP	DATE	TIME	400G_Fre	VOA_XI	VOA_XQ	VOA_YI	VOA_YQ
    Test	Test	Test	Test	Test	190.1125	0	98	83	0
    ......
    ...
    .
    :return:NA
    '''
    page_lis=[i for i in range(20,52)]
    page_id=[]
    group_num=[]
    for i in range(len(page_lis)):
        page_id.append(str(i))
        page_id.append(str(i))
        page_id.append(str(i))
        group_num.append('0')
        group_num.append('1')
        group_num.append('2')

    #read raw data or direct pass the dataframe
    df=voaCal
    freq   = df.iloc[:, 5]
    voa_XI = df.iloc[:, 6]
    voa_XQ = df.iloc[:, 7]
    voa_YI = df.iloc[:, 8]
    voa_YQ = df.iloc[:, 9]

    for i,j in enumerate(page_id):
        page_na=str(int(j)+20)
        group_na=group_num[i]
        if (np.mod(i,3)==0):
            #print(page_id[i])
            print('Now write page{}:'.format(page_na))
        commd='fs400_set_eep_voa_table {} {} {} {} {} {} {} {} {} {}\n\r'.format(j,group_na,
                                                                                 str(voa_XI[i]),str(voa_XI[i]),
                                                                                 str(voa_XQ[i]),str(voa_XQ[i]),
                                                                                 str(voa_YI[i]),str(voa_YI[i]),
                                                                                 str(voa_YQ[i]),str(voa_YQ[i])).encode('utf-8')
        print(commd)
        obj.CtrlB.write(commd)
        print(obj.CtrlB.read_until(b'shell'))
        #print(obj.CtrlB.read_all())
        if (np.mod(i,3)==2):
            print(page_id[i])
            obj.CtrlB.write('fs400_eep_arg_save {}\n\r'.format(page_na).encode('utf-8'));
            time.sleep(0.2)
            print(obj.CtrlB.read_until(b'Save ok'))
            print(obj.CtrlB.read_all())
            print('Page{} write and saved!'.format(page_na))

    #The end of VOA write
    print('VOA write done!')

def TIA_on(obj):
    '''
    TIA on before ICR test
    :param obj:
    :return: NA
    '''
    if serial.Serial.isOpen(obj.CtrlB):
        obj.CtrlB.write(b'cpld_wr 0x7f 0x8\n')
        time.sleep(0.1)
        obj.CtrlB.write(b'set_adj 3 100\n')
        time.sleep(0.1)
        obj.CtrlB.write(b'set_adj 4 100\n')
        time.sleep(0.1)
        obj.CtrlB.write(b'en_adj 3 1\n')
        time.sleep(0.1)
        obj.CtrlB.write(b'en_adj 4 1\n')
        time.sleep(0.1)
        obj.CtrlB.write(b'slave_on\n')
        time.sleep(0.1)
        print('TIA on!')
        # a=obj.CtrlB.read_until(b'check ok')
        # obj.CtrlB.close()

def TIA_write(obj,l='0000',c='0000'):
    '''
    TIA write and check function
    :param obj:
    :param l: location
    :param c: configuration
    :return: True/False
    '''
    if l=='0026':
        r= b'fs400_tia_read 0x0027\n'
    elif l=='0024':
        r= b'fs400_tia_read 0x0025\n'
    elif l=='000e':
        r= b'fs400_tia_read 0x000f\n'
    elif l=='002a':
        r= b'fs400_tia_read 0x002b\n'
    elif l=='012a':
        r= b'fs400_tia_read 0x012b\n'
    elif l=='022a':
        r= b'fs400_tia_read 0x022b\n'
    elif l=='032a':
        r= b'fs400_tia_read 0x032b\n'
    elif l=='002c':
        r= b'fs400_tia_read 0x002d\n'
    elif l=='012c':
        r= b'fs400_tia_read 0x012d\n'
    elif l=='022c':
        r= b'fs400_tia_read 0x022d\n'
    elif l=='032c':
        r= b'fs400_tia_read 0x032d\n'
    elif l=='000c':
        r= b'fs400_tia_read 0x000d\n'

    #write config
    w='fs400_tia_write 0x{} 0x{}\n'.format(l,c).encode('utf-8')
    s=''
    for i in range(4):
        obj.CtrlB.flushInput()
        obj.CtrlB.flushOutput()
        obj.CtrlB.write(w)
        time.sleep(0.2)
        obj.CtrlB.write(r)
        time.sleep(0.2)
        a=obj.CtrlB.read_all().decode('utf-8').split('\r\n')[1].strip()
        s=a[18:20]+a[25:27]
        if s!=c:
            continue
        else:
            print('TIA write in {} with {} OK!'.format(l,c))
            return True
    print('TIA write in {} with {} error!!!'.format(l,c))
    return False

def TIA_config(obj):
    '''
    TIA configuration before ICR test
    :param obj: mawin obj
    :return: NA
    '''
    if serial.Serial.isOpen(obj.CtrlB):
        obj.CtrlB.flushInput()
        #内部寄存器2
        if not TIA_write(obj,'0026','000a'):return
        #内部寄存器1
        if not TIA_write(obj,'0024','af88'):return
        #MGC
        if not TIA_write(obj,'000e','0002'):return
        #GC0
        if not TIA_write(obj,'002a','00a0'):return
        #GC1
        if not TIA_write(obj,'012a','00a0'):return
        #GC2
        if not TIA_write(obj,'022a','00a0'):return
        #GC3
        if not TIA_write(obj,'032a','00a0'):return
        #BW0
        if not TIA_write(obj,'002c','003f'):return
        #BW1
        if not TIA_write(obj,'012c','003f'):return
        #BW2
        if not TIA_write(obj,'022c','003f'):return
        #BW3
        if not TIA_write(obj,'032c','003f'):return
        #输出开启
        if not TIA_write(obj,'000c','0000'):return
        #DAC开启，PD上电
        obj.CtrlB.write(b'switch_set 10 0\n')
        time.sleep(0.1)
        obj.CtrlB.write(b'cpld_spi_wr 0x2c 2150\n')
        time.sleep(0.1)
        obj.CtrlB.write(b'cpld_spi_wr 0x2f 2150\n')
        time.sleep(0.1)
        print('TIA config Done!')

def TIA_config_ALU(obj,t):
    '''
    ALU TIA configuration before ICR test
    :param obj: mawin obj,t:type of test function
    :return: NA
    '''
    if serial.Serial.isOpen(obj.CtrlB):
        obj.CtrlB.flushInput()
        if t:
            print('ALU 响应度测试')
            #DAC开启，PD上电
            obj.CtrlB.write(b'switch_set 10 0\n')
            time.sleep(0.1)
            obj.CtrlB.write(b'cpld_spi_wr 0x2c 2150\n')
            time.sleep(0.1)
            obj.CtrlB.write(b'cpld_spi_wr 0x2f 2150\n')
            time.sleep(0.1)
            # obj.CtrlB.write(b'cpld_wr 0x7f 0x7\n')
            # time.sleep(0.1)
            # obj.CtrlB.write(b'set_adj 9 180\n')
            # time.sleep(0.1)
            # obj.CtrlB.write(b'en_adj 9 1\n')
            # time.sleep(0.1)
            # obj.CtrlB.write(b'cpld_wr 0x7f 0x8\n')
            # time.sleep(0.1)
            # obj.CtrlB.write(b'set_adj 1 220\n')
            # time.sleep(0.1)
            # obj.CtrlB.write(b'en_adj 1 1\n')
            # time.sleep(0.1)
            print('ALU TIA config Done!')
        else:
            print('ALU 带宽测试')
            #DAC开启，PD上电
            # obj.CtrlB.write(b'switch_set 10 0\n')
            # time.sleep(0.1)
            # obj.CtrlB.write(b'cpld_spi_wr 0x2c 2150\n')
            # time.sleep(0.1)
            # obj.CtrlB.write(b'cpld_spi_wr 0x2f 2150\n')
            # time.sleep(0.1)
            obj.CtrlB.write(b'cpld_wr 0x7f 0x7\n')
            time.sleep(0.1)
            obj.CtrlB.write(b'set_adj 9 180\n')
            time.sleep(0.1)
            obj.CtrlB.write(b'en_adj 9 1\n')
            time.sleep(0.1)
            obj.CtrlB.write(b'cpld_wr 0x7f 0x8\n')
            time.sleep(0.1)
            obj.CtrlB.write(b'set_adj 1 220\n')
            time.sleep(0.1)
            obj.CtrlB.write(b'en_adj 1 1\n')
            time.sleep(0.1)
            print('ALU TIA config Done!')

def status_color(obj,color='green'):
    '''
    change the color of corrsponding widget
    :param obj: widget
    :param color: color string
    :return: NA
    '''
    if color=='green':
        obj.setStyleSheet("background-color: rgb(0, 255, 0);")
    if color=='yellow':
        obj.setStyleSheet("background-color: rgb(255, 255, 0);")
    if color=='red':
        obj.setStyleSheet("background-color: rgb(255, 0, 0);")
    if color=='blue':
        obj.setStyleSheet("background-color: rgb(0, 0, 255);")

def VOA_cal_data(da,x_tofit,x_toget,fre_full):
    '''
    :param da: DataFrame of voa raw data
    :param x_tofit: channel(list of fre) of voa raw data
    :param x_toget: channel(list of fre) of data fitted
    :param cppITLA: whether C++ ITLA to determine wavelength config
    :return: data calibrated to write to EEPROM， (voa_XI,voa_XQ,voa_YI,voa_YQ)
    1.Calculate voa wavelength frequency range 190.1125~197.2375THz,75G step;
    2.Fre and VOA value polyfit:
        a=[0.2,0.4,0.6,0.8,0.9,1.5,1.6]
        x=range(1,len(a)+1)
        p=np.poly1d(np.polyfit(x,a,2))
        plt.plot(x,p(x))
        np.diff(p(x))
        In [40]: p(x)
        Out[40]:array([0.21666667, 0.37857143, 0.57142857, 0.7952381 , 1.05      ,
               1.33571429, 1.65238095])
    3.Temperature medium value:
        da.median(0)[1]...
    4.Get round, up to round: math.ceil()
                down to round:math.floor()
    '''
    tempe=da.iloc[:,6]
    voa_value=da.iloc[:,7:]
    voa_XIQ_dif=voa_value.iloc[:,0]-voa_value.iloc[:,1]
    voa_YIQ_dif = voa_value.iloc[:, 2] - voa_value.iloc[:, 3]

    #remove the data to jump,namely 1000 in it
    voa_deal=voa_XIQ_dif.to_list()
    ch_jump=[]
    for t,j in enumerate(voa_deal):
        if j==1000:
            ch_jump.append(t)
    if not len(ch_jump)==0:
        voa_XIQ_dif=voa_XIQ_dif.drop(ch_jump)
        voa_XIQ_dif=voa_XIQ_dif.reindex()
        voa_YIQ_dif=voa_YIQ_dif.drop(ch_jump)
        voa_YIQ_dif=voa_YIQ_dif.reindex()
        x_tofit=np.delete(x_tofit,ch_jump)
    #first calculate the temperatrue corelate factor
    T_median=tempe.median()
    if T_median<40:
        shift_dire='UP'
        N=math.ceil((40-T_median)/4.4)
    elif T_median>45:
        shift_dire='DOWN'
        N=math.ceil((T_median-45)/4.4)
    else:
        shift_dire='NONE'
        N=0
    print('\nShift direction: ',shift_dire)
    print('Shift position qty: ',N)
    #Then calculate VOA fit data, from 6 degree to 2, until dVOA<=10
    result=[]

    # VOA XIQ first
    for deg in range(7,1,-1):
        p = np.poly1d(np.polyfit(x_tofit, voa_XIQ_dif, deg))
        p1= p(x_toget)
        if max(abs(np.diff(p1)))<=10:
            print('get XIXQ dVOA<=10, fit data at {} degree.'.format(str(deg)))
            print('dVOA XIXQ:\n',max(abs(np.diff(p1))))
            break
        print('Did not get XIXQ dVOA<=10, fit data at {} degree, go to next or end...'.format(str(deg)))
    result.append(p1)
    print('XIXQ:\n',p1)
    # VOA YIQ now
    for deg in range(7, 1, -1):
        p = np.poly1d(np.polyfit(x_tofit, voa_YIQ_dif, deg))
        p1 = p(x_toget)
        if max(abs(np.diff(p1))) <= 10:
            print('get YIYQ dVOA<=10, fit data at {} degree fit.'.format(str(deg)))
            print('dVOA YIYQ:\n', max(abs(np.diff(p1))))
            break
        print('Did not get YIYQ dVOA<=10, fit data at {} degree fit, go to next or end...'.format(str(deg)))
    result.append(p1)
    print('YIYQ:\n', p1)
    #result now have two columns, transfer and transpose to DataFrame format
    resu_df=pd.DataFrame(result).transpose()
    resu_df.columns=['XIXQ_diff','YIYQ_diff']
    #shift the frequency according to the temperature
    fre_fix=x_toget
    if shift_dire=='UP':
        fre_fix=[i-0.075*N for i in x_toget]
    elif shift_dire=='DOWN':
        fre_fix = [i + 0.075 * N for i in x_toget]
    resu_df.insert(0,'Freq(THz)',fre_fix)
    # print(resu_df)
    # return resu_df
    print('Raw data:\n',resu_df)
    voa_cal_resu = resu_df
    voa_fre_shift = np.array(voa_cal_resu['Freq(THz)'])
    voa_XIXQ_shift = np.array(voa_cal_resu['XIXQ_diff'].round(),dtype='int').tolist()
    voa_YIYQ_shift = np.array(voa_cal_resu['YIYQ_diff'].round(),dtype='int').tolist()
    sta_ind = voa_fre_shift[0]
    end_ind = voa_fre_shift[-1]
    star_fre = fre_full[np.argwhere(fre_full < sta_ind)]
    end_fre = fre_full[np.argwhere(fre_full > end_ind)]
    star_fre1 = [i[0] for i in star_fre]
    end_fre1 = [i[0] for i in end_fre]
    #fre_list = star_fre1 + voa_fre_shift.tolist() + end_fre1
    voa_XIXQ_shift_all = [voa_XIXQ_shift[0]] * len(star_fre1) + voa_XIXQ_shift + [voa_XIXQ_shift[-1]] * len(end_fre1)
    voa_YIYQ_shift_all = [voa_YIYQ_shift[0]] * len(star_fre1) + voa_YIYQ_shift + [voa_YIYQ_shift[-1]] * len(end_fre1)
    voa_XI, voa_XQ, voa_YI, voa_YQ = [[], [], [], []]
    # restore the four pairs voa
    for i in voa_XIXQ_shift_all:
        if i > 0:
            voa_XI.append(i)
            voa_XQ.append(0)
        else:
            voa_XI.append(0)
            voa_XQ.append(0-i)
    for i in voa_YIYQ_shift_all:
        if i > 0:
            voa_YI.append(i)
            voa_YQ.append(0)
        else:
            voa_YI.append(0)
            voa_YQ.append(0-i)

    return(voa_XI,voa_XQ,voa_YI,voa_YQ)
    # # write data into the fit voa data,'VOA_XI','VOA_XQ', 'VOA_YI', 'VOA_YQ'
    # test_result_fit = DataFrame(columns=('SN', 'RX_BW_DESK', 'TEMP', 'DATE', 'TIME', '400G_Fre',
    #                                      'VOA_XI', 'VOA_XQ', 'VOA_YI', 'VOA_YQ'))
    # for i in len(fre_list):
    #     da_add = config + [fre_list[i]] + [voa_XI[i]] + [voa_XQ[i]] + [voa_YI[i]] + [voa_YQ[i]]
    #     test_result_fit.loc[i] = da_add
