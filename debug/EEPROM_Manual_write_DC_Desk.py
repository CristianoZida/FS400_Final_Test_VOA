# -*- coding: utf-8 -*-
import pandas as pd
from pandas import DataFrame
import numpy as np
import os
import sys
import Common_functions.General_functions as gf
import Instruments.CtrlBoard as cb
import time

class a():
    CtrlB=object
    port='com5'
    drv_up=''
    drv_down=''
    config_path=''
    device_type=''

# Config default configuration files and paths
# this code ensure the path to be switched to where the script is
abspath = os.path.dirname(__file__)
sys.path.append(abspath)
print(abspath)
if abspath == '':
    os.chdir(sys.path[0])
    script_path = sys.path[0]
else:
    os.chdir(abspath)
    script_path = abspath

mawin=a()
cb.open_board(mawin,mawin.port)

mawin.CtrlB.write(b'slave_on\n\r')
print(mawin.CtrlB.read_until(b'shell'))
print(mawin.CtrlB.read_all())
mawin.CtrlB.write(b'fs400_stop_fs400_performance\n\r')
print(mawin.CtrlB.read_until(b'shell'))
print(mawin.CtrlB.read_all())
mawin.CtrlB.write(b'cpld_spi_wr 0x31 1\n\r')
print(mawin.CtrlB.read_until(b'shell'))
print(mawin.CtrlB.read_all())

sn='WRRC280049'
config_path=r'C:\Test_script\FS400_Final_Test_Debug\Configuration'
path=r'C:\Test_result\DC\Normal\WRRC280049_20220901_172148'
f1=os.path.join(path,'WRRC280049_DC_20220901_172148.csv')
df1=pd.read_csv(f1)
print(df1)
device_type='IDT'
if device_type=='ALU':
    drv_up = os.path.join(config_path, 'Setup_driverup_Ctrlboard56017837A002_20211203_ALU.txt')
    drv_down = os.path.join(config_path, 'Setup_driverdown_CtrlboardA001_20210820_ALU.txt')
if device_type=='IDT':
    drv_up = os.path.join(config_path, 'Setup_driverup_Ctrlboard56017837A002_20211203.txt')
    drv_down = os.path.join(config_path, 'Setup_driverdown_CtrlboardA001_20210820.txt')

board_up = os.path.join(config_path, 'Setup_brdup_CtrlboardA001_20220113_C80.txt')

#Drv up
#Driver power up
cb.board_set(mawin,board_up)
cb.board_set(mawin,drv_up)
gf.write_eeprom_final(mawin,sn,df1,device_type,'0x12d2')
cb.board_set(mawin,drv_down)
mawin.CtrlB.close()