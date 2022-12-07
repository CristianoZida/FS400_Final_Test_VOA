# -*- coding: utf-8 -*-
import pandas as pd
from pandas import DataFrame
import numpy as np
import os
import sys
import Common_functions.General_functions as gf
import Instruments.CtrlBoard as cb
import Instruments.PowerMeter as pm
import time

class a():
    CtrlB=object
    PM=object
    board_port='com11'
    PM_port='com6'
pm_ch=2
itla_ch=1
mawin=a()
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

ITLA_type=input('Please choose ITLA type, \nC++: enter \'Y\'\nC  : enter \'N\'\n')
if ITLA_type.upper()=='Y':
    cpp_itla = True
    typ='C++'
elif ITLA_type.upper()=='N':
    cpp_itla = False
    typ = 'C'
else:
    a=input('Error input, you should enter Y or N to choose ITLA type, press any key to exit...')
    exit()

cb.open_board(mawin,mawin.board_port)
pm.open_PM(mawin, mawin.PM_port)

mawin.CtrlB.write(b'slave_on\n\r')
print(mawin.CtrlB.read_until(b'shell'))
print(mawin.CtrlB.read_all())
mawin.CtrlB.write(b'fs400_stop_fs400_performance\n\r')
print(mawin.CtrlB.read_until(b'shell'))
print(mawin.CtrlB.read_all())
mawin.CtrlB.write(b'cpld_spi_wr 0x31 1\n\r')
print(mawin.CtrlB.read_until(b'shell'))
print(mawin.CtrlB.read_all())

time_stamp=gf.get_timestamp(1)
report_name='Res_Optical_calibration_ITLA_CH{}_{}_{}.csv'.format(itla_ch,typ,time_stamp)
cal_file=os.path.join(script_path,report_name)

comlist = []
fre = list
if cpp_itla:
    if itla_ch==0:
        comlist = ['slave_on',
                   'fs400_stop_fs400_performance',
                   'cpld_spi_wr 0x31 1',
                   'itla_wr 0 0x32 0',
                   'itla_wr 0 0x34 0x2ee',
                   'itla_wr 0 0x35 0xbe',
                   'itla_wr 0 0x36 0x1bd5',
                   'itla_wr 0 0x31 0x5dc',
                   'itla_wr 0 0x30 1',
                   'itla_wr 0 0x32 8']
    else:
        comlist = ['slave_on',
                   'fs400_stop_fs400_performance',
                   'cpld_spi_wr 0x31 1',
                   'itla_wr 1 0x32 0',
                   'itla_wr 1 0x34 0x2ee',
                   'itla_wr 1 0x35 0xbe',
                   'itla_wr 1 0x36 0x1bd5',
                   'itla_wr 1 0x31 0x3e8',
                   'itla_wr 1 0x30 1',
                   'itla_wr 1 0x32 8']
    mawin.channel = [i for i in range(1, 81)]
else:
    if itla_ch==0:
        comlist = ['slave_on',
                   'fs400_stop_fs400_performance',
                   'cpld_spi_wr 0x31 1',
                   'itla_wr 0 0x32 0',
                   'itla_wr 0 0x34 0x2ee',
                   'itla_wr 0 0x35 0xbf',
                   'itla_wr 0 0x36 0xc35',
                   'itla_wr 0 0x31 0x5dc'
                   'itla_wr 0 0x30 1',
                   'itla_wr 0 0x32 8']
    else:
        comlist = ['slave_on',
                   'fs400_stop_fs400_performance',
                   'cpld_spi_wr 0x31 1',
                   'itla_wr 1 0x32 0',
                   'itla_wr 1 0x34 0x2ee',
                   'itla_wr 1 0x35 0xbf',
                   'itla_wr 1 0x36 0xc35',
                   'itla_wr 1 0x31 0x3e8'
                   'itla_wr 1 0x30 1',
                   'itla_wr 1 0x32 8']
    mawin.channel = [i for i in range(9, 73)]

print('ITLA configed and opened, set output power to 15dBm!')
cb.board_set_CmdList_DC(mawin, comlist)

test_result = DataFrame(columns=('Fre_ch','pwr'))
tt=[]

input('Please connect ITLA output fiber directly to Power Meter port {}!\nPress any key and the calibration will start...'.format(pm_ch))
mawin.CtrlB.write('itla_wr {} 0x32 8\r\n'.format(itla_ch).encode('utf-8'))
print(mawin.CtrlB.read_until(b'Write itla OK'))

for i in range(len(mawin.channel)):
    ch=mawin.channel[i]
    print('Now calibrate channel {}'.format(ch))
    #switch ITLA wavelength
    if itla_ch==0:
        cmClose = b'itla_wr 0 0x32 0\n\r'
        cmOpen = b'itla_wr 0 0x32 8\n\r'

        if cpp_itla:
            ch_toset = int(ch)
            cm1 = 'itla_wr 0 0x30 {}\n\r'.format(str(hex(int(ch)))).encode('utf-8')
        else:
            ch_toset = int(ch) - 8
            cm1 = 'itla_wr 0 0x30 {}\n\r'.format(str(hex(int(ch) - 8))).encode('utf-8')
    else:
        cmClose = b'itla_wr 1 0x32 0\n\r'
        cmOpen = b'itla_wr 1 0x32 8\n\r'

        if cpp_itla:
            ch_toset = int(ch)
            cm1 = 'itla_wr 1 0x30 {}\n\r'.format(str(hex(int(ch)))).encode('utf-8')
        else:
            ch_toset = int(ch) - 8
            cm1 = 'itla_wr 1 0x30 {}\n\r'.format(str(hex(int(ch) - 8))).encode('utf-8')

    mawin.CtrlB.timeout=60
    mawin.CtrlB.write(cm1)
    print(mawin.CtrlB.read_until(b'Write itla'))
    time.sleep(0.2)
    mawin.CtrlB.write(cm1)
    print(mawin.CtrlB.read_until(b'Write itla OK'))
    #time.sleep(2)
    #Need to check whether the wavelength of PM need to be switched
    #read the power
    pwr=pm.read_PM(mawin,pm_ch)
    tt=[i+1,pwr]
    print(tt)
    test_result.loc[i]=tt

test_result.to_csv(cal_file,index=False)
print('Calibration completed!')
if itla_ch==0:
    comlist = ['slave_off',
            'itla_wr 0 0x32 0']
else:
    comlist = ['slave_off',
               'itla_wr 1 0x32 0']
print('Close the ITLA and slave off board.')
cb.board_set_CmdList_DC(mawin, comlist)

cb.close_board(mawin)
pm.close_PM(mawin)

'''
//ITLA Cband 400G
itla_wr 0 0x32 0
itla_wr 0 0x34 0x02ee
->"Write itla OK!."
itla_wr 0 0x35 0xbf
->"Write itla OK!."
itla_wr 0 0x36 3125
->"Write itla OK!."
//itla_wr 0 0x30 0x05
//15dB
itla_wr 0 0x31 0x5dc
->"Write itla OK!."
\p
itla_wr 0 0x32 0x08
->"Write itla"
\p
\p
itla_wr 0 0x32 0x08
->"Write itla OK!."
\p

//first DRVout
cpld_wr 0x7f 0x9
set_adj 5 105
en_adj 5 1
\p

en_adj 5 0
\p

//VPN
cpld_wr 0x7f 0x9
en_adj 6 0
\p

//DRIVER VCC IN
cpld_wr 0x7f 0x8
set_adj 1 105
en_adj 1 0
\p

//TIA pwr
cpld_wr 0x7f 0x8
en_adj 3 0
en_adj 4 0

//断电
slave_off

//关ITLA
itla_wr 0 0x32 0

End.
'''