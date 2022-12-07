# -*- coding: utf-8 -*-
# Create date:3/17/2022
# Update on 4/6/2022 V1.4
# Update on 04/25/2022 :add ICR test functions
# Updated on 05/22/2022 V2.0: Finish the ICR test main function and save middle data
# author:Jiawen Zhou
import sys
import os
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
import pandas as pd
from pandas import DataFrame
from Main_win_GUI import Ui_MainWindow
#import other classes
import PowerMeter as pwm
import CtrlBoard as cb
import GOCA as goca
import ZVA as zva
import RFswitch as rfsw
import General_functions as gf
import Oscilloscope as osc
import ICRTestPlatform as icr
#others
import numpy as np
import operator
import time
import configparser
from ThreadCalient import DebugGui
import xlwings as xw
import serial
import smooth
#for phase and skew calculation
import Skew_phaseError_calculate as peSkew
#Plot the curve module
import matplotlib as mpl
import matplotlib.pyplot as plt

class main_test(QMainWindow, Ui_MainWindow):
    """
    Class documentation goes here.
    """

    def __init__(self, parent=None):
        """
        Constructor
        """
        global PM
        global CtrlB
        global ZVA
        global GOCA
        global RFSW
        global OScope
        global ICRtf
        global channel
        global temp
        global DSO

        self.PM     = object
        self.CtrlB  = object
        self.ZVA    = object
        self.GOCA   = object
        self.RFSW   = object
        self.OScope = object
        self.ICRtf  = object
        self.DSO    = object

        self.sn=''
        self.start=time.time()
        self.end=time.time()
        self.finalResult='None'
        self.test_flag='DC'
        self.test_type='Normal'
        self.gocflag=False #Aim to set timeout for GOCA, tigger the timeout
        self.log=''
        self.limit=0#暗电流较小
        self.limit_samp=1000 #取样点数默认1000

        #config which parameter to test in ICR test
        self.test_pe  = True#True
        self.test_bw  = True#True
        self.test_res = True#True

        self.config_path  = os.path.join(sys.path[0],'Configuration')
        self.report_path  = os.path.join(sys.path[0],'Test_report_TXDC')
        self.board_up     = os.path.join(self.config_path,'Setup_brdup_CtrlboardA001_20220113.txt')
        self.drv_up       = os.path.join(self.config_path,'Setup_driverup_Ctrlboard56017837A002_20211203.txt')
        self.drv_down     = os.path.join(self.config_path,'Setup_driverdown_CtrlboardA001_20210820.txt')
        self.config_file  = os.path.join(self.config_path,'config.ini')
        self.report_model = os.path.join(self.config_path,'test_report.xlsx')
        #self.report_judge= os.path.join(self.report_path,'test_report.xlsx')
        self.ITLA_config  = os.path.join(self.config_path,'ITLA1300pwr_Cband_ch75G_20220107T113903.csv')
        self.ICRpower_cal = os.path.join(self.config_path,'ICR_Light_Power_cal.csv')
        self.Tx_line_loss = os.path.join(self.config_path,'S21_lineIL.csv')
        self.Tx_pd_loss   = os.path.join(self.config_path,'S21_PD.csv')
        self.Rx_line_loss = os.path.join(self.config_path,'S21_RxlineIL20220516.csv')

        #Read the config file
        conf=configparser.ConfigParser()
        # print(sys.path)
        # print(self.config_file)
        # a=input('Waiting...')
        conf.read(self.config_file)
        sections=conf.sections()
        self.sw       = conf.get(sections[0],'SoftwareVersion')
        self.desk     = conf.get(sections[0],'TestDesk')
        self.temp     = conf.get(sections[0],'Temperature').split(',')
        self.channel  = conf.get(sections[0],'Channel').split(',')
        #perform full channel test
        if self.channel[0]=='all':
            self.channel=[i for i in range(9,73)]
        #self.ITLA_pwr = conf.get(sections[0],'ITLApower').split(',') #this is used for 3 channel test
        self.ITLA_pwr    = pd.read_csv(self.ITLA_config) #DataFrame format
        self.ICR_pwr     = pd.read_csv(self.ICRpower_cal) #DataFrame format
        self.PM_ch       = conf.get(sections[0],'PowerMeterChannel')
        self.ctrl_port   = conf.get(sections[0],'ControlBoardPort')
        self.pow_port    = conf.get(sections[0],'PowerMeterPort')
        self.RFSW_port   = conf.get(sections[0],'RFswPort')
        self.ZVA_port    = conf.get(sections[0],'ZVAport')
        self.GOCA_port   = conf.get(sections[0],'GOCAport')
        self.RFcal_path  = conf.get(sections[0],'RFcalPath')
        self.OScope_port = conf.get(sections[0],'OSCport')
        self.DSO_port    = 'IP:'+self.OScope_port.split('::')[1]
        self.ICRtf_port  = conf.get(sections[0],'ICRtfport')

        #Set the GUI
        QMainWindow.__init__(self, parent)
        self.setupUi(self)
        self.timer=QTimer(self) #Timer for the main GUI to monitor main test thread
        self.workthread=WorkThread()
        self.workthread.sig_progress.connect(self.update_pro)
        self.workthread.sig_but.connect(self.update_but)
        self.workthread.sig_print.connect(self.print_out_status)
        self.workthread.sig_status.connect(self.updata_status)
        self.workthread.sig_staColor.connect(self.updata_staColor)
        self.workthread.sig_clear.connect(self.text_clear)
        self.workthread.sig_goca.connect(self.goca_timer)
        self.timer.timeout.connect(self.periodly_check)
        self.start_test_butt_3.clicked.connect(self.debug_mode)
        self.start_test_butt_4.clicked.connect(self.openConfig)
        self.start_test_butt.clicked.connect(self.Start_test)
        self.read_SN.clicked.connect(self.readSN)
        #set up timer to monitor the GOCA equipment visit timeout
        self.timer1=QTimer(self)
        self.timer1.timeout.connect(self.periodly_checkEQ)
        #update the configuration on the label
        self.text_show='测试软件版本：{}\n测试机台号：{}\n测试温度:{}\n测试通道:{}\n光功率计通道:{}\n控制板串口号:{}\n功率计串口号：{}\nRF开关串口号：{}\nZVA端口号：{}\nGOCA端口号：{}\n示波器端口号：{}\nICR平台端口号：{}'.format(self.sw,self.desk
                                                                                         ,self.temp,self.channel,self.PM_ch,
                                                                                           self.ctrl_port,self.pow_port,self.RFSW_port,
                                                                                        self.ZVA_port,self.GOCA_port,self.OScope_port,self.ICRtf_port
                                                                                           )
        self.label_2.setText(self.text_show)

    def update_config(self):
        # self.text_show='测试软件版本：{}\n测试机台号：{}\n测试温度:{}\n测试通道:{}\n光功率计通道:{}\n控制板串口号:{}\n功率计串口号：{}\nRF开关串口号：{}\nZVA端口号：{}\nGOCA端口号：{}'.format(self.sw,self.desk
        #                                                                                                                                   ,self.temp,self.channel,self.PM_ch,
        #                                                                                                                                   self.ctrl_port,self.pow_port,self.RFSW_port,self.ZVA_port,self.GOCA_port
        #                                                                                                                                   )
        self.text_show='测试软件版本：{}\n测试机台号：{}\n测试温度:{}\n测试通道:{}\n光功率计通道:{}\n控制板串口号:{}\n功率计串口号：{}\nRF开关串口号：{}\nZVA端口号：{}\nGOCA端口号：{}\n示波器端口号：{}\nICR平台端口号：{}'.format(self.sw,self.desk,self.temp,self.channel,self.PM_ch,self.ctrl_port,self.pow_port,self.RFSW_port,self.ZVA_port,self.GOCA_port,self.OScope_port,self.ICRtf_port
                                                                                                                                                                      )
        self.label_2.setText(self.text_show)

    #read the SN from EEPROM
    def readSN(self):
        try:
            cb.open_board(self,self.ctrl_port)
            s=gf.read_eeprom(self)
            self.print_out_status('已获取SN:{}'.format(s))
            self.lineEdit.setText(s)
        except Exception as e:
            self.print_out_status(str(e))
            print(e)

    def openConfig(self):
        os.system("explorer.exe %s" % self.config_file)

    def debug_mode(self):
        print('Debug mode started...')
        self.debug_gui=DebugGui()
        self.debug_gui.pushButton.clicked.connect(self.passval_debug)
        self.debug_gui.pushButton_2.clicked.connect(self.setdrv_down)
        self.debug_gui.show()

    def setdrv_down(self):
        if not cb.open_board(self,self.ctrl_port):
            self.show()
            self.start_test_butt.setText('开始')
            self.test_status.setText('请检查控制板串口！')
            gf.status_color(self.test_status,'red')
            return
        cb.board_set(self,self.drv_down)
        self.print_out_status('手动下电完成！')
        cb.close_board(self)


    def passval_debug(self):
        '''
        #pass the value from the debug window to the main window
        :return:
        '''
        #config the channel
        if self.debug_gui.radioButton.isChecked():
            self.channel=['13']
        elif self.debug_gui.radioButton_2.isChecked():
            self.channel=['39']
        elif self.debug_gui.radioButton_3.isChecked():
            self.channel=['65']
        elif self.debug_gui.radioButton_4.isChecked():
            self.channel=['13','39','65']
        elif self.debug_gui.radioButton_9.isChecked():
            self.channel=[i for i in range(9,73)]

        #config the temperature
        if self.debug_gui.radioButton_7.isChecked():
            self.temp=['-5']
        elif self.debug_gui.radioButton_6.isChecked():
            self.temp=['25']
        elif self.debug_gui.radioButton_8.isChecked():
            self.temp=['75']
        elif self.debug_gui.radioButton_5.isChecked():
            self.temp=['25','-5','75']

        #config the ICR test items
        if self.debug_gui.checkBox_12.isChecked():
            self.test_pe=True
        else:
            self.test_pe=False

            self.test_res=False
        if self.debug_gui.checkBox_11.isChecked():
            self.test_bw=True
        else:
            self.test_bw=False
        if self.debug_gui.checkBox_13.isChecked():
            self.test_res=True
        else:
            self.test_res=False

        self.debug_gui.close()
        print('Test channels:{}\nTest temperature:{}'.format(self.channel,self.temp))
        print('ICR test items:\nPhase error={}\nRx BW={}\nResponsivity={}'.format(str(self.test_pe),str(self.test_bw),str(self.test_res)))
        self.update_config()




    def periodly_check(self):
        '''
        #test QTimer when workthread is not running recover GUI widgets and flags
        :return: NA
        '''
        if not self.workthread.isRunning():
            self.start_test_butt_3.setEnabled(True)
            self.start_test_butt_4.setEnabled(True)
            self.read_SN.setEnabled(True)
            self.Test_item.setEnabled(True)
            self.Test_type.setEnabled(True)
            if self.finalResult=='Pass':
                gf.status_color(self.test_status,'green')
                self.test_status.setText("Pass")
            elif self.finalResult=='Fail':
                self.test_status.setText("Fail")
                gf.status_color(self.test_status,'red')
                self.finalResult='None'
            else:
                #self.test_status.setText("Aborted")
                #gf.status_color(self.test_status,'blue')
                self.finalResult='None'
            self.progressBar.setValue(100)
            self.test_end()
            #self.start_test_butt.setText('开始')
            self.end=time.time()
            utime=str(time.strftime("%H:%M:%S", time.gmtime(self.end-self.start)))
            self.print_out_status('测试完成，用时：'+utime)
            print('测试完成，用时：',utime)
            self.timer.stop()


    def periodly_checkEQ(self,timeout=20):
        '''
        Test QTimer when timeout and no response from the equipment GOCA
        :param timeout:
        :return: NA
        '''
        t1=time.time()
        while not self.gocflag:
            t2=time.time()-t1
            if t2>timeout:
                self.workthread.terminate()
                self.timer1.stop()
                print('GOCA初始化失败...请检查GOCA Romote模式是否开启！')
                self.print_out_status('GOCA初始化失败...请检查GOCA Romote模式是否开启！')
                self.finalResult=False
                return
        self.print_out_status('GOCA初始化成功...')
        self.gocflag=False
        self.timer1.stop()

    #start the test
    @pyqtSlot()
    def Start_test(self):
        self.test_flag=self.Test_item.currentText()
        self.test_type=self.Test_type.currentText()
        self.finalResult=='None'
        if self.start_test_butt.text()=='开始':
            self.start_test_butt.setText('停止')
            self.start =time.time()
            self.timer.start(1000)
            self.start_test_butt_3.setEnabled(False)
            self.start_test_butt_4.setEnabled(False)
            self.read_SN.setEnabled(False)
            self.Test_item.setEnabled(False)
            self.Test_type.setEnabled(False)
            self.workthread.start()
            #print (self.workthread.isRunning())
        else:
            self.Stop_test()
            self.start_test_butt.setText('开始')

    @pyqtSlot()
    def Stop_test(self):
        status_show=''
        if self.workthread.isRunning():
            self.workthread.terminate()
            self.test_status.setText('测试停止!')
            self.print_out_status('测试停止!')
            gf.status_color(self.test_status,'blue')
            self.progressBar.setValue(0)
        else:
            self.test_status.setText('测试已经停止了!')
            gf.status_color(self.test_status,'blue')
        self.test_end()

    def test_end(self):
        '''
        define the end of the test
        :return: NA
        '''
        try:
            self.gocflag=False
            #Driver 下电，关闭串口连接
            self.start_test_butt.setText('开始')
            if self.test_flag=='DC':
                if serial.Serial.isOpen(self.CtrlB):
                    cb.board_set(self,self.drv_down)
                    self.print_out_status('Driver下电完成！')
                    cb.close_board(self)
                    self.print_out_status('测试板串口关闭完成！')
                pwm.close_PM(self)
            elif self.test_flag=='TxBW':
                if serial.Serial.isOpen(self.CtrlB):
                    cb.board_set(self,self.drv_down)
                    self.print_out_status('Driver下电完成！')
                    cb.close_board(self)
                    self.print_out_status('测试板串口关闭完成！')
                rfsw.close_RFSW(self)
                goca.close_GOCA(self)
                zva.close_ZVA(self)
            else:
                #ICR test close power off
                self.drv_down=os.path.join(self.config_path,'Setup_brddown_CtrlboardA001_20210820_ICR.txt')
                if serial.Serial.isOpen(self.CtrlB):
                    cb.board_set(self,self.drv_down)
                    self.print_out_status('Driver下电完成！')
                    cb.close_board(self)
                    self.print_out_status('测试板串口关闭完成！')
                self.ICRtf.write('OUTP1:CHAN1:STATE OFF');time.sleep(0.1)
                self.ICRtf.write('OUTP1:CHAN2:STATE OFF');time.sleep(0.1)
                self.print_out_status('Laser关闭成功！')
                osc.close_OSc(self)
                icr.close_ICR(self)
                self.print_out_status('示波器和测试平台连接关闭！')
            #self.test_status.setText('测试停止!')
        except Exception as e:
            print(e)
            self.print_out_status(str(e))

    def updata_status(self,f):
        mawin.test_status.setText(f)

    def updata_staColor(self,f):
        gf.status_color(self.test_status,f)

    def goca_timer(self):
        self.timer1.start(1000)# added

    #print the status into the GUi and update the log file and the screen
    def print_out_status(self,s=''):
        status_show=''
        status_show=gf.get_timestamp(0)+': '+s+'\n'
        self.log+=status_show
        print(status_show)
        self.plainTextEdit.appendPlainText(status_show)

    def update_but(self,f):
        self.start_test_butt.setText(f)

    #update progressbar
    def update_pro(self,s):
        self.progressBar.setValue(s)

    def text_clear(self):
        mawin.plainTextEdit.clear()

    #Only for program test
    def time_consuming(self):
        self.a=0
        #self.gocflag=False
        while self.a<100000000:
            self.a+=1
            if (self.a%100)==0:
                print(self.a)
        self.gocflag=True

    #RF test functions here
    '''
    balance the XY polorization
    '''

    def scan_RF(self,ch,ave=3):
        ''''
        function:scan RF from 1 to 10GHz to perform phase error test
        input:self,Current channel,average times
        :return:data array 30X4, every frequency test three times
        '''
        print('Phase Error calculating')
        #print('Phase Error calculating')
        data2=float(str(self.ICRtf.query('SOUR1:CHAN1:POW? SET')).strip())
        time.sleep(0.1)
        self.OScope.write('tdiv 500e-12')
        self.OScope.write('VBS app.Acquisition.Horizontal.HorOffset=0')
        data=[['']*4]*ave*10
        for i in range(10):
            #switch to test wavelength offset
            icr.set_wavelength(self,1,ch,i+1)
            ready=float(str(self.ICRtf.query('slot1:OPC?')).strip())
            while ready==0.0:
                time.sleep(0.2)
                ready=float(str(self.ICRtf.query('slot1:OPC?')).strip())
            #control the scope
            self.OScope.write('TRMD SINGLE')
            time.sleep(1)
            #query the data of amtiplitude
            A1=self.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')#;time.sleep(0.01)
            A2=self.OScope.query('VBS? return=app.Measure.P2.last.Result.Value')#;time.sleep(0.01)
            A3=self.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')#;time.sleep(0.01)
            A4=self.OScope.query('VBS? return=app.Measure.P4.last.Result.Value')#;time.sleep(0.01)

            while not icr.isfloat(A1) or not icr.isfloat(A2) or not icr.isfloat(A3) or not icr.isfloat(A4):
                A1=self.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')#;time.sleep(0.01)
                A2=self.OScope.query('VBS? return=app.Measure.P2.last.Result.Value')#;time.sleep(0.01)
                A3=self.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')#;time.sleep(0.01)
                A4=self.OScope.query('VBS? return=app.Measure.P4.last.Result.Value')#;time.sleep(0.01)
            while float(A1)/float(A3)>1.15 or float(A1)/float(A3)<0.85 or float(A1)/float(A4)>1.15 or float(A1)/float(A3)<0.87:
                time.sleep(0.1)
                icr.ICR_scamblePol(self)
                A1=self.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')#;time.sleep(0.01)
                A2=self.OScope.query('VBS? return=app.Measure.P2.last.Result.Value')#;time.sleep(0.01)
                A3=self.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')#;time.sleep(0.01)
                A4=self.OScope.query('VBS? return=app.Measure.P4.last.Result.Value')#;time.sleep(0.01)
                while not icr.isfloat(A1) or not icr.isfloat(A2) or not icr.isfloat(A3) or not icr.isfloat(A4):
                    A1=self.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')#;time.sleep(0.01)
                    A2=self.OScope.query('VBS? return=app.Measure.P2.last.Result.Value')#;time.sleep(0.01)
                    A3=self.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')#;time.sleep(0.01)
                    A4=self.OScope.query('VBS? return=app.Measure.P4.last.Result.Value')#;time.sleep(0.01)
            j=0
            while j<ave:
                self.OScope.write('TRMD SINGLE');time.sleep(0.5)
                osc.switch_to_DSO(self,self.OScope_port);time.sleep(0.5)
                data[i*ave+j]=osc.get_data_DSO(self)#data={'CH2':'','CH3':'','CH6':'','CH7':''}
                time.sleep(0.1)
                osc.switch_to_VISA(self,self.DSO_port);time.sleep(0.5)
                j+=1
            print('CE{} PE calculating...{}%'.format(str(ch),str(round((i+1)*10))))
        return data


    def scan_S21(self,ch,bw_start,bw_stop,bw_step):
        '''
        function:scan RF S21 from 1 to 40GHz to perform Rx BW test
        input:self
        output:[float(A1),float(A2),float(A3),float(A4)]
        '''
        print('S21 testing')
        self.OScope.write('TRMD AUTO')
        m=0
        num=round((bw_stop-bw_start)/bw_step)+1
        data=[[0.0]*4]*num
        for j in range(bw_start,bw_stop+1,bw_step):
            icr.set_wavelength(self,1,ch,j)
            ready=float(str(self.ICRtf.query('slot1:OPC?')).strip())
            while ready==0.0:
                time.sleep(0.2)
                ready=float(str(self.ICRtf.query('slot1:OPC?')).strip())
            self.OScope.write('tdiv 1e-9')
            time.sleep(1)
            #query the data of amtiplitude and balance the XY light distribute
            A1=self.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')#;time.sleep(0.01)
            A2=self.OScope.query('VBS? return=app.Measure.P2.last.Result.Value')#;time.sleep(0.01)
            A3=self.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')#;time.sleep(0.01)
            A4=self.OScope.query('VBS? return=app.Measure.P4.last.Result.Value')#;time.sleep(0.01)

            while not icr.isfloat(A1) or not icr.isfloat(A2) or not icr.isfloat(A3) or not icr.isfloat(A4):
                A1=self.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')#;time.sleep(0.01)
                A2=self.OScope.query('VBS? return=app.Measure.P2.last.Result.Value')#;time.sleep(0.01)
                A3=self.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')#;time.sleep(0.01)
                A4=self.OScope.query('VBS? return=app.Measure.P4.last.Result.Value')#;time.sleep(0.01)
            while (float(A1)/float(A3)>1.08 or float(A1)/float(A3)<0.92) and (float(A1)/float(A4)>1.08 or float(A1)/float(A3)<0.92):
            #while (float(A1)+float(A2))/(float(A3)+float(A4))>1.08 or (float(A1)+float(A2))/(float(A3)+float(A4))<0.92:
                time.sleep(0.1)
                icr.ICR_scamblePol(self)
                A1=self.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')#;time.sleep(0.01)
                A2=self.OScope.query('VBS? return=app.Measure.P2.last.Result.Value')#;time.sleep(0.01)
                A3=self.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')#;time.sleep(0.01)
                A4=self.OScope.query('VBS? return=app.Measure.P4.last.Result.Value')#;time.sleep(0.01)
                while not icr.isfloat(A1) or not icr.isfloat(A2) or not icr.isfloat(A3) or not icr.isfloat(A4):
                    A1=self.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')#;time.sleep(0.01)
                    A2=self.OScope.query('VBS? return=app.Measure.P2.last.Result.Value')#;time.sleep(0.01)
                    A3=self.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')#;time.sleep(0.01)
                    A4=self.OScope.query('VBS? return=app.Measure.P4.last.Result.Value')#;time.sleep(0.01)
            data[m]=[float(A1),float(A2),float(A3),float(A4)]
            m+=1
            print('CE{} Fre {}GHz S21 testing...{}%'.format(str(ch),str(j),str(round(((j-bw_start)/bw_step+1)/num*100,1))))
        return data

    def get_pd_resp_sig(self,ch):
        '''
        Get signal(Rx) port PD responsivity
        :return:[res_sig_x,res_sig_y,per_x,per_y,dark_x,dark_y]
        '''
        #load sig calibration data here:'D:\Fisilink\RX_Test\power_cal\power_sig.txt'
        if ch=='13':
            pwr=float(self.ICR_pwr.iloc[0,1])
        elif ch=='39':
            pwr=float(self.ICR_pwr.iloc[1,1])
        elif ch=='65':
            pwr=float(self.ICR_pwr.iloc[2,1])
        else:
            pwr=0
            print('Selected channel not in the calibration list, set power as 0dBm, please check')
        #Firstly get dark current and set the flags to determine the test point and time
        self.ICRtf.write('OUTP1:CHAN1:STATE OFF')
        self.ICRtf.write('OUTP1:CHAN2:STATE OFF')
        dark_x,dark_y=icr.TIA_getPDcurrent(self)[0:2] #[current_x,current_y,VPDX,VPDY]
        if dark_x>0.004 or dark_y>0.004:
            self.limit=1#暗电流较大
            self.limit_samp=1000#如果暗电流比较大，最多测500个点
        else:
            self.limit=0#暗电流较小
            self.limit_samp=1500

        self.ICRtf.write('OUTP1:CHAN1:STATE OFF');time.sleep(0.1)
        self.ICRtf.write( 'OUTP3:CHAN2:POW 6');time.sleep(1)#laser2切换为6dBm
        self.ICRtf.write('OUTP1:CHAN2:STATE ON')
        ready=float(str(self.ICRtf.query('slot1:OPC?')).strip())
        while ready==0.0:
            time.sleep(0.2)
            ready=float(str(self.ICRtf.query('slot1:OPC?')).strip())
        print('Laser set ready!')
        #自动旋转偏振态找最大最小电流
        cur_x,cur_y,per_x,per_y,PD_currentX,PD_currentY=self.ICR_scamblePol_XYmax_Auto()
        #calculate the responsivity
        #pwr=1,Responsivity unit:mA/mW
        res_sig_x=cur_x/(10**(pwr/10))/8
        res_sig_y=cur_y/(10**(pwr/10))/8
        print('Res_Sig_X,Res_Sig_Y,PER_X,PER_Y:\n',[res_sig_x,res_sig_y,per_x,per_y,dark_x,dark_y])
        return [res_sig_x,res_sig_y,per_x,per_y,dark_x,dark_y,PD_currentX,PD_currentY]

    def get_pd_resp_lo(self,ch):
        '''
        Get local port PD responsivity
        :parameter:Channel currently tested
        :return:[res_lo_x,res_lo_y]
        '''
        #load sig calibration data here:'D:\Fisilink\RX_Test\power_cal\power_lo.txt'
        if ch=='13':
            pwr=float(self.ICR_pwr.iloc[0,2])
        elif ch=='39':
            pwr=float(self.ICR_pwr.iloc[1,2])
        elif ch=='65':
            pwr=float(self.ICR_pwr.iloc[2,2])
        else:
            pwr=0
            print('Selected channel not in the calibration list, set power as 0dBm, please check')

        #Firstly get dark current and set the flags to determine the test point and time
        #self.ICRtf.write('OUTP1:CHAN1:STATE OFF')
        self.ICRtf.write('OUTP1:CHAN2:STATE OFF');time.sleep(0.1)
        self.ICRtf.write( 'OUTP3:CHAN1:POW 9');time.sleep(1)#laser1切换为9dBm
        self.ICRtf.write('OUTP1:CHAN1:STATE ON');time.sleep(0.1)
        ready=float(str(self.ICRtf.query('slot1:OPC?')).strip())
        while ready==0.0:
            time.sleep(0.2)
            ready=float(str(self.ICRtf.query('slot1:OPC?')).strip())
        print('Laser set ready!')
        cur_x,cur_y=icr.TIA_getPDcurrent_LO(self)[0:2]
        '''
        resp_x=current_X/(10^(power(find(ch400==i))/10))/4 %find(ch400==i)返回值为ch400中数值i对应的行/列，如果是二位矩阵，[m,n]=find(ch400==i)
        resp_y=current_Y/(10^(power(find(ch400==i))/10))/4
        '''
        #calculate the responsivity
        #pwr=1
        res_lo_x=cur_x/(10**(pwr/10))/4
        res_lo_y=cur_y/(10**(pwr/10))/4
        print('Res_LO_X,Res_LO_Y:\n',[res_lo_x,res_lo_y])
        return [res_lo_x,res_lo_y]

    def ICR_scamblePol_XYmax_Auto(self):
        '''
        Auto scamble mode to find XY max status
        :param:GUI Class obj
        :return:[current_X,current_Y,PER_X,PER_Y]
        '''
        try:
            #旋转偏振控制器，使得X/Y两路输出基本一致（相差10%）
            self.ICRtf.write(':POL2:MODE SCRAMBLE');time.sleep(0.1)
            self.ICRtf.write(':SCRAMBLE:FUNCTION SIN');time.sleep(0.1)
            self.ICRtf.write(':POL2:SCRAMBLE:FREQUENCY1 0.010');time.sleep(0.1)
            self.ICRtf.write(':POL2:SCRAMBLE:FREQUENCY2 0.018');time.sleep(0.1)
            self.ICRtf.write(':POL2:SCRAMBLE:FREQUENCY3 0.02');time.sleep(0.1)

            num=1
            current_X=0
            current_Y=0
            PD_current_X=[]
            PD_current_Y=[]
            cur_x,cur_y=[0,0]
            while current_X*current_Y==0:
                if self.limit==0:#暗电流正常（暗电流过大，limit_samp有值；暗电流很小，limit_samp为空）
                    cur_x,cur_y=icr.TIA_getPDcurrent(self)[0:2]
                    if cur_x<=0.004:
                        if current_Y==0:
                            current_Y=cur_y
                        else:
                            current_Y=max(current_Y,cur_y)
                        current_Xmin=cur_x
                        PER_X=10*np.log10(current_Y/current_Xmin)
                    if cur_y<=0.004:
                        if current_X==0:
                            current_X=cur_x
                        else:
                            current_X=max(current_X,cur_x)
                        current_Ymin=cur_y
                        PER_Y=10*np.log10(current_X/current_Ymin)
                    PD_current_X.append(cur_x)
                    PD_current_Y.append(cur_y)
                    num+=1
                    print('---Rx signal port test---\nTotal {} points, now at NO.{}\nCurrent_X:{}\nCurrent_Y:{}'.format(str(self.limit_samp),
                                                                                    str(num),str(cur_x),str(cur_y)))
                    if num>=self.limit_samp:
                        current_Xmin=min(PD_current_X)
                        num_minX=PD_current_X.index(current_Xmin)
                        current_Y=PD_current_Y[num_minX]

                        current_Ymin=min(PD_current_Y)
                        num_minY=PD_current_Y.index(current_Ymin)
                        current_X=PD_current_X[num_minY]

                        PER_X=10*np.log10(current_Y/current_Xmin)
                        PER_Y=10*np.log10(current_X/current_Ymin)
                        break
                else:#暗电流很大的情况
                    cur_x,cur_y=icr.TIA_getPDcurrent(self)[0:2]
                    PD_current_X.append(cur_x)
                    PD_current_Y.append(cur_y)
                    num+=1
                    print('---Rx signal port test---\nTotal {} points, now at NO.{}\nCurrent_X:{}\nCurrent_Y:{}'.format(str(self.limit_samp),
                                                                                                                        str(num),str(cur_x),str(cur_y)))

                    if num>=self.limit_samp:
                        current_Xmin=min(PD_current_X)
                        num_minX=PD_current_X.index(current_Xmin)
                        current_Y=PD_current_Y[num_minX]

                        current_Ymin=min(PD_current_Y)
                        num_minY=PD_current_Y.index(current_Ymin)
                        current_X=PD_current_X[num_minY]

                        PER_X=10*np.log10(current_Y/current_Xmin)
                        PER_Y=10*np.log10(current_X/current_Ymin)
                        break

            self.ICRtf.write(':POL2:MODE MANUAL')
            #work left to do to plot the curve and save the picture
            pass
            #work left to do to count the time
            pass
            return [current_X,current_Y,PER_X,PER_Y,PD_current_X,PD_current_Y]#,PD_current_X,PD_current_Y]

        except Exception as e:
            print(e)


class WorkThread(QThread):
    sig_progress=pyqtSignal(int)
    sig_status=pyqtSignal(str)
    sig_staColor=pyqtSignal(str)
    sig_print=pyqtSignal(str)
    sig_but=pyqtSignal(str)
    sig_clear=pyqtSignal()
    sig_goca=pyqtSignal(int)

    def __init__(self):
        super(WorkThread, self).__init__()

    def run(self):
        # try:
            if mawin.test_flag=="DC":
                self.DC_test()
                #self.test_unit()
            elif mawin.test_flag=="TxBW":
                #mawin.channel=[13,39,65]
                self.TxBW_test()
            elif mawin.test_flag=="ICR":
                #mawin.channel=[13,39,65]
                i=0
                while i<50:
                    print('执行第{}次循环测试中。。。共50次'.format(str(i+1)))
                    self.ICR_test()
                    i=i+1

            else:
                print('No correct test flag selected, this will not happen :)')
        # except Exception as e:
        #     print(e)
        #     self.sig_print.emit(str(e))
        #     self.sig_status.emit('Aborted!')
        #     self.sig_status.emit('blue')

    def DC_test(self):
        '''
        #DC test main process
        :return:
        '''
        #Get and judge the SN format
        sn=str(mawin.lineEdit.text()).strip()
        if not gf.SN_check(sn):
            #mawin.test_status.setText('SN输入有误，请检查SN！')
            self.sig_status.emit('SN输入有误，请检查SN！')
            self.sig_staColor.emit('red')
            self.sig_but.emit('开始')
            return
        self.sig_status.emit('测试进行中...')
        self.sig_staColor.emit('yellow')
        self.sig_clear.emit()
        self.sig_print.emit(sn)
        self.sig_progress.emit(5)

        ###设备连接
        if not cb.open_board(mawin,mawin.ctrl_port):
            self.sig_status.emit('请检查控制板串口！')
            self.sig_but.emit('开始')
            self.sig_staColor.emit('red')
            self.sig_progress.emit(0)
            return
        if not pwm.open_PM(mawin,mawin.pow_port):
            self.sig_status.emit('请检查功率计串口！')
            self.sig_but.emit('开始')
            self.sig_staColor.emit('red')
            self.sig_progress.emit(0)
            return

        ###connectivity test
        self.sig_progress.emit(10)
        self.sig_print.emit('检查pin脚连接中...')
        con=0
        con=cb.test_connectivity(mawin)
        if con==1:
            self.sig_print.emit('请检查，Driver未正确连接')
            self.sig_staColor.emit('blue')
            self.sig_but.emit('开始')
            self.sig_status.emit('请检查Driver连接!')
            #self.sig_stoptest.emit()
            self.sig_progress.emit(0)
            return
        elif con==2:
            self.sig_print.emit('请检查，TIA未正确连接')
            self.sig_staColor.emit('blue')
            self.sig_but.emit('开始')
            self.sig_status.emit('请检查TIA连接!')
            #self.sig_stoptest.emit()
            self.sig_progress.emit(0)
            return
        elif con==0:
            self.sig_print.emit('请检查，无正确返回值，连接性检查失败！')
            self.sig_staColor.emit('blue')
            self.sig_but.emit('开始')
            self.sig_status.emit('请检查器件连接!')
            #self.sig_stoptest.emit()
            self.sig_progress.emit(0)
            return
        '''
        *****Start the whole test steps*****
        '''
        self.sig_progress.emit(15)
        self.sig_print.emit('器件连接成功, FS400初始化中...')
        #Driver power up
        cb.board_set(mawin,mawin.drv_up)
        ##Considering add the status judgement here or in the board_set function

        #get noise power
        self.sig_print.emit('获取底噪(Get noise power)...')
        noipwr=cb.get_noipwr(mawin)
        self.sig_print.emit('获取底噪完成...')
        print('Noise power:',noipwr)

        #load ITLA power
        self.sig_progress.emit(20)

        #Test group by wavelength
        self.sig_print.emit('开始IL,PDL,bias测试...')
        ###if normal ITLA(Not C++ ITLA) then wavelength minus 8
        result_tmp=[]
        ##define the data frame to store the test data
        test_result=DataFrame(columns=('SN','TX_DC_DESK','TEMP','DATE','TIME',
                                       '400G_CH','TX_PDL','TX_IL','TX_Vbias_XI',
                                       'TX_Vbias_XQ','TX_Vbias_XP','TX_Vbias_YI','TX_Vbias_YQ','TX_Vbias_YP',
                                       'TX_ER_XI','TX_ER_XQ','TX_ER_XP','TX_ER_YI','TX_ER_YQ','TX_ER_YP',
                                       'TVpi_XI','TVpi_XQ','TVpi_XP','TVpi_YI','TVpi_YQ','TVpi_YP','TX_maxVbias_XI','TX_maxVbias_XQ',
                                       'TX_maxVbias_XP','TX_maxVbias_YI','TX_maxVbias_YQ','TX_maxVbias_YP'
                                       ))
        #Here is to add full channel test in the program
        #mawin.channel=[i for i in range(9,73)]
        for i in range(len(mawin.channel)):
            self.sig_print.emit("CH%s 测试开始...\n"%(str(mawin.channel[i])))
            data=np.zeros(6)
            data=cb.get_ER(mawin,str(int(mawin.channel[i])-8),noipwr[0],noipwr[1],1,1)
            if data==False or data==[]:
                self.sig_print.emit('ER获取失败，任务中止...')
                break
            self.sig_print.emit('ER获取成功...')
            print('max,min,abc,ER,Tvpi:\n',data)
            max=data[0][:]
            min=data[1][:]
            #get power meter reading of X-max Y-max power
            abc_ok=max[:]
            abc_tmp=np.zeros(6)
            while not operator.eq(abc_ok,abc_tmp):
                cb.set_abc(mawin,abc_ok)
                abc_tmp=cb.get_abc(mawin)
            pwr=pwm.read_PM(mawin,mawin.PM_ch)
            if pwr==0:
                self.sig_print.emit('功率获取失败,请检查串口连接...')
                return
            #get power meter reading of X-max Y-min power
            abc_ok=max[0:3]+min[3:6]
            abc_tmp=np.zeros(6)
            while not operator.eq(abc_ok,abc_tmp):
                cb.set_abc(mawin,abc_ok)
                abc_tmp=cb.get_abc(mawin)
            pwr_x=pwm.read_PM(mawin,mawin.PM_ch)
            if pwr_x==0:
                self.sig_print.emit('功率获取失败,请检查串口连接...')
                return
            #get power meter reading of X-min Y-max power
            abc_ok=min[0:3]+max[3:6]
            abc_tmp=np.zeros(6)
            while not operator.eq(abc_ok,abc_tmp):
                cb.set_abc(mawin,abc_ok)
                abc_tmp=cb.get_abc(mawin)
            pwr_y=pwm.read_PM(mawin,mawin.PM_ch)
            if pwr_y==0:
                self.sig_print.emit('功率获取失败,请检查串口连接...')
                return
            #set abc to max
            abc_ok=max[:]
            abc_tmp=np.zeros(6)
            while not operator.eq(abc_ok,abc_tmp):
                cb.set_abc(mawin,abc_ok)
                abc_tmp=cb.get_abc(mawin)
            self.sig_print.emit('PDL和IL获取成功...')
            CH=str(mawin.channel[i])
            PDL=pwr_x-pwr_y
            IL=pwr-float(mawin.ITLA_pwr.iloc[mawin.channel[i]-9,1])
            ABC=data[2]
            ER=data[3]
            Tvpi=data[4]
            Max=max[:]
            tt=[CH]+[PDL]+[IL]+ABC+ER+Tvpi+Max
            result_tmp.append(tt)
            print('CH:',CH)
            print('PDL:',PDL)
            print('IL:',IL)
            print('ABC:',ABC)
            print('ER:',ER)
            print('Tvpi:',Tvpi)
            print('Max:',Max)
            self.sig_progress.emit(round(20+(75/len(mawin.channel)*(i+1))))

        #Test data storage
        if not os.path.exists(mawin.report_path):
            os.mkdir(mawin.report_path)
        report_path1=os.path.join(mawin.report_path,mawin.test_flag)#create the child folder to store data
        if not os.path.exists(report_path1):
            os.mkdir(report_path1)
        report_path2=os.path.join(report_path1,mawin.test_type)#create the child folder to store data
        if not os.path.exists(report_path2):
            os.mkdir(report_path2)
        timestamp=gf.get_timestamp(1)
        config=[sn,mawin.desk,mawin.temp[0]]+timestamp.split('_')
        report_name=sn+'_'+timestamp+'.csv'
        report_judgename=sn+'_DC_Report_'+timestamp+'.xlsx'
        report_file=os.path.join(report_path2,report_name)
        report_judge=os.path.join(report_path2,report_judgename)
        print(report_file)
        for i in result_tmp:
            result_tmp[result_tmp.index(i)]=config+i
        #Write the result into data frame
        for i in range(len(mawin.channel)):
            test_result.loc[i]=result_tmp[i]
        #generate the report and print out the log file
        test_result.to_csv(report_file,index=False)
        self.sig_print.emit('测试完成!')
        self.sig_staColor.emit('green')
        self.sig_but.emit('开始')
        self.sig_status.emit('测试完成!')
        self.sig_progress.emit(100)
        print('测试完成')

        # Bob try to add
        cb.close_board(mawin)
        osc.close_OSc(mawin)
        icr.close_ICR(mawin)


        ##Write the data into report model and open the report after finished
        wb=xw.Book(mawin.report_model.replace('test_report.xlsx','test_report_DC.xlsx'))
        worksht=wb.sheets(1)
        worksht.activate()
        worksht.range((1,2)).value=test_result.iloc[0,0]
        worksht.range((2,2)).value=test_result.iloc[0,1]
        worksht.range((3,2)).value=test_result.iloc[0,2]
        worksht.range((4,2)).value=test_result.iloc[0,3]
        worksht.range((5,2)).value=test_result.iloc[0,4]
        worksht.range((8,2)).options(index=False,header=False,transpose=True).value=test_result.iloc[:,6:32]
        mawin.finalResult=worksht.range((6,2)).value
        wb.sheets(2).activate()
        wb.save(report_judge)

    def test_unit(self):
        '''
        #test function to verify the timer for goca timeout test
        :return:
        '''
        self.sig_goca.emit(20)
        mawin.time_consuming()#goca.init_GOCA(mawin)
        mawin.time_consuming()
        mawin.time_consuming()
        mawin.time_consuming()

    def TxBW_test(self):
        '''
        #Tx BW test main process
        :return:
        '''
        #Get and judge the SN format
        sn=str(mawin.lineEdit.text()).strip()
        if not gf.SN_check(sn):
            #mawin.test_status.setText('SN输入有误，请检查SN！')
            self.sig_status.emit('SN输入有误，请检查SN！')
            self.sig_staColor.emit('red')
            self.sig_but.emit('开始')
            return
        self.sig_status.emit('测试进行中...')
        self.sig_staColor.emit('yellow')
        self.sig_clear.emit()
        self.sig_print.emit(sn)
        self.sig_progress.emit(5)

        ###设备连接
        if not cb.open_board(mawin,mawin.ctrl_port):
            self.sig_status.emit('请检查控制板串口！')
            self.sig_but.emit('开始')
            self.sig_staColor.emit('red')
            self.sig_progress.emit(0)
            return
        if not rfsw.open_RFSW(mawin,mawin.RFSW_port):
            self.sig_status.emit('请检查RF Switch串口！')
            self.sig_but.emit('开始')
            self.sig_staColor.emit('red')
            self.sig_progress.emit(0)
            return
        if not goca.open_GOCA(mawin,mawin.GOCA_port):
            self.sig_status.emit('请检查GOCA光座端口(检查是否开启Remote模式)！')
            self.sig_but.emit('开始')
            self.sig_staColor.emit('red')
            self.sig_progress.emit(0)
            return
        if not zva.open_ZVA(mawin,mawin.ZVA_port):
            self.sig_status.emit('请检查ZVA端口！')
            self.sig_but.emit('开始')
            self.sig_staColor.emit('red')
            self.sig_progress.emit(0)
            return


        ###connectivity test
        self.sig_progress.emit(10)
        self.sig_print.emit('检查pin脚连接中...')
        con=0
        con=cb.test_connectivity(mawin)
        if con==1:
            self.sig_print.emit('请检查，Driver未正确连接')
            self.sig_staColor.emit('blue')
            self.sig_but.emit('开始')
            self.sig_status.emit('请检查Driver连接!')
            #self.sig_stoptest.emit()
            self.sig_progress.emit(0)
            return
        elif con==2:
            self.sig_print.emit('请检查，TIA未正确连接')
            self.sig_staColor.emit('blue')
            self.sig_but.emit('开始')
            self.sig_status.emit('请检查TIA连接!')
            #self.sig_stoptest.emit()
            self.sig_progress.emit(0)
            return
        elif con==0:
            self.sig_print.emit('请检查，无正确返回值，连接性检查失败！')
            self.sig_staColor.emit('blue')
            self.sig_but.emit('开始')
            self.sig_status.emit('请检查器件连接!')
            #self.sig_stoptest.emit()
            self.sig_progress.emit(0)
            return
        '''
        *****Start the whole test steps*****
        '''
        self.sig_progress.emit(15)
        self.sig_print.emit('器件连接成功, FS400初始化中...')
        #Driver power up
        cb.board_set(mawin,mawin.drv_up)
        ##Considering add the status judgement here or in the board_set function

        # #get calibration file path
        # rfcal=mawin.RFcal_path+''
        self.sig_print.emit('Driver上电完成，设备初始化中...')
        #initiate GOCA and ZVA
        self.sig_goca.emit(30)
        mawin.gocflag = goca.init_GOCA(mawin)
        if not zva.init_ZVA(mawin)[0]=='1':
            self.sig_print.emit('ZVA初始化失败...')
            return
        else:
            self.sig_print.emit('ZVA初始化完成...')
        self.sig_progress.emit(20)

        #Test group by wavelength
        self.sig_print.emit('开始TxBW测试...')
        ###if normal ITLA(Not C++ ITLA) then wavelength minus 8
        result_tmp=[]
        ##define the data frame to store the test data
        test_result=DataFrame(columns=('SN','TX_BW_DESK','TEMP','DATE','TIME','400G_CH',
                                       'TX_ROLLOFF_XI','TX_BW3DB_XI','TX_BW6DB_XI','Kink_XI','sigma_XI',
                                       'TX_ROLLOFF_XQ','TX_BW3DB_XQ','TX_BW6DB_XQ','Kink_XQ','sigma_XQ',
                                       'TX_ROLLOFF_YI','TX_BW3DB_YI','TX_BW6DB_YI','Kink_YI','sigma_YI',
                                       'TX_ROLLOFF_YQ','TX_BW3DB_YQ','TX_BW6DB_YQ','Kink_YQ','sigma_YQ'))
        #Test data storage
        report_path1=os.path.join(mawin.report_path,mawin.test_flag)#create the child folder to store data
        if not os.path.exists(report_path1):
            os.mkdir(report_path1)
        report_path2=os.path.join(report_path1,mawin.test_type)#create the child folder to store data
        if not os.path.exists(report_path2):
            os.mkdir(report_path2)
        timestamp=gf.get_timestamp(1)
        linecal=pd.read_csv(mawin.Tx_line_loss).iloc[:-1,:]
        pdcal=pd.read_csv(mawin.Tx_pd_loss).iloc[:-1,:]
        fre_ref=1.5e9
        for i in range(len(mawin.channel)):
            self.sig_print.emit("CH%s 测试开始...\n"%(str(mawin.channel[i])))
            data=np.zeros(6)
            data=cb.get_ER(mawin,str(int(mawin.channel[i])-8),10,10,0,0)
            if data==False or data==[]:
                self.sig_print.emit('ABC point获取失败，任务中止...')
                break
            self.sig_print.emit('ABC获取成功...')
            print('max,min,abc,ER,Tvpi:\n',data)
            # max=data[0]
            # min=data[1]
            abc=data[2][:]
            # #get power meter reading of X-max Y-max power
            abc_ok=abc[:]
            abc_tmp=np.zeros(6)
            while not operator.eq(abc_ok,abc_tmp):
                cb.set_abc(mawin,abc_ok)
                abc_tmp=cb.get_abc(mawin)
            self.sig_print.emit('Open EDFA')
            mawin.CtrlB.write(b'edfa_on\n')
            bias_ch=1;sw_ch=2#XI
            tt=[]
            for j in range(4):
                if j==0:
                    bias_ch=1;sw_ch=2#XI
                    self.sig_print.emit('start XI channel test')
                elif j==1:
                    bias_ch=2;sw_ch=3#XQ
                    self.sig_print.emit('start XQ channel test')
                elif j==2:
                    bias_ch=5;sw_ch=1#YI
                    self.sig_print.emit('start YI channel test')
                elif j==3:
                    bias_ch=4;sw_ch=4#YQ
                    self.sig_print.emit('start YQ channel test')
                zva.recall_cal(mawin,mawin.RFcal_path,sw_ch)#Load the calibration file according to the Date in the config file
                rfsw.switch_RFchannel(mawin,sw_ch)
                point=cb.get_quad(mawin,bias_ch,500,4000,15)
                if point<1000:
                    abc_ok[bias_ch-1]='0'+str(point)
                else:
                    abc_ok[bias_ch-1]=str(point)
                abc_tmp=np.zeros(6)
                while not operator.eq(abc_ok,abc_tmp):
                    cb.set_abc(mawin,abc_ok)
                    abc_tmp=cb.get_abc(mawin)
                time.sleep(8)
                filename=['']*4
                PCpath=''
                if '1'==zva.singleTest(mawin)[0]:
                    info=zva.savedata(mawin,sn,str(mawin.channel[i]),mawin.temp[0],timestamp,j+1)
                    if info[0][0]=='1':
                        filename[j]=info[1]
                        time.sleep(0.5)
                        r=zva.read_data(mawin,filename[j])
                        r=r.replace(';',',')
                        lis=r.split('\r\n')
                        lis1=[i.split(',')[:-1] for i in lis[:-1]]
                        df=pd.DataFrame(lis1[3:],columns=lis1[2]).astype('float')
                        df.to_csv(os.path.join(report_path2,filename[j]),index=None)
                        raw=df.iloc[:,3]
                        emb=raw-pdcal.iloc[:,1]-linecal.iloc[:,1]
                        fre=df.iloc[:,0]
                        ind=(abs(fre-fre_ref)).idxmin()
                        emb=emb-emb.iloc[ind]
                        #calculate the rolloff
                        rolloff=emb.iloc[5:10].mean()-emb.iloc[ind] #Use this one
                        #calculate the Kink
                        emb1=smooth.smooth(emb,30)
                        ss=np.diff(emb1)
                        kink=min(ss[149:3500])#to record the kink
                        sigma_ss=7*np.std(ss[149:3500])#to record the sigma
                        #calculate the BW
                        smo=smooth.smooth(emb,50)#to smooth the curve with 50 points window
                        aa=[x for x in range(smo.shape[0]) if smo[x]<-3]
                        if len(aa)==0:
                            BW3dB=0
                        else:
                            BW3dB=fre.iloc[min(aa)]/1e9
                        aa1=[x for x in range(smo.shape[0]) if smo[x]<-6]
                        if len(aa1)==0:
                            BW6dB=0
                        else:
                            BW6dB=fre.iloc[min(aa1)]/1e9
                        #BW6dB=fre.iloc[min(aa1)]/1e9
                        abc_ok=abc[:]
                        abc_tmp=np.zeros(6)
                        while not operator.eq(abc_ok,abc_tmp):
                            cb.set_abc(mawin,abc_ok)
                            abc_tmp=cb.get_abc(mawin)
                        tt.append([rolloff,BW3dB,BW6dB,kink,sigma_ss])
            tt1=[str(mawin.channel[i])]+tt[0]+tt[1]+tt[2]+tt[3]
            result_tmp.append(tt1)
            self.sig_progress.emit(round(20+(75/len(mawin.channel)*(i+1))))

        # timestamp=gf.get_timestamp(1)
        config=[sn,mawin.desk,mawin.temp[0]]+timestamp.split('_')
        report_name=sn+'_'+timestamp+'.csv'
        report_judgename=sn+'_TxBW_Report_'+timestamp+'.xlsx'
        report_file=os.path.join(report_path2,report_name)
        report_judge=os.path.join(report_path2,report_judgename)
        # print(report_file)
        for i in result_tmp:
            result_tmp[result_tmp.index(i)]=config+i
        #Write the result into data frame
        for i in range(len(mawin.channel)):
            test_result.loc[i]=result_tmp[i]
        #generate the report and print out the log file
        test_result.to_csv(report_file,index=False)
        #wb.close()
        #os.system("explorer "+report_judge)
        self.sig_print.emit('测试完成!')
        self.sig_staColor.emit('green')
        self.sig_but.emit('开始')
        self.sig_status.emit('测试完成!')
        self.sig_progress.emit(100)
        print('测试完成')
        ##Write the data into report model and open the report after finished
        wb=xw.Book(mawin.report_model.replace('test_report.xlsx','test_report_TxBW.xlsx'))
        worksht=wb.sheets(1)
        worksht.activate()
        worksht.range((1,2)).value=test_result.iloc[0,0]
        worksht.range((2,2)).value=test_result.iloc[0,1]
        worksht.range((3,2)).value=test_result.iloc[0,2]
        worksht.range((4,2)).value=test_result.iloc[0,3]
        worksht.range((5,2)).value=test_result.iloc[0,4]
        worksht.range((8,2)).options(index=False,header=False,transpose=True).value=test_result.iloc[:,6:26]
        mawin.finalResult=worksht.range((6,2)).value
        wb.save(report_judge)


    def ICR_test(self):
        '''
        #ICR test main process
        :return:
        '''
        #firstly rename the file of ICR test config of 'Board up,Drv up,Drv down'
        #Because no need to open ITLA, and need to power down PDs
        mawin.board_up     = os.path.join(mawin.config_path,'Setup_brdup_CtrlboardA001_20220113_ICR.txt')
        #mawin.drv_up       = os.path.join(mawin.config_path,'Setup_driverup_Ctrlboard56017837A002_20211203.txt')
        #mawin.drv_down     = os.path.join(mawin.config_path,'Setup_driverdown_CtrlboardA001_20210820.txt')
        #Get and judge the SN format
        sn=str(mawin.lineEdit.text()).strip()
        if not gf.SN_check(sn):
            #mawin.test_status.setText('SN输入有误，请检查SN！')
            self.sig_status.emit('SN输入有误，请检查SN！')
            self.sig_staColor.emit('red')
            self.sig_but.emit('开始')
            return
        self.sig_status.emit('测试进行中...')
        self.sig_staColor.emit('yellow')
        self.sig_clear.emit()
        self.sig_print.emit(sn)
        self.sig_progress.emit(5)

        ###设备连接
        if not cb.open_board(mawin,mawin.ctrl_port):
            self.sig_status.emit('请检查控制板串口！')
            self.sig_but.emit('开始')
            self.sig_staColor.emit('red')
            self.sig_progress.emit(0)
            return
        #####Oscilloscope and ICR tester connection
        if not osc.open_OScope(mawin,mawin.OScope_port):
            self.sig_status.emit('请检查示波器端口！')
            self.sig_but.emit('开始')
            self.sig_staColor.emit('red')
            self.sig_progress.emit(0)
            return

        if not icr.open_ICRtf(mawin,mawin.ICRtf_port):
            self.sig_status.emit('请检查ICR测试平台串口！')
            self.sig_but.emit('开始')
            self.sig_staColor.emit('red')
            self.sig_progress.emit(0)
            return

        ###connectivity test
        self.sig_progress.emit(10)
        self.sig_print.emit('检查pin脚连接中...')
        con=0
        con=cb.test_connectivity(mawin)
        if con==1:
            self.sig_print.emit('请检查，Driver未正确连接')
            self.sig_staColor.emit('blue')
            self.sig_but.emit('开始')
            self.sig_status.emit('请检查Driver连接!')
            #self.sig_stoptest.emit()
            self.sig_progress.emit(0)
            return
        elif con==2:
            self.sig_print.emit('请检查，TIA未正确连接')
            self.sig_staColor.emit('blue')
            self.sig_but.emit('开始')
            self.sig_status.emit('请检查TIA连接!')
            #self.sig_stoptest.emit()
            self.sig_progress.emit(0)
            return
        elif con==0:
            self.sig_print.emit('请检查，无正确返回值，连接性检查失败！')
            self.sig_staColor.emit('blue')
            self.sig_but.emit('开始')
            self.sig_status.emit('请检查器件连接!')
            #self.sig_stoptest.emit()
            self.sig_progress.emit(0)
            return
        '''
        *****Start the whole test steps*****
        '''
        self.sig_progress.emit(15)
        self.sig_print.emit('器件连接成功, FS400初始化, 配置TIA及上电...')
        #TIA Power on and config
        gf.TIA_on(mawin)
        gf.TIA_config(mawin)
        self.sig_print.emit('TIA上电完成，设备初始化中...')
        #initiate Oscilloscope and switch to visa connection
        osc.switch_to_DSO(mawin,mawin.OScope_port)
        print('DSO object created')
        osc.switch_to_VISA(mawin,mawin.DSO_port)
        if mawin.test_pe or mawin.test_bw:
            osc.init_OSc(mawin)
            self.sig_print.emit('示波器初始化成功...')
        self.sig_print.emit('开启并设置Laser波长...')
        icr.set_laser(mawin,1,15.4,10)
        icr.set_laser(mawin,2,15.4,-5,1550.000)
        self.sig_progress.emit(20)
        #Test group by wavelength
        self.sig_print.emit('配置完成!开始ICR测试...')
        ###if normal ITLA(Not C++ ITLA) then wavelength minus 8
        #result_tmp=[]
        ##define the data frame to store the test data
        test_result=DataFrame(columns=('SN','TX_BW_DESK','TEMP','DATE','TIME','400G_CH',
                                       'RX_BW_XI','RX_BW_XQ','RX_BW_YI','RX_BW_YQ','RX_PE_X',
                                       'RX_PE_Y','RX_X_Skew','RX_Y_Skew','RX_1PD_RES_LO_X','RX_1PD_RES_LO_Y',
                                       'RX_1PD_RES_SIG_X','RX_1PD_RES_SIG_Y','RX_PER_X','RX_PER_Y','Dark_X',
                                       'Dark_Y'))
        #Test data storage
        timestamp=gf.get_timestamp(1)
        if not os.path.exists(mawin.report_path):
            os.mkdir(mawin.report_path)
        report_path1=os.path.join(mawin.report_path,mawin.test_flag)#create the child folder to store data
        if not os.path.exists(report_path1):
            os.mkdir(report_path1)
        report_path2=os.path.join(report_path1,mawin.test_type)#create the child folder to store data
        if not os.path.exists(report_path2):
            os.mkdir(report_path2)
        #Add the folder named as SN_Date to store all the test data
        report_path3=os.path.join(report_path2,sn+'_'+timestamp)#create the child folder to store data
        if not os.path.exists(report_path3):
            os.mkdir(report_path3)
        #create list to store the result
        tt=[]
        for i in range(len(mawin.channel)):
            ch=str(mawin.channel[i])
            self.sig_print.emit("CH%s 测试开始...\n"%(ch))
            #ICR Test start and set the wavelength
            icr.set_wavelength(mawin,2,ch,0)
            icr.set_wavelength(mawin,1,ch,1)
            mawin.ICRtf.write('OUTP3:CHAN1:POW 10')
            mawin.ICRtf.write('OUTP3:CHAN2:POW -3')
            mawin.ICRtf.write('OUTP1:CHAN1:STATE ON')
            mawin.ICRtf.write('OUTP1:CHAN2:STATE ON')
            pe_x,pe_y,skew_x,skew_y=[0.0]*4
            XIBW3dB,XQBW3dB,YIBW3dB,YQBW3dB=[0.0]*4
            res_sig_x,res_sig_y,per_x,per_y,dark_x,dark_y,res_lo_x,res_lo_y=[0.0]*8
            if mawin.test_pe:
                self.sig_print.emit('Phase 和 Skew 测试中...')
                #phase error test
                lamb=icr.cal_wavelength(ch,0)
                m=3 #平均次数
                icr.ICR_scamblePol(mawin)
                #query the data of amtiplitude
                A1=mawin.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')
                A3=mawin.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')
                while not icr.isfloat(A1) or not icr.isfloat(A3):
                    A1=mawin.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')
                    A3=mawin.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')
                while float(A1)/float(A3)>1.08 or float(A1)/float(A3)<0.92:
                    time.sleep(0.2)
                    icr.ICR_scamblePol(mawin)
                    A1=mawin.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')
                    A3=mawin.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')
                    while not icr.isfloat(A1) or not icr.isfloat(A3):
                        A1=mawin.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')
                        A3=mawin.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')
                #Scan RF start
                data=mawin.scan_RF(ch)
                #save the result
                #pe_skew='{}_CH{}_PeSkew_{}'.format(sn,ch,timestamp)
                #pd.DataFrame(data).to_csv(pe_skew)
                pe_x,pe_y,skew_x,skew_y,new_pahseX,new_pahseY=peSkew.deal_with_data(data)
                #save raw and fitted s21 data
                raw_phase_report=os.path.join(report_path3,'{}_CH{}_phase_RawData_{}.csv'.format(sn,ch,timestamp))
                col=['Phase error','Skew']+[str(i)+'GHz' for i in range(1,11)]
                pha=pd.DataFrame([[pe_x,skew_x]+new_pahseX,[pe_y,skew_y]+new_pahseY],index=('phase_X','phase_Y'),columns=col)
                pha.to_csv(raw_phase_report)
                #work left here to draw the curve
                fre_1=[i for i in range(0,11)]
                y1=np.array(fre_1)*skew_x+pe_x+90
                y2=np.array(fre_1)*skew_y+pe_y+90
                #X plot fig
                fig_X=os.path.join(report_path3,'{}_CH{}_XIXQ_PeSkew_{}.png'.format(sn,ch,timestamp))
                fig, ax = plt.subplots()  # Create a figure containing a single axes.
                ax.plot(fre_1,y1,new_pahseX,'*')  # Plot some data on the axes.
                ax.set_title('XI-XQ Phase Error and Skew CH{}'.format(ch))
                ax.set_xlabel('Fre(GHz)')
                ax.set_ylabel('Angle [°]')
                fig.savefig(fig_X)
                #Y plot fig
                fig_Y=os.path.join(report_path3,'{}_CH{}_YIYQ_PeSkew_{}.png'.format(sn,ch,timestamp))
                fig1, ax1 = plt.subplots()  # Create a figure containing a single axes.
                ax1.plot(fre_1,y2,new_pahseY,'*')  # Plot some data on the axes.
                ax1.set_title('YI-YQ Phase Error and Skew CH{}'.format(ch))
                ax1.set_xlabel('Fre(GHz)')
                ax1.set_ylabel('Angle [°]')
                fig1.savefig(fig_Y)

                #fig.show()
                #fig1.show()

            #Work left to do here to calculate SKEW PE
            if mawin.test_bw:
                #Rx BW test
                self.sig_print.emit('Rx BW 测试中...')
                icr.set_wavelength(mawin,2,ch,0)
                icr.set_wavelength(mawin,1,ch,1)
                time.sleep(0.1)
                error=True
                eg=1 #表示在某一个波长处测试的次数
                while error:
                    icr.ICR_scamblePol(mawin)
                    #query the data of amtiplitude
                    A1=mawin.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')
                    A3=mawin.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')
                    while not icr.isfloat(A1) or not icr.isfloat(A3):
                        time.sleep(0.2)
                        A1=mawin.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')
                        A3=mawin.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')
                    while float(A1)/float(A3)>1.15 or float(A1)/float(A3)<0.85:
                        time.sleep(0.2)
                        icr.ICR_scamblePol(mawin)
                        A1=mawin.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')
                        A3=mawin.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')
                        while not icr.isfloat(A1) or not icr.isfloat(A3):
                            A1=mawin.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')
                            A3=mawin.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')
                    #scan S21
                    print('A1:{}\nA3:{}'.format(A1,A3))
                    data1=mawin.scan_S21(ch,1,40,1)
                    #config the frequency
                    numfre=40
                    da=[[0.0]*4]*numfre
                    fre=[i*1e9 for i in range(1,41)]
                    #归一化数据Y
                    raw=pd.DataFrame(data1)
                    for i in range(4):
                        raw.iloc[:,i]=20*np.log10(raw.iloc[:,i]/raw.iloc[0,i])
                    #线损
                    linelo=pd.read_csv(mawin.Rx_line_loss,encoding='gb2312')
                    fre_loss=linelo.iloc[:-1,0]
                    loss=linelo.iloc[:-1,1]
                    Yi=np.interp(fre,fre_loss,loss) #插值求得对应频率的线损
                    #减去线损
                    Ysmo=raw.sub(Yi,axis=0)
                    #save raw and fitted s21 data
                    raw_s21_report=os.path.join(report_path3,'{}_CH{}_RxBw_RawData_{}.csv'.format(sn,ch,timestamp))
                    fit_s21_report=os.path.join(report_path3,'{}_CH{}_RxBw_FitData_{}.csv'.format(sn,ch,timestamp))
                    colBW=['XI','XQ','YI','YQ']
                    indBW=[i for i in range(1,41)]
                    raw.columns=colBW
                    raw.index=indBW
                    raw.to_csv(raw_s21_report)#,index=False)
                    ###This point needs to be verified, whether the smooth data is equal to matlab gaussian method
                    for i in range(4):
                        tem=smooth.smooth(Ysmo.iloc[:,i],3)[1:-1]
                        tem=tem-tem[0] #归一化数据
                        Ysmo.iloc[:,i]=tem
                    Ysmo.columns=colBW
                    Ysmo.index=indBW
                    Ysmo.to_csv(fit_s21_report)#,index=False)
                    #draw the curves
                    #work left here to draw the curve
                    x_bw=[i for i in range(1,41)]
                    fig_bw, ax_bw = plt.subplots()  # Create a figure containing a single axes.
                    ax_bw.plot(x_bw,Ysmo.iloc[:,0],label='XI')
                    ax_bw.plot(x_bw,Ysmo.iloc[:,1],label='XQ')
                    ax_bw.plot(x_bw,Ysmo.iloc[:,2],label='YI')
                    ax_bw.plot(x_bw,Ysmo.iloc[:,3],label='YQ')# Plot some data on the axes.
                    ax_bw.set_title('Rx Band Width CH{}'.format(ch))
                    ax_bw.set_xlabel('Fre(GHz)')
                    ax_bw.set_ylabel('Loss(dB)')
                    ax_bw.legend()
                    #fig_bw.show()
                    fig_bwPic=os.path.join(report_path3,'{}_CH{}_RxBW_{}.png'.format(sn,ch,timestamp))
                    fig_bw.savefig(fig_bwPic)

                    # #calculate the result, base on Ysmo is a list type
                    # for i in range(4):
                    #     aa=[x for x in range(len(Ysmo[i])) if Ysmo[i][x]<-3]
                    #     if len(aa)==0:
                    #         BW3dB=40
                    #     else:
                    #         BW3dB=min(aa)#/1e9
                    #     if i==0:
                    #         XIBW3dB=BW3dB
                    #     elif i==1:
                    #         XQBW3dB=BW3dB
                    #     elif i==2:
                    #         YIBW3dB=BW3dB
                    #     elif i==3:
                    #         YQBW3dB=BW3dB
                    #
                    # YsmoMean=Ysmo.mean(axis=1)
                    # aa=[x for x in range(len(YsmoMean)) if YsmoMean[x]<-3]
                    # if len(aa)==0:
                    #     MeanBW3dB=40
                    # else:
                    #     MeanBW3dB=min(aa)#/1e9

                    #calculate the result, base on Ysmo is a dataframe type
                    for i in range(4):
                        aa=[x for x in range(Ysmo.shape[0]) if Ysmo.iloc[x,i]<-3]
                        if len(aa)==0:
                            BW3dB=40
                        else:
                            BW3dB=min(aa)#+1#/1e9 index plus 1 as Freq
                        #each channel
                        if i==0:
                            XIBW3dB=BW3dB
                            if min(Ysmo.iloc[:,i])<-20.0:
                                XIBW3dB=0.0
                        elif i==1:
                            XQBW3dB=BW3dB
                            if min(Ysmo.iloc[:,i])<-20.0:
                                XQBW3dB=0.0
                        elif i==2:
                            YIBW3dB=BW3dB
                            if min(Ysmo.iloc[:,i])<-20.0:
                                YIBW3dB=0.0
                        elif i==3:
                            YQBW3dB=BW3dB
                            if min(Ysmo.iloc[:,i])<-20.0:
                                YQBW3dB=0.0

                    YsmoMean=Ysmo.mean(axis=1)
                    aa=[x for x in range(YsmoMean.shape[0]) if YsmoMean[x+1]<-3]
                    if len(aa)==0:
                        MeanBW3dB=40
                    else:
                        MeanBW3dB=min(aa)#+1#/1e9
                        if min(YsmoMean)<-20:
                            MeanBW3dB=0

                    #Need to check whether to perform the judgement of >28G Hz and each channel > mean bw by 3GHz
                    #if not meet the requirement then perform the retest
                    #left to write the result
                    error=False #Not check BW and retest
            #TIA amplitude test not performed
            #scan_Amp:to test the TIA output amplitude
            #left blank here
            if mawin.test_res:
                #Responsivity test
                # self.sig_print.emit('响应度测试中...')
                mawin.CtrlB.write(b'switch_set 10 0\n');time.sleep(0.1)
                mawin.CtrlB.write(b'cpld_spi_wr 0x2c 2700\n');time.sleep(0.1)
                mawin.CtrlB.write(b'cpld_spi_wr 0x2f 2700\n');time.sleep(0.1)
                self.sig_print.emit('Rx响应度测试中...')
                res_sig_x,res_sig_y,per_x,per_y,dark_x,dark_y,PD_currentX,PD_currentY=mawin.get_pd_resp_sig(ch)
                self.sig_print.emit('LO响应度测试中...')
                res_lo_x,res_lo_y=mawin.get_pd_resp_lo(ch)

                #Draw the Rx PD current data point and save the image
                #Save the raw data
                res_raw=os.path.join(report_path3,'{}_CH{}_RxPDcurrentXY_{}.csv'.format(sn,ch,timestamp))
                raw_res=pd.DataFrame([PD_currentX,PD_currentY]).transpose()
                raw_res.columns=['PD_currentX','PD_currentY']
                raw_res.to_csv(res_raw)
                #Draw the curve X Y phase
                x_px=[i for i in range(1,len(PD_currentX)+1)]
                fig_px, ax_px = plt.subplots()  # Create a figure containing a single axes.
                ax_px.plot(x_px,PD_currentX,label='Current X')
                ax_px.plot(x_px,PD_currentY,label='Current Y')
                ax_px.set_title('Rx PD current CH{}'.format(ch))
                ax_px.set_xlabel('Points')
                ax_px.set_ylabel('Current(mA)')
                ax_px.legend()
                #fig_px.show()
                fig_pxPic=os.path.join(report_path3,'{}_CH{}_Rx_PDcurrentXY_{}.png'.format(sn,ch,timestamp))
                fig_px.savefig(fig_pxPic)

            self.sig_progress.emit(round(20+(75/len(mawin.channel)*(i+1))))
            tt.append([ch,XIBW3dB,XQBW3dB,YIBW3dB,YQBW3dB,pe_x,pe_y,skew_x,skew_y,res_lo_x,res_lo_y,
                       res_sig_x,res_sig_y,per_x,per_y,dark_x,dark_y])
        #close the laser output
        mawin.ICRtf.write('OUTP1:CHAN1:STATE OFF')
        mawin.ICRtf.write('OUTP1:CHAN2:STATE OFF')

        #'SN','TX_BW_DESK','TEMP','DATE','TIME','400G_CH'
        # # timestamp=gf.get_timestamp(1)
        config=[sn,mawin.desk,mawin.temp[0]]+timestamp.split('_')
        report_name=sn+'_'+timestamp+'.csv'
        report_judgename=sn+'_ICR_Report_'+timestamp+'.xlsx'
        report_file=os.path.join(report_path3,report_name)
        report_judge=os.path.join(report_path3,report_judgename)
        # print(report_file)
        for i in tt:
            tt[tt.index(i)]=config+i
        #Write the result into data frame
        for i in range(len(mawin.channel)):
            test_result.loc[i]=tt[i]
        #generate the report and print out the log file
        test_result.to_csv(report_file,index=False)
        # print('Close the figs')
        # plt.close(fig)
        # plt.close(fig1)
        # plt.close(fig_bw)
        # plt.close(fig_px)

        #wb.close()
        #os.system("explorer "+report_judge)
        # self.sig_print.emit('测试完成!')
        # self.sig_staColor.emit('green')
        # self.sig_but.emit('开始')
        # self.sig_status.emit('测试完成!')
        # self.sig_progress.emit(100)
        # print('测试完成')

        #close the equipment connection
        cb.close_board(mawin)
        icr.close_ICR(mawin)
        osc.close_OSc(mawin)

        ##Write the data into report model and open the report after finished
        # wb=xw.Book(mawin.report_model.replace('test_report.xlsx','test_report_ICR.xlsx'))
        # worksht=wb.sheets(1)
        # worksht.activate()
        # worksht.range((1,2)).value=test_result.iloc[0,0]
        # worksht.range((2,2)).value=test_result.iloc[0,1]
        # worksht.range((3,2)).value=test_result.iloc[0,2]
        # worksht.range((4,2)).value=test_result.iloc[0,3]
        # worksht.range((5,2)).value=test_result.iloc[0,4]
        # worksht.range((8,2)).options(index=False,header=False,transpose=True).value=test_result.iloc[:,6:]
        # mawin.finalResult=worksht.range((6,2)).value
        # wb.save(report_judge)

if __name__ == "__main__":
    app=QApplication(sys.argv)
    mawin = main_test()
    mawin.show()
    sys.exit(app.exec_())