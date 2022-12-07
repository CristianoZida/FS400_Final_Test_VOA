# -*- coding: utf-8 -*-
# Create date:3/17/2022
# Update on 4/6/2022 V1.4
# Update on 04/25/2022 :add ICR test functions
# Updated on 05/22/2022 V2.0: Finish the ICR test main function and save middle data
# updated on 07/20/2022: gf.write for ch64 only; replace Test_ER_retest.py (by Bob)
# author:Jiawen Zhou
import sys
import os
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
from GUI.Main_win_GUI import Ui_MainWindow
from GUI.ThreadCalient import DebugGui

#equipment control related classes
import Instruments.PowerMeter as pwm
import Instruments.CtrlBoard as cb
import Instruments.GOCA as goca
import Instruments.ZVA as zva
import Instruments.RFswitch as rfsw
import Common_functions.General_functions as gf
import Instruments.Oscilloscope as osc
import Instruments.ICRTestPlatform as icr
import Instruments.MPC_GBS as mpc
import Instruments.JinLi_test as jl
import serial
#calculation and files handle related classes
import numpy as np
import operator
import time
import configparser
import pandas as pd
from pandas import DataFrame
import xlwings as xw
import Common_functions.smooth as smooth
import shutil
import re
import math
#for phase and skew calculation
import Common_functions.Skew_phaseError_calculate as peSkew
#calculate and judge the ER retest condition in DC test
import Common_functions.Test_ER_retest as erRetest
#Plot the curve module
import matplotlib as mpl
import matplotlib.pyplot as plt

class main_test(QMainWindow, Ui_MainWindow):
    """
    Main window of FS400 final test program
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
        self.MPC    = object
        self.JinLi  = object

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
        #for curve plot
        self.fig=object
        self.fig1=object
        self.fig_px=object
        self.fig_px_1=object
        self.fig_bw=object
        #define device type
        self.device_type='IDT'
        self.ITLA_type='C80'

        #define the resistance of Tx MPD circuit
        self.R_txmpdx=27e3 #Om
        self.R_txmpdy=27.2e3 #Om

        #define the frequency to calculate the Tx BW 3dB/6dB loss
        self.cal_fre=1.5e9

        #C++ and C band wavelength:
        sta1 = 190.7125
        end1 = 196.6375
        self.wl_cpp=np.linspace(sta1,end1,80)

        sta1 = 191.3125
        end1 = 196.0375
        self.wl_c = np.linspace(sta1, end1, 64)

        #VOA calibration range:
        sta1 = 190.1125
        end1 = 197.2375
        self.voa_wl=np.linspace(sta1,end1,round((end1-sta1)/0.075)+1).round(4)

        #config which parameter to test in ICR test
        self.test_pe  = True#True
        self.test_bw  = True#True
        self.test_res = True#True

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

        self.config_path = os.path.join(script_path, 'Configuration')
        # self.report_path  = os.path.join(sys.path[0],'Test_report_TXDC')
        self.report_path = os.path.join('C:\\', 'Test_result')
        self.config_file = os.path.join(self.config_path, 'config.ini')
        self.report_model = os.path.join(self.config_path, 'test_report.xlsx')

        #Read the config file
        conf=configparser.ConfigParser()
        conf.read(self.config_file)
        sections=conf.sections()
        self.sw          = conf.get(sections[0],'SoftwareVersion')
        self.desk        = conf.get(sections[0],'TestDesk')
        self.temp        = conf.get(sections[0],'Temperature').split(',')
        self.channel     = conf.get(sections[0],'Channel').split(',')
        self.ITLA_type   = conf.get(sections[0], 'ITLA')
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
        self.MPC_port    = conf.get(sections[0],'MPCport')
        self.JinLi_port  = conf.get(sections[0],'JINLIport')

        #config default driver up and down configuration file
        self.drv_up = os.path.join(self.config_path, 'Setup_driverup_Ctrlboard56017837A002_20211203.txt')
        self.drv_down = os.path.join(self.config_path, 'Setup_driverdown_CtrlboardA001_20210820.txt')

        if self.ITLA_type=='C80':
            self.board_up = os.path.join(self.config_path, 'Setup_brdup_CtrlboardA001_20220113_C80.txt')
            self.ITLA_config = os.path.join(self.config_path,
                                        'ITLA1300pwr_Cband_ch75G_C80.csv')  # 待修改 20220722
        elif self.ITLA_type=='C64':
            self.board_up = os.path.join(self.config_path, 'Setup_brdup_CtrlboardA001_20220113.txt')
            self.ITLA_config = os.path.join(self.config_path,
                                            'ITLA1300pwr_Cband_ch75G_20220107T113903.csv')
        else:
            print('ITLA type 配置错误，请填写\'C80\'或\'C64\，5s后退出程序。。。')
            for i in range(5):
                time.sleep(1)
                print(i + 1)
            sys.exit()

        self.ICRpower_cal = os.path.join(self.config_path, 'ICR_Light_Power_cal.csv')
        self.Res_power_cal0 = os.path.join(self.config_path, 'Res_Optical_calibration_ITLA_CH0_C++_20220927_103321.csv')
        self.Res_power_cal1 = os.path.join(self.config_path, 'Res_Optical_calibration_ITLA_CH1_C++_20220927_112730.csv')
        self.Tx_line_loss = os.path.join(self.config_path, 'S21_lineIL.csv')
        self.Tx_pd_loss = os.path.join(self.config_path, 'S21_PD.csv')
        self.Rx_line_loss = os.path.join(self.config_path, 'S21_RxlineIL20220516.csv')

        #Read DC ITLA optical calibration and ICR optical calibration
        self.ITLA_pwr    = pd.read_csv(self.ITLA_config)
        self.ICR_pwr     = pd.read_csv(self.ICRpower_cal)
        self.Res_ITLA_pwr_0 = pd.read_csv(self.Res_power_cal0)
        self.Res_ITLA_pwr_1 = pd.read_csv(self.Res_power_cal1)

        #judge if all the configuration file are exists
        #work left here to do

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
        #Test plot function
        self.workthread.sig_plot.connect(self.plotTest)
        self.workthread.sig_plotClose.connect(self.plotClose)
        self.timer.timeout.connect(self.periodly_check)
        self.start_test_butt_3.clicked.connect(self.debug_mode)
        self.start_test_butt_4.clicked.connect(self.openConfig)
        self.start_test_butt.clicked.connect(self.Start_test)
        self.read_SN.clicked.connect(self.readSN)
        #set up timer to monitor the GOCA equipment visit timeout
        self.timer1=QTimer(self)
        self.timer1.timeout.connect(self.periodly_checkEQ)
        self.setWindowTitle('FS400 Final Test {}'.format(self.sw))
        #update the configuration on the label
        self.text_show = '软件版本:{}\n机台编号:{}\nITLA类型:{}\n测试温度:{}\n' \
                         '测试通道:{}\n光功率计通道:{}\n控制板串口:{}\n功率计串口:{}\n' \
                         'RF开关串口:{}\n偏振控制器串口:{}\n锦鲤串口:{}\nZVA端口:{}\n' \
                         'GOCA端口:{}\n示波器端口:{}\nMTP端口:{}'.format(self.sw, self.desk,
                                                                         self.ITLA_type, self.temp,
                                                                         self.channel, self.PM_ch,
                                                                         self.ctrl_port, self.pow_port,
                                                                         self.RFSW_port, self.MPC_port,
                                                                         self.JinLi_port, self.ZVA_port,
                                                                         self.GOCA_port, self.OScope_port,
                                                                         self.ICRtf_port)
        self.label_2.setText(self.text_show)

    def update_config(self):
       '''
       Update test config information in the GUI
       :return: NA
       '''
       self.text_show = '软件版本:{}\n机台编号:{}\nITLA类型:{}\n测试温度:{}\n' \
                        '测试通道:{}\n光功率计通道:{}\n控制板串口:{}\n功率计串口:{}\n' \
                        'RF开关串口:{}\n偏振控制器串口:{}\n锦鲤串口:{}\nZVA端口:{}\n' \
                        'GOCA端口:{}\n示波器端口:{}\nMTP端口:{}'.format(self.sw, self.desk,
                                                                        self.ITLA_type, self.temp,
                                                                        self.channel, self.PM_ch,
                                                                        self.ctrl_port, self.pow_port,
                                                                        self.RFSW_port, self.MPC_port,
                                                                        self.JinLi_port, self.ZVA_port,
                                                                        self.GOCA_port, self.OScope_port,
                                                                        self.ICRtf_port)
       self.label_2.setText(self.text_show)

    #Close the plots
    def plotClose(self):
        try:
            plt.close('all')
            if self.test_pe:
                plt.close(self.fig)
                plt.close(self.fig1)
            if self.test_bw:
                plt.close(self.fig_bw)
            if self.test_res:
                plt.close(self.fig_px)
                plt.close(self.fig_px_1)
        except Exception as e:
            print(e)

    def plotTest(self, df, f, ch, sn, timetmp, typ,ind_markYmin=[],ind_markXmin=[]):
        '''
        this is to plot and save the drawings in the ICR test
        :param df: data to draw
        :param f: file name to store
        :param ch: channel to test
        :param sn: Serial Number
        :param timetmp: Test time to record
        :param typ: Test type selected
        :param ind_markYmin: Rx res test find Ymin process, qty of each process
        :param ind_markXmin: Rx res test find Xmin process, qty of each process
        :return:
        '''
        if typ == 'PE':
            fre_1 = [i for i in range(1, 11)]
            skew_x = df.iloc[0, :][1]
            skew_y = df.iloc[1, :][1]
            pe_x = df.iloc[0, :][0]
            pe_y = df.iloc[1, :][0]
            new_pahseX = df.iloc[0, :][2:]
            new_pahseY = df.iloc[1, :][2:]
            y1 = np.array(fre_1) * skew_x + pe_x + 90
            y2 = np.array(fre_1) * skew_y + pe_y + 90
            # X plot fig
            fig_X = os.path.join(f, '{}_CH{}_XIXQ_PeSkew_{}.png'.format(sn, ch, timetmp))
            self.fig, ax = plt.subplots()  # Create a figure containing a single axes.
            ax.plot(fre_1, new_pahseX, '*')  # Plot some data on the axes.
            ax.plot(fre_1, y1)
            ax.set_title('XI-XQ Phase Error and Skew CH{}'.format(ch))
            ax.set_xlabel('Fre(GHz)')
            ax.set_ylabel('Angle [°]')
            self.fig.show()
            self.fig.savefig(fig_X)
            # Y plot fig
            fig_Y = os.path.join(f, '{}_CH{}_YIYQ_PeSkew_{}.png'.format(sn, ch, timetmp))
            self.fig1, ax1 = plt.subplots()  # Create a figure containing a single axes.
            ax1.plot(fre_1, new_pahseY, '*')  # Plot some data on the axes.
            ax1.plot(fre_1, y2)
            ax1.set_title('YI-YQ Phase Error and Skew CH{}'.format(ch))
            ax1.set_xlabel('Fre(GHz)')
            ax1.set_ylabel('Angle [°]')
            self.fig1.show()
            self.fig1.savefig(fig_Y)
        elif typ == 'RES':
            x_px = [i for i in range(1, df.shape[0] + 1)]
            self.fig_px, ax_px = plt.subplots()  # Create a figure containing a single axes.
            ax_px.plot(x_px, df.iloc[:, 0], label='Current X')
            ax_px.plot(x_px, df.iloc[:, 1], label='Current Y')
            axvertical=ind_markYmin+[i+ind_markYmin[-1] for i in ind_markXmin]
            for t in axvertical:
                if t==ind_markYmin[-1]:
                    ax_px.axvline(t, color='r')
                else:
                    ax_px.axvline(t,ls='--',color='g')
            ax_px.set_title('Rx PD current CH{}'.format(ch))
            ax_px.set_xlabel('Points')
            ax_px.set_ylabel('Current(mA)')
            ax_px.legend()
            self.fig_px.show()
            fig_pxPic = os.path.join(f, '{}_CH{}_Rx_PDcurrentXY_{}.png'.format(sn, ch, timetmp))
            self.fig_px.savefig(fig_pxPic)
        elif typ == 'RES_LO':
            x_px = [i for i in range(1, df.shape[0] + 1)]
            self.fig_px_1, ax_px = plt.subplots()  # Create a figure containing a single axes.
            ax_px.plot(x_px, df.iloc[:, 0], label='Current X')
            ax_px.plot(x_px, df.iloc[:, 1], label='Current Y')
            ax_px.set_title('LO PD current CH{}'.format(ch))
            ax_px.set_xlabel('Points')
            ax_px.set_ylabel('Current(mA)')
            ax_px.legend()
            self.fig_px_1.show()
            fig_pxPic = os.path.join(f, '{}_CH{}_LO_PDcurrentXY_{}.png'.format(sn, ch, timetmp))
            self.fig_px_1.savefig(fig_pxPic)
        elif typ == 'BW':
            Ysmo = df
            x_bw = [i for i in range(1, 41)]
            self.fig_bw, ax_bw = plt.subplots()  # Create a figure containing a single axes.
            ax_bw.plot(x_bw, Ysmo.iloc[:, 0], label='XI')
            ax_bw.plot(x_bw, Ysmo.iloc[:, 1], label='XQ')
            ax_bw.plot(x_bw, Ysmo.iloc[:, 2], label='YI')
            ax_bw.plot(x_bw, Ysmo.iloc[:, 3], label='YQ')  # Plot some data on the axes.
            ax_bw.axhline(0,color='k')
            ax_bw.axhline(-3,ls='--')
            ax_bw.set_title('Rx Band Width CH{}'.format(ch))
            ax_bw.set_xlabel('Fre(GHz)')
            ax_bw.set_ylabel('Loss(dB)')
            ax_bw.legend()
            self.fig_bw.show()
            fig_bwPic = os.path.join(f, '{}_CH{}_RxBW_{}.png'.format(sn, ch, timetmp))
            self.fig_bw.savefig(fig_bwPic)

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

    # write the SN to EEPROM
    def writeSN(self):
        try:
            sn=self.lineEdit.text()
            cb.open_board(self, self.ctrl_port)
            gf.write_SN(self,sn)
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
        self.sn=self.lineEdit.text()
        self.debug_gui.pushButton_3.clicked.connect(self.writeSN)
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
            if self.ITLA_type == 'C80':
                self.channel=[i for i in range(1,81)]
            elif self.ITLA_type=='C64':
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
                self.finalResult = 'None'
            elif self.finalResult=='Fail':
                self.test_status.setText("Fail")
                gf.status_color(self.test_status,'red')
                self.finalResult='None'
            else:
                gf.status_color(self.test_status,'blue')
                self.finalResult='None'
            self.progressBar.setValue(100)
            self.test_end()
            self.end=time.time()
            utime=str(time.strftime("%H:%M:%S", time.gmtime(self.end-self.start)))
            self.print_out_status('测试完成，用时：'+utime)
            print('测试完成，用时：',utime)
            self.timer.stop()
            try:
                sys.stdout.CloseLogFile()
                sys.stdout.log_enabled = False
            except Exception as e:
                print(e)

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
        #Config the channel according to test type selected
        if self.test_type=='Normal':
            if self.test_flag=='DC':
                if self.ITLA_type=='C80':
                    self.channel=[i for i in range(1,81)]
                    self.update_config()
                else:
                    self.channel=[i for i in range(9,73)]
                    self.update_config()
            elif self.test_flag=='TxBW' or self.test_flag=='ICR':
                self.channel=['13','39','65']
                self.update_config()
            elif self.test_flag=='RxRes':
                self.channel=['1','13','39','65','80']
                self.update_config()
                if not self.ITLA_type=='C80':
                    print('Not C80 C++ ITLA, please check config, Res test should use C++ ITLA!!!')
                    return
            # elif self.test_flag=='PeSkew':
            #     self.channel=[i for i in range(9,73)]
            #     self.update_config()
        elif self.test_type=='金样':
            self.channel=['13']
            self.update_config()
            print('Golden sample test starts...')
        elif self.test_type=='客户返回复测':
            self.channel = ['13']
            self.update_config()
            print('客户返回产品复测...')
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
        else:
            self.Stop_test()
            self.start_test_butt.setText('开始')

    # start the test
    @pyqtSlot()
    def Start_test_OLD(self):
        self.test_flag = self.Test_item.currentText()
        self.test_type = self.Test_type.currentText()
        self.finalResult == 'None'
        if self.start_test_butt.text() == '开始':
            self.start_test_butt.setText('停止')
            self.start = time.time()
            self.timer.start(1000)
            self.start_test_butt_3.setEnabled(False)
            self.start_test_butt_4.setEnabled(False)
            self.read_SN.setEnabled(False)
            self.Test_item.setEnabled(False)
            self.Test_type.setEnabled(False)
            self.workthread.start()
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
            elif self.test_flag == 'RxRes':
                if serial.Serial.isOpen(self.CtrlB):
                    cb.board_set(self,self.drv_down)
                    #self.print_out_status('Driver下电完成！')
                    cb.close_board(self)
                    self.print_out_status('测试板串口关闭完成！')
                if serial.Serial.isOpen(self.MPC):
                    mpc.close_MPC(self)
                    self.print_out_status('GBS偏振控制器串口关闭完成！')
            elif self.test_flag=='ICR':
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
            elif self.test_flag=='VOAcal':
                if serial.Serial.isOpen(self.CtrlB):
                    cb.close_board(self)
                    self.print_out_status('测试板串口关闭完成！')
            elif self.test_flag=='PeSkew':
                if serial.Serial.isOpen(self.CtrlB):
                    cb.close_board_PeSkew(self)
                    self.print_out_status('测试板串口关闭完成！')
                if serial.Serial.isOpen(self.JinLi):
                    jl.close_JINLI(self)
                    self.print_out_status('锦鲤串口关闭完成！')
        except Exception as e:
            print(e)
            self.print_out_status(str(e))

    def updata_status(self,f):
        mawin.test_status.setText(f)

    def updata_staColor(self,f):
        gf.status_color(self.test_status,f)

    def goca_timer(self):
        self.timer1.start(1000)# added

    def print_out_status(self,s=''):
        '''
        #print the status into the GUi and update the log file and the screen
        :param s:
        :return:
        '''
        status_show=''
        status_show=gf.get_timestamp(0)+': '+s+'\n'
        self.log+=status_show
        print(status_show)
        self.plainTextEdit.appendPlainText(status_show)

    def dc_scapOPdata(self, s=''):
        '''
        This is to deal with the scan op return data in the DC test
        :param:s -string with data inside
        :return:datafram of sweep result,columns=[Vcode,Delay1,Delay2....]
        '''
        lo = s.split('\r\n')
        lis = []
        dela = []
        for i in lo:
            if 'delay' in i:
                lis.append(lo.index(i))
                dela.append(i.split(':')[1])
        resu = []
        for t in range(len(lis)):
            if t < len(lis) - 1:
                tem = lo[lis[t] + 1:lis[t + 1]]
            else:
                tem = lo[lis[t] + 1:-1]
            inde = [j.split(' ')[0] for j in tem]
            dat = [j.split(' ')[1] for j in tem]
            if t == 0: resu.append(inde)
            resu.append(dat)
        df = pd.DataFrame(resu).transpose()
        df.columns = ['Vcode'] + dela
        return df

    def update_but(self,f):
        self.start_test_butt.setText(f)


    def update_pro(self,s):
        '''
        #update progressbar
        :param s:
        :return:
        '''
        self.progressBar.setValue(s)

    def text_clear(self):
        mawin.plainTextEdit.clear()

    def time_consuming(self):
        '''
        #Only for program test
        :return:
        '''
        self.a=0
        #self.gocflag=False
        while self.a<100000000:
            self.a+=1
            if (self.a%100)==0:
                print(self.a)
        self.gocflag=True

    ###******-------------RF test functions here-------------******
    def scan_RF(self, ch, ave=3):
        ''''
        function:scan RF from 1 to 10GHz to perform phase error test
        input:self,Current channel,average times
        :return:data array 30X4, every frequency test three times
        '''
        print('Phase Error calculating')
        # print('Phase Error calculating')
        data2 = float(str(self.ICRtf.query('SOUR1:CHAN1:POW? SET')).strip())
        time.sleep(0.1)
        self.OScope.write('tdiv 500e-12')
        self.OScope.write('VBS app.Acquisition.Horizontal.HorOffset=0')
        data = [[''] * 4] * ave * 10
        for i in range(10):
            # switch to test wavelength offset
            icr.set_wavelength(self, 1, ch, i + 1)
            ready = float(str(self.ICRtf.query('slot1:OPC?')).strip())
            while ready == 0.0:
                time.sleep(0.2)
                ready = float(str(self.ICRtf.query('slot1:OPC?')).strip())
            # control the scope
            self.OScope.write('TRMD SINGLE')
            time.sleep(1)
            # query the data of amtiplitude
            A1 = self.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')  # ;time.sleep(0.01)
            A2 = self.OScope.query('VBS? return=app.Measure.P2.last.Result.Value')  # ;time.sleep(0.01)
            A3 = self.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')  # ;time.sleep(0.01)
            A4 = self.OScope.query('VBS? return=app.Measure.P4.last.Result.Value')  # ;time.sleep(0.01)

            while not icr.isfloat(A1) or not icr.isfloat(A2) or not icr.isfloat(A3) or not icr.isfloat(A4):
                A1 = self.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')  # ;time.sleep(0.01)
                A2 = self.OScope.query('VBS? return=app.Measure.P2.last.Result.Value')  # ;time.sleep(0.01)
                A3 = self.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')  # ;time.sleep(0.01)
                A4 = self.OScope.query('VBS? return=app.Measure.P4.last.Result.Value')  # ;time.sleep(0.01)
            while float(A1) / float(A3) > 1.15 or float(A1) / float(A3) < 0.85 or float(A1) / float(A4) > 1.15 or float(
                    A1) / float(A3) < 0.87:
                time.sleep(0.1)
                icr.ICR_scamblePol(self)
                A1 = self.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')  # ;time.sleep(0.01)
                A2 = self.OScope.query('VBS? return=app.Measure.P2.last.Result.Value')  # ;time.sleep(0.01)
                A3 = self.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')  # ;time.sleep(0.01)
                A4 = self.OScope.query('VBS? return=app.Measure.P4.last.Result.Value')  # ;time.sleep(0.01)
                while not icr.isfloat(A1) or not icr.isfloat(A2) or not icr.isfloat(A3) or not icr.isfloat(A4):
                    A1 = self.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')  # ;time.sleep(0.01)
                    A2 = self.OScope.query('VBS? return=app.Measure.P2.last.Result.Value')  # ;time.sleep(0.01)
                    A3 = self.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')  # ;time.sleep(0.01)
                    A4 = self.OScope.query('VBS? return=app.Measure.P4.last.Result.Value')  # ;time.sleep(0.01)
            j = 0
            while j < ave:
                self.OScope.write('TRMD SINGLE');
                time.sleep(0.5)
                osc.switch_to_DSO(self, self.OScope_port);
                time.sleep(0.5)
                data[i * ave + j] = osc.get_data_DSO(self)  # data={'CH2':'','CH3':'','CH6':'','CH7':''}
                time.sleep(0.1)
                osc.switch_to_VISA(self, self.DSO_port);
                time.sleep(0.5)
                j += 1
            print('CE{} PE calculating...{}%'.format(str(ch), str(round((i + 1) * 10))))
        return data

    def scan_RF_ManualPol_PDbalance(self, ch, ave=3):
        ''''
        function:scan RF from 1 to 10GHz to perform phase error test
        input:self,Current channel,average times
        :return:data array 30X4, every frequency test three times
        '''
        print('Phase Error calculating')
        # print('Phase Error calculating')
        data2 = float(str(self.ICRtf.query('SOUR1:CHAN1:POW? SET')).strip())
        time.sleep(0.1)
        self.OScope.write('tdiv 500e-12')
        self.OScope.write('VBS app.Acquisition.Horizontal.HorOffset=0')
        data = [[''] * 4] * ave * 10
        self.OScope.write('TRMD SINGLE')
        icr.ICR_manualPol_PD(self)
        #icr.ICR_manualPol_PDbalance_ScopeJudge(self)
        for i in range(10):
            # switch to test wavelength offset
            icr.set_wavelength(self, 1, ch, i + 1)
            ready = float(str(self.ICRtf.query('slot1:OPC?')).strip())
            while ready == 0.0:
                time.sleep(0.2)
                ready = float(str(self.ICRtf.query('slot1:OPC?')).strip())
            current_x, current_y = icr.TIA_getPDcurrent(self)[0:2]
            #找PD电流的平衡点
            while current_x/current_y>1.09 or current_x/current_y<0.92:
                self.MPC_Find_Position_Fine('DIFF', False)
                current_x,current_y=icr.TIA_getPDcurrent(self)[0:2]
                print('find the balance point, X:{},Y:{}'.format(current_x,current_y))
            j = 0
            while j < ave:
                self.OScope.write('TRMD SINGLE');
                time.sleep(0.5)
                osc.switch_to_DSO(self, self.OScope_port);
                time.sleep(0.5)
                data[i * ave + j] = osc.get_data_DSO(self)  # data={'CH2':'','CH3':'','CH6':'','CH7':''}
                time.sleep(0.1)
                osc.switch_to_VISA(self, self.DSO_port);
                time.sleep(0.5)
                j += 1
            print('CE{} PE calculating...{}%'.format(str(ch), str(round((i + 1) * 10))))
        return data

    def scan_RF_ManualPol_SCOPEbalance(self, ch, ave=3):
        ''''
        function:scan RF from 1 to 10GHz to perform phase error test
        input:self,Current channel,average times
        :return:data array 30X4, every frequency test three times
        '''
        print('Phase Error calculating')
        # print('Phase Error calculating')
        data2 = float(str(self.ICRtf.query('SOUR1:CHAN1:POW? SET')).strip())
        time.sleep(0.1)
        self.OScope.write('tdiv 500e-12')
        self.OScope.write('VBS app.Acquisition.Horizontal.HorOffset=0')
        data = [[''] * 4] * ave * 10
        self.OScope.write('TRMD SINGLE')
        icr.ICR_manualPol_PD(self)
        #icr.ICR_manualPol_PDbalance_ScopeJudge(self)
        for i in range(10):
            # switch to test wavelength offset
            icr.set_wavelength(self, 1, ch, i + 1)
            ready = float(str(self.ICRtf.query('slot1:OPC?')).strip())
            while ready == 0.0:
                time.sleep(0.2)
                ready = float(str(self.ICRtf.query('slot1:OPC?')).strip())
            # current_x, current_y = icr.TIA_getPDcurrent(self)[0:2]
            # #找PD电流的平衡点
            # while current_x/current_y>1.09 or current_x/current_y<0.92:
            #     self.MPC_Find_Position_Fine('DIFF', False)
            #     current_x,current_y=icr.TIA_getPDcurrent(self)[0:2]
            #     print('find the balance point, X:{},Y:{}'.format(current_x,current_y))

            #control the scope
            self.OScope.write('TRMD SINGLE')
            time.sleep(1)
            icr.ICR_manualPol_ScopeBalance_Judge(self)
            # query the data of amtiplitude
            A1 = self.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')  # ;time.sleep(0.01)
            A2 = self.OScope.query('VBS? return=app.Measure.P2.last.Result.Value')  # ;time.sleep(0.01)
            A3 = self.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')  # ;time.sleep(0.01)
            A4 = self.OScope.query('VBS? return=app.Measure.P4.last.Result.Value')  # ;time.sleep(0.01)

            while not icr.isfloat(A1) or not icr.isfloat(A2) or not icr.isfloat(A3) or not icr.isfloat(A4):
                A1 = self.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')  # ;time.sleep(0.01)
                A2 = self.OScope.query('VBS? return=app.Measure.P2.last.Result.Value')  # ;time.sleep(0.01)
                A3 = self.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')  # ;time.sleep(0.01)
                A4 = self.OScope.query('VBS? return=app.Measure.P4.last.Result.Value')  # ;time.sleep(0.01)
            while float(A1) / float(A3) > 1.15 or float(A1) / float(A3) < 0.85 or float(A1) / float(A4) > 1.15 or float(
                    A1) / float(A3) < 0.87:
                time.sleep(0.1)
                #icr.ICR_scamblePol(self)
                #self.MPC_Find_Position_ScopeBalance_Fine()
                #icr.ICR_manualPol_ScopeBalance_Judge(self)
                self.MPC_Find_Position_ScopeBalance_Fine()
                A1 = self.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')  # ;time.sleep(0.01)
                A2 = self.OScope.query('VBS? return=app.Measure.P2.last.Result.Value')  # ;time.sleep(0.01)
                A3 = self.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')  # ;time.sleep(0.01)
                A4 = self.OScope.query('VBS? return=app.Measure.P4.last.Result.Value')  # ;time.sleep(0.01)
                while not icr.isfloat(A1) or not icr.isfloat(A2) or not icr.isfloat(A3) or not icr.isfloat(A4):
                    A1 = self.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')  # ;time.sleep(0.01)
                    A2 = self.OScope.query('VBS? return=app.Measure.P2.last.Result.Value')  # ;time.sleep(0.01)
                    A3 = self.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')  # ;time.sleep(0.01)
                    A4 = self.OScope.query('VBS? return=app.Measure.P4.last.Result.Value')  # ;time.sleep(0.01)
            j = 0
            while j < ave:
                self.OScope.write('TRMD SINGLE');
                time.sleep(0.5)
                osc.switch_to_DSO(self, self.OScope_port);
                time.sleep(0.5)
                data[i * ave + j] = osc.get_data_DSO(self)  # data={'CH2':'','CH3':'','CH6':'','CH7':''}
                time.sleep(0.1)
                osc.switch_to_VISA(self, self.DSO_port);
                time.sleep(0.5)
                j += 1
            print('CE{} PE calculating...{}%'.format(str(ch), str(round((i + 1) * 10))))
        return data

    def scan_S21(self, ch, bw_start, bw_stop, bw_step):
        '''
        function:scan RF S21 from 1 to 40GHz to perform Rx BW test
        input:self
        output:[float(A1),float(A2),float(A3),float(A4)]
        '''
        print('S21 testing')
        self.OScope.write('TRMD AUTO')
        m = 0
        num = round((bw_stop - bw_start) / bw_step) + 1
        data = [[0.0] * 4] * num
        for j in range(bw_start, bw_stop + 1, bw_step):
            icr.set_wavelength(self, 1, ch, j)
            ready = float(str(self.ICRtf.query('slot1:OPC?')).strip())
            while ready == 0.0:
                time.sleep(0.2)
                ready = float(str(self.ICRtf.query('slot1:OPC?')).strip())
            self.OScope.write('tdiv 1e-9')
            time.sleep(1)
            # query the data of amtiplitude and balance the XY light distribute
            A1 = self.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')  # ;time.sleep(0.01)
            A2 = self.OScope.query('VBS? return=app.Measure.P2.last.Result.Value')  # ;time.sleep(0.01)
            A3 = self.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')  # ;time.sleep(0.01)
            A4 = self.OScope.query('VBS? return=app.Measure.P4.last.Result.Value')  # ;time.sleep(0.01)

            while not icr.isfloat(A1) or not icr.isfloat(A2) or not icr.isfloat(A3) or not icr.isfloat(A4):
                A1 = self.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')  # ;time.sleep(0.01)
                A2 = self.OScope.query('VBS? return=app.Measure.P2.last.Result.Value')  # ;time.sleep(0.01)
                A3 = self.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')  # ;time.sleep(0.01)
                A4 = self.OScope.query('VBS? return=app.Measure.P4.last.Result.Value')  # ;time.sleep(0.01)
            while (float(A1) / float(A3) > 1.08 or float(A1) / float(A3) < 0.92) and (
                    float(A1) / float(A4) > 1.08 or float(A1) / float(A3) < 0.92):
                # while (float(A1)+float(A2))/(float(A3)+float(A4))>1.08 or (float(A1)+float(A2))/(float(A3)+float(A4))<0.92:
                time.sleep(0.1)
                icr.ICR_scamblePol(self)
                A1 = self.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')  # ;time.sleep(0.01)
                A2 = self.OScope.query('VBS? return=app.Measure.P2.last.Result.Value')  # ;time.sleep(0.01)
                A3 = self.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')  # ;time.sleep(0.01)
                A4 = self.OScope.query('VBS? return=app.Measure.P4.last.Result.Value')  # ;time.sleep(0.01)
                while not icr.isfloat(A1) or not icr.isfloat(A2) or not icr.isfloat(A3) or not icr.isfloat(A4):
                    A1 = self.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')  # ;time.sleep(0.01)
                    A2 = self.OScope.query('VBS? return=app.Measure.P2.last.Result.Value')  # ;time.sleep(0.01)
                    A3 = self.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')  # ;time.sleep(0.01)
                    A4 = self.OScope.query('VBS? return=app.Measure.P4.last.Result.Value')  # ;time.sleep(0.01)
            data[m] = [float(A1), float(A2), float(A3), float(A4)]
            m += 1
            print('CE{} Fre {}GHz S21 testing...{}%'.format(str(ch), str(j),
                                                            str(round(((j - bw_start) / bw_step + 1) / num * 100, 1))))
        return data

    def scan_S21_ManualPol_PDbalance(self, ch, bw_start, bw_stop, bw_step):
        '''
        function:scan RF S21 from 1 to 40GHz to perform Rx BW test
        input:self
        output:[float(A1),float(A2),float(A3),float(A4)]
        '''
        print('S21 testing')
        self.OScope.write('TRMD AUTO')
        m = 0
        num = round((bw_stop - bw_start) / bw_step) + 1
        data = [[0.0] * 4] * num
        icr.ICR_manualPol_PD(self)
        for j in range(bw_start, bw_stop + 1, bw_step):
            icr.set_wavelength(self, 1, ch, j)
            ready = float(str(self.ICRtf.query('slot1:OPC?')).strip())
            while ready == 0.0:
                time.sleep(0.2)
                ready = float(str(self.ICRtf.query('slot1:OPC?')).strip())
            self.OScope.write('tdiv 1e-9')
            time.sleep(1)
            current_x, current_y = icr.TIA_getPDcurrent(self)[0:2]
            #找PD电流的平衡点
            while current_x/current_y>1.09 or current_x/current_y<0.92:
                self.MPC_Find_Position_Fine('DIFF', False)
                current_x,current_y=icr.TIA_getPDcurrent(self)[0:2]
                print('find the balance point, X:{},Y:{}'.format(current_x,current_y))
            # query the data of amtiplitude and balance the XY light distribute
            A1 = self.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')  # ;time.sleep(0.01)
            A2 = self.OScope.query('VBS? return=app.Measure.P2.last.Result.Value')  # ;time.sleep(0.01)
            A3 = self.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')  # ;time.sleep(0.01)
            A4 = self.OScope.query('VBS? return=app.Measure.P4.last.Result.Value')  # ;time.sleep(0.01)

            while not icr.isfloat(A1) or not icr.isfloat(A2) or not icr.isfloat(A3) or not icr.isfloat(A4):
                A1 = self.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')  # ;time.sleep(0.01)
                A2 = self.OScope.query('VBS? return=app.Measure.P2.last.Result.Value')  # ;time.sleep(0.01)
                A3 = self.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')  # ;time.sleep(0.01)
                A4 = self.OScope.query('VBS? return=app.Measure.P4.last.Result.Value')  # ;time.sleep(0.01)
            data[m] = [float(A1), float(A2), float(A3), float(A4)]
            m += 1
            print('CE{} Fre {}GHz S21 testing...{}%'.format(str(ch), str(j),
                                                            str(round(((j - bw_start) / bw_step + 1) / num * 100, 1))))
        return data

    def scan_S21_ManualPol_SCOPEbalance(self, ch, bw_start, bw_stop, bw_step):
        '''
        function:scan RF S21 from 1 to 40GHz to perform Rx BW test
        input:self
        output:[float(A1),float(A2),float(A3),float(A4)]
        '''
        print('S21 testing')
        self.OScope.write('TRMD AUTO')
        m = 0
        num = round((bw_stop - bw_start) / bw_step) + 1
        data = [[0.0] * 4] * num
        icr.ICR_manualPol_ScopeBalance_Judge(self)
        # icr.ICR_manualPol_PD(self)
        for j in range(bw_start, bw_stop + 1, bw_step):
            icr.set_wavelength(self, 1, ch, j)
            ready = float(str(self.ICRtf.query('slot1:OPC?')).strip())
            while ready == 0.0:
                time.sleep(0.2)
                ready = float(str(self.ICRtf.query('slot1:OPC?')).strip())
            self.OScope.write('tdiv 1e-9')
            time.sleep(1)
            # current_x, current_y = icr.TIA_getPDcurrent(self)[0:2]
            # #找PD电流的平衡点
            # while current_x/current_y>1.09 or current_x/current_y<0.92:
            #     self.MPC_Find_Position_Fine('DIFF', False)
            #     current_x,current_y=icr.TIA_getPDcurrent(self)[0:2]
            #     print('find the balance point, X:{},Y:{}'.format(current_x,current_y))
            # query the data of amtiplitude and balance the XY light distribute
            A1 = self.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')  # ;time.sleep(0.01)
            A2 = self.OScope.query('VBS? return=app.Measure.P2.last.Result.Value')  # ;time.sleep(0.01)
            A3 = self.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')  # ;time.sleep(0.01)
            A4 = self.OScope.query('VBS? return=app.Measure.P4.last.Result.Value')  # ;time.sleep(0.01)

            while not icr.isfloat(A1) or not icr.isfloat(A2) or not icr.isfloat(A3) or not icr.isfloat(A4):
                A1 = self.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')  # ;time.sleep(0.01)
                A2 = self.OScope.query('VBS? return=app.Measure.P2.last.Result.Value')  # ;time.sleep(0.01)
                A3 = self.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')  # ;time.sleep(0.01)
                A4 = self.OScope.query('VBS? return=app.Measure.P4.last.Result.Value')  # ;time.sleep(0.01)
            while (float(A1) / float(A3) > 1.08 or float(A1) / float(A3) < 0.92) and (
                    float(A1) / float(A4) > 1.08 or float(A1) / float(A3) < 0.92):
                # while (float(A1)+float(A2))/(float(A3)+float(A4))>1.08 or (float(A1)+float(A2))/(float(A3)+float(A4))<0.92:
                time.sleep(0.1)
                self.MPC_Find_Position_ScopeBalance_Fine()
                #icr.ICR_manualPol_ScopeBalance_Judge(self)
                A1 = self.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')  # ;time.sleep(0.01)
                A2 = self.OScope.query('VBS? return=app.Measure.P2.last.Result.Value')  # ;time.sleep(0.01)
                A3 = self.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')  # ;time.sleep(0.01)
                A4 = self.OScope.query('VBS? return=app.Measure.P4.last.Result.Value')  # ;time.sleep(0.01)
                while not icr.isfloat(A1) or not icr.isfloat(A2) or not icr.isfloat(A3) or not icr.isfloat(A4):
                    A1 = self.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')  # ;time.sleep(0.01)
                    A2 = self.OScope.query('VBS? return=app.Measure.P2.last.Result.Value')  # ;time.sleep(0.01)
                    A3 = self.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')  # ;time.sleep(0.01)
                    A4 = self.OScope.query('VBS? return=app.Measure.P4.last.Result.Value')  # ;time.sleep(0.01)
            data[m] = [float(A1), float(A2), float(A3), float(A4)]
            m += 1
            print('CE{} Fre {}GHz S21 testing...{}%'.format(str(ch), str(j),
                                                            str(round(((j - bw_start) / bw_step + 1) / num * 100, 1))))
        return data

    def get_pd_resp_sig(self, ch):
        '''
        Get signal(Rx) port PD responsivity
        :return:[res_sig_x,res_sig_y,per_x,per_y,dark_x,dark_y]
        '''
        # load sig calibration data here:'D:\Fisilink\RX_Test\power_cal\power_sig.txt'
        if ch == '13':
            pwr = float(self.ICR_pwr.iloc[0, 1])
        elif ch == '39':
            pwr = float(self.ICR_pwr.iloc[1, 1])
        elif ch == '65':
            pwr = float(self.ICR_pwr.iloc[2, 1])
        else:
            pwr = 0
            print('Selected channel not in the calibration list, set power as 0dBm, please check')
        # Firstly get dark current and set the flags to determine the test point and time
        self.ICRtf.write('OUTP1:CHAN1:STATE OFF')
        self.ICRtf.write('OUTP1:CHAN2:STATE OFF')
        dark_x, dark_y = icr.TIA_getPDcurrent(self)[0:2]  # [current_x,current_y,VPDX,VPDY]
        if dark_x > 0.004 or dark_y > 0.004:
            dark_x, dark_y = icr.TIA_getPDcurrent(self)[0:2]  # [current_x,current_y,VPDX,VPDY]
            if dark_x > 0.004 or dark_y > 0.004:
                self.limit = 1  # 暗电流较大
                self.limit_samp = 1000  # 如果暗电流比较大，最多测500个点
            else:
                self.limit = 0  # 暗电流较小
                self.limit_samp = 2000
        else:
            self.limit = 0  # 暗电流较小
            self.limit_samp = 2000

        self.ICRtf.write('OUTP1:CHAN1:STATE OFF');
        time.sleep(0.1)
        self.ICRtf.write('OUTP3:CHAN2:POW 5');
        time.sleep(1)  # laser2切换为6dBm 20220704改成5dBm
        self.ICRtf.write('OUTP1:CHAN2:STATE ON')
        ready = float(str(self.ICRtf.query('slot1:OPC?')).strip())
        while ready == 0.0:
            time.sleep(0.2)
            ready = float(str(self.ICRtf.query('slot1:OPC?')).strip())
        time.sleep(3)
        print('Laser set ready!')
        # 自动旋转偏振态找最大最小电流
        cur_x, cur_y, per_x, per_y, PD_currentX, PD_currentY = self.ICR_scamblePol_XYmax_Auto()
        # calculate the responsivity
        # pwr=1,Responsivity unit:mA/mW
        res_sig_x = cur_x / (10 ** (pwr / 10)) / 8
        res_sig_y = cur_y / (10 ** (pwr / 10)) / 8
        print('Res_Sig_X,Res_Sig_Y,PER_X,PER_Y:\n', [res_sig_x, res_sig_y, per_x, per_y, dark_x, dark_y])
        return [res_sig_x, res_sig_y, per_x, per_y, dark_x, dark_y, PD_currentX, PD_currentY]

    def get_pd_resp_sig_New_GBSmpc_old(self, ch):
        '''
        Using the GBS MPC
        The new function to manual control MPC to...
        Get signal(Rx) port PD responsivity
        :return:[res_sig_x,res_sig_y,per_x,per_y,dark_x,dark_y,PD_currentX,PD_currentY,ind_markYmin,ind_markXmin]
        '''
        # load sig calibration data here:'D:\Fisilink\RX_Test\power_cal\power_sig.txt'
        if ch == '13':
            pwr = float(self.ICR_pwr.iloc[0, 1])
        elif ch == '39':
            pwr = float(self.ICR_pwr.iloc[1, 1])
        elif ch == '65':
            pwr = float(self.ICR_pwr.iloc[2, 1])
        else:
            pwr = 0
            return
            print('Selected channel not in the calibration list, set power as 0dBm, please check')
        # Firstly get dark current and set the flags to determine the test point and time
        self.ICRtf.write('OUTP1:CHAN1:STATE OFF')
        self.ICRtf.write('OUTP1:CHAN2:STATE OFF')
        dark_x, dark_y = icr.TIA_getPDcurrent(self)[0:2]  # [current_x,current_y,VPDX,VPDY]
        if dark_x > 0.004 or dark_y > 0.004:
            dark_x, dark_y = icr.TIA_getPDcurrent(self)[0:2]  # [current_x,current_y,VPDX,VPDY]
            if dark_x > 0.004 or dark_y > 0.004:
                self.limit = 1  # 暗电流较大
                self.limit_samp = 1000  # 如果暗电流比较大，最多测500个点
            else:
                self.limit = 0  # 暗电流较小
                self.limit_samp = 2000
        else:
            self.limit = 0  # 暗电流较小
            self.limit_samp = 2000

        self.ICRtf.write('OUTP1:CHAN1:STATE OFF');
        time.sleep(0.1)
        self.ICRtf.write('OUTP3:CHAN2:POW 5');
        time.sleep(1)  # laser2切换为6dBm 20220704改成5dBm
        self.ICRtf.write('OUTP1:CHAN2:STATE ON')
        ready = float(str(self.ICRtf.query('slot1:OPC?')).strip())
        while ready == 0.0:
            time.sleep(0.2)
            ready = float(str(self.ICRtf.query('slot1:OPC?')).strip())
        time.sleep(3)
        print('Laser set ready!')
        Xmax_pd_currentX=[]
        Xmax_pd_currentY=[]
        Ymax_pd_currentX = []
        Ymax_pd_currentY = []
        pd_currentX = []
        pd_currentY = []
        ind_markXmin = []
        ind_markYmin = []
        if self.limit==0:
            print('暗电流较小，可以找到4uA以下，否则重新找一遍')
            #Now start manual find the status:X max Y min
            Xmax_pd_currentX,Xmax_pd_currentY,ind_markYmin= self.MPC_GBS_Find_Position('PDY', False)
            count=0
            #进行两次X微调
            while count<3:
                if not (min(Xmax_pd_currentY)) <0.004:
                    count+=1
                    print('CurYmin 大于4uA，进行第{}次重调'.format(count))
                    Xmax_pd_currentX,Xmax_pd_currentY,ind_markYmin= self.MPC_GBS_Find_Position('PDY', False)
                    #curXmax, curYmin = icr.TIA_getPDcurrent(mawin)[0:2]
                else:
                    break
            Ymax_pd_currentX,Ymax_pd_currentY,ind_markXmin= self.MPC_GBS_Find_Position('PDX', False)
            #curXmin, curYmax = icr.TIA_getPDcurrent(mawin)[0:2]
            count = 0
            # 进行两次Y微调
            while count < 3:
                if not (min(Ymax_pd_currentX)) < 0.004:
                    count += 1
                    print('CurXmin 大于4uA，进行第{}次重调'.format(count))
                    Ymax_pd_currentX,Ymax_pd_currentY,ind_markXmin= self.MPC_GBS_Find_Position('PDX', False)
                    #curXmin, curYmax = icr.TIA_getPDcurrent(mawin)[0:2]
                else:
                    break
        else:
            print('暗电流较大，不进行重调')
            #this is to compare result of find maximun and minimum direction
            Xmax_pd_currentX,Xmax_pd_currentY,ind_markYmin= self.MPC_GBS_Find_Position('PDY', False)
            c_x1, c_y1 = icr.TIA_getPDcurrent(self)[0:2]
            Ymax_pd_currentX,Ymax_pd_currentY,ind_markXmin= self.MPC_GBS_Find_Position('PDX', False)
            c_x2, c_y2 = icr.TIA_getPDcurrent(self)[0:2]
        curXmax=max(Xmax_pd_currentX)
        curYmin=min(Xmax_pd_currentY)
        curXmin=min(Ymax_pd_currentX)
        curYmax=max(Ymax_pd_currentY)
        pd_currentX=Xmax_pd_currentX+Ymax_pd_currentX
        pd_currentY=Xmax_pd_currentY+Ymax_pd_currentY
        print('curXmax, curYmin:{}\ncurXmin, curYmax:{}'.format([curXmax, curYmin],[curXmin, curYmax]))

        res_sig_x = curXmax / (10 ** (pwr / 10)) / 8
        res_sig_y = curYmax / (10 ** (pwr / 10)) / 8
        per_x = 10 * np.log10(curXmax / curYmin)
        per_y = 10 * np.log10(curYmax / curXmin)
        print('Res_Sig_X,Res_Sig_Y,PER_X,PER_Y,Dark_X,Dark_Y:\n', [res_sig_x, res_sig_y, per_x, per_y, dark_x, dark_y])
        return [res_sig_x, res_sig_y, per_x, per_y, dark_x, dark_y,pd_currentX,pd_currentY,ind_markYmin,ind_markXmin]

    def get_pd_resp_sig_New_GBSmpc(self, ch):
        '''
        Using the GBS MPC
        The new function to manual control MPC to...
        Get signal(Rx) port PD responsivity
        :return:[res_sig_x,res_sig_y,per_x,per_y,dark_x,dark_y,PD_currentX,PD_currentY,ind_markYmin,ind_markXmin]
        '''
        # load sig calibration data here:'D:\Fisilink\RX_Test\power_cal\power_sig.txt'
        # if ch == '1':
        #     pwr = float(self.ICR_pwr.iloc[0, 0])
        # elif ch == '39':
        #     pwr = float(self.ICR_pwr.iloc[1, 1])
        # elif ch == '65':
        #     pwr = float(self.ICR_pwr.iloc[2, 1])
        # else:
        #     pwr = 0
        #     return
        #     print('Selected channel not in the calibration list, set power as 0dBm, please check')
        pwr = float(self.Res_ITLA_pwr_0.iloc[int(ch) - 1, 1])
        # Firstly get dark current and set the flags to determine the test point and time

        self.CtrlB.write(b'itla_wr 0 0x32 0x00\n');time.sleep(0.1)
        print(self.CtrlB.read_until(b'Write itla'))
        print('Rx laser closed...')
        self.CtrlB.write(b'itla_wr 1 0x32 0x00\n');time.sleep(0.1) #shut down LO port
        print(self.CtrlB.read_until(b'Write itla'))
        print('Lo laser closed...')
        dark_x, dark_y = icr.TIA_getPDcurrent(self)[0:2]  # [current_x,current_y,VPDX,VPDY]
        if dark_x > 0.004 or dark_y > 0.004:
            dark_x, dark_y = icr.TIA_getPDcurrent(self)[0:2]  # [current_x,current_y,VPDX,VPDY]
            if dark_x > 0.004 or dark_y > 0.004:
                self.limit = 1  # 暗电流较大
                self.limit_samp = 1000  # 如果暗电流比较大，最多测500个点
            else:
                self.limit = 0  # 暗电流较小
                self.limit_samp = 2000
        else:
            self.limit = 0  # 暗电流较小
            self.limit_samp = 2000

        self.CtrlB.write(b'itla_wr 0 0x32 0x08\n');time.sleep(0.1)
        print(self.CtrlB.read_until(b'Write itla'))
        cu_x, cu_y = icr.TIA_getPDcurrent(self)[0:2]
        # while cu_x<0.01 and cu_y<0.01:
        #     self.CtrlB.write(b'itla_wr 0 0x32 0x08\n');time.sleep(0.1)
        #     print(self.CtrlB.read_until(b'Write itla'))
        count = 0
        while cu_x < 0.01 and cu_y < 0.01:
            self.CtrlB.write(b'itla_wr 0 0x32 0x08\n');
            time.sleep(0.1)
            print('The {} time Rx port light open retry...'.format(count))
            print(self.CtrlB.read_until(b'Write itla'))
            cu_x, cu_y = icr.TIA_getPDcurrent(self)[0:2]
            count += 1
            if count == 5:
                print('Error open Rx ITLA or Rx fiber not connectted, please check!')
                return
        #print('Lo laser open!')
        print('Rx laser open...')
        #print('Laser set ready!')
        Xmax_pd_currentX=[]
        Xmax_pd_currentY=[]
        Ymax_pd_currentX = []
        Ymax_pd_currentY = []
        pd_currentX = []
        pd_currentY = []
        ind_markXmin = []
        ind_markYmin = []
        if self.limit==0:
            print('暗电流较小，可以找到4uA以下，否则重新找一遍')
            #Now start manual find the status:X max Y min
            Xmax_pd_currentX,Xmax_pd_currentY,ind_markYmin= self.MPC_GBS_Find_Position('PDY', False)
            count=0
            #进行两次X微调
            while count<3:
                if not (min(Xmax_pd_currentY)) <0.004:
                    count+=1
                    print('CurYmin 大于4uA，进行第{}次重调'.format(count))
                    Xmax_pd_currentX,Xmax_pd_currentY,ind_markYmin= self.MPC_GBS_Find_Position('PDY', False)
                    #curXmax, curYmin = icr.TIA_getPDcurrent(mawin)[0:2]
                else:
                    break
            Ymax_pd_currentX,Ymax_pd_currentY,ind_markXmin= self.MPC_GBS_Find_Position('PDX', False)
            #curXmin, curYmax = icr.TIA_getPDcurrent(mawin)[0:2]
            count = 0
            # 进行两次Y微调
            while count < 3:
                if not (min(Ymax_pd_currentX)) < 0.004:
                    count += 1
                    print('CurXmin 大于4uA，进行第{}次重调'.format(count))
                    Ymax_pd_currentX,Ymax_pd_currentY,ind_markXmin= self.MPC_GBS_Find_Position('PDX', False)
                    #curXmin, curYmax = icr.TIA_getPDcurrent(mawin)[0:2]
                else:
                    break
        else:
            print('暗电流较大，不进行重调')
            #this is to compare result of find maximun and minimum direction
            Xmax_pd_currentX,Xmax_pd_currentY,ind_markYmin= self.MPC_GBS_Find_Position('PDY', False)
            c_x1, c_y1 = icr.TIA_getPDcurrent(self)[0:2]
            Ymax_pd_currentX,Ymax_pd_currentY,ind_markXmin= self.MPC_GBS_Find_Position('PDX', False)
            c_x2, c_y2 = icr.TIA_getPDcurrent(self)[0:2]
        curXmax=max(Xmax_pd_currentX)
        curYmin=min(Xmax_pd_currentY)
        curXmin=min(Ymax_pd_currentX)
        curYmax=max(Ymax_pd_currentY)
        pd_currentX=Xmax_pd_currentX+Ymax_pd_currentX
        pd_currentY=Xmax_pd_currentY+Ymax_pd_currentY
        print('curXmax, curYmin:{}\ncurXmin, curYmax:{}'.format([curXmax, curYmin],[curXmin, curYmax]))

        res_sig_x = curXmax / (10 ** (pwr / 10)) / 8
        res_sig_y = curYmax / (10 ** (pwr / 10)) / 8
        per_x = 10 * np.log10(curXmax / curYmin)
        per_y = 10 * np.log10(curYmax / curXmin)
        print('Res_Sig_X,Res_Sig_Y,PER_X,PER_Y,Dark_X,Dark_Y:\n', [res_sig_x, res_sig_y, per_x, per_y, dark_x, dark_y])
        return [res_sig_x, res_sig_y, per_x, per_y, dark_x, dark_y,pd_currentX,pd_currentY,ind_markYmin,ind_markXmin]


    def get_pd_resp_sig_New(self, ch):
        '''
        The new function to manual control MPC to...
        Get signal(Rx) port PD responsivity
        :return:[res_sig_x,res_sig_y,per_x,per_y,dark_x,dark_y,PD_currentX,PD_currentY,ind_markYmin,ind_markXmin]
        '''
        # load sig calibration data here:'D:\Fisilink\RX_Test\power_cal\power_sig.txt'
        if ch == '13':
            pwr = float(self.ICR_pwr.iloc[0, 1])
        elif ch == '39':
            pwr = float(self.ICR_pwr.iloc[1, 1])
        elif ch == '65':
            pwr = float(self.ICR_pwr.iloc[2, 1])
        else:
            pwr = 0
            return
            print('Selected channel not in the calibration list, set power as 0dBm, please check')
        # Firstly get dark current and set the flags to determine the test point and time
        self.ICRtf.write('OUTP1:CHAN1:STATE OFF')
        self.ICRtf.write('OUTP1:CHAN2:STATE OFF')
        dark_x, dark_y = icr.TIA_getPDcurrent(self)[0:2]  # [current_x,current_y,VPDX,VPDY]
        if dark_x > 0.004 or dark_y > 0.004:
            dark_x, dark_y = icr.TIA_getPDcurrent(self)[0:2]  # [current_x,current_y,VPDX,VPDY]
            if dark_x > 0.004 or dark_y > 0.004:
                self.limit = 1  # 暗电流较大
                self.limit_samp = 1000  # 如果暗电流比较大，最多测500个点
            else:
                self.limit = 0  # 暗电流较小
                self.limit_samp = 2000
        else:
            self.limit = 0  # 暗电流较小
            self.limit_samp = 2000

        self.ICRtf.write('OUTP1:CHAN1:STATE OFF');
        time.sleep(0.1)
        self.ICRtf.write('OUTP3:CHAN2:POW 5');
        time.sleep(1)  # laser2切换为6dBm 20220704改成5dBm
        self.ICRtf.write('OUTP1:CHAN2:STATE ON')
        ready = float(str(self.ICRtf.query('slot1:OPC?')).strip())
        while ready == 0.0:
            time.sleep(0.2)
            ready = float(str(self.ICRtf.query('slot1:OPC?')).strip())
        time.sleep(3)
        print('Laser set ready!')
        Xmax_pd_currentX=[]
        Xmax_pd_currentY=[]
        Ymax_pd_currentX = []
        Ymax_pd_currentY = []
        pd_currentX = []
        pd_currentY = []
        ind_markXmin = []
        ind_markYmin = []
        if self.limit==0:
            print('暗电流较小，可以找到4uA以下，否则重新找一遍')
            #Now start manual find the status:X max Y min
            Xmax_pd_currentX,Xmax_pd_currentY,ind_markYmin= self.MPC_Find_Position('PDY', False)
            count=0
            #进行两次X微调
            while count<3:
                if not (min(Xmax_pd_currentY)) <0.004:
                    count+=1
                    print('CurYmin 大于4uA，进行第{}次重调'.format(count))
                    Xmax_pd_currentX,Xmax_pd_currentY,ind_markYmin= self.MPC_Find_Position('PDY', False)
                    #curXmax, curYmin = icr.TIA_getPDcurrent(mawin)[0:2]
                else:
                    break
            Ymax_pd_currentX,Ymax_pd_currentY,ind_markXmin= self.MPC_Find_Position('PDX', False)
            #curXmin, curYmax = icr.TIA_getPDcurrent(mawin)[0:2]
            count = 0
            # 进行两次Y微调
            while count < 3:
                if not (min(Ymax_pd_currentX)) < 0.004:
                    count += 1
                    print('CurXmin 大于4uA，进行第{}次重调'.format(count))
                    Ymax_pd_currentX,Ymax_pd_currentY,ind_markXmin= self.MPC_Find_Position('PDX', False)
                    #curXmin, curYmax = icr.TIA_getPDcurrent(mawin)[0:2]
                else:
                    break
        else:
            print('暗电流较大，不进行重调')
            #this is to compare result of find maximun and minimum direction
            Xmax_pd_currentX,Xmax_pd_currentY,ind_markYmin= self.MPC_Find_Position('PDY', False)
            c_x1, c_y1 = icr.TIA_getPDcurrent(self)[0:2]
            Ymax_pd_currentX,Ymax_pd_currentY,ind_markXmin= self.MPC_Find_Position('PDX', False)
            c_x2, c_y2 = icr.TIA_getPDcurrent(self)[0:2]
        curXmax=max(Xmax_pd_currentX)
        curYmin=min(Xmax_pd_currentY)
        curXmin=min(Ymax_pd_currentX)
        curYmax=max(Ymax_pd_currentY)
        pd_currentX=Xmax_pd_currentX+Ymax_pd_currentX
        pd_currentY=Xmax_pd_currentY+Ymax_pd_currentY
        print('curXmax, curYmin:{}\ncurXmin, curYmax:{}'.format([curXmax, curYmin],[curXmin, curYmax]))

        res_sig_x = curXmax / (10 ** (pwr / 10)) / 8
        res_sig_y = curYmax / (10 ** (pwr / 10)) / 8
        per_x = 10 * np.log10(curXmax / curYmin)
        per_y = 10 * np.log10(curYmax / curXmin)
        print('Res_Sig_X,Res_Sig_Y,PER_X,PER_Y,Dark_X,Dark_Y:\n', [res_sig_x, res_sig_y, per_x, per_y, dark_x, dark_y])
        return [res_sig_x, res_sig_y, per_x, per_y, dark_x, dark_y,pd_currentX,pd_currentY,ind_markYmin,ind_markXmin]



    def get_pd_resp_lo(self, ch):
        '''
        Get local port PD responsivity
        :parameter:Channel currently tested
        :return:[res_lo_x,res_lo_y]
        '''
        # load sig calibration data here:'D:\Fisilink\RX_Test\power_cal\power_lo.txt'
        if ch == '13':
            pwr = float(self.ICR_pwr.iloc[0, 2])
        elif ch == '39':
            pwr = float(self.ICR_pwr.iloc[1, 2])
        elif ch == '65':
            pwr = float(self.ICR_pwr.iloc[2, 2])
        else:
            pwr = 0
            print('Selected channel not in the calibration list, set power as 0dBm, please check')

        # Firstly get dark current and set the flags to determine the test point and time
        self.ICRtf.write('OUTP1:CHAN2:STATE OFF');
        time.sleep(0.1)
        self.ICRtf.write('OUTP3:CHAN1:POW 6');
        time.sleep(1)  # laser1切换为9dBm 20220704改成6dBm
        self.ICRtf.write('OUTP1:CHAN1:STATE ON');
        time.sleep(0.1)
        ready = float(str(self.ICRtf.query('slot1:OPC?')).strip())
        while ready == 0.0:
            time.sleep(0.2)
            ready = float(str(self.ICRtf.query('slot1:OPC?')).strip())
        print('Laser set ready!')
        cur_x, cur_y = icr.TIA_getPDcurrent_LO(self)[0:2]
        '''
        resp_x=current_X/(10^(power(find(ch400==i))/10))/4 %find(ch400==i)返回值为ch400中数值i对应的行/列，如果是二位矩阵，[m,n]=find(ch400==i)
        resp_y=current_Y/(10^(power(find(ch400==i))/10))/4
        '''
        # calculate the responsivity
        res_lo_x = cur_x / (10 ** (pwr / 10)) / 4
        res_lo_y = cur_y / (10 ** (pwr / 10)) / 4
        print('Res_LO_X,Res_LO_Y,Cur_X,Cur_Y:\n', [res_lo_x, res_lo_y, cur_x, cur_y])
        return [res_lo_x, res_lo_y, cur_x, cur_y]

    def get_pd_resp_lo_ITLA80C(self, ch):
        '''
        Get local port PD responsivity using 80C ITLA on the control board
        :parameter:Channel currently tested
        :return:[res_lo_x, res_lo_y, cur_x, cur_y]
        '''
        # load sig calibration data here:'D:\Fisilink\RX_Test\power_cal\power_lo.txt'
        # if ch == '13':
        #     pwr = float(self.ICR_pwr.iloc[0, 2])
        # elif ch == '39':
        #     pwr = float(self.ICR_pwr.iloc[1, 2])
        # elif ch == '65':
        #     pwr = float(self.ICR_pwr.iloc[2, 2])
        # else:
        #     pwr = 0
        #     print('Selected channel not in the calibration list, set power as 0dBm, please check')
        pwr = float(self.Res_ITLA_pwr_1.iloc[int(ch) - 1, 1])

        # Firstly get dark current and set the flags to determine the test point and time
        # print('Laser set ready!')
        ##Work left here to close the Rx ITLA and open LO ITLA
        self.CtrlB.write(b'itla_wr 0 0x32 0x00\n');
        time.sleep(0.1)
        print(self.CtrlB.read_until(b'Write itla'))
        print('Rx laser closed...')
        self.CtrlB.write(b'itla_wr 1 0x32 0x08\n');time.sleep(0.1) #Open LO port
        print(self.CtrlB.read_until(b'Write itla'))
        cu_x, cu_y = icr.TIA_getPDcurrent(self)[0:2]
        count=0
        while cu_x < 0.01 and cu_y < 0.01:
            self.CtrlB.write(b'itla_wr 1 0x32 0x08\n');
            time.sleep(0.1)
            print('The {} time LO port light open retry...'.format(count))
            print(self.CtrlB.read_until(b'Write itla'))
            cu_x, cu_y = icr.TIA_getPDcurrent(self)[0:2]
            count+=1
            if count ==5:
                print('Error open LO ITLA, please check!')
                return
        print('Lo laser open!')
        cur_x, cur_y = icr.TIA_getPDcurrent_LO(self)[0:2]
        '''
        resp_x=current_X/(10^(power(find(ch400==i))/10))/4 %find(ch400==i)返回值为ch400中数值i对应的行/列，如果是二位矩阵，[m,n]=find(ch400==i)
        resp_y=current_Y/(10^(power(find(ch400==i))/10))/4
        '''
        # calculate the responsivity
        # pwr=1
        res_lo_x = cur_x / (10 ** (pwr / 10)) / 4
        res_lo_y = cur_y / (10 ** (pwr / 10)) / 4
        print('Res_LO_X,Res_LO_Y,Cur_X,Cur_Y:\n', [res_lo_x, res_lo_y, cur_x, cur_y])
        return [res_lo_x, res_lo_y, cur_x, cur_y]

    def ICR_scamblePol_XYmax_Auto(self):
        '''
        Auto scamble mode to find XY max status
        :param:GUI Class obj
        :return:[current_X,current_Y,PER_X,PER_Y]
        '''
        try:
            # 旋转偏振控制器，使得X/Y两路输出基本一致（相差10%）
            self.ICRtf.write(':POL2:MODE SCRAMBLE');
            time.sleep(0.1)
            self.ICRtf.write(':SCRAMBLE:FUNCTION SIN');
            time.sleep(0.1)
            self.ICRtf.write(':POL2:SCRAMBLE:FREQUENCY1 0.010');
            time.sleep(0.1)
            self.ICRtf.write(':POL2:SCRAMBLE:FREQUENCY2 0.018');
            time.sleep(0.1)
            self.ICRtf.write(':POL2:SCRAMBLE:FREQUENCY3 0.02');
            time.sleep(0.1)

            num = 1
            current_X = 0
            current_Y = 0
            PD_current_X = []
            PD_current_Y = []
            cur_x, cur_y = [0, 0]
            while current_X * current_Y == 0:
                if self.limit == 0:  # 暗电流正常（暗电流过大，limit_samp有值；暗电流很小，limit_samp为空）
                    cur_x, cur_y = icr.TIA_getPDcurrent(self)[0:2]
                    if cur_x <= 0.004:
                        if current_Y == 0:
                            current_Y = cur_y
                        else:
                            current_Y = max(current_Y, cur_y)
                        current_Xmin = cur_x
                        PER_X = 10 * np.log10(current_Y / current_Xmin)
                    if cur_y <= 0.004:
                        if current_X == 0:
                            current_X = cur_x
                        else:
                            current_X = max(current_X, cur_x)
                        current_Ymin = cur_y
                        PER_Y = 10 * np.log10(current_X / current_Ymin)
                    PD_current_X.append(cur_x)
                    PD_current_Y.append(cur_y)
                    num += 1
                    print('---Rx signal port test---\nTotal {} points, now at NO.{}\nCurrent_X:{}\nCurrent_Y:{}'.format(
                        str(self.limit_samp),
                        str(num), str(cur_x), str(cur_y)))
                    if num >= self.limit_samp:
                        # update 0617/2022 improve judgement condition
                        if current_Y == 0:
                            current_Xmin = min(PD_current_X)
                            num_minX = PD_current_X.index(current_Xmin)
                            current_Y = PD_current_Y[num_minX]

                        if current_X == 0:
                            current_Ymin = min(PD_current_Y)
                            num_minY = PD_current_Y.index(current_Ymin)
                            current_X = PD_current_X[num_minY]

                        PER_X = 10 * np.log10(current_Y / current_Xmin)
                        PER_Y = 10 * np.log10(current_X / current_Ymin)
                        break
                else:  # 暗电流很大的情况
                    cur_x, cur_y = icr.TIA_getPDcurrent(self)[0:2]
                    PD_current_X.append(cur_x)
                    PD_current_Y.append(cur_y)
                    num += 1
                    print('---Rx signal port test---\nTotal {} points, now at NO.{}\nCurrent_X:{}\nCurrent_Y:{}'.format(
                        str(self.limit_samp),
                        str(num), str(cur_x), str(cur_y)))

                    if num >= self.limit_samp:
                        current_Xmin = min(PD_current_X)
                        num_minX = PD_current_X.index(current_Xmin)
                        current_Y = PD_current_Y[num_minX]

                        current_Ymin = min(PD_current_Y)
                        num_minY = PD_current_Y.index(current_Ymin)
                        current_X = PD_current_X[num_minY]

                        PER_X = 10 * np.log10(current_Y / current_Xmin)
                        PER_Y = 10 * np.log10(current_X / current_Ymin)
                        break

            self.ICRtf.write(':POL2:MODE MANUAL')
            # work left to do to plot the curve and save the picture
            pass
            # work left to do to count the time
            pass
            return [current_X, current_Y, PER_X, PER_Y, PD_current_X, PD_current_Y]  # ,PD_current_X,PD_current_Y]

        except Exception as e:
            print(e)

    def ICR_scamblePol_XYmax_Manual(self):
        '''
        Manual mode to find XY max status
        :param:GUI Class obj
        :return:[current_X,current_Y,PER_X,PER_Y]
        '''
        try:
            # 旋转偏振控制器，使得X/Y两路输出基本一致（相差10%）
            self.ICRtf.write(':POL2:MODE MANUAL');
            self.ICRtf.write(':POL2:MANUAL:SET1 0.01')
            self.ICRtf.write(':POL2:MANUAL:SET2 0.01')
            self.ICRtf.write(':POL2:MANUAL:SET3 0.01');

            num = 1
            current_X = 0
            current_Y = 0
            PD_current_X = []
            PD_current_Y = []
            cur_x, cur_y = [0, 0]
            while current_X * current_Y == 0:
                if self.limit == 0:  # 暗电流正常（暗电流过大，limit_samp有值；暗电流很小，limit_samp为空）
                    cur_x, cur_y = icr.TIA_getPDcurrent(self)[0:2]
                    if cur_x <= 0.004:
                        if current_Y == 0:
                            current_Y = cur_y
                        else:
                            current_Y = max(current_Y, cur_y)
                        current_Xmin = cur_x
                        PER_X = 10 * np.log10(current_Y / current_Xmin)
                    if cur_y <= 0.004:
                        if current_X == 0:
                            current_X = cur_x
                        else:
                            current_X = max(current_X, cur_x)
                        current_Ymin = cur_y
                        PER_Y = 10 * np.log10(current_X / current_Ymin)
                    PD_current_X.append(cur_x)
                    PD_current_Y.append(cur_y)
                    num += 1
                    print('---Rx signal port test---\nTotal {} points, now at NO.{}\nCurrent_X:{}\nCurrent_Y:{}'.format(
                        str(self.limit_samp),
                        str(num), str(cur_x), str(cur_y)))
                    if num >= self.limit_samp:
                        # update 0617/2022 improve judgement condition
                        if current_Y == 0:
                            current_Xmin = min(PD_current_X)
                            num_minX = PD_current_X.index(current_Xmin)
                            current_Y = PD_current_Y[num_minX]

                        if current_X == 0:
                            current_Ymin = min(PD_current_Y)
                            num_minY = PD_current_Y.index(current_Ymin)
                            current_X = PD_current_X[num_minY]

                        PER_X = 10 * np.log10(current_Y / current_Xmin)
                        PER_Y = 10 * np.log10(current_X / current_Ymin)
                        break
                else:  # 暗电流很大的情况
                    cur_x, cur_y = icr.TIA_getPDcurrent(self)[0:2]
                    PD_current_X.append(cur_x)
                    PD_current_Y.append(cur_y)
                    num += 1
                    print('---Rx signal port test---\nTotal {} points, now at NO.{}\nCurrent_X:{}\nCurrent_Y:{}'.format(
                        str(self.limit_samp),
                        str(num), str(cur_x), str(cur_y)))

                    if num >= self.limit_samp:
                        current_Xmin = min(PD_current_X)
                        num_minX = PD_current_X.index(current_Xmin)
                        current_Y = PD_current_Y[num_minX]

                        current_Ymin = min(PD_current_Y)
                        num_minY = PD_current_Y.index(current_Ymin)
                        current_X = PD_current_X[num_minY]

                        PER_X = 10 * np.log10(current_Y / current_Xmin)
                        PER_Y = 10 * np.log10(current_X / current_Ymin)
                        break

            self.ICRtf.write(':POL2:MODE MANUAL')
            # work left to do to plot the curve and save the picture
            pass
            # work left to do to count the time
            pass
            return [current_X, current_Y, PER_X, PER_Y, PD_current_X, PD_current_Y]  # ,PD_current_X,PD_current_Y]

        except Exception as e:
            print(e)

    def MPC_Find_XYZ(self,Func,Plate,C,R,Inc,Maxmize,Delay=0):
        '''
        Function to find the position of X, Y, or Z that either maximizes or minimizes or balance the output power.
        :param Func:to choose RF PD X or Y, input with 'PDX' or 'PDY' or 'DIFF'
        :param Plate:choose the MPC X Y Z plate
        :param C:Central of the adjust range
        :param R:Half range of the chosed central
        :param Inc:Increasement of adjust step
        :param Maxmize:True/False to choose whether find maximun or minimun value
        :param Delay: the first two iteration with 0.2s Delay before reading,the last and fine tune without Delay
        :return:Updated plate position(Current to set)...
            self.ICRtf.write(':POL2:MODE MANUAL');
            # self.ICRtf.write(':SCRAMBLE:FUNCTION SIN');
            # self.ICRtf.write(':POL2:SCRAMBLE:FREQUENCY1 0.010');
            # self.ICRtf.write(':POL2:SCRAMBLE:FREQUENCY2 0.018');
            # self.ICRtf.write(':POL2:SCRAMBLE:FREQUENCY3 0.02');
            self.ICRtf.write(':POL2:MANUAL:SET1 0.01')
            self.ICRtf.write(':POL2:MANUAL:SET2 0.01')
            self.ICRtf.write(':POL2:MANUAL:SET3 0.01');
        '''
        self.ICRtf.write(':POL2:MODE MANUAL')
        cur_x, cur_y = [0, 0]
        pd_currentX = []
        pd_currentY = []
        steps=round(2*R/Inc)
        start=C-R
        stop=C+R
        ArrayValue=np.linspace(start,stop,steps).round(2)
        ArrayCur=np.zeros(steps)
        Current =0
        for i in range(steps):
            setCur=ArrayValue[i]
            if setCur<=0 or setCur>=0.998:
                if Maxmize:
                    Current=-200
                else:
                    Current=200
                ArrayCur[i]=Current
                continue
            if Plate == 'X':
                self.ICRtf.write(':POL2:MANUAL:SET1 {}'.format(str(setCur)))
                time.sleep(Delay)
                while not str(setCur)==str(round(float(self.ICRtf.query(':POL2:MANUAL:SET1?')),2)):
                    self.ICRtf.write(':POL2:MANUAL:SET1 {}'.format(str(setCur)))
                    time.sleep(Delay)
            if Plate == 'Y':
                self.ICRtf.write(':POL2:MANUAL:SET2 {}'.format(str(setCur)))
                time.sleep(Delay)
                while not str(setCur) == str(round(float(self.ICRtf.query(':POL2:MANUAL:SET2?')),2)):
                    self.ICRtf.write(':POL2:MANUAL:SET2 {}'.format(str(setCur)))
                    time.sleep(Delay)
            if Plate == 'Z':
                self.ICRtf.write(':POL2:MANUAL:SET3 {}'.format(str(setCur)))
                time.sleep(Delay)
                while not str(setCur) == str(round(float(self.ICRtf.query(':POL2:MANUAL:SET3?')),2)):
                    self.ICRtf.write(':POL2:MANUAL:SET3 {}'.format(str(setCur)))
                    time.sleep(Delay)
            #time.sleep(Delay)
            cur_x, cur_y = icr.TIA_getPDcurrent(self)[0:2]
            pd_currentX.append(cur_x)
            pd_currentY.append(cur_y)
            print('{} Plate position {} get PD current, \nX: {}\nY: {}'.format(Plate,setCur,cur_x,cur_y))
            if Func=='PDX':
                Current = cur_x
            if Func == 'PDY':
                Current = cur_y
            if Func == 'DIFF':
                Current = abs(cur_x-cur_y)
            ArrayCur[i] = Current

        List_Cur=ArrayCur.tolist()
        if Maxmize:
            ArrayIndex=List_Cur.index(max(List_Cur))
        else:
            ArrayIndex = List_Cur.index(min(List_Cur))

        setCur=ArrayValue[ArrayIndex]
        if Plate == 'X':
            self.ICRtf.write(':POL2:MANUAL:SET1 {}'.format(str(setCur)))
            #time.sleep(Delay)
            while not str(setCur) == str(round(float(self.ICRtf.query(':POL2:MANUAL:SET1?')),2)):
                self.ICRtf.write(':POL2:MANUAL:SET1 {}'.format(str(setCur)))
                #time.sleep(Delay)
        if Plate == 'Y':
            self.ICRtf.write(':POL2:MANUAL:SET2 {}'.format(str(setCur)))
            #time.sleep(Delay)
            while not str(setCur) == str(round(float(self.ICRtf.query(':POL2:MANUAL:SET2?')),2)):
                self.ICRtf.write(':POL2:MANUAL:SET2 {}'.format(str(setCur)))
                #time.sleep(Delay)
        if Plate == 'Z':
            self.ICRtf.write(':POL2:MANUAL:SET3 {}'.format(str(setCur)))
            #time.sleep(Delay)
            while not str(setCur) == str(round(float(self.ICRtf.query(':POL2:MANUAL:SET3?')),2)):
                self.ICRtf.write(':POL2:MANUAL:SET3 {}'.format(str(setCur)))
                #time.sleep(Delay)
        return setCur,pd_currentX,pd_currentY

    def MPC_GBS_Find_XYZ(self, Func, Plate, C, R, Inc, Maxmize, Delay=0):
        '''
        Model GBS-PDL-1
        Function to find the position of W, X, Y, or Z that either maximizes or minimizes or balance the output power.
        :param Func:to choose RF PD X or Y, input with 'PDX' or 'PDY' or 'DIFF'
        :param Plate:choose the MPC W X Y Z plate corresponding to ch 1,2,3,4
        :param C:Central of the adjust range
        :param R:Half range of the chosed central
        :param Inc:Increasement of adjust step
        :param Maxmize:True/False to choose whether find maximun or minimun value
        :param Delay: the first two iteration with 0.2s Delay before reading,the last and fine tune without Delay
        :return:Updated plate position(Current to set)...
            self.ICRtf.write(':POL2:MODE MANUAL');
            # self.ICRtf.write(':SCRAMBLE:FUNCTION SIN');
            # self.ICRtf.write(':POL2:SCRAMBLE:FREQUENCY1 0.010');
            # self.ICRtf.write(':POL2:SCRAMBLE:FREQUENCY2 0.018');
            # self.ICRtf.write(':POL2:SCRAMBLE:FREQUENCY3 0.02');
            self.ICRtf.write(':POL2:MANUAL:SET1 0.01')
            self.ICRtf.write(':POL2:MANUAL:SET2 0.01')
            self.ICRtf.write(':POL2:MANUAL:SET3 0.01');
        '''
        cur_x, cur_y = [0, 0]
        pd_currentX = []
        pd_currentY = []
        steps = round(2 * R / Inc)
        start = C - R
        stop = C + R
        ArrayValue = np.array(np.linspace(start, stop, steps).round(),dtype='int')
        ArrayCur = np.zeros(steps)
        Current = 0
        for i in range(steps):
            setCur = ArrayValue[i]
            if setCur <= 0 or setCur >= 139.998:
                if Maxmize:
                    Current = -200
                else:
                    Current = 200
                ArrayCur[i] = Current
                continue
            if Plate == 'W':
                mpc.set_vol(self,1,setCur)
                time.sleep(Delay)
                while not str(setCur) == str(mpc.get_ch(self,1)):
                    mpc.set_vol(self,1,setCur)
                    time.sleep(Delay)
            if Plate == 'X':
                mpc.set_vol(self, 2, setCur)
                time.sleep(Delay)
                while not str(setCur) == str(mpc.get_ch(self, 2)):
                    mpc.set_vol(self, 2, setCur)
                    time.sleep(Delay)
            if Plate == 'Y':
                mpc.set_vol(self, 3, setCur)
                time.sleep(Delay)
                while not str(setCur) == str(mpc.get_ch(self, 3)):
                    mpc.set_vol(self, 3, setCur)
                    time.sleep(Delay)
            if Plate == 'Z':
                mpc.set_vol(self, 4, setCur)
                time.sleep(Delay)
                while not str(setCur) == str(mpc.get_ch(self, 4)):
                    mpc.set_vol(self, 4, setCur)
                    time.sleep(Delay)
            # time.sleep(Delay)
            cur_x, cur_y = icr.TIA_getPDcurrent(self)[0:2]
            pd_currentX.append(cur_x)
            pd_currentY.append(cur_y)
            print('{} Plate position {} get PD current, \nX: {}\nY: {}'.format(Plate, setCur, cur_x, cur_y))
            if Func == 'PDX':
                Current = cur_x
            if Func == 'PDY':
                Current = cur_y
            if Func == 'DIFF':
                Current = abs(cur_x - cur_y)
            ArrayCur[i] = Current
        List_Cur = ArrayCur.tolist()
        if Maxmize:
            ArrayIndex = List_Cur.index(max(List_Cur))
        else:
            ArrayIndex = List_Cur.index(min(List_Cur))
        setCur = ArrayValue[ArrayIndex]
        if Plate == 'W':
            mpc.set_vol(self, 1, setCur)
            time.sleep(Delay)
            while not str(setCur) == str(mpc.get_ch(self, 1)):
                mpc.set_vol(self, 1, setCur)
                time.sleep(Delay)
        if Plate == 'X':
            mpc.set_vol(self, 2, setCur)
            time.sleep(Delay)
            while not str(setCur) == str(mpc.get_ch(self, 2)):
                mpc.set_vol(self, 2, setCur)
                time.sleep(Delay)
        if Plate == 'Y':
            mpc.set_vol(self, 3, setCur)
            time.sleep(Delay)
            while not str(setCur) == str(mpc.get_ch(self, 3)):
                mpc.set_vol(self, 3, setCur)
                time.sleep(Delay)
        if Plate == 'Z':
            mpc.set_vol(self, 4, setCur)
            time.sleep(Delay)
            while not str(setCur) == str(mpc.get_ch(self, 4)):
                mpc.set_vol(self, 4, setCur)
                time.sleep(Delay)
        return setCur, pd_currentX, pd_currentY

    def MPC_Find_XYZ_ScopeBalance(self, Func, Plate, C, R, Inc, Maxmize, Delay=0):
        '''
        Function to find the position of X, Y, or Z that either maximizes or minimizes or balance the Scope output power.
        :param Func:to choose RF PD X or Y, input with 'PDX' or 'PDY' or 'DIFF',in this case on 'DIFF' can be choose
        :param Plate:choose the MPC X Y Z plate
        :param C:Central of the adjust range
        :param R:Half range of the chosed central
        :param Inc:Increasement of adjust step
        :param Maxmize:True/False to choose whether find maximun or minimun value,in this case only minimun differ
        :param Delay: the first two iteration with 0.2s Delay before reading,the last and fine tune without Delay
        :return:Updated plate position(Current to set)...
            self.ICRtf.write(':POL2:MODE MANUAL');
            # self.ICRtf.write(':SCRAMBLE:FUNCTION SIN');
            # self.ICRtf.write(':POL2:SCRAMBLE:FREQUENCY1 0.010');
            # self.ICRtf.write(':POL2:SCRAMBLE:FREQUENCY2 0.018');
            # self.ICRtf.write(':POL2:SCRAMBLE:FREQUENCY3 0.02');
            self.ICRtf.write(':POL2:MANUAL:SET1 0.01')
            self.ICRtf.write(':POL2:MANUAL:SET2 0.01')
            self.ICRtf.write(':POL2:MANUAL:SET3 0.01');
        '''
        self.ICRtf.write(':POL2:MODE MANUAL')
        cur_x, cur_y = [0, 0]
        steps = round(2 * R / Inc)
        start = C - R
        stop = C + R
        ArrayValue = np.linspace(start, stop, steps).round(2)
        ArrayCur = np.zeros(steps)
        Current = 0
        for i in range(steps):
            setCur = ArrayValue[i]
            if setCur <= 0 or setCur >= 0.998:
                if Maxmize:
                    Current = -200
                else:
                    Current = 200000
                ArrayCur[i] = Current
                continue
            if Plate == 'X':
                self.ICRtf.write(':POL2:MANUAL:SET1 {}'.format(str(setCur)))
                time.sleep(Delay)
                while not str(setCur) == str(round(float(self.ICRtf.query(':POL2:MANUAL:SET1?')), 2)):
                    self.ICRtf.write(':POL2:MANUAL:SET1 {}'.format(str(setCur)))
                    time.sleep(Delay)
            if Plate == 'Y':
                self.ICRtf.write(':POL2:MANUAL:SET2 {}'.format(str(setCur)))
                time.sleep(Delay)
                while not str(setCur) == str(round(float(self.ICRtf.query(':POL2:MANUAL:SET2?')), 2)):
                    self.ICRtf.write(':POL2:MANUAL:SET2 {}'.format(str(setCur)))
                    time.sleep(Delay)
            if Plate == 'Z':
                self.ICRtf.write(':POL2:MANUAL:SET3 {}'.format(str(setCur)))
                time.sleep(Delay)
                while not str(setCur) == str(round(float(self.ICRtf.query(':POL2:MANUAL:SET3?')), 2)):
                    self.ICRtf.write(':POL2:MANUAL:SET3 {}'.format(str(setCur)))
                    time.sleep(Delay)

            # query the data of amtiplitude
            A1 = self.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')
            A3 = self.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')
            while not icr.isfloat(A1) or not icr.isfloat(A3):
                A1 = self.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')
                A3 = self.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')
            print('{} Plate position {} get Scope amplitude, \nA1: {}\nA3: {}'.format(Plate, setCur, A1, A3))
            if Func == 'PDX':
                Current = cur_x
            if Func == 'PDY':
                Current = cur_y
            if Func == 'DIFF':
                Current = float(A1)/float(A3)
                # Current = abs(float(A1) - float(A3))
            print('The differ is:{}'.format(Current))
            ArrayCur[i] = Current
        List_Cur = ArrayCur.tolist()
        if Maxmize:
            ArrayIndex = List_Cur.index(max(List_Cur))
        else:
            ArrayIndex = List_Cur.index(min(List_Cur))
        setCur = ArrayValue[ArrayIndex]
        AmpDiff= ArrayCur[ArrayIndex]
        if Plate == 'X':
            self.ICRtf.write(':POL2:MANUAL:SET1 {}'.format(str(setCur)))
            # time.sleep(Delay)
            while not str(setCur) == str(round(float(self.ICRtf.query(':POL2:MANUAL:SET1?')), 2)):
                self.ICRtf.write(':POL2:MANUAL:SET1 {}'.format(str(setCur)))
                # time.sleep(Delay)
        if Plate == 'Y':
            self.ICRtf.write(':POL2:MANUAL:SET2 {}'.format(str(setCur)))
            # time.sleep(Delay)
            while not str(setCur) == str(round(float(self.ICRtf.query(':POL2:MANUAL:SET2?')), 2)):
                self.ICRtf.write(':POL2:MANUAL:SET2 {}'.format(str(setCur)))
                # time.sleep(Delay)
        if Plate == 'Z':
            self.ICRtf.write(':POL2:MANUAL:SET3 {}'.format(str(setCur)))
            # time.sleep(Delay)
            while not str(setCur) == str(round(float(self.ICRtf.query(':POL2:MANUAL:SET3?')), 2)):
                self.ICRtf.write(':POL2:MANUAL:SET3 {}'.format(str(setCur)))
                # time.sleep(Delay)
        return setCur,List_Cur

    '''
        First iteration:C=0.5;R=0.5;Inc=0.08
        Second iteration:C=Updated one;R=0.25;Inc=0.03
        Third iteration:C=Updated one;R=0.08;Inc=0.01
    '''

    def MPC_Find_Position(self, Func, Maxmize, limit=0.0012):
        '''
        Function to find the position either to maximize or minimize the output power.
        :param Func:PDX,PDY or DIFF to choose the function to adjust, either X min/max Y min/max or balance
        :param Maxmize:choose whether to find the maximun current of the PD
        :return:pd_currentX,pd_currentY, this algorithm adjust the MPC to selection status and keep.
        '''
        setCurX = 0
        setCurY = 0
        setCurZ = 0
        curX = 0
        curY = 0
        pd_currentX = []
        pd_currentY = []
        index_mark = []
        # first set the three plate to the middle
        # self.ICRtf.write(':POL2:MANUAL:SET1 0.5')
        # self.ICRtf.write(':POL2:MANUAL:SET2 0.5')
        # self.ICRtf.write(':POL2:MANUAL:SET3 0.5')
        # time.sleep(0.5)
        # the first iteration
        C = 0.5
        R = 0.5
        Inc = 0.08
        print('Iteration NO.1:')
        setCurY, curX, curY = self.MPC_Find_XYZ(Func, 'Y', C, R, Inc, Maxmize, 0.1)
        pd_currentX += curX
        pd_currentY += curY
        index_mark.append(len(pd_currentX))
        if (Func == 'PDX' and not Maxmize) or (Func == 'PDY' and Maxmize):
            if min(pd_currentX) < limit: return pd_currentX, pd_currentY, index_mark
        else:
            if min(pd_currentY) < limit: return pd_currentX, pd_currentY, index_mark
        setCurX, curX, curY = self.MPC_Find_XYZ(Func, 'X', C, R, Inc, Maxmize, 0.1)
        pd_currentX += curX
        pd_currentY += curY
        index_mark.append(len(pd_currentX))
        if (Func == 'PDX' and not Maxmize) or (Func == 'PDY' and Maxmize):
            if min(pd_currentX) < limit: return pd_currentX, pd_currentY, index_mark
        else:
            if min(pd_currentY) < limit: return pd_currentX, pd_currentY, index_mark
        setCurZ, curX, curY = self.MPC_Find_XYZ(Func, 'Z', C, R, Inc, Maxmize, 0.1)
        pd_currentX += curX
        pd_currentY += curY
        index_mark.append(len(pd_currentX))
        if (Func == 'PDX' and not Maxmize) or (Func == 'PDY' and Maxmize):
            if min(pd_currentX) < limit: return pd_currentX, pd_currentY, index_mark
        else:
            if min(pd_currentY) < limit: return pd_currentX, pd_currentY, index_mark
        # the second iteration
        R = 0.4
        Inc = 0.04
        print('Iteration NO.2:')
        setCurY, curX, curY = self.MPC_Find_XYZ(Func, 'Y', setCurY, R, Inc, Maxmize)
        pd_currentX += curX
        pd_currentY += curY
        index_mark.append(len(pd_currentX))
        if (Func == 'PDX' and not Maxmize) or (Func == 'PDY' and Maxmize):
            if min(pd_currentX) < limit: return pd_currentX, pd_currentY, index_mark
        else:
            if min(pd_currentY) < limit: return pd_currentX, pd_currentY, index_mark
        setCurX, curX, curY = self.MPC_Find_XYZ(Func, 'X', setCurX, R, Inc, Maxmize)
        pd_currentX += curX
        pd_currentY += curY
        index_mark.append(len(pd_currentX))
        if (Func == 'PDX' and not Maxmize) or (Func == 'PDY' and Maxmize):
            if min(pd_currentX) < limit: return pd_currentX, pd_currentY, index_mark
        else:
            if min(pd_currentY) < limit: return pd_currentX, pd_currentY, index_mark
        setCurZ, curX, curY = self.MPC_Find_XYZ(Func, 'Z', setCurZ, R, Inc, Maxmize)
        pd_currentX += curX
        pd_currentY += curY
        index_mark.append(len(pd_currentX))
        if (Func == 'PDX' and not Maxmize) or (Func == 'PDY' and Maxmize):
            if min(pd_currentX) < limit: return pd_currentX, pd_currentY, index_mark
        else:
            if min(pd_currentY) < limit: return pd_currentX, pd_currentY, index_mark
        # the third iteration
        R = 0.2
        Inc = 0.02
        print('Iteration NO.3:')
        setCurY, curX, curY = self.MPC_Find_XYZ(Func, 'Y', setCurY, R, Inc, Maxmize)
        pd_currentX += curX
        pd_currentY += curY
        index_mark.append(len(pd_currentX))
        if (Func == 'PDX' and not Maxmize) or (Func == 'PDY' and Maxmize):
            if min(pd_currentX) < limit: return pd_currentX, pd_currentY, index_mark
        else:
            if min(pd_currentY) < limit: return pd_currentX, pd_currentY, index_mark
        setCurX, curX, curY = self.MPC_Find_XYZ(Func, 'X', setCurX, R, Inc, Maxmize)
        pd_currentX += curX
        pd_currentY += curY
        index_mark.append(len(pd_currentX))
        if (Func == 'PDX' and not Maxmize) or (Func == 'PDY' and Maxmize):
            if min(pd_currentX) < limit: return pd_currentX, pd_currentY, index_mark
        else:
            if min(pd_currentY) < limit: return pd_currentX, pd_currentY, index_mark
        setCurZ, curX, curY = self.MPC_Find_XYZ(Func, 'Z', setCurZ, R, Inc, Maxmize)
        pd_currentX += curX
        pd_currentY += curY
        index_mark.append(len(pd_currentX))
        if (Func == 'PDX' and not Maxmize) or (Func == 'PDY' and Maxmize):
            if min(pd_currentX) < limit: return pd_currentX, pd_currentY, index_mark
        else:
            if min(pd_currentY) < limit: return pd_currentX, pd_currentY, index_mark

        # the forth iteration
        R = 0.1
        Inc = 0.01
        print('Iteration NO.4:')
        setCurY, curX, curY = self.MPC_Find_XYZ(Func, 'Y', setCurY, R, Inc, Maxmize)
        pd_currentX += curX
        pd_currentY += curY
        index_mark.append(len(pd_currentX))
        if (Func == 'PDX' and not Maxmize) or (Func == 'PDY' and Maxmize):
            if min(pd_currentX) < limit: return pd_currentX, pd_currentY, index_mark
        else:
            if min(pd_currentY) < limit: return pd_currentX, pd_currentY, index_mark
        setCurX, curX, curY = self.MPC_Find_XYZ(Func, 'X', setCurX, R, Inc, Maxmize)
        pd_currentX += curX
        pd_currentY += curY
        index_mark.append(len(pd_currentX))
        if (Func == 'PDX' and not Maxmize) or (Func == 'PDY' and Maxmize):
            if min(pd_currentX) < limit: return pd_currentX, pd_currentY, index_mark
        else:
            if min(pd_currentY) < limit: return pd_currentX, pd_currentY, index_mark
        setCurZ, curX, curY = self.MPC_Find_XYZ(Func, 'Z', setCurZ, R, Inc, Maxmize)
        pd_currentX += curX
        pd_currentY += curY
        index_mark.append(len(pd_currentX))
        return pd_currentX, pd_currentY, index_mark

    def MPC_Find_Position_Fine(self,Func,Maxmize):
        '''
        fine tune to small adjust
        Function to find the position either to maximize or minimize the output power.
        :param Func:PDX,PDY or DIFF to choose the function to adjust, either X min/max Y min/max or balance
        :param Maxmize:choose whether to find the maximun current of the PD
        :return:NA, this algorithm adjust the MPC to selection status and keep.
        '''
        setCurX = round(float(self.ICRtf.query(':POL2:MANUAL:SET1?')),2)
        setCurY = round(float(self.ICRtf.query(':POL2:MANUAL:SET2?')),2)
        setCurZ = round(float(self.ICRtf.query(':POL2:MANUAL:SET3?')),2)
        # the fine iteration
        R = 0.05
        Inc = 0.01
        print('Iteration Fine tune:')
        setCurX = self.MPC_Find_XYZ(Func, 'X', setCurX, R, Inc, Maxmize)
        setCurY = self.MPC_Find_XYZ(Func, 'Y', setCurY, R, Inc, Maxmize)
        setCurZ = self.MPC_Find_XYZ(Func, 'Z', setCurZ, R, Inc, Maxmize)


    def MPC_Find_Position_ScopeBalance(self,Func='DIFF',Maxmize=False):
        '''
        Function to find the position either to maximize or minimize the output power by scope voltage reading.
        :param Func:PDX,PDY or DIFF to choose the function to adjust, either X min/max Y min/max or balance, onnly balance this case
        :param Maxmize:choose whether to find the maximun current of the diff, only minimun this case
        :return:NA, this algorithm adjust the MPC to selection status and keep.
        '''
        setCurX = 0
        setCurY = 0
        setCurZ = 0
        ampDiff=0
        ampDiffRecord=[]
        count=0
        while count<12 or not(0.92<min(ampDiffRecord)<1.08):
            # the first iteration
            C = 0.5
            R = 0.5
            Inc = 0.1
            print('Iteration NO.1:')
            setCurY,ampDiff = self.MPC_Find_XYZ_ScopeBalance(Func, 'Y', C, R, Inc, Maxmize, 0.1)
            ampDiffRecord+=ampDiff
            count+=1
            setCurX,ampDiff = self.MPC_Find_XYZ_ScopeBalance(Func, 'X', C, R, Inc, Maxmize, 0.1)
            ampDiffRecord += ampDiff
            count += 1
            setCurZ,ampDiff = self.MPC_Find_XYZ_ScopeBalance(Func, 'Z', C, R, Inc, Maxmize, 0.1)
            ampDiffRecord += ampDiff
            count += 1
            # the second iteration
            R = 0.25
            Inc = 0.05
            print('Iteration NO.2:')
            setCurY,ampDiff  = self.MPC_Find_XYZ_ScopeBalance(Func, 'Y', setCurY, R, Inc, Maxmize, 0.1)
            ampDiffRecord += ampDiff
            count += 1
            setCurX,ampDiff  = self.MPC_Find_XYZ_ScopeBalance(Func, 'X', setCurX, R, Inc, Maxmize, 0.1)
            ampDiffRecord += ampDiff
            count += 1
            setCurZ,ampDiff  = self.MPC_Find_XYZ_ScopeBalance(Func, 'Z', setCurZ, R, Inc, Maxmize, 0.1)
            ampDiffRecord += ampDiff
            count += 1
            # the third iteration
            R = 0.08
            Inc = 0.02
            print('Iteration NO.3:')
            setCurY,ampDiff  = self.MPC_Find_XYZ_ScopeBalance(Func, 'Y', setCurY, R, Inc, Maxmize)
            ampDiffRecord += ampDiff
            count += 1
            setCurX,ampDiff  = self.MPC_Find_XYZ_ScopeBalance(Func, 'X', setCurX, R, Inc, Maxmize)
            ampDiffRecord += ampDiff
            count += 1
            setCurZ,ampDiff  = self.MPC_Find_XYZ_ScopeBalance(Func, 'Z', setCurZ, R, Inc, Maxmize)
            ampDiffRecord += ampDiff
            count += 1
            # the forth iteration
            R = 0.04
            Inc = 0.01
            print('Iteration NO.4:')
            setCurY,ampDiff  = self.MPC_Find_XYZ_ScopeBalance(Func, 'Y', setCurY, R, Inc, Maxmize)
            ampDiffRecord += ampDiff
            count += 1
            setCurX,ampDiff  = self.MPC_Find_XYZ_ScopeBalance(Func, 'X', setCurX, R, Inc, Maxmize)
            ampDiffRecord += ampDiff
            count += 1
            setCurZ,ampDiff  = self.MPC_Find_XYZ_ScopeBalance(Func, 'Z', setCurZ, R, Inc, Maxmize)
            ampDiffRecord += ampDiff
            count += 1

    def MPC_Find_Position_ScopeBalance_Fine(self,Func='DIFF',Maxmize=False):
        '''
        fine tune to small adjust scope balance output
        Function to find the position either to maximize or minimize the output power.
        :param Func:PDX,PDY or DIFF to choose the function to adjust, either X min/max Y min/max or balance
        :param Maxmize:choose whether to find the maximun current of the PD
        :return:NA, this algorithm adjust the MPC to selection status and keep.
        '''
        setCurX = round(float(self.ICRtf.query(':POL2:MANUAL:SET1?')),2)
        setCurY = round(float(self.ICRtf.query(':POL2:MANUAL:SET2?')),2)
        setCurZ = round(float(self.ICRtf.query(':POL2:MANUAL:SET3?')),2)
        # the fine iteration
        R = 0.03
        Inc = 0.01
        print('Iteration Fine tune:')
        setCurY = self.MPC_Find_XYZ_ScopeBalance(Func, 'Y', setCurY, R, Inc, Maxmize)
        setCurX = self.MPC_Find_XYZ_ScopeBalance(Func, 'X', setCurX, R, Inc, Maxmize)
        setCurZ = self.MPC_Find_XYZ_ScopeBalance(Func, 'Z', setCurZ, R, Inc, Maxmize)

    def MPC_GBS_Find_Position_OLD(self,Func,Maxmize,limit=0.0012):
        '''
        MPC model is GBS-PDL-1
        Function to find the position either to maximize or minimize the output power.
        :param Func:PDX,PDY or DIFF to choose the function to adjust, either X min/max Y min/max or balance
        :param Maxmize:choose whether to find the maximun current of the PD
        :return:pd_currentX,pd_currentY, this algorithm adjust the MPC to selection status and keep.
        '''
        setCurW = 0
        setCurX = 0
        setCurY = 0
        setCurZ = 0
        curX=0
        curY=0
        pd_currentX=[]
        pd_currentY=[]
        index_mark=[]
        #first set the three plate to the middle
        # self.ICRtf.write(':POL2:MANUAL:SET1 0.5')
        # self.ICRtf.write(':POL2:MANUAL:SET2 0.5')
        # self.ICRtf.write(':POL2:MANUAL:SET3 0.5')
        # time.sleep(0.5)
        # the first iteration
        C = 70
        R = 70
        Inc = 12
        print('Iteration NO.1:')
        setCurW, curX, curY = self.MPC_Find_XYZ(Func, 'W', C, R, Inc, Maxmize, 0.1)
        pd_currentX += curX
        pd_currentY += curY
        index_mark.append(len(pd_currentX))
        if (Func == 'PDX' and not Maxmize) or (Func == 'PDY' and Maxmize):
            if min(pd_currentX) < limit: return pd_currentX, pd_currentY
        else:
            if min(pd_currentY) < limit: return pd_currentX, pd_currentY
        setCurY,curX,curY = self.MPC_Find_XYZ(Func, 'Y', C, R, Inc, Maxmize,0.1)
        pd_currentX += curX
        pd_currentY += curY
        index_mark.append(len(pd_currentX))
        if (Func=='PDX' and not Maxmize) or (Func=='PDY' and Maxmize):
            if min(pd_currentX)<limit:return pd_currentX,pd_currentY
        else:
            if min(pd_currentY)<limit:return pd_currentX,pd_currentY
        setCurX,curX,curY = self.MPC_Find_XYZ(Func, 'X', C, R, Inc, Maxmize,0.1)
        pd_currentX+=curX
        pd_currentY+=curY
        index_mark.append(len(pd_currentX))
        if (Func=='PDX' and not Maxmize) or (Func=='PDY' and Maxmize):
            if min(pd_currentX)<limit:return pd_currentX,pd_currentY
        else:
            if min(pd_currentY)<limit:return pd_currentX,pd_currentY
        setCurZ,curX,curY = self.MPC_Find_XYZ(Func, 'Z', C, R, Inc, Maxmize,0.1)
        pd_currentX += curX
        pd_currentY += curY
        index_mark.append(len(pd_currentX))
        if (Func=='PDX' and not Maxmize) or (Func=='PDY' and Maxmize):
            if min(pd_currentX)<limit:return pd_currentX,pd_currentY
        else:
            if min(pd_currentY)<limit:return pd_currentX,pd_currentY
        # the second iteration
        R = 36
        Inc = 6
        print('Iteration NO.2:')
        setCurW, curX, curY = self.MPC_Find_XYZ(Func, 'W', setCurW, R, Inc, Maxmize, 0.1)
        pd_currentX += curX
        pd_currentY += curY
        index_mark.append(len(pd_currentX))
        if (Func == 'PDX' and not Maxmize) or (Func == 'PDY' and Maxmize):
            if min(pd_currentX) < limit: return pd_currentX, pd_currentY
        else:
            if min(pd_currentY) < limit: return pd_currentX, pd_currentY
        setCurY,curX,curY = self.MPC_Find_XYZ(Func, 'Y', setCurY, R, Inc, Maxmize)
        pd_currentX += curX
        pd_currentY += curY
        index_mark.append(len(pd_currentX))
        if (Func=='PDX' and not Maxmize) or (Func=='PDY' and Maxmize):
            if min(pd_currentX)<limit:return pd_currentX,pd_currentY
        else:
            if min(pd_currentY)<limit:return pd_currentX,pd_currentY
        setCurX,curX,curY = self.MPC_Find_XYZ(Func, 'X', setCurX, R, Inc, Maxmize)
        pd_currentX += curX
        pd_currentY += curY
        index_mark.append(len(pd_currentX))
        if (Func=='PDX' and not Maxmize) or (Func=='PDY' and Maxmize):
            if min(pd_currentX)<limit:return pd_currentX,pd_currentY
        else:
            if min(pd_currentY)<limit:return pd_currentX,pd_currentY
        setCurZ,curX,curY = self.MPC_Find_XYZ(Func, 'Z', setCurZ, R, Inc, Maxmize)
        pd_currentX += curX
        pd_currentY += curY
        index_mark.append(len(pd_currentX))
        if (Func=='PDX' and not Maxmize) or (Func=='PDY' and Maxmize):
            if min(pd_currentX)<limit:return pd_currentX,pd_currentY
        else:
            if min(pd_currentY)<limit:return pd_currentX,pd_currentY
        # the third iteration
        R = 18
        Inc = 3
        print('Iteration NO.3:')
        setCurW, curX, curY = self.MPC_Find_XYZ(Func, 'W', setCurW, R, Inc, Maxmize, 0.1)
        pd_currentX += curX
        pd_currentY += curY
        index_mark.append(len(pd_currentX))
        if (Func == 'PDX' and not Maxmize) or (Func == 'PDY' and Maxmize):
            if min(pd_currentX) < limit: return pd_currentX, pd_currentY
        else:
            if min(pd_currentY) < limit: return pd_currentX, pd_currentY
        setCurY,curX,curY = self.MPC_Find_XYZ(Func, 'Y', setCurY, R, Inc, Maxmize)
        pd_currentX += curX
        pd_currentY += curY
        index_mark.append(len(pd_currentX))
        if (Func=='PDX' and not Maxmize) or (Func=='PDY' and Maxmize):
            if min(pd_currentX)<limit:return pd_currentX,pd_currentY
        else:
            if min(pd_currentY)<limit:return pd_currentX,pd_currentY
        setCurX,curX,curY = self.MPC_Find_XYZ(Func, 'X', setCurX, R, Inc, Maxmize)
        pd_currentX += curX
        pd_currentY += curY
        index_mark.append(len(pd_currentX))
        if (Func=='PDX' and not Maxmize) or (Func=='PDY' and Maxmize):
            if min(pd_currentX)<limit:return pd_currentX,pd_currentY
        else:
            if min(pd_currentY)<limit:return pd_currentX,pd_currentY
        setCurZ,curX,curY = self.MPC_Find_XYZ(Func, 'Z', setCurZ, R, Inc, Maxmize)
        pd_currentX += curX
        pd_currentY += curY
        index_mark.append(len(pd_currentX))
        if (Func=='PDX' and not Maxmize) or (Func=='PDY' and Maxmize):
            if min(pd_currentX)<limit:return pd_currentX,pd_currentY
        else:
            if min(pd_currentY)<limit:return pd_currentX,pd_currentY

        # the forth iteration
        R = 6
        Inc = 1
        print('Iteration NO.4:')
        setCurW, curX, curY = self.MPC_Find_XYZ(Func, 'W', setCurW, R, Inc, Maxmize, 0.1)
        pd_currentX += curX
        pd_currentY += curY
        index_mark.append(len(pd_currentX))
        if (Func == 'PDX' and not Maxmize) or (Func == 'PDY' and Maxmize):
            if min(pd_currentX) < limit: return pd_currentX, pd_currentY
        else:
            if min(pd_currentY) < limit: return pd_currentX, pd_currentY
        setCurY,curX,curY = self.MPC_Find_XYZ(Func, 'Y', setCurY, R, Inc, Maxmize)
        pd_currentX += curX
        pd_currentY += curY
        index_mark.append(len(pd_currentX))
        if (Func=='PDX' and not Maxmize) or (Func=='PDY' and Maxmize):
            if min(pd_currentX)<limit:return pd_currentX,pd_currentY
        else:
            if min(pd_currentY)<limit:return pd_currentX,pd_currentY
        setCurX,curX,curY = self.MPC_Find_XYZ(Func, 'X', setCurX, R, Inc, Maxmize)
        pd_currentX += curX
        pd_currentY += curY
        index_mark.append(len(pd_currentX))
        if (Func=='PDX' and not Maxmize) or (Func=='PDY' and Maxmize):
            if min(pd_currentX)<limit:return pd_currentX,pd_currentY
        else:
            if min(pd_currentY)<limit:return pd_currentX,pd_currentY
        setCurZ,curX,curY = self.MPC_Find_XYZ(Func, 'Z', setCurZ, R, Inc, Maxmize)
        pd_currentX += curX
        pd_currentY += curY
        index_mark.append(len(pd_currentX))
        return pd_currentX,pd_currentY,index_mark

    def MPC_GBS_Find_Position(self,Func,Maxmize,limit=0.0012):
        '''
        MPC model is GBS-PDL-1
        Function to find the position either to maximize or minimize the output power.
        :param Func:PDX,PDY or DIFF to choose the function to adjust, either X min/max Y min/max or balance
        :param Maxmize:choose whether to find the maximun current of the PD
        :return:pd_currentX,pd_currentY, this algorithm adjust the MPC to selection status and keep.
        '''
        setCurW = 0
        setCurX = 0
        setCurY = 0
        setCurZ = 0
        curX=0
        curY=0
        pd_currentX=[]
        pd_currentY=[]
        index_mark=[]

        # the first iteration
        C = 95
        R = 45
        Inc = 10
        print('Iteration NO.1:')
        setCurW, curX, curY = self.MPC_GBS_Find_XYZ(Func, 'W', C, R, Inc, Maxmize,0.1)
        time.sleep(0.5)
        pd_currentX += curX
        pd_currentY += curY
        index_mark.append(len(pd_currentX))
        if (Func == 'PDX' and not Maxmize) or (Func == 'PDY' and Maxmize):
            if min(pd_currentX) < limit: return pd_currentX, pd_currentY,index_mark
        else:
            if min(pd_currentY) < limit: return pd_currentX, pd_currentY,index_mark
        setCurY,curX,curY = self.MPC_GBS_Find_XYZ(Func, 'Y', C, R, Inc, Maxmize,0.1)
        time.sleep(0.5)
        pd_currentX += curX
        pd_currentY += curY
        index_mark.append(len(pd_currentX))
        if (Func=='PDX' and not Maxmize) or (Func=='PDY' and Maxmize):
            if min(pd_currentX)<limit:return pd_currentX,pd_currentY,index_mark
        else:
            if min(pd_currentY)<limit:return pd_currentX,pd_currentY,index_mark
        setCurX,curX,curY = self.MPC_GBS_Find_XYZ(Func, 'X', C, R, Inc, Maxmize,0.1)
        time.sleep(0.5)
        pd_currentX+=curX
        pd_currentY+=curY
        index_mark.append(len(pd_currentX))
        if (Func=='PDX' and not Maxmize) or (Func=='PDY' and Maxmize):
            if min(pd_currentX)<limit:return pd_currentX,pd_currentY,index_mark
        else:
            if min(pd_currentY)<limit:return pd_currentX,pd_currentY,index_mark
        setCurZ,curX,curY = self.MPC_GBS_Find_XYZ(Func, 'Z', C, R, Inc, Maxmize,0.1)
        time.sleep(0.5)
        pd_currentX += curX
        pd_currentY += curY
        index_mark.append(len(pd_currentX))
        if (Func=='PDX' and not Maxmize) or (Func=='PDY' and Maxmize):
            if min(pd_currentX)<limit:return pd_currentX,pd_currentY,index_mark
        else:
            if min(pd_currentY)<limit:return pd_currentX,pd_currentY,index_mark

        #sweep channel 1 again:
        setCurW, curX, curY = self.MPC_GBS_Find_XYZ(Func, 'W', C, R, Inc, Maxmize, 0.1)
        time.sleep(0.5)
        pd_currentX += curX
        pd_currentY += curY
        index_mark.append(len(pd_currentX))
        if (Func == 'PDX' and not Maxmize) or (Func == 'PDY' and Maxmize):
            if min(pd_currentX) < limit: return pd_currentX, pd_currentY, index_mark
        else:
            if min(pd_currentY) < limit: return pd_currentX, pd_currentY, index_mark

        # the second iteration
        R = 18
        Inc = 3
        print('Iteration NO.2:')
        setCurW, curX, curY = self.MPC_GBS_Find_XYZ(Func, 'W', setCurW, R, Inc, Maxmize,0.1)
        time.sleep(0.5)
        pd_currentX += curX
        pd_currentY += curY
        index_mark.append(len(pd_currentX))
        if (Func == 'PDX' and not Maxmize) or (Func == 'PDY' and Maxmize):
            if min(pd_currentX) < limit: return pd_currentX, pd_currentY,index_mark
        else:
            if min(pd_currentY) < limit: return pd_currentX, pd_currentY,index_mark
        setCurY,curX,curY = self.MPC_GBS_Find_XYZ(Func, 'Y', setCurY, R, Inc, Maxmize,0.1)
        time.sleep(0.5)
        pd_currentX += curX
        pd_currentY += curY
        index_mark.append(len(pd_currentX))
        if (Func=='PDX' and not Maxmize) or (Func=='PDY' and Maxmize):
            if min(pd_currentX)<limit:return pd_currentX,pd_currentY,index_mark
        else:
            if min(pd_currentY)<limit:return pd_currentX,pd_currentY,index_mark
        # setCurX,curX,curY = self.MPC_GBS_Find_XYZ(Func, 'X', setCurX, R, Inc, Maxmize,0.1)
        # time.sleep(0.5)
        # pd_currentX += curX
        # pd_currentY += curY
        # index_mark.append(len(pd_currentX))
        # if (Func=='PDX' and not Maxmize) or (Func=='PDY' and Maxmize):
        #     if min(pd_currentX)<limit:return pd_currentX,pd_currentY,index_mark
        # else:
        #     if min(pd_currentY)<limit:return pd_currentX,pd_currentY,index_mark
        setCurZ,curX,curY = self.MPC_GBS_Find_XYZ(Func, 'Z', setCurZ, R, Inc, Maxmize,0.1)
        time.sleep(0.5)
        pd_currentX += curX
        pd_currentY += curY
        index_mark.append(len(pd_currentX))
        if (Func=='PDX' and not Maxmize) or (Func=='PDY' and Maxmize):
            if min(pd_currentX)<limit:return pd_currentX,pd_currentY,index_mark
        else:
            if min(pd_currentY)<limit:return pd_currentX,pd_currentY,index_mark

        # the third iteration
        R = 6
        Inc = 1
        print('Iteration NO.3:')
        setCurW, curX, curY = self.MPC_GBS_Find_XYZ(Func, 'W', setCurW, R, Inc, Maxmize,0.1)
        time.sleep(0.5)
        pd_currentX += curX
        pd_currentY += curY
        index_mark.append(len(pd_currentX))
        if (Func == 'PDX' and not Maxmize) or (Func == 'PDY' and Maxmize):
            if min(pd_currentX) < limit: return pd_currentX, pd_currentY,index_mark
        else:
            if min(pd_currentY) < limit: return pd_currentX, pd_currentY,index_mark
        setCurY,curX,curY = self.MPC_GBS_Find_XYZ(Func, 'Y', setCurY, R, Inc, Maxmize,0.1)
        time.sleep(0.5)
        pd_currentX += curX
        pd_currentY += curY
        index_mark.append(len(pd_currentX))
        if (Func=='PDX' and not Maxmize) or (Func=='PDY' and Maxmize):
            if min(pd_currentX)<limit:return pd_currentX,pd_currentY,index_mark
        else:
            if min(pd_currentY)<limit:return pd_currentX,pd_currentY,index_mark
        # setCurX,curX,curY = self.MPC_GBS_Find_XYZ(Func, 'X', setCurX, R, Inc, Maxmize,0.1)
        # time.sleep(0.5)
        # pd_currentX += curX
        # pd_currentY += curY
        # index_mark.append(len(pd_currentX))
        # if (Func=='PDX' and not Maxmize) or (Func=='PDY' and Maxmize):
        #     if min(pd_currentX)<limit:return pd_currentX,pd_currentY,index_mark
        # else:
        #     if min(pd_currentY)<limit:return pd_currentX,pd_currentY,index_mark
        setCurZ,curX,curY = self.MPC_GBS_Find_XYZ(Func, 'Z', setCurZ, R, Inc, Maxmize,0.1)
        time.sleep(0.5)
        pd_currentX += curX
        pd_currentY += curY
        index_mark.append(len(pd_currentX))
        if (Func=='PDX' and not Maxmize) or (Func=='PDY' and Maxmize):
            if min(pd_currentX)<limit:return pd_currentX,pd_currentY,index_mark
        else:
            if min(pd_currentY)<limit:return pd_currentX,pd_currentY,index_mark

        # the forth iteration
        R = 6
        Inc = 1
        print('Iteration NO.4:')
        setCurW, curX, curY = self.MPC_GBS_Find_XYZ(Func, 'W', setCurW, R, Inc, Maxmize,0.1)
        time.sleep(0.5)
        pd_currentX += curX
        pd_currentY += curY
        index_mark.append(len(pd_currentX))
        if (Func == 'PDX' and not Maxmize) or (Func == 'PDY' and Maxmize):
            if min(pd_currentX) < limit: return pd_currentX, pd_currentY,index_mark
        else:
            if min(pd_currentY) < limit: return pd_currentX, pd_currentY,index_mark
        setCurY,curX,curY = self.MPC_GBS_Find_XYZ(Func, 'Y', setCurY, R, Inc, Maxmize,0.1)
        time.sleep(0.5)
        pd_currentX += curX
        pd_currentY += curY
        index_mark.append(len(pd_currentX))
        if (Func=='PDX' and not Maxmize) or (Func=='PDY' and Maxmize):
            if min(pd_currentX)<limit:return pd_currentX,pd_currentY,index_mark
        else:
            if min(pd_currentY)<limit:return pd_currentX,pd_currentY,index_mark
        # setCurX,curX,curY = self.MPC_GBS_Find_XYZ(Func, 'X', setCurX, R, Inc, Maxmize,0.1)
        # time.sleep(0.5)
        # pd_currentX += curX
        # pd_currentY += curY
        # index_mark.append(len(pd_currentX))
        # if (Func=='PDX' and not Maxmize) or (Func=='PDY' and Maxmize):
        #     if min(pd_currentX)<limit:return pd_currentX,pd_currentY,index_mark
        # else:
        #     if min(pd_currentY)<limit:return pd_currentX,pd_currentY,index_mark
        setCurZ,curX,curY = self.MPC_GBS_Find_XYZ(Func, 'Z', setCurZ, R, Inc, Maxmize,0.1)
        time.sleep(0.5)
        pd_currentX += curX
        pd_currentY += curY
        index_mark.append(len(pd_currentX))
        return pd_currentX,pd_currentY,index_mark


    def create_report_folders(self,sn,timeStamp):
        '''
        This function is to create the local folder and check the network share folder namely 'R:\' is exists
        :param sn:SN
        :param test_flag:DC,TxBW or ICR test
        :param test_type:Normal,debug or others... test type
        :return1: True:'R:\' exists False: not exist
        :return2: the folder path to copy the local test data to
        '''
        # Test data storage
        if not os.path.exists(self.report_path):
            os.mkdir(self.report_path)
        report_path1 = os.path.join(self.report_path, self.test_flag)  # create the child folder to store data
        if not os.path.exists(report_path1):
            os.mkdir(report_path1)
        report_path2 = os.path.join(report_path1, self.test_type)  # create the child folder to store data
        if not os.path.exists(report_path2):
            os.mkdir(report_path2)
        # Add the folder named as SN_Date to store all the test data
        report_path3 = os.path.join(report_path2, sn + '_' + timeStamp)  # create the child folder to store data
        if not os.path.exists(report_path3):
            os.mkdir(report_path3)
        #Create the network paths
        networkFolder=r'R:\\'
        if not os.path.exists(networkFolder):
            print('网盘共享路径{}不存在，请检查网络连接或其他原因！'.format(networkFolder))
            return report_path3,False
        #Normal,debug or others... test type
        network_path=os.path.join(networkFolder,self.test_type)
        if not os.path.exists(network_path):
            os.mkdir(network_path)
        #SN folder
        network_path1 = os.path.join(network_path, sn)
        if not os.path.exists(network_path1):
            os.mkdir(network_path1)
        #test type:DC,TxBW or ICR
        network_path2 = os.path.join(network_path1, self.test_flag)
        if not os.path.exists(network_path2):
            os.mkdir(network_path2)
        #Date time folder
        network_path3 = os.path.join(network_path2, timeStamp)
        if not os.path.exists(network_path2):
            os.mkdir(network_path2)
        return report_path3,network_path3


class WorkThread(QThread):
    '''
    This class is to perform the multiple threads stand alone the main GUI.
    including the main test processes:1.DC;2.Tx BW;3.ICR test(PE skew,Rx BW,Responsivity)
    '''
    sig_progress=pyqtSignal(int)
    sig_status=pyqtSignal(str)
    sig_staColor=pyqtSignal(str)
    sig_print=pyqtSignal(str)
    sig_but=pyqtSignal(str)
    sig_clear=pyqtSignal()
    sig_goca=pyqtSignal(int)
    #plot function
    sig_plot=pyqtSignal(DataFrame,str,str,str,str,str,list,list)
    sig_plotClose=pyqtSignal()

    def __init__(self):
        '''
        inheriate the propertity from class WorkThread
        '''
        super(WorkThread, self).__init__()

    def run(self):
        try:
            if mawin.test_flag=="DC":
                self.DC_test()
            elif mawin.test_flag=="TxBW":
                self.TxBW_test()
            elif mawin.test_flag=="ICR":
                self.ICR_test_main()#PE-BW old program,Res with new algorithm and MTP1000
            elif mawin.test_flag=="RxRes":
                self.Res_test_GBS()
            elif mawin.test_flag=="PeSkew":
                self.PEskew_test()
            elif mawin.test_flag == "VOAcal":
                self.VOA_calibration_New()
            else:
                print('No correct test flag selected, this will not happen :)')
        except Exception as e:
            print(e)
            self.sig_print.emit(str(e))

    def DC_test_OLD(self):
        '''
        DC test main process
        :return:NA
        '''
        #Get and judge the SN format
        sn=str(mawin.lineEdit.text()).strip()
        if not gf.SN_check(sn):
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
        #New method to check the connectivity and check TIA/DRV is ALU or IDT
        con=0
        con=cb.test_connectivity_new(mawin,sn)
        if con==1:
            self.sig_print.emit('Driver is ALU')
            mawin.device_type='ALU'
            print('检测到ALU器件,将执行ALU器件相关的测试...')
        elif con==2:
            self.sig_print.emit('Driver is IDT')
            mawin.device_type='IDT'
            print('检测到IDT器件,将执行IDT器件相关的测试...')
        elif con==3:
            self.sig_print.emit('请检查光路或压接，连接性检查失败！！！')
            self.sig_staColor.emit('blue')
            self.sig_but.emit('开始')
            self.sig_status.emit('请检查器件连接!')
            #self.sig_stoptest.emit()
            self.sig_progress.emit(0)
            return
        if mawin.device_type=='ALU':
            #mawin.board_up = os.path.join(self.config_path, 'Setup_brdup_CtrlboardA001_20220113_ALU.txt')
            mawin.drv_up = os.path.join(mawin.config_path, 'Setup_driverup_Ctrlboard56017837A002_20211203_ALU.txt')
            mawin.drv_down = os.path.join(mawin.config_path, 'Setup_driverdown_CtrlboardA001_20210820_ALU.txt')
        if mawin.device_type=='IDT':
            #mawin.board_up = os.path.join(self.config_path, 'Setup_brdup_CtrlboardA001_20220113_ALU.txt')
            mawin.drv_up = os.path.join(mawin.config_path, 'Setup_driverup_Ctrlboard56017837A002_20211203.txt')
            mawin.drv_down = os.path.join(mawin.config_path, 'Setup_driverdown_CtrlboardA001_20210820.txt')
        '''
        *****Start the whole test steps*****
        '''
        #Read the CLPD,FPGA,MCU information
        testBrdClp='NA'
        ctlBrdClp='NA'
        ctlBrdModVer='NA'
        ctlBrdFPGAver='NA'
        MCUver='NA'
        testBrdClp,ctlBrdClp,ctlBrdModVer,ctlBrdFPGAver,MCUver=gf.read_MCUandFPGA(mawin)

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
        #if normal ITLA(Not C++ ITLA) then wavelength minus 8
        #define the data frame to store the test data
        test_result=DataFrame(columns=('SN','TX_DC_DESK','TEMP','DATE','TIME',
                                       '400G_CH','TX_PDL','TX_IL','TX_Vbias_XI',
                                       'TX_Vbias_XQ','TX_Vbias_XP','TX_Vbias_YI','TX_Vbias_YQ','TX_Vbias_YP',
                                       'TX_ER_XI','TX_ER_XQ','TX_ER_XP','TX_ER_YI','TX_ER_YQ','TX_ER_YP',
                                       'TVpi_XI','TVpi_XQ','TVpi_XP','TVpi_YI','TVpi_YQ','TVpi_YP','TX_maxVbias_XI','TX_maxVbias_XQ',
                                       'TX_maxVbias_XP','TX_maxVbias_YI','TX_maxVbias_YQ','TX_maxVbias_YP',
                                       'TX_minVbias_XI','TX_minVbias_XQ',
                                       'TX_minVbias_XP','TX_minVbias_YI','TX_minVbias_YQ','TX_minVbias_YP',
                                       'Res_TxMPDX','Res_TxMPDY'
                                       ))

        timestamp = gf.get_timestamp(1)
        report_path3,network_path=mawin.create_report_folders(sn,timestamp)
        if network_path==False:
            pass
        config = [sn, mawin.desk, mawin.temp[0]] + timestamp.split('_')
        report_name = sn + '_DC_' + timestamp + '.csv'
        report_judgename = sn + '_DC_Report_' + timestamp + '.xlsx'
        report_file = os.path.join(report_path3, report_name)
        report_judge = os.path.join(report_path3, report_judgename)
        print(report_file)
        for i in range(len(mawin.channel)):
            self.sig_print.emit("CH%s 测试开始...\n"%(str(mawin.channel[i])))
            #different loop according to ITLA type
            data = np.zeros(6)
            pwr_cal=0.0
            if mawin.ITLA_type=='C80':
                pwr_cal = float(mawin.ITLA_pwr.iloc[int(mawin.channel[i]) - 1, 1])  # 修改20220722 -9改成-1
                data = cb.get_ER(mawin, str(int(mawin.channel[i])), noipwr[0], noipwr[1], 1, 1)  # 修改20220722 去掉-8
            elif mawin.ITLA_type=='C64':
                pwr_cal = float(mawin.ITLA_pwr.iloc[int(mawin.channel[i]) - 9, 1])
                data = cb.get_ER(mawin, str(int(mawin.channel[i]-8)), noipwr[0], noipwr[1], 1, 1)
            if data==False or data==[]:
                self.sig_print.emit('ER获取失败，任务中止...')
                break
            self.sig_print.emit('ER获取成功...')
            print('max,min,abc,ER,Tvpi:\n',data)
            max_1=data[0][:]
            min_1=data[1][:]
            #get power meter reading of X-max Y-max power
            abc_ok=max_1[:]
            abc_tmp=[0]*6#np.zeros(6)
            while not operator.eq(abc_ok,abc_tmp):
                cb.set_abc(mawin,abc_ok)
                abc_tmp=cb.get_abc(mawin)
            pwr=pwm.read_PM(mawin,mawin.PM_ch)
            if pwr==0:
                self.sig_print.emit('功率获取失败,请检查串口连接...')
                return
            #X/Y max case read the vcode and calculate responsivity of Tx MPD X/Y
            vcodeX,vcodeY=cb.Tx_MPD_responsivity_Test(mawin,'XmaxYmax')
            res_TxMPDX=1000*(vcodeX-int(noipwr[0]))*3.3/(mawin.R_txmpdx*65535*10**(pwr_cal/10))
            res_TxMPDY = 1000*(vcodeY - int(noipwr[1])) * 3.3 / (mawin.R_txmpdy * 65535 * 10**(pwr_cal/10))
            print('vcodeX:', vcodeX)
            print('vcodeY:', vcodeY)
            print('Res_TX_MPDX(mA/mW):',res_TxMPDX)
            print('Res_TX_MPDY(mA/mW):',res_TxMPDY)
            #get power meter reading of X-max Y-min power
            abc_ok=max_1[0:3]+min_1[3:6]
            abc_tmp=[0]*6#np.zeros(6)
            while not operator.eq(abc_ok,abc_tmp):
                cb.set_abc(mawin,abc_ok)
                abc_tmp=cb.get_abc(mawin)
            pwr_x=pwm.read_PM(mawin,mawin.PM_ch)
            if pwr_x==0:
                self.sig_print.emit('功率获取失败,请检查串口连接...')
                return
            #get power meter reading of X-min Y-max power
            abc_ok=min_1[0:3]+max_1[3:6]
            abc_tmp=[0]*6#np.zeros(6)
            while not operator.eq(abc_ok,abc_tmp):
                cb.set_abc(mawin,abc_ok)
                abc_tmp=cb.get_abc(mawin)
            pwr_y=pwm.read_PM(mawin,mawin.PM_ch)
            if pwr_y==0:
                self.sig_print.emit('功率获取失败,请检查串口连接...')
                return
            #set abc to max
            abc_ok=max_1[:]
            abc_tmp=[0]*6#np.zeros(6)
            while not operator.eq(abc_ok,abc_tmp):
                cb.set_abc(mawin,abc_ok)
                abc_tmp=cb.get_abc(mawin)
            self.sig_print.emit('PDL和IL获取成功...')
            CH=str(mawin.channel[i])
            PDL=pwr_x-pwr_y
            IL=pwr-pwr_cal#float(mawin.ITLA_pwr.iloc[int(mawin.channel[i])-1,1])  #修改20220722 -9改成-1
            ABC=data[2]
            ER=data[3]
            Tvpi=data[4]
            Max=max_1[:]
            #05252022:add min value to the rusult to record
            Min=min_1[:]
            tt=config+[CH]+[PDL]+[IL]+ABC+ER+Tvpi+Max+Min+[res_TxMPDX]+[res_TxMPDY]
            test_result.loc[i]=tt
            test_result.to_csv(report_file, index=False)
            #result_tmp.append(tt)
            print('CH:',CH)
            print('PDL:',PDL)
            print('IL:',IL)
            print('ABC:',ABC)
            print('ER:',ER)
            print('Tvpi:',Tvpi)
            print('Max:',Max)
            print('Min:',Min)
            print('Res_Tx_MPD_X and Y:', res_TxMPDX,res_TxMPDY)
            self.sig_progress.emit(round(20+(75/len(mawin.channel)*(i+1))))

        #Judge if retest ER
        retest_ch=[]
        test_result,retest_ch=self.er_retest(test_result,config,noipwr,20,20,-3)
        #generate the report and print out the log file
        if len(retest_ch)>0:
            report_name=sn+'_Retest_'+timestamp+'.csv'
            report_file=os.path.join(report_path3,report_name)
            test_result.to_csv(report_file,index=False)

        #Write data to EEPROM
        # 07172022:add full channel condition to write eeprom
        if test_result.shape[0]>63:
            print('>63通道测试，执行EEPROM写入...')
            gf.write_eeprom_alu(mawin,test_result,mawin.device_type)
        else:
            print('非>63通道测试，不执行EEPROM写入...')
        self.sig_print.emit('测试完成!')
        self.sig_staColor.emit('green')
        self.sig_but.emit('开始')
        self.sig_status.emit('测试完成!')
        self.sig_progress.emit(100)
        print('测试完成')
        ##Write the data into report model and open the report after finished
        wb=xw.Book(mawin.report_model.replace('test_report.xlsx','test_report_DC.xlsx'))
        worksht=wb.sheets(1)
        worksht.activate()
        worksht.range((1,2)).value=test_result.iloc[0,0]
        worksht.range((2,2)).value=test_result.iloc[0,1]
        worksht.range((3,2)).value=test_result.iloc[0,2]
        worksht.range((4,2)).value=test_result.iloc[0,3]
        worksht.range((5,2)).value=test_result.iloc[0,4]
        #write CLPD,FPGA,MCU information here into the test report
        #testBrdClp,ctlBrdClp,ctlBrdModVer,ctlBrdFPGAver,MCUver
        worksht.range((2, 4)).value = testBrdClp
        worksht.range((3, 4)).value = ctlBrdClp
        worksht.range((4, 4)).value = ctlBrdModVer
        worksht.range((5, 4)).value = ctlBrdFPGAver
        worksht.range((6, 4)).value = MCUver
        worksht.range((1, 4)).value = mawin.sw

        worksht.range((7,2)).options(index=False,header=False,transpose=True).value=test_result.iloc[:,5:]#38]
        mawin.finalResult=worksht.range((6,2)).value
        if len(retest_ch)>0:
            for i in retest_ch:
                worksht.range((7,i+2)).color=(255,255,0) #将复测通道颜色标黄
        wb.sheets(2).activate()
        wb.save(report_judge)
        wb.close()
        #copy the local data to network folder
        if not network_path==False:
            shutil.copytree(report_path3,network_path)
        wb = xw.Book(mawin.report_model.replace('test_report.xlsx', 'test_report_DC.xlsx'))
        worksht = wb.sheets(1)
        worksht.activate()

    def DC_test(self):
        '''
        DC test main process
        Update the get ER algorithm in the MCU
        :return:NA
        '''
        #Get and judge the SN format
        sn=str(mawin.lineEdit.text()).strip()
        if not gf.SN_check(sn):
            #mawin.test_status.setText('SN输入有误，请检查SN！')
            self.sig_status.emit('SN输入有误，请检查SN！')
            self.sig_staColor.emit('red')
            self.sig_but.emit('开始')
            return

        #report and config prepare, generate log
        timestamp = gf.get_timestamp(1)
        report_path3, network_path = mawin.create_report_folders(sn, timestamp)
        if network_path == False:
            pass
            # return
        shutil.copy(mawin.config_file, report_path3)
        sys.stdout = gf.Logger(path=report_path3)
        #print(sn)

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
        #New method to check the connectivity and check TIA/DRV is ALU or IDT
        cb.board_set(mawin,mawin.board_up)
        con=0
        con=cb.test_connectivity_new_DC(mawin,sn)
        if con==1:
            self.sig_print.emit('Driver is ALU')
            mawin.device_type='ALU'
            print('检测到ALU器件,将执行ALU器件相关的测试...')
        elif con==2:
            self.sig_print.emit('Driver is IDT')
            mawin.device_type='IDT'
            print('检测到IDT器件,将执行IDT器件相关的测试...')
        elif con==3:
            self.sig_print.emit('请检查光路或压接，连接性检查失败！！！')
            self.sig_staColor.emit('blue')
            self.sig_but.emit('开始')
            self.sig_status.emit('请检查器件连接!')
            #self.sig_stoptest.emit()
            self.sig_progress.emit(0)
            return
        if mawin.device_type=='ALU':
            mawin.drv_up = os.path.join(mawin.config_path, 'Setup_driverup_Ctrlboard56017837A002_20211203_ALU.txt')
            mawin.drv_down = os.path.join(mawin.config_path, 'Setup_driverdown_CtrlboardA001_20210820_ALU.txt')
        if mawin.device_type=='IDT':
            mawin.drv_up = os.path.join(mawin.config_path, 'Setup_driverup_Ctrlboard56017837A002_20211203.txt')
            mawin.drv_down = os.path.join(mawin.config_path, 'Setup_driverdown_CtrlboardA001_20210820.txt')

        '''
        *****Start the whole test steps*****
        '''
        #Read the CLPD,FPGA,MCU information
        testBrdClp='NA'
        ctlBrdClp='NA'
        ctlBrdModVer='NA'
        ctlBrdFPGAver='NA'
        MCUver='NA'
        testBrdClp,ctlBrdClp,ctlBrdModVer,ctlBrdFPGAver,MCUver=gf.read_MCUandFPGA(mawin)

        self.sig_progress.emit(15)
        self.sig_print.emit('器件连接成功, FS400初始化中...')
        #Driver power up
        cb.board_set(mawin,mawin.drv_up)

        #get noise power
        self.sig_print.emit('获取底噪(Get noise power)...')
        noipwr=cb.get_noipwr(mawin)
        self.sig_print.emit('获取底噪完成...')
        print('Noise power:',noipwr)

        self.sig_progress.emit(20)
        self.sig_print.emit('开始IL,PDL,bias测试...')
        #if normal ITLA(Not C++ ITLA) then wavelength minus 8
        #define the data frame to store the test data
        test_result=DataFrame(columns=('SN','TX_DC_DESK','TEMP','DATE','TIME',
                                       '400G_CH','TX_PDL','TX_IL','TX_Vbias_XI',
                                       'TX_Vbias_XQ','TX_Vbias_XP','TX_Vbias_YI','TX_Vbias_YQ','TX_Vbias_YP',
                                       'TX_ER_XI','TX_ER_XQ','TX_ER_XP','TX_ER_YI','TX_ER_YQ','TX_ER_YP',
                                       'TVpi_XI','TVpi_XQ','TVpi_XP','TVpi_YI','TVpi_YQ','TVpi_YP','TX_maxVbias_XI','TX_maxVbias_XQ',
                                       'TX_maxVbias_XP','TX_maxVbias_YI','TX_maxVbias_YQ','TX_maxVbias_YP',
                                       'TX_minVbias_XI','TX_minVbias_XQ',
                                       'TX_minVbias_XP','TX_minVbias_YI','TX_minVbias_YQ','TX_minVbias_YP',
                                       'Res_TxMPDX','Res_TxMPDY'
                                       ))

        config = [sn, mawin.desk, mawin.temp[0]] + timestamp.split('_')
        report_name = sn + '_DC_' + timestamp + '.csv'
        report_judgename = sn + '_DC_Report_' + timestamp + '.xlsx'
        report_file = os.path.join(report_path3, report_name)
        report_judge = os.path.join(report_path3, report_judgename)
        print(report_file)
        for i in range(len(mawin.channel)):
            self.sig_print.emit("CH%s 测试开始...\n"%(str(mawin.channel[i])))
            #different loop according to ITLA type
            data = np.zeros(6)
            pwr_cal=0.0
            if mawin.ITLA_type=='C80':
                pwr_cal = float(mawin.ITLA_pwr.iloc[int(mawin.channel[i]) - 1, 1])  # 修改20220722 -9改成-1
                data = cb.get_ER(mawin, str(int(mawin.channel[i])), noipwr[0], noipwr[1], 1, 1)
                # if i==0:
                #     data = cb.get_ER(mawin, str(int(mawin.channel[i])), noipwr[0], noipwr[1], 1, 1)  # 修改20220722 去掉-8
                # else:
                #     data = cb.get_ER_New1(mawin, str(int(mawin.channel[i])), noipwr[0], noipwr[1], 1, 1)
            elif mawin.ITLA_type=='C64':
                pwr_cal = float(mawin.ITLA_pwr.iloc[int(mawin.channel[i]) - 9, 1])
                data = cb.get_ER(mawin, str(int(mawin.channel[i]-8)), noipwr[0], noipwr[1], 1, 1)
                # if i==0:
                #     data = cb.get_ER(mawin, str(int(mawin.channel[i]-8)), noipwr[0], noipwr[1], 1, 1)
                # else:
                #     data = cb.get_ER_New1(mawin, str(int(mawin.channel[i] - 8)), noipwr[0], noipwr[1], 1, 1)
            if data==False or data==[]:
                print('Get ER error, retest with get er 0')
                if mawin.ITLA_type=='C80':
                    pwr_cal = float(mawin.ITLA_pwr.iloc[int(mawin.channel[i]) - 1, 1])  # 修改20220722 -9改成-1
                    data = cb.get_ER(mawin, str(int(mawin.channel[i])), noipwr[0], noipwr[1], 1, 1)
                    # if i==0:
                    #     data = cb.get_ER(mawin, str(int(mawin.channel[i])), noipwr[0], noipwr[1], 1, 1)  # 修改20220722 去掉-8
                    # else:
                    #     data = cb.get_ER_New1(mawin, str(int(mawin.channel[i])), noipwr[0], noipwr[1], 1, 1)
                elif mawin.ITLA_type=='C64':
                    pwr_cal = float(mawin.ITLA_pwr.iloc[int(mawin.channel[i]) - 9, 1])
                    data = cb.get_ER(mawin, str(int(mawin.channel[i]-8)), noipwr[0], noipwr[1], 1, 1)
                    # if i==0:
                    #     data = cb.get_ER(mawin, str(int(mawin.channel[i]-8)), noipwr[0], noipwr[1], 1, 1)
                    # else:
                    #     data = cb.get_ER_New1(mawin, str(int(mawin.channel[i] - 8)), noipwr[0], noipwr[1], 1, 1)
                if data==False or data==[]:
                    self.sig_print.emit('ER获取失败，任务中止...')
                    break
            self.sig_print.emit('ER获取成功...')
            print('max,min,abc,ER,Tvpi:\n',data)
            max_1=data[0][:]
            min_1=data[1][:]
            #get power meter reading of X-max Y-max power
            abc_ok=max_1[:]
            abc_tmp=[0]*6
            while not operator.eq(abc_ok,abc_tmp):
                cb.set_abc(mawin,abc_ok)
                abc_tmp=cb.get_abc(mawin)
            pwr=pwm.read_PM(mawin,mawin.PM_ch)
            if pwr==0:
                self.sig_print.emit('功率获取失败,请检查串口连接...')
                return
            # 2022/09/05 check IL, if unresonable then break and prompt
            mawin.CtrlB.write(b'itla_rd 0 0x30\n');
            time.sleep(0.3)
            print('ITLA wavelength:\n', mawin.CtrlB.read_all().decode('utf-8'))
            mawin.CtrlB.write(b'itla_rd 0 0x31\n');
            time.sleep(0.3)
            print('ITLA power:\n', mawin.CtrlB.read_all().decode('utf-8'))
            mawin.CtrlB.write(b'itla_rd 0 0x32\n');
            time.sleep(0.3)
            print('ITLA status on/off:\n', mawin.CtrlB.read_all().decode('utf-8'))
            print(mawin.CtrlB.read_all())
            print('Power meter reading:\n', pwr)
            if pwr - pwr_cal < -50:
                print('检测到IL过大，请debug问题，检查是否Tx接上功率计！')
                return
            #X/Y max case read the vcode and calculate responsivity of Tx MPD X/Y
            vcodeX,vcodeY=cb.Tx_MPD_responsivity_Test(mawin,'XmaxYmax')
            res_TxMPDX=1000*(vcodeX-int(noipwr[0]))*3.3/(mawin.R_txmpdx*65535*10**(pwr_cal/10))
            res_TxMPDY = 1000*(vcodeY - int(noipwr[1])) * 3.3 / (mawin.R_txmpdy * 65535 * 10**(pwr_cal/10))
            print('vcodeX:', vcodeX)
            print('vcodeY:', vcodeY)
            print('Res_TX_MPDX(mA/mW):',res_TxMPDX)
            print('Res_TX_MPDY(mA/mW):',res_TxMPDY)
            #get power meter reading of X-max Y-min power
            abc_ok=max_1[0:3]+min_1[3:6]
            abc_tmp=[0]*6#np.zeros(6)
            while not operator.eq(abc_ok,abc_tmp):
                cb.set_abc(mawin,abc_ok)
                abc_tmp=cb.get_abc(mawin)
            pwr_x=pwm.read_PM(mawin,mawin.PM_ch)
            if pwr_x==0:
                self.sig_print.emit('功率获取失败,请检查串口连接...')
                return

            #get power meter reading of X-min Y-max power
            abc_ok=min_1[0:3]+max_1[3:6]
            abc_tmp=[0]*6#np.zeros(6)
            while not operator.eq(abc_ok,abc_tmp):
                cb.set_abc(mawin,abc_ok)
                abc_tmp=cb.get_abc(mawin)
            pwr_y=pwm.read_PM(mawin,mawin.PM_ch)
            if pwr_y==0:
                self.sig_print.emit('功率获取失败,请检查串口连接...')
                return

            if str(mawin.channel[i]) == '13':
                # X - state 1: XI quad and XQ min --- scan XP
                # X - state 2: XI min and XQ quad --- scan XP
                # Y - state 3: YI quad and YQ min --- scan YP
                # Y - state 4: YI min and YQ quad --- scan YP
                # Command: scan_op 0-5 start stop step
                # Firstly X part
                ABC_1 = data[2][:]
                Max_1 = max_1
                Min_1 = min_1
                abc_x1 = [Max_1[0]] + [Min_1[1]] + [ABC_1[2]] + Min_1[3:6]
                abc_x2 = [Min_1[0]] + [Max_1[1]] + [ABC_1[2]] + Min_1[3:6]
                abc_y1 = Min_1[0:3] + [Max_1[3]] + [Min_1[4]] + [ABC_1[5]]
                abc_y2 = Min_1[0:3] + [Min_1[3]] + [Max_1[4]] + [ABC_1[5]]
                st=''
                data_sw1=''
                data_scan=DataFrame
                # X - state 1: XI quad and XQ min --- scan XP
                abc_ok = abc_x1[:]
                abc_tmp = [0] * 6
                while not operator.eq(abc_ok, abc_tmp):
                    cb.set_abc(mawin, abc_ok)
                    abc_tmp = cb.get_abc(mawin)
                (heater, start, stop, step) = (2, 1000, 2500, 10)
                comand = 'scan_op {} {} {} {}\n'.format(heater, start, stop, step).encode('utf-8')
                mawin.CtrlB.write(comand)
                st = mawin.CtrlB.read_until(b'test end').decode('utf-8')
                # left work here to store the data
                data_scan=mawin.dc_scapOPdata(st)
                print('CH13_scanOp state 1: XIquaXQmin_SweepXP:\n',data_scan)
                #data_scan.to_csv(data_sw1)
                data_sw1 = os.path.join(report_path3, 'CH{}_XIquaXQmin_SweepXP.csv'.format(mawin.channel[i]))
                data_scan.to_csv(data_sw1)
                # X - state 2: XI min and XQ quad --- scan XP
                abc_ok = abc_x2[:]
                abc_tmp = [0] * 6  # np.zeros(6)
                while not operator.eq(abc_ok, abc_tmp):
                    cb.set_abc(mawin, abc_ok)
                    abc_tmp = cb.get_abc(mawin)
                (heater, start, stop, step) = (2, 1000, 2500, 10)
                comand = 'scan_op {} {} {} {}\n'.format(heater, start, stop, step).encode('utf-8')
                mawin.CtrlB.write(comand)
                time.sleep(0.1)
                st = mawin.CtrlB.read_until(b'test end').decode('utf-8')
                # left work here to store the data
                data_scan = mawin.dc_scapOPdata(st)
                print('CH13_scanOp state 2: XIminXQqua_SweepXP:\n', data_scan)
                # data_scan.to_csv(data_sw1)
                data_sw1 = os.path.join(report_path3, 'CH{}_XIminXQqua_SweepXP.csv'.format(mawin.channel[i]))
                data_scan.to_csv(data_sw1)
                # Y - state 3: YI quad and YQ min --- scan YP
                abc_ok = abc_y1[:]
                abc_tmp = [0] * 6  # np.zeros(6)
                while not operator.eq(abc_ok, abc_tmp):
                    cb.set_abc(mawin, abc_ok)
                    abc_tmp = cb.get_abc(mawin)
                (heater, start, stop, step) = (5, 1000, 2500, 10)
                comand = 'scan_op {} {} {} {}\n'.format(heater, start, stop, step).encode('utf-8')
                mawin.CtrlB.write(comand)
                time.sleep(0.1)
                st = mawin.CtrlB.read_until(b'test end').decode('utf-8')
                # left work here to store the data
                data_scan = mawin.dc_scapOPdata(st)
                print('CH13_scanOp state 3: YIquaYQmin_SweepYP:\n', data_scan)
                # data_scan.to_csv(data_sw1)
                data_sw1 = os.path.join(report_path3, 'CH{}_YIquaYQmin_SweepYP.csv'.format(mawin.channel[i]))
                data_scan.to_csv(data_sw1)
                # Y - state 4: YI min and YQ quad --- scan YP
                abc_ok = abc_y2[:]
                abc_tmp = [0] * 6  # np.zeros(6)
                while not operator.eq(abc_ok, abc_tmp):
                    cb.set_abc(mawin, abc_ok)
                    abc_tmp = cb.get_abc(mawin)
                (heater, start, stop, step) = (2, 1000, 2500, 10)
                comand = 'scan_op {} {} {} {}\n'.format(heater, start, stop, step).encode('utf-8')
                mawin.CtrlB.write(comand)
                time.sleep(0.1)
                st = mawin.CtrlB.read_until(b'test end').decode('utf-8')
                # left work here to store the data
                data_scan = mawin.dc_scapOPdata(st)
                print('CH13_scanOp state 3: YIminYQqua_SweepYP:\n', data_scan)
                # data_scan.to_csv(data_sw1)
                data_sw1 = os.path.join(report_path3, 'CH{}_YIminYQqua_SweepYP.csv'.format(mawin.channel[i]))
                data_scan.to_csv(data_sw1)
                # The end of the CH13 XP/YP sweep test

            #set abc to max
            abc_ok=max_1[:]
            abc_tmp=[0]*6#np.zeros(6)
            while not operator.eq(abc_ok,abc_tmp):
                cb.set_abc(mawin,abc_ok)
                abc_tmp=cb.get_abc(mawin)
            self.sig_print.emit('PDL和IL获取成功...')
            CH=str(mawin.channel[i])
            PDL=pwr_x-pwr_y
            IL=pwr-pwr_cal #float(mawin.ITLA_pwr.iloc[int(mawin.channel[i])-1,1])  #修改20220722 -9改成-1
            ABC=data[2]
            ER=data[3]
            Tvpi=data[4]
            Max=max_1[:]
            #05252022:add min value to the rusult to record
            Min=min_1[:]
            tt=config+[CH]+[PDL]+[IL]+ABC+ER+Tvpi+Max+Min+[res_TxMPDX]+[res_TxMPDY]
            test_result.loc[i]=tt
            test_result.to_csv(report_file, index=False)
            #result_tmp.append(tt)
            print('CH:',CH)
            print('PDL:',PDL)
            print('IL:',IL)
            print('ABC:',ABC)
            print('ER:',ER)
            print('Tvpi:',Tvpi)
            print('Max:',Max)
            print('Min:',Min)
            print('Res_Tx_MPD_X and Y:', res_TxMPDX,res_TxMPDY)
            self.sig_progress.emit(round(20+(75/len(mawin.channel)*(i+1))))

        retest_ch=[]
        test_result,retest_ch=self.er_retest(test_result,config,noipwr,10,15,-3)
        #generate the report and print out the log file
        if len(retest_ch)>0:
            report_name=sn+'_Retest_'+timestamp+'.csv'
            report_file=os.path.join(report_path3,report_name)
            test_result.to_csv(report_file,index=False)

        #Copy golden CSV data to folder under golden sample
        if mawin.test_type == '金样':
            golden_path=os.path.join(os.path.split(os.path.split(report_path3)[0])[0],'Monitor_data')
            if not os.path.exists(golden_path):os.mkdir(golden_path)
            shutil.copy(report_file, golden_path)
        #Write data to EEPROM
        # 07172022:add full channel condition to write eeprom
        if mawin.test_type=='Normal':
            print('正常生产测试，将执行EEPROM写入...')
            gf.write_eeprom_final(mawin,sn,test_result, mawin.device_type)
        else:
            print('非正常生产测试，不执行EEPROM写入...')
        self.sig_print.emit('测试完成!')
        self.sig_staColor.emit('green')
        self.sig_but.emit('开始')
        self.sig_status.emit('测试完成!')
        self.sig_progress.emit(100)
        print('测试完成')

        #Write the data into report model and open the report after finished
        wb=xw.Book(mawin.report_model.replace('test_report.xlsx','test_report_DC.xlsx'))
        worksht=wb.sheets(1)
        # app = xw.App(visible=False, add_book=False)
        # wb = app.books.open(self.report_model)
        worksht.activate()
        worksht.range((1,2)).value=test_result.iloc[0,0]
        worksht.range((2,2)).value=test_result.iloc[0,1]
        worksht.range((3,2)).value=test_result.iloc[0,2]
        worksht.range((4,2)).value=test_result.iloc[0,3]
        worksht.range((5,2)).value=test_result.iloc[0,4]

        #write CLPD,FPGA,MCU information here into the test report
        #testBrdClp,ctlBrdClp,ctlBrdModVer,ctlBrdFPGAver,MCUver
        worksht.range((2, 4)).value = testBrdClp
        worksht.range((3, 4)).value = ctlBrdClp
        worksht.range((4, 4)).value = ctlBrdModVer
        worksht.range((5, 4)).value = ctlBrdFPGAver
        worksht.range((6, 4)).value = MCUver
        worksht.range((1, 4)).value = mawin.sw

        #paste result to the report
        worksht.range((7,2)).options(index=False,header=False,transpose=True).value=test_result.iloc[:,5:]#38]
        mawin.finalResult=worksht.range((6,2)).value
        if len(retest_ch)>0:
            for i in retest_ch:
                worksht.range((7,i+2)).color=(255,255,0) #将复测通道颜色标黄
        wb.sheets(2).activate()
        wb.save(report_judge)
        wb.close()
        #copy the local data to network folder
        if not network_path==False:
            shutil.copytree(report_path3,network_path)
        wb = xw.Book(report_judge)
        wb.sheets(2).activate()

    def er_retest(self,test_result,config,noipwr,er_judgeIQ=10,er_judgeXY=15,di_judge=-3):
        '''
        Updated on 05242022:add the function to check if :
        IQ_ER<10 or XY_ER<10
        and at the same time
        diff<-3dB
        Then perform the corresponding channel retest
        :parameter:dataframe of DC test result,config,noipwr,er_judgeIQ=15,er_judgeXY=10,di_judge=-3 in default
        :return: updated dataframe of DC test result,the retested channels,retest_flag=the times needs to do retest
        '''
        ch_retest,er_retest=erRetest.er_judge(test_result,er_judgeIQ,er_judgeXY,di_judge)
        retest_flag=3
        count=0
        ch_judge=ch_retest
        while count<retest_flag:
            if len(ch_judge)>0:
                print(ch_judge)
                print('复测通道：',ch_judge)
                self.sig_print.emit('检测到ER跳变过大，将执行第{}次对应通道复测：{}'.format(str(count+1),ch_judge))
                #start the retest
                res_retest=[]
                for i in ch_judge:
                    self.sig_print.emit("CH%s 复测开始...\n"%(str(mawin.channel[i])))
                    # different loop according to ITLA type
                    data = np.zeros(6)
                    pwr_cal = 0.0
                    if mawin.ITLA_type == 'C80':
                        pwr_cal = float(mawin.ITLA_pwr.iloc[int(mawin.channel[i]) - 1, 1])  # 修改20220722 -9改成-1
                        data = cb.get_ER(mawin, str(int(mawin.channel[i])), noipwr[0], noipwr[1], 1,
                                         1)  # 修改20220722 去掉-8
                    elif mawin.ITLA_type == 'C64':
                        pwr_cal = float(mawin.ITLA_pwr.iloc[int(mawin.channel[i]) - 9, 1])
                        data = cb.get_ER(mawin, str(int(mawin.channel[i] - 8)), noipwr[0], noipwr[1], 1, 1)
                    # pwr_cal = float(mawin.ITLA_pwr.iloc[int(mawin.channel[i]) - 1, 1])  # 修改20220722 -9改成-1
                    # data=np.zeros(6)
                    # data=cb.get_ER(mawin,str(int(mawin.channel[i])),noipwr[0],noipwr[1],1,1) # 去除 -8
                    if data==False or data==[]:
                        self.sig_print.emit('ER获取失败，任务中止...')
                        break
                    self.sig_print.emit('ER获取成功...')
                    print('max,min,abc,ER,Tvpi:\n',data)
                    max_1=data[0][:]
                    min_1=data[1][:]
                    #get power meter reading of X-max Y-max power
                    abc_ok=max_1[:]
                    abc_tmp=[0]*6#np.zeros(6)
                    while not operator.eq(abc_ok,abc_tmp):
                        cb.set_abc(mawin,abc_ok)
                        abc_tmp=cb.get_abc(mawin)
                    pwr=pwm.read_PM(mawin,mawin.PM_ch)
                    if pwr==0:
                        self.sig_print.emit('功率获取失败,请检查串口连接...')
                        return
                    # X/Y max case read the vcode and calculate responsivity of Tx MPD X/Y
                    vcodeX, vcodeY = cb.Tx_MPD_responsivity_Test(mawin, 'XmaxYmax')
                    res_TxMPDX = 1000 * (vcodeX - int(noipwr[0])) * 3.3 / (
                                mawin.R_txmpdx * 65535 * 10 ** (pwr_cal / 10))
                    res_TxMPDY = 1000 * (vcodeY - int(noipwr[1])) * 3.3 / (
                                mawin.R_txmpdy * 65535 * 10 ** (pwr_cal / 10))
                    print('vcodeX:', vcodeX)
                    print('vcodeY:', vcodeY)
                    print('Res_TX_MPDX(mA/mW):', res_TxMPDX)
                    print('Res_TX_MPDY(mA/mW):', res_TxMPDY)
                    #get power meter reading of X-max Y-min power
                    abc_ok=max_1[0:3]+min_1[3:6]
                    abc_tmp=[0]*6#np.zeros(6)
                    while not operator.eq(abc_ok,abc_tmp):
                        cb.set_abc(mawin,abc_ok)
                        abc_tmp=cb.get_abc(mawin)
                    pwr_x=pwm.read_PM(mawin,mawin.PM_ch)
                    if pwr_x==0:
                        self.sig_print.emit('功率获取失败,请检查串口连接...')
                        return
                    #get power meter reading of X-min Y-max power
                    abc_ok=min_1[0:3]+max_1[3:6]
                    abc_tmp=[0]*6#np.zeros(6)
                    while not operator.eq(abc_ok,abc_tmp):
                        cb.set_abc(mawin,abc_ok)
                        abc_tmp=cb.get_abc(mawin)
                    pwr_y=pwm.read_PM(mawin,mawin.PM_ch)
                    if pwr_y==0:
                        self.sig_print.emit('功率获取失败,请检查串口连接...')
                        return
                    #set abc to max
                    abc_ok=max_1[:]
                    abc_tmp=[0]*6#np.zeros(6)
                    while not operator.eq(abc_ok,abc_tmp):
                        cb.set_abc(mawin,abc_ok)
                        abc_tmp=cb.get_abc(mawin)
                    self.sig_print.emit('PDL和IL获取成功...')
                    CH=str(mawin.channel[i])
                    PDL=pwr_x-pwr_y
                    IL=pwr-pwr_cal#float(mawin.ITLA_pwr.iloc[int(mawin.channel[i])-1,1]) #-9 --》-1 20220728
                    ABC=data[2]
                    ER=data[3]
                    Tvpi=data[4]
                    Max=max_1[:]
                    #05252022:add min value to the rusult to record
                    Min=min_1[:]
                    res_retest=config+[CH]+[PDL]+[IL]+ABC+ER+Tvpi+Max+Min+[res_TxMPDX]+[res_TxMPDY]
                    test_result.iloc[i,:]=res_retest
                    print('CH:',CH)
                    print('PDL:',PDL)
                    print('IL:',IL)
                    print('ABC:',ABC)
                    print('ER:',ER)
                    print('Tvpi:',Tvpi)
                    print('Max:',Max)
                    print('Min:',Min)
                count+=1
                ch_judge=erRetest.er_judge(test_result,er_judgeIQ,er_judgeXY,di_judge)[0]
            else:
                self.sig_print.emit('复测判定OK或无需复测')
                break
        return (test_result,ch_retest)


    def test_unit(self):
        '''
        #test function to verify the timer for goca timeout test
        :return:NA
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

        # Add the folder named as SN_Date to store all the test data
        timestamp = gf.get_timestamp(1)
        report_path3, network_path = mawin.create_report_folders(sn, timestamp)
        if network_path == False:
            pass
        shutil.copy(mawin.config_file, report_path3)
        sys.stdout = gf.Logger(path=report_path3)

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
        # New method to check the connectivity and check TIA/DRV is ALU or IDT
        con = 0
        con = cb.test_connectivity_new(mawin, sn)
        if con == 1:
            self.sig_print.emit('Driver is ALU')
            mawin.device_type = 'ALU'
            print('检测到ALU器件,将执行ALU器件相关的测试...')
        elif con == 2:
            self.sig_print.emit('Driver is IDT')
            mawin.device_type = 'IDT'
            print('检测到IDT器件,将执行IDT器件相关的测试...')
        elif con == 3:
            self.sig_print.emit('请检查光路或压接，连接性检查失败！！！')
            self.sig_staColor.emit('blue')
            self.sig_but.emit('开始')
            self.sig_status.emit('请检查器件连接!')
            # self.sig_stoptest.emit()
            self.sig_progress.emit(0)
            return
        if mawin.device_type=='ALU':
            #mawin.board_up = os.path.join(self.config_path, 'Setup_brdup_CtrlboardA001_20220113_ALU.txt')
            mawin.drv_up = os.path.join(mawin.config_path, 'Setup_driverup_Ctrlboard56017837A002_20211203_ALU.txt')
            mawin.drv_down = os.path.join(mawin.config_path, 'Setup_driverdown_CtrlboardA001_20210820_ALU.txt')
        if mawin.device_type=='IDT':
            #mawin.board_up = os.path.join(self.config_path, 'Setup_brdup_CtrlboardA001_20220113_ALU.txt')
            mawin.drv_up = os.path.join(mawin.config_path, 'Setup_driverup_Ctrlboard56017837A002_20211203.txt')
            mawin.drv_down = os.path.join(mawin.config_path, 'Setup_driverdown_CtrlboardA001_20210820.txt')

        # Read the CLPD,FPGA,MCU information
        testBrdClp = 'NA'
        ctlBrdClp = 'NA'
        ctlBrdModVer = 'NA'
        ctlBrdFPGAver = 'NA'
        MCUver = 'NA'
        testBrdClp, ctlBrdClp, ctlBrdModVer, ctlBrdFPGAver, MCUver = gf.read_MCUandFPGA(mawin)
        '''
        *****Start the whole test steps*****
        '''
        self.sig_progress.emit(15)
        self.sig_print.emit('器件连接成功, FS400初始化中...')
        #Driver power up
        cb.board_set(mawin,mawin.drv_up)
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
        self.sig_print.emit('开始TxBW测试...')
        #if normal ITLA(Not C++ ITLA) then wavelength minus 8
        #define the data frame to store the test data
        test_result=DataFrame(columns=('SN','TX_BW_DESK','TEMP','DATE','TIME','400G_CH',
                                       'TX_ROLLOFF_XI','TX_BW3DB_XI','TX_BW6DB_XI','Kink_XI','sigma_XI',
                                       'TX_ROLLOFF_XQ','TX_BW3DB_XQ','TX_BW6DB_XQ','Kink_XQ','sigma_XQ',
                                       'TX_ROLLOFF_YI','TX_BW3DB_YI','TX_BW6DB_YI','Kink_YI','sigma_YI',
                                       'TX_ROLLOFF_YQ','TX_BW3DB_YQ','TX_BW6DB_YQ','Kink_YQ','sigma_YQ'))
        config = [sn, mawin.desk, mawin.temp[0]] + timestamp.split('_')
        report_name = sn + '_TxBW_' + timestamp + '.csv'
        report_judgename = sn + '_TxBW_Report_' + timestamp + '.xlsx'
        report_file = os.path.join(report_path3, report_name)
        report_judge = os.path.join(report_path3, report_judgename)

        linecal=pd.read_csv(mawin.Tx_line_loss).iloc[:-1,:]
        pdcal=pd.read_csv(mawin.Tx_pd_loss).iloc[:-1,:]
        fre_ref=mawin.cal_fre#1.5e9
        for i in range(len(mawin.channel)):
            self.sig_print.emit("CH%s 测试开始...\n"%(str(mawin.channel[i])))
            data = np.zeros(6)
            if mawin.ITLA_type == 'C80':
                data=cb.get_ER(mawin,str(int(mawin.channel[i])),10,10,0,0) #20220728 -8去掉
            elif mawin.ITLA_type == 'C64':
                data = cb.get_ER(mawin, str(int(mawin.channel[i] - 8)), 10, 10, 0, 0)
            if data==False or data==[]:
                self.sig_print.emit('ABC point获取失败，任务中止...')
                break
            self.sig_print.emit('ABC获取成功...')
            print('max,min,abc,ER,Tvpi:\n',data)
            abc=data[2][:]
            #get power meter reading of X-max Y-max power
            abc_ok=abc[:]
            abc_tmp=[0]*6
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
                abc_tmp=[0]
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
                        df.to_csv(os.path.join(report_path3,filename[j]),index=None)
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
                        abc_ok=abc[:]
                        abc_tmp=[0]*6
                        while not operator.eq(abc_ok,abc_tmp):
                            cb.set_abc(mawin,abc_ok)
                            abc_tmp=cb.get_abc(mawin)
                        tt.append([rolloff,BW3dB,BW6dB,kink,sigma_ss])
            tt1=config+[str(mawin.channel[i])]+tt[0]+tt[1]+tt[2]+tt[3]
            test_result.loc[i]=tt1
            test_result.to_csv(report_file, index=False)
            self.sig_progress.emit(round(20+(75/len(mawin.channel)*(i+1))))

            #add here to set work point to max if the last channel
            if i==(len(mawin.channel)-1):
                print('最后一个通道，设置工作点到max')
                abc_ok = data[0][:]
                abc_tmp = [0] * 6  # np.zeros(6)
                while not operator.eq(abc_ok, abc_tmp):
                    cb.set_abc(mawin, abc_ok)
                    abc_tmp = cb.get_abc(mawin)
        #Set the RF switch channel to empty namely CH0
        rfsw.switch_RFchannel(mawin, 0)
        self.sig_print.emit('测试完成!')
        self.sig_staColor.emit('green')
        self.sig_but.emit('开始')
        self.sig_status.emit('测试完成!')
        self.sig_progress.emit(100)
        print('测试完成')
        # Copy golden CSV data to folder under golden sample
        if mawin.test_type == '金样':
            golden_path = os.path.join(os.path.split(os.path.split(report_path3)[0])[0], 'Monitor_data')
            if not os.path.exists(golden_path): os.mkdir(golden_path)
            shutil.copy(report_file, golden_path)
        ##Write the data into report model and open the report after finished
        wb=xw.Book(mawin.report_model.replace('test_report.xlsx','test_report_TxBW.xlsx'))
        worksht=wb.sheets(1)
        worksht.activate()
        worksht.range((1,2)).value=test_result.iloc[0,0]
        worksht.range((2,2)).value=test_result.iloc[0,1]
        worksht.range((3,2)).value=test_result.iloc[0,2]
        worksht.range((4,2)).value=test_result.iloc[0,3]
        worksht.range((5,2)).value=test_result.iloc[0,4]
        # write CLPD,FPGA,MCU information here into the test report
        # testBrdClp,ctlBrdClp,ctlBrdModVer,ctlBrdFPGAver,MCUver
        worksht.range((2, 4)).value = testBrdClp
        worksht.range((3, 4)).value = ctlBrdClp
        worksht.range((4, 4)).value = ctlBrdModVer
        worksht.range((5, 4)).value = ctlBrdFPGAver
        worksht.range((6, 4)).value = MCUver
        worksht.range((1, 4)).value = mawin.sw

        worksht.range((7,2)).options(index=False,header=False,transpose=True).value=test_result.iloc[:,5:]#26]
        mawin.finalResult=worksht.range((6,2)).value
        wb.save(report_judge)
        wb.close()
        # copy the local data to network folder
        if not network_path == False:
            shutil.copytree(report_path3, network_path)
        wb = xw.Book(report_judge)
        wb.sheets(1).activate()

    def ICR_test_main(self):
        '''
        #ICR test main process
        :return:NA
        '''
        # firstly rename the file of ICR test config of 'Board up,Drv up,Drv down'
        # Because no need to open ITLA, and need to power down PDs
        mawin.board_up = os.path.join(mawin.config_path, 'Setup_brdup_CtrlboardA001_20220113_ICR.txt')
        # Get and judge the SN format
        sn = str(mawin.lineEdit.text()).strip()
        if not gf.SN_check(sn):
            self.sig_status.emit('SN输入有误，请检查SN！')
            self.sig_staColor.emit('red')
            self.sig_but.emit('开始')
            return

        # Test data storage
        # Add the folder named as SN_Date to store all the test data
        timestamp = gf.get_timestamp(1)
        report_path3, network_path = mawin.create_report_folders(sn, timestamp)
        if network_path == False:
            pass
        shutil.copy(mawin.config_file, report_path3)
        sys.stdout = gf.Logger(path=report_path3)

        self.sig_status.emit('测试进行中...')
        self.sig_staColor.emit('yellow')
        self.sig_clear.emit()
        self.sig_print.emit(sn)
        self.sig_progress.emit(5)

        #设备连接
        if not cb.open_board(mawin, mawin.ctrl_port):
            self.sig_status.emit('请检查控制板串口！')
            self.sig_but.emit('开始')
            self.sig_staColor.emit('red')
            self.sig_progress.emit(0)
            return
        #Oscilloscope and ICR tester connection
        if not osc.open_OScope(mawin, mawin.OScope_port):
            self.sig_status.emit('请检查示波器端口！')
            self.sig_but.emit('开始')
            self.sig_staColor.emit('red')
            self.sig_progress.emit(0)
            return
        if not icr.open_ICRtf(mawin, mawin.ICRtf_port):
            self.sig_status.emit('请检查ICR测试平台串口！')
            self.sig_but.emit('开始')
            self.sig_staColor.emit('red')
            self.sig_progress.emit(0)
            return

        #connectivity test
        self.sig_progress.emit(10)
        self.sig_print.emit('检查pin脚连接中...')
        #New method to check the connectivity and check TIA/DRV is ALU or IDT
        con = 0
        con = cb.test_connectivity_new_ICR(mawin, sn)
        if con == 1:
            self.sig_print.emit('Driver is ALU')
            mawin.device_type = 'ALU'
            print('检测到ALU器件,将执行ALU器件相关的测试...')
        elif con == 2:
            self.sig_print.emit('Driver is IDT')
            mawin.device_type = 'IDT'
            print('检测到IDT器件,将执行IDT器件相关的测试...')
        elif con == 3:
            self.sig_print.emit('请检查光路或压接，连接性检查失败！！！')
            self.sig_staColor.emit('blue')
            self.sig_but.emit('开始')
            self.sig_status.emit('请检查器件连接!')
            # self.sig_stoptest.emit()
            self.sig_progress.emit(0)
            return

        # Read the CLPD,FPGA,MCU information
        testBrdClp = 'NA'
        ctlBrdClp = 'NA'
        ctlBrdModVer = 'NA'
        ctlBrdFPGAver = 'NA'
        MCUver = 'NA'
        # testBrdClp, ctlBrdClp, ctlBrdModVer, ctlBrdFPGAver, MCUver = gf.read_MCUandFPGA(mawin)

        '''
        *****Start the whole test steps*****
        '''
        self.sig_progress.emit(15)
        self.sig_print.emit('器件连接成功, FS400初始化, 配置TIA及上电...')
        # TIA Power on and config
        if mawin.device_type=='IDT':
            mawin.drv_up = os.path.join(mawin.config_path, 'Setup_driverup_Ctrlboard56017837A002_20211203.txt')
            mawin.drv_down = os.path.join(mawin.config_path, 'Setup_driverdown_CtrlboardA001_20210820.txt')
            gf.TIA_on(mawin)
            gf.TIA_config(mawin)
            print('IDT器件测试')
            self.sig_print.emit('TIA上电完成，设备初始化中...')
        elif mawin.device_type == 'ALU':
            mawin.drv_up = os.path.join(mawin.config_path, 'Setup_driverup_Ctrlboard56017837A002_20211203_ALU.txt')
            mawin.drv_down = os.path.join(mawin.config_path, 'Setup_driverdown_CtrlboardA001_20210820_ALU.txt')
            gf.TIA_on(mawin)
            gf.TIA_config_ALU(mawin,mawin.test_res)
            print('ALU器件测试')

        # initiate Oscilloscope and switch to visa connection
        osc.switch_to_DSO(mawin, mawin.OScope_port)
        print('DSO object created')
        osc.switch_to_VISA(mawin, mawin.DSO_port)
        if mawin.test_pe or mawin.test_bw:
            osc.init_OSc(mawin)
            self.sig_print.emit('示波器初始化成功...')
        self.sig_print.emit('开启并设置Laser波长...')
        icr.set_laser(mawin, 1, 15.4, 10)
        icr.set_laser(mawin, 2, 15.4, -5, 1550.000)
        self.sig_progress.emit(20)
        self.sig_print.emit('配置完成!开始ICR测试...')
        #if normal ITLA(Not C++ ITLA) then wavelength minus 8
        #define the data frame to store the test data
        test_result = DataFrame(columns=('SN', 'RX_BW_DESK', 'TEMP', 'DATE', 'TIME', '400G_CH',
                                         'RX_BW_XI', 'RX_BW_XQ', 'RX_BW_YI', 'RX_BW_YQ', 'RX_PE_X',
                                         'RX_PE_Y', 'RX_X_Skew', 'RX_Y_Skew', 'RX_1PD_RES_LO_X', 'RX_1PD_RES_LO_Y',
                                         'RX_1PD_RES_SIG_X', 'RX_1PD_RES_SIG_Y', 'RX_PER_X', 'RX_PER_Y', 'I_Dark_X',
                                         'I_Dark_Y'))
        # create list to store the result
        tt = []
        config = [sn, mawin.desk, mawin.temp[0]] + timestamp.split('_')
        # report_name=sn+'_'+timestamp+'.csv'
        report_name = sn + '_ICR_' + timestamp + '.csv'  # BOB改了csv报告名字
        report_judgename = sn + '_ICR_Report_' + timestamp + '.xlsx'
        report_file = os.path.join(report_path3, report_name)
        report_judge = os.path.join(report_path3, report_judgename)

        for i in range(len(mawin.channel)):
            ch = str(mawin.channel[i])
            p = i + 1
            self.sig_print.emit("CH%s 测试开始...\n" % (ch))
            # ICR Test start and set the wavelength
            icr.set_wavelength(mawin, 2, ch, 0)
            icr.set_wavelength(mawin, 1, ch, 1)
            mawin.ICRtf.write('OUTP3:CHAN1:POW 6')  # modify the value from 10dBm to 9dBm
            mawin.ICRtf.write('OUTP3:CHAN2:POW -6')
            mawin.ICRtf.write('OUTP1:CHAN1:STATE ON')
            mawin.ICRtf.write('OUTP1:CHAN2:STATE ON')
            pe_x, pe_y, skew_x, skew_y = [0.0] * 4
            XIBW3dB, XQBW3dB, YIBW3dB, YQBW3dB = [0.0] * 4
            res_sig_x, res_sig_y, per_x, per_y, dark_x, dark_y, res_lo_x, res_lo_y = [0.0] * 8
            # ***---------PE and SKEW test--------***
            if mawin.test_pe:
                self.sig_print.emit('Phase 和 Skew 测试中...')
                # phase error test
                lamb = icr.cal_wavelength(ch, 0)
                m = 3  # 平均次数
                icr.ICR_scamblePol(mawin)
                # query the data of amtiplitude
                A1 = mawin.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')
                A3 = mawin.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')
                while not icr.isfloat(A1) or not icr.isfloat(A3):
                    A1 = mawin.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')
                    A3 = mawin.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')
                while float(A1) / float(A3) > 1.08 or float(A1) / float(A3) < 0.92:
                    time.sleep(0.2)
                    icr.ICR_scamblePol(mawin)
                    A1 = mawin.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')
                    A3 = mawin.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')
                    while not icr.isfloat(A1) or not icr.isfloat(A3):
                        A1 = mawin.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')
                        A3 = mawin.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')
                # Scan RF start
                data = mawin.scan_RF(ch)
                # save the result
                pe_x, pe_y, skew_x, skew_y, new_pahseX, new_pahseY = peSkew.deal_with_data(data)
                # save raw and fitted s21 data
                raw_phase_report = os.path.join(report_path3, '{}_CH{}_phase_RawData_{}.csv'.format(sn, ch, timestamp))
                col = ['Phase error', 'Skew'] + [str(i) + 'GHz' for i in range(1, 11)]
                pha = pd.DataFrame([[pe_x, skew_x] + new_pahseX, [pe_y, skew_y] + new_pahseY],
                                   index=('phase_X', 'phase_Y'), columns=col)
                pha.to_csv(raw_phase_report)
                # draw the curve in the main thread
                self.sig_plot.emit(pha, report_path3, ch, sn, timestamp, 'PE',[],[])

            #***---------BW test--------***
            if mawin.test_bw and ch=='13':
                # Rx BW test
                self.sig_print.emit('Rx BW 测试中...')
                icr.set_wavelength(mawin, 2, ch, 0)
                icr.set_wavelength(mawin, 1, ch, 1)
                time.sleep(0.1)
                error = True
                eg = 1  # 表示在某一个波长处测试的次数
                while error:
                    icr.ICR_scamblePol(mawin)
                    # query the data of amtiplitude
                    A1 = mawin.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')
                    A3 = mawin.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')
                    while not icr.isfloat(A1) or not icr.isfloat(A3):
                        time.sleep(0.2)
                        A1 = mawin.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')
                        A3 = mawin.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')
                    while float(A1) / float(A3) > 1.15 or float(A1) / float(A3) < 0.85:
                        time.sleep(0.2)
                        icr.ICR_scamblePol(mawin)
                        A1 = mawin.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')
                        A3 = mawin.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')
                        while not icr.isfloat(A1) or not icr.isfloat(A3):
                            A1 = mawin.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')
                            A3 = mawin.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')
                    # scan S21
                    print('A1:{}\nA3:{}'.format(A1, A3))
                    data1 = mawin.scan_S21(ch, 1, 40, 1)
                    # config the frequency
                    numfre = 40
                    da = [[0.0] * 4] * numfre
                    fre = [i * 1e9 for i in range(1, 41)]
                    # 归一化数据Y
                    raw = pd.DataFrame(data1)
                    for i in range(4):
                        raw.iloc[:, i] = 20 * np.log10(raw.iloc[:, i] / raw.iloc[0, i])
                    # 线损
                    linelo = pd.read_csv(mawin.Rx_line_loss, encoding='gb2312')
                    fre_loss = linelo.iloc[:-1, 0]
                    loss = linelo.iloc[:-1, 1]
                    Yi = np.interp(fre, fre_loss, loss)  # 插值求得对应频率的线损
                    # 减去线损
                    Ysmo = raw.sub(Yi, axis=0)
                    # save raw and fitted s21 data
                    raw_s21_report = os.path.join(report_path3, '{}_CH{}_RxBw_RawData_{}.csv'.format(sn, ch, timestamp))
                    fit_s21_report = os.path.join(report_path3, '{}_CH{}_RxBw_FitData_{}.csv'.format(sn, ch, timestamp))
                    colBW = ['XI', 'XQ', 'YI', 'YQ']
                    indBW = [i for i in range(1, 41)]
                    raw.columns = colBW
                    raw.index = indBW
                    raw.to_csv(raw_s21_report)  # ,index=False)
                    ##This point needs to be verified, whether the smooth data is equal to matlab gaussian method
                    for i in range(4):
                        tem = smooth.smooth(Ysmo.iloc[:, i], 3)[1:-1]
                        tem = tem - tem[0]  # 归一化数据
                        Ysmo.iloc[:, i] = tem
                    Ysmo.columns = colBW
                    Ysmo.index = indBW
                    Ysmo.to_csv(fit_s21_report)  # ,index=False)
                    # draw the curves
                    self.sig_plot.emit(Ysmo, report_path3, ch, sn, timestamp, 'BW',[],[])

                    # calculate the result, base on Ysmo is a dataframe type
                    for i in range(4):
                        aa = [x for x in range(Ysmo.shape[0]) if Ysmo.iloc[x, i] < -3]
                        if len(aa) == 0:
                            BW3dB = 40
                        else:
                            BW3dB = min(aa) + 1  # /1e9 index plus 1 as Freq
                        # each channel
                        if i == 0:
                            XIBW3dB = BW3dB
                            if min(Ysmo.iloc[:, i]) < -20.0:
                                XIBW3dB = 0.0
                        elif i == 1:
                            XQBW3dB = BW3dB
                            if min(Ysmo.iloc[:, i]) < -20.0:
                                XQBW3dB = 0.0
                        elif i == 2:
                            YIBW3dB = BW3dB
                            if min(Ysmo.iloc[:, i]) < -20.0:
                                YIBW3dB = 0.0
                        elif i == 3:
                            YQBW3dB = BW3dB
                            if min(Ysmo.iloc[:, i]) < -20.0:
                                YQBW3dB = 0.0

                    YsmoMean = Ysmo.mean(axis=1)
                    aa = [x for x in range(YsmoMean.shape[0]) if YsmoMean[x + 1] < -3]
                    if len(aa) == 0:
                        MeanBW3dB = 40
                    else:
                        MeanBW3dB = min(aa) + 1  # /1e9
                        if min(YsmoMean) < -20:
                            MeanBW3dB = 0

                    # Need to check whether to perform the judgement of >28G Hz and each channel > mean bw by 3GHz
                    # if not meet the requirement then perform the retest
                    # left to write the result
                    error = False  # Not check BW and retest
            # TIA amplitude test not performed
            # scan_Amp:to test the TIA output amplitude
            #***---------Responsivity test--------***---To Be Removed
            if mawin.test_res:
                # PD voltage set to 2150
                mawin.CtrlB.write(b'switch_set 10 0\n');
                time.sleep(0.1)
                mawin.CtrlB.write(b'cpld_spi_wr 0x2c 2150\n');
                time.sleep(0.1)
                mawin.CtrlB.write(b'cpld_spi_wr 0x2f 2150\n');
                time.sleep(0.1)

                self.sig_print.emit('Rx响应度测试中...')
                res_sig_x, res_sig_y, per_x, per_y, dark_x, dark_y, PD_currentX, PD_currentY,id_markYmin,id_markXmin= mawin.get_pd_resp_sig_New(ch)
                #res_sig_x, res_sig_y, per_x, per_y, dark_x, dark_y, PD_currentX, PD_currentY = mawin.get_pd_resp_sig(ch)

                # Draw the Rx PD current data point and save the image
                # Save the raw data
                res_raw = os.path.join(report_path3, '{}_CH{}_RxPDcurrentXY_{}.csv'.format(sn, ch, timestamp))
                raw_res = pd.DataFrame([PD_currentX, PD_currentY]).transpose()
                raw_res.columns = ['PD_currentX', 'PD_currentY']
                raw_res.to_csv(res_raw)
                # Draw the curve X Y phase
                self.sig_plot.emit(raw_res, report_path3, ch, sn, timestamp, 'RES',id_markYmin,id_markXmin)
                self.sig_print.emit('LO响应度测试中...')
                # res_lo_x,res_lo_y,curX,curY=mawin.get_pd_resp_lo(ch)

                # Bob 瞎搞 !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

                PD_current_X_LO = []
                PD_current_Y_LO = []
                for t in range(10):
                    res_lo_x, res_lo_y, curX, curY = mawin.get_pd_resp_lo(ch)
                    PD_current_X_LO.append(curX)
                    PD_current_Y_LO.append(curY)

                res_raw = os.path.join(report_path3, '{}_CH{}_RxPDcurrentXY_LO_{}.csv'.format(sn, ch, timestamp))
                raw_res = pd.DataFrame([PD_current_X_LO, PD_current_Y_LO]).transpose()
                raw_res.columns = ['PD_currentX_LO', 'PD_currentY_LO']
                raw_res.to_csv(res_raw)

            self.sig_progress.emit(round(20 + (75 / len(mawin.channel) * p)))
            tt=config+[ch, XIBW3dB, XQBW3dB, YIBW3dB, YQBW3dB, pe_x, pe_y, skew_x, skew_y, res_lo_x, res_lo_y,
                       res_sig_x, res_sig_y, per_x, per_y, dark_x, dark_y]
            # generate the report and print out the log file
            test_result.loc[p-1]=tt
            test_result.to_csv(report_file, index=False)
            # Close the figs
            self.sig_plotClose.emit()
        # close the laser output
        mawin.ICRtf.write('OUTP1:CHAN1:STATE OFF')
        mawin.ICRtf.write('OUTP1:CHAN2:STATE OFF')

        # wb.close()
        # os.system("explorer "+report_judge)
        self.sig_print.emit('测试完成!')
        self.sig_staColor.emit('green')
        self.sig_but.emit('开始')
        self.sig_status.emit('测试完成!')
        self.sig_progress.emit(100)
        print('测试完成')
        # Copy golden CSV data to folder under golden sample
        if mawin.test_type == '金样':
            golden_path = os.path.join(os.path.split(os.path.split(report_path3)[0])[0], 'Monitor_data')
            if not os.path.exists(golden_path): os.mkdir(golden_path)
            shutil.copy(report_file, golden_path)
        ##Write the data into report model and open the report after finished
        wb = xw.Book(mawin.report_model.replace('test_report.xlsx', 'test_report_ICR.xlsx'))
        worksht = wb.sheets(1)
        worksht.activate()
        worksht.range((1, 2)).value = test_result.iloc[0, 0]
        worksht.range((2, 2)).value = test_result.iloc[0, 1]
        worksht.range((3, 2)).value = test_result.iloc[0, 2]
        worksht.range((4, 2)).value = test_result.iloc[0, 3]
        worksht.range((5, 2)).value = test_result.iloc[0, 4]
        # write CLPD,FPGA,MCU information here into the test report
        # testBrdClp,ctlBrdClp,ctlBrdModVer,ctlBrdFPGAver,MCUver
        worksht.range((2, 4)).value = testBrdClp
        worksht.range((3, 4)).value = ctlBrdClp
        worksht.range((4, 4)).value = ctlBrdModVer
        worksht.range((5, 4)).value = ctlBrdFPGAver
        worksht.range((6, 4)).value = MCUver
        worksht.range((1, 4)).value = mawin.sw

        worksht.range((7, 2)).options(index=False, header=False, transpose=True).value = test_result.iloc[:, 5:]
        mawin.finalResult = worksht.range((6, 2)).value
        wb.save(report_judge)
        wb.close()
        # copy the local data to network folder
        if not network_path == False:
            shutil.copytree(report_path3, network_path)
        wb = xw.Book(report_judge)
        wb.sheets(1).activate()

    def ICR_test_old(self):
        '''
        #ICR test main process
        :return:NA
        '''
        # firstly rename the file of ICR test config of 'Board up,Drv up,Drv down'
        # Because no need to open ITLA, and need to power down PDs
        mawin.board_up = os.path.join(mawin.config_path, 'Setup_brdup_CtrlboardA001_20220113_ICR.txt')
        # Get and judge the SN format
        sn = str(mawin.lineEdit.text()).strip()
        if not gf.SN_check(sn):
            # mawin.test_status.setText('SN输入有误，请检查SN！')
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
        if not cb.open_board(mawin, mawin.ctrl_port):
            self.sig_status.emit('请检查控制板串口！')
            self.sig_but.emit('开始')
            self.sig_staColor.emit('red')
            self.sig_progress.emit(0)
            return
        #####Oscilloscope and ICR tester connection
        if not osc.open_OScope(mawin, mawin.OScope_port):
            self.sig_status.emit('请检查示波器端口！')
            self.sig_but.emit('开始')
            self.sig_staColor.emit('red')
            self.sig_progress.emit(0)
            return

        if not icr.open_ICRtf(mawin, mawin.ICRtf_port):
            self.sig_status.emit('请检查ICR测试平台串口！')
            self.sig_but.emit('开始')
            self.sig_staColor.emit('red')
            self.sig_progress.emit(0)
            return

        ###connectivity test
        self.sig_progress.emit(10)
        self.sig_print.emit('检查pin脚连接中...')
        # New method to check the connectivity and check TIA/DRV is ALU or IDT
        con = 0
        con = cb.test_connectivity_new_ICR(mawin, sn)
        if con == 1:
            self.sig_print.emit('Driver is ALU')
            mawin.device_type = 'ALU'
            print('检测到ALU器件,将执行ALU器件相关的测试...')
        elif con == 2:
            self.sig_print.emit('Driver is IDT')
            mawin.device_type = 'IDT'
            print('检测到IDT器件,将执行IDT器件相关的测试...')
        elif con == 3:
            self.sig_print.emit('请检查光路或压接，连接性检查失败！！！')
            self.sig_staColor.emit('blue')
            self.sig_but.emit('开始')
            self.sig_status.emit('请检查器件连接!')
            # self.sig_stoptest.emit()
            self.sig_progress.emit(0)
            return

        # Read the CLPD,FPGA,MCU information
        testBrdClp = 'NA'
        ctlBrdClp = 'NA'
        ctlBrdModVer = 'NA'
        ctlBrdFPGAver = 'NA'
        MCUver = 'NA'
        testBrdClp, ctlBrdClp, ctlBrdModVer, ctlBrdFPGAver, MCUver = gf.read_MCUandFPGA(mawin)

        '''
        *****Start the whole test steps*****
        '''
        self.sig_progress.emit(15)
        self.sig_print.emit('器件连接成功, FS400初始化, 配置TIA及上电...')
        # TIA Power on and config
        if mawin.device_type=='IDT':
            mawin.drv_up = os.path.join(mawin.config_path, 'Setup_driverup_Ctrlboard56017837A002_20211203.txt')
            mawin.drv_down = os.path.join(mawin.config_path, 'Setup_driverdown_CtrlboardA001_20210820.txt')
            gf.TIA_on(mawin)
            gf.TIA_config(mawin)
            print('IDT器件测试')
            self.sig_print.emit('TIA上电完成，设备初始化中...')
        elif mawin.device_type == 'ALU':
            mawin.drv_up = os.path.join(mawin.config_path, 'Setup_driverup_Ctrlboard56017837A002_20211203_ALU.txt')
            mawin.drv_down = os.path.join(mawin.config_path, 'Setup_driverdown_CtrlboardA001_20210820_ALU.txt')
            gf.TIA_on(mawin)
            gf.TIA_config_ALU(mawin,mawin.test_res)
            print('ALU器件测试')

        # initiate Oscilloscope and switch to visa connection
        osc.switch_to_DSO(mawin, mawin.OScope_port)
        print('DSO object created')
        osc.switch_to_VISA(mawin, mawin.DSO_port)
        if mawin.test_pe or mawin.test_bw:
            osc.init_OSc(mawin)
            self.sig_print.emit('示波器初始化成功...')
        self.sig_print.emit('开启并设置Laser波长...')
        icr.set_laser(mawin, 1, 15.4, 10)
        icr.set_laser(mawin, 2, 15.4, -5, 1550.000)
        self.sig_progress.emit(20)
        # Test group by wavelength
        self.sig_print.emit('配置完成!开始ICR测试...')
        ###if normal ITLA(Not C++ ITLA) then wavelength minus 8
        # result_tmp=[]
        ##define the data frame to store the test data
        test_result = DataFrame(columns=('SN', 'RX_BW_DESK', 'TEMP', 'DATE', 'TIME', '400G_CH',
                                         'RX_BW_XI', 'RX_BW_XQ', 'RX_BW_YI', 'RX_BW_YQ', 'RX_PE_X',
                                         'RX_PE_Y', 'RX_X_Skew', 'RX_Y_Skew', 'RX_1PD_RES_LO_X', 'RX_1PD_RES_LO_Y',
                                         'RX_1PD_RES_SIG_X', 'RX_1PD_RES_SIG_Y', 'RX_PER_X', 'RX_PER_Y', 'I_Dark_X',
                                         'I_Dark_Y'))

        # # Add the folder named as SN_Date to store all the test data
        timestamp = gf.get_timestamp(1)
        report_path3, network_path = mawin.create_report_folders(sn, timestamp)
        if network_path == False:
            pass
        # create list to store the result
        tt = []
        config = [sn, mawin.desk, mawin.temp[0]] + timestamp.split('_')
        # report_name=sn+'_'+timestamp+'.csv'
        report_name = sn + '_ICR_' + timestamp + '.csv'  # BOB改了csv报告名字
        report_judgename = sn + '_ICR_Report_' + timestamp + '.xlsx'
        report_file = os.path.join(report_path3, report_name)
        report_judge = os.path.join(report_path3, report_judgename)

        for i in range(len(mawin.channel)):
            ch = str(mawin.channel[i])
            p = i + 1
            self.sig_print.emit("CH%s 测试开始...\n" % (ch))
            # ICR Test start and set the wavelength
            icr.set_wavelength(mawin, 2, ch, 0)
            icr.set_wavelength(mawin, 1, ch, 1)
            mawin.ICRtf.write('OUTP3:CHAN1:POW 6')  # modify the value from 10dBm to 9dBm
            mawin.ICRtf.write('OUTP3:CHAN2:POW -6')
            mawin.ICRtf.write('OUTP1:CHAN1:STATE ON')
            mawin.ICRtf.write('OUTP1:CHAN2:STATE ON')
            pe_x, pe_y, skew_x, skew_y = [0.0] * 4
            XIBW3dB, XQBW3dB, YIBW3dB, YQBW3dB = [0.0] * 4
            res_sig_x, res_sig_y, per_x, per_y, dark_x, dark_y, res_lo_x, res_lo_y = [0.0] * 8
            if mawin.test_pe:
                self.sig_print.emit('Phase 和 Skew 测试中...')
                # phase error test
                lamb = icr.cal_wavelength(ch, 0)
                m = 3  # 平均次数
                #icr.ICR_manualPol_ScopeBalance_Judge(mawin)
                icr.ICR_scamblePol(mawin)
                # query the data of amtiplitude
                A1 = mawin.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')
                A3 = mawin.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')
                while not icr.isfloat(A1) or not icr.isfloat(A3):
                    A1 = mawin.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')
                    A3 = mawin.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')
                while float(A1) / float(A3) > 1.08 or float(A1) / float(A3) < 0.92:
                    time.sleep(0.2)
                    #icr.ICR_manualPol_ScopeBalance_Judge(mawin)
                    icr.ICR_scamblePol(mawin)
                    A1 = mawin.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')
                    A3 = mawin.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')
                    while not icr.isfloat(A1) or not icr.isfloat(A3):
                        A1 = mawin.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')
                        A3 = mawin.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')
                # Scan RF start
                #data = mawin.scan_RF_ManualPol(ch)
                data = mawin.scan_RF(ch)
                # save the result
                # pe_skew='{}_CH{}_PeSkew_{}'.format(sn,ch,timestamp)
                # pd.DataFrame(data).to_csv(pe_skew)
                pe_x, pe_y, skew_x, skew_y, new_pahseX, new_pahseY = peSkew.deal_with_data(data)
                # save raw and fitted s21 data
                raw_phase_report = os.path.join(report_path3, '{}_CH{}_phase_RawData_{}.csv'.format(sn, ch, timestamp))
                col = ['Phase error', 'Skew'] + [str(i) + 'GHz' for i in range(1, 11)]
                pha = pd.DataFrame([[pe_x, skew_x] + new_pahseX, [pe_y, skew_y] + new_pahseY],
                                   index=('phase_X', 'phase_Y'), columns=col)
                pha.to_csv(raw_phase_report)
                # draw the curve in the main thread
                self.sig_plot.emit(pha, report_path3, ch, sn, timestamp, 'PE')

            # Work left to do here to calculate SKEW PE
            if mawin.test_bw:
                # Rx BW test
                self.sig_print.emit('Rx BW 测试中...')
                icr.set_wavelength(mawin, 2, ch, 0)
                icr.set_wavelength(mawin, 1, ch, 1)
                time.sleep(0.1)
                error = True
                eg = 1  # 表示在某一个波长处测试的次数
                while error:
                    icr.ICR_scamblePol(mawin)
                    #icr.ICR_manualPol_ScopeBalance_Judge(mawin)
                    # query the data of amtiplitude
                    A1 = mawin.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')
                    A3 = mawin.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')
                    while not icr.isfloat(A1) or not icr.isfloat(A3):
                        time.sleep(0.2)
                        A1 = mawin.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')
                        A3 = mawin.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')
                    while float(A1) / float(A3) > 1.15 or float(A1) / float(A3) < 0.85:
                        time.sleep(0.2)
                        icr.ICR_scamblePol(mawin)
                        #icr.ICR_manualPol_ScopeBalance_Judge(mawin)
                        A1 = mawin.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')
                        A3 = mawin.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')
                        while not icr.isfloat(A1) or not icr.isfloat(A3):
                            A1 = mawin.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')
                            A3 = mawin.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')
                    # scan S21
                    print('A1:{}\nA3:{}'.format(A1, A3))
                    data1 = mawin.scan_S21(ch, 1, 40, 1)
                    #data1 = mawin.scan_S21_ManualPol(ch, 1, 40, 1)
                    # config the frequency
                    numfre = 40
                    da = [[0.0] * 4] * numfre
                    fre = [i * 1e9 for i in range(1, 41)]
                    # 归一化数据Y
                    raw = pd.DataFrame(data1)
                    for i in range(4):
                        raw.iloc[:, i] = 20 * np.log10(raw.iloc[:, i] / raw.iloc[0, i])
                    # 线损
                    linelo = pd.read_csv(mawin.Rx_line_loss, encoding='gb2312')
                    fre_loss = linelo.iloc[:-1, 0]
                    loss = linelo.iloc[:-1, 1]
                    Yi = np.interp(fre, fre_loss, loss)  # 插值求得对应频率的线损
                    # 减去线损
                    Ysmo = raw.sub(Yi, axis=0)
                    # save raw and fitted s21 data
                    raw_s21_report = os.path.join(report_path3, '{}_CH{}_RxBw_RawData_{}.csv'.format(sn, ch, timestamp))
                    fit_s21_report = os.path.join(report_path3, '{}_CH{}_RxBw_FitData_{}.csv'.format(sn, ch, timestamp))
                    colBW = ['XI', 'XQ', 'YI', 'YQ']
                    indBW = [i for i in range(1, 41)]
                    raw.columns = colBW
                    raw.index = indBW
                    raw.to_csv(raw_s21_report)  # ,index=False)
                    ###This point needs to be verified, whether the smooth data is equal to matlab gaussian method
                    for i in range(4):
                        tem = smooth.smooth(Ysmo.iloc[:, i], 3)[1:-1]
                        tem = tem - tem[0]  # 归一化数据
                        Ysmo.iloc[:, i] = tem
                    Ysmo.columns = colBW
                    Ysmo.index = indBW
                    Ysmo.to_csv(fit_s21_report)  # ,index=False)
                    # draw the curves
                    self.sig_plot.emit(Ysmo, report_path3, ch, sn, timestamp, 'BW')

                    # calculate the result, base on Ysmo is a dataframe type
                    for i in range(4):
                        aa = [x for x in range(Ysmo.shape[0]) if Ysmo.iloc[x, i] < -3]
                        if len(aa) == 0:
                            BW3dB = 40
                        else:
                            BW3dB = min(aa) + 1  # /1e9 index plus 1 as Freq
                        # each channel
                        if i == 0:
                            XIBW3dB = BW3dB
                            if min(Ysmo.iloc[:, i]) < -20.0:
                                XIBW3dB = 0.0
                        elif i == 1:
                            XQBW3dB = BW3dB
                            if min(Ysmo.iloc[:, i]) < -20.0:
                                XQBW3dB = 0.0
                        elif i == 2:
                            YIBW3dB = BW3dB
                            if min(Ysmo.iloc[:, i]) < -20.0:
                                YIBW3dB = 0.0
                        elif i == 3:
                            YQBW3dB = BW3dB
                            if min(Ysmo.iloc[:, i]) < -20.0:
                                YQBW3dB = 0.0

                    YsmoMean = Ysmo.mean(axis=1)
                    aa = [x for x in range(YsmoMean.shape[0]) if YsmoMean[x + 1] < -3]
                    if len(aa) == 0:
                        MeanBW3dB = 40
                    else:
                        MeanBW3dB = min(aa) + 1  # /1e9
                        if min(YsmoMean) < -20:
                            MeanBW3dB = 0

                    # Need to check whether to perform the judgement of >28G Hz and each channel > mean bw by 3GHz
                    # if not meet the requirement then perform the retest
                    # left to write the result
                    error = False  # Not check BW and retest
            # TIA amplitude test not performed
            # scan_Amp:to test the TIA output amplitude
            # left blank here
            if mawin.test_res:
                # DAC开启，PD上电
                mawin.CtrlB.write(b'switch_set 10 0\n');
                time.sleep(0.1)
                mawin.CtrlB.write(b'cpld_spi_wr 0x2c 2150\n');
                time.sleep(0.1)
                mawin.CtrlB.write(b'cpld_spi_wr 0x2f 2150\n');
                time.sleep(0.1)

                self.sig_print.emit('Rx响应度测试中...')
                res_sig_x, res_sig_y, per_x, per_y, dark_x, dark_y, PD_currentX, PD_currentY= mawin.get_pd_resp_sig(ch)
                res_raw = os.path.join(report_path3, '{}_CH{}_RxPDcurrentXY_{}.csv'.format(sn, ch, timestamp))
                raw_res = pd.DataFrame([PD_currentX, PD_currentY]).transpose()
                raw_res.columns = ['PD_currentX', 'PD_currentY']
                raw_res.to_csv(res_raw)
                # Draw the curve X Y phase
                self.sig_plot.emit(raw_res, report_path3, ch, sn, timestamp, 'RES')
                self.sig_print.emit('LO响应度测试中...')
                PD_current_X_LO = []
                PD_current_Y_LO = []
                for t in range(10):
                    res_lo_x, res_lo_y, curX, curY = mawin.get_pd_resp_lo(ch)
                    PD_current_X_LO.append(curX)
                    PD_current_Y_LO.append(curY)

                res_raw = os.path.join(report_path3, '{}_CH{}_RxPDcurrentXY_LO_{}.csv'.format(sn, ch, timestamp))
                raw_res = pd.DataFrame([PD_current_X_LO, PD_current_Y_LO]).transpose()
                raw_res.columns = ['PD_currentX_LO', 'PD_currentY_LO']
                raw_res.to_csv(res_raw)

            self.sig_progress.emit(round(20 + (75 / len(mawin.channel) * p)))
            tt=config+[ch, XIBW3dB, XQBW3dB, YIBW3dB, YQBW3dB, pe_x, pe_y, skew_x, skew_y, res_lo_x, res_lo_y,
                       res_sig_x, res_sig_y, per_x, per_y, dark_x, dark_y]
            # generate the report and print out the log file
            test_result.loc[p-1]=tt
            test_result.to_csv(report_file, index=False)
            # Close the figs
            self.sig_plotClose.emit()
        # close the laser output
        mawin.ICRtf.write('OUTP1:CHAN1:STATE OFF')
        mawin.ICRtf.write('OUTP1:CHAN2:STATE OFF')

        # wb.close()
        # os.system("explorer "+report_judge)
        self.sig_print.emit('测试完成!')
        self.sig_staColor.emit('green')
        self.sig_but.emit('开始')
        self.sig_status.emit('测试完成!')
        self.sig_progress.emit(100)
        print('测试完成')
        ##Write the data into report model and open the report after finished
        wb = xw.Book(mawin.report_model.replace('test_report.xlsx', 'test_report_ICR.xlsx'))
        worksht = wb.sheets(1)
        worksht.activate()
        worksht.range((1, 2)).value = test_result.iloc[0, 0]
        worksht.range((2, 2)).value = test_result.iloc[0, 1]
        worksht.range((3, 2)).value = test_result.iloc[0, 2]
        worksht.range((4, 2)).value = test_result.iloc[0, 3]
        worksht.range((5, 2)).value = test_result.iloc[0, 4]
        # write CLPD,FPGA,MCU information here into the test report
        # testBrdClp,ctlBrdClp,ctlBrdModVer,ctlBrdFPGAver,MCUver
        worksht.range((2, 4)).value = testBrdClp
        worksht.range((3, 4)).value = ctlBrdClp
        worksht.range((4, 4)).value = ctlBrdModVer
        worksht.range((5, 4)).value = ctlBrdFPGAver
        worksht.range((6, 4)).value = MCUver
        worksht.range((1, 4)).value = mawin.sw

        worksht.range((7, 2)).options(index=False, header=False, transpose=True).value = test_result.iloc[:, 5:]
        mawin.finalResult = worksht.range((6, 2)).value
        wb.save(report_judge)
        # copy the local data to network folder
        if not network_path == False:
            shutil.copytree(report_path3, network_path)

    def ICR_test_GBS(self):
        '''
        #ICR test Res experiment
        :return:NA
        '''
        # firstly rename the file of ICR test config of 'Board up,Drv up,Drv down'
        # Because no need to open ITLA, and need to power down PDs
        mawin.board_up = os.path.join(mawin.config_path, 'Setup_brdup_CtrlboardA001_20220113_Res.txt')
        # Get and judge the SN format
        sn = str(mawin.lineEdit.text()).strip()
        if not gf.SN_check(sn):
            self.sig_status.emit('SN输入有误，请检查SN！')
            self.sig_staColor.emit('red')
            self.sig_but.emit('开始')
            return
        self.sig_status.emit('测试进行中...')
        self.sig_staColor.emit('yellow')
        self.sig_clear.emit()
        self.sig_print.emit(sn)
        self.sig_progress.emit(5)

        #设备连接
        if not cb.open_board(mawin, mawin.ctrl_port):
            self.sig_status.emit('请检查控制板串口！')
            self.sig_but.emit('开始')
            self.sig_staColor.emit('red')
            self.sig_progress.emit(0)
            return
        #Oscilloscope and ICR tester connection
        if not osc.open_OScope(mawin, mawin.OScope_port):
            self.sig_status.emit('请检查示波器端口！')
            self.sig_but.emit('开始')
            self.sig_staColor.emit('red')
            self.sig_progress.emit(0)
            return

        if not icr.open_ICRtf(mawin, mawin.ICRtf_port):
            self.sig_status.emit('请检查ICR测试平台串口！')
            self.sig_but.emit('开始')
            self.sig_staColor.emit('red')
            self.sig_progress.emit(0)
            return

        #MPC connection
        if not mpc.open_MPC(mawin, mawin.MPC_port):
            self.sig_status.emit('请检查MPC串口！')
            self.sig_but.emit('开始')
            self.sig_staColor.emit('red')
            self.sig_progress.emit(0)
            return

        ###connectivity test
        self.sig_progress.emit(10)
        self.sig_print.emit('检查pin脚连接中...')
        # New method to check the connectivity and check TIA/DRV is ALU or IDT
        con = 0
        con = cb.test_connectivity_new_ICR(mawin, sn)
        if con == 1:
            self.sig_print.emit('Driver is ALU')
            mawin.device_type = 'ALU'
            print('检测到ALU器件,将执行ALU器件相关的测试...')
        elif con == 2:
            self.sig_print.emit('Driver is IDT')
            mawin.device_type = 'IDT'
            print('检测到IDT器件,将执行IDT器件相关的测试...')
        elif con == 3:
            self.sig_print.emit('请检查光路或压接，连接性检查失败！！！')
            self.sig_staColor.emit('blue')
            self.sig_but.emit('开始')
            self.sig_status.emit('请检查器件连接!')
            # self.sig_stoptest.emit()
            self.sig_progress.emit(0)
            return

        # Read the CLPD,FPGA,MCU information
        testBrdClp = 'NA'
        ctlBrdClp = 'NA'
        ctlBrdModVer = 'NA'
        ctlBrdFPGAver = 'NA'
        MCUver = 'NA'
        # testBrdClp, ctlBrdClp, ctlBrdModVer, ctlBrdFPGAver, MCUver = gf.read_MCUandFPGA(mawin)

        '''
        *****Start the whole test steps*****
        '''
        self.sig_progress.emit(15)
        self.sig_print.emit('器件连接成功, FS400初始化, 配置TIA及上电...')
        # TIA Power on and config
        # if mawin.device_type=='IDT':
        #     gf.TIA_on(mawin)
        #     gf.TIA_config(mawin)
        #     print('IDT器件测试')
        #     self.sig_print.emit('TIA上电完成，设备初始化中...')
        # elif mawin.device_type == 'ALU':
        #     mawin.drv_up = os.path.join(mawin.config_path, 'Setup_driverup_Ctrlboard56017837A002_20211203_ALU.txt')
        #     mawin.drv_down = os.path.join(mawin.config_path, 'Setup_driverdown_CtrlboardA001_20210820_ALU.txt')
        #     gf.TIA_on(mawin)
        #     gf.TIA_config_ALU(mawin,mawin.test_res)
        #     print('ALU器件测试')

        # # initiate Oscilloscope and switch to visa connection
        # osc.switch_to_DSO(mawin, mawin.OScope_port)
        # print('DSO object created')
        # osc.switch_to_VISA(mawin, mawin.DSO_port)
        # if mawin.test_pe or mawin.test_bw:
        #     osc.init_OSc(mawin)
        #     self.sig_print.emit('示波器初始化成功...')
        # self.sig_print.emit('开启并设置Laser波长...')
        # icr.set_laser(mawin, 1, 15.4, 10)
        # icr.set_laser(mawin, 2, 15.4, -5, 1550.000)
        self.sig_progress.emit(20)
        # Test group by wavelength
        self.sig_print.emit('配置完成!开始ICR测试...')
        ###if normal ITLA(Not C++ ITLA) then wavelength minus 8
        # result_tmp=[]
        ##define the data frame to store the test data
        test_result = DataFrame(columns=('SN', 'RX_BW_DESK', 'TEMP', 'DATE', 'TIME', '400G_CH',
                                         'RX_BW_XI', 'RX_BW_XQ', 'RX_BW_YI', 'RX_BW_YQ', 'RX_PE_X',
                                         'RX_PE_Y', 'RX_X_Skew', 'RX_Y_Skew', 'RX_1PD_RES_LO_X', 'RX_1PD_RES_LO_Y',
                                         'RX_1PD_RES_SIG_X', 'RX_1PD_RES_SIG_Y', 'RX_PER_X', 'RX_PER_Y', 'I_Dark_X',
                                         'I_Dark_Y'))

        # # Add the folder named as SN_Date to store all the test data
        timestamp = gf.get_timestamp(1)
        report_path3, network_path = mawin.create_report_folders(sn, timestamp)
        if network_path == False:
            pass
        # create list to store the result
        tt = []
        config = [sn, mawin.desk, mawin.temp[0]] + timestamp.split('_')
        # report_name=sn+'_'+timestamp+'.csv'
        report_name = sn + '_ICR_' + timestamp + '.csv'  # BOB改了csv报告名字
        report_judgename = sn + '_ICR_Report_' + timestamp + '.xlsx'
        report_file = os.path.join(report_path3, report_name)
        report_judge = os.path.join(report_path3, report_judgename)

        for i in range(len(mawin.channel)):
            ch = str(mawin.channel[i])
            p = i + 1
            self.sig_print.emit("CH%s 测试开始...\n" % (ch))
            pe_x, pe_y, skew_x, skew_y = [0.0] * 4
            XIBW3dB, XQBW3dB, YIBW3dB, YQBW3dB = [0.0] * 4
            res_sig_x, res_sig_y, per_x, per_y, dark_x, dark_y, res_lo_x, res_lo_y = [0.0] * 8

            if mawin.test_res:
                # Responsivity test
                # DAC开启，PD上电
                # PD voltage set to 2150
                mawin.CtrlB.write(b'switch_set 10 0\n');
                time.sleep(0.1)
                mawin.CtrlB.write(b'cpld_spi_wr 0x2c 2150\n');
                time.sleep(0.1)
                mawin.CtrlB.write(b'cpld_spi_wr 0x2f 2150\n');
                time.sleep(0.1)

                self.sig_print.emit('Rx响应度测试中...')
                res_sig_x, res_sig_y, per_x, per_y, dark_x, dark_y, PD_currentX, PD_currentY,id_markYmin,id_markXmin= mawin.get_pd_resp_sig_New_GBSmpc(ch)

                # Draw the Rx PD current data point and save the image
                # Save the raw data
                res_raw = os.path.join(report_path3, '{}_CH{}_RxPDcurrentXY_{}.csv'.format(sn, ch, timestamp))
                raw_res = pd.DataFrame([PD_currentX, PD_currentY]).transpose()
                raw_res.columns = ['PD_currentX', 'PD_currentY']
                raw_res.to_csv(res_raw)
                # Draw the curve X Y phase
                self.sig_plot.emit(raw_res, report_path3, ch, sn, timestamp, 'RES',id_markYmin,id_markXmin)
                self.sig_print.emit('LO响应度测试中...')
                res_lo_x,res_lo_y,curX,curY=mawin.get_pd_resp_lo(ch)

                # PD_current_X_LO = []
                # PD_current_Y_LO = []
                # for t in range(10):
                #     res_lo_x, res_lo_y, curX, curY = mawin.get_pd_resp_lo(ch)
                #     PD_current_X_LO.append(curX)
                #     PD_current_Y_LO.append(curY)
                #
                # res_raw = os.path.join(report_path3, '{}_CH{}_RxPDcurrentXY_LO_{}.csv'.format(sn, ch, timestamp))
                # raw_res = pd.DataFrame([PD_current_X_LO, PD_current_Y_LO]).transpose()
                # raw_res.columns = ['PD_currentX_LO', 'PD_currentY_LO']
                # raw_res.to_csv(res_raw)

            self.sig_progress.emit(round(20 + (75 / len(mawin.channel) * p)))
            tt=config+[ch, XIBW3dB, XQBW3dB, YIBW3dB, YQBW3dB, pe_x, pe_y, skew_x, skew_y, res_lo_x, res_lo_y,
                       res_sig_x, res_sig_y, per_x, per_y, dark_x, dark_y]
            # generate the report and print out the log file
            test_result.loc[p-1]=tt
            test_result.to_csv(report_file, index=False)
            # Close the figs
            self.sig_plotClose.emit()
        # close the laser output
        mawin.ICRtf.write('OUTP1:CHAN1:STATE OFF')
        mawin.ICRtf.write('OUTP1:CHAN2:STATE OFF')

        # wb.close()
        # os.system("explorer "+report_judge)
        self.sig_print.emit('测试完成!')
        self.sig_staColor.emit('green')
        self.sig_but.emit('开始')
        self.sig_status.emit('测试完成!')
        self.sig_progress.emit(100)
        print('测试完成')
        ##Write the data into report model and open the report after finished
        wb = xw.Book(mawin.report_model.replace('test_report.xlsx', 'test_report_ICR.xlsx'))
        worksht = wb.sheets(1)
        worksht.activate()
        worksht.range((1, 2)).value = test_result.iloc[0, 0]
        worksht.range((2, 2)).value = test_result.iloc[0, 1]
        worksht.range((3, 2)).value = test_result.iloc[0, 2]
        worksht.range((4, 2)).value = test_result.iloc[0, 3]
        worksht.range((5, 2)).value = test_result.iloc[0, 4]
        # write CLPD,FPGA,MCU information here into the test report
        # testBrdClp,ctlBrdClp,ctlBrdModVer,ctlBrdFPGAver,MCUver
        worksht.range((2, 4)).value = testBrdClp
        worksht.range((3, 4)).value = ctlBrdClp
        worksht.range((4, 4)).value = ctlBrdModVer
        worksht.range((5, 4)).value = ctlBrdFPGAver
        worksht.range((6, 4)).value = MCUver
        worksht.range((1, 4)).value = mawin.sw

        worksht.range((7, 2)).options(index=False, header=False, transpose=True).value = test_result.iloc[:, 5:]
        mawin.finalResult = worksht.range((6, 2)).value
        wb.save(report_judge)
        # copy the local data to network folder
        # if not network_path == False:
        #     shutil.copytree(report_path3, network_path)
        #For auto test
        wb.close()
        mawin.CtrlB.close()
        mawin.MPC.close()

    def Res_test_GBS_old(self):
        '''
        #Responsivity test use GBS polorization controller
        :return:NA
        '''
        # firstly rename the file of ICR test config of 'Board up,Drv up,Drv down'
        mawin.board_up = os.path.join(mawin.config_path, 'Setup_brdup_CtrlboardA001_20220113_Res.txt')
        # Get and judge the SN format
        sn = str(mawin.lineEdit.text()).strip()
        if not gf.SN_check(sn):
            # mawin.test_status.setText('SN输入有误，请检查SN！')
            self.sig_status.emit('SN输入有误，请检查SN！')
            self.sig_staColor.emit('red')
            self.sig_but.emit('开始')
            return

        # # Add the folder named as SN_Date to store all the test data
        timestamp = gf.get_timestamp(1)
        report_path3, network_path = mawin.create_report_folders(sn, timestamp)
        if network_path == False:
            pass
        shutil.copy(mawin.config_file, report_path3)
        sys.stdout = gf.Logger(path=report_path3)
        #print(sn)

        self.sig_status.emit('测试进行中...')
        self.sig_staColor.emit('yellow')
        self.sig_clear.emit()
        self.sig_print.emit(sn)
        self.sig_progress.emit(5)

        ###设备连接
        if not cb.open_board(mawin, mawin.ctrl_port):
            self.sig_status.emit('请检查控制板串口！')
            self.sig_but.emit('开始')
            self.sig_staColor.emit('red')
            self.sig_progress.emit(0)
            return

        #MPC connection
        if not mpc.open_MPC(mawin, mawin.MPC_port):
            self.sig_status.emit('请检查MPC串口！')
            self.sig_but.emit('开始')
            self.sig_staColor.emit('red')
            self.sig_progress.emit(0)
            return

        ###connectivity test
        self.sig_progress.emit(10)
        self.sig_print.emit('检查pin脚连接中...')
        # New method to check the connectivity and check TIA/DRV is ALU or IDT
        con = 0
        con = cb.test_connectivity_new_ICR(mawin, sn)
        if con == 1:
            self.sig_print.emit('Driver is ALU')
            mawin.device_type = 'ALU'
            print('检测到ALU器件,将执行ALU器件相关的测试...')
        elif con == 2:
            self.sig_print.emit('Driver is IDT')
            mawin.device_type = 'IDT'
            print('检测到IDT器件,将执行IDT器件相关的测试...')
        elif con == 3:
            self.sig_print.emit('请检查光路或压接，连接性检查失败！！！')
            self.sig_staColor.emit('blue')
            self.sig_but.emit('开始')
            self.sig_status.emit('请检查器件连接!')
            # self.sig_stoptest.emit()
            self.sig_progress.emit(0)
            return

        # TIA Power on and config <<this is also provide the voltage to RF PDs>>
        gf.TIA_on(mawin)

        # Read the CLPD,FPGA,MCU information
        testBrdClp = 'NA'
        ctlBrdClp = 'NA'
        ctlBrdModVer = 'NA'
        ctlBrdFPGAver = 'NA'
        MCUver = 'NA'
        # testBrdClp, ctlBrdClp, ctlBrdModVer, ctlBrdFPGAver, MCUver = gf.read_MCUandFPGA(mawin)

        '''
        *****Start the whole test steps*****
        '''
        self.sig_progress.emit(20)
        # Test group by wavelength
        self.sig_print.emit('配置完成!开始Res测试...')
        ###if normal ITLA(Not C++ ITLA) then wavelength minus 8
        # result_tmp=[]
        ##define the data frame to store the test data
        test_result = DataFrame(columns=('SN', 'RX_BW_DESK', 'TEMP', 'DATE', 'TIME', '400G_CH',
                                         'RX_1PD_RES_LO_X', 'RX_1PD_RES_LO_Y',
                                         'RX_1PD_RES_SIG_X', 'RX_1PD_RES_SIG_Y', 'RX_PER_X', 'RX_PER_Y', 'I_Dark_X',
                                         'I_Dark_Y'))


        # create list to store the result
        tt = []
        config = [sn, mawin.desk, mawin.temp[0]] + timestamp.split('_')
        # report_name=sn+'_'+timestamp+'.csv'
        report_name = sn + '_Res_' + timestamp + '.csv'
        report_judgename = sn + '_Res_Report_' + timestamp + '.xlsx'
        report_file = os.path.join(report_path3, report_name)
        report_judge = os.path.join(report_path3, report_judgename)
        code_diffJudge=True

        for i in range(len(mawin.channel)):
            ch = str(mawin.channel[i])
            p = i + 1
            self.sig_print.emit("CH%s 测试开始...\n" % (ch))
            res_sig_x, res_sig_y, per_x, per_y, dark_x, dark_y, res_lo_x, res_lo_y = [0.0] * 8
            #Switch the channel,Rx first and Lo follows
            #Switch wavelength
            cm1='itla_wr 0 0x30 {}\n'.format(str(hex(int(ch)))).encode('utf-8')
            cm2='itla_wr 1 0x30 {}\n'.format(str(hex(int(ch)))).encode('utf-8')
            mawin.CtrlB.write(cm1);
            time.sleep(0.1)
            print(mawin.CtrlB.read_until(b'Write itla'))
            #ITLA1 needs to write wavelength twice
            mawin.CtrlB.write(cm2);
            time.sleep(0.1)
            print(mawin.CtrlB.read_until(b'Write itla'))
            mawin.CtrlB.write(cm2);
            time.sleep(0.1)
            print(mawin.CtrlB.read_until(b'Write itla'))
            # Responsivity test
            # PD voltage set to 2150
            mawin.CtrlB.write(b'switch_set 10 0\n');time.sleep(0.1)
            mawin.CtrlB.write(b'cpld_spi_wr 0x2c 2150\n');time.sleep(0.1)
            mawin.CtrlB.write(b'cpld_spi_wr 0x2f 2150\n');time.sleep(0.1)
            self.sig_print.emit('Rx响应度测试中...')
            res_sig_x, res_sig_y, per_x, per_y, dark_x, dark_y, PD_currentX, PD_currentY,id_markYmin,id_markXmin= mawin.get_pd_resp_sig_New_GBSmpc(ch)
            # Draw the Rx PD current data point and save the image
            # Save the raw data
            res_raw = os.path.join(report_path3, '{}_CH{}_RxPDcurrentXY_{}.csv'.format(sn, ch, timestamp))
            raw_res = pd.DataFrame([PD_currentX, PD_currentY]).transpose()
            raw_res.columns = ['PD_currentX', 'PD_currentY']
            raw_res.to_csv(res_raw)
            # Draw the curve X Y phase
            self.sig_plot.emit(raw_res, report_path3, ch, sn, timestamp, 'RES',id_markYmin,id_markXmin)
            self.sig_print.emit('LO响应度测试中...')
            #res_lo_x,res_lo_y,curX,curY=mawin.get_pd_resp_lo_ITLA80C(ch)
            PD_current_X_LO = []
            PD_current_Y_LO = []
            for t in range(10):
                res_lo_x,res_lo_y,curX,curY=mawin.get_pd_resp_lo_ITLA80C(ch)
                PD_current_X_LO.append(curX)
                PD_current_Y_LO.append(curY)

            res_raw = os.path.join(report_path3, '{}_CH{}_RxPDcurrentXY_LO_{}.csv'.format(sn, ch, timestamp))
            raw_res = pd.DataFrame([PD_current_X_LO, PD_current_Y_LO]).transpose()
            raw_res.columns = ['PD_currentX_LO', 'PD_currentY_LO']
            raw_res.to_csv(res_raw)
            self.sig_plot.emit(raw_res, report_path3, ch, sn, timestamp, 'RES_LO', [], [])


            # Res retest with higher voltage
            # PD voltage set to 2700
            mawin.CtrlB.write(b'switch_set 10 0\n');
            time.sleep(0.1)
            mawin.CtrlB.write(b'cpld_spi_wr 0x2c 2700\n');
            time.sleep(0.1)
            mawin.CtrlB.write(b'cpld_spi_wr 0x2f 2700\n');
            time.sleep(0.1)
            # self.sig_print.emit('Rx响应度测试中...')
            # timestamp_retest='2700vcode_'+timestamp
            # res_sig_x1, res_sig_y1, per_x1, per_y1, dark_x1, dark_y1, PD_currentX1, PD_currentY1, id_markYmin1, id_markXmin1 = mawin.get_pd_resp_sig_New_GBSmpc(
            #     ch)
            # # Draw the Rx PD current data point and save the image
            # # Save the raw data
            # res_raw = os.path.join(report_path3, '{}_CH{}_RxPDcurrentXY_{}.csv'.format(sn, ch, timestamp_retest))
            # raw_res = pd.DataFrame([PD_currentX1, PD_currentY1]).transpose()
            # raw_res.columns = ['PD_currentX', 'PD_currentY']
            # raw_res.to_csv(res_raw)
            # # Draw the curve X Y phase
            # self.sig_plot.emit(raw_res, report_path3, ch, sn, timestamp_retest, 'RES', id_markYmin1, id_markXmin1)
            self.sig_print.emit('LO响应度测试中...')
            PD_current_X_LO = []
            PD_current_Y_LO = []
            for t in range(10):
                res_lo_x1, res_lo_y1, curX1, curY1 = mawin.get_pd_resp_lo_ITLA80C(ch)
                PD_current_X_LO.append(curX1)
                PD_current_Y_LO.append(curY1)

            res_raw = os.path.join(report_path3, '{}_CH{}_RxPDcurrentXY_LO_{}.csv'.format(sn, ch, timestamp_retest))
            raw_res = pd.DataFrame([PD_current_X_LO, PD_current_Y_LO]).transpose()
            raw_res.columns = ['PD_currentX_LO', 'PD_currentY_LO']
            raw_res.to_csv(res_raw)
            self.sig_plot.emit(raw_res, report_path3, ch, sn, timestamp_retest, 'RES_LO', [], [])
            # PD_current_X_LO = []
            # PD_current_Y_LO = []
            # for t in range(10):
            #     res_lo_x, res_lo_y, curX, curY = mawin.get_pd_resp_lo(ch)
            #     PD_current_X_LO.append(curX)
            #     PD_current_Y_LO.append(curY)
            #
            # res_raw = os.path.join(report_path3, '{}_CH{}_RxPDcurrentXY_LO_{}.csv'.format(sn, ch, timestamp))
            # raw_res = pd.DataFrame([PD_current_X_LO, PD_current_Y_LO]).transpose()
            # raw_res.columns = ['PD_currentX_LO', 'PD_currentY_LO']
            # raw_res.to_csv(res_raw)

            self.sig_progress.emit(round(20 + (75 / len(mawin.channel) * p)))
            # if not 0.5<res_sig_x/res_sig_x1<2 or not 0.5<res_sig_y/res_sig_y1<2 or not 0.5<res_lo_x/res_lo_x1<2 or not 0.5<res_lo_y/res_lo_y1<2:
            if not 0.5 < res_lo_x / res_lo_x1 < 2 or not 0.5 < res_lo_y / res_lo_y1 < 2:
                print('vcode 2700 and 2150 responsivity variation too large,plese check!')
                print('continue test and record the result')
                code_diffJudge = False
            tt=config+[ch, res_lo_x, res_lo_y,res_sig_x, res_sig_y, per_x, per_y, dark_x, dark_y]
            # generate the report and print out the log file
            test_result.loc[p-1]=tt
            test_result.to_csv(report_file, index=False)
            # Close the figs
            self.sig_plotClose.emit()
        # close the laser output
        mawin.CtrlB.write(b'itla_wr 0 0x32 0x00\n');
        time.sleep(0.1)
        print(mawin.CtrlB.read_until(b'Write itla'))
        print('Rx laser closed...')
        mawin.CtrlB.write(b'itla_wr 1 0x32 0x00\n');
        time.sleep(0.1)  # shut down LO port
        print(mawin.CtrlB.read_until(b'Write itla'))
        print('Lo laser closed...')

        # os.system("explorer "+report_judge)
        self.sig_print.emit('测试完成!')
        self.sig_staColor.emit('green')
        self.sig_but.emit('开始')
        self.sig_status.emit('测试完成!')
        self.sig_progress.emit(100)
        print('测试完成')
        ##Write the data into report model and open the report after finished
        wb = xw.Book(mawin.report_model.replace('test_report.xlsx', 'test_report_Res.xlsx'))
        worksht = wb.sheets(1)
        worksht.activate()
        worksht.range((1, 2)).value = test_result.iloc[0, 0]
        worksht.range((2, 2)).value = test_result.iloc[0, 1]
        worksht.range((3, 2)).value = test_result.iloc[0, 2]
        worksht.range((4, 2)).value = test_result.iloc[0, 3]
        worksht.range((5, 2)).value = test_result.iloc[0, 4]
        # write CLPD,FPGA,MCU information here into the test report
        # testBrdClp,ctlBrdClp,ctlBrdModVer,ctlBrdFPGAver,MCUver
        worksht.range((2, 4)).value = testBrdClp
        worksht.range((3, 4)).value = ctlBrdClp
        worksht.range((4, 4)).value = ctlBrdModVer
        worksht.range((5, 4)).value = ctlBrdFPGAver
        worksht.range((6, 4)).value = MCUver
        worksht.range((1, 4)).value = mawin.sw

        worksht.range((7, 2)).options(index=False, header=False, transpose=True).value = test_result.iloc[:, 5:]
        mawin.finalResult = worksht.range((6, 2)).value
        if not code_diffJudge:
            mawin.finalResult='Fail'
            print('Vcode test result differ,please check the test data...')
            #worksht.range((6, 2)).value='Vcode differ'
        wb.save(report_judge)
        # copy the local data to network folder
        # if not network_path == False:
        #     shutil.copytree(report_path3, network_path)
        #For auto test
        #wb.close()
        mawin.CtrlB.close()
        mawin.MPC.close()

    def Res_test_GBS(self):
        '''
        #Responsivity test use GBS polorization controller
        :return:NA
        '''
        # firstly rename the file of ICR test config of 'Board up,Drv up,Drv down'
        mawin.board_up = os.path.join(mawin.config_path, 'Setup_brdup_CtrlboardA001_20220113_Res.txt')
        # Get and judge the SN format
        sn = str(mawin.lineEdit.text()).strip()
        if not gf.SN_check(sn):
            # mawin.test_status.setText('SN输入有误，请检查SN！')
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
        if not cb.open_board(mawin, mawin.ctrl_port):
            self.sig_status.emit('请检查控制板串口！')
            self.sig_but.emit('开始')
            self.sig_staColor.emit('red')
            self.sig_progress.emit(0)
            return

        #MPC connection
        if not mpc.open_MPC(mawin, mawin.MPC_port):
            self.sig_status.emit('请检查MPC串口！')
            self.sig_but.emit('开始')
            self.sig_staColor.emit('red')
            self.sig_progress.emit(0)
            return

        ###connectivity test
        self.sig_progress.emit(10)
        self.sig_print.emit('检查pin脚连接中...')
        # New method to check the connectivity and check TIA/DRV is ALU or IDT
        con = 0
        con = cb.test_connectivity_new_ICR(mawin, sn)
        if con == 1:
            self.sig_print.emit('Driver is ALU')
            mawin.device_type = 'ALU'
            print('检测到ALU器件,将执行ALU器件相关的测试...')
        elif con == 2:
            self.sig_print.emit('Driver is IDT')
            mawin.device_type = 'IDT'
            print('检测到IDT器件,将执行IDT器件相关的测试...')
        elif con == 3:
            self.sig_print.emit('请检查光路或压接，连接性检查失败！！！')
            self.sig_staColor.emit('blue')
            self.sig_but.emit('开始')
            self.sig_status.emit('请检查器件连接!')
            # self.sig_stoptest.emit()
            self.sig_progress.emit(0)
            return

        # TIA Power on and config <<this is also provide the voltage to RF PDs>>
        gf.TIA_on(mawin)

        # Read the CLPD,FPGA,MCU information
        testBrdClp = 'NA'
        ctlBrdClp = 'NA'
        ctlBrdModVer = 'NA'
        ctlBrdFPGAver = 'NA'
        MCUver = 'NA'
        # testBrdClp, ctlBrdClp, ctlBrdModVer, ctlBrdFPGAver, MCUver = gf.read_MCUandFPGA(mawin)

        '''
        *****Start the whole test steps*****
        '''
        self.sig_progress.emit(20)
        # Test group by wavelength
        self.sig_print.emit('配置完成!开始Res测试...')
        ###if normal ITLA(Not C++ ITLA) then wavelength minus 8
        # result_tmp=[]
        ##define the data frame to store the test data
        test_result = DataFrame(columns=('SN', 'RX_BW_DESK', 'TEMP', 'DATE', 'TIME', '400G_CH',
                                         'RX_1PD_RES_LO_X', 'RX_1PD_RES_LO_Y',
                                         'RX_1PD_RES_SIG_X', 'RX_1PD_RES_SIG_Y', 'RX_PER_X', 'RX_PER_Y', 'I_Dark_X',
                                         'I_Dark_Y'))

        # # Add the folder named as SN_Date to store all the test data
        timestamp = gf.get_timestamp(1)
        report_path3, network_path = mawin.create_report_folders(sn, timestamp)
        if network_path == False:
            pass
        shutil.copy(mawin.config_file, report_path3)
        sys.stdout = gf.Logger(path=report_path3)
        print(sn)
        # create list to store the result
        tt = []
        config = [sn, mawin.desk, mawin.temp[0]] + timestamp.split('_')
        # report_name=sn+'_'+timestamp+'.csv'
        report_name = sn + '_Res_' + timestamp + '.csv'
        report_judgename = sn + '_Res_Report_' + timestamp + '.xlsx'
        report_file = os.path.join(report_path3, report_name)
        report_judge = os.path.join(report_path3, report_judgename)
        code_diffJudge=True

        for i in range(len(mawin.channel)):
            ch = str(mawin.channel[i])
            p = i + 1
            self.sig_print.emit("CH%s 测试开始...\n" % (ch))
            res_sig_x, res_sig_y, per_x, per_y, dark_x, dark_y, res_lo_x, res_lo_y = [0.0] * 8
            # Switch the channel,Rx first and Lo follows
            # # close the laser output
            # mawin.CtrlB.write(b'itla_wr 0 0x32 0x00\n');
            # time.sleep(0.1)
            # print(mawin.CtrlB.read_until(b'Write itla'))
            # print('Rx laser closed...')
            # mawin.CtrlB.write(b'itla_wr 1 0x32 0x00\n');
            # time.sleep(0.1)  # shut down LO port
            # print(mawin.CtrlB.read_until(b'Write itla'))
            # print('Lo laser closed...')
            #Switch wavelength
            cm1='itla_wr 0 0x30 {}\n'.format(str(hex(int(ch)))).encode('utf-8')
            cm2='itla_wr 1 0x30 {}\n'.format(str(hex(int(ch)))).encode('utf-8')
            mawin.CtrlB.write(cm1);
            time.sleep(0.1)
            print(mawin.CtrlB.read_until(b'Write itla'))
            #ITLA1 needs to write wavelength twice
            mawin.CtrlB.write(cm2);
            time.sleep(0.1)
            print(mawin.CtrlB.read_until(b'Write itla'))
            mawin.CtrlB.write(cm2);
            time.sleep(0.1)
            print(mawin.CtrlB.read_until(b'Write itla'))
            # Responsivity test
            # self.sig_print.emit('响应度测试中...')
            # DAC开启，PD上电
            # mawin.CtrlB.write(b'switch_set 10 0\n');time.sleep(0.1)
            # mawin.CtrlB.write(b'cpld_spi_wr 0x2c 2700\n');time.sleep(0.1)
            # mawin.CtrlB.write(b'cpld_spi_wr 0x2f 2700\n');time.sleep(0.1)
            # PD voltage set to 2150
            mawin.CtrlB.write(b'switch_set 10 0\n');time.sleep(0.1)
            mawin.CtrlB.write(b'cpld_spi_wr 0x2c 2150\n');time.sleep(0.1)
            mawin.CtrlB.write(b'cpld_spi_wr 0x2f 2150\n');time.sleep(0.1)
            self.sig_print.emit('Rx响应度测试中...')
            res_sig_x, res_sig_y, per_x, per_y, dark_x, dark_y, PD_currentX, PD_currentY,id_markYmin,id_markXmin= mawin.get_pd_resp_sig_New_GBSmpc(ch)
            # Draw the Rx PD current data point and save the image
            # Save the raw data
            res_raw = os.path.join(report_path3, '{}_CH{}_RxPDcurrentXY_{}.csv'.format(sn, ch, timestamp))
            raw_res = pd.DataFrame([PD_currentX, PD_currentY]).transpose()
            raw_res.columns = ['PD_currentX', 'PD_currentY']
            raw_res.to_csv(res_raw)
            # Draw the curve X Y phase
            self.sig_plot.emit(raw_res, report_path3, ch, sn, timestamp, 'RES',id_markYmin,id_markXmin)
            self.sig_print.emit('LO响应度测试中...')
            #res_lo_x,res_lo_y,curX,curY=mawin.get_pd_resp_lo_ITLA80C(ch)
            PD_current_X_LO = []
            PD_current_Y_LO = []
            for t in range(4):
                res_lo_x,res_lo_y,curX,curY=mawin.get_pd_resp_lo_ITLA80C(ch)
                PD_current_X_LO.append(curX)
                PD_current_Y_LO.append(curY)

            res_raw = os.path.join(report_path3, '{}_CH{}_RxPDcurrentXY_LO_{}.csv'.format(sn, ch, timestamp))
            raw_res = pd.DataFrame([PD_current_X_LO, PD_current_Y_LO]).transpose()
            raw_res.columns = ['PD_currentX_LO', 'PD_currentY_LO']
            raw_res.to_csv(res_raw)
            self.sig_plot.emit(raw_res, report_path3, ch, sn, timestamp, 'RES_LO', [], [])


            # Res retest with higher voltage
            # PD voltage set to 2700
            mawin.CtrlB.write(b'switch_set 10 0\n');
            time.sleep(0.1)
            mawin.CtrlB.write(b'cpld_spi_wr 0x2c 2700\n');
            time.sleep(0.1)
            mawin.CtrlB.write(b'cpld_spi_wr 0x2f 2700\n');
            time.sleep(0.1)
            # self.sig_print.emit('Rx响应度测试中...')
            timestamp_retest='2700vcode_'+timestamp
            # res_sig_x1, res_sig_y1, per_x1, per_y1, dark_x1, dark_y1, PD_currentX1, PD_currentY1, id_markYmin1, id_markXmin1 = mawin.get_pd_resp_sig_New_GBSmpc(
            #     ch)
            # # Draw the Rx PD current data point and save the image
            # # Save the raw data
            # res_raw = os.path.join(report_path3, '{}_CH{}_RxPDcurrentXY_{}.csv'.format(sn, ch, timestamp_retest))
            # raw_res = pd.DataFrame([PD_currentX1, PD_currentY1]).transpose()
            # raw_res.columns = ['PD_currentX', 'PD_currentY']
            # raw_res.to_csv(res_raw)
            # # Draw the curve X Y phase
            # self.sig_plot.emit(raw_res, report_path3, ch, sn, timestamp_retest, 'RES', id_markYmin1, id_markXmin1)
            self.sig_print.emit('LO@2700响应度测试中...')
            PD_current_X_LO = []
            PD_current_Y_LO = []
            for t in range(2):
                res_lo_x1, res_lo_y1, curX1, curY1 = mawin.get_pd_resp_lo_ITLA80C(ch)
                PD_current_X_LO.append(curX1)
                PD_current_Y_LO.append(curY1)

            res_raw = os.path.join(report_path3, '{}_CH{}_RxPDcurrentXY_LO_{}.csv'.format(sn, ch, timestamp_retest))
            raw_res = pd.DataFrame([PD_current_X_LO, PD_current_Y_LO]).transpose()
            raw_res.columns = ['PD_currentX_LO', 'PD_currentY_LO']
            raw_res.to_csv(res_raw)
            self.sig_plot.emit(raw_res, report_path3, ch, sn, timestamp_retest, 'RES_LO', [], [])
            # PD_current_X_LO = []
            # PD_current_Y_LO = []
            # for t in range(10):
            #     res_lo_x, res_lo_y, curX, curY = mawin.get_pd_resp_lo(ch)
            #     PD_current_X_LO.append(curX)
            #     PD_current_Y_LO.append(curY)
            #
            # res_raw = os.path.join(report_path3, '{}_CH{}_RxPDcurrentXY_LO_{}.csv'.format(sn, ch, timestamp))
            # raw_res = pd.DataFrame([PD_current_X_LO, PD_current_Y_LO]).transpose()
            # raw_res.columns = ['PD_currentX_LO', 'PD_currentY_LO']
            # raw_res.to_csv(res_raw)

            self.sig_progress.emit(round(20 + (75 / len(mawin.channel) * p)))
            # if not 0.5<res_sig_x/res_sig_x1<2 or not 0.5<res_sig_y/res_sig_y1<2 or not 0.5<res_lo_x/res_lo_x1<2 or not 0.5<res_lo_y/res_lo_y1<2:
            if not 0.5 < res_lo_x / res_lo_x1 < 2 or not 0.5 < res_lo_y / res_lo_y1 < 2:
                print('vcode 2700 and 2150 responsivity variation too large,plese check!')
                print('continue test and record the result')
                code_diffJudge = False
            tt=config+[ch, res_lo_x, res_lo_y,res_sig_x, res_sig_y, per_x, per_y, dark_x, dark_y]
            # generate the report and print out the log file
            test_result.loc[p-1]=tt
            test_result.to_csv(report_file, index=False)
            # Close the figs
            self.sig_plotClose.emit()
        # close the laser output
        mawin.CtrlB.write(b'itla_wr 0 0x32 0x00\n');
        time.sleep(0.1)
        print(mawin.CtrlB.read_until(b'Write itla'))
        print('Rx laser closed...')
        mawin.CtrlB.write(b'itla_wr 1 0x32 0x00\n');
        time.sleep(0.1)  # shut down LO port
        print(mawin.CtrlB.read_until(b'Write itla'))
        print('Lo laser closed...')

        # os.system("explorer "+report_judge)
        self.sig_print.emit('测试完成!')
        self.sig_staColor.emit('green')
        self.sig_but.emit('开始')
        self.sig_status.emit('测试完成!')
        self.sig_progress.emit(100)
        print('测试完成')
        # Copy golden CSV data to folder under golden sample
        if mawin.test_type == '金样':
            golden_path = os.path.join(os.path.split(os.path.split(report_path3)[0])[0], 'Monitor_data')
            if not os.path.exists(golden_path): os.mkdir(golden_path)
            shutil.copy(report_file, golden_path)
        ##Write the data into report model and open the report after finished
        wb = xw.Book(mawin.report_model.replace('test_report.xlsx', 'test_report_Res.xlsx'))
        worksht = wb.sheets(1)
        worksht.activate()
        worksht.range((1, 2)).value = test_result.iloc[0, 0]
        worksht.range((2, 2)).value = test_result.iloc[0, 1]
        worksht.range((3, 2)).value = test_result.iloc[0, 2]
        worksht.range((4, 2)).value = test_result.iloc[0, 3]
        worksht.range((5, 2)).value = test_result.iloc[0, 4]
        # write CLPD,FPGA,MCU information here into the test report
        # testBrdClp,ctlBrdClp,ctlBrdModVer,ctlBrdFPGAver,MCUver
        worksht.range((2, 4)).value = testBrdClp
        worksht.range((3, 4)).value = ctlBrdClp
        worksht.range((4, 4)).value = ctlBrdModVer
        worksht.range((5, 4)).value = ctlBrdFPGAver
        worksht.range((6, 4)).value = MCUver
        worksht.range((1, 4)).value = mawin.sw

        worksht.range((7, 2)).options(index=False, header=False, transpose=True).value = test_result.iloc[:, 5:]
        mawin.finalResult = worksht.range((6, 2)).value
        if not code_diffJudge:
            mawin.finalResult='Fail'
            print('Vcode test result differ,please check the test data...')
            #worksht.range((6, 2)).value='Vcode differ'
        wb.save(report_judge)
        wb.close()
        # copy the local data to network folder
        if not network_path == False:
            shutil.copytree(report_path3, network_path)
        #For auto test

        mawin.CtrlB.close()
        mawin.MPC.close()
        wb = xw.Book(report_judge)
        wb.sheets(1).activate()

    def PEskew_test(self):
        '''
        #Phase error and Skew test with jinli and nianyu namely control FS400 to receive and module to transmit
        :return:NA
        '''
        # firstly rename the file of ICR test config of 'Board up,Drv up,Drv down'
        mawin.board_up = os.path.join(mawin.config_path, 'Setup_brdup_CtrlboardA001_20220113_PEskew.txt')
        # Get and judge the SN format
        sn = str(mawin.lineEdit.text()).strip()
        if not gf.SN_check(sn):
            # mawin.test_status.setText('SN输入有误，请检查SN！')
            self.sig_status.emit('SN输入有误，请检查SN！')
            self.sig_staColor.emit('red')
            self.sig_but.emit('开始')
            return

        # # Add the folder named as SN_Date to store all the test data
        timestamp = gf.get_timestamp(1)
        report_path3, network_path = mawin.create_report_folders(sn, timestamp)
        if network_path == False:
            pass
        shutil.copy(mawin.config_file, report_path3)
        sys.stdout = gf.Logger(path=report_path3)

        self.sig_status.emit('测试进行中...')
        self.sig_staColor.emit('yellow')
        self.sig_clear.emit()
        self.sig_print.emit(sn)
        self.sig_progress.emit(5)

        #***------设备连接---------***
        if not cb.open_board_PeSkew(mawin, mawin.ctrl_port):
            self.sig_status.emit('请检查控制板串口！')
            self.sig_but.emit('开始')
            self.sig_staColor.emit('red')
            self.sig_progress.emit(0)
            return

        #MPC connection
        if not jl.open_JINLI(mawin, mawin.JinLi_port):
            self.sig_status.emit('请检查JINLI串口！')
            self.sig_but.emit('开始')
            self.sig_staColor.emit('red')
            self.sig_progress.emit(0)
            return

        # firstly judge if ITLA is CPP
        cpp_itla = True
        mawin.CtrlB.write(b'show_itla 0\n\r');
        time.sleep(0.5)
        re_itla = mawin.CtrlB.read_until(b'run_state').decode('utf-8').split('\n')
        print(re_itla)
        try:
            re_itla1 = [i.split(' ')[-1] for i in re_itla if 'g_itla_model' in i][0]
        except Exception as e:
            print('ITLA model get failed:\n', e)
            return
        print(re_itla1)
        if not '0x3' in re_itla1:
            print('ITLA is C band')
            cpp_itla = False
        else:
            print('ITLA is C ++')
            cpp_itla = True
        comlist = []
        fre = list
        if cpp_itla:
            comlist = ['task_ctrl 4 0',
                       'itla_write 0 0x32 0',
                       'itla_write 0 0x34 0x2ee',
                       'itla_write 0 0x35 0xbe',
                       'itla_write 0 0x36 0x1bd5',
                       'itla_write 0 0x30 1',
                       'itla_write 0 0x32 8']
            mawin.channel = [i for i in range(1, 81, 4)] + [80]
            fre = np.interp(mawin.channel, range(1, 81), mawin.wl_cpp)
        else:
            comlist = ['task_ctrl 4 0',
                       'itla_write 0 0x32 0',
                       'itla_write 0 0x34 0x2ee',
                       'itla_write 0 0x35 0xbf',
                       'itla_write 0 0x36 0xc35',
                       'itla_write 0 0x30 1',
                       'itla_write 0 0x32 8']
            mawin.channel = [i for i in range(9, 73, 4)] + [72]
            fre = np.interp(mawin.channel, range(9, 73), mawin.wl_c)
        cb.board_set_CmdList(mawin, comlist)
        mawin.channel = [i for i in range(9, 73)]
        #mawin.channel = [i for i in range(9, 12)]
        mawin.update_config()

        print('Laser opened!!!')
        # connectivity test
        self.sig_progress.emit(10)
        self.sig_print.emit('检查pin脚连接中...')
        # New method to check the connectivity and check TIA/DRV is ALU or IDT
        con = 0
        con = cb.test_connectivity_PeSkew(mawin, sn)
        if con == 1:
            self.sig_print.emit('Driver is ALU')
            mawin.device_type = 'ALU'
            print('检测到ALU器件,将执行ALU器件相关的测试...')
        elif con == 2:
            self.sig_print.emit('Driver is IDT')
            mawin.device_type = 'IDT'
            print('检测到IDT器件,将执行IDT器件相关的测试...')
        elif con == 3:
            self.sig_print.emit('请检查光路或压接，连接性检查失败！！！')
            self.sig_staColor.emit('blue')
            self.sig_but.emit('开始')
            self.sig_status.emit('请检查器件连接!')
            self.sig_progress.emit(0)
            return

        # Jinli start wait for 3mins about, switch to module control
        # mawin.JinLi.write(b'set_uart_out_print_flag 0x22\n\r')
        # mawin.JinLi.timeout = 210
        # ##mawin.JinLi.read_until(b'driver gain[0]')
        # print(mawin.JinLi.read_until(b'test_cb_moduleready'))
        #
        # mawin.JinLi.timeout = 2
        # # TIA Power on and config <<this is also provide the voltage to RF PDs>>
        # gf.TIA_on(mawin)

        # Read the CLPD,FPGA,MCU information
        testBrdClp = 'NA'
        ctlBrdClp = 'NA'
        ctlBrdModVer = 'NA'
        ctlBrdFPGAver = 'NA'
        MCUver = 'NA'
        # testBrdClp, ctlBrdClp, ctlBrdModVer, ctlBrdFPGAver, MCUver = gf.read_MCUandFPGA(mawin)

        '''
        *****Start the whole test steps*****
        '''
        self.sig_progress.emit(20)
        # Test group by wavelength
        self.sig_print.emit('配置完成!开始PE_SKEW测试...')
        #define the data frame to store the test data
        test_result = DataFrame(columns=('SN', 'RX_BW_DESK', 'TEMP', 'DATE', 'TIME', '400G_CH',
                                         'RX_PE_X_ave', 'RX_PE_Y_ave',
                                         'RX_PE_X_max', 'RX_PE_Y_max',
                                         'RX_PE_X_min', 'RX_PE_Y_min',
                                         'RX_PE_X','RX_PE_Y', 'RX_X_Skew', 'RX_Y_Skew'))

        # create list to store the result
        tt = []
        config = [sn, mawin.desk, mawin.temp[0]] + timestamp.split('_')
        report_name = sn + '_PeSkew_' + timestamp + '.csv'
        report_judgename = sn + '_PeSkew_Report_' + timestamp + '.xlsx'
        report_file = os.path.join(report_path3, report_name)
        report_judge = os.path.join(report_path3, report_judgename)

        for i in range(len(mawin.channel)):
            ch = str(mawin.channel[i])
            cmClose = b'itla_write 0 0x32 0\n\r'
            cmOpen = b'itla_write 0 0x32 8\n\r'
            if cpp_itla:
                ch_toset = int(ch)
                cm1 = 'itla_write 0 0x30 {}\n\r'.format(str(hex(int(ch)))).encode('utf-8')
            else:
                ch_toset = int(ch) - 8
                cm1 = 'itla_write 0 0x30 {}\n\r'.format(str(hex(int(ch) - 8))).encode('utf-8')

            p = i + 1
            self.sig_print.emit("CH%s 测试开始...\n" % (ch))
            pe_x, pe_y, skew_x, skew_y = [0.0] * 4

            #***-------Switch wavelength-------***
            # mawin.CtrlB.write(cmClose);
            # time.sleep(0.1)
            # print(mawin.CtrlB.read_until(b'itla_0_write'))
            # print('ITLA closed...')
            #ITLA1 needs to write wavelength twice
            # mawin.CtrlB.write(cm1);
            # time.sleep(0.1)
            # print(mawin.CtrlB.read_until(b'itla_0_write'))
            mawin.CtrlB.write(cm1)
            print(mawin.CtrlB.read_until(b'itla_0_write'))
            time.sleep(1)

            #Confirm the wavelength is right
            mawin.CtrlB.flushInput()
            time.sleep(0.1)
            mawin.CtrlB.flushOutput()
            time.sleep(0.1)
            mawin.CtrlB.write(b'itla_read 0 0x30\n\r')
            re_wl=mawin.CtrlB.read_until(b'shell').decode('utf-8')
            print(re_wl)
            while not 'read_itla' in re_wl:
                time.sleep(1)
                mawin.CtrlB.write(b'itla_read 0 0x30\n\r')
                re_wl = mawin.CtrlB.read_until(b'shell').decode('utf-8')
                print(re_wl)
            fre_get=int(re_wl.split('\n')[2].strip().replace('\r',''),16)
            cou1=0
            while not fre_get == ch_toset:
                print('The {} time set wavelength...'.format(cou1+1))
                mawin.CtrlB.write(cm1)
                print(mawin.CtrlB.read_until(b'itla_0_write'))
                time.sleep(1)
                mawin.CtrlB.write(b'itla_read 0 0x30\n\r')
                re_wl = mawin.CtrlB.read_until(b'shell').decode('utf-8')
                print(re_wl)
                while not 'read_itla' in re_wl:
                    time.sleep(1)
                    mawin.CtrlB.write(b'itla_read 0 0x30\n\r')
                    re_wl = mawin.read_until(b'shell').decode('utf-8')
                    print(re_wl)
                fre_get = int(re_wl.split('\n')[2].strip().replace('\r', ''), 16)
                cou1+=1
                if cou1>3:
                    print('Wavelength get error after 4 times retry: ', re_wl)
                    return

            print('ITLA wavelength switched to ch{}...'.format(ch))
            # mawin.CtrlB.write(cmOpen);
            # time.sleep(0.1)
            # print(mawin.CtrlB.read_until(b'itla_0_write'))
            # print('ITLA opened...')
            #***---------Switch JinLi wavelength(Tx)--------***
            jl.set_JINLIwave(mawin,ch)
            print('JINLI wavelength switched to ch{}...'.format(ch))
            #***---------Judge system connectivity--------***
            sta='0'
            count=0
            while sta!='255':
                if count>19:
                    print('Not connected after 20 times loop, please check connectivity!')
                    self.sig_print.emit('系统不通，请检查！')
                    return
                mawin.CtrlB.flushInput()
                mawin.CtrlB.flushOutput()
                mawin.CtrlB.write(b'show_dsp_ber_info\n\r')#;time.sleep(0.5)
                #ret=mawin.CtrlB.read_all().decode('utf-8')
                ret=mawin.CtrlB.read_until(b'line_fas_ber').decode('utf-8')
                print('DSP information:\n',ret)
                #tmp1=re.split('[:,\r]',tmp[ind+1]+tmp[ind+2])
                ret_1=ret.split('\n')
                sta_list=[i[i.index('(')+1:i.index(')')] for i in ret_1 if 'dsp_state' in i]
                if not len(sta_list)==0:sta=sta_list[0]
                count += 1
            self.sig_print.emit('系统连通，开始获取Phase error和Skew...')
            mawin.CtrlB.write(b'dsp_control_task 26 1 1\n\r')
            time.sleep(2)
            print(mawin.CtrlB.read_all().decode('utf-8'))
            mawin.CtrlB.write(b'dsp_set_op_trig_source 0 0 1\n\r')
            time.sleep(2)
            print(mawin.CtrlB.read_all().decode('utf-8'))
            mawin.CtrlB.write(b'trigger_monitors 1\n\r')
            time.sleep(2)
            print(mawin.CtrlB.read_all().decode('utf-8'))
            #***----Read the return value to get the result----***
            mawin.CtrlB.flushInput()
            mawin.CtrlB.flushOutput()
            mawin.CtrlB.write(b'get_line_optical_channel_monitors_all\n\r')
            mawin.CtrlB.timeout=60
            resu=mawin.CtrlB.read_until(b'rx_csr_max_v').decode('utf-8')
            print(resu)
            result_tmp=resu.split('\n')
            flag_PEx='rx_angle_average_h'
            flag_PEy = 'rx_angle_average_v'
            flag_SKEWx = 'rx_skew_average_h'
            flag_SKEWy = 'rx_skew_average_v'
            result_peskew_ave=[i for i in result_tmp if flag_PEx in i or flag_PEy in i or flag_SKEWx in i or flag_SKEWy in i]
            print('Average:\n',result_peskew_ave)
            flag_PEx = 'rx_angle_max_h'
            flag_PEy = 'rx_angle_max_v'
            flag_SKEWx = 'rx_skew_max_h'
            flag_SKEWy = 'rx_skew_max_v'
            result_peskew_max = [i for i in result_tmp if
                             flag_PEx in i or flag_PEy in i or flag_SKEWx in i or flag_SKEWy in i]
            print('Max:\n', result_peskew_max)
            flag_PEx = 'rx_angle_min_h'
            flag_PEy = 'rx_angle_min_v'
            flag_SKEWx = 'rx_skew_min_h'
            flag_SKEWy = 'rx_skew_min_v'
            result_peskew_min = [i for i in result_tmp if
                             flag_PEx in i or flag_PEy in i or flag_SKEWx in i or flag_SKEWy in i]
            print('Min:\n', result_peskew_min)
            if len(result_peskew_ave)!=4 or len(result_peskew_max)!=4 or len(result_peskew_min)!=4:
                print('Average or Max or Min返回值数量为{} {} {},不等于4个，请检查！'.format(len(result_peskew_ave),len(result_peskew_max),len(result_peskew_min)))
                return
            pe_x_ave   = re.split('[: ]', result_peskew_ave[0])[1]
            pe_y_ave   = re.split('[: ]', result_peskew_ave[1])[1]
            skew_x_ave = re.split('[: ]', result_peskew_ave[2])[1]
            skew_y_ave = re.split('[: ]', result_peskew_ave[3])[1]

            pe_x_max   = re.split('[: ]', result_peskew_max[0])[1]
            pe_y_max   = re.split('[: ]', result_peskew_max[1])[1]
            skew_x_max = re.split('[: ]', result_peskew_max[2])[1]
            skew_y_max = re.split('[: ]', result_peskew_max[3])[1]

            pe_x_min    = re.split('[: ]', result_peskew_min[0])[1]
            pe_y_min    = re.split('[: ]', result_peskew_min[1])[1]
            skew_x_min  = re.split('[: ]', result_peskew_min[2])[1]
            skew_y_min  = re.split('[: ]', result_peskew_min[3])[1]

            if abs(float(pe_x_max))>=abs(float(pe_x_min)):
                pe_x   = pe_x_max
            else:
                pe_x = pe_x_min
            if abs(float(pe_y_max))>=abs(float(pe_y_min)):
                pe_y   = pe_y_max
            else:
                pe_y = pe_y_min
            skew_x = skew_x_max
            skew_y = skew_y_max

            tt=config+[ch,pe_x_ave, pe_y_ave,pe_x_max, pe_y_max,pe_x_min, pe_y_min,pe_x, pe_y, skew_x, skew_y]
            # generate the report and print out the log file
            test_result.loc[p-1]=tt
            test_result.to_csv(report_file, index=False)
            mawin.CtrlB.timeout = 2
            self.sig_progress.emit(round(20 + (75 / len(mawin.channel) * p)))
        # close the laser output
        mawin.CtrlB.write(b'itla_wr 0 0x32 0x00\n\r');
        time.sleep(0.1)
        print(mawin.CtrlB.read_until(b'itla_0_write'))
        print('LO laser closed...')

        # os.system("explorer "+report_judge)
        self.sig_print.emit('测试完成!')
        self.sig_staColor.emit('green')
        self.sig_but.emit('开始')
        self.sig_status.emit('测试完成!')
        self.sig_progress.emit(100)
        print('测试完成')
        # Copy golden CSV data to folder under golden sample
        if mawin.test_type == '金样':
            golden_path = os.path.join(os.path.split(os.path.split(report_path3)[0])[0], 'Monitor_data')
            if not os.path.exists(golden_path): os.mkdir(golden_path)
            shutil.copy(report_file, golden_path)
        ##Write the data into report model and open the report after finished
        wb = xw.Book(mawin.report_model.replace('test_report.xlsx', 'test_report_PeSkew.xlsx'))
        worksht = wb.sheets(1)
        worksht.activate()
        worksht.range((1, 2)).value = test_result.iloc[0, 0]
        worksht.range((2, 2)).value = test_result.iloc[0, 1]
        worksht.range((3, 2)).value = test_result.iloc[0, 2]
        worksht.range((4, 2)).value = test_result.iloc[0, 3]
        worksht.range((5, 2)).value = test_result.iloc[0, 4]
        # write CLPD,FPGA,MCU information here into the test report
        # testBrdClp,ctlBrdClp,ctlBrdModVer,ctlBrdFPGAver,MCUver
        worksht.range((2, 4)).value = testBrdClp
        worksht.range((3, 4)).value = ctlBrdClp
        worksht.range((4, 4)).value = ctlBrdModVer
        worksht.range((5, 4)).value = ctlBrdFPGAver
        worksht.range((6, 4)).value = MCUver
        worksht.range((1, 4)).value = mawin.sw

        worksht.range((7, 2)).options(index=False, header=False, transpose=True).value = test_result.iloc[:, 5:]
        mawin.finalResult = worksht.range((6, 2)).value
        wb.sheets(2).activate()
        wb.save(report_judge)
        wb.close()
        mawin.CtrlB.close()
        mawin.JinLi.close()
        # copy the local data to network folder
        if not network_path == False:
            shutil.copytree(report_path3, network_path)
        wb = xw.Book(report_judge)
        wb.sheets(2).activate()

    def VOA_calibration(self):
        '''
        #VOA calibration with FS400 control board
        :return:NA
        '''
        # Get and judge the SN format
        sn = str(mawin.lineEdit.text()).strip()
        if not gf.SN_check(sn):
            # mawin.test_status.setText('SN输入有误，请检查SN！')
            self.sig_status.emit('SN输入有误，请检查SN！')
            self.sig_staColor.emit('red')
            self.sig_but.emit('开始')
            return

        # # Add the folder named as SN_Date to store all the test data
        timestamp = gf.get_timestamp(1)
        report_path3, network_path = mawin.create_report_folders(sn, timestamp)
        if network_path == False:
            pass
        shutil.copy(mawin.config_file, report_path3)
        sys.stdout = gf.Logger(path=report_path3)
        #print(sn)
        self.sig_status.emit('测试进行中...')
        self.sig_staColor.emit('yellow')
        self.sig_clear.emit()
        self.sig_print.emit(sn)
        self.sig_progress.emit(5)

        #***------设备连接---------***
        if not cb.open_board_VOAcal(mawin, mawin.ctrl_port):
            self.sig_status.emit('请检查控制板串口！')
            self.sig_but.emit('开始')
            self.sig_staColor.emit('red')
            self.sig_progress.emit(0)
            return

        # First ignore VOA 1 judgement and revise VOA setting as 0
        comlist = ['set_ignore_fs400 1',
                   'fs400_eep_arg_read 2',
                   'dbg_set_eep_voa_para 0 -170 4280',
                   'fs400_eep_arg_save 2',
                   'set_ignore_fs400 0',
                   ]
        cb.board_set_CmdList(mawin, comlist)

        # firstly judge if ITLA is CPP
        cpp_itla=True
        mawin.CtrlB.write(b'show_itla 0\n\r');time.sleep(0.5)
        re_itla=mawin.CtrlB.read_until(b'run_state').decode('utf-8').split('\n')
        print(re_itla)
        try:
            re_itla1=[i.split(' ')[-1] for i in re_itla if 'g_itla_model' in i][0]
        except Exception as e:
            print('ITLA model get failed:\n',e)
            return
        print(re_itla1)
        if not '0x3' in re_itla1:
            print('ITLA is C band')
            cpp_itla=False
        else:
            print('ITLA is C ++')
            cpp_itla = True
        # config ITLA,
        # 标定分C和C++
        # C标定波长为：191.3125THz~间隔300G~  追加一个波长196.0375
        # C+标定波长为:190.7125THz~间隔300G~  追加一个波长196.6375
        comlist=[]
        fre=list
        if cpp_itla:
            comlist = ['task_ctrl 4 0',
                       'itla_write 0 0x32 0',
                       'itla_write 0 0x34 0x2ee',
                       'itla_write 0 0x35 0xbe',
                       'itla_write 0 0x36 0x1bd5',
                       'itla_write 0 0x30 1',
                       'itla_write 0 0x32 8']
            mawin.channel=[i for i in range(1,81,4)]+[80]
            fre=np.interp(mawin.channel,range(1,81),mawin.wl_cpp)
        else:
            comlist = ['task_ctrl 4 0',
                       'itla_write 0 0x32 0',
                       'itla_write 0 0x34 0x2ee',
                       'itla_write 0 0x35 0xbf',
                       'itla_write 0 0x36 0xc35',
                       'itla_write 0 0x30 1',
                       'itla_write 0 0x32 8']
            mawin.channel = [i for i in range(9 , 73,4)] + [72]
            fre = np.interp(mawin.channel, range(9 , 73), mawin.wl_c)

        cb.board_set_CmdList(mawin,comlist)
        print('Laser opened!!!')
        #Close DSP
        mawin.CtrlB.write(b'close_dsp_before_voa_scan\n\r')
        print('Wait for 60s to close DSP and wait for temperature to be stable...')
        #time.sleep(60)
        for i in range(60,0,-5):
            print('Please wait,time left:{}s'.format(str(i)))
            time.sleep(5)
        #Wait for temperature to be stable
        tem_list=[3.0]*10
        count=0
        while max(tem_list)>0.15:
            count+=1
            print('The {} time get temperature, target is all <0.15:'.format(count))
            time.sleep(5)
            mawin.CtrlB.write(b'show_abc_adjust_speed_info\n\r')
            time.sleep(0.1)
            s_tem=mawin.CtrlB.read_until(b'low').decode('utf-8')
            print(s_tem)
            tem_list=[float(i[i.index(']')+2:i.index(']')+6]) for i in s_tem.split('\n') if 'g_dtemp_recored' in i]
            print(tem_list)
            if len(tem_list)<10:
                tem_list = [3.0] * 10
            if count==60:
                print('Not found after 2 mins, stop the test...please check!')
                return

        # connectivity test
        self.sig_progress.emit(10)
        self.sig_print.emit('检查pin脚连接中...')
        # New method to check the connectivity and check TIA/DRV is ALU or IDT
        con = 0
        con = cb.test_connectivity_VoaCalibration(mawin, sn)
        if con == 1:
            self.sig_print.emit('Driver is ALU')
            mawin.device_type = 'ALU'
            print('检测到ALU器件,将执行ALU器件相关的测试...')
        elif con == 2:
            self.sig_print.emit('Driver is IDT')
            mawin.device_type = 'IDT'
            print('检测到IDT器件,将执行IDT器件相关的测试...')
        elif con == 3:
            self.sig_print.emit('请检查光路或压接，连接性检查失败！！！')
            self.sig_staColor.emit('blue')
            self.sig_but.emit('开始')
            self.sig_status.emit('请检查器件连接!')
            # self.sig_stoptest.emit()
            self.sig_progress.emit(0)
            return

        # Read the CLPD,FPGA,MCU information
        testBrdClp = 'NA'
        ctlBrdClp = 'NA'
        ctlBrdModVer = 'NA'
        ctlBrdFPGAver = 'NA'
        MCUver = 'NA'
        # testBrdClp, ctlBrdClp, ctlBrdModVer, ctlBrdFPGAver, MCUver = gf.read_MCUandFPGA(mawin)

        '''
        *****Start the whole test steps*****
        '''
        self.sig_progress.emit(20)
        self.sig_print.emit('开始VOA Calibration...')
        ##define the data frame to store the test data
        test_result = DataFrame(columns=('SN', 'RX_BW_DESK', 'TEMP', 'DATE', 'TIME', '400G_CH','Temperature',
                                         'VOA_XI','VOA_XQ', 'VOA_YI', 'VOA_YQ'))

        # create list to store the result
        tt = []
        config = [sn, mawin.desk, mawin.temp[0]] + timestamp.split('_')
        report_name = sn + '_VOA_Calibration_raw_' + timestamp + '.csv'
        report_judgename = sn + '_VOA_Calibration_Report_' + timestamp + '.xlsx'
        report_file = os.path.join(report_path3, report_name)
        report_judge = os.path.join(report_path3, report_judgename)

        for i in range(len(mawin.channel)):
            # if i<10:
            #     continue
            ch = str(mawin.channel[i])
            p = i + 1
            self.sig_print.emit("CH%s 测试开始...\n" % (ch))
            voa_result = [0.0] * 4
            cmClose = b'itla_write 0 0x32 0\n\r'
            cmOpen =b'itla_write 0 0x32 8\n\r'
            if cpp_itla:
                ch_toset=int(ch)
                cm1='itla_write 0 0x30 {}\n\r'.format(str(hex(int(ch)))).encode('utf-8')
            else:
                ch_toset = int(ch)-8
                cm1 = 'itla_write 0 0x30 {}\n\r'.format(str(hex(int(ch) - 8))).encode('utf-8')

            #***-------Switch LO wavelength-------***
            # mawin.CtrlB.write(cmClose);
            # time.sleep(0.1)
            # print(mawin.CtrlB.read_until(b'itla_0_write'))
            # print('ITLA closed...')
            #need to switch wavelength twice to ensure, otherwise will not response
            mawin.CtrlB.write(cm1);
            time.sleep(1)
            print(mawin.CtrlB.read_until(b'itla_0_write'))
            mawin.CtrlB.write(cm1);
            time.sleep(1)
            print(mawin.CtrlB.read_until(b'itla_0_write'))
            # Confirm the wavelength is right
            mawin.CtrlB.flushInput()
            time.sleep(0.1)
            mawin.CtrlB.flushOutput()
            time.sleep(0.1)
            mawin.CtrlB.write(b'itla_read 0 0x30\n\r')
            re_wl = mawin.CtrlB.read_until(b'shell').decode('utf-8')
            print(re_wl)
            while not 'read_itla' in re_wl:
                time.sleep(1)
                mawin.CtrlB.write(b'itla_read 0 0x30\n\r')
                re_wl = mawin.CtrlB.read_until(b'shell').decode('utf-8')
                print(re_wl)
            fre_get = int(re_wl.split('\n')[2].strip().replace('\r', ''), 16)
            cou1 = 0
            while not fre_get == ch_toset:
                print('The {} time set wavelength...'.format(cou1 + 1))
                mawin.CtrlB.write(b'itla_read 0 0x30\n\r')
                re_wl = mawin.CtrlB.read_until(b'shell').decode('utf-8')
                print(re_wl)
                while not 'read_itla' in re_wl:
                    time.sleep(1)
                    mawin.CtrlB.write(b'itla_read 0 0x30\n\r')
                    re_wl = mawin.read_until(b'shell').decode('utf-8')
                    print(re_wl)
                fre_get = int(re_wl.split('\n')[2].strip().replace('\r', ''), 16)
                cou1 += 1
                if cou1 > 3:
                    print('Wavelength get error after 4 times retry: ', re_wl)
                    return
            # mawin.CtrlB.timeout=120
            # mawin.CtrlB.write(b'show_itla 0\n\r');
            # time.sleep(0.1)
            # re_wl = mawin.CtrlB.read_all().decode('utf-8')
            # while not 'g_itla_freqh' in re_wl and not 'g_itla_freql' in re_wl:
            #     mawin.CtrlB.write(b'show_itla 0\n\r');
            #     time.sleep(0.1)
            #     re_wl = mawin.CtrlB.read_all().decode('utf-8')
            # print(re_wl)
            # re_wl=re_wl.split('\n')
            # # try:
            # re_wl1 = [i.split(' ')[-1] for i in re_wl if 'g_itla_freqh' in i or 'g_itla_freql' in i]
            #     # count=0
            #     # while not len(re_wl1)==2:
            #     #     mawin.CtrlB.write(b'show_itla 0\n\r');
            #     #     time.sleep(0.2)
            #     #     re_wl = mawin.CtrlB.read_until(b'run_state').decode('utf-8').split('\n')
            #     #     print(re_wl)
            #     #     re_wl1 = [i.split(' ')[-1] for i in re_wl if 'g_itla_fcfreqh' in i or 'g_itla_fcfreql' in i]
            #     #     count+=1
            #     #     print('Count:',count)
            #     #     if count>600:
            #     #         print('Wavelength get error after 120s: ',re_wl1)
            #     #         return
            # # except Exception as e:
            # #     print('ITLA model get failed:\n', e)
            # #     return
            # sta1=re_wl1[0].index('(')
            # sto1=re_wl1[0].index(')')
            # sta2 = re_wl1[1].index('(')
            # sto2 = re_wl1[1].index(')')
            # fre_ge1=re_wl1[0][sta1+1:sto1]
            # fre_get2 = re_wl1[1][sta2 + 1:sto2]
            # fre_get=fre_ge1+'.'+fre_get2
            # # if cpp_itla:
            # cou=0
            # while not fre_get==str(fre[i]):
            #     print('Wavelength incorrect, retry for the {} times...'.format(str(cou+1)))
            #     mawin.CtrlB.write(cm1);
            #     time.sleep(1)
            #     print(mawin.CtrlB.read_until(b'itla_0_write'))
            #     cou+=1
            #     mawin.CtrlB.write(b'show_itla 0\n\r');
            #     time.sleep(0.1)
            #     re_wl = mawin.CtrlB.read_all().decode('utf-8')
            #     while not 'g_itla_freqh' in re_wl and not 'g_itla_freql' in re_wl:
            #         mawin.CtrlB.write(b'show_itla 0\n\r');
            #         time.sleep(0.1)
            #         re_wl = mawin.CtrlB.read_all().decode('utf-8')
            #     print(re_wl)
            #     re_wl = re_wl.split('\n')
            #     # try:
            #     re_wl1 = [i.split(' ')[-1] for i in re_wl if 'g_itla_freqh' in i or 'g_itla_freql' in i]
            #     sta1 = re_wl1[0].index('(')
            #     sto1 = re_wl1[0].index(')')
            #     sta2 = re_wl1[1].index('(')
            #     sto2 = re_wl1[1].index(')')
            #     fre_ge1 = re_wl1[0][sta1 + 1:sto1]
            #     fre_get2 = re_wl1[1][sta2 + 1:sto2]
            #     fre_get = fre_ge1 + '.' + fre_get2
            #     if cou==3:
            #         print('ITLA wl not set correct, target is {}, read is {}...'.format(str(fre[i]),fre_get))
            #         return
            # mawin.CtrlB.write(cmOpen);
            # time.sleep(0.1)
            # print(mawin.CtrlB.read_until(b'itla_0_write'))
            # print('ITLA opened...')
            # else:
            #     if not fre_get==str(fre[i]):
            #         print('ITLA wl not set correct, read is {}, target is {}...'.format(str(fre[i]),fre_get))
            #***---------Read temperature--------***
            mawin.CtrlB.write(b'show_temperature\n\r');
            time.sleep(0.5)
            re_temp_all=mawin.CtrlB.read_all().decode('utf-8')
            while not 'module_temperature' in re_temp_all:
                mawin.CtrlB.write(b'show_temperature\n\r');
                time.sleep(0.5)
                re_temp_all = mawin.CtrlB.read_all().decode('utf-8')
            # re_temp = mawin.CtrlB.read_until(b'module_temperature').decode('utf-8').split('\n')
            re_temp = re_temp_all.split('\n')
            try:
                re_temp1 = [i[i.index('current')+8:i.index('current')+16] for i in re_temp if 'fs400_sensor_temp' in i ][0]
                print(re_temp1)
                tempera=float(re_temp1[0:4])*(10**int(re_temp1[-2:]))
                print('FS400 senser temperature:{}℃'.format(tempera))
            except Exception as e:
                print('Temperature get failed:\n', e)
                return

            # ***---------Read voa--------***
            mawin.CtrlB.write(b'voa_scan_no_dsp 0\n\r');
            time.sleep(0.2)
            mawin.CtrlB.timeout=120
            str_wait=''
            count=0
            ch_jump=False
            while not 'voa scan verify success-' in str_wait:
                str_wait = mawin.CtrlB.readline().decode('utf-8')
                print(str_wait)
                time.sleep(0.1)
                if 'voa scan verify fail' in str_wait:
                    print('VOA value got failed, try again...')
                    mawin.CtrlB.write(b'voa_scan_no_dsp 0\n\r');
                    time.sleep(0.1)
                    count1 = 0
                    while not 'voa scan verify success-' in str_wait:
                        str_wait = mawin.CtrlB.readline().decode('utf-8')
                        print(str_wait)
                        time.sleep(0.1)
                        if 'voa scan verify fail' in str_wait:
                            print('VOA value got failed again，will jump over this channel...')
                            ch_jump=True
                            break
                        count1 += 1
                        if count > 3000:
                            print('Retest 300s not found value valid, end process...')
                            return
                count+=1
                if count>3000:
                    print('300s not found value valid, end process...')
                    return
            if not ch_jump:
                #No channel to jump, normal to go!
                #print('jump this channel as no valid voa result get')
                voa_read=str_wait#.decode('utf-8')
                print(voa_read)
                mawin.CtrlB.timeout = 3
                try:
                    voa_read1=voa_read[voa_read.index('-')+1:].split('\n')[0].strip().split(',')
                    print(voa_read1)
                    if not len(voa_read1)==8:
                        print('VOA value got failed, not 8!')
                        return
                    voa_result=[voa_read1[0],voa_read1[2],voa_read1[4],voa_read1[6]]
                    voa_result=[float(t) for t in voa_result]
                    print(voa_result)
                except Exception as e:
                    print('Temperature get failed:\n', e)
                    return
            else:
                print('jump this channel as no valid voa result get')
                voa_result = [1000, 0, 1000, 0]
            tt=config+[ch]+[tempera]+voa_result
            test_result.loc[p-1]=tt
            test_result.to_csv(report_file, index=False)
            mawin.CtrlB.timeout = 2
            self.sig_progress.emit(round(20 + (75 / len(mawin.channel) * p)))
        # close the laser output
        mawin.CtrlB.write(cmClose)
        time.sleep(0.1)
        print(mawin.CtrlB.read_until(b'itla_0_write'))
        print('LO laser closed...')

        #calculate VOA data
        print('Calculate VOA data...')
        report_name = sn + '_VOA_Calibration_fit_' + timestamp + '.csv'
        report_file = os.path.join(report_path3, report_name)

        if cpp_itla:
            voa_XI, voa_XQ, voa_YI, voa_YQ=gf.VOA_cal_data(test_result,fre, mawin.wl_cpp,mawin.voa_wl)
        else:
            voa_XI, voa_XQ, voa_YI, voa_YQ = gf.VOA_cal_data(test_result, fre, mawin.wl_c, mawin.voa_wl)
        T_median = test_result.iloc[:,6].median()
        if T_median < 40:
            #shift_dire = 'UP'
            N = math.ceil((40 - T_median) / 4.4)
            T=int(round(T_median+N*4.4,2)*100)
        elif T_median > 45:
            #shift_dire = 'DOWN'
            N = math.ceil((T_median - 45) / 4.4)
            T = int(round(T_median - N * 4.4, 2) * 100)
        else:
            #shift_dire = 'NONE'
            N=0
            T = int(round(T_median + N * 4.4, 2) * 100)

        #write data into the fit voa data,'VOA_XI','VOA_XQ', 'VOA_YI', 'VOA_YQ'
        test_result_fit = DataFrame(columns=('SN', 'RX_BW_DESK', 'TEMP', 'DATE', 'TIME', '400G_Fre',
                                         'VOA_XI', 'VOA_XQ', 'VOA_YI', 'VOA_YQ'))
        for i in range(len(mawin.voa_wl)):
            da_add=config+[mawin.voa_wl[i]]+[voa_XI[i]]+[voa_XQ[i]]+[voa_YI[i]]+[voa_YQ[i]]
            test_result_fit.loc[i]=da_add
        test_result_fit.to_csv(report_file, index=False)
        print('Save VOA final fit result done!')
        # in the end ignore VOA 1 judgement and revise VOA setting as 1, and pass temperature into EEPROM
        comlist = ['fs400_eep_arg_read 2',
                    'dbg_set_eep_voa_para 1 -170 {}'.format(str(T)),
                   'fs400_eep_arg_save 2']
        cb.board_set_CmdList(mawin, comlist)

        #write VOA result to EEPROM
        #need to power up first? or no need?
        gf.write_VOAcal2EEP(mawin,test_result_fit)
        # os.system("explorer "+report_judge)
        self.sig_print.emit('测试完成!')
        self.sig_staColor.emit('green')
        self.sig_but.emit('开始')
        self.sig_status.emit('测试完成!')
        self.sig_progress.emit(100)
        print('测试完成')
        ##Write the data into report model and open the report after finished
        wb = xw.Book(mawin.report_model.replace('test_report.xlsx', 'test_report_VOAcal.xlsx'))
        worksht = wb.sheets(1)
        worksht.activate()
        worksht.range((1, 2)).value = test_result.iloc[0, 0]
        worksht.range((2, 2)).value = test_result.iloc[0, 1]
        worksht.range((3, 2)).value = test_result.iloc[0, 2]
        worksht.range((4, 2)).value = test_result.iloc[0, 3]
        worksht.range((5, 2)).value = test_result.iloc[0, 4]
        # write CLPD,FPGA,MCU information here into the test report
        # testBrdClp,ctlBrdClp,ctlBrdModVer,ctlBrdFPGAver,MCUver
        worksht.range((2, 4)).value = testBrdClp
        worksht.range((3, 4)).value = ctlBrdClp
        worksht.range((4, 4)).value = ctlBrdModVer
        worksht.range((5, 4)).value = ctlBrdFPGAver
        worksht.range((6, 4)).value = MCUver
        worksht.range((1, 4)).value = mawin.sw

        worksht.range((7, 2)).options(index=False, header=False, transpose=True).value = test_result_fit.iloc[:, 5:]
        mawin.finalResult = worksht.range((6, 2)).value
        wb.sheets(2).activate()
        wb.save(report_judge)
        mawin.CtrlB.write(b'wrmsa 0xb010 0x4000\n\r')
        # copy the local data to network folder
        # if not network_path == False:
        #     shutil.copytree(report_path3, network_path)

    def VOA_calibration_New(self):
        '''
        0915_auto_detect and write Drv vcc out
        #VOA calibration with FS400 control board
        :return:NA
        '''
        # Get and judge the SN format
        sn = str(mawin.lineEdit.text()).strip()
        if not gf.SN_check(sn):
            self.sig_status.emit('SN输入有误，请检查SN！')
            self.sig_staColor.emit('red')
            self.sig_but.emit('开始')
            return

        # Add the folder named as SN_Date to store all the test data
        timestamp = gf.get_timestamp(1)
        report_path3, network_path = mawin.create_report_folders(sn, timestamp)
        if network_path == False:
            pass
        shutil.copy(mawin.config_file, report_path3)
        sys.stdout = gf.Logger(path=report_path3)

        self.sig_status.emit('测试进行中...')
        self.sig_staColor.emit('yellow')
        self.sig_clear.emit()
        self.sig_print.emit(sn)
        self.sig_progress.emit(5)

        #***------设备连接---------***
        if not cb.open_board_VOAcal(mawin, mawin.ctrl_port):
            self.sig_status.emit('请检查控制板串口！')
            self.sig_but.emit('开始')
            self.sig_staColor.emit('red')
            self.sig_progress.emit(0)
            return

        # First ignore VOA 1 judgement and revise VOA setting as 0
        comlist = ['set_ignore_fs400 1',
                   'fs400_eep_arg_read 2',
                   'dbg_set_eep_voa_para 0 -170 4280',
                   'fs400_eep_arg_save 2',
                   'set_ignore_fs400 0',
                   ]
        cb.board_set_CmdList(mawin, comlist)

        #auto get drv out setting to make drv vpd as 2.2v~2.3v
        cb.drv_VCC_set_EEPROM(mawin)
        print('DRV VCC out write done!')

        mawin.CtrlB.flushInput()
        mawin.CtrlB.flushOutput()
        #Judge if ITLA is CPP
        cpp_itla=True
        mawin.CtrlB.timeout=10
        mawin.CtrlB.write(b'show_itla 0\n\r')
        re_itla=mawin.CtrlB.read_until(b'run_state').decode('utf-8')
        coun=0
        while not 'g_itla_model' in re_itla:
            time.sleep(1)
            mawin.CtrlB.write(b'show_itla 0\n\r')
            re_itla = mawin.CtrlB.read_until(b'run_state').decode('utf-8')
            coun+=1
            if coun>2:
                print('3 times read error, please check!')
                return
        s=re_itla.split('\n')
        print(re_itla)
        try:
            re_itla1=[i.split(' ')[-1] for i in s if 'g_itla_model' in i][0]
        except Exception as e:
            print('ITLA model get failed:\n',e)
            return
        print(re_itla1)
        if not '0x3' in re_itla1:
            print('ITLA is C band')
            cpp_itla=False
        else:
            print('ITLA is C ++')
            cpp_itla = True
        # config ITLA,
        # 标定分C和C++
        # C标定波长为：191.3125THz~间隔300G~  追加一个波长196.0375
        # C+标定波长为:190.7125THz~间隔300G~  追加一个波长196.6375
        comlist=[]
        fre=list
        if cpp_itla:
            comlist = ['task_ctrl 4 0',
                       'itla_write 0 0x32 0',
                       'itla_write 0 0x34 0x2ee',
                       'itla_write 0 0x35 0xbe',
                       'itla_write 0 0x36 0x1bd5',
                       'itla_write 0 0x30 1',
                       'itla_write 0 0x32 8']
            mawin.channel=[i for i in range(1,81,4)]+[80]
            fre=np.interp(mawin.channel,range(1,81),mawin.wl_cpp)
        else:
            comlist = ['task_ctrl 4 0',
                       'itla_write 0 0x32 0',
                       'itla_write 0 0x34 0x2ee',
                       'itla_write 0 0x35 0xbf',
                       'itla_write 0 0x36 0xc35',
                       'itla_write 0 0x30 1',
                       'itla_write 0 0x32 8']
            mawin.channel = [i for i in range(9 , 73,4)] + [72]
            fre = np.interp(mawin.channel, range(9 , 73), mawin.wl_c)
        mawin.update_config()
        cb.board_set_CmdList(mawin,comlist)
        print('Laser opened!!!')
        #Close DSP
        mawin.CtrlB.write(b'close_dsp_before_voa_scan\n\r')
        print('Wait for 60s to close DSP and wait for temperature to be stable...')
        for i in range(60,0,-5):
            print('Please wait,time left:{}s'.format(str(i)))
            time.sleep(5)
        #Wait for temperature to be stable
        tem_list=[3.0]*10
        count=0
        while max(tem_list)>0.15:
            count+=1
            print('The {} time get temperature, target is all <0.15:'.format(count))
            time.sleep(5)
            mawin.CtrlB.write(b'show_abc_adjust_speed_info\n\r')
            time.sleep(0.1)
            s_tem=mawin.CtrlB.read_until(b'low').decode('utf-8')
            print(s_tem)
            tem_list=[float(i[i.index(']')+2:i.index(']')+6]) for i in s_tem.split('\n') if 'g_dtemp_recored' in i]
            print(tem_list)
            if len(tem_list)<10:
                tem_list = [3.0] * 10
            if count==60:
                print('Not found after 2 mins, stop the test...please check!')
                return

        #connectivity test
        self.sig_progress.emit(10)
        self.sig_print.emit('检查pin脚连接中...')
        # New method to check the connectivity and check TIA/DRV is ALU or IDT
        con = 0
        con = cb.test_connectivity_VoaCalibration(mawin, sn)
        if con == 1:
            self.sig_print.emit('Driver is ALU')
            mawin.device_type = 'ALU'
            print('检测到ALU器件,将执行ALU器件相关的测试...')
        elif con == 2:
            self.sig_print.emit('Driver is IDT')
            mawin.device_type = 'IDT'
            print('检测到IDT器件,将执行IDT器件相关的测试...')
        elif con == 3:
            self.sig_print.emit('请检查光路或压接，连接性检查失败！！！')
            self.sig_staColor.emit('blue')
            self.sig_but.emit('开始')
            self.sig_status.emit('请检查器件连接!')
            # self.sig_stoptest.emit()
            self.sig_progress.emit(0)
            return


        # Read the CLPD,FPGA,MCU information
        testBrdClp = 'NA'
        ctlBrdClp = 'NA'
        ctlBrdModVer = 'NA'
        ctlBrdFPGAver = 'NA'
        MCUver = 'NA'
        # testBrdClp, ctlBrdClp, ctlBrdModVer, ctlBrdFPGAver, MCUver = gf.read_MCUandFPGA(mawin)

        '''
        *****Start the whole test steps*****
        '''
        self.sig_progress.emit(20)
        self.sig_print.emit('开始VOA Calibration...')
        #define the data frame to store the test data
        test_result = DataFrame(columns=('SN', 'RX_BW_DESK', 'TEMP', 'DATE', 'TIME', '400G_CH','Temperature',
                                         'VOA_XI','VOA_XQ', 'VOA_YI', 'VOA_YQ'))

        # create list to store the result
        tt = []
        config = [sn, mawin.desk, mawin.temp[0]] + timestamp.split('_')
        report_name = sn + '_VOA_Calibration_raw_' + timestamp + '.csv'
        report_judgename = sn + '_VOA_Calibration_Report_' + timestamp + '.xlsx'
        report_file = os.path.join(report_path3, report_name)
        report_judge = os.path.join(report_path3, report_judgename)

        for i in range(len(mawin.channel)):
            # if i<10:
            #     continue
            ch = str(mawin.channel[i])
            p = i + 1
            self.sig_print.emit("CH%s 测试开始...\n" % (ch))
            voa_result = [0.0] * 4
            cmClose = b'itla_write 0 0x32 0\n\r'
            cmOpen =b'itla_write 0 0x32 8\n\r'
            if cpp_itla:
                ch_toset=int(ch)
                cm1='itla_write 0 0x30 {}\n\r'.format(str(hex(int(ch)))).encode('utf-8')
            else:
                ch_toset = int(ch)-8
                cm1 = 'itla_write 0 0x30 {}\n\r'.format(str(hex(int(ch) - 8))).encode('utf-8')

            #***-------Switch LO wavelength-------***
            # mawin.CtrlB.write(cmClose);
            # time.sleep(0.1)
            # print(mawin.CtrlB.read_until(b'itla_0_write'))
            # print('ITLA closed...')
            #need to switch wavelength twice to ensure, otherwise will not response
            mawin.CtrlB.write(cm1)
            print(mawin.CtrlB.read_until(b'itla_0_write'))
            time.sleep(1)
            # mawin.CtrlB.write(cm1)
            # print(mawin.CtrlB.read_until(b'itla_0_write'))
            # Confirm the wavelength is right
            mawin.CtrlB.flushInput()
            mawin.CtrlB.flushOutput()
            time.sleep(0.1)
            mawin.CtrlB.write(b'itla_read 0 0x30\n\r')
            re_wl = mawin.CtrlB.read_until(b'shell').decode('utf-8')
            print(re_wl)
            while not 'read_itla' in re_wl:
                time.sleep(1)
                mawin.CtrlB.write(b'itla_read 0 0x30\n\r')
                re_wl = mawin.CtrlB.read_until(b'shell').decode('utf-8')
                print(re_wl)
            fre_get = int(re_wl.split('\n')[2].strip().replace('\r', ''), 16)
            cou1 = 0
            while not fre_get == ch_toset:
                print('The {} time set wavelength...'.format(cou1 + 1))
                mawin.CtrlB.write(cm1)
                print(mawin.CtrlB.read_until(b'itla_0_write'))
                time.sleep(1)
                mawin.CtrlB.write(b'itla_read 0 0x30\n\r')
                re_wl = mawin.CtrlB.read_until(b'shell').decode('utf-8')
                print(re_wl)
                while not 'read_itla' in re_wl:
                    time.sleep(1)
                    mawin.CtrlB.write(b'itla_read 0 0x30\n\r')
                    re_wl = mawin.read_until(b'shell').decode('utf-8')
                    print(re_wl)
                fre_get = int(re_wl.split('\n')[2].strip().replace('\r', ''), 16)
                cou1 += 1
                if cou1 > 3:
                    print('Wavelength get error after 4 times retry: ', re_wl)
                    return
            print('ITLA wavelength switched to ch{}...'.format(ch))

            #***---------Read temperature--------***
            mawin.CtrlB.write(b'show_temperature\n\r');
            time.sleep(0.5)
            re_temp_all=mawin.CtrlB.read_all().decode('utf-8')
            while not 'module_temperature' in re_temp_all:
                mawin.CtrlB.write(b'show_temperature\n\r');
                time.sleep(0.5)
                re_temp_all = mawin.CtrlB.read_all().decode('utf-8')
            # re_temp = mawin.CtrlB.read_until(b'module_temperature').decode('utf-8').split('\n')
            re_temp = re_temp_all.split('\n')
            try:
                re_temp1 = [i[i.index('current')+8:i.index('current')+16] for i in re_temp if 'fs400_sensor_temp' in i ][0]
                print(re_temp1)
                tempera=float(re_temp1[0:4])*(10**int(re_temp1[-2:]))
                print('FS400 senser temperature:{}℃'.format(tempera))
            except Exception as e:
                print('Temperature get failed:\n', e)
                return

            # ***---------Read voa--------***
            mawin.CtrlB.write(b'voa_scan_no_dsp 0\n\r');
            time.sleep(0.2)
            mawin.CtrlB.timeout=120
            str_wait=''
            count=0
            ch_jump=False
            while not 'voa scan verify success-' in str_wait:
                str_wait = mawin.CtrlB.readline().decode('utf-8')
                print(str_wait)
                time.sleep(0.1)
                if 'voa scan verify fail' in str_wait:
                    print('VOA value got failed, try again...')
                    mawin.CtrlB.write(b'voa_scan_no_dsp 0\n\r');
                    time.sleep(0.1)
                    count1 = 0
                    while not 'voa scan verify success-' in str_wait:
                        str_wait = mawin.CtrlB.readline().decode('utf-8')
                        print(str_wait)
                        time.sleep(0.1)
                        if 'voa scan verify fail' in str_wait:
                            print('VOA value got failed again，will jump over this channel...')
                            ch_jump=True
                            break
                        count1 += 1
                        if count > 3000:
                            print('Retest 300s not found value valid, end process...')
                            return
                count+=1
                if count>3000:
                    print('300s not found value valid, end process...')
                    return
            if not ch_jump:
                #No channel to jump, normal to go!
                #print('jump this channel as no valid voa result get')
                voa_read=str_wait#.decode('utf-8')
                print(voa_read)
                mawin.CtrlB.timeout = 3
                try:
                    voa_read1=voa_read[voa_read.index('-')+1:].split('\n')[0].strip().split(',')
                    print(voa_read1)
                    if not len(voa_read1)==8:
                        print('VOA value got failed, not 8!')
                        return
                    voa_result=[voa_read1[0],voa_read1[2],voa_read1[4],voa_read1[6]]
                    voa_result=[float(t) for t in voa_result]
                    print(voa_result)
                except Exception as e:
                    print('Temperature get failed:\n', e)
                    return
            else:
                print('jump this channel as no valid voa result get')
                voa_result = [1000, 0, 1000, 0]
            tt=config+[ch]+[tempera]+voa_result
            test_result.loc[p-1]=tt
            test_result.to_csv(report_file, index=False)
            mawin.CtrlB.timeout = 2
            self.sig_progress.emit(round(20 + (75 / len(mawin.channel) * p)))
        # close the laser output
        mawin.CtrlB.write(cmClose)
        print(mawin.CtrlB.read_until(b'itla_0_write'))
        print('LO laser closed...')

        #calculate VOA data
        print('Calculate VOA data...')
        report_name = sn + '_VOA_Calibration_fit_' + timestamp + '.csv'
        report_file = os.path.join(report_path3, report_name)
        if cpp_itla:
            voa_XI, voa_XQ, voa_YI, voa_YQ=gf.VOA_cal_data(test_result,fre, mawin.wl_cpp,mawin.voa_wl)
        else:
            voa_XI, voa_XQ, voa_YI, voa_YQ = gf.VOA_cal_data(test_result, fre, mawin.wl_c, mawin.voa_wl)
        T_median = test_result.iloc[:,6].median()
        if T_median < 40:
            #shift_dire = 'UP'
            N = math.ceil((40 - T_median) / 4.4)
            T=int(round(T_median+N*4.4,2)*100)
        elif T_median > 45:
            #shift_dire = 'DOWN'
            N = math.ceil((T_median - 45) / 4.4)
            T = int(round(T_median - N * 4.4, 2) * 100)
        else:
            #shift_dire = 'NONE'
            N=0
            T = int(round(T_median + N * 4.4, 2) * 100)

        #T=int(round(test_result.iloc[:,6].median(),2)*100)
        #write data into the fit voa data,'VOA_XI','VOA_XQ', 'VOA_YI', 'VOA_YQ'
        test_result_fit = DataFrame(columns=('SN', 'RX_BW_DESK', 'TEMP', 'DATE', 'TIME', '400G_Fre',
                                         'VOA_XI', 'VOA_XQ', 'VOA_YI', 'VOA_YQ'))
        for i in range(len(mawin.voa_wl)):
            da_add=config+[mawin.voa_wl[i]]+[voa_XI[i]]+[voa_XQ[i]]+[voa_YI[i]]+[voa_YQ[i]]
            test_result_fit.loc[i]=da_add
        test_result_fit.to_csv(report_file, index=False)
        print('Save VOA final fit result done!')
        # in the end ignore VOA 1 judgement and revise VOA setting as 1, and pass temperature into EEPROM
        comlist = ['fs400_eep_arg_read 2',
                    'dbg_set_eep_voa_para 1 -170 {}'.format(str(T)),
                   'fs400_eep_arg_save 2']
        cb.board_set_CmdList(mawin, comlist)

        #write VOA result to EEPROM
        #need to power up first? or no need?
        gf.write_VOAcal2EEP(mawin,test_result_fit)
        # os.system("explorer "+report_judge)
        self.sig_print.emit('测试完成!')
        self.sig_staColor.emit('green')
        self.sig_but.emit('开始')
        self.sig_status.emit('测试完成!')
        self.sig_progress.emit(100)
        print('测试完成')
        ##Write the data into report model and open the report after finished
        wb = xw.Book(mawin.report_model.replace('test_report.xlsx', 'test_report_VOAcal.xlsx'))
        worksht = wb.sheets(1)
        worksht.activate()
        worksht.range((1, 2)).value = test_result.iloc[0, 0]
        worksht.range((2, 2)).value = test_result.iloc[0, 1]
        worksht.range((3, 2)).value = test_result.iloc[0, 2]
        worksht.range((4, 2)).value = test_result.iloc[0, 3]
        worksht.range((5, 2)).value = test_result.iloc[0, 4]
        # write CLPD,FPGA,MCU information here into the test report
        # testBrdClp,ctlBrdClp,ctlBrdModVer,ctlBrdFPGAver,MCUver
        worksht.range((2, 4)).value = testBrdClp
        worksht.range((3, 4)).value = ctlBrdClp
        worksht.range((4, 4)).value = ctlBrdModVer
        worksht.range((5, 4)).value = ctlBrdFPGAver
        worksht.range((6, 4)).value = MCUver
        worksht.range((1, 4)).value = mawin.sw

        worksht.range((7, 2)).options(index=False, header=False, transpose=True).value = test_result_fit.iloc[:, 5:]
        mawin.finalResult = worksht.range((6, 2)).value
        wb.sheets(2).activate()
        wb.save(report_judge)
        wb.close()
        mawin.CtrlB.write(b'wrmsa 0xb010 0x4000\n\r')
        # copy the local data to network folder
        if not network_path == False:
            shutil.copytree(report_path3, network_path)
        wb = xw.Book(report_judge)
        wb.sheets(2).activate()

    def ICR_test_PDbalance(self):
        '''
        #ICR test main process
        :return:NA
        '''
        # firstly rename the file of ICR test config of 'Board up,Drv up,Drv down'
        # Because no need to open ITLA, and need to power down PDs
        mawin.board_up = os.path.join(mawin.config_path, 'Setup_brdup_CtrlboardA001_20220113_ICR.txt')
        # mawin.drv_up       = os.path.join(mawin.config_path,'Setup_driverup_Ctrlboard56017837A002_20211203.txt')
        # mawin.drv_down     = os.path.join(mawin.config_path,'Setup_driverdown_CtrlboardA001_20210820.txt')
        # Get and judge the SN format
        sn = str(mawin.lineEdit.text()).strip()
        if not gf.SN_check(sn):
            # mawin.test_status.setText('SN输入有误，请检查SN！')
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
        if not cb.open_board(mawin, mawin.ctrl_port):
            self.sig_status.emit('请检查控制板串口！')
            self.sig_but.emit('开始')
            self.sig_staColor.emit('red')
            self.sig_progress.emit(0)
            return
        #####Oscilloscope and ICR tester connection
        if not osc.open_OScope(mawin, mawin.OScope_port):
            self.sig_status.emit('请检查示波器端口！')
            self.sig_but.emit('开始')
            self.sig_staColor.emit('red')
            self.sig_progress.emit(0)
            return

        if not icr.open_ICRtf(mawin, mawin.ICRtf_port):
            self.sig_status.emit('请检查ICR测试平台串口！')
            self.sig_but.emit('开始')
            self.sig_staColor.emit('red')
            self.sig_progress.emit(0)
            return

        ###connectivity test
        self.sig_progress.emit(10)
        self.sig_print.emit('检查pin脚连接中...')
        # New method to check the connectivity and check TIA/DRV is ALU or IDT
        con = 0
        con = cb.test_connectivity_new_ICR(mawin, sn)
        if con == 1:
            self.sig_print.emit('Driver is ALU')
            mawin.device_type = 'ALU'
            print('检测到ALU器件,将执行ALU器件相关的测试...')
        elif con == 2:
            self.sig_print.emit('Driver is IDT')
            mawin.device_type = 'IDT'
            print('检测到IDT器件,将执行IDT器件相关的测试...')
        elif con == 3:
            self.sig_print.emit('请检查光路或压接，连接性检查失败！！！')
            self.sig_staColor.emit('blue')
            self.sig_but.emit('开始')
            self.sig_status.emit('请检查器件连接!')
            # self.sig_stoptest.emit()
            self.sig_progress.emit(0)
            return

        #This is for MPC test
        #res_sig_x = cur_x / (10 ** (pwr / 10)) / 8
        # mawin.CtrlB.write(b'switch_set 10 0\n');time.sleep(0.1)
        # mawin.CtrlB.write(b'cpld_spi_wr 0x2c 2150\n');time.sleep(0.1)
        # mawin.CtrlB.write(b'cpld_spi_wr 0x2f 2150\n');time.sleep(0.1)
        # icr.set_laser(mawin, 2, 15.4, -5, 1550.000)
        # mawin.ICRtf.write('OUTP3:CHAN2:POW 5')
        # mawin.MPC_Find_Position('PDX',True)
        # curX,curY=icr.TIA_getPDcurrent(mawin)[0:2]
        # mawin.MPC_Find_Position('PDY',True)
        # curX,curY=icr.TIA_getPDcurrent(mawin)[0:2]
        # mawin.MPC_Find_Position('DIFF',False)
        # curX,curY=icr.TIA_getPDcurrent(mawin)[0:2]
        # mawin.ICRtf.write('OUTP1:CHAN2:STATE OFF')

        # con = 0
        # # con=3
        # con = cb.test_connectivity(mawin)
        # if con == 1:
        #     self.sig_print.emit('请检查，Driver未正确连接')
        #     self.sig_staColor.emit('blue')
        #     self.sig_but.emit('开始')
        #     self.sig_status.emit('请检查Driver连接!')
        #     # self.sig_stoptest.emit()
        #     self.sig_progress.emit(0)
        #     return
        # elif con == 2:
        #     self.sig_print.emit('请检查，TIA未正确连接')
        #     self.sig_staColor.emit('blue')
        #     self.sig_but.emit('开始')
        #     self.sig_status.emit('请检查TIA连接!')
        #     # self.sig_stoptest.emit()
        #     self.sig_progress.emit(0)
        #     return
        # elif con == 0:
        #     self.sig_print.emit('请检查，无正确返回值，连接性检查失败！')
        #     self.sig_staColor.emit('blue')
        #     self.sig_but.emit('开始')
        #     self.sig_status.emit('请检查器件连接!')
        #     # self.sig_stoptest.emit()
        #     self.sig_progress.emit(0)
        #     return

        # Read the CLPD,FPGA,MCU information
        testBrdClp = 'NA'
        ctlBrdClp = 'NA'
        ctlBrdModVer = 'NA'
        ctlBrdFPGAver = 'NA'
        MCUver = 'NA'
        testBrdClp, ctlBrdClp, ctlBrdModVer, ctlBrdFPGAver, MCUver = gf.read_MCUandFPGA(mawin)

        '''
        *****Start the whole test steps*****
        '''
        self.sig_progress.emit(15)
        self.sig_print.emit('器件连接成功, FS400初始化, 配置TIA及上电...')
        # TIA Power on and config
        if mawin.device_type=='IDT':
            gf.TIA_on(mawin)
            gf.TIA_config(mawin)
            print('IDT器件测试')
            self.sig_print.emit('TIA上电完成，设备初始化中...')
        elif mawin.device_type == 'ALU':
            mawin.drv_up = os.path.join(mawin.config_path, 'Setup_driverup_Ctrlboard56017837A002_20211203_ALU.txt')
            mawin.drv_down = os.path.join(mawin.config_path, 'Setup_driverdown_CtrlboardA001_20210820_ALU.txt')
            gf.TIA_on(mawin)
            gf.TIA_config_ALU(mawin,mawin.test_res)
            print('ALU器件测试')

        # initiate Oscilloscope and switch to visa connection
        osc.switch_to_DSO(mawin, mawin.OScope_port)
        print('DSO object created')
        osc.switch_to_VISA(mawin, mawin.DSO_port)
        if mawin.test_pe or mawin.test_bw:
            osc.init_OSc(mawin)
            self.sig_print.emit('示波器初始化成功...')
        self.sig_print.emit('开启并设置Laser波长...')
        icr.set_laser(mawin, 1, 15.4, 10)
        icr.set_laser(mawin, 2, 15.4, -5, 1550.000)
        self.sig_progress.emit(20)
        # Test group by wavelength
        self.sig_print.emit('配置完成!开始ICR测试...')
        ###if normal ITLA(Not C++ ITLA) then wavelength minus 8
        # result_tmp=[]
        ##define the data frame to store the test data
        test_result = DataFrame(columns=('SN', 'RX_BW_DESK', 'TEMP', 'DATE', 'TIME', '400G_CH',
                                         'RX_BW_XI', 'RX_BW_XQ', 'RX_BW_YI', 'RX_BW_YQ', 'RX_PE_X',
                                         'RX_PE_Y', 'RX_X_Skew', 'RX_Y_Skew', 'RX_1PD_RES_LO_X', 'RX_1PD_RES_LO_Y',
                                         'RX_1PD_RES_SIG_X', 'RX_1PD_RES_SIG_Y', 'RX_PER_X', 'RX_PER_Y', 'I_Dark_X',
                                         'I_Dark_Y'))
        # Test data storage
        # timestamp = gf.get_timestamp(1)
        # if not os.path.exists(mawin.report_path):
        #     os.mkdir(mawin.report_path)
        # report_path1 = os.path.join(mawin.report_path, mawin.test_flag)  # create the child folder to store data
        # if not os.path.exists(report_path1):
        #     os.mkdir(report_path1)
        # report_path2 = os.path.join(report_path1, mawin.test_type)  # create the child folder to store data
        # if not os.path.exists(report_path2):
        #     os.mkdir(report_path2)
        # # Add the folder named as SN_Date to store all the test data
        # report_path3 = os.path.join(report_path2, sn + '_' + timestamp)  # create the child folder to store data
        # if not os.path.exists(report_path3):
        #     os.mkdir(report_path3)

        # # Add the folder named as SN_Date to store all the test data
        timestamp = gf.get_timestamp(1)
        report_path3, network_path = mawin.create_report_folders(sn, timestamp)
        if network_path == False:
            pass
        # create list to store the result
        tt = []
        config = [sn, mawin.desk, mawin.temp[0]] + timestamp.split('_')
        # report_name=sn+'_'+timestamp+'.csv'
        report_name = sn + '_ICR_' + timestamp + '.csv'  # BOB改了csv报告名字
        report_judgename = sn + '_ICR_Report_' + timestamp + '.xlsx'
        report_file = os.path.join(report_path3, report_name)
        report_judge = os.path.join(report_path3, report_judgename)

        for i in range(len(mawin.channel)):
            ch = str(mawin.channel[i])
            p = i + 1
            self.sig_print.emit("CH%s 测试开始...\n" % (ch))
            # ICR Test start and set the wavelength
            icr.set_wavelength(mawin, 2, ch, 0)
            icr.set_wavelength(mawin, 1, ch, 1)
            mawin.ICRtf.write('OUTP3:CHAN1:POW 6')  # modify the value from 10dBm to 9dBm
            mawin.ICRtf.write('OUTP3:CHAN2:POW -6')
            mawin.ICRtf.write('OUTP1:CHAN1:STATE ON')
            mawin.ICRtf.write('OUTP1:CHAN2:STATE ON')
            pe_x, pe_y, skew_x, skew_y = [0.0] * 4
            XIBW3dB, XQBW3dB, YIBW3dB, YQBW3dB = [0.0] * 4
            res_sig_x, res_sig_y, per_x, per_y, dark_x, dark_y, res_lo_x, res_lo_y = [0.0] * 8
            if mawin.test_pe:
                self.sig_print.emit('Phase 和 Skew 测试中...')
                # phase error test
                lamb = icr.cal_wavelength(ch, 0)
                m = 3  # 平均次数
                # icr.ICR_manualPol_ScopeBalance_Judge(mawin)
                # #icr.ICR_scamblePol(mawin)
                # # query the data of amtiplitude
                # A1 = mawin.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')
                # A3 = mawin.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')
                # while not icr.isfloat(A1) or not icr.isfloat(A3):
                #     A1 = mawin.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')
                #     A3 = mawin.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')
                # while float(A1) / float(A3) > 1.08 or float(A1) / float(A3) < 0.92:
                #     time.sleep(0.2)
                #     icr.ICR_manualPol_ScopeBalance_Judge(mawin)
                #     #icr.ICR_scamblePol(mawin)
                #     A1 = mawin.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')
                #     A3 = mawin.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')
                #     while not icr.isfloat(A1) or not icr.isfloat(A3):
                #         A1 = mawin.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')
                #         A3 = mawin.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')
                # Scan RF start
                data = mawin.scan_RF_ManualPol(ch)
                # data = mawin.scan_RF(ch)
                # save the result
                # pe_skew='{}_CH{}_PeSkew_{}'.format(sn,ch,timestamp)
                # pd.DataFrame(data).to_csv(pe_skew)
                pe_x, pe_y, skew_x, skew_y, new_pahseX, new_pahseY = peSkew.deal_with_data(data)
                # save raw and fitted s21 data
                raw_phase_report = os.path.join(report_path3, '{}_CH{}_phase_RawData_{}.csv'.format(sn, ch, timestamp))
                col = ['Phase error', 'Skew'] + [str(i) + 'GHz' for i in range(1, 11)]
                pha = pd.DataFrame([[pe_x, skew_x] + new_pahseX, [pe_y, skew_y] + new_pahseY],
                                   index=('phase_X', 'phase_Y'), columns=col)
                pha.to_csv(raw_phase_report)
                # draw the curve in the main thread
                self.sig_plot.emit(pha, report_path3, ch, sn, timestamp, 'PE')
                # #work left here to draw the curve
                # fre_1=[i for i in range(0,11)]
                # y1=np.array(fre_1)*skew_x+pe_x+90
                # y2=np.array(fre_1)*skew_y+pe_y+90
                # #X plot fig
                # fig_X=os.path.join(report_path3,'{}_CH{}_XIXQ_PeSkew_{}.png'.format(sn,ch,timestamp))
                # fig, ax = plt.subplots()  # Create a figure containing a single axes.
                # ax.plot(fre_1,y1,new_pahseX,'*')  # Plot some data on the axes.
                # ax.set_title('XI-XQ Phase Error and Skew CH{}'.format(ch))
                # ax.set_xlabel('Fre(GHz)')
                # ax.set_ylabel('Angle [°]')
                # fig.savefig(fig_X)
                # #Y plot fig
                # fig_Y=os.path.join(report_path3,'{}_CH{}_YIYQ_PeSkew_{}.png'.format(sn,ch,timestamp))
                # fig1, ax1 = plt.subplots()  # Create a figure containing a single axes.
                # ax1.plot(fre_1,y2,new_pahseY,'*')  # Plot some data on the axes.
                # ax1.set_title('YI-YQ Phase Error and Skew CH{}'.format(ch))
                # ax1.set_xlabel('Fre(GHz)')
                # ax1.set_ylabel('Angle [°]')
                # fig1.savefig(fig_Y)

                # fig.show()
                # fig1.show()

            # Work left to do here to calculate SKEW PE
            if mawin.test_bw:
                # Rx BW test
                self.sig_print.emit('Rx BW 测试中...')
                icr.set_wavelength(mawin, 2, ch, 0)
                icr.set_wavelength(mawin, 1, ch, 1)
                time.sleep(0.1)
                error = True
                eg = 1  # 表示在某一个波长处测试的次数
                while error:
                    #icr.ICR_scamblePol(mawin)
                    # icr.ICR_manualPol_ScopeBalance_Judge(mawin)
                    # # query the data of amtiplitude
                    # A1 = mawin.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')
                    # A3 = mawin.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')
                    # while not icr.isfloat(A1) or not icr.isfloat(A3):
                    #     time.sleep(0.2)
                    #     A1 = mawin.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')
                    #     A3 = mawin.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')
                    # while float(A1) / float(A3) > 1.15 or float(A1) / float(A3) < 0.85:
                    #     time.sleep(0.2)
                    #     #icr.ICR_scamblePol(mawin)
                    #     icr.ICR_manualPol_ScopeBalance_Judge(mawin)
                    #     A1 = mawin.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')
                    #     A3 = mawin.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')
                    #     while not icr.isfloat(A1) or not icr.isfloat(A3):
                    #         A1 = mawin.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')
                    #         A3 = mawin.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')
                    # # scan S21
                    # print('A1:{}\nA3:{}'.format(A1, A3))
                    #data1 = mawin.scan_S21(ch, 1, 40, 1)
                    data1 = mawin.scan_S21_ManualPol(ch, 1, 40, 1)
                    # config the frequency
                    numfre = 40
                    da = [[0.0] * 4] * numfre
                    fre = [i * 1e9 for i in range(1, 41)]
                    # 归一化数据Y
                    raw = pd.DataFrame(data1)
                    for i in range(4):
                        raw.iloc[:, i] = 20 * np.log10(raw.iloc[:, i] / raw.iloc[0, i])
                    # 线损
                    linelo = pd.read_csv(mawin.Rx_line_loss, encoding='gb2312')
                    fre_loss = linelo.iloc[:-1, 0]
                    loss = linelo.iloc[:-1, 1]
                    Yi = np.interp(fre, fre_loss, loss)  # 插值求得对应频率的线损
                    # 减去线损
                    Ysmo = raw.sub(Yi, axis=0)
                    # save raw and fitted s21 data
                    raw_s21_report = os.path.join(report_path3, '{}_CH{}_RxBw_RawData_{}.csv'.format(sn, ch, timestamp))
                    fit_s21_report = os.path.join(report_path3, '{}_CH{}_RxBw_FitData_{}.csv'.format(sn, ch, timestamp))
                    colBW = ['XI', 'XQ', 'YI', 'YQ']
                    indBW = [i for i in range(1, 41)]
                    raw.columns = colBW
                    raw.index = indBW
                    raw.to_csv(raw_s21_report)  # ,index=False)
                    ###This point needs to be verified, whether the smooth data is equal to matlab gaussian method
                    for i in range(4):
                        tem = smooth.smooth(Ysmo.iloc[:, i], 3)[1:-1]
                        tem = tem - tem[0]  # 归一化数据
                        Ysmo.iloc[:, i] = tem
                    Ysmo.columns = colBW
                    Ysmo.index = indBW
                    Ysmo.to_csv(fit_s21_report)  # ,index=False)
                    # draw the curves
                    self.sig_plot.emit(Ysmo, report_path3, ch, sn, timestamp, 'BW')
                    # #work left here to draw the curve
                    # x_bw=[i for i in range(1,41)]
                    # fig_bw, ax_bw = plt.subplots()  # Create a figure containing a single axes.
                    # ax_bw.plot(x_bw,Ysmo.iloc[:,0],label='XI')
                    # ax_bw.plot(x_bw,Ysmo.iloc[:,1],label='XQ')
                    # ax_bw.plot(x_bw,Ysmo.iloc[:,2],label='YI')
                    # ax_bw.plot(x_bw,Ysmo.iloc[:,3],label='YQ')# Plot some data on the axes.
                    # ax_bw.set_title('Rx Band Width CH{}'.format(ch))
                    # ax_bw.set_xlabel('Fre(GHz)')
                    # ax_bw.set_ylabel('Loss(dB)')
                    # ax_bw.legend()
                    # #fig_bw.show()
                    # fig_bwPic=os.path.join(report_path3,'{}_CH{}_RxBW_{}.png'.format(sn,ch,timestamp))
                    # fig_bw.savefig(fig_bwPic)

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

                    # calculate the result, base on Ysmo is a dataframe type
                    for i in range(4):
                        aa = [x for x in range(Ysmo.shape[0]) if Ysmo.iloc[x, i] < -3]
                        if len(aa) == 0:
                            BW3dB = 40
                        else:
                            BW3dB = min(aa) + 1  # /1e9 index plus 1 as Freq
                        # each channel
                        if i == 0:
                            XIBW3dB = BW3dB
                            if min(Ysmo.iloc[:, i]) < -20.0:
                                XIBW3dB = 0.0
                        elif i == 1:
                            XQBW3dB = BW3dB
                            if min(Ysmo.iloc[:, i]) < -20.0:
                                XQBW3dB = 0.0
                        elif i == 2:
                            YIBW3dB = BW3dB
                            if min(Ysmo.iloc[:, i]) < -20.0:
                                YIBW3dB = 0.0
                        elif i == 3:
                            YQBW3dB = BW3dB
                            if min(Ysmo.iloc[:, i]) < -20.0:
                                YQBW3dB = 0.0

                    YsmoMean = Ysmo.mean(axis=1)
                    aa = [x for x in range(YsmoMean.shape[0]) if YsmoMean[x + 1] < -3]
                    if len(aa) == 0:
                        MeanBW3dB = 40
                    else:
                        MeanBW3dB = min(aa) + 1  # /1e9
                        if min(YsmoMean) < -20:
                            MeanBW3dB = 0

                    # Need to check whether to perform the judgement of >28G Hz and each channel > mean bw by 3GHz
                    # if not meet the requirement then perform the retest
                    # left to write the result
                    error = False  # Not check BW and retest
            # TIA amplitude test not performed
            # scan_Amp:to test the TIA output amplitude
            # left blank here
            if mawin.test_res:
                # Responsivity test
                # self.sig_print.emit('响应度测试中...')
                # DAC开启，PD上电
                # mawin.CtrlB.write(b'switch_set 10 0\n');time.sleep(0.1)
                # mawin.CtrlB.write(b'cpld_spi_wr 0x2c 2700\n');time.sleep(0.1)
                # mawin.CtrlB.write(b'cpld_spi_wr 0x2f 2700\n');time.sleep(0.1)
                # PD voltage set to 2150
                mawin.CtrlB.write(b'switch_set 10 0\n');
                time.sleep(0.1)
                mawin.CtrlB.write(b'cpld_spi_wr 0x2c 2150\n');
                time.sleep(0.1)
                mawin.CtrlB.write(b'cpld_spi_wr 0x2f 2150\n');
                time.sleep(0.1)

                self.sig_print.emit('Rx响应度测试中...')
                res_sig_x, res_sig_y, per_x, per_y, dark_x, dark_y, PD_currentX, PD_currentY,ind_markYmin,ind_markXmin= mawin.get_pd_resp_sig_New(ch)
                #res_sig_x, res_sig_y, per_x, per_y, dark_x, dark_y, PD_currentX, PD_currentY = mawin.get_pd_resp_sig(ch)
                # #For test
                # a=[]
                # for i in range(1500,2701,100):
                #     mawin.CtrlB.write(b'switch_set 10 0\n');time.sleep(0.1)
                #     mawin.CtrlB.write('cpld_spi_wr 0x2c {}\n'.format(str(i)).encode('utf-8'));time.sleep(0.5)
                #     mawin.CtrlB.write('cpld_spi_wr 0x2f {}\n'.format(str(i)).encode('utf-8'));time.sleep(0.5)
                #     self.sig_print.emit('LO响应度测试中...')
                #     res_lo_x,res_lo_y,curX,curY=mawin.get_pd_resp_lo(ch)
                #     a.append([i,curX,curY])
                # dfff=pd.DataFrame(a)
                # dfff.columns=['PD Voltage','PD_LO_currentX','PD_LO_currentY']
                # res_temp=os.path.join(report_path3,'{}_CH{}_temp_{}.csv'.format(sn,ch,timestamp))
                # dfff.to_csv(res_temp)
                # #PD 2150 test ten times
                # a=[]
                # for t in range(10):
                #     i=2150
                #     mawin.CtrlB.write(b'switch_set 10 0\n');time.sleep(0.1)
                #     mawin.CtrlB.write('cpld_spi_wr 0x2c {}\n'.format(str(i)).encode('utf-8'));time.sleep(0.5)
                #     mawin.CtrlB.write('cpld_spi_wr 0x2f {}\n'.format(str(i)).encode('utf-8'));time.sleep(0.5)
                #     self.sig_print.emit('LO响应度测试中...')
                #     res_lo_x,res_lo_y,curX,curY=mawin.get_pd_resp_lo(ch)
                #     a.append([t,curX,curY])
                # dfff=pd.DataFrame(a)
                # dfff.columns=['PD times','PD_LO_currentX','PD_LO_currentY']
                # res_temp=os.path.join(report_path3,'{}_CH{}_temp_10times_{}.csv'.format(sn,ch,timestamp))
                # dfff.to_csv(res_temp)

                # Draw the Rx PD current data point and save the image
                # Save the raw data
                res_raw = os.path.join(report_path3, '{}_CH{}_RxPDcurrentXY_{}.csv'.format(sn, ch, timestamp))
                raw_res = pd.DataFrame([PD_currentX, PD_currentY]).transpose()
                raw_res.columns = ['PD_currentX', 'PD_currentY']
                raw_res.to_csv(res_raw)
                # Draw the curve X Y phase
                self.sig_plot.emit(raw_res, report_path3, ch, sn, timestamp, 'RES',ind_markYmin,ind_markXmin)

                self.sig_print.emit('LO响应度测试中...')
                # res_lo_x,res_lo_y,curX,curY=mawin.get_pd_resp_lo(ch)

                # Bob 瞎搞 !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

                PD_current_X_LO = []
                PD_current_Y_LO = []
                for t in range(10):
                    res_lo_x, res_lo_y, curX, curY = mawin.get_pd_resp_lo(ch)
                    PD_current_X_LO.append(curX)
                    PD_current_Y_LO.append(curY)

                res_raw = os.path.join(report_path3, '{}_CH{}_RxPDcurrentXY_LO_{}.csv'.format(sn, ch, timestamp))
                raw_res = pd.DataFrame([PD_current_X_LO, PD_current_Y_LO]).transpose()
                raw_res.columns = ['PD_currentX_LO', 'PD_currentY_LO']
                raw_res.to_csv(res_raw)
                # self.sig_plot.emit(raw_res,report_path3,ch,sn,timestamp,'RES')
                # !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

                # x_px=[i for i in range(1,len(PD_currentX)+1)]
                # fig_px, ax_px = plt.subplots()  # Create a figure containing a single axes.
                # ax_px.plot(x_px,PD_currentX,label='Current X')
                # ax_px.plot(x_px,PD_currentY,label='Current Y')
                # ax_px.set_title('Rx PD current CH{}'.format(ch))
                # ax_px.set_xlabel('Points')
                # ax_px.set_ylabel('Current(mA)')
                # ax_px.legend()
                # #fig_px.show()
                # fig_pxPic=os.path.join(report_path3,'{}_CH{}_Rx_PDcurrentXY_{}.png'.format(sn,ch,timestamp))
                # fig_px.savefig(fig_pxPic)

            self.sig_progress.emit(round(20 + (75 / len(mawin.channel) * p)))
            tt=config+[ch, XIBW3dB, XQBW3dB, YIBW3dB, YQBW3dB, pe_x, pe_y, skew_x, skew_y, res_lo_x, res_lo_y,
                       res_sig_x, res_sig_y, per_x, per_y, dark_x, dark_y]
            # generate the report and print out the log file
            test_result.loc[p-1]=tt
            test_result.to_csv(report_file, index=False)
            # Close the figs
            self.sig_plotClose.emit()
        # close the laser output
        mawin.ICRtf.write('OUTP1:CHAN1:STATE OFF')
        mawin.ICRtf.write('OUTP1:CHAN2:STATE OFF')

        # # 'SN','TX_BW_DESK','TEMP','DATE','TIME','400G_CH'
        # # # timestamp=gf.get_timestamp(1)
        # config = [sn, mawin.desk, mawin.temp[0]] + timestamp.split('_')
        # # report_name=sn+'_'+timestamp+'.csv'
        # report_name = sn + '_ICR_' + timestamp + '.csv'  # BOB改了csv报告名字
        # report_judgename = sn + '_ICR_Report_' + timestamp + '.xlsx'
        # report_file = os.path.join(report_path3, report_name)
        # report_judge = os.path.join(report_path3, report_judgename)
        # print(report_file)
        # for i in tt:
        #     tt[tt.index(i)] = config + i
        # # Write the result into data frame
        # for i in range(len(mawin.channel)):
        #     test_result.loc[i] = tt[i]
        # # generate the report and print out the log file
        # test_result.to_csv(report_file, index=False)
        # print('Close the figs')
        # plt.close(fig)
        # plt.close(fig1)
        # plt.close(fig_bw)
        # plt.close(fig_px)

        # wb.close()
        # os.system("explorer "+report_judge)
        self.sig_print.emit('测试完成!')
        self.sig_staColor.emit('green')
        self.sig_but.emit('开始')
        self.sig_status.emit('测试完成!')
        self.sig_progress.emit(100)
        print('测试完成')
        ##Write the data into report model and open the report after finished
        wb = xw.Book(mawin.report_model.replace('test_report.xlsx', 'test_report_ICR.xlsx'))
        worksht = wb.sheets(1)
        worksht.activate()
        worksht.range((1, 2)).value = test_result.iloc[0, 0]
        worksht.range((2, 2)).value = test_result.iloc[0, 1]
        worksht.range((3, 2)).value = test_result.iloc[0, 2]
        worksht.range((4, 2)).value = test_result.iloc[0, 3]
        worksht.range((5, 2)).value = test_result.iloc[0, 4]
        # write CLPD,FPGA,MCU information here into the test report
        # testBrdClp,ctlBrdClp,ctlBrdModVer,ctlBrdFPGAver,MCUver
        worksht.range((2, 4)).value = testBrdClp
        worksht.range((3, 4)).value = ctlBrdClp
        worksht.range((4, 4)).value = ctlBrdModVer
        worksht.range((5, 4)).value = ctlBrdFPGAver
        worksht.range((6, 4)).value = MCUver
        worksht.range((1, 4)).value = mawin.sw

        worksht.range((7, 2)).options(index=False, header=False, transpose=True).value = test_result.iloc[:, 5:]
        mawin.finalResult = worksht.range((6, 2)).value
        wb.save(report_judge)
        # copy the local data to network folder
        if not network_path == False:
            shutil.copytree(report_path3, network_path)

    # plot test function
    def plot_test(self):
        '''
        This is to test the plot function outside of the main thread
        def plotTest(self,df,f,ch,sn,timetmp,typ):
        :return:
        '''
        f = r'C:\backup\Desktop\FS400\Test_Script\Test_report_TXDC\ICR\Normal\#215_TestAll_20220524_170351\#215_TestAll_CH13_RxBw_FitData_20220524_170351.csv'
        df = pd.read_csv(f).iloc[:,1:]
        self.sig_plot.emit(df, r'C:\backup', '13', 'test', '2022', 'BW')
        time.sleep(60)
        self.sig_plotClose.emit()
        pass

if __name__ == "__main__":
    app=QApplication(sys.argv)
    mawin = main_test()
    mawin.show()
    sys.exit(app.exec_())