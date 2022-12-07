# -*- coding: utf-8 -*-
# date:3/17/2022
# Update on 4/6/2022 V1.4
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
#others
import numpy as np
import operator
import time
import configparser
from ThreadCalient import DebugGui
import xlwings as xw
import serial
import smooth

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
        #global log
        global channel
        global temp
        #global a
        self.PM=object
        self.CtrlB=object
        self.ZVA=object
        self.GOCA=object
        self.RFSW=object
        self.log=''
        self.sn=''
        #self.a=0
        self.start=time.time()
        self.end=time.time()
        self.finalResult='Pass'
        self.test_flag='DC'
        self.test_type='Normal'
        self.gocflag=False #Aim to set timeout for GOCA, tigger the timeout

        self.config_path=os.path.join(sys.path[0],'Configuration')
        self.report_path=os.path.join(sys.path[0],'Test_report_TXDC')
        self.board_up=os.path.join(self.config_path,'Setup_brdup_CtrlboardA001_20220113.txt')
        self.drv_up=os.path.join(self.config_path,'Setup_driverup_Ctrlboard56017837A002_20211203.txt')
        self.drv_down=os.path.join(self.config_path,'Setup_driverdown_CtrlboardA001_20210820.txt')
        self.config_file=os.path.join(self.config_path,'config.ini')
        self.report_model=os.path.join(self.config_path,'test_report.xlsx')
        self.report_judge=os.path.join(self.report_path,'test_report.xlsx')

        #Read the config file
        conf=configparser.ConfigParser()
        conf.read(self.config_file)
        sections=conf.sections()
        self.sw       = conf.get(sections[0],'SoftwareVersion')
        self.desk     = conf.get(sections[0],'TestDesk')
        self.temp     = conf.get(sections[0],'Temperature').split(',')
        self.channel  = conf.get(sections[0],'Channel').split(',')
        self.ITLA_pwr = conf.get(sections[0],'ITLApower').split(',')
        self.PM_ch    = conf.get(sections[0],'PowerMeterChannel')
        self.ctrl_port= conf.get(sections[0],'ControlBoardPort')
        self.pow_port = conf.get(sections[0],'PowerMeterPort')
        self.RFSW_port = conf.get(sections[0],'RFswPort')
        self.ZVA_port = conf.get(sections[0],'ZVAport')
        self.GOCA_port = conf.get(sections[0],'GOCAport')
        self.RFcal_path = conf.get(sections[0],'RFcalPath')
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
        #self.workthread.sig_stoptest.connect(self.test_end)
        self.workthread.sig_goca.connect(self.goca_timer)
        self.timer.timeout.connect(self.periodly_check)
        self.start_test_butt_3.clicked.connect(self.debug_mode)
        self.start_test_butt_4.clicked.connect(self.openConfig)
        self.start_test_butt.clicked.connect(self.Start_test)
        #set up timer to monitor the GOCA equipment visit timeout
        self.timer1=QTimer(self)
        self.timer1.timeout.connect(self.periodly_checkEQ)
        #update the configuration on the label
        self.text_show='测试软件版本：{}\n测试机台号：{}\n测试温度:{}\n测试通道:{}\n光功率计通道:{}\n控制板串口号:{}\n功率计串口号：{}\nRF开关串口号：{}\nZVA端口号：{}\nGOCA端口号：{}'.format(self.sw,self.desk
                                                                                         ,self.temp,self.channel,self.PM_ch,
                                                                                           self.ctrl_port,self.pow_port,self.RFSW_port,self.ZVA_port,self.GOCA_port
                                                                                           )
        self.label_2.setText(self.text_show)

    def update_config(self):
        self.text_show='测试软件版本：{}\n测试机台号：{}\n测试温度:{}\n测试通道:{}\n光功率计通道:{}\n控制板串口号:{}\n功率计串口号：{}\nRF开关串口号：{}\nZVA端口号：{}\nGOCA端口号：{}'.format(self.sw,self.desk
                                                                                                                                          ,self.temp,self.channel,self.PM_ch,
                                                                                                                                          self.ctrl_port,self.pow_port,self.RFSW_port,self.ZVA_port,self.GOCA_port
                                                                                                                                          )
        self.label_2.setText(self.text_show)

    def openConfig(self):
        os.system("explorer.exe %s" % self.config_file)

    def debug_mode(self):
        print('Debug mode started...')
        debug_gui.pushButton.clicked.connect(mawin.passval_debug)
        debug_gui.pushButton_2.clicked.connect(mawin.setdrv_down)
        debug_gui.show()

    def setdrv_down(self):
        if not cb.open_board(mawin,mawin.ctrl_port):
            mawin.show()
            mawin.start_test_butt.setText('开始')
            mawin.test_status.setText('请检查控制板串口！')
            gf.status_color(mawin.test_status,'red')
            return
        cb.board_set(self,self.drv_down)
        self.print_out_status('手动下电完成！')
        cb.close_board(mawin)

    #pass the value from the debug window to the main window
    def passval_debug(self):
        #config the channel
        if debug_gui.radioButton.isChecked():
            self.channel=['13']
        elif debug_gui.radioButton_2.isChecked():
            self.channel=['39']
        elif debug_gui.radioButton_3.isChecked():
            self.channel=['65']
        elif debug_gui.radioButton_4.isChecked():
            self.channel=['13','39','65']

        #config the temperature
        if debug_gui.radioButton_7.isChecked():
            self.temp=['-5']
        elif debug_gui.radioButton_6.isChecked():
            self.temp=['25']
        elif debug_gui.radioButton_8.isChecked():
            self.temp=['75']
        elif debug_gui.radioButton_5.isChecked():
            self.temp=['25','-5','75']

        debug_gui.close()
        print(self.channel,self.temp)
        self.update_config()
        pass


    #test QTimer when workthread is not running recover GUI widgets and flags
    def periodly_check(self):
        if not self.workthread.isRunning():
            self.start_test_butt_3.setEnabled(True)
            self.start_test_butt_4.setEnabled(True)
            self.Test_item.setEnabled(True)
            self.Test_type.setEnabled(True)
            if self.finalResult=='Pass':
                gf.status_color(self.test_status,'green')
                self.test_status.setText("Pass")
            else:
                self.test_status.setText("Fail")
                gf.status_color(self.test_status,'red')
                self.finalResult='Pass'
            self.progressBar.setValue(100)
            self.test_end()
            #self.start_test_butt.setText('开始')
            self.end=time.time()
            utime=str(time.strftime("%H:%M:%S", time.gmtime(self.end-self.start)))
            self.print_out_status('测试完成，用时：'+utime)
            print('测试完成，用时：',utime)
            self.timer.stop()

    #test QTimer when timeout and no response from the equipment GOCA
    def periodly_checkEQ(self,timeout=20):
        t1=time.time()
        while not self.gocflag:
            t2=time.time()-t1
            if t2>timeout:
                self.workthread.terminate()
                self.timer1.stop()
                print('GOCA初始化失败...请检查GOCA Romote模式是否开启！')
                self.test_status.setText('GOCA初始化失败...请检查GOCA Romote模式是否开启！')
                gf.status_color(self.test_status,'red')
                return
        self.test_status.setText('GOCA初始化成功...')
        self.gocflag=False
        self.timer1.stop()

    #start the test
    @pyqtSlot()
    def Start_test(self):
        self.test_flag=self.Test_item.currentText()
        self.test_type=self.Test_type.currentText()
        if self.start_test_butt.text()=='开始':
            self.start_test_butt.setText('停止')
            self.start =time.time()
            self.timer.start(1000)
            self.start_test_butt_3.setEnabled(False)
            self.start_test_butt_4.setEnabled(False)
            self.Test_item.setEnabled(False)
            self.Test_type.setEnabled(False)
            self.workthread.start()
            print (self.workthread.isRunning())
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
        self.gocflag=False
        #Driver 下电，关闭串口连接
        self.start_test_butt.setText('开始')
        if serial.Serial.isOpen(self.CtrlB):
            cb.board_set(self,self.drv_down)
            self.print_out_status('Driver下电完成！')
            cb.close_board(self)
            self.print_out_status('测试板串口关闭完成！')
        if self.test_flag=='DC':
            pwm.close_PM(self)
        else:
            rfsw.close_RFSW(self)
            goca.close_GOCA(self)
            zva.close_ZVA(self)

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


class WorkThread(QThread):
    sig_progress=pyqtSignal(int)
    sig_status=pyqtSignal(str)
    sig_staColor=pyqtSignal(str)
    sig_print=pyqtSignal(str)
    sig_but=pyqtSignal(str)
    sig_clear=pyqtSignal()
    sig_goca=pyqtSignal(int)
    #sig_stoptest=pyqtSignal()

    def __init__(self):
        super(WorkThread, self).__init__()

    def run(self):
        try:
            if mawin.test_flag=="DC":
                self.DC_test()
                #self.test_unit()
            elif mawin.test_flag=="TxBW":
                self.TxBW_test()
            else:
                pass
        except Exception as e:
            print(e)
            return


    #DC test main process
    def DC_test(self):
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
        for i in range(len(mawin.channel)):
            self.sig_print.emit("CH%s 测试开始...\n"%(str(mawin.channel[i])))
            data=np.zeros(6)
            data=cb.get_ER(mawin,str(int(mawin.channel[i])-8),noipwr[0],noipwr[1],1,1)
            if data==False or data==[]:
                self.sig_print.emit('ER获取失败，任务中止...')
                break
            self.sig_print.emit('ER获取成功...')
            print('max,min,abc,ER,Tvpi:\n',data)
            max=data[0]
            min=data[1]
            #get power meter reading of X-max Y-max power
            abc_ok=max
            abc_tmp=np.zeros(6)
            while not operator.eq(abc_ok,abc_tmp):
                cb.set_abc(mawin,abc_ok)
                abc_tmp=cb.get_abc(mawin)
            pwr=pwm.read_PM(mawin,mawin.PM_ch)
            #get power meter reading of X-max Y-min power
            abc_ok=max[0:3]+min[3:6]
            abc_tmp=np.zeros(6)
            while not operator.eq(abc_ok,abc_tmp):
                cb.set_abc(mawin,abc_ok)
                abc_tmp=cb.get_abc(mawin)
            pwr_x=pwm.read_PM(mawin,mawin.PM_ch)
            #get power meter reading of X-min Y-max power
            abc_ok=min[0:3]+max[3:6]
            abc_tmp=np.zeros(6)
            while not operator.eq(abc_ok,abc_tmp):
                cb.set_abc(mawin,abc_ok)
                abc_tmp=cb.get_abc(mawin)
            pwr_y=pwm.read_PM(mawin,mawin.PM_ch)
            #set abc to max
            abc_ok=max
            abc_tmp=np.zeros(6)
            while not operator.eq(abc_ok,abc_tmp):
                cb.set_abc(mawin,abc_ok)
                abc_tmp=cb.get_abc(mawin)
            self.sig_print.emit('PDL和IL获取成功...')
            CH=str(mawin.channel[i])
            PDL=pwr_x-pwr_y
            IL=pwr-float(mawin.ITLA_pwr[i])
            ABC=data[2]
            ER=data[3]
            Tvpi=data[4]
            Max=max
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
        report_path1=os.path.join(mawin.report_path,mawin.test_flag)#create the child folder to store data
        if not os.path.exists(report_path1):
            os.mkdir(report_path1)
        report_path2=os.path.join(report_path1,mawin.test_type)#create the child folder to store data
        if not os.path.exists(report_path2):
            os.mkdir(report_path2)
        timestamp=gf.get_timestamp(1)
        config=[sn,mawin.desk,mawin.temp[0]]+timestamp.split('_')
        report_name=sn+'_'+timestamp+'.csv'
        report_judgename=sn+'_'+timestamp+'.xlsx'
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
        #wb.close()
        #os.system("explorer "+report_judge)
        self.sig_print.emit('测试完成!')
        self.sig_staColor.emit('green')
        self.sig_but.emit('开始')
        self.sig_status.emit('测试完成!')
        self.sig_progress.emit(100)
        print('测试完成')
        ##Write the data into report model and open the report after finished
        wb=xw.Book(mawin.report_model)
        worksht=wb.sheets(1)
        worksht.activate()
        worksht.range((1,2)).value=test_result.iloc[0,0]
        worksht.range((2,2)).value=test_result.iloc[0,1]
        worksht.range((3,2)).value=test_result.iloc[0,2]
        worksht.range((4,2)).value=test_result.iloc[0,3]
        worksht.range((5,2)).value=test_result.iloc[0,4]
        worksht.range((8,2)).options(index=False,header=False,transpose=True).value=test_result.iloc[:,6:32]
        mawin.finalResult=worksht.range((6,2)).value
        # if mawin.finalResult=='Pass':
        #     mawin.test_status.setText("Pass")
        # else:
        #     mawin.test_status.setText("Fail")
        #     gf.status_color(mawin.test_status,'red')
        #     mawin.finalResult='Pass'
        wb.save(report_judge)
        #self.sig_stoptest.emit()

    #test function
    def test_unit(self):
        self.sig_goca.emit(20)
        mawin.time_consuming()#goca.init_GOCA(mawin)
        mawin.time_consuming()
        mawin.time_consuming()
        mawin.time_consuming()


    #Tx BW test main process
    def TxBW_test(self):
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

        #get calibration file path
        rfcal=mawin.RFcal_path+''
        self.sig_print.emit('Driver上电完成，设备初始化中...')
        #initiate GOCA and ZVA
        self.sig_goca.emit(20)
        #mawin.timer1.start(1000)
        mawin.gocflag = goca.init_GOCA(mawin)
        #     self.sig_print.emit('GOCA初始化失败...')
        #     return
        # else:
        #     self.sig_print.emit('GOCA初始化完成...')
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
        linecal=pd.read_csv(os.path.join(mawin.config_path,'S21_lineIL.csv')).iloc[:-1,:]
        pdcal=pd.read_csv(os.path.join(mawin.config_path,'S21_PD.csv')).iloc[:-1,:]
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
            abc=data[2]
            # #get power meter reading of X-max Y-max power
            abc_ok=abc
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
                zva.recall_cal(mawin,mawin.RFcal_path,sw_ch)
                rfsw.switch_RFchannel(mawin,sw_ch)
                point=cb.get_quad(mawin,bias_ch,500,4000,15)
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
                        abc_ok=abc
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
        report_judgename=sn+'_'+timestamp+'.xlsx'
        report_file=os.path.join(report_path2,report_name)
        report_judge=os.path.join(mawin.report_path,report_judgename)
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
        # wb=xw.Book(mawin.report_model)
        # worksht=wb.sheets(1)
        # worksht.activate()
        # worksht.range((1,2)).value=test_result.iloc[0,0]
        # worksht.range((2,2)).value=test_result.iloc[0,1]
        # worksht.range((3,2)).value=test_result.iloc[0,2]
        # worksht.range((4,2)).value=test_result.iloc[0,3]
        # worksht.range((5,2)).value=test_result.iloc[0,4]
        # worksht.range((8,2)).options(index=False,header=False,transpose=True).value=test_result.iloc[:,6:32]
        # mawin.finalResult=worksht.range((6,2)).value
        # wb.save(report_judge)
        # self.sig_stoptest.emit()

if __name__ == "__main__":
    app=QApplication(sys.argv)
    mawin = main_test()
    debug_gui=DebugGui()
    mawin.show()
    sys.exit(app.exec_())