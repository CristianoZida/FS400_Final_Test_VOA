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
import CtrlBoard
from Main_win_GUI import Ui_MainWindow
import PowerMeter as pwm
import CtrlBoard as cb
import General_functions as gf
import numpy as np
import operator
import time
import configparser
from ThreadCalient import DebugGui
import xlwings as xw
import serial

class main_test(QMainWindow, Ui_MainWindow):
    """
    Class documentation goes here.
    """
    global PM
    global CtrlB
    global log
    global channel
    global temp
    global a
    PM=object
    CtrlB=object


    def __init__(self, parent=None):
        """
        Constructor
        """
        self.log=''
        self.sn=''
        self.a=0
        self.start=time.time()
        self.end=time.time()
        self.finalResult='Pass'


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

        QMainWindow.__init__(self, parent)
        self.setupUi(self)
        self.timer=QTimer(self)
        self.workthread=WorkThread()
        self.workthread.sig_progress.connect(self.update_pro)
        self.workthread.sig_but.connect(self.update_but)
        self.workthread.sig_print.connect(self.print_out_status)
        self.workthread.sig_status.connect(self.updata_status)
        self.workthread.sig_staColor.connect(self.updata_staColor)
        self.workthread.sig_clear.connect(self.text_clear)
        self.workthread.sig_stoptest.connect(self.test_end)
        self.timer.timeout.connect(self.periodly_check)
        self.start_test_butt_3.clicked.connect(self.debug_mode)
        self.start_test_butt_4.clicked.connect(self.openConfig)
        #update the configuration on the label
        self.text_show='测试软件版本：{}\n测试机台号：{}\n测试温度:{}\n测试通道:{}\n光功率计通道:{}\n控制板串口号:{}\n功率计串口号：{}'.format(self.sw,self.desk
                                                                                         ,self.temp,self.channel,self.PM_ch,
                                                                                           self.ctrl_port,self.pow_port
                                                                                           )
        self.label_2.setText(self.text_show)

    def update_config(self):
        self.text_show='测试软件版本：{}\n测试机台号：{}\n测试温度:{}\n测试通道:{}\n光功率计通道:{}\n控制板串口号:{}\n功率计串口号：{}'.format(self.sw,self.desk
                                                                                                       ,self.temp,self.channel,self.PM_ch,
                                                                                                       self.ctrl_port,self.pow_port
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

    #Only for program test
    def time_consuming(self):
        self.a=0
        while self.a<10000000:
            self.a+=1
            if (self.a%100)==0:
                print(self.a)

    #test QTimer
    def periodly_check(self):
        if not self.workthread.isRunning():
            self.start_test_butt_3.setEnabled(True)
            self.start_test_butt_4.setEnabled(True)
            if self.finalResult=='Pass':
                self.test_status.setText("Pass")
            else:
                self.test_status.setText("Fail")
                gf.status_color(self.test_status,'red')
                self.finalResult='Pass'
            self.end=time.time()
            utime=str(time.strftime("%H:%M:%S", time.gmtime(self.end-self.start)))
            self.print_out_status('测试完成，用时：'+utime)
            print('测试完成，用时：',utime)
            self.timer.stop()

    #start the test
    @pyqtSlot()
    def Start_test(self):
        if self.start_test_butt.text()=='开始':
            self.start_test_butt.setText('停止')
            self.start =time.time()
            #self.print_out_status('OK')
            self.timer.start(1000)
            self.start_test_butt_3.setEnabled(False)
            self.start_test_butt_4.setEnabled(False)
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

    def updata_status(self,f):
        mawin.test_status.setText(f)

    def updata_staColor(self,f):
        gf.status_color(self.test_status,f)

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

    def test_end(self):
        #Driver 下电，关闭串口连接
        if serial.Serial.isOpen(self.CtrlB):
            cb.board_set(self,self.drv_down)
            self.print_out_status('Driver下电完成！')
            cb.close_board(self)
            self.print_out_status('测试板串口关闭完成！')
        pwm.close_PM(self)

class WorkThread(QThread):
    sig_progress=pyqtSignal(int)
    sig_status=pyqtSignal(str)
    sig_staColor=pyqtSignal(str)
    sig_print=pyqtSignal(str)
    sig_but=pyqtSignal(str)
    sig_clear=pyqtSignal()

    sig_stoptest=pyqtSignal()

    def __init__(self):
        super(WorkThread, self).__init__()

    def run(self):
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
            self.sig_stoptest.emit()
            self.sig_progress.emit(0)
            return
        elif con==2:
            self.sig_print.emit('请检查，TIA未正确连接')
            self.sig_staColor.emit('blue')
            self.sig_but.emit('开始')
            self.sig_status.emit('请检查TIA连接!')
            self.sig_stoptest.emit()
            self.sig_progress.emit(0)
            return
        elif con==0:
            self.sig_print.emit('请检查，无正确返回值，连接性检查失败！')
            self.sig_staColor.emit('blue')
            self.sig_but.emit('开始')
            self.sig_status.emit('请检查器件连接!')
            self.sig_stoptest.emit()
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
        mawin.channel=[i for i in range(9,73)]
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
            # #get power meter reading of X-max Y-max power
            # abc_ok=max
            # abc_tmp=np.zeros(6)
            # while not operator.eq(abc_ok,abc_tmp):
            #     cb.set_abc(mawin,abc_ok)
            #     abc_tmp=cb.get_abc(mawin)
            # pwr=pwm.read_PM(mawin,mawin.PM_ch)
            # #get power meter reading of X-max Y-min power
            # abc_ok=max[0:3]+min[3:6]
            # abc_tmp=np.zeros(6)
            # while not operator.eq(abc_ok,abc_tmp):
            #     cb.set_abc(mawin,abc_ok)
            #     abc_tmp=cb.get_abc(mawin)
            # pwr_x=pwm.read_PM(mawin,mawin.PM_ch)
            # #get power meter reading of X-min Y-max power
            # abc_ok=min[0:3]+max[3:6]
            # abc_tmp=np.zeros(6)
            # while not operator.eq(abc_ok,abc_tmp):
            #     cb.set_abc(mawin,abc_ok)
            #     abc_tmp=cb.get_abc(mawin)
            # pwr_y=pwm.read_PM(mawin,mawin.PM_ch)
            # #set abc to max
            # abc_ok=max
            # abc_tmp=np.zeros(6)
            # while not operator.eq(abc_ok,abc_tmp):
            #     cb.set_abc(mawin,abc_ok)
            #     abc_tmp=cb.get_abc(mawin)
            # self.sig_print.emit('PDL和IL获取成功...')
            CH=str(mawin.channel[i])
            PDL='0'
            IL='0'
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
        if not os.path.exists(mawin.report_path):
            os.mkdir(mawin.report_path)
        timestamp=gf.get_timestamp(1)
        config=[sn,mawin.desk,mawin.temp[0]]+timestamp.split('_')
        report_name=sn+'_'+timestamp+'.csv'
        report_judgename=sn+'_'+timestamp+'.xlsx'
        report_file=os.path.join(mawin.report_path,report_name)
        report_judge=os.path.join(mawin.report_path,report_judgename)
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
        wb.save(report_judge)
        self.sig_stoptest.emit()

if __name__ == "__main__":
    app=QApplication(sys.argv)
    mawin = main_test()
    debug_gui=DebugGui()
    mawin.show()
    sys.exit(app.exec_())