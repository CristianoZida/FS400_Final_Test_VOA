# -*- coding: utf-8 -*-
# this module is to write the control functions of ZVA
# date:4/14/2022
# author:Jiawen Zhou
#'TCPIP0::192.168.0.15::inst0::INSTR'

import time
import pyvisa
from win32com import client
import win32com

rm=pyvisa.ResourceManager()
rm.list_resources()

def open_OScope(obj,osc_port):
    '''
    Open ZVA connection
    :param obj: GUI Class obj
    :param osc_port: port number
    :return: True/False
    '''
    try:
        obj.OScope = rm.open_resource(osc_port)
        obj.OScope.timeout=3000
        ins=obj.OScope.query('*IDN?')
        print(ins)
        if "LECROY" or "TCPIP" in ins:
            return True
        else:
            return False
    except Exception as e:
        print(e)

def close_OSc(obj):
    '''
    Close Oscilloscope VISA connection and DSO object
    :param obj: GUI Class obj
    :return: NA
    '''
    try:
        if obj.OScope in rm.list_opened_resources():
            obj.OScope.close()
            print('示波器VISA端口关闭成功!')
        else:
            print('示波器VISA端口未打开!')
        obj.DSO.Disconnect()
        print('示波器DSO端口关闭成功!')
    except Exception as e:
        print(e)

def init_OSc(obj):
    '''
    initiate the Oscilloscope settings
    :param obj: GUI Class obj
    :return: NA
    '''
    try:
        if obj.OScope in rm.list_opened_resources():
            #初始化Display
            obj.OScope.write('TRMD AUTO');time.sleep(0.1)
            #初始化Measure
            obj.OScope.write('VBS app.Display.GridMode="Dual"');time.sleep(0.1)
            #设置测量对应的通道
            obj.OScope.write('VBS app.Measure.P1.View=1');time.sleep(0.1)
            obj.OScope.write('VBS app.Measure.P1.Source1="C2"');time.sleep(0.1)
            obj.OScope.write('VBS app.Measure.P1.ParamEngine="PeakToPeak"');time.sleep(0.1)
            obj.OScope.write('VBS app.Measure.P2.View=1');time.sleep(0.1)
            obj.OScope.write('VBS app.Measure.P2.Source1="C3"');time.sleep(0.1)
            obj.OScope.write('VBS app.Measure.P2.ParamEngine="PeakToPeak"');time.sleep(0.1)
            obj.OScope.write('VBS app.Measure.P3.View=1');time.sleep(0.1)
            obj.OScope.write('VBS app.Measure.P3.Source1="C6"');time.sleep(0.1)
            obj.OScope.write('VBS app.Measure.P3.ParamEngine="PeakToPeak"');time.sleep(0.1)
            obj.OScope.write('VBS app.Measure.P4.View=1');time.sleep(0.1)
            obj.OScope.write('VBS app.Measure.P4.Source1="C7"');time.sleep(0.1)
            obj.OScope.write('VBS app.Measure.P4.ParamEngine="PeakToPeak"');time.sleep(0.1)
            obj.OScope.write('VBS app.Measure.P5.View=1');time.sleep(0.1)
            obj.OScope.write('VBS app.Measure.P5.Source1="C2"');time.sleep(0.1)
            obj.OScope.write('VBS app.Measure.P5.ParamEngine="Frequency"');time.sleep(0.1)
            obj.OScope.write('VBS app.Measure.P6.View=1');time.sleep(0.1)
            obj.OScope.write('VBS app.Measure.P6.Source1="C6"');time.sleep(0.1)
            obj.OScope.write('VBS app.Measure.P6.ParamEngine="Frequency"');time.sleep(0.1)
            #打开Std Vertical的测量（打开mean/min/max等的测量）
            obj.OScope.write('VBS app.Measure.StatsOn=1');time.sleep(0.1)
            #初始化4条波形设置
            obj.OScope.write('CHDR OFF');time.sleep(0.1)
            obj.OScope.write('CFMT DEF9,WORD,BIN');time.sleep(0.1)
            obj.OScope.write('WFSU SP,0,NP,0,FP,0,SN,0');time.sleep(0.1)
            obj.OScope.write('CORD LO');time.sleep(0.1)
            #设置通道对应的信号源
            obj.OScope.write('VBS app.Acquisition.C6.ActiveInput="InputB"');time.sleep(0.1)
            obj.OScope.write('C6:TRA ON');time.sleep(0.1)
            obj.OScope.write('VBS app.Acquisition.C7.ActiveInput="InputB"');time.sleep(0.1)
            obj.OScope.write('C7:TRA ON');time.sleep(0.1)
            obj.OScope.write('VBS app.Acquisition.C2.ActiveInput="InputB"');time.sleep(0.1)
            obj.OScope.write('C2:TRA ON');time.sleep(0.1)
            obj.OScope.write('VBS app.Acquisition.C3.ActiveInput="InputB"');time.sleep(0.1)
            obj.OScope.write('C3:TRA ON');time.sleep(0.1)
            obj.OScope.write('VBS app.Acquisition.Horizontal.Maxmize="FixedSampleRate"');time.sleep(0.1)
            obj.OScope.write('VBS app.Acquisition.Horizontal.SampleRate="40000000000"');time.sleep(0.1)
            #设置水平方向采样时间
            obj.OScope.write('tdiv 10e-9');time.sleep(0.1)
            obj.OScope.write('TRMD AUTO');time.sleep(0.1)
            #设置垂直方向参数
            obj.OScope.write('C6:VDIV 0.08');time.sleep(0.1)
            obj.OScope.write('C6:OFFSET 0');time.sleep(0.1)
            obj.OScope.write('C7:VDIV 0.08');time.sleep(0.1)
            obj.OScope.write('C7:OFFSET 0');time.sleep(0.1)
            obj.OScope.write('C2:VDIV 0.08');time.sleep(0.1)
            obj.OScope.write('C2:OFFSET 0');time.sleep(0.1)
            obj.OScope.write('C3:VDIV 0.08');time.sleep(0.1)
            obj.OScope.write('C3:OFFSET 0');time.sleep(0.1)
            #设置（水平方向）采样速率
            #设置触发条件，延时，带宽
            obj.OScope.write('VBS app.Acquisition.Trigger.Type="Edge"');time.sleep(0.1)
            obj.OScope.write('VBS app.Acquisition.Trigger.Edge.Slope="Positive"');time.sleep(0.1)
            obj.OScope.write('VBS app.Acquisition.Trigger.Edge.Level=0');time.sleep(0.1)
            obj.OScope.write('VBS app.Acquisition.C6.OptimizeGroupDelay=2');time.sleep(0.1)
            obj.OScope.write('VBS app.Acquisition.C7.OptimizeGroupDelay=2');time.sleep(0.1)
            obj.OScope.write('VBS app.Acquisition.C2.OptimizeGroupDelay=2');time.sleep(0.1)
            obj.OScope.write('VBS app.Acquisition.C3.OptimizeGroupDelay=2');time.sleep(0.1)
            obj.OScope.write('VBS app.Acquisition.C6.BandwidthLimit="Full"');time.sleep(0.1)
            obj.OScope.write('VBS app.Acquisition.C7.BandwidthLimit="Full"');time.sleep(0.1)
            obj.OScope.write('VBS app.Acquisition.C2.BandwidthLimit="Full"');time.sleep(0.1)
            obj.OScope.write('VBS app.Acquisition.C3.BandwidthLimit="Full"');time.sleep(0.1)
            #设置模式
            obj.OScope.write('VBS app.findAllVerScaleAtCurrentTimebase');time.sleep(0.1)
            print('等待10s,示波器初始化中...')
            time.sleep(10)
            print('示波器初始化成功!')
        else:
            print('示波器VISA端口未打开!')
    except Exception as e:
        print(e)

def switch_to_VISA(obj,dso_port):
    '''
    Switch the Oscilloscope connectiong mode to VISA connection and set up DSO object here
    :param obj:GUI Class obj
    :param dso_port:
    :return:NA
    '''
    try:
        #The next code is to solve the issue:-2147221008, '尚未调用 CoInitialize。', None, None, don`t know why right now
        import pythoncom
        pythoncom.CoInitialize()
        obj.DSO=win32com.client.Dispatch('LeCroy.ActiveDSOCtrl.1')
        obj.DSO.MakeConnection(dso_port)
        obj.DSO.WriteString('VBS app.Utility.Remote.Interface="LXI"',True)
        print('示波器端口切换VISA模式成功!')
    except Exception as e:
        print(e)


def switch_to_DSO(obj,VISA_port):
    '''
    Switch the Oscilloscope connectiong mode to DSO(TCP/IP) connection
    :param obj: GUI Class obj
    :param VISA_port:
    :return: NA
    '''
    try:
        if not obj.OScope in rm.list_opened_resources():
            open_OScope(obj,VISA_port)
        obj.OScope.write('VBS app.Utility.Remote.Interface="TCPIP"')
        print('示波器端口切换DSO模式成功!')
    except Exception as e:
        print(e)

def get_data_DSO(obj):
    '''
    get data in DSO mode
    :param obj: GUI Class obj
    :return: data={'CH2':'','CH3':'','CH6':'','CH7':''}
    '''
    obj.DSO=win32com.client.Dispatch('LeCroy.ActiveDSOCtrl.1')
    obj.DSO.MakeConnection(obj.DSO_port)
    #data={'CH2':'','CH3':'','CH6':'','CH7':''}
    data=['']*4
    switch_to_DSO(obj,obj.OScope_port)
    # data['CH2']=dso.GetScaledWaveformWithTimes('C2',802,0)
    # data['CH3']=dso.GetScaledWaveformWithTimes('C3',802,0)
    # data['CH6']=dso.GetScaledWaveformWithTimes('C6',802,0)
    # data['CH7']=dso.GetScaledWaveformWithTimes('C7',802,0)

    data[0]=obj.DSO.GetScaledWaveformWithTimes('C2',802,0)
    data[1]=obj.DSO.GetScaledWaveformWithTimes('C3',802,0)
    data[2]=obj.DSO.GetScaledWaveformWithTimes('C6',802,0)
    data[3]=obj.DSO.GetScaledWaveformWithTimes('C7',802,0)
    return data

def get_data_VISA(obj):
    '''
    function:get data in VISA mode
    :param obj: GUI Class obj
    :return: data={'CH2':'','CH3':'','CH6':'','CH7':''}
    '''
    data=['']*4
    switch_to_VISA(obj,obj.DSO_port)
    dataS100=obj.OScope.query('C2:INSPECT? "SIMPLE"')
    time.sleep(0.1)
    while dataS100=='':
        dataS100=obj.OScope.query('C2:INSPECT? "SIMPLE"')
        time.sleep(0.1)
    dataS103=obj.OScope.query('C2:INSPECT? "HORIZ_OFFSET"')
    time.sleep(0.1)
    dataS104=obj.OScope.query('C2:INSPECT? "HORIZ_INTERVAL"')
    time.sleep(0.1)

    dataS200=obj.OScope.query('C3:INSPECT? "SIMPLE"')
    time.sleep(0.1)
    while dataS200=='':
        dataS200=obj.OScope.query('C3:INSPECT? "SIMPLE"')
        time.sleep(0.1)
    dataS203=dataS103
    dataS204=dataS104

    dataS300=obj.OScope.query('C6:INSPECT? "SIMPLE"')
    time.sleep(0.1)
    while dataS200=='':
        dataS300=obj.OScope.query('C6:INSPECT? "SIMPLE"')
        time.sleep(0.1)
    dataS303=dataS103
    dataS304=dataS104

    dataS400=obj.OScope.query('C7:INSPECT? "SIMPLE"')
    time.sleep(0.1)
    while dataS400=='':
        dataS400=obj.OScope.query('C7:INSPECT? "SIMPLE"')
        time.sleep(0.1)
    dataS403=dataS103
    dataS404=dataS104
    # data handle
    # dataS103 = get_num(dataS103);
    # dataS104 = get_num(dataS104);
    # xC2=getxdata(dataS103,dataS104);
    # yC2=str2num_new(dataS100);
    # xC2=xC2(2:end-1);
    # yC2=yC2(2:end-1);
    # data1(:,1)=xC2;
    # data1(:,2)=yC2;
    #
    # dataS203 = get_num(dataS203);
    # dataS204 = dataS104;%%%相同的时间间隔
    # xC3=getxdata(dataS203,dataS204);
    # yC3=str2num_new(dataS200);
    # xC3=xC3(2:end-1);
    # yC3=yC3(2:end-1);
    # data2(:,1)=xC3;
    # data2(:,2)=yC3;
    #
    # dataS303 = get_num(dataS303);
    # dataS304 = dataS104;%%%相同的时间间隔
    # xC6=getxdata(dataS303,dataS304);
    # yC6=str2num_new(dataS300);
    # xC6=xC6(2:end-1);
    # yC6=yC6(2:end-1);
    # data3(:,1)=xC6;
    # data3(:,2)=yC6;
    #
    # dataS403 = get_num(dataS403);
    # dataS404 = dataS104;%%%相同的时间间隔
    # xC7=getxdata(dataS403,dataS404);
    # yC7=str2num_new(dataS400);
    # xC7=xC7(2:end-1);
    # yC7=yC7(2:end-1);
    # data4(:,1)=xC7;
    # data4(:,2)=yC7;

    return data

def auto_Osc(obj):
    '''
    #Unused!!!
function:Automate Oscilloscope
    :param obj: GUI Class obj
    :return: NA
    '''
    try:
        if obj.OScope in rm.list_opened_resources():
            #设置扫描模式
            obj.OScope.write('TRMD AUTO')
            #设置要测量的内容
            obj.OScope.write('VBS app.Measure.MeasurementSet="MyMeasure"')
            obj.OScope.write('VBS app.Measure.P1.MeasurementType="measure"')
            obj.OScope.write('VBS app.Measure.P1.Source="C6"')
            obj.OScope.write('VBS app.Measure.P1.View=1')
            dataS32=obj.OScope.query('VBS? return=app.Measure.P1.View')
            obj.OScope.write('VBS app.Measure.P1.ParamEngine="PeakToPeak"')
            obj.OScope.write('VBS app.Measure.P2.MeasurementType="measure"')
            obj.OScope.write('VBS app.Measure.P2.Source="C"')
            obj.OScope.write('VBS app.Measure.P2.View=1')
            dataS33=obj.OScope.query('VBS? return=app.Measure.P2.View')
            obj.OScope.write('VBS app.Measure.P2.ParamEngine="PeakToPeak"')
            obj.OScope.write('VBS app.Measure.P3.MeasurementType="measure"')
            obj.OScope.write('VBS app.Measure.P3.Source="C7"')
            dataS34=obj.OScope.query('VBS? return=app.Measure.P2.View')

        else:
            print('示波器VISA端口未打开!')
    except Exception as e:
        print(e)

