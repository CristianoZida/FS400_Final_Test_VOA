# -*- coding: utf-8 -*-
#this module is translated from the control functions of ctrol board provided by Fslink
# date:3/17/2022
# author:Jiawen Zhou
#0620/2022:update connectivity test with PD current judgement(ITLA on)
import serial
import time
import numpy as np
import re
import math


def open_board(obj,board_port):
    '''
    Open board connection
    :param obj:
    :param board_port:
    :return: True/False
    '''
    try:
        obj.CtrlB = serial.Serial(board_port,115200,timeout=20)
        #Update 05.23.2022
        obj.CtrlB.write(b'fs400_stop_fs400_performance\n') #stop BER readings
        i=0
        count=0
        while i<25:
            time.sleep(0.2)
            obj.CtrlB.write(b'\r\n')
            time.sleep(5)
            a=obj.CtrlB.read_all()
            print(a)
            if len(a)>17:
                count=0
            if len(a)==17:
                count+=1
                if count>2:
                    print('控制板通信正常！')
                    return True
            if len(a)==0:
                print('控制板未上电！')
                return False
            i+=1
        print('读取120s超时，请检查控制板是否开启')
        return False

    except Exception as e:
        print(e)

def close_board(obj):
    '''
    Close board serial connection
    :param obj:
    :return: NA
    '''
    try:
        if serial.Serial.isOpen(obj.CtrlB):
            obj.CtrlB.close()
            print('控制板串口关闭!')
        else:
            print('控制板已是关闭状态!')
    except Exception as e:
        print(e)

def read_comd(file):
    '''
    Read command from txt file
    :param file:
    :return: valid command list with "\p" stands for pause 1 sec(time.sleep(1))
    '''
    with open(file,'r',encoding='utf-8') as f:
        comd=[i for i in f.readlines() if not '//' in i and not '->' in i][:-1]
        for i in comd:
            if i =='\n':
                comd.remove(i)
    return comd

def board_set(obj,f):
    '''
    Board up/drv up/drv down
    :param obj: GUI Class obj
    :param f: file absolote path(including three status)
    :return: Ture/False
    '''
    obj.CtrlB.flushInput()
    obj.CtrlB.flushOutput()
    obj.CtrlB.timeout=30
    try:
        if "brdup" in f:
            stri="board power on!"
            wai=b'en_adj 4 1'
        elif "driverdown" in f:
            stri="Driver power off!"
        elif "driverup" in f:
            stri="Driver power on!"
        colist=read_comd(f)
        for i in colist:
            if i=='\\p\n':
                time.sleep(1)
                continue
            if 'itla_wr' in i:# 0 0x32 0x08' in i:
                obj.CtrlB.write(i.encode('utf-8'))
                time.sleep(3)
                a=obj.CtrlB.read_until(b'Write itla')
                time.sleep(0.5)
            else:
                obj.CtrlB.write(i.encode('utf-8'))
                time.sleep(0.1)
                a=obj.CtrlB.read_all()
            time.sleep(0.1)
            print(a)
        #a=obj.CtrlB.read_until('\r\n')
        # if "brdup" in f:
        #     a=obj.CtrlB.read_until(wai)
        # else:
        #     a=obj.CtrlB.read_all()
        obj.CtrlB.timeout=10
        #print(a)
        print(stri)
        #return a
        #return obj.CtrlB.read_all()
    except Exception as e:
        print(e)

def get_abc(obj):
    '''
    Get the abc status
    :param obj: GUI Class obj
    :return: abc data 6 voltage items array-(XI,XQ,XP,YI,YQ,YP)
    '''
    obj.CtrlB.flushInput()
    obj.CtrlB.flushOutput()
    data=np.zeros(6)
    obj.CtrlB.write(b'abc_show\n')
    time.sleep(0.1)
    tm=obj.CtrlB.read_until(b'shell').decode('utf-8')
    print(tm)
    #tm=obj.CtrlB.read_all().decode('utf-8')
    if not 'abc_show' in tm:
        obj.CtrlB.write(b'abc_show\n')
        time.sleep(0.1)
        tm=obj.CtrlB.read_until(b'shell').decode('utf-8')
        #tm=obj.CtrlB.read_all().decode('utf-8')#.split('\n')
    tmp=tm.split('\n')
    ind=[tmp.index(i) for i in tmp if 'abc_show' in i][0]
    #tmp=tmp[[tmp.index(i) for i in tmp if 'abc_show' in i][0]]+tmp[[tmp.index(i) for i in tmp if 'abc_show' in i][1]]
    tmp1=re.split('[:,\r]',tmp[ind+1]+tmp[ind+2])
    data=[tmp1[i] for i in range(1,12,2)]
    return data

def set_abc(obj,data):
    '''
    Set the abc value
    :param obj: GUI Class obj
    :param data: data of 6 items array-(XI,XQ,XP,YI,YQ,YP)
    :return: NA
    '''
    obj.CtrlB.write(b'fpga_spi_wr 0x70 0 0x0\n')
    time.sleep(0.1)
    obj.CtrlB.write(b'fpga_spi_wr 0x80 0 0x0\n')
    time.sleep(0.1)
    obj.CtrlB.write(b'fpga_spi_wr 0x90 0 0x3f\n')
    time.sleep(0.1)
    for i in range(6):
        comd='fpga_spi_wr 0x9'+str(i+1)+' 0 '+str(int(data[i]))+'\n' #Need to delete the '0' etc..'0900' otherwise will return '0000'
        obj.CtrlB.write(comd.encode('utf-8'))
        time.sleep(0.1)
    print(obj.CtrlB.read_all())

def pow_scan(obj,ch,start,stop,step):
    '''
    pow scan to get the min/max according to Tx QMPD readings by adjusting the phase shifter
    :param obj: GUI Class obj
    :param ch: channel(phase tuner)
    :param start: start voltage
    :param stop: stop voltage
    :param step:voltage step
    :return: ideal voltage point array
    '''
    data=[]
    if 1<=int(ch)<=6 and start>=0 and stop<=4095 and start<=stop and step>1:
        if (stop-start)/step>300:
            step=20
        abc=get_abc(obj)
        obj.CtrlB.flushInput()
        obj.CtrlB.flushOutput()
        time.sleep(0.1)
        str_1='pow_scan '+str(ch-1)+' '+str(start)+' '+str(stop)+' '+str(step)+'\n'
        obj.CtrlB.write(str_1.encode('utf-8'))
        time.sleep(0.3)
        #Decode the output message of the voltage
        obj.timeout=30
        p=obj.CtrlB.read_until(b'Mission completed!\r\n').decode('utf-8')
        obj.timeout=10
        a1=p.index('min_num')
        a2=p.index('max_num')
        a3=p.index('Mission')
        minn=p[a1:a2]
        maxn=p[a2:(a3-2)]
        te=re.split('[=\r\n]',minn)
        te=[i for i in te if not 'min_' in i and i !='']
        te1=re.split('[=\r\n]',maxn)
        te1=[i for i in te1 if not 'max' in i and i !='']
        data=[te[0],te[1:],te1[0],te1[1:]] #For example:['8',['1140', '1700', '2300', '2640', '2920', '3200', '3460', '3740'],'9',['700', '1440', '1920', '2480', '2780', '3060', '3340', '3580', '3800']]
        set_abc(obj,abc)
    else:
        print('Error input, please check!')
    print("Pow Scan value:\n",data)
    return data

def get_chvpi_old(obj,ch,start,stop,step):
    '''
    get CH VPI value
    :param obj:
    :param ch:
    :param start:
    :param stop:
    :param step:
    :return: CHVPI value
    '''
    Tvpi=0.0
    data=pow_scan(obj,ch,start,stop,step)
    if len(data)==4 and data[1]!='' and data[3]!='':
        minv=[int(i) for i in data[1]]
        maxv=[int(i) for i in data[3]]
        cbias=[i for i in sorted(minv+maxv) if i>500]
        pbias=[(i/4096*3.3)**2 for i in cbias]
        x=[i for i in range(1,len(pbias)+1)]
        R2=(np.corrcoef(x,pbias))**2
        if np.shape(R2)[0]==2:
            R2=R2[0,1]
            if R2>0.9989:#相关系数需要调节0.9982不行
                p=np.polyfit(x,pbias,1)
                Tvpi=3.3**2/p[0]/2
                return Tvpi
            else:
                delta=np.diff(pbias)
                tmp=[i for i in delta if i<(np.mean(delta)+2*np.std(delta))]
                tmp_1=[i for i in tmp if i>(np.mean(tmp)-2*np.std(tmp))]
                if len(tmp_1)==0:
                    Tvpi=0
                else:
                    Tvpi=3.3**2/np.mean(tmp_1)/2
                    return Tvpi
    else:
        return Tvpi
        data=[]
        print('No data get from pow scan or min/max not exists')

def get_chvpi(obj,ch,start,stop,step):
    '''
    09.20.2022:Update the algorithm
    get CH VPI value
    :param obj:
    :param ch:
    :param start:
    :param stop:
    :param step:
    :return: CHVPI value
    '''
    Tvpi=0.0
    data=pow_scan(obj,ch,start,stop,step)
    if len(data)==4 and data[1]!='' and data[3]!='':
        minv=[int(i) for i in data[1]]
        maxv=[int(i) for i in data[3]]
        cbias=[i for i in sorted(minv+maxv) if i>880]  #500-->880 BY BOB
        pbias=[(i/4096*3.3)**2 for i in cbias]
        x=[i for i in range(1,len(pbias)+1)]
        R2=(np.corrcoef(x,pbias))**2
        if np.shape(R2)[0]==2:
            R2=R2[0,1]
            # if R2>0.9989:#相关系数需要调节0.9982不行
            #     p=np.polyfit(x,pbias,1)
            #     Tvpi=3.3**2/p[0]/2
            #     return Tvpi
            # else:
            delta=np.diff(pbias)
            tmp_0=[i for i in delta if i>(0.3)]  # add by bob  对应 <500000 code^2 的最大最小差
            tmp=[i for i in tmp_0 if i<(np.median(tmp_0)+2*np.std(tmp_0))]  # mean-> median by bob
            tmp_1=[i for i in tmp if i>(np.median(tmp)-2*np.std(tmp))]  # mean-> median by bob
            if len(tmp_1)==0:
                Tvpi=0
            else:
                Tvpi=3.3**2/np.median(tmp_1)/2   # mean-> median by bob
                return Tvpi
    else:
        return Tvpi
        data=[]
        print('No data get from pow scan or min/max not exists')

def get_noipwr(obj):
    '''
    get noise power value
    :param obj:
    :return: Noise power value([noipower1,noipower2])
    '''
    obj.CtrlB.flushOutput()
    obj.CtrlB.timeout=30
    obj.CtrlB.write(b'get_noipwr\n')
    time.sleep(0.1)
    p=obj.CtrlB.read_until(b'Mission completed!\r\n').decode('utf-8')
    s=p[p.index('Xp_'):p.index('Mission')-2]
    da=[i for i in re.split('[\r\n:]',s) if i!='' and not '_' in i]
    obj.CtrlB.timeout=10
    return da #[Xp_noipwr0,Yp_noipwr1]

def get_quad(obj,ch,start,stop,step):
    '''
    get quad value
    :param obj:
    :param ch:
    :param start:
    :param stop:
    :param step:
    :return: quad voltage value
    '''
    data=pow_scan(obj,ch,start,stop,step)
    if len(data)==4 and data[1]!='' and data[3]!='':
        minv=[int(i) for i in data[1]]
        maxv=[int(i) for i in data[3]]
        result=round(math.sqrt(min(minv)**2+min(maxv)**2)/math.sqrt(2))
    else:
        data=[]
        print('No data get from pow scan or min/max not exists')
    return result

def get_ER_OLD(obj,ch,noi1,noi2,erflag,vpiflag):
    '''
    get ER value
    :param obj:
    :param ch:
    :param noi1:
    :param noi2:
    :param erflag:
    :param vpiflag:
    :return: ER value and others(list in list)
            #max:maxbias - 6 items   XI/XQ/XP/YI/YQ/YP
            #min:minbias - 6 items   XI/XQ/XP/YI/YQ/YP
            #abc:abc point - 6 items XI/XQ/XP/YI/YQ/YP
            #min:er - 6 items of corresponding channel
            #Tvpi - 6 items, T of corrsponding channel
    '''
    #for example:get_er ch 0 noi1 noi2 1
    str_1=('get_er '+str(ch)+' 0 '+str(noi1)+' '+str(noi2)+' '+str(erflag)+'\n').encode('utf-8')
    #3 times to search ER, if mission failed 3 times,will return empty result, other wise continue
    # for i in range(3):
    #     obj.CtrlB.flushInput()
    #     obj.CtrlB.flushOutput()
    #     obj.CtrlB.write(str_1)
    #     time.sleep(0.1)
    #     obj.CtrlB.timeout=180
    #     log=obj.CtrlB.read_until(b'Mission completed!').decode('utf-8')
    #     obj.CtrlB.timeout=10
    #     if log.find("Mission failed!")!=-1:
    #         break
    #     elif i==2:
    #         result=[]
    #         return
    #     abc=get_abc(obj)
    #     for i in range(len(abc)):
    #         if int(abc[i])<800:
    #             abc[i]='1500'
    #     set_abc(obj,abc)
    #     #log=''
    obj.CtrlB.flushInput()
    obj.CtrlB.flushOutput()
    obj.CtrlB.write(str_1)
    time.sleep(0.1)
    obj.CtrlB.timeout=180
    log=obj.CtrlB.read_until(b'Mission completed!').decode('utf-8')
    obj.CtrlB.timeout=10
    if log.find("Mission failed!")!=-1:
        result=[]
        return result

    abc=get_abc(obj)
    for i in range(len(abc)):
        if int(abc[i])<800:
            abc[i]='1500'
    set_abc(obj,abc)

    log_list=log.split('\r\n')
    #Create function to screen the index of ABC and Maxbias
    fin=[]
    biamax=[]
    for i in log_list:
        if 'ABC point' in i:
            fin.append(log_list.index(i))
        if 'Biasmax' in i:
            biamax.append(log_list.index(i))
    if len(fin)>1 or len(fin)==0:
        print('Multiple ABC point or no ABC point returned!')
        return False
    if len(biamax)>1 or len(biamax)==0:
        print('Multiple biamax or no biamax returned!')
        return False

    abc_x=re.split('[:,]',log_list[fin[0]+1])
    abc_y=re.split('[:,]',log_list[fin[0]+2])
    abc=[abc_x[1],abc_x[3],abc_x[5],abc_y[1],abc_y[3],abc_y[5]]
    print(abc)

    er=np.zeros(6)
    if erflag>0:
        tmp=re.split('[\r\n ]',log_list[biamax[0]-1])
        er=[float(i) for i in tmp if i!='']
        print(er)
    else:
        er=er.tolist()

    tmp_1=re.split('[\r\n,:]',log_list[biamax[0]])
    max=[i for i in tmp_1[1:7]]

    tmp_2=re.split('[\r\n,:]',log_list[biamax[0]+1])
    min=[i for i in tmp_2[1:7]]

    Tvpi=np.zeros(6)
    if vpiflag>0:
        abc_ok=min
        set_abc(obj,abc_ok)
        for ch in [1,2,4,5]:
            Tvpi[ch-1]=get_chvpi(obj,ch,0,4000,20)
            print(Tvpi)
        abc_ok=max
        set_abc(obj,abc_ok)
        for ch in [3,6]:
            Tvpi[ch-1]=get_chvpi(obj,ch,0,4000,20)
            print(Tvpi)
    '''return the result by list in list
    #max:maxbias - 6 items
    #min:minbias - 6 items
    #abc:abc point - 6 items XI/XQ/XP/YI/YQ/YP
    #min:er - 6 items of corresponding channel
    #Tvpi - 6 items, T of corrsponding channel
    '''
    Tvpi=Tvpi.tolist()
    print ([max,min,abc,er,Tvpi])
    return [max,min,abc,er,Tvpi]

def get_ER(obj,ch,noi1,noi2,erflag,vpiflag):
    '''
    get ER value with 'get_er 0 ch 0 noi1 noi2 1'
    :param obj:
    :param ch:
    :param noi1:
    :param noi2:
    :param erflag:
    :param vpiflag:
    :return: ER value and others(list in list)
            #max:maxbias - 6 items   XI/XQ/XP/YI/YQ/YP
            #min:minbias - 6 items   XI/XQ/XP/YI/YQ/YP
            #abc:abc point - 6 items XI/XQ/XP/YI/YQ/YP
            #min:er - 6 items of corresponding channel
            #Tvpi - 6 items, T of corrsponding channel
    '''
    #for example:get_er 0 ch 0 noi1 noi2 1
    str_1=('get_er 0 '+str(ch)+' 0 '+str(noi1)+' '+str(noi2)+' '+str(erflag)+'\n').encode('utf-8')
    obj.CtrlB.flushInput()
    obj.CtrlB.flushOutput()
    obj.CtrlB.write(str_1)
    time.sleep(0.1)
    obj.CtrlB.timeout=180
    log=obj.CtrlB.read_until(b'Mission completed!').decode('utf-8')
    obj.CtrlB.timeout=10
    if log.find("Mission failed!")!=-1:
        #if 'Mission failed found,retest with get_er 0...'
        obj.CtrlB.write(str_1)
        time.sleep(0.1)
        obj.CtrlB.timeout = 180
        log = obj.CtrlB.read_until(b'Mission completed!').decode('utf-8')
        obj.CtrlB.timeout = 10
        if log.find("Mission failed!") != -1:
            result = []
            return result

    abc=get_abc(obj)
    for i in range(len(abc)):
        if int(abc[i])<800:
            abc[i]='1500'
    set_abc(obj,abc)

    log_list=log.split('\r\n')
    #Create function to screen the index of ABC and Maxbias
    fin=[]
    biamax=[]
    for i in log_list:
        if 'ABC point' in i:
            fin.append(log_list.index(i))
        if 'Biasmax' in i:
            biamax.append(log_list.index(i))
    if len(fin)>1 or len(fin)==0:
        print('Multiple ABC point or no ABC point returned!')
        return False
    if len(biamax)>1 or len(biamax)==0:
        print('Multiple biamax or no biamax returned!')
        return False

    abc_x=re.split('[:,]',log_list[fin[0]+1])
    abc_y=re.split('[:,]',log_list[fin[0]+2])
    abc=[abc_x[1],abc_x[3],abc_x[5],abc_y[1],abc_y[3],abc_y[5]]
    print(abc)

    er=np.zeros(6)
    if erflag>0:
        tmp=re.split('[\r\n ]',log_list[biamax[0]-1])
        er=[float(i) for i in tmp if i!='']
        print(er)
    else:
        er=er.tolist()

    tmp_1=re.split('[\r\n,:]',log_list[biamax[0]])
    max=[i for i in tmp_1[1:7]]

    tmp_2=re.split('[\r\n,:]',log_list[biamax[0]+1])
    min=[i for i in tmp_2[1:7]]

    Tvpi=np.zeros(6)
    if vpiflag>0:
        abc_ok=min
        set_abc(obj,abc_ok)
        for ch in [1,2,4,5]:
            Tvpi[ch-1]=get_chvpi(obj,ch,0,4000,20)
            print(Tvpi)
        abc_ok=max
        set_abc(obj,abc_ok)
        for ch in [3,6]:
            Tvpi[ch-1]=get_chvpi(obj,ch,0,4000,20)
            print(Tvpi)
    '''return the result by list in list
    #max:maxbias - 6 items
    #min:minbias - 6 items
    #abc:abc point - 6 items XI/XQ/XP/YI/YQ/YP
    #min:er - 6 items of corresponding channel
    #Tvpi - 6 items, T of corrsponding channel
    '''
    Tvpi=Tvpi.tolist()
    print ([max,min,abc,er,Tvpi])
    return [max,min,abc,er,Tvpi]

def get_ER_New1(obj,ch,noi1,noi2,erflag,vpiflag):
    '''
    get ER value with 'get_er 1 ch 0 noi1 noi2 1'
    :param obj:
    :param ch:
    :param noi1:
    :param noi2:
    :param erflag:
    :param vpiflag:
    :return: ER value and others(list in list)
            #max:maxbias - 6 items   XI/XQ/XP/YI/YQ/YP
            #min:minbias - 6 items   XI/XQ/XP/YI/YQ/YP
            #abc:abc point - 6 items XI/XQ/XP/YI/YQ/YP
            #min:er - 6 items of corresponding channel
            #Tvpi - 6 items, T of corrsponding channel
    '''
    #for example:get_er 1 ch 0 noi1 noi2 1
    str_1=('get_er 1 '+str(ch)+' 0 '+str(noi1)+' '+str(noi2)+' '+str(erflag)+'\n').encode('utf-8')
    obj.CtrlB.flushInput()
    obj.CtrlB.flushOutput()
    obj.CtrlB.write(str_1)
    time.sleep(0.1)
    obj.CtrlB.timeout=360
    log=obj.CtrlB.read_until(b'Mission completed!').decode('utf-8')
    print(log)
    obj.CtrlB.timeout=10
    if log.find("Mission failed!")!=-1:
        # if 'Mission failed found,retest with get_er 0...'
        str_1 = ('get_er 0 ' + str(ch) + ' 0 ' + str(noi1) + ' ' + str(noi2) + ' ' + str(erflag) + '\n').encode('utf-8')
        obj.CtrlB.write(str_1)
        time.sleep(0.1)
        obj.CtrlB.timeout = 180
        log = obj.CtrlB.read_until(b'Mission completed!').decode('utf-8')
        obj.CtrlB.timeout = 10
        if log.find("Mission failed!") != -1:
            result = []
            return result

    abc=get_abc(obj)
    for i in range(len(abc)):
        if int(abc[i])<800:
            abc[i]='1500'
    set_abc(obj,abc)

    log_list=log.split('\r\n')
    #Create function to screen the index of ABC and Maxbias
    fin=[]
    biamax=[]
    for i in log_list:
        if 'ABC point' in i:
            fin.append(log_list.index(i))
        if 'Biasmax' in i:
            biamax.append(log_list.index(i))
    if len(fin)>1 or len(fin)==0:
        print('Multiple ABC point or no ABC point returned!')
        return False
    if len(biamax)>1 or len(biamax)==0:
        print('Multiple biamax or no biamax returned!')
        return False

    abc_x=re.split('[:,]',log_list[fin[0]+1])
    abc_y=re.split('[:,]',log_list[fin[0]+2])
    abc=[abc_x[1],abc_x[3],abc_x[5],abc_y[1],abc_y[3],abc_y[5]]
    print(abc)

    er=np.zeros(6)
    if erflag>0:
        tmp=re.split('[\r\n ]',log_list[biamax[0]-1])
        er=[float(i) for i in tmp if i!='']
        print(er)
    else:
        er=er.tolist()

    tmp_1=re.split('[\r\n,:]',log_list[biamax[0]])
    max=[i for i in tmp_1[1:7]]

    tmp_2=re.split('[\r\n,:]',log_list[biamax[0]+1])
    min=[i for i in tmp_2[1:7]]

    Tvpi=np.zeros(6)
    if vpiflag>0:
        abc_ok=min
        set_abc(obj,abc_ok)
        for ch in [1,2,4,5]:
            Tvpi[ch-1]=get_chvpi(obj,ch,0,4000,20)
            print(Tvpi)
        abc_ok=max
        set_abc(obj,abc_ok)
        for ch in [3,6]:
            Tvpi[ch-1]=get_chvpi(obj,ch,0,4000,20)
            print(Tvpi)
    '''return the result by list in list
    #max:maxbias - 6 items
    #min:minbias - 6 items
    #abc:abc point - 6 items XI/XQ/XP/YI/YQ/YP
    #min:er - 6 items of corresponding channel
    #Tvpi - 6 items, T of corrsponding channel
    '''
    Tvpi=Tvpi.tolist()
    print ([max,min,abc,er,Tvpi])
    return [max,min,abc,er,Tvpi]


def test_connectivity(obj):
    '''
    Check pin connectivity
    :param obj:
    :return: 1 # Driver not connected
            2 # TIA not connected
            3 # connetivity ok
    '''
    ##first board up
    board_set(obj,obj.board_up)
    print('控制板初始化完成！')
    print(obj.board_up)
    obj.CtrlB.flushInput()
    obj.CtrlB.flushOutput()
    #Driver check e0cf
    print('Driver 连接性检查中！')
    i=0
    while i <4:
        time.sleep(0.2)
        obj.CtrlB.write(b'fs400_drv_read 0x0401\n')
        time.sleep(0.2)
        log=obj.CtrlB.read_all()#.decode('utf-8')
        print('Driver check times NO.{} : {}'.format(str(i+1),log.decode('utf-8')))
        #log=obj.CtrlB.read_until(b'0x00ff 0x00ff').decode('utf-8')
        #to be continued
        if b'0x00e0 0x00cf' in log:
            break
        else:
            i+=1
            if i==3:
                return 1

    #TIA check e0cf
    obj.CtrlB.flushInput()
    obj.CtrlB.flushOutput()
    print('TIA 连接性检查中！')
    j=0
    while j<4:
        time.sleep(0.2)
        obj.CtrlB.write(b'fs400_tia_read 0x0001\n')
        time.sleep(0.2)
        log=obj.CtrlB.read_all()#.decode('utf-8')
        print('TIA check times NO.{} : {}'.format(str(j+1),log.decode('utf-8')))
        #log=obj.CtrlB.read_until(b'0x00ff 0x00ff').decode('utf-8')
        if b'0x00e0 0x00cf' in log:
            break
        else:
            j+=1
            if j==3:
                return 2
    #to becontinued
    return 3

def test_connectivity_new(obj,sn):
    '''
    update 0620:make sure ITLA is on to check connectivity and judge whether IDT or ALU
    Check pin connectivity
    :param obj:NA sn:SN to judge whether ALU or IDT
    :return: 1 # Driver is ALU
            2 # Driver is IDT
            3 # connectivity fail
    '''
    ##first board up
    board_set(obj,obj.board_up)
    print('控制板初始化完成！')
    print(obj.board_up)
    obj.CtrlB.flushInput()
    obj.CtrlB.flushOutput()
    print('FS400连接性检查中！')
    obj.CtrlB.write(b'fpga_spi_rd 0x4a\n');time.sleep(0.5)
    ret=obj.CtrlB.read_all().decode('utf-8')
    try:
        cur=int(re.split('[(,)]',ret)[-2])
        print('Connectivity test current Tx MPD X:',cur)
    except Exception as e:
        print(e,'Current value get failed!')
        return 3 #current value get failed
    obj.CtrlB.write(b'fpga_spi_rd 0x47\n');time.sleep(0.5)
    ret = obj.CtrlB.read_all().decode('utf-8')
    try:
        cur1 = int(re.split('[(,)]', ret)[-2])
        print('Connectivity test current Tx MPD Y:', cur1)
    except Exception as e:
        print(e, 'Current value get failed!')
        return 3  # current value get failed
    if not (cur>11500 and cur1>11500):
        print('MPD connection failed! please check!!!')
        return 3
    judge_Drv=True
    judge_TIA=True
    #Driver check e0cf
    print('Driver 连接性检查中！')
    i=0
    while i <4:
        time.sleep(0.2)
        obj.CtrlB.write(b'fs400_drv_read 0x0401\n')
        time.sleep(0.2)
        log=obj.CtrlB.read_all()#.decode('utf-8')
        print('Driver check times NO.{} : {}'.format(str(i+1),log.decode('utf-8')))
        #log=obj.CtrlB.read_until(b'0x00ff 0x00ff').decode('utf-8')
        #to be continued
        if b'0x00e0 0x00cf' in log:
            break
        else:
            i+=1
            if i==3:
                judge_Drv=False
                #return 1 #Driver is ALU

    #TIA check e0cf
    obj.CtrlB.flushInput()
    obj.CtrlB.flushOutput()
    print('TIA 连接性检查中！')
    j=0
    while j<4:
        time.sleep(0.2)
        obj.CtrlB.write(b'fs400_tia_read 0x0001\n')
        time.sleep(0.2)
        log=obj.CtrlB.read_all()#.decode('utf-8')
        print('TIA check times NO.{} : {}'.format(str(j+1),log.decode('utf-8')))
        #log=obj.CtrlB.read_until(b'0x00ff 0x00ff').decode('utf-8')
        if b'0x00e0 0x00cf' in log:
            break
        else:
            j+=1
            if j==3:
                judge_TIA=False
                # print('Please check whether TIA is powered on, driver is IDT and connected, TIA not connected!')
                # return 3
    if judge_Drv and judge_TIA:
        print('Driver is IDT')
        return 2 #Driver is IDT
    elif not judge_Drv and not judge_TIA and sn[1:3]=='AA':
        print('Driver is ALU')
        return 1 #Driver is ALU
    else:
        print('Please check whether TIA is powered on, driver is IDT and connected, TIA not connected!')
        return 3
    #to becontinued

def test_connectivity_new_DC(obj,sn):
    '''
    update 0620:make sure ITLA is on to check connectivity and judge whether IDT or ALU
    Check pin connectivity
    :param obj:NA sn:SN to judge whether ALU or IDT
    :return: 1 # Driver is ALU
            2 # Driver is IDT
            3 # connectivity fail
    '''
    ##first board up
    #board_set(obj,obj.board_up)
    print('控制板初始化完成！')
    print(obj.board_up)
    obj.CtrlB.flushInput()
    obj.CtrlB.flushOutput()
    print('FS400连接性检查中！')
    obj.CtrlB.write(b'fpga_spi_rd 0x4a\n');time.sleep(0.5)
    ret=obj.CtrlB.read_all().decode('utf-8')
    try:
        cur=int(re.split('[(,)]',ret)[-2])
        print('Connectivity test current Tx MPD X:',cur)
    except Exception as e:
        print(e,'Current value get failed!')
        return 3 #current value get failed
    obj.CtrlB.write(b'fpga_spi_rd 0x47\n');time.sleep(0.5)
    ret = obj.CtrlB.read_all().decode('utf-8')
    try:
        cur1 = int(re.split('[(,)]', ret)[-2])
        print('Connectivity test current Tx MPD Y:', cur1)
    except Exception as e:
        print(e, 'Current value get failed!')
        return 3  # current value get failed
    if not (cur>11500 and cur1>11500):
        print('MPD connection failed! please check!!!')
        return 3
    judge_Drv=True
    judge_TIA=True
    #Driver check e0cf
    print('Driver 连接性检查中！')
    i=0
    while i <4:
        time.sleep(0.2)
        obj.CtrlB.write(b'fs400_drv_read 0x0401\n')
        time.sleep(0.2)
        log=obj.CtrlB.read_all()#.decode('utf-8')
        print('Driver check times NO.{} : {}'.format(str(i+1),log.decode('utf-8')))
        #log=obj.CtrlB.read_until(b'0x00ff 0x00ff').decode('utf-8')
        #to be continued
        if b'0x00e0 0x00cf' in log:
            break
        else:
            i+=1
            if i==3:
                judge_Drv=False
                #return 1 #Driver is ALU

    #TIA check e0cf
    obj.CtrlB.flushInput()
    obj.CtrlB.flushOutput()
    print('TIA 连接性检查中！')
    j=0
    while j<4:
        time.sleep(0.2)
        obj.CtrlB.write(b'fs400_tia_read 0x0001\n')
        time.sleep(0.2)
        log=obj.CtrlB.read_all()#.decode('utf-8')
        print('TIA check times NO.{} : {}'.format(str(j+1),log.decode('utf-8')))
        #log=obj.CtrlB.read_until(b'0x00ff 0x00ff').decode('utf-8')
        if b'0x00e0 0x00cf' in log:
            break
        else:
            j+=1
            if j==3:
                judge_TIA=False
                # print('Please check whether TIA is powered on, driver is IDT and connected, TIA not connected!')
                # return 3
    if judge_Drv and judge_TIA:
        print('Driver is IDT')
        return 2 #Driver is IDT
    elif not judge_Drv and not judge_TIA and sn[1:3]=='AA':
        print('Driver is ALU')
        return 1 #Driver is ALU
    else:
        print('Please check whether TIA is powered on, driver is IDT and connected, TIA not connected!')
        return 3
    #to becontinued

def test_connectivity_new_ICR(obj,sn):
    '''
    update 0620:make sure ITLA is on to check connectivity and judge whether IDT or ALU
    Check pin connectivity
    :param obj:NA sn:SN to judge whether ALU or IDT
    :return: 1 # Driver is ALU
            2 # Driver is IDT
            3 # connectivity fail
    '''
    ##first board up
    board_set(obj,obj.board_up)
    print('控制板初始化完成！')
    print(obj.board_up)
    obj.CtrlB.flushInput()
    obj.CtrlB.flushOutput()
    # print('FS400连接性检查中！')
    # obj.CtrlB.write(b'fpga_spi_rd 0x4a\n');time.sleep(0.5)
    # ret=obj.CtrlB.read_all().decode('utf-8')
    # try:
    #     cur=int(re.split('[(,)]',ret)[-2])
    #     print('Connectivity test current Tx MPD X:',cur)
    # except Exception as e:
    #     print(e,'Current value get failed!')
    #     return 3 #current value get failed
    # obj.CtrlB.write(b'fpga_spi_rd 0x47\n');time.sleep(0.5)
    # ret = obj.CtrlB.read_all().decode('utf-8')
    # try:
    #     cur1 = int(re.split('[(,)]', ret)[-2])
    #     print('Connectivity test current Tx MPD Y:', cur1)
    # except Exception as e:
    #     print(e, 'Current value get failed!')
    #     return 3  # current value get failed
    # if not (cur>12000 and cur1>12000):
    #     print('MPD connection failed! please check!!!')
    #     return 3
    judge_Drv=True
    judge_TIA=True
    #Driver check e0cf
    print('Driver 连接性检查中！')
    i=0
    while i <4:
        time.sleep(0.2)
        obj.CtrlB.write(b'fs400_drv_read 0x0401\n')
        time.sleep(0.2)
        log=obj.CtrlB.read_all()#.decode('utf-8')
        print('Driver check times NO.{} : {}'.format(str(i+1),log.decode('utf-8')))
        #log=obj.CtrlB.read_until(b'0x00ff 0x00ff').decode('utf-8')
        #to be continued
        if b'0x00e0 0x00cf' in log:
            break
        else:
            i+=1
            if i==3:
                judge_Drv=False
                #return 1 #Driver is ALU

    #TIA check e0cf
    obj.CtrlB.flushInput()
    obj.CtrlB.flushOutput()
    print('TIA 连接性检查中！')
    j=0
    while j<4:
        time.sleep(0.2)
        obj.CtrlB.write(b'fs400_tia_read 0x0001\n')
        time.sleep(0.2)
        log=obj.CtrlB.read_all()#.decode('utf-8')
        print('TIA check times NO.{} : {}'.format(str(j+1),log.decode('utf-8')))
        #log=obj.CtrlB.read_until(b'0x00ff 0x00ff').decode('utf-8')
        if b'0x00e0 0x00cf' in log:
            break
        else:
            j+=1
            if j==3:
                judge_TIA=False
                # print('Please check whether TIA is powered on, driver is IDT and connected, TIA not connected!')
                # return 3
    if judge_Drv and judge_TIA:
        print('Driver is IDT')
        return 2 #Driver is IDT
    elif not judge_Drv and not judge_TIA and sn[1:3]=='AA':
        print('Driver is ALU')
        return 1 #Driver is ALU
    else:
        print('Please check whether TIA is powered on, driver is IDT and connected, TIA not connected!')
        return 3
    #to becontinued


def Tx_MPD_responsivity_Test(obj, bias):
    '''
    Update Tx MPD responsivity test here, input noise of X/Y,input power,
    then read the voltage code of Tx MPD X/Y then calculate the responsivity
    :param obj:NA
        noiX:Tx_MPD_X dark current
        noiY:Tx_MPD_Y dark current
        pwr:input light power of LO port(must confider the loss of path)
        bias:[] a list of lenth 6; change to indicate the current status for test,XmaxYmax/XmaxYmin/XminYmax
        R:Resistance of the sampling circult to calculate the I(Current of MPD)
    :return: v_Tx_MPDX,v_Tx_MPDY
    '''
    if serial.Serial.isOpen(obj.CtrlB):
        # Need fistly set the bias of all the tuners to MAX point(X max and Y min or X min and Ymax)
        if bias=='XmaxYmax':
            # Read Tx MPD X
            obj.CtrlB.write(b'fpga_spi_rd 0x47\n')
            time.sleep(0.1)
            s = obj.CtrlB.read_all().decode('utf-8')
            # left work here to deal with the output message,for example:
            s1=s.split('\r\n')[1]
            ind1=s1.index('(')
            ind2=s1.index(')')
            vCodeX = s1[ind1+1:ind2]
            # Read Tx MPD Y
            obj.CtrlB.write(b'fpga_spi_rd 0x4a\n')
            time.sleep(0.1)
            s = obj.CtrlB.read_all().decode('utf-8')
            # left work here to deal with the output message, for example:
            s1=s.split('\r\n')[1]
            ind1=s1.index('(')
            ind2=s1.index(')')
            vCodeY = s1[ind1+1:ind2]

        elif bias=='XmaxYmin':
            # Read Tx MPD X
            obj.CtrlB.write(b'fpga_spi_rd 0x47\n')
            time.sleep(0.1)
            s = obj.CtrlB.read_all().decode('utf-8')
            # left work here to deal with the output message,for example:
            s1=s.split('\r\n')[1]
            ind1=s1.index('(')
            ind2=s1.index(')')
            vCodeX = s1[ind1+1:ind2]
            # Read Tx MPD Y
            # obj.CtrlB.write(b'fpga_spi_rd 0x4a\n')
            # time.sleep(0.1)
            # s = obj.CtrlB.read_all().decode('utf-8')
            # left work here to deal with the output message, for example:
            vCodeY = 10 #No return in this situation
        elif bias=='XminYmax':
            # Read Tx MPD X
            #obj.CtrlB.write(b'fpga_spi_rd 0x47\n')
            #time.sleep(0.1)
            #s = obj.CtrlB.read_all().decode('utf-8')
            # left work here to deal with the output message,for example:
            vCodeX = 10 #No return in this situation
            # Read Tx MPD Y
            obj.CtrlB.write(b'fpga_spi_rd 0x4a\n')
            time.sleep(0.1)
            s = obj.CtrlB.read_all().decode('utf-8')
            # left work here to deal with the output message, for example:
            s1=s.split('\r\n')[1]
            ind1=s1.index('(')
            ind2=s1.index(')')
            vCodeY = s1[ind1+1:ind2]
        else:
            print('bias error input...')
            return
        # return the vcode
        return int(vCodeX),int(vCodeY)

'''
#***--------The code below is for FS400 PeSkew board control--------***
20220902
'''
def open_board_PeSkew(obj,board_port):
    '''
    Open board connection for PeSkew test
    :param obj:
    :param board_port:
    :return: True/False
    '''
    try:
        obj.CtrlB = serial.Serial(board_port,115200,timeout=2)
        # Check board status:
        obj.CtrlB.write(b'rdmsa 0xb016\r\n')
        time.sleep(0.2)
        r = obj.CtrlB.read_all().decode('utf-8')
        print(r)
        count = 0
        count1 = 0
        while not '0x0020' in r:
            time.sleep(5)
            count += 1
            print('The {} times read not ready, wait for 5s to read...'.format(count))
            obj.CtrlB.write(b'rdmsa 0xb016\r\n')
            time.sleep(0.2)
            r = obj.CtrlB.read_all().decode('utf-8')
            print(r)
            if count > 30:
                print('After 150s read not ready, please check...')
                return False
            if len(r) == 0:
                count1 += 1
                if count1 > 3:
                    print('控制板未上电！')
                    return False
        print('Board ready...')
        return True

        # #Update 09.02.2022
        # obj.CtrlB.write(b'fs400_stop_fs400_performance\n') #stop BER readings
        # i=0
        # count=0
        # while i<25:
        #     time.sleep(0.2)
        #     obj.CtrlB.write(b'\r\n')
        #     time.sleep(5)
        #     a=obj.CtrlB.read_all()
        #     print(a)
        #     if len(a)>9:
        #         count=0
        #     if len(a)==9:
        #         count+=1
        #         if count>2:
        #             print('控制板通信正常！')
        #             return True
        #     if len(a)==0:
        #         print('控制板未上电！')
        #         return False
        #     i+=1
        # print('读取120s超时，请检查控制板是否开启')
        # return False
    except Exception as e:
        print(e)
        return False

def open_board_VOAcal(obj,board_port):
    '''
    Open board connection for VOA calibration test, this version is for NianYu control board
    :param obj:
    :param board_port:
    :return: True/False
    '''
    try:
        obj.CtrlB = serial.Serial(board_port,115200,timeout=2)
        # #Update 09.02.2022
        obj.CtrlB.write(b'set_ignore_fs400 1\r\n') #stop VOA judgement
        time.sleep(0.2)
        #Check board status:
        obj.CtrlB.write(b'rdmsa 0xb016\r\n')
        time.sleep(0.2)
        r=obj.CtrlB.read_all().decode('utf-8',errors='ignore')
        print(r)
        count=0
        count1 = 0
        while not '0xB016: 0x0020' in r:
            time.sleep(5)
            count+=1
            print('The {} times read not ready, wait for 5s to read...'.format(count))
            obj.CtrlB.write(b'rdmsa 0xb016\r\n')
            time.sleep(0.2)
            r = obj.CtrlB.read_all().decode('utf-8',errors='ignore')
            print(r)
            if count>50:
                print('')
                print('After 250s read not ready, please check...')
                return False
            if len(r) == 0:
                count1 += 1
                if count1 > 3:
                    print('控制板未上电！')
                    return False
        print('Board ready...')
        return True

    except Exception as e:
        print(e)
        return False

def close_board_PeSkew(obj):
    '''
    Close board serial connection for PeSkew test
    :param obj:
    :return: NA
    '''
    try:
        if serial.Serial.isOpen(obj.CtrlB):
            obj.CtrlB.close()
            print('FS400控制板串口关闭!')
        else:
            print('FS400控制板已是关闭状态!')
    except Exception as e:
        print(e)

def read_comd_PeSkew(file):
    '''
    Read command from txt file
    :param file:
    :return: valid command list with "\p" stands for pause 1 sec(time.sleep(1))
    '''
    with open(file,'r',encoding='utf-8') as f:
        comd=[i for i in f.readlines() if not '//' in i and not '->' in i][:-1]
        for i in comd:
            if i =='\n':
                comd.remove(i)
            if not '\r' in i:
                comd[comd.index(i)]=i+'\r'
    return comd

def board_set_PeSkew(obj,f):
    '''
    Board up/drv up/drv down for PeSkew test
    :param obj: GUI Class obj
    :param f: file absolote path(including three status)
    :return: Ture/False
    '''
    obj.CtrlB.flushInput()
    obj.CtrlB.flushOutput()
    obj.CtrlB.timeout=30
    try:
        if "brdup" in f:
            stri="FS400 board set!"
            wai=b'en_adj 4 1'
        elif "driverdown" in f:
            stri="Driver power off!"
        elif "driverup" in f:
            stri="Driver power on!"
        colist=read_comd_PeSkew(f)
        for i in colist:
            if i=='\\p\n':
                time.sleep(1)
                continue
            if 'itla_wr' in i:# 0 0x32 0x08' in i:
                obj.CtrlB.write(i.encode('utf-8'))
                time.sleep(3)
                a=obj.CtrlB.read_until(b'itla_0_write')
                time.sleep(0.5)
            else:
                obj.CtrlB.write(i.encode('utf-8'))
                time.sleep(0.1)
                a=obj.CtrlB.read_all()
            time.sleep(0.1)
            print(a)

        obj.CtrlB.timeout=2
        print(stri)
    except Exception as e:
        print(e)

def board_set_CmdList_DC(obj,lis):
    '''
    list of command to batch config to the control board of DC/Tx bw/ICR
    :param obj: GUI Class obj
    :param lis: For example-['itla_write 0 0x32 0','itla_write 0 0x34 0x2ee']
                will convert to bytes and add '\n\r' in the end
    :return: NA
    '''
    obj.CtrlB.flushInput()
    obj.CtrlB.flushOutput()
    obj.CtrlB.timeout=60
    try:
        for i in lis:
            i=i.strip()
            i+='\n\r'
            if 'itla_wr' in i:# 0 0x32 0x08' in i:
                obj.CtrlB.write(i.encode('utf-8'))
                time.sleep(2)
                a=obj.CtrlB.read_until(b'Write itla')
                time.sleep(0.5)
            else:
                obj.CtrlB.write(i.encode('utf-8'))
                time.sleep(0.5)
                a=obj.CtrlB.read_all()
            #time.sleep(0.1)
            print(a)
        obj.CtrlB.timeout=2
    except Exception as e:
        print(e)

def board_set_CmdList(obj,lis):
    '''
    list of command to batch config to the control board
    :param obj: GUI Class obj
    :param lis: For example-['itla_write 0 0x32 0','itla_write 0 0x34 0x2ee']
                will convert to bytes and add '\n\r' in the end
    :return: NA
    '''
    obj.CtrlB.flushInput()
    obj.CtrlB.flushOutput()
    obj.CtrlB.timeout=60
    try:
        for i in lis:
            i=i.strip()
            i+='\n\r'
            if 'itla_wr' in i:# 0 0x32 0x08' in i:
                obj.CtrlB.write(i.encode('utf-8'))
                time.sleep(2)
                a=obj.CtrlB.read_until(b'itla_0_write')
                time.sleep(0.5)
            else:
                obj.CtrlB.write(i.encode('utf-8'))
                time.sleep(0.5)
                a=obj.CtrlB.read_all()
            #time.sleep(0.1)
            print(a)
        obj.CtrlB.timeout=2
    except Exception as e:
        print(e)

def get_VCC_VDC(obj):
    '''
    Get the result of VCCout and VDCout
    :param obj:
    :return:
    '''
    #obj.CtrlB.read_all()
    drv_VCCout=0
    drv_VDCout=0
    obj.CtrlB.write(b'fpga_ad_get_value 255\r\n')
    # time.sleep(0.2)
    r = obj.CtrlB.readline().decode('utf-8')
    print(r)
    count=0
    while not 'ADC_DRV_VDCOUT' in r:
        if 'ADC_DRV_VCCOUT' in r:
            tem=r.split('=')[1]
            drv_VCCout=int(tem[0:tem.index('mV')].strip())
        r = obj.CtrlB.readline().decode('utf-8')
        print(r)
        count+=1
        if count>40:
            print('VCCout get failed, please check!')
            break
    if 'ADC_DRV_VDCOUT' in r:
        tem = r.split('=')[1]
        drv_VDCout = int(tem[0:tem.index('mV')].strip())
    return(drv_VCCout,drv_VDCout)

def drv_VCC_set_EEPROM_old(obj,volmin=2200,volmax=2300,limit=3195,step=40):
    '''
    NianYu board to auto set DRV VCCout to make DRV VDCout to a defined range
    Remark:For ALU test, please revise the limit when the value pass in
    :param obj: GUI Class obj
    :param volmin: v limit min(unit:mV)
    :param volmax: v limit max(unit:mV)
    :param limit:VCC out code to set limit max(unit:NA,code)
    :param step: step of adjust code(unit:NA,code)
    :return: NA
    '''
    obj.CtrlB.flushInput()
    obj.CtrlB.flushOutput()
    # obj.CtrlB.timeout=60
    try:
        #Firstly get start point
        obj.CtrlB.write(b'show_fpga_da\r\n')
        # time.sleep(0.2)
        r=obj.CtrlB.read_until(b'DAC_TIA_GC_YQ').decode('utf-8')
        # print(r.split('\n'))
        # print([i for i in r.split('\n') if 'DAC_DRV_VCCOUT' in i ])
        resu=[int(i.strip().split(' '*4)[1]) for i in r.split('\n') if 'DAC_DRV_VCCOUT' in i ]
        vccCode=resu[0]
        #get VCC VDC voltage
        drv_VCCout, drv_VDCout=get_VCC_VDC(obj)
        print('Vcode read is {}'.format(vccCode))
        print('VCCout read is {}'.format(drv_VCCout))
        print('VDCout read is {}'.format(drv_VDCout))
        if drv_VDCout==0 or drv_VCCout==0:
            return
        while not volmin<=drv_VDCout<=volmax:
            if drv_VDCout<volmin:
                vccCode+=step
                if vccCode>limit:
                    print('can`t get voltage within {} code range, please check if the limit is too low!maybe ALU device?')
                    return
                obj.CtrlB.write('fpga_da_set_value 19 {}\r\n'.format(str(vccCode)).encode('utf-8'))
            elif drv_VDCout>volmax:
                vccCode -= step
                obj.CtrlB.write('fpga_da_set_value 19 {}\r\n'.format(str(vccCode)).encode('utf-8'))
            time.sleep(0.2)
            #set vcode
            drv_VCCout, drv_VDCout=get_VCC_VDC(obj)
            print('Vcode set as {}'.format(vccCode))
            print('VCCout set as {}'.format(drv_VCCout))
            print('VDCout read is {}'.format(drv_VDCout))

        #get the value and write to eeprom
        page=[7,8,9,11,13,15]
        ind='0'
        for i in page:
            if i==7 or i==11 or i==13 or i==15:
                ind='0'
            elif i==8:
                ind='1'
            elif i==9:
                ind='2'
            print('Write page {}'.format(i))
            obj.CtrlB.write('fs400_eep_arg_read {}\r\n'.format(str(i)).encode('utf-8'))
            s=obj.CtrlB.read_until(b'rfpd_adc_vol')
            print(s)
            obj.CtrlB.write('fs400_set_eep_drv_config4 {} 0x01 6000 {}\r\n'.format(ind,str(drv_VCCout)).encode('utf-8'))
            time.sleep(0.4)
            s = obj.CtrlB.read_all()
            print(s)
            obj.CtrlB.write('fs400_eep_arg_save {}\r\n'.format(str(i)).encode('utf-8'))
            s = obj.CtrlB.read_until(b'ok!\r\n')
            print(s)
            obj.CtrlB.write('fs400_eep_arg_read {}\r\n'.format(str(i)).encode('utf-8'))
            s = obj.CtrlB.read_until(b'rfpd_adc_vol')
            print(s)
            #print(s)
    except Exception as e:
        print(e)

def drv_VCC_set_EEPROM_fixVPN(obj,volmin=2200,volmax=2300,limit=3195,step=40):
    '''
    NianYu board to auto set DRV VCCout to make DRV VDCout to a defined range
    Remark:For ALU test, please revise the limit when the value pass in
    :param obj: GUI Class obj
    :param volmin: v limit min(unit:mV)
    :param volmax: v limit max(unit:mV)
    :param limit:VCC out code to set limit max(unit:NA,code)
    :param step: step of adjust code(unit:NA,code)
    :return: NA
    '''
    obj.CtrlB.flushInput()
    obj.CtrlB.flushOutput()
    # obj.CtrlB.timeout=60
    try:
        #Firstly get start point
        obj.CtrlB.write(b'show_fpga_da\r\n')
        # time.sleep(0.2)
        r=obj.CtrlB.read_until(b'DAC_TIA_GC_YQ').decode('utf-8', errors='ignore')
        # print(r.split('\n'))
        # print([i for i in r.split('\n') if 'DAC_DRV_VCCOUT' in i ])
        resu=[int(i.strip().split(' '*4)[1]) for i in r.split('\n') if 'DAC_DRV_VCCOUT' in i ]
        vccCode=resu[0]
        #get VCC VDC voltage
        drv_VCCout, drv_VDCout=get_VCC_VDC(obj)
        print('Vcode read is {}'.format(vccCode))
        print('VCCout read is {}'.format(drv_VCCout))
        print('VDCout read is {}'.format(drv_VDCout))
        if drv_VDCout==0 or drv_VCCout==0 or vccCode==0:
            print('vccCode,drv_VDCout or drv_VCCout equal 0, attention!!')
            return False
        while not volmin<=drv_VDCout<=volmax:
            if drv_VDCout<volmin:
                vccCode+=step
                if vccCode>limit:
                    print('can`t get voltage within code range, please check if the limit is too low!maybe ALU device?')
                    return False
                obj.CtrlB.write('fpga_da_set_value 19 {}\r\n'.format(str(vccCode)).encode('utf-8'))
            elif drv_VDCout>volmax:
                vccCode -= step
                obj.CtrlB.write('fpga_da_set_value 19 {}\r\n'.format(str(vccCode)).encode('utf-8'))
                if vccCode<limit-2000:
                    print('can`t get voltage within code range, vcc Code below {}, please check it.'.format(limit-2000))
                    return False
            time.sleep(0.2)
            #set vcode
            drv_VCCout, drv_VDCout=get_VCC_VDC(obj)
            print('Vcode set as {}'.format(vccCode))
            print('VCCout set as {}'.format(drv_VCCout))
            print('VDCout read is {}'.format(drv_VDCout))

        #get the value and write to eeprom
        page=[7,8,9,11,13,15]
        ind='0'
        for i in page:
            if i==7 or i==11 or i==13 or i==15:
                ind='0'
            elif i==8:
                ind='1'
            elif i==9:
                ind='2'
            print('Write page {}'.format(i))
            obj.CtrlB.write('fs400_eep_arg_read {}\r\n'.format(str(i)).encode('utf-8'))
            s=obj.CtrlB.read_until(b'rfpd_adc_vol')
            print(s)
            obj.CtrlB.write('fs400_set_eep_drv_config4 {} 0x01 6000 {}\r\n'.format(ind,str(drv_VCCout)).encode('utf-8'))
            time.sleep(0.4)
            s = obj.CtrlB.read_all()
            print(s)
            obj.CtrlB.write('fs400_eep_arg_save {}\r\n'.format(str(i)).encode('utf-8'))
            s = obj.CtrlB.read_until(b'ok!\r\n')
            print(s)
            obj.CtrlB.write('fs400_eep_arg_read {}\r\n'.format(str(i)).encode('utf-8'))
            s = obj.CtrlB.read_until(b'rfpd_adc_vol')
            print(s)
        return True
    except Exception as e:
        print(e)
        return False

def drv_VCC_set_EEPROM(obj,volmin=2200,volmax=2300,limit=3195,step=40):
    '''
    NianYu board to auto set DRV VCCout to make DRV VDCout to a defined range
    Remark:For ALU test, please revise the limit when the value pass in
    :param obj: GUI Class obj
    :param volmin: v limit min(unit:mV)
    :param volmax: v limit max(unit:mV)
    :param limit:VCC out code to set limit max(unit:NA,code)
    :param step: step of adjust code(unit:NA,code)
    :return: NA
    '''
    obj.CtrlB.flushInput()
    obj.CtrlB.flushOutput()
    # obj.CtrlB.timeout=60
    try:
        #Firstly get start point
        obj.CtrlB.write(b'show_fpga_da\r\n')
        # time.sleep(0.2)
        r=obj.CtrlB.read_until(b'DAC_TIA_GC_YQ').decode('utf-8', errors='ignore')
        # print(r.split('\n'))
        # print([i for i in r.split('\n') if 'DAC_DRV_VCCOUT' in i ])
        resu=[int(i.strip().split(' '*4)[1]) for i in r.split('\n') if 'DAC_DRV_VCCOUT' in i ]
        vccCode=resu[0]
        #get VCC VDC voltage
        drv_VCCout, drv_VDCout=get_VCC_VDC(obj)
        print('Vcode read is {}'.format(vccCode))
        print('VCCout read is {}'.format(drv_VCCout))
        print('VDCout read is {}'.format(drv_VDCout))
        if drv_VDCout==0 or drv_VCCout==0 or vccCode==0:
            print('vccCode,drv_VDCout or drv_VCCout equal 0, attention!!')
            return False
        while not volmin<=drv_VDCout<=volmax:
            if drv_VDCout<volmin:
                vccCode+=step
                if vccCode>limit:
                    print('can`t get voltage within code range, please check if the limit is too low!maybe ALU device?')
                    return False
                obj.CtrlB.write('fpga_da_set_value 19 {}\r\n'.format(str(vccCode)).encode('utf-8'))
            elif drv_VDCout>volmax:
                vccCode -= step
                obj.CtrlB.write('fpga_da_set_value 19 {}\r\n'.format(str(vccCode)).encode('utf-8'))
                if vccCode<limit-2000:
                    print('can`t get voltage within code range, vcc Code below {}, please check it.'.format(limit-2000))
                    return False
            time.sleep(0.2)
            #set vcode
            drv_VCCout, drv_VDCout=get_VCC_VDC(obj)
            print('Vcode set as {}'.format(vccCode))
            print('VCCout set as {}'.format(drv_VCCout))
            print('VDCout read is {}'.format(drv_VDCout))

        #get the value and write to eeprom
        page=[7,8,9,11,13,15]
        ind='0'
        for i in page:
            if i==7 or i==11 or i==13 or i==15:
                ind='0'
            elif i==8:
                ind='1'
            elif i==9:
                ind='2'
            print('Write page {}'.format(i))
            obj.CtrlB.write('fs400_eep_arg_read {}\r\n'.format(str(i)).encode('utf-8'))
            s=obj.CtrlB.read_until(b'rfpd_adc_vol').decode('utf-8', errors='ignore')
            print(s)
            time.sleep(1)
            drv_OA,drv_VPN=[i.split(':')[1].strip() for i in s.split('\r\n') if 'drv_output_amp' in i or 'drv_vpn' in i]
            print('\n\n*******************Read DRV OA,VPN******************')
            print('drv_OA: '+drv_OA)
            print('drv_VPN: ' + drv_VPN)
            obj.CtrlB.write('fs400_set_eep_drv_config4 {} {} {} {}\r\n'.format(ind,drv_OA,drv_VPN,str(drv_VCCout)).encode('utf-8'))
            time.sleep(1)
            s = obj.CtrlB.read_all()
            print(s)
            obj.CtrlB.write('fs400_eep_arg_save {}\r\n'.format(str(i)).encode('utf-8'))
            s = obj.CtrlB.read_until(b'ok!\r\n')
            print(s)
            obj.CtrlB.write('fs400_eep_arg_read {}\r\n'.format(str(i)).encode('utf-8'))
            s = obj.CtrlB.read_until(b'rfpd_adc_vol')
            print(s)
        return True
    except Exception as e:
        print(e)
        return False

def drv_VCC_Monitor_EEPROM(obj, file_store, delay=60, cycle=100, volmin=2200, volmax=2300, limit=3195, step=40):
    '''
    Monitor the Drv VCC and VDC after adjusted to stable status...
    NianYu board to auto set DRV VCCout to make DRV VDCout to a defined range
    Remark:For ALU test, please revise the limit when the value pass in
    :param obj: GUI Class obj
    :param file_store:the absolute path of the file to store the data:For example,'c:\\test\DRV.csv'
    :param delay:The time delay set in the loop
    :param cycle:Cycle of the loop monitor the result
    :param volmin: v limit min(unit:mV)
    :param volmax: v limit max(unit:mV)
    :param limit:VCC out code to set limit max(unit:NA,code)
    :param step: step of adjust code(unit:NA,code)
    :return: NA
    '''
    import pandas as pd
    obj.CtrlB.flushInput()
    obj.CtrlB.flushOutput()
    # obj.CtrlB.timeout=60
    try:
        # Firstly get start point
        obj.CtrlB.write(b'show_fpga_da\r\n')
        # time.sleep(0.2)
        r = obj.CtrlB.read_until(b'DAC_TIA_GC_YQ').decode('utf-8', errors='ignore')
        # print(r.split('\n'))
        # print([i for i in r.split('\n') if 'DAC_DRV_VCCOUT' in i ])
        resu = [int(i.strip().split(' ' * 4)[1]) for i in r.split('\n') if 'DAC_DRV_VCCOUT' in i]
        vccCode = resu[0]
        # get VCC VDC voltage
        drv_VCCout, drv_VDCout = get_VCC_VDC(obj)
        print('Vcode read is {}'.format(vccCode))
        print('VCCout read is {}'.format(drv_VCCout))
        print('VDCout read is {}'.format(drv_VDCout))
        if drv_VDCout == 0 or drv_VCCout == 0 or vccCode == 0:
            print('vccCode,drv_VDCout or drv_VCCout equal 0, attention!!')
            return False
        while not volmin <= drv_VDCout <= volmax:
            if drv_VDCout < volmin:
                vccCode += step
                if vccCode > limit:
                    print(
                        'can`t get voltage within code range, please check if the limit is too low!maybe ALU device?')
                    return False
                obj.CtrlB.write('fpga_da_set_value 19 {}\r\n'.format(str(vccCode)).encode('utf-8'))
            elif drv_VDCout > volmax:
                vccCode -= step
                obj.CtrlB.write('fpga_da_set_value 19 {}\r\n'.format(str(vccCode)).encode('utf-8'))
                if vccCode < limit - 2000:
                    print('can`t get voltage within code range, vcc Code below {}, please check it.'.format(
                        limit - 2000))
                    return False
            time.sleep(0.2)
            # set vcode
            drv_VCCout, drv_VDCout = get_VCC_VDC(obj)
            print('Vcode set as {}'.format(vccCode))
            print('VCCout set as {}'.format(drv_VCCout))
            print('VDCout read is {}'.format(drv_VDCout))
        # monitor the value
        DRV_VCCmonitor = pd.DataFrame(columns=('Time', 'DRV_VCC', 'DRV_VDC'))
        time_stamp = 0
        for i in range(cycle):
            time.sleep(delay)
            time_stamp += delay
            drv_VCCout, drv_VDCout = get_VCC_VDC(obj)
            row_data = [time_stamp, drv_VCCout, drv_VDCout]
            print(row_data)
            DRV_VCCmonitor.loc[i] = row_data
            DRV_VCCmonitor.to_csv(file_store)
        return True
    except Exception as e:
        print(e)
        return False

def test_connectivity_PeSkew(obj,sn):
    '''
    PeSkew connectivity test on the base of original one
    update 0620:make sure ITLA is on to check connectivity and judge whether IDT or ALU
    Check pin connectivity
    :param obj:NA sn:SN to judge whether ALU or IDT
    :return: 1 # Driver is ALU
            2 # Driver is IDT
            3 # connectivity fail
    '''
    ##first board up
    #board_set_PeSkew(obj,obj.board_up)
    #print('控制板初始化完成！')
    #print(obj.board_up)
    # obj.CtrlB.flushInput()
    # obj.CtrlB.flushOutput()
    print('FS400连接性检查中！')
    #judge LO light through
    #judge_lo=False
    #judge LO connectivity
    #print('check LO light connect, start op scan...')
    #obj.CtrlB.timeout=60
    #obj.CtrlB.write(b'scan_op 0 0 5000 20000 500\r\n')
    #time.sleep(0.1)
    #ss=obj.CtrlB.read_until(b'test end').decode('utf-8').split('\n')[7:-1]
    #print(ss)
    # ind=0
    # for i in ss:
    #     mpdX_c=int(i.replace('\r', ' ').strip().split(' ')[1])
    #     if mpdX_c>10500:
    #         print('LO connectivity test ok, Tx MPD X vcode:',mpdX_c)
    #         ind+=1
    #         break
    # if ind==0:
    #     print('LO connectivity test failed, Tx MPD X vcode:', mpdX_c)
    #     return 3
    #work left here to do
    # cmClose = b'itla_wr 0 0x32 0x00\n\r'
    # cmOpen = b'itla_wr 0 0x32 0x08\n\r'
    # mawin.CtrlB.write(cmOpen);
    # time.sleep(0.1)
    # print(mawin.CtrlB.read_until(b'itla_0_write'))
    # #check lo
    # #TXPDX_IN_PMON 4:
    # if not judge_lo:
    #     print('请检查LO光路连接...')
    #     return 3
    # mawin.CtrlB.write(cmClose);
    # time.sleep(0.1)
    # print(mawin.CtrlB.read_until(b'Write itla'))
    # #check lo
    # if judge_lo:
    #     print('请检查光路连接,可能LO Rx接反了...')
    #     return 3
    judge_TIA=True
    #TIA check e0cf
    obj.CtrlB.flushInput()
    obj.CtrlB.flushOutput()
    print('TIA 连接性检查中！')
    j=0
    while j<4:
        time.sleep(0.2)
        obj.CtrlB.write(b'fs400_tia_read 0\r\n')
        time.sleep(0.5)
        log=obj.CtrlB.read_all()#.decode('utf-8')
        print('TIA check times NO.{} : {}'.format(str(j+1),log.decode('utf-8')))
        #log=obj.CtrlB.read_until(b'0x00ff 0x00ff').decode('utf-8')
        if b'0xe0cf' in log:
            break
        else:
            j+=1
            if j==3:
                judge_TIA=False
    if judge_TIA: print('TIA is IDT...')
    ## judge Rx light through
    # left work here
    obj.CtrlB.read_all()
    obj.CtrlB.timeout=20
    judge_RxLight = False
    obj.CtrlB.write(b'fpga_ad_get_value 255\r\n')
    # Rx MPD 5mv no light, light 1280mv,spec 100mv
    #time.sleep(0.2)
    ret = obj.CtrlB.read_until(b'ADC_DRV_VDCOUT').decode('utf-8')
    print('RxMPD information:\n', ret)
    # tmp1=re.split('[:,\r]',tmp[ind+1]+tmp[ind+2])
    ret_1 = ret.split('\n')
    sta_list = int([i[i.index('=') + 2:i.index('mV')] for i in ret_1 if 'ADC_RXMPDXY' in i][0])
    print('PD current get: {}mv'.format(sta_list))
    if sta_list>100:judge_RxLight=True

    if judge_RxLight and judge_TIA:
        print('Device is IDT and Rx light path is OK')
        return 2 #Driver is IDT
    elif judge_RxLight and not judge_TIA:
        if sn[1:3]=='AA':
            print('Device is ALU and Rx light path is OK')
            return 1 #Driver is ALU
        else:
            print('Device is IDT and TIA not connected, please check pin connect.')
            return 3
    else:
        print('Please check whether TIA is connected, and whether Rx light path is connected!')
        return 3
    #to becontinued

def test_connectivity_VoaCalibration(obj,sn):
    '''
    VOA connectivity test on the base of original one
    update 0620:make sure ITLA is on to check connectivity and judge whether IDT or ALU
    Check pin connectivity
    :param obj:NA sn:SN to judge whether ALU or IDT
    :return: 1 # Driver is ALU
            2 # Driver is IDT
            3 # connectivity fail
    '''
    ##first board up
    #board_set_PeSkew(obj,obj.board_up)
    #print('控制板初始化完成！')
    print(obj.board_up)
    obj.CtrlB.flushInput()
    obj.CtrlB.flushOutput()
    print('FS400连接性检查中！')

    judge_TIA=True
    #TIA check e0cf
    obj.CtrlB.flushInput()
    obj.CtrlB.flushOutput()
    print('TIA 连接性检查中！')
    j=0
    while j<4:
        time.sleep(0.2)
        obj.CtrlB.write(b'fs400_tia_read 0\r\n')
        time.sleep(0.2)
        log=obj.CtrlB.read_all()#.decode('utf-8')
        print('TIA check times NO.{} : {}'.format(str(j+1),log.decode('utf-8')))
        #log=obj.CtrlB.read_until(b'0x00ff 0x00ff').decode('utf-8')
        if b'0xe0cf' in log:
            break
        else:
            j+=1
            if j==3:
                judge_TIA=False
    if judge_TIA:print('TIA is IDT...')
    #judge LO connectivity
    #print('check LO light connect, start op scan...')

    # obj.CtrlB.timeout=100
    # obj.CtrlB.write(b'scan_op 0 0 5000 30500 500\r\n')
    # time.sleep(0.2)
    # ss=obj.CtrlB.read_until(b'test end').decode('utf-8').split('\n')[7:-1]
    # while ss==[]:
    #     obj.CtrlB.write(b'scan_op 0 0 5000 31000 500\r\n')
    #     time.sleep(0.2)
    #     ss = obj.CtrlB.read_until(b'test end').decode('utf-8').split('\n')[7:-1]
    # print(ss)
    # ind=0
    # for i in ss:
    #     mpdX_c=int(i.replace('\r', ' ').strip().split(' ')[1])
    #     if mpdX_c>10500:
    #         print('LO connectivity test ok, Tx MPD X vcode:',mpdX_c)
    #         ind+=1
    #         break
    # if ind==0:
    #     print('LO connectivity test failed, Tx MPD X vcode:', mpdX_c)
    #     return 3


    # ## judge Rx light through
    # #left work here
    # judge_RxLight=False
    # obj.CtrlB.write(b'fpga_ad_get_value 255\r\n')
    # #Rx MPD 5mv no light, light 1280mv,spec 100mv
    # time.sleep(0.2)
    # ret = obj.CtrlB.read_all().decode('utf-8')
    # print('RxMPD information:\n', ret)
    # # tmp1=re.split('[:,\r]',tmp[ind+1]+tmp[ind+2])
    # ret_1 = ret.split('\n')
    # sta_list = int([i[i.index('=') + 2:i.index('mV')] for i in ret_1 if 'ADC_RXMPDXY' in i][0])
    # print('PD current get: {}mv'.format(sta_list))
    # #Judge connectivity
    # if sta_list>100:
    #     judge_RxLight=True
    #     print('Rx light path Connectivity test OK!')
    if judge_TIA:
        print('Device is IDT and connect OK')
        return 2 #Driver is IDT
    elif not judge_TIA:
        if sn[1:3]=='AA':
            print('Device is ALU and connect OK')
            return 1 #Driver is ALU
        else:
            print('Device is IDT and TIA not connected, please check pin connect.')
            return 3
    # else:
    #     print('Please check whether TIA is connected, and whether Rx light path is connected!')
    #     return 3
    # #to becontinued
