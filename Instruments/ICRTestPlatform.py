# -*- coding: utf-8 -*-
# this module is to write the control functions of Test Platform
# date:4/14/2022
# author:Jiawen Zhou
import time
import pyvisa
import numpy as np

rm=pyvisa.ResourceManager()
rm.list_resources()

def open_ICRtf(obj,ICRtf_port):
    '''
    Open ICR test platform connection
    :param obj: GUI Class obj
    :param ICRtf_port: ICR port number
    :return: True/False
    '''
    try:
        obj.ICRtf = rm.open_resource(ICRtf_port)
        obj.ICRtf.timeout=2000
        ins=obj.ICRtf.query('*IDN?')
        if "Coherent Solutions" in ins:
            return True
        else:
            return False
    except Exception as e:
        print(e)

def close_ICR(obj):
    '''
    Close ICR test platform connection
    :param obj: GUI Class obj
    :return: NA
    '''
    try:
        if obj.ICRtf in rm.list_opened_resources():
            obj.ICRtf.close()
            print('ICR平台端口关闭成功!')
    except Exception as e:
        print(e)

def set_laser(obj,ch,p0,pout,lamb=1550.008):
    '''
    Set ICR test platform laser power by controlling the laser and VOA, finally open the laser light
    :param obj: GUI Class obj
    :param ch: channel of laser
    :param p0:
    :param pout:
    :param lamb: wavelength corresponding to the frequency
    :return: NA
    '''
    p1=float(str(obj.ICRtf.query('SOUR1:CHAN{}:POW? MIN'.format(str(ch)))).strip())#;time.sleep(0.1)
    p2=float(str(obj.ICRtf.query('SOUR1:CHAN{}:POW? MAX'.format(str(ch)))).strip())#;time.sleep(0.1)
    lamb1=float(str(obj.ICRtf.query('SOUR1:CHAN{}:WAV? MIN'.format(str(ch)))).strip())*1e9#;time.sleep(0.1)
    lamb2=float(str(obj.ICRtf.query('SOUR1:CHAN{}:WAV? MAX'.format(str(ch)))).strip())*1e9#;time.sleep(0.1)

    if p0<p1 or p0>p2:
        print('error:Power of Laser is not right')
        return
    if lamb<lamb1 or lamb>lamb2:
        print('error:Lambda of Laser is not right')
        return
    if pout>p0:
        print('error:Output power of VOA is not right')
        return
    #Open the laser and set wavelength first
    obj.ICRtf.write('OUTP1:CHAN{}:STATE ON'.format(str(ch)))
    time.sleep(3)
    obj.ICRtf.write('SOUR1:CHAN{}:WAV {}e-9'.format(str(ch),str(lamb)))#;time.sleep(0.1)

    #Set power now
    obj.ICRtf.write('CONT3:CHAN{}:MODE POW'.format(str(ch)))#;time.sleep(0.1)
    data1=str(obj.ICRtf.query('CONT3:CHAN{}:MODE?'.format(str(ch)))).strip()#;time.sleep(0.1)
    while data1!='POWER':
        obj.ICRtf.write('CONT3:CHAN{}:MODE POW'.format(str(ch)))#;time.sleep(0.1)
        data1=str(obj.ICRtf.query('CONT3:CHAN{}:MODE?'.format(str(ch)))).strip()#;time.sleep(0.1)
    # set p0 power,namely the output of laser1
    obj.ICRtf.write('SOUR1:CHAN{}:POW {}'.format(str(ch),str(p0)))#;time.sleep(0.1)
    time.sleep(3)
    data2=float(str(obj.ICRtf.query('SOUR1:CHAN{}:POW?'.format(str(ch)))).strip())#;time.sleep(0.1)
    while abs(data2-float(p0))>0.05:
        time.sleep(0.5)
    #   obj.ICRtf.write('SOUR1:CHAN{}:POW {}'.format(str(ch),str(p0)));time.sleep(0.1)
        data2=float(str(obj.ICRtf.query('SOUR1:CHAN{}:POW?'.format(str(ch)))).strip())#;time.sleep(0.1)
    # set pout power,namely the output of VOA
    obj.ICRtf.write('OUTP3:CHAN{}:POW {}'.format(str(ch),str(pout)))#;time.sleep(0.1)
    data3=float(str(obj.ICRtf.query('OUTP3:CHAN{}:POW?'.format(str(ch)))).strip())#;time.sleep(0.1)
    while abs(data3-float(pout))>0.05:
        time.sleep(0.5)
    #     obj.ICRtf.write('OUTP3:CHAN{}:POW {}'.format(str(ch),str(pout)));time.sleep(0.1)
        data3=float(str(obj.ICRtf.query('OUTP3:CHAN{}:POW?'.format(str(ch)))).strip())#;time.sleep(0.1)

    #set the wavelength
    #obj.ICRtf.write('SOUR1:CHAN1:WAV 1550.008e-9');time.sleep(0.1)
    #obj.ICRtf.write('SOUR1:CHAN{}:WAV {}e-9'.format(str(ch),str(lamb)));time.sleep(0.1)
    #obj.ICRtf.write('OUTP1:CHAN{}:STATE ON'.format(str(ch)));time.sleep(0.1)

def set_wavelength(obj,ch,wl,offset):
    '''
    Set ICR test platform laser wavelength according to the set value and offset
    :param obj: GUI Class obj
    :param ch: channel of laser
    :param wl: wavelength
    :param offset: times 0.008nm
    :return: NA
    '''
    #74 channels corresponding to channel NO. selected
    # b=[1571.960,1571.342,1570.725,1570.108,1569.491,1568.875,1568.26,1567.645,1567.0301630,1566.4160830,1565.8024840,1565.1893650,1564.5767260,1563.9645670,1563.3528870,1562.7416850,1562.1309610,1561.5207140,1560.9109430,1560.3016490,1559.6928300,1559.0844850,1558.4766160,1557.8692200,1557.2622970,1556.6558470,1556.0498700,1555.4443630,1554.8393280,1554.2347640,1553.6306690,1553.0270440,1552.4238880,1551.8212000,1551.2189790,1550.6172260,1550.0159400,1549.4151200,1548.8147650,1548.2148760,1547.6154510,1547.0164900,1546.4179920,1545.8199570,1545.2223850,1544.6252750,1544.0286260,1543.4324370,1542.8367090,1542.2414400,1541.6466310,1541.0522800,1540.4583880,1539.8649530,1539.2719750,1538.6794530,1538.0873880,1537.4957780,1536.9046230,1536.3139220,1535.7236750,1535.1338820,1534.5445420,1533.9556530,1533.3672170,1532.7792320,1532.1916970,1531.6046130,1531.0179790,1530.4317940,1529.8460570,1529.2607690,1528.6759280,1528.0915350]
    # c=[i*1e-9 for i in b]
    # d=c(int(wl)-1)+offset*0.008e-9
    d=cal_wavelength(wl,offset)
    obj.ICRtf.write('SOUR1:CHAN{}:WAV {}'.format(str(ch),str(d)))
    time.sleep(5)
    print('Laser channel {} wavelength set as {}.'.format(str(ch),str(d)))

def cal_wavelength(i,j):
    '''
    calculate the wavelength with offset
    :param i: channel
    :param j: offset
    :return: wavelength
    '''
    b=[1571.960,1571.342,1570.725,1570.108,1569.491,1568.875,1568.26,1567.645,1567.0301630,1566.4160830,1565.8024840,1565.1893650,1564.5767260,1563.9645670,1563.3528870,1562.7416850,1562.1309610,1561.5207140,1560.9109430,1560.3016490,1559.6928300,1559.0844850,1558.4766160,1557.8692200,1557.2622970,1556.6558470,1556.0498700,1555.4443630,1554.8393280,1554.2347640,1553.6306690,1553.0270440,1552.4238880,1551.8212000,1551.2189790,1550.6172260,1550.0159400,1549.4151200,1548.8147650,1548.2148760,1547.6154510,1547.0164900,1546.4179920,1545.8199570,1545.2223850,1544.6252750,1544.0286260,1543.4324370,1542.8367090,1542.2414400,1541.6466310,1541.0522800,1540.4583880,1539.8649530,1539.2719750,1538.6794530,1538.0873880,1537.4957780,1536.9046230,1536.3139220,1535.7236750,1535.1338820,1534.5445420,1533.9556530,1533.3672170,1532.7792320,1532.1916970,1531.6046130,1531.0179790,1530.4317940,1529.8460570,1529.2607690,1528.6759280,1528.0915350]
    c=[i*1e-9 for i in b]
    d=c[int(i)-1]+j*0.008e-9
    return d

def TIA_tra_data(data):
    '''
    PD current handle to decimal value from hex
    :param data:
    :return: PD current decimal value
    '''
    try:
        if not 'ok' in data:
            print('TIA PD current return error, please check manually!!')
            return
        ind=data.index('ok')
        return int(data[ind-9:ind-3],16)
    except Exception as e:
        print(e)
        return 0


def TIA_getPDcurrent(obj):
    '''
    get TIA current signal port
    :parameter:GUI Class obj
    :return:[current_x,current_y,VPDX,VPDY]
    '''
    try:
        #DAC开启，PD上电
        # obj.CtrlB.write(i.encode('utf-8'))
        #print('---Rx signal port PD current---')
        # obj.CtrlB.write(b'switch_set 10 0\n');time.sleep(0.1)
        # obj.CtrlB.write(b'cpld_spi_wr 0x2c 2700\n');time.sleep(0.1)
        # obj.CtrlB.write(b'cpld_spi_wr 0x2f 2700\n');time.sleep(0.1)
        #get PD current X
        obj.CtrlB.flushInput()
        obj.CtrlB.flushOutput()
        #b'cpld_spi_ave 0x1e 10 1\n' is read Y path
        obj.CtrlB.write(b'cpld_spi_ave 0x20 10 1\n');time.sleep(0.1)
        data1=obj.CtrlB.read_all().decode('utf-8')#need test
        i=0
        while not 'ok' in data1:
            time.sleep(0.05)
            obj.CtrlB.write(b'cpld_spi_ave 0x20 10 1\n');time.sleep(0.1)
            data1=obj.CtrlB.read_all().decode('utf-8')#need test
            i+=1
            if i>10:
                print('get PD current X Error!No valid current returned after ten times')
                return
        out1=TIA_tra_data(data1)#Where it is?_need to check and rewrite
        current_x=(out1/65535)*4.096/10/150*1000 #mA
        VPDX=2150/4096*4.85-(out1/65535)*4.096/10
        #get PD current Y
        obj.CtrlB.flushInput()
        obj.CtrlB.flushOutput()
        obj.CtrlB.write(b'cpld_spi_ave 0x1e 10 1\n');time.sleep(0.1)
        data2=obj.CtrlB.read_all().decode('utf-8')#need test
        i=0
        while not 'ok' in data2:
            obj.CtrlB.write(b'cpld_spi_ave 0x1e 10 1\n');time.sleep(0.1)
            data2=obj.CtrlB.read_all().decode('utf-8')#need test
            i+=1
            if i>10:
                print('get PD current Y Error!No valid current returned after ten times')
                return
        out2=TIA_tra_data(data2)
        current_y=(out2/65535)*4.096/10/150*1000 #mA
        VPDY=2150/4096*4.85-(out2/65535)*4.096/10
        #print('---Rx signal port---\nPoint {}\nCurrent_X:{}\nCurrent_Y:{}'.format(str(),str(current_x),str(current_y))) #\nVPD_X:{}\nVPD_Y:{}'.format(str(current_x),str(current_y),str(VPDX),str(VPDY)))
        return [current_x,current_y,VPDX,VPDY]
        #left work here to draw the curve and save
    except Exception as e:
        print(e)

def TIA_getPDcurrent_LO(obj):
    '''
    get TIA current LO port
    :param obj: GUI Class obj
    :return: [current_x,current_y,VPDX,VPDY]
    '''
    try:
        # #DAC开启，PD上电
        # # obj.CtrlB.write(i.encode('utf-8'))
        # obj.CtrlB.write(b'switch_set 10 0\n');time.sleep(0.1)
        # obj.CtrlB.write(b'cpld_spi_wr 0x2c 2700\n');time.sleep(0.1)
        # obj.CtrlB.write(b'cpld_spi_wr 0x2f 2700\n');time.sleep(0.1)
        obj.CtrlB.flushInput()
        obj.CtrlB.flushOutput()
        obj.CtrlB.write(b'cpld_spi_ave 0x20 100 1\n');time.sleep(0.8)
        data1=obj.CtrlB.read_all().decode('utf-8')#need test
        i=0
        while not 'ok' in data1:
            obj.CtrlB.write(b'cpld_spi_ave 0x20 100 1\n');time.sleep(0.8)
            data1=obj.CtrlB.read_all().decode('utf-8')#need test
            i+=1
            if i>10:
                print('get PD current X Error!No valid current returned after ten times')
                return
        out1=TIA_tra_data(data1)#Where it is?_need to check and rewrite
        #add one more time in case error
        if out1==0:
            obj.CtrlB.write(b'cpld_spi_ave 0x20 100 1\n');time.sleep(0.8)
            data1=obj.CtrlB.read_all().decode('utf-8')#need test
            out1=TIA_tra_data(data1)
        current_x=(out1/65535)*4.096/10/150*1000 #mA
        VPDX=2150/4096*4.85-(out1/65535)*4.096/10
        obj.CtrlB.write(b'cpld_spi_ave 0x1e 100 1\n');time.sleep(0.8)
        data2=obj.CtrlB.read_all().decode('utf-8')#need test
        '''
        %data1/data2=0x00f0（DEC=240）对应10nA
        %data1/data2=0x0960（DEC=2400）对应100nA
        %运放最大输出3.25V，对应data1/data2=0xc9c9
        '''
        i=0
        while not 'ok' in data2:
            obj.CtrlB.write(b'cpld_spi_ave 0x1e 100 1\n');time.sleep(0.8)
            data2=obj.CtrlB.read_all().decode('utf-8')#need test
            i+=1
            if i>10:
                print('get PD current Y Error!No valid current returned after ten times')
                return
        out2=TIA_tra_data(data2)
        #add one more time in case error
        if out2==0:
            obj.CtrlB.write(b'cpld_spi_ave 0x1e 100 1\n');time.sleep(0.8)
            data2=obj.CtrlB.read_all().decode('utf-8')#need test
            out2=TIA_tra_data(data2)
        current_y=(out2/65535)*4.096/10/150*1000 #mA
        VPDY=2150/4096*4.85-(out2/65535)*4.096/10
        print('---LO port---\nCurrent_X:{}\nCurrent_Y:{}\nVPD_X:{}\nVPD_Y:{}'.format(str(current_x),
                                                                str(current_y),str(VPDX),str(VPDY)))
        return [current_x,current_y,VPDX,VPDY]
        #left work here to draw the curve and save
    except Exception as e:
        print(e)

def ICR_scamblePol(obj):
    '''
    Scamble mode to find balance through Oscilloscope reading
    :param obj: GUI Class obj
    :return: NA
    '''
    try:
        #旋转偏振控制器，使得X/Y两路输出基本一致（相差10%）
        print('旋转偏振控制器，使得X/Y两路输出基本一致')
        obj.ICRtf.write(':POL2:MODE SCRAMBLE')
        obj.ICRtf.write(':SCRAMBLE:FUNCTION SIN')
        obj.ICRtf.write(':POL2:SCRAMBLE:FREQUENCY1 0.005')
        obj.ICRtf.write(':POL2:SCRAMBLE:FREQUENCY2 0.009')
        obj.ICRtf.write(':POL2:SCRAMBLE:FREQUENCY3 0.01')

        #read data with Oscilloscope
        obj.OScope.timeout=3000
        obj.OScope.write('C2:VDIV 0.08')
        obj.OScope.write('C2:OFFSET 0')
        obj.OScope.write('C3:VDIV 0.08')
        obj.OScope.write('C3:OFFSET 0')
        obj.OScope.write('C6:VDIV 0.08')
        obj.OScope.write('C6:OFFSET 0')
        obj.OScope.write('C7:VDIV 0.08')
        obj.OScope.write('C7:OFFSET 0')
        obj.OScope.write('TRMD AUTO')
        #query the data of amtiplitude
        A1=obj.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')
        A3=obj.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')
        while not isfloat(A1) or not isfloat(A3):
            A1=obj.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')
            A3=obj.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')
        #Judge by amtiplitude
        while float(A1)/float(A3)>1.08 or float(A1)/float(A3)<0.92:
            time.sleep(0.2)
            A1=obj.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')
            A3=obj.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')
        obj.ICRtf.write(':POL2:MODE MANUAL')
    except Exception as e:
        print(e)

def ICR_manualPol_PD(obj):
    '''
    Manual mode to find balance through PD reading
    :param obj: GUI Class obj
    :return: NA
    '''
    try:
        obj.MPC_Find_Position('DIFF', False)
        # TIA get PD current
        current_x, current_y = TIA_getPDcurrent(obj)[0:2]
        #找PD电流的平衡点
        while current_x/current_y>1.09 or current_x/current_y<0.92:
            obj.MPC_Find_Position_Fine('DIFF', False)
            current_x,current_y=TIA_getPDcurrent(obj)[0:2]
        print('find the balance point, X:{},Y:{}'.format(current_x,current_y))
    except Exception as e:
        print(e)

def ICR_manualPol_ScopeBalance_Judge(obj):
    '''
    Manual mode to find Scope output balance through Oscilloscope reading judge
    :param obj: GUI Class obj
    :return: NA
    '''
    try:
        #旋转偏振控制器，使得X/Y两路输出基本一致（相差10%）

        #read data with Oscilloscope
        obj.OScope.timeout=3000
        obj.OScope.write('C2:VDIV 0.08')
        obj.OScope.write('C2:OFFSET 0')
        obj.OScope.write('C3:VDIV 0.08')
        obj.OScope.write('C3:OFFSET 0')
        obj.OScope.write('C6:VDIV 0.08')
        obj.OScope.write('C6:OFFSET 0')
        obj.OScope.write('C7:VDIV 0.08')
        obj.OScope.write('C7:OFFSET 0')
        obj.OScope.write('TRMD AUTO')
        #query the data of amtiplitude
        # A1=obj.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')
        # A3=obj.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')
        # while not isfloat(A1) or not isfloat(A3):
        #     A1=obj.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')
        #     A3=obj.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')
        obj.MPC_Find_Position_ScopeBalance()
        A1 = obj.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')
        A3 = obj.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')
        while not isfloat(A1) or not isfloat(A3):
            A1 = obj.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')
            A3 = obj.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')
        #Judge by amtiplitude
        count=0
        while float(A1)/float(A3)>1.08 or float(A1)/float(A3)<0.92 and count<3:
            obj.MPC_Find_Position_ScopeBalance_Fine()
            count+=1
            print('Scope balance执行第{}次微调'.format(count))
            A1 = obj.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')
            A3 = obj.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')
            while not isfloat(A1) or not isfloat(A3):
                A1 = obj.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')
                A3 = obj.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')
        #obj.ICRtf.write(':POL2:MODE MANUAL')
    except Exception as e:
        print(e)

def ICR_manualPol_PDbalance_ScopeJudge(obj):
    '''
    Manual adjust to find the X Y PD balance to find balance through Oscilloscope reading
    :param obj: GUI Class obj
    :return: NA
    '''
    try:
        #read data with Oscilloscope
        obj.OScope.timeout=3000
        obj.OScope.write('C2:VDIV 0.08')
        obj.OScope.write('C2:OFFSET 0')
        obj.OScope.write('C3:VDIV 0.08')
        obj.OScope.write('C3:OFFSET 0')
        obj.OScope.write('C6:VDIV 0.08')
        obj.OScope.write('C6:OFFSET 0')
        obj.OScope.write('C7:VDIV 0.08')
        obj.OScope.write('C7:OFFSET 0')
        obj.OScope.write('TRMD AUTO')
        print('旋转偏振控制器，使得X/Y两路PD输出基本一致')
        obj.MPC_Find_Position('DIFF', False)
        #curX, curY = TIA_getPDcurrent(mawin)[0:2]
        # #旋转偏振控制器，使得X/Y两路输出基本一致（相差10%）
        #
        # obj.ICRtf.write(':POL2:MODE SCRAMBLE')
        # obj.ICRtf.write(':SCRAMBLE:FUNCTION SIN')
        # obj.ICRtf.write(':POL2:SCRAMBLE:FREQUENCY1 0.005')
        # obj.ICRtf.write(':POL2:SCRAMBLE:FREQUENCY2 0.009')
        # obj.ICRtf.write(':POL2:SCRAMBLE:FREQUENCY3 0.01')

        #query the data of amtiplitude
        A1=obj.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')
        A3=obj.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')
        while not isfloat(A1) or not isfloat(A3):
            A1=obj.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')
            A3=obj.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')
        #Judge by amtiplitude
        while float(A1)/float(A3)>1.08 or float(A1)/float(A3)<0.92:
            obj.MPC_Find_Position_Fine('DIFF', False)
            A1=obj.OScope.query('VBS? return=app.Measure.P1.last.Result.Value')
            A3=obj.OScope.query('VBS? return=app.Measure.P3.last.Result.Value')
    except Exception as e:
        print(e)

def isfloat(value):
    '''
    check if is float
    :param value: input value
    :return: True/False
    '''
    try:
        float(value)
        return True
    except ValueError:
        return False

def ICR_scamblePol_XYmax(obj):
    '''
    Ununsed for now as this is mannual adjust
    scamble mode to find XY max status
    :param obj: GUI Class obj
    :return: NA
    '''
    try:
        #旋转偏振控制器，使得X/Y两路输出基本一致（相差10%）
        obj.ICRtf.write(':POL2:MODE MANUAL');time.sleep(0.1)
        obj.ICRtf.write(':POL2:MANUAL:SET1 0.000');time.sleep(0.1)
        obj.ICRtf.write(':POL2:MANUAL:SET2 0.000');time.sleep(0.1)
        obj.ICRtf.write(':POL2:MANUAL:SET3 0.000');time.sleep(0.1)
        v1=np.arange(0,1,0.02)
        v2=np.arange(0,1,0.02)
        v3=np.arange(0,1,0.02)
        length1=len(v1)
        length2=len(v2)
        length3=len(v3)

        k1=k2=k3=1
        '''
        cur11扫描一个玻片时，任意时刻记录的x路电流；cur12扫描一个玻片时，任意时刻记录的y路电流；
        cur1表示扫描第一个玻片，获得x和y路电流合集/向量；cur2表示扫描第二个玻片，获得x和y路电流合集/向量；cur3表示扫描第二个玻片，获得x和y路电流合集/向量；
        '''
        current_X=[]
        current_Y=[]

    except Exception as e:
        print(e)
