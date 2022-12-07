# -*- coding: utf-8 -*-
#this module is to write the control functions of power meter
# date:3/17/2022
# updated:9/17/2022
# author:Jiawen Zhou
import serial
import time

def open_JINLI(obj,JinLi_port):
    '''
    Open MPC connection
    :param obj: GUI Class obj
    :param mpc_port: MPC port number
    :return: True/False
    '''
    try:
        obj.JinLi = serial.Serial(JinLi_port,115200,timeout=3)
        time.sleep(5)
        # Check board status:
        obj.JinLi.write(b'set_uart_out_print_flag 0x22\n\r')
        time.sleep(1)
        obj.JinLi.write(b'msa_read 0xb016\r\n')
        time.sleep(0.2)
        r = obj.JinLi.read_all().decode('utf-8')
        print(r)
        count = 0
        count1 = 0
        while not '0x0020' in r:
            time.sleep(5)
            count += 1
            print('The {} times read not ready, wait for 5s to read...'.format(count))
            obj.JinLi.write(b'set_uart_out_print_flag 0x22\n\r')
            time.sleep(1)
            obj.JinLi.write(b'msa_read 0xb016\r\n')
            time.sleep(0.2)
            r = obj.JinLi.read_all().decode('utf-8')
            print(r)
            if count > 40:
                print('After 200s read JinLi not ready, please check...')
                return False
            if len(r) == 0:
                count1 += 1
                if count1 > 3:
                    print('JinLi未上电！')
                    return False
        print('JinLI ready...')
        return True
    except Exception as e:
        print(e)
        return False

def close_JINLI(obj):
    '''
    Close MPC connection
    :param obj: GUI Class obj:
    :return: True/False
    '''
    try:
        if serial.Serial.isOpen(obj.JinLi_port):
            obj.JinLi.close()
            print('JINLI串口关闭成功!')
    except Exception as e:
        print(e)
        return

def set_JINLIwave(obj,ch):
    '''
    set the wavelength of Jinli
    :param obj:mawin
    :param ch:9-72
    :return: NA
    '''
    cha=['A00B',
 'A017',
 'A023',
 'A02F',
 'A03B',
 'A047',
 'A053',
 'A05F',
 'A06B',
 'A077',
 'A083',
 'A08F',
 'A09B',
 'A0A7',
 'A0B3',
 'A0BF',
 'A0CB',
 'A0D7',
 'A0E3',
 'A0EF',
 'A0FB',
 'A107',
 'A113',
 'A11F',
 'A12B',
 'A137',
 'A143',
 'A14F',
 'A15B',
 'A167',
 'A173',
 'A17F',
 'A18B',
 'A197',
 'A1A3',
 'A1AF',
 'A1BB',
 'A1C7',
 'A1D3',
 'A1DF',
 'A1EB',
 'A1F7',
 'A203',
 'A20F',
 'A21B',
 'A227',
 'A233',
 'A23F',
 'A24B',
 'A257',
 'A263',
 'A26F',
 'A27B',
 'A287',
 'A293',
 'A29F',
 'A2AB',
 'A2B7',
 'A2C3',
 'A2CF',
 'A2DB',
 'A2E7',
 'A2F3',
 'A2FF'] #from ch9 to ch72
    if not 8<int(ch)<73:
        print('Error channel input, should be 9-72!!!')
        return
    obj.JinLi.read_all()
    time.sleep(0.1)
    cmd='msa_write 0xb400 0x{}\r\n'.format(cha[int(ch)-9]).encode('utf-8')
    obj.JinLi.write(cmd)
    #time.sleep(0.2)
    obj.JinLi.timeout=30
    print(obj.JinLi.read_until(b'ITTRA_Driver target'))
    obj.JinLi.read_all()
    obj.JinLi.timeout = 3

#Should not be here need to add to a new module
def set_CATFishwave(obj,ch):
    '''
    set the wavelength of CatFish
    :param obj:mawin
    :param ch:9-72
    :return: NA
    '''
    cha=['8002',
 '8008',
 '800E',
 '8014',
 '801A',
 '8020',
 '8026',
 '802C',
 '8032',
 '8038',
 '803E',
 '8044',
 '804A',
 '8050',
 '8056',
 '805C',
 '8062',
 '8068',
 '806E',
 '8074',
 '807A',
 '8080',
 '8086',
 '808C',
 '8092',
 '8098',
 '809E',
 '80A4',
 '80AA',
 '80B0',
 '80B6',
 '80BC',
 '80C2',
 '80C8',
 '80CE',
 '80D4',
 '80DA',
 '80E0',
 '80E6',
 '80EC',
 '80F2',
 '80F8',
 '80FE',
 '8104',
 '810A',
 '8110',
 '8116',
 '811C',
 '8122',
 '8128',
 '812E',
 '8134',
 '813A',
 '8140',
 '8146',
 '814C',
 '8152',
 '8158',
 '815E',
 '8164',
 '816A',
 '8170',
 '8176',
 '817C'] #from ch9 to ch72
    if not 8<int(ch)<73:
        print('Error channel input, should be 9-72!!!')
        return
    cmd='msa_write 0xb400 0x{}\r\n'.format(cha[int(ch)-9]).encode('utf-8')
    obj.JinLi.write(cmd);time.sleep(0.2)
    print(obj.JinLi.read_all())