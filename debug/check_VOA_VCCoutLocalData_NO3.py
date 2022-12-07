import shutil
import os
#import xlwings as xw
import pandas as pd
import numpy as np
import time
import configparser
import Common_functions.General_functions as gf

path=r'C:\Test_result\VOAcal\Normal'
report_path=r'C:\Users\Jiawen_zhou\OneDrive\桌面\DataUploadTest\DataSheet'
SN_list_path=r'D:\SN_all.txt'
report_model=r'C:\Users\Jiawen_zhou\OneDrive\桌面\DataUploadTest\FS400发货数据模板1021.xlsx'

sn_list=[]
with open(SN_list_path) as f:
    for i in f:
        sn_list.append(i.strip())
print('SN to query:\n',sn_list)
resu=[i for i in os.listdir(path) if len(i)>15 and not '.' in i]

valid_sn=[]
invalid_sn=[]
invalidConfig_sn=[]
invalid_sn_notReady=[]
invalid_sn_notDecode=[]
error_read_sn=[]
result_pageRead=[]
invalid_DRV_SN=[]

for i in sn_list:
    for j in resu:
        if i in j:
            data_time=j.split('_')[1]+'_'+j.split('_')[2]
            filefolder=os.path.join(path,i+'_'+data_time)
            file_toread=os.path.join(filefolder,'logfile.log')
            if not os.path.exists(file_toread):
                invalid_sn.append(i+'_noFileFound')
                continue
            with open(file_toread, 'r') as ff:
                a = ff.read()
            if not 'DRV VCC out write done!' in a:
                invalid_sn.append(i)
            elif not '0xB016: 0x0020' in a:
                invalid_sn_notReady.append(i)
            elif 'codec can\'t decode byte' in a:
                invalid_sn_notDecode.append(i)
            else:
                #now check drv vccout value is 4860 or not
                try:
                    sta7 = a.index('Write page 7')
                    sta8 = a.index('Write page 8')
                    sta9 = a.index('Write page 9')

                    sta11 = a.index('Write page 11')
                    sta13 = a.index('Write page 13')
                    sta15 = a.index('Write page 15')
                    stop15 = a.index('DRV VCC out write done!')

                    page7 = a[sta7:sta8]
                    page8 = a[sta8:sta9]
                    page9 = a[sta9:sta11]
                    page11 = a[sta11:sta13]
                    page13 = a[sta13:sta15]
                    page15 = a[sta15:stop15]

                    drv_flag7 = [i for i in page7.split('\\n') if '4860' in i]
                    drv_flag8 = [i for i in page8.split('\\n') if '4860' in i]
                    drv_flag9 = [i for i in page9.split('\\n') if '4860' in i]
                    drv_flag11 = [i for i in page11.split('\\n') if '4860' in i]
                    drv_flag13 = [i for i in page13.split('\\n') if '4860' in i]
                    drv_flag15 = [i for i in page15.split('\\n') if '4860' in i]

                    result_pageRead.append([i]+drv_flag7+drv_flag8+drv_flag9+drv_flag11+drv_flag13+drv_flag15)
                    if len(drv_flag7)+len(drv_flag8)+len(drv_flag9)+len(drv_flag11)+len(drv_flag13)+len(drv_flag15)!=6:
                        invalid_DRV_SN.append(i)
                    else:
                        valid_sn.append(i)

                except Exception as e:
                    print(e)
                    error_read_sn.append(i)
                    continue
                #valid_sn.append(i)


# for k in sn_list:
#     if not k in resu:
#         print('Invalid SN not in the target folder: \n{}\nSN:{}...'.format(path,k))
#         invalid_sn.append(k)
#         continue
#     #valid SN and valid folder, now to check whether data exists
#     print('SN to go to the next check: ', k)
#
#     dt_folder=os.path.join(path,k)
#     repo_fol = [i for i in os.listdir(dt_folder) if not '.' in i]
#     print(repo_fol)
#     if 'VOAcal' in repo_fol:
#         repo_fol.remove('VOAcal')
#     else:
#         invalid_sn.append(k)
#         print('No VOA calibration data found for SN : {}'.format(k))
#         continue
#
#     # *********-------Now to perform the VOA calibration data check-----**************
#     print('\n\n\nNow go to the VOA calibration data check...')
#     # *****VOA calibration status check******#
#     print(k, repo_fol, 'VOA calibration status check...')
#     VOAcal_folder = [i for i in os.listdir(os.path.join(path, k, 'VOAcal')) if len(i) == 15]
#     if len(VOAcal_folder) == 0:
#         print('Empty folder in VOAcal')
#     print(k, VOAcal_folder)
#     if len(VOAcal_folder) > 0:
#         VOAcal_folder.sort()
#         latest_folder = VOAcal_folder[-1]
#         latest_folder_path = os.path.join(path, k, 'VOAcal', latest_folder)
#         config_file = os.path.join(latest_folder_path, 'logfile.log')
#         latest_folder_report = os.path.join(latest_folder_path, k + '_VOA_Calibration_Report_' + latest_folder + '.xlsx')
#         latest_folder_file=os.path.join(latest_folder_path, k + '_VOA_Calibration_raw_' + latest_folder + '.csv')
#         if not os.path.exists(config_file):
#             print('No log file found for SN: '+k)
#             invalidConfig_sn.append(k)
#             continue
#         # Read the log file
#         with open(config_file,'r') as ff:
#             a=ff.read()
#         if not 'DRV VCC out write done!' in a:
#             invalid_sn.append(k)

print('Result of DRV check:\n',result_pageRead)
print('Valid SN list: ')
print(valid_sn)
print('Invalie SN list of Drv not written:')
print(invalid_sn)
print('Invalie SN list of board not ready:')
print(invalid_sn_notReady)
print('Invalie SN list of error decode:')
print(invalid_sn_notDecode)
print('Invalie SN list of log not exists:')
print(invalidConfig_sn)
print('VOA DRV vcc checked done!')

print('\n***************************DRV result************************\n')

print('Invalid DRV sn:\n',invalid_DRV_SN)
print('ERROR DRV sn:\n',error_read_sn)