# -*- coding: utf-8 -*-
import pandas as pd
from pandas import DataFrame
import numpy as np
import Common_functions.General_functions as gf
import os
# import CtrlBoard as cb

class a():
    CtrlB=object
    port='com10'

sta1 = 191.3125
end1 = 196.0375
wl_c = np.linspace(sta1, end1, 64)
#VOA calibration range:
sta1 = 190.1125
end1 = 197.2375
voa_wl=np.linspace(sta1,end1,round((end1-sta1)/0.075)+1).round(4)
#self.wl_c = np.linspace(sta1, end1, 64)

channel = [i for i in range(9 , 73,4)] + [72]
fre = np.interp(channel, range(9 , 73), wl_c)

path=r'C:\Test_result\VOAcal\Normal\WRRC280045_20220914_194709'
f=os.path.join(path,'WRRC270045_VOA_Calibration_raw_20220914_194709.csv')
df=pd.read_csv(f)
print(df)
#obj=a()
#cb.open_board(obj,'com22')
voa_XI, voa_XQ, voa_YI, voa_YQ=gf.VOA_cal_data(df,fre, wl_c,voa_wl)
print(voa_XI, voa_XQ, voa_YI, voa_YQ)
test_result_fit = DataFrame(columns=('SN', 'RX_BW_DESK', 'TEMP', 'DATE', 'TIME', '400G_Fre',
                                         'VOA_XI', 'VOA_XQ', 'VOA_YI', 'VOA_YQ'))
config = ['Test', 'Test', 'Test' ,'Test','Test']
for i in range(len(voa_wl)):
    da_add = config + [voa_wl[i]] + [voa_XI[i]] + [voa_XQ[i]] + [voa_YI[i]] + [voa_YQ[i]]
    test_result_fit.loc[i] = da_add
report_file=os.path.join(path,'WRRC270045_fit.csv')
test_result_fit.to_csv(report_file, index=False)

#path=r'C:\backup\Desktop\FS400\Test_Script\VOA_calibration\EEPROM_write\Test_data\WRRC270027_20220906_113534'
# f1=os.path.join(path,'WRRC270027.csv')
# df1=pd.read_csv(f1)
print(test_result_fit)
import CtrlBoard as cb
mawin=a()
cb.open_board_VOAcal(mawin,mawin.port)
gf.write_VOAcal2EEP(mawin,test_result_fit)
