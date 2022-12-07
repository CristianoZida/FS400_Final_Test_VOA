# -*- coding: utf-8 -*-
import Instruments.CtrlBoard as cb
file_store=r'C:\Users\Jiawen_zhou\testgit\FS400_Final_Test\debug\DrvVCC_Monitor_VOAcal.csv'

class a():
    CtrlB=object
    port='com3'

mawin=a()
cb.open_board_VOAcal(mawin,mawin.port)
# auto get drv out setting to make drv vpd as 2.2v~2.3v, then monitor periodly
DRV_readOK = cb.drv_VCC_Monitor_EEPROM(mawin,file_store,60,60)
if DRV_readOK:
    print('DRV VCC out monitor done!')
else:
    print('DRV VCC out monitor failed!')
