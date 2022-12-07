# -*- coding: utf-8 -*-
import Instruments.CtrlBoard as cb

class a():
    CtrlB=object
    port='com3'

mawin=a()
cb.open_board_VOAcal(mawin,mawin.port)
# auto get drv out setting to make drv vpd as 2.2v~2.3v
DRV_writeOK = cb.drv_VCC_set_EEPROM(mawin)
if DRV_writeOK:
    print('***********DRV VCC out write done!************')
else:
    print('!!!!DRV VCC out write failed!!!!!!!!!!!!!!!!!!!!')
