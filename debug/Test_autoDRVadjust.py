import pandas as pd
from pandas import DataFrame
import numpy as np
import General_functions as gf
import os
# import CtrlBoard as cb

class a():
    CtrlB=object
    port='com10'

import CtrlBoard as cb
mawin=a()
cb.open_board_VOAcal(mawin,mawin.port)
cb.drv_VCC_set_EEPROM(mawin)
cb.close_board(mawin)

