import pandas as pd
from pandas import DataFrame
import numpy as np
# import General_functions as gf
# import CtrlBoard as cb

# class a():
#     CtrlB=object

def er_judge(test_result,er_judgeIQ=20,er_judgeXY=20,di_judge=0,il_judge=0.7):
    '''
    Updated on 05242022:add the function to check if :
    IQ_ER<15 or XY_ER<10
    and at the same time
    diff<-3dB
    Then perform the corresponding channel retest
    07172022 update: add IL diff retest condition
    :parameter:dataframe of DC test result,config,noipwr,er_judgeIQ=15,er_judgeXY=10,di_judge=-3 in default
    :return: updated dataframe of DC test result,the retested channels
    '''
    df_er=test_result.iloc[:,14:20].astype('float')
    #add max and min bias to judgement
    df_maxBias=test_result.iloc[:,26:32].astype('int')
    #df_minBias=test_result.iloc[:,32:38] #This code is to select the min values
    df_minBias=test_result.iloc[:,8:14].astype('int') #This code is to select the min\min\quad values
    df_IL=test_result.iloc[:,7].astype('float')
    #The judge condition
    diff_judge=di_judge
    ch_retest=set()
    er_retest= []
    ind_er=set()
    for i in range(6):
        if i == 0 or i==1 or i==3 or i==4:
            er_judge=er_judgeIQ
        else:
            er_judge=er_judgeXY
        type_er=df_er.iloc[:,i]
        type_max=df_maxBias.iloc[:,i]
        type_min=df_minBias.iloc[:,i]
        ind=type_er.loc[type_er<er_judge].index.to_list() #return the index of er not meet condition
        #not meet ER spec
        for i1 in ind:
            ind_er.add(i1)
        diff_er  = np.diff(type_er)
        diff_max = np.diff(type_max)
        diff_min = np.diff(type_min)
        if len(ind)==0 or len(diff_er)==0 or len(diff_max)==0 or len(diff_min)==0:
            continue
        else:
            # if np.min(ind)==0:
            #     er_retest.append([i,ind[0]])
            #     ch_retest.add(ind[0])
            dif=diff_er[[i2-1 for i2 in ind if i2>0]]
            di_max=diff_max[[i2-1 for i2 in ind if i2>0]]
            di_min=diff_min[[i2-1 for i2 in ind if i2>0]]
            for t,j in enumerate(dif):
                if j<diff_judge:#er decrease condition
                    if 25<abs(di_max[t])<400 or 25<abs(di_min[t])<400:
                        if not 0 in ind:
                            er_retest.append([i,ind[t]])
                            ch_retest.add(ind[t])
                        else:
                            er_retest.append([i,ind[t+1]])
                            ch_retest.add(ind[t+1])
    if 1 in ch_retest and 0 in ind_er:
        ch_retest.add(0)

    #add IL judge condition according to channel quantity
    cha_qty=test_result.shape[0]
    if cha_qty==3:
        il_judge=1
    elif cha_qty>=64:
        il_judge=0.7
    else:
        il_judge=5 #never met
    dif_IL=np.diff(df_IL)
    for i,j in enumerate(dif_IL):
        if abs(j)>il_judge:
            ch_retest.add(i+1)

    return (ch_retest,er_retest)

# f=r"C:\Users\Jiawen_zhou\OneDrive\桌面\WRRC270038_Retest_20220803_180649.csv"
# df=pd.read_csv(f)
# obj=a()
# cb.open_board(obj,'com22')
# gf.write_eeprom(obj,df)
# er_judge(df,15,10,-3)