# -*- coding: utf-8 -*-
"""
Created on Fri Dec 21 15:55:48 2018

@author: jkern
"""

import pandas as pd
import numpy as np

df_monthly = pd.read_excel('hydromonthlyavgs.xlsx',sheetname='hourlydispatch',header=0)

daily = np.zeros((365,7))

zones = ['CT','ME','NH','NEMA','RI', 'VT','WCMA']

for z in zones:
    
    z_index = zones.index(z)
    
    daily[0:31,z_index] = df_monthly.loc[0,z]*24
    daily[31:60,z_index] = df_monthly.loc[1,z]*24
    daily[60:91,z_index] = df_monthly.loc[2,z]*24
    daily[91:121,z_index] = df_monthly.loc[3,z]*24
    daily[121:152,z_index] = df_monthly.loc[4,z]*24
    daily[152:182,z_index] = df_monthly.loc[5,z]*24
    daily[182:213,z_index] = df_monthly.loc[6,z]*24
    daily[213:243,z_index] = df_monthly.loc[7,z]*24
    daily[243:274,z_index] = df_monthly.loc[8,z]*24
    daily[274:305,z_index] = df_monthly.loc[9,z]*24
    daily[305:335,z_index] = df_monthly.loc[10,z]*24
    daily[335:365,z_index] = df_monthly.loc[11,z]*24
        

df_daily = pd.DataFrame(daily)
df_daily.columns = zones
df_daily.to_excel('daily_hydro.xlsx')