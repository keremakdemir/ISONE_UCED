# -*- coding: utf-8 -*-
"""
Created on Fri Dec 21 15:55:48 2018

@author: jkern
"""

import pandas as pd
import numpy as np

df_monthly = pd.read_excel('hydromonthlyavgs.xlsx',sheetname='hourlydispatch',header=0)

daily = 