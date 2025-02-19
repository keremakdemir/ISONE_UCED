# -*- coding: utf-8 -*-
"""
Created on Wed Oct  3 21:29:55 2018

@author: jkern
"""

############################################################################
#                               DATA SETUP

# This file selects a single year from the synthetic record and organizes the 
# data in a form that is accessible to the unit commitment/economic dispatch
# (UC/ED) simulation model. This script can be interfaced with a data mining
# scheme for selecting specific years to run.
############################################################################


############################################################################
#                         YEAR SELECTION

import pandas as pd
import NEISO_exchange_time_series
import NEISO_data_setup

Load_data = pd.read_csv('../Time_series_data/Synthetic_demand_pathflows/Sim_hourly_load.csv',header=0)
year_length = len(Load_data)
years = int(year_length/365)
print('Possible hub heights are 53, 60, 80, 90, 100, 110, 120, 140, 160, 180 and 200 m.')
Hub_height = int(input("Enter hub height: "))
Offshore_capacity = int(input("Enter offshore wind power capacity (this is just used to name the folders): "))

for i in range(0,69):
    
    year = i
    
    print(i)
    
############################################################################
#                          UC/ED Data File Setup
    
    NEISO_exchange_time_series.exchange(year)
    NEISO_data_setup.setup(year, Hub_height, Offshore_capacity)

