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

# Default is that a random year from the synthetic record is selected to be run
# through the UC/ED model. 

import pandas as pd
import numpy as np

Load_data = pd.read_csv('../Time_series_data/Synthetic_demand_pathflows/Sim_hourly_load.csv',header=0)
year_length = len(Load_data)
years = int(year_length/365)
Hub_height = input("Enter hub height: ") 
Hub_height = int(Hub_height)

for i in range(0,10):
    
    year = i
    
    print(i)
    
############################################################################
#                          UC/ED Data File Setup
    
    import NEISO_exchange_time_series
    NEISO_exchange_time_series.exchange(year)

    import NEISO_data_setup
    NEISO_data_setup.setup(year, Hub_height)

