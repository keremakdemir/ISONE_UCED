# -*- coding: utf-8 -*-
"""
Created on Tue Feb  4 15:38:18 2020

@author: kakdemi
"""

import pandas as pd 
import matplotlib.pyplot as plt
from pandas.plotting import register_matplotlib_converters
register_matplotlib_converters()

#importing all hourly data for different heights from CSV file
All_Hourly_Data = pd.read_excel("Historical_Hourly_Wind_Speeds.xlsx", index_col="Date&Time", parse_dates=True)

#applying moving averages to smooth the data to see overall wind speed trend
smoothed_hourly_110 = All_Hourly_Data[110].rolling(window=8760).mean()

#plotting the overall wind speed trend for 110 m
plt.plot(All_Hourly_Data.index, smoothed_hourly_110)
plt.xticks(rotation=45)
plt.title("Hourly Wind Speeds at 110 m")
plt.ylabel("Wind Speed (m/s)")
plt.show()
#clearing the plot
plt.clf()

#defining different turbine heights
turbine_height = [53, 60, 80, 90, 100, 110, 120, 140, 160, 180, 200]

#creating an empty list to add all cut out frequencies
all_cut_out_freq = []

#creating empty dictionaries to keep hub height specific lists
hourly_energy_production_MWh = {}
cut_out_dict = {}
hourly_wind_dict = {}


for height in turbine_height:
    
    #creating the empty lists for every hub height in the dictionary
    hourly_energy_production_MWh['hourly_energy_%s' % height] = []
    cut_out_dict['cut_out_freq_%s' % height] = []
    hourly_wind_dict['hourly_wind_%s' % height] = []
    
    #adding raw wind speeds to the dictionary
    height_wind_speeds = list(All_Hourly_Data[height])
    hourly_wind_dict['hourly_wind_'+str(height)].extend(height_wind_speeds)
    
    for speed in height_wind_speeds:
        
        #finding wind energy production from regression on power curve of the MHI Vestas Offshore V164/9500
        if 0 <= speed < 3.5:
            energy_kwh = 0
            energy_mwh = 0
        elif 3.5 <= speed < 5.5:
            energy_kwh = (386.8*speed) - 1279.2
            energy_mwh = energy_kwh/1000
        elif 5.5 <= speed < 7.5:
            energy_kwh = (828.8*speed) - 3722
            energy_mwh = energy_kwh/1000
        elif 7.5 <= speed < 12.5:
            energy_kwh = (1347.8*speed) - 7611.7
            energy_mwh = energy_kwh/1000
        elif 12.5 <= speed < 14:
            energy_kwh = (279.6*speed) + 5610.8
            energy_mwh = energy_kwh/1000
        elif 14 <= speed <= 25:
            energy_kwh = 9500
            energy_mwh = energy_kwh/1000
        elif 25 <= speed:
            energy_kwh = 0
            energy_mwh = 0
        
        #adding energy produced to the hub specific lists in dictionary
        hourly_energy_production_MWh['hourly_energy_'+str(height)].append(energy_mwh)

#creating a dataframe for the hourly energy production dataset
list_labels = ["Date&Time", 53, 60, 80, 90, 100, 110, 120, 140, 160, 180, 200]
list_columns = [list(All_Hourly_Data.index), hourly_energy_production_MWh['hourly_energy_53'], hourly_energy_production_MWh['hourly_energy_60'], 
                hourly_energy_production_MWh['hourly_energy_80'], hourly_energy_production_MWh['hourly_energy_90'], 
                hourly_energy_production_MWh['hourly_energy_100'], hourly_energy_production_MWh['hourly_energy_110'], 
                hourly_energy_production_MWh['hourly_energy_120'], hourly_energy_production_MWh['hourly_energy_140'], 
                hourly_energy_production_MWh['hourly_energy_160'], hourly_energy_production_MWh['hourly_energy_180'], 
                hourly_energy_production_MWh['hourly_energy_200']]
zipped_list = list(zip(list_labels, list_columns))
hourly_energy_dict = dict(zipped_list)
Predicted_Hourly_Energy_MWh = pd.DataFrame(hourly_energy_dict)
#declaring the index of the dataframe as date column
Predicted_Hourly_Energy_MWh = Predicted_Hourly_Energy_MWh.set_index("Date&Time") 

#dropping all February 29ths
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2016-02-29 00:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2016-02-29 01:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2016-02-29 02:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2016-02-29 03:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2016-02-29 04:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2016-02-29 05:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2016-02-29 06:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2016-02-29 07:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2016-02-29 08:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2016-02-29 09:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2016-02-29 10:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2016-02-29 11:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2016-02-29 12:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2016-02-29 13:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2016-02-29 14:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2016-02-29 15:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2016-02-29 16:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2016-02-29 17:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2016-02-29 18:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2016-02-29 19:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2016-02-29 20:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2016-02-29 21:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2016-02-29 22:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2016-02-29 23:00:00'), inplace=True)

Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2012-02-29 00:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2012-02-29 01:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2012-02-29 02:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2012-02-29 03:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2012-02-29 04:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2012-02-29 05:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2012-02-29 06:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2012-02-29 07:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2012-02-29 08:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2012-02-29 09:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2012-02-29 10:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2012-02-29 11:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2012-02-29 12:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2012-02-29 13:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2012-02-29 14:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2012-02-29 15:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2012-02-29 16:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2012-02-29 17:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2012-02-29 18:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2012-02-29 19:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2012-02-29 20:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2012-02-29 21:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2012-02-29 22:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2012-02-29 23:00:00'), inplace=True)

Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2008-02-29 00:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2008-02-29 01:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2008-02-29 02:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2008-02-29 03:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2008-02-29 04:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2008-02-29 05:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2008-02-29 06:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2008-02-29 07:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2008-02-29 08:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2008-02-29 09:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2008-02-29 10:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2008-02-29 11:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2008-02-29 12:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2008-02-29 13:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2008-02-29 14:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2008-02-29 15:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2008-02-29 16:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2008-02-29 17:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2008-02-29 18:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2008-02-29 19:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2008-02-29 20:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2008-02-29 21:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2008-02-29 22:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2008-02-29 23:00:00'), inplace=True)

Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2004-02-29 00:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2004-02-29 01:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2004-02-29 02:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2004-02-29 03:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2004-02-29 04:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2004-02-29 05:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2004-02-29 06:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2004-02-29 07:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2004-02-29 08:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2004-02-29 09:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2004-02-29 10:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2004-02-29 11:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2004-02-29 12:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2004-02-29 13:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2004-02-29 14:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2004-02-29 15:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2004-02-29 16:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2004-02-29 17:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2004-02-29 18:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2004-02-29 19:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2004-02-29 20:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2004-02-29 21:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2004-02-29 22:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2004-02-29 23:00:00'), inplace=True)

Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2000-02-29 00:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2000-02-29 01:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2000-02-29 02:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2000-02-29 03:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2000-02-29 04:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2000-02-29 05:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2000-02-29 06:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2000-02-29 07:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2000-02-29 08:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2000-02-29 09:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2000-02-29 10:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2000-02-29 11:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2000-02-29 12:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2000-02-29 13:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2000-02-29 14:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2000-02-29 15:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2000-02-29 16:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2000-02-29 17:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2000-02-29 18:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2000-02-29 19:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2000-02-29 20:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2000-02-29 21:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2000-02-29 22:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('2000-02-29 23:00:00'), inplace=True)

Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1996-02-29 00:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1996-02-29 01:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1996-02-29 02:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1996-02-29 03:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1996-02-29 04:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1996-02-29 05:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1996-02-29 06:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1996-02-29 07:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1996-02-29 08:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1996-02-29 09:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1996-02-29 10:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1996-02-29 11:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1996-02-29 12:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1996-02-29 13:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1996-02-29 14:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1996-02-29 15:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1996-02-29 16:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1996-02-29 17:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1996-02-29 18:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1996-02-29 19:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1996-02-29 20:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1996-02-29 21:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1996-02-29 22:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1996-02-29 23:00:00'), inplace=True)

Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1992-02-29 00:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1992-02-29 01:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1992-02-29 02:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1992-02-29 03:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1992-02-29 04:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1992-02-29 05:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1992-02-29 06:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1992-02-29 07:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1992-02-29 08:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1992-02-29 09:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1992-02-29 10:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1992-02-29 11:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1992-02-29 12:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1992-02-29 13:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1992-02-29 14:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1992-02-29 15:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1992-02-29 16:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1992-02-29 17:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1992-02-29 18:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1992-02-29 19:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1992-02-29 20:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1992-02-29 21:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1992-02-29 22:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1992-02-29 23:00:00'), inplace=True)

Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1988-02-29 00:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1988-02-29 01:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1988-02-29 02:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1988-02-29 03:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1988-02-29 04:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1988-02-29 05:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1988-02-29 06:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1988-02-29 07:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1988-02-29 08:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1988-02-29 09:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1988-02-29 10:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1988-02-29 11:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1988-02-29 12:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1988-02-29 13:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1988-02-29 14:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1988-02-29 15:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1988-02-29 16:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1988-02-29 17:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1988-02-29 18:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1988-02-29 19:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1988-02-29 20:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1988-02-29 21:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1988-02-29 22:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1988-02-29 23:00:00'), inplace=True)

Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1984-02-29 00:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1984-02-29 01:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1984-02-29 02:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1984-02-29 03:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1984-02-29 04:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1984-02-29 05:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1984-02-29 06:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1984-02-29 07:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1984-02-29 08:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1984-02-29 09:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1984-02-29 10:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1984-02-29 11:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1984-02-29 12:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1984-02-29 13:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1984-02-29 14:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1984-02-29 15:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1984-02-29 16:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1984-02-29 17:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1984-02-29 18:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1984-02-29 19:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1984-02-29 20:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1984-02-29 21:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1984-02-29 22:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1984-02-29 23:00:00'), inplace=True)

Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1980-02-29 00:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1980-02-29 01:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1980-02-29 02:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1980-02-29 03:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1980-02-29 04:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1980-02-29 05:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1980-02-29 06:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1980-02-29 07:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1980-02-29 08:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1980-02-29 09:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1980-02-29 10:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1980-02-29 11:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1980-02-29 12:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1980-02-29 13:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1980-02-29 14:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1980-02-29 15:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1980-02-29 16:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1980-02-29 17:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1980-02-29 18:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1980-02-29 19:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1980-02-29 20:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1980-02-29 21:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1980-02-29 22:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1980-02-29 23:00:00'), inplace=True)

Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1976-02-29 00:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1976-02-29 01:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1976-02-29 02:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1976-02-29 03:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1976-02-29 04:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1976-02-29 05:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1976-02-29 06:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1976-02-29 07:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1976-02-29 08:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1976-02-29 09:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1976-02-29 10:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1976-02-29 11:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1976-02-29 12:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1976-02-29 13:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1976-02-29 14:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1976-02-29 15:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1976-02-29 16:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1976-02-29 17:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1976-02-29 18:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1976-02-29 19:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1976-02-29 20:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1976-02-29 21:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1976-02-29 22:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1976-02-29 23:00:00'), inplace=True)

Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1972-02-29 00:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1972-02-29 01:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1972-02-29 02:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1972-02-29 03:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1972-02-29 04:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1972-02-29 05:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1972-02-29 06:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1972-02-29 07:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1972-02-29 08:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1972-02-29 09:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1972-02-29 10:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1972-02-29 11:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1972-02-29 12:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1972-02-29 13:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1972-02-29 14:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1972-02-29 15:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1972-02-29 16:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1972-02-29 17:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1972-02-29 18:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1972-02-29 19:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1972-02-29 20:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1972-02-29 21:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1972-02-29 22:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1972-02-29 23:00:00'), inplace=True)

Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1968-02-29 00:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1968-02-29 01:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1968-02-29 02:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1968-02-29 03:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1968-02-29 04:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1968-02-29 05:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1968-02-29 06:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1968-02-29 07:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1968-02-29 08:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1968-02-29 09:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1968-02-29 10:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1968-02-29 11:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1968-02-29 12:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1968-02-29 13:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1968-02-29 14:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1968-02-29 15:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1968-02-29 16:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1968-02-29 17:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1968-02-29 18:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1968-02-29 19:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1968-02-29 20:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1968-02-29 21:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1968-02-29 22:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1968-02-29 23:00:00'), inplace=True)

Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1964-02-29 00:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1964-02-29 01:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1964-02-29 02:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1964-02-29 03:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1964-02-29 04:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1964-02-29 05:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1964-02-29 06:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1964-02-29 07:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1964-02-29 08:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1964-02-29 09:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1964-02-29 10:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1964-02-29 11:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1964-02-29 12:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1964-02-29 13:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1964-02-29 14:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1964-02-29 15:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1964-02-29 16:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1964-02-29 17:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1964-02-29 18:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1964-02-29 19:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1964-02-29 20:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1964-02-29 21:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1964-02-29 22:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1964-02-29 23:00:00'), inplace=True)

Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1960-02-29 00:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1960-02-29 01:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1960-02-29 02:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1960-02-29 03:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1960-02-29 04:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1960-02-29 05:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1960-02-29 06:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1960-02-29 07:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1960-02-29 08:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1960-02-29 09:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1960-02-29 10:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1960-02-29 11:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1960-02-29 12:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1960-02-29 13:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1960-02-29 14:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1960-02-29 15:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1960-02-29 16:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1960-02-29 17:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1960-02-29 18:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1960-02-29 19:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1960-02-29 20:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1960-02-29 21:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1960-02-29 22:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1960-02-29 23:00:00'), inplace=True)

Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1956-02-29 00:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1956-02-29 01:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1956-02-29 02:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1956-02-29 03:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1956-02-29 04:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1956-02-29 05:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1956-02-29 06:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1956-02-29 07:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1956-02-29 08:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1956-02-29 09:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1956-02-29 10:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1956-02-29 11:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1956-02-29 12:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1956-02-29 13:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1956-02-29 14:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1956-02-29 15:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1956-02-29 16:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1956-02-29 17:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1956-02-29 18:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1956-02-29 19:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1956-02-29 20:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1956-02-29 21:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1956-02-29 22:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1956-02-29 23:00:00'), inplace=True)

Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1952-02-29 00:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1952-02-29 01:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1952-02-29 02:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1952-02-29 03:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1952-02-29 04:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1952-02-29 05:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1952-02-29 06:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1952-02-29 07:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1952-02-29 08:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1952-02-29 09:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1952-02-29 10:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1952-02-29 11:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1952-02-29 12:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1952-02-29 13:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1952-02-29 14:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1952-02-29 15:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1952-02-29 16:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1952-02-29 17:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1952-02-29 18:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1952-02-29 19:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1952-02-29 20:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1952-02-29 21:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1952-02-29 22:00:00'), inplace=True)
Predicted_Hourly_Energy_MWh.drop(pd.to_datetime('1952-02-29 23:00:00'), inplace=True)

for height in turbine_height:
    
    #showing how many times wind speed exceeds cut-out speed
    cut_out_dict['cut_out_freq_'+str(height)].append(sum(1 for individual_speed in hourly_wind_dict['hourly_wind_'+str(height)] if individual_speed > 25))
    print("For hub height of "+str(height)+" m, hourly wind speed exceeds cut-out speed " + str(cut_out_dict['cut_out_freq_'+str(height)]).replace("[", "").replace("]", "") + " times.")
    #adding cut out frequency information to the general list
    all_cut_out_freq.extend(cut_out_dict['cut_out_freq_'+str(height)])
    
#plotting cut-out frequency vs hub height       
plt.plot(turbine_height, all_cut_out_freq) 
plt.xticks([53, 60, 80, 90, 100, 110, 120, 140, 160, 180, 200])  
plt.xlabel("Hub Height (m)")
plt.ylabel("Number of Hours that exceeds cut-out\n wind speed between 1949-2018")
plt.show()
#clearing the plot
plt.clf() 

#calculating the total energy produced for different hub heights
Total_Energy_53 = Predicted_Hourly_Energy_MWh[53].sum()
Total_Energy_60 = Predicted_Hourly_Energy_MWh[60].sum()
Total_Energy_80 = Predicted_Hourly_Energy_MWh[80].sum()
Total_Energy_90 = Predicted_Hourly_Energy_MWh[90].sum()
Total_Energy_100 = Predicted_Hourly_Energy_MWh[100].sum()
Total_Energy_110 = Predicted_Hourly_Energy_MWh[110].sum()
Total_Energy_120 = Predicted_Hourly_Energy_MWh[120].sum()
Total_Energy_140 = Predicted_Hourly_Energy_MWh[140].sum()
Total_Energy_160 = Predicted_Hourly_Energy_MWh[160].sum()
Total_Energy_180 = Predicted_Hourly_Energy_MWh[180].sum()
Total_Energy_200 = Predicted_Hourly_Energy_MWh[200].sum()

#adding total energy produced to a general list
Energy_List = [Total_Energy_53, Total_Energy_60, Total_Energy_80, Total_Energy_90, Total_Energy_100, Total_Energy_110, Total_Energy_120, Total_Energy_140, Total_Energy_160, Total_Energy_180, Total_Energy_200]
        
#plotting total energy produced vs hub height       
plt.plot(turbine_height, Energy_List) 
plt.xticks([53, 60, 80, 90, 100, 110, 120, 140, 160, 180, 200])  
plt.xlabel("Hub Height (m)")
plt.ylabel("Total Energy Produced between 1949-2018 (MWh)")
plt.show()
#clearing the plot
plt.clf() 

#resampling hourly energy generation data into yearly
Yearly_Total_Generation = Predicted_Hourly_Energy_MWh.resample('Y').sum()

#finding average annual energy produced between 1949-2018 for different hub heights
Annual_Energy_53 = Yearly_Total_Generation[53].mean()
Annual_Energy_60 = Yearly_Total_Generation[60].mean()
Annual_Energy_80 = Yearly_Total_Generation[80].mean()
Annual_Energy_90 = Yearly_Total_Generation[90].mean()
Annual_Energy_100 = Yearly_Total_Generation[100].mean()
Annual_Energy_110 = Yearly_Total_Generation[110].mean()
Annual_Energy_120 = Yearly_Total_Generation[120].mean()
Annual_Energy_140 = Yearly_Total_Generation[140].mean()
Annual_Energy_160 = Yearly_Total_Generation[160].mean()
Annual_Energy_180 = Yearly_Total_Generation[180].mean()
Annual_Energy_200 = Yearly_Total_Generation[200].mean()

#adding average annual energy produced to a general list
Annual_Energy_List = [Annual_Energy_53, Annual_Energy_60, Annual_Energy_80, Annual_Energy_90, Annual_Energy_100, Annual_Energy_110, Annual_Energy_120, Annual_Energy_140, Annual_Energy_160, Annual_Energy_180, Annual_Energy_200]

#plotting average annual energy produced between 1949-2018 vs hub height       
plt.plot(turbine_height, Annual_Energy_List) 
plt.xticks([53, 60, 80, 90, 100, 110, 120, 140, 160, 180, 200])  
plt.xlabel("Hub Height (m)")
plt.ylabel("Average Annual Energy Produced between\n 1949-2018 (MWh)")
plt.show()
#clearing the plot
plt.clf() 

#Excluding year 1948 and 1959 from the analysis
Partial_Generation_1 = Predicted_Hourly_Energy_MWh.loc['1949':'1958']
Partial_Generation_2 = Predicted_Hourly_Energy_MWh.loc['1960':'2018']
#adding all dataframes together
All_Generation_MWh = pd.concat([Partial_Generation_1, Partial_Generation_2])
All_Generation_MWh = All_Generation_MWh.sort_index()
#resetting the index and dropping the index column
All_Generation_MWh = All_Generation_MWh.reset_index(drop=True)
#including energy from all 84 turbines
All_Generation_MWh.loc[:, [53, 60, 80, 90, 100, 110, 120, 140, 160, 180, 200]] *= 84

#exporting data to a excel file
All_Generation_MWh.to_excel("Historical_Wind_Generation.xlsx")
