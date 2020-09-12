# -*- coding: utf-8 -*-
"""
Created on Sat Jun 13 12:52:08 2020

@author: kakdemi
"""

import pandas as pd

sim_2015 = pd.read_csv('shadow_price_2015.csv', usecols= ['Constraint', 'Time', 'Value'])
sim_2016 = pd.read_csv('shadow_price_2016.csv', usecols= ['Constraint', 'Time', 'Value'])
sim_2017 = pd.read_csv('shadow_price_2017.csv', usecols= ['Constraint', 'Time', 'Value'])

bal_list = ['Bal1Constraint', 'Bal2Constraint', 'Bal3Constraint', 'Bal4Constraint', 'Bal5Constraint', 'Bal6Constraint', 'Bal7Constraint', 'Bal8Constraint']

for b in bal_list:
    
    name_1 = b + '_2015'
    globals()[name_1] = sim_2015.loc[sim_2015['Constraint'] == b]
    globals()[name_1].reset_index(inplace=True)
    del globals()[name_1]['Time']
    
    name_2 = b + '_2016'
    globals()[name_2] = sim_2016.loc[sim_2016['Constraint'] == b]
    globals()[name_2].reset_index(inplace=True)
    del globals()[name_2]['Time']
    
    name_3 = b + '_2017'
    globals()[name_3] = sim_2017.loc[sim_2017['Constraint'] == b]
    globals()[name_3].reset_index(inplace=True)
    del globals()[name_3]['Time']
    
    name_t = b + '_total'
    globals()[name_t] = pd.concat([globals()[name_1], globals()[name_2], globals()[name_3]])
    globals()[name_t].reset_index(inplace=True, drop=True)
    del globals()[name_t]['index']
    
Total_sim = pd.concat([Bal1Constraint_total['Value'], Bal2Constraint_total['Value'], Bal3Constraint_total['Value'], Bal4Constraint_total['Value'], Bal5Constraint_total['Value'], Bal6Constraint_total['Value'], Bal7Constraint_total['Value'], Bal8Constraint_total['Value']], axis=1)
Total_sim.columns = ['Bal1Constraint', 'Bal2Constraint', 'Bal3Constraint', 'Bal4Constraint', 'Bal5Constraint', 'Bal6Constraint', 'Bal7Constraint', 'Bal8Constraint']

hist_hourly_2015 = pd.read_excel('hist_hourly_2015.xls', sheet_name = 'ISONE CA', usecols = ['DA_LMP'])
hist_hourly_2016 = pd.read_excel('hist_hourly_2016.xls', sheet_name = 'ISO NE CA', usecols = ['DA_LMP'])
hist_hourly_2017 = pd.read_excel('hist_hourly_2017.xlsx', sheet_name = 'ISO NE CA', usecols = ['DA_LMP'])

hist_daily_2015 = pd.read_excel('hist_daily_2015.xls', sheet_name = 'ISONE CA', usecols = ['AvgDALMP'])
hist_daily_2016 = pd.read_excel('hist_daily_2016.xlsx', sheet_name = 'ISO NE CA', usecols = ['Avg_DA_LMP'])
hist_daily_2017 = pd.read_excel('hist_daily_2017.xlsx', sheet_name = 'ISO NE CA', usecols = ['Avg_DA_LMP'])

hist_daily_2016.columns = ['AvgDALMP']
hist_daily_2017.columns = ['AvgDALMP']

hist_hourly = pd.concat([hist_hourly_2015, hist_hourly_2016, hist_hourly_2017])
hist_hourly.columns = ['LMP']
hist_hourly.reset_index(inplace=True, drop=True)

hist_daily = pd.concat([hist_daily_2015, hist_daily_2016, hist_daily_2017])
hist_daily.columns = ['LMP']
hist_daily.reset_index(inplace=True, drop=True)

for year in range(3):
    hist_1 = hist_hourly.loc[year*8760:year*8760+8759, 'LMP'].copy()
    hist_1.reset_index(drop=True, inplace=True)
    globals()['Hourly_'+str(year)] = pd.DataFrame(hist_1).copy()

    
hist_hourly_final = pd.concat([Hourly_0, Hourly_1, Hourly_2])
hist_hourly_final.reset_index(inplace=True, drop=True)

for year in range(3):
    hist_2 = hist_daily.loc[year*365:year*365+364, 'LMP'].copy()
    hist_2.reset_index(drop=True, inplace=True)
    globals()['Daily_'+str(year)] = pd.DataFrame(hist_2).copy()
    
hist_daily_final = pd.concat([Daily_0, Daily_1, Daily_2])
hist_daily_final.reset_index(inplace=True, drop=True)

time_range = pd.date_range(start='2009-01-01 00:00:00', end='2011-12-31 23:00:00', freq='H')
Hourly_sim = Total_sim.copy()
Hourly_sim['Date'] = time_range
Hourly_sim.set_index('Date', inplace=True)

Daily_sim = Hourly_sim.resample('D').mean()
Daily_sim.reset_index(inplace=True, drop=True)

writer = pd.ExcelWriter('validation_prices.xlsx', engine='xlsxwriter')
Daily_sim.to_excel(writer, sheet_name='Daily_sim', index=False)
Total_sim.to_excel(writer, sheet_name='Hourly_sim', index=False)
hist_hourly_final.to_excel(writer, sheet_name='hist_hourly', index=False)
hist_daily_final.to_excel(writer, sheet_name='hist_daily', index=False)
writer.save()

