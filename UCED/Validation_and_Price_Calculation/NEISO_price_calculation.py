# -*- coding: utf-8 -*-
"""
Created on Sat Jun 13 21:23:17 2020

@author: kakdemi
"""

from sklearn import linear_model
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

#########################################################
#                    Weight by zone
#########################################################

#importing LMP data
Daily_sim = pd.read_excel('validation_prices.xlsx', sheet_name='Daily_sim')
Hourly_sim = pd.read_excel('validation_prices.xlsx', sheet_name='Hourly_sim')
Hourly_hist = pd.read_excel('validation_prices.xlsx', sheet_name='hist_hourly')
Daily_hist = pd.read_excel('validation_prices.xlsx', sheet_name='hist_daily')

#determining number of days and hours for multivariate regression
num_days = int(len(Hourly_sim)/24)
num_hours = len(Hourly_sim)

#creating a linear regression model for hourly simulations
hourly_x = Hourly_sim.copy()
hourly_y = Hourly_hist.loc[:,'LMP'].copy()
hourly_reg = linear_model.ElasticNet(positive=True, max_iter=100000)
hourly_reg.fit(hourly_x,hourly_y)

sim_hourly = np.zeros((num_hours,1))

#finding zonal weights by using the regression and predicting overall LMP
for i in range(0,num_hours):
    
    s = Hourly_sim.loc[i,:].values
    s = s.reshape((1,len(s)))
    sim_hourly[i] = hourly_reg.predict(s)
    
#saving the weighted hourly prices   
SH = pd.DataFrame(sim_hourly)
SH.columns = ['NEISO']
SH.to_excel('weighted_hourly_prices.xlsx')    
     
#creating a linear regression model for daily simulations
daily_x = Daily_sim.copy()
daily_y = Daily_hist.loc[:,'LMP'].copy()
daily_reg = linear_model.ElasticNet(positive=True, max_iter=100000)
daily_reg.fit(daily_x,daily_y)
    
sim_daily = np.zeros((num_days,1))

#finding zonal weights by using the regression and predicting overall LMP
for i in range(0,num_days):
    
    s = Daily_sim.loc[i,:].values
    s = s.reshape((1,len(s)))
    sim_daily[i] = daily_reg.predict(s)

#saving the weighted daily prices 
SD = pd.DataFrame(sim_daily)
SD.columns = ['NEISO']
SD.to_excel('weighted_daily_prices.xlsx')  

#showing the R^2 value for both linear regressions
W_daily = pd.read_excel('weighted_daily_prices.xlsx')
W_daily = np.asarray(W_daily['NEISO'])
H_daily = np.asarray(Daily_hist['LMP'])
print('R-squared value for hourly prices is ', hourly_reg.score(hourly_x, hourly_y))
print('R-squared value for daily prices is ', daily_reg.score(daily_x, daily_y))

#Creating a figure showing both simulated and historical daily LMPs
plt.style.use('seaborn-whitegrid')
plt.rcParams.update({'font.size': 30}) 
plt.rcParams['font.sans-serif'] = "Arial"
axis_fontsize=45
axis_label_pad = 10
tick_pad = 5
fig, ax = plt.subplots(figsize=(25,10))
ax.plot(W_daily, label='Simulated Daily LMP', linewidth=2)
ax.plot(H_daily, label='Historical Daily LMP', linewidth=2)
ax.legend(loc='best')
ax.set_ylabel('Day-Ahead Price ($/MWh)', labelpad=axis_label_pad, weight = 'bold', fontsize=axis_fontsize)
ax.set_xlabel('Date', labelpad=axis_label_pad, weight = 'bold', fontsize=axis_fontsize)
# ax.set_xlim([0,1092])
ax.set_xticks([0,91,182,273,364,455,546,637,728,819,910,1000,1095])
ax.set_xticklabels(['Jan \n2015','Apr \n2015','Jul \n2015','Oct \n2015','Jan \n2016','Apr \n2016','Jul \n2016','Oct \n2016','Jan \n2017','Apr \n2017','Jul \n2017','Oct \n2017','Jan \n2018'])
ax.text(476, 150, "Daily $\mathregular{R^{2}}$"+'='+str(round(daily_reg.score(daily_x, daily_y),2)), color='black')
# ax.title('Locational Marginal Price (LMP) Validation')
ax.tick_params(axis='both', which='both', pad=tick_pad)
plt.tight_layout()
plt.savefig('Validation.png', bbox_inches='tight', dpi=200)
plt.show()
plt.clf()

#Creating figures for the daily and hourly errors
Daily_Error = Daily_hist['LMP'] - SD['NEISO']
Hourly_Error = Hourly_hist['LMP'] - SH['NEISO']

plt.style.use('seaborn-whitegrid')
plt.rcParams.update({'font.size': 14}) 
plt.rcParams['font.sans-serif'] = "Arial"
axis_fontsize=16
axis_label_pad = 10
tick_pad = 5
fig, ax = plt.subplots(figsize=(10,6))
sns.histplot(Daily_Error, bins=40,ax=ax)
# ax.set_title('Daily Errors')
ax.set_xlabel('Difference between Historical and Simulated Daily Price ($/MWh)', labelpad=axis_label_pad, weight = 'bold', fontsize=axis_fontsize)
ax.set_ylabel('Frequency', labelpad=axis_label_pad, weight = 'bold', fontsize=axis_fontsize)
ax.tick_params(axis='both', which='both', pad=tick_pad)
plt.savefig('Daily_Error.png', bbox_inches='tight', dpi=200)
plt.show()
plt.clf()

plt.style.use('seaborn-whitegrid')
plt.rcParams.update({'font.size': 14}) 
plt.rcParams['font.sans-serif'] = "Arial"
axis_fontsize=16
axis_label_pad = 10
tick_pad = 5
fig, ax = plt.subplots(figsize=(10,6))
sns.histplot(Hourly_Error, bins=40, ax=ax)
# plt.title('Hourly Errors')
ax.set_xlabel('Difference between Historical and Simulated Hourly Price ($/MWh)', labelpad=axis_label_pad, weight = 'bold', fontsize=axis_fontsize)
ax.set_ylabel('Frequency', labelpad=axis_label_pad, weight = 'bold', fontsize=axis_fontsize)
ax.tick_params(axis='both', which='both', pad=tick_pad)
plt.savefig('Hourly_Error.png', bbox_inches='tight', dpi=200)
plt.show()
plt.clf()
    
