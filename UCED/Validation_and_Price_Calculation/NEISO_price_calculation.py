# -*- coding: utf-8 -*-
"""
Created on Sat Jun 13 21:23:17 2020

@author: kakdemi
"""

from sklearn import linear_model
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

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
hourly_reg = linear_model.ElasticNet(positive=True, max_iter=10000)
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
daily_reg = linear_model.ElasticNet(positive=True, max_iter=10000)
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
W_hourly = pd.read_excel('weighted_hourly_prices.xlsx')
W_daily = np.asarray(W_daily['NEISO'])
W_hourly = np.asarray(W_hourly['NEISO'])
H_hourly = np.asarray(Hourly_hist['LMP'])
H_daily = np.asarray(Daily_hist['LMP'])
print('R-squared value for hourly prices is ', hourly_reg.score(hourly_x, hourly_y))
print('R-squared value for daily prices is ', daily_reg.score(daily_x, daily_y))

#Creating a figure showing both simulated and historical daily LMPs
plt.style.use('seaborn-whitegrid')
plt.figure(figsize=(35,10))
plt.rcParams.update({'font.size': 20})
plt.plot(W_daily, label='Simulated Daily LMP')
plt.plot(H_daily, label='Historical Daily LMP')
plt.legend(loc='best')
plt.ylabel('Day-Ahead Locational Marginal Price ($/MWh)')
plt.xlim([0,1092])
plt.xticks([0,95,190,285,380,475,570,665,760,855,950,1045 ], ['Jan 2015','Apr 2015','July 2015','Oct 2015','Jan 2016','Apr 2016','July 2016','Oct 2016','Jan 2017','Apr 2017','July 2017','Oct 2017'])
plt.savefig('Validation.png', bbox_inches='tight', dpi=200)
plt.show()
plt.clf()

Daily_Error = Daily_hist['LMP'] - SD['NEISO']
Hourly_Error = Hourly_hist['LMP'] - SH['NEISO']

plt.style.use('default')
plt.style.use('seaborn-whitegrid')
plt.hist(Daily_Error, bins=40)
plt.title('Daily Errors')
plt.xlabel('Difference between Historical and Simulated Daily LMP ($/MWh)')
plt.ylabel('Frequency')
plt.savefig('Daily_Error.png', bbox_inches='tight', dpi=200)
plt.show()
plt.clf()
plt.hist(Hourly_Error, bins=40)
plt.title('Hourly Errors')
plt.xlabel('Difference between Historical and Simulated Hourly LMP ($/MWh)')
plt.ylabel('Frequency')
plt.savefig('Hourly_Error.png', bbox_inches='tight', dpi=200)
plt.show()
plt.clf()
    
