# -*- coding: utf-8 -*-
"""
Created on Wed Mar 21 10:00:33 2018

@author: jdkern
"""

from __future__ import division
from sklearn import linear_model
from statsmodels.tsa.api import VAR
import pandas as pd
import numpy as np

######################################################################
#                                LOAD
######################################################################

#import data
df_load = pd.read_excel('All_Data.xlsx',sheet_name='Hourly_load',header=0)
df_weather = pd.read_excel('All_Data.xlsx',sheet_name='Regression_weather',header=0)
NEISO_weights = pd.read_excel('All_Data.xlsx',sheet_name='NEISO_location_weights',header=0)
synthetic_weather_all = pd.read_excel('All_Data.xlsx',sheet_name='Weather',header=0)

# Name_list=pd.read_csv('Synthetic_demand_pathflows/Covariance_Calculation.csv')
# Name_list=list(Name_list.loc['SALEM_T':])
# Name_list=Name_list[1:]

Name_list = ['CT_T', 'CT_W', 'ME_T', 'ME_W', 'NEMA_T', 'NEMA_W', 'NH_T', 'NH_W', 'RI_T', 'RI_W', 'SEMA_T', 'SEMA_W', 'VT_T', 'VT_W', 'WCMA_T', 'WCMA_W']

# df_wind=pd.read_csv('Synthetic_wind_power/wind_power_sim.csv',header=0)
sim_years = int(len(df_load)/8760)
sim_weather=pd.read_csv('synthetic_weather_data.csv', header=0)
# sim_weather = sim_weather.iloc[0:365*sim_years,:]
# sim_weather = sim_weather.iloc[365:len(sim_weather)-730,:]
# sim_weather = sim_weather.reset_index(drop=True)

#weekday designation
dow = df_weather.loc[:,'Weekday']


#generate simulated day of the week assuming starts from monday
count=0
sim_dow= np.zeros(len(sim_weather))
for i in range(0,len(sim_weather)):
    count = count +1
    if count <=5:
        sim_dow[i]=1
    elif count > 5:
        sim_dow[i]=0
    
    if count ==7:
        count =0

#Generate a datelist
datelist=pd.date_range(pd.datetime(2017,1,1),periods=365).tolist()  
sim_month=np.zeros(len(sim_weather))  
sim_day=np.zeros(len(sim_weather))
sim_year=np.zeros(len(sim_weather))

count=0
for i in range(0,len(sim_weather)):
    
    if count <=364:
        sim_month[i]=datelist[count].month
        sim_day[i]=datelist[count].day
        sim_year[i]=datelist[count].year
    else:
        count=0
        sim_month[i]=datelist[count].month
        sim_day[i]=datelist[count].day
        sim_year[i]=datelist[count].year        
    count=count+1

######################################################################
#                                DEMAND_SIMULATION
######################################################################
#Find the simulated data at the sites

col_T = ['CT_T', 'ME_T', 'NEMA_T', 'NH_T', 'RI_T', 'SEMA_T', 'VT_T', 'WCMA_T']
col_W = ['CT_W', 'ME_W', 'NEMA_W', 'NH_W', 'RI_W', 'SEMA_W', 'VT_W', 'WCMA_W']

sim_T=sim_weather[col_T].values
sim_W=sim_weather[col_W].values

hist_T = df_weather[col_T].values
hist_W = df_weather[col_W].values

sim_days = len(sim_weather)

weighted_SimT = np.zeros((sim_days,1))

zones = ['CT','ME','NEMA','NH','RI','SEMA','VT','WCMA']

for zone in zones:
    
    zone_T = zone + '_T'
    ###########################################
    num_days = len(df_weather)
        
    #Convert simulated temperature to F
    weighted_SimT=sim_weather[zone_T].values 
    weighted_histT = df_weather[zone_T].values
    # BPA_sim_T_F=(BPA_sim_T * 9/5) +32   
           
    #convert to degree days
    HDD = np.zeros((num_days,len(col_T)))
    CDD = np.zeros((num_days,len(col_T)))
    
    HDD_sim = np.zeros((sim_days,len(col_T)))
    CDD_sim = np.zeros((sim_days,len(col_T)))
    
    for i in range(0,num_days):
        for j in range(0,len(col_T)):
            HDD[i,j] = np.max((0,65-hist_T[i,j]))
            CDD[i,j] = np.max((0,hist_T[i,j] - 65))
    
    for i in range(0,sim_days):
        for j in range(0,len(col_T)):
            HDD_sim[i,j] = np.max((0,65-sim_T[i,j]))
            CDD_sim[i,j] = np.max((0,sim_T[i,j] - 65)) 
    
    #separate wind speed by cooling/heating degree day
    binary_CDD = CDD>0
    binary_HDD = HDD>0
    CDD_wind = np.multiply(hist_W,binary_CDD)
    HDD_wind = np.multiply(hist_W,binary_HDD)
    
    binary_CDD_sim = CDD_sim > 0
    binary_HDD_sim = HDD_sim > 0
    CDD_wind_sim = np.multiply(sim_W,binary_CDD_sim)
    HDD_wind_sim = np.multiply(sim_W,binary_HDD_sim)
    
    #convert load to array 
    load = df_load.loc[:,zone].values
    
    #remove NaNs
    a = np.argwhere(np.isnan(load))
    for i in a:
        load[i] = load[i+24]   
    
    peaks = np.zeros((num_days,1))
    
    #find peaks
    for i in range(0,num_days):
        peaks[i] = np.max(load[i*24:i*24+24])
    
    d_max = max(peaks)
    d_min = min(peaks)
     
    #Separate data by weighted temperature
    M = np.column_stack((weighted_histT,peaks,dow,HDD,CDD,HDD_wind,CDD_wind))
    M_sim=np.column_stack((weighted_SimT,sim_dow,HDD_sim,CDD_sim,HDD_wind_sim,CDD_wind_sim))
    
    X80p = M[(M[:,0] >= 80),2:]
    y80p = M[(M[:,0] >= 80),1] 
    X75_80 = M[(M[:,0] >= 75) & (M[:,0] < 80),2:]
    y75_80 = M[(M[:,0] >= 75) & (M[:,0] < 80),1]
    X70_75 = M[(M[:,0] >= 70) & (M[:,0] < 75),2:]
    y70_75 = M[(M[:,0] >= 70) & (M[:,0] < 75),1]
    X65_70 = M[(M[:,0] >= 65) & (M[:,0] < 70),2:]
    y65_70 = M[(M[:,0] >= 65) & (M[:,0] < 70),1]  
    X60_65 = M[(M[:,0] >= 60) & (M[:,0] < 65),2:]
    y60_65 = M[(M[:,0] >= 60) & (M[:,0] < 65),1]  
    X55_60 = M[(M[:,0] >= 55) & (M[:,0] < 60),2:]
    y55_60 = M[(M[:,0] >= 55) & (M[:,0] < 60),1]  
    X50_55 = M[(M[:,0] >= 50) & (M[:,0] < 55),2:]
    y50_55 = M[(M[:,0] >= 50) & (M[:,0] < 55),1]      
    X40_50 = M[(M[:,0] >= 40) & (M[:,0] < 50),2:]
    y40_50 = M[(M[:,0] >= 40) & (M[:,0] < 50),1]      
    X30_40 = M[(M[:,0] >= 30) & (M[:,0] < 40),2:]
    y30_40 = M[(M[:,0] >= 30) & (M[:,0] < 40),1] 
    X25_30 = M[(M[:,0] >= 25) & (M[:,0] < 30),2:]
    y25_30 = M[(M[:,0] >= 25) & (M[:,0] < 30),1]  
    X20_25 = M[(M[:,0] >= 20) & (M[:,0] < 25),2:]
    y20_25 = M[(M[:,0] >= 20) & (M[:,0] < 25),1] 
    X15_20 = M[(M[:,0] >= 15) & (M[:,0] < 20),2:]
    y15_20 = M[(M[:,0] >= 15) & (M[:,0] < 20),1] 
    X10_15 = M[(M[:,0] >= 10) & (M[:,0] < 15),2:]
    y10_15 = M[(M[:,0] >= 10) & (M[:,0] < 15),1] 
    X10m = M[(M[:,0] < 10),2:]
    y10m = M[(M[:,0] < 10),1]  
    
    X80p_Sim = M_sim[(M_sim[:,0] >= 80),1:]
    X75_80_Sim = M_sim[(M_sim[:,0] >= 75) & (M_sim[:,0] < 80),1:]
    X70_75_Sim = M_sim[(M_sim[:,0] >= 70) & (M_sim[:,0] < 75),1:]
    X65_70_Sim = M_sim[(M_sim[:,0] >= 65) & (M_sim[:,0] < 70),1:]
    X60_65_Sim = M_sim[(M_sim[:,0] >= 60) & (M_sim[:,0] < 65),1:]
    X55_60_Sim = M_sim[(M_sim[:,0] >= 55) & (M_sim[:,0] < 60),1:]
    X50_55_Sim = M_sim[(M_sim[:,0] >= 50) & (M_sim[:,0] < 55),1:]
    X40_50_Sim = M_sim[(M_sim[:,0] >= 40) & (M_sim[:,0] < 50),1:]
    X30_40_Sim = M_sim[(M_sim[:,0] >= 30) & (M_sim[:,0] < 40),1:]
    X25_30_Sim = M_sim[(M_sim[:,0] >= 25) & (M_sim[:,0] < 30),1:]
    X20_25_Sim = M_sim[(M_sim[:,0] >= 20) & (M_sim[:,0] < 25),1:]
    X15_20_Sim = M_sim[(M_sim[:,0] >= 15) & (M_sim[:,0] < 20),1:]
    X10_15_Sim = M_sim[(M_sim[:,0] >= 10) & (M_sim[:,0] < 15),1:]
    X10m_Sim = M_sim[(M_sim[:,0] < 10),1:]
    #multivariate regression
    
    #Create linear regression object
    reg80p = linear_model.LinearRegression()
    reg75_80 = linear_model.LinearRegression()
    reg70_75 = linear_model.LinearRegression()
    reg65_70 = linear_model.LinearRegression()
    reg60_65 = linear_model.LinearRegression()
    reg55_60 = linear_model.LinearRegression()
    reg50_55 = linear_model.LinearRegression()
    reg40_50 = linear_model.LinearRegression()
    reg30_40 = linear_model.LinearRegression()
    reg25_30 = linear_model.LinearRegression()
    reg20_25 = linear_model.LinearRegression()
    reg15_20 = linear_model.LinearRegression()
    reg10_15 = linear_model.LinearRegression()
    reg10m = linear_model.LinearRegression()
    
    # Train the model using the training sets
    if len(y80p) > 0:
        reg80p.fit(X80p,y80p)    
    if len(y75_80) > 0:
        reg75_80.fit(X75_80,y75_80)       
    if len(y70_75) > 0:
        reg70_75.fit(X70_75,y70_75)      
    if len(y65_70) > 0:
        reg65_70.fit(X65_70,y65_70)
    if len(y60_65) > 0:
        reg60_65.fit(X60_65,y60_65)
    if len(y55_60) > 0:
        reg55_60.fit(X55_60,y55_60)
    if len(y50_55) > 0:
        reg50_55.fit(X50_55,y50_55)
    if len(y40_50) > 0:
        reg40_50.fit(X40_50,y40_50)
    if len(y30_40) > 0:
        reg30_40.fit(X30_40,y30_40)
    if len(y25_30) > 0:
        reg25_30.fit(X25_30,y25_30)  
    if len(y20_25) > 0:
        reg20_25.fit(X20_25,y20_25)      
    if len(y15_20) > 0:
        reg15_20.fit(X15_20,y15_20)       
    if len(y10_15) > 0:
        reg10_15.fit(X10_15,y10_15)                 
    if len(y10m) > 0:
        reg10m.fit(X10m,y10m)
    
    # Make predictions using the testing set
    
    predicted = []
    for i in range(0,num_days):
        s = M[i,2:]
        s = s.reshape((1,len(s)))
        if M[i,0]>=80:      
            y_hat = reg80p.predict(s)    
        elif M[i,0] >= 75 and M[i,0] < 80:
            y_hat = reg75_80.predict(s)          
        elif M[i,0] >= 70 and M[i,0] < 75:
            y_hat = reg70_75.predict(s)                 
        elif M[i,0] >= 65 and M[i,0] < 70:
            y_hat = reg65_70.predict(s)
        elif M[i,0] >= 60 and M[i,0] < 65:
            y_hat = reg60_65.predict(s)        
        elif M[i,0] >= 55 and M[i,0] < 60:
            y_hat = reg55_60.predict(s)
        elif M[i,0] >= 50 and M[i,0] < 55:
            y_hat = reg50_55.predict(s)          
        elif M[i,0] >= 40 and M[i,0] < 50:
            y_hat = reg40_50.predict(s)
        elif M[i,0] >= 30 and M[i,0] < 40:
            y_hat = reg30_40.predict(s)
        elif M[i,0] >= 25 and M[i,0] < 30:
            y_hat = reg25_30.predict(s)
        elif M[i,0] >= 20 and M[i,0] < 25:
            y_hat = reg20_25.predict(s)              
        elif M[i,0] >= 15 and M[i,0] < 20:
            y_hat = reg15_20.predict(s)          
        elif M[i,0] >= 10 and M[i,0] < 15:
            y_hat = reg10_15.predict(s)            
        elif M[i,0] < 10:
            y_hat = reg10m.predict(s)
            
        predicted = np.append(predicted,y_hat)
        p = predicted.reshape((len(predicted),1))
    
    #Simulate using the regression above
    simulated=[]
    for i in range(0,sim_days):
        s = M_sim[i,1:]
        s = s.reshape((1,len(s)))
        if M_sim[i,0]>=80:      
            y_hat = reg80p.predict(s)
        elif M_sim[i,0] >= 75 and M_sim[i,0] < 80:
            y_hat = reg75_80.predict(s)   
        elif M_sim[i,0] >= 70 and M_sim[i,0] < 75:
            y_hat = reg70_75.predict(s)      
        elif M_sim[i,0] >= 65 and M_sim[i,0] < 70:
            y_hat = reg65_70.predict(s)
        elif M_sim[i,0] >= 60 and M_sim[i,0] < 65:
            y_hat = reg60_65.predict(s)        
        elif M_sim[i,0] >= 55 and M_sim[i,0] < 60:
            y_hat = reg55_60.predict(s)
        elif M_sim[i,0] >= 50 and M_sim[i,0] < 55:
            y_hat = reg50_55.predict(s)          
        elif M_sim[i,0] >= 40 and M_sim[i,0] < 50:
            y_hat = reg40_50.predict(s)
        elif M_sim[i,0] >= 30 and M_sim[i,0] < 40:
            y_hat = reg30_40.predict(s)
        elif M_sim[i,0] >= 25 and M_sim[i,0] < 30:
            y_hat = reg25_30.predict(s)
        elif M_sim[i,0] >= 20 and M_sim[i,0] < 25:
            y_hat = reg20_25.predict(s)        
        elif M_sim[i,0] >= 15 and M_sim[i,0] < 20:
            y_hat = reg15_20.predict(s)           
        elif M_sim[i,0] >= 10 and M_sim[i,0] < 15:
            y_hat = reg10_15.predict(s)                    
        elif M_sim[i,0] < 10:
            y_hat = reg10m.predict(s)
    
        if y_hat > 1.10*d_max:
            y_hat = 1.10*d_max
        elif y_hat <.95*d_min:
            y_hat = .95*d_min
        simulated = np.append(simulated,y_hat)
        sim = simulated.reshape((len(simulated),1))  
    
    #a=st.pearsonr(peaks,p)
    #print(a[0]**2, a[1])
    
    # Residuals
    Residuals = p - peaks
    
#    # RMSE
#    RMSE = (np.sum((Residuals**2))/len(Residuals))**.5
    
# #Collect residuals from load regression    
    zone_index = zones.index(zone)
    if zone_index < 1:
        C = Residuals
    else:
        C = np.column_stack((C,Residuals))
        
    #Collect simulated data

    if zone_index < 1:
        combined_sim = sim
    else:
        combined_sim = np.column_stack((combined_sim,sim))    
    

 #####################################################################
 #                       Residual Analysis
 #####################################################################
    
R = C
rc = np.shape(R)
cols = rc[1]
mus = np.zeros((cols,1))
stds = np.zeros((cols,1))
R_w = np.zeros(np.shape(R))
sim_days = len(R_w)

#whiten residuals
for i in range(0,cols):
    mus[i] = np.mean(R[:,i])
    stds[i] = np.std(R[:,i])
    R_w[:,i] = (R[:,i] - mus[i])/stds[i]

#Vector autoregressive model on residuals
model = VAR(R_w)
results = model.fit(1)
sim_residuals = np.zeros((sim_days,cols))
errors = np.zeros((sim_days,cols))
p = results.params
y_seeds = R_w[-1]
C = results.sigma_u
means = [0,0,0,0,0,0,0,0]
E  = np.random.multivariate_normal(means,C,sim_days)
ys = np.zeros((cols,1))

# Generate cross correlated residuals
for i in range(0,sim_days):
    
    for j in range(1,cols+1):
        name='y' + str(j) 
        locals()[name]= p[0,j-1] + p[1,j-1]*y_seeds[0]+ p[2,j-1]*y_seeds[1]+ p[3,j-1]*y_seeds[2]+ p[4,j-1]*y_seeds[3]+ p[5,j-1]*y_seeds[4]+ p[6,j-1]*y_seeds[5]+ p[7,j-1]*y_seeds[6]

    for j in range(1,cols+1):
        name='y' + str(j) 
        y_seeds[j-1]=locals()[name]
    
    sim_residuals[i,:] = [y1,y2,y3,y4,y5,y6,y7,y8]
            

for i in range(0,cols):
    sim_residuals[:,i] = sim_residuals[:,i]*stds[i]*(1/np.std(sim_residuals[:,i])) + mus[i]

# ##########################################################################################################################################################
# #Simulating demand
# #########################################################################################################################################################

#Sim Residual
simulation_length=len(sim_weather)

syn_residuals = np.zeros((simulation_length,cols))
errors = np.zeros((simulation_length,cols))
y_seeds = R_w[-1]
C = results.sigma_u
means = [0,0,0,0,0,0,0,0]
E  = np.random.multivariate_normal(means,C,simulation_length)
ys = np.zeros((cols,1))


for i in range(0,simulation_length):
   
    for n in range(0,cols):
        
        ys[n] = p[0,n] 
        
        for m in range(0,cols):
    
            ys[n] = ys[n] + p[m+1,n]*y_seeds[n]
                
        ys[n] = ys[n] + E[i,n]
        
    for n in range(0,cols):
        y_seeds[n] = ys[n]
        
    syn_residuals[i,:] = np.reshape([ys],(1,cols))
            

for i in range(0,cols):
    syn_residuals[:,i] = syn_residuals[:,i]*stds[i]*(1/np.std(syn_residuals[:,i])) + mus[i]


# ############################################################################

daily_load = combined_sim + syn_residuals


###############################################################################
Demand=pd.DataFrame()
for zone in zones:
    zone_index = zones.index(zone)
    Demand[zone]=daily_load[:,zone_index].tolist()

Demand.to_csv('Load_Sim.csv')

######################################################################
##                         Hourly Demand Simulation
######################################################################

# number of historical days
hist_days = num_days
hist_years = int(hist_days/365)
sim_years = int(len(sim_weather)/365)

# create profiles
# simulate using synthetic peaks
hourly = np.zeros((8760*sim_years,len(zones)))

for zone in zones:
    
    profile = np.zeros((24,365))
    
    zone_index = zones.index(zone)
    
    for i in range(0,hist_years):
        for j in range(0,365):
            
            # pull 24 hours of demand
            sample = df_load.loc[i*8760+j*24:i*8760+j*24+23,zone]
            sample = sample.values
#            sample = np.reshape(sample,(24,1))
            
            # create fractional profile (relative to peak demand)
            sample_peak = np.max(sample)
            fraction = sample/sample_peak
            profile[:,j] = profile[:,j] + fraction*(1/hist_years)
     
    for i in range(0,sim_years):
        for j in range(0,365):
             v = Demand.loc[i*365+j,zone]*profile[:,j]
#             a = np.reshape(v,(24,1))
             hourly[i*8760+24*j:i*8760+24*j+24,zone_index] = v
            
df_H = pd.DataFrame(hourly)
df_H.columns = zones
df_H.to_csv('Sim_hourly_load.csv')


