# -*- coding: utf-8 -*-
"""
Created on Tue Mar 10 15:57:24 2020

@author: jdkern
"""

from __future__ import division
from sklearn import linear_model
from statsmodels.tsa.api import VAR
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

###################################
#      INTERCHANGE PATHFLOWS
###################################

#import data
synthetic_weather_all = pd.read_excel('All_Data.xlsx',sheet_name='Weather',header=0)
df_data1 = pd.read_excel('Interchange.xlsx',sheet_name='Daily', header=0)
df_data1 = df_data1.iloc[1460:,:]
#find average temps 
cities = ['CT','ME','NEMA','NH','RI','SEMA','VT','WCMA']
num_cities = len(cities)
num_days = len(df_data1)

AvgT = np.zeros((num_days,num_cities))
Wind = np.zeros((num_days,num_cities))

for i in cities:
    n1 = i + '_T'
    n2 = i + '_W'
    
    j = int(cities.index(i))
    
    AvgT[:,j] = df_data1.loc[:,n1] 
    Wind[:,j] = df_data1.loc[:,n2]
       
#convert to degree days
HDD = np.zeros((num_days,num_cities))
CDD = np.zeros((num_days,num_cities))


for i in range(0,num_days):
    for j in range(0,num_cities):
        HDD[i,j] = np.max((0,65-AvgT[i,j]))
        CDD[i,j] = np.max((0,AvgT[i,j] - 65))


#separate wind speed by cooling/heating degree day
binary_CDD = CDD>0
binary_HDD = HDD>0
CDD_wind = np.multiply(Wind,binary_CDD)
HDD_wind = np.multiply(Wind,binary_HDD)


X1 = np.array(df_data1.loc[:,'Month':'Year'])
X2 = np.array(df_data1.loc[:,'Weekday'])
X3 = np.column_stack((HDD,CDD,HDD_wind,CDD_wind))

Y = np.array(df_data1.loc[:,'SALBRYNB':'NORTHPORT'])
cXY = np.column_stack((Y,X1,X2,X3))
df_data = pd.DataFrame(cXY)
df_data.rename(columns={6:'Month'}, inplace=True)
df_data.rename(columns={9:'Weekday'}, inplace=True)


jan = df_data.loc[df_data['Month'] == 1,:]
feb = df_data.loc[df_data['Month'] == 2,:]
mar = df_data.loc[df_data['Month'] == 3,:]
apr = df_data.loc[df_data['Month'] == 4,:]
may = df_data.loc[df_data['Month'] == 5,:]
jun = df_data.loc[df_data['Month'] == 6,:]
jul = df_data.loc[df_data['Month'] == 7,:]
aug = df_data.loc[df_data['Month'] == 8,:]
sep = df_data.loc[df_data['Month'] == 9,:]
oct = df_data.loc[df_data['Month'] == 10,:]
nov = df_data.loc[df_data['Month'] == 11,:]
dec = df_data.loc[df_data['Month'] == 12,:] 


#############################################################################
#      Synthetic data setup
#############################################################################


syn_days = len(synthetic_weather_all)

AvgT_all = np.zeros((syn_days,num_cities))
Wind_all = np.zeros((syn_days,num_cities))

for i in cities:
    n1_new = i + '_T'
    n2_new = i + '_W'
    
    j = int(cities.index(i))
    
    AvgT_all[:,j] = synthetic_weather_all.loc[:,n1_new] 
    Wind_all[:,j] = synthetic_weather_all.loc[:,n2_new]
       
#convert to degree days
HDD_new = np.zeros((syn_days,num_cities))
CDD_new = np.zeros((syn_days,num_cities))


for i in range(0,syn_days):
    for j in range(0,num_cities):
        HDD_new[i,j] = np.max((0,65-AvgT_all[i,j]))
        CDD_new[i,j] = np.max((0,AvgT_all[i,j] - 65))


#separate Wind_all speed by cooling/heating degree day
binary_CDD_new = CDD_new>0
binary_HDD_new = HDD_new>0
CDD_new_Wind_all = np.multiply(Wind_all,binary_CDD_new)
HDD_new_Wind_all = np.multiply(Wind_all,binary_HDD_new)


X1_new = np.array(synthetic_weather_all.loc[:,'Month':'Year'])
X2_new = np.array(synthetic_weather_all.loc[:,'Weekday'])
X3_new = np.column_stack((HDD_new,CDD_new,HDD_new_Wind_all,CDD_new_Wind_all))

cXY_new = np.column_stack((X1_new,X2_new,X3_new))
df_data_new = pd.DataFrame(cXY_new)
df_data_new.rename(columns={0:'Month'}, inplace=True)
df_data_new.rename(columns={3:'Weekday'}, inplace=True)


jan_new = df_data_new.loc[df_data_new['Month'] == 1,:]
feb_new = df_data_new.loc[df_data_new['Month'] == 2,:]
mar_new = df_data_new.loc[df_data_new['Month'] == 3,:]
apr_new = df_data_new.loc[df_data_new['Month'] == 4,:]
may_new = df_data_new.loc[df_data_new['Month'] == 5,:]
jun_new = df_data_new.loc[df_data_new['Month'] == 6,:]
jul_new = df_data_new.loc[df_data_new['Month'] == 7,:]
aug_new = df_data_new.loc[df_data_new['Month'] == 8,:]
sep_new = df_data_new.loc[df_data_new['Month'] == 9,:]
oct_new = df_data_new.loc[df_data_new['Month'] == 10,:]
nov_new = df_data_new.loc[df_data_new['Month'] == 11,:]
dec_new = df_data_new.loc[df_data_new['Month'] == 12,:] 

############################################################################

Interchange_line = ['SALBRYNB', 'ROSETON', 'HQ_P1_P2', 'HQHIGATE', 'SHOREHAM', 'NORTHPORT']
for line in Interchange_line:
    line_index = Interchange_line.index(line)
    y = df_data.iloc[:,line_index]

    #multivariate regression
    jan_reg = linear_model.LinearRegression()
    feb_reg = linear_model.LinearRegression()
    mar_reg = linear_model.LinearRegression()
    apr_reg = linear_model.LinearRegression()
    may_reg = linear_model.LinearRegression()
    jun_reg = linear_model.LinearRegression()
    jul_reg = linear_model.LinearRegression()
    aug_reg = linear_model.LinearRegression()
    sep_reg = linear_model.LinearRegression()
    oct_reg = linear_model.LinearRegression()
    nov_reg = linear_model.LinearRegression()
    dec_reg = linear_model.LinearRegression()


    # Train the model using the training sets
    jan_reg.fit(jan.loc[:,'Weekday':],jan.iloc[:,line_index])
    feb_reg.fit(feb.loc[:,'Weekday':],feb.iloc[:,line_index])
    mar_reg.fit(mar.loc[:,'Weekday':],mar.iloc[:,line_index])
    apr_reg.fit(apr.loc[:,'Weekday':],apr.iloc[:,line_index])
    may_reg.fit(may.loc[:,'Weekday':],may.iloc[:,line_index])
    jun_reg.fit(jun.loc[:,'Weekday':],jun.iloc[:,line_index])
    jul_reg.fit(jul.loc[:,'Weekday':],jul.iloc[:,line_index])
    aug_reg.fit(aug.loc[:,'Weekday':],aug.iloc[:,line_index])
    sep_reg.fit(sep.loc[:,'Weekday':],sep.iloc[:,line_index])
    oct_reg.fit(oct.loc[:,'Weekday':],oct.iloc[:,line_index])
    nov_reg.fit(nov.loc[:,'Weekday':],nov.iloc[:,line_index])
    dec_reg.fit(dec.loc[:,'Weekday':],dec.iloc[:,line_index])

    # Make predictions using the testing set
    predicted = []
    rc = np.shape(jan.loc[:,'Weekday':])
    n = rc[1] 
    
    for i in range(0,len(y)):
        
        m = df_data.loc[i,'Month']
        
        if m==1:
            s = jan.loc[i,'Weekday':]
            s = np.reshape(s[:,None],(1,n))
            p = jan_reg.predict(s)
            predicted = np.append(predicted,p)
        elif m==2:
            s = feb.loc[i,'Weekday':]
            s = np.reshape(s[:,None],(1,n))        
            p = feb_reg.predict(s)
            predicted = np.append(predicted,p)
        elif m==3:
            s = mar.loc[i,'Weekday':]
            s = np.reshape(s[:,None],(1,n))
            p = mar_reg.predict(s)
            predicted = np.append(predicted,p)
        elif m==4:
            s = apr.loc[i,'Weekday':]
            s = np.reshape(s[:,None],(1,n))
            p = apr_reg.predict(s)
            predicted = np.append(predicted,p)
        elif m==5:
            s = may.loc[i,'Weekday':]
            s = np.reshape(s[:,None],(1,n))
            p = may_reg.predict(s)
            predicted = np.append(predicted,p)
        elif m==6:
            s = jun.loc[i,'Weekday':]
            s = np.reshape(s[:,None],(1,n))
            p = jun_reg.predict(s)
            predicted = np.append(predicted,p)
        elif m==7:
            s = jul.loc[i,'Weekday':]
            s = np.reshape(s[:,None],(1,n))
            p = jul_reg.predict(s)
            predicted = np.append(predicted,p)
        elif m==8:
            s = aug.loc[i,'Weekday':]
            s = np.reshape(s[:,None],(1,n))
            p = aug_reg.predict(s)
            predicted = np.append(predicted,p)
        elif m==9:
            s = sep.loc[i,'Weekday':]
            s = np.reshape(s[:,None],(1,n))
            p = sep_reg.predict(s)
            predicted = np.append(predicted,p)
        elif m==10:
            s = oct.loc[i,'Weekday':]
            s = np.reshape(s[:,None],(1,n))
            p = oct_reg.predict(s)
            predicted = np.append(predicted,p)
        elif m==11:
            s = nov.loc[i,'Weekday':]
            s = np.reshape(s[:,None],(1,n))
            p = nov_reg.predict(s)
            predicted = np.append(predicted,p)
        else:
            s = dec.loc[i,'Weekday':]
            s = np.reshape(s[:,None],(1,n))
            p = dec_reg.predict(s)
            predicted = np.append(predicted,p)
        
    # Residuals
    residuals = predicted - y.values
    line_index = Interchange_line.index(line)
    
    if line_index < 1:
        
        Residuals_line = pd.DataFrame(residuals)
        Residuals_line.columns = [line]
        
    else:
        Residuals_line[line] = residuals     


######################## Interchange simulation #####################

    # Make predictions using the synthetic set
    predicted_new = []
    rc_new = np.shape(jan_new.loc[:,'Weekday':])
    n_new = rc_new[1] 
    
    for i in range(0,syn_days):
        
        t = df_data_new.loc[i,'Month']
        
        if t==1:
            s = jan_new.loc[i,'Weekday':]
            s = np.reshape(s[:,None],(1,n_new))
            p = jan_reg.predict(s)
            predicted_new = np.append(predicted_new,p)
        elif t==2:
            s = feb_new.loc[i,'Weekday':]
            s = np.reshape(s[:,None],(1,n_new))        
            p = feb_reg.predict(s)
            predicted_new = np.append(predicted_new,p)
        elif t==3:
            s = mar_new.loc[i,'Weekday':]
            s = np.reshape(s[:,None],(1,n_new))
            p = mar_reg.predict(s)
            predicted_new = np.append(predicted_new,p)
        elif t==4:
            s = apr_new.loc[i,'Weekday':]
            s = np.reshape(s[:,None],(1,n_new))
            p = apr_reg.predict(s)
            predicted_new = np.append(predicted_new,p)
        elif t==5:
            s = may_new.loc[i,'Weekday':]
            s = np.reshape(s[:,None],(1,n_new))
            p = may_reg.predict(s)
            predicted_new = np.append(predicted_new,p)
        elif t==6:
            s = jun_new.loc[i,'Weekday':]
            s = np.reshape(s[:,None],(1,n_new))
            p = jun_reg.predict(s)
            predicted_new = np.append(predicted_new,p)
        elif t==7:
            s = jul_new.loc[i,'Weekday':]
            s = np.reshape(s[:,None],(1,n_new))
            p = jul_reg.predict(s)
            predicted_new = np.append(predicted_new,p)
        elif t==8:
            s = aug_new.loc[i,'Weekday':]
            s = np.reshape(s[:,None],(1,n_new))
            p = aug_reg.predict(s)
            predicted_new = np.append(predicted_new,p)
        elif t==9:
            s = sep_new.loc[i,'Weekday':]
            s = np.reshape(s[:,None],(1,n_new))
            p = sep_reg.predict(s)
            predicted_new = np.append(predicted_new,p)
        elif t==10:
            s = oct_new.loc[i,'Weekday':]
            s = np.reshape(s[:,None],(1,n_new))
            p = oct_reg.predict(s)
            predicted_new = np.append(predicted_new,p)
        elif t==11:
            s = nov_new.loc[i,'Weekday':]
            s = np.reshape(s[:,None],(1,n_new))
            p = nov_reg.predict(s)
            predicted_new = np.append(predicted_new,p)
        else:
            s = dec_new.loc[i,'Weekday':]
            s = np.reshape(s[:,None],(1,n_new))
            p = dec_reg.predict(s)
            predicted_new = np.append(predicted_new,p)

    if line_index < 1:

        predictions_all = pd.DataFrame(predicted_new)
        predictions_all.columns = [line]
    
    else:
        predictions_all[line] = predicted_new
        
        
  #####################################################################
  #                       Residual Analysis for Interchange
  #####################################################################

R = Residuals_line.values
rc = np.shape(R)
cols = rc[1]
mus = np.zeros((cols,1))
stds = np.zeros((cols,1))
R_w = np.zeros(np.shape(R))
sim_days = len(predictions_all)

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
means = [0,0,0,0,0,0]
E  = np.random.multivariate_normal(means,C,sim_days)
ys = np.zeros((cols,1))

# Generate cross correlated residuals
for i in range(0,sim_days):
    
    for j in range(1,cols+1):
        name='y' + str(j) 
        locals()[name]= p[0,j-1] + p[1,j-1]*y_seeds[0]+ p[2,j-1]*y_seeds[1]+ p[3,j-1]*y_seeds[2]+ p[4,j-1]*y_seeds[3]+ p[5,j-1]*y_seeds[4]+ p[6,j-1]*y_seeds[5]+E[i,j-1]

    for j in range(1,cols+1):
        name='y' + str(j) 
        y_seeds[j-1]=locals()[name]
    
    sim_residuals[i,:] = [y1,y2,y3,y4,y5,y6]
            

for i in range(0,cols):
    sim_residuals[:,i] = sim_residuals[:,i]*stds[i]*(1/np.std(sim_residuals[:,i])) + mus[i]


simulated_errors = pd.DataFrame(sim_residuals)

# enact logical bounds
for line in Interchange_line:
    line_index = Interchange_line.index(line)
    for i in range(0,len(sim_residuals)):
        if predictions_all.loc[i,line] > np.max(Y[:,line_index]):
            predictions_all.loc[i,line] = np.max(Y[:,line_index])
        elif predictions_all.loc[i,line] < np.min(Y[:,line_index]):
            predictions_all.loc[i,line] = np.min(Y[:,line_index])
            
SALBRYNB_total = predictions_all.loc[:,'SALBRYNB'] + simulated_errors.loc[:,0]
ROSETON_total = predictions_all.loc[:,'ROSETON'] + simulated_errors.loc[:,1]
HQ_P1_P2_total = predictions_all.loc[:,'HQ_P1_P2'] + simulated_errors.loc[:,2]
HQHIGATE_total = predictions_all.loc[:,'HQHIGATE'] + simulated_errors.loc[:,3]
SHOREHAM_total = predictions_all.loc[:,'SHOREHAM'] + simulated_errors.loc[:,4]
NORTHPORT_total = predictions_all.loc[:,'NORTHPORT'] + simulated_errors.loc[:,5]
list_labels = ['SALBRYNB', 'ROSETON', 'HQ_P1_P2', 'HQHIGATE', 'SHOREHAM', 'NORTHPORT']       
list_columns = [SALBRYNB_total, ROSETON_total, HQ_P1_P2_total, HQHIGATE_total, SHOREHAM_total, NORTHPORT_total]
zipped_list = list(zip(list_labels, list_columns))
interchange_simulated = dict(zipped_list)
interchange_simulated_final = pd.DataFrame(interchange_simulated)
interchange_simulated_final.to_csv("Sim_daily_interchange.csv")

plt.figure()
plt.plot(SALBRYNB_total[0:3650])
plt.title('SALBRYNB')

plt.figure()
plt.plot(ROSETON_total[0:3650])
plt.title('ROSETON')

plt.figure()
plt.plot(HQ_P1_P2_total[0:3650])
plt.title('HQ_P1_P2')

plt.figure()
plt.plot(HQHIGATE_total[0:3650])
plt.title('HQHIGATE')

plt.figure()
plt.plot(SHOREHAM_total[0:3650])
plt.title('SHOREHAM')

plt.figure()
plt.plot(NORTHPORT_total[0:3650])
plt.title('NORTHPORT')
