# -*- coding: utf-8 -*-
"""
Created on Tue Jun 20 22:14:07 2017

@author: YSu
"""

from pyomo.opt import SolverFactory
from NEISO_dispatch import model
from pyomo.core import Var
from pyomo.core import Param
from operator import itemgetter
import pandas as pd
import numpy as np
from datetime import datetime

def sim(days):
    
    instance = model.create_instance('../Model_setup/NEISO_data_file/data.dat')
    
    opt = SolverFactory("cplex")
    H = instance.HorizonHours
    D = 2
    K=range(1,H+1)
    
    
    #Space to store results
    mwh_1=[]
    mwh_2=[]
    mwh_3=[]
    on=[]
    switch=[]
    srsv=[]
    nrsv=[]
    solar=[]
    wind=[]
    flow=[]
    Generator=[]
    
    df_generators = pd.read_excel('../Model_setup/NEISO_data_file/generators.xlsx',header=0)
    
    #max here can be (1,365)
    for day in range(1,days):
        
         #load time series data
        for z in instance.zones:
            
            instance.GasPrice[z] = instance.SimGasPrice[z,day]
            
            for i in K:
                instance.HorizonDemand[z,i] = instance.SimDemand[z,(day-1)*24+i]
                instance.HorizonWind[z,i] = instance.SimWind[z,(day-1)*24+i]
                instance.HorizonSolar[z,i] = instance.SimSolar[z,(day-1)*24+i]
                instance.HorizonMustRun[z,i] = instance.SimMustRun[z,(day-1)*24+i]
              
        
        for d in range(1,D+1):
            instance.HorizonNY_imports_CT[d] = instance.SimNY_imports_CT[day-1+d]
            instance.HorizonNY_imports_WCMA[d] = instance.SimNY_imports_WCMA[day-1+d]
            instance.HorizonNY_imports_VT[d] = instance.SimNY_imports_VT[day-1+d]
            instance.HorizonHQ_imports_VT[d] = instance.SimHQ_imports_VT[day-1+d]
            instance.HorizonHQ_imports_WCMA[d] = instance.SimHQ_imports_WCMA[day-1+d]
            instance.HorizonNB_imports_ME[d] = instance.SimNB_imports_ME[day-1+d]
            instance.HorizonME_hydro[d] = instance.SimME_hydro[day-1+d]
            instance.HorizonCT_hydro[d] = instance.SimCT_hydro[day-1+d]
            instance.HorizonNH_hydro[d] = instance.SimNH_hydro[day-1+d]
            instance.HorizonNEMA_hydro[d] = instance.SimNEMA_hydro[day-1+d]
            instance.HorizonRI_hydro[d] = instance.SimRI_hydro[day-1+d]
            instance.HorizonVT_hydro[d] = instance.SimVT_hydro[day-1+d]
            instance.HorizonWCMA_hydro[d] = instance.SimWCMA_hydro[day-1+d]
            
        for i in K:
            instance.HorizonReserves[i] = instance.SimReserves[(day-1)*24+i] 
            instance.HorizonCT_exports_NY[i] = instance.SimCT_exports_NY[(day-1)*24+i] 
            instance.HorizonWCMA_exports_NY[i] = instance.SimWCMA_exports_NY[(day-1)*24+i] 
            instance.HorizonVT_exports_NY[i] = instance.SimVT_exports_NY[(day-1)*24+i]             
            instance.HorizonVT_exports_HQ[i] = instance.SimVT_exports_HQ[(day-1)*24+i]  
            instance.HorizonWCMA_exports_HQ[i] = instance.SimWCMA_exports_HQ[(day-1)*24+i]             
            instance.HorizonME_exports_NB[i] = instance.SimME_exports_NB[(day-1)*24+i]             
            
            
            instance.HorizonNY_imports_CT_minflow[i] = instance.SimNY_imports_CT_minflow[(day-1)*24+i]             
            instance.HorizonNY_imports_WCMA_minflow[i] = instance.SimNY_imports_WCMA_minflow[(day-1)*24+i] 
            instance.HorizonNY_imports_VT_minflow[i] = instance.SimNY_imports_VT_minflow[(day-1)*24+i] 
            instance.HorizonHQ_imports_VT_minflow[i] = instance.SimHQ_imports_VT_minflow[(day-1)*24+i]  
            instance.HorizonHQ_imports_WCMA_minflow[i] = instance.SimHQ_imports_WCMA_minflow[(day-1)*24+i]
            instance.HorizonNB_imports_ME_minflow[i] = instance.SimNB_imports_ME_minflow[(day-1)*24+i]
            
            instance.HorizonME_hydro_minflow[i] = instance.SimME_hydro_minflow[(day-1)*24+i] 
            instance.HorizonCT_hydro_minflow[i] = instance.SimCT_hydro_minflow[(day-1)*24+i] 
            instance.HorizonNH_hydro_minflow[i] = instance.SimNH_hydro_minflow[(day-1)*24+i] 
            instance.HorizonNEMA_hydro_minflow[i] = instance.SimNEMA_hydro_minflow[(day-1)*24+i] 
            instance.HorizonRI_hydro_minflow[i] = instance.SimRI_hydro_minflow[(day-1)*24+i] 
            instance.HorizonVT_hydro_minflow[i] = instance.SimVT_hydro_minflow[(day-1)*24+i] 
            instance.HorizonWCMA_hydro_minflow[i] = instance.SimWCMA_hydro_minflow[(day-1)*24+i]             
         
         
        NEISO_result = opt.solve(instance)
        instance.solutions.load_from(NEISO_result)   
     
        #The following section is for storing and sorting results
        for v in instance.component_objects(Var, active=True):
            varobject = getattr(instance, str(v))
            a=str(v)
            if a=='mwh_1':
             
             for index in varobject:
                 
               name = index[0]   
               
               g = df_generators[df_generators['name']==name]
               seg1 = g['seg1'].values
               seg1 = seg1[0]
               
               if int(index[1]>0 and index[1]<25):
                if index[0] in instance.Zone1Generators:
                    
                    gas_price = instance.GasPrice['CT'].value
                    
                    if index[0] in instance.Gas:
                        marginal_cost = seg1*gas_price
                        mwh_1.append((index[0],index[1]+((day-1)*24),varobject[index].value,'CT','Gas',marginal_cost))
                    elif index[0] in instance.Coal:
                        marginal_cost = seg1*2
                        mwh_1.append((index[0],index[1]+((day-1)*24),varobject[index].value,'CT','Coal',marginal_cost))
                    elif index[0] in instance.Oil:
                        marginal_cost = seg1*20
                        mwh_1.append((index[0],index[1]+((day-1)*24),varobject[index].value,'CT','Oil',marginal_cost))
                    elif index[0] in instance.PSH:
                        marginal_cost = 10
                        mwh_1.append((index[0],index[1]+((day-1)*24),varobject[index].value,'CT','PSH',marginal_cost))
                    elif index[0] in instance.Slack:
                        marginal_cost = 700
                        mwh_1.append((index[0],index[1]+((day-1)*24),varobject[index].value,'CT','Slack',marginal_cost))               
                    elif index[0] in instance.Hydro:
                        marginal_cost = 0
                        mwh_1.append((index[0],index[1]+((day-1)*24),varobject[index].value,'CT','Hydro',marginal_cost))
                        
                        
                elif index[0] in instance.Zone2Generators:
                    
                    gas_price = instance.GasPrice['ME'].value
                    
                    if index[0] in instance.Gas:
                        marginal_cost = seg1*gas_price
                        mwh_1.append((index[0],index[1]+((day-1)*24),varobject[index].value,'ME','Gas',marginal_cost))
                    elif index[0] in instance.Coal:
                        marginal_cost = seg1*2
                        mwh_1.append((index[0],index[1]+((day-1)*24),varobject[index].value,'ME','Coal',marginal_cost))
                    elif index[0] in instance.Oil:
                        marginal_cost = seg1*20
                        mwh_1.append((index[0],index[1]+((day-1)*24),varobject[index].value,'ME','Oil',marginal_cost))
                    elif index[0] in instance.PSH:
                        marginal_cost = 10
                        mwh_1.append((index[0],index[1]+((day-1)*24),varobject[index].value,'ME','PSH',marginal_cost))
                    elif index[0] in instance.Slack:
                        marginal_cost = 700
                        mwh_1.append((index[0],index[1]+((day-1)*24),varobject[index].value,'ME','Slack',marginal_cost))  
        
                elif index[0] in instance.Zone3Generators:
                    
                    gas_price = instance.GasPrice['NH'].value
                    
                    if index[0] in instance.Gas:
                        marginal_cost = seg1*gas_price
                        mwh_1.append((index[0],index[1]+((day-1)*24),varobject[index].value,'NH','Gas',marginal_cost))
                    elif index[0] in instance.Coal:
                        marginal_cost = seg1*2
                        mwh_1.append((index[0],index[1]+((day-1)*24),varobject[index].value,'NH','Coal',marginal_cost))
                    elif index[0] in instance.Oil:
                        marginal_cost = seg1*20
                        mwh_1.append((index[0],index[1]+((day-1)*24),varobject[index].value,'NH','Oil',marginal_cost))
                    elif index[0] in instance.PSH:
                        marginal_cost = 10
                        mwh_1.append((index[0],index[1]+((day-1)*24),varobject[index].value,'NH','PSH',marginal_cost))
                    elif index[0] in instance.Slack:
                        marginal_cost = 700
                        mwh_1.append((index[0],index[1]+((day-1)*24),varobject[index].value,'NH','Slack',marginal_cost))  
                    elif index[0] in instance.Hydro:
                        marginal_cost = 0
                        mwh_1.append((index[0],index[1]+((day-1)*24),varobject[index].value,'NH','Hydro',marginal_cost))  
            
                elif index[0] in instance.Zone4Generators:
                    
                    gas_price = instance.GasPrice['NEMA'].value
                    
                    if index[0] in instance.Gas:
                        marginal_cost = seg1*gas_price
                        mwh_1.append((index[0],index[1]+((day-1)*24),varobject[index].value,'NEMA','Gas',marginal_cost))
                    elif index[0] in instance.Coal:
                        marginal_cost = seg1*2
                        mwh_1.append((index[0],index[1]+((day-1)*24),varobject[index].value,'NEMA','Coal',marginal_cost))
                    elif index[0] in instance.Oil:
                        marginal_cost = seg1*20
                        mwh_1.append((index[0],index[1]+((day-1)*24),varobject[index].value,'NEMA','Oil',marginal_cost))
                    elif index[0] in instance.PSH:
                        marginal_cost = 10
                        mwh_1.append((index[0],index[1]+((day-1)*24),varobject[index].value,'NEMA','PSH',marginal_cost))
                    elif index[0] in instance.Slack:
                        marginal_cost = 700
                        mwh_1.append((index[0],index[1]+((day-1)*24),varobject[index].value,'NEMA','Slack',marginal_cost))  
        

                elif index[0] in instance.Zone5Generators:
                    
                    gas_price = instance.GasPrice['RI'].value
                    
                    if index[0] in instance.Gas:
                        marginal_cost = seg1*gas_price
                        mwh_1.append((index[0],index[1]+((day-1)*24),varobject[index].value,'RI','Gas',marginal_cost))
                    elif index[0] in instance.Coal:
                        marginal_cost = seg1*2
                        mwh_1.append((index[0],index[1]+((day-1)*24),varobject[index].value,'RI','Coal',marginal_cost))
                    elif index[0] in instance.Oil:
                        marginal_cost = seg1*20
                        mwh_1.append((index[0],index[1]+((day-1)*24),varobject[index].value,'RI','Oil',marginal_cost))
                    elif index[0] in instance.PSH:
                        marginal_cost = 10
                        mwh_1.append((index[0],index[1]+((day-1)*24),varobject[index].value,'RI','PSH',marginal_cost))
                    elif index[0] in instance.Slack:
                        marginal_cost = 700
                        mwh_1.append((index[0],index[1]+((day-1)*24),varobject[index].value,'RI','Slack',marginal_cost))  

                elif index[0] in instance.Zone6Generators:
                    
                    gas_price = instance.GasPrice['SEMA'].value
                    
                    if index[0] in instance.Gas:
                        marginal_cost = seg1*gas_price
                        mwh_1.append((index[0],index[1]+((day-1)*24),varobject[index].value,'SEMA','Gas',marginal_cost))
                    elif index[0] in instance.Coal:
                        marginal_cost = seg1*2
                        mwh_1.append((index[0],index[1]+((day-1)*24),varobject[index].value,'SEMA','Coal',marginal_cost))
                    elif index[0] in instance.Oil:
                        marginal_cost = seg1*20
                        mwh_1.append((index[0],index[1]+((day-1)*24),varobject[index].value,'SEMA','Oil',marginal_cost))
                    elif index[0] in instance.PSH:
                        marginal_cost = 10
                        mwh_1.append((index[0],index[1]+((day-1)*24),varobject[index].value,'SEMA','PSH',marginal_cost))
                    elif index[0] in instance.Slack:
                        marginal_cost = 700
                        mwh_1.append((index[0],index[1]+((day-1)*24),varobject[index].value,'SEMA','Slack',marginal_cost))  
                    

                elif index[0] in instance.Zone7Generators:
                    
                    gas_price = instance.GasPrice['VT'].value
                    
                    if index[0] in instance.Gas:
                        marginal_cost = seg1*gas_price
                        mwh_1.append((index[0],index[1]+((day-1)*24),varobject[index].value,'VT','Gas',marginal_cost))
                    elif index[0] in instance.Coal:
                        marginal_cost = seg1*2
                        mwh_1.append((index[0],index[1]+((day-1)*24),varobject[index].value,'VT','Coal',marginal_cost))
                    elif index[0] in instance.Oil:
                        marginal_cost = seg1*20
                        mwh_1.append((index[0],index[1]+((day-1)*24),varobject[index].value,'VT','Oil',marginal_cost))
                    elif index[0] in instance.PSH:
                        marginal_cost = 10
                        mwh_1.append((index[0],index[1]+((day-1)*24),varobject[index].value,'VT','PSH',marginal_cost))
                    elif index[0] in instance.Slack:
                        marginal_cost = 700
                        mwh_1.append((index[0],index[1]+((day-1)*24),varobject[index].value,'VT','Slack',marginal_cost))  


                elif index[0] in instance.Zone8Generators:
                    
                    gas_price = instance.GasPrice['WCMA'].value
                    
                    if index[0] in instance.Gas:
                        marginal_cost = seg1*gas_price
                        mwh_1.append((index[0],index[1]+((day-1)*24),varobject[index].value,'WCMA','Gas',marginal_cost))
                    elif index[0] in instance.Coal:
                        marginal_cost = seg1*2
                        mwh_1.append((index[0],index[1]+((day-1)*24),varobject[index].value,'WCMA','Coal',marginal_cost))
                    elif index[0] in instance.Oil:
                        marginal_cost = seg1*20
                        mwh_1.append((index[0],index[1]+((day-1)*24),varobject[index].value,'WCMA','Oil',marginal_cost))
                    elif index[0] in instance.PSH:
                        marginal_cost = 10
                        mwh_1.append((index[0],index[1]+((day-1)*24),varobject[index].value,'WCMA','PSH',marginal_cost))
                    elif index[0] in instance.Slack:
                        marginal_cost = 700
                        mwh_1.append((index[0],index[1]+((day-1)*24),varobject[index].value,'WCMA','Slack',marginal_cost))  

    
            if a=='mwh_2':
    
             for index in varobject:
                 
               name = index[0]
               g = df_generators[df_generators['name']==name]
               seg2 = g['seg2'].values
               seg2 = seg2[0]
                 
               if int(index[1]>0 and index[1]<25):
                if index[0] in instance.Zone1Generators:
                    
                    gas_price = instance.GasPrice['CT'].value
                    
                    if index[0] in instance.Gas:
                        marginal_cost = seg2*gas_price
                        mwh_2.append((index[0],index[1]+((day-1)*24),varobject[index].value,'CT','Gas',marginal_cost))
                    elif index[0] in instance.Coal:
                        marginal_cost = seg2*2
                        mwh_2.append((index[0],index[1]+((day-1)*24),varobject[index].value,'CT','Coal',marginal_cost))
                    elif index[0] in instance.Oil:
                        marginal_cost = seg2*20
                        mwh_2.append((index[0],index[1]+((day-1)*24),varobject[index].value,'CT','Oil',marginal_cost))
                    elif index[0] in instance.PSH:
                        marginal_cost = 10
                        mwh_2.append((index[0],index[1]+((day-1)*24),varobject[index].value,'CT','PSH',marginal_cost))
                    elif index[0] in instance.Slack:
                        marginal_cost = 700
                        mwh_2.append((index[0],index[1]+((day-1)*24),varobject[index].value,'CT','Slack',marginal_cost))               
                    elif index[0] in instance.Hydro:
                        marginal_cost = 0
                        mwh_2.append((index[0],index[1]+((day-1)*24),varobject[index].value,'CT','Hydro',marginal_cost))
                        
                        
                elif index[0] in instance.Zone2Generators:
                    
                    gas_price = instance.GasPrice['ME'].value
                    
                    if index[0] in instance.Gas:
                        marginal_cost = seg2*gas_price
                        mwh_2.append((index[0],index[1]+((day-1)*24),varobject[index].value,'ME','Gas',marginal_cost))
                    elif index[0] in instance.Coal:
                        marginal_cost = seg2*2
                        mwh_2.append((index[0],index[1]+((day-1)*24),varobject[index].value,'ME','Coal',marginal_cost))
                    elif index[0] in instance.Oil:
                        marginal_cost = seg2*20
                        mwh_2.append((index[0],index[1]+((day-1)*24),varobject[index].value,'ME','Oil',marginal_cost))
                    elif index[0] in instance.PSH:
                        marginal_cost = 10
                        mwh_2.append((index[0],index[1]+((day-1)*24),varobject[index].value,'ME','PSH',marginal_cost))
                    elif index[0] in instance.Slack:
                        marginal_cost = 700
                        mwh_2.append((index[0],index[1]+((day-1)*24),varobject[index].value,'ME','Slack',marginal_cost))  
        
                elif index[0] in instance.Zone3Generators:
                    
                    gas_price = instance.GasPrice['NH'].value
                    
                    if index[0] in instance.Gas:
                        marginal_cost = seg2*gas_price
                        mwh_2.append((index[0],index[1]+((day-1)*24),varobject[index].value,'NH','Gas',marginal_cost))
                    elif index[0] in instance.Coal:
                        marginal_cost = seg2*2
                        mwh_2.append((index[0],index[1]+((day-1)*24),varobject[index].value,'NH','Coal',marginal_cost))
                    elif index[0] in instance.Oil:
                        marginal_cost = seg2*20
                        mwh_2.append((index[0],index[1]+((day-1)*24),varobject[index].value,'NH','Oil',marginal_cost))
                    elif index[0] in instance.PSH:
                        marginal_cost = 10
                        mwh_2.append((index[0],index[1]+((day-1)*24),varobject[index].value,'NH','PSH',marginal_cost))
                    elif index[0] in instance.Slack:
                        marginal_cost = 700
                        mwh_2.append((index[0],index[1]+((day-1)*24),varobject[index].value,'NH','Slack',marginal_cost))  
                    elif index[0] in instance.Hydro:
                        marginal_cost = 0
                        mwh_2.append((index[0],index[1]+((day-1)*24),varobject[index].value,'NH','Hydro',marginal_cost))  
            
                elif index[0] in instance.Zone4Generators:
                    
                    gas_price = instance.GasPrice['NEMA'].value
                    
                    if index[0] in instance.Gas:
                        marginal_cost = seg2*gas_price
                        mwh_2.append((index[0],index[1]+((day-1)*24),varobject[index].value,'NEMA','Gas',marginal_cost))
                    elif index[0] in instance.Coal:
                        marginal_cost = seg2*2
                        mwh_2.append((index[0],index[1]+((day-1)*24),varobject[index].value,'NEMA','Coal',marginal_cost))
                    elif index[0] in instance.Oil:
                        marginal_cost = seg2*20
                        mwh_2.append((index[0],index[1]+((day-1)*24),varobject[index].value,'NEMA','Oil',marginal_cost))
                    elif index[0] in instance.PSH:
                        marginal_cost = 10
                        mwh_2.append((index[0],index[1]+((day-1)*24),varobject[index].value,'NEMA','PSH',marginal_cost))
                    elif index[0] in instance.Slack:
                        marginal_cost = 700
                        mwh_2.append((index[0],index[1]+((day-1)*24),varobject[index].value,'NEMA','Slack',marginal_cost))  
        
                elif index[0] in instance.Zone5Generators:
                    
                    gas_price = instance.GasPrice['RI'].value
                    
                    if index[0] in instance.Gas:
                        marginal_cost = seg2*gas_price
                        mwh_2.append((index[0],index[1]+((day-1)*24),varobject[index].value,'RI','Gas',marginal_cost))
                    elif index[0] in instance.Coal:
                        marginal_cost = seg2*2
                        mwh_2.append((index[0],index[1]+((day-1)*24),varobject[index].value,'RI','Coal',marginal_cost))
                    elif index[0] in instance.Oil:
                        marginal_cost = seg2*20
                        mwh_2.append((index[0],index[1]+((day-1)*24),varobject[index].value,'RI','Oil',marginal_cost))
                    elif index[0] in instance.PSH:
                        marginal_cost = 10
                        mwh_2.append((index[0],index[1]+((day-1)*24),varobject[index].value,'RI','PSH',marginal_cost))
                    elif index[0] in instance.Slack:
                        marginal_cost = 700
                        mwh_2.append((index[0],index[1]+((day-1)*24),varobject[index].value,'RI','Slack',marginal_cost))  
        
     
                elif index[0] in instance.Zone6Generators:
                    
                    gas_price = instance.GasPrice['SEMA'].value
                    
                    if index[0] in instance.Gas:
                        marginal_cost = seg2*gas_price
                        mwh_2.append((index[0],index[1]+((day-1)*24),varobject[index].value,'SEMA','Gas',marginal_cost))
                    elif index[0] in instance.Coal:
                        marginal_cost = seg2*2
                        mwh_2.append((index[0],index[1]+((day-1)*24),varobject[index].value,'SEMA','Coal',marginal_cost))
                    elif index[0] in instance.Oil:
                        marginal_cost = seg2*20
                        mwh_2.append((index[0],index[1]+((day-1)*24),varobject[index].value,'SEMA','Oil',marginal_cost))
                    elif index[0] in instance.PSH:
                        marginal_cost = 10
                        mwh_2.append((index[0],index[1]+((day-1)*24),varobject[index].value,'SEMA','PSH',marginal_cost))
                    elif index[0] in instance.Slack:
                        marginal_cost = 700
                        mwh_2.append((index[0],index[1]+((day-1)*24),varobject[index].value,'SEMA','Slack',marginal_cost))  
        
      
                elif index[0] in instance.Zone7Generators:
                    
                    gas_price = instance.GasPrice['VT'].value
                    
                    if index[0] in instance.Gas:
                        marginal_cost = seg2*gas_price
                        mwh_2.append((index[0],index[1]+((day-1)*24),varobject[index].value,'VT','Gas',marginal_cost))
                    elif index[0] in instance.Coal:
                        marginal_cost = seg2*2
                        mwh_2.append((index[0],index[1]+((day-1)*24),varobject[index].value,'VT','Coal',marginal_cost))
                    elif index[0] in instance.Oil:
                        marginal_cost = seg2*20
                        mwh_2.append((index[0],index[1]+((day-1)*24),varobject[index].value,'VT','Oil',marginal_cost))
                    elif index[0] in instance.PSH:
                        marginal_cost = 10
                        mwh_2.append((index[0],index[1]+((day-1)*24),varobject[index].value,'VT','PSH',marginal_cost))
                    elif index[0] in instance.Slack:
                        marginal_cost = 700
                        mwh_2.append((index[0],index[1]+((day-1)*24),varobject[index].value,'VT','Slack',marginal_cost))  
        
                elif index[0] in instance.Zone8Generators:
                    
                    gas_price = instance.GasPrice['WCMA'].value
                    
                    if index[0] in instance.Gas:
                        marginal_cost = seg2*gas_price
                        mwh_2.append((index[0],index[1]+((day-1)*24),varobject[index].value,'WCMA','Gas',marginal_cost))
                    elif index[0] in instance.Coal:
                        marginal_cost = seg2*2
                        mwh_2.append((index[0],index[1]+((day-1)*24),varobject[index].value,'WCMA','Coal',marginal_cost))
                    elif index[0] in instance.Oil:
                        marginal_cost = seg2*20
                        mwh_2.append((index[0],index[1]+((day-1)*24),varobject[index].value,'WCMA','Oil',marginal_cost))
                    elif index[0] in instance.PSH:
                        marginal_cost = 10
                        mwh_2.append((index[0],index[1]+((day-1)*24),varobject[index].value,'WCMA','PSH',marginal_cost))
                    elif index[0] in instance.Slack:
                        marginal_cost = 700
                        mwh_2.append((index[0],index[1]+((day-1)*24),varobject[index].value,'WCMA','Slack',marginal_cost))  
        
              
            if a=='mwh_3':
               
             for index in varobject:
                 
               name = index[0]
               g = df_generators[df_generators['name']==name]
               seg3 = g['seg3'].values
               seg3 = seg3[0]
                 
               if int(index[1]>0 and index[1]<25):
                if index[0] in instance.Zone1Generators:
                    
                    gas_price = instance.GasPrice['CT'].value
                    
                    if index[0] in instance.Gas:
                        marginal_cost = seg3*gas_price
                        mwh_3.append((index[0],index[1]+((day-1)*24),varobject[index].value,'CT','Gas',marginal_cost))
                    elif index[0] in instance.Coal:
                        marginal_cost = seg3*2
                        mwh_3.append((index[0],index[1]+((day-1)*24),varobject[index].value,'CT','Coal',marginal_cost))
                    elif index[0] in instance.Oil:
                        marginal_cost = seg3*20
                        mwh_3.append((index[0],index[1]+((day-1)*24),varobject[index].value,'CT','Oil',marginal_cost))
                    elif index[0] in instance.PSH:
                        marginal_cost = 10
                        mwh_3.append((index[0],index[1]+((day-1)*24),varobject[index].value,'CT','PSH',marginal_cost))
                    elif index[0] in instance.Slack:
                        marginal_cost = 700
                        mwh_3.append((index[0],index[1]+((day-1)*24),varobject[index].value,'CT','Slack',marginal_cost))               
                    elif index[0] in instance.Hydro:
                        marginal_cost = 0
                        mwh_3.append((index[0],index[1]+((day-1)*24),varobject[index].value,'CT','Hydro',marginal_cost))
                        
                        
                elif index[0] in instance.Zone2Generators:
                    
                    gas_price = instance.GasPrice['ME'].value
                    
                    if index[0] in instance.Gas:
                        marginal_cost = seg3*gas_price
                        mwh_3.append((index[0],index[1]+((day-1)*24),varobject[index].value,'ME','Gas',marginal_cost))
                    elif index[0] in instance.Coal:
                        marginal_cost = seg3*2
                        mwh_3.append((index[0],index[1]+((day-1)*24),varobject[index].value,'ME','Coal',marginal_cost))
                    elif index[0] in instance.Oil:
                        marginal_cost = seg3*20
                        mwh_3.append((index[0],index[1]+((day-1)*24),varobject[index].value,'ME','Oil',marginal_cost))
                    elif index[0] in instance.PSH:
                        marginal_cost = 10
                        mwh_3.append((index[0],index[1]+((day-1)*24),varobject[index].value,'ME','PSH',marginal_cost))
                    elif index[0] in instance.Slack:
                        marginal_cost = 700
                        mwh_3.append((index[0],index[1]+((day-1)*24),varobject[index].value,'ME','Slack',marginal_cost))  
        
                elif index[0] in instance.Zone3Generators:
                    
                    gas_price = instance.GasPrice['NH'].value
                    
                    if index[0] in instance.Gas:
                        marginal_cost = seg3*gas_price
                        mwh_3.append((index[0],index[1]+((day-1)*24),varobject[index].value,'NH','Gas',marginal_cost))
                    elif index[0] in instance.Coal:
                        marginal_cost = seg3*2
                        mwh_3.append((index[0],index[1]+((day-1)*24),varobject[index].value,'NH','Coal',marginal_cost))
                    elif index[0] in instance.Oil:
                        marginal_cost = seg3*20
                        mwh_3.append((index[0],index[1]+((day-1)*24),varobject[index].value,'NH','Oil',marginal_cost))
                    elif index[0] in instance.PSH:
                        marginal_cost = 10
                        mwh_3.append((index[0],index[1]+((day-1)*24),varobject[index].value,'NH','PSH',marginal_cost))
                    elif index[0] in instance.Slack:
                        marginal_cost = 700
                        mwh_3.append((index[0],index[1]+((day-1)*24),varobject[index].value,'NH','Slack',marginal_cost))  
                    elif index[0] in instance.Hydro:
                        marginal_cost = 0
                        mwh_3.append((index[0],index[1]+((day-1)*24),varobject[index].value,'NH','Hydro',marginal_cost))  
            
                elif index[0] in instance.Zone4Generators:
                    
                    gas_price = instance.GasPrice['NEMA'].value
                    
                    if index[0] in instance.Gas:
                        marginal_cost = seg3*gas_price
                        mwh_3.append((index[0],index[1]+((day-1)*24),varobject[index].value,'NEMA','Gas',marginal_cost))
                    elif index[0] in instance.Coal:
                        marginal_cost = seg3*2
                        mwh_3.append((index[0],index[1]+((day-1)*24),varobject[index].value,'NEMA','Coal',marginal_cost))
                    elif index[0] in instance.Oil:
                        marginal_cost = seg3*20
                        mwh_3.append((index[0],index[1]+((day-1)*24),varobject[index].value,'NEMA','Oil',marginal_cost))
                    elif index[0] in instance.PSH:
                        marginal_cost = 10
                        mwh_3.append((index[0],index[1]+((day-1)*24),varobject[index].value,'NEMA','PSH',marginal_cost))
                    elif index[0] in instance.Slack:
                        marginal_cost = 700
                        mwh_3.append((index[0],index[1]+((day-1)*24),varobject[index].value,'NEMA','Slack',marginal_cost))  
        
                elif index[0] in instance.Zone5Generators:
                    
                    gas_price = instance.GasPrice['RI'].value
                    
                    if index[0] in instance.Gas:
                        marginal_cost = seg3*gas_price
                        mwh_3.append((index[0],index[1]+((day-1)*24),varobject[index].value,'RI','Gas',marginal_cost))
                    elif index[0] in instance.Coal:
                        marginal_cost = seg3*2
                        mwh_3.append((index[0],index[1]+((day-1)*24),varobject[index].value,'RI','Coal',marginal_cost))
                    elif index[0] in instance.Oil:
                        marginal_cost = seg3*20
                        mwh_3.append((index[0],index[1]+((day-1)*24),varobject[index].value,'RI','Oil',marginal_cost))
                    elif index[0] in instance.PSH:
                        marginal_cost = 10
                        mwh_3.append((index[0],index[1]+((day-1)*24),varobject[index].value,'RI','PSH',marginal_cost))
                    elif index[0] in instance.Slack:
                        marginal_cost = 700
                        mwh_3.append((index[0],index[1]+((day-1)*24),varobject[index].value,'RI','Slack',marginal_cost))  
  

                elif index[0] in instance.Zone6Generators:
                    
                    gas_price = instance.GasPrice['SEMA'].value
                    
                    if index[0] in instance.Gas:
                        marginal_cost = seg3*gas_price
                        mwh_3.append((index[0],index[1]+((day-1)*24),varobject[index].value,'SEMA','Gas',marginal_cost))
                    elif index[0] in instance.Coal:
                        marginal_cost = seg3*2
                        mwh_3.append((index[0],index[1]+((day-1)*24),varobject[index].value,'SEMA','Coal',marginal_cost))
                    elif index[0] in instance.Oil:
                        marginal_cost = seg3*20
                        mwh_3.append((index[0],index[1]+((day-1)*24),varobject[index].value,'SEMA','Oil',marginal_cost))
                    elif index[0] in instance.PSH:
                        marginal_cost = 10
                        mwh_3.append((index[0],index[1]+((day-1)*24),varobject[index].value,'SEMA','PSH',marginal_cost))
                    elif index[0] in instance.Slack:
                        marginal_cost = 700
                        mwh_3.append((index[0],index[1]+((day-1)*24),varobject[index].value,'SEMA','Slack',marginal_cost))  

                elif index[0] in instance.Zone7Generators:
                    
                    gas_price = instance.GasPrice['VT'].value
                    
                    if index[0] in instance.Gas:
                        marginal_cost = seg3*gas_price
                        mwh_3.append((index[0],index[1]+((day-1)*24),varobject[index].value,'VT','Gas',marginal_cost))
                    elif index[0] in instance.Coal:
                        marginal_cost = seg3*2
                        mwh_3.append((index[0],index[1]+((day-1)*24),varobject[index].value,'VT','Coal',marginal_cost))
                    elif index[0] in instance.Oil:
                        marginal_cost = seg3*20
                        mwh_3.append((index[0],index[1]+((day-1)*24),varobject[index].value,'VT','Oil',marginal_cost))
                    elif index[0] in instance.PSH:
                        marginal_cost = 10
                        mwh_3.append((index[0],index[1]+((day-1)*24),varobject[index].value,'VT','PSH',marginal_cost))
                    elif index[0] in instance.Slack:
                        marginal_cost = 700
                        mwh_3.append((index[0],index[1]+((day-1)*24),varobject[index].value,'VT','Slack',marginal_cost))  

                elif index[0] in instance.Zone8Generators:
                    
                    gas_price = instance.GasPrice['WCMA'].value
                    
                    if index[0] in instance.Gas:
                        marginal_cost = seg3*gas_price
                        mwh_3.append((index[0],index[1]+((day-1)*24),varobject[index].value,'WCMA','Gas',marginal_cost))
                    elif index[0] in instance.Coal:
                        marginal_cost = seg3*2
                        mwh_3.append((index[0],index[1]+((day-1)*24),varobject[index].value,'WCMA','Coal',marginal_cost))
                    elif index[0] in instance.Oil:
                        marginal_cost = seg3*20
                        mwh_3.append((index[0],index[1]+((day-1)*24),varobject[index].value,'WCMA','Oil',marginal_cost))
                    elif index[0] in instance.PSH:
                        marginal_cost = 10
                        mwh_3.append((index[0],index[1]+((day-1)*24),varobject[index].value,'WCMA','PSH',marginal_cost))
                    elif index[0] in instance.Slack:
                        marginal_cost = 700
                        mwh_3.append((index[0],index[1]+((day-1)*24),varobject[index].value,'WCMA','Slack',marginal_cost))  
          
              
            if a=='on':
                
             for index in varobject:
               if int(index[1]>0 and index[1]<25):
                if index[0] in instance.Zone1Generators:
                 on.append((index[0],index[1]+((day-1)*24),varobject[index].value,'CT'))
                elif index[0] in instance.Zone2Generators:
                 on.append((index[0],index[1]+((day-1)*24),varobject[index].value,'ME'))
                elif index[0] in instance.Zone3Generators:
                 on.append((index[0],index[1]+((day-1)*24),varobject[index].value,'NH'))
                elif index[0] in instance.Zone4Generators:
                 on.append((index[0],index[1]+((day-1)*24),varobject[index].value,'NEMA'))
                elif index[0] in instance.Zone5Generators:
                 on.append((index[0],index[1]+((day-1)*24),varobject[index].value,'RI'))
                elif index[0] in instance.Zone6Generators:
                 on.append((index[0],index[1]+((day-1)*24),varobject[index].value,'SEMA'))
                elif index[0] in instance.Zone7Generators:
                 on.append((index[0],index[1]+((day-1)*24),varobject[index].value,'VT'))
                elif index[0] in instance.Zone8Generators:
                 on.append((index[0],index[1]+((day-1)*24),varobject[index].value,'WCMA'))
          
             
            if a=='switch':
            
             for index in varobject:
               if int(index[1]>0 and index[1]<25):
                if index[0] in instance.Zone1Generators:
                 switch.append((index[0],index[1]+((day-1)*24),varobject[index].value,'CT'))
                elif index[0] in instance.Zone2Generators:
                 switch.append((index[0],index[1]+((day-1)*24),varobject[index].value,'ME'))
                elif index[0] in instance.Zone3Generators:
                 switch.append((index[0],index[1]+((day-1)*24),varobject[index].value,'NH'))
                elif index[0] in instance.Zone4Generators:
                 switch.append((index[0],index[1]+((day-1)*24),varobject[index].value,'NEMA'))   
                elif index[0] in instance.Zone5Generators:
                 switch.append((index[0],index[1]+((day-1)*24),varobject[index].value,'RI'))   
                elif index[0] in instance.Zone6Generators:
                 switch.append((index[0],index[1]+((day-1)*24),varobject[index].value,'SEMA'))   
                elif index[0] in instance.Zone7Generators:
                 switch.append((index[0],index[1]+((day-1)*24),varobject[index].value,'VT'))   
                elif index[0] in instance.Zone8Generators:
                 switch.append((index[0],index[1]+((day-1)*24),varobject[index].value,'WCMA'))   
        
             
            if a=='srsv':
            
             for index in varobject:
               if int(index[1]>0 and index[1]<25):
                if index[0] in instance.Zone1Generators:
                 srsv.append((index[0],index[1]+((day-1)*24),varobject[index].value,'CT'))
                elif index[0] in instance.Zone2Generators:
                 srsv.append((index[0],index[1]+((day-1)*24),varobject[index].value,'ME'))
                elif index[0] in instance.Zone3Generators:
                 srsv.append((index[0],index[1]+((day-1)*24),varobject[index].value,'NH'))
                elif index[0] in instance.Zone4Generators:
                 srsv.append((index[0],index[1]+((day-1)*24),varobject[index].value,'NEMA'))  
                elif index[0] in instance.Zone5Generators:
                 srsv.append((index[0],index[1]+((day-1)*24),varobject[index].value,'RI'))  
                elif index[0] in instance.Zone6Generators:
                 srsv.append((index[0],index[1]+((day-1)*24),varobject[index].value,'SEMA'))  
                elif index[0] in instance.Zone7Generators:
                 srsv.append((index[0],index[1]+((day-1)*24),varobject[index].value,'VT'))  
                elif index[0] in instance.Zone8Generators:
                 srsv.append((index[0],index[1]+((day-1)*24),varobject[index].value,'WCMA'))  
             
            if a=='nrsv':
           
             for index in varobject:
               if int(index[1]>0 and index[1]<25):
                if index[0] in instance.Zone1Generators:
                 nrsv.append((index[0],index[1]+((day-1)*24),varobject[index].value,'CT'))
                elif index[0] in instance.Zone2Generators:
                 nrsv.append((index[0],index[1]+((day-1)*24),varobject[index].value,'ME'))
                elif index[0] in instance.Zone3Generators:
                 nrsv.append((index[0],index[1]+((day-1)*24),varobject[index].value,'NH'))
                elif index[0] in instance.Zone4Generators:
                 nrsv.append((index[0],index[1]+((day-1)*24),varobject[index].value,'NEMA'))
                elif index[0] in instance.Zone5Generators:
                 nrsv.append((index[0],index[1]+((day-1)*24),varobject[index].value,'RI'))
                elif index[0] in instance.Zone6Generators:
                 nrsv.append((index[0],index[1]+((day-1)*24),varobject[index].value,'SEMA'))
                elif index[0] in instance.Zone7Generators:
                 nrsv.append((index[0],index[1]+((day-1)*24),varobject[index].value,'VT'))
                elif index[0] in instance.Zone8Generators:
                 nrsv.append((index[0],index[1]+((day-1)*24),varobject[index].value,'WCMA'))
             
            if a=='solar':
               
             for index in varobject:
               if int(index[1]>0 and index[1]<25):
                solar.append((index[0],index[1]+((day-1)*24),varobject[index].value))   
             
              
            if a=='wind':
               
             for index in varobject:
               if int(index[1]>0 and index[1]<25):
                wind.append((index[0],index[1]+((day-1)*24),varobject[index].value))  
                 
            if a=='flow':
               
             for index in varobject:
               if int(index[2]>0 and index[2]<25):
                flow.append((index[0],index[1],index[2]+((day-1)*24),varobject[index].value))   
             
        
            for j in instance.Generators:
                if instance.on[j,H] == 1:
                    instance.on[j,0] = 1
                else: 
                    instance.on[j,0] = 0
                instance.on[j,0].fixed = True
                           
                if instance.mwh_1[j,H].value <=0 and instance.mwh_1[j,H].value>= -0.0001:
                    newval_1=0
                else:
                    newval_1=instance.mwh_1[j,H].value
                instance.mwh_1[j,0] = newval_1
                instance.mwh_1[j,0].fixed = True
                              
                if instance.mwh_2[j,H].value <=0 and instance.mwh_2[j,H].value>= -0.0001:
                    newval=0
                else:
                    newval=instance.mwh_2[j,H].value
                                         
                if instance.mwh_3[j,H].value <=0 and instance.mwh_3[j,H].value>= -0.0001:
                    newval2=0
                else:
                    newval2=instance.mwh_3[j,H].value
                                          
                                          
                instance.mwh_2[j,0] = newval
                instance.mwh_2[j,0].fixed = True
                instance.mwh_3[j,0] = newval2
                instance.mwh_3[j,0].fixed = True 
                if instance.switch[j,H] == 1:
                    instance.switch[j,0] = 1
                else:
                    instance.switch[j,0] = 0
                instance.switch[j,0].fixed = True
              
                if instance.srsv[j,H].value <=0 and instance.srsv[j,H].value>= -0.0001:
                    newval_srsv=0
                else:
                    newval_srsv=instance.srsv[j,H].value
                instance.srsv[j,0] = newval_srsv 
                instance.srsv[j,0].fixed = True        
        
                if instance.nrsv[j,H].value <=0 and instance.nrsv[j,H].value>= -0.0001:
                    newval_nrsv=0
                else:
                    newval_nrsv=instance.nrsv[j,H].value
                instance.nrsv[j,0] = newval_nrsv 
                instance.nrsv[j,0].fixed = True      
                
        print(day)
                     
    mwh_1_pd=pd.DataFrame(mwh_1,columns=('Generator','Time','Value','Zones','Type','$/MWh'))
    mwh_2_pd=pd.DataFrame(mwh_2,columns=('Generator','Time','Value','Zones','Type','$/MWh'))
    mwh_3_pd=pd.DataFrame(mwh_3,columns=('Generator','Time','Value','Zones','Type','$/MWh'))
    on_pd=pd.DataFrame(on,columns=('Generator','Time','Value','Zones'))
    switch_pd=pd.DataFrame(switch,columns=('Generator','Time','Value','Zones'))
    srsv_pd=pd.DataFrame(srsv,columns=('Generator','Time','Value','Zones'))
    nrsv_pd=pd.DataFrame(nrsv,columns=('Generator','Time','Value','Zones'))
    solar_pd=pd.DataFrame(solar,columns=('Zone','Time','Value'))
    wind_pd=pd.DataFrame(wind,columns=('Zone','Time','Value'))
    flow_pd=pd.DataFrame(flow,columns=('Source','Sink','Time','Value'))
    
    flow_pd.to_csv('NEISO/flow.csv')
    mwh_1_pd.to_csv('NEISO/mwh_1.csv')
    mwh_2_pd.to_csv('NEISO/mwh_2.csv')
    mwh_3_pd.to_csv('NEISO/mwh_3.csv')
    on_pd.to_csv('NEISO/on.csv')
    switch_pd.to_csv('NEISO/switch.csv')
    srsv_pd.to_csv('NEISO/srsv.csv')
    nrsv_pd.to_csv('NEISO/nrsv.csv')
    solar_pd.to_csv('NEISO/solar_out.csv')
    wind_pd.to_csv('NEISO/wind_out.csv')
    
    return None
