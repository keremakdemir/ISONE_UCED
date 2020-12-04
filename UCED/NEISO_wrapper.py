# -*- coding: utf-8 -*-
"""
Created on Tue Jun 20 22:14:07 2017

@author: YSu
"""

from pyomo.opt import SolverFactory
from NEISO_dispatch import model as m1
from NEISO_dispatchLP import model as m2
from pyomo.core import Var
from pyomo.core import Constraint
from pyomo.core import Param
from operator import itemgetter
import pandas as pd
import numpy as np
from datetime import datetime
import pyomo.environ as pyo

Solvername = 'gurobi'
Timelimit = 1800 # for the simulation of one day in seconds
# Threadlimit = 8 # maximum number of threads to use

def sim(days):

    instance = m1.create_instance('data.dat')
    instance2 = m2.create_instance('data.dat')

    instance2.dual = pyo.Suffix(direction=pyo.Suffix.IMPORT)
    opt = SolverFactory(Solvername)
    
    if Solvername == 'cplex':
        opt.options['timelimit'] = Timelimit
    elif Solvername == 'gurobi':           
        opt.options['TimeLimit'] = Timelimit
        
    # opt.options['threads'] = Threadlimit

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
    onshore_wind = []
    offshore_wind=[]
    must_run=[]
    flow=[]
    Generator=[]
    Duals=[]
    System_cost = []
    df_generators = pd.read_excel('generators.xlsx',header=0)

    #max here can be (1,365)
    for day in range(1,days):
        
        if day == days-1:
            horizon_end = 49
        else:
            horizon_end = 25

        #load time series data
        for z in instance.zones:

            instance.GasPrice[z] = 3.6

            for i in K:
                
                instance.HorizonDemand[z,i] = instance.SimDemand[z,(day-1)*24+i]
                instance.HorizonOffshoreWind[z,i] = instance.SimOffshoreWind[z,(day-1)*24+i]
                instance.HorizonSolar[z,i] = instance.SimSolar[z,(day-1)*24+i]
                instance.HorizonOnshoreWind[z,i] = instance.SimOnshoreWind[z,(day-1)*24+i]
                instance.HorizonMustRun[z,i] = instance.SimMustRun[z,(day-1)*24+i]
        
        for d in range(1,D+1):
            
            instance.HorizonNY_imports_CT[d] = instance.SimNY_imports_CT[day-1+d]
            instance.HorizonNY_imports_WCMA[d] = instance.SimNY_imports_CT[day-1+d]
            instance.HorizonNY_imports_VT[d] = instance.SimNY_imports_CT[day-1+d]
            instance.HorizonHQ_imports_VT[d] = instance.SimNY_imports_CT[day-1+d]
            instance.HorizonNB_imports_ME[d] = instance.SimNY_imports_CT[day-1+d]
            
            instance.HorizonME_hydro[d] = instance.SimME_hydro[day-1+d]
            instance.HorizonVT_hydro[d] = instance.SimVT_hydro[day-1+d]
            instance.HorizonRI_hydro[d] = instance.SimRI_hydro[day-1+d]
            instance.HorizonNH_hydro[d] = instance.SimNH_hydro[day-1+d]
            instance.HorizonCT_hydro[d] = instance.SimCT_hydro[day-1+d]
            instance.HorizonWCMA_hydro[d] = instance.SimWCMA_hydro[day-1+d]
            instance.HorizonNEMA_hydro[d] = instance.SimNEMA_hydro[day-1+d]

        for i in K:
            
            instance.HorizonReserves[i] = instance.SimReserves[(day-1)*24+i]
            
            instance.HorizonCT_exports_NY[i] =  instance.SimCT_exports_NY[(day-1)*24+i]
            instance.HorizonWCMA_exports_NY[i] = instance.SimWCMA_exports_NY[(day-1)*24+i]
            instance.HorizonVT_exports_NY[i] = instance.SimVT_exports_NY[(day-1)*24+i]
            instance.HorizonVT_exports_HQ[i] = instance.SimVT_exports_HQ[(day-1)*24+i]
            instance.HorizonME_exports_NB[i] = instance.SimME_exports_NB[(day-1)*24+i]

        NEISO_result = opt.solve(instance, tee=True, symbolic_solver_labels=True, load_solutions=False)
        instance.solutions.load_from(NEISO_result)

        # record objective function value
        coal = 0
        gas1_1 = 0
        gas2_1 = 0
        gas3_1 = 0
        gas1_2 = 0
        gas2_2 = 0
        gas3_2 = 0
        gas1_3 = 0
        gas2_3 = 0
        gas3_3 = 0
        gas1_4 = 0
        gas2_4 = 0
        gas3_4 = 0
        gas1_5 = 0
        gas2_5 = 0
        gas3_5 = 0
        gas1_6 = 0
        gas2_6 = 0
        gas3_6 = 0
        gas1_7 = 0
        gas2_7 = 0
        gas3_7 = 0
        gas1_8 = 0
        gas2_8 = 0
        gas3_8 = 0
        oil = 0
        NY_Imports_CT_all = 0
        NY_Imports_WCMA_all = 0
        NY_Imports_VT_all = 0
        HQ_Imports_VT_all = 0
        NB_Imports_ME_all = 0
        slack = 0
        fix_cost = 0
        st = 0
        exchn = 0
        hydro = 0
        onshore_all = 0
        offshore_all = 0
        solar_all = 0
        mustrun_all = 0
        
        for i in range(1,horizon_end):
            for j in instance.Coal:
                coal = coal + instance.mwh_1[j,i].value*(instance.seg1[j]*2.2 + instance.var_om[j]) + instance.mwh_2[j,i].value*(instance.seg2[j]*2.2 + instance.var_om[j]) + instance.mwh_3[j,i].value*(instance.seg3[j]*2.2 + instance.var_om[j])  
            for j in instance.Zone1Gas:
                gas1_1 = gas1_1 + instance.mwh_1[j,i].value*(instance.seg1[j]*instance.GasPrice['CT'].value + instance.var_om[j]) 
                gas2_1 = gas2_1 + instance.mwh_2[j,i].value*(instance.seg2[j]*instance.GasPrice['CT'].value + instance.var_om[j]) 
                gas3_1 = gas3_1 + instance.mwh_3[j,i].value*(instance.seg3[j]*instance.GasPrice['CT'].value + instance.var_om[j]) 
            for j in instance.Zone2Gas:
                gas1_2 = gas1_2 + instance.mwh_1[j,i].value*(instance.seg1[j]*instance.GasPrice['ME'].value + instance.var_om[j]) 
                gas2_2 = gas2_2 + instance.mwh_2[j,i].value*(instance.seg2[j]*instance.GasPrice['ME'].value + instance.var_om[j]) 
                gas3_2 = gas3_2 + instance.mwh_3[j,i].value*(instance.seg3[j]*instance.GasPrice['ME'].value + instance.var_om[j]) 
            for j in instance.Zone3Gas:
                gas1_3 = gas1_3 + instance.mwh_1[j,i].value*(instance.seg1[j]*instance.GasPrice['NH'].value + instance.var_om[j]) 
                gas2_3 = gas2_3 + instance.mwh_2[j,i].value*(instance.seg2[j]*instance.GasPrice['NH'].value + instance.var_om[j]) 
                gas3_3 = gas3_3 + instance.mwh_3[j,i].value*(instance.seg3[j]*instance.GasPrice['NH'].value + instance.var_om[j]) 
            for j in instance.Zone4Gas:
                gas1_4 = gas1_4 + instance.mwh_1[j,i].value*(instance.seg1[j]*instance.GasPrice['NEMA'].value + instance.var_om[j]) 
                gas2_4 = gas2_4 + instance.mwh_2[j,i].value*(instance.seg2[j]*instance.GasPrice['NEMA'].value + instance.var_om[j]) 
                gas3_4 = gas3_4 + instance.mwh_3[j,i].value*(instance.seg3[j]*instance.GasPrice['NEMA'].value + instance.var_om[j])    
            for j in instance.Zone5Gas:
                gas1_5 = gas1_5 + instance.mwh_1[j,i].value*(instance.seg1[j]*instance.GasPrice['RI'].value + instance.var_om[j]) 
                gas2_5 = gas2_5 + instance.mwh_2[j,i].value*(instance.seg2[j]*instance.GasPrice['RI'].value + instance.var_om[j]) 
                gas3_5 = gas3_5 + instance.mwh_3[j,i].value*(instance.seg3[j]*instance.GasPrice['RI'].value + instance.var_om[j])   
            for j in instance.Zone6Gas:
                gas1_6 = gas1_6 + instance.mwh_1[j,i].value*(instance.seg1[j]*instance.GasPrice['SEMA'].value + instance.var_om[j]) 
                gas2_6 = gas2_6 + instance.mwh_2[j,i].value*(instance.seg2[j]*instance.GasPrice['SEMA'].value + instance.var_om[j]) 
                gas3_6 = gas3_6 + instance.mwh_3[j,i].value*(instance.seg3[j]*instance.GasPrice['SEMA'].value + instance.var_om[j])                
            for j in instance.Zone7Gas:
                gas1_7 = gas1_7 + instance.mwh_1[j,i].value*(instance.seg1[j]*instance.GasPrice['VT'].value + instance.var_om[j]) 
                gas2_7 = gas2_7 + instance.mwh_2[j,i].value*(instance.seg2[j]*instance.GasPrice['VT'].value + instance.var_om[j]) 
                gas3_7 = gas3_7 + instance.mwh_3[j,i].value*(instance.seg3[j]*instance.GasPrice['VT'].value + instance.var_om[j])   
            for j in instance.Zone8Gas:
                gas1_8 = gas1_8 + instance.mwh_1[j,i].value*(instance.seg1[j]*instance.GasPrice['WCMA'].value + instance.var_om[j]) 
                gas2_8 = gas2_8 + instance.mwh_2[j,i].value*(instance.seg2[j]*instance.GasPrice['WCMA'].value + instance.var_om[j]) 
                gas3_8 = gas3_8 + instance.mwh_3[j,i].value*(instance.seg3[j]*instance.GasPrice['WCMA'].value + instance.var_om[j])   
            for j in instance.Oil:
                oil = oil + instance.mwh_1[j,i].value*(instance.seg1[j]*8 + instance.var_om[j]) + instance.mwh_2[j,i].value*(instance.seg2[j]*8 + instance.var_om[j]) + instance.mwh_3[j,i].value*(instance.seg3[j]*8 + instance.var_om[j])
            for j in instance.NY_Imports_CT:
                NY_Imports_CT_all = NY_Imports_CT_all + instance.mwh_1[j,i].value*4.3 + instance.mwh_2[j,i].value*4.3 + instance.mwh_3[j,i].value*4.3
            for j in instance.NY_Imports_WCMA:
                NY_Imports_WCMA_all = NY_Imports_WCMA_all + instance.mwh_1[j,i].value*4.3 + instance.mwh_2[j,i].value*4.3 + instance.mwh_3[j,i].value*4.3
            for j in instance.NY_Imports_VT:
                NY_Imports_VT_all = NY_Imports_VT_all + instance.mwh_1[j,i].value*4.3 + instance.mwh_2[j,i].value*4.3 + instance.mwh_3[j,i].value*4.3
            for j in instance.HQ_Imports_VT:
                HQ_Imports_VT_all = HQ_Imports_VT_all + instance.mwh_1[j,i].value*4.3 + instance.mwh_2[j,i].value*4.3 + instance.mwh_3[j,i].value*4.3  
            for j in instance.NB_Imports_ME:
                NB_Imports_ME_all = NB_Imports_ME_all + instance.mwh_1[j,i].value*4.3 + instance.mwh_2[j,i].value*4.3 + instance.mwh_3[j,i].value*4.3
            for j in instance.Slack:
                slack = slack + instance.mwh_1[j,i].value*(instance.seg1[j]*2000 + instance.var_om[j]) + instance.mwh_2[j,i].value*(instance.seg2[j]*2000 + instance.var_om[j]) + instance.mwh_3[j,i].value*(instance.seg3[j]*2000 + instance.var_om[j])  
            for j in instance.Generators:
                fix_cost = fix_cost + instance.no_load[j]*instance.on[j,i].value
            for j in instance.Generators:
                st = st + instance.st_cost[j]*instance.switch[j,i].value
            for s in instance.sources:
                for k in instance.sinks:
                    exchn = exchn + instance.flow[s,k,i].value*instance.hurdle[s,k] 
            for j in instance.Hydro:
                hydro = hydro + instance.mwh_1[j,i].value*0.01 + instance.mwh_2[j,i].value*0.01 + instance.mwh_3[j,i].value*0.01
            for j in instance.zones:
                onshore_all = onshore_all + instance.onshorewind[j,i].value*0.01 + instance.onshorewind[j,i].value*0.01 + instance.onshorewind[j,i].value*0.01
            for j in instance.zones:
                offshore_all = offshore_all + instance.offshorewind[j,i].value*0.01 + instance.offshorewind[j,i].value*0.01 + instance.offshorewind[j,i].value*0.01
            for j in instance.zones:
                solar_all = solar_all + instance.solar[j,i].value*0.01 + instance.solar[j,i].value*0.01 + instance.solar[j,i].value*0.01 
            for j in instance.zones:
                mustrun_all = mustrun_all + instance.mustrun[j,i].value*0.01 + instance.mustrun[j,i].value*0.01 + instance.mustrun[j,i].value*0.01

            S = coal + gas1_1 + gas2_1 + gas3_1 + gas1_2 + gas2_2 + gas3_2 + gas1_3 + gas2_3 + gas3_3 + gas1_4 + gas2_4 + gas3_4 + gas1_5 + gas2_5 + gas3_5 + gas1_6 + gas2_6 + gas3_6 + gas1_7 + gas2_7 + gas3_7 + gas1_8 + gas2_8 + gas3_8 + oil + slack + fix_cost + st + exchn + NY_Imports_CT_all + NY_Imports_WCMA_all + NY_Imports_VT_all + HQ_Imports_VT_all + NB_Imports_ME_all + hydro + onshore_all + offshore_all + solar_all + mustrun_all
            System_cost.append(round(S, 3))


        for z in instance2.zones:

            instance2.GasPrice[z] = 3.6

            for i in K:
                
                instance2.HorizonDemand[z,i] = instance2.SimDemand[z,(day-1)*24+i]
                instance2.HorizonOffshoreWind[z,i] = instance2.SimOffshoreWind[z,(day-1)*24+i]
                instance2.HorizonSolar[z,i] = instance2.SimSolar[z,(day-1)*24+i]
                instance2.HorizonOnshoreWind[z,i] = instance2.SimOnshoreWind[z,(day-1)*24+i]
                instance2.HorizonMustRun[z,i] = instance2.SimMustRun[z,(day-1)*24+i]

        for d in range(1,D+1):
            
            instance2.HorizonNY_imports_CT[d] = instance2.SimNY_imports_CT[day-1+d]
            instance2.HorizonNY_imports_WCMA[d] = instance2.SimNY_imports_CT[day-1+d]
            instance2.HorizonNY_imports_VT[d] = instance2.SimNY_imports_CT[day-1+d]
            instance2.HorizonHQ_imports_VT[d] = instance2.SimNY_imports_CT[day-1+d]
            instance2.HorizonNB_imports_ME[d] = instance2.SimNY_imports_CT[day-1+d]
            
            instance2.HorizonME_hydro[d] = instance2.SimME_hydro[day-1+d]
            instance2.HorizonVT_hydro[d] = instance2.SimVT_hydro[day-1+d]
            instance2.HorizonRI_hydro[d] = instance2.SimRI_hydro[day-1+d]
            instance2.HorizonNH_hydro[d] = instance2.SimNH_hydro[day-1+d]
            instance2.HorizonCT_hydro[d] = instance2.SimCT_hydro[day-1+d]
            instance2.HorizonWCMA_hydro[d] = instance2.SimWCMA_hydro[day-1+d]
            instance2.HorizonNEMA_hydro[d] = instance2.SimNEMA_hydro[day-1+d]
            
        for i in K:
            
            instance2.HorizonReserves[i] = instance2.SimReserves[(day-1)*24+i]

            instance2.HorizonCT_exports_NY[i] =  instance2.SimCT_exports_NY[(day-1)*24+i]
            instance2.HorizonWCMA_exports_NY[i] = instance2.SimWCMA_exports_NY[(day-1)*24+i]
            instance2.HorizonVT_exports_NY[i] = instance2.SimVT_exports_NY[(day-1)*24+i]
            instance2.HorizonVT_exports_HQ[i] = instance2.SimVT_exports_HQ[(day-1)*24+i]
            instance2.HorizonME_exports_NB[i] = instance2.SimME_exports_NB[(day-1)*24+i]

        for j in instance.Generators:
            for t in K:
                if instance.on[j,t] == 1:
                    instance2.on[j,t] = 1
                    instance2.on[j,t].fixed = True
                else:
                    instance.on[j,t] = 0
                    instance2.on[j,t] = 0
                    instance2.on[j,t].fixed = True

                if instance.switch[j,t] == 1:
                    instance2.switch[j,t] = 1
                    instance2.switch[j,t].fixed = True
                else:
                    instance2.switch[j,t] = 0
                    instance2.switch[j,t] = 0
                    instance2.switch[j,t].fixed = True
                    
        results = opt.solve(instance2, tee=True, symbolic_solver_labels=True, load_solutions=False)
        instance2.solutions.load_from(results)

        print ("Duals")

        for c in instance2.component_objects(Constraint, active=True):
    #        print ("   Constraint",c)
            cobject = getattr(instance2, str(c))
            if str(c) in ['Bal1Constraint','Bal2Constraint','Bal3Constraint','Bal4Constraint', 'Bal5Constraint','Bal6Constraint','Bal7Constraint','Bal8Constraint']:
                for index in cobject:
                     if index>0 and index<horizon_end:
    #                print ("   Constraint",c)
                         try:
                             Duals.append((str(c),index+((day-1)*24), round(instance2.dual[cobject[index]], 3)))
                         except KeyError:
                             Duals.append((str(c),index+((day-1)*24),-999))
    #            print ("      ", index, instance2.dual[cobject[index]])
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
                    zone = g['zone'] 
                    
                    if int(index[1]>0 and index[1]<horizon_end):
                        
                        gas_price = 3.6
                        
                        if index[0] in instance.Gas:
                            marginal_cost = seg1*gas_price
                            mwh_1.append((index[0],index[1]+((day-1)*24),round(varobject[index].value, 3),zone.values[0],'Gas',round(marginal_cost, 3)))
                        elif index[0] in instance.Coal:
                            marginal_cost = seg1*2.2
                            mwh_1.append((index[0],index[1]+((day-1)*24),round(varobject[index].value, 3),zone.values[0],'Coal',round(marginal_cost, 3)))
                        elif index[0] in instance.Oil:
                            marginal_cost = seg1*8
                            mwh_1.append((index[0],index[1]+((day-1)*24),round(varobject[index].value, 3),zone.values[0],'Oil',round(marginal_cost, 3)))
                        elif index[0] in instance.Slack:
                            marginal_cost = 700
                            mwh_1.append((index[0],index[1]+((day-1)*24),round(varobject[index].value, 3),zone.values[0],'Slack',round(marginal_cost, 3)))
                        elif index[0] in instance.Hydro:
                            marginal_cost = 0
                            mwh_1.append((index[0],index[1]+((day-1)*24),round(varobject[index].value, 3),zone.values[0],'Hydro',round(marginal_cost, 3)))
                        elif index[0] in instance.NY_Imports_CT:
                            marginal_cost = 4.3
                            mwh_1.append((index[0],index[1]+((day-1)*24),round(varobject[index].value, 3),zone.values[0],'Imports',round(marginal_cost, 3)))
                        elif index[0] in instance.NY_Imports_WCMA:
                            marginal_cost = 4.3
                            mwh_1.append((index[0],index[1]+((day-1)*24),round(varobject[index].value, 3),zone.values[0],'Imports',round(marginal_cost, 3)))                
                        elif index[0] in instance.NY_Imports_VT:
                            marginal_cost = 4.3
                            mwh_1.append((index[0],index[1]+((day-1)*24),round(varobject[index].value, 3),zone.values[0],'Imports',round(marginal_cost, 3))) 
                        elif index[0] in instance.HQ_Imports_VT:
                            marginal_cost = 4.3
                            mwh_1.append((index[0],index[1]+((day-1)*24),round(varobject[index].value, 3),zone.values[0],'Imports',round(marginal_cost, 3))) 
                        elif index[0] in instance.NB_Imports_ME:
                            marginal_cost = 4.3
                            mwh_1.append((index[0],index[1]+((day-1)*24),round(varobject[index].value, 3),zone.values[0],'Imports',round(marginal_cost, 3)))


            if a=='mwh_2':

                for index in varobject:

                    name = index[0]
                    g = df_generators[df_generators['name']==name]
                    seg2 = g['seg2'].values
                    seg2 = seg2[0]
                    zone = g['zone'] 
                    
                    if int(index[1]>0 and index[1]<horizon_end):

                        gas_price = 3.6
                        
                        if index[0] in instance.Gas:
                            marginal_cost = seg2*gas_price
                            mwh_2.append((index[0],index[1]+((day-1)*24),round(varobject[index].value, 3),zone.values[0],'Gas',round(marginal_cost, 3)))
                        elif index[0] in instance.Coal:
                            marginal_cost = seg2*2.2
                            mwh_2.append((index[0],index[1]+((day-1)*24),round(varobject[index].value, 3),zone.values[0],'Coal',round(marginal_cost, 3)))
                        elif index[0] in instance.Oil:
                            marginal_cost = seg2*8
                            mwh_2.append((index[0],index[1]+((day-1)*24),round(varobject[index].value, 3),zone.values[0],'Oil',round(marginal_cost, 3)))
                        elif index[0] in instance.Slack:
                            marginal_cost = 700
                            mwh_2.append((index[0],index[1]+((day-1)*24),round(varobject[index].value, 3),zone.values[0],'Slack',round(marginal_cost, 3)))
                        elif index[0] in instance.Hydro:
                            marginal_cost = 0
                            mwh_2.append((index[0],index[1]+((day-1)*24),round(varobject[index].value, 3),zone.values[0],'Hydro',round(marginal_cost, 3)))    
                        elif index[0] in instance.NY_Imports_CT:
                            marginal_cost = 4.3
                            mwh_2.append((index[0],index[1]+((day-1)*24),round(varobject[index].value, 3),zone.values[0],'Imports',round(marginal_cost, 3)))
                        elif index[0] in instance.NY_Imports_WCMA:
                            marginal_cost = 4.3
                            mwh_2.append((index[0],index[1]+((day-1)*24),round(varobject[index].value, 3),zone.values[0],'Imports',round(marginal_cost, 3)))                
                        elif index[0] in instance.NY_Imports_VT:
                            marginal_cost = 4.3
                            mwh_2.append((index[0],index[1]+((day-1)*24),round(varobject[index].value, 3),zone.values[0],'Imports',round(marginal_cost, 3))) 
                        elif index[0] in instance.HQ_Imports_VT:
                            marginal_cost = 4.3
                            mwh_2.append((index[0],index[1]+((day-1)*24),round(varobject[index].value, 3),zone.values[0],'Imports',round(marginal_cost, 3))) 
                        elif index[0] in instance.NB_Imports_ME:
                            marginal_cost = 4.3
                            mwh_2.append((index[0],index[1]+((day-1)*24),round(varobject[index].value, 3),zone.values[0],'Imports',round(marginal_cost, 3)))


            if a=='mwh_3':

                for index in varobject:

                    name = index[0]
                    g = df_generators[df_generators['name']==name]
                    seg3 = g['seg3'].values
                    seg3 = seg3[0]
                    zone = g['zone']
                    
                    if int(index[1]>0 and index[1]<horizon_end):

                        gas_price = 3.6
    
                        if index[0] in instance.Gas:
                            marginal_cost = seg3*gas_price
                            mwh_3.append((index[0],index[1]+((day-1)*24),round(varobject[index].value, 3),zone.values[0],'Gas',round(marginal_cost, 3)))
                        elif index[0] in instance.Coal:
                            marginal_cost = seg3*2.2
                            mwh_3.append((index[0],index[1]+((day-1)*24),round(varobject[index].value, 3),zone.values[0],'Coal',round(marginal_cost, 3)))
                        elif index[0] in instance.Oil:
                            marginal_cost = seg3*8
                            mwh_3.append((index[0],index[1]+((day-1)*24),round(varobject[index].value, 3),zone.values[0],'Oil',round(marginal_cost, 3)))
                        elif index[0] in instance.Slack:
                            marginal_cost = 700
                            mwh_3.append((index[0],index[1]+((day-1)*24),round(varobject[index].value, 3),zone.values[0],'Slack',round(marginal_cost, 3)))
                        elif index[0] in instance.Hydro:
                            marginal_cost = 0
                            mwh_3.append((index[0],index[1]+((day-1)*24),round(varobject[index].value, 3),zone.values[0],'Hydro',round(marginal_cost, 3)))    
                        elif index[0] in instance.NY_Imports_CT:
                            marginal_cost = 4.3
                            mwh_3.append((index[0],index[1]+((day-1)*24),round(varobject[index].value, 3),zone.values[0],'Imports',round(marginal_cost, 3)))
                        elif index[0] in instance.NY_Imports_WCMA:
                            marginal_cost = 4.3
                            mwh_3.append((index[0],index[1]+((day-1)*24),round(varobject[index].value, 3),zone.values[0],'Imports',round(marginal_cost, 3)))                
                        elif index[0] in instance.NY_Imports_VT:
                            marginal_cost = 4.3
                            mwh_3.append((index[0],index[1]+((day-1)*24),round(varobject[index].value, 3),zone.values[0],'Imports',round(marginal_cost, 3))) 
                        elif index[0] in instance.HQ_Imports_VT:
                            marginal_cost = 4.3
                            mwh_3.append((index[0],index[1]+((day-1)*24),round(varobject[index].value, 3),zone.values[0],'Imports',round(marginal_cost, 3))) 
                        elif index[0] in instance.NB_Imports_ME:
                            marginal_cost = 4.3
                            mwh_3.append((index[0],index[1]+((day-1)*24),round(varobject[index].value, 3),zone.values[0],'Imports',round(marginal_cost, 3)))


            # if a=='on':

            #     for index in varobject:
                    
            #         name = index[0]
            #         g = df_generators[df_generators['name']==name]
            #         zone = g['zone']
                
            #         if int(index[1]>0 and index[1]<horizon_end):
                
            #             on.append((index[0],index[1]+((day-1)*24),varobject[index].value,zone.values[0]))
                

            # if a=='switch':
              

            #     for index in varobject:
                    
            #         name = index[0]
            #         g = df_generators[df_generators['name']==name]
            #         zone = g['zone']
                    
            #         if int(index[1]>0 and index[1]<horizon_end):
                        
            #             switch.append((index[0],index[1]+((day-1)*24),varobject[index].value,zone.values[0]))
                

            # if a=='srsv':

            #     for index in varobject:
                    
            #         name = index[0]
            #         g = df_generators[df_generators['name']==name]
            #         zone = g['zone']
                    
            #         if int(index[1]>0 and index[1]<horizon_end):
                        
            #             srsv.append((index[0],index[1]+((day-1)*24),round(varobject[index].value, 3),zone.values[0]))
                

            # if a=='nrsv':

            #     for index in varobject:
                    
            #         name = index[0]
            #         g = df_generators[df_generators['name']==name]
            #         zone = g['zone']
                    
            #         if int(index[1]>0 and index[1]<horizon_end):
                        
            #             nrsv.append((index[0],index[1]+((day-1)*24),round(varobject[index].value, 3),zone.values[0]))
                

            if a=='offshorewind':

                for index in varobject:
                    
                    if int(index[1]>0 and index[1]<horizon_end):
                        
                        offshore_wind.append((index[0],index[1]+((day-1)*24),round(varobject[index].value, 3)))
            
            
            if a=='solar':

                for index in varobject:
                    
                    if int(index[1]>0 and index[1]<horizon_end):
                        
                        solar.append((index[0],index[1]+((day-1)*24),round(varobject[index].value, 3)))
                        
                        
            if a=='onshorewind':

                for index in varobject:
                    
                    if int(index[1]>0 and index[1]<horizon_end):
                        
                        onshore_wind.append((index[0],index[1]+((day-1)*24),round(varobject[index].value, 3)))
                   
                        
            if a=='mustrun':

                for index in varobject:
                    
                    if int(index[1]>0 and index[1]<horizon_end):
                        
                        must_run.append((index[0],index[1]+((day-1)*24),round(varobject[index].value, 3)))


            if a=='flow':

                for index in varobject:
                    
                    if int(index[2]>0 and index[2]<horizon_end):
                        
                        flow.append((index[0],index[1],index[2]+((day-1)*24),round(varobject[index].value, 3)))


            for j in instance.Generators:
                if instance.on[j,24] == 1:
                    instance.on[j,0] = 1
                else:
                    instance.on[j,0] = 0
                instance.on[j,0].fixed = True

                if instance.mwh_1[j,24].value <=0 and instance.mwh_1[j,24].value>= -0.0001:
                    newval_1=0
                else:
                    newval_1=instance.mwh_1[j,24].value
                instance.mwh_1[j,0] = newval_1
                instance.mwh_1[j,0].fixed = True

                if instance.mwh_2[j,24].value <=0 and instance.mwh_2[j,24].value>= -0.0001:
                    newval=0
                else:
                    newval=instance.mwh_2[j,24].value

                if instance.mwh_3[j,24].value <=0 and instance.mwh_3[j,24].value>= -0.0001:
                    newval2=0
                else:
                    newval2=instance.mwh_3[j,24].value


                instance.mwh_2[j,0] = newval
                instance.mwh_2[j,0].fixed = True
                instance.mwh_3[j,0] = newval2
                instance.mwh_3[j,0].fixed = True
                if instance.switch[j,24] == 1:
                    instance.switch[j,0] = 1
                else:
                    instance.switch[j,0] = 0
                instance.switch[j,0].fixed = True

                if instance.srsv[j,24].value <=0 and instance.srsv[j,24].value>= -0.0001:
                    newval_srsv=0
                else:
                    newval_srsv=instance.srsv[j,24].value
                instance.srsv[j,0] = newval_srsv
                instance.srsv[j,0].fixed = True

                if instance.nrsv[j,24].value <=0 and instance.nrsv[j,24].value>= -0.0001:
                    newval_nrsv=0
                else:
                    newval_nrsv=instance.nrsv[j,24].value
                instance.nrsv[j,0] = newval_nrsv
                instance.nrsv[j,0].fixed = True

        print(day)

    mwh_1_pd=pd.DataFrame(mwh_1,columns=['Generator','Time','Value','Zones','Type','$/MWh'])
    mwh_2_pd=pd.DataFrame(mwh_2,columns=['Generator','Time','Value','Zones','Type','$/MWh'])
    mwh_3_pd=pd.DataFrame(mwh_3,columns=['Generator','Time','Value','Zones','Type','$/MWh'])
    # on_pd=pd.DataFrame(on,columns=['Generator','Time','Value','Zones'])
    # switch_pd=pd.DataFrame(switch,columns=['Generator','Time','Value','Zones'])
    # srsv_pd=pd.DataFrame(srsv,columns=['Generator','Time','Value','Zones'])
    # nrsv_pd=pd.DataFrame(nrsv,columns=['Generator','Time','Value','Zones'])
    solar_pd=pd.DataFrame(solar,columns=['Zone','Time','Value'])
    onshore_wind_pd=pd.DataFrame(onshore_wind,columns=['Zone','Time','Value'])
    offshore_wind_pd=pd.DataFrame(offshore_wind,columns=['Zone','Time','Value'])
    must_run_pd=pd.DataFrame(must_run,columns=['Zone','Time','Value'])
    flow_pd=pd.DataFrame(flow,columns=['Source','Sink','Time','Value'])
    shadow_price=pd.DataFrame(Duals,columns=['Constraint','Time','Value'])
    objective = pd.DataFrame(System_cost,columns=['Value'])

    flow_pd.to_csv('flow.csv',index=False)
    mwh_1_pd.to_csv('mwh_1.csv',index=False)
    mwh_2_pd.to_csv('mwh_2.csv',index=False)
    mwh_3_pd.to_csv('mwh_3.csv',index=False)
    # on_pd.to_csv('on.csv',index=False)
    # switch_pd.to_csv('switch.csv',index=False)
    # srsv_pd.to_csv('srsv.csv',index=False)
    # nrsv_pd.to_csv('nrsv.csv',index=False)
    solar_pd.to_csv('solar_out.csv',index=False)
    onshore_wind_pd.to_csv('onshore_wind_out.csv',index=False)
    offshore_wind_pd.to_csv('offshore_wind_out.csv',index=False)
    must_run_pd.to_csv('mustrun_out.csv',index=False)
    shadow_price.to_csv('shadow_price.csv',index=False)
    objective.to_csv('obj_function.csv',index=False)

    return None
