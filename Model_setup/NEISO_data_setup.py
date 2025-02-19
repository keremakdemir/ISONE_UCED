# -*- coding: utf-8 -*-
"""
Created on Wed May 03 15:01:31 2017

@author: jdkern
"""

import pandas as pd
import numpy as np

#read generator parameters into DataFrame
df_gen = pd.read_excel('NEISO_data_file/generators.xlsx',header=0)

#read transmission path parameters into DataFrame
df_paths = pd.read_csv('NEISO_data_file/paths.csv',header=0)

#list zones
zones = ['CT', 'ME', 'NH', 'NEMA', 'RI', 'SEMA', 'VT', 'WCMA']

##time series of load for each zone
df_load_all = pd.read_csv('../Time_series_data/Synthetic_demand_pathflows/Sim_hourly_load.csv',header=0)
df_load_all = df_load_all[zones]

##daily hydropower availability 
df_hydro = pd.read_csv('Hydro_setup/NEISO_dispatchable_hydro.csv',header=0)

#must run resources (LFG,ag_waste,nuclear)
df_must = pd.read_excel('NEISO_data_file/must_run.xlsx',header=0)

# must run generation 
must_run_CT = []
must_run_ME = []
must_run_NEMA = []
must_run_NH = []   
must_run_RI = []
must_run_SEMA = []
must_run_VT = []
must_run_WCMA = [] 

must_run_CT = np.ones((8760,1))*df_must.loc[0,'CT']
must_run_ME = np.ones((8760,1))*df_must.loc[0,'ME']
must_run_NEMA = np.ones((8760,1))*df_must.loc[0,'NEMA']
must_run_NH = np.ones((8760,1))*df_must.loc[0,'NH']
must_run_RI = np.ones((8760,1))*df_must.loc[0,'RI']
must_run_SEMA = np.ones((8760,1))*df_must.loc[0,'SEMA']
must_run_VT = np.ones((8760,1))*df_must.loc[0,'VT']
must_run_WCMA = np.ones((8760,1))*df_must.loc[0,'WCMA']
must_run = np.column_stack((must_run_CT,must_run_ME,must_run_NEMA,must_run_NH,must_run_RI,must_run_SEMA,must_run_VT,must_run_WCMA))
df_total_must_run = pd.DataFrame(must_run,columns=('CT','ME','NEMA','NH','RI','SEMA','VT','WCMA'))
df_total_must_run.to_csv('NEISO_data_file/must_run_hourly.csv')

#natural gas prices
df_ng_all = pd.read_excel('../Time_series_data/Gas_prices/NG.xlsx', header=0)
df_ng_all = df_ng_all[zones]

#oil prices
df_oil_all = pd.read_excel('../Time_series_data/Oil_prices/Oil_prices.xlsx', header=0)
df_oil_all = df_oil_all[zones]

# time series of offshore wind generation for each zone
df_offshore_wind_all = pd.read_excel('../Time_series_data/Synthetic_wind_power/offshore_wind_power_sim.xlsx',header=0)

# time series of solar generation
df_solar = pd.read_excel('NEISO_data_file/hourly_solar_gen.xlsx',header=0)
solar_caps = pd.read_excel('NEISO_data_file/solar_caps.xlsx',header=0)

# time series of onshore wind generation
df_onshore_wind = pd.read_excel('NEISO_data_file/hourly_onshore_wind_gen.xlsx',header=0)
onshore_wind_caps = pd.read_excel('NEISO_data_file/wind_onshore_caps.xlsx',header=0)


def setup(year, Hub_height, Offshore_capacity):
    
    ##time series of natural gas prices for each zone
    df_ng = globals()['df_ng_all'].copy()
    df_ng = df_ng.reset_index()
    
    ##time series of oil prices for each zone
    df_oil = globals()['df_oil_all'].copy()
    df_oil = df_oil.reset_index()
    
    ##time series of load for each zone
    df_load = globals()['df_load_all'].loc[year*8760:year*8760+8759].copy()
    df_load = df_load.reset_index(drop=True)
    
    ##time series of operational reserves for each zone
    rv= df_load.values
    reserves = np.zeros((len(rv),1))
    for i in range(0,len(rv)):
            reserves[i] = np.sum(rv[i,:])*.04
    df_reserves = pd.DataFrame(reserves)
    df_reserves.columns = ['reserves']
    
    ##daily time series of dispatchable imports by path
    df_imports = pd.read_csv('Path_setup/NEISO_dispatchable_imports.csv',header=0)
    
    ##hourly time series of exports by zone
    df_exports = pd.read_csv('Path_setup/NEISO_exports.csv',header=0)
        
    # time series of offshore wind generation for each zone
    df_offshore_wind = globals()['df_offshore_wind_all'].loc[:, Hub_height].copy()
    df_offshore_wind = df_offshore_wind.loc[year*8760:year*8760+8759]
    df_offshore_wind = df_offshore_wind.reset_index()
    offshore_wind_caps = pd.read_excel('NEISO_data_file/wind_offshore_caps.xlsx')
            
    ############
    #  sets    #
    ############
    
    #write data.dat file
    import os
    from shutil import copy
    from pathlib import Path
    
    path = str(Path.cwd().parent) + str(Path('/UCED/LR/NEISO' +'_'+ str(Hub_height) +'_'+ str(Offshore_capacity) +'_'+ str(year)))
    os.makedirs(path,exist_ok=True)
    
    generators_file='NEISO_data_file/generators.xlsx'
    dispatch_file='../UCED/NEISO_dispatch.py'
    dispatchLP_file='../UCED/NEISO_dispatchLP.py'
    wrapper_file='../UCED/NEISO_wrapper.py'
    simulation_file='../UCED/NEISO_simulation.py'
    
    copy(dispatch_file,path)
    copy(wrapper_file,path)
    copy(simulation_file,path)
    copy(dispatchLP_file,path)
    copy(generators_file,path)
    
    filename = path + '/data.dat'
    
    #write data.dat file
    # filename = 'NEISO_data_file/data.dat'
    with open(filename, 'w') as f:
        
        # generator sets by zone
        for z in zones:
            # zone string
            z_int = zones.index(z)
            f.write('set Zone%dGenerators :=\n' % (z_int+1))
            # pull relevant generators
            for gen in range(0,len(df_gen)):
                if df_gen.loc[gen,'zone'] == z:
                    unit_name = df_gen.loc[gen,'name']
                    unit_name = unit_name.replace(' ','_')
                    f.write(unit_name + ' ')
            f.write(';\n\n')   
            
        # NY imports
        f.write('set NY_Imports_CT :=\n')
        # pull relevant generators
        for gen in range(0,len(df_gen)):
            if df_gen.loc[gen,'typ'] == 'imports' and df_gen.loc[gen,'zone'] == 'NYCT_I':
                unit_name = df_gen.loc[gen,'name']
                unit_name = unit_name.replace(' ','_')
                f.write(unit_name + ' ')
        f.write(';\n\n')   
    
        # NY imports
        f.write('set NY_Imports_WCMA :=\n')
        # pull relevant generators
        for gen in range(0,len(df_gen)):
            if df_gen.loc[gen,'typ'] == 'imports' and df_gen.loc[gen,'zone'] == 'NYWCMA_I':
                unit_name = df_gen.loc[gen,'name']
                unit_name = unit_name.replace(' ','_')
                f.write(unit_name + ' ')
        f.write(';\n\n') 
        
        # NY imports
        f.write('set NY_Imports_VT :=\n')
        # pull relevant generators
        for gen in range(0,len(df_gen)):
            if df_gen.loc[gen,'typ'] == 'imports' and df_gen.loc[gen,'zone'] == 'NYVT_I':
                unit_name = df_gen.loc[gen,'name']
                unit_name = unit_name.replace(' ','_')
                f.write(unit_name + ' ')
        f.write(';\n\n')
        
        # HQ imports
        f.write('set HQ_Imports_VT :=\n')
        # pull relevant generators
        for gen in range(0,len(df_gen)):
            if df_gen.loc[gen,'typ'] == 'imports' and df_gen.loc[gen,'zone'] == 'HQVT_I':
                unit_name = df_gen.loc[gen,'name']
                unit_name = unit_name.replace(' ','_')
                f.write(unit_name + ' ')
        f.write(';\n\n')
            
        # NB imports
        f.write('set NB_Imports_ME :=\n')
        # pull relevant generators
        for gen in range(0,len(df_gen)):
            if df_gen.loc[gen,'typ'] == 'imports' and df_gen.loc[gen,'zone'] == 'NBME_I':
                unit_name = df_gen.loc[gen,'name']
                unit_name = unit_name.replace(' ','_')
                f.write(unit_name + ' ')
        f.write(';\n\n')
            
        # generator sets by type
        # coal
        f.write('set Coal :=\n')
        # pull relevant generators
        for gen in range(0,len(df_gen)):
            if df_gen.loc[gen,'typ'] == 'coal':
                unit_name = df_gen.loc[gen,'name']
                unit_name = unit_name.replace(' ','_')
                f.write(unit_name + ' ')
        f.write(';\n\n')    
       
        # # oil
        # f.write('set Oil :=\n')
        # # pull relevant generators
        # for gen in range(0,len(df_gen)):
        #     if df_gen.loc[gen,'typ'] == 'oil':
        #         unit_name = df_gen.loc[gen,'name']
        #         unit_name = unit_name.replace(' ','_')
        #         f.write(unit_name + ' ')
        # f.write(';\n\n')        
     
        
        # Slack
        f.write('set Slack :=\n')
        # pull relevant generators
        for gen in range(0,len(df_gen)):
            if df_gen.loc[gen,'typ'] == 'slack':
                unit_name = df_gen.loc[gen,'name']
                unit_name = unit_name.replace(' ','_')
                f.write(unit_name + ' ')
        f.write(';\n\n')  
    
        # Hydro
        f.write('set Hydro :=\n')
        # pull relevant generators
        for gen in range(0,len(df_gen)):
            if df_gen.loc[gen,'typ'] == 'hydro':
                unit_name = df_gen.loc[gen,'name']
                unit_name = unit_name.replace(' ','_')
                f.write(unit_name + ' ')
        f.write(';\n\n')   
        
        # Ramping
        f.write('set Ramping :=\n')
        # pull relevant generators
        for gen in range(0,len(df_gen)):
            if df_gen.loc[gen,'typ'] == 'hydro' or df_gen.loc[gen,'typ'] == 'imports':
                unit_name = df_gen.loc[gen,'name']
                unit_name = unit_name.replace(' ','_')
                f.write(unit_name + ' ')
        f.write(';\n\n')   
    
    
        # gas generator sets by zone and type
        for z in zones:
            # zone string
            z_int = zones.index(z)
            
            # Natural Gas
            # find relevant generators
            trigger = 0
            for gen in range(0,len(df_gen)):
                if df_gen.loc[gen,'zone'] == z and (df_gen.loc[gen,'typ'] == 'ngcc' or df_gen.loc[gen,'typ'] == 'ngct' or df_gen.loc[gen,'typ'] == 'ngst'):
                    trigger = 1
            if trigger > 0:
                # pull relevant generators
                f.write('set Zone%dGas :=\n' % (z_int+1))      
                for gen in range(0,len(df_gen)):
                    if df_gen.loc[gen,'zone'] == z and (df_gen.loc[gen,'typ'] == 'ngcc' or df_gen.loc[gen,'typ'] == 'ngct' or df_gen.loc[gen,'typ'] == 'ngst'):
                        unit_name = df_gen.loc[gen,'name']
                        unit_name = unit_name.replace(' ','_')
                        f.write(unit_name + ' ')
                f.write(';\n\n')
                
        
        # oil generator sets by zone and type
        for z in zones:
            # zone string
            z_int = zones.index(z)
            
            # find relevant generators
            trigger = 0
            for gen in range(0,len(df_gen)):
                if (df_gen.loc[gen,'zone'] == z) and (df_gen.loc[gen,'typ'] == 'oil'):
                    trigger = 1
            if trigger > 0:
                # pull relevant generators
                f.write('set Zone%dOil :=\n' % (z_int+1))      
                for gen in range(0,len(df_gen)):
                    if (df_gen.loc[gen,'zone'] == z) and (df_gen.loc[gen,'typ'] == 'oil'):
                        unit_name = df_gen.loc[gen,'name']
                        unit_name = unit_name.replace(' ','_')
                        f.write(unit_name + ' ')
                f.write(';\n\n')
       
                
        # zones
        f.write('set zones :=\n')
        for z in zones:
            f.write(z + ' ')
        f.write(';\n\n')
        
        # sources
        f.write('set sources :=\n')
        for z in zones:
            f.write(z + ' ')
        f.write(';\n\n')
        
        # sinks
        f.write('set sinks :=\n')
        for z in zones:
            f.write(z + ' ')
        f.write(';\n\n')
        
    ################
    #  parameters  #
    ################
        
        # simulation details
        SimHours = 8760
        f.write('param SimHours := %d;' % SimHours)
        f.write('\n')
        f.write('param SimDays:= %d;' % int(SimHours/24))
        f.write('\n\n')
        HorizonHours = 48
        f.write('param HorizonHours := %d;' % HorizonHours)
        f.write('\n\n')
        HorizonDays = int(HorizonHours/24)
        f.write('param HorizonDays := %d;' % HorizonDays)
        f.write('\n\n')
        
    
        # create parameter matrix for transmission paths (source and sink connections)
        f.write('param:' + '\t' + 'limit' + '\t' +'hurdle :=' + '\n')
        for z in zones:
            for x in zones:           
                f.write(z + '\t' + x + '\t')
                match = 0
                for p in range(0,len(df_paths)):
                    source = df_paths.loc[p,'start_zone']
                    sink = df_paths.loc[p,'end_zone']
                    if source == z and sink == x:
                        match = 1
                        p_match = p
                if match > 0:
                    f.write(str(round(df_paths.loc[p_match,'limit'],3)) + '\t' + str(round(df_paths.loc[p_match,'hurdle'],3)) + '\n')
                else:
                    f.write('0' + '\t' + '0' + '\n')
        f.write(';\n\n')
        
    # create parameter matrix for generators
        f.write('param:' + '\t')
        for c in df_gen.columns:
            if c != 'name':
                f.write(c + '\t')
        f.write(':=\n\n')
        for i in range(0,len(df_gen)):    
            for c in df_gen.columns:
                if c == 'name':
                    unit_name = df_gen.loc[i,'name']
                    unit_name = unit_name.replace(' ','_')
                    f.write(unit_name + '\t')  
                elif c == 'typ' or c == 'zone':
                    f.write(str(df_gen.loc[i,c]) + '\t')    
                else:
                    f.write(str(round(df_gen.loc[i,c],3)) + '\t')               
            f.write('\n')
                
        f.write(';\n\n')     
        
        # times series data
        # zonal (hourly)
        f.write('param:' + '\t' + 'SimDemand' + '\t' + 'SimOffshoreWind' \
        + '\t' + 'SimSolar' + '\t' + 'SimOnshoreWind' + '\t' + 'SimMustRun:=' + '\n')      
        for z in zones:
            wz = offshore_wind_caps.loc[0,z]
            sz = solar_caps.loc[0,z]
            owz = onshore_wind_caps.loc[0,z]
            for h in range(0,len(df_load)): 
                f.write(z + '\t' + str(h+1) + '\t' + str(round(df_load.loc[h,z],3))\
                + '\t' + str(round(df_offshore_wind.loc[h,Hub_height]*wz,3))\
                + '\t' + str(round(df_solar.loc[h,'Solar_Output_MWh']*sz,3))\
                + '\t' + str(round(df_onshore_wind.loc[h,'Onshore_Output_MWh']*owz,3))\
                + '\t' + str(round(df_total_must_run.loc[h,z],3)) + '\n')
        f.write(';\n\n')
        
        # zonal (daily)
        f.write('param:' + '\t' + 'SimGasPrice' + '\t' + 'SimOilPrice:=' + '\n')      
        for z in zones:
            for d in range(0,int(SimHours/24)): 
                f.write(z + '\t' + str(d+1) + '\t' + str(round(df_ng.loc[d,z], 3)) + '\t' + str(round(df_oil.loc[d,z], 3)) + '\n')
        f.write(';\n\n')
    
        #system wide (daily)
        f.write('param:' + '\t' + 'SimNY_imports_CT' + '\t' + 'SimNY_imports_VT' + '\t' + 'SimNY_imports_WCMA' + '\t' + 'SimNB_imports_ME' + '\t' + 'SimHQ_imports_VT' + '\t' + 'SimCT_hydro' + '\t' + 'SimME_hydro' + '\t' +  'SimNH_hydro' + '\t' +  'SimNEMA_hydro' + '\t' +  'SimRI_hydro' + '\t' +  'SimVT_hydro' + '\t' + 'SimWCMA_hydro:=' + '\n')
        for d in range(0,len(df_imports)):
                f.write(str(d+1) + '\t' + str(round(df_imports.loc[d,'NY_imports_CT'],3)) + '\t' + str(round(df_imports.loc[d,'NY_imports_VT'],3)) + '\t' + str(round(df_imports.loc[d,'NY_imports_WCMA'],3)) + '\t' + str(round(df_imports.loc[d,'NB_imports_ME'],3)) + '\t' + str(round(df_imports.loc[d,'HQ_imports_VT'],3)) + '\t' + str(round(df_hydro.loc[d,'CT'],3)) + '\t' + str(round(df_hydro.loc[d,'ME'],3)) + '\t' + str(round(df_hydro.loc[d,'NH'],3)) + '\t' + str(round(df_hydro.loc[d,'NEMA'],3)) + '\t' + str(round(df_hydro.loc[d,'RI'],3)) + '\t' + str(round(df_hydro.loc[d,'VT'],3)) + '\t' + str(round(df_hydro.loc[d,'WCMA'],3)) + '\n')
        f.write(';\n\n')
            
        #system wide (hourly)
        f.write('param:' + '\t' + 'SimCT_exports_NY' + '\t' + 'SimWCMA_exports_NY' + '\t' + 'SimVT_exports_NY' + '\t' + 'SimVT_exports_HQ' + '\t' + 'SimME_exports_NB' + '\t' + 'SimReserves:=' + '\n')
        for h in range(0,len(df_load)):
                f.write(str(h+1) + '\t' + str(round(df_exports.loc[h,'CT_exports_NY'],3)) + '\t' + str(round(df_exports.loc[h,'WCMA_exports_NY'],3)) + '\t' + str(round(df_exports.loc[h,'VT_exports_NY'],3)) + '\t' + str(round(df_exports.loc[h,'VT_exports_HQ'],3)) + '\t' + str(round(df_exports.loc[h,'ME_exports_NB'],3)) + '\t' + str(round(df_reserves.loc[h,'reserves'],3))  + '\n')
        f.write(';\n\n')
    
    return None
    
    