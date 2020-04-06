# -*- coding: utf-8 -*-
"""
Created on Mon May 14 17:29:16 2018

@author: jdkern
"""
from __future__ import division
import pandas as pd
import numpy as np

def exchange(year):

    df_data = pd.read_csv('../Time_series_data/Synthetic_demand_pathflows/Sim_daily_interchange.csv',header=0)
    paths = ['SALBRYNB', 'ROSETON', 'HQ_P1_P2', 'HQHIGATE', 'SHOREHAM', 'NORTHPORT']
    df_data = df_data[paths]
    df_data = df_data.loc[year*365:year*365+364,:]
    
    # select dispatchable imports (positive flow days)
    imports = df_data
    imports = imports.reset_index()
    
    for p in paths:
        for i in range(0,len(imports)):     
            
            if imports.loc[i,p] < 0:
                imports.loc[i,p] = 0
            else:
                pass  
            
    imports_total = imports.copy()
    imports_total.rename(columns={'SALBRYNB':'NB_imports_ME'}, inplace=True)
    imports_total['HYDRO_QUEBEC'] = imports_total['HQ_P1_P2'] + imports_total['HQHIGATE']
    imports_total['NEWYORK'] = imports_total['ROSETON'] + imports_total['SHOREHAM'] + imports_total['NORTHPORT']
    imports_total['NY_imports_CT'] = imports_total['NEWYORK'].mul(4).div(9)
    imports_total['NY_imports_WCMA'] = imports_total['NEWYORK'].div(9)
    imports_total['NY_imports_VT'] = imports_total['NEWYORK'].mul(4).div(9)
    imports_total.rename(columns={'HYDRO_QUEBEC':'HQ_imports_VT'}, inplace=True)
    del imports_total['ROSETON']
    del imports_total['HQ_P1_P2']
    del imports_total['HQHIGATE']
    del imports_total['SHOREHAM']
    del imports_total['NORTHPORT']
    del imports_total['NEWYORK']
    imports_total.to_csv('Path_setup/NEISO_dispatchable_imports.csv') 
    
    # hourly exports
    df_data = pd.read_csv('../Time_series_data/Synthetic_demand_pathflows/Sim_daily_interchange.csv',header=0)
    df_data = df_data[paths]
    df_data = df_data.loc[year*365:year*365+364,:]
    df_data = df_data.reset_index()
    
    e = np.zeros((8760,len(paths)))
    
    #SALBRYNB
    path_profiles = pd.read_excel('Path_setup/NEISO_path_export_profiles.xlsx',sheet_name='SALBRYNB',header=None)
    pp = path_profiles.values
    
    for i in range(0,len(df_data)):
        if df_data.loc[i,'SALBRYNB'] < 0:
            e[i*24:i*24+24,0] = pp[i,:]*df_data.loc[i,'SALBRYNB']*-1
      
    #ROSETON
    path_profiles = pd.read_excel('Path_setup/NEISO_path_export_profiles.xlsx',sheet_name='ROSETON',header=None)
    pp = path_profiles.values
    
    for i in range(0,len(df_data)):
        if df_data.loc[i,'ROSETON'] < 0:
            e[i*24:i*24+24,1] = pp[i,:]*df_data.loc[i,'ROSETON']*-1
    
    #HQ_P1_P2
    path_profiles = pd.read_excel('Path_setup/NEISO_path_export_profiles.xlsx',sheet_name='HQ_P1_P2',header=None)
    pp = path_profiles.values
    
    for i in range(0,len(df_data)):
        if df_data.loc[i,'HQ_P1_P2'] < 0:
            e[i*24:i*24+24,2] = pp[i,:]*df_data.loc[i,'HQ_P1_P2']*-1
    
    #HQHIGATE
    path_profiles = pd.read_excel('Path_setup/NEISO_path_export_profiles.xlsx',sheet_name='HQHIGATE',header=None)
    pp = path_profiles.values
    
    for i in range(0,len(df_data)):
        if df_data.loc[i,'HQHIGATE'] < 0:
            e[i*24:i*24+24,3] = pp[i,:]*df_data.loc[i,'HQHIGATE']*-1
    
    #SHOREHAM
    path_profiles = pd.read_excel('Path_setup/NEISO_path_export_profiles.xlsx',sheet_name='SHOREHAM',header=None)
    pp = path_profiles.values
    
    for i in range(0,len(df_data)):
        if df_data.loc[i,'SHOREHAM'] < 0:
            e[i*24:i*24+24,4] = pp[i,:]*df_data.loc[i,'SHOREHAM']*-1
    
    #NORTHPORT
    path_profiles = pd.read_excel('Path_setup/NEISO_path_export_profiles.xlsx',sheet_name='NORTHPORT',header=None)
    pp = path_profiles.values
    
    for i in range(0,len(df_data)):
        if df_data.loc[i,'NORTHPORT'] < 0:
            e[i*24:i*24+24,5] = pp[i,:]*df_data.loc[i,'NORTHPORT']*-1
    
    exports = pd.DataFrame(e) 
    exports.columns = paths
    exports_total = exports.copy()
    exports_total.rename(columns={'SALBRYNB':'ME_exports_NB'}, inplace=True)
    exports_total['HYDRO_QUEBEC'] = exports_total['HQ_P1_P2'] + exports_total['HQHIGATE']
    exports_total['NEWYORK'] = exports_total['ROSETON'] + exports_total['SHOREHAM'] + exports_total['NORTHPORT']
    exports_total['CT_exports_NY'] = exports_total['NEWYORK'].mul(4).div(9)
    exports_total['WCMA_exports_NY'] = exports_total['NEWYORK'].div(9)
    exports_total['VT_exports_NY'] = exports_total['NEWYORK'].mul(4).div(9)
    exports_total.rename(columns={'HYDRO_QUEBEC':'VT_exports_HQ'}, inplace=True)
    del exports_total['ROSETON']
    del exports_total['HQ_P1_P2']
    del exports_total['HQHIGATE']
    del exports_total['SHOREHAM']
    del exports_total['NORTHPORT']
    del exports_total['NEWYORK']
    exports_total.to_csv('Path_setup/NEISO_exports.csv')
    
    return None
