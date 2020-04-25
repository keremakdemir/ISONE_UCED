# -*- coding: utf-8 -*-
"""
Created on Fri Apr 24 18:45:34 2020

@author: kakdemi
"""

import pandas as pd

#importing generators
all_generators = pd.read_excel('generators2.xlsx', sheet_name='NEISO generators (dispatch)')

#getting all oil generators
all_oil = all_generators[all_generators['typ']=='oil'].copy()

#getting all generators in every zone
CT_oil = all_oil[all_oil['zone']=='CT'].copy()
ME_oil = all_oil[all_oil['zone']=='ME'].copy()
NEMA_oil = all_oil[all_oil['zone']=='NEMA'].copy()
NH_oil = all_oil[all_oil['zone']=='NH'].copy()
RI_oil = all_oil[all_oil['zone']=='RI'].copy()
SEMA_oil = all_oil[all_oil['zone']=='SEMA'].copy()
VT_oil = all_oil[all_oil['zone']=='VT'].copy()
WCMA_oil = all_oil[all_oil['zone']=='WCMA'].copy()

#defining zones
zones = ['CT','ME','NEMA','NH','RI','SEMA','VT','WCMA']

#getting all slack generators
all_slack = all_generators[all_generators['typ']=='slack'].copy()  

#getting generators other than slack and oil
all_other = all_generators[(all_generators['typ']!='oil') & (all_generators['typ']!='slack')].copy()

#defining a function to downsample oil generators
def oil_downsampler(zone):
    
    #copying the oil generators in that zone and sorting wrt to their seg1 heat rate
    Selected_line_oil = globals()[zone+'_oil'].copy()
    sorted_df = Selected_line_oil.sort_values(by=['seg1'])
    sorted_df_reset = sorted_df.reset_index(drop=True)
    
    #creating 3 chunks wrt their heatrates
    heat_rate = list(sorted_df_reset.loc[:,'seg1'])
    num = int(len(heat_rate)/3)
    First_plant = sorted_df_reset.iloc[:num,:].copy()
    Second_plant = sorted_df_reset.iloc[num:num*2,:].copy()
    Third_plant = sorted_df_reset.iloc[num*2:,:].copy()
    
    #finding the relevant parameters for the downsampled oil plants
    First_cap = First_plant.loc[:,'netcap'].sum()
    Second_cap = Second_plant.loc[:,'netcap'].sum()
    Third_cap = Third_plant.loc[:,'netcap'].sum()
    netcap = [First_cap, Second_cap, Third_cap]
    ramp_1 = First_cap
    ramp_2 = Second_cap
    ramp_3 = Third_cap
    ramp = [ramp_1, ramp_2, ramp_3]
    First_min_cap = First_cap*0.35
    Second_min_cap = Second_cap*0.35
    Third_min_cap = Third_cap*0.35
    min_cap = [First_min_cap, Second_min_cap, Third_min_cap]
    Min_u = [1, 1, 1]
    Min_d = [1, 1, 1]
    zones = [zone, zone, zone]
    types = ['oil', 'oil', 'oil']
    seg_1_1 = First_plant.loc[:,'netcap'] * First_plant.loc[:,'seg1']
    seg_1_1_new = seg_1_1.sum()/First_plant.loc[:,'netcap'].sum()
    seg_1_2 = First_plant.loc[:,'netcap'] * First_plant.loc[:,'seg2']
    seg_1_2_new = seg_1_2.sum()/First_plant.loc[:,'netcap'].sum()
    seg_1_3 = First_plant.loc[:,'netcap'] * First_plant.loc[:,'seg3']
    seg_1_3_new = seg_1_3.sum()/First_plant.loc[:,'netcap'].sum()
    seg_2_1 = Second_plant.loc[:,'netcap'] * Second_plant.loc[:,'seg1']
    seg_2_1_new = seg_2_1.sum()/Second_plant.loc[:,'netcap'].sum()
    seg_2_2 = Second_plant.loc[:,'netcap'] * Second_plant.loc[:,'seg2']
    seg_2_2_new = seg_2_2.sum()/Second_plant.loc[:,'netcap'].sum()
    seg_2_3 = Second_plant.loc[:,'netcap'] * Second_plant.loc[:,'seg3']
    seg_2_3_new = seg_2_3.sum()/Second_plant.loc[:,'netcap'].sum()
    seg_3_1 = Third_plant.loc[:,'netcap'] * Third_plant.loc[:,'seg1']
    seg_3_1_new = seg_3_1.sum()/Third_plant.loc[:,'netcap'].sum()
    seg_3_2 = Third_plant.loc[:,'netcap'] * Third_plant.loc[:,'seg2']
    seg_3_2_new = seg_3_2.sum()/Third_plant.loc[:,'netcap'].sum()
    seg_3_3 = Third_plant.loc[:,'netcap'] * Third_plant.loc[:,'seg3']
    seg_3_3_new = seg_3_3.sum()/Third_plant.loc[:,'netcap'].sum()
    seg_1 = [seg_1_1_new, seg_2_1_new, seg_3_1_new]
    seg_2 = [seg_1_2_new, seg_2_2_new, seg_3_2_new]
    seg_3 = [seg_1_3_new, seg_2_3_new, seg_3_3_new]
    var_om_1 = First_plant.loc[:,'netcap'] * First_plant.loc[:,'var_om']
    var_om_1_new = var_om_1.sum()/First_plant.loc[:,'netcap'].sum()
    var_om_2 = Second_plant.loc[:,'netcap'] * Second_plant.loc[:,'var_om']
    var_om_2_new = var_om_2.sum()/Second_plant.loc[:,'netcap'].sum()
    var_om_3 = Third_plant.loc[:,'netcap'] * Third_plant.loc[:,'var_om']
    var_om_3_new = var_om_3.sum()/Third_plant.loc[:,'netcap'].sum()
    var_om = [var_om_1_new, var_om_2_new, var_om_3_new]
    no_load_1 = First_plant.loc[:,'netcap'] * First_plant.loc[:,'no_load']
    no_load_1_new = no_load_1.sum()/First_plant.loc[:,'netcap'].sum()
    no_load_2 = Second_plant.loc[:,'netcap'] * Second_plant.loc[:,'no_load']
    no_load_2_new = no_load_2.sum()/Second_plant.loc[:,'netcap'].sum()
    no_load_3 = Third_plant.loc[:,'netcap'] * Third_plant.loc[:,'no_load']
    no_load_3_new = no_load_3.sum()/Third_plant.loc[:,'netcap'].sum()
    no_load = [no_load_1_new, no_load_2_new, no_load_3_new]
    st_cost_1 = First_plant.loc[:,'netcap'] * First_plant.loc[:,'st_cost']
    st_cost_1_new = st_cost_1.sum()/First_plant.loc[:,'netcap'].sum()
    st_cost_2 = Second_plant.loc[:,'netcap'] * Second_plant.loc[:,'st_cost']
    st_cost_2_new = st_cost_2.sum()/Second_plant.loc[:,'netcap'].sum()
    st_cost_3 = Third_plant.loc[:,'netcap'] * Third_plant.loc[:,'st_cost']
    st_cost_3_new = st_cost_3.sum()/Third_plant.loc[:,'netcap'].sum()
    st_cost = [st_cost_1_new, st_cost_2_new, st_cost_3_new]
    name = [zone+'_agg_oil_1', zone+'_agg_oil_2', zone+'_agg_oil_3']
    
    #creating a dataframe that includes downsampled oil generators
    list_labels = list(WCMA_oil.columns)
    list_columns = [name, types, zones, netcap, seg_1, seg_2, seg_3, min_cap, ramp, Min_u,
                    Min_d, var_om, no_load, st_cost]
    zipped_list = list(zip(list_labels, list_columns))
    gen_df = dict(zipped_list)
    df_oils = pd.DataFrame(gen_df)
    
    return df_oils

#downsampling oil generators in every zone by using the defined function
for z in zones:
    
    globals()[z+'_agg_oil_df'] = oil_downsampler(z)
    
#adding downsampled oil generators to create a complete list of generators
final_generators = pd.concat([all_other, CT_agg_oil_df, ME_agg_oil_df, NEMA_agg_oil_df, 
                              NH_agg_oil_df, RI_agg_oil_df, SEMA_agg_oil_df, VT_agg_oil_df, 
                              WCMA_agg_oil_df, all_slack], ignore_index=True) 

#exporting the generators as an Excel file
final_generators.to_excel('generators.xlsx', sheet_name='NEISO generators (dispatch)', index=False)

