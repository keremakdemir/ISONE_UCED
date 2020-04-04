# -*- coding: utf-8 -*-
"""
Created on Thu Mar 26 19:54:12 2020

@author: kakdemi
"""

import pandas as pd

#importing data
df_data_new = pd.read_excel('Interchange.xlsx',sheet_name='Daily', header=0)

#importing data
SALBRYNB_hourly = pd.read_excel('Interchange.xlsx',sheet_name='SALBRYNB', header=0)
ROSETON_hourly = pd.read_excel('Interchange.xlsx',sheet_name='ROSETON', header=0)
HQ_P1_P2_hourly = pd.read_excel('Interchange.xlsx',sheet_name='HQ_P1_P2', header=0)
HQHIGATE_hourly = pd.read_excel('Interchange.xlsx',sheet_name='HQHIGATE', header=0)
SHOREHAM_hourly = pd.read_excel('Interchange.xlsx',sheet_name='SHOREHAM', header=0)
NORTHPORT_hourly = pd.read_excel('Interchange.xlsx',sheet_name='NORTHPORT', header=0)

#creating a daily index and removing Feb 29s
daily_index = pd.date_range(start='2008-01-01', end='2018-12-31', freq='D')
daily_index = pd.DatetimeIndex.delete(daily_index, [59, 1520, 2981])
daily_interchange = df_data_new.set_index(daily_index)

#Adding all hourly data together and renaming the columns
SALBRYNB_all_hourly = pd.concat([SALBRYNB_hourly[2008], SALBRYNB_hourly[2009], SALBRYNB_hourly[2010], SALBRYNB_hourly[2011], SALBRYNB_hourly[2012], SALBRYNB_hourly[2013], SALBRYNB_hourly[2014], SALBRYNB_hourly[2015], SALBRYNB_hourly[2016], SALBRYNB_hourly[2017], SALBRYNB_hourly[2018]] ,ignore_index=True)
ROSETON_all_hourly = pd.concat([ROSETON_hourly[2008], ROSETON_hourly[2009], ROSETON_hourly[2010], ROSETON_hourly[2011], ROSETON_hourly[2012], ROSETON_hourly[2013], ROSETON_hourly[2014], ROSETON_hourly[2015], ROSETON_hourly[2016], ROSETON_hourly[2017], ROSETON_hourly[2018]] ,ignore_index=True)
HQ_P1_P2_all_hourly = pd.concat([HQ_P1_P2_hourly[2008], HQ_P1_P2_hourly[2009], HQ_P1_P2_hourly[2010], HQ_P1_P2_hourly[2011], HQ_P1_P2_hourly[2012], HQ_P1_P2_hourly[2013], HQ_P1_P2_hourly[2014], HQ_P1_P2_hourly[2015], HQ_P1_P2_hourly[2016], HQ_P1_P2_hourly[2017], HQ_P1_P2_hourly[2018]] ,ignore_index=True)
HQHIGATE_all_hourly = pd.concat([HQHIGATE_hourly[2008], HQHIGATE_hourly[2009], HQHIGATE_hourly[2010], HQHIGATE_hourly[2011], HQHIGATE_hourly[2012], HQHIGATE_hourly[2013], HQHIGATE_hourly[2014], HQHIGATE_hourly[2015], HQHIGATE_hourly[2016], HQHIGATE_hourly[2017], HQHIGATE_hourly[2018]] ,ignore_index=True)
SHOREHAM_all_hourly = pd.concat([SHOREHAM_hourly[2008], SHOREHAM_hourly[2009], SHOREHAM_hourly[2010], SHOREHAM_hourly[2011], SHOREHAM_hourly[2012], SHOREHAM_hourly[2013], SHOREHAM_hourly[2014], SHOREHAM_hourly[2015], SHOREHAM_hourly[2016], SHOREHAM_hourly[2017], SHOREHAM_hourly[2018]] ,ignore_index=True)
NORTHPORT_all_hourly = pd.concat([NORTHPORT_hourly[2008], NORTHPORT_hourly[2009], NORTHPORT_hourly[2010], NORTHPORT_hourly[2011], NORTHPORT_hourly[2012], NORTHPORT_hourly[2013], NORTHPORT_hourly[2014], NORTHPORT_hourly[2015], NORTHPORT_hourly[2016], NORTHPORT_hourly[2017], NORTHPORT_hourly[2018]] ,ignore_index=True)
All_Lines_hourly = pd.concat([SALBRYNB_all_hourly, ROSETON_all_hourly, HQ_P1_P2_all_hourly, HQHIGATE_all_hourly, SHOREHAM_all_hourly, NORTHPORT_all_hourly], axis=1)
All_Lines_hourly.columns = ['SALBRYNB', 'ROSETON', 'HQ_P1_P2', 'HQHIGATE', 'SHOREHAM', 'NORTHPORT']

#creating list to locate Feb 29s
omitted_dates = []
feb_2008 = list(range(1416,1440))
feb_2012 = list(range(36480,36504))
feb_2016 = list(range(71544,71568))
omitted_dates.extend(feb_2008)
omitted_dates.extend(feb_2012)
omitted_dates.extend(feb_2016)

#creating an hourly index and removing Feb 29s
hourly_index = pd.date_range(start='2008-01-01 00:00:00', end='2018-12-31 23:00:00', freq='H')
hourly_index = pd.DatetimeIndex.delete(hourly_index, omitted_dates)
hourly_interchange = All_Lines_hourly.set_index(hourly_index)

#turning daily dates into strings and saving them in a list
daily_date_list = list(daily_index)
daily_date_list = [str(a) for a in daily_date_list]

#defining interchange lines
Interchange_line = ['SALBRYNB', 'ROSETON', 'HQ_P1_P2', 'HQHIGATE', 'SHOREHAM', 'NORTHPORT']

#creating empty dictionaries to store hourly profiles
SALBRYNB_profile_dict = {}
ROSETON_profile_dict = {}
HQ_P1_P2_profile_dict = {}
HQHIGATE_profile_dict = {}
SHOREHAM_profile_dict = {}
NORTHPORT_profile_dict = {}

for line in Interchange_line:
    
    for day_sp in range(len(daily_date_list)):
    
        day_specific = daily_interchange.loc[daily_date_list[day_sp][:-9], line]
        
        if day_specific < 0:
            
            #if specific day's value is negative, pull out hourly data for that day, turn positive values into zero, 
            #take absolute value of negative values, find hourly profile by dividing individually by total sum of that day
            hour_specific = hourly_interchange.loc[daily_date_list[day_sp][:-9], line]
            hour_list = list(hour_specific)
            transform_hour_list = [0 if i > 0 else i for i in hour_list] 
            final_hour_list = [abs(a) if a < 0 else a for a in transform_hour_list]
            total_change = sum(final_hour_list)
            hourly_profile = [b/total_change for b in final_hour_list]
            
            #saving that days into relevant dictionaries
            if line == 'SALBRYNB':
                SALBRYNB_profile_dict[daily_date_list[day_sp][:-9]] = hourly_profile
            elif line == 'ROSETON':
                ROSETON_profile_dict[daily_date_list[day_sp][:-9]] = hourly_profile
            elif line == 'HQ_P1_P2':
                HQ_P1_P2_profile_dict[daily_date_list[day_sp][:-9]] = hourly_profile
            elif line == 'HQHIGATE':
                HQHIGATE_profile_dict[daily_date_list[day_sp][:-9]] = hourly_profile
            elif line == 'SHOREHAM':
                SHOREHAM_profile_dict[daily_date_list[day_sp][:-9]] = hourly_profile
            elif line == 'NORTHPORT':
                NORTHPORT_profile_dict[daily_date_list[day_sp][:-9]] = hourly_profile

        else:
            pass

###############################################
######## AVERAGE YEAR FOR SALBRYNB ############
###############################################

#turn dictionary into dataframe, transpose it and sotre month and day into a list
SALBRYNB_df = pd.DataFrame(SALBRYNB_profile_dict).T
if not SALBRYNB_df.empty:
    
    Month_day = []  
    for row in range(SALBRYNB_df.shape[0]):
        Month_day.append(str(SALBRYNB_df.index[row])[5:10])
    
    #take average of the hourly profiles to find an average year, and reindex the dataframe
    SALBRYNB_df['Date'] = Month_day    
    Average_Year = SALBRYNB_df.groupby(['Date']).mean()     
    New_Index = [str(value) for value in list(Average_Year.index)]
    Average_Year["New_Index"] = New_Index
    Average_Year.set_index("New_Index")
                
    #dropping unused column, and sorting
    del Average_Year['New_Index']                 
    Average_Year_SALBRYNB = Average_Year.sort_index()
    
    #creating a hypothetical index to use get_loc function later
    index_change = list(Average_Year_SALBRYNB.index)
    index_change = [i+'-2010' for i in index_change]
    Average_Year_SALBRYNB['Hyp_Date'] = index_change
    Average_Year_SALBRYNB = Average_Year_SALBRYNB.set_index('Hyp_Date')
    Average_Year_SALBRYNB.index = pd.to_datetime(Average_Year_SALBRYNB.index)
    Average_Year_SALBRYNB = Average_Year_SALBRYNB.sort_index()

else:
    pass

###############################################
######## AVERAGE YEAR FOR ROSETON #############
###############################################

#turn dictionary into dataframe, transpose it and sotre month and day into a list
ROSETON_df = pd.DataFrame(ROSETON_profile_dict).T
if not ROSETON_df.empty:
    
    Month_day = []  
    for row in range(ROSETON_df.shape[0]):
        Month_day.append(str(ROSETON_df.index[row])[5:10])
    
    #take average of the hourly profiles to find an average year, and reindex the dataframe
    ROSETON_df['Date'] = Month_day    
    Average_Year = ROSETON_df.groupby(['Date']).mean()     
    New_Index = [str(value) for value in list(Average_Year.index)]
    Average_Year["New_Index"] = New_Index
    Average_Year.set_index("New_Index")
                
    #dropping unused column, and sorting
    del Average_Year['New_Index']                 
    Average_Year_ROSETON = Average_Year.sort_index()
    
    #creating a hypothetical index to use get_loc function later
    index_change = list(Average_Year_ROSETON.index)
    index_change = [i+'-2010' for i in index_change]
    Average_Year_ROSETON['Hyp_Date'] = index_change
    Average_Year_ROSETON = Average_Year_ROSETON.set_index('Hyp_Date')
    Average_Year_ROSETON.index = pd.to_datetime(Average_Year_ROSETON.index)
    Average_Year_ROSETON = Average_Year_ROSETON.sort_index()

else:
    pass

###############################################
######## AVERAGE YEAR FOR HQ_P1_P2 ############
###############################################

#turn dictionary into dataframe, transpose it and sotre month and day into a list
HQ_P1_P2_df = pd.DataFrame(HQ_P1_P2_profile_dict).T
if not HQ_P1_P2_df.empty:
    
    Month_day = []  
    for row in range(HQ_P1_P2_df.shape[0]):
        Month_day.append(str(HQ_P1_P2_df.index[row])[5:10])
    
    #take average of the hourly profiles to find an average year, and reindex the dataframe
    HQ_P1_P2_df['Date'] = Month_day    
    Average_Year = HQ_P1_P2_df.groupby(['Date']).mean()     
    New_Index = [str(value) for value in list(Average_Year.index)]
    Average_Year["New_Index"] = New_Index
    Average_Year.set_index("New_Index")
                
    #dropping unused column, and sorting
    del Average_Year['New_Index']                 
    Average_Year_HQ_P1_P2 = Average_Year.sort_index()
    
    #creating a hypothetical index to use get_loc function later
    index_change = list(Average_Year_HQ_P1_P2.index)
    index_change = [i+'-2010' for i in index_change]
    Average_Year_HQ_P1_P2['Hyp_Date'] = index_change
    Average_Year_HQ_P1_P2 = Average_Year_HQ_P1_P2.set_index('Hyp_Date')
    Average_Year_HQ_P1_P2.index = pd.to_datetime(Average_Year_HQ_P1_P2.index)
    Average_Year_HQ_P1_P2 = Average_Year_HQ_P1_P2.sort_index()

else:
    pass

###############################################
######## AVERAGE YEAR FOR HQHIGATE ############
###############################################

#turn dictionary into dataframe, transpose it and sotre month and day into a list
HQHIGATE_df = pd.DataFrame(HQHIGATE_profile_dict).T
if not HQHIGATE_df.empty:
    
    Month_day = []  
    for row in range(HQHIGATE_df.shape[0]):
        Month_day.append(str(HQHIGATE_df.index[row])[5:10])
    
    #take average of the hourly profiles to find an average year, and reindex the dataframe
    HQHIGATE_df['Date'] = Month_day    
    Average_Year = HQHIGATE_df.groupby(['Date']).mean()     
    New_Index = [str(value) for value in list(Average_Year.index)]
    Average_Year["New_Index"] = New_Index
    Average_Year.set_index("New_Index")
                
    #dropping unused column, and sorting
    del Average_Year['New_Index']                 
    Average_Year_HQHIGATE = Average_Year.sort_index()
    
    #creating a hypothetical index to use get_loc function later
    index_change = list(Average_Year_HQHIGATE.index)
    index_change = [i+'-2010' for i in index_change]
    Average_Year_HQHIGATE['Hyp_Date'] = index_change
    Average_Year_HQHIGATE = Average_Year_HQHIGATE.set_index('Hyp_Date')
    Average_Year_HQHIGATE.index = pd.to_datetime(Average_Year_HQHIGATE.index)
    Average_Year_HQHIGATE = Average_Year_HQHIGATE.sort_index()
    
else:
    pass

###############################################
######## AVERAGE YEAR FOR SHOREHAM ############
###############################################

#turn dictionary into dataframe, transpose it and sotre month and day into a list
SHOREHAM_df = pd.DataFrame(SHOREHAM_profile_dict).T
if not SHOREHAM_df.empty:
    
    Month_day = []  
    for row in range(SHOREHAM_df.shape[0]):
        Month_day.append(str(SHOREHAM_df.index[row])[5:10])
    
    #take average of the hourly profiles to find an average year, and reindex the dataframe
    SHOREHAM_df['Date'] = Month_day    
    Average_Year = SHOREHAM_df.groupby(['Date']).mean()     
    New_Index = [str(value) for value in list(Average_Year.index)]
    Average_Year["New_Index"] = New_Index
    Average_Year.set_index("New_Index")
                
    #dropping unused column, and sorting
    del Average_Year['New_Index']                 
    Average_Year_SHOREHAM = Average_Year.sort_index()
    
    #creating a hypothetical index to use get_loc function later
    index_change = list(Average_Year_SHOREHAM.index)
    index_change = [i+'-2010' for i in index_change]
    Average_Year_SHOREHAM['Hyp_Date'] = index_change
    Average_Year_SHOREHAM = Average_Year_SHOREHAM.set_index('Hyp_Date')
    Average_Year_SHOREHAM.index = pd.to_datetime(Average_Year_SHOREHAM.index)
    Average_Year_SHOREHAM = Average_Year_SHOREHAM.sort_index()
    
else:
    pass

###############################################
######## AVERAGE YEAR FOR NORTHPORT ###########
###############################################

#turn dictionary into dataframe, transpose it and sotre month and day into a list
NORTHPORT_df = pd.DataFrame(NORTHPORT_profile_dict).T
if not NORTHPORT_df.empty:
    
    Month_day = []  
    for row in range(NORTHPORT_df.shape[0]):
        Month_day.append(str(NORTHPORT_df.index[row])[5:10])
    
    #take average of the hourly profiles to find an average year, and reindex the dataframe
    NORTHPORT_df['Date'] = Month_day    
    Average_Year = NORTHPORT_df.groupby(['Date']).mean()     
    New_Index = [str(value) for value in list(Average_Year.index)]
    Average_Year["New_Index"] = New_Index
    Average_Year.set_index("New_Index")
                
    #dropping unused column, and sorting
    del Average_Year['New_Index']                 
    Average_Year_NORTHPORT = Average_Year.sort_index()
    
    #creating a hypothetical index to use get_loc function later
    index_change = list(Average_Year_NORTHPORT.index)
    index_change = [i+'-2010' for i in index_change]
    Average_Year_NORTHPORT['Hyp_Date'] = index_change
    Average_Year_NORTHPORT = Average_Year_NORTHPORT.set_index('Hyp_Date')
    Average_Year_NORTHPORT.index = pd.to_datetime(Average_Year_NORTHPORT.index)
    Average_Year_NORTHPORT = Average_Year_NORTHPORT.sort_index()
    
else:
    pass

################## FILLING AND EXPORTING HOURLY PROFILES #####################

#creating a hypothetical profile days and turning into strings and saving them in a list
profile_index = pd.date_range(start='2010-01-01', end='2010-12-31', freq='D')
profile_date_list = list(profile_index)
profile_date_list = [str(a) for a in profile_date_list]

#defining a function to fill the missing hourly profiles by copying from the nearest profile
def data_filler(line_name):
    
    Filled_Average_Year = globals()['Average_Year_'+line_name].copy()
    
    for ave_date in profile_date_list:
        
        lookup_date = pd.to_datetime(ave_date)
        
        if lookup_date in Filled_Average_Year.index:
            
            pass
        
        else:
            
            selected_profile = list(globals()['Average_Year_'+line_name].iloc[globals()['Average_Year_'+line_name].index.get_loc(lookup_date, method='nearest'), :])
            Filled_Average_Year.loc[lookup_date] = selected_profile
            Filled_Average_Year = Filled_Average_Year.sort_index()
            
    return Filled_Average_Year

#completing hourly profiles by using the function we created
Final_Hourly_Profile_SALBRYNB = data_filler('SALBRYNB')
Final_Hourly_Profile_ROSETON = data_filler('ROSETON')
Final_Hourly_Profile_HQ_P1_P2 = data_filler('HQ_P1_P2')
Final_Hourly_Profile_HQHIGATE = data_filler('HQHIGATE')
Final_Hourly_Profile_SHOREHAM = data_filler('SHOREHAM')
Final_Hourly_Profile_NORTHPORT = data_filler('NORTHPORT')
  
#exporting data into an excel file
writer = pd.ExcelWriter('../../Model_setup/Path_setup/NEISO_path_export_profiles.xlsx', engine='xlsxwriter')
Final_Hourly_Profile_SALBRYNB.to_excel(writer, sheet_name='SALBRYNB', index=False, header=False)
Final_Hourly_Profile_ROSETON.to_excel(writer, sheet_name='ROSETON', index=False, header=False)
Final_Hourly_Profile_HQ_P1_P2.to_excel(writer, sheet_name='HQ_P1_P2', index=False, header=False)
Final_Hourly_Profile_HQHIGATE.to_excel(writer, sheet_name='HQHIGATE', index=False, header=False)
Final_Hourly_Profile_SHOREHAM.to_excel(writer, sheet_name='SHOREHAM', index=False, header=False)
Final_Hourly_Profile_NORTHPORT.to_excel(writer, sheet_name='NORTHPORT', index=False, header=False)
writer.save()