# -*- coding: utf-8 -*-
"""
Created on Fri Jul 07 12:23:45 2017

@author: jdkern
"""

#######################################################################################################
# a basic unit commitment model for CAISO system                                                       #
# This is the trial version of the electricity market model                                            #
# 4 Zone system                                                                                        #                                                                                #
#######################################################################################################


from __future__ import division # This line is used to ensure that int or long division arguments are converted to floating point values before division is performed 
from pyomo.environ import * # This command makes the symbols used by Pyomo known to Python
from pyomo.opt import SolverFactory
import itertools

##Create a solver
opt = SolverFactory('gurobi')

model = AbstractModel()
#
######################################################################
## string indentifiers for the set of generators in different zones. #
######################################################################
#

model.Zone1Gas = Set()
model.Zone1Oil = Set()
model.Zone1Generators =  Set()
# Connecticut (CT) 

model.Zone2Gas = Set()
model.Zone2Oil = Set()
model.Zone2Generators = Set()
# Maine (ME)

model.Zone3Gas = Set()
model.Zone3Oil = Set()
model.Zone3Generators = Set()
# New Hampshire (NH)

model.Zone4Gas = Set()
model.Zone4Oil = Set()
model.Zone4Generators = Set()
# Northeast Massachusetts (NEMA)

model.Zone5Gas = Set()
model.Zone5Oil = Set()
model.Zone5Generators =  Set()
# Rhode Island (RI)

model.Zone6Gas = Set()
model.Zone6Oil = Set()
model.Zone6Generators = Set()
# Southeast Massachusetts (SEMA)

model.Zone7Gas = Set()
model.Zone7Oil = Set()
model.Zone7Generators = Set()
# Vermont (VT)

model.Zone8Gas = Set()
model.Zone8Oil = Set()
model.Zone8Generators = Set()
#Western and Central Massachusetts (WCMA)

model.Coal = Set()
model.Gas = model.Zone1Gas | model.Zone2Gas | model.Zone3Gas | model.Zone4Gas | model.Zone5Gas | model.Zone6Gas | model.Zone7Gas | model.Zone8Gas 
model.Oil = model.Zone1Oil | model.Zone2Oil | model.Zone3Oil | model.Zone4Oil | model.Zone5Oil | model.Zone6Oil | model.Zone7Oil | model.Zone8Oil
model.Slack = Set()
model.Hydro = Set()
model.NY_Imports_CT = Set()
model.NY_Imports_WCMA = Set()
model.NY_Imports_VT = Set()
model.HQ_Imports_VT = Set()
model.NB_Imports_ME = Set()

model.Ramping = model.Hydro | model.NY_Imports_CT | model.NY_Imports_WCMA |  model.NY_Imports_VT | model.HQ_Imports_VT | model.NB_Imports_ME
model.Generators = model.Zone1Generators | model.Zone2Generators | model.Zone3Generators | model.Zone4Generators | model.Zone5Generators | model.Zone6Generators | model.Zone7Generators | model.Zone8Generators | model.NY_Imports_CT | model.NY_Imports_WCMA  |  model.NY_Imports_VT | model.HQ_Imports_VT |  model.NB_Imports_ME

#
model.zones =Set()
model.sources = Set(within=model.zones)
model.sinks = Set(within=model.zones)

#########################################################
# These are the generators parameters from model input  #
#########################################################


#Generator Type
model.typ = Param(model.Generators)

#Zone parameters
model.zone = Param(model.Generators)

#Max Generating Capacity
model.netcap = Param(model.Generators)

#Min Generating Capacity
model.mincap = Param(model.Generators,mutable=True)

#Minimun up time
model.minu = Param(model.Generators)

#Minmun down time
model.mind = Param(model.Generators)

#Ramp rate
model.ramp  = Param(model.Generators)

#Start cost
model.st_cost = Param(model.Generators)

#Piecewice varible cost segments
model.seg1= Param(model.Generators)
model.seg2= Param(model.Generators)
model.seg3= Param(model.Generators)

#Variable O&M
model.var_om = Param(model.Generators)

#No load cost
model.no_load  = Param(model.Generators)

#Transmission Path parameters
model.hurdle = Param(model.sources, model.sinks)
model.limit = Param(model.sources, model.sinks)
#
###########################################################
### These are the detailed parameters for model runs      #
###########################################################
##
## Full range of time series information provided in .dat file (1 year)
model.SimHours = Param(within=PositiveIntegers)
model.SH_periods = RangeSet(1,model.SimHours)
model.SimDays = Param(within=PositiveIntegers)
model.SD_periods = RangeSet(1,model.SimDays)

# Operating horizon information 
model.HorizonHours = Param(within=PositiveIntegers)
model.HH_periods = RangeSet(0,model.HorizonHours)
model.hh_periods = RangeSet(1,model.HorizonHours)
model.HorizonDays = Param(within=PositiveIntegers)
model.hd_periods = RangeSet(1,model.HorizonDays)
model.h1_periods = RangeSet(1,24)
model.h2_periods = RangeSet(25,48)
model.ramp1_periods = RangeSet(2,24)
model.ramp2_periods = RangeSet(26,48)

#Demand over simulation period
model.SimDemand = Param(model.zones*model.SH_periods, within=NonNegativeReals)

#Must run generation over simulation period
model.SimMustRun = Param(model.zones*model.SH_periods, within=NonNegativeReals)

#Horizon demand
model.HorizonDemand = Param(model.zones*model.hh_periods,within=NonNegativeReals,mutable=True)

#Horizon must run generation
model.HorizonMustRun = Param(model.zones*model.hh_periods,within=NonNegativeReals,mutable=True)

#Reserve for the entire system
model.SimReserves = Param(model.SH_periods, within=NonNegativeReals)
model.HorizonReserves = Param(model.hh_periods, within=NonNegativeReals,mutable=True)

##Variable resources over simulation period
model.SimOffshoreWind = Param(model.zones, model.SH_periods, within=NonNegativeReals)
model.SimSolar = Param(model.zones, model.SH_periods, within=NonNegativeReals)
model.SimOnshoreWind = Param(model.zones, model.SH_periods, within=NonNegativeReals)

#Exports
model.SimCT_exports_NY = Param(model.SH_periods, within=NonNegativeReals)
model.SimWCMA_exports_NY = Param(model.SH_periods, within=NonNegativeReals)
model.SimVT_exports_NY = Param(model.SH_periods, within=NonNegativeReals)
model.SimVT_exports_HQ = Param(model.SH_periods, within=NonNegativeReals)
model.SimME_exports_NB = Param(model.SH_periods, within=NonNegativeReals)

model.HorizonCT_exports_NY = Param(model.hh_periods, within=NonNegativeReals,mutable=True)
model.HorizonWCMA_exports_NY = Param(model.hh_periods, within=NonNegativeReals,mutable=True)
model.HorizonVT_exports_NY = Param(model.hh_periods, within=NonNegativeReals,mutable=True)
model.HorizonVT_exports_HQ = Param(model.hh_periods, within=NonNegativeReals,mutable=True)
model.HorizonME_exports_NB = Param(model.hh_periods, within=NonNegativeReals,mutable=True)

#Natural gas prices over simulation period
model.SimGasPrice = Param(model.zones,model.SD_periods, within=NonNegativeReals)
model.GasPrice = Param(model.zones,within = NonNegativeReals, mutable=True,initialize=0)

#Oil prices over simulation period
model.SimOilPrice = Param(model.zones,model.SD_periods, within=NonNegativeReals)
model.OilPrice = Param(model.zones,within = NonNegativeReals, mutable=True,initialize=0)

#Daily path and hydro parameters
model.SimNY_imports_CT = Param(model.SD_periods, within=NonNegativeReals)
model.SimNY_imports_WCMA = Param(model.SD_periods, within=NonNegativeReals)
model.SimNY_imports_VT = Param(model.SD_periods, within=NonNegativeReals)
model.SimHQ_imports_VT = Param(model.SD_periods, within=NonNegativeReals)
model.SimNB_imports_ME = Param(model.SD_periods, within=NonNegativeReals)

model.SimME_hydro = Param(model.SD_periods, within=NonNegativeReals)
model.SimVT_hydro = Param(model.SD_periods, within=NonNegativeReals)
model.SimRI_hydro = Param(model.SD_periods, within=NonNegativeReals)
model.SimNH_hydro = Param(model.SD_periods, within=NonNegativeReals)
model.SimCT_hydro = Param(model.SD_periods, within=NonNegativeReals)
model.SimWCMA_hydro = Param(model.SD_periods, within=NonNegativeReals)
model.SimNEMA_hydro = Param(model.SD_periods, within=NonNegativeReals)

model.HorizonNY_imports_CT = Param(model.hd_periods, within=NonNegativeReals,mutable=True)
model.HorizonNY_imports_WCMA = Param(model.hd_periods, within=NonNegativeReals,mutable=True)
model.HorizonNY_imports_VT = Param(model.hd_periods, within=NonNegativeReals,mutable=True)
model.HorizonHQ_imports_VT = Param(model.hd_periods, within=NonNegativeReals,mutable=True)
model.HorizonNB_imports_ME = Param(model.hd_periods, within=NonNegativeReals,mutable=True)

model.HorizonME_hydro = Param(model.hd_periods, within=NonNegativeReals,mutable=True)
model.HorizonVT_hydro = Param(model.hd_periods, within=NonNegativeReals,mutable=True)
model.HorizonRI_hydro = Param(model.hd_periods, within=NonNegativeReals,mutable=True)
model.HorizonNH_hydro = Param(model.hd_periods, within=NonNegativeReals,mutable=True)
model.HorizonCT_hydro = Param(model.hd_periods, within=NonNegativeReals,mutable=True)
model.HorizonWCMA_hydro = Param(model.hd_periods, within=NonNegativeReals,mutable=True)
model.HorizonNEMA_hydro = Param(model.hd_periods, within=NonNegativeReals,mutable=True)

#Variable resources over horizon
model.HorizonOffshoreWind = Param(model.zones,model.hh_periods,within=NonNegativeReals,mutable=True)
model.HorizonSolar = Param(model.zones,model.hh_periods,within=NonNegativeReals,mutable=True)
model.HorizonOnshoreWind = Param(model.zones,model.hh_periods,within=NonNegativeReals,mutable=True)
model.HorizonHydro = Param(model.zones,model.hh_periods,within=NonNegativeReals,mutable=True)

##Initial conditions
model.ini_on = Param(model.Generators, within=NonNegativeReals, initialize=0,mutable=True) 
model.ini_mwh_1 = Param(model.Generators,initialize=0,mutable=True) #seg1
model.ini_mwh_2 = Param(model.Generators,initialize=0,mutable=True) #seg2
model.ini_mwh_3 = Param(model.Generators,initialize=0,mutable=True) #seg3

###########################################################
### Decision variables                                    #
###########################################################

##Amount of day-ahead energy generated by each thermal unit's 3 segments at each hour
model.mwh_1 = Var(model.Generators,model.HH_periods, within=NonNegativeReals,initialize=0)
model.mwh_2 = Var(model.Generators,model.HH_periods, within=NonNegativeReals,initialize=0)
model.mwh_3 = Var(model.Generators,model.HH_periods, within=NonNegativeReals,initialize=0)

#1 if unit is on in hour i
model.on = Var(model.Generators,model.HH_periods, within=NonNegativeReals, initialize=0)

#1 if unit is switching on in hour i
model.switch = Var(model.Generators,model.HH_periods, within=NonNegativeReals,initialize=0)

#Amount of spining reserce offered by each unit in each hour
model.srsv = Var(model.Generators,model.HH_periods, within=NonNegativeReals,initialize=0)

#Amount of non-sping reserve ovvered by each unit in each hour
model.nrsv = Var(model.Generators,model.HH_periods, within=NonNegativeReals,initialize=0)

#Renewable energy production
model.solar = Var(model.zones,model.HH_periods,within=NonNegativeReals)
model.onshorewind = Var(model.zones,model.HH_periods,within=NonNegativeReals)
model.offshorewind = Var(model.zones,model.HH_periods,within=NonNegativeReals)

#Power flows on each path
model.flow = Var(model.sources*model.sinks*model.HH_periods, within=NonNegativeReals)

#Must-run generation 
model.mustrun = Var(model.zones,model.HH_periods,within=NonNegativeReals)

#
#
####################################################################
##Objective function                                               #
##To minimize overall system cost while satistfy system constraints#
####################################################################
#
##
def SysCost(model):
    fixed = sum(model.no_load[j]*model.on[j,i] for i in model.hh_periods for j in model.Generators)
    coal1 = sum(model.mwh_1[j,i]*(model.seg1[j]*2.2 + model.var_om[j]) for i in model.hh_periods for j in model.Coal) 
    coal2 = sum(model.mwh_2[j,i]*(model.seg2[j]*2.2 + model.var_om[j]) for i in model.hh_periods for j in model.Coal) 
    coal3 = sum(model.mwh_3[j,i]*(model.seg3[j]*2.2 + model.var_om[j]) for i in model.hh_periods for j in model.Coal) 
    gas1_1 = sum(model.mwh_1[j,i]*(model.seg1[j]*model.GasPrice['CT'] + model.var_om[j]) for i in model.hh_periods for j in model.Zone1Gas) 
    gas2_1 = sum(model.mwh_2[j,i]*(model.seg2[j]*model.GasPrice['CT'] + model.var_om[j]) for i in model.hh_periods for j in model.Zone1Gas)  
    gas3_1 = sum(model.mwh_3[j,i]*(model.seg3[j]*model.GasPrice['CT'] + model.var_om[j]) for i in model.hh_periods for j in model.Zone1Gas)  
    gas1_2 = sum(model.mwh_1[j,i]*(model.seg1[j]*model.GasPrice['ME'] + model.var_om[j]) for i in model.hh_periods for j in model.Zone2Gas) 
    gas2_2 = sum(model.mwh_2[j,i]*(model.seg2[j]*model.GasPrice['ME'] + model.var_om[j]) for i in model.hh_periods for j in model.Zone2Gas)  
    gas3_2 = sum(model.mwh_3[j,i]*(model.seg3[j]*model.GasPrice['ME'] + model.var_om[j]) for i in model.hh_periods for j in model.Zone2Gas)  
    gas1_3 = sum(model.mwh_1[j,i]*(model.seg1[j]*model.GasPrice['NH'] + model.var_om[j]) for i in model.hh_periods for j in model.Zone3Gas) 
    gas2_3 = sum(model.mwh_2[j,i]*(model.seg2[j]*model.GasPrice['NH'] + model.var_om[j]) for i in model.hh_periods for j in model.Zone3Gas)  
    gas3_3 = sum(model.mwh_3[j,i]*(model.seg3[j]*model.GasPrice['NH'] + model.var_om[j]) for i in model.hh_periods for j in model.Zone3Gas)  
    gas1_4 = sum(model.mwh_1[j,i]*(model.seg1[j]*model.GasPrice['NEMA'] + model.var_om[j]) for i in model.hh_periods for j in model.Zone4Gas) 
    gas2_4 = sum(model.mwh_2[j,i]*(model.seg2[j]*model.GasPrice['NEMA'] + model.var_om[j]) for i in model.hh_periods for j in model.Zone4Gas)  
    gas3_4 = sum(model.mwh_3[j,i]*(model.seg3[j]*model.GasPrice['NEMA'] + model.var_om[j]) for i in model.hh_periods for j in model.Zone4Gas)    
    gas1_5 = sum(model.mwh_1[j,i]*(model.seg1[j]*model.GasPrice['RI'] + model.var_om[j]) for i in model.hh_periods for j in model.Zone5Gas) 
    gas2_5 = sum(model.mwh_2[j,i]*(model.seg2[j]*model.GasPrice['RI'] + model.var_om[j]) for i in model.hh_periods for j in model.Zone5Gas)  
    gas3_5 = sum(model.mwh_3[j,i]*(model.seg3[j]*model.GasPrice['RI'] + model.var_om[j]) for i in model.hh_periods for j in model.Zone5Gas)  
    gas1_6 = sum(model.mwh_1[j,i]*(model.seg1[j]*model.GasPrice['SEMA'] + model.var_om[j]) for i in model.hh_periods for j in model.Zone6Gas) 
    gas2_6 = sum(model.mwh_2[j,i]*(model.seg2[j]*model.GasPrice['SEMA'] + model.var_om[j]) for i in model.hh_periods for j in model.Zone6Gas)  
    gas3_6 = sum(model.mwh_3[j,i]*(model.seg3[j]*model.GasPrice['SEMA'] + model.var_om[j]) for i in model.hh_periods for j in model.Zone6Gas)  
    gas1_7 = sum(model.mwh_1[j,i]*(model.seg1[j]*model.GasPrice['VT'] + model.var_om[j]) for i in model.hh_periods for j in model.Zone7Gas) 
    gas2_7 = sum(model.mwh_2[j,i]*(model.seg2[j]*model.GasPrice['VT'] + model.var_om[j]) for i in model.hh_periods for j in model.Zone7Gas)  
    gas3_7 = sum(model.mwh_3[j,i]*(model.seg3[j]*model.GasPrice['VT'] + model.var_om[j]) for i in model.hh_periods for j in model.Zone7Gas)  
    gas1_8 = sum(model.mwh_1[j,i]*(model.seg1[j]*model.GasPrice['WCMA'] + model.var_om[j]) for i in model.hh_periods for j in model.Zone8Gas) 
    gas2_8 = sum(model.mwh_2[j,i]*(model.seg2[j]*model.GasPrice['WCMA'] + model.var_om[j]) for i in model.hh_periods for j in model.Zone8Gas)  
    gas3_8 = sum(model.mwh_3[j,i]*(model.seg3[j]*model.GasPrice['WCMA'] + model.var_om[j]) for i in model.hh_periods for j in model.Zone8Gas)
    NY_Imports_CT1 = sum(model.mwh_1[j,i]*4.3 for i in model.hh_periods for j in model.NY_Imports_CT)
    NY_Imports_CT2 = sum(model.mwh_2[j,i]*4.3 for i in model.hh_periods for j in model.NY_Imports_CT)
    NY_Imports_CT3 = sum(model.mwh_3[j,i]*4.3 for i in model.hh_periods for j in model.NY_Imports_CT) 
    NY_Imports_WCMA1 = sum(model.mwh_1[j,i]*4.3 for i in model.hh_periods for j in model.NY_Imports_WCMA)
    NY_Imports_WCMA2 = sum(model.mwh_2[j,i]*4.3 for i in model.hh_periods for j in model.NY_Imports_WCMA)
    NY_Imports_WCMA3 = sum(model.mwh_3[j,i]*4.3 for i in model.hh_periods for j in model.NY_Imports_WCMA)
    NY_Imports_VT1 = sum(model.mwh_1[j,i]*4.3 for i in model.hh_periods for j in model.NY_Imports_VT)
    NY_Imports_VT2 = sum(model.mwh_2[j,i]*4.3 for i in model.hh_periods for j in model.NY_Imports_VT)
    NY_Imports_VT3 = sum(model.mwh_3[j,i]*4.3 for i in model.hh_periods for j in model.NY_Imports_VT)
    HQ_Imports_VT1 = sum(model.mwh_1[j,i]*4.3 for i in model.hh_periods for j in model.HQ_Imports_VT)
    HQ_Imports_VT2 = sum(model.mwh_2[j,i]*4.3 for i in model.hh_periods for j in model.HQ_Imports_VT)
    HQ_Imports_VT3 = sum(model.mwh_3[j,i]*4.3 for i in model.hh_periods for j in model.HQ_Imports_VT)
    NB_Imports_ME1 = sum(model.mwh_1[j,i]*4.3 for i in model.hh_periods for j in model.NB_Imports_ME)
    NB_Imports_ME2 = sum(model.mwh_2[j,i]*4.3 for i in model.hh_periods for j in model.NB_Imports_ME)
    NB_Imports_ME3 = sum(model.mwh_3[j,i]*4.3 for i in model.hh_periods for j in model.NB_Imports_ME)
    Oil1_1 = sum(model.mwh_1[j,i]*(model.seg1[j]*model.OilPrice['CT'] + model.var_om[j]) for i in model.hh_periods for j in model.Zone1Oil) 
    Oil2_1 = sum(model.mwh_2[j,i]*(model.seg2[j]*model.OilPrice['CT'] + model.var_om[j]) for i in model.hh_periods for j in model.Zone1Oil)  
    Oil3_1 = sum(model.mwh_3[j,i]*(model.seg3[j]*model.OilPrice['CT'] + model.var_om[j]) for i in model.hh_periods for j in model.Zone1Oil)  
    Oil1_2 = sum(model.mwh_1[j,i]*(model.seg1[j]*model.OilPrice['ME'] + model.var_om[j]) for i in model.hh_periods for j in model.Zone2Oil) 
    Oil2_2 = sum(model.mwh_2[j,i]*(model.seg2[j]*model.OilPrice['ME'] + model.var_om[j]) for i in model.hh_periods for j in model.Zone2Oil)  
    Oil3_2 = sum(model.mwh_3[j,i]*(model.seg3[j]*model.OilPrice['ME'] + model.var_om[j]) for i in model.hh_periods for j in model.Zone2Oil)  
    Oil1_3 = sum(model.mwh_1[j,i]*(model.seg1[j]*model.OilPrice['NH'] + model.var_om[j]) for i in model.hh_periods for j in model.Zone3Oil) 
    Oil2_3 = sum(model.mwh_2[j,i]*(model.seg2[j]*model.OilPrice['NH'] + model.var_om[j]) for i in model.hh_periods for j in model.Zone3Oil)  
    Oil3_3 = sum(model.mwh_3[j,i]*(model.seg3[j]*model.OilPrice['NH'] + model.var_om[j]) for i in model.hh_periods for j in model.Zone3Oil)  
    Oil1_4 = sum(model.mwh_1[j,i]*(model.seg1[j]*model.OilPrice['NEMA'] + model.var_om[j]) for i in model.hh_periods for j in model.Zone4Oil) 
    Oil2_4 = sum(model.mwh_2[j,i]*(model.seg2[j]*model.OilPrice['NEMA'] + model.var_om[j]) for i in model.hh_periods for j in model.Zone4Oil)  
    Oil3_4 = sum(model.mwh_3[j,i]*(model.seg3[j]*model.OilPrice['NEMA'] + model.var_om[j]) for i in model.hh_periods for j in model.Zone4Oil)    
    Oil1_5 = sum(model.mwh_1[j,i]*(model.seg1[j]*model.OilPrice['RI'] + model.var_om[j]) for i in model.hh_periods for j in model.Zone5Oil) 
    Oil2_5 = sum(model.mwh_2[j,i]*(model.seg2[j]*model.OilPrice['RI'] + model.var_om[j]) for i in model.hh_periods for j in model.Zone5Oil)  
    Oil3_5 = sum(model.mwh_3[j,i]*(model.seg3[j]*model.OilPrice['RI'] + model.var_om[j]) for i in model.hh_periods for j in model.Zone5Oil)  
    Oil1_6 = sum(model.mwh_1[j,i]*(model.seg1[j]*model.OilPrice['SEMA'] + model.var_om[j]) for i in model.hh_periods for j in model.Zone6Oil) 
    Oil2_6 = sum(model.mwh_2[j,i]*(model.seg2[j]*model.OilPrice['SEMA'] + model.var_om[j]) for i in model.hh_periods for j in model.Zone6Oil)  
    Oil3_6 = sum(model.mwh_3[j,i]*(model.seg3[j]*model.OilPrice['SEMA'] + model.var_om[j]) for i in model.hh_periods for j in model.Zone6Oil)  
    Oil1_7 = sum(model.mwh_1[j,i]*(model.seg1[j]*model.OilPrice['VT'] + model.var_om[j]) for i in model.hh_periods for j in model.Zone7Oil) 
    Oil2_7 = sum(model.mwh_2[j,i]*(model.seg2[j]*model.OilPrice['VT'] + model.var_om[j]) for i in model.hh_periods for j in model.Zone7Oil)  
    Oil3_7 = sum(model.mwh_3[j,i]*(model.seg3[j]*model.OilPrice['VT'] + model.var_om[j]) for i in model.hh_periods for j in model.Zone7Oil)  
    Oil1_8 = sum(model.mwh_1[j,i]*(model.seg1[j]*model.OilPrice['WCMA'] + model.var_om[j]) for i in model.hh_periods for j in model.Zone8Oil) 
    Oil2_8 = sum(model.mwh_2[j,i]*(model.seg2[j]*model.OilPrice['WCMA'] + model.var_om[j]) for i in model.hh_periods for j in model.Zone8Oil)  
    Oil3_8 = sum(model.mwh_3[j,i]*(model.seg3[j]*model.OilPrice['WCMA'] + model.var_om[j]) for i in model.hh_periods for j in model.Zone8Oil)  
    slack1 = sum(model.mwh_1[j,i]*model.seg1[j]*10000 for i in model.hh_periods for j in model.Slack)
    slack2 = sum(model.mwh_2[j,i]*model.seg2[j]*10000 for i in model.hh_periods for j in model.Slack)
    slack3 = sum(model.mwh_3[j,i]*model.seg3[j]*10000 for i in model.hh_periods for j in model.Slack)
    starts = sum(model.st_cost[j]*model.switch[j,i] for i in model.hh_periods for j in model.Generators) 
    exchange = sum(model.flow[s,k,i]*model.hurdle[s,k] for s in model.sources for k in model.sinks for i in model.hh_periods)
    hydro1 = sum(model.mwh_1[j,i]*0.01 for i in model.hh_periods for j in model.Hydro)
    hydro2 = sum(model.mwh_2[j,i]*0.01 for i in model.hh_periods for j in model.Hydro)
    hydro3 = sum(model.mwh_3[j,i]*0.01 for i in model.hh_periods for j in model.Hydro)
    onshore_wind_cost = sum(model.onshorewind[j,i]*0.01 for i in model.hh_periods for j in model.zones)
    offshore_wind_cost = sum(model.offshorewind[j,i]*0.01 for i in model.hh_periods for j in model.zones)
    solar_cost = sum(model.solar[j,i]*0.01 for i in model.hh_periods for j in model.zones)
    mustrun_cost = sum(model.mustrun[j,i]*0.01 for i in model.hh_periods for j in model.zones)
    return fixed + coal1 + coal2 + coal3 + gas1_1 + gas1_2 + gas1_3 + gas1_4 + gas1_5 + gas1_6 + gas1_7 + gas1_8 + gas2_1 + gas2_2 + gas2_3 + gas2_4 + gas2_5 + gas2_6 + gas2_7 + gas2_8 + gas3_1 + gas3_2 + gas3_3 + gas3_4 + gas3_5 + gas3_6 + gas3_7 + gas3_8 + Oil1_1 + Oil1_2 + Oil1_3 + Oil1_4 + Oil1_5 + Oil1_6 + Oil1_7 + Oil1_8 + Oil2_1 + Oil2_2 + Oil2_3 + Oil2_4 + Oil2_5 + Oil2_6 + Oil2_7 + Oil2_8 + Oil3_1 + Oil3_2 + Oil3_3 + Oil3_4 + Oil3_5 + Oil3_6 + Oil3_7 + Oil3_8 + slack1 + slack2 + slack3 + starts + exchange + NY_Imports_CT1 + NY_Imports_CT2 + NY_Imports_CT3 + NY_Imports_WCMA1 + NY_Imports_WCMA2 + NY_Imports_WCMA3 + NY_Imports_VT1 + NY_Imports_VT2 + NY_Imports_VT3 + HQ_Imports_VT1 + HQ_Imports_VT2 + HQ_Imports_VT3 + NB_Imports_ME1 + NB_Imports_ME2 + NB_Imports_ME3 + hydro1 + hydro2 + hydro3 + onshore_wind_cost + offshore_wind_cost + solar_cost + mustrun_cost
model.SystemCost = Objective(rule=SysCost, sense=minimize)
    
   
####################################################################
#   Constraints                                                    #
####################################################################
   
#WECC Constraint 25% of internal demand must be generated locally
def WECC1(model,i):
    s1 = sum(model.mwh_1[j,i] for j in model.Zone1Generators) 
    s2 = sum(model.mwh_2[j,i] for j in model.Zone1Generators) 
    s3 = sum(model.mwh_3[j,i] for j in model.Zone1Generators) 
    renew = model.offshorewind['CT',i] + model.solar['CT',i] + model.onshorewind['CT',i]
    must_run = model.mustrun['CT',i]
    return s1 + s2 + s3 + renew + must_run >=  0.25*model.HorizonDemand['CT',i]
model.Local1= Constraint(model.hh_periods,rule=WECC1)
##
def WECC2(model,i):
    s1 = sum(model.mwh_1[j,i] for j in model.Zone2Generators) 
    s2 = sum(model.mwh_2[j,i] for j in model.Zone2Generators) 
    s3 = sum(model.mwh_3[j,i] for j in model.Zone2Generators) 
    renew = model.offshorewind['ME',i] + model.solar['ME',i] + model.onshorewind['ME',i]
    must_run = model.mustrun['ME',i]
    return s1 + s2 + s3 + renew + must_run >=  0.25*model.HorizonDemand['ME',i]
model.Local2= Constraint(model.hh_periods,rule=WECC2)
###
def WECC3(model,i):
    s1 = sum(model.mwh_1[j,i] for j in model.Zone3Generators) 
    s2 = sum(model.mwh_2[j,i] for j in model.Zone3Generators) 
    s3 = sum(model.mwh_3[j,i] for j in model.Zone3Generators) 
    renew = model.offshorewind['NH',i] + model.solar['NH',i] + model.onshorewind['NH',i]
    must_run = model.mustrun['NH',i]
    return s1 + s2 + s3 + renew + must_run >=  0.25*model.HorizonDemand['NH',i]
model.Local3= Constraint(model.hh_periods,rule=WECC3)
#
def WECC4(model,i):
    s1 = sum(model.mwh_1[j,i] for j in model.Zone4Generators) 
    s2 = sum(model.mwh_2[j,i] for j in model.Zone4Generators) 
    s3 = sum(model.mwh_3[j,i] for j in model.Zone4Generators) 
    renew = model.offshorewind['NEMA',i] + model.solar['NEMA',i] + model.onshorewind['NEMA',i]
    must_run = model.mustrun['NEMA',i]
    return s1 + s2 + s3 + renew + must_run >=  0.25*model.HorizonDemand['NEMA',i]
model.Local4= Constraint(model.hh_periods,rule=WECC4)

def WECC5(model,i):
    s1 = sum(model.mwh_1[j,i] for j in model.Zone5Generators) 
    s2 = sum(model.mwh_2[j,i] for j in model.Zone5Generators) 
    s3 = sum(model.mwh_3[j,i] for j in model.Zone5Generators) 
    renew = model.offshorewind['RI',i] + model.solar['RI',i] + model.onshorewind['RI',i]
    must_run = model.mustrun['RI',i]
    return s1 + s2 + s3 + renew + must_run >=  0.25*model.HorizonDemand['RI',i]
model.Local5= Constraint(model.hh_periods,rule=WECC5)

def WECC6(model,i):
    s1 = sum(model.mwh_1[j,i] for j in model.Zone6Generators) 
    s2 = sum(model.mwh_2[j,i] for j in model.Zone6Generators) 
    s3 = sum(model.mwh_3[j,i] for j in model.Zone6Generators) 
    renew = model.offshorewind['SEMA',i] + model.solar['SEMA',i] + model.onshorewind['SEMA',i]
    must_run = model.mustrun['SEMA',i]
    return s1 + s2 + s3 + renew + must_run >=  0.25*model.HorizonDemand['SEMA',i]
model.Local6= Constraint(model.hh_periods,rule=WECC6)

def WECC7(model,i):
    s1 = sum(model.mwh_1[j,i] for j in model.Zone7Generators) 
    s2 = sum(model.mwh_2[j,i] for j in model.Zone7Generators) 
    s3 = sum(model.mwh_3[j,i] for j in model.Zone7Generators) 
    renew = model.offshorewind['VT',i] + model.solar['VT',i] + model.onshorewind['VT',i]
    must_run = model.mustrun['VT',i]
    return s1 + s2 + s3 + renew + must_run >=  0.25*model.HorizonDemand['VT',i]
model.Local7= Constraint(model.hh_periods,rule=WECC7)

def WECC8(model,i):
    s1 = sum(model.mwh_1[j,i] for j in model.Zone8Generators) 
    s2 = sum(model.mwh_2[j,i] for j in model.Zone8Generators) 
    s3 = sum(model.mwh_3[j,i] for j in model.Zone8Generators) 
    renew = model.offshorewind['WCMA',i] + model.solar['WCMA',i] + model.onshorewind['WCMA',i]
    must_run = model.mustrun['WCMA',i]
    return s1 + s2 + s3 + renew + must_run >=  0.25*model.HorizonDemand['WCMA',i]
model.Local8= Constraint(model.hh_periods,rule=WECC8)


###Power Balance 
def Zone1_Balance(model,i):
    s1 = sum(model.mwh_1[j,i] for j in model.Zone1Generators)
    s2 = sum(model.mwh_2[j,i] for j in model.Zone1Generators)  
    s3 = sum(model.mwh_3[j,i] for j in model.Zone1Generators)  
    other = model.offshorewind['CT',i] + model.mustrun['CT',i] + model.solar['CT',i] + model.onshorewind['CT',i]
    imports = sum(model.flow[s,'CT',i] for s in model.sources) + model.mwh_1['NYCTI',i] + model.mwh_2['NYCTI',i] + model.mwh_3['NYCTI',i] 
    exports = sum(model.flow['CT',k,i] for k in model.sinks) + model.HorizonCT_exports_NY[i]        
    return s1 + s2 + s3 + other + imports - exports >= model.HorizonDemand['CT',i]
model.Bal1Constraint= Constraint(model.hh_periods,rule=Zone1_Balance)

def Zone2_Balance(model,i):
    s1 = sum(model.mwh_1[j,i] for j in model.Zone2Generators)
    s2 = sum(model.mwh_2[j,i] for j in model.Zone2Generators)  
    s3 = sum(model.mwh_3[j,i] for j in model.Zone2Generators)  
    other = model.offshorewind['ME',i] + model.mustrun['ME',i] + model.solar['ME',i] + model.onshorewind['ME',i]
    imports = sum(model.flow[s,'ME',i] for s in model.sources) + model.mwh_1['NBMEI',i] + model.mwh_2['NBMEI',i] + model.mwh_3['NBMEI',i] 
    exports = sum(model.flow['ME',k,i] for k in model.sinks) + model.HorizonME_exports_NB[i]        
    return s1 + s2 + s3 + other + imports - exports >= model.HorizonDemand['ME',i]
model.Bal2Constraint= Constraint(model.hh_periods,rule=Zone2_Balance)
#
def Zone3_Balance(model,i):
    s1 = sum(model.mwh_1[j,i] for j in model.Zone3Generators)
    s2 = sum(model.mwh_2[j,i] for j in model.Zone3Generators)  
    s3 = sum(model.mwh_3[j,i] for j in model.Zone3Generators)  
    other = model.offshorewind['NH',i] + model.mustrun['NH',i] + model.solar['NH',i] + model.onshorewind['NH',i]
    imports = sum(model.flow[s,'NH',i] for s in model.sources) 
    exports = sum(model.flow['NH',k,i] for k in model.sinks)        
    return s1 + s2 + s3 + other + imports - exports >= model.HorizonDemand['NH',i]
model.Bal3Constraint= Constraint(model.hh_periods,rule=Zone3_Balance)
# 
def Zone4_Balance(model,i):
    s1 = sum(model.mwh_1[j,i] for j in model.Zone4Generators)
    s2 = sum(model.mwh_2[j,i] for j in model.Zone4Generators)  
    s3 = sum(model.mwh_3[j,i] for j in model.Zone4Generators)  
    other = model.offshorewind['NEMA',i] + model.mustrun['NEMA',i] + model.solar['NEMA',i] + model.onshorewind['NEMA',i]
    imports = sum(model.flow[s,'NEMA',i] for s in model.sources) 
    exports = sum(model.flow['NEMA',k,i] for k in model.sinks)        
    return s1 + s2 + s3 + other + imports - exports >= model.HorizonDemand['NEMA',i]
model.Bal4Constraint= Constraint(model.hh_periods,rule=Zone4_Balance)

def Zone5_Balance(model,i):
    s1 = sum(model.mwh_1[j,i] for j in model.Zone5Generators)
    s2 = sum(model.mwh_2[j,i] for j in model.Zone5Generators)  
    s3 = sum(model.mwh_3[j,i] for j in model.Zone5Generators)  
    other = model.offshorewind['RI',i] + model.mustrun['RI',i] + model.solar['RI',i] + model.onshorewind['RI',i]
    imports = sum(model.flow[s,'RI',i] for s in model.sources) 
    exports = sum(model.flow['RI',k,i] for k in model.sinks)        
    return s1 + s2 + s3 + other + imports - exports >= model.HorizonDemand['RI',i]
model.Bal5Constraint= Constraint(model.hh_periods,rule=Zone5_Balance)

def Zone6_Balance(model,i):
    s1 = sum(model.mwh_1[j,i] for j in model.Zone6Generators)
    s2 = sum(model.mwh_2[j,i] for j in model.Zone6Generators)  
    s3 = sum(model.mwh_3[j,i] for j in model.Zone6Generators)  
    other = model.offshorewind['SEMA',i] + model.mustrun['SEMA',i] + model.solar['SEMA',i] + model.onshorewind['SEMA',i]
    imports = sum(model.flow[s,'SEMA',i] for s in model.sources) 
    exports = sum(model.flow['SEMA',k,i] for k in model.sinks)        
    return s1 + s2 + s3 + other + imports - exports >= model.HorizonDemand['SEMA',i]
model.Bal6Constraint= Constraint(model.hh_periods,rule=Zone6_Balance)

def Zone7_Balance(model,i):
    s1 = sum(model.mwh_1[j,i] for j in model.Zone7Generators)
    s2 = sum(model.mwh_2[j,i] for j in model.Zone7Generators)  
    s3 = sum(model.mwh_3[j,i] for j in model.Zone7Generators)  
    other = model.offshorewind['VT',i] + model.mustrun['VT',i] + model.solar['VT',i] + model.onshorewind['VT',i]
    imports = sum(model.flow[s,'VT',i] for s in model.sources) + model.mwh_1['NYVTI',i] + model.mwh_2['NYVTI',i] + model.mwh_3['NYVTI',i] + model.mwh_1['HQVTI',i] + model.mwh_2['HQVTI',i] + model.mwh_3['HQVTI',i] 
    exports = sum(model.flow['VT',k,i] for k in model.sinks) + model.HorizonVT_exports_NY[i]        
    return s1 + s2 + s3 + other + imports - exports >= model.HorizonDemand['VT',i]
model.Bal7Constraint= Constraint(model.hh_periods,rule=Zone7_Balance)

def Zone8_Balance(model,i):
    s1 = sum(model.mwh_1[j,i] for j in model.Zone8Generators)
    s2 = sum(model.mwh_2[j,i] for j in model.Zone8Generators)  
    s3 = sum(model.mwh_3[j,i] for j in model.Zone8Generators)  
    other = model.offshorewind['WCMA',i] + model.mustrun['WCMA',i] + model.solar['WCMA',i] + model.onshorewind['WCMA',i]
    imports = sum(model.flow[s,'WCMA',i] for s in model.sources) + model.mwh_1['NYWCMAI',i] + model.mwh_2['NYWCMAI',i] + model.mwh_3['NYWCMAI',i] 
    exports = sum(model.flow['WCMA',k,i] for k in model.sinks) + model.HorizonWCMA_exports_NY[i]        
    return s1 + s2 + s3 + other + imports - exports >= model.HorizonDemand['WCMA',i]
model.Bal8Constraint= Constraint(model.hh_periods,rule=Zone8_Balance)

#Max capacity constraints on variable resources 
def SolarC(model,z,i):
    return model.solar[z,i] <= model.HorizonSolar[z,i]
model.SolarConstraint= Constraint(model.zones,model.hh_periods,rule=SolarC)

def OnshoreWindC(model,z,i):
    return model.onshorewind[z,i] <= model.HorizonOnshoreWind[z,i]
model.OnshoreWindConstraint= Constraint(model.zones,model.hh_periods,rule=OnshoreWindC)

def OffshoreWindC(model,z,i):
    return model.offshorewind[z,i] <= model.HorizonOffshoreWind[z,i]
model.OffshoreWindConstraint= Constraint(model.zones,model.hh_periods,rule=OffshoreWindC)

#Max capacity constraint for must-run resources
def MustRunC(model,z,i):
    return model.mustrun[z,i] <= model.HorizonMustRun[z,i]
model.MustRunConstraint= Constraint(model.zones,model.hh_periods,rule=MustRunC)

# Daily production limits on dispatchable hydropower
def HydroC1(model,i):
    m1 = sum(model.mwh_1['CT_hydro',i] for i in model.h1_periods)
    m2 = sum(model.mwh_2['CT_hydro',i] for i in model.h1_periods)
    m3 = sum(model.mwh_3['CT_hydro',i] for i in model.h1_periods)
    return m1 + m2 + m3 <= model.HorizonCT_hydro[1]
model.HydroConstraint1= Constraint(model.h1_periods,rule=HydroC1)

def HydroC2(model,i):
    m1 = sum(model.mwh_1['CT_hydro',i] for i in model.h2_periods)
    m2 = sum(model.mwh_2['CT_hydro',i] for i in model.h2_periods)
    m3 = sum(model.mwh_3['CT_hydro',i] for i in model.h2_periods)
    return m1 + m2 + m3 <= model.HorizonCT_hydro[2]
model.HydroConstraint2= Constraint(model.h2_periods,rule=HydroC2)

def HydroC3(model,i):
    m1 = sum(model.mwh_1['ME_hydro',i] for i in model.h1_periods)
    m2 = sum(model.mwh_2['ME_hydro',i] for i in model.h1_periods)
    m3 = sum(model.mwh_3['ME_hydro',i] for i in model.h1_periods)
    return m1 + m2 + m3 <= model.HorizonME_hydro[1]
model.HydroConstraint3= Constraint(model.h1_periods,rule=HydroC3)

def HydroC4(model,i):
    m1 = sum(model.mwh_1['ME_hydro',i] for i in model.h2_periods)
    m2 = sum(model.mwh_2['ME_hydro',i] for i in model.h2_periods)
    m3 = sum(model.mwh_3['ME_hydro',i] for i in model.h2_periods)
    return m1 + m2 + m3 <= model.HorizonME_hydro[2]
model.HydroConstraint4= Constraint(model.h2_periods,rule=HydroC4)

def HydroC5(model,i):
    m1 = sum(model.mwh_1['NH_hydro',i] for i in model.h1_periods)
    m2 = sum(model.mwh_2['NH_hydro',i] for i in model.h1_periods)
    m3 = sum(model.mwh_3['NH_hydro',i] for i in model.h1_periods)
    return m1 + m2 + m3 <= model.HorizonNH_hydro[1]
model.HydroConstraint5= Constraint(model.h1_periods,rule=HydroC5)

def HydroC6(model,i):
    m1 = sum(model.mwh_1['NH_hydro',i] for i in model.h2_periods)
    m2 = sum(model.mwh_2['NH_hydro',i] for i in model.h2_periods)
    m3 = sum(model.mwh_3['NH_hydro',i] for i in model.h2_periods)
    return m1 + m2 + m3 <= model.HorizonNH_hydro[2]
model.HydroConstraint6= Constraint(model.h2_periods,rule=HydroC6)

def HydroC7(model,i):
    m1 = sum(model.mwh_1['NEMA_hydro',i] for i in model.h1_periods)
    m2 = sum(model.mwh_2['NEMA_hydro',i] for i in model.h1_periods)
    m3 = sum(model.mwh_3['NEMA_hydro',i] for i in model.h1_periods)
    return m1 + m2 + m3 <= model.HorizonNEMA_hydro[1]
model.HydroConstraint7= Constraint(model.h1_periods,rule=HydroC7)

def HydroC8(model,i):
    m1 = sum(model.mwh_1['NEMA_hydro',i] for i in model.h2_periods)
    m2 = sum(model.mwh_2['NEMA_hydro',i] for i in model.h2_periods)
    m3 = sum(model.mwh_3['NEMA_hydro',i] for i in model.h2_periods)
    return m1 + m2 + m3 <= model.HorizonNEMA_hydro[2]
model.HydroConstraint8= Constraint(model.h2_periods,rule=HydroC8)

def HydroC9(model,i):
    m1 = sum(model.mwh_1['RI_hydro',i] for i in model.h1_periods)
    m2 = sum(model.mwh_2['RI_hydro',i] for i in model.h1_periods)
    m3 = sum(model.mwh_3['RI_hydro',i] for i in model.h1_periods)
    return m1 + m2 + m3 <= model.HorizonRI_hydro[1]
model.HydroConstraint9= Constraint(model.h1_periods,rule=HydroC9)

def HydroC10(model,i):
    m1 = sum(model.mwh_1['RI_hydro',i] for i in model.h2_periods)
    m2 = sum(model.mwh_2['RI_hydro',i] for i in model.h2_periods)
    m3 = sum(model.mwh_3['RI_hydro',i] for i in model.h2_periods)
    return m1 + m2 + m3 <= model.HorizonRI_hydro[2]
model.HydroConstraint10= Constraint(model.h2_periods,rule=HydroC10)

def HydroC11(model,i):
    m1 = sum(model.mwh_1['VT_hydro',i] for i in model.h1_periods)
    m2 = sum(model.mwh_2['VT_hydro',i] for i in model.h1_periods)
    m3 = sum(model.mwh_3['VT_hydro',i] for i in model.h1_periods)
    return m1 + m2 + m3 <= model.HorizonVT_hydro[1]
model.HydroConstraint11= Constraint(model.h1_periods,rule=HydroC11)

def HydroC12(model,i):
    m1 = sum(model.mwh_1['VT_hydro',i] for i in model.h2_periods)
    m2 = sum(model.mwh_2['VT_hydro',i] for i in model.h2_periods)
    m3 = sum(model.mwh_3['VT_hydro',i] for i in model.h2_periods)
    return m1 + m2 + m3 <= model.HorizonVT_hydro[2]
model.HydroConstraint12= Constraint(model.h2_periods,rule=HydroC12)

def HydroC13(model,i):
    m1 = sum(model.mwh_1['WCMA_hydro',i] for i in model.h1_periods)
    m2 = sum(model.mwh_2['WCMA_hydro',i] for i in model.h1_periods)
    m3 = sum(model.mwh_3['WCMA_hydro',i] for i in model.h1_periods)
    return m1 + m2 + m3 <= model.HorizonWCMA_hydro[1]
model.HydroConstraint13= Constraint(model.h1_periods,rule=HydroC13)

def HydroC14(model,i):
    m1 = sum(model.mwh_1['WCMA_hydro',i] for i in model.h2_periods)
    m2 = sum(model.mwh_2['WCMA_hydro',i] for i in model.h2_periods)
    m3 = sum(model.mwh_3['WCMA_hydro',i] for i in model.h2_periods)
    return m1 + m2 + m3 <= model.HorizonWCMA_hydro[2]
model.HydroConstraint14= Constraint(model.h2_periods,rule=HydroC14)


# Daily production limits on imported power
def ImportsC1(model,i):
    m1 = sum(model.mwh_1['NYCTI',i] for i in model.h1_periods)
    m2 = sum(model.mwh_2['NYCTI',i] for i in model.h1_periods)
    m3 = sum(model.mwh_3['NYCTI',i] for i in model.h1_periods)
    return m1 + m2 + m3 <= model.HorizonNY_imports_CT[1]
model.ImportsConstraint1= Constraint(model.h1_periods,rule=ImportsC1)

def ImportsC2(model,i):
    m1 = sum(model.mwh_1['NYCTI',i] for i in model.h2_periods)
    m2 = sum(model.mwh_2['NYCTI',i] for i in model.h2_periods)
    m3 = sum(model.mwh_3['NYCTI',i] for i in model.h2_periods)
    return m1 + m2 + m3 <= model.HorizonNY_imports_CT[2]
model.ImportsConstraint2= Constraint(model.h2_periods,rule=ImportsC2)

def ImportsC3(model,i):
    m1 = sum(model.mwh_1['NBMEI',i] for i in model.h1_periods)
    m2 = sum(model.mwh_2['NBMEI',i] for i in model.h1_periods)
    m3 = sum(model.mwh_3['NBMEI',i] for i in model.h1_periods)
    return m1 + m2 + m3 <= model.HorizonNB_imports_ME[1]
model.ImportsConstraint3= Constraint(model.h1_periods,rule=ImportsC3)

def ImportsC4(model,i):
    m1 = sum(model.mwh_1['NBMEI',i] for i in model.h2_periods)
    m2 = sum(model.mwh_2['NBMEI',i] for i in model.h2_periods)
    m3 = sum(model.mwh_3['NBMEI',i] for i in model.h2_periods)
    return m1 + m2 + m3 <= model.HorizonNB_imports_ME[2]
model.ImportsConstraint4= Constraint(model.h2_periods,rule=ImportsC4)

def ImportsC5(model,i):
    m1 = sum(model.mwh_1['NYVTI',i] for i in model.h1_periods)
    m2 = sum(model.mwh_2['NYVTI',i] for i in model.h1_periods)
    m3 = sum(model.mwh_3['NYVTI',i] for i in model.h1_periods)
    return m1 + m2 + m3 <= model.HorizonNY_imports_VT[1]
model.ImportsConstraint5= Constraint(model.h1_periods,rule=ImportsC5)

def ImportsC6(model,i):
    m1 = sum(model.mwh_1['NYVTI',i] for i in model.h2_periods)
    m2 = sum(model.mwh_2['NYVTI',i] for i in model.h2_periods)
    m3 = sum(model.mwh_3['NYVTI',i] for i in model.h2_periods)
    return m1 + m2 + m3 <= model.HorizonNY_imports_VT[2]
model.ImportsConstraint6= Constraint(model.h2_periods,rule=ImportsC6)

def ImportsC7(model,i):
    m1 = sum(model.mwh_1['HQVTI',i] for i in model.h1_periods)
    m2 = sum(model.mwh_2['HQVTI',i] for i in model.h1_periods)
    m3 = sum(model.mwh_3['HQVTI',i] for i in model.h1_periods)
    return m1 + m2 + m3 <= model.HorizonHQ_imports_VT[1]
model.ImportsConstraint7= Constraint(model.h1_periods,rule=ImportsC7)

def ImportsC8(model,i):
    m1 = sum(model.mwh_1['HQVTI',i] for i in model.h2_periods)
    m2 = sum(model.mwh_2['HQVTI',i] for i in model.h2_periods)
    m3 = sum(model.mwh_3['HQVTI',i] for i in model.h2_periods)
    return m1 + m2 + m3 <= model.HorizonHQ_imports_VT[2]
model.ImportsConstraint8= Constraint(model.h2_periods,rule=ImportsC8)

def ImportsC9(model,i):
    m1 = sum(model.mwh_1['NYWCMAI',i] for i in model.h1_periods)
    m2 = sum(model.mwh_2['NYWCMAI',i] for i in model.h1_periods)
    m3 = sum(model.mwh_3['NYWCMAI',i] for i in model.h1_periods)
    return m1 + m2 + m3 <= model.HorizonNY_imports_WCMA[1]
model.ImportsConstraint9= Constraint(model.h1_periods,rule=ImportsC9)

def ImportsC10(model,i):
    m1 = sum(model.mwh_1['NYWCMAI',i] for i in model.h2_periods)
    m2 = sum(model.mwh_2['NYWCMAI',i] for i in model.h2_periods)
    m3 = sum(model.mwh_3['NYWCMAI',i] for i in model.h2_periods)
    return m1 + m2 + m3 <= model.HorizonNY_imports_WCMA[2]
model.ImportsConstraint10= Constraint(model.h2_periods,rule=ImportsC10)

##max capacity constraints on flows, etc. 
def FlowC(model,s,k,i):
    return model.flow[s,k,i] <= model.limit[s,k]        
model.FlowConstraint= Constraint(model.sources,model.sinks,model.hh_periods,rule=FlowC)


#Max Capacity Constraint
def MaxC(model,j,i):
    return model.mwh_1[j,i] + model.mwh_2[j,i] + model.mwh_3[j,i] <= model.on[j,i] * model.netcap[j]
model.MaxCap= Constraint(model.Generators,model.hh_periods,rule=MaxC)


##Min Capacity Constraint
def MinC1(model,j,i):
    return model.mwh_1[j,i] + model.mwh_2[j,i] + model.mwh_3[j,i] >= model.on[j,i] * model.mincap[j]
model.MinCap1= Constraint(model.Generators,model.hh_periods,rule=MinC1)



##System Reserve Requirement (excludes pumped storage)
def SysReserve(model,i):
    return sum(model.nrsv[j,i] for j in model.Coal) + sum(model.nrsv[j,i] for j in model.Gas) + sum(model.nrsv[j,i] for j in model.Oil) + sum(model.nrsv[j,i] for j in model.Slack) >= model.HorizonReserves[i]
model.SystemReserve = Constraint(model.hh_periods,rule=SysReserve)
##
##
###Segment capacity requirements
def Seg1(model,j,i):
    return model.mwh_1[j,i] <= .6*model.netcap[j]
model.Segment1 = Constraint(model.Generators,model.hh_periods,rule=Seg1)
#
def Seg2(model,j,i):
    return model.mwh_2[j,i] <= .2*model.netcap[j]
model.Segment2 = Constraint(model.Generators,model.hh_periods,rule=Seg2)

def Seg3(model,j,i):
    return model.mwh_3[j,i] <= .2*model.netcap[j]
model.Segment3 = Constraint(model.Generators,model.hh_periods,rule=Seg3)
##
#
##Zero Sum Constraint
def ZeroSum(model,j,i):
    return model.mwh_1[j,i] + model.mwh_2[j,i] + model.mwh_3[j,i] + model.srsv[j,i] + model.nrsv[j,i] <= model.netcap[j]
model.ZeroSumConstraint=Constraint(model.Generators,model.hh_periods,rule=ZeroSum)
#
#
###Switch is 1 if unit is turned on in current period
#def SwitchCon(model,j,i):
#    return model.switch[j,i] >= 1 - model.on[j,i-1] - (1 - model.on[j,i])
#model.SwitchConstraint = Constraint(model.Generators,model.hh_periods,rule = SwitchCon)
#
#
###Min Up time
#def MinUp(model,j,i,k):
#    if i > 0 and k > i and k < min(i+model.minu[j]-1,model.HorizonHours):
#        return model.on[j,i] - model.on[j,i-1] <= model.on[j,k]
#    else: 
#        return Constraint.Skip
#model.MinimumUp = Constraint(model.Generators,model.HH_periods,model.HH_periods,rule=MinUp)
##
###Min Down time
#def MinDown(model,j,i,k):
#   if i > 0 and k > i and k < min(i+model.mind[j]-1,model.HorizonHours):
#       return model.on[j,i-1] - model.on[j,i] <= 1 - model.on[j,k]
#   else:
#       return Constraint.Skip
#model.MinimumDown = Constraint(model.Generators,model.HH_periods,model.HH_periods,rule=MinDown)

#Ramp Rate Constraints
def Ramp1(model,j,i):
    a = model.mwh_1[j,i] + model.mwh_2[j,i] + model.mwh_3[j,i]
    b = model.mwh_1[j,i-1] + model.mwh_2[j,i-1] + model.mwh_3[j,i-1]
    return a - b <= model.ramp[j] 
model.RampCon1 = Constraint(model.Ramping,model.ramp1_periods,rule=Ramp1)

def Ramp2(model,j,i):
    a = model.mwh_1[j,i] + model.mwh_2[j,i] + model.mwh_3[j,i]
    b = model.mwh_1[j,i-1] + model.mwh_2[j,i-1] + model.mwh_3[j,i-1]
    return b - a <= model.ramp[j] 
model.RampCon2 = Constraint(model.Ramping,model.ramp2_periods,rule=Ramp2)





