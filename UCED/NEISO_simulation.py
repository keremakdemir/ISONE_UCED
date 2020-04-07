# -*- coding: utf-8 -*-
"""
Created on Mon Oct  8 11:45:39 2018

@author: jkern
"""

############################################################################
#                           UC/ED Simulation

# This file simulates power system/market operations for the NEISO
# market, stores the data in appropriate places and calculates wholesale
# electricity prices. 
############################################################################

############################################################################
# Simulates power system operations for as many simulation days as 
# specified (max is 365)
days = 365
                           
# NEISO
import NEISO_wrapper
NEISO_wrapper.sim(days)

#
############################################################################
#                    WHOLESALE ELECTRICITY PRICES
# 
