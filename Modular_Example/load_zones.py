from __future__ import division
import pyomo.environ
from pyomo.opt import SolverFactory, SolverStatus
from pyomo.environ import *
import pandas as pd
import os



def build_abstract_model(model):

    model.LOADZONES = Set()
#    model.TIMEPOINT
#    model.ZONE_TIMEPOINTS = Set(dimen=2, initialize=lambda m: m.LOAD_ZONES * m.TIMEPOINTS)
        
    model.Zone_Power_Injections = []
    model.Zone_Power_Withdrawals = []
    
#    model.Zone_Energy_Balance = Constraint(
#    model.ZONE_TIMEPOINTS,
#    rule=lambda m, z, t: 
#        (sum(getattr(m, component)[z, t] for component in m.Zone_Power_Injections
#             ) == sum(getattr(m, component)[z, t] for component in m.Zone_Power_Withdrawals))
#    )
        
    return model

def load_data(model, data):

    data.load(filename='load_zones.tab', set=model.LOADZONES, format="set")
    
    return data
