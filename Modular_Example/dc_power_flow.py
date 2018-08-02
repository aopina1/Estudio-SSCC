from __future__ import division
import pyomo.environ
from pyomo.opt import SolverFactory, SolverStatus
from pyomo.environ import *
import pandas as pd
import os



def build_abstract_model(model):
    
    model.LINE = Set()
    
    model.line_capacity = Param(model.LINE)
    model.line_b = Param(model.LINE)
        
    
    return model

def load_data(model, data):
    
    data.load(filename='line.tab', set=model.LINE, format="set")

    return data