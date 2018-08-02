from __future__ import division
import pyomo.environ
from pyomo.opt import SolverFactory, SolverStatus
from pyomo.environ import *
import pandas as pd
import os



def build_abstract_model(model):

    model.GEN = Set()
    
    model.gen_capacity = Param(model.GEN)
    model.gen_min_up = Param(model.GEN)
    model.gen_min_down = Param(model.GEN)
    
    model.GenCommit = Var(model.GEN, domain=NonNegativeReals)
    model.GenPg = Var(model.GEN, domain=NonNegativeReals)

    return model

def load_data(m, data):

    data.load(filename='gen.tab', param=(m.gen_capacity,m.gen_min_up,m.gen_min_down), index=m.GEN)

    return data