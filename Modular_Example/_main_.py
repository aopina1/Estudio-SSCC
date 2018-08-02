""" The birth of NEWEN Operation"""

from __future__ import division
import pyomo.environ
from pyomo.opt import SolverFactory, SolverStatus
from pyomo.environ import *
import pandas as pd
import os
from timeit import default_timer as timer


supertest = AbstractModel()
othertest = DataPortal()

imports = (pd.read_csv('modules.txt'))['MODULES']
modules = []

for x in imports:
    try:
        modules.append(__import__(x))
        print "Successfully imported ", x, '.'
    except ImportError:
        print "Error importing ", x, '.'

for module in modules:
    supertest = module.build_abstract_model(supertest)
    print "Successfully executed ", module, '.'

for module in modules:
    othertest = module.load_data(supertest, othertest)
    print "Data successfully loaded for ", module, '.'
    
def obj_expression(m):
    return sum(m.GenCommit[g]*m.gen_capacity[g] for g in m.GEN)

supertest.obj = Objective(rule=obj_expression)
print "Test OF successfully created"

instance = supertest.create_instance(othertest)
print "Model instance successfully created"


instance.preprocess()
solver = 'gurobi' # ipopt gurobi
if solver == 'gurobi':
    opt = SolverFactory(solver)
elif solver == 'ipopt':
    solver_io = 'nl'
    opt = SolverFactory(solver,solver_io=solver_io)
if opt is None:
    print("ERROR: Unable to create solver plugin for %s using the %s interface" % (solver, solver_io))
    exit(1)
        
stream_solver =False # True prints solver output to screen
keepfiles = False # True prints intermediate file names (.nl,.sol,...)
start = timer()
results = opt.solve(instance,keepfiles=keepfiles,tee=stream_solver)
end = timer()
print("Problem solved in: %.4fseconds" % (end-start))
time = end-start
print time


