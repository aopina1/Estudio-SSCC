""" The birth of NEWEN Operation: UNIT COMMITMENT"""

from __future__ import division
import pyomo.environ
from pyomo.opt import SolverFactory, SolverStatus
from pyomo.environ import *
import pandas as pd
import os
from timeit import default_timer as timer

""" 
ISSUES:
1. Consider including an initial value for GenCommit and GenPg as a parameter, for specific constraints Gen_On_Off and Gen_Ramps. So far,
all generators start with an OFF status and Pg=0.
2. This Commit is still uninodal. Consider coding the extension to a multinodal system.
3. Define outputs format
"""

model = AbstractModel()

model.LOADZONES = Set()
model.GEN = Set()
#model.LINE = Set()
model.TIMEPOINT = Set()

model.GenCommit = Var(model.GEN, model.TIMEPOINT, domain=Binary)
model.GenStartUp = Var(model.GEN, model.TIMEPOINT, domain=Binary)
model.GenShutDown = Var(model.GEN, model.TIMEPOINT, domain=Binary)
model.GenPg = Var(model.GEN, model.TIMEPOINT, domain=NonNegativeReals)

model.OverGeneration = Var(model.LOADZONES, model.TIMEPOINT, domain=NonNegativeReals)
model.LoadShedding = Var(model.LOADZONES, model.TIMEPOINT, domain=NonNegativeReals)

model.noloadcost = Param(model.GEN)
model.startupcost = Param(model.GEN)
model.variablecost = Param(model.GEN)
model.mindowntime = Param(model.GEN)
model.minuptime = Param(model.GEN)
model.genpmin = Param(model.GEN)
model.genpmax = Param(model.GEN)
model.rampup = Param(model.GEN)
model.rampdown = Param(model.GEN)
model.shutdownramp = Param(model.GEN)
model.startupramp = Param(model.GEN)

model.overgencost = Param(model.LOADZONES)
model.loadsheddingcost = Param(model.LOADZONES)

###INCIDENCE MATRICES
#
#model.generationshiftfactor = Param(model.LINE)
#model.flowlimit = Param(model.LINE)

model.zonedemand = Param(model.TIMEPOINT,model.LOADZONES)

def obj_expression(m):
    return sum(sum(
            (m.noloadcost[gen]*m.GenCommit[gen,t] 
            + m.startupcost[gen]*m.GenStartUp[gen,t] 
            + m.variablecost[gen]*m.GenPg[gen,t])
            for gen in m.GEN)
            +sum(m.overgencost[loadzone]*m.OverGeneration[loadzone,t]
            +m.loadsheddingcost[loadzone]*m.LoadShedding[loadzone,t]
                    for loadzone in m.LOADZONES)
               for t in m.TIMEPOINT)
model.Obj = Objective(rule=obj_expression)  

def gen_p_min_rule(m, gen, t):
    return (m.GenCommit[gen,t]*m.genpmin[gen]<= m.GenPg[gen,t])
model.Gen_P_Min = Constraint(model.GEN, model.TIMEPOINT, rule=gen_p_min_rule)

def gen_p_max_rule(m, gen, t):
    return (m.GenPg[gen,t] <= m.GenCommit[gen,t]*m.genpmax[gen])
model.Gen_P_Max = Constraint(model.GEN, model.TIMEPOINT, rule=gen_p_max_rule)

def on_off_rule(m, gen, t):
    if t>1:
        return m.GenCommit[gen,t]-m.GenCommit[gen,(t-1)] == m.GenStartUp[gen,t] - m.GenShutDown[gen,t]
    if t==1:
        return m.GenCommit[gen,t] == m.GenStartUp[gen,t] - m.GenShutDown[gen,t]
model.Gen_On_Off = Constraint(model.GEN, model.TIMEPOINT, rule=on_off_rule)
    
def min_start_up_rule(m, gen, t):
    if (t <= len(m.TIMEPOINT) - m.minuptime[gen]+1):
        return sum(m.GenCommit[gen,tau] for tau in range(t,t+m.minuptime[gen])) >= m.minuptime[gen]*m.GenStartUp[gen,t]
    else:
        return Constraint.Feasible
model.Gen_Min_Start_Up = Constraint(model.GEN, model.TIMEPOINT, rule=min_start_up_rule)

def min_shut_down_rule(m, gen, t):
    if (t <= len(m.TIMEPOINT) - m.mindowntime[gen]+1):
        return sum((1-m.GenCommit[gen,tau]) for tau in range(t,t+m.mindowntime[gen])) >= m.mindowntime[gen]*m.GenShutDown[gen,t]
    else:
        return Constraint.Feasible
model.Gen_Min_Shut_Down = Constraint(model.GEN, model.TIMEPOINT, rule=min_shut_down_rule)

def min_start_up_bound_rule(m, gen, t):
    if (t >= len(m.TIMEPOINT) - m.minuptime[gen]+1):
        return sum(m.GenCommit[gen,tau]-m.GenStartUp[gen,t] for tau in range(t, len(m.TIMEPOINT)+1)) >=0
    else:
        return Constraint.Feasible
model.Gen_Min_Start_Up_Bound = Constraint(model.GEN, model.TIMEPOINT, rule=min_start_up_bound_rule)

def min_shut_down_bound_rule(m, gen, t):
    if (t >= len(m.TIMEPOINT) - m.mindowntime[gen]+1):
        return sum(1-m.GenCommit[gen,tau]-m.GenShutDown[gen,t] for tau in range(t,len(m.TIMEPOINT)+1)) >= 0 
    else:
        return Constraint.Feasible
model.Gen_Min_Shut_Down_Bound = Constraint(model.GEN, model.TIMEPOINT, rule=min_shut_down_bound_rule)

def lower_ramp_rule(m, gen, t):
    if t>1:
        return (-m.rampdown[gen]*m.GenCommit[gen,t] - m.shutdownramp[gen]*m.GenShutDown[gen,t] 
                <= m.GenPg[gen,t]-m.GenPg[gen,(t-1)] 
                )
    if t==1:
        return (-m.rampdown[gen]*m.GenCommit[gen,t] - m.shutdownramp[gen]*m.GenShutDown[gen,t] 
                <= m.GenPg[gen,t] 
                )
model.Lower_Gen_Ramps = Constraint(model.GEN, model.TIMEPOINT, rule=lower_ramp_rule)

def upper_ramp_rule(m, gen, t):
    if t>1:
        return (m.GenPg[gen,t]-m.GenPg[gen,(t-1)] 
                <= m.rampup[gen]*m.GenCommit[gen,t] + m.startupramp[gen]*m.GenStartUp[gen,t])
    if t==1:
        return (m.GenPg[gen,t] 
                <= m.rampup[gen]*m.GenCommit[gen,t] + m.startupramp[gen]*m.GenStartUp[gen,t])
model.Upper_Gen_Ramps = Constraint(model.GEN, model.TIMEPOINT, rule=upper_ramp_rule)


def load_balance_rule(m,t):
    return (sum(m.zonedemand[t, loadzone]for loadzone in m.LOADZONES) 
            + sum(m.OverGeneration[loadzone,t] for loadzone in m.LOADZONES)
            - sum(m.LoadShedding[loadzone,t] for loadzone in m.LOADZONES)
            == sum(m.GenPg[gen,t] for gen in m.GEN) )
model.Load_Balance = Constraint(model.TIMEPOINT, rule=load_balance_rule)

#def power_flow_rule(m, line, t):
#    return Constraint.Feasible
#model.Power_Flow = Constraint(model.LINE, model.TIMEPOINT, rule=power_flow_rule)


data = DataPortal()
inputs_dir = 'uc_inputs'

data.load(filename=os.path.join(inputs_dir,'load_zones.tab'), 
          param=(model.overgencost, model.loadsheddingcost), 
          index=model.LOADZONES)

data.load(filename=os.path.join(inputs_dir,'timepoints.tab'), set=model.TIMEPOINT, format="set")

data.load(filename=os.path.join(inputs_dir,'gen.tab'), 
          param=(model.noloadcost, model.startupcost, model.variablecost, 
                 model.mindowntime, model.minuptime, 
                 model.genpmin, model.genpmax, 
                 model.rampup, model.rampdown,
                 model.startupramp, model.shutdownramp),
          index=model.GEN)

#data.load(filename=os.path.join(inputs_dir,'line.tab'), 
#          param=(model.generationshiftfactor, model.flowlimit), 
#          index=model.LINE)

data.load(filename=os.path.join(inputs_dir,'zone_demand.tab'),
          param=model.zonedemand,
          format='array')

instance = model.create_instance(data)
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
        
stream_solver =True # True prints solver output to screen
keepfiles = False # True prints intermediate file names (.nl,.sol,...)
start = timer()
results = opt.solve(instance,keepfiles=keepfiles,tee=stream_solver)
end = timer()
print("Problem solved in: %.4fseconds" % (end-start))
time = end-start
print time

GEN_COMMIT = pd.DataFrame(data=[[instance.GenCommit[gen,t].value for gen in instance.GEN] for t in instance.TIMEPOINT], 
                          columns=instance.GEN, index=instance.TIMEPOINT)
GEN_START_UP = pd.DataFrame(data=[[instance.GenStartUp[gen,t].value for gen in instance.GEN] for t in instance.TIMEPOINT], 
                          columns=instance.GEN, index=instance.TIMEPOINT)
GEN_SHUT_DOWN = pd.DataFrame(data=[[instance.GenShutDown[gen,t].value for gen in instance.GEN] for t in instance.TIMEPOINT], 
                          columns=instance.GEN, index=instance.TIMEPOINT)
GEN_PG = pd.DataFrame(data=[[instance.GenPg[gen,t].value for gen in instance.GEN] for t in instance.TIMEPOINT], 
                          columns=instance.GEN, index=instance.TIMEPOINT)
LOAD_SHEDDING = pd.DataFrame(data=[[instance.LoadShedding[lz,t].value for lz in instance.LOADZONES] for t in instance.TIMEPOINT],
                             columns= instance.LOADZONES, index=instance.TIMEPOINT)
OVER_GEN = pd.DataFrame(data=[[instance.OverGeneration[lz,t].value for lz in instance.LOADZONES] for t in instance.TIMEPOINT],
                             columns= instance.LOADZONES, index=instance.TIMEPOINT)

script_dir = os.path.dirname(__file__)
results_dir = 'uc_outputs'

if not os.path.isdir(results_dir):
    os.makedirs(results_dir)
    
GEN_COMMIT.to_csv(os.path.join(results_dir,'gen_commit.tab'),sep='\t')
GEN_START_UP.to_csv(os.path.join(results_dir,'gen_start_up.tab'),sep='\t')
GEN_SHUT_DOWN.to_csv(os.path.join(results_dir,'gen_shut_down.tab'),sep='\t')
GEN_PG.to_csv(os.path.join(results_dir,'gen_pg.tab'),sep='\t')
LOAD_SHEDDING.to_csv(os.path.join(results_dir,'load_shedding.tab'),sep='\t')
OVER_GEN.to_csv(os.path.join(results_dir,'over_gen.tab'),sep='\t')