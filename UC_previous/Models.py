import pyomo.environ
from pyomo.opt import SolverFactory, SolverStatus

BaseMVA = 100.0
PrintOutput = True # False True
PrintDispatchSolverOutput = False # False True
PrintUnitCommitmentSolverOutput = True # False True

def SimulateTrajectoryForWindAndSolarPower(Data):
    import numpy as np
    CoefficientOfVariation = 0.3
    Trajectory = [[0.0 for iIndex in range(len(Data['Generator']))] for t in range(len(Data['TimePeriodsSimulation']))]
    for t in range(len(Data['TimePeriodsSimulation'])):
        for i in range(len(Data['Generator'])):
            if not Data['PowerAvailabilityProfile'][t][i] == 0.0:
                mu = Data['PowerAvailabilityProfile'][t][i]
                sigma = CoefficientOfVariation * Data['PowerAvailabilityProfile'][t][i]
                Trajectory[t][i] = max([0.0,np.random.normal(mu,sigma,1)[0]])
                
    if PrintOutput == True:
        TotalDemandProfile = [sum(Data['PdProfile'][t]) for t in range(len(Data['TimePeriodsSimulation']))]
        TotalWindPower = [sum([Trajectory[t][i] for i in range(len(Data['Generator'])) if Data['GeneratorType'][i] == 'Wind']) for t in range(len(Data['TimePeriodsSimulation']))]
        TotalSolarPower = [sum([Trajectory[t][i] for i in range(len(Data['Generator'])) if Data['GeneratorType'][i] == 'Solar']) for t in range(len(Data['TimePeriodsSimulation']))]
        TotalNetLoad = [TotalDemandProfile[t]-TotalWindPower[t]-TotalSolarPower[t] for t in range(len(Data['TimePeriodsSimulation']))]
        print '\nTrajectory simulated'
        print 't, TotalDemandProfile, TotalWindPowerProfile, TotalSolarPowerProfile, TotalNetLoadProfile'
        for t in range(len(Data['TimePeriodsSimulation'])):
            print t, ',', TotalDemandProfile[t], ',', TotalWindPower[t], ',', TotalSolarPower[t], ',', TotalNetLoad[t]
            
    return Trajectory
    
def SolveDCOPF(Data):
    ######### Create Problem #########
    ### Create Pyomo instance ###
    Problem = pyomo.environ.ConcreteModel()
    
    ### Create variables ###
    Problem.Pg = pyomo.environ.Var( range(len(Data['TimePeriod'])), range(len(Data['Generator'])) )
    for t in range(len(Data['TimePeriod'])):
        for i in range(len(Data['Generator'])):
            Problem.Pg[t,i].setlb(Data['GeneratorLBInDispatchModel'][t][i])
            Problem.Pg[t,i].setub(Data['GeneratorUBInDispatchModel'][t][i])
            if Data['GeneratorLBInDispatchModel'][t][i] > Data['GeneratorUBInDispatchModel'][t][i]:
                print 'Error: Inconsistent bounds in dispatch:', Data['GeneratorLBInDispatchModel'][t][i], Data['GeneratorUBInDispatchModel'][t][i]
                
    Problem.PUnderGeneration = pyomo.environ.Var( range(len(Data['TimePeriod'])), range(len(Data['Bus'])), within = pyomo.environ.NonNegativeReals )
    Problem.POverGeneration = pyomo.environ.Var( range(len(Data['TimePeriod'])), range(len(Data['Bus'])), within = pyomo.environ.NonNegativeReals )
    
    Problem.Theta = pyomo.environ.Var( range(len(Data['TimePeriod'])), range(len(Data['Bus'])), within = pyomo.environ.Reals )
    for t in range(len(Data['TimePeriod'])):
        ### Reference bus ###
        Problem.Theta[t,0].setlb(0.0)
        Problem.Theta[t,0].setub(0.0)
        
    ### Create objective function ###
    OFExpression = 0.0
    for t in range(len(Data['TimePeriod'])):
        for i in range(len(Data['Generator'])):
            OFExpression += Data['GeneratorVariableCostInUSDPerMWh'][i] * Problem.Pg[t,i]
        for i in range(len(Data['Bus'])):
            OFExpression += Data['CostOfPUnderOrOverGenerationInUSDPerMWh'] * Problem.PUnderGeneration[t,i]
            OFExpression += Data['CostOfPUnderOrOverGenerationInUSDPerMWh'] * Problem.POverGeneration[t,i]
            
    Problem.ObjectiveFunction = pyomo.environ.Objective(expr = OFExpression, sense=pyomo.environ.minimize)
    
    ### Create list of constraints ###
    Problem.ListOfConstraints = pyomo.environ.ConstraintList()
    
    ### Ramping constraints ###
    for t in range(len(Data['TimePeriod'])):
        if t == 0:
            if Data['tSimulation'] > 0:
                for i in range(len(Data['Generator'])):
                    if not (Data['GeneratorType'][i] == 'Wind' or Data['GeneratorType'][i] == 'Solar'):
                        Problem.ListOfConstraints.add( Problem.Pg[t,i] - Data['GeneratorInitialPgInDispatchModel'][i] >= -Data['GeneratorRDInDispatchModel'][t][i] )
                        Problem.ListOfConstraints.add( Problem.Pg[t,i] - Data['GeneratorInitialPgInDispatchModel'][i] <= Data['GeneratorRUInDispatchModel'][t][i] )
        else:
            for i in range(len(Data['Generator'])):
                if not (Data['GeneratorType'][i] == 'Wind' or Data['GeneratorType'][i] == 'Solar'):
                    Problem.ListOfConstraints.add( Problem.Pg[t,i] - Problem.Pg[t-1,i] >= -Data['GeneratorRDInDispatchModel'][t][i] )
                    Problem.ListOfConstraints.add( Problem.Pg[t,i] - Problem.Pg[t-1,i] <= Data['GeneratorRUInDispatchModel'][t][i] )
                    
    ### Energy balance constraints ###
    for t in range(len(Data['TimePeriod'])):
        for i in range(len(Data['Bus'])):
            LHSExpression = 0.0; RHSExpression = 0.0;
            for k in Data['SetOfGeneratorsAtBus'][i]:
                LHSExpression += Problem.Pg[t,k]
            LHSExpression -= Data['PdInDispatchModel'][t][i]
            LHSExpression += Problem.PUnderGeneration[t,i]
            LHSExpression -= Problem.POverGeneration[t,i]
            for j in Data['SetOfNeighborsOfBus'][i]:
                RHSExpression += (BaseMVA * Data['B'][i][j]) * ( Problem.Theta[t,i] - Problem.Theta[t,j] )
            Problem.ListOfConstraints.add( LHSExpression == RHSExpression )
            
    ### Transmission constraints ###
    for t in range(len(Data['TimePeriod'])):
        for [i,j] in Data['SetL']:
            PijExpression = (BaseMVA * Data['B'][i][j]) * ( Problem.Theta[t,i] - Problem.Theta[t,j] )
            Problem.ListOfConstraints.add( PijExpression <= Data['LineMaxFlowForPairOfBuses'][i][j])
            
    ######### Solve Problem #########
    ### Call Solver ###
    Problem.preprocess()
    
    solver = 'gurobi' # ipopt gurobi
    if solver == 'gurobi':
        opt = SolverFactory(solver)
    elif solver == 'ipopt':
        solver_io = 'nl'
        opt = SolverFactory(solver,solver_io=solver_io)
    if opt is None:
        print("ERROR: Unable to create solver plugin for %s using the %s interface" % (solver, solver_io))
        exit(1)
        
    stream_solver = PrintDispatchSolverOutput # True prints solver output to screen
    keepfiles = False # True prints intermediate file names (.nl,.sol,...)
    results = opt.solve(Problem,keepfiles=keepfiles,tee=stream_solver)
    Problem.solutions.load_from(results)
    
    ### Extract Solution ###
    Solution = {}
    if not results.solver.status == SolverStatus.ok:
        print 'results.solver.status =', results.solver.status
    else:
        Solution['ObjectiveValue'] = pyomo.environ.value(Problem.ObjectiveFunction)
        Solution['Pg'] = [[Problem.Pg[t,i].value for i in range(len(Data['Generator']))] for t in range(len(Data['TimePeriod']))]
        for t in range(len(Data['TimePeriod'])):
            for i in range(len(Data['Generator'])):
                if Solution['Pg'][t][i] == None:
                    Solution['Pg'][t][i] = 0.0
        Solution['PUnderGeneration'] = [[Problem.PUnderGeneration[t,i].value for i in range(len(Data['Bus']))] for t in range(len(Data['TimePeriod']))]
        Solution['POverGeneration'] = [[Problem.POverGeneration[t,i].value for i in range(len(Data['Bus']))] for t in range(len(Data['TimePeriod']))]
        #Solution['Theta'] = [[Problem.Theta[t,i].value for i in range(len(Data['Bus']))] for t in range(len(Data['TimePeriod']))]
        #Solution['Pij'] = [[BaseMVA * Data['B'][i][j] * ( Problem.Theta[t,i].value - Problem.Theta[t,j].value) for [i,j] in Data['SetL']] for t in range(len(Data['TimePeriod']))]
        
    return Solution