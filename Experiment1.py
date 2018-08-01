########################### Preamble ###########################
from copy import deepcopy
import DataInterface
import Models
import pickle

CostOfPUnderOrOverGenerationInUSDPerMWh = 5000
GenerateTrajectories = True # True False
NumberOfSimulations = 2
NumberOfTimePeriodsInDispatchModel = 1
PrintOutput = True # True False



########################### Read Data ###########################
FileName = 'Case014.xlsx';
#FileName = 'Case118.xlsx';
Data = DataInterface.ReadDataFile(FileName)
Data['CostOfPUnderOrOverGenerationInUSDPerMWh'] = CostOfPUnderOrOverGenerationInUSDPerMWh



########################### Trajectories for wind and solar power ###########################
if GenerateTrajectories == True:
    Trajectories = [[] for k in range(NumberOfSimulations)]
    for k in range(NumberOfSimulations):
        Trajectories[k] = Models.SimulateTrajectoryForWindAndSolarPower(Data)
    pickle.dump(Trajectories, open('Trajectories' + FileName + '.pkl','wb'))
else:
    Trajectories = pickle.load(open('Trajectories' + FileName + '.pkl','rb'))
    
    
    
########################### Simulation ###########################
### Detemine a unit commitment solution where all generators are ON all the time ###
if PrintOutput == True:
    'Using unit commitment solution where all generators are ON all the time'
UCSolution = {}
UCSolution['x'] = [[1.0 for i in range(len(Data['Generator']))] for t in range(len(Data['TimePeriodsSimulation']))]
UCSolution['u'] = [[0.0 for i in range(len(Data['Generator']))] for t in range(len(Data['TimePeriodsSimulation']))]
UCSolution['u'][0] = [1.0 for i in range(len(Data['Generator']))]
UCSolution['v'] = [[0.0 for i in range(len(Data['Generator']))] for t in range(len(Data['TimePeriodsSimulation']))]

### Real-time simulation for all trajectories ###
DispatchSolution = [[[] for tSimulation in range(len(Data['TimePeriodsSimulation']))] for k in range(NumberOfSimulations)]
for k in range(NumberOfSimulations):
    if PrintOutput == True:
        print '\nSimulating real-time dispatch under trajectory', k
        
    for tSimulation in range(len(Data['TimePeriodsSimulation'])):
        ### In tSimulation we solve an optimization problem that needs the data below ###
        Data['tSimulation'] = tSimulation
        Data['TimePeriod'] = range(min([NumberOfTimePeriodsInDispatchModel,len(Data['TimePeriodsSimulation'])-tSimulation]))
        Data['PdInDispatchModel'] = [[Data['PdProfile'][tSimulation+t][BusIndex] for BusIndex in range(len(Data['Bus']))] for t in range(len(Data['TimePeriod']))]
        Data['QdInDispatchModel'] = [[Data['QdProfile'][tSimulation+t][BusIndex] for BusIndex in range(len(Data['Bus']))] for t in range(len(Data['TimePeriod']))]
        
        ### The dispatch model observes the realized renewable power at time tSimulation and the forecast for renewable power in subsequent time periods ###
        Data['PowerAvailabilityInDispatchModel'] = [[Trajectories[k][tSimulation+t][iIndex] for iIndex in range(len(Data['Generator']))] for t in range(len(Data['TimePeriod']))]
        for t in range(len(Data['TimePeriod'])):
            if t == 0:
                Data['PowerAvailabilityInDispatchModel'][t] = [Trajectories[k][tSimulation+t][iIndex] for iIndex in range(len(Data['Generator']))]
            else:
                Data['PowerAvailabilityInDispatchModel'][t] = [Data['PowerAvailabilityProfile'][t][iIndex] for iIndex in range(len(Data['Generator']))]
                
        ### UC decision determines bounds and ramping capacities for conventional generators, and the trajectory and profile for renewables determines bounds for renewable generators ###
        Data['GeneratorLBInDispatchModel'] = [[[] for iIndex in range(len(Data['Generator']))] for t in range(len(Data['TimePeriod']))]
        Data['GeneratorUBInDispatchModel'] = [[[] for iIndex in range(len(Data['Generator']))] for t in range(len(Data['TimePeriod']))]
        Data['GeneratorRDInDispatchModel'] = [[[] for iIndex in range(len(Data['Generator']))] for t in range(len(Data['TimePeriod']))]
        Data['GeneratorRUInDispatchModel'] = [[[] for iIndex in range(len(Data['Generator']))] for t in range(len(Data['TimePeriod']))]
        for t in range(len(Data['TimePeriod'])):
            for i in range(len(Data['Generator'])):
                if Data['GeneratorType'][i] == 'Wind' or Data['GeneratorType'][i] == 'Solar':
                    Data['GeneratorLBInDispatchModel'][t][i] = 0.0
                    Data['GeneratorUBInDispatchModel'][t][i] = Data['PowerAvailabilityInDispatchModel'][t][i]
                    Data['GeneratorRDInDispatchModel'][t][i] = Data['PowerAvailabilityInDispatchModel'][t][i]
                    Data['GeneratorRUInDispatchModel'][t][i] = Data['PowerAvailabilityInDispatchModel'][t][i]
                else:
                    Data['GeneratorLBInDispatchModel'][t][i] = Data['GeneratorPmin'][i] * UCSolution['x'][tSimulation+t][i]
                    Data['GeneratorUBInDispatchModel'][t][i] = Data['GeneratorPmax'][i] * UCSolution['x'][tSimulation+t][i]
                    if t > 0 or (t == 0 and tSimulation > 0):
                        Data['GeneratorRDInDispatchModel'][t][i] = Data['GeneratorRampDownInMWPerHour'][i] * UCSolution['x'][tSimulation+t][i] + Data['GeneratorSRampDownInMW'][i] * UCSolution['v'][tSimulation+t][i]
                        Data['GeneratorRUInDispatchModel'][t][i] = Data['GeneratorRampUpInMWPerHour'][i] * UCSolution['x'][tSimulation+t-1][i] + Data['GeneratorSRampUpInMW'][i] * UCSolution['u'][tSimulation+t][i]
                        
        ### We might need to reduce Data['GeneratorUBInDispatchModel'][t][i] in the last time period if the generator will be shut down afterwards ###
        T = len(Data['TimePeriod']) - 1
        for i in range(len(Data['Generator'])):
            if not (Data['GeneratorType'][i] == 'Wind' or Data['GeneratorType'][i] == 'Solar') and UCSolution['x'][tSimulation+T][i] == 1.0:
                for tPrime in range(tSimulation + T + 1, len(Data['TimePeriodsSimulation'])):
                    if UCSolution['x'][tPrime][i] == 0.0:
                        if (tPrime - tSimulation - T - 1) * Data['GeneratorRampDownInMWPerHour'][i] + Data['GeneratorSRampDownInMW'][i] < Data['GeneratorUBInDispatchModel'][T][i]:
                            print 'GeneratorUBInDispatchModel updated:', (tPrime - tSimulation - T - 1) * Data['GeneratorRampDownInMWPerHour'][i] + Data['GeneratorSRampDownInMW'][i], 'instead of', Data['GeneratorUBInDispatchModel'][T][i]
                        Data['GeneratorUBInDispatchModel'][T][i] = min([Data['GeneratorUBInDispatchModel'][T][i], (tPrime - tSimulation - T - 1) * Data['GeneratorRampDownInMWPerHour'][i] + Data['GeneratorSRampDownInMW'][i]])
                        break
                        
        ### We also need the decisions implemented from the last dispatch to use in the ramping constraints of the dispatch model ###
        Data['GeneratorInitialPgInDispatchModel'] = [[] for iIndex in range(len(Data['Generator']))]
        if tSimulation > 0:
            Data['GeneratorInitialPgInDispatchModel'] = [DispatchSolution[k][tSimulation-1]['Pg'][0][iIndex] for iIndex in range(len(Data['Generator']))]
            
        ### Save Data to a file: This is useful to study the dispatch problem if there is an error. ###
        if True: # False True
            pickle.dump(Data,open('Data' + FileName + '.pkl','wb'))
            
        ### We can now solve the dispatch problem ###
        DispatchSolution[k][tSimulation] = Models.SolveDCOPF(Data)
        
        ### Calculate costs ###
        DispatchSolution[k][tSimulation]['DispatchCost'] = sum([Data['GeneratorVariableCostInUSDPerMWh'][i] * DispatchSolution[k][tSimulation]['Pg'][0][i] for i in range(len(Data['Generator']))]) 
        DispatchSolution[k][tSimulation]['PenaltyCost'] = Data['CostOfPUnderOrOverGenerationInUSDPerMWh'] * sum(DispatchSolution[k][tSimulation]['PUnderGeneration'][0]) + Data['CostOfPUnderOrOverGenerationInUSDPerMWh'] * sum(DispatchSolution[k][tSimulation]['POverGeneration'][0])
        DispatchSolution[k][tSimulation]['TotalCost'] = DispatchSolution[k][tSimulation]['DispatchCost'] + DispatchSolution[k][tSimulation]['PenaltyCost']
        
        if PrintOutput == True:
            print '\nk:', k, '; tSimulation:', tSimulation
            print 'TotalDemand:', sum(Data['PdInDispatchModel'][0]), '; TotalWind:', sum([Data['PowerAvailabilityInDispatchModel'][0][i] for i in range(len(Data['Generator'])) if Data['GeneratorType'][i] == 'Wind']), '; TotalSolar:', sum([Data['PowerAvailabilityInDispatchModel'][0][i] for i in range(len(Data['Generator'])) if Data['GeneratorType'][i] == 'Solar']), '; TotalNetLoad:', sum(Data['PdInDispatchModel'][0]) - sum([Data['PowerAvailabilityInDispatchModel'][0][i] for i in range(len(Data['Generator'])) if Data['GeneratorType'][i] == 'Wind' or Data['GeneratorType'][i] == 'Solar'])
            print 'ConventionalGeneration:', sum([DispatchSolution[k][tSimulation]['Pg'][0][i] for i in range(len(Data['Generator'])) if not (Data['GeneratorType'][i] == 'Wind' or Data['GeneratorType'][i] == 'Solar')]), '; WindGeneration:', sum([DispatchSolution[k][tSimulation]['Pg'][0][i] for i in range(len(Data['Generator'])) if Data['GeneratorType'][i] == 'Wind']), '; SolarGeneration:', sum([DispatchSolution[k][tSimulation]['Pg'][0][i] for i in range(len(Data['Generator'])) if Data['GeneratorType'][i] == 'Solar']), '; PUnderGeneration:', sum(DispatchSolution[k][tSimulation]['PUnderGeneration'][0]), '; POverGeneration:', sum(DispatchSolution[k][tSimulation]['POverGeneration'][0])
            print 'DispatchCost:', DispatchSolution[k][tSimulation]['DispatchCost'], '; PenaltyCost:', DispatchSolution[k][tSimulation]['PenaltyCost'], '; TotalCost:', DispatchSolution[k][tSimulation]['TotalCost']
            if True: # False True
                print 'Pg:', DispatchSolution[k][tSimulation]['Pg'][0]
                
                
                
########################### Analysis of Results ###########################
print '\nAnalysis of Results'

DispatchCostAverage = (1.0/float(NumberOfSimulations)) * sum([sum([DispatchSolution[k][tSimulation]['DispatchCost'] for tSimulation in range(len(Data['TimePeriodsSimulation']))]) for k in range(NumberOfSimulations)])
PenaltyCostAverage = (1.0/float(NumberOfSimulations)) * sum([sum([DispatchSolution[k][tSimulation]['PenaltyCost'] for tSimulation in range(len(Data['TimePeriodsSimulation']))]) for k in range(NumberOfSimulations)])
TotalCostAverage = (1.0/float(NumberOfSimulations)) * sum([sum([DispatchSolution[k][tSimulation]['TotalCost'] for tSimulation in range(len(Data['TimePeriodsSimulation']))]) for k in range(NumberOfSimulations)])

print 'DispatchCostAverage:', DispatchCostAverage
print 'PenaltyCostAverage:', PenaltyCostAverage
print 'TotalCostAverage:', TotalCostAverage
