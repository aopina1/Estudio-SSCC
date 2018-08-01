BaseMVA = 100.0
PrintOutput = True # False True

def ReadDataFile(FileName):
    ### This function reads an Excel file called "FileName" and returns a dictionary with all the data in this file ###
    
    ### Important note ###
    # In this file, voltages and admittance matrix elements are in p.u.
    # Given this, Bij (Thetai-Thetaj) will be in p.u. and needs to be multiplied by BaseMVA to transform it to MW
    # An analogous conversion needs to be carried out for other expressions involving G and B
    # In summary, to transform p.u. to MW multiply by BaseMVA and to transfrom p.u. to MVAR also multiply by BaseMVA
    
    ### Open Excel file and create dictionary ###
    if PrintOutput == True:
        print 'Reading data file:', FileName
        
    import xlrd
    book = xlrd.open_workbook(FileName)
    
    Data = {}
    Data['FileName'] = FileName
    Data['TimePeriodsSimulation'] = range(1,25)
    
    ### Read worksheet Buses ###
    sh = book.sheet_by_name('Buses')
    Data['Bus'] = []
    Data['BusVmax'] = []
    Data['BusVmin'] = []
    Data['BusShuntConductance'] = []
    Data['BusShuntSusceptance'] = []
    row = 1
    while True:
        col = 0
        if str(sh.cell_value(row, col)) == 'END':
            break
        else:
            col = 0; Data['Bus'].append(str(sh.cell_value(row, col)));
            col += 1; Data['BusVmax'].append(sh.cell_value(row, col));
            col += 1; Data['BusVmin'].append(sh.cell_value(row, col));
            col += 1; Data['BusShuntConductance'].append(sh.cell_value(row, col));
            col += 1; Data['BusShuntSusceptance'].append(sh.cell_value(row, col));
            row += 1
            
    ### Read worksheet Demand ###
    sh = book.sheet_by_name('Demand')
    Data['PdProfile'] = [[0.0 for BusIndex in range(len(Data['Bus']))] for t in range(len(Data['TimePeriodsSimulation']))]
    Data['QdProfile'] = [[0.0 for BusIndex in range(len(Data['Bus']))] for t in range(len(Data['TimePeriodsSimulation']))]
    row = 2
    while True:
        col = 0
        if str(sh.cell_value(row, col)) == 'END':
            break
        else:
            col = 0; BusIndex = Data['Bus'].index(str(sh.cell_value(row,col)));
            for t in range(len(Data['TimePeriodsSimulation'])):
                col = t+1; Data['PdProfile'][t][BusIndex] = sh.cell_value(row, col);
                col = t+1+26; Data['QdProfile'][t][BusIndex] = sh.cell_value(row, col);
            row += 1
            
    ### Read worksheet Generators ###
    sh = book.sheet_by_name('Generators')
    Data['Generator'] = []
    Data['GeneratorBusIndex'] = []
    Data['GeneratorPmax'] = []
    Data['GeneratorPmin'] = []
    Data['GeneratorQmax'] = []
    Data['GeneratorQmin'] = []
    Data['GeneratorRampUpInMWPerHour'] = []
    Data['GeneratorRampDownInMWPerHour'] = []
    Data['GeneratorSRampUpInMW'] = []
    Data['GeneratorSRampDownInMW'] = []
    Data['GeneratorMinUp'] = []
    Data['GeneratorMinDw'] = []
    Data['GeneratorInitS'] = []
    Data['GeneratorInitP'] = []
    Data['GeneratorStartUpCostInUSD'] = []
    Data['GeneratorFixedCostInUSD'] = []
    Data['GeneratorVariableCostInUSDPerMWh'] = []
    Data['GeneratorType'] = []
    row = 1
    while True:
        try:
            col = 0
            if str(sh.cell_value(row, col)) == 'END':
                break
            else:
                col = 0; c = str(sh.cell_value(row, col)); Data['Generator'].append(c);
                col += 1; Data['GeneratorBusIndex'].append(Data['Bus'].index(str(sh.cell_value(row,col))));
                col += 1; Data['GeneratorPmax'].append(sh.cell_value(row,col));
                col += 1; Data['GeneratorPmin'].append(sh.cell_value(row,col));
                col += 1; Data['GeneratorQmax'].append(sh.cell_value(row,col));
                col += 1; Data['GeneratorQmin'].append(sh.cell_value(row,col));
                col += 1; Data['GeneratorRampUpInMWPerHour'].append(sh.cell_value(row,col));
                col += 0; Data['GeneratorRampDownInMWPerHour'].append(sh.cell_value(row,col));
                col += 1; Data['GeneratorSRampUpInMW'].append(sh.cell_value(row,col));
                col += 0; Data['GeneratorSRampDownInMW'].append(sh.cell_value(row,col));
                col += 1; Data['GeneratorMinUp'].append(sh.cell_value(row,col));
                col += 1; Data['GeneratorMinDw'].append(sh.cell_value(row,col));
                col += 1; Data['GeneratorInitS'].append(sh.cell_value(row,col));
                col += 1; Data['GeneratorInitP'].append(sh.cell_value(row,col));
                col += 1; Data['GeneratorStartUpCostInUSD'].append(sh.cell_value(row,col));
                col += 1; Data['GeneratorFixedCostInUSD'].append(sh.cell_value(row,col));
                col += 1; Data['GeneratorVariableCostInUSDPerMWh'].append(sh.cell_value(row,col));
                col += 1; Data['GeneratorType'].append(str(sh.cell_value(row,col)));
                row += 1
        except IndexError:
            break
        
    Data['SetOfGeneratorsAtBus'] = [[i for i in range(len(Data['Generator'])) if Data['GeneratorBusIndex'][i] == BusIndexx] for BusIndexx in range(len(Data['Bus']))]
    
    ### Read worksheet Lines ###
    sh = book.sheet_by_name('Lines')
    Data['Line'] = []
    Data['FromBusIndex'] = []
    Data['ToBusIndex'] = []
    Data['LineResistanceR'] = []
    Data['LineReactanceX'] = []
    Data['LineChargingB'] = []
    Data['LineMaxFlow'] = []
    Data['InverseReactance'] = []
    Data['Lineg'] = []
    Data['Lineb'] = []
    row = 1
    while True:
        try:
            col = 0
            if str(sh.cell_value(row, col)) == 'END':
                break
            else:
                col = 0; Data['Line'].append(str(sh.cell_value(row, col)));
                col += 1; Data['FromBusIndex'].append(Data['Bus'].index(str(sh.cell_value(row,col))));
                col += 1; Data['ToBusIndex'].append(Data['Bus'].index(str(sh.cell_value(row,col))));
                col += 1; Data['LineResistanceR'].append(sh.cell_value(row,col));
                col += 1; Data['LineReactanceX'].append(sh.cell_value(row,col));
                col += 1; Data['LineChargingB'].append(sh.cell_value(row,col));
                col += 1; Data['LineMaxFlow'].append(sh.cell_value(row,col));
                Data['InverseReactance'].append(1.0 / Data['LineReactanceX'][-1]);
                Data['Lineg'].append( Data['LineResistanceR'][-1] / ( Data['LineResistanceR'][-1]**2 + Data['LineReactanceX'][-1]**2));
                Data['Lineb'].append( -Data['LineReactanceX'][-1] / ( Data['LineResistanceR'][-1]**2 + Data['LineReactanceX'][-1]**2));
                row += 1
        except IndexError:
            break
        
    Data['SetOfLinesThatHaveThisBusAsFromBus'] = [[LineIndex for LineIndex in range(len(Data['Line'])) if Data['FromBusIndex'][LineIndex] == BusIndexx] for BusIndexx in range(len(Data['Bus']))]
    Data['SetOfLinesThatHaveThisBusAsToBus'] = [[LineIndex for LineIndex in range(len(Data['Line'])) if Data['ToBusIndex'][LineIndex] == BusIndexx] for BusIndexx in range(len(Data['Bus']))]
    
    Data['BusG'] = [Data['BusShuntConductance'][ip] / BaseMVA for ip in range(len(Data['Bus']))]
    Data['BusB'] = [Data['BusShuntSusceptance'][ip] / BaseMVA for ip in range(len(Data['Bus']))]
    Data['LineG'] = [0.0 for LineIndex in range(len(Data['Line']))]
    Data['LineB'] = [0.0 for LineIndex in range(len(Data['Line']))]
    for LineIndex in range(len(Data['Line'])):
        i = Data['FromBusIndex'][LineIndex]
        j = Data['ToBusIndex'][LineIndex]
        Data['BusG'][i] += Data['Lineg'][LineIndex]
        Data['BusB'][i] += Data['Lineb'][LineIndex] + 0.5 * Data['LineChargingB'][LineIndex]
        Data['BusG'][j] += Data['Lineg'][LineIndex]
        Data['BusB'][j] += Data['Lineb'][LineIndex] + 0.5 * Data['LineChargingB'][LineIndex]
        Data['LineG'][LineIndex] += -Data['Lineg'][LineIndex]
        Data['LineB'][LineIndex] += -Data['Lineb'][LineIndex]
        
    Data['LinegTilde'] = [0.0 for LineIndexPrime in range(len(Data['Line']))]
    Data['LinebTilde'] = [0.0 for LineIndexPrime in range(len(Data['Line']))]
    for LineIndex in range(len(Data['Line'])):
        i = Data['FromBusIndex'][LineIndex]
        j = Data['ToBusIndex'][LineIndex]
        Data['LinegTilde'][LineIndex] += Data['Lineg'][LineIndex]
        Data['LinebTilde'][LineIndex] += -(Data['Lineb'][LineIndex] + 0.5 * Data['LineChargingB'][LineIndex])
        
    Data['SetOfNeighborsOfBus'] = [[] for BusIndexx in range(len(Data['Bus']))]
    for BusIndex in range(len(Data['Bus'])):
        for LineIndex in Data['SetOfLinesThatHaveThisBusAsFromBus'][BusIndex]:
            Data['SetOfNeighborsOfBus'][BusIndex].append(Data['ToBusIndex'][LineIndex])
        for LineIndex in Data['SetOfLinesThatHaveThisBusAsToBus'][BusIndex]:
            Data['SetOfNeighborsOfBus'][BusIndex].append(Data['FromBusIndex'][LineIndex])
        
    Data['SetL'] = []
    Data['LineMaxFlowForPairOfBuses'] = [[[] for jp in range(len(Data['Bus']))] for ip in range(len(Data['Bus']))]
    for LineIndex in range(len(Data['Line'])):
        Data['SetL'].append([Data['FromBusIndex'][LineIndex], Data['ToBusIndex'][LineIndex]]);
        Data['SetL'].append([Data['ToBusIndex'][LineIndex], Data['FromBusIndex'][LineIndex]]);
        Data['LineMaxFlowForPairOfBuses'][Data['FromBusIndex'][LineIndex]][Data['ToBusIndex'][LineIndex]] = Data['LineMaxFlow'][LineIndex]
        Data['LineMaxFlowForPairOfBuses'][Data['ToBusIndex'][LineIndex]][Data['FromBusIndex'][LineIndex]] = Data['LineMaxFlow'][LineIndex]
        
    Data['G'] = [[[] for jp in range(len(Data['Bus']))] for ip in range(len(Data['Bus']))]
    Data['B'] = [[[] for jp in range(len(Data['Bus']))] for ip in range(len(Data['Bus']))]
    for i in range(len(Data['Bus'])):
        Data['G'][i][i] = Data['BusG'][i]
        Data['B'][i][i] = Data['BusB'][i]
    for LineIndex in range(len(Data['Line'])):
        iIndex = Data['FromBusIndex'][LineIndex]
        jIndex = Data['ToBusIndex'][LineIndex]
        Data['G'][iIndex][jIndex] = Data['LineG'][LineIndex]
        Data['G'][jIndex][iIndex] = Data['LineG'][LineIndex]
        Data['B'][iIndex][jIndex] = Data['LineB'][LineIndex]
        Data['B'][jIndex][iIndex] = Data['LineB'][LineIndex]
        
    ### Read worksheet Renewables ###
    sh = book.sheet_by_name('Renewables')
    Data['PowerAvailabilityProfile'] = [[0.0 for GeneratorIndex in range(len(Data['Generator']))] for t in range(len(Data['TimePeriodsSimulation']))]
    row = 2
    while True:
        col = 0
        if str(sh.cell_value(row, col)) == 'END':
            break
        else:
            col = 0; GeneratorIndex = Data['Generator'].index(str(sh.cell_value(row,col)));
            for t in range(len(Data['TimePeriodsSimulation'])):
                col = t+1; Data['PowerAvailabilityProfile'][t][GeneratorIndex] = sh.cell_value(row, col);
            row += 1
            
    ### Print some stuff ###
    if False and PrintOutput == True:
        print ''
        for BusIndex in range(len(Data['Bus'])):
            print 'Index of Bus with name', Data['Bus'][BusIndex],':', BusIndex
            
        print ''
        for BusIndex in range(len(Data['Bus'])):
            print 'Indices of Generators at Bus', BusIndex,':', Data['SetOfGeneratorsAtBus'][BusIndex]
            
        print ''
        for BusIndex in range(len(Data['Bus'])):
            print 'Buses that are neighbors of Bus', BusIndex,':', Data['SetOfNeighborsOfBus'][BusIndex]
            
        print ''
        for LineIndex in range(len(Data['Line'])):
            print 'FromBus and ToBus of Line', LineIndex,':', Data['FromBusIndex'][LineIndex], ',', Data['ToBusIndex'][LineIndex]
        
        print '\nSetL:', Data['SetL']
        print '\nLineMaxFlowForPairOfBuses:', Data['LineMaxFlowForPairOfBuses']
        print '\nAdmittance Matrix: G:', Data['G']
        print '\nAdmittance Matrix: B:', Data['B']
        
    if PrintOutput == True:
        TotalDemandProfile = [sum(Data['PdProfile'][t]) for t in range(len(Data['TimePeriodsSimulation']))]
        TotalWindPowerProfile = [sum([Data['PowerAvailabilityProfile'][t][i] for i in range(len(Data['Generator'])) if Data['GeneratorType'][i] == 'Wind']) for t in range(len(Data['TimePeriodsSimulation']))]
        TotalSolarPowerProfile = [sum([Data['PowerAvailabilityProfile'][t][i] for i in range(len(Data['Generator'])) if Data['GeneratorType'][i] == 'Solar']) for t in range(len(Data['TimePeriodsSimulation']))]
        TotalNetLoadProfile = [TotalDemandProfile[t]-TotalWindPowerProfile[t]-TotalSolarPowerProfile[t] for t in range(len(Data['TimePeriodsSimulation']))]
        print '\nProfiles for demand and renewables in the simulation'
        print 't, TotalDemandProfile, TotalWindPowerProfile, TotalSolarPowerProfile, TotalNetLoadProfile'
        for t in range(len(Data['TimePeriodsSimulation'])):
            print t, ',', TotalDemandProfile[t], ',', TotalWindPowerProfile[t], ',', TotalSolarPowerProfile[t], ',', TotalNetLoadProfile[t]
            
    return Data