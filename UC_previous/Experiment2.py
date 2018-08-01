import Models
import pickle

FileName = 'Case014.xlsx';
#FileName = 'Case118.xlsx';

Data = pickle.load(open('Data' + FileName + '.pkl','rb'))

DispatchSolution = Models.SolveDCOPF(Data)

print 'Pg:', DispatchSolution['Pg'][0]
