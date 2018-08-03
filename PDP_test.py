# -*- coding: utf-8 -*-
################################################################################
################################################################################
##########            Importamos los paquetes necesarios            ############
################################################################################ 
################################################################################

from __future__ import division
from copy import deepcopy
#Importamos el ambiente de PYOMO. 
from pyomo.environ import * 
#Paquetes para la resolucion del problema.
from pyomo.opt import SolverFactory, SolverStatus
#Paquete que nos permitira utilizar algunas funciones matematicas. 
import numpy as np
#Paquete que nos permite leer archivos Excel. 
import xlrd
#Definimos lo siguiente para resolver el modelo.
PrintSolverOutput=True 





################################################################################
################################################################################
##########   Preparamos la lectura de datos desde el archivo Excel  ############
################################################################################ 
################################################################################

#Creamos el objeto 'book' que contendra el libro del archico Excel.
book = xlrd.open_workbook('case14Data.xlsx')
#Creamos los objetos que contendran cada una de las paginas del libro.
GenData = book.sheet_by_name('GenData')
BusData = book.sheet_by_name('BusData')
BranchData = book.sheet_by_name('BranchData')
ActiveLoadData = book.sheet_by_name('ActiveLoadData')
ReactiveLoadData = book.sheet_by_name('ReactiveLoadData')
BESSData = book.sheet_by_name('BESSData')

#Inicializamos parametros que nos permitiran saber el numero de Generadores,
#Barras y Lineas que tendra el sistema
NumGen = 0
NumBus = 0
NumLines = 0
NumTime = 0
NumBESS = 0

#Obtenemos el numero total de Generadores, Barras y Lineas. Es importante que 
#el Excel tenga la palabra 'END' al finalizar las columnas donde se listan los 
#elementos del modelo.
while GenData.cell_value(NumGen+1,0)!='END':
    NumGen = NumGen+1
while BusData.cell_value(NumBus+1,0)!='END':
    NumBus = NumBus+1
while BranchData.cell_value(NumLines+1,0)!='END':
    NumLines = NumLines+1
while ActiveLoadData.cell_value(0,NumTime+1)!='END':
    NumTime = NumTime+1
while BESSData.cell_value(NumBESS+1,0)!='END':
    NumBESS = NumBESS+1





################################################################################
################################################################################
##########              Creamos el modelo y los SET                 ############
################################################################################    
################################################################################

#Creamos el objeto 'm', el cual sera usado para definir los SETS, Variables,
#parametros, y restricciones del modelo.
#Es importante destacar que todos los objetos del modelo comienzan con 'm.'.  
m=ConcreteModel()

#Definimos los SETs que usaremos en el modelo. Adicionalmente, leemos los datos 
#del Excel y los vamos agregando a los SETs creados.
#Esta es una primera manera de poder agregar datos a los sets, mas adelante se
#muestra otra forma

#SET de pasos de tiempo (En el ejemplo, seran: 1, 2, 3, 4).
m.TIME = Set() 
for i in range(NumTime):
    m.TIME.add(ActiveLoadData.cell_value(0,i+1))

#SET de generadores (En el ejemplo, seran: 'A', 'B' y 'C').
m.GEN = Set() 
for i in range(NumGen):
    m.GEN.add(GenData.cell_value(i+1,0))

m.BESS = Set() 
for i in range(NumBESS):
    m.BESS.add(BESSData.cell_value(i+1,0)) 

m.BUS = Set() #SET de barras (En el ejemplo, seran: 1, 2, 3, 4 y 5).
m.BUS_S = Set() #SET de barras subestaciones
m.BUS_DG = Set() #SET de barras de generacion distribuida
m.BUS_P = Set() #SET de barras de transferencia
m.BUS_AUX = Set(dimen=2)
# Tipo de nodo: (1) Subestación; (2) DG o Fuente de Reactivos; (3) Transferencia.
NumSub = 0
for i in range(NumBus):
    m.BUS.add(BusData.cell_value(i+1,0))
    if BusData.cell_value(i+1,5)==1: 
    	m.BUS_S.add(BusData.cell_value(i+1,0))
    	NumSub = NumSub+1
    if BusData.cell_value(i+1,5)==2: m.BUS_DG.add(BusData.cell_value(i+1,0))
    for j in range(NumTime): 
    	if BusData.cell_value(i+1,j+6)==3: m.BUS_AUX.add((BusData.cell_value(i+1,0),BusData.cell_value(0,j+6)))

m.BUS_P = Set(m.TIME,initialize=lambda m, t: set(
            z for z in m.BUS if (z,t) in m.BUS_AUX))

#SET de barras que cuentan con generadores 
#(En el ejemplo, seran: 1, 3 y 4).
m.BUS_W_GEN = Set()
for i in range(NumGen):
    m.BUS_W_GEN.add(GenData.cell_value(i+1,7))

#SET de barras que no cuentan con generadores 
#(En el ejemplo, seran: 2 y 5).
m.BUS_WO_GEN = Set() 
for i in m.BUS:
    if not i in m.BUS_W_GEN: m.BUS_WO_GEN.add(i)

#Aqui es importante destacar que cuando un SET se define como 'dimen=2',
#este se considera compuesto de tuplas.

#SET que contiene las lineas definidas en el archivo Excel 
#(En el ejemplo, seran las tuplas: (i,k)). 
m.BRANCHES = Set(dimen=2)
#SET que contiene todas posibles lineas de forma bidireccional.
#(En el ejemplo, seran las tuplas: (i,k) y (k,i)). 
m.BRANCHES_AUX = Set(dimen=2)
for i in range(NumLines):
    a=int(BranchData.cell_value(i+1,0))
    b=int(BranchData.cell_value(i+1,1))  
    m.BRANCHES.add((a,b))
    m.BRANCHES_AUX.add((a,b))
    m.BRANCHES_AUX.add((b,a))

#Esta es una forma distinta de generar SETs. Notese que estos son diferentes a 
#los creados anteriormente, ya que se definen en funcion de otros SETs.

#m.BUS_FROM[i] corresponde al SET de barras 'k' tal que la tupla (i,k)
#pertenece al SET de lineas definidas en el archivo Excel (i.e. m.BRANCHES).
# LOS WARNINGS QUE APARECEN EN EL MODELO ES POR QUE ESTA TRATANDO DE AGREGAR
# ELEMENTOS A LOS SETS QUE YA EXISTEN
m.BUS_FROM = Set()
m.BUS_FROM = Set(m.BUS,initialize=lambda m, i: set(
            z for z in m.BUS if (i,z) in m.BRANCHES))

m.BUS_TO = Set()
m.BUS_TO = Set(m.BUS,initialize=lambda m, i: set(
            z for z in m.BUS if (z,i) in m.BRANCHES))

#m.BUS_CON_TO[i] corresponde al SET de barras 'k' conectadas a la barra 'i'.
m.BUS_CON_TO = Set()
m.BUS_CON_TO = Set(m.BUS,initialize=lambda m, lz: set(
            z for z in m.BUS if (lz,z) in m.BRANCHES_AUX))

#SET que contiene las tuplas de generadores 'g' y barras asociadas 'i' (g,i).
m.GEN_AUX = Set(dimen=2)
for i in range(NumGen):
    a=GenData.cell_value(i+1,0)
    b=int(GenData.cell_value(i+1,7))
    m.GEN_AUX.add((a,b))

#m.GEN_IN_BUS[i] corresponde al SET de generadores 'g' conectadas a la barra 'i'.    
m.GEN_IN_BUS = Set() 
m.GEN_IN_BUS = Set(m.BUS,initialize=lambda m, i: 
	               set(g for g in m.GEN if (g,i) in m.GEN_AUX))

#SET que contiene las tuplas de BESS 'bess' y barras asociadas 'i' (bess,i).
m.BESS_AUX = Set(dimen=2)
for i in range(NumBESS):
    a=BESSData.cell_value(i+1,0)
    b=int(BESSData.cell_value(i+1,7))
    m.BESS_AUX.add((a,b))

#m.BESS_IN_BUS[i] corresponde al SET de BESS 'bess' conectadas a la barra 'i'.    
m.BESS_IN_BUS = Set() 
m.BESS_IN_BUS = Set(m.BUS,initialize=lambda m, i: 
	               set(g for g in m.BESS if (g,i) in m.BESS_AUX))










mu1=6 #Parametro de construccion de la primera aproximacion polihedral
mu2=6
#A continuacion se describen los SETs determinados por estos parametros
m.J1= Set(within=Integers)
m.J2= Set(within=Integers)
aux=0
while(aux<=mu1):
    m.J1.add(aux)
    aux = aux+1
aux=0
while(aux<=mu2):
    m.J2.add(aux)
    aux = aux+1

################################################################################
################################################################################
##########              Creamos los parametros del modelo           ############
################################################################################
################################################################################

#Datos BESS
Psto_max_init={}
Ppro_max_init={}
Qsto_max_init={}
Qpro_max_init={}
Csto_init={}
Cpro_init={}
Cinvbess_init={}
Pmaxbess_init={}
Qmaxbess_init={}

for i in range(NumBESS):
    a=BESSData.cell_value(i+1,0)
    Psto_max_init[a]=BESSData.cell_value(i+1,1)
    Ppro_max_init[a]=BESSData.cell_value(i+1,2)
    Qsto_max_init[a]=BESSData.cell_value(i+1,3)
    Qpro_max_init[a]=BESSData.cell_value(i+1,4)
    Csto_init[a]=BESSData.cell_value(i+1,5) 
    Cpro_init[a]=BESSData.cell_value(i+1,6)
    Cinvbess_init[a]=BESSData.cell_value(i+1,8)
    Pmaxbess_init[a]=BESSData.cell_value(i+1,10)
    Qmaxbess_init[a]=BESSData.cell_value(i+1,11)
 
m.Psto_max=Param(m.BESS, initialize=Psto_max_init, default = 0) 
m.Ppro_max=Param(m.BESS, initialize=Ppro_max_init, default = 0)
m.Qsto_max=Param(m.BESS, initialize=Qsto_max_init, default = 0)
m.Qpro_max=Param(m.BESS, initialize=Qpro_max_init, default = 0)
m.Csto=Param(m.BESS, initialize=Csto_init, default = 0)   
m.Cpro=Param(m.BESS, initialize=Cpro_init, default = 0) 
m.Cinvbess=Param(m.BESS, initialize=Cinvbess_init, default = 0)   
m.Pmaxbess=Param(m.BESS, initialize=Pmaxbess_init, default = 0) 
m.Qmaxbess=Param(m.BESS, initialize=Qmaxbess_init, default = 0)  

#Datos DG/SE
Pmin_init={}
Pmax_init={}
Qmin_init={}
Qmax_init={}
Clineal_init={}
Cquad_init={}
Cinv_init={}

for i in range(NumGen):
    a=GenData.cell_value(i+1,0)
    Pmin_init[a]=GenData.cell_value(i+1,1)
    Pmax_init[a]=GenData.cell_value(i+1,2)
    Qmin_init[a]=GenData.cell_value(i+1,3)
    Qmax_init[a]=GenData.cell_value(i+1,4)
    Clineal_init[a]=GenData.cell_value(i+1,5) 
    Cquad_init[a]=GenData.cell_value(i+1,6)
    Cinv_init[a]=GenData.cell_value(i+1,8)
 
m.Pmin=Param(m.GEN, initialize=Pmin_init, default = 0) 
m.Pmax=Param(m.GEN, initialize=Pmax_init, default = 0)
m.Qmin=Param(m.GEN, initialize=Qmin_init, default = 0)
m.Qmax=Param(m.GEN, initialize=Qmax_init, default = 0)
m.Clineal=Param(m.GEN, initialize=Clineal_init, default = 0)   
m.Cquad=Param(m.GEN, initialize=Cquad_init, default = 0) 
m.Cinv=Param(m.GEN, initialize=Cinv_init, default = 0)      

# Se inicializan los parametros asociados a las barras.
Pl_init={}
Ql_init={}
Vmin_init={}
Vmax_init={}
gS_init={}
bS_init={}

for i in range(NumBus):
    a=BusData.cell_value(i+1,0)
    Vmin_init[a]=BusData.cell_value(i+1,1)
    Vmax_init[a]=BusData.cell_value(i+1,2)
    gS_init[a]=BusData.cell_value(i+1,3)
    bS_init[a]=BusData.cell_value(i+1,4)
    for j in range(NumTime):
        t=ActiveLoadData.cell_value(0,j+1)
        Pl_init[a,t]=ActiveLoadData.cell_value(i+1,j+1)
        Ql_init[a,t]=ReactiveLoadData.cell_value(i+1,j+1)

m.Pl=Param(m.BUS,m.TIME, initialize=Pl_init, default = 0) 
m.Ql=Param(m.BUS,m.TIME, initialize=Ql_init, default = 0)
m.Vmin=Param(m.BUS, initialize=Vmin_init, default = 0)
m.Vmax=Param(m.BUS, initialize=Vmax_init, default = 0)
m.gS=Param(m.BUS, initialize=gS_init, default = 0)
m.bS=Param(m.BUS, initialize=bS_init, default = 0)

#Creamos los diccionarios de inicializacion de los parametros de las lineas.
rik_init={}
xik_init={}
gshik_init={}
bshik_init={}
tik_init={}
phiik_init={}
plimik_init={}
qlimik_init={}
slimik_init={}
#Cargamos los datos de los parametros de las lineas.
for i in range(NumLines):
    a=int(BranchData.cell_value(i+1,0))
    b=int(BranchData.cell_value(i+1,1))
    rik_init[a,b]=BranchData.cell_value(i+1,2)
    xik_init[a,b]=BranchData.cell_value(i+1,3)
    gshik_init[a,b]=BranchData.cell_value(i+1,4)
    bshik_init[a,b]=BranchData.cell_value(i+1,5)
    tik_init[a,b]=BranchData.cell_value(i+1,6)
    phiik_init[a,b]=BranchData.cell_value(i+1,7)
    plimik_init[a,b]=BranchData.cell_value(i+1,9)
    qlimik_init[a,b]=BranchData.cell_value(i+1,10)
    slimik_init[a,b]=BranchData.cell_value(i+1,11)

#Inicializamos los parametros correspondientes a las lineas.
m.rik=Param(m.BUS,m.BUS, initialize=rik_init, default = 0) 
m.xik=Param(m.BUS,m.BUS, initialize=xik_init, default = 0)
m.bshik=Param(m.BUS,m.BUS, initialize=bshik_init, default = 0)
m.tik=Param(m.BUS,m.BUS, initialize=tik_init, default = 0)
m.phiik=Param(m.BUS,m.BUS, initialize=phiik_init, default = 0)
m.plimik=Param(m.BUS,m.BUS, initialize=plimik_init, default = 0)
m.qlimik=Param(m.BUS,m.BUS, initialize=qlimik_init, default = 0)
m.slimik=Param(m.BUS,m.BUS, initialize=slimik_init, default = 0)

gik={}
bik={}
G_init={}
B_init={}

for i in m.BUS:         
    for j in m.BUS:
        gik[i,j]=0
        bik[i,j]=0
        G_init[i,j]=0
        B_init[i,j]=0

for i in m.BUS:          
    for j in m.BUS_FROM[i]:
        gik[i,j] = m.rik[i,j]/(m.rik[i,j]**2+m.xik[i,j]**2)
        bik[i,j] = -m.xik[i,j]/(m.rik[i,j]**2+m.xik[i,j]**2)

for i in m.BUS:
	G_init[i,i] += m.gS[i]
	B_init[i,i] += m.bS[i]
	for j in m.BUS_CON_TO[i]:
		if j in m.BUS_FROM[i]:
			G_init[i,i] += (1/m.tik[i,j]**2)*gik[i,j]
			B_init[i,i] += (1/m.tik[i,j]**2)*(bik[i,j]+(1/2)*m.bshik[i,j])
			G_init[i,j] = -(1/m.tik[i,j])*(gik[i,j]*cos(m.phiik[i,j])
                -bik[i,j]*sin(m.phiik[i,j]))
			B_init[i,j] = -(1/m.tik[i,j])*(gik[i,j]*sin(m.phiik[i,j])
                +bik[i,j]*cos(m.phiik[i,j]))
		if i in m.BUS_FROM[j]:
			G_init[i,i] += gik[j,i]
			B_init[i,i] += (bik[j,i]+(1/2)*m.bshik[j,i])
			G_init[i,j] = -(1/m.tik[j,i])*(gik[j,i]*cos(m.phiik[j,i])
                +bik[j,i]*sin(m.phiik[j,i]))
			B_init[i,j] = -(1/m.tik[j,i])*(-gik[j,i]*sin(m.phiik[j,i])
                +bik[j,i]*cos(m.phiik[j,i]))

m.G = Param(m.BUS,m.BUS, initialize = G_init)
m.B = Param(m.BUS,m.BUS, initialize = B_init)





################################################################################
################################################################################
##########         Creamos las Variables y Funcion Objetivo         ############
################################################################################
################################################################################

m.P = Var(m.GEN,m.TIME, within = NonNegativeReals)
m.Psto = Var(m.BESS, m.TIME, within=NonNegativeReals)
m.Ppro = Var(m.BESS, m.TIME, within=NonNegativeReals)
m.Qsto = Var(m.BESS, m.TIME, within=NonNegativeReals)
m.Qpro = Var(m.BESS, m.TIME, within=NonNegativeReals)
m.Pcont = Var(m.BESS, m.TIME, within=NonNegativeReals)
m.Qcont = Var(m.BESS, m.TIME, within=NonNegativeReals)
#Variables de Load Shedding Activo y Reactivo
m.LSP = Var(m.BUS,m.TIME, within = NonNegativeReals)
m.LSQ = Var(m.BUS,m.TIME, within = NonNegativeReals)
m.Xg = Var(m.GEN,m.TIME, within = Binary) ##Variable de Disponibilidad generadores##
m.Xl = Var(m.BUS,m.BUS,m.TIME, within = Binary) ##Variable de Disponibilidad lineas##
m.Xbess = Var(m.BESS,m.TIME, within = Binary)
m.Xstor = Var(m.BESS,m.TIME, within = Binary)
m.Y = Var(m.BUS,m.TIME, within = Binary) #Variable de Operación de barra de transferencia
m.kij = Var(m.BUS,m.BUS,m.TIME) #Variable flujo imaginario
m.Ki = Var(m.BUS,m.TIME) #Variable inyeccion de flujo imaginario
#Variables de bancos de capacitores
m.YCu = Var(m.BUS,m.TIME, within = NonNegativeIntegers) #Variable de Uso de cap
m.XCu = Var(m.BUS,m.TIME, within = NonNegativeIntegers) #Variable de Disponibilidad de cap
m.XCb = Var(m.BUS,m.TIME, within = Binary) #Variable Disponibilidad de Banco de cap
m.Q = Var(m.GEN,m.TIME)
m.Cii = Var(m.BUS,m.TIME)
m.Cij = Var(m.BUS,m.BUS,m.TIME)
m.Sij = Var(m.BUS,m.BUS,m.TIME)

#Variable auxiliar para la aproximacion lineal del SOCP
m.Yaux = Var(m.BUS,m.BUS,m.TIME)
m.Eta1 = Var(m.BUS,m.BUS,m.TIME,m.J1)
m.Xi1 = Var(m.BUS,m.BUS,m.TIME,m.J1)
m.Eta2 = Var(m.BUS,m.BUS,m.TIME,m.J2)
m.Xi2 = Var(m.BUS,m.BUS,m.TIME,m.J2)





################################################################################
################################################################################
##########################     Funcion Objetivo      ###########################
################################################################################
################################################################################

OFExpression = 0.0

for t in m.TIME:
	for b in m.BESS:
		OFExpression += m.Cpro[b]*m.Ppro[b,t] + m.Cinvbess[b]*m.Xbess[b,t] + m.Csto[b]*m.Psto[b,t]
	for i in m.GEN:
		OFExpression += m.Clineal[i]*m.P[i,t] + m.Cquad[i]*m.P[i,t]**2 + m.Cinv[i]*m.Xg[i,t]
	for i in m.BUS:
		OFExpression += m.LSP[i,t]*500000 + m.LSQ[i,t]*500000 + 100*m.XCb[i,t] + 20*m.XCu[i,t]
        for k in m.BUS:
            if k in m.BUS_CON_TO[i]:
                if k in m.BUS_FROM[i]:
                    OFExpression += m.Xl[i,k,t]*100

m.ObjectiveFunction = Objective(expr = OFExpression, sense =minimize)





################################################################################
################################################################################
##########     Comenzamos a crear las Restricciones del modelo      ############
################################################################################
################################################################################

#Restricciones de balance de potencia activa en cada una de las barras.
m.Active_Balance = Constraint(m.BUS,m.TIME, rule=lambda m, i, t: 
    - m.Pl[i,t] + m.LSP[i,t] 
    + sum(m.Ppro[b,t] for b in m.BESS_IN_BUS[i]) 
    - sum(m.Psto[b,t] for b in m.BESS_IN_BUS[i]) 
    + sum(m.P[g,t] for g in m.GEN_IN_BUS[i]) 
    - (m.G[i,i]*m.Cii[i,t] 
    + sum(m.G[i,j]*m.Cij[i,j,t]
      -m.B[i,j]*m.Sij[i,j,t] for j in m.BUS_CON_TO[i])) == 0)

#Restricciones de balance de potencia reactiva en cada una de las barras.
m.Reactive_Balance = Constraint(m.BUS,m.TIME, rule=lambda m, i, t: 
    - m.Ql[i,t] + m.LSQ[i,t]
    + 0.005*m.YCu[i,t]
	+ sum(m.Qpro[b,t] for b in m.BESS_IN_BUS[i]) 
    - sum(m.Qsto[b,t] for b in m.BESS_IN_BUS[i])
    + sum(m.Q[g,t] for g in m.GEN_IN_BUS[i]) 
    - (- m.B[i,i]*m.Cii[i,t] 
    + sum(-m.G[i,j]*m.Sij[i,j,t]
      -m.B[i,j]*m.Cij[i,j,t] for j in m.BUS_CON_TO[i])) == 0)

#Flujos bidireccionales vuelven necesaria la limitación de ambas direcciones.
def Pflow_max_rule(m,i,k,t):
	if k in m.BUS_CON_TO[i]:
		if k in m.BUS_FROM[i]:
			return m.plimik[i,k]*m.Xl[i,k,t] >= ((1/tik_init[i,k])**2)*(gik[i,k]+(1/2)*gshik_init[i,k])*m.Cii[i,t] + m.G[i,k]*m.Cij[i,k,t] -m.B[i,k]*m.Sij[i,k,t]
		else:
			return m.plimik[k,i]*m.Xl[k,i,t] >= (gik[k,i]+(1/2)*gshik_init[k,i])*m.Cii[i,t] +m.G[i,k]*m.Cij[i,k,t] - m.B[i,k]*m.Sij[i,k,t]
	else: return Constraint.Skip
m.Pflow_lim_max=Constraint(m.BUS,m.BUS,m.TIME, rule= Pflow_max_rule)

def Qflow_max_rule(m,i,k,t):
    if k in m.BUS_CON_TO[i]:
    	if k in m.BUS_FROM[i]:
    		return m.qlimik[i,k]*m.Xl[i,k,t] >= -((1/tik_init[i,k])**2)*(bik[i,k]+(1/2)*bshik_init[i,k])*m.Cii[i,t]-m.B[i,k]*m.Cij[i,k,t]-m.G[i,k]*m.Sij[i,k,t]
    	else:
    		return m.qlimik[k,i]*m.Xl[k,i,t] >= -(bik[k,i]+(1/2)*bshik_init[k,i])*m.Cii[i,t]-m.B[i,k]*m.Cij[i,k,t]-m.G[i,k]*m.Sij[i,k,t]
    else: return Constraint.Skip
m.Qflow_lim_max=Constraint(m.BUS,m.BUS,m.TIME, rule= Qflow_max_rule)

def Pflow_min_rule(m,i,k,t):
    if k in m.BUS_CON_TO[i]:
        if k in m.BUS_FROM[i]:
            return -m.plimik[i,k]*m.Xl[i,k,t] <= ((1/tik_init[i,k])**2)*(gik[i,k]+(1/2)*gshik_init[i,k])*m.Cii[i,t] + m.G[i,k]*m.Cij[i,k,t] -m.B[i,k]*m.Sij[i,k,t]
        else:
            return -m.plimik[k,i]*m.Xl[k,i,t] <= (gik[k,i]+(1/2)*gshik_init[k,i])*m.Cii[i,t] +m.G[i,k]*m.Cij[i,k,t] - m.B[i,k]*m.Sij[i,k,t]
    else: return Constraint.Skip
m.Pflow_lim_min=Constraint(m.BUS,m.BUS,m.TIME, rule= Pflow_min_rule)

def Qflow_min_rule(m,i,k,t):
    if k in m.BUS_CON_TO[i]:
        if k in m.BUS_FROM[i]:
            return -m.qlimik[i,k]*m.Xl[i,k,t] <= -((1/tik_init[i,k])**2)*(bik[i,k]+(1/2)*bshik_init[i,k])*m.Cii[i,t]-m.B[i,k]*m.Cij[i,k,t]-m.G[i,k]*m.Sij[i,k,t]
        else:
            return -m.qlimik[k,i]*m.Xl[k,i,t] <= -(bik[k,i]+(1/2)*bshik_init[k,i])*m.Cii[i,t]-m.B[i,k]*m.Cij[i,k,t]-m.G[i,k]*m.Sij[i,k,t]
    else: return Constraint.Skip
m.Qflow_lim_min=Constraint(m.BUS,m.BUS,m.TIME, rule= Qflow_min_rule)

#Restricciones de limites de operacion de los BESS dados por la variable de disponibilidad.
m.Ppro_max_xb = Constraint(m.BESS, m.TIME, rule=lambda m, b, t: m.Ppro_max[b]*m.Xbess[b,t]>=m.Ppro[b,t])
m.Psto_max_xb = Constraint(m.BESS, m.TIME, rule=lambda m, b, t: m.Psto_max[b]*m.Xbess[b,t]>=m.Psto[b,t])
m.Qpro_max_xb = Constraint(m.BESS, m.TIME, rule=lambda m, b, t: m.Qpro_max[b]*m.Xbess[b,t]>=m.Qpro[b,t])
m.Qsto_max_xb = Constraint(m.BESS, m.TIME, rule=lambda m, b, t: m.Qsto_max[b]*m.Xbess[b,t]>=m.Qsto[b,t])
m.Qpro_max_xb_log = Constraint(m.BESS, m.TIME, rule=lambda m, b, t: m.Qpro_max[b]*m.Xstor[b,t]>=m.Qpro[b,t])
m.Qsto_max_xb_log = Constraint(m.BESS, m.TIME, rule=lambda m, b, t: m.Qsto_max[b]*(1-m.Xstor[b,t])>=m.Qsto[b,t])
m.P_max_xb = Constraint(m.BESS, m.TIME, rule=lambda m, b, t: m.Pmaxbess[b]*m.Xbess[b,t]>=m.Pcont[b,t])
m.Q_max_xb = Constraint(m.BESS, m.TIME, rule=lambda m, b, t: m.Qmaxbess[b]*m.Xbess[b,t]>=m.Qcont[b,t])

#Restricciones de limites de operacion de los generadores dados por la variable de disponibilidad.
m.Pmin_Generation_xg = Constraint(m.GEN, m.TIME, rule=lambda m, g, t: m.Pmin[g]*m.Xg[g,t]<=m.P[g,t])
m.Pmax_Generation_xg = Constraint(m.GEN, m.TIME, rule=lambda m, g, t: m.Pmax[g]*m.Xg[g,t]>=m.P[g,t])
m.Qmin_Generation_xg = Constraint(m.GEN, m.TIME, rule=lambda m, g, t: m.Qmin[g]*m.Xg[g,t]<=m.Q[g,t])
m.Qmax_Generation_xg = Constraint(m.GEN, m.TIME, rule=lambda m, g, t: m.Qmax[g]*m.Xg[g,t]>=m.Q[g,t])

#Restricciones de limites de operacion de las barras.
m.Vmin_Level = Constraint(m.BUS, m.TIME, rule=lambda m, b, t: m.Vmin[b]**2<=m.Cii[b,t])
m.Vmax_Level = Constraint(m.BUS, m.TIME, rule=lambda m, b, t: m.Vmax[b]**2>=m.Cii[b,t])

m.Logica1 = Constraint(m.BUS,m.BUS, m.TIME, rule=lambda m, g, h, t: m.Cij[g,h,t] ==  m.Cij[h,g,t])
m.Logica2 = Constraint(m.BUS,m.BUS, m.TIME, rule=lambda m, g, h, t: m.Sij[g,h,t] == -m.Sij[h,g,t])














################################################################################
################################################################################
###################    APROXIMACION POLIHEDRAL DE L2-1     #####################
################################################################################
################################################################################

#m.Logica3 = Constraint(m.BUS,m.BUS, m.TIME, rule=lambda m, g, h, t: m.Sij[g,h,t]**2 + m.Cij[g,h,t]**2 <= m.Cii[g,t]*m.Cii[h,t])

#Restricciones para APROXIMACION POLIHEDRAL DE L2 para Xi1
def PolApproxL1_Xi1_a1b1_rule(m,i,k,t,j):
    if k in m.BUS_CON_TO[i]:
        if k in m.BUS_FROM[i]:
            if j==0: return (m.Xi1[i,k,t,j]>=m.Cij[i,k,t])
            if j>=1: return (m.Xi1[i,k,t,j]==cos(np.pi/(2**(j+1)))*m.Xi1[i,k,t,j-1] +sin(np.pi/(2**(j+1)))*m.Eta1[i,k,t,j-1])
        else: return Constraint.Skip
    else: return Constraint.Skip
m.PolApproxL1_Xi1_a1b1 = Constraint(m.BUS,m.BUS,m.TIME,m.J1, rule= PolApproxL1_Xi1_a1b1_rule)

def PolApproxL1_Xi1_a1neg_rule(m,i,k,t):
    if k in m.BUS_CON_TO[i]:
        if k in m.BUS_FROM[i]:
            return (m.Xi1[i,k,t,0]>= -m.Cij[i,k,t])
        else: return Constraint.Skip
    else: return Constraint.Skip
m.PolApproxL1_Xi1_a1neg = Constraint(m.BUS,m.BUS,m.TIME, rule= PolApproxL1_Xi1_a1neg_rule)

def PolApproxL1_Xi1_c1_rule(m,i,k,t):
    if k in m.BUS_CON_TO[i]:
        if k in m.BUS_FROM[i]:
            return (m.Xi1[i,k,t,mu1]<= m.Yaux[i,k,t])
        else: return Constraint.Skip
    else: return Constraint.Skip
m.PolApproxL1_Xi1_c1 = Constraint(m.BUS,m.BUS,m.TIME, rule= PolApproxL1_Xi1_c1_rule)

#Restricciones para APROXIMACION POLIHEDRAL DE L2 para Eta1
def PolApproxL1_Eta1_a1b1_rule(m,i,k,t,j):
    if k in m.BUS_CON_TO[i]:
        if k in m.BUS_FROM[i]:
            if j==0: return (m.Eta1[i,k,t,j]>=m.Sij[i,k,t])
            if j>=1: return (m.Eta1[i,k,t,j]>= -sin(np.pi/(2**(j+1)))*m.Xi1[i,k,t,j-1] + cos(np.pi/(2**(j+1)))*m.Eta1[i,k,t,j-1])
        else: return Constraint.Skip
    else: return Constraint.Skip
m.PolApproxL1_Eta1_a1b1= Constraint(m.BUS,m.BUS,m.TIME,m.J1, rule= PolApproxL1_Eta1_a1b1_rule)

def PolApproxL1_Eta1_a1b1neg_rule(m,i,k,t,j):
    if k in m.BUS_CON_TO[i]:
        if k in m.BUS_FROM[i]:
            if j==0: return (m.Eta1[i,k,t,j]>= -m.Sij[i,k,t])
            if j>=1: return (m.Eta1[i,k,t,j]>= sin(np.pi/(2**(j+1)))*m.Xi1[i,k,t,j-1] - cos(np.pi/(2**(j+1)))*m.Eta1[i,k,t,j-1])
        else: return Constraint.Skip
    else: return Constraint.Skip
m.PolApproxL1_Eta1_a1b1neg = Constraint(m.BUS,m.BUS,m.TIME,m.J1, rule= PolApproxL1_Eta1_a1b1neg_rule)

def PolApproxL1_Eta1_c1_rule(m,i,k,t):
    if k in m.BUS_CON_TO[i]:
        if k in m.BUS_FROM[i]:
            return (m.Eta1[i,k,t,mu1]<= tan(np.pi/(2**(mu1+1)))*m.Xi1[i,k,t,mu1])
        else: return Constraint.Skip
    else: return Constraint.Skip
m.PolApproxL1_Eta1_c1 = Constraint(m.BUS,m.BUS,m.TIME, rule= PolApproxL1_Eta1_c1_rule)


################################################################################
################################################################################
###################    APROXIMACION POLIHEDRAL DE L2-1     #####################
################################################################################
################################################################################

#Restricciones para APROXIMACION POLIHEDRAL DE L2 para Xi2
def PolApproxL2_Xi2_a2b2_rule(m,i,k,t,j):
    if k in m.BUS_CON_TO[i]:
        if k in m.BUS_FROM[i]:
            if j==0: return (m.Xi2[i,k,t,j]>=m.Yaux[i,k,t])
            if j>=1: return (m.Xi2[i,k,t,j]==cos(np.pi/(2**(j+1)))*m.Xi2[i,k,t,j-1] +sin(np.pi/(2**(j+1)))*m.Eta2[i,k,t,j-1])
        else: return Constraint.Skip
    else: return Constraint.Skip
m.PolApproxL2_Xi2_a2b2 = Constraint(m.BUS,m.BUS,m.TIME,m.J2, rule= PolApproxL2_Xi2_a2b2_rule)

def PolApproxL2_Xi2_a2neg_rule(m,i,k,t):
    if k in m.BUS_CON_TO[i]:
        if k in m.BUS_FROM[i]:
            return (m.Xi2[i,k,t,0]>= -m.Yaux[i,k,t])
        else: return Constraint.Skip
    else: return Constraint.Skip
m.PolApproxL2_Xi2_a2neg = Constraint(m.BUS,m.BUS,m.TIME, rule= PolApproxL2_Xi2_a2neg_rule)

def PolApproxL2_Xi2_c2_rule(m,i,k,t):
    if k in m.BUS_CON_TO[i]:
        if k in m.BUS_FROM[i]:
            return (m.Xi2[i,k,t,mu2]<= ((1/2)*(m.Cii[i,t]+m.Cii[k,t])) ) #AQUI X4
        else: return Constraint.Skip
    else: return Constraint.Skip
m.PolApproxL2_Xi2_c2 = Constraint(m.BUS,m.BUS,m.TIME, rule= PolApproxL2_Xi2_c2_rule)

#Restricciones para APROXIMACION POLIHEDRAL DE L2 para Eta2
def PolApproxL2_Eta2_a2b2_rule(m,i,k,t,j):
    if k in m.BUS_CON_TO[i]:
        if k in m.BUS_FROM[i]: 
            if j==0: return (m.Eta2[i,k,t,j]>= ((1/2)*(m.Cii[i,t]-m.Cii[k,t])) ) #AQUI X3
            if j>=1: return (m.Eta2[i,k,t,j]>= -sin(np.pi/(2**(j+1)))*m.Xi2[i,k,t,j-1] + cos(np.pi/(2**(j+1)))*m.Eta2[i,k,t,j-1])
        else: return Constraint.Skip
    else: return Constraint.Skip
m.PolApproxL2_Eta2_a2b2= Constraint(m.BUS,m.BUS,m.TIME,m.J2, rule= PolApproxL2_Eta2_a2b2_rule)

def PolApproxL2_Eta2_a2b2neg_rule(m,i,k,t,j):
    if k in m.BUS_CON_TO[i]:
        if k in m.BUS_FROM[i]:
            if j==0: return (m.Eta2[i,k,t,j]>= -((1/2)*(m.Cii[i,t]-m.Cii[k,t])) ) #AQUI X3
            if j>=1: return (m.Eta2[i,k,t,j]>= sin(np.pi/(2**(j+1)))*m.Xi2[i,k,t,j-1] - cos(np.pi/(2**(j+1)))*m.Eta2[i,k,t,j-1])
        else: return Constraint.Skip
    else: return Constraint.Skip
m.PolApproxL2_Eta2_a2b2neg = Constraint(m.BUS,m.BUS,m.TIME,m.J2, rule= PolApproxL2_Eta2_a2b2neg_rule)

def PolApproxL2_Eta2_c2_rule(m,i,k,t):
    if k in m.BUS_CON_TO[i]:
        if k in m.BUS_FROM[i]:
            return (m.Eta2[i,k,t,mu2]<= tan(np.pi/(2**(mu2+1)))*m.Xi2[i,k,t,mu2])
        else: return Constraint.Skip
    else: return Constraint.Skip
m.PolApproxL2_Eta2_c2 = Constraint(m.BUS,m.BUS,m.TIME, rule= PolApproxL2_Eta2_c2_rule)


























































#Restriccion logica de inversion.
def Xbesslogica_rule(m,b,t):
	if t == NumTime:
		return Constraint.Skip
	else: return (m.Xbess[b,t]<=m.Xbess[b,t+1])
m.Xbesslogica = Constraint(m.BESS,m.TIME, rule= Xbesslogica_rule)

def Xglogica_rule(m,g,t):
	if t == NumTime:
		return Constraint.Skip
	else: return (m.Xg[g,t]<=m.Xg[g,t+1])
m.Xglogica = Constraint(m.GEN,m.TIME, rule= Xglogica_rule)

def Xllogica_rule(m,i,k,t):
    if t == NumTime:
        return Constraint.Skip
    else:
        if k in m.BUS_CON_TO[i]:
            if k in m.BUS_FROM[i]: return (m.Xl[i,k,t]<=m.Xl[i,k,t+1])
            else: return Constraint.Skip
        else: return Constraint.Skip
m.Xllogica = Constraint(m.BUS,m.BUS,m.TIME, rule= Xllogica_rule)

# #Restricciones de radialidad
m.Radial13 = Constraint(m.TIME, rule=lambda m, t: 
	sum( sum(m.Xl[i,j,t] for j in m.BUS_FROM[i]) for i in m.BUS) == NumBus - NumSub - sum((1-m.Y[b,t]) for b in m.BUS_P[t]))

def Radial14_rule(m,i,j,t):
	if j in m.BUS_P[t]:
		if i in m.BUS_CON_TO[j]:
			if j in m.BUS_FROM[i]: return (m.Xl[i,j,t]<=m.Y[j,t]) 
			if i in m.BUS_FROM[j]: return (m.Xl[j,i,t]<=m.Y[j,t])
		else: return Constraint.Skip #Segundo if genera restricciones para aquellas lineas existentes
	else: return Constraint.Skip #Primer if asegura que la restricción se genera solo para barras de transferencia en t
m.Radial14 = Constraint(m.BUS,m.BUS,m.TIME, rule= Radial14_rule)

def Radial16_rule(m,j,t):
	if j in m.BUS_P[t]:
		return ( (sum((m.Xg[g,t]) for g in m.GEN_IN_BUS[j]))
				+(sum((m.Xbess[b,t]) for b in m.BESS_IN_BUS[j]))
			    +(sum((m.Xl[i,j,t]) for i in m.BUS_TO[j]))
			    +(sum((m.Xl[j,i,t]) for i in m.BUS_FROM[j])) >= 2*m.Y[j,t])
	else: return Constraint.Skip
m.Radial16 = Constraint(m.BUS,m.TIME, rule= Radial16_rule)

#En principio este conjunto de restricciones no es necesario utilizarlo para BESS, VAR, Bancos de Capacitores, etc..
#Esto porque no existen cargas puramente reactivas que permitan el islanding.. 
#De la misma forma, las baterias pueden considerarse inicialmente sin carga...

def Radial9_rule(m,i,t):
	return(sum(m.kij[j,i,t] for j in m.BUS_TO[i])-sum(m.kij[i,j,t] for j in m.BUS_FROM[i])==m.Ki[i,t])
m.Radial9 = Constraint(m.BUS,m.TIME, rule= Radial9_rule)

def Radial10_rule(m,i,t):
	if i in m.BUS_DG: return (m.Ki[i,t]==sum(m.Xg[g,t] for g in m.GEN_IN_BUS[i]))
	if i in m.BUS_S: return (m.Ki[i,t]>=-(NumGen-NumSub)*sum(m.Xg[g,t] for g in m.GEN_IN_BUS[i])) #Probando
	else: return (m.Ki[i,t]==0)
m.Radial10 = Constraint(m.BUS,m.TIME, rule= Radial10_rule)

def Radial12neg_rule(m,i,j,t):
	if j in m.BUS_FROM[i]: return (m.kij[i,j,t]<=(NumGen-NumSub)*m.Xl[i,j,t])
	else: return Constraint.Skip
m.Radial12neg = Constraint(m.BUS,m.BUS,m.TIME, rule=Radial12neg_rule)

def Radial12pos_rule(m,i,j,t):
	if j in m.BUS_FROM[i]: return (m.kij[i,j,t]>=-1*(NumGen-NumSub)*m.Xl[i,j,t])
	else: return Constraint.Skip
m.Radial12pos = Constraint(m.BUS,m.BUS,m.TIME, rule=Radial12pos_rule)

##Restricciones de BESS
efch=0.9
efdis=0.9
#InventarioP
def BESSinventarioP_rule(m,b,t):
    if t == 1:
    	return (m.Pcont[b,1]==0+efch*m.Psto[b,1]-(1/efdis)*m.Ppro[b,1]) #0 es lo almacenado inicialmente en 0
    else:
        return (m.Pcont[b,t]==m.Pcont[b,t-1]+efch*m.Psto[b,t]-(1/efdis)*m.Ppro[b,t])
m.BESSinventarioP = Constraint(m.BESS,m.TIME, rule= BESSinventarioP_rule)
#InventarioQ
def BESSinventarioQ_rule(m,b,t):
    if t == 1:
    	return (m.Qcont[b,1]==0+efch*m.Qsto[b,1]-(1/efdis)*m.Qpro[b,1]) #0 es lo almacenado inicialmente
    else:
        return (m.Qcont[b,t]==m.Qcont[b,t-1]+efch*m.Qsto[b,t]-(1/efdis)*m.Qpro[b,t])
m.BESSinventarioQ = Constraint(m.BESS,m.TIME, rule= BESSinventarioQ_rule)

#Bancos de capacitores
N = 3 #Numero maximo de capacitores del banco
def BancoCapacitoreslog1_rule(m,b,t):
	SumaXCb = 0.0
	for k in m.TIME:
		if int(k)<=int(t): SumaXCb += m.XCb[b,k]
	return m.XCu[b,t]<=N*SumaXCb
m.BancoCapacitoreslog1 = Constraint(m.BUS,m.TIME, rule = BancoCapacitoreslog1_rule)

m.BancoCapacitoreslog2 = Constraint(m.BUS, rule = lambda m, b: 
	sum(m.XCu[b,t] for t in m.TIME)<=N)

def BancoCapacitoreslog3_rule(m,b,t):
	SumaXCu = 0.0
	for k in m.TIME:
		if int(k)<=int(t): SumaXCu += m.XCu[b,k]
	return m.YCu[b,t]<=SumaXCu
m.BancoCapacitoreslog3 = Constraint(m.BUS,m.TIME, rule = BancoCapacitoreslog3_rule)

M = 30
m.BancoCapacitoreslog4 = Constraint(rule = lambda m: 
	sum(sum(m.XCu[b,t] for b in m.BUS) for t in m.TIME)<=M)











































################################################################################
################################################################################
##########       Entregamos el solver para resolver el modelo      #############
################################################################################
################################################################################

#m.dual = Suffix(direction=Suffix.IMPORT_EXPORT)
solver = 'gurobi'
if solver == 'gurobi':
    opt = SolverFactory(solver)
    #opt.options['mipgap'] =0.02
elif solver == 'ipopt':
    solver_io = 'nl'
    opt = SolverFactory(solver,solver_io=solver_io)
    opt.options['tol'] =1e-10
if opt is None:
    print("ERROR: Unable to create solver plugin for %s using the %s interface" 
    % (solver, solver_io))
    exit(1)
stream_solver = PrintSolverOutput # True prints solver output to screen.
keepfiles = False # True prints intermediate file names (.nl,.sol,...).
results = opt.solve(m,keepfiles=keepfiles,tee=stream_solver)
m.solutions.load_from(results) 

print '\n'
print 'Valor de Funcion Objetivo: '
print value(OFExpression)





################################################################################
################################################################################
##########       Creamos expresiones para determinar flujos        #############
################################################################################
################################################################################

sumaP=0
Pflow = {}
for i in m.BUS:
    for k in m.BUS_CON_TO[i]:
        for t in m.TIME:
            if k in m.BUS_FROM[i]:
                Pflow[i,k,t] = ((1/tik_init[i,k])**2)*(gik[i,k]+(1/2)*gshik_init[i,k])*m.Cii.get_values()[i,t] + m.G[i,k]*m.Cij.get_values()[i,k,t] -m.B[i,k]*m.Sij.get_values()[i,k,t]
            else:
                Pflow[i,k,t] = (gik[k,i]+(1/2)*gshik_init[k,i])*m.Cii.get_values()[i,t] +m.G[i,k]*m.Cij.get_values()[i,k,t] - m.B[i,k]*m.Sij.get_values()[i,k,t]
            sumaP = sumaP + Pflow[i,k,t]     

sumaQ=0
Qflow = {}
for i in m.BUS:
    for k in m.BUS_CON_TO[i]:
        for t in m.TIME:
            if k in m.BUS_FROM[i]:
                Qflow[i,k,t] = -((1/tik_init[i,k])**2)*(bik[i,k]+(1/2)*bshik_init[i,k])*m.Cii.get_values()[i,t]-m.B[i,k]*m.Cij.get_values()[i,k,t]-m.G[i,k]*m.Sij.get_values()[i,k,t]
            else:
                Qflow[i,k,t] = -(bik[k,i]+(1/2)*bshik_init[k,i])*m.Cii.get_values()[i,t]-m.B[i,k]*m.Cij.get_values()[i,k,t]-m.G[i,k]*m.Sij.get_values()[i,k,t]
            sumaQ = sumaQ + Qflow[i,k,t] 

Voltage = {}
for i in m.BUS:
    for t in m.TIME:
        Voltage[i,t] = m.Cii.get_values()[i,t]**(0.5)

Dif_Theta = {}
for i in m.BUS:
    for k in m.BUS_FROM[i]:
        for t in m.TIME:
            Dif_Theta [i,k,t] = np.arctan2(m.Sij.get_values()[i,k,t],m.Cij.get_values()[i,k,t])
            Dif_Theta [k,i,t] = -Dif_Theta [i,k,t]

Theta = {}
for i in m.BUS:
    for t in m.TIME:
        Theta[i,t]=0

clue = True
while clue:
    for i in m.BUS:
        for j in m.BUS_CON_TO[i]:
            for t in m.TIME:
                if not i==1:
                    if not Theta[i,t]==0:
                        Theta[j,t] = Dif_Theta[i,j,t]+Theta[i,t]
                else:
                    Theta[j,t] = Dif_Theta[i,j,t]+Theta[i,t]
    clue = False
    for i in m.BUS:
        for t in m.TIME:
            if not i==1:
                if Theta[i,t]==0: clue=True      





################################################################################
################################################################################
##########    Impresion de los resultados en una planilla Excel    #############
################################################################################
################################################################################                

###Impresion de los resultados en una planilla Excel.
import xlwt #Importamos el paquete de escritura en Excel.
outputbook = xlwt.Workbook() #Creamos un libro de salida.
#Creamos planillas dentro del libro.
GenOutput = outputbook.add_sheet('GenOutput') 
BusOutput = outputbook.add_sheet('BusOutput') 
FlowsOutput = outputbook.add_sheet('FlowsOutput')
BESSOutput = outputbook.add_sheet('BESSOutput')
CBOutput = outputbook.add_sheet('CBOutput')

#Impresion de CB.
for t in m.TIME:
    CBOutput.write(int((t-1)*(NumBus+5)),0,'Time: ')
    CBOutput.write(int((t-1)*(NumBus+5)),1,str(t))
    CBOutput.write(int((t-1)*(NumBus+5))+2,0,'Bus')
    CBOutput.write(int((t-1)*(NumBus+5))+2,1,'Instalacion de Banco')
    CBOutput.write(int((t-1)*(NumBus+5))+2,2,'Instalacion de Caps')
    CBOutput.write(int((t-1)*(NumBus+5))+2,3,'Numero de Caps utilizados')
    row=int((t-1)*(NumBus+5))+3
    for b in m.BUS:  
	   CBOutput.write(row,0,str(b))
	   CBOutput.write(row,1,float(m.XCb.get_values()[b,t]))
	   CBOutput.write(row,2,float(m.XCu.get_values()[b,t]))
	   CBOutput.write(row,3,float(m.YCu.get_values()[b,t]))
	   row +=1

#Impresion de BESS.
for t in m.TIME:
	BESSOutput.write(int((t-1)*(NumBESS+5)),0,'Time: ')
	BESSOutput.write(int((t-1)*(NumBESS+5)),1,str(t))
	BESSOutput.write(int((t-1)*(NumBESS+5))+2,0,'BESS')
	BESSOutput.write(int((t-1)*(NumBESS+5))+2,1,'Active Power Storage')
	BESSOutput.write(int((t-1)*(NumBESS+5))+2,2,'Active Power Production')
	BESSOutput.write(int((t-1)*(NumBESS+5))+2,3,'Active Power Almacenado')
	BESSOutput.write(int((t-1)*(NumBESS+5))+2,4,'Reactive Power Storage')
	BESSOutput.write(int((t-1)*(NumBESS+5))+2,5,'Reactive Power Production')
	BESSOutput.write(int((t-1)*(NumBESS+5))+2,6,'Reactive Power Almacenado')
	BESSOutput.write(int((t-1)*(NumBESS+5))+2,7,'Xbess')
	row=int((t-1)*(NumBESS+5))+3
	for b in m.BESS:
		BESSOutput.write(row,0,str(b))
		BESSOutput.write(row,1,float(m.Psto.get_values()[b,t]))
		BESSOutput.write(row,2,float(m.Ppro.get_values()[b,t]))
		BESSOutput.write(row,3,float(m.Pcont.get_values()[b,t]))
		BESSOutput.write(row,4,float(m.Qsto.get_values()[b,t]))
		BESSOutput.write(row,5,float(m.Qpro.get_values()[b,t]))
		BESSOutput.write(row,6,float(m.Qcont.get_values()[b,t]))
		BESSOutput.write(row,7,float(m.Xbess.get_values()[b,t]))
		row +=1

#Impresion de Potencias Generadas.
for t in m.TIME:
    GenOutput.write(int((t-1)*(NumGen+5)),0,'Time: ')
    GenOutput.write(int((t-1)*(NumGen+5)),1,str(t))
    GenOutput.write(int((t-1)*(NumGen+5))+2,0,'Generator')
    GenOutput.write(int((t-1)*(NumGen+5))+2,1,'Activer Power Generation')
    GenOutput.write(int((t-1)*(NumGen+5))+2,2,'Reactive Power Generation')
    GenOutput.write(int((t-1)*(NumGen+5))+2,3,'Operational Cost')
    GenOutput.write(int((t-1)*(NumGen+5))+2,4,'Marginal Cost')
    GenOutput.write(int((t-1)*(NumGen+5))+2,5,'Xg')
    row=int((t-1)*(NumGen+5))+3
    for g in m.GEN:
	   GenOutput.write(row,0,str(g))
	   GenOutput.write(row,1,float(m.P.get_values()[g,t]))
	   GenOutput.write(row,2,float(m.Q.get_values()[g,t])) 
	   GenOutput.write(row,3,float(m.Clineal[g]*m.P.get_values()[g,t] 
							    +m.Cquad[g]*(m.P.get_values()[g,t])**2)) 
	   GenOutput.write(row,4,float(m.Clineal[g] 
							    +2*m.Cquad[g]*m.P.get_values()[g,t]))
	   GenOutput.write(row,5,float(m.Xg.get_values()[g,t]))
	   row +=1

#Impresion de Voltajes, Angulos y Costos Marginales en barras.
for t in m.TIME:
    BusOutput.write(int((t-1)*(NumBus+5)),0,'Time: ')
    BusOutput.write(int((t-1)*(NumBus+5)),1,str(t))
    BusOutput.write(int((t-1)*(NumBus+5))+2,0,'Bus')
    BusOutput.write(int((t-1)*(NumBus+5))+2,1,'Voltage')
    BusOutput.write(int((t-1)*(NumBus+5))+2,2,'Angle')
    BusOutput.write(int((t-1)*(NumBus+5))+2,3,'ActiveLoadShedding')
    BusOutput.write(int((t-1)*(NumBus+5))+2,4,'ReactiveLoadShedding')
    BusOutput.write(int((t-1)*(NumBus+5))+1,5,'Transfer Node Used')
    BusOutput.write(int((t-1)*(NumBus+5))+1,6,'Ki inyeccion')
    #BusOutput.write(2,int((t-1)*5)+3,'Local Marginal Price')
    row=int((t-1)*(NumBus+5))+3
    for b in m.BUS:  
	   BusOutput.write(row,0,str(b))
	   BusOutput.write(row,1,float(Voltage[b,t]))
	   BusOutput.write(row,2,(Theta[b,t]*180/np.pi))
	   BusOutput.write(row,3,float(m.LSP.get_values()[b,t]))
	   BusOutput.write(row,4,float(m.LSQ.get_values()[b,t]))
	   BusOutput.write(row,5,(m.Y.get_values()[b,t]))
	   BusOutput.write(row,6,(m.Ki.get_values()[b,t]))
	   #BusOutput.write(row,3,float(m.dual[m.Active_Balance[b]]))  
	   row +=1

#Impresion de flujos en lineas.
for t in m.TIME:
    FlowsOutput.write(int((t-1)*(NumLines+5)),0,'Time: ')
    FlowsOutput.write(int((t-1)*(NumLines+5)),1,str(t))
    FlowsOutput.write(int((t-1)*(NumLines+5))+2,0,'Bus i')
    FlowsOutput.write(int((t-1)*(NumLines+5))+2,1,'Bus j')
    FlowsOutput.write(int((t-1)*(NumLines+5))+2,2,'Active Power Flow From i to j')
    FlowsOutput.write(int((t-1)*(NumLines+5))+2,3,'Reactive Power Flow From i to j')
    FlowsOutput.write(int((t-1)*(NumLines+5))+2,4,'Active Power Flow From j to i')
    FlowsOutput.write(int((t-1)*(NumLines+5))+2,5,'Reactive Power Flow From j to i')
    FlowsOutput.write(int((t-1)*(NumLines+5))+2,6,'Xl')
    FlowsOutput.write(int((t-1)*(NumLines+5))+2,7,'Kij Flow From i to j')
    FlowsOutput.write(int((t-1)*(NumLines+5))+2,8,'kij Flow From j to i')
    row=int((t-1)*(NumLines+5))+3
    for i in m.BUS:  
	   for k in m.BUS_FROM[i]:
		  FlowsOutput.write(row,0,str(i))
		  FlowsOutput.write(row,1,str(k))
		  FlowsOutput.write(row,2,Pflow[i,k,t])
		  FlowsOutput.write(row,3,Qflow[i,k,t])
		  FlowsOutput.write(row,4,Pflow[k,i,t])
		  FlowsOutput.write(row,5,Qflow[k,i,t])
		  FlowsOutput.write(row,6,int(m.Xl.get_values()[i,k,t]))
		  FlowsOutput.write(row,7,m.kij.get_values()[i,k,t])
		  FlowsOutput.write(row,8,m.kij.get_values()[k,i,t]) 
		  row +=1

outputbook.save('Resultados_PDP_full16.30LP.xls')


