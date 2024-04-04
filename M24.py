#Mathematical model formulated by Moein QAISARI HASAN ABADI
#Mathematical model coded and optimized with pyomo by Russell Sadeghi


import pyomo.environ as pyo
from pyomo.environ import *
from pyomo.opt import SolverFactory
import pandas as pd

df = pd.read_excel('G:\DataPowerCAR2023.xlsx',engine='openpyxl', sheet_name='Data',header=0, index_col=0)

Demand = (df['D'].iloc[:14]).values.tolist()
Price = (df['Price'].iloc[:14]).values.tolist()
print(Demand)
print(Price)
m = pyo.ConcreteModel("BlockChain")
#Set
m.I = pyo.RangeSet(1, 14)

d = {i+1: Demand[i] for i in range(0, 14)}
p = {i+1: Price[i] for i in range(0, 14)}

m.d = pyo.Param(m.I, initialize=d)
m.p = pyo.Param(m.I, initialize=p)


def bParam(m, i):
    return sum(Demand[i]/14 for i in range(0,14))
    #return  sum(Demand[i] for i in range(0,24))/24
m.Param1 = pyo.Param(m.I,initialize = bParam)

def cParam(m,i):
    return 1.5 * Demand[i-1]
m.Param2 = pyo.Param(m.I,initialize = cParam)

def tParam(m,i):
    return 0.5 * Price[i-1]
m.Param3 = pyo.Param(m.I,initialize = tParam)


m.D = pyo.Var(m.I, domain=pyo.PositiveIntegers)
m.e = pyo.Var(m.I, domain=pyo.Reals)

def obj(m):
    return sum(m.D[i]*m.p[i] for i in m.I) + sum(m.Param3[i]*(m.D[i] - m.Param1[i]) for i in m.I)
m.obj = pyo.Objective(rule=obj,sense = pyo.minimize)

def Constraint1(m,i):
    return m.D[i] >= m.d[i]/2- m.e[i]*(m.p[i]+m.Param3[i])
m.Co1 = pyo.Constraint(m.I, rule=Constraint1)

def Constraint2(m,i):
    return m.D[i] <= m.Param2[i]
m.Co2 = pyo.Constraint(m.I, rule=Constraint2)

def eParam(m, i):
    if i == m.I.first():
        return m.e[i] == (((m.d[14] - m.d[1])) / ((m.p[14] + m.Param3[14]) - (m.p[1] + m.Param3[1])) * ((m.p[14] + m.Param3[14] + m.p[1] + m.Param3[1]) / (m.d[14] + m.d[1])))
    return m.e[i] ==(((m.d[i-1] -m.d[i])) / ((m.p[i-1] + m.Param3[i-1]) - (m.p[i] + m.Param3[i])) * ((m.p[i-1] + m.Param3[i-1] + m.p[i] + m.Param3[i]) / (m.d[i-1] + m.d[i])))
m.Param4 = pyo.Constraint(m.I, rule=eParam)

opt = SolverFactory('ipopt', executable='C:\\Ipopt\\bin\\ipopt.exe')
opt.solve(m)

# Print the results
m.pprint()

print('-----------------------------------------------------------')
for i in m.I:
    print('D Period: {0}, Prod. Amount: {1}'.format(i, pyo.value(m.D[i])))
print('-------------------------Optimal----------------------------------')

Optimal = pyo.value(obj(m))
print("Optimal Answer is :",Optimal)