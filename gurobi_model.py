# -*- coding: utf-8 -*-
"""
Created on Sun Apr 24 10:33:18 2022

@author: Admin
"""
import gurobipy as gp
from gurobipy import GRB
import pandas as pd
Food_bank = pd.read_excel("Data_Food_Bank_copy.xlsx",sheet_name ="FoodBank")
agency = pd.read_excel("Data_Food_Bank_copy.xlsx",sheet_name ="Agency")
Demand = pd.read_excel("Data_Food_Bank_copy.xlsx",sheet_name ="DemandPoints")
FbtoAgencyDist = pd.read_excel("Data_Food_Bank_copy.xlsx",sheet_name ="FBtoAgencyDist")
FbtoDPdist = pd.read_excel("Data_Food_Bank_copy.xlsx",sheet_name ="FBtoDPdist")
DPtoAgencydist = pd.read_excel("Data_Food_Bank_copy.xlsx",sheet_name ="DPtoAgencydist")
agency_id = []
for i in agency["Agency_ID"]:
    agency_id.append(i)
Demand_id = []
for i in Demand["DP_ID"]:
    Demand_id.append(i)
CfbtoAg = {}
for i in agency_id:
    CfbtoAg[i] = 0.1*FbtoAgencyDist["Distance from FB (km)"][i-1]
CfbtoDp ={}
for i in Demand_id:
    CfbtoDp[i] =0.3*FbtoDPdist["distance from FB (km)"][i-1]
CDptoAg ={}
for i in agency_id:
    CDptoAg[i] ={}
    for j in Demand_id:
        CDptoAg[i][j] = 0.5*DPtoAgencydist["Agency_"+str(i)][j-1]
           
Pc= [100,500]
Pnc =[1000,5000]
qFB = 100000
B= [10000,50000]
q = [48000,40000,50000,42000,49000,51000,40000,43000,45000]
S = [100000,500000]
theta1 = [0.1,0.2,0.3,0.4,0.5,0.6,0.7,0.8,0.9]
theta2 = [0.1,0.2,0.3,0.4,0.5,0.6,0.7,0.8,0.9]
Dnc = {}
for i in Demand_id:
    Dnc[i] = Demand["Demand_poor_no_veh_people"][i-1]
Dc = {}
for i in Demand_id:
    Dc[i] = Demand["Demand_poor_with_vehicle"][i-1]
file = open("gurobiOutputtext.txt","w")
index =0
array =["theta1","theta2","B","S","Pc","Pnc","Objective Value","eFB"]
for i in agency_id:
    array.append("eA" + str(i))
array.extend(["Cost1","Cost2","Cost3","Cost4","Cost5"])    
df = pd.DataFrame({str(index):array})  
for a in Pc:
    for b in Pnc:
        for c in B:
          for d in S:
              for e in theta1:
                  for f in theta2:
                      problem = gp.Model("LpProblem")
                      #adding variables
                      x ={}
                      for i in agency_id:
                          x[i] = {}
                          for j in Demand_id:
                              x[i][j] = problem.addVar(lb=0,ub=1,vtype =GRB.CONTINUOUS,name="x"+str(i)+str(j))
                      eFB = problem.addVar(lb =0,vtype =GRB.CONTINUOUS,name ="eB")
                      eA={}
                      for i in agency_id:
                          eA[i] = problem.addVar(lb =0,vtype =GRB.CONTINUOUS,name="eA")
                      y ={}
                      for i in agency_id:
                          y[i]= problem.addVar(lb =0,vtype = GRB.CONTINUOUS,name ="y"+str(i))
                      xbar={}
                      for j in Demand_id:
                          xbar[j] = problem.addVar(lb =0,vtype = GRB.CONTINUOUS,name ="xbar")
                      Sc ={}
                      for j in Demand_id:
                          Sc[j] = (problem.addVar(lb =0,vtype =GRB.CONTINUOUS,name ="Sc"))
                      Snc ={}
                      for j in Demand_id:
                          Snc[j] = problem.addVar(lb =0,vtype =GRB.CONTINUOUS,name ="Snc")
                      #constraint1
                      problem.addConstr(eFB + gp.quicksum(eA[i] for i in agency_id) <= c)
                      #constraint2
                      problem.addConstrs(y[i] <= q[i-1] + eA[i] for i in agency_id)
                      #constraint3
                      problem.addConstr(gp.quicksum([xbar[j] for j in Demand_id] +[y[i] for i in agency_id]) <= qFB + eFB)
                      #constraint4
                      problem.addConstr(gp.quicksum([xbar[j] for j in Demand_id] +[y[i] for i in agency_id]) <= d)
                      #constraint5
                      problem.addConstrs(y[i] >= gp.quicksum(Dc[j]*x[i][j]*30 for j in Demand_id) for i in agency_id)
                      #constraint7
                      problem.addConstrs(xbar[j] + Snc[j] == 30*Dnc[j] for j in Demand_id)
                      #constraint9a
                      for j in Demand_id:
                          for k in Demand_id:
                              if(j == k):
                                  pass
                              else:
                                  problem.addConstr((Sc[j]+Snc[j]/Dc[j]+Dnc[j]) - (Sc[k] + Snc[k]/Dc[k]+Dnc[k]) <= e)
                      #constraint9b
                      for j in Demand_id:
                            for k in Demand_id:
                                if(j == k):
                                    pass
                                else:
                                    problem.addConstr(Sc[j]+Snc[j]/(Dc[j]+Dnc[j]) - (Sc[k] + Snc[k])/(Dc[k]+Dnc[k]) >= -e)
                      #constraint12a
                      #constraint 12a
                      for j in Demand_id:
                         problem.addConstr(Sc[j]*Dnc[j] - Snc[j]*Dc[j] <= f*Dc[j]*Dnc[j])
                      #constarint 12
                      for j in Demand_id:
                         problem.addConstr((Sc[j]*Dnc[j]) - (Snc[j]*Dc[j]) >= -f*Dc[j]*Dnc[j])
                      #objective function
                      problem.setObjective(gp.quicksum( [CfbtoAg[i]*y[i] for i in agency_id] + [CfbtoDp[j]*xbar[j] for j in Demand_id] + [CDptoAg[i][j]*Dc[j]*x[i][j]*30 for i in agency_id for j in Demand_id] + [a*Sc[j] for j in Demand_id] + [b*Snc[j] for j in Demand_id]),GRB.MAXIMIZE)
                      #problemsolve    
                      problem.write("gurobioutput.lp")
                      problem.optimize()
                      print(problem.objVal)
                      u =0
                      file.write("theta1: "+ str(e) +" theta2: "+str(f) + " B: " + str(c) + " S: " +str(d)+ " Pc: " +str(a) + " Pnc: " + str(b) +"\n")
                      file.write("Objective Function:")
                      file.write(str(problem.objVal)+"\n")
                      file.write("eFB:")
                      file.write(str(eFB.x)+"\n")
                      array1 =[str(e),str(f),str(c),str(d),str(a),str(b),str(problem.objVal),str(eFB.x)]
                      for i in agency_id:
                        array1.append(str(eA[i].x))
                        file.write("eA" + str(i)  +" " + str(eA[i].x)+"\n")
                      for j in Demand_id:
                        for i in agency_id:
                            u += CDptoAg[i][j]*Dc[j]*(x[i][j].x)*30
                      array1.append(str(u))
                      file.write("Cost3: "+str(u)+"\n")
                      u =0
                      for j in Demand_id:
                        if(xbar[j].x != 0):
                            u += CfbtoDp[j]*xbar[j].x 
                      array1.append(str(u))
                      file.write("Cost2:  "+str(u)+"\n")
                      u=0
                      for i in agency_id:
                        if(y[i].x != 0):
                            u += CfbtoAg[i]*y[i].x 
                      array1.append(str(u))
                      file.write("Cost1: "+ str(u)+"\n")
                      u =0
                      v =0
                      for j in Demand_id:
                        u+= a*Snc[j].x 
                        v+= b*Snc[j].x
                      array1.append(str(u))
                      array1.append(str(v))
                      file.write("Cost4"+str(u)+"\n")
                      file.write("Cost5"+str(v)+ "\n")
                      index =index+1
                      df[str(index)] = array1

df.to_excel('./output_excel.xlsx',sheet_name= "Sheet1")
file.close()
                      
                              