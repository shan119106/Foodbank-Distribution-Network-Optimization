# -*- coding: utf-8 -*-
"""
Created on Thu Feb 24 22:43:38 2022

@author: Admin
"""

import pulp as p
import pandas as pd
import openpyxl
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
theta1 = [0,0.2,0.4,0.6,0.8,1]
theta2 = [0,0.2,0.4,0.6,0.8,1]
#Pc,Pnc,B,S,theta1,theta2 are iterated 
Dnc = {}
for i in Demand_id:
    Dnc[i] = Demand["Demand_poor_no_veh_people"][i-1]
Dc = {}
for i in Demand_id:
    Dc[i] = Demand["Demand_poor_with_vehicle"][i-1]
file = open("Outputtext.txt","w") 
index =0
array =["theta1","theta2","B","S","Pc","Pnc","Objective Value","eFB"]
for i in agency_id:
    array.append("eA" + str(i))
array.extend(["Cost1","Cost2","Cost3","Cost4","Cost5"])    
df = pd.DataFrame({str(index):array})  
print(df)
for a in Pc:
    for b in Pnc:
        for c in B:
          for d in S:
              for e in theta1:
                  for f in theta2:
                        problem = p.LpProblem("FoodBank",p.LpMinimize)
                        x = p.LpVariable.dicts("xij",[(i,j) for i in agency_id for j in Demand_id],lowBound =0,upBound =1,cat="Continuous")
                        eFB = p.LpVariable("eB",lowBound =0,cat ="Continuous")
                        eA = p.LpVariable.dicts("eA",agency_id,lowBound =0,cat="Continuous")
                        y = p.LpVariable.dicts("yi",agency_id,lowBound =0,cat = "Continuous")
                        xbar =p.LpVariable.dicts("xbar",Demand_id,lowBound =0,cat ="Continuous")
                        Sc = p.LpVariable.dicts("shortage",Demand_id,lowBound = 0,cat ="Continuous")
                        Snc = p.LpVariable.dicts("shortage_without_car",Demand_id,lowBound =0,cat ="Continuous")
                        #objective function
                        #constraint1
                        problem += eFB +p.lpSum(eA[i] for i in agency_id) <= c
                        #constraint2
                        for i in agency_id:
                            problem += y[i] <= q[i-1] + eA[i]
                        #constraint3
                        problem += p.lpSum([xbar[j] for j in Demand_id] +[y[i] for i in agency_id]) <= qFB + eFB
                        #constraint4
                        problem += p.lpSum([xbar[j] for j in Demand_id] +[y[i] for i in agency_id]) <= d
                        #constraint5
                        for i in agency_id:
                            problem += y[i] >= p.lpSum(Dc[j]*x[(i,j)]*30 for j in Demand_id)
                        #constraint6
                        for j in Demand_id:
                            problem += p.lpSum(30*Dc[j]*x[(i,j)] for i in agency_id) + Sc[j] == 30*Dc[j] 
                        #constraint7
                        for j in Demand_id:
                            problem += xbar[j] + Snc[j] == 30*Dnc[j]
                        #constraint 9a
                        for j in Demand_id:
                            for k in Demand_id:
                                if(j == k):
                                    pass
                                else:
                                    problem += (Sc[j]+Snc[j])/(Dc[j]+Dnc[j]) - ((Sc[k] + Snc[k])/(Dc[k]+Dnc[k])) <= e
                        #constarint 9b
                        for j in Demand_id:
                            for k in Demand_id:
                                if(j == k):
                                    pass
                                else:
                                    problem += (Sc[j]+Snc[j])/(Dc[j]+Dnc[j]) - ((Sc[k] + Snc[k])/(Dc[k]+Dnc[k])) >= -e
                                    
                        
                        #constraint 12a
                        for j in Demand_id:
                            problem += Sc[j]*Dnc[j] - Snc[j]*Dc[j] <= f*Dc[j]*Dnc[j]
                        #constarint 12
                        for j in Demand_id:
                            problem += (Sc[j]*Dnc[j]) - (Snc[j]*Dc[j]) >= -f*Dc[j]*Dnc[j]
                        #objective function
                        problem += p.lpSum( [CfbtoAg[i]*y[i] for i in agency_id] + [CfbtoDp[j]*xbar[j] for j in Demand_id] + [CDptoAg[i][j]*Dc[j]*x[(i,j)]*30 for i in agency_id for j in Demand_id] + [a*Sc[j] for j in Demand_id] + [b*Snc[j] for j in Demand_id])
                        #LPfile
                        
                        problem.writeLP("output",writeSOS=1, mip=1, max_length=100)
                        #Problem Solve
                        problem.solve()
                        u =0
                        file.write("theta1: "+ str(e) +" theta2: "+str(f) + " B: " + str(c) + " S: " +str(d)+ " Pc: " +str(a) + " Pnc: " + str(b) +"\n")
                        file.write("Objective Function:")
                        file.write(str(p.value(problem.objective))+"\n")
                        file.write("eFB:")
                        file.write(str(eFB.varValue)+"\n")
                        array1 =[]
                        array1 =[str(e),str(f),str(c),str(d),str(a),str(b),str(p.value(problem.objective)),str(eFB.varValue)]
                        for i in agency_id:
                            array1.append(str(eA[i].varValue))
                            file.write("eA" + str(i)  +" " + str(eA[i].varValue)+"\n")
                        for j in Demand_id:
                            for i in agency_id:
                                u += CDptoAg[i][j]*Dc[j]*(x[(i,j)].varValue)*30
                        array1.append(str(u))
                        file.write("Cost3: "+str(u)+"\n")
                        u =0
                        for j in Demand_id:
                            if(xbar[j].varValue != 0):
                                u += CfbtoDp[j]*xbar[j].varValue 
                        array1.append(str(u))
                        file.write("Cost2:  "+str(u)+"\n")
                        u=0
                        for i in agency_id:
                            if(y[i].varValue != 0):
                                u += CfbtoAg[i]*y[i].varValue 
                        array1.append(str(u))
                        file.write("Cost1: "+ str(u)+"\n")
                        u =0
                        v =0
                        for j in Demand_id:
                            u+= a*Snc[j].varValue 
                            v+= b*Snc[j].varValue
                        array1.append(str(u))
                        array1.append(str(v))
                        file.write("Cost4"+str(u)+"\n")
                        file.write("Cost5"+str(v)+ "\n")
                        index =index+1
                        df[str(index)] = array1

df.to_excel('./output_excel.xlsx',sheet_name= "Sheet1")
file.close()


 


