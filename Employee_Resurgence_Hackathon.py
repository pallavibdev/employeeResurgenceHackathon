import os
import pandas as pd
import numpy as np

def calcZoneScore(Zone):
    score=0
    if(Zone.__eq__("Green")):
        score =0.3 *3
    elif(Zone.__eq__("Orange")):
        score =0.3 *2
    elif(Zone.__eq__("Red")):
        score =0.3 *1
    elif(Zone.__eq__("Containment")):
        score =0.3 *0
    return round(score,4)
def calcMedConScore(MedCon):
    score=0
    if(MedCon.__eq__("Healthy")):
        score =0.25 *1
    elif(MedCon.__eq__("Susceptible")):
        score =0.25 *0
    return round(score,4)
def calcEmpCritcalityScore(EmpCritcality):
    score=0
    if(EmpCritcality.__eq__("Y")):
        score =0.2 *1
    elif(EmpCritcality.__eq__("N")):
        score =0.2 *0
    return round(score,4)
def calcProximityScore(Proximity):
    score=0
    if(Proximity.__eq__("Closest to office")):
        score =0.1 *2
    elif(Proximity.__eq__("Average distant")):
        score =0.1 *1
    elif(Proximity.__eq__("Most distant")):
        score =0.1 *0
    return round(score,4)
def calcCommuteScore(PubTrans):
    score=0
    if(PubTrans.__eq__("Private transport")):
        score =0.1 *2
    elif(PubTrans.__eq__("Office  transport")):
        score =0.1 *1
    elif(PubTrans.__eq__("Public  transport")):
        score =0.1 *0
    return round(score,2)
def calcEmpPrefScore(EmpPref):
    score=0
    if(EmpPref.__eq__("Comfortable")):
        score =0.05 *1
    elif(EmpPref.__eq__("Not Comfortable")):
        score =0.05 *0
    return round(score,4)
def calculateEligiblity(src_file,trg_file):
    df= pd.read_excel(src_file)
    df['Zone Score'] =df.Zone.apply(calcZoneScore)
    df['Medical Condition Score'] =df.Medical_Condition.apply(calcMedConScore)
    df['Employee Criticality Score'] =df.Employee_Criticality.apply(calcEmpCritcalityScore)
    df['Proximity Score'] =df.Proximity.apply(calcProximityScore)
    df['Commute Score'] =df.Commute.apply(calcCommuteScore)
    df['Employee Preference Score'] =df.Employee_Preference.apply(calcEmpPrefScore)
    df.loc[df['Zone Score'] == 0,'Total Score'] =0
    df.loc[df['Zone Score'] > 0,'Total Score'] = round( (df['Zone Score'] + df['Medical Condition Score'] + df['Employee Criticality Score'] + df['Proximity Score'] + df['Commute Score'] + df['Employee Preference Score']),4)   
    #df['Total Score'].rank(ascending=0)
    df.sort_values('Total Score', inplace = True, ascending=False)
    df['Rank'] = np.arange(len(df))+1
    score_max = df['Total Score'].max()
    score_min = df['Total Score'].min()
    score_avg = round((score_max + score_min )/2)
    df.loc[df['Total Score'] >= score_avg,'Eligble To ROTA']='Y'
    df.loc[df['Total Score'] < score_avg,'Eligble To ROTA'] ='N'
    if(os.path.isfile(trg_file)):
        os.remove(trg_file)
    df.to_excel(trg_file, index =False)
	
###################################DRIVER CODE#############################################    
loc= input("Enter File Path : ") +"\\"
src_file = loc +  input("Enter Source Excel File Name : ")
trg_file= loc +  input("Enter Target Excel File Name : ")
calculateEligiblity(src_file,trg_file)
print("Employee Resurgence File available at- " , trg_file)
