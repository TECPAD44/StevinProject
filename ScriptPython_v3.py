# -*- coding: utf-8 -*-
"""
Éditeur de Spyder

Ceci est un script temporaire.
"""
import pandas as pd;
#print("Quelle année ?");
#File=input();
for a in range(2020, 2021):
    File=str(a);
    print("Traitement de l'année "+File+"\n");
    File="DonneesEolienBelge"+File+".xlsx";    
    print("Lecture du fichier "+File);
    
    Inshore=pd.read_excel(File, usecols="A:D", header=7, na_values="-");
    Offshore=pd.read_excel(File, usecols="A,E:G", header=7, na_values="-");
    
    print("Construction de la série des erreurs");
    
    Inshore["DayAheadError"]=Inshore["Current"]-Inshore["Day Ahead"]/Inshore["Current"];
    Offshore["DayAheadError.1"]=Offshore["Current.1"]-Offshore["Day Ahead.1"]/Offshore["Current.1"];
    
    Inshore["IntradayError"]=Inshore["Current"]-Inshore["Intraday"]/Inshore["Current"];
    Offshore["IntradayError.1"]=Offshore["Current.1"]-Offshore["Intraday.1"]/Offshore["Current.1"];
    
    
    print("Calcul de scores basiques");
    
    
    InshoreMeanErrorDA=Inshore["DayAheadError"].mean();
    InshoreMedianErrorDA=Inshore["DayAheadError"].median();
    InshoreRMSEDA=(Inshore["DayAheadError"]**2).mean()**0.5;
    OffShoreMeanErrorDA=Offshore["DayAheadError.1"].mean();
    OffShoreMedianErrorDA=Offshore["DayAheadError.1"].median();
    OffshoreRMSEDA=(Offshore["DayAheadError.1"]**2).mean()**0.5;
    
    InshoreMeanErrorID=Inshore["IntradayError"].mean();
    InshoreMedianErrorID=Inshore["IntradayError"].median();
    InshoreRMSEID=(Inshore["IntradayError"]**2).mean()**0.5;
    OffshoreMeanErrorID=Offshore["IntradayError.1"].mean();
    OffshoreMedianErrorID=Offshore["IntradayError.1"].median();
    OffshoreRMSEID=(Offshore["IntradayError.1"]**2).mean()**0.5;
    
    print("Scores calculés\n");
    
    print("Calcul des modèles de climatologie et de persistence");
    
    
    Inshore["InshorePersistence"]=pd.Series([Inshore["Current"][0]]).append(Inshore["Current"], ignore_index=True).drop(Inshore["Current"].size);
    Inshore["InshorePersistenceError"]=Inshore["Current"]-Inshore["InshorePersistence"];
    Offshore["OffshorePersistence"]=pd.Series([Offshore["Current.1"][0]]).append(Offshore["Current.1"], ignore_index=True).drop(Offshore["Current.1"].size);
    Offshore["OffshorePersistenceError"]=Offshore["Current.1"]-Offshore["OffshorePersistence"];
    
    Inshore["InshoreClimatology"]=Inshore["Current"].mean();
    Inshore["InshoreClimatologyError"]=Inshore["Current"]-Inshore["InshoreClimatology"]
    Offshore["OffshoreClimatology"]=Offshore["Current.1"].mean();
    Offshore["OffshoreClimatologyError"]=Offshore["Current.1"]-Offshore["OffshoreClimatology"];
    
    InshorePersistenceMeanError=Inshore["InshorePersistenceError"].mean();
    InshorePersistenceMedianError=Inshore["InshorePersistenceError"].median();
    InshorePersistenceRMSE=(Inshore["InshorePersistenceError"]**2).mean()**0.5;
    OffshorePersistenceMeanError=Offshore["OffshorePersistenceError"].mean();
    OffshorePersistenceMedianError=Offshore["OffshorePersistenceError"].median();
    OffshorePersistenceRMSE=(Offshore["OffshorePersistenceError"]**2).mean()**0.5;
    
    InshoreClimatologyMeanError=Inshore["InshoreClimatologyError"].mean();
    InshoreClimatologyMedianError=Inshore["InshoreClimatologyError"].median();
    InshoreClimatologyRMSE=(Inshore["InshoreClimatologyError"]**2).mean()**0.5;
    OffshoreClimatologyMeanError=Offshore["OffshoreClimatologyError"].mean();
    OffshoreClimatologyMedianError=Offshore["OffshoreClimatologyError"].median();
    OffshoreClimatologyRMSE=(Offshore["OffshoreClimatologyError"]**2).mean()**0.5;
    
    print("Modèles établis\n");
    
    
    print("Identification des plages de lignes correspondantes aux jours et aux mois");
    
    
    LignesMois=[0]; #Contient les numéros de ligne correspondants aux changements de mois
    LignesJours=[[0]]; #Contient la liste des listes des numéros de liste associés à la ligne de début de chaque jour de chaque mois
    m=0; #Mois en cours d'écriture
    k='01'; #identifiant du mois en cours d'écriture
    n=0; 
    x=Inshore["Unnamed: 0"][n];
    for n in range(Inshore["Unnamed: 0"].size):
        if pd.isnull(x) and pd.isnull(Inshore["Unnamed: 0"][n+1]): #Si la ligne est comprise entre deux lignes vide
            l=Inshore["Unnamed: 0"][n][3:5] #on acquiert les chiffres du mois
            if k!=l: #si le mois a changé
                LignesMois.append(n); #on ajoute le numéro de la ligne
                k=l;
                m+=1;
                LignesJours.append([n]);
            else:
                LignesJours[m].append(n);
        x=Inshore["Unnamed: 0"][n];
        
    print("Plages de données identifiées\n");
    
    
    print("Construction d'une liste des séries de mois et des séries de jours");
    
    InshoreMois=[];
    InshoreJours=[];
    k=0; #compteur de jours
    
   
    #Initialisation
   
    InshoreMois.append(pd.DataFrame({"Hour": Inshore["Unnamed: 0"][LignesMois[0]:LignesMois[1]-1], "Day": Inshore["Unnamed: 0"][LignesMois[0]:LignesMois[1]-1]}));
    InshoreMois[0]["Current"]=Inshore["Current"][LignesMois[0]:LignesMois[1]-1];
    InshoreMois[0]["Intraday"]=Inshore["Intraday"][LignesMois[0]:LignesMois[1]-1];
    InshoreMois[0]["Day Ahead"]=Inshore["Day Ahead"][LignesMois[0]:LignesMois[1]-1];
    InshoreMois[0]["Hour"][0:24]=str(1)+"."+str(1)+"."+InshoreMois[0]["Hour"][0:24];#Ajout du jour à la ligne de temps du mois
    InshoreMois[0]["Day"][0:24]=str(1)+"."+str(1);
    for j in range(1,len(LignesJours[0])-1): #Construction de la liste de séries des jours
        InshoreJours.append(pd.DataFrame({"Hour": Inshore["Unnamed: 0"][LignesJours[0][j]+2:LignesJours[0][j+1]-1]}));#On initialise la liste avec le dataframe avec la colonne "Hour"
        InshoreJours[k]["Current"]=Inshore["Current"][LignesJours[0][j]+2:LignesJours[0][j+1]-1];
        InshoreJours[k]["Intraday"]=Inshore["Intraday"][LignesJours[0][j]+2:LignesJours[0][j+1]-1];
        InshoreJours[k]["Day Ahead"]=Inshore["Day Ahead"][LignesJours[0][j]+2:LignesJours[0][j+1]-1];
        InshoreMois[0]=InshoreMois[0].drop([LignesJours[0][j]-1,LignesJours[0][j],LignesJours[0][j]+1]);#Retrait des lignes vides
        InshoreMois[0]["Hour"][24*j:24*(j+1)]=str(j+1)+"."+str(1)+"."+InshoreMois[0]["Hour"][24*j:24*(j+1)];#Ajout du jour à la ligne de temps du mois
        InshoreMois[0]["Day"][24*j:24*(j+1)]=str(j+1)+"."+str(1);
        k+=1;
    InshoreJours.append(pd.DataFrame({"Hour": Inshore["Unnamed: 0"][LignesJours[0][len(LignesJours[0])-1]+2:LignesMois[1]-1]})); #Traitement de la fin de la série des jours du mois en cours d'écriture
    InshoreJours[k]["Current"]=Inshore["Current"][LignesJours[0][len(LignesJours[0])-1]+2:LignesMois[1]-1];
    InshoreJours[k]["Intraday"]=Inshore["Intraday"][LignesJours[0][len(LignesJours[0])-1]+2:LignesMois[1]-1];
    InshoreJours[k]["Day Ahead"]=Inshore["Day Ahead"][LignesJours[0][len(LignesJours[0])-1]+2:LignesMois[1]-1];
    InshoreMois[0]=InshoreMois[0].drop([LignesJours[0][len(LignesJours[0])-1]-1,LignesJours[0][len(LignesJours[0])-1],LignesJours[0][len(LignesJours[0])-1]+1]);
    InshoreMois[0]["Hour"][24*(len(LignesJours[0])-1):24*len(LignesJours[0])]=str(len(LignesJours[0]))+"."+str(1)+"."+InshoreMois[0]["Hour"][24*(len(LignesJours[0])-1):24*len(LignesJours[0])];
    InshoreMois[0]["Day"][24*(len(LignesJours[0])-1):24*len(LignesJours[0])]=str(len(LignesJours[0]))+"."+str(1)
    #Fin initialisation
    
    for i in range(1,len(LignesMois)-1): #Construction de la liste de séries des mois
        InshoreMois.append(pd.DataFrame({"Hour": Inshore["Unnamed: 0"][LignesMois[i]+2:LignesMois[i+1]-1], "Day": Inshore["Unnamed: 0"][LignesMois[i]+2:LignesMois[i+1]-1]}));
        InshoreMois[i]["Current"]=Inshore["Current"][LignesMois[i]+2:LignesMois[i+1]-1];
        InshoreMois[i]["Intraday"]=Inshore["Intraday"][LignesMois[i]+2:LignesMois[i+1]-1];
        InshoreMois[i]["Day Ahead"]=Inshore["Day Ahead"][LignesMois[i]+2:LignesMois[i+1]-1];
        for j in range(len(LignesJours[i])-1): #Construction de la liste de séries des jours
            InshoreJours.append(pd.DataFrame({"Hour": Inshore["Unnamed: 0"][LignesJours[i][j]+2:LignesJours[i][j+1]-1]}));
            InshoreJours[k]["Current"]=Inshore["Current"][LignesJours[i][j]+2:LignesJours[i][j+1]-1];
            InshoreJours[k]["Intraday"]=Inshore["Intraday"][LignesJours[i][j]+2:LignesJours[i][j+1]-1];
            InshoreJours[k]["Day Ahead"]=Inshore["Day Ahead"][LignesJours[i][j]+2:LignesJours[i][j+1]-1];
            if j!=0:
                InshoreMois[i]=InshoreMois[i].drop([LignesJours[i][j]-1,LignesJours[i][j],LignesJours[i][j]+1]);
            InshoreMois[i]["Hour"][24*j:24*(j+1)]=str(j+1)+"."+str(i+1)+"."+InshoreMois[i]["Hour"][24*j:24*(j+1)];#Ajout du jour à la ligne de temps du mois
            InshoreMois[i]["Day"][24*j:24*(j+1)]=str(j+1)+"."+str(i+1);
            k+=1;
        InshoreJours.append(pd.DataFrame({"Hour": Inshore["Unnamed: 0"][LignesJours[i][len(LignesJours[i])-1]+2:LignesMois[i+1]-1]})); #Traitement de la fin de la série des jours du mois en cours d'écriture
        InshoreJours[k]["Current"]=Inshore["Current"][LignesJours[i][len(LignesJours[i])-1]+2:LignesMois[i+1]-1];
        InshoreJours[k]["Intraday"]=Inshore["Intraday"][LignesJours[i][len(LignesJours[i])-1]+2:LignesMois[i+1]-1];
        InshoreJours[k]["Day Ahead"]=Inshore["Day Ahead"][LignesJours[i][len(LignesJours[i])-1]+2:LignesMois[i+1]-1];
        InshoreMois[i]=InshoreMois[i].drop([LignesJours[i][len(LignesJours[i])-1]-1,LignesJours[i][len(LignesJours[i])-1],LignesJours[i][len(LignesJours[i])-1]+1]);
        InshoreMois[i]["Hour"][24*(len(LignesJours[i])-1):24*len(LignesJours[i])]=str(len(LignesJours[i]))+"."+str(i+1)+"."+InshoreMois[i]["Hour"][24*(len(LignesJours[i])-1):24*len(LignesJours[i])];        
        InshoreMois[i]["Day"][24*(len(LignesJours[i])-1):24*len(LignesJours[i])]=str(len(LignesJours[i]))+"."+str(i+1);
        
    InshoreMois.append(pd.DataFrame({"Hour": Inshore["Unnamed: 0"][LignesMois[11]+2:Inshore["Unnamed: 0"].size], "Day": Inshore["Unnamed: 0"][LignesMois[11]+2:Inshore["Unnamed: 0"].size]}));#Traitement de la fin de la liste des mois (i.e. décembre)
    InshoreMois[11]["Current"]=Inshore["Current"][LignesMois[11]+2:Inshore["Unnamed: 0"].size];
    InshoreMois[11]["Intraday"]=Inshore["Intraday"][LignesMois[11]+2:Inshore["Unnamed: 0"].size];
    InshoreMois[11]["Day Ahead"]=Inshore["Day Ahead"][LignesMois[11]+2:Inshore["Unnamed: 0"].size];
    for j in range(len(LignesJours[11])-1): #Construction de la liste de séries des jours de décembre
            InshoreJours.append(pd.DataFrame({"Hour": Inshore["Unnamed: 0"][LignesJours[11][j]+2:LignesJours[11][j+1]-1]}));
            InshoreJours[k]["Current"]=Inshore["Current"][LignesJours[11][j]+2:LignesJours[11][j+1]-1];
            InshoreJours[k]["Intraday"]=Inshore["Intraday"][LignesJours[11][j]+2:LignesJours[11][j+1]-1];
            InshoreJours[k]["Day Ahead"]=Inshore["Day Ahead"][LignesJours[11][j]+2:LignesJours[11][j+1]-1];
            InshoreMois[11]["Hour"][24*j:24*(j+1)]=str(j+1)+"."+str(12)+"."+InshoreMois[11]["Hour"][24*j:24*(j+1)];#Ajout du jour à la ligne de temps du mois
            InshoreMois[11]["Day"][24*j:24*(j+1)]=str(j+1)+"."+str(12)
            if j!=0:
                InshoreMois[11]=InshoreMois[11].drop([LignesJours[11][j]-1,LignesJours[11][j],LignesJours[11][j]+1]);
            k+=1;
    InshoreJours.append(pd.DataFrame({"Hour": Inshore["Unnamed: 0"][LignesJours[11][len(LignesJours[11])-1]+2:Inshore["Unnamed: 0"].size]}));#Traitement de la fin de la série des jours de décembre
    InshoreJours[k]["Current"]=Inshore["Current"][LignesJours[11][len(LignesJours[11])-1]+2:Inshore["Unnamed: 0"].size];
    InshoreJours[k]["Intraday"]=Inshore["Intraday"][LignesJours[11][len(LignesJours[11])-1]+2:Inshore["Unnamed: 0"].size];
    InshoreJours[k]["Day Ahead"]=Inshore["Day Ahead"][LignesJours[11][len(LignesJours[11])-1]+2:Inshore["Unnamed: 0"].size];
    InshoreMois[11]["Hour"][24*(len(LignesJours[11])-1):24*len(LignesJours[11])]=str(len(LignesJours[11]))+"."+str(12)+"."+InshoreMois[11]["Hour"][24*(len(LignesJours[11])-1):24*len(LignesJours[11])];
    InshoreMois[11]["Day"][24*(len(LignesJours[11])-1):24*len(LignesJours[11])]=str(len(LignesJours[11]))+"."+str(12)
    
    print("Listes construites\n");
