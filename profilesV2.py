# -*- coding: utf-8 -*-
"""
Created on Fri Dec  6 11:50:44 2019
@author: TintM
"""

# -*- coding: utf-8 -*-
"""
Created on Mon Nov 25 11:04:34 2019
@author: TintM
"""

import xlrd 
import pandas as pd
import os
import numpy as np
from openpyxl import load_workbook
import time
import sys, os 

def Name(profiles):
    
    swap = str(profiles["Prog_Name"])
    
    return swap

def Descriptor(profiles,data):
 
    swap = str(profiles["Prog_Cat"])
    
    if swap == 'Maritime, C4ISR & Gunships':
        swap = '<mission>'        
   
    elif swap == 'Light 1-2 Seaters (Piston / Turboprop / Jet)':
        swap = 'light 1-2 seat aircraft'
        
    elif swap == 'Military Transports & Tankers':
        data = pd.Series(data['Fixed Wing Weight'].unique())         
        if len(data)==1:    
            swap = data[0] + ' transport aircraft'
            swap = swap.lower()
        else:
            swap = 'transport aircraft'

    elif swap == 'Civil Propeller (3+ Seats)':
        swap = 'civil propeller aircraft'
    
    elif swap == 'UAVs':
        data = pd.Series(data['AircraftCat'].unique())                    
        swap = SeriesList(data,' or')
        swap = swap.replace("UAV","")
        swap = swap.replace("Group","")
        swap = "Group" + swap + "UAV"
        
    elif swap == 'Helicopters & Tiltrotors':
        swap = 'helicopter'
    
    elif swap == 'Fighters':
        swap = 'fighter'
    
    elif swap == 'Commercial Jets':
        swap = 'commercial jet'
    
    elif swap == 'Bombers':
        swap = 'bomber'
    
    elif swap == 'Amphibians':
        swap = 'amphibious aircraft'
    
    elif swap == 'Business Jets':
        swap = 'business jet'

    return swap

def BuiltBy(data):
    prime = pd.Series(data['Aircraft Mfr'].unique())
    lastmfr = pd.Series(data['LastMFR'].unique())
    
    if prime.equals(lastmfr):
        swap = 'built by ' + SeriesList(prime,' and')
    else:
        swap = ""
        
    return swap    
    

def Operators(data):
    operators = pd.Series(data['Operator'].unique())
    
    if len(operators)==0:
        swap=""
    
    elif len(operators)==1:
        swap = str(operators[0])
        swap = 'the ' + swap 
    
    else:
        swap = str(len(operators))
        swap = swap + ' operators'               
        
    return swap

def ServiceOrOrder(data):
    
    isf = sum(data['Current Fleet'])
    
    try:
        uncertainisf = len(data[data['QuestionableNbr']=="X"])
    except KeyError:
        uncertainisf = 0
    
    try:    
        orders = sum(data['Total_Deliveries_Y0toY10'])
    except KeyError:
        orders = 0

    if uncertainisf > 0:
        swap = " in military service or on order in "     
    else:
        if isf == 0:
            if orders == 0:
                swap = " in military service in "
            else:
                swap = " on order in "
        else:
            if orders == 0:
                swap = " in military service in "
            else:
                swap = " in military service or on order in "
    
    return swap

def Countries(data):
    
    countries = pd.Series(data['Country'].unique())

    if len(countries)==0:
        swap= ""
    
    elif len(countries)<6:
        swap = SeriesList(countries,' and')
        swap = TheCountries(swap)

    else:
        swap = str(len(countries)) + ' countries'
    
    return swap


def ISFandOrders(profiles,data):
 
    isf = sum(data['Current Fleet'])
    
    try:
        uncertainisf = len(data[data['QuestionableNbr']=="X"])
    except KeyError:
        uncertainisf = 0
    
    try:    
        orders = sum(data['Total_Deliveries_Y0toY10'])
    except KeyError:
        orders = 0
    
    swap = " As of December 2019,"
    
    if uncertainisf > 0:
        swap = swap + " it was in military service or on order with"
    else:
        if isf == 0:
            if orders == 0:
                swap = swap + " it was in military service with"
            elif orders == 1:
                swap = swap + " there was one new delivery on contract with"
            else:
                swap = swap + " there were " + format(orders, ',') + " new deliveries on contract with"
        
        elif isf == 1:   
            swap = swap + " there was one in military service"
            if orders == 0:
                swap = swap + " with"
            elif orders == 1:
                swap = swap + " and one on order with"
            else:
                swap = swap + " and " + format(orders, ',') + " new deliveries on contract with"            
        else:
            swap = swap + " there were " + format(isf, ',') + " in military service"
            if orders == 0:
                swap = swap + " with"
            elif orders == 1:
                swap = swap + " and one on order with"
            else:
                swap = swap + " and " + format(orders, ',') + " new deliveries on contract with"             

    return swap


def TypeCount(profiles,data):
    
    data = pd.Series(data['Type'].unique()) 

    swap = len(data)

    if swap == 1:
        swap = "one type"
    else:
        swap = str(swap) + " types"

    return swap

def TypeList(profiles,data):  
    
    data = pd.Series(data['Type'].unique()) 
    
    if len(data)>20:
        swap = ""
    
    elif len(data)==1:
        swap = data[0]
        swap = ', the ' + swap 
    
    else:
        swap = SwapLast(", the " + data.str.cat(sep=', ', na_rep='?'),","," and")
                 
    return swap

def OtherProfiles(profiles,current):

    if len(profiles)==1:
        swap = ""
    
    elif len(profiles)==2:
        profiles = profiles[profiles['Prog_Name'] != current]
        swap = "This family also includes the " + SeriesList(profiles['Prog_FullName']," and") + " profile."                
    else:            
        profiles = profiles[profiles['Prog_Name'] != current]
        swap = "Other profiles in this family include the " + SeriesList(profiles['Prog_FullName'],' and') + "."
      
    return swap

def EngineCount(count):
    
    ecount = SeriesList(count.astype(str)," or") + " "
    ecount = TextNumbers(ecount)
    
    return ecount

def EngineFamily(family):  

    if len(family)==1:
        if family[0] == "Indeterminate":
            efamily = ""
        else:
            efamily = family[0] + " " 
    else:
        family = family[family != "Indeterminate"]
        efamily = SeriesList(family,' or') + " "
    
    return efamily

def EngineType(propulsion):
          
    etype = SeriesList(propulsion,' or') + " "
    etype = etype.lower()
    
    return etype

def AllEngines(profiles,data):
    
    ecount = EngineCount(pd.Series(data['NbrEngs'].unique()))
    efamily = EngineFamily(pd.Series(data['EngineFamily'].unique()))
    etype = EngineType(pd.Series(data['Propulsion'].unique()))
    
    swap = ecount + efamily + etype
    
    if ecount == "one ":
        swap = swap + "engine"
    else:
        swap = swap + "engines"
            
    return swap

def TheCountries(swap):  
        
    swap = swap.replace("USA","the United States")    
    swap = swap.replace("International [Europe]","NATO")    
    swap = swap.replace("International [Caribbean]","the Caribbean")    
    swap = swap.replace("United Kingdom","the United Kingdom")
    swap = swap.replace("Netherlands","the Netherlands")
    swap = swap.replace("Philippines","the Philippines")
    swap = swap.replace("Czech Republic","the Czech Republic")
    swap = swap.replace("Dominican Republic","the Dominican Republic")
    swap = swap.replace("Maldives","the Maldives")
    swap = swap.replace("United Arab Emirates","the United Arab Emirates")
   
    return swap

def SeriesList(data,final):
    
    temp = data.str.cat(sep=', ', na_rep='?')
    temp = SwapLast(temp,",",final)
    return temp
    
def SwapLast(s, old, new):
    return (s[::-1].replace(old[::-1],new[::-1], 1))[::-1]

def TextNumbers(string):
    
    string = string.replace('1','one')
    string = string.replace('2','two')
    string = string.replace('3','three')
    string = string.replace('4','four')
       
    return string

def EliminateDoubleSpaces(swap):
    
    while "  " in swap:
        swap = swap.replace("  "," ")
        
    return swap

def ImportData(inputfile):
    xlsx = pd.ExcelFile(inputfile)
    
    data = pd.read_excel(xlsx,"data")
    
    profiles = pd.read_excel(xlsx,"profiles")  
    profiles['Prog_AWIN_Acct_ID'] = profiles[['Prog_AWIN_Acct_ID']].applymap("{:.0f}".format)
    profiles['output'] = profiles['input']
    
    return profiles,data

def ExportData(table,output_file,sheetname):
    book = load_workbook(output_file)
    book[sheetname].delete_rows(1,2000)
    
    writer = pd.ExcelWriter(output_file, engine='openpyxl') 
    writer.book = book
    writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
    table.to_excel(writer, sheetname,index=False)
    
    writer.save()

    
    
def SwapCodes(inputfile,outputfile,outputfields):
    
    profiles, data = ImportData(inputfile)
    
    for i in range(0,len(profiles['output'])):      
        
        currentprofile = profiles.loc[i,:]
        currentdata = data[data["Prog_Name"]==currentprofile["Prog_Name"]]
    
        profiles.loc[i,"output"] = profiles.loc[i,"output"].replace("<name>",Name(currentprofile))    
        profiles.loc[i,"output"] = profiles.loc[i,"output"].replace("<descriptor>",Descriptor(currentprofile,currentdata))   
        profiles.loc[i,"output"] = profiles.loc[i,"output"].replace("<builtby>",BuiltBy(currentdata))
        profiles.loc[i,"output"] = profiles.loc[i,"output"].replace("<serviceororder>",ServiceOrOrder(currentdata))    
        profiles.loc[i,"output"] = profiles.loc[i,"output"].replace("<countries>",Countries(currentdata))
        profiles.loc[i,"output"] = profiles.loc[i,"output"].replace("<allengines>",AllEngines(currentprofile,currentdata))   
        profiles.loc[i,"output"] = profiles.loc[i,"output"].replace("<ISFandOrders>",ISFandOrders(currentprofile,currentdata))  
        profiles.loc[i,"output"] = profiles.loc[i,"output"].replace("<operators>",Operators(currentdata)) 
        profiles.loc[i,"output"] = profiles.loc[i,"output"].replace("<typecount>",TypeCount(currentprofile,currentdata)) 
        profiles.loc[i,"output"] = profiles.loc[i,"output"].replace("<typelist>",TypeList(currentprofile,currentdata)) 
        profiles.loc[i,"output"] = profiles.loc[i,"output"].replace("<otherprofiles>",OtherProfiles(profiles[profiles['Family'] == currentprofile["Family"]],currentprofile['Prog_Name']))  
        profiles.loc[i,"output"] = profiles.loc[i,"output"].replace("<lb>","\n")   
        profiles.loc[i,"output"] = EliminateDoubleSpaces(currentprofile["output"])
        
    ExportData(profiles[outputfields],outputfile,'output')    

    return profiles[outputfields]

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

print("started")
start_time = time.time()

inputfile = resource_path('rawdata.xlsx')
outputfile = resource_path('output2.xlsx')


outputfields = ['Prog_AWIN_Acct_ID','Prog_FullName','Prog_Cat','output']

output = SwapCodes(inputfile,outputfile,outputfields)
outputcheck = output[['output','Prog_FullName']]


print(" --- %s seconds ---" % (time.time() - start_time))
