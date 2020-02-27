# -*- coding: utf-8 -*-
"""
Modified for Cassiopea project
test di update della brach customizzata
"""

import xlwings as xw
from datetime import datetime
import pandas as pd
import numpy as np
from datetime import timedelta



"""wb = xw.Book('Verification_List.xslx')  # connect to an existing file in the current working directory
#wb=xw.Book()
sht = wb.sheets['Sheet1']
sht.range('A10').value = pd.DataFrame(table).T"""

def structure_by_package(mel):
    """this function re-organizes the MEL by packages and systems and products."""
    """receives in a pandas dataframe"""
    string='K10024-'
    WP='00'
    l={}
    mel['Level 1','Level 2','Level 3','Level 4']=''
    mel['WP']=mel['Level'].astype(str)  
    for i,row in mel.iterrows():
        print (WP)
        if (type(row['WP Activity/ Part No.']) is str) and (string in row['WP Activity/ Part No.']) :
            #new section starts:
            WP=row['WP Activity/ Part No.']
            l[row['Level']]=row['Equipment Description']
            
        mel.loc[i,'WP']=WP
        for key in l.keys():
            mel.loc[i,'Level ' +key]=l[key]
            
    mel.dropna(subset=['Delivery','WP'], inplace=True)
    
    mel['WP']=mel['WP'].str.replace('K10024-','',regex=False)   
    mel['WP']=mel['WP'].str[:2]
    #mel.drop(columns=['Level'],inplace=True)    
    mel.to_excel('packages_MEL.xlsx')
    return mel

def consolidate_saipem_mel(mel):
    """not used normally -- ignore"""
    c_MEL={}
    mel=mel.dropna(subset=['Part No.'])
    mel['Part No.']=mel['Part No.'].astype(str)
    mel['Quantity']=mel['Quantity'].fillna(value=0).astype(str) 
    mel['Spare Quantity']=mel['Spare Quantity'].fillna(value=0).astype(str)

    mel['Quantity']=mel['Quantity'].str.replace(' TBC','',regex=False)  
    mel['Quantity']=mel['Quantity'].str.replace('TBC','0',regex=False)  
    mel['Quantity']=mel['Quantity'].str.replace('HOLD','0',regex=False)  
    mel['Spare Quantity']=mel['Spare Quantity'].str.replace(' TBC','',regex=False)  
    mel['Spare Quantity']=mel['Spare Quantity'].str.replace('TBC','0',regex=False)  
    mel['Spare Quantity']=mel['Spare Quantity'].str.replace('HOLD','0',regex=False)  
    #print (mel['Quantity'][mel['Quantity'].isna()])

    print (mel['Quantity'][mel['Quantity'].isna()])

    mel['Quantity']=mel['Quantity'].astype('int64')+mel['Spare Quantity'].astype('int64')

    for i, row in mel.iterrows():
        c_MEL[(str(row['Part No.']))]={'Quantity':mel['Quantity'][mel['Part No.'].astype(str)==str(row['Part No.'])].sum(),
                  'Part No.':row['Part No.'],
                  'Equipment Description':row['Equipment Description'],
                  'System':row['SYSTEM']}
    c_MEL=pd.DataFrame(c_MEL).T
    return c_MEL
    
def consolidate_mel(mel,delivery=False):
    """this sums all the delivery across packages and optionally across delivery types SOS, REN, CPP. . this can be useful in different scenarios (e.g. CPI list logistics) etc
    the delivery boolean allows for aggregating across delivery types."""
    c_MEL={}
    WP=00
    
    mel['Part No.']=mel['WP Activity/ Part No.']
    mel['Part No.']=mel['Part No.'].astype(str)

    #mel['Quantity']=mel['Quantity'].str.replace('m','',regex=False)  

    mel['Quantity']=mel['Quantity'].fillna(value=0).astype(str)   
    mel['Quantity']=mel['Quantity'].str.replace('meters','',regex=True)  
    mel['Quantity']=mel['Quantity'].str.replace('m','',regex=False)  


    mel['Quantity']=mel['Quantity'].astype('float')
    if delivery:
        for i, row in mel.iterrows():
            c_MEL[(str(row['Part No.'])+row['Delivery'])]={'Quantity':mel['Quantity'][(mel['Part No.'].astype(str)==str(row['Part No.'])) & (mel['Delivery']==row['Delivery'])].sum(),
                  'Part No.':row['Part No.'],
                  'Delivery':row['Delivery'],
                  'Equipment Description':row['Equipment Description'],
                  'Company Work Package':row['Company Work Package'],
                  'WP':row['WP']}
    else:
        for i, row in mel.iterrows():
            c_MEL[(str(row['Part No.']))]={'Quantity':mel['Quantity'][mel['Part No.'].astype(str)==str(row['Part No.'])].sum(),
                  'Part No.':row['Part No.'],
                  'Company Work Package':row['Company Work Package'],                
                  'Equipment Description':row['Equipment Description']}
        
    c_MEL=pd.DataFrame(c_MEL).T  
    return c_MEL

def rev_check(me1l,mel2):
    c_mel1=consolidate_mel(mel1,False)
    c_mel2=consolidate_mel(mel2,False)
    
    #check removed
    rem={}
    c_mel=pd.concat([c_mel1,c_mel2])
    """c_mel['Track Change']='New'
    removed=mel2.index.isin(mel1.index)
    c_mel.loc[c_mel1.index,'Track Change']=''
    c_mel['Track Change'][~removed]='Removed'"""
    add={}
    diff={}
    rem={}
    for i,row in c_mel.iterrows():
        try:
            c_mel1.loc[i]
            try:
                c_mel2.loc[i]
                if c_mel1.loc[i]['Quantity']!=c_mel2.loc[i]['Quantity']:
                    diff[i]=row
            except:
                rem[i]=row
        except:
            add[i]=row
    
    diff=pd.DataFrame(diff).T
    rem=pd.DataFrame(rem).T
    add=pd.DataFrame(add).T
    diff['Change']='Change Quantity'
    rem['Change']='Removed from 01 to 02'
    add['Change']='New from 01 to 02'
    
    c_mel=pd.concat([diff,rem,add])
    c_mel.to_excel('track_changes.xlsx')



#removed1=mel1[~mel1[mel2.index]]
if __name__ == "__main__":
    mel=pd.read_excel('082140DGEAS0051_EXDE01.xlsm', sheet_name='LIST')
    melcons=structure_by_package(mel)
    #melcons.to_excel("consolidated_MEL.xlsx")
    
    
   # mel1=pd.read_excel('packages_MEL.xlsx')
   #mel2=pd.read_excel('packages_MEL02.xlsx')
    
