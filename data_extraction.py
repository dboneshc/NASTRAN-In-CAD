# -*- coding: utf-8 -*-
"""
Created on Thu Dec 19 16:42:02 2019

@author: davida
"""

#import regular expressions as re
import re

#import pandas as pd
import pandas as pd

#import xlwings as xw
import xlwings as xw

#loads NASTRAN OUT file data - example uses Plate Frame Assy Shrink_ShrinkWrap_1.ipt - Analysis 3 [Linear Static] OUT file
strFolder=r'C:\Users\david\Documents\Python Scripts'
strFile='ikjo1ocmm.OUT'
strPath=r''+strFolder+'\\'+strFile
f=open(strPath, mode='r')
strText=f.read()
f.close()

#looking for the information between the specified subcase and next page heading
subCase='4'
# see https://regexper.com for a visual representation of the regular expression below
# be careful of the concatenation of subCase into the regular expression string
strPattern=r'SUBCASE.*'+subCase+'\s*F O R C E S   I N   B A R   E L E M E N T S\n([\s\S]*?)PAGE\s*\d+\n'
pattern=re.compile(strPattern)

#initialize a counter - in the test case there are 7 pages of data of interest
ct=0

#initialize a big concatenation of the string results looking for all forces in bar elements
strResults=''

#uses regular expression object and generates an iterator that contains the results from the regular expression search
matches=pattern.finditer(strText)
for match in matches:
    #group index 1 is selected because only information between forces and page is desired
    #all the bar force results are concatenated across all the pages - this will be parsed further later
    strResults=strResults+match.group(1)
    ct=ct+1
    
#list of elements of interest from the NASTRAN analysis
elements=['1097','1076','1041','1037','1145','1187']
elementForcesMoments=[];

for iter in elements:
    strPattern2=r'\n\s*('+iter+'.*\n.*\n)'
    pattern2=re.compile(strPattern2)
    match=pattern2.search(strResults)
    if match:
        #extracts forces and moments for each element and appends them to a list
        elementForcesMoments.append(match.group(1))
        
#selection of either Node A or B from the elements of interest
nodes=['A','A','B','B','B','A']
nodalForcesMoments=[]

ct=0
for iter in nodes:
    if iter=='A':
        strPattern3=r'0.0000(.*)\n'
    else:
        strPattern3=r'1.0000(.*)\n'
    
    pattern3=re.compile(strPattern3)
    match=pattern3.search(elementForcesMoments[ct])
    if match:
        #extracts forces and moments for each node and appends them to the list
        nodalForcesMoments.append(match.group(1))
    ct=ct+1
 
#initialize a counter
ct=1
df=pd.DataFrame()   
    
#extract forces and moments and write them into a dataframe as real numbers
for iter in nodalForcesMoments:

    match=re.search(r'\s*(\S*)\s',iter)
    if match:
        moment1=float(match.group(1))
    else:
        moment1=float('nan')
        
    match=re.search(r'\s*\S*\s*(\S*)\s',iter)
    if match:
        moment2=float(match.group(1))
    else:
        moment2=float('nan')
        
    match=re.search(r'\s*\S*\s*\S*\s*(\S*)\s',iter)
    if match:
        shear1=float(match.group(1))
    else:
        shear1=float('nan')
    
    match=re.search(r'\s*\S*\s*\S*\s*\S*\s*(\S*)\s',iter)
    if match:
        shear2=float(match.group(1))
    else:
        shear2=float('nan')
        
    match=re.search(r'\s*\S*\s*\S*\s*\S*\s*\S*\s*(\S*)\s',iter)
    if match:
        axial=float(match.group(1))
    else:
        axial=float('nan')
    
    match=re.search(r'\s*\S*\s*\S*\s*\S*\s*\S*\s*\S*\s*(\S*)',iter)
    if match:
        torque=float(match.group(1))
    else:
        torque=float('nan')
    
    #append the data to the DataFrame
    dftemp=pd.DataFrame([axial, shear1, shear2, torque, moment2, moment1])
    dftemp=dftemp.transpose()
    df=df.append(dftemp)
    ct=ct+1
    
#add header to DataFrame
df.columns=['Fx', 'Fy','Fz', 'Mx','My', 'Mz']
#initial index list
indexList=[]
#add index to DataFrame
for i in range(len(elements)):
    indexList.append(elements[i]+'/'+nodes[i])
df.index=indexList

#write data to a new Excel file
wb=xw.Book()
sheet=wb.sheets['Sheet1']
sheet.range('A6').value=df
sheet.range('A6').value='Element/Node'
sheet.range('A4').value='Caution: The bar element coordinate system is independent of the global coordinate system.'
sheet.range('A1').value='File: '+strFile
sheet.range('A2').value='Subcase: '+subCase

#print the DataFrame in the console
print(df)