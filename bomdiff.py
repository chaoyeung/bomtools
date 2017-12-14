# -*- coding: utf-8 -*-
"""
Bom Compare Script

This script read two boms exported form windchill in the format of *.csv and compare them to find the differneces.
Use "Reference Designator" as the index for comparison, which is unique is a bom file.
The difference report generated will have the 3 sections below:
    1. A list of the parts only in bom A
    2. A list of the parts only in bom B
    3. Ref show in bom A and B, but the comcode is different

"""

import glob
import pandas as pd
import time

file = glob.glob("*.csv") #seach csv filr in current folder

bom_name_a = file[0]
bom_name_b = file[1]

adf = pd.read_csv(file[0])
bdf = pd.read_csv(file[1])

# Check the file format and set the header

adf.iat[0,0] = 'TX03' # Add TX03
bdf.iat[0,0] = 'TX03'

adf.iat[0,1] = adf.iloc[0,1] + '_' + file[0]
bdf.iat[0,1] = bdf.iloc[0,1] + '_' + file[1]

# Drop the column Identity
adf = adf.drop(['Quantity','Version','Name'], axis = 1)
bdf = bdf.drop(['Quantity','Version','Name'], axis = 1)

#Rename the Columns, short names
adf = adf.rename_axis({"Number": "Comcode", "Reference Designator":"Ref"}, axis="columns")
bdf = bdf.rename_axis({"Number": "Comcode", "Reference Designator":"Ref"}, axis="columns")


d = pd.merge(adf,bdf,how='inner',on='Ref') #Retain only rows whose Ref is in both sets

#All rows in adf that do not have a match in d
#filtering the Refs only in adf
a = adf[~adf.Ref.isin(d.Ref)]

a = a.reset_index(drop=True)

#Filering the Refs only in bdf using the same method
b = bdf[~bdf.Ref.isin(d.Ref)]
b = b.reset_index(drop=True)

print "--------------------------------------------------------------"
print "Reference Designator only in " + bom_name_a
print a

print "--------------------------------------------------------------"
print "Reference Designator only in " + bom_name_b
print b

y = pd.merge(adf,bdf) #Retain rows have same ref and comcode
x = adf[~adf.Ref.isin(y.Ref)] #different rows in adf,
y = bdf[~bdf.Ref.isin(y.Ref)]
c = pd.merge(x,y,how='inner',on='Ref') #Same ref but different Comcode
c = c.reset_index(drop=True)

print "--------------------------------------------------------------"
print "Same Reference Designator with different comcode"
print c

#add time stamp to xlsx file
current_time = time.strftime("%Y%m%d%H%M%S",time.localtime())

print "Compare reslut has been save to diff.xlsx in current directory"
excel_file_name = current_time + '_diff' + '_' + bom_name_a+ '_' + bom_name_b + '.xlsx'
writer = pd.ExcelWriter(excel_file_name)
a.to_excel(writer,bom_name_a +' ONLY')
b.to_excel(writer,bom_name_b +' ONLY')
c.to_excel(writer,'Replacement')
writer.save()
