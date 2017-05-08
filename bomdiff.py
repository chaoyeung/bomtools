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

file = glob.glob("*.csv")

adf = pd.read_csv(file[0])
bdf = pd.read_csv(file[1])

# Get the Version of the bom
# Remove the last n characters from a string [:-n]
adf.iat[0,2] = adf.iloc[0,1] + '_' + adf.iloc[0,2] + '_V' + adf.iloc[0,4][:-8]
bdf.iat[0,2] = bdf.iloc[0,1] + '_' + bdf.iloc[0,2] + '_v' + bdf.iloc[0,4][:-8]

adf.iat[0,1] = adf.iloc[0,1] + '_V' + adf.iloc[0,4][:-8]
bdf.iat[0,1] = bdf.iloc[0,1] + '_V' + bdf.iloc[0,4][:-8]

adf.iat[0,0] = 'TX03'
bdf.iat[0,0] = 'TX03'

# Drop the column Identity
adf = adf.drop(['Quantity','Version'], axis = 1)
bdf = bdf.drop(['Quantity','Version'], axis = 1)

#Rename the Columns
adf = adf.rename_axis({"Number": "Comcode", "Name": "Description","Reference Designator":"Ref"}, axis="columns")
bdf = bdf.rename_axis({"Number": "Comcode", "Name": "Description","Reference Designator":"Ref"}, axis="columns")

bom_name_a = adf.iloc[0,2]
bom_name_b = bdf.iloc[0,2]

bom_name_a_short = adf.iloc[0,1] 
bom_name_b_short = bdf.iloc[0,1]

print "boms loaded"


d = pd.merge(adf,bdf,how='inner',on='Ref')
a = adf[~adf.Ref.isin(d.Ref)]
a = a.reset_index(drop=True)
b = bdf[~bdf.Ref.isin(d.Ref)]
b = b.reset_index(drop=True)

y = pd.merge(adf,bdf)
x = adf[~adf.Ref.isin(y.Ref)]
y = bdf[~bdf.Ref.isin(y.Ref)]
c = pd.merge(x,y,how='inner',on='Ref')
c = c.reset_index(drop=True)

excel_file_name = 'diff' + '_' + bom_name_a_short + '_' + bom_name_b_short + '.xlsx'
writer = pd.ExcelWriter(excel_file_name)
a.to_excel(writer,bom_name_a_short +' ONLY')
b.to_excel(writer,bom_name_b_short +' ONLY')
c.to_excel(writer,'Replacement')
writer.save()

print "--------------------------------------------------------------"
print "Reference Designator only in " + bom_name_a
print a.drop(['Description'], axis = 1)

print "--------------------------------------------------------------"
print "Reference Designator only in " + bom_name_b
print b.drop(['Description'], axis = 1)

print "--------------------------------------------------------------"
print "Same Reference Designator with different comcode"
print c.drop(['Description_x','Description_y'], axis = 1)

print "Compare reslut has been save to diff.xlsx in current directory"