# -*- coding: utf-8 -*-
"""
Created on Thu Nov 11 13:03:40 2021

@author: godin033
"""

import pandas as pd
import numpy as np
import os
import math

#loaded in in calendar
file_name = input('regular file input:') 
calendar=pd.read_excel(file_name)


#Calendar Timepoint
calendar_2=calendar.drop([0,1,2,3,4])
#reset column names
calendar_2.columns = calendar.iloc[5]

#drop rows with nothing in form columns

calendar_2.dropna(subset = ['Forms'], inplace=True) # have to replace column names but most likely blank

calendar_3 = calendar_2.T
calendar_3.columns = calendar_3.iloc[1]
#calendar_3=calendar_3.drop(['Procedure','Forms'])

calendar_3=calendar_3.reset_index()
#calendar_3 = calendar_3.iloc[: , 1:]

calendar_3=calendar_3.where(calendar_3 != 'R', calendar_3.columns.to_series(), axis=1)
calendar_3=calendar_3.where(calendar_3 != 'INS', calendar_3.columns.to_series(), axis=1)
calendar_3=calendar_3.where(calendar_3 != 'INS1', calendar_3.columns.to_series(), axis=1)
calendar_3=calendar_3.where(calendar_3 != '2R', calendar_3.columns.to_series(), axis=1)
calendar_3=calendar_3.where(calendar_3 != '3R', calendar_3.columns.to_series(), axis=1)
calendar_3=calendar_3.drop([0,1])
calendar_3 = calendar_3.iloc[: , 1:]


#can choose one
##calendar_3.drop(columns=['Relabeled visits'], inplace=True)
#calendar_3.drop(columns=['Procedure'], inplace=True)

##do not need to drop forms
calendar_4=calendar_3
calendar_4.columns
calendar_4=calendar_4.drop(columns=['Forms'])

calendar_3['Forms_at_visit']=calendar_4.apply(lambda x: '\n'.join(x.dropna()), axis=1)


calendar_3['Forms_at_visit'] = calendar_3['Forms_at_visit'].str.split('\n')

# convert list of pd.Series then stack it
calendar_test= (calendar_3
 .set_index(['Forms'])['Forms_at_visit']
 .apply(pd.Series)
 .stack()
 .reset_index()
 .drop('level_1', axis=1)
 .rename(columns={'Forms':'Timepoints',0:'Forms'}))


#================write nicely into xcel=============


# Create the list where we 'll capture the cells that appear for 1st time,
# add the 1st row and we start checking from 2nd row until end of df
startCells = [1]
for row in range(2,len(calendar_test)+1):
    if (calendar_test.loc[row-1,'Timepoints'] != calendar_test.loc[row-2,'Timepoints']):
        startCells.append(row)

output=input('outputfile name or complete path:')
writer = pd.ExcelWriter(output, engine='xlsxwriter')

calendar_test.to_excel(writer, sheet_name='Sheet1', index=False)
workbook = writer.book
worksheet = writer.sheets['Sheet1']
merge_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 2})


lastRow = len(calendar_test)

for row in startCells:
    try:
        endRow = startCells[startCells.index(row)+1]-1
        if row == endRow:
            worksheet.write(row, 0, calendar_test.loc[row-1,'Timepoints'], merge_format)
        else:
            worksheet.merge_range(row, 0, endRow, 0, calendar_test.loc[row-1,'Timepoints'], merge_format)
    except IndexError:
        if row == lastRow:
            worksheet.write(row, 0, calendar_test.loc[row-1,'Timepoints'], merge_format)
        else:
            worksheet.merge_range(row, 0, lastRow, 0, calendar_test.loc[row-1,'Timepoints'], merge_format)


writer.save()








