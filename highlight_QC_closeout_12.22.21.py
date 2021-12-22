# -*- coding: utf-8 -*-
"""
Created on Wed Oct 20 15:49:45 2021

@author: godin033
"""

import pandas as pd
import numpy as np
import os
import math

cwd = input('Copy File path: ')

files_DF = os.listdir(cwd)
files_DF

os.chdir(cwd)
exportDF=[]
for nombres in files_DF:
    g=nombres.split('.')
    final1=g[0]
    final1=final1[0:30]
    exportDF.append(final1)
    
    
excels_DF = [pd.ExcelFile(name) for name in files_DF]

col_names_DF=[x.parse(x.sheet_names[0], header=0,index_col=None,keep_default_na=False) for x in excels_DF]

demographics=col_names_DF[1]
#only keep rows with On Study dates
demographics.dropna(subset=['On Study Date'], inplace=True)

on_study_seq=list(demographics['Sequence No.'])

new=[]
for inde, data in enumerate(col_names_DF):
            elig_filter= data[data['Sequence No.'].isin(on_study_seq)]
            list_colnames=elig_filter.columns.to_list()
            sub1 ='date'
            index_list1=[j for j in list_colnames if sub1 in j.lower() and 'ext' not in j.lower()]
            for name1 in index_list1:
                elig_filter[name1] = pd.to_datetime(elig_filter[name1],errors='coerce').dt.strftime('%Y-%m-%d')
            #You may need to edit this by modifiying the unessary columns based on protocol calendar and arm set up
            elig_filter=elig_filter.drop(columns=[ 'Initials','Form','Form Desc.','Phase','Level','Form Status'], errors='ignore')
            new.append(elig_filter)
            
            

            
            
filtered =[]   
for test in new:
    test.replace(r'^\s*$', np.nan, regex=True, inplace = True)
    if 'Not Applicable or Missing' in test.columns:
        test1=test.loc[(test['Not Applicable or Missing'].isnull())].fillna('empty')
        col_filter=test1.columns[[0,1,2,3,4,5,6]]
        test1[col_filter]=test1[col_filter].replace({'empty':''})
        test2=test.loc[~(test['Not Applicable or Missing'].isnull())]
        final_missing=test1.append(test2)
        final_test = final_missing.sort_values(by = ['Sequence No.', 'Visit Date']) 
    else:
        final_test=test.fillna('empty')
        
        
    #test = test.replace(r'^\s*$', np.nan, regex=True)
    #test1 =test.loc[ (test['Not Applicable or Missing'].isnull())& (test['Form Status'] == 'Started') ].fillna('empty')
    #test1=test.loc[ (test['Not Applicable or Missing'].isnull())& (test['Form Status'] == 'Started') ].fillna('empty')

    #test=test.loc[(test['Not Applicable or Missing'].isnull())& (test['Form Status'] == 'Completed')].fillna('empty')
    filtered.append(final_test)   
    
    
output=input('Please enter output file: (you can write all file path or just file name and it will go to original input foler):')

import xlsxwriter
writer = pd.ExcelWriter(output, engine='xlsxwriter')
#formatyellow = workbook.add_format({'bg_color':'#F7FE2E'})


for num, sets in enumerate(filtered, start=0):
    sets.to_excel(writer, sheet_name=exportDF[num])
    workbook  = writer.book
    worksheet = writer.sheets[exportDF[num]]
    colnumber=len(sets.columns)
    worksheet.autofilter(0, 0, colnumber, colnumber)
    #format1 = workbook.add_format({'text_wrap': True})
    # format2.set_bg_color('green')
    worksheet.set_column('B:I', 15)
    worksheet.set_column('K:CZ', 18)
    header_format = workbook.add_format({
    'bold': True,
    'text_wrap': True,
    'valign': 'middle',
    'fg_color': '#B0C4DE',
    'border': 1})

# Write the column headers with the defined format.
    for col_num, value in enumerate(sets.columns.values,start=0):
        worksheet.write(0, col_num + 1, value, header_format)
    formatyellow = workbook.add_format({'bg_color':'#F7FE2E'})
    worksheet.conditional_format('K1:DZ500', {'type':   'text',
                                          'criteria': 'containing',
                                          'value':   'empty',
                                       'format': formatyellow})
workbook.close()