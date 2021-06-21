#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Jul  6 14:36:14 2020

@author: FANI
"""

#load excel
import pandas as pd
import numpy as np
import os
import math

#loaded in all dataframes
cwd = input('regular file input:') 
files = os.listdir(cwd)
os.chdir(cwd)
#files

#files.remove('~$out_data_model.xlsx')
export=[]
for nombres in files:
    g=nombres.split('.')
    final=g[0]
    export.append(final)
    
export1 = [item[:30] for item in export]

    

excels = [pd.ExcelFile(name) for name in files]

col_names=[x.parse(x.sheet_names[0], header=None,index_col=None) for x in excels]
domain = [x.parse(x.sheet_names[1], header=None,index_col=None) for x in excels]

#download all the variable with data with


cwd_var = input('variable_name file input:')
files_var = os.listdir(cwd)
files_var

#files.remove('~$out_data_model.xlsx')
os.chdir(cwd_var)
excels_var = [pd.ExcelFile(name) for name in files]

var_names=[x.parse(x.sheet_names[0], header=None,index_col=None) for x in excels_var]

final=[]
for index, col in enumerate(col_names):
    for indexx, colx in enumerate(var_names):
        if index==indexx:
            eso=col.iloc[0:1]
            eso2=colx.iloc[0:1]
            nise=eso.append(eso2,ignore_index=True)
            nise.columns=nise.iloc[0]
            lol = nise.drop(0,0)
            lolx=lol.drop(columns=[ 'Initials','Form', 'Form Desc.', 'Phase','Cycle','Form Status'],errors='coerce')
            #lolx=lol.drop(columns=[ 'Initials'])
            final.append(lolx)


test1=[]
for df1 in domain:
    new=df1.drop([1,2] ,axis=1)
    new.columns=new.loc[0]
    new2=new.drop(0)
    new3=new2.pivot(columns="Column Name", values="Description")
    new4=new3.apply(lambda x: pd.Series(x.dropna().values))
    #new4=new3.drop(columns=['Column Name'])
    test1.append(new4)
    
    
    
master=[]   
for idx, val in enumerate(test1):
    for i, v in enumerate(final):
        if idx ==i:
            if val.empty:
                master.append(v)
            else:
                wtf=v.append(val, sort= False)
                master.append(wtf)
                

#saving the data

import xlsxwriter
output=input('outputfile name or complete path:')
writer = pd.ExcelWriter(output, engine='xlsxwriter')
#formatyellow = workbook.add_format({'bg_color':'#F7FE2E'})


for num, sets in enumerate(master, start=0):
    sets.to_excel(writer, sheet_name=export1[num])
    workbook  = writer.book
    worksheet = writer.sheets[export1[num]]
    colnumber=len(sets.columns)
    worksheet.autofilter(0, 0, colnumber, colnumber)
    format1 = workbook.add_format({'text_wrap': True})
    # format2.set_bg_color('green')
    worksheet.set_column('B:I', 15, format1)
    worksheet.set_column('J:CZ', 18, format1)
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


