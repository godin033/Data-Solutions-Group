# -*- coding: utf-8 -*-
"""
Created on Thu May 27 13:59:33 2021

@author: godin033
"""

import pandas as pd
import numpy as np
import os


cwd = input('Copy File path: ')
#load in trike automated report 
xls=pd.ExcelFile(cwd)

list_frames=xls.sheet_names
sheet_to_df_map = pd.read_excel(xls, sheet_name=None)
dfs_list=list(sheet_to_df_map.values())

dataframe_formname=[]
for ind, data in enumerate(dfs_list):
    # data.columns=data.iloc[0]
    data=data.drop(['Unnamed: 0','Arm','Level'] ,axis=1,errors='ignore')
    for ind1, form in enumerate(list_frames):
        if ind==ind1:
            data['Form']=str(form)
            dataframe_formname.append(data)


#table of form counts

masterlist=[]
for num1 , df_count in enumerate(dataframe_formname):
    for num2,formname in enumerate(list_frames):
        if num1==num2:
            counts=df_count.groupby(['Sequence No.'])['Sequence No.'].count()
            list_holder=[]
            list_holder.append(formname)
            list_holder.append(counts)
            masterlist.append(list_holder)
            
df_count_entry = pd.DataFrame(masterlist, columns = ['Form', 'Count'])
cwd1 = input('Enter where to save the row counts per form: ')
df_count_entry.to_excel(cwd1)
#========================
QC_report=[]
QC_report2=[]
for test in dataframe_formname:

    # test=test[(test['Not Applicable or Missing'] != 'Not Applicable')]
    # test=test[(test['Not Applicable or Missing'] != 'Missing')]
    
    index_test=test.set_index(['Form','Not Applicable or Missing','Sequence No.',  'Segment','Visit Date',
        ])
    #test.drop(columns={'Unnamed: 0','Phase'}, inplace=True)

   # desc_clean = index_test[index_test.columns.drop(list(index_test.filter(regex='Description')))]
   # comments_clean = desc_clean[desc_clean.columns.drop(list(desc_clean.filter(regex='Comments')))]
   # other_clean = comments_clean[comments_clean.columns.drop(list(comments_clean.filter(regex='Other')))]

    stacked=index_test.stack()
    stacked_2=stacked.to_frame()
    stacked_2 = stacked_2.rename(columns = {0:'Blank fields'})
    stacked_clean=stacked_2[stacked_2['Blank fields'].str.contains('empty', na=False)]
    stacked_clean2=stacked_2[stacked_2['Blank fields'].str.contains('No Source Data', na=False)]
    QC_report.append(stacked_clean)
    QC_report2.append(stacked_clean2)
    

df_merged = pd.concat(QC_report)
df_merged2=pd.concat(QC_report2)
df_merged[['QC Issue','DMA Comment']] = df_merged['Blank fields'].str.split(";",n=1,expand=True)
#df_merged2[['QC Issue','DMA Comment']] = df_merged2['Blank fields'].str.split("ta ",n=1,expand=True)

#df_merged_final=pd.concat([df_merged, df_merged2])
cwd2= input('file path to save final comments: ')
df_merged.to_excel(cwd2, merge_cells=False)
    