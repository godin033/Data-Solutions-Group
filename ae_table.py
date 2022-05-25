
"""
Created on Wed Jan 19 09:37:06 2022

@author: godin033
"""

#load in libraries
import pandas as pd
import numpy as np
import os
import math

# load in all of

file_name=input('File path to pull data:')

ae_list=pd.read_excel(file_name)

ae_list_clean=ae_list[['CDUS Toxicity Type Code', 'Toxicity Category', 'Grade']]



test=pd.pivot_table(ae_list_clean, values='Toxicity Category',index=['CDUS Toxicity Type Code'], columns='Grade',
               aggfunc='count')
test=test.reset_index()
class_dict=dict(zip(ae_list['CDUS Toxicity Type Code'],ae_list['Toxicity Category']))

	
test['Toxicity Category']= test['CDUS Toxicity Type Code'].map(class_dict)
test.columns

final_table=test[['Toxicity Category','CDUS Toxicity Type Code','1 Mild', '2 Moderate', '3 Severe',
       '4 Life-threatening or disabling']]



#write file out for TABLE 1 with all the AEs
final_table.to_excel(r'M:/CRTI/DSG/DSG trials/2020LS001-FT516 (CLOSE OUT)/ae_table_5.25.22_V1.xlsx', index = False)


#----TABLE 2------
ae_list_v1=ae_list[['CDUS Toxicity Type Code', 'Toxicity Category', 'Grade', 'Attribution']]
#Filter only certain attibutions
ae_list_v2 = ae_list_v1[(ae_list_v1["Attribution"] =='Possible') | (ae_list_v1["Attribution"] =='Definite')|(ae_list_v1["Attribution"] =='Probable')]


test1=pd.pivot_table(ae_list_v2, values='Toxicity Category',index=['CDUS Toxicity Type Code'], columns='Grade',
               aggfunc='count')
test1=test1.reset_index()

class_dict1=dict(zip(ae_list_v2['CDUS Toxicity Type Code'],ae_list_v2['Toxicity Category']))

	
test1['Toxicity Category']= test1['CDUS Toxicity Type Code'].map(class_dict1)
test1.columns

final_table1=test1[['Toxicity Category','CDUS Toxicity Type Code','1 Mild', '2 Moderate', '3 Severe',
       '4 Life-threatening or disabling']]

#Make sure to update the OUTPUT path:

final_table1.to_excel(r'M:/CRTI/DSG/DSG trials/2020LS001-FT516 (CLOSE OUT)/ae_table_5.25.22_2.xlsx', index = False)
