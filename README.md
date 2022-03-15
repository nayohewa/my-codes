# my-codes
codes of sql
import pandas as pd
import numpy as np
import openpyxl
from openpyxl import Workbook, load_workbook
from os import listdir
import os
dir=os.path.join('/Users/prabhathsamaranayake/Desktop/task01/data_files')
dir
file_found = False
for files in os.listdir(dir):
    print(f"processing file: '{files}'")
    if files[-4::] == 'xlsx':
        file_found = True
    else:
        print(f"current file does not end with xlsx. Its last 4 chars are: '{files[-4:]}'")
if not file_found:
    print("ERROR: There were no files with ending xlsx")
file_1 = pd.ExcelFile(os.path.join(dir,files),engine='openpyxl')
print('Path of File: ', os.path.join(dir,files))
print('Student: ', pd.read_excel(file_1, sheet_name=0).iloc[0,1])
sheets_names = ['Yr1', 'Yr2', 'Yr3','Subjects']
for names in sheets_names:
    sheet = file_1.sheet_names.index(names)
    print('Sheet: ', file_1.sheet_names[sheet])
    file_original = pd.read_excel(file_1, sheet_name=sheet,engine='openpyxl')
    file_copy = file_original.copy()
    print(file_copy.columns)
    #print(file_copy)   
credits={"AMAT":1,"BFIN":1,"DELT":0,"ELEC":1,"MGMT":1,"PHYS":1,"PMAT":1,"COST":1,"MAPS":1,"COSC":1,"STAT":1,"BOTA":1}
grades = {'**':0, 'AB':0, 'A+':4, 'A':4, 'A-':3.7, 'B+':3.3, 'B':3, 'B-':2.7, 'C+':2.3, 
          'C':2, 'C-':1.7, 'D+':1.3, 'D':1, 'E':0}
    for i in range(len(file_copy)):
                file_copy.loc[i,'Grades'] = grades[file_copy.loc[i,'Grade']]
                #sortedby=file_copy.sort_values(file_copy.loc[:,'Course Code'])                               
            # Rows Repeated
                dupli = file_copy.loc[file_copy.duplicated(['Course Code'], keep='first')].reset_index()
                cont = 0
                rows_to_delete = []

                for i in range(int(len(dupli))):
                    if dupli.loc[cont,'Grades'] >= dupli.loc[cont+1,'Grades']:
                        rows_to_delete.append(dupli.loc[cont+1,'index'])
                else:
                    rows_to_delete.append(dupli.loc[cont,'index'])
                    cont += 2
                    file_copy.drop(index=rows_to_delete, inplace=True)
                    file_copy.reset_index(drop=True, inplace=True)
              # Convertion the grades
            
                    #file_copy.loc[i,'Grades']=format(file_copy.loc[i,'Grades'],"0.2f")

            # Total of credit (all last digits of Coruse Code)
                    file_copy.loc[:,'subject_credits'] = [int(i[-1]) for i in file_copy.loc[:,'Course Code']]
                    print('')
                    file_copy.loc[:,'Credit_Multiplier'] = [str(i[0:4]) for i in file_copy.loc[:,'Course Code']]
                    credits=[("AMAT",1),("BFIN",1),("DELT",0),("ELEC",1),("MGMT",1),("PHYS",1),("PMAT",1),("COST",1),("MAPS",1),("COSC",1),("STAT",1),("BOTA",1)]
                    file_copy.loc[:,'course_unit']=file_copy.loc[:,'Credit_Multiplier'].map(dict(credits))
                    file_copy.loc[:,'subject_credits'].sum()#need to give specific cell
                    for index in range(len(file_copy)):
                #total_credits[index] = last digit of Course Code
                            gpv= file_copy.loc[:,'Grades']*file_copy.loc[:,'subject_credits'] *file_copy.loc[:,'Course_Unit']
                            GPA=gpv/file_copy.loc[:,'subject_credits'].sum()
                #file_copy.loc[i,'gpa']= (total_credits[index]*Grades*Course_Unit)/sum(total_credits)
                            GPA=format(file_copy.loc[:,'gpv'],"0.2f")
                            print(GPA.sum())
                            print(GPA)
                 
