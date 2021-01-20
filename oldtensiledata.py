# -*- coding: utf-8 -*-
"""
Created on Wed Apr 22 16:14:31 2020

@author: connor.folkman

Script was used to pull data from pdf files before Mark-10 Tensile tester was used
"""
import pandas as pd
from pandas import DataFrame
from pandas import ExcelWriter
import glob
import numpy as np
import tabula
from pathlib import Path
import os.path, time

#output file paths where csv files will be located
output_filepath = 'T:/MAESTRO/MAESTRO/TENSILE TESTER DATA FILES/OLD TENSILE DATA CSV FILES/'

#Empty lists to store date file was created and the file name 
date_file_created = []
file_name = []

#for loop that reads from every pdf file in the directory and pulls the table data from the document and makes a csv file from that table
for filepath in glob.iglob('T:/MAESTRO/MAESTRO/TENSILE TESTER DATA FILES/OLD TENSILE TEST DATA 4_21_2020/*.pdf'):
    date = time.strftime('%m-%d-%Y', time.localtime(os.path.getmtime(filepath)))
    date_file_created.append(date)
    name = Path(filepath).stem
    file_name.append(name[:-8])
    #converts the table in the pdf into a csv file if the csv file for that pdf doesnt already exist
    if os.path.exists(output_filepath + filepath[77:-12] + 'csv') == False:
        tabula.convert_to(filepath, output_filepath + filepath[77:-12] + 'csv', output_formate='csv', area=(50,351,408,650), lattic=True)
    
#For loop that matches the file name to the current file in order to associate the correct creation date with it, then populates the files into a dataframe
i = 0
df = DataFrame()       
for filepath in glob.iglob('T:/MAESTRO/MAESTRO/TENSILE TESTER DATA FILES/OLD TENSILE DATA CSV FILES/*.csv'):
    if filepath[72:-4] == file_name[i]:
        filename = np.array(['Type of test: ', filepath[72:-4], 'Date Test Performed: ', date_file_created[i]])
        fileseries = pd.Series(filename)
        df.append(fileseries, ignore_index=True)
        df.append(pd.read_csv(filepath, names = list(range(0,14))))
        i = i + 1
    else: 
        while filepath[72:-4] != file_name[i]:
            i = i + 1
            if len(file_name) == i:
                i = 0
        filename = np.array(['Type of test: ', filepath[72:-4], 'Date Test Performed: ', date_file_created[i]])
        fileseries = pd.Series(filename)
        df.append(fileseries, ignore_index=True)
        df.append(pd.read_csv(filepath, names = list(range(0,14))))
        i = i + 1

#writes the dataframe into an excel file 
writer = ExcelWriter('T:/MAESTRO\MAESTRO/ENGINEERING/PYTHON SCRIPTS/OUTPUT DATA/OLD TENSILE 2013-2020 DATA.xlsx')
df.to_excel(writer, 'Tensile', index=False)
writer.save()


        

