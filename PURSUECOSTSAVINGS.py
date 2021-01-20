# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""
import pandas as pd
pursue_pn = [504280001, 504280002,504280003,504280004,504280005,504280006,504280007,504280008,504280009,
             504281001, 504281002,504281003,504281004,504281005,504281006,504281007,504281008,504281009]
#filepaths = [f for f in listdir('T:/MAESTRO/MAESTRO/ENGINEERING/CONNORS FOLDER/Python Data Scripts/') if f.endswith('.csv')]
#df = pd.concat(map(pd.read_csv, filepaths))
df = pd.read_csv(r'T:/MAESTRO/MAESTRO/ENGINEERING/CONNORS FOLDER/PYTHON SCRIPTS/DATA/PursueScrap/pursue_2020.csv')
pursue_WO_list = df['part_number'].to_list()

i = 0
pursue_WO_count = 0

while i < len(pursue_WO_list):
    j = 0
    while j < len(pursue_pn):
        if int(pursue_WO_list[i]) == int(pursue_pn[j]):
            pursue_WO_count = pursue_WO_count + 1
        j+=1
    i+=1

savings_per_WO = 240
print(pursue_WO_count)
print(savings_per_WO*pursue_WO_count)




