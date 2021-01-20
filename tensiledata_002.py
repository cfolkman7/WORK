# -*- coding: utf-8 -*-
"""
Created on Fri Mar 27 14:15:19 2020

@author: connor.folkman

This rev is used if data is corrupted and fresh start of spreadsheet is needed
"""
import pandas as pd
from pandas import ExcelWriter, DataFrame
import glob
import os.path, time

#creates empty list to be populated with peak tensile test data and date created from each file for the hub
peak_tensile_data_hub = []
file_creation_date_hub = []
WO_number_hub = []
part_number_hub = []

#creates empty list to be populated with peak tensile test data and date created from each file for the tip
peak_tensile_data_tip = []
file_creation_date_tip = []
WO_number_tip = []
part_number_tip = []

#creates empty list to be populated with peak tensile test data and date created from each file for the mb
peak_tensile_data_mb = []
file_creation_date_mb = []
WO_number_mb = []
part_number_mb = []

#created empty data frames to put the data into
df = DataFrame()

#opens directory where tensile data is saved and reads from each file to find the peak force from colume 'Load' for the hub
for filepath in glob.iglob('//merit.com/Merit_Shares/South_Jordan/public/MAESTRO/MAESTRO/TENSILE TESTER DATA FILES/HUB TENSILE/*.xlsx'):
    df = pd.read_excel(filepath)
    file_creation_date_hub.append(time.strftime('%m/%d/%Y', time.localtime(os.path.getmtime(filepath))))
    WO_number_hub.append(filepath[99:107])
    part_number_hub.append(filepath[108:-5])
    try: 
        peak_tensile_data_hub.append(df['Load [N]'].max())
    except:
        try: 
            peak_tensile_data_hub.append(df['Load [lbF]'].max()/0.73756)
        except:
            peak_tensile_data_hub.append(0)
   
#opens directory where tensile data is saved and reads from each file to find the peak force from colume 'Load' for the tip
for filepath in glob.iglob('//merit.com/Merit_Shares/South_Jordan/public/MAESTRO/MAESTRO/TENSILE TESTER DATA FILES/TIP TENSILE/*.xlsx'):
    df = pd.read_excel(filepath)
    file_creation_date_tip.append(time.strftime('%m/%d/%Y', time.localtime(os.path.getmtime(filepath))))
    WO_number_tip.append(filepath[99:107])
    part_number_tip.append(filepath[108:-5])
    try: 
        peak_tensile_data_tip.append(df['Load [N]'].max())
    except:
        try: 
            peak_tensile_data_tip.append(df['Load [lbF]'].max()/0.73756)
        except:
            peak_tensile_data_tip.append(0)
            
#opens directory where tensile data is saved and reads from each file to find the peak force from colume 'Load' for the mb
for filepath in glob.glob('//merit.com/Merit_Shares/South_Jordan/public/MAESTRO/MAESTRO/TENSILE TESTER DATA FILES/MB TENSILE/*.xlsx'):
    df = pd.read_excel(filepath)
    file_creation_date_mb.append(time.strftime('%m/%d/%Y', time.localtime(os.path.getmtime(filepath))))
    WO_number_mb.append(filepath[98:106])
    part_number_mb.append(filepath[107:-5])
    try: 
        peak_tensile_data_mb.append(df['Load [lbF]'].max())
    except:
        try: 
            peak_tensile_data_mb.append(df['Load [N]'].max()*0.73756)
        except:
            peak_tensile_data_mb.append(0)
    

#creates new dataframe to load data into, sorting by date file was created
df2 = pd.DataFrame({'Date Test Performed for MB' : file_creation_date_mb, 'WO:' : WO_number_mb, 'Part Number:' : part_number_mb, 'Peak Tensile [lbf] for MB' : peak_tensile_data_mb})
df3 = pd.DataFrame({'Date Test Performed for Hub' : file_creation_date_hub, 'WO:' : WO_number_hub, 'Part Number:' : part_number_hub, 'Peak Tensile [N] for Hub' : peak_tensile_data_hub}) 
df4 = pd.DataFrame({'Date Test Performed for Tip' : file_creation_date_tip, 'WO:' : WO_number_tip, 'Part Number:' : part_number_tip, 'Peak Tensile [N] for Tip' : peak_tensile_data_tip})

#creates new excel file and transfers the dataframe into the new file
writer = ExcelWriter('T:/MAESTRO/MAESTRO/ENGINEERING/TENSILE DATA/TENSILE DATA SPREADSHEET 2.xlsx')
df2.to_excel(writer, 'MB Tensile', index=False)
df3.to_excel(writer, 'Hub Tensile', index=False)
df4.to_excel(writer, 'Tip Tensile', index=False)
writer.save()
  

