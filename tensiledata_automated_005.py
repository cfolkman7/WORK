# -*- coding: utf-8 -*-
"""
Created on Fri Mar 27 14:15:19 2020

@author: connor.folkman

Program written for gathering tensile testing data taken from mark-10 and parsing it into an excel spreadsheet. 
"""
import pandas as pd
from pandas import DataFrame
import glob
import os.path, time
import openpyxl

#creates empty list to be populated with peak tensile test data and date created from each file for the hub
peak_tensile_data_hub = []
file_creation_date_hub = []
WO_number_hub = []
part_number_hub = []
below_spec_hub = []
sheet_name_hub = 'Hub Tensile'

#creates empty list to be populated with peak tensile test data and date created from each file for the tip
peak_tensile_data_tip = []
file_creation_date_tip = []
WO_number_tip = []
part_number_tip = []
below_spec_tip = []
sheet_name_tip = 'Tip Tensile'

#creates empty list to be populated with peak tensile test data and date created from each file for the mb
peak_tensile_data_mb = []
file_creation_date_mb = []
WO_number_mb = []
part_number_mb = []
below_spec_mb = []
sheet_name_mb = 'MB Tensile'

#created empty data frames to put the data into
df = DataFrame()

#read and write spreadsheet for all data
read_file = '//USSJFS1/SJ_PUBLIC/MAESTRO/MAESTRO/ENGINEERING/TENSILE DATA/TENSILE DATA SPREADSHEET.xlsx'

#creates tuple from each list associated with the test and appends that tuple to the tests sheet in the workbook
def append_tup_to_excel(file_date, WO_number, part_number, peak_tensile_data, sheet_name, below_spec_warning):
    i=0
    while i < len(file_date):
        test_tup = (file_date[i], WO_number[i], part_number[i], peak_tensile_data[i], below_spec_warning[i])
        wb = openpyxl.load_workbook(filename=read_file)
        ws = wb[sheet_name]
        ws.append(test_tup)
        wb.save(read_file)
        i+=1

#opens directory where tensile data is saved and reads from each file to find the peak force from colume 'Load' for the hub
for filepath in glob.iglob('//USSJFS1/SJ_PUBLIC/MAESTRO/MAESTRO/TENSILE TESTER DATA FILES/HUB TENSILE/*'):
    
    if filepath[-4:] == 'xlsx':
        df = pd.read_excel(filepath)
        file_creation_date_hub.append(time.strftime('%m/%d/%Y', time.localtime(os.path.getmtime(filepath))))
        WO_number_hub.append(filepath[74:82])
        part_number_hub.append(filepath[83:-5])
        os.rename(filepath, '//USSJFS1/SJ_PUBLIC/MAESTRO/MAESTRO/ENGINEERING/TENSILE DATA/ARCHIVED TENSILE DATA (PRESENT)/HUB/' + filepath[74:])
    elif filepath[-3:] == 'csv':
        df = pd.read_csv(filepath)
        file_creation_date_hub.append(time.strftime('%m/%d/%Y', time.localtime(os.path.getmtime(filepath))))
        WO_number_hub.append(filepath[74:82])
        part_number_hub.append(filepath[83:-5])
        os.rename(filepath, '//USSJFS1/SJ_PUBLIC/MAESTRO/MAESTRO/ENGINEERING/TENSILE DATA/ARCHIVED TENSILE DATA (PRESENT)/HUB/' + filepath[74:])
    else:
        os.rename(filepath, '//USSJFS1/SJ_PUBLIC/MAESTRO/MAESTRO/ENGINEERING/TENSILE DATA/ARCHIVED TENSILE DATA (PRESENT)/HUB/' + filepath[74:] + ' filetype error (not recorded in spreadsheet)')
    
    try: 
        if df['Load [N]'].max() > 0.0:
            peak_tensile_data_hub.append(df['Load [N]'].max())
            if df['Load [N]'].max() < 10.0:
                below_spec_hub.append('LOW')
            else: below_spec_hub.append('  ')
        else: peak_tensile_data_hub.append('error')
    except:
        try: 
            if df['Load [lbF]'].max() > 0.0:
                peak_tensile_data_hub.append(df['Load [lbF]'].max()*4.448)
                if df['Load [lbF]'].max() < 2.24809:
                    below_spec_hub.append('LOW')
                else: below_spec_hub.append('  ')
            else: peak_tensile_data_hub.append('error')
        except:
            peak_tensile_data_hub.append('error')

#opens directory where tensile data is saved and reads from each file to find the peak force from colume 'Load' for the tip
for filepath in glob.iglob('//USSJFS1/SJ_PUBLIC/MAESTRO/MAESTRO/TENSILE TESTER DATA FILES/TIP TENSILE/*'):
    
    if filepath[-4:] == 'xlsx':
        df = pd.read_excel(filepath)
        file_creation_date_tip.append(time.strftime('%m/%d/%Y', time.localtime(os.path.getmtime(filepath))))
        WO_number_tip.append(filepath[74:82])
        part_number_tip.append(filepath[83:-5])
        os.rename(filepath, '//USSJFS1/SJ_PUBLIC/MAESTRO/MAESTRO/ENGINEERING/TENSILE DATA/ARCHIVED TENSILE DATA (PRESENT)/TIP/' + filepath[74:])
    elif filepath[-3:] == 'csv':
        df = pd.read_csv(filepath)
        file_creation_date_tip.append(time.strftime('%m/%d/%Y', time.localtime(os.path.getmtime(filepath))))
        WO_number_tip.append(filepath[74:82])
        part_number_tip.append(filepath[83:-5])
        os.rename(filepath, '//USSJFS1/SJ_PUBLIC/MAESTRO/MAESTRO/ENGINEERING/TENSILE DATA/ARCHIVED TENSILE DATA (PRESENT)/TIP/' + filepath[74:])
    else:
        os.rename(filepath, '//USSJFS1/SJ_PUBLIC/MAESTRO/MAESTRO/ENGINEERING/TENSILE DATA/ARCHIVED TENSILE DATA (PRESENT)/TIP/' + filepath[74:] + ' filetype error (not recorded in spreadsheet)')
        
    try: 
        if df['Load [N]'].max() > 0.0:
            peak_tensile_data_tip.append(df['Load [N]'].max())
            if df['Load [N]'].max() < 10.0:
                below_spec_tip.append('LOW')
            else: below_spec_tip.append('  ')
        else: peak_tensile_data_tip.append('error')
    except:
        try: 
            if df['Load [lbF]'].max() > 0.0:
                peak_tensile_data_tip.append(df['Load [lbF]'].max()*4.448)
                if df['Load [lbF]'].max() < 2.24809:
                    below_spec_tip.append('LOW')
                else: below_spec_tip.append('  ')
            else: peak_tensile_data_tip.append('error')
        except:
            peak_tensile_data_tip.append('error')
               

#opens directory where tensile data is saved and reads from each file to find the peak force from colume 'Load' for the mb
for filepath in glob.glob('//USSJFS1/SJ_PUBLIC/MAESTRO/MAESTRO/TENSILE TESTER DATA FILES/MB TENSILE/*'):
    
    if filepath[-4:] == 'xlsx':
        df = pd.read_excel(filepath)
        file_creation_date_mb.append(time.strftime('%m/%d/%Y', time.localtime(os.path.getmtime(filepath))))
        WO_number_mb.append(filepath[73:81])
        part_number_mb.append(filepath[82:-5])
        os.rename(filepath, '//USSJFS1/SJ_PUBLIC/MAESTRO/MAESTRO/ENGINEERING/TENSILE DATA/ARCHIVED TENSILE DATA (PRESENT)/MB/' + filepath[73:])
    elif filepath[-3:] == 'csv':
        df = pd.read_csv(filepath)
        file_creation_date_mb.append(time.strftime('%m/%d/%Y', time.localtime(os.path.getmtime(filepath))))
        WO_number_mb.append(filepath[73:81])
        part_number_mb.append(filepath[82:-5])
        os.rename(filepath, '//USSJFS1/SJ_PUBLIC/MAESTRO/MAESTRO/ENGINEERING/TENSILE DATA/ARCHIVED TENSILE DATA (PRESENT)/MB/' + filepath[73:])
    else:
        os.rename(filepath, '//USSJFS1/SJ_PUBLIC/MAESTRO/MAESTRO/ENGINEERING/TENSILE DATA/ARCHIVED TENSILE DATA (PRESENT)/MB/' + filepath[73:] + ' filetype error (not recorded in spreadsheet)')
    
    try: 
        if df['Load [lbF]'].max() > 0.0:
            peak_tensile_data_mb.append(df['Load [lbF]'].max())
            if df['Load [lbF]'].max() < 0.25:
                below_spec_mb.append('LOW')
            else: below_spec_mb.append('  ')
        else: peak_tensile_data_mb.append('error')
    except:
        try: 
            if df['Load [N]'].max() > 0.0:
                peak_tensile_data_mb.append(df['Load [N]'].max()/4.448)
                if df['Load [N]'].max() < 1.11206:
                    below_spec_mb.append('LOW')
                else: below_spec_mb.append('  ')
            else: peak_tensile_data_mb.append('error')
        except:
            peak_tensile_data_mb.append('error')
            
    else: continue

append_tup_to_excel(file_creation_date_mb, WO_number_mb, part_number_mb, peak_tensile_data_mb, sheet_name_mb, below_spec_mb)
append_tup_to_excel(file_creation_date_hub, WO_number_hub, part_number_hub, peak_tensile_data_hub, sheet_name_hub, below_spec_hub)
append_tup_to_excel(file_creation_date_tip , WO_number_tip, part_number_tip, peak_tensile_data_tip, sheet_name_tip, below_spec_tip)
