# -*- coding: utf-8 -*-
"""
Created on Fri Jan 22 11:29:37 2021

@author: connor.folkman
"""

import telnetlib
import time
import openpyxl
from pandas import DataFrame
from shutil import copyfile

#List for Outer Diameter Marker Band Measurement
OD_measure_list_24 = []
OD_measure_list_HF = []
OD_measure_list_29Mandrel = []
OD_measure_list_28Mandrel = []
OD_measure_list_24Mandrel = []

#function to append MB OD data to excel file
def to_excel_mb(ave, measure_max, measure_min, std, var, med, quan25, quan50, quan75, sheet_name, date, time):
    test_tup = (ave, measure_max, measure_min, std, var, med, quan25, quan50, quan75, date, time)
    wb = openpyxl.load_workbook(filename='//USSJFS1/SJ_PUBLIC/MAESTRO/MAESTRO/ENGINEERING/MB_DATA_SMOOTHING/MB_DATA_ETHERNET_TEST.xlsx')
    ws = wb[sheet_name]
    ws.append(test_tup)
    wb.save(filename='//USSJFS1/SJ_PUBLIC/MAESTRO/MAESTRO/ENGINEERING/MB_DATA_SMOOTHING/MB_DATA_ETHERNET_TEST.xlsx')

#function to append the mandrel OD data to excel file
def to_excel_mandrel(ave, sheet_name, date, time):
    test_tup = (ave, date, time)
    wb = openpyxl.load_workbook(filename='//USSJFS1/SJ_PUBLIC/MAESTRO/MAESTRO/ENGINEERING/MB_DATA_SMOOTHING/MB_DATA_ETHERNET_TEST.xlsx')
    ws = wb[sheet_name]
    ws.append(test_tup)
    wb.save(filename='//USSJFS1/SJ_PUBLIC/MAESTRO/MAESTRO/ENGINEERING/MB_DATA_SMOOTHING/MB_DATA_ETHERNET_TEST.xlsx')

#main function where data collection and sparsing takes place
def main():
    #Empty dataframe for statistical analysis of data
    df = DataFrame()
    
    #bool to determine connection to lasermic
    connection_bool = False
    
    #counts number of data points after 1000 data points a backup copy of the excel file will be created
    data_points = 0
    
    while connection_bool == False:
        try:   
            tn = telnetlib.Telnet('10.37.20.145')
            time.sleep(.1)
            tn.write(bytes('I', 'ascii'))
            time.sleep(.1)
            connection_bool = True
            
        except TimeoutError:
            time.sleep(5)
            continue
    try:
        while True:
            #gets current time 
            current_time = time.localtime()
            
            tn.write(bytes('J', 'ascii'))
            open_gate = tn.read_until(b"\r").decode()
            open_gate = open_gate[6]
            tn.write(bytes('F', 'ascii'))
            gate_position = tn.read_until(b"\r").decode()
            gate_position = int(gate_position[2:4])
            tn.write(bytes('D', 'ascii'))
            OD_measure = tn.read_until(b"\r").decode()
            OD_measure = float('0.' + OD_measure[2:5])
            
            if open_gate == '0':
                if gate_position < 15:
                    if (OD_measure > 0.77) & (OD_measure < 0.86):
                        OD_measure_list_24.append(OD_measure)
                        
                    elif (OD_measure > 0.88) & (OD_measure < 1.1):
                        OD_measure_list_HF.append(OD_measure)
                        
                    elif (OD_measure > .495) & (OD_measure < .53):
                        OD_measure_list_24Mandrel.append(OD_measure)
                        
                    elif (OD_measure > .6) & (OD_measure < .63):
                        OD_measure_list_28Mandrel.append(OD_measure)
                        
                    elif (OD_measure > .665) & (OD_measure < .685):
                        OD_measure_list_29Mandrel.append(OD_measure)
                        
            elif open_gate == '1':
                if len(OD_measure_list_24) > 200:
                    df = DataFrame(OD_measure_list_24,columns=['MB OD (mm)'])
                    OD_measure_ave = round(df['MB OD (mm)'].mean(), 3)
                    OD_measure_max = round(df['MB OD (mm)'].max(), 3)
                    OD_measure_min = round(df['MB OD (mm)'].min(), 3)
                    OD_measure_std = round(df['MB OD (mm)'].std(), 4)
                    OD_measure_var = round(df['MB OD (mm)'].var(), 5)
                    OD_measure_med = round(df['MB OD (mm)'].median(), 3)
                    OD_measure_quan25 = round(df['MB OD (mm)'].quantile(q=.25), 3)
                    OD_measure_quan50 = round(df['MB OD (mm)'].quantile(q=.50), 3)
                    OD_measure_quan75 = round(df['MB OD (mm)'].quantile(q=.75), 3)
                    to_excel_mb(OD_measure_ave, OD_measure_max, OD_measure_min, OD_measure_std, OD_measure_var, OD_measure_med, OD_measure_quan25, 
                             OD_measure_quan50, OD_measure_quan75, '2.4F', time.strftime("%Y-%m-%d", current_time), time.strftime("%H:%M:%S", current_time))
                    OD_measure_list_24.clear()
                    data_points+=1
                    
                elif len(OD_measure_list_HF) > 200:
                    df = DataFrame(OD_measure_list_HF,columns=['MB OD (mm)'])
                    OD_measure_ave = round(df['MB OD (mm)'].mean(), 3)
                    OD_measure_max = round(df['MB OD (mm)'].max(), 3)
                    OD_measure_min = round(df['MB OD (mm)'].min(), 3)
                    OD_measure_std = round(df['MB OD (mm)'].std(), 4)
                    OD_measure_var = round(df['MB OD (mm)'].var(), 5)
                    OD_measure_med = round(df['MB OD (mm)'].median(), 3)
                    OD_measure_quan25 = round(df['MB OD (mm)'].quantile(q=.25), 3)
                    OD_measure_quan50 = round(df['MB OD (mm)'].quantile(q=.50), 3)
                    OD_measure_quan75 = round(df['MB OD (mm)'].quantile(q=.75), 3)
                    if OD_measure_ave < 0.96:
                        to_excel_mb(OD_measure_ave, OD_measure_max, OD_measure_min, OD_measure_std, OD_measure_var, OD_measure_med, OD_measure_quan25, 
                                 OD_measure_quan50, OD_measure_quan75, '2.8F', time.strftime("%Y-%m-%d", current_time), time.strftime("%H:%M:%S", current_time))
                        
                    elif OD_measure_ave > 0.96:
                        to_excel_mb(OD_measure_ave, OD_measure_max, OD_measure_min, OD_measure_std, OD_measure_var, OD_measure_med, OD_measure_quan25, 
                                 OD_measure_quan50, OD_measure_quan75, '2.9F', time.strftime("%Y-%m-%d", current_time), time.strftime("%H:%M:%S", current_time))
                        
                    OD_measure_list_HF.clear()
                    data_points+=1
                    
                elif len(OD_measure_list_24Mandrel) > 100:
                    df = DataFrame(OD_measure_list_24Mandrel,columns=['Mandrel OD (mm)'])
                    OD_measure_ave = round(df['Mandrel OD (mm)'].mean(), 3)
                    to_excel_mandrel(OD_measure_ave, '24MANDRELS', time.strftime("%Y-%m-%d", current_time), time.strftime("%H:%M:%S", current_time))
                    OD_measure_list_24Mandrel.clear()
                    
                elif len(OD_measure_list_28Mandrel) > 100:
                    df = DataFrame(OD_measure_list_28Mandrel,columns=['Mandrel OD (mm)'])
                    OD_measure_ave = round(df['Mandrel OD (mm)'].mean(), 3)
                    to_excel_mandrel(OD_measure_ave, '28MANDRELS', time.strftime("%Y-%m-%d", current_time), time.strftime("%H:%M:%S", current_time))
                    OD_measure_list_28Mandrel.clear()
                    
                elif len(OD_measure_list_29Mandrel) > 100:
                    df = DataFrame(OD_measure_list_29Mandrel,columns=['Mandrel OD (mm)'])
                    OD_measure_ave = round(df['Mandrel OD (mm)'].mean(), 3)
                    to_excel_mandrel(OD_measure_ave, '29MANDRELS', time.strftime("%Y-%m-%d", current_time), time.strftime("%H:%M:%S", current_time))
                    OD_measure_list_29Mandrel.clear()
                    
                else: OD_measure_list_24.clear(), OD_measure_list_HF.clear(), OD_measure_list_24Mandrel.clear(), OD_measure_list_28Mandrel.clear(), OD_measure_list_29Mandrel.clear()
                
            if data_points == 5:
                data_points = 0
                copyfile('//USSJFS1/SJ_PUBLIC/MAESTRO/MAESTRO/ENGINEERING/MB_DATA_SMOOTHING/MB_DATA_ETHERNET_TEST.xlsx', '//USSJFS1/SJ_PUBLIC/MAESTRO/MAESTRO/ENGINEERING/MB_DATA_SMOOTHING/MB_DATA_ETHERNET_TEST_BACKUP.xlsx')
                
    except KeyboardInterrupt:
        tn.close()
        
    except:
        tn.close()
        time.sleep(5)
        main()

main()



