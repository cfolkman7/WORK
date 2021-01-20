# -*- coding: utf-8 -*-
"""
Created on Tue Jul 14 15:49:29 2020

@author: connor.folkman
"""

import tkinter as tk
import serial
import openpyxl
import time 

#dictionary containing the min and max od values for the marker band at smoothing
SM_MIN_MAX = {'2.1F MAX': float(0.72), '2.1F MIN': float(0.67),
              '2.4F MAX': float(0.82), '2.4F MIN': float(0.78),
              '2.8F MAX': float(0.94), '2.8F MIN': float(0.90),
              '2.9F MAX': float(0.97), '2.9F MIN': float(0.93)}

#gets the current time 
current_time = time.localtime()

beg_mid_end = ' '

#global variable that manages which cell the program is being run on
cell_num = '1'

#function to change cell number
def cell_num_change(num):
    global cell_num 
    cell_num = str(num)

#function creates window that shows button selection for cell number 
def window_cell():
    #function creates window that displays french size choices that will be run in the cell
    def window_21F(french_size):
        #gathers data from lasermic and diplays on screen with ability to save data to excel
        def measure(french_size):
            #function saves each measurement to an excel file 
            def store_data(string, data, french_size):
                if (data > SM_MIN_MAX[french_size + ' MIN']) & (data < SM_MIN_MAX[french_size + ' MAX']):
                    tup_measure = (string, data, time.strftime("%Y-%m-%d", current_time), time.strftime("%H:%M:%S", current_time), 'GOOD')
                else: 
                    if data < SM_MIN_MAX[french_size + ' MIN']:
                        tup_measure = (string, data, time.strftime("%Y-%m-%d", current_time), time.strftime("%H:%M:%S", current_time), 'LOW')
                    if data > SM_MIN_MAX[french_size + ' MAX']:
                        tup_measure = (string, data, time.strftime("%Y-%m-%d", current_time), time.strftime("%H:%M:%S", current_time), 'HIGH')
                wb = openpyxl.load_workbook(filename='T:\MAESTRO\MAESTRO\ENGINEERING\MB_DATA_SMOOTHING\MB_DATA_CELL' + cell_num + '.xlsx')
                ws = wb[french_size + ' DATA']
                ws.append(tup_measure)
                wb.save('T:\MAESTRO\MAESTRO\ENGINEERING\MB_DATA_SMOOTHING\MB_DATA_CELL' + cell_num + '.xlsx')
                
            with serial.Serial('COM9', 115200, stopbits = 1, parity='N', bytesize=8, timeout=.5) as ser:
                char = 'H\n'
                ser.write(bytes(char, 'ASCII'))
                try:
                    lasermik_data = str(ser.read(8).decode('utf-8'))
                    lasermik_data = lasermik_data[3:4] + '.' + lasermik_data[4:7]
                    label_lasermik_data.config(text=lasermik_data +'mm')
                except UnicodeDecodeError:
                    label_lasermik_data.config(text='RE-MEASURE')
            
            button_pass.config(text = 'PASS', command=lambda: store_data('PASS', float(lasermik_data[:7]), french_size))
            button_fail.config(text = 'RE-SMOOTH', command=lambda: store_data('RE-SMOOTHED', float(lasermik_data[:7]), french_size))
            button_fail.pack()
            
        #this function takes multiple measurements to create an average and saves data to excel
        def average(french_size):
            average_data = []
            
            #this function facilitates the beg, mid, and end sampling 
            def button_pass_reveal(string):
                global beg_mid_end
                beg_mid_end = string
                label_lasermik_data.config(text='Press "MEASURE" Button and \n slightly rotate MB', font=("Courier, 80"))
                button_beg.pack_forget()
                button_fail.pack_forget()
                button_mid.pack_forget()
                button_end.pack_forget()
                button_pass.config(text = 'MEASURE', command = lambda: measure_ave(average_data, french_size))
                button_pass.pack()
            
            #This function allows data to be read from lasermic and stores the 8 measurements to calc the ave
            def measure_ave(average_data, french_size):
                with serial.Serial('COM9', 115200, stopbits = 1, parity='N', bytesize=8, timeout=1) as ser:
                    char = 'H\n'
                    ser.write(bytes(char, 'ASCII'))
                    try:
                        data = str(ser.read(8).decode('utf-8'))
                        data = data[3:4] + '.' + data[4:7]
                        average_data.append(float(data))
                    except UnicodeDecodeError:
                        label_lasermik_data.config(text='Error: RE-MEASURE')
                if len(average_data) > 0:
                    running_ave = 0
                    i = 0
                    while i < len(average_data):
                        running_ave = average_data[i] + running_ave
                        i = i + 1
                    label_lasermik_data.config(text='Average: ' + str(round((running_ave/len(average_data)), 3)) + '\n Measurement: ' + str(len(average_data)))
                if len(average_data) == 8:
                    i = 0
                    average_calc = 0.0
                    average_num = 0.0
                    while i < 8:
                        average_calc += average_data[i]
                        i+=1
                    average_num = round((average_calc/8), 3)
                    str_average_num = str(average_num)
                    label_lasermik_data.config(text='MEASUREMENTS COMPLETED \n AVERAGE: ' + str_average_num[:5] + 'mm')
                    if (average_num > SM_MIN_MAX[french_size + ' MIN']) & (average_num < SM_MIN_MAX[french_size + ' MAX']):
                        tup_average = ('PASS', beg_mid_end, average_num, time.strftime("%Y-%m-%d", current_time), time.strftime("%H:%M:%S", current_time), 'GOOD', beg_mid_end)
                    else: 
                        if average_num < SM_MIN_MAX[french_size + ' MIN']:
                            print(beg_mid_end)
                            tup_average = ('FAIL', average_num, time.strftime("%Y-%m-%d", current_time), time.strftime("%H:%M:%S", current_time), 'LOW', beg_mid_end)
                        if average_num > SM_MIN_MAX[french_size + ' MAX']:
                            tup_average = ('FAIL', average_num, time.strftime("%Y-%m-%d", current_time), time.strftime("%H:%M:%S", current_time), 'HIGH', beg_mid_end)
                    wb = openpyxl.load_workbook(filename='T:\MAESTRO\MAESTRO\ENGINEERING\MB_DATA_SMOOTHING\MB_DATA_CELL' + cell_num + '.xlsx')
                    ws = wb[french_size + ' DATA']
                    ws.append(tup_average)
                    wb.save('T:\MAESTRO\MAESTRO\ENGINEERING\MB_DATA_SMOOTHING\MB_DATA_CELL' + cell_num + '.xlsx')
                    
            button_fail.pack_forget()
            button_pass.pack_forget()
            
            button_beg = tk.Button(master = frame4, text = 'BEGINNING', font = ('Courier, 90'), command = lambda: button_pass_reveal('BEG'))
            button_beg.pack(side='left', padx = 10)
            
            button_mid = tk.Button(master = frame4, text = 'MIDDLE', font = ('Courier, 90'), command = lambda: button_pass_reveal('MID'))
            button_mid.pack(side='left', padx = 10)
            
            button_end = tk.Button(master = frame4, text = 'END', font = ('Courier, 90'), command = lambda: button_pass_reveal('END'))
            button_end.pack(side='left', padx = 10)
            
            label_lasermik_data.config(text='Press "BEG" "MID" "END" Buttons \n to select sample', font=("Courier, 80"))
            
        window_21F = tk.Tk()
        window_21F.state('zoomed')
        
        frame1 = tk.Frame(master = window_21F)
        frame2 = tk.Frame(master = window_21F)
        frame3 = tk.Frame(master = window_21F)
        frame4 = tk.Frame(master = window_21F)
        frame1.pack()
        frame2.pack()
        frame3.pack()
        frame4.pack()
        
        label_title = tk.Label(master=frame1, text=french_size + ' MEASUREMENT', font=("Courier, 90"))
        button_measure = tk.Button(master = frame2, text = 'MEASURE', font = ('Courier, 90'), command = lambda: measure(french_size))
        button_average = tk.Button(master = frame2, text = 'AVERAGE', font = ('Courier, 90'), command = lambda: average(french_size))
        label_lasermik_data = tk.Label(master=frame3, text='0.000', font=("Courier, 140"))
        button_pass = tk.Button(master = frame4, font = 'Times 90', bg='green')
        button_fail = tk.Button(master = frame4, font = 'Times 90', bg='red')
        label_title.pack()
        button_fail.pack(side='right', padx = 30, fill='x')
        button_pass.pack(side='left', padx = 30, fill='x')
        button_measure.pack(side="left", padx=20)
        button_average.pack(side="right", padx = 20)
        label_lasermik_data.pack()
        window_21F.mainloop()
        
    window_cell = tk.Tk()
    window_cell.state('zoomed')
    
    button_21F = tk.Button(master=window_cell, text='2.1F', bg='blue', fg='white', font=("Courier, 90"), command=lambda: window_21F('2.1F'))
    button_21F.pack(side='left', expand=2, fill='both', pady=5, padx=5)
    
    button_24F = tk.Button(master=window_cell, text='2.4F', bg='red', fg='white', font=("Courier, 90"), command=lambda: window_21F('2.4F'))
    button_24F.pack(side='left', expand=2, fill='both', pady=5, padx=5)
    
    button_28F = tk.Button(master=window_cell, text='2.8F', bg='purple', fg='white', font=("Courier, 90"), command=lambda: window_21F('2.8F'))
    button_28F.pack(side='left', expand=2, fill='both', pady=5, padx=5)
    
    button_29F = tk.Button(master=window_cell, text='2.9F', bg='green', fg='white', font=("Courier, 90"), command=lambda: window_21F('2.9F'))
    button_29F.pack(side='left', expand=2, fill='both', pady=5, padx=5)
    
    window_cell.mainloop()
    
root_window = tk.Tk()
root_window.state('zoomed')

button_cell1 = tk.Button(master=root_window, text='CELL #1', bg='red', fg='white', font=("Courier, 90"), command=lambda: [cell_num_change('1'), window_cell()])
button_cell1.pack(side='left', fill='y', pady=5, padx=5)

button_cell2 = tk.Button(master=root_window, text='CELL #2', bg='blue', fg='white', font=("Courier, 90"), command=lambda: [cell_num_change('2'), window_cell()])
button_cell2.pack(side='left', fill='y', pady=5, padx=5)

button_cell3 = tk.Button(master=root_window, text='CELL #3', bg='green', fg='white', font=("Courier, 90"), command=lambda: [cell_num_change('3'), window_cell()])
button_cell3.pack(side='left', fill='y', pady=5, padx=5)

root_window.mainloop()