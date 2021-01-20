# -*- coding: utf-8 -*-
"""
Created on Wed Aug  5 16:32:02 2020

@author: connor.folkman
"""

import tkinter as tk
import serial
import time
import openpyxl
from openpyxl import Workbook
import matplotlib.pyplot as plt
import os  

bool_stop = True

WO = ''

PN = ''

tensile_max = 0.0

def main():

    def submit():
        global WO, PN
        WO = WO_entry.get()
        PN = PN_entry.get()
        if (len(WO) == 8 & len(PN) == 8):
            test_selection()
        else: 
            warning_label.config(text = 'Error: Check the work order and part number, \n than re-submit.')
            
    def test_selection():
        def test(test_type):
            
            tensile_data_list = []
            unit = 0
            
            if test_type == 'MB':
                unit = -6
            else: unit = -5
            
            def stop_test():
                global bool_stop
                bool_stop = False
            
            def max_tensile():
                global tensile_max
                i = 0
                min_tensile = 0.0
                if test_type == 'MB':
                    min_tensile = 0.25
                else: min_tensile = 10.0
                while i < len(tensile_data_list):
                    if tensile_data_list[i] > tensile_max:
                        tensile_max = tensile_data_list[i]
                    i+=1
                MB_label.config(text = "Max Tensile: " + str(float('{:.2f}'.format(tensile_max))))
                if tensile_max > min_tensile:
                    MB_label3.config(text = 'Test: Pass')
                else: MB_label3.config(text = 'Test: Fail \n Alert Engineering')
                
            def test_begin():
                if bool_stop:
                    try:
                        with serial.Serial('COM4', 115200, stopbits = 1, parity='N', bytesize=8) as ser:
                            char = 'X'
                            ser.write(bytes(char, 'ASCII'))
                            tensile_data = ser.readline().decode('UTF-8')
                            tensile_data_list.append(abs(float(tensile_data[:unit]))) 
                            MB_Stop.config(text = 'Click to Stop Data Collection', command = lambda: [(stop_test(), max_tensile())])
                            MB_label3.config(text = str(abs(float(tensile_data[:unit])))+tensile_data[unit:])
                            MB_label3.after(150, test_begin)
                    except PermissionError:
                        ser.open()
                        time.sleep(.05)
                        with serial.Serial('COM4', 115200, stopbits = 1, parity='N', bytesize=8) as ser:
                            char = '?'
                            ser.write(bytes(char, 'ASCII'))
                            tensile_data = ser.readline().decode('UTF-8')
                            tensile_data_list.append(abs(float(tensile_data[:unit]))) 
                            MB_Stop.config(text = 'Click to Stop Data Collection', command = lambda: [(stop_test(), max_tensile())])
                            MB_label3.config(text = str(abs(float(tensile_data[:unit])))+tensile_data[unit:])
                            MB_label3.after(150, test_begin)
                            
                else: 
                    def destroy_window():
                        global bool_stop, tensile_max
                        bool_stop = True
                        tensile_max = 0.0
                        time.sleep(.5)
                        window_test.destroy()
                        window_test_selection.destroy()
                        root_window.destroy()
                        main()
                    
                    def print_out():
                        wb = Workbook()
                        ws = wb.active
                        plt.plot(tensile_data_list)
                        if test_type == 'MB':
                            plt.ylabel('Tensile Force (lbF)')
                        else: plt.ylabel('Tensile Force (N)')
                        plt.xlabel('Read points')
                        plt.savefig('tensileplot.png', dpi = 110)
                        img = openpyxl.drawing.image.Image('tensileplot.png')
                        tup_WO = ('WO: ', WO)
                        tup_TIME = ('Date Performed:', time.strftime("%m-%d-%y", time.localtime()))
                        tup_PN = ('PN: ', PN)
                        if test_type == 'MB':
                            tup_MAX = ('MAX FORCE (lbF): ', tensile_max)
                        else: tup_MAX = ('MAX FORCE (N): ', tensile_max)
                        ws.append(tup_TIME)
                        ws.append(tup_WO)
                        ws.append(tup_PN)
                        ws.append(tup_MAX)
                        ws.column_dimensions['A'].width = 15
                        ws.column_dimensions['B'].width = 10
                        ws.add_image(img, 'A5')
                        wb.save('T:/MAESTRO/MAESTRO/TENSILE TESTER DATA FILES/CHATILLON TESTER FILES/GRAPH.xlsx')
                        #os.startfile('T:/MAESTRO/MAESTRO/TENSILE TESTER DATA FILES/CHATILLON TESTER FILES/GRAPH.xlsx', 'print')
                        plt.clf()
                        
                    def save_data():  
                        tup = (WO, PN, tensile_max, time.strftime("%m-%d-%y", time.localtime()))
                        write_file = 'T:/MAESTRO/MAESTRO/TENSILE TESTER DATA FILES/CHATILLON TESTER FILES/TENSILE_DATA.xlsx'
                        wb = openpyxl.load_workbook(filename=write_file)
                        ws = wb[test_type + ' Tensile']
                        ws.append(tup)
                        wb.save(write_file) 
                    
                    print_out()
                    save_data()
                    MB_Stop.config(text = 'Press to Continue', command = lambda: destroy_window())
                    
            window_test = tk.Tk()
            window_test.state('zoomed')
            
            MB_label = tk.Button(master = window_test, text = 'Click to Start Data Collection', font = ('Courier, 90'), command = lambda: test_begin())
            MB_label.pack()
            
            MB_label3 = tk.Label(master = window_test, text = '0.000', font = ('Courier, 90'))
            MB_label3.pack()
            
            MB_Stop = tk.Button(master = window_test, font = ('Courier, 90'))
            MB_Stop.pack()
            
            window_test.mainloop()
        
        window_test_selection = tk.Tk()
        window_test_selection.state('zoomed')
        
        button_MB = tk.Button(master=window_test_selection, text = 'MB TENSILE TEST', font = ('Courier, 90'), background = 'blue', command = lambda: test('MB'))
    
        button_TIP = tk.Button(master=window_test_selection, text = 'TIP TENSILE TEST', font = ('Courier, 90'), background = 'red', command = lambda: test('Tip'))
        
        button_HUB = tk.Button(master=window_test_selection, text = 'HUB TENSILE TEST', font = ('Courier, 90'), background = 'orange', command = lambda: test('Hub'))
        
        button_MB.pack(fill = 'both', pady=30)
        button_TIP.pack(fill = 'both', pady=30)
        button_HUB.pack(fill = 'both', pady=30)
        
        window_test_selection.mainloop() 
            
    root_window = tk.Tk()
    root_window.state('zoomed')
    
    label = tk.Label(root_window, font = ('Courier, 35'), text = 'Enter Work Order number and Part number than hit submit.')
    label.pack()
    
    frame1 = tk.Frame()
    frame2 = tk.Frame()
    frame1.pack()
    frame2.pack()
    
    label2 = tk.Label(frame1, font = ('Courier, 90'), text = 'WO:')
    label2.pack(side='left')
    
    WO_entry = tk.Entry(frame1, font = ('Courier, 90'), text = 'WO:')
    WO_entry.pack(side='right')
    
    label3 = tk.Label(frame2, font = ('Courier, 90'), text = 'PN:')
    label3.pack(side='left')
    
    PN_entry = tk.Entry(frame2, font = ('Courier, 90'), text = 'PN:')
    PN_entry.pack(side='right')
    
    submit_button = tk.Button(root_window, font = ('Courier, 50'), text = 'Submit', command = lambda: submit())
    submit_button.pack()
    
    warning_label = tk.Label(root_window, font = ('Courier, 50'))
    warning_label.pack()
    
    root_window.mainloop()
    
if __name__ == '__main__':
    main()