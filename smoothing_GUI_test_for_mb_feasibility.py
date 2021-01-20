# -*- coding: utf-8 -*-
"""
Created on Tue Jul 14 15:49:29 2020

Used for gathering marker band data that was needed for the qualification of a new MB vendor

Used on BETA lasermike AS4012 unit with RS232 connections

@author: connor.folkman
"""

import tkinter as tk
import serial
import openpyxl

#list of marker band vendors used in test
vendor_list = ['stanford', 'lakeregion', 'prince']

#list of french sizes used on Maestro
frenchsize_list = ['2.9', '2.8', '2.4', '2.1']

#list of swaging dyes used on Maestro
swagdye_list = ['2-4-3', '2-4-4', '2-4-C',
                '2-8-1', '2-8-4',
                '2-9-3', '2-9-4', '2-9-A', '2-9-B', '2-9-10']

#global variable that manages which swaging dye was used
swagdye_num = 'none'

#global variable that manages the vendor 
vendorname = 'None'

#global variable that manages the french size
frenchsize = 'None'

#global variable that manages the number of measurements used to calc the average
ave_num = 6

#function that verifies the input information from the first window (french size, vendor, swag dye)
def submit():
    global swagdye_num, vendorname, frenchsize
    i = 0
    bool_swagdye = False
    bool_vendor = False
    bool_frenchsize = False
    swagdye_num = str(input_swagdye.get())
    vendorname = str(input_vendorname.get())
    frenchsize = str(input_frenchsize.get())
    while i < len(swagdye_list):
        if swagdye_num == swagdye_list[i]:
            bool_swagdye = True
            break
        i=i+1
    i=0
    while i < len(vendor_list):
        if vendorname == vendor_list[i]:
            bool_vendor = True
            break
        i=i+1
    i=0
    while i < len(frenchsize_list):
        if frenchsize == frenchsize_list[i]:
            bool_frenchsize = True
            break
        i=i+1
    swagdye_frenchsize = swagdye_num[0] + '.' + swagdye_num[2]
    if bool_swagdye == True & bool_vendor == True & bool_frenchsize == True & (swagdye_frenchsize == frenchsize) == True: 
        warning_label.config(text='')
        measure()
    elif bool_swagdye == False: 
        warning_label.config(text = 'Error: Check the swag die # and try again')
    elif bool_vendor == False: 
        warning_label.config(text = 'Error: Check the vendor and try again')
    elif bool_frenchsize == False: 
        warning_label.config(text = 'Error: Check the french size and try again')
    elif str(swagdye_frenchsize) != str(frenchsize):
        warning_label.config(text = 'Error: Swag die # and Frenchsize dont match')
    else: 
        warning_label.config(text = 'Error')

#function used to take single measurements of data 
def measure():
    def measurement():
        with serial.Serial('COM9', 115200, stopbits = 1, parity='N', bytesize=8, timeout=.5) as ser:
            char = 'H\n'
            ser.write(bytes(char, 'ASCII'))
            try:
                lasermik_data = str(ser.read(8).decode('utf-8'))
                lasermik_data = lasermik_data[3:4] + '.' + lasermik_data[4:8]
                label_lasermik_data.config(text=lasermik_data +'mm')
            except UnicodeDecodeError:
                label_lasermik_data.config(text='RE-MEASURE')
            
        button_pass.config(text = 'VISUAL PASS', command=lambda: store_data('PASS', float(lasermik_data[:8])))
        button_pass.pack(side='left')
        button_fail.config(text = 'VISUAL FAIL', command=lambda: store_data('FAIL', float(lasermik_data[:8])))
        button_fail.pack(side='right')
        
    #function used to store the single measurement data in an excel file
    def store_data(string, data):
        tup_measure = (data, string, 1, swagdye_num)
        wb = openpyxl.load_workbook(filename='T:\MAESTRO\MAESTRO\ENGINEERING\MAESTRO_NEW_MB_VENDOR\MB FEASIBILITY DATA\\' + frenchsize + 'DATA' + '.xlsx')
        ws = wb[vendorname.upper()]
        ws.append(tup_measure)
        wb.save('T:\MAESTRO\MAESTRO\ENGINEERING\MAESTRO_NEW_MB_VENDOR\MB FEASIBILITY DATA\\' + frenchsize + 'DATA' + '.xlsx')
    
    #function that controls the 8 measurments
    def average():
        
        average_data = []
        button_fail.pack_forget()
        label_lasermik_data.config(text='Press "Measure (Ave)" Button and \n slightly rotate MB', font=("Courier, 80"))
        button_pass.config(text = 'Measure (Ave)', command = lambda: start_ave())
        button_pass.pack()
        
        #function takes the 8 measurments and calculates the average
        def start_ave():
            failure_note = tk.Text(frame4, font = ("Courier, 45"))
            def forget_note():
                failure_note.pack_forget()
            def store_ave(average_num, string):
                    tup_average = (average_num, string, ave_num, swagdye_num)
                    wb = openpyxl.load_workbook(filename='T:\MAESTRO\MAESTRO\ENGINEERING\MAESTRO_NEW_MB_VENDOR\MB FEASIBILITY DATA\\' + frenchsize + 'DATA' + '.xlsx')
                    ws = wb[vendorname.upper()]
                    ws.append(tup_average)
                    wb.save('T:\MAESTRO\MAESTRO\ENGINEERING\MAESTRO_NEW_MB_VENDOR\MB FEASIBILITY DATA\\' + frenchsize + 'DATA' + '.xlsx')
                    
            with serial.Serial('COM9', 115200, stopbits = 1, parity='N', bytesize=8, timeout=1) as ser:
                char = 'H\n'
                ser.write(bytes(char, 'ASCII'))
                try:
                    data = str(ser.read(8).decode('utf-8'))
                    data = data[3:4] + '.' + data[4:8]
                    average_data.append(float(data))
                except UnicodeDecodeError:
                    label_lasermik_data.config(text='Error: RE-MEASURE')
            if len(average_data) > 0:
                running_ave = 0
                i = 0
                while i < len(average_data):
                    running_ave = average_data[i] + running_ave
                    i = i + 1
                label_lasermik_data.config(text='Running Average: ' + str(round((running_ave/len(average_data)), 3)) + 'mm' + '\n Measurement: ' + str(len(average_data)))
            if len(average_data) == ave_num:
                i = 0
                average_calc = 0.0
                average_num = 0.0
                while i < ave_num:
                    average_calc += average_data[i]
                    i+=1
                average_num = round((average_calc/ave_num), 3)
                str_average_num = str(average_num)
                label_lasermik_data.config(text='MEASUREMENTS COMPLETED \n AVERAGE: ' + str_average_num[:6] + 'mm')
                
                failure_note.pack(side='left', expand='YES', fill='both')
                button_pass.config(text='Submit Note', command = lambda: [average(), store_ave(average_num, failure_note.get('1.0','end')), forget_note()])
                button_pass.pack(side='right')
            
    window_measure = tk.Tk()
    window_measure.state('zoomed')
    
    frame1 = tk.Frame(master = window_measure)
    frame2 = tk.Frame(master = window_measure)
    frame3 = tk.Frame(master = window_measure)
    frame4 = tk.Frame(master = window_measure)
    frame1.pack()
    frame2.pack()
    frame3.pack()
    frame4.pack()
    
    label_title = tk.Label(master=frame1, text=frenchsize + ' MEASUREMENT', font=("Courier, 90"))
    button_measure = tk.Button(master = frame2, text = 'MEASURE', font = ('Courier, 90'), command = lambda: measurement())
    button_average = tk.Button(master = frame2, text = 'AVERAGE', font = ('Courier, 90'), command = lambda: average())
    label_lasermik_data = tk.Label(master=frame3, text='0.000', font=("Courier, 120"))
    button_pass = tk.Button(frame4, font=("Courier, 75"), text = 'VISUAL GOOD')
    button_fail = tk.Button(frame4, font=("Courier, 75"), text = 'VISUAL BAD')
    label_title.pack()
    button_fail.pack(side='right', padx = 30, fill='x')
    button_pass.pack(side='left', padx = 30, fill='x')
    button_measure.pack(side="left", padx=20)
    button_average.pack(side="right", padx = 20)
    label_lasermik_data.pack()
    
    window_measure.mainloop()
            
root_window = tk.Tk()
root_window.state('zoomed')

frame1 = tk.Frame()
frame2 = tk.Frame()
frame3 = tk.Frame()
frame4 = tk.Frame()
frame1.pack()
frame2.pack()
frame3.pack()
frame4.pack()

input_swagdye = tk.Entry(master=frame1, font=("Courier, 70"))
input_swagdye.pack(side='right')

label_swagdye = tk.Label(master=frame1, font=("Courier, 70"), text='Swaging Die #:')
label_swagdye.pack(side='left')

input_vendorname = tk.Entry(master=frame2, font=("Courier, 70"))
input_vendorname.pack(side = 'right')

label_vendorname = tk.Label(master=frame2, font=("Courier, 70"), text='Vendor:')
label_vendorname.pack(side = 'left')

input_frenchsize = tk.Entry(master=frame3, font=("Courier, 70"))
input_frenchsize.pack(side='right')

label_frenchsize = tk.Label(master=frame3, font=("Courier, 70"), text='French Size:')
label_frenchsize.pack(side = 'left')

button_submit = tk.Button(master=frame4, font=("Courier, 70"), text='Submit', command = lambda: submit())
button_submit.pack()

warning_label = tk.Label(master=frame4, font = ('Courier, 50'))
warning_label.pack()
    
root_window.mainloop()


    
    
    