
# IMPORTING LIBRARIES -------------------------------------
#region
import PySimpleGUI as sg
import pandas as pd
import numpy as np
import os
import re
import glob
import csv
import shutil
import datetime
from datetime import datetime
print('Libs Imported')
#endregion

# GUI ------------------------------------------------------

sg.theme('DarkAmber')   # Add a touch of color
# All the stuff inside your window.
layout = [  [sg.Text('Example -   C:\FILES   ')],
            [sg.Text('Enter folder path'), sg.InputText()],
            [sg.Text('Example -    data if data.csv    ')],
            [sg.Text('Enter file name'), sg.InputText()],
            [sg.Text('CLICK ON OK TO MAKE IT RUN')],
            [sg.Button('Ok'), sg.Button('Cancel')] 
        ]

# Create the Window
window = sg.Window('Window Title', layout)
# Event Loop to process "events" and get the "values" of the inputs
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Cancel': # if user closes window or clicks cancel
        break
    input_folder = str(values[0])
    
    input_filename = str(values[1])

    # LOGIC TO DO ---------------------------------------------------------
    input_filename = input_filename + '.csv'
    filepath_name =str(input_folder + str('\\') + input_filename)
    print('input')
    print(filepath_name)
    output_folder = input_folder.replace('\\','/')
    output_filenamepath = output_folder + '/PPD-Data.xlsx'
    print('output is here')
    print(output_filenamepath)

    # READ FILE -----------------------------------------------
    print('reading csv file')
    df_file = pd.read_csv(filepath_name)                
    print('Datafram object generated from csv file')

    # Delete Unessesary Rows ----------------------------------

    print('deleting non PPD rows')
    print('')
    # String to be searched in start of string
    search_term ="PPD"
    # boolean series returned 
    Bool_Ser_Filter = df_file["Event name"].str.startswith(search_term)
    df_file = df_file[Bool_Ser_Filter]
    df_file.reset_index(inplace = True, drop = True)
    #replace blank teachers with 'Unkown, Teacher'
    values = {'Teachers': 'Unkown, Teacher'}
    df_file = df_file.fillna(value=values)
    print('deleted')

    #Loop though and make each ppd a row
    row_number = 0
    df_out = pd.DataFrame(columns= df_file.columns)
    for item in df_file['Teachers']:
        str_teacher = str(item)
        str_teacher = str_teacher.replace(', ', ',')
        lst_teacher = str_teacher.split(',')
        lst_f_names = lst_teacher[0::2]
        lst_l_names = lst_teacher[1::2]
        lst_teacher2 = [i+', '+j for i,j in zip(lst_f_names, lst_l_names)]
        for ateacher in lst_teacher2:
            new_row = df_file.iloc[row_number]
            new_row.iloc[7] = ateacher
            df_out = df_out.append(new_row)
        row_number += 1
    df_out.reset_index(inplace = True, drop = True)
    

    # Calculate difference in time
    periodto = pd.to_datetime(df_out['Period to'])
    periodfrom = pd.to_datetime(df_out['Period from'])
    difference_time = periodto - periodfrom

    currentrow = 0
    for item in difference_time:
        difference_time.iloc[currentrow] = difference_time.iloc[currentrow].total_seconds()/3600
        currentrow +=1

    df_out['TIME'] = difference_time
    df_out['TIME'].values[df_out['TIME'].values > 7.6] = 7.6



    # Make each teacher a row
    lst_uniq_teachers = df_out['Teachers'].unique()

    lst_of_lsts = []
    for teacher_name in lst_uniq_teachers:
        bools_temp = df_out['Teachers'] == teacher_name
        df_temp_teacher = df_out[bools_temp]
        df_temp_teacher.reset_index(inplace = True, drop = True)
        df_temp_teacher.fillna(value = ' ', inplace = True)

        lst_row = [teacher_name]
        for index, row in df_temp_teacher.iterrows():
            datecorr = row['Date from']
            datecorr = str (datecorr)
            timecorr = row['TIME']
            timecorr = str(timecorr)
            timecorr = timecorr[ 0 : 4 ]
            duration_date = timecorr + ' | (' + datecorr + ')'
            lst_row.append(duration_date)
        lst_of_lsts.append(lst_row)

    df=pd.DataFrame(lst_of_lsts)
    col_mapper = { 0: 'Name', 1: 'PPD1', 2: 'PPD2', 3: 'PPD3', 4: 'PPD4'}
    df.rename( columns = col_mapper, inplace=True)
    df.to_excel(output_filenamepath)
    print('DONE')

window.close()


