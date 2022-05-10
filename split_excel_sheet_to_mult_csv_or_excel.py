#!/usr/bin/env python3

import os, glob, pandas as pd
from datetime import datetime

path_of_input_folder = input("Enter the path of input folder where excel file/s present: ")

while len(glob.glob(os.path.join(path_of_input_folder, "*.xlsx"))) == 0:
    print('No excel files in the folder.')
    path_of_input_folder = input("Enter the path of input folder where excel file is saved: ")

output_format = int(input("Specify Output Format: csv=1 or excel=2: "))

while output_format < 1 or output_format > 2:
     output_format = int(input("Specify Output Format: csv=1 or excel=2: "))
     
error_log = open("Error"+str(datetime.now())+".log", "a")

for file in glob.glob(os.path.join(path_of_input_folder, "*.xlsx")):
    
    sheet_names = pd.ExcelFile(os.path.join(path_of_input_folder, file)).sheet_names    
    
    isDir = os.path.isdir(r"Split_files")
    if isDir == False:
        os.mkdir("Split_files")
        
    for sheet in sheet_names:
        df = pd.read_excel(os.path.join(path_of_input_folder, file), sheet_name=sheet)
        if output_format == 1:
            try:
                df.to_csv("./Split_files/"+file.split("/")[-1].split('.')[0]+"_"+sheet+'.csv', quoting=2, doublequote=True, index=False)
                print('Complete writing csv file for sheet {} of file {}'.format(sheet,file.split("/")[-1]))
            except Exception as e:
                error_log.write(str(e) + ' has occured in sheet '+ str(sheet) + 'of file ' + str(file) + '\n')
                print('Opps! ',e, ' has occured in sheet ', sheet, 'of file ' + file.split("/")[-1])
        else: 
            try:
                df.to_excel("./Split_files/"+file.split("/")[-1].split('.')[0]+"_"+sheet+'.xlsx'.format(sheet),sheet_name=sheet, index=False)
                print('Complete writing excel file for sheet {} of file {}'.format(sheet,file.split("/")[-1]))
            except Exception as e:
                error_log.write(str(e) + ' has occured in sheet ' + str(sheet) + 'of file ' + str(file) + '\n')
                print('Opps! ',e, ' has occured in sheet ', sheet, 'of file ' + file.split("/")[-1])

error_log.close()

