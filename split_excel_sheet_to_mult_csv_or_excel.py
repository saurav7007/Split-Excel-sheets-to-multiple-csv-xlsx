#!/usr/bin/env python3

import pandas as pd

path_of_excel = input("Enter the path of input excel file: ")
output_path = input("Enter the path of output folder: ")

output_format = int(input("Specify Output Format: csv=1 or excel=2: "))

while output_format < 1 or output_format > 2:
     output_format = int(input("Specify Output Format: csv=1 or excel=2: "))

sheet_names = pd.ExcelFile(path_of_excel).sheet_names

for sheet in sheet_names:
    df = pd.read_excel(path_of_excel, sheet_name=sheet)
    if output_format == 1:
        try:
    	    df.to_csv(output_path+'/'+sheet+'.csv', quoting=2, doublequote=True, index=False)
    	    print('Complete writing csv file for sheet {}'.format(sheet))
        except Exception as e:
    	    print('Opps! ',e, ' has occured in sheet ', sheet)
    else: 
       try:
    	   df.to_excel(output_path+'/'+sheet+'.xlsx'.format(sheet),sheet_name=sheet, index=False)
    	   print('Complete writing excel file for sheet {}'.format(sheet))
       except Exception as e:
    	   print('Opps! ',e, ' has occured in sheet ', sheet)

