# Split Excel sheets to multiple csv/xlsx

**Description:** The python script "split_excel_sheet_to_mult_csv_or_excel.py" reads each sheets from the input excel file and creates seperate files for them (csv or excel based on user input).

**Pre-requsites:**
* python 3.6+
* Modules: glob, configparser, pandas, openpyxl

**Steps:**
1. Download the zip 'Split-Excel-sheets-to-multiple-csv-xlsx-main.zip' and extract the content.
2. Go to the folder 'Split-Excel-sheets-to-multiple-csv-xlsx-main' and open the configfile.ini file.
3. Update the value for "path_of_input_folder", and "file_format" and save it.
4. Open the terminal in the same folder and run the following command: ```python3 split_excel_sheet_to_mult_csv_or_excel.py```

**Output**
* ./Split_files/<FilesName_sheetname.csv/xlsx>
* ./Error.TimeStamp.csv
