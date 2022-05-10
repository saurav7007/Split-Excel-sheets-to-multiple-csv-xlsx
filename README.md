# Split Excel sheets to multiple csv/xlsx

**Description:** The python script "split_excel_sheet_to_mult_csv_or_excel.py" reads each sheets from the input excel file and creates seperate files for them (csv or excel based on user input).

**Pre-requsites:**
* python 3.6+
* pandas module
* Openpyxl module

**Steps:**
1. Download the python script "split_excel_sheet_to_mult_csv_or_excel.py".
2. Go to the folder where the file is downloaded and open the terminal.
3. Run the following command: ```python3 split_excel_sheet_to_mult_csv_or_excel.py```
4. Enter the path of input folder where excel file/s present: <**path for folder where excel file/s present**>
5. Specify Output Format: csv=1 or excel=2: <**1** to get output in csv format / **2** to get output as xlsx format>

**Output**
* ./Split_files/<FilesName_sheetname.csv/xlsx>
* ./Error.<TimeStamp>.log
