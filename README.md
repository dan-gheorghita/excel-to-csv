# excel-to-csv.py

**Description of the Python Code**

This Python script is designed to automate the conversion of Excel (.xlsx) files to CSV files. Here's a breakdown of its functionality:

### Importing Libraries

The script starts by importing necessary libraries:

* `os`: for interacting with the operating system and working with file paths
* `csv`: for reading and writing CSV files
* `openpyxl`: for reading and manipulating Excel files (.xlsx)

### Main Loop: Looping through Excel Files

The script then enters a loop that iterates through every file in the current directory. It skips non-Excel files by using a conditional statement (`if not excelFile.endswith('.xlsx')`) to check if the file name ends with '.xlsx'.

### Loading Excel Workbooks and Sheets

For each Excel file found, the script:

1. Loads the workbook using `openpyxl`.
2. Extracts the base name of the Excel file (i.e., the name without the extension) using `os