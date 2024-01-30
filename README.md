# CFD_ITR_Tool
Tool that scrapes and searches legacy in service training records from the old Share Point form

Training Record Lookup Tool
This is a Python script that allows you to search and extract information from a specific Excel file containing CFD in service training records. It provides a graphical user interface (GUI) created using the tkinter library for user interaction.

How to Use
Select Excel File: Upon running the script, you will be prompted to select an Excel file containing the training records. The script assumes that the relevant data is stored in a sheet named 'ITRs' within the selected Excel file.

Enter File Number: After selecting the Excel file, a dialog box will appear asking you to enter a "File Number." This file number is used to filter the training records based on the selected file number.

Output Directory: You will be asked to choose an output directory where the script will save the generated Excel file.

Search and Write: Click the "Search and Write to Excel" button to execute the search and extraction process. The script will perform the following steps:

Filter the training records based on the entered file number.
Sum the durations of different training classes.
Create an Excel file with two sheets:
"Class Summary": Lists training classes and their total hours in descending order.
"ITRs": Contains the filtered training records, excluding 'Item Type' and 'Path' columns.
Result: The script will display a message indicating the location where the generated Excel file is saved.

Requirements
Python 3.x
Pandas library
tkinter library (usually included with Python)
xlsxwriter library
