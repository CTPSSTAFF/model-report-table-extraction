# model-report-table-extraction
Python tool and accompanying Access database to streamline extraction of tabular data from CTPS model-generated report file in PRN (columnar) text format.

## Installation
You must install Python version 3 or higher. You can obtain Python be going to https://www.python.org/downloads/.

This tool also requires installation of a Python library called wxPython, which supports user interface elements native to the Windows (or other) operating system. Once you have installed Python itself, you can add the wxPython library by opening a command window (in Windows, click the Windows button in the start bar and type "command prompt" in the search box) and entering the command `python -m pip install wxPython`.

The last step is to get the tool files themselves onto your system. If you are viewing this README on the GitHub.com website, the easiest way to do this is to download a zip file of the project by clicking the green "Code" button and choosing "Download ZIP." Extract the downloaded zip file to the folder of your choice.

## Usage
There are two main steps to use the tool:
1. Run the Python script model_report_table_extraction_gui.py to break out the data-containing lines from certain tables found in the PRN file.
2. Open the Access database and run the "saved imports" in turn to ingest the files output from the previous step and convert them to Access tables

### Run the Python script
In the file browser (Windows Explorer) find the Python script model_report_table_extraction_gui.py that you downloaded during installation and double-click it to launch it. In the dialog window that appears work your way through each of the three buttons:
1. Select model report PRN file. Use the file dialog window to locate and select the model report file from which you wish to extract tables.
2. Specify output directory. Use the file dialog window to locate and select the folder where temporary files for the tables will be created. You may choose to use the same folder where the model report file itself is stored. The temporary files have fixed names of the form <table number>.txt, and will overwrite any files of the same name that already exist in the folder.
3. Break out separate table files. Click "OK" to start the processing.

### Import the temporary table files to Access
Open the (mostly) empty Access database model_report_table_extraction.accdb. Activate the "External Data" ribbon and click "Saved Imports" to open the saved imports window. There is a saved import for each temporary table. For each one, perform the following steps.
1. Click on the file path that begins "C:\Users\dknudsen" and edit it so that the path specifies the actual location of the temporary file from the Python script that is to be loaded.
2. Click the "Run" button.

### Export the Access tables for use in other software
Once your desired tables have been added to Access, you can right-click on any of them, open the "Export" sub-menu, and choose
- "Excel" to export the table to an Excel file
- "Text File" to export the table to a CSV file
- "dBase File" to export the table to a dBase file
