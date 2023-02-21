N-Able Report Manager Extraction Tool
This program is a data processing tool designed to import reports sent by N-Able's Report Manager, clean up the data, extract specific tables from the data, and save the extracted data to an XLSX file. The program uses the Python programming language and several popular Python libraries for data processing, including pandas and tkinter.

Getting Started
Prerequisites
Python 3.6 or higher
pandas
tkinter
XlsxWriter

Usage
Put all of the XLSX files you want to process in a single directory.
Run the script and select the directory containing the files when prompted.
Enter customer names to be filtered out in the filter.py file.
Select a file path and name for the output file.
The script will process the files and generate a new Excel file with the merged data.
The script will prompt the user if they want to also save data to a single excel file where data from many reports can be used for long-term data analysis

Built With
Python - The programming language used
pandas - Library used for data manipulation
tkinter - Library used for GUI development
XlsxWriter - Library used for Excel file generation


Authors
Zach Radabaugh (zach.radaba@gmail.com)
License
This program is licensed under the MIT License. See the LICENSE file for more information.