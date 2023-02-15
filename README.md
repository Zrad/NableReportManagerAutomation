N-Able Report Manager Processing Tool
This program is a data processing tool designed to import reports sent by N-Able's Report Manager, clean up the data, extract specific tables from the data, and save the extracted data to a CSV file. The program uses the Python programming language and several popular Python libraries for data processing, including pandas and tkinter.

How to Use the Program
Open a command prompt or terminal window and navigate to the directory where the data_processing.py file is located.

Run the command python data_processing.py to start the program.

When prompted, select the Excel file you want to process. The program will load the data from the file and display a message indicating that the data has been imported successfully.

The program will then clean up the data by removing unwanted sheets and columns, renaming columns, and filling in missing values. The cleaned data will be displayed in the terminal window.

The program will then extract specific tables from the cleaned data based on certain criteria. The extracted tables will be displayed in the terminal window.

Finally, the program will prompt you to select a file name and location for the output CSV file. The extracted tables will be saved to the selected file.

Requirements
Python 3.x
pandas
tkinter
Known Issues
If the Excel file contains any sheets that do not match the expected format, the program may fail or produce unexpected results.
The program currently only supports Excel files with the .xlsx file extension.
Authors
Zach Radabaugh (zach.radaba@gmail.com)
License
This program is licensed under the MIT License. See the LICENSE file for more information.