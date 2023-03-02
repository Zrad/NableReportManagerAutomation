# N-Able Report Manager Extraction Tool
This program is a data processing tool designed to import reports sent by N-Able's Report Manager, clean up the data, and exports the extracted data to MariaDB. The program uses the Python programming language and several popular Python libraries for data processing, including pandas and tkinter.

## Getting Started
### Prerequisites
* Python 3.6 or higher
* pandas
* tkinter
* TQDM

## Usage
* Put all of the CSV files you want to process in a single directory.
* Run the script and select the directory containing the files when prompted.
* The script will process the files and export the data to a DB using the configuration in config.py.

## Built With
* Python - The programming language used
* pandas - Library used for data manipulation
* tkinter - Library used for GUI development
* TQDM - Libray used to display progress

## Branch Notes
As I found a better export option in N-Able which saves in csv format and also includes all of the individual patch information which is helpful, I've rewritten this script to just pull in the data from the CSV file and export the data to a DB. It is much simpler this way, however at the moment, it doesn't export out a summary xlsx file. This likely will be added in at a later stage.

### Authors
Zach Radabaugh (zach.radaba@gmail.com)
### License
This program is licensed under the MIT License. See the LICENSE file for more information.