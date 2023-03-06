# N-Able Report Manager Extraction Tool
This program is a data processing tool designed to import reports sent by N-Able's Report Manager, clean up the data, and exports the extracted data to MariaDB. The program uses the Python programming language and several popular Python libraries for data processing, including pandas and tkinter.

## Getting Started
### Prerequisites
* Python 3.6 or higher
* pandas
* os
* mariadb
* pathlib
* tkinter
* TQDM

## Usage
* Put all of the CSV files you want to process in a single directory.
* Run the script and select the directory containing the files when prompted.
* The script will process the files and export the data to a DB using the configuration in config.py.
* The script will prompt the user to choose where to save an exported status summary in XLSX format.

## Built With
* Python - The programming language used
* pandas - Library used for data manipulation
* os - A standard Python library for interacting with the operating system.
* mariadb - A module for connecting to a MariaDB database.
* pathlib - A module that provides an object-oriented way to interact with file system paths.
* tkinter - Library used for GUI development
* TQDM - Libray used to display progress

### Authors
Zach Radabaugh (zach.radaba@gmail.com)
### License
This program is licensed under the MIT License. See the LICENSE file for more information.