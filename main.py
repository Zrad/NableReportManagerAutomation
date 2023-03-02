# Import require modules
import tkinter as tk
from tkinter import filedialog
from pathlib import Path
import pandas as pd
import os
import mariadb
from tqdm import tqdm
from config import db_config

# Create a Tkinter root window
root = tk.Tk()
root.withdraw()

# Define global variables
source = str
date = str
columns = ['Customer_Name', 'Installation_Status', 'Device_Class', 'Device_Name', 'Patch_Name', 'Patch_Products', 'Patch_Category', 'Publish_Date', 'Approval_Status', 'Approval_Date', 'Status_Change_Date']
master_df = pd.DataFrame(columns=columns)

# Prompt user to select directory containing CSV files
def get_dir():

    default_directory = str(Path.home().joinpath('Desktop'))
    dir_path = filedialog.askdirectory(initialdir=default_directory, title='Select Directory Containing CSV Files')

    print(f"Using directory path, {dir_path}")

    return dir_path

# Import CSV File Function
def import_csv(file):
    # Set variables from the filename

    source = file.split(' ')[0] # Split the filename by the first space, get the item before
    date1 = file.split('Patch Status')[1] # Split the filename by "Patch Status", take the word item after
    date = date1.split('T')[0] # Split the date by the T, take the item before
    print(f"Pulling information from the filename: Server is {source}, date is {date}")

    import_df = pd.read_csv(file_path, header=None, names=columns, usecols=range(11))

    return import_df, source, date

# Clean Data Function
def cleanup_df(import_df, source, date):
    # Look for valid data by finding where there is data in column 7
    row_index = import_df[import_df.iloc[:, 7].notnull()].index[0]
    # Drop all rows above row_idex (including row_index)
    clean_df = import_df.drop(import_df.index[:row_index + 1])
    # Remove the Customer: text from the Customer_Name Column
    clean_df['Customer_Name'] = clean_df['Customer_Name'].str.replace('Customer: ', '')
    # Remove the Patch Status: text from the Customer_Name Column
    clean_df['Installation_Status'] = clean_df['Installation_Status'].str.replace('Patch Status: ', '')
    # Reset the row index
    clean_df = clean_df.reset_index(drop=True)
    # Add a column with the server source
    clean_df = clean_df.assign(Source=source) 
    # Add a column with the date of this report
    clean_df = clean_df.assign(Report_Date=date) 

    # Format the Publish Date and Approval Date columns using YYYY-MM-DD instead of Feb 16, 2023
    clean_df['Publish_Date'] = pd.to_datetime(clean_df['Publish_Date'], errors='coerce')
    clean_df['Approval_Date'] = pd.to_datetime(clean_df['Approval_Date'], errors='coerce')
    clean_df['Status_Change_Date'] = pd.to_datetime(clean_df['Status_Change_Date'], errors='coerce')

    # Replace NaTType values with default value
    clean_df['Approval_Date'].fillna('1970-01-01', inplace=True)
    clean_df['Status_Change_Date'].fillna('1970-01-01', inplace=True)
    clean_df['Patch_Products'].fillna('', inplace=True)

    # Remove any commas in the Patch_Category column
    clean_df['Patch_Products'] = clean_df['Patch_Products'].replace(',', ' ', regex=True)

    return clean_df

# Save data to MariaDB
def output_db(master_df, date):
    # Create a connection object
    print("Connecting to MariaDB")
    conn = mariadb.connect(**db_config)

    # Create a cursor object
    cursor = conn.cursor()

    # Execute an SQL query to create a table
    print("Creating table if one doesn't exist")
    cursor.execute("CREATE TABLE IF NOT EXISTS nable_table (id INT AUTO_INCREMENT PRIMARY KEY, customer_name VARCHAR(255), installation_status VARCHAR(255), device_class VARCHAR(255), device_name VARCHAR(255), patch_name VARCHAR(255), patch_products VARCHAR(255), patch_category VARCHAR(1255), publish_date DATE, approval_status VARCHAR(255), approval_date DATE, status_change_date DATE, source VARCHAR(255), report_date DATE)")

    # Prepare the insert query
    insert_query = "INSERT INTO nable_table (customer_name, installation_status, device_class, device_name, patch_name, patch_products, patch_category, publish_date, approval_status, approval_date, status_change_date, source, report_date) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"


    # Execute the query to search for the date in the report_date column
    print("Searching DB to see if data already exists based on report date")
    cursor.execute("SELECT * FROM nable_table WHERE report_date = %s", (date,))

    # Fetch the results of the query
    result = cursor.fetchall()

    # Check if the date was found in the report_date column
    if result:
        # Code to execute if the date was found
        print("The date was found in the report_date column!")
    else:
        # Code to execute if the date was not found
        print("The date was not found in the report_date column.")

        # Loop through the data and insert each row into the table in batches
        print("Saving data to MariaDB")
        pd.options.mode.chained_assignment = None  # default='warn'
        batch_size = 1000
        for i in tqdm(range(0, len(master_df), batch_size)):
            batch_df = master_df.iloc[i:i+batch_size]

            # Convert the date columns to datetime objects
            batch_df['Publish_Date'] = pd.to_datetime(batch_df['Publish_Date'], format='%b %d, %Y', errors='coerce').dt.date
            batch_df['Approval_Date'] = pd.to_datetime(batch_df['Approval_Date'], format='%b %d, %Y', errors='coerce').dt.date
            batch_df['Status_Change_Date'] = pd.to_datetime(batch_df['Status_Change_Date'], format='%b %d, %Y', errors='coerce').dt.date

            # Convert the dataframe to a list of tuples for batch insert
            batch_values = batch_df.to_records(index=False).tolist()

            # Execute the batch insert
            cursor.executemany(insert_query, batch_values)

        # Commit the changes to the database
        print("Committing data")
        conn.commit()

    # Close the cursor and connection objects
    print("Exiting from MariaDB")
    cursor.close()
    conn.close()

    return

# Main funtion of program
if __name__ == '__main__':

    # Prompt user for directory path of csv files to import
    dir_path = get_dir()

    # Look at each file in the folder, if they are a csv file, process the file.
    for file in os.listdir(dir_path):
        print(f"Looking at file: {file}")
        if file.endswith('.csv'):
            file_path = os.path.join(dir_path, file)
            # Import csv file
            import_df, source, date = import_csv(file)

            # Clean csv file
            clean_df = cleanup_df(import_df, source, date)

            # Append data to a single dataframe
            master_df = pd.concat([master_df, clean_df])

        else:
            print("File does not end with .csv, skipping to next file")
            continue

    # Export data to MariaDB
    output_db(master_df, date)