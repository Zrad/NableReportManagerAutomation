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

# Prompt user to select directory containing CSV files
def get_dir():

    default_directory = str(Path.home().joinpath('Desktop'))
    dir_path = filedialog.askdirectory(initialdir=default_directory, title='Select Directory Containing CSV Files')

    print(f"Using directory path, {dir_path}")

    return dir_path

# Import Patch Status CSV File Function
def import_status(file):
    # Set variables from the filename
    
    source = file.split(' ')[0] # Split the filename by the first space, get the item before
    type = file.split(' ')[1] # Split the filename by the first space, get the item after
    date1 = file.split('Patch Status')[1] # Split the filename by "Patch Status", take the word item after
    date = date1.split('T')[0] # Split the date by the T, take the item before
    ##print(f"Pulling information from the filename: Server is {source}, type is {type}, date is {date}")

    import_df = pd.read_csv(file_path, header=None, names=import_status_columns, usecols=range(8))

    return import_df, source, type, date

# Import Patch Details CSV File Function
def import_details(file):
    # Set variables from the filename

    source = file.split(' ')[0] # Split the filename by the first space, get the item before
    date1 = file.split('Patch Details')[1] # Split the filename by "Patch Details", take the word item after
    date = date1.split('T')[0] # Split the date by the T, take the item before
    ##print(f"Pulling information from the filename: Server is {source}, date is {date}")

    import_df = pd.read_csv(file_path, header=None, names=details_columns, usecols=range(11))

    return import_df, source, date, details_columns

# Clean Data Function
def cleanup_df(import_df, source, type, date):

    if 'Patch Status' in file:
        
        # Look for cell in column 1 that has txtPS_Details_CustomerName
        table_start_search_value = "txtPS_Details_CustomerName"
        row_index_start = import_df.loc[import_df['Customer_Name'] == table_start_search_value].index[0] + 1

        # Create a new dataframe using the row indexes and resetting the row index
        clean_df = import_df.iloc[row_index_start:].reset_index(drop=True) 

        # Remove the Customer: text from the Customer_Name Column and Patch Status: text from the Customer_Name Column
        clean_df['Customer_Name'] = clean_df['Customer_Name'].str.replace('Customer: ', '')
        clean_df['Installation_Status'] = clean_df['Installation_Status'].str.replace('Patch Status: ', '')

        # Add columns for source, type, and date
        clean_df = clean_df.assign(Source=source)
        clean_df = clean_df.assign(Type=type)
        clean_df = clean_df.assign(date=date)

        # Remove the columns Details and Managed_Link
        clean_df = clean_df.drop(['Details', 'Managed_Link'], axis=1)

        return clean_df
    elif 'Patch Details' in file:

        # Look for valid data by finding where there is data in column 7
        row_index = import_df[import_df.iloc[:, 7].notnull()].index[0]
        # Drop all rows above row_idex (including row_index)
        clean_df = import_df.drop(import_df.index[:row_index + 1])

        # Remove the Customer: text from the Customer_Name Column and Patch Status: text from the Customer_Name Column
        clean_df['Customer_Name'] = clean_df['Customer_Name'].str.replace('Customer: ', '')
        clean_df['Installation_Status'] = clean_df['Installation_Status'].str.replace('Patch Status: ', '')

        # Add a column with the server source and date
        clean_df = clean_df.assign(Source=source) 
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

        # Reset the row index
        clean_df = clean_df.reset_index(drop=True)

        return clean_df

# Save data to MariaDB
def output_db(status_df, details_df, date, status_query, details_query):
    # Create a connection object
    print("Connecting to MariaDB")
    conn = mariadb.connect(**db_config)

    # Create a cursor object
    cursor = conn.cursor()

    # Execute an SQL query to create a table
    print("Creating table if one doesn't exist")
    cursor.execute("CREATE TABLE IF NOT EXISTS status_table (id INT AUTO_INCREMENT PRIMARY KEY, customer_name VARCHAR(255), installation_status VARCHAR(255), device_class VARCHAR(255), device_name VARCHAR(255), patch_category VARCHAR(255), count INT, source VARCHAR(255), type VARCHAR(255), report_date DATE)")
    cursor.execute("CREATE TABLE IF NOT EXISTS details_table (id INT AUTO_INCREMENT PRIMARY KEY, customer_name VARCHAR(255), installation_status VARCHAR(255), device_class VARCHAR(255), device_name VARCHAR(255), patch_name VARCHAR(255), patch_products VARCHAR(255), patch_category VARCHAR(1255), publish_date DATE, approval_status VARCHAR(255), approval_date DATE, status_change_date DATE, source VARCHAR(255), report_date DATE)")

    # Execute the query to search for the date in the report_date column
    print(f"Searching status_table to see if any rows contain {date}")
    cursor.execute("SELECT * FROM status_table WHERE report_date = %s", (date,))

    # Fetch the results of the query
    status_result = cursor.fetchall()

    # Execute the query to search for the date in the report_date column
    print(f"Searching details_table to see if any rows contain {date}")
    cursor.execute("SELECT * FROM details_table WHERE report_date = %s", (date,))

    # Fetch the results of the query
    details_result = cursor.fetchall()
    
    # Setting amount of rows to send to DB per batch
    batch_size = 1000
    pd.options.mode.chained_assignment = None  # default='warn'


    # Check if the date was found in the report_date column
    if status_result:
        # Code to execute if the date was found
        print("The date was found in the report_date column! The data is likely already on the database.")
    else:
        # Code to execute if the date was not found
        print("The date was not found in the report_date column.")

        # Loop through the data and insert each row into the table in batches
        print("Saving data to MariaDB")
        
        for i in tqdm(range(0, len(status_df), batch_size)):
            batch_df = status_df.iloc[i:i+batch_size]

            # Convert the dataframe to a list of tuples for batch insert
            batch_values = batch_df.to_records(index=False).tolist()

            # Execute the batch insert
            cursor.executemany(status_query, batch_values)

        # Commit the changes to the database
        print("Committing data")
        conn.commit()

    if details_result:
        # Code to execute if the date was found
        print("The date was found in the report_date column! The data is likely already on the database.")
    else:
        # Code to execute if the date was not found
        print("The date was not found in the report_date column.")

        # Loop through the data and insert each row into the table in batches
        print("Saving data to MariaDB")

        for i in tqdm(range(0, len(details_df), batch_size)):
            batch_df = details_df.iloc[i:i+batch_size]

            # Convert the date columns to datetime objects
            batch_df['Publish_Date'] = pd.to_datetime(batch_df['Publish_Date'], format='%b %d, %Y', errors='coerce').dt.date
            batch_df['Approval_Date'] = pd.to_datetime(batch_df['Approval_Date'], format='%b %d, %Y', errors='coerce').dt.date
            batch_df['Status_Change_Date'] = pd.to_datetime(batch_df['Status_Change_Date'], format='%b %d, %Y', errors='coerce').dt.date

            # Convert the dataframe to a list of tuples for batch insert
            batch_values = batch_df.to_records(index=False).tolist()

            # Execute the batch insert
            cursor.executemany(details_query, batch_values)

        # Commit the changes to the database
        print("Committing data")
        conn.commit()

    # Close the cursor and connection objects
    print("Exiting from MariaDB")
    cursor.close()
    conn.close()

    return

# Save a summary report to an XLSX file
def output_summary(status_df, date):
    print("Please select a filepath to save the output file to.")

    # Prompt user to select file path and name
    default_directory = str(Path.home().joinpath('Desktop'))
    output_file_path = filedialog.asksaveasfilename(initialdir=default_directory, initialfile=f"Patch Status - {date}.xlsx",defaultextension='.xlsx', title="Save The XLSX File", filetypes=(('XLSX Files', '*.xlsx'), ('All files', '*.*')))
    print("Using", output_file_path, "to save the file to.")

    # create an Excel writer object
    writer = pd.ExcelWriter(output_file_path, engine='xlsxwriter')

    # Convert Count column to integers for better data compatibility
    status_df['Count'] = status_df['Count'].astype('int64')

    # Save merged dataframe to XLSX file
    try:
        print("Attempting to save file:", output_file_path)
        status_df.to_excel(writer, sheet_name='Data', index=False)
        print("Output file saved successfully")
    except FileNotFoundError:
        print("The specified file path could not be found. Please try again.")
    except PermissionError:
        print("You do not have permission to save a file at the specified location. Please try again with a different file path.")
    except Exception as e:
        print("An unexpected error occurred while saving the output file: {}".format(e))

    # create a pivot table
    print("Creating pivot table")
    pivot_table = pd.pivot_table(status_df, values='Count', index=['Source', 'Type'], columns='Installation_Status', aggfunc='sum', fill_value=0, margins=True)

    # Pulling installed and all patch counts for each row
    all_values = pivot_table['All'].drop(index='All')
    installed_values = pivot_table['Installed'].drop(index='All')

    # Calculate the percentage of installed values
    percent_installed = installed_values / all_values * 100
    # Format as percentage with no decimal places
    percent_installed_str = percent_installed.map('{:.1f}%'.format)
    pivot_table = pivot_table.assign(Percent_Installed=percent_installed_str)

    # write the pivot table to a second sheet
    print("Writing pivot table to second sheet.")
    pivot_table.to_excel(writer, sheet_name='Summary')

    writer.close()
    print("File save complete.")

    return

# Main funtion of program
if __name__ == '__main__':

    # Prompt user for directory path of csv files to import
    dir_path = get_dir()

    # Prepare the insert query
    status_query = "INSERT INTO status_table (customer_name, installation_status, device_class, device_name,  patch_category, count, source, type, report_date) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)"
    details_query = "INSERT INTO details_table (customer_name, installation_status, device_class, device_name, patch_name, patch_products, patch_category, publish_date, approval_status, approval_date, status_change_date, source, report_date) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"

    # Set column names
    import_status_columns = ['Customer_Name', 'Details', 'Managed_Link', 'Installation_Status', 'Device_Class', 'Device_Name', 'Patch_Category', 'Count']
    status_columns = ['Customer_Name', 'Installation_Status', 'Device_Class', 'Device_Name', 'Patch_Category', 'Count']
    details_columns = ['Customer_Name', 'Installation_Status', 'Device_Class', 'Device_Name', 'Patch_Name', 'Patch_Products', 'Patch_Category', 'Publish_Date', 'Approval_Status', 'Approval_Date', 'Status_Change_Date']

    # Create empty dataframes with defined columns
    status_df = pd.DataFrame(columns=status_columns)
    details_df = pd.DataFrame(columns=details_columns)    

    # Look at each file in the folder, if they are a csv file, process the file.
    for file in os.listdir(dir_path):
        print(f"Looking at file: {file}")
        if file.endswith('.csv'):
            file_path = os.path.join(dir_path, file)

            if 'Patch Status' in file: # Search the filename for Patch Status
                print(f"Importing: {file}")
                
                # Import csv file
                import_df, source, type, date = import_status(file)

                # Clean csv file
                clean_df = cleanup_df(import_df, source, type, date)

                # Append cleaned dataframe to status dataframe
                status_df = pd.concat([status_df, clean_df])

            elif 'Patch Details' in file: # Search the filename for Patch Details
                print(f"Importing: {file}")
                # Import csv file
                import_df, source, date, details_columns = import_details(file)

                # Clean csv file
                clean_df = cleanup_df(import_df, source, type, date)

                # Append cleaned dataframe to details dataframe
                details_df = pd.concat([details_df, clean_df])

            else: # Return error message if the file doesn't contain the correct names
                print("File name doesn't contain Status or Details")
                continue

        else:
            print("File does not end with .csv, skipping to next file")
            continue

    output_summary(status_df, date)

    # Export data to MariaDB
    output_db(status_df, details_df, date, status_query, details_query)