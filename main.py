# Import required modules
import tkinter as tk
from tkinter import filedialog
from pathlib import Path
import pandas as pd
import os
from filter import customer_filter
import powerbi_export

# Create a Tkinter root window
root = tk.Tk()
root.withdraw()

# Define global variables
source = str
date = str
type = str
master_df = pd.DataFrame(columns=['Date', 'Source', 'Type', 'Customer', 'Device Class', 'Device Name', 'Patch Category', 'Patch Status', 'Count'])

# Prompt user to select directory containing XLSX files
def get_dir():
    """
    Prompts user to select the directory containing XLSX files.
    Args:
    None

    Returns:
    dir_path (string): The path of the selected directory.
    """
    default_directory = str(Path.home().joinpath('Desktop'))
    dir_path = filedialog.askdirectory(initialdir=default_directory, title='Select Directory Containing XLSX Files')

    print("Using directory path", dir_path)

    return dir_path

# Function to import the Excel file
def import_excel(file):
    """
    Imports an Excel file and sets variables based on the filename.

    Args:
    file_path (string): The path of the Excel file.

    Returns:
    df (Pandas dataframe): The dataframe containing the data from the Excel file.
    source(string): The automatically pulled server source taken from the filename.
    type(string): The automatically pulled endpoint type taken from the filename.
    date(string):The automatically pulled date taken from the filename.
    """

    # Set variables from the filename

    source = file.split(' ')[0] # Split the filename by the first space, get the item before
    type = file.split(' ')[1] # Split the filename by the first space, get the item after
    date1 = file.split('Patch Status')[1] # Split the filename by "Patch Status", take the word item after
    date = date1.split('T')[0] # Split the date by the T, take the item before
    print("Pulling information from the filename, found the source server is" , source , ". The endpoint type is" , type , ". The full date is" , date1 , " and the core date is" , date , ".")
    
    # Load the Excel file into a pandas dataframe
    df = pd.read_excel(file_path, sheet_name=None)
    return df, source, type, date

# Function to clean up each sheet
def clean_data(df):
    """
    Cleans up the data in each sheet of the input dataframe.

    Args:
    df (Pandas dataframe): The dataframe containing the data from the Excel file.

    Returns:
    df (Pandas dataframe): The cleaned dataframe.
    """
    # Remove the first 5 sheets of the dataframe
    print("Cleaing up data")
    for i in range(5):
        df.pop(list(df.keys())[0])

    for sheet_name in df:
        # Clean up data of Sheet 6
        if sheet_name == 'Sheet6':
            customer_name = df[sheet_name].iloc[1, 1].split(":")[1].strip() # Find the customer name
            df[sheet_name].columns = df[sheet_name].iloc[4] # Set column values to row 4
            df[sheet_name] = df[sheet_name].drop(columns=[df[sheet_name].columns[0]]) # Drop first column
            df[sheet_name] = df[sheet_name].assign(Customer=customer_name) # Create new column called Customer Name, assign value of customer_name

        # Clean up the data for other sheets
        else:
            customer_name = df[sheet_name].columns[1] # Set customer_name to the value of the B1 cell
            df[sheet_name].columns = df[sheet_name].iloc[2] # Set column values to row 2
            df[sheet_name] = df[sheet_name].drop(columns=[df[sheet_name].columns[0]]) # Drop first column
            customer_name = customer_name[10:] # Remove the text "Customer: "
            df[sheet_name] = df[sheet_name].assign(Customer=customer_name) #Add a new column with the customer_name value
    return df

# Function to pull each patch status table from the sheets
def extract_table(df):
    """
    Extracts data from each sheet of the input dataframe.

    Args:
    df (Pandas dataframe): The dataframe containing the data from the Excel file.

    Returns:
    sheet_dfs (dictionary): A dictionary containing the extracted data.
    """
    print("Extracting tables from each sheet")
    statuses = ['Aborted', 'Failed', 'In Progress', 'Installed With Errors', 'Not Installed', 'Installed']
    sheet_dfs = {}

    for sheet_name in df:
        for status in statuses:
            startIndex = df[sheet_name][df[sheet_name]['Device Class'] == 'Patch Status: {}'.format(status)].index.tolist()
            endIndex = df[sheet_name][df[sheet_name]['Device Class'] == '{} Patch Count: '.format(status)].index.tolist()

            # Check that there is data for this status
            if startIndex and endIndex:
                sheet_dfs[sheet_name + '_' + status] = df[sheet_name].iloc[startIndex[0]+2:endIndex[0]].copy()
                sheet_dfs[sheet_name + '_' + status]['Patch Status'] = status
                sheet_dfs[sheet_name + '_' + status] = sheet_dfs[sheet_name + '_' + status].fillna(method='ffill')
    return sheet_dfs

# Function to merge dataframes into one dataframe
def merge_df(df, source, type, date):
    """
    Merges data from each dataframe of the input dataframe.

    Args:
    df (Pandas dataframe): The dataframe containing the data from the Excel file.
    source(string): The automatically pulled server source taken from the filename in import_excel()
    type(string): The automatically pulled endpoint type taken from the filename in import_excel()
    date(string):The automatically pulled date taken from the filename in import_excel()

    Returns:
    merged_df (dataframe): A dataframe containing the merged data.
    """
    print("Merging each dataframe from df into one merged_df")
    merged_df = pd.concat(df) #Concat all data into one df
    merged_df = merged_df.iloc[:, :7] #Drop all columns after 7
    merged_df = merged_df.assign(Source=source) # Add a column with the server source
    merged_df = merged_df.assign(Type=type) # Add a column with the endpoint type
    merged_df = merged_df.assign(Date=date) # Add a column with the date
    
    return (merged_df)

# Filter clients from the dataset
def filter_df(master_df, customer_filter):
    print("Filtering dataframe")
    for customer in customer_filter:

        print("Filtering out", customer)
        # filter out rows where Customer is 'Bob'
        master_df = master_df.loc[master_df['Customer'] != customer]
    return(master_df)

# Save file
def output_file(master_df, date):
    """
    Saves data from dataframe into csv file.

    Args:
    df (Pandas dataframe): The dataframe containing the data from the Excel file.

    Returns:
    None
    """
    print("Please select a filepath to save the output file to.")
    # Prompt user to select file path and name
    default_directory = str(Path.home().joinpath('Desktop'))
    output_file_path = filedialog.asksaveasfilename(initialdir=default_directory, initialfile=f"Patch Status - {date}.xlsx",defaultextension='.xlsx', title="Save The XLSX File", filetypes=(('XLSX Files', '*.xlsx'), ('All files', '*.*')))
    print("Using", output_file_path, "to save the file to.")

    # create an Excel writer object
    writer = pd.ExcelWriter(output_file_path, engine='xlsxwriter')

    # Save merged dataframe to XLSX file
    try:
        print("Attempting to save file:", output_file_path)
        master_df.to_excel(writer, sheet_name='Data', index=False)
        print("Output file saved successfully")
    except FileNotFoundError:
        print("The specified file path could not be found. Please try again.")
    except PermissionError:
        print("You do not have permission to save a file at the specified location. Please try again with a different file path.")
    except Exception as e:
        print("An unexpected error occurred while saving the output file: {}".format(e))

    # create a pivot table
    print("Creating pivot table")
    pivot_table = pd.pivot_table(master_df, values='Count', index=['Source', 'Type'], columns='Patch Status', aggfunc='sum', fill_value=0, margins=True)

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

    # Prompt user for directory path of excel files to import
    dir_path = get_dir()

    # Look at each file in the folder, if they are an excel file, process the file.
    for file in os.listdir(dir_path):
        print("Looking at file:" , file)
        if file.endswith('.xlsx'):
            file_path = os.path.join(dir_path, file)
            # Import excel file
            df, source, type, date = import_excel(file)

            # Clean up each sheet
            df = clean_data(df)

            # Extract tables from each sheet
            df = extract_table(df)

            # Merge all dataframes from df into a single dataframe called merged_df
            merged_df = merge_df(df, source, type, date)

            # Append each file's merged_df into a single master_df
            print("Appending each file's merged_df into the master_df")
            master_df = pd.concat([master_df, merged_df])
        else:
            print("File does not end with .xlsx, skipping to next file")
            continue

    #Filter out clients
    master_df = filter_df(master_df, customer_filter)

    # Save data into csv file
    output_file(master_df, date)

    power_bi_choice = input("Would you like to save this data to the Power BI dataset? (Y/N): ")
    if power_bi_choice == "Y":
    
        # Prompt user to select where to save the dataset
        default_directory = str(Path.home().joinpath('Desktop'))
        power_bi_path = filedialog.asksaveasfilename(initialdir=default_directory, initialfile=f"PowerBiReport.xlsx",defaultextension='.xlsx', title="Save The XLSX File", filetypes=(('XLSX Files', '*.xlsx'), ('All files', '*.*')))

        #Use the create_report funciton from powerbi_export.py to create a report if not already present
        powerbi_export.create_report(power_bi_path, master_df)

        #Use the check_report function from powerbi_export.py to check if a report already has the data about to be saved
        powerbi_export.check_report(power_bi_path, master_df)


    elif power_bi_choice == "N":
        print("Not saving data to Power BI, exiting now.")
        exit()
    else:
        print("Invalid response, exiting now.")
        exit()


    input("Script completed, press enter to close the window...")