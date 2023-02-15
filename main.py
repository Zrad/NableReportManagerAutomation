# Import required modules
import tkinter as tk
from tkinter import filedialog
from pathlib import Path
import pandas as pd
import os

# Create a Tkinter root window
root = tk.Tk()
root.withdraw()

# Define global variables
source = str
date = str
type = str
master_df = pd.DataFrame(columns=['Date', 'Source', 'Type', 'Customer', 'Device Class', 'Device Name', 'Patch Category', 'Patch Status', 'Count'])

def get_dir():
    default_directory = str(Path.home().joinpath('Desktop'))
    dir_path = filedialog.askdirectory(initialdir=default_directory, title='Select Directory Containing XLSX Files')

    return dir_path

# Function to import the Excel file
def import_excel(file_path):

    # Set variables from the filename
    test = os.path.basename(file_path) # Get name of the file
    source = test.split(' ')[0] # Split the filename by the first space, get the item before
    type = test.split(' ')[1] # Split the filename by the first space, get the item after
    date = test.split('Patch Status')[1] # Split the filename by "Patch Status", take the word item after
    date = date.split('T')[0] # Split the date by the T, take the item before
    
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
            #print(df[sheet_name].head(15))
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

    merged_df = pd.concat(df) #Concat all data into one df
    merged_df = merged_df.iloc[:, :7] #Drop all columns after 7
    merged_df = merged_df.assign(Source=source) # Add a column with the server source
    merged_df = merged_df.assign(Type=type) # Add a column with the endpoint type
    merged_df = merged_df.assign(Date=date) # Add a column with the date
    
    return (merged_df)

# Save file
def output_file(merged_df):
    """
    Saves data from dataframe into csv file.

    Args:
    df (Pandas dataframe): The dataframe containing the data from the Excel file.

    Returns:
    None
    """
    # Prompt user to select file path and name
    default_directory = str(Path.home().joinpath('Desktop'))
    output_file_path = filedialog.asksaveasfilename(initialdir=default_directory, defaultextension='.csv', title="Save CSV File", filetypes=(('CSV Files', '*.csv'), ('All files', '*.*')))

    # Save merged dataframe to CSV file
    try:
        merged_df.to_csv(output_file_path, index=False)
        print("Output file saved successfully")
    except FileNotFoundError:
        print("The specified file path could not be found. Please try again.")
    except PermissionError:
        print("You do not have permission to save a file at the specified location. Please try again with a different file path.")
    except Exception as e:
        print("An unexpected error occurred while saving the output file: {}".format(e))


    return

if __name__ == '__main__':

    # Prompt user for directory path of excel files to import
    dir_path = get_dir()

    # Look at each file in the folder, if they are an excel file, process the file.
    for file in os.listdir(dir_path):
        if file.endswith('.xlsx'):
            file_path = os.path.join(dir_path, file)
            # Import excel file
            df, source, type, date = import_excel(file_path)

            # Clean up each sheet
            df = clean_data(df)

            # Extract tables from each sheet
            df = extract_table(df)

            # Merge all dataframes from df into a single dataframe called merged_df
            merged_df = merge_df(df, source, type, date)

            master_df = pd.concat([master_df, merged_df])

    # Save data into csv file
    output_file(master_df)