import tkinter as tk
from tkinter import filedialog
from pathlib import Path
import pandas as pd
from openpyxl import load_workbook
import os

# Create a Tkinter root window
root = tk.Tk()
root.withdraw()

# Prompt user to select XLSX file
def get_file():
    """
    Prompts user to select an XLSX file.
    Args:
    None

    Returns:
    file_path (string): The path of the selected XLSX file.
    """
    default_directory = str(Path.home().joinpath('Desktop'))
    file_path = filedialog.askopenfilename(initialdir=default_directory, title='Select XLSX File')

    return file_path

# Read data from the imported xlsx file
def import_data(file_path):
    """
    Reads the data from the first sheet of an XLSX file into a pandas dataframe.

    Args:
    file_path (string): The file path of the XLSX file to be imported.

    Returns:
    import_df (pandas dataframe): A pandas dataframe containing the data from the first sheet of the XLSX file.
    """
    import_df = pd.read_excel(file_path, sheet_name= 0)
    return import_df

# Check to see if a xlsx file has already been created. If not, output dataframe to new file and exit program
def create_report(power_bi_path, import_df):
    """
    Looks for an existing Power BI dataset using the power_bi_path. If one does not exist, writes the contents of the dataframe to a new file and exits program.

    Args:
    power_bi_path (string): The file path of the xlsx file containing the Power BI dataset.
    import_df (Pandas dataframe): The dataframe containing the data from the Excel file.

    Returns:
    None
    """
    print("Checking to see if Power BI dataset exists")
    if not os.path.isfile(power_bi_path):
        print("Power BI dataset does not exist, creating report now and exiting")
        with pd.ExcelWriter(power_bi_path) as writer:
            import_df.to_excel(writer, index=False, sheet_name='Sheet1')
        exit()
    print("Current Power BI dataset found")
    return

# Check to see if data to import already exists in Power BI dataset
def check_report(power_bi_path, import_df):
    """
    Compares the date from import_df to power_bi_path's data to determine if data is already present

    Args:
    power_bi_path (string): The file path of the xlsx file containing the Power BI dataset.
    import_df (Pandas dataframe): The dataframe containing the data from the Excel file.

    Returns:
    None
    """
    # Pull data from existing power_bi_path xlsx into a dataframe for evaluation
    powerbi_df = pd.read_excel(power_bi_path) 
    # Extract the date from the imported xlsx file
    import_date = import_df['Date'].iloc[0]
    # If the import_date is found somewhere in the Date column, exit program
    if import_date in powerbi_df['Date'].values:
        print("Date exists")
        exit()
    else:
        print("Date not found in dataframe. Appending dataframe to Power BI datset")
        # Append import_df to power_bi_path
        updated_df = pd.concat([powerbi_df, import_df], ignore_index=True, sort=False)
        with pd.ExcelWriter(power_bi_path, mode='w') as writer:
            updated_df.to_excel(writer, index=False, sheet_name='Sheet1')
        print("Data appended to Power BI dataset.")
    return

# Main function of program
if __name__ == '__main__':
    # Get filepath of patch status to add to power bi report
    file_path = get_file()

    # Import data from the selected xlsx file
    import_df = import_data(file_path)

    # Prompt user to select where to save the dataset
    #power_bi_path = 'C:/Users/zach.radabaugh/OneDrive - Nexus Technologies Inc/Documents/Patch logs/PowerBiReport.xlsx' # Used for internal testing
    default_directory = str(Path.home().joinpath('Desktop'))
    power_bi_path = filedialog.asksaveasfilename(initialdir=default_directory, initialfile=f"PowerBiReport.xlsx",defaultextension='.xlsx', title="Save The XLSX File", filetypes=(('XLSX Files', '*.xlsx'), ('All files', '*.*')))

    # Check if power bi report already exists or not. If no, create the file and exit.
    create_report(power_bi_path, import_df)
    
    # If power bi report exists, check the date to verify if the data already exists in the report
    check_report(power_bi_path, import_df)