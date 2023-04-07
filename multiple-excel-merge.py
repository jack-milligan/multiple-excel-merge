"""

Author: Jack, https://github.com/jack-milligan

This Python program combines multiple Excel files into a single file using a GUI.
The program includes several functions that perform different tasks such as checking if a file is an Excel file,
getting a list of file paths for the files to be combined, reading each file into a separate pandas dataframe,
saving the combined dataframe to an Excel file, and combining multiple Excel files into a single file.

Functions:
----------
is_excel_file(filename):
    Check if a file is an Excel file.

is_csv_file(filename):
    Check if a file is a csv file.

get_dataframe_path_list():
    Prompts the user to select the number of files to combine and allows the user to select the files to combine.
    Only files with the extension ".xlsx", ".xls", ".xlsm", or ".csv" are accepted.

read_dataframes(dataframe_path_list):
    Reads each file in the provided list of file paths into a separate pandas dataframe.

save_dataframe_to_excel(combined_df):
    Prompts the user to select an output file path and writes the provided pandas dataframe to an Excel file.

combine_files():
    Combines multiple Excel files into a single file.
    Prompts the user to select the number of files to combine and allows the user to select the files to combine.
    Only files with the extension ".xlsx", ".xls", ".xlsm", or ".csv" are accepted.
    After selecting the files, the function reads each file into a separate pandas dataframe,
    and concatenates all dataframes into a single dataframe.
    The user is then prompted to select an output file path,
    and the combined dataframe is written to an Excel file at that location.

Dependencies:
- os
- pandas
- tkinter
- filedialog
- messagebox
"""
import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog,messagebox


def is_excel_file(filename):
    """
    Check if a file is an Excel file.

    Args:
        filename: A file path to check.

    Returns:
        bool: True if the file is an Excel or csv file (i.e., has the extension ".xlsx" or ".csv"),
        False otherwise.

    Raises:
        None
    """
    if filename.endswith('.xlsx') or filename.endswith('.csv') \
            or filename.endswith('.xls') or filename.endswith('.xlsm'):
        return True
    else:
        return False


def is_csv_file(filename):
    """
    Check if a file is a csv file.

    Args:
        filename: A file path to check.

    Returns:
        bool: True if the file is a csv file (ends with ".csv"),
        False otherwise.

    Raises:
        None
    """
    if filename.endswith('.csv'):
        return True
    else:
        return False


def get_dataframe_path_list():
    """
    Prompts the user to select the number of files to combine, and then opens a file dialog
    for each file, allowing the user to select the files to combine. Only files with the
    extension ".xlsx", ".xls", ".xlsm", or ".csv" are accepted.

    Returns:
        list: A list of file paths for the files to be combined.
    """
    dataframe_path_list = []

    # get the number of files to combine from the user
    number_files = int(num_files_entry.get())

    # prompt the user to select the files to combine
    for i in range(number_files):
        while True:
            # use a file dialog to get the path of each file
            file_path = filedialog.askopenfilename()
            if is_excel_file(file_path) and os.path.isfile(file_path):
                dataframe_path_list.append(file_path)
                break
            else:
                tk.messagebox.showerror("Error", "You did not select a valid Excel file.")
                continue

    return dataframe_path_list


def read_dataframes(dataframe_path_list):
    """
    Reads each file in the provided list of file paths into a separate pandas dataframe.

    Args:
        dataframe_path_list (list): A list of file paths for the files to be combined.

    Returns:
        list: A list of pandas dataframes, each containing the data from one of the input files.
    """
    dataframes = []

    for file in dataframe_path_list:
        if is_csv_file(file):
            df = pd.read_csv(file)
        else:
            df = pd.read_excel(file)
        dataframes.append(df)

    return dataframes


def save_dataframe_to_excel(combined_df):
    """
    Prompts the user to select an output file path, and writes the provided pandas dataframe
    to an Excel file at that location.

    Args:
        combined_df (pandas.DataFrame): The combined pandas dataframe to be written to an Excel file.

    Returns:
        None
    """
    # use a file dialog to select the output file path
    output_path = filedialog.asksaveasfilename(defaultextension=".xlsx")

    # export the dataframe to an Excel document
    combined_df.to_excel(output_path, index=False)

def combine_files():
    """
    Combines multiple Excel files into a single file.

    Prompts the user to select the number of files to combine, and then opens a file dialog
    for each file, allowing the user to select the files to combine. Only files with the
    extension ".xlsx", ".xls", ".xlsm", or ".csv" are accepted. After selecting the files,
    the function reads each file into a separate pandas dataframe, and concatenates all
    dataframes into a single dataframe. The user is then prompted to select an output file
    path, and the combined dataframe is written to an Excel file at that location.

    Returns:
        None
    """

    # get the list of files to combine
    dataframe_path_list = get_dataframe_path_list()

    # read each file into a separate pandas dataframe
    dataframes = read_dataframes(dataframe_path_list)

    # concatenate all the dataframes into a single dataframe
    combined_df = pd.concat(dataframes, ignore_index=True)

    # get the output file path from the user and write the combined dataframe to an Excel file
    save_dataframe_to_excel(combined_df)

    tk.messagebox.showinfo("Success", "The Excel files have been combined.")

# create the GUI window
root = tk.Tk()
root.title("Combine Excel Spreadsheets")

# create a label and entry widget for the number of files
num_files_label = tk.Label(root, text="Number of Excel Spreadsheets to Combine:")
num_files_label.pack()
num_files_entry = tk.Entry(root)
num_files_entry.pack()

# create a button to initiate the file selection process
select_files_button = tk.Button(root, text="Select Files", command=combine_files)
select_files_button.pack()

# start the GUI
root.mainloop()
