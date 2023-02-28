"""
Author: Jack, https://github.com/jack-milligan
This script combines multiple Excel spreadsheets of the same structure (same columns, data types, etc.) into a
single Excel file. The user is prompted to enter the number of spreadsheets to combine and the file paths (local) for each
spreadsheet. The script checks that each input is a valid path of an Excel file with the extension ".xlsx". Once all the
spreadsheets have been read into separate pandas dataframes, the dataframes are concatenated into a single dataframe.
The resulting dataframe is then exported to an Excel file named "combined_excels.xlsx" in the same directory as the
script file.

Functions:
- is_excel_file(filename): checks if a file is an Excel file with the extension ".xlsx".

Usage:
- Run the script in a Python environment (e.g., Anaconda, Jupyter Notebook, etc.).
- Follow the prompts to enter the number of Excel spreadsheets to combine and the file paths for each spreadsheet.
- The resulting dataframe will be saved to a file named "combined_excels.xlsx" in the same directory as the script.
"""
import os
import pandas as pd


def is_excel_file(filename):
    """Check if a file is an Excel file.

    Args:
        filename: A file path to check.

    Returns:
        bool: True if the file is an Excel file (i.e., has the extension ".xlsx"),
        False otherwise.

    Raises:
        None
    """
    if filename.endswith('.xlsx'):
        return True
    else:
        return False


while True:
    # ensuring valid input
    try:
        numberFiles = int(input('How many excel spreadsheets of same structure (same columns and data types) would '
                                'you like to combine?:'))
        break
    except ValueError:
        print('Invalid input. Please enter a number.')

dataframePathList = list()

for i in range(numberFiles):
    # ensuring valid input
    while True:
        userInput = input(f'Enter path of spreadsheet {i + 1}:')
        if is_excel_file(userInput) and os.path.isfile(userInput):
            dataframePathList.append(userInput)
            break
        else:
            print('You did not enter the path of an excel file in .xlsx format, please try again.')

# list to house each dataframe
dataframes = list()

# loop through each file and read it into a dataframe
for file in dataframePathList:
    df = pd.read_excel(file)
    dataframes.append(df)

# concatenate all the dataframes into a single dataframe
combined_df = pd.concat(dataframes, ignore_index=True)

# export the dataframe to an Excel document
combined_df.to_excel('combined_excels.xlsx', index=False)
