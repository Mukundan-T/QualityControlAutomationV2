"""
Authored by James Gaskell

08/19/2024

Edited by:

"""

import pandas as pd
import openpyxl, easygui
from openpyxl.utils.dataframe import dataframe_to_rows



"""Opens an EasyGUI window to allow the user to select the file they want to parse
Returns:
    String: filepath of the selected file
"""
def get_file():
        path = easygui.fileopenbox()
        return path


"""Uses pandas to read excel document selected by the user
Args:
    String: Filepath decide dby the user in the selection process
Returns:
    Dict: Dataframes for each sheet in the file
    List[String]:Sheetnames in the file
"""
def get_dataFrames(filepath):
    xl_file = pd.ExcelFile(filepath)
    dfs = pd.read_excel(xl_file, sheet_name=None)
    sheets = xl_file.sheet_names

    xl_file.close()
    return dfs,sheets

def write_dataframes(dfs, sheetnames, filepath):
    xl_writer = pd.ExcelWriter("filepath", engine='openpyxl')
    for sheet in sheetnames:
         dfs[sheet].to_excel(xl_writer, sheet_name=sheet, index=False)
    xl_writer.save()
    return True

def df_to_excel(df, sheetnames, filepath, header=True, index=True, startrow=0, startcol=0):

    wb = openpyxl.load_workbook(filepath)  # load as openpyxl workbook; useful to keep the original layout
                                 # which is discarded in the following dataframe
    
    for sheet in sheetnames:
         
        ws = wb[sheet]

        rows = dataframe_to_rows(df, header=header, index=index)

        for r_idx, row in enumerate(rows, startrow + 1):
            for c_idx, value in enumerate(row, startcol + 1):
                ws.cell(row=r_idx, column=c_idx).value = value

    wb.save(filepath)
