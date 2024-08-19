"""
Authored by James Gaskell

08/19/2024

Edited by:

"""
import pandas as pd
import openpyxl, easygui


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

def write_dataframes(filepath):
    return True
