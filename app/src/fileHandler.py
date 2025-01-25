"""
Authored by James Gaskell

01/24/2025

Edited by:

"""
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
from openpyxl.formatting.formatting import ConditionalFormattingList

"""Openpyxl does not have permission to edit an open file
    This function checks if the file is open in an editor so an exception can be raised
Args:
    filepath: the path to the excel file
Returns:
    Boolean: True if the file is open in an editor
"""
def file_open_check(filepath):
    wb = openpyxl.load_workbook(filepath)
    try:
        wb.save(filepath)
        wb.close()
        return False
    except:
        return True
    
"""Currently used to set date format to the correct ISO
    Can be expanded to include formatting options for other fields
Args:
    ws: the worksheet object
    column_name: the name of the column
    column_index: the index of the column
"""
    
def set_field_format(ws, column_name, column_index):
    if column_name == "date_created":
        for cell in ws[get_column_letter(column_index+1)]:
             cell.alignment = Alignment(horizontal='right')
             cell.number_format = "YYYY-MM-DD"


# Can we keep this method mostly the same? 
# We don't currently store sheetnames in the object structure - does it make sense to just pass in the whole file object?
# If we do it this way, df_to_excel could call out color method
def df_to_excel(dfs, sheetnames, filepath):

    wb = openpyxl.load_workbook(filepath)  # load as openpyxl workbook; useful to keep the original layout
                                 # which is discarded in the following dataframe
    
    for sheet in sheetnames:
         
        ws = wb[sheet]
        for index, column_name in enumerate(list(dfs[sheet].columns.values)): #Writes the column headings to the file
             ws.cell(1, index+1).value = column_name

             # This is where we call the field formatter
             # Any expansion to formatter requires an additional call here
             if column_name == ("date_created"):
                  set_field_format(ws, "date_created", index)


        for r_idx, row in dfs[sheet].iterrows():
             for c_idx, value in enumerate(row, 1):
                  ws.cell(row=r_idx+2, column=c_idx).value = value

    wb.save(filepath)
    wb.close()