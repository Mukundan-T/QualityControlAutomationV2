"""
Authored by James Gaskell

01/24/2025

Edited by:

"""
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment


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
    

#Computationally expensive but should work for now
def reset_colors(ExcelFile, wb, colors_to_remove):
    fill_reset = openpyxl.styles.PatternFill(fill_type=None)
    for sheet in ExcelFile.sheetList:
        ws = wb[sheet.sheetName]
        for row in ws.iter_rows():
            for cell in row:
                if cell.fill in colors_to_remove.values():
                    cell.fill = fill_reset


def highlight_errors(ExcelFile):

    xl_file = pd.ExcelFile(ExcelFile.filePath)       
    wb = openpyxl.load_workbook(xl_file)

    reset_colors(ExcelFile, wb, ExcelFile.errorColors)

    for sheet in ExcelFile.sheetList:
        dt = pd.read_excel(xl_file, sheet.sheetName)
        ws = wb[sheet.sheetName]

        for file in sheet.getSheetErrorDict():
            error_color = None
            try:
                if file.errors['DupFilename']:
                    error_color = ExcelFile.errorColors['Duplicate']
                elif file.errors['Filename']:
                    error_color = ExcelFile.errorColors['Filename']
                elif file.errors['Date']:
                    error_color = ExcelFile.errorColors['DateFormat']
            except:
                error_color = ExcelFile.errorColors['Filename'] # If try fails to an exception we know fileName is likely Nan
                
            if error_color != None:
                 fill = openpyxl.styles.PatternFill(start_color=error_color, end_color=error_color, fill_type="solid")
                 for index, row in dt.iterrows():
                    try:
                        if file.fileName == dt['Filename'][index]:
                            for y in range(1, ws.max_column+1):
                                ws.cell(row=index+2, column=y).fill = fill
                    except:
                        pass
    

    try: 
        wb.save(ExcelFile.filepath)
        wb.close()
        xl_file.close()
    except:
        wb.close()
        xl_file.close()
        return False

    return True #If false returned then the file is open in an editor
    

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

# Need a method to check if a file has been auto failed in the past by the program
# Delete pass/fail, comment and intials from dataframes
def clear_auto_fails():
    pass

# Test this
# Can we call row highlighter from this method?
# ^^ Shouldn't do it this way because sheet not written to for spreadsheetChecks only PreliminaryQC
def write_excelfile(ExcelFile):
    wb = openpyxl.load_workbook(ExcelFile.filePath)
 
    reset_colors(ExcelFile, wb, ExcelFile.failColors) # Only removes fail colors since this is independent of spreadsheetChecks

    for sheet in ExcelFile.sheetList:
        ws = wb[sheet.sheetName]

        for index, column_name in enumerate(list(ExcelFile.dataFrames[sheet.sheetName].columns.values)):
            ws.cell(1, index+1).value = column_name

            # This is where we call the field formatter
            # Any expansion to formatter requires an additional call here
            if column_name == ("date_created"):
                set_field_format(ws, "date_created", index)

        for r_idx, row in ExcelFile.dataFrames[sheet.sheetName].iterrows():
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx+2, column=c_idx).value = value

    wb.save(ExcelFile.filePath)
    wb.close()