"""
Authored by James Gaskell

08/12/2024

Edited by:

"""

import numpy as np
import pandas as pd
import tkinter as tk
import time, sys, random, openpyxl, os, Date_formatter, Excel_reader_writer, Location_checker


"""Print slow function to simulate typing
    makes interfce more friendly but not necessarily needed. Typing speed changed by altering typing_speed to reflect words per minute
Args:
    String t: The string typed out
"""
def print_slow(t):
    typing_speed = 60 #wpm
    for l in t:
        sys.stdout.write(l)
        sys.stdout.flush()
        time.sleep(random.random()*10.0/typing_speed)
    print ('')


"""Resets the color of rows in the worksheet so they can be reassessed based on the progra's findings
Args:
    worksheet: current sheet opened as part of a Workbook with the openpyxl library
    Int: i refering to the row index after opening the worksheet
    Dict: the colors that are used by the program to highlight errors
        This way we can avoid removing colorson the spreadsheet not caused by previous program runs

"""
def reset_color_rows(ws, i, error_fills):
    fill_reset = openpyxl.styles.PatternFill(fill_type=None)
    if ws.cell(row=i+2, column=1).fill.start_color.index in error_fills.values():
        for y in range(1, ws.max_column+1):
            ws.cell(row=i+2, column=y).fill = fill_reset


"""Identifies the problem rows in the spreadsheet and highlights them depending on the error type
Args:
    String: sheetname - the current sheet being worked on
    Array[String]: the current problem rows, identified by their filename
    String: the type of error currently being added to the spreadsheet - this dictates highlight color
    Boolean: informs the subroutine wether a color reset is required for the spreadsheet - runs when the first error highlight is being added
Returns:
    Boolean: depending on wether the process was successful or not. Fails usually occur due to the excel spreadsheet being open in another application
"""
def identify_problem_rows(Spreadsheet, sheetname,  problem_rows, type, reset_rows):

    #type will be used to determine if the passed problem rows are filename, date format etc.

    error_fills = {"Filename":"FFFFADB0", "Duplicate":"FFADD8E6", "DateFormat":"FFFDFD96"}

    if type in error_fills.keys():
        fill_hex = error_fills[type]
    else:
        fill_hex = type

    err_fill = openpyxl.styles.PatternFill(start_color=fill_hex, end_color=fill_hex, fill_type="solid")

    xl_file = pd.ExcelFile(Spreadsheet.filepath)       
    dt = pd.read_excel(xl_file, sheetname)

    wb = openpyxl.load_workbook(Spreadsheet.filepath)
    ws = wb[sheetname]

    for index, row in dt.iterrows():

        if reset_rows:
            reset_color_rows(ws, index, error_fills)

        for p in problem_rows:
            #print(dt['Filename'][index])
            if p == dt['Filename'][index]:
                for y in range(1, ws.max_column+1):
                    ws.cell(row=index+2, column=y).fill = err_fill
    
    try: 
        wb.save(Spreadsheet.filepath)
    except:
        wb.close()
        xl_file.close()
        return False

    wb.close()
    xl_file.close()
    
    return True


"""Creates the report text files containing errors
Args:
    String: sheetName denoting which sheet is currently being worked on - Box 5, Box 6 etc.
    Array[String]: reportMajor containing the error messages ad filenames to be outputted in the report
    Array[String]: reportNoLocation containing minor location errors for the report
    Array[String]: reportMinor will be used for other errors that aren't in these categories
"""   
def createReportText(sheetName, reportMajor, reportNoLocation, reportMinor):

    newpath = (os.getcwd() + "\\Reports\\")
    if not os.path.exists(newpath):
        os.makedirs(newpath)

    outputFile = open(newpath + sheetName + r" Report.txt", "w+")
    outputFile.write("------------Major Issues------------\n\n")
    for error in reportMajor:
        outputFile.write(error + "\n")
    outputFile.write("\n\n\n\n------------Minor Issues------------\n\n")
    for error in reportNoLocation:
        outputFile.write(error + "\n")
    for error in reportMinor:
        outputFile.write(error + "\n")
    outputFile.close()


"""The sheet loop that checks each sheet for errors and creates the error reports
    Contains error checking for the whole program - if any process is not successful and is unhandled an error message is produced
Args:
    Dict: dataframes - all the dataframes in the document split by sheet
    Array[String]: the names of the sheets in the document
"""
def run_checks(Spreadsheet):

    print_slow("File loaded...")

    errors = False

    for sheet in Spreadsheet.sheetnames:

        reportNoLocation = [] # These must be in this subroutine loop to prevent report data carrying over to otherr boxes
        reportMajor = []
        reportMinor = []
        location_problem_rows = []
        duplicate_problem_rows = []
        date_problem_rows = []
        sheet_success = []

        success, Spreadsheet.dataframes[sheet] = Date_formatter.check_date_format(Spreadsheet.dataframes[sheet], reportMajor, date_problem_rows) #First check for date formatting issues

        if success:
            sheet_success.append(identify_problem_rows(Spreadsheet, sheet, date_problem_rows, "DateFormat", True))
            success = Location_checker.check_location_filename(Spreadsheet.dataframes[sheet], reportMajor, reportNoLocation, location_problem_rows) #Check for location/filename discrepancies
        if success:
            sheet_success.append(identify_problem_rows(Spreadsheet, sheet, location_problem_rows,"Filename", False)) #Colors the problem rows in red for location/filename problems, resets rows from previous runs of the program
            success = Location_checker.check_duplicate_filenames(Spreadsheet.dataframes[sheet], reportMajor, duplicate_problem_rows) #Second check - duplicate filenames
        if success:
            sheet_success.append(identify_problem_rows(Spreadsheet, sheet, duplicate_problem_rows, "Duplicate", False)) #Colors the problem rows in blue for duplicates, doesn't reset rows as it is the scond step
        if False not in sheet_success:
            if len(reportMajor) != 0:
                errors = True
            createReportText(sheet, reportMajor, reportNoLocation, reportMinor)
            error_rate = round((len(reportMajor)/ (Spreadsheet.dataframes[sheet].shape[0]) * 100) , 1)
            print(sheet + " Completed.     Error Rate: " + str(round(error_rate,1)) + "%")
        else:
            break

    if False in sheet_success:
        tk.messagebox.showerror("File error", "Save Failed. Please ensure the excel file is closed and try again")
    else:
        Excel_reader_writer.df_to_excel(Spreadsheet.dataframes, Spreadsheet.sheetnames, Spreadsheet.filepath)
        if errors:
            msg = "Success! Please check the excel file for issues"
        else:
            msg = "Success! There were no recorded issues. Continue to Quality Control"
        tk.messagebox.showinfo("Success", msg)



