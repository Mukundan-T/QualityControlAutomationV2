"""
Authored by James Gaskell

08/21/2024

Edited by:

"""

import Files, easygui
import Spreadsheet_checks, Preliminary_QC, Excel_reader_writer
import tkinter as tk

class Program():

    Files_List = {}
    Spreadsheet = None
    Parent_Directory = None
    Error_Colors = {"Filename":"FFFFADB0", "Duplicate":"FFADD8E6", "DateFormat":"FFFDFD96", "QCFail":"FF7F7F7F"}

    def __init__(self):
        pass
    
    def get_excel_file(self):
        path = easygui.fileopenbox()
        self.Spreadsheet = Files.ExcelFile(path)
        self.Spreadsheet.read_data_frames()
    
    def get_parent_directory(self):
        path = easygui.diropenbox()
        self.Parent_Directory = path

    def run_spreadsheet_checks(self):
        if self.Spreadsheet != None:
            File_open = Excel_reader_writer.file_open_check(self.Spreadsheet.filepath)
            if File_open:
                tk.messagebox.showerror("Error", "The excel file is currently open. Please close it before continuing")
            else:
                Spreadsheet_checks.run_checks(self.Spreadsheet)
        else:
            tk.messagebox.showerror("Error", "No file loaded")

    def run_prelim_QC(self):
        if self.Spreadsheet != None:
            File_open = Excel_reader_writer.file_open_check(self.Spreadsheet.filepath)
            if File_open:
                tk.messagebox.showerror("Error", "The excel file is currently open. Please close it before continuing")
            else:
                Preliminary_QC.run_checks(self)



def under_construction():
    tk.messagebox.showerror("Error", "Under Construction")