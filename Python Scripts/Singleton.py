"""
Authored by James Gaskell

08/21/2024

Edited by:

"""

import Files, easygui
import Spreadsheet_checks
import tkinter as tk

class Program():

    Files_List = []
    Spreadsheet = None
    Parent_Directory = None
    Error_Colors = {"Filename":"FFFFADB0", "Duplicate":"FFADD8E6", "DateFormat":"FFFDFD96"}

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
            Spreadsheet_checks.run_checks(self.Spreadsheet)
        else:
            tk.messagebox.showerror("Error", "No file loaded")


def under_construction():
    tk.messagebox.showerror("Error", "Under Construction")