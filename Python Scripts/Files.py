"""
Authored by James Gaskell

08/21/2024

Edited by:

"""

import pandas as pd
import tkinter as tk

class File():

    filepath = None
    extent = None
    file_size = None
    too_large = None
    exists = False
    filename = None
    parentDirectory = None

    def __init__(self, filename):
        self.filename = filename

class ExcelFile(File):

    sheetnames = []
    dataframes = None

    def __init__(self, filepath):
        self.filepath = filepath
    
    def read_data_frames(self):
        try:
            xl_file = pd.ExcelFile(self.filepath)
            self.dataframes = pd.read_excel(xl_file, sheet_name=None)
            self.sheetnames = xl_file.sheet_names
            xl_file.close()
        except:
            tk.messagebox.showinfo("Error", "File error. Please try again")

    