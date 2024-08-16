"""
Authored by James Gaskell and Corinne Chatnik

08/14/2024

"""

import pandas as pd
import time, sys, easygui, random, string, datetime, openpyxl

def get_file():
    path = easygui.fileopenbox()
    return path

def get_dataFrames(filepath):
    xl_file = pd.ExcelFile(filepath)
    dfs = pd.read_excel(xl_file, sheet_name=None)
    sheets = xl_file.sheet_names
    return dfs,sheets

def main():
    filepath = get_file()
    dfs, sheets = get_dataFrames(filepath)

    