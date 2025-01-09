"""
Authored by James Gaskell

11/06/2024

Edited by:

"""
from typing import List
import openpyxl
import pandas as pd

class ScanFile():

    fileName = None 
    # Real filename read in from spreadsheet
    location = None 
    # This will be used to construct the expected filename
    date = None
    # Stores the date for format checking
    exists = False
    # Results of searching for the file
    filePath = None 
    # Only populates if exists is true
    errors = None 
    # Used to track date errors, filename errors etc.
    extent = 0
    # Allows comparison between real extent and spreadsheet recorded extent
    too_large = False
    # If the file size is > 200Mb (subject to change)

    def __init__(self, filename, loc, date, pages):
        self.fileName = filename
        self.location = loc
        self.date = date
        self.extent = pages

class ExcelSheet():

    sheetName = None
    # the name of the sheet in the spreadsheet - used as a unique identifier
    fileList: List[ScanFile] = list()
    # List of files on that sheet

    def __init__(self, sheetname):
        self.sheetName = sheetname

    def createScanFileList(self, df):
        for index, row in df.iterrows():
            record = ScanFile(row['Filename'], row['Physical Location'], row['date_created'], row['extent (total page count including covers)'])
            self.fileList.append(record)




class ExcelFile():

    filePath = None
    # Filepath of the spreadsheet chosen in UI
    sheetList: List[ExcelSheet] = list()
    # List of sheet objects in the spreadsheet
    dataFrames = None
    # Pandas dataframe containing all of the spreadsheet data for each record

    def __init__(self, filepath):
        self.filePath = filepath


    def setFilePath(self, newpath):
        """
        Allows the user to change the spreadsheet in the UI without creating a new object instance
        Clears the old sheet list, which in turn clearsthe file storage

        Args:
            newpath (str): the new filepath selected
        """
        self.filePath = newpath
        self.sheetList: List[ExcelSheet] = list()
    
    def createFileStructure(self):
        try:
            if self.filePath == None:
                raise Exception ("No Spreadsheet Selected")
            else:
                xl_file = pd.ExcelFile(self.filePath)
                self.dataFrames = pd.read_excel(xl_file, sheet_name=None)
                [self.sheetList.append(ExcelSheet(sheet)) for sheet in xl_file.sheet_names] #Creates a new sheet object from the sheetname
                [sheet.createScanFileList(self.dataFrames[sheet.sheetName]) for sheet in self.sheetList] #Creates the list of files from the sheet dataframe
        except Exception as error:
            return error
        
                
