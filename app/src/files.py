"""
Authored by James Gaskell

11/06/2024

Edited by:

"""
from typing import List
import pandas as pd

class ScanFile():

    def __init__(self, filename, loc, date, pages):
        self.fileName = filename
        self.location = loc
        self.date = date
        self.exists = False
        self.filePath = None 
        self.extent = pages
        self.errors = {'Date': False, 'Filename': False, 'DupFilename': False}
        self.too_large = too_large = False

class ExcelSheet():

    def __init__(self, sheetname):
        self.sheetName = sheetname
        self.fileList: List[ScanFile] = list()

    def createScanFileList(self, df):
        for index, row in df.iterrows():
            record = ScanFile(row['Filename'],
                            row['Physical Location'],
                            row['date_created'] if not pd.isna(row['date_created']) else None,
                            row['extent (total page count including covers)'])
            
            self.fileList.append(record)


class ExcelFile():

    def __init__(self, filepath):
        self.filePath = filepath
        self.sheetList: List[ExcelSheet] = list()
        self.dataFrames = None


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
        
                
