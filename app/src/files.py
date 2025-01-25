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
        self.filePath = None
        self.exists = False
        self.extent = pages # This is the extent read from the spreadsheet - not the actual extent
        self.errors = {'Date': False,
                       'Filename': False,
                       'DupFilename': False,
                       'Extent': False,
                       'Existance': False,
                       'Filesize': False}




class ExcelSheet():

    def __init__(self, sheetname):
        self.sheetName = sheetname
        self.fileList: List[ScanFile] = list()

    """Creates a list of ScanFile objects from the dataframe"""
    def createScanFileList(self, df):
        for _, row in df.iterrows():
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

    """Allows the user to change the spreadsheet in the UI without creating a new object instance
        Clears the old sheet list, which in turn clearsthe file storage
    Args:
        newpath (str): the new filepath selected
    """
    def setFilePath(self, newpath):
        self.filePath = newpath
        self.sheetList: List[ExcelSheet] = list()
    
    """Creates the file structure of an excel file creating each sheet and adding it to the sheetlist
        Method calls sheet.createScanFileList to create the list of files from the sheet dataframe
    """
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
        
                
