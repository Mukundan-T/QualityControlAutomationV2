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

        # The order of these is important for the error returns
        # The file is failed or highlighted based on the last error that matches
        # The dictionary must go in descending order of importance
        # Order:
        # Date
        # Filename
        # DupFilename
        # Extent
        # Filesize
        # Existance

        self.errors = {'Date': False,
                       'Filename': False,
                       'DupFilename': False,
                       'Extent': False,
                       'Filesize': False,
                       'Existance': False}




class ExcelSheet():

    def __init__(self, sheetname):
        self.sheetName = sheetname
        self.errors = 0
        self.failures = 0
        self.fileList: List[ScanFile] = list()

    def getErrorRate(self):
        return self.errors / len(self.fileList)
    
    def getFailureRate(self):
        return self.failures / len(self.fileList)

    """Creates a list of ScanFile objects from the dataframe"""
    def createScanFileList(self, df):
        for _, row in df.iterrows():
            record = ScanFile(row['Filename'],
                            row['Physical Location'],
                            row['date_created'] if not pd.isna(row['date_created']) else None,
                            row['extent (total page count including covers)'])
            
            self.fileList.append(record)

    """Supplies a dictionary of the errors in the sheet so the errors can be highlighted using openpyxl
    Returns:
        Dict: Dictionary of the errors in the sheet
    """
    def getSheetErrorDict(self):
        errorDict = {}
        for file in self.fileList:
            for key in file.errors:
                if file.errors[key]:
                    errorDict[file.fileName] = key
        return errorDict




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

    def getTotalError(self):
        return sum([sheet.errors for sheet in self.sheetList])
    
    def getTotalFailures(self):
        return sum([sheet.failures for sheet in self.sheetList])
    
    def getTotalFiles(self):
        return sum([len(sheet.fileList) for sheet in self.sheetList])
    
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
    
    """Supplies a dictionary of sheets with the errors nested inside
    Returns:
        Dict: Dictionary of the errors in the file
    """
    def getFileErrorDict(self):
        return {sheet.sheetName: sheet.getSheetErrorDict() for sheet in self.sheetList}
    
    """Method to add auto fail comments to the correct cells in the spreadsheet
        Updates the QC Results, QC Comments, QC initials columns depending on fail types"""
    def updateDataFrames(self):
        for sheet in self.sheetList:
            sheetDict = sheet.getSheetErrorDict()
            for key in sheetDict:
                self.dataFrames[sheet.sheetName].loc[self.dataFrames[sheet.sheetName]['Filename'] == key, 'QC Results'] = 'Fail'
                self.dataFrames[sheet.sheetName].loc[self.dataFrames[sheet.sheetName]['Filename'] == key, 'QC initials'] = 'AUTO'
                if sheetDict[key] == 'Extent':
                    self.dataFrames[sheet.sheetName].loc[self.dataFrames[sheet.sheetName]['Filename'] == key, 'QC Comments'] = 'Incorrect page count'
                if sheetDict[key] == 'Filesize':
                    self.dataFrames[sheet.sheetName].loc[self.dataFrames[sheet.sheetName]['Filename'] == key, 'QC Comments'] = 'Filesize too large'
                if sheetDict[key] == 'Existance':
                    self.dataFrames[sheet.sheetName].loc[self.dataFrames[sheet.sheetName]['Filename'] == key, 'QC Comments'] = 'File does not exist'
