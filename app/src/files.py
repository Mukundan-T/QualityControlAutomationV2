"""
Authored by James Gaskell

11/06/2024

Edited by:

"""
from typing import List
import pandas as pd
import re

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
        # Existence
        self.errors = {'Date': False,
                       'Filename': False,
                       'DupFilename': False}
        
        self.failures = {
                        'Extent': False,
                        'Filesize': False,
                        'Existence': False}
    




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
            try:
                spreadsheet_extent = int(re.findall(r'\d+', row['extent (total page count including covers)'])[0])
            except:
                spreadsheet_extent = None
            record = ScanFile(row['Filename'],
                            row['Physical Location'],
                            row['date_created'] if not pd.isna(row['date_created']) else None,
                            spreadsheet_extent)
            
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
    
    def getSheetFailureDict(self):
        failureDict ={}
        for file in self.fileList:
            for key in file.failures:
                if file.failures[key]:
                    failureDict[file.fileName] = key
        return failureDict
    



class ExcelFile():

    # Colors
    # Incorrect Filename - Orange
    # Duplicate Filename - Blue
    # Incorrect Date - Yellow
    # Failure - Red


    errorColors = {"Filename":"FFEFBE7D", "DupFilename":"FF8BD3E6", "Date":"FFE9EC6B"} # This is outside the __init__ method so it is shared between all instances of the class
    failColors = {"Extent":"FFFFADB0", "Filesize":"FFFFADB0", "Existence":"FFFFADB0"} #Done this way to allow expansion of the fail types

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
    
    def setErrorColor(self, errorType, newColor):
        self.errorColors[errorType] = newColor

    def setFailColor(self, newColor):
        self.failColor = newColor


    """Creates the file structure of an excel file creating each sheet and adding it to the sheetlist
        Method calls sheet.createScanFileList to create the list of files from the sheet dataframe
    """
    # It might be a better design decision to put this in FileHandler but this works for now
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
    # Can't test until I can access the onedrive file structure
    def updateDataFrames(self):
        for sheet in self.sheetList:
            sheetDict = sheet.getSheetFailureDict()
            for key in sheetDict:


                # Clear failures from dataframe before adding new ones
                # This means corrected autofails can be overwritten
                #if self.dataFrames[sheet.sheetName]['QC Results'] == "Fail" and self.dataFrames[sheet.sheetName]['QC Initials'] == "AUTO":
                #
                #    self.dataFrames[sheet.sheetName]['QC Results'] = ""
                #    self.dataFrames[sheet.sheetName]['QC Comments'] = ""
                #    self.dataFrames[sheet.sheetName]['QC Initials'] = ""

                self.dataFrames[sheet.sheetName]['QC Results'] = self.dataFrames[sheet.sheetName]['QC Results'].astype(str).replace(to_replace='nan', value='', regex=True)
                self.dataFrames[sheet.sheetName]['QC Comments'] = self.dataFrames[sheet.sheetName]['QC Comments'].astype(str).replace(to_replace='nan', value='', regex=True)
                self.dataFrames[sheet.sheetName]['QC Initials'] = self.dataFrames[sheet.sheetName]['QC Initials'].astype(str) .replace(to_replace='nan', value='', regex=True)
                # Order matters here - if an error is matched in the order; Existence, Filesize, Extent then the others are ignored
                # This means the failures go in order of precedence
                # A file may fail for more than one of these reasons, only the mopst important reason is recorded
                if sheetDict[key] in ['Extent', 'Filesize', 'Existence']:
                    self.dataFrames[sheet.sheetName].loc[self.dataFrames[sheet.sheetName]['Filename'] == key, 'QC Results'] = 'Fail'
                    self.dataFrames[sheet.sheetName].loc[self.dataFrames[sheet.sheetName]['Filename'] == key, 'QC Initials'] = 'AUTO'
                if sheetDict[key] == 'Existence':
                    self.dataFrames[sheet.sheetName].loc[self.dataFrames[sheet.sheetName]['Filename'] == key, 'QC Comments'] = 'File does not exist'
                elif sheetDict[key] == 'Filesize':
                    self.dataFrames[sheet.sheetName].loc[self.dataFrames[sheet.sheetName]['Filename'] == key, 'QC Comments'] = 'Filesize too large'
                elif sheetDict[key] == 'Extent':
                    self.dataFrames[sheet.sheetName].loc[self.dataFrames[sheet.sheetName]['Filename'] == key, 'QC Comments'] = 'Incorrect page count'
