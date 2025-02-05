"""
Authored by James Gaskell

1/24/2025

Edited by:

"""
import os, glob, re
from pathlib import Path


"""Conducts the preliminary QC checks
    Checks if the file exists in the file structure, if the extent is correct 
    and if the filesize is less than 300mb
    Adds the respecive failures to the count and to the file's failure dictionary
Args:
    sheet: excel sheet containing a list of files
    parent_directory: the OneDrive parent folder to search through
"""
def check_files(sheet, parent_directory):

    failures = 0

    for file in sheet.fileList:
        try:
            for path in Path(parent_directory).rglob(file.fileName + '.[pdf jpg]*'):

                file.filePath = (parent_directory + "\\" + file.fileName + ".pdf") # Can I do this using path above?
                # We can't assume this will be a pdf as single pages stored as jpg

                file.exists = True

                # We perhaps can't assume this is correct since some don't have parent folders
                if file.extent == len(os.listdir(path.parent.absolute())) - 1 or file.extent == None:
                    file.failures['Extent'] = False
                else:
                    file.failures['Extent'] = True
                    failures += 1

                if not (os.path.getsize(file.filePath) >> 20) < 300:
                    file.failures['Filesize'] = True
                    failures += 1

        except:
            pass
        
        if not file.exists:
            file.failures['Existence'] = True
            failures += 1

    sheet.failures = failures # Add number of failures to the sheet object