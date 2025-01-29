"""
Authored by James Gaskell

1/24/2025

Edited by:

"""
import os, glob, re
from pathlib import Path

def check_files(sheet, parent_directory):

    failures = 0

    for file in sheet.fileList:
        try:
            for path in Path(parent_directory).rglob(file.fileName + '.[pdf jpg]*'):

                file.filePath = (parent_directory + "\\" + file.fileName + ".pdf") # Can I do this using path above?
                # We can't assume this will be a pdf as single pages stored as jpg

                file.exists = True

                # We perhaps can't assume this is correct since some don't have parent folders
                file.failures['Extent'] = False if file.extent == len(os.listdir(path.parent.absolute())) - 1 or file.extent == None else True

                # 
                file.failures['Filesize'] = False if (os.path.getsize(file.filePath) >> 20) < 300 else True

        except:
            pass
        
        if not file.exists:
            file.failures['Existance'] = True
            failures += 1

    sheet.failures = failures # Add number of failures to the sheet object