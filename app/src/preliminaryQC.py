"""
Authored by James Gaskell

1/24/2025

Edited by:

"""
import os, glob, re
from pathlib import Path

def check_files(file_list, parent_directory):

    failures = 0

    for file in file_list:
        for path in Path(parent_directory).rglob(file.fileName + '.[pdf jpg]*'):

            file.filePath = (parent_directory + "\\" + file.fileName + ".pdf") # Can I do this using path above?
            # We can't assume this will be a pdf as single pages stored as jpg

            file.exists = True

            file.errors['Extent'] = True if os.path.getsize(file.filePath) >> 20 else False

            # We can't assume this is correct since some don't have parent folders
            file.extent = len(os.listdir(path.parent.absolute())) - 1
        
        if not file.exists:
            file.errors['Existance'] = True
            failures += 1

    return (failures/len(file_list) * 100) # Return percentage of failures