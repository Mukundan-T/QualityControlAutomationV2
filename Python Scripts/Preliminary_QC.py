"""
Authored by James Gaskell

08/22/2024

Edited by:

"""

import Files, glob, os
from pathlib import Path
from pdfreader import PDFDocument

def create_possible_files(ProgramData):
    for sheet in ProgramData.Spreadsheet.sheetnames:
        for index, row in ProgramData.Spreadsheet.dataframes[sheet].iterrows():
            new_file = Files.File(row['Filename'])
            if sheet not in ProgramData.Files_List:
                ProgramData.Files_List[sheet]=[new_file]
            else:
                ProgramData.Files_List[sheet].append(new_file)

def file_exists(ProgramData):
    for key in ProgramData.Files_List:
        for File in ProgramData.Files_List[key]:
            for path in Path(ProgramData.Parent_Directory).rglob(File.filename + '.pdf'):
                File.filepath = path
                File.folderName = path.parent.absolute()
                File.exists = True
                File.file_size = os.path.getsize(path) >> 20
                File.extent = len(os.listdir(File.folderName)) - 1

            print(File.filename + " Exists :" + str(File.exists) + ". Has " + str(File.extent) + " pages. And is " + str(File.file_size) +"MB")
    print("done")

def run_checks(ProgramData):
    create_possible_files(ProgramData)
    file_exists(ProgramData)
