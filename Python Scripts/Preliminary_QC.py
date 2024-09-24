"""
Authored by James Gaskell

08/22/2024

Edited by:

"""

import Files, glob, os, Excel_reader_writer, re
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

def file_checker(ProgramData):
    files = 0
    failures = 0

    for key in ProgramData.Files_List:

        ProgramData.Spreadsheet.dataframes[key]['QC Pass/Fail'] = ProgramData.Spreadsheet.dataframes[key]['QC Pass/Fail'].astype(str)
        ProgramData.Spreadsheet.dataframes[key]['QC Comments'] = ProgramData.Spreadsheet.dataframes[key]['QC Comments'].astype(str)
        ProgramData.Spreadsheet.dataframes[key]['QC Initials'] = ProgramData.Spreadsheet.dataframes[key]['QC Initials'].astype(str)

        Linenumber = 0
        for File in ProgramData.Files_List[key]:
            files += 1
            for path in Path(ProgramData.Parent_Directory).rglob(File.filename + '.[pdf jpg]*'):
                File.filepath = (ProgramData.Parent_Directory + "\\" + File.filename + ".pdf")
                File.folderName = path.parent.absolute()
                File.exists = True
                File.file_size = os.path.getsize(path) >> 20
                File.extent = len(os.listdir(File.folderName)) - 1
                if File.file_size > 300:
                    File.too_large = True #Slightly redundant but better for data visualization should it be needed
            
            #Needs this or the script writes out Nans to the spreadsheet
            ProgramData.Spreadsheet.dataframes[key].loc[Linenumber, "QC Pass/Fail"] = ""
            ProgramData.Spreadsheet.dataframes[key].loc[Linenumber, "QC Comments"] = ""
            ProgramData.Spreadsheet.dataframes[key].loc[Linenumber, "QC Initials"] = ""
            
            if not File.exists:
                ProgramData.Spreadsheet.dataframes[key].loc[Linenumber, "QC Pass/Fail"] = "Fail"
                ProgramData.Spreadsheet.dataframes[key].loc[Linenumber, "QC Comments"] = "File cannot be located"
                ProgramData.Spreadsheet.dataframes[key].loc[Linenumber, "QC Initials"] = "AUTO"
                # print(File.filename + " Could not be located, QC: " + ProgramData.Spreadsheet.dataframes[key]["QC Pass/Fail"][Linenumber])
                failures += 1
            if File.too_large:
                ProgramData.Spreadsheet.dataframes[key].loc[Linenumber, "QC Pass/Fail"] = "Fail"
                ProgramData.Spreadsheet.dataframes[key].loc[Linenumber, "QC Comments"] = "File too large for upload"
                ProgramData.Spreadsheet.dataframes[key].loc[Linenumber, "QC Initials"] = "AUTO"
                # print(File.filename + " Too large for upload, QC: " + ProgramData.Spreadsheet.dataframes[key]["QC Pass/Fail"][Linenumber])
                failures += 1

            spreadsheet_extent = ProgramData.Spreadsheet.dataframes[key]["extent (total page count including covers)"][Linenumber]
            try:
                spreadsheet_extent = int(re.findall(r'\d+', spreadsheet_extent)[0])
                if File.extent != spreadsheet_extent and File.exists:
                    ProgramData.Spreadsheet.dataframes[key].loc[Linenumber, "QC Pass/Fail"] = "Fail"
                    ProgramData.Spreadsheet.dataframes[key].loc[Linenumber, "QC Comments"] = "Extent and pagecount do not match"
                    ProgramData.Spreadsheet.dataframes[key].loc[Linenumber, "QC Initials"] = "AUTO"
                    print(File.filename + " Has an incorrect extent, QC: " + ProgramData.Spreadsheet.dataframes[key]["QC Pass/Fail"][Linenumber])
            except TypeError:
                spreadsheet_extent = 0
                

            Linenumber += 1
    print("Failure rate: " + str(failures/files * 100) + "%")




def run_checks(ProgramData):
    create_possible_files(ProgramData)
    file_checker(ProgramData)
    Excel_reader_writer.df_to_excel(ProgramData.Spreadsheet.dataframes, ProgramData.Spreadsheet.sheetnames, ProgramData.Spreadsheet.filepath)
