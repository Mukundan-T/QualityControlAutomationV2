"""
Authored by James Gaskell

1/24/2025

Edited by:

"""
from config import DEFAULT_INPUT_FILE, DEFAULT_ONEDRIVE_FOLDER
import files, spreadsheetChecks, preliminaryQC, fileHandler
from tkinter import messagebox

file = DEFAULT_INPUT_FILE

file_open = fileHandler.file_open_check(DEFAULT_INPUT_FILE) #No point in doing any analysis unless spreadsheet is closed at the start
while file_open:
    messagebox.showerror("Error", "The excel file is open in editor. Please close it and try again")
    file_open = fileHandler.file_open_check(DEFAULT_INPUT_FILE)

spreadsheet = files.ExcelFile(DEFAULT_INPUT_FILE)
spreadsheet.createFileStructure()

sheet_num = 0

for sheet in spreadsheet.sheetList:
    spreadsheetChecks.check_date_format(sheet)
    spreadsheetChecks.check_duplicate_filenames(sheet)
    spreadsheetChecks.check_location_filename(sheet)
    preliminaryQC.check_files(sheet, DEFAULT_ONEDRIVE_FOLDER)

    date = 0
    dups = 0
    bad_filenames = 0
    for file in sheet.fileList:
        if file.errors['Date'] == True:
            date +=1
        if file.errors['DupFilename'] == True:
            dups +=1
        if file.errors['Filename'] == True:
            bad_filenames +=1

    print("found "+ str(date) + " date errors")
    print("found "+ str(dups) + " duplicate filename errors")
    print("found " + str(bad_filenames) + " filename errors")

total_errors = spreadsheet.getTotalError()
print("total errors accross all sheets: " + str(total_errors))
total_failures = spreadsheet.getTotalFailures()
print("total failures accross all sheets: " + str(total_failures))

print("error rate: " + str(total_errors/spreadsheet.getTotalFiles() * 100))
print("failure rate: " + str(total_failures/spreadsheet.getTotalFiles() * 100))

spreadsheet.updateDataFrames()

success = fileHandler.write_excelfile(spreadsheet)
if success:
    success = fileHandler.highlight_errors(spreadsheet)
else:
    messagebox.showerror("Error","The excel file is open in editor so changes could not be saved")


print("Success")

