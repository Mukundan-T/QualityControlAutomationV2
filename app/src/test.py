from config import DEFAULT_INPUT_FILE
import files, spreadsheetChecks

file = DEFAULT_INPUT_FILE

spreadsheet = files.ExcelFile(DEFAULT_INPUT_FILE)
spreadsheet.createFileStructure()

#print(spreadsheet)
#print(spreadsheet.sheetList)
#print(spreadsheet.sheetList[0].fileList)

spreadsheetChecks.check_date_format(spreadsheet.sheetList[1].fileList)
spreadsheetChecks.check_duplicate_filenames(spreadsheet.sheetList[1].fileList)
file_prefix = spreadsheetChecks.find_file_prefix(spreadsheet.sheetList[1].fileList)
spreadsheetChecks.check_location_filename(spreadsheet.sheetList[1].fileList)

date = 0
dups = 0
bad_filenames = 0
for file in spreadsheet.sheetList[1].fileList:
    if file.errors['Date'] == True:
        print(file.date)
        date +=1
    if file.errors['DupFilename'] == True:
        dups +=1
    if file.errors['Filename'] == True:
        bad_filenames +=1

print("found "+ str(date) + " date errors")
print("found "+ str(dups) + " duplicate filename errors")
#print("file prefix is: " + file_prefix)

print("found " + str(bad_filenames) + " filename errors")


