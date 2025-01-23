from config import DEFAULT_INPUT_FILE
import files, spreadsheetChecks

file = DEFAULT_INPUT_FILE

spreadsheet = files.ExcelFile(DEFAULT_INPUT_FILE)
spreadsheet.createFileStructure()

#print(spreadsheet)
#print(spreadsheet.sheetList)
#print(spreadsheet.sheetList[0].fileList)

spreadsheetChecks.check_date_format(spreadsheet.sheetList[1].fileList)

found = 0
for file in spreadsheet.sheetList[1].fileList:
    if file.errors['Date'] == True:
        print(file.date)
        found +=1

print("found "+ str(found) + " errors")
