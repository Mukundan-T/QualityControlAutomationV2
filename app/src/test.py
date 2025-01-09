from config import DEFAULT_INPUT_FILE
import files

file = DEFAULT_INPUT_FILE

spreadsheet = files.ExcelFile(DEFAULT_INPUT_FILE)
spreadsheet.createFileStructure()

print(spreadsheet)
print(spreadsheet.sheetList)
print(spreadsheet.sheetList[0].fileList)

