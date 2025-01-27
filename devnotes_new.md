# Refactoring development notes

* Should 'undated', 'none' etc be valid for the date column? In which case I can attempt formatting and leave all else - i.e. don't have to highlight errors if it won't lead to an upload error

* Could really do with a cleanup of the date checks. The filename checks are now a lot more pallateable, date checks need to be just as readable

* What is the best way to handle file output? We need to reference the objects to see if they have errors or fails but want to do this as efficiently as possible.

    1. We could request the failures of each type from the file objects using getters
        + If we could make this a dictionary of its own then it becomes more efficient to reference dictionary keys
        + nested dictionary? - split the method into a sheet error getter, use this with an excelfile to make a dictionary of sheets with a dictionary of errors inside?
        - Adds complexity to the objects

* Added a lot of getter methods to the object structure. At the end we need to look through these and decide which are necessary and which aren't. We can probably do this by just ctrl-f for calls to thee functions.

* Still looking like we maintain the following order:
    1. Update dataframes with auto fails
    2. Output to the excel file using openpyxl
    3. Clear the colors that are hard programmed so we can run the program in a recursive process
    4. Apply color formatting to errors
    5. Close file

* Does the singleton class need to be in the same file as the UI to prevent circular imports?
    * Ui should be able to both get and set the filepath
    * UI should have access to the single instance of ExcelFile

* Broadly speaking would it be a good idea to extract the fails into another metadata spreadsheet? This could make the QC process more recursive.
    1. Scan
    2. Conduct checks
    3. Generate new, shorter spreadsheet with just fails
    4. Back to step 1

    Could be interesting to add this functionality - would definitely make the fails a little clearer for the person that has to go through and itentify records to sort out

* Should we split the failures and the errors into separate dictionaries inside the excel file/sheet objects? The prelimQC and spreadsheetChecks are supposedly independent of one another so this would reduce complexity when run on their own. Less errors to check through for each item.