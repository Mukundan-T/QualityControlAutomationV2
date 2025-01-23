# Design report and work diary

## Initial Design Brief

Create a program to automate aspects of the current quality control process for digitized documents that are arduous and time consuming.

### What we look for in Quality Control

1. File Exists?
    - The file should be locatable within the Onedrive folder.
    - Any files that cannot be found have not uploaded correctly so should be marked as fail.
2. Number of pages
    - Extent field of the metadata spreadsheet informs the QCer of the expected page count of the PDF.
    - Any PDFs that do not have the correct number of pages should automatically fail.
3. Filenames
    - The physical filename must match the filename in the spreadsheet.
    - Since the spreadsheet may contain errors to the filenames that have been carried over, it may be useful to develop a check to ensure location and filename match before scanning.
    - Filenames currently do not have stringent conventions, they must allow for bulletins, sheets, items etc.
    - Repeat filenames may exist if there is a duplicate item within each box/folder.
4. Image Size
    - Any image files over 500mb cannot be processed by ARCHES (the online archive).
    - Any folders containing images too large should automatically fail.
5. Date Format
    - The upload program for ARCHES only recognizes ISO format, date_created should be converted to ISO in the spreadsheet.
    - Blanks should be ignored, any other issues should be indicated.

### Work Diary and Implementations By Date

<details>

<summary><b><u><font size="+1">08-14-2024</font></u></b></summary>

<b>Implementations:</b>

* Spreadsheet validation - Physical Location and Filename
    - Some of the filenames are incorrect given the location.
    - Assumes naming conventions - Sheet, Bull, and a preceding 0 for single digits. Proceeding 0 convention could be relaxed with further implementation.

* Main Menu UI
    - Asks the user to select a spreadsheet file usin file explorer simple UI package.
    - Buttons for 'Preliminary Spreadsheet Checks', 'Quality Control'.
    - Built on PyQt5 so can be improved easily and made more visually appealing with QSS.

<b>Notes:</b>

* Could we create a Fragile column? Since fragile items currently are manually indicated with highlighting, preliminary checks would be able to color the spreadsheet rows in blue for fragile items.
* Use Python’s PDF reader to view x random pdf pages to automate pass/fail.
* Initial entry to fill in who the QCer is each time? Happens once at the start - current workflow uses initials to show who QC'd.
* Any changes to the naming conventions of the files. Current convention is:
    - Bull used to denote a bulletin -> ZWU_SCA0319.B06.F01.Bull.107.
    - Sheet denotes a sheet -> ZWU_SCA0319.B06.F01.Sheet.564.
    - Item has no name -> ZWU_SCA0319.B06.F05.01.

</details>

<details>

<summary><b><u><font size="+1">08-15-2024</font></u></b></summary>

<b>Implementations:</b>

* Filename validation
    - Now colors discrepancies between location and filename in red.
    - Ignores other colors on the spreadsheet, but removes the program's own error colors from the sheet so it can be used .recursively. I.e. after errors are fixed they will return to no fill.
    - Can therefore be continually run into no color is left on the spreadsheet.

* Added error rate to terminal output
    - Should be on the PyQt5 window in later versions.

* Improved GUI using PyQt5 built in styles.

* Added exception handling to ensure the file is closed before the program attempts to access anything.

* Expanded sheet naming conventions to include:
    - Sheet &rarr; .Sheet.
    - Bulletin &rarr; .Bull.
    - Item &rarr; no prefix

<b>Notes:</b>

* Item repeats should be allowed a character identifier (a, b, c etc.). The process must be reworked slightly so it does not have to appear in the location
* Needs funtionality to allow for differences in input convention for 0 padding item number (i.e. .04 or .4)

</details>

<details>

<summary><b><u><font size="+1">08-16-2024</font></u></b></summary>

<b>Implementations:</b>

* Duplicate Filenames.
    - Runs after name/location error check so supersedes in importance
    - Colors rows blue where filenames are duplicated.

* Expanded naming conventions to allow the character after a filename which does not have to be reflected in location.
    - Since duplicates are highlighted this allows the user to but b, c, d next to the duplicate and run the program again.

* Modularized the color section so error type determines fill color, making it easier to add more error types in the future.

* Changed the success check to include multiple subrocesses on each sheet (filename, duplicate, datecheck etc.).

<b>Notes:</b>

Currently working on - 
* Date validation for date created.
* Date ISO formatting.

* Perhaps error colors should be on a palette and chosen by the user on the GUI to ensure no clashes with current spreadsheet highlighting?

</details>

<details>

<summary><b><u><font size="+1">08-20-2024</font></u></b></summary>

<b>Implementations:</b>

* Modularized the program further to make it easier to add functionality later on. 
    - Processes such as opening and closing files are within their own python file and have limited relience on current code.
    - Improved the code's readability and maintainability

* ISO date formatting
    - date_created checked for datetime, then converted to ISO format
    - date_times that are not type(datetime) go through a conversion process that includes spellcheck for written dates
    - Dates that cannot be converted by the program are added to a list for highlighing - since we cannot be certain that the program will work on every instance it is better to let the user decide in the spreadsheet
    - Highlights rows with date errors yellow in the spreadsheet

* File_writer
    - Can be called at any time after data is changed in the main dataframe
    - Contains functionality to format cells - this currently sets the excel format for date_time to YYYY-MM-DD but can be used to add color or font styles to single cells
    - Successfully writes the dataframe over the original excel file, and maintains original formatting unless otherwise specified

<b>Notes:</b>

* Could File_reader_writer be an object since it has an increasing number of instance variables?
* Need to look into how we can recursively check the source folder for filenames to implement file_exists?
* Does quality control need to be performed by sheet? If so we could add another page to the UI that uses the sheetnames of the read file to ask the user which sheet they would like to QC.


</details>

<details>

<summary><b><u><font size="+1">08-21-2024</font></u></b></summary>

<b>Implementations:</b>

* Started to improve the reusability of my code by creating file and excel file objects to store data that is referred to multiple times throughout my program by different subroutines
    1. Object File: contains Filepath, PDF_filepath, Extent, File_size and Exists.
    2. Object Excel File: contains sheetnames, dataframes and Filepath.

* Designed a singleton to contain a list of files, a spreadsheet, a dictionary of error colors and a parent directory for sharepoint. Subprocesses should wuery the singlton for file data, excel data etc.

* Redesigned the User Interface on paper to make it more intuative and to understand what is needed in the singleton

<b>Notes:</b>

</details>

<details>

<summary><b><u><font size="+1">09-17-2024</font></u></b></summary>

<b>Implementations:</b>

* Added pdf pages functionality to count the number of files in the subfolder since the individual image files are stored alongside the full pdf
    * Added folderName to File objects to support this
* Added functionality to identify the files that do not exist using a boolean flag
    * Changed the print output to inform wether the file exists or not

<b>Notes:</b>

* Need to add an autofail by changing the Pass/Fail field in the dataframe and using ExcelReaderWriter to update the spreadsheet accordingly

* Would be nice to group the widgets that do the same things together in the UI and have subwidgets to make them easier to move around

* Could I add a spreadsheetExtent variable to the file pbject and compare this to the real extent of the file to determine the pass/fail?


</details>


<details>

<summary><b><u><font size="+1">09-18-2024</font></u></b></summary>

<b>Implementations</b>

* Added new datapoints to the File object for the extent and to show wether the file is too big for upload to AWS

* Added auto pass/fail if:
    - The filename cannot be located in the file structure selected by the user
    - The extent is incorrect (user dependent since the feature only works if the extent has been inputted at the time of scanning)
    - The file is larger than 300MB
On current spreadsheet the automatic failure rate is roughly 37% when all of these checks are complete

* Added functiuonality to output a fail message tailored to the type of failure. This way the fails can be double checked manually

<b>Notes</b>

* Could I add the conditional formatting rules into the program so openpyxl adds them rather than having to be done manually?

* Could I add a key for the spreadsheet highlightings? This may not be necessary if I manage to fully implement the color selection
    Current colors:
    1. Red - Filename mismatch
    2. Blue - Duplicate filename
    3. Yellow - Date format error

* Current goals
    1. Improve user interface. Add lines to break up sections since there are 3 distinct parts. Maybe place buttons on subwidgets making it easier to add and take away features
    2. Begin to work on full QC (third option on the UI)
    3. Test the spreadsheet checks on the new spreadsheet and see how it holds up


### Refactoring to Object Oriented

