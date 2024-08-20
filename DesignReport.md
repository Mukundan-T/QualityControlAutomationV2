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

Implementations:

* Spreadsheet validation - Physical Location and Filename
    - Some of the filenames are incorrect given the location.
    - Assumes naming conventions - Sheet, Bull, and a preceding 0 for single digits. Proceeding 0 convention could be relaxed with further implementation.

* Main Menu UI
    - Asks the user to select a spreadsheet file usin file explorer simple UI package.
    - Buttons for 'Preliminary Spreadsheet Checks', 'Quality Control'.
    - Built on PyQt5 so can be improved easily and made more visually appealing with QSS.

Notes:

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

Implementations:

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

Notes:

* Item repeats should be allowed a character identifier (a, b, c etc.). The process must be reworked slightly so it does not have to appear in the location
* Needs funtionality to allow for differences in input convention for 0 padding item number (i.e. .04 or .4)

</details>

<details>

<summary><b><u><font size="+1">08-16-2024</font></u></b></summary>

Implementations:

* Duplicate Filenames.
    - Runs after name/location error check so supersedes in importance
    - Colors rows blue where filenames are duplicated.
* Expanded naming conventions to allow the character after a filename which does not have to be reflected in location.
    - Since duplicates are highlighted this allows the user to but b, c, d next to the duplicate and run the program again.
* Modularized the color section so error type determines fill color, making it easier to add more error types in the future.
* Changed the success check to include multiple subrocesses on each sheet (filename, duplicate, datecheck etc.).

Notes:

Currently working on - 
* Date validation for date created.
* Date ISO formatting.

* Perhaps error colors should be on a palette and chosen by the user on the GUI to ensure no clashes with current spreadsheet highlighting?

</details>

<details>

<summary><b><u><font size="+1">08-20-2024</font></u></b></summary>

Implementations:

* Modularized the program further to make it easier to add functionality later on. 
    - Processes such as opening and closing files are within their own python file and have limited relience on current code.
    - Improved the code's readability and maintainability
* ISO date formatting



</details>





