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
    - Blanks should be ignored, any other issues should be indicated .


