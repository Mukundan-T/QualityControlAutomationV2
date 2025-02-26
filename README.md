# Quality Control Automation Program

This is the current working version of the Quality Control Program designed to fit into Schaffer Library's digitization workflow.

## Setup

To set up the program:

1. Navigate to scripts
2. Click on setup.cmd. This will install requirements and create a shortcut to the main program file
3. Locate the desktop shortcut and run the program!

The program works with python 3.12.5, and will likely work with later versions of python but you may need an update if any of the requirements fail to install. This will be indicated with an error upon running the setup.

## Features

1. Spreadsheet checks - this function checks for:
    * Incorrect date formats. Highlighted or corrected based on how extreme the issue is.
    * Mismacthed locations and filenames.
    * Duplicate filenames.

2. Preliminary QC - this function checks for:
    * Page extent. Does the real pagecount match what is listed on the spreadsheet?
    * File existance. Has the scanned file made it to the OneDrive folder successfully?
    * File size. Is the file too big for upload to Archipelago? (Union College's archival manager)

The Spreadsheet checks are either remedied by the program or highlighted in the metadata spreadsheet for manual consideration. The Preliminary QC function auto passes/fails items and highlights them also for manual review.

The user has the option to personalize the highlighting colors in the UI for each error type since there will likely be highlighting already on the sheet by the time it makes it to QC. The program is designed to be run recursively. If a user remedies errors, their highlighting or auto Pass/Fails will be removed

## Developments

The object file structure lends itself well to a more complete quality control process where the QCer (Quality Controller) would not have to interact with the file structure at all. This is a potential future development and would improve productivity beyond the measures this program already introduces. This functionality is present with a button on the UI but is currently non-functional.

## Testing

Given the limited output of the digitization department at Union College we have only been able to test the program on a limited number of records (circa 500.) If you discover any issues please add these to the Issues section in github and we will work to resolve it.