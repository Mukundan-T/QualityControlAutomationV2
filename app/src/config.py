"""
Authored by James Gaskell

11/06/2024

This is a simple check for platform. In the early stages the
program worked on both MacOS and Windows, it is not clear if
this is still the case or not. The sample.xlsx file can be used
for testing and must be placed in app/spreadsheets folder. When
testing DEAFAULT_ONEDRIVE_FOLDER should be set to the root of the
scanned documents or test.py will incorrectly indicate a 100%
failure rate.


Edited by:

"""

import os
import platform

ROOT = os.path.dirname(__file__)
MAINROOT = os.path.dirname(ROOT)
BATCH_SIZE = 10
APP_ROOT = os.path.join(ROOT)
PLATFORM = ""

input_filename = 'sample.xlsx'

# Accounting fot the differences in file structure between Mac and Windows
# Can be expanded for more operating systems.
# Currently works on Mac and windows - other functions may break if the program is ported to a different OS
# Needs calling by the main class or method to populate the PLATFORM constant
# This constant can then be used where there are interactions with the file structure

plat = platform.platform()

if "macOS" in plat:
    PLATFORM = 'macOS'
    DEFAULT_INPUT_FILE = os.path.join(MAINROOT, 'spreadsheets/'+ input_filename)
    DEFAULT_ONEDRIVE_FOLDER = None
elif 'Windows' in plat:
    PLATFORM = 'windows'
    DEFAULT_INPUT_FILE = os.path.join(MAINROOT, 'spreadsheets\\'+ input_filename)
    DEFAULT_ONEDRIVE_FOLDER = "C:\\Users\\libstudent.UNION\\OneDrive - Union College\\SCA-0319 William Stanley Jr"

