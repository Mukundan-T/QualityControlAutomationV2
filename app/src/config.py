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

plat = platform.platform()

if "macOS" in plat:
    PLATFORM = 'macOS'
    DEFAULT_INPUT_FILE = os.path.join(MAINROOT, 'spreadsheets/'+ input_filename)
elif 'Windows' in plat:
    PLATFORM = 'windows'
    DEFAULT_INPUT_FILE = os.path.join(MAINROOT, 'spreadsheets\\'+ input_filename)

