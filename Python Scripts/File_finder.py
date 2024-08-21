"""
Authored by James Gaskell

08/20/2024

Edited by:

"""

import glob, easygui
import Excel_reader_writer

"""Opens an EasyGUI window to allow the user to select the file they want to parse
Returns:
    String: filepath of the selected file
"""
def get_file():
        path = easygui.fileopenbox()
        return path

def find_folder():
        return True

