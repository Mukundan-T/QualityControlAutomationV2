import glob, easygui
from pathlib import Path

"""Opens an EasyGUI window to allow the user to select the file they want to parse
Returns:
    String: filepath of the selected file
"""
def get_file():
        path = easygui.diropenbox()
        return path

def find_file(folder):
        for path in Path(folder).rglob('*.pyc'):
                print(path.name)
        return True

