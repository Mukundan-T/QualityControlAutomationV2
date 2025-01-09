"""Checks for duplicate filenames in the sheet
Args:
    files_list: list of file objects from excel sheet
Returns:
    Boolean: True or False may be used to verify the process was executed successfully
"""
def check_duplicate_filenames(files_list):
    for file in files_list:
        for comp_file in files_list:
            if file.filepath == comp_file.filepath and not "duplicate" in file.errors:
                file.errors += "duplicate"
    return True