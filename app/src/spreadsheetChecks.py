"""
Authored by James Gaskell

1/24/2025

Edited by:

"""

from autocorrect import Speller
from dateutil.parser import parse
import datetime, string

"""Converts dates from year to year and day. E.g. 1980 to 1980-01-01
Args:
    date: date in year form
Returns:
    [Boolean, date]: success flag and date in ISO format
"""
def year_to_date(date):
    try:
        date_new = datetime.datetime(date, 1, 1)
        return [True, date_new]
    except:
        return [False, date]
    
"""
Attempts to format the date if it is in an unexpected string form
Args:
    date: date in string form
Returns:
    [Boolean, date]: success flag and date in ISO format
"""
    
def attempt_format(date):
    date_new = ""

    if date.rstrip()[-3] in string.punctuation:
        date_new = date.translate(str.maketrans('', '', string.punctuation))

        if date[-2:] == '00':
            date_new = date[:-2] + '01'
    
        try:
            date_new = parse(date_new)
            return [True, date_new]
        except:
            return [False, date]
        
    return [False, date]

"""Checks the date format and attempts to format the date if in unexpected form
Args:
    fileList: list of file objects from excel sheet
Returns:
    Boolean: True to verify the process was executed successfully
"""
def check_date_format(sheet):

    spell = Speller(lang="en")
    for file in sheet.fileList:
        success = False
        if file.date != None:
            if not type(file.date) is datetime.datetime:
                try:
                    date = (parse(file.date.rstrip()))
                    success = True
                except:
                    if type(file.date) is str and not success:
                        success, date = attempt_format((file.date))
                    elif type(file.date) is int and not success:
                        success, date = year_to_date((file.date))
                    if not success: #Last ditch effort, successful if incorrect spelling in date
                        try:
                            date = (parse(spell(file.date.rstrip())))
                            date = date.strftime("%Y-%m-%d")
                            file.date = date
                            success = True
                        except:
                            file.errors['Date'] = True
                            sheet.errors += 1
                    else:
                        file.date = date.strftime("%Y-%m-%d")
            else:
                file.date = file.date.strftime("%Y-%m-%d")
    return True


"""Checks for duplicate filenames in the list of files
Args:
    fileList: list of file objects from excel sheet
"""
def check_duplicate_filenames(sheet):
    for file in sheet.fileList:
        for comp_file in sheet.fileList:
            if file.fileName == comp_file.fileName and file != comp_file:
                file.errors['DupFilename'] = True
                sheet.errors += 1


"""Determines the correct prefix for the files by finding the most common from the file
    Means the script can be used for different collections
Args:
    List[ScanFile]: list of file objects from excel sheet
Returns:
    String: most likely prefix given all the entries
"""
def find_file_prefix(fileList):
    filenameDict = {} #Dictionary containing the prefixes in the document and their count
    for file in fileList:
        try: #ignore any funky filenames
            prefix = file.fileName.split('.')[0]
            if filenameDict.get(prefix) == None:
                filenameDict[prefix] = 1
            else:
                filenameDict[prefix] += 1
            return(max(filenameDict)) #Returns the most common prefix - assumes this is correct
        except:
            pass


def check_location_filename(sheet):
    prefix = find_file_prefix(sheet.fileList)
    for file in sheet.fileList:
        pred_filename = prefix
        if not file.location == None:

            #Does this filter work? I think so...
            Location = list(filter(None, file.location.translate(str.maketrans('', '', string.punctuation)).split(" ")))

            #ignore any funky filenames
            try:
                file.fileName = file.fileName.replace(" ", "") #Removes any spaces that shouldn't be in the filename
            except:
                pass

            if "Box" in Location:
                pred_filename += ".B" + Location[1].zfill(2)
            if "Folder" in Location:
                pred_filename += ".F" + Location[3].zfill(2)
            if "Bulletin" in Location: #Assumes .Bull. for bulletins
                pred_filename += ".Bull." + Location[5].zfill(2)
            elif "Sheet" in Location: #Assumes .Sheet. for sheets
                pred_filename += ".Sheet." + Location[5].zfill(2)   
            elif len(Location) >= 6: #Assumes no identifier for Items
                pred_filename += "." + Location[5].zfill(2)

            if pred_filename != file.fileName:
                file.errors['Filename'] = True
                sheet.errors += 1