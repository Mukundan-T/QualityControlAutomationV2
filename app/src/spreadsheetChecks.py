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
    
def attempt_format(date):
    date_new = ""

    if date.rstrip()[-3] in string.punctuation:
        date_new = date.translate(str.maketrans('', '', string.punctuation))

        if date[-2:] == '00':
            date_new = date[:-2] + '01'
    
        try:
            date_new = parse(date)
            return [True, date_new]
        except:
            return [False, date]
        
    return [False, date]

"""Checks the date format and attempts to format the date if in unexpected form
Args:
    files_list: list of file objects from excel sheet
Returns:
    Boolean: True to verify the process was executed successfully
"""
def check_date_format(files_list):
    spell = Speller(lang="en")
    for file in files_list:
        success = False
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
                else:
                    file.date = date.strftime("%Y-%m-%d")
        else:
            file.date = file.date.strftime("%Y-%m-%d")
    return True