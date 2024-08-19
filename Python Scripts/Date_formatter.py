from dateutil.parser import parse
import datetime
import string
import pandas as pd
from autocorrect import Speller

def check_date_format(df, reportMajor, problem_rows):
    spell = Speller(lang="en")
    for index, row in df.iterrows():
        #print(row['date_created'])
        if not pd.isna(row['date_created']):
            if not type(row['date_created']) is datetime.datetime:
                try:
                    parse(row['date_created'].rstrip())
                    success = True
                except:
                    success = False
                    if type(row['date_created']) is str and not success:
                        success = attempt_format(row['date_created'])
                    elif type(row['date_created']) is int and not success:
                        success = year_to_date(row['date_created'])
                if not success:
                    try:
                        row['date_created'] = parse(spell(row['date_created']))
                    except:
                        problem_rows.append(row['Filename'])
                        reportMajor.append("The date for Filename: " + row['Filename'] + " is not in a readable format")

    return True

def attempt_format(df_date):
    success = False

    if df_date.rstrip()[-3] in string.punctuation:
        date = df_date.translate(str.maketrans('', '', string.punctuation))

        if date[-2:] == '00':
            date = date[:-2] + '01'
    
        try:
            date = parse(date)
            df_date = date
            success = True
        except:
            success = False

    return success


def year_to_date(df_date):
    date = df_date

    try:
        date = datetime.datetime(df_date, 1, 1)
        success = True
    except:
        success = False

    return success