"""
Authored by James Gaskell

08/19/2024

Edited by:

"""

import List_operations, string
import pandas as pd

"""Checks for duplicate filenames in the sheet
Args:
    Dataframe: current pandas dataframe being worked on
    Array[String]: ReportMajor containing duplicate filenames be exported to the report
    Array[String]: Problem_rows containing the dulicate filenames so they can be highlighted and added to the text report
Returns:
    Boolean: True or False may be used to verify the process was executed successfully - not yet implemented
"""
def check_duplicate_filenames(df, reportMajor, problem_rows):
    filenames = df["Filename"][df['Filename'].duplicated(keep=False)]
    for name in filenames:
        if name not in problem_rows:
            problem_rows.append(name)
            reportMajor.append("Filename: " + name +" is duplicated")
    return True


"""Determines the correct prefix for the files by finding the most common from the file
    Means the script can be used for different collections
Args:
    List[String]: filenames used to determine most likely correct prefix
Returns:
    String: most likely prefix given all the entries
"""
def find_file_prefix(filenames):
    filenameDict = {} #Dictionary containing the prefixes in the document and their count
    for name in filenames:
        prefix = name.split(".")[0]
        if filenameDict.get(prefix) == None:
            filenameDict[prefix] = 1
        else:
            filenameDict[prefix] += 1
    return(max(filenameDict)) #Returns the most common prefix - assumes this is correct


"""Splits the Physical location into Box, Folder and, if available, Item
    Uses the prefix to reconstruct the filename
    Searches the folder to locate the files by the determined filenames
Args:
    Dataframe: current pandas dataframe being worked on
    Array[String]: ReportMajor containing major errors to be exported to the report
    Array[String]: ReportNoLocation exports rows with no physical location to the Report in Minor section
    Array[String]: Problem_rows containing the filenames that do not match so they can be highlighted and added to the text report
Returns:
    Boolean: Used to determine if the process was successful or not
"""
def check_location_filename(df, reportMajor, reportNoLocation, problem_rows):
    locations = df[['aspace_id','Physical Location','Filename']].copy()
    prefix = find_file_prefix(locations['Filename'])
    for index, row in locations.iterrows():
        if not pd.isna(row['Physical Location']):
            Location = (row['Physical Location']).translate(str.maketrans('', '', string.punctuation))
            Location = Location.split(" ")
            Location = List_operations.removeBlanks(Location) #Removes spaces added at the end of entries

            if "Bulletin" in row['Physical Location']: #Assumes .Bull. for bulletins
                Type = "Bull."
            elif "Sheet" in row['Physical Location']: #Assumes .Sheet. for sheets
                Type = "Sheet."
            else:
                Type = "" #Assumes no type identifier for items

            detFilename = prefix
            detFilenameNoPadding = None
            Filename_no_suffix = None
            if len(Location) >= 2:
                detFilename += "."
                if Location[0][0] == "B":
                    detFilename += "B"
                detFilename += Location[1].zfill(2)
            if len(Location) >= 4:
                detFilename += "."
                if Location[2][0] == "F":
                    detFilename += "F"
                detFilename  += Location[3].zfill(2)
            if len(Location) >= 6:
                detFilenameNoPadding = detFilename + "." + Type + Location[5]
                detFilename += "." + Type + Location[5].zfill(2)

                # for test
            if (len(Location) > 6) and len(Location[6]) == 1 and Location[6].isalpha():
                try:
                    more_characters = Location[7]
                except IndexError:
                    detFilename += Location[6]
                    detFilenameNoPadding += Location[6]

            if row['Filename'].rstrip()[-1].isalpha():
                Filename_no_suffix = row['Filename'].rstrip()[:-1]

            if (detFilename != row['Filename'].rstrip() and detFilenameNoPadding != row['Filename'].rstrip() and detFilename != Filename_no_suffix and detFilenameNoPadding != Filename_no_suffix):
                reportMajor.append("Location and Filename do not match for File: " + row['Filename'])
                problem_rows.append(row['Filename']) #Rows with major problems identified for highlighting
        else:
            reportNoLocation.append("No Physical Location for File: " + row['Filename'])

    return True