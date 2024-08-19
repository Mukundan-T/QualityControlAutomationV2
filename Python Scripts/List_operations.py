"""
Authored by James Gaskell

08/19/2024

Edited by:

"""


"""Removes blank spaces added to spreadsheet during data entry
Args:
    List: String broken up into parts
Returns:
    List without blank spaces
"""
def removeBlanks(lst):
    while '' in lst:
        lst.remove('')
    return (lst)