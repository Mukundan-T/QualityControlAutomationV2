# Refactoring development notes

* Should 'undated', 'none' etc be valid for the date column? In which case I can attempt formatting and leave all else - i.e. don't have to highlight errors if it won't lead to an upload error

* Could really do with a cleanup of the date checks. The filename checks are now a lot more pallateable, date checks need to be just as readable

* What is the best way to handle file output? We need to reference the objects to see if they have errors or fails but want to do this as efficiently as possible.

    1. We could request the failures of each type from the file objects using getters
        + If we could make this a dictionary of its own then it becomes more efficient to reference dictionary keys
        + nested dictionary? - split the method into a sheet error getter, use this with an excelfile to make a dictionary of sheets with a dictionary of errors inside?
        - Adds complexity to the objects
    2. We could 
