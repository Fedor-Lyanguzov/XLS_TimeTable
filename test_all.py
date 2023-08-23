from Main import *


def test_import_timetable():
    timetable = import_timetable("./test_data/have.xlsx")
    print(timetable.__dict__)
