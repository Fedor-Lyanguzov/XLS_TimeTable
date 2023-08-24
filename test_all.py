import pytest
import tkinter as tk
from Main import *
from xml_to_db import *


@pytest.mark.skip()
def test_import_timetable():
    timetable = import_timetable("./test_data/have.xlsx")
    # print(timetable.__dict__)


@pytest.mark.skip()
def test_import_old_input():
    inp, out = tk.StringVar(None, "./test_data/old_input.xlsx"), tk.StringVar(
        None, "./test_data/output.xlsx"
    )
    # timetable = import_timetable("./test_data/old_input.xlsx")
    start(inp, out)


def test_convert_xml_to_db():
    create_cache("./test_data/input.xml")


def test_output():
    make_output()


def test_shorten_subject():
    with sqlite3.connect(".cache.db") as db:
        print(list(db.execute("select name from subject")))
