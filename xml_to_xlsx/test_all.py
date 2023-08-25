import pytest
import tkinter as tk
from xml_to_db import *


def test_convert_xml_to_db():
    create_cache("xml_to_xlsx/test_data/input.xml")


def test_output():
    make_output()


def test_shorten_subject():
    with sqlite3.connect(".cache.db") as db:
        print(list(db.execute("select name from subject")))
