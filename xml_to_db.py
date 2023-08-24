import xml.etree.ElementTree as ET
import sqlite3
import os
import shutil

full_query = """
select subject.name, teacher.name, class.name, `group`.name, classroom.name, card.period, card.days 
from card
join classroom on card.classroomid=classroom.id
join lesson on lesson.id=card.lessonid
join lesson_to_teacher on lesson.id=lesson_to_teacher.lessonid
join teacher on teacher.id=lesson_to_teacher.teacherid
join subject on subject.id=lesson.subjectid
join class on lesson.classid=class.id
join `group` on lesson.groupid=`group`.id
order by teacher.name
;
"""

def make_output(filename):
    shutil.copy("template.xlsx", "output.xlsx")


def create_cache(xml_in, cache_name="cache.db"):
    if os.path.exists(cache_name):
        os.remove(cache_name)
    tree = ET.parse(xml_in) 
    with sqlite3.connect(cache_name) as db:
        for e in tree.getroot():
            if e.tag=="cards":
                db.execute("CREATE TABLE card (lessonid, period, days, classroomid);")
                for t in e:
                    d = t.attrib
                    x = (d["lessonid"], d["period"], d["days"], d["classroomids"])
                    db.execute("INSERT INTO card VALUES (?, ?, ?, ?);", x)
            elif e.tag=="lessons":
                db.execute("CREATE TABLE lesson (id, subjectid, classid, groupid);")
                db.execute("CREATE TABLE lesson_to_teacher (lessonid, teacherid);")
                for t in e:
                    d = t.attrib
                    x = (d["id"], d["subjectid"], d["classids"], d["groupids"])
                    for teacher in d["teacherids"].split(","):
                        db.execute("INSERT INTO lesson_to_teacher VALUES (?, ?);", [d["id"], teacher])
                    db.execute("INSERT INTO lesson VALUES (?, ?, ?, ?);", x)
            elif e.tag=="classrooms":
                db.execute("CREATE TABLE classroom (id, name);")
                for t in e:
                    d = t.attrib
                    x = (d["id"], d["short"])
                    db.execute("INSERT INTO classroom VALUES (?, ?);", x)
            elif e.tag=="classes":
                db.execute("CREATE TABLE class (id, name);")
                for t in e:
                    d = t.attrib
                    x = (d["id"], d["short"])
                    db.execute("INSERT INTO class VALUES (?, ?);", x)
            elif e.tag=="subjects":
                db.execute("CREATE TABLE subject (id, name);")
                for t in e:
                    d = t.attrib
                    x = (d["id"], d["name"])
                    db.execute("INSERT INTO subject VALUES (?, ?);", x)
            elif e.tag=="teachers":
                db.execute("CREATE TABLE teacher (id, name);")
                for t in e:
                    d = t.attrib
                    x = (d["id"], d["name"])
                    db.execute("INSERT INTO teacher VALUES (?, ?);", x)
            elif e.tag=="groups":
                db.execute("CREATE TABLE `group` (id, classid, name);")
                for t in e:
                    d = t.attrib
                    x = (d["id"], d["classid"], d["name"])
                    db.execute("INSERT INTO `group` VALUES (?, ?, ?);", x)
