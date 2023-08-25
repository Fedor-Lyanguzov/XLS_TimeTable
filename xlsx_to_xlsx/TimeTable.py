from dataclasses import dataclass


class TimeTable:
    def __init__(self):
        self.classes = []
        self.classes_timetable = {}
        self.teachers = []
        self.teachers_timetable = {}
        self.classrooms = []
        self.classrooms_timetable = {}


@dataclass
class TeacherLesson:
    class_name: str = "0-0"
    number: int = 0
    day: int = 0
    classroom: str = "0"


@dataclass
class Lesson:
    name: str = "lesson name"
    number: int = 0
    group: int = 0
    auditory: str = ""


def getDay(day):
    if day == 0:
        return "Понед."
    if day == 1:
        return "Вторн."
    if day == 2:
        return "Среда"
    if day == 3:
        return "Четв."
    if day == 4:
        return "Пятн."
    if day == 5:
        return "Субб."


def getDayFull(day):
    if day == 0:
        return "Понедельник"
    if day == 1:
        return "Вторник"
    if day == 2:
        return "Среда"
    if day == 3:
        return "Четверг"
    if day == 4:
        return "Пятница"
    if day == 5:
        return "Суббота"


def getLessonsTime(lesson):
    if lesson == 0:
        return "08:20 - 09:05"
    if lesson == 1:
        return "09:15 - 10:00"
    if lesson == 2:
        return "10:10 - 10:55"
    if lesson == 3:
        return "11:10 - 11:55"
    if lesson == 4:
        return "12:10 - 12:55"
    if lesson == 5:
        return "13:25 - 14:10"
    if lesson == 6:
        return "14:20 - 15:05"
    if lesson == 7:
        return "15:15 - 16:00"
