class TimeTable:
    classes = []
    classes_timetable = {}
    teachers = []
    teachers_timetable = {}
    classrooms = []
    classrooms_timetable = {}


class TeacherLesson:
    class_name = "0-0"
    number = 0
    day = 0
    classroom = "0"

    def __init__(self, class_name, number, day):
        self.class_name = class_name
        self.number = number
        self.day = day


class Lesson:
    name = "lesson name"
    group = 0
    number = 0
    auditory = ""

    def __init__(self, name, number, group, auditory):
        self.number = number
        self.name = name
        self.group = group
        self.auditory = auditory


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
