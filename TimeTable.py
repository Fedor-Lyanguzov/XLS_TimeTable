class TimeTable:
    classes = []
    classes_timetable = {}

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