class Item:
    student_class = ""
    teachers = []
    subject = ""
    group = 1
    day = 0
    auditory = ""
    lesson_number = ""

    def __init__(self, student_class, teacher, subject, group, day, auditory, lesson_number):
        self.student_class = student_class
        self.teachers = list(teacher.split(","))
        self.subject = subject
        self.group = group
        self.day = day
        self.auditory = auditory
        self.lesson_number = lesson_number