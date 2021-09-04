from openpyxl import load_workbook
from Item import *
from TimeTable import *
import xlsxwriter


#  Import timetable
workbook = load_workbook('cl_sum.xlsx')
first_sheet = workbook.get_sheet_names()[0]
worksheet = workbook.get_sheet_by_name(first_sheet)

print(worksheet.cell(row=4, column=1).value)

items = []

# Parse input file
for student_class in range(5, 320, 10):
    name = worksheet.cell(row=student_class, column=1).value
    for day in range(0, 6):
        for lesson_number in range(8):
            lesson = []
            for param in range(10):
                cell = worksheet.cell(row=student_class + param, column=2 + day * 8 + lesson_number).value
                if cell != None:
                    lesson.append(cell)
            if len(lesson) == 4:
                items.append(Item(student_class=name, teacher=lesson[1], subject=lesson[0], group=1, day=day, auditory=lesson[3], lesson_number=lesson_number))
            elif len(lesson) == 10:
                items.append(Item(student_class=name, teacher=lesson[1], subject=lesson[0], group=1, day=day, auditory=lesson[3], lesson_number=lesson_number))
                items.append(Item(student_class=name, teacher=lesson[6], subject=lesson[5], group=2, day=day, auditory=lesson[8], lesson_number=lesson_number))

timetable = TimeTable()

# Create timetable
for item in items:
    if item.student_class not in timetable.classes:
        timetable.classes.append(item.student_class)
        timetable.classes_timetable[item.student_class] = []
    lesson = Lesson(item.subject, item.lesson_number, item.group - 1, item.auditory)
    if len(timetable.classes_timetable[item.student_class]) <= item.day:
        timetable.classes_timetable[item.student_class].append([lesson])
    else:
        timetable.classes_timetable[item.student_class][item.day].append(lesson)


# Create excel file

workbook = xlsxwriter.Workbook("output_timetable.xlsx")

students_worksheet = workbook.add_worksheet("students")
merge_format = workbook.add_format({
    'bold': 1,
    'border': 2,
    'align': 'center',
    'valign': 'vcenter'})

cell_format = workbook.add_format({
    'bold': 1,
    'border': 1
})

start_x = 4
start_y = 2
lessons = 8
width = 4

for y in range(start_y - 1, start_y + 7 * lessons):
    for x in range(start_x - 3, start_x + (len(timetable.classes) + 1) * width):
        students_worksheet.write(y, x, "", cell_format)

ind = 0
for st_class in timetable.classes:
    for day in range(6):
        for lesson in timetable.classes_timetable[st_class][day]:
            students_worksheet.write(start_y + day * lessons + lesson.number, start_x + ind * width + lesson.group * 2, lesson.name, cell_format)
            students_worksheet.write(start_y + day * lessons + lesson.number, start_x + ind * width + lesson.group * 2 + 1, lesson.auditory, cell_format)
        # Add classes names in header
        students_worksheet.merge_range(start_y - 1, start_x + ind * width, start_y - 1, start_x + ind * width + 3, st_class, merge_format)
    ind += 1

for day in range(0, 6):
    for lesson in range(lessons):
        students_worksheet.write(start_y + day * lessons + lesson, start_x - 2, str(lesson), cell_format)
    students_worksheet.merge_range(start_y + day * lessons, start_x - 3, start_y + day * lessons + lessons - 1, start_x - 3, getDay(day), merge_format)

print("done!")
workbook.close()