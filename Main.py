from openpyxl import load_workbook
from Item import *
from TimeTable import *
import xlsxwriter

# from xls2xlsx import *
# x2x = HTMLXLS2XLSX("Classes_Summaryhtml.html")
# x2x.to_xlsx("spreadsheet.xlsx")

from html2excel import ExcelParser

input_file = 'Classes_Summary.xls'
output_file = 'spreadsheet.xlsx'

parser = ExcelParser(input_file)
parser.to_excel(output_file)


def box(work_book, work_sheet, first_row, first_col, rows_count, cols_count):
    # top left corner
    work_sheet.conditional_format(first_row, first_col,
                                  first_row, first_col,
                                  {'type': 'formula', 'criteria': 'True',
                                   'format': work_book.add_format({'top': 2, 'left': 2})})
    # top right corner
    work_sheet.conditional_format(first_row, first_col + cols_count - 1,
                                  first_row, first_col + cols_count - 1,
                                  {'type': 'formula', 'criteria': 'True',
                                   'format': work_book.add_format({'top': 2, 'right': 2})})
    # bottom left corner
    work_sheet.conditional_format(first_row + rows_count - 1, first_col,
                                  first_row + rows_count - 1, first_col,
                                  {'type': 'formula', 'criteria': 'True',
                                   'format': work_book.add_format({'bottom': 2, 'left': 2})})
    # bottom right corner
    work_sheet.conditional_format(first_row + rows_count - 1, first_col + cols_count - 1,
                                  first_row + rows_count - 1, first_col + cols_count - 1,
                                  {'type': 'formula', 'criteria': 'True',
                                   'format': work_book.add_format({'bottom': 2, 'right': 2})})

    # top
    work_sheet.conditional_format(first_row, first_col + 1,
                                  first_row, first_col + cols_count - 2,
                                  {'type': 'formula', 'criteria': 'True', 'format': work_book.add_format({'top': 2})})
    # left
    work_sheet.conditional_format(first_row + 1, first_col,
                                  first_row + rows_count - 2, first_col,
                                  {'type': 'formula', 'criteria': 'True', 'format': work_book.add_format({'left': 2})})
    # bottom
    work_sheet.conditional_format(first_row + rows_count - 1, first_col + 1,
                                  first_row + rows_count - 1, first_col + cols_count - 2,
                                  {'type': 'formula', 'criteria': 'True',
                                   'format': work_book.add_format({'bottom': 2})})
    # right
    work_sheet.conditional_format(first_row + 1, first_col + cols_count - 1,
                                  first_row + rows_count - 2, first_col + cols_count - 1,
                                  {'type': 'formula', 'criteria': 'True', 'format': work_book.add_format({'right': 2})})


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
                if cell is not None:
                    lesson.append(cell)
            if len(lesson) == 4:
                items.append(Item(student_class=name, teacher=lesson[1], subject=lesson[0], group=1, day=day,
                                  auditory=lesson[3], lesson_number=lesson_number))
            elif len(lesson) == 10:
                items.append(Item(student_class=name, teacher=lesson[1], subject=lesson[0], group=1, day=day,
                                  auditory=lesson[3], lesson_number=lesson_number))
                items.append(Item(student_class=name, teacher=lesson[6], subject=lesson[5], group=2, day=day,
                                  auditory=lesson[8], lesson_number=lesson_number))

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

students_worksheet.set_column(start_x - 1, start_x - 1, 12)
students_worksheet.set_column(start_x - 2, start_x - 2, 2)

ind = 0
for st_class in timetable.classes:
    for day in range(6):
        box(workbook, students_worksheet, start_y + day * lessons, start_x + ind * width, lessons, width)
        for lesson in timetable.classes_timetable[st_class][day]:
            students_worksheet.write(start_y + day * lessons + lesson.number, start_x + ind * width + lesson.group * 2,
                                     lesson.name)
            students_worksheet.write(start_y + day * lessons + lesson.number,
                                     start_x + ind * width + lesson.group * 2 + 1, lesson.auditory)
        # Add classes names in header
        students_worksheet.merge_range(start_y - 1, start_x + ind * width, start_y - 1, start_x + ind * width + 3,
                                       st_class, merge_format)
    ind += 1

for day in range(0, 6):
    for lesson in range(lessons):
        box(workbook, students_worksheet, start_y + day * lessons, start_x - 2, lessons, 2)
        students_worksheet.write(start_y + day * lessons + lesson, start_x - 2, str(lesson))
        students_worksheet.write(start_y + day * lessons + lesson, start_x - 1, getLessonsTime(lesson))
    students_worksheet.merge_range(start_y + day * lessons, start_x - 3, start_y + day * lessons + lessons - 1,
                                   start_x - 3, getDay(day), merge_format)

for item in items:
    for teacher in item.teachers:
        if teacher not in timetable.teachers:
            timetable.teachers.append(teacher)
            timetable.teachers_timetable[teacher] = []
        lesson = TeacherLesson(item.student_class, item.lesson_number, item.day)
        timetable.teachers_timetable[teacher].append(lesson)

teachers_worksheet = workbook.add_worksheet("teachers")
teachers_worksheet.set_column(0, 0, 18)
teachers_worksheet.set_column(1, 50, 5)

start_x = 1
start_y = 3
lessons = 8

ind = 0
for teacher in sorted(timetable.teachers):
    teachers_worksheet.write(start_y + ind, start_x - 1, teacher)
    for lesson in timetable.teachers_timetable[teacher]:
        class_name = lesson.class_name
        if class_name == "10-8интернат":
            class_name = "10-8"
        teachers_worksheet.write(start_y + ind, start_x + lesson.day * lessons + lesson.number, class_name)
    ind += 1

for day in range(6):
    box(workbook, teachers_worksheet, start_y, start_x + day * lessons, len(timetable.teachers), lessons)
    box(workbook, teachers_worksheet, start_y - 2, start_x + day * lessons, 2, lessons)
    teachers_worksheet.merge_range(start_y - 2, start_x + day * lessons, start_y - 2,
                                   start_x + day * lessons + lessons - 1, getDayFull(day), merge_format)
    for lesson in range(lessons):
        teachers_worksheet.write(start_y - 1, start_x + day * lessons + lesson, lesson)

print("done!")
workbook.close()
