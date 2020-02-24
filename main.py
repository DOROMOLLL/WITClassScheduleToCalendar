import xlrd
from ics import Event,Calendar
from datetime import timedelta,datetime
import sys
import string
import re
import pytz

weeks, times, clist = [], [], []
lesson = []
availWeeks = []
col_start, col_end = 1, 7
row_start, row_end = 3, 7
order_start, order_end = 0, 0
sheet_index=0
filepath="/example/example.xls"
# filepath specified here
output_filepath="my.ics"
# output filepath specified here

def find_content(sh):
    weeksNames = ['星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期日']
    timesKeys = ['第', '节']
    global col_start, row_start

    for row in range(sh.nrows):
        for col in range(sh.ncols):
            myCell = sh.cell(row, col)
            for value in weeksNames:
                if value in myCell.value:
                    weeks.append([row, col])
                    row_start, row_end = row + 1, row + 5
            if all(elem in list(myCell.value) for elem in timesKeys):
                times.append([row, col, myCell.value])
                col_start, col_end = col + 1, col + 7


def input_sheet():
    global sheet_index
    while True:
        i = input("Input the sheet index number of the class schedule（default is 1 , \'quit\' to exit）: ")
        if i == "":
            sheet_index=0
            return 0
        if i.isnumeric():
            sheet_index=i-1
            return int(i-1)
        if i == "quit":
            sys.exit()
        else:
            print("wrong input, input should be integer")
            continue
        return 0


def input_semester_start_date():
    global semester_start_date
    while True:
        while True:
            i = input("Enter date of semester\'s first MONDAY in YYYY-MM-DD format, \'quit\' to exit. \n"
                      "Please make sure the date is MONDAY , otherwise serious problems may occur : ")
            try:
                if i == "quit":
                    sys.exit()
                else:
                    year, month, day = map(int, i.split('-'))
                    semester_start_date = datetime(year, month, day, 0, 0, 0,tzinfo=pytz.timezone('Asia/Shanghai'))
                return 1
            except:
                print("wrong input, input should be YYYY-MM-DD format")
                continue


def get_lessons(row, col):
    global lesson, order_start, order_end
    global availWeeks
    l = sh.cell(row, col).value
    lesson = l.split('\n')
    if len(lesson) > 1:
        # filter blank item
        match1 = re.search("\(*\)", lesson[2])
        # find subtitle of class
        if match1:
            lesson[1] = lesson[1] + lesson[2]
            lesson.pop(2)
            # remove duplicates
        lesson = list(filter(lambda x: x != "", lesson))
        print("found lesson info "+str(lesson))

        pattern2 = re.compile(r'\d\d?')
        w = pattern2.findall(lesson[2])
        pattern1 = re.compile(r'\d\d?')
        order = pattern1.findall(lesson[4])

        order_start, order_end = int(order[0]), int(order[1])
        if len(w) == 2:
            availWeeks = list(range(int(w[0]), int(w[1]) + 1))
        elif len(w) == 4:
            availWeeks = list(range(int(w[0]), int(w[1]) + 1)) + list(range(int(w[2]), int(w[3]) + 1))


def get_time_start(orderStart, week, dayInWeek):
    lesson_start_hour = {
        1: 8,
        3: 9,
        6: 14,
        8: 15,
        10: 19,
    }.get(orderStart)
    lesson_start_minute = {
        1: 0,
        3: 55,
        6: 0,
        8: 55,
        10: 0,
    }.get(orderStart)

    lesson_start_time = semester_start_date \
                        + timedelta(weeks=week - 1, days=(int(dayInWeek) - 1),
                                    hours=lesson_start_hour - semester_start_date.hour,
                                    minutes=lesson_start_minute - semester_start_date.minute,
                                    seconds=-semester_start_date.second,
                                    milliseconds=-semester_start_date.microsecond)


    return lesson_start_time


def get_time_end(timeStart, orderEnd):
    # lesson_end_hour = {
    #     2: 9,
    #     4: 11,
    #     5: 12,
    #     7: 15,
    #     9: 17,
    #     11: 20,
    #     12: 21,
    # }.get(orderEnd)
    # lesson_end_minute = {
    #     2: 45,
    #     4: 30,
    #     5: 20,
    #     6: 35,
    #     8: 30,
    #     11: 35,
    #     12: 25,
    # }.get(orderEnd)
    timeEnd = timeStart + timedelta(minutes={
        2: 95,
        4: 95,
        5: 145,
        7: 95,
        9: 95,
        11: 95,
        12: 145,
    }.get(orderEnd))
    timeEnd.replace(tzinfo=pytz.timezone('Asia/Shanghai'))
    print(timeEnd)
    return timeEnd


xlrd.Book.encoding = "utf-8"
rf = xlrd.open_workbook(filepath)
sheet = rf.sheet_by_index(input_sheet())
nrows = sheet.nrows
ncols = sheet.ncols
input_semester_start_date()
for sh in rf.sheets():
    find_content(sh)
tz = pytz.timezone('Asia/Shanghai')

c = Calendar()
for row in range(row_start, row_end):
    for col in range(col_start, col_end):
        get_lessons(row, col)
        if len(rf.sheet_by_index(sheet_index).cell(row, col).value) > 1:
            print("start adding "+str(availWeeks)+" of "+lesson[0])
            for w in availWeeks:
                print("add week " + str(w)+" of lesson "+lesson[0])

                e = Event()
                e.name = str(lesson[0] + " " + lesson[1])
                e.location = lesson[3]

                print(tz)
                e.begin = get_time_start(order_start, w, col)
                e.end= get_time_end(e.begin, order_end)

                e.location = lesson[3]
                c.events.add(e)
                c.events
                print("added week " + str(w)+" of lesson "+lesson[0])

        else:
            print("skip null lesson")

with open(output_filepath, 'w') as f:
    f.write(str(c) + "\n")