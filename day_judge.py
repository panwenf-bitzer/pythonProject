import datetime
import os
from dateutil.parser import parse
from Data_Fetch import data_fetch
from Assy_data_fetch import ASSY_Attendance_fetch
def day_judgement():
    path = r"\\10.15.8.42\TAData"
    file = os.listdir(path)
    delta_date = 1
    match = file[-delta_date][12:20]
    filter_date = parse(match)
    Last_date = filter_date - datetime.timedelta(days=1)
    Last_date_str = str(Last_date)[0:10]
    filter_date = parse(match)
    Current_date_str = str(filter_date)[0:10]
    # dayofweek = datetime.datetime.strptime(Current_date_str, '%Y-%m-%d').isoweekday()
    if True:
        Last_2date = filter_date - datetime.timedelta(days=2)
        Last_3date = filter_date - datetime.timedelta(days=3)
        Last_4date = filter_date - datetime.timedelta(days=4)
        Last_5date = filter_date - datetime.timedelta(days=5)
        Last_6date = filter_date - datetime.timedelta(days=6)
        Last_7date = filter_date - datetime.timedelta(days=7)
        Last_8date = filter_date - datetime.timedelta(days=8)
        Last_9date = filter_date - datetime.timedelta(days=9)
        Last_10date = filter_date - datetime.timedelta(days=10)
        Last_11date = filter_date - datetime.timedelta(days=11)
        Last_12date = filter_date - datetime.timedelta(days=12)
        Last_2date_str = str(Last_2date)[0:10]
        Last_3date_str = str(Last_3date)[0:10]
        Last_4date_str = str(Last_4date)[0:10]
        Last_5date_str = str(Last_5date)[0:10]
        Last_6date_str = str(Last_6date)[0:10]
        Last_7date_str = str(Last_7date)[0:10]
        Last_8date_str = str(Last_8date)[0:10]
        Last_9date_str = str(Last_9date)[0:10]
        Last_10date_str = str(Last_10date)[0:10]
        Last_11date_str = str(Last_11date)[0:10]
        Last_12date_str = str(Last_12date)[0:10]

        for i in [ASSY_Attendance_fetch(firstday=Current_date_str, secondday=Last_date_str),
        ASSY_Attendance_fetch(firstday=Last_date_str, secondday= Last_2date_str),
        ASSY_Attendance_fetch(firstday=Last_2date_str, secondday=Last_3date_str),
        ASSY_Attendance_fetch(firstday=Last_3date_str, secondday=Last_4date_str),
        ASSY_Attendance_fetch(firstday=Last_4date_str, secondday=Last_5date_str),
        ASSY_Attendance_fetch(firstday=Last_5date_str, secondday=Last_6date_str),
        ASSY_Attendance_fetch(firstday=Last_6date_str, secondday=Last_7date_str),
        ASSY_Attendance_fetch(firstday=Last_7date_str, secondday=Last_8date_str),
        ASSY_Attendance_fetch(firstday=Last_8date_str, secondday=Last_9date_str),
        ASSY_Attendance_fetch(firstday=Last_9date_str, secondday=Last_10date_str),
        ASSY_Attendance_fetch(firstday=Last_10date_str, secondday=Last_11date_str),
        ASSY_Attendance_fetch(firstday=Last_11date_str, secondday=Last_12date_str)]:
            return i

if __name__ == '__main__':
    day_judgement()

