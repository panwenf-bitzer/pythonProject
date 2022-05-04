# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
from day_judge import day_judgement
import sqlalchemy
import time,datetime
import pandas as pd
import os,sys
from Email import Email
def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    print_hi('PyCharm')
    search_flag=True
    email_flag = "start"
    current_date=datetime.datetime.now().date()
    dayofweek = datetime.datetime.strptime(str(current_date), '%Y-%m-%d').isoweekday()
    path = r"\\10.15.8.42\TAData"
    file = os.listdir(path)
    while search_flag:
        current_time = datetime.datetime.now()
        try:
            day_judgement()
            search_flag = False
            for i in file:
                file_name = path + "\\" + i
                os.remove(file_name)
                sys.exit()
        except IndexError:
            if dayofweek!=6 and dayofweek!=7 and current_time.hour==9 and current_time.minute==30 and email_flag=="start":
                message = "We don't receive your data untill now (9:30),kind remind you submit attendance sheet to server.\n the email generate automaticlly"
                email_flag=Email('pan.wenfang@bitzer.cn', 'Attendance submit remind', message)
                Email('zhu.zhengran@bitzer.cn', 'Attendance submit remind', message)
                Email('liang.yanbo@bitzer.cn', 'Attendance submit remind', message)
                Email('lv.na@bitzer.cn', 'Attendance submit remind', message)
                print(email_flag)
            pass
        except Exception:
            sys.exit()


# See PyCharm help at https://www.jetbrains.com/help/pycharm/