import xlwings as xw
import pandas as pd
import time,datetime
import numpy as np
import pymssql
import sqlalchemy
import os
from dateutil.parser import parse
import re
import sys
class data_fetch():
    def attendace_fetch(self,firstday,secondday):
        try:
            # 确定读取excell文件的搜索日期，锁定两天进行清洗
            path = r"\\10.15.8.42\TAData"
            file = os.listdir(path)
            Last_date_str = secondday
            Record_date = Last_date_str.replace("-", '')
            Search_date = firstday
            # 设置调试窗口的pandas显示最大的行数和列数
            pd.set_option('display.max_columns', None)
            pd.set_option('display.max_rows', None)
            # 读取HR存放的excell文件
            excell_file = r"\\10.15.8.42\TAData\\" + file[-1]
            # 启动xlwings读取excell文件
            app = xw.App(visible=False, add_book=False)
            app.display_alerts = False
            app.screen_updating = False
            wb = app.books.open(excell_file)
            data = app.books[0].sheets[0].range('B5:L8000').value
            # 将读取的数据dataframe化，并且将第一行设为表头
            pd_data = pd.DataFrame(data, columns=data[0])
            # 设定索引
            pd_data.set_index("Company ID", inplace=True)
            # 清洗空格及多余的项目
            del pd_data[None]
            del pd_data["Name"]
            del pd_data["Org. Unit"]
            del pd_data["Name of Organizational Unit"]
            del pd_data["Prev.PersNo."]
            pd_data.dropna(axis=0, inplace=True)
            pd_data = pd_data.iloc[1:]
            c = pd_data.iloc[pd_data.index.str.startswith("PD")]
            output = pd.DataFrame(c)
            output.drop(output.index[(output["Time Event Type"] == 'Clock-in')])
            output["Attendance_time"] = 0
            output["State"] = 0
            # 存储并关闭excell
            output.to_excel('attendance.xlsx')
            wb.close()
            app.quit()
            # 再次读取excell
            new_attendance = pd.read_excel("attendance.xlsx", index_col=[1], parse_dates=[4],
                                           date_parser=lambda x: pd.to_datetime(x, format='%Y-%m-%d'))
            new_attendance_dataframe = pd.DataFrame(new_attendance)
            filter_attendance = new_attendance_dataframe[
                (new_attendance_dataframe["Log.date"].astype(str) == Search_date) | (
                            new_attendance_dataframe["Log.date"].astype(str) == Last_date_str)]
            counter = list(filter_attendance["Log.date"].astype(str)).count(Last_date_str)
            print(counter)
            # if counter == 0:
            #     for i in file:
            #         file_name = path + "\\" + i
            #         os.remove(file_name)
            #     sys.exit()
            # 再次清洗数据
            filter_attendance.reset_index(inplace=True)
            filter_attendance.drop(filter_attendance.index[(filter_attendance["Time Event Type"] == 'Clock-in') & (
                        filter_attendance["Log.date"].astype(str) == Search_date)], inplace=True)
            filter_attendance.drop(filter_attendance.index[filter_attendance["Company ID"] == 'PD-LOGISTIC'],
                                   inplace=True)
            filter_attendance.set_index(["Personnel No."], inplace=True)

            # 遍历所有的元素并开始计算出勤时间
            start_time = 0
            stop_time = 0
            for index, row in filter_attendance.iterrows():
                if row["Time Event Type"] == 'Clock-in':
                    start_time = row["Time"]
                    if row.Time <= 0.261 and row.Time > 0.23:
                        # new2.loc[new2["Time"]==row["Time"],"Time"]=0.271
                        start_time = 0.261
                    if row.Time <= 0.594 and row.Time > 0.563:
                        start_time = 0.594
                    if row.Time <= 0.927 and row.Time > 0.896:
                        start_time = 0.927
                if row["Time Event Type"] == 'Clock-out':
                    stop_time = row["Time"]
                    if row.Time <= 0.311 and row.Time >= 0.271:
                        stop_time = 0.271
                    if row.Time <= 0.644 and row.Time >= 0.604:
                        stop_time = 0.604
                    if row.Time <= 0.978 and row.Time >= 0.938:
                        stop_time = 0.938
                    attendance = stop_time - start_time
                    if attendance < 0:
                        attendance = attendance + 1
                    filter_attendance.loc[[index], "Attendance_time"] = round(attendance * 24, 2)
            # 考勤异常写入
            filter_attendance['State'] = filter_attendance["Attendance_time"].apply(
                lambda x: "abnormal" if x != 8 else "normal")

            # 去重操作
            tem_attendance_dataframe = filter_attendance.groupby(filter_attendance.index).first()

            final_attendance_dataframe = pd.DataFrame(tem_attendance_dataframe)
            final_attendance_dataframe.reset_index(drop=False, inplace=True)
            # 重新设定列名
            final_attendance_dataframe.columns = ['PersonnelNo', 'ProductionLine', "PersonnelName", 'TimeEventType',
                                                  "LogDate", "LogTime", "AttendanceTime", "State"]
            # 数据汇总
            final_attendance_dataframe.to_excel("attendace2.xlsx")

            return [final_attendance_dataframe, Record_date,counter]
        except Exception:
            pass

