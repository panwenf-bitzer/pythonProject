import pymssql
import pandas as pd
from Data_Fetch import data_fetch
import datetime
import sqlalchemy
import os,sys
def ASSY_Attendance_fetch(firstday,secondday):
    try:
        conn = pymssql.connect(server="10.15.1.199", user="sa", password="bitzer,.123", database="Tadata")
        sql_ASS = "select * from ASS"
        sql_CNC = "select * from CNC"
        sql_maintenance= "select * from PDMT"
        pd.set_option('display.max_columns', None)
        pd.set_option('display.max_rows', None)
        df0 = pd.read_sql(sql_ASS, conn)
        df1 = pd.read_sql(sql_CNC, conn)
        df2 = pd.read_sql(sql_maintenance, conn)
        path = r"\\10.15.8.42\TAData"
        file = os.listdir(path)

        attendace_Data_list = data_fetch().attendace_fetch(firstday=firstday, secondday=secondday)
        attendace_Data = attendace_Data_list[0]

        df0["date"] = attendace_Data_list[1]
        df1["date"] = attendace_Data_list[1]
        df2["date"] = attendace_Data_list[1]

        ShAssembly = 0
        ScAssembly = 0
        AluMskAssembly = 0
        ScrAssembly = 0
        Assembly = 0
        Maintenance=0
        ALL = 0
        CNC = 0
        Date = attendace_Data_list[1]

        for index1, row1 in df0.iterrows():
            for index2, row2 in attendace_Data.iterrows():
                if row1["Pers#No#"] == row2["PersonnelNo"]:
                    df0.loc[[index1], "AttendanceTime"] = row2["AttendanceTime"]
        df0.to_excel("attendance_assy.xlsx")
        df0.dropna(inplace=True)
# assembly calculation
        for index, row in df0.iterrows():
            Assembly = Assembly + row["AttendanceTime"]
            if row["Cost Center"] in [40602100, 40602110, 40602120, 40602130, 40602140]:
                AluMskAssembly = AluMskAssembly + row["AttendanceTime"]
            if row["Cost Center"] in [40602200, 40602210, 40602220, 40602230, 40602240, 40602250]:
                ScAssembly = ScAssembly + row["AttendanceTime"]
            if row["Cost Center"] in [40602300, 40602310, 40602320, 40602330, 40602340, 40602150]:
                ShAssembly = ShAssembly + row["AttendanceTime"]
            if row["Cost Center"] in [40602400, 40602410, 40602420, 40602430, 40602440, 40602450]:
                ScrAssembly = ScrAssembly + row["AttendanceTime"]
        df0.drop("Description", axis=1, inplace=True)
        df0.columns = ["Pers_No", "Name", "CC", "Description", "Date_id", "Attendance_Time", ]
        df0.to_excel("assy_database.xlsx")
        print(ScAssembly, ShAssembly, ScrAssembly, AluMskAssembly, Assembly)
# cnc calculation
        for index1, row1 in df1.iterrows():
            for index2, row2 in attendace_Data.iterrows():
                if row1["Pers#No#"] == row2["PersonnelNo"]:
                    df1.loc[[index1], "AttendanceTime"] = row2["AttendanceTime"]
        df1.to_excel("attendance_cnc.xlsx")
        df1.dropna(inplace=True)

        for index, row in df1.iterrows():
            if row["Cost Center"] in [40602010, 40602040, 40602020, 40602030, 40602050]:
                CNC = CNC + row["AttendanceTime"]

# maintenance calculation
        for index3, row3 in df2.iterrows():
            for index4, row4 in attendace_Data.iterrows():
                if row3["Pers#No#"] == row4["PersonnelNo"]:
                    df2.loc[[index3], "AttendanceTime"] = row4["AttendanceTime"]
        df2.dropna(inplace=True)

        for index, row in df2.iterrows():
            if row["Cost Center"] in [40602910]:
                Maintenance+=row["AttendanceTime"]
        df2.columns = ["CC","Description","Name","Pers_No","Date_id","Attendance_Time",]
        df2.to_excel("maintenance.xlsx")
        print(Maintenance)

        Total_attendance = [round(ShAssembly, 2), round(ScAssembly, 2), round(ScrAssembly, 2), round(AluMskAssembly, 2),
                            round(Assembly, 2), round(CNC, 2), round(Assembly + CNC+Maintenance, 2),Date,round(Maintenance,2),]
        df1.drop("Description", axis=1, inplace=True)
        df1.columns = ["Pers_No", "Name", "CC", "Description", "Date_id", "Attendance_Time", ]
        df1.to_excel("CNC_database.xlsx")
        print(Total_attendance)
        sum_column = ["ShAssembly", "ScAssembly", "ScrAssembly", "AluMskAssembly", "Assembly", "Cnc", "All", "Date_id","Maintenance"]

        attendace_dic = dict(zip(sum_column, Total_attendance))
        final_attendance_dataframe = pd.DataFrame(data=attendace_dic, index=[0])
        if attendace_Data_list[2]!=0:
            engine1 = sqlalchemy.create_engine(r'mssql+pymssql://s00015:Start123!@DW-SQL\DW/ThroughPutTime?charset=utf8', )
            final_attendance_dataframe.to_sql(name='Data_aquisition_productionlineattendancefact', con=engine1,
                                          if_exists='append', index=False)
            df0.to_sql(name='Data_aquisition_dailyattendanceassfact', con=engine1, if_exists='append', index=False)
            df1.to_sql(name='Data_aquisition_dailyattendancecncfact', con=engine1, if_exists='append', index=False)
            df2.to_sql(name='Data_aquisition_dailyattendancemaintenancefact', con=engine1, if_exists='append',index=False)

            engine2 = sqlalchemy.create_engine(
            r'mssql+pymssql://sa:bitzer,.123@CNBJS16226/KPI_visualization?charset=utf8', )
            final_attendance_dataframe.to_sql(name='Data_aquisition_attendancefact', con=engine2, if_exists='append',
                                          index=False)
            df0.to_sql(name='Data_aquisition_dailyattendanceassfact', con=engine2, if_exists='append', index=False)
            df1.to_sql(name='Data_aquisition_dailyattendancecncfact', con=engine2, if_exists='append', index=False)
            df2.to_sql(name='Data_aquisition_dailyattendancemaintenancefact', con=engine2, if_exists='append',index=False)
    except Exception:
        pass


if __name__ == '__main__':
    ASSY_Attendance_fetch()