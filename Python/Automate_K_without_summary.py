import pandas as pd
import numpy as np
import datetime
import random
from pandas.tseries.offsets import CustomBusinessDay

def generate_attendance_records(employees, dates, data):
    rows = []
    for date in dates:
        for employee in employees:
            check_in_time = ''
            check_out_time = ''
            employee_data = data[(data['Name'] == employee) & (data['Date'] == date)]
            if not employee_data.empty:
                check_in_time = employee_data['Time'].min().strftime('%H:%M')
                check_out_time = employee_data['Time'].max().strftime('%H:%M')
            row = {'Date': date, 'Name': employee, 'CHECK IN TIME': check_in_time, 'CHECK OUT TIME': check_out_time}
            rows.append(row)
    return pd.DataFrame(rows)

employees = ['Rachael Mashao', 'Vanessa Odundo', 'Peter Omanyo', 'Perpetual Kung`u',
             'Joan Khabugwi', 'Ann Ngige', 'Betty Kimani', 'Florence Njunguna',
             'Lorna Kuria', 'Everlynn Nundu', 'Hans zablon']

data = pd.read_excel('data/April/2.xlsx', engine='openpyxl', sheet_name='Sheet2')
data['Time'] = pd.to_datetime(data['Time'])
data['Date'] = data['Time'].dt.date

weekmask = 'Mon Tue Wed Thu Fri Sat'
holidays = []
custom_bday = CustomBusinessDay(weekmask=weekmask, holidays=holidays)
dates = pd.date_range(start='2024-04-08', end='2024-04-09', freq=custom_bday).date
dates = [date for date in dates if date.weekday() < 6]

final_data = generate_attendance_records(employees, dates, data)

# Add additional rows
final_data = pd.concat([final_data,
                       pd.DataFrame([{'Date': '', 'Name': '', 'CHECK IN TIME': 'KENCOM FEB ATTENDANCE', 'CHECK OUT TIME': ''}]),
                       pd.DataFrame([{'Date': 'NAME', 'Name': 'CHECK-IN AFTER 7:30 A.M', 'CHECK IN TIME': 'NO CHECK IN/OUT RECORDS', 'CHECK OUT TIME': ''}])],
                      ignore_index=True)

final_data = final_data.sort_values(['Date', 'Name'])
final_data.to_excel('report/April/Kencom_Report.xlsx', index=False)