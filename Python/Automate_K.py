import pandas as pd
import numpy as np
import datetime
from pandas.tseries.offsets import CustomBusinessDay

def generate_attendance_records(employees, dates, grouped):
    rows = []
    holidays = [datetime.date(2024, 6, 1), datetime.date(2024, 6, 17)]
    for date in dates:
        for employee in employees:
            if date in holidays:
                row = {'Date': date, 'Name': employee, 'CHECK IN TIME': 'HOLIDAY', 'CHECK OUT TIME': 'HOLIDAY'}
            elif (employee, date) in grouped.groups:
                group = grouped.get_group((employee, date)).reset_index(drop=True)
                check_in_time = group.loc[0, 'Time'].time().strftime('%H:%M') if len(group) > 0 else '_'
                check_out_time = group.loc[len(group) - 1, 'Time'].time().strftime('%H:%M') if len(group) > 1 else '_'
                row = {'Date': date, 'Name': employee, 'CHECK IN TIME': check_in_time, 'CHECK OUT TIME': check_out_time}
            else:
                row = {'Date': date, 'Name': employee, 'CHECK IN TIME': '_', 'CHECK OUT TIME': '_'}
            rows.append(row)
    return pd.DataFrame(rows)

# List of all employees
employees = ['Rachael Mashao', 'Vanessa Odundo', 'Munene Naftaly', 'Peter Omanyo', 'Perpetual Kung`u',
             'Joan Khabugwi', 'Ann Ngige', 'Betty Kimani', 'Florence Njunguna', 'Lorna Kuria', 'Hans Zablon']

# Load the data
data = pd.read_csv('data/june/report_JUNE.csv')

# Convert 'Time' to datetime format
data['Time'] = pd.to_datetime(data['Time'])

# Sort by 'Name' and 'Time'
data = data.sort_values(['Name', 'Time'])

# Group by 'Name' and date
data['Date'] = data['Time'].dt.date
grouped = data.groupby(['Name', 'Date'])

# Define a custom business day frequency that excludes only Sundays
weekmask = 'Mon Tue Wed Thu Fri Sat'
holidays = [datetime.date(2024, 6, 1), datetime.date(2024, 6, 17)]
custom_bday = CustomBusinessDay(weekmask=weekmask, holidays=holidays)

# Create a record for each employee for each custom business day
dates = pd.date_range(start='2024-06-01', end='2024-06-30', freq=custom_bday).date

# Manually filter out Sundays
dates = [date for date in dates if date.weekday() < 6]

final_data = generate_attendance_records(employees, dates, grouped)

# Add a summary after the 29th day
summary = final_data.copy()

# Remove holidays from summary calculation
summary = summary[~summary['Date'].isin(holidays)]

# Initialize mask as an empty Series
mask = pd.Series(dtype=bool)

if 'CHECK IN TIME' in summary.columns:
    mask = summary['CHECK IN TIME'].str.contains(':', na=False)
else:
    print("Column 'CHECK IN TIME' does not exist in the DataFrame.")

# Convert '07:30' to a string
time_to_compare = '07:30'

# Now you can do the comparison
summary.loc[mask, 'Late'] = summary.loc[mask, 'CHECK IN TIME'] > time_to_compare

# Count rows with '_' or NaN as 'NO CHECK IN/OUT RECORDS'
summary['No Record'] = summary['CHECK IN TIME'] == '_'

# Group by 'Name' and sum 'Late' and 'No Record'
summary = summary.groupby('Name').agg({'Late': 'sum', 'No Record': 'sum'}).reset_index()

# Rename columns
summary.columns = ['Name', 'Late', 'No Record']

# Define a function to add a row to the DataFrame
def add_row(data, row_values):
    row_df = pd.DataFrame([row_values], columns=['Date', 'Name', 'CHECK IN TIME', 'CHECK OUT TIME'])
    return pd.concat([data, row_df], ignore_index=True)

# Add a row with the heading 'KENCOM JUNE ATTENDANCE' before the summary
final_data = add_row(final_data, ['', '', 'KENCOM JUNE ATTENDANCE', ''])

# Add a row with 'NAME', 'CHECK-IN AFTER 7:30 A.M', 'NO CHECK IN/OUT RECORDS'
final_data = add_row(final_data, ['NAME', 'CHECK-IN AFTER 7:30 A.M', 'NO CHECK IN/OUT RECORDS', ''])

# Rename summary columns and concatenate it with final_data
summary.columns = ['Name', 'CHECK IN TIME', 'CHECK OUT TIME']
final_data = pd.concat([final_data, summary], ignore_index=True)

# Sort final_data by 'Date' and 'Name'
final_data = final_data.sort_values(['Date', 'Name'])

# Insert a blank line after each date
unique_dates = final_data['Date'].unique()
for date in unique_dates:
    idx = final_data[final_data['Date'] == date].index.max()
    final_data = pd.concat([final_data.loc[:idx], pd.DataFrame([{'Date': date, 'Name': np.nan, 'CHECK IN TIME': np.nan, 'CHECK OUT TIME': np.nan}]), final_data.loc[idx+1:]]).reset_index(drop=True)

# Set 'Date' to NaN for all rows except the first one for each unique date
final_data['Date'] = final_data['Date'].where(final_data['Date'].shift() != final_data['Date'])

# Save the final report
output_path = 'report/june/Kencom_Report_June.xlsx'
final_data.to_excel(output_path, index=False)

