import pandas as pd
import numpy as np
import datetime

# List of all employees
employees = ['Moreen Naitore', 'Sammy Walumoli']

# Load the data
data = pd.read_excel('/Users/nyagaderrick/Developer/Report_automation/ATT JAN 24_ruaraka.xlsx')

# Convert 'Time' to datetime format
data['Time'] = pd.to_datetime(data['Time'])

# Sort by 'Name' and 'Time'
data = data.sort_values(['Name', 'Time'])

# Group by 'Name' and date
data['Date'] = data['Time'].dt.date
grouped = data.groupby(['Name', 'Date'])

# Create a new DataFrame to hold the final data
final_data = pd.DataFrame()

# Create a record for each employee for each day
for date in pd.date_range(start='2024-01-01', end='2024-01-31').date:
    # Skip Sundays
    if date.weekday() == 6:
        continue
    date_written = False
    for employee in employees:
        if (employee, date) in grouped.groups:
            group = grouped.get_group((employee, date)).reset_index(drop=True)
            check_in_time = group.loc[0, 'Time'].time().strftime('%H:%M') if len(group) > 0 else '_'
            check_out_time = group.loc[len(group) - 1, 'Time'].time().strftime('%H:%M') if len(group) > 1 else '_'
        else:
            check_in_time = '_'
            check_out_time = '_'
        row = pd.DataFrame({'Date': [date if not date_written else ''], 'Name': [employee], 'CHECK IN TIME': [check_in_time], 'CHECK OUT TIME': [check_out_time]})
        final_data = pd.concat([final_data, row], ignore_index=True)
        date_written = True
    # Add a blank row after each date
    blank_row = pd.DataFrame([['', '', '', '']], columns=['Date', 'Name', 'CHECK IN TIME', 'CHECK OUT TIME'])
    final_data = pd.concat([final_data, blank_row], ignore_index=True)

# Add a summary after the 31st day
summary = final_data.copy()

# Create a mask for rows that contain time data
mask = summary['CHECK IN TIME'].str.contains(':', na=False)

# Convert '07:30' to a string
time_to_compare = '07:30'

# Now you can do the comparison
summary.loc[mask, 'Late'] = summary.loc[mask, 'CHECK IN TIME'] > time_to_compare

# Count rows with '_' or NaN as 'NO CHECK IN/OUT RECORDS'
summary['No Record'] = summary['CHECK IN TIME'] == '_'

# Group by 'Name' and sum 'Late' and 'No Record'
summary = summary.groupby('Name').agg({'Late': 'sum', 'No Record': 'sum'}).reset_index()

# Rename columns
summary.columns = ['Name', 'CHECK IN TIME', 'CHECK OUT TIME']

# Convert 'Date' column to datetime format only for rows that have a date
# Add a row with the heading 'KENCOM DECEMBER ATTENDANCE' before the summary
heading = pd.DataFrame([['RUARAKA JANUARY SUMMARY', '', '', '']], columns=['Date', 'Name', 'CHECK IN TIME', 'CHECK OUT TIME'])
final_data = pd.concat([final_data, heading], ignore_index=True)

# Add a row with 'NAME', 'CHECK-IN AFTER 7:30 A.M', 'NO CHECK IN/OUT RECORDS'
column_names = pd.DataFrame([['NAME', 'CHECK-IN AFTER 7:30 A.M', 'NO CHECK IN/OUT RECORDS', '']], columns=['Date', 'Name', 'CHECK IN TIME', 'CHECK OUT TIME'])
final_data = pd.concat([final_data, column_names], ignore_index=True)

# Concatenate final_data and summary
summary.columns = ['Name', 'CHECK IN TIME', 'CHECK OUT TIME']
final_data = pd.concat([final_data, summary], ignore_index=True)

# Save the final report
final_data.to_excel('/Users/nyagaderrick/Developer/Report_automation/JAN_2024_RUARAKA_report.xlsx', index=False)