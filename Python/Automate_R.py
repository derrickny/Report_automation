import pandas as pd
import numpy as np
import datetime

# List of all employees
employees = ['Moreen Naitore', 'Sammy Walumoli']

# Load the data
data = pd.read_excel('data/june/JUNE ATT EDITED..xlsx', engine='openpyxl')

# Convert 'Time' to datetime format
data['Time'] = pd.to_datetime(data['Time'])

# Sort by 'Name' and 'Time'
data = data.sort_values(['Name', 'Time'])

# Group by 'Name' and date
data['Date'] = data['Time'].dt.date
grouped = data.groupby(['Name', 'Date'])

# Create a new DataFrame to hold the final data
final_data = pd.DataFrame()

# Define holidays
holidays = [datetime.date(2024, 6, 1), datetime.date(2024, 6, 17)]

# Create a record for each employee for each day
for date in pd.date_range(start='2024-06-01', end='2024-06-30').date:
    # Skip Sundays
    if date.weekday() == 6:
        continue
    date_written = False
    for employee in employees:
        if date in holidays:
            row = pd.DataFrame({'Date': [date if not date_written else ''], 'Name': [employee], 'CHECK IN TIME': ['HOLIDAY'], 'CHECK OUT TIME': ['HOLIDAY']})
        elif (employee, date) in grouped.groups:
            group = grouped.get_group((employee, date)).reset_index(drop=True)
            check_in_time = group.loc[0, 'Time'].time().strftime('%H:%M') if len(group) > 0 else '_'
            check_out_time = group.loc[len(group) - 1, 'Time'].time().strftime('%H:%M') if len(group) > 1 else '_'
            row = pd.DataFrame({'Date': [date if not date_written else ''], 'Name': [employee], 'CHECK IN TIME': [check_in_time], 'CHECK OUT TIME': [check_out_time]})
        else:
            row = pd.DataFrame({'Date': [date if not date_written else ''], 'Name': [employee], 'CHECK IN TIME': ['_'], 'CHECK OUT TIME': ['_']})
        final_data = pd.concat([final_data, row], ignore_index=True)
        date_written = True
    # Add a blank row after each date
    blank_row = pd.DataFrame([['', '', '', '']], columns=['Date', 'Name', 'CHECK IN TIME', 'CHECK OUT TIME'])
    final_data = pd.concat([final_data, blank_row], ignore_index=True)

# Add a summary after the last day
summary = final_data.copy()

# Exclude holidays from summary calculation
summary = summary[~summary['CHECK IN TIME'].isin(['HOLIDAY'])]

# Create a mask for rows that contain time data
mask = summary['CHECK IN TIME'].str.contains(':', na=False)

# Convert '08:30' to a string
time_to_compare = '08:30'

# Now you can do the comparison
summary.loc[mask, 'Late'] = summary.loc[mask, 'CHECK IN TIME'] > time_to_compare

# Count rows with '_' as 'NO CHECK IN/OUT RECORDS'
summary['No Record'] = summary['CHECK IN TIME'] == '_'

# Group by 'Name' and sum 'Late' and 'No Record'
summary = summary.groupby('Name').agg({'Late': 'sum', 'No Record': 'sum'}).reset_index()

# Rename columns
summary.columns = ['Name', 'Late', 'No Record']

# Add a row with the heading 'RUARAKA JUNE ATTENDANCE' before the summary
heading = pd.DataFrame([['RUARAKA JUNE ATTENDANCE', '', '', '']], columns=['Date', 'Name', 'CHECK IN TIME', 'CHECK OUT TIME'])
final_data = pd.concat([final_data, heading], ignore_index=True)

# Add a row with 'NAME', 'CHECK-IN AFTER 8:30 A.M', 'NO CHECK IN/OUT RECORDS'
column_names = pd.DataFrame([['NAME', 'CHECK-IN AFTER 8:30 A.M', 'NO CHECK IN/OUT RECORDS', '']], columns=['Date', 'Name', 'CHECK IN TIME', 'CHECK OUT TIME'])
final_data = pd.concat([final_data, column_names], ignore_index=True)

# Concatenate final_data and summary
summary.columns = ['Name', 'CHECK IN TIME', 'CHECK OUT TIME']
final_data = pd.concat([final_data, summary], ignore_index=True)

# Save the final report as a CSV file
output_path = 'report/june/RUARAKA_report_June.xlsx'
final_data.to_excel(output_path, index=False)

print(f"Report saved to {output_path}")