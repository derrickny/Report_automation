import pandas as pd
import numpy as np

# List of all employees
employees = ['Rachael Mashao', 'Vanessa Odundo', 'Bilal Hasan', 'Munene Naftaly', 'Peter Omanya', 'Perpetual Kung`u', 'Joan Khabugwi', 'Lucy Mukasia', 'Fridah Murugi', 'Ann Ngige', 'Beril Achieng', 'Betty Kimani', 'Florence Njunguna', 'Rita Wangechi', 'Margaret Wanjiku']

# Load the data
data = pd.read_excel('/Users/nyagaderrick/Developer/Report_automation/JAN ATT REPORT.xlsx')

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
            check_in_time = group.loc[0, 'Time'].time() if len(group) > 0 else '_'
            check_out_time = group.loc[len(group) - 1, 'Time'].time() if len(group) > 1 else '_'
        else:
            check_in_time = '_'
            check_out_time = '_'
        formatted_date = date.strftime('%A ,%d/ %m / %Y') if not date_written else ''
        row = pd.DataFrame({'Date': [formatted_date], 'Name': [employee], 'CHECK IN TIME': [check_in_time], 'CHECK OUT TIME': [check_out_time]})
        final_data = pd.concat([final_data, row], ignore_index=True)
        date_written = True
    # Add a blank row after each date
    blank_row = pd.DataFrame([['', '', '', '']], columns=['Date', 'Name', 'CHECK IN TIME', 'CHECK OUT TIME'])
    final_data = pd.concat([final_data, blank_row], ignore_index=True)

# Save the final report
final_data.to_excel('/Users/nyagaderrick/Developer/Report_automation/JAN 2024 REPORT.xlsx', index=False)