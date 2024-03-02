import pandas as pd
import numpy as np

# List of all employees
employees = ['Jackline Oira', 'Kennedy Mungai', 'Rhone Kwamboka', 'Laurraine Esisari', 'Humphrey Magawi',
             'Cecilia Mugure', 'Faith Chepkirui', 'Nancy Macharia', 'Hidaya Saidi', 'Grace Kagwe', 'Sammy Mukhwana',
             'Jimnah Njue',  'Bernice Kata']

# Load the data from the Excel file
df = pd.read_excel('/Users/nyagaderrick/Developer/Report_automation/January 2024 Attendance-USIU.xlsx')

# Create a new DataFrame to hold the final data
final_data = pd.DataFrame(columns=['Date', 'Name', 'Clock In', 'Clock Out'])

# For each unique date, filter the employees who worked that day and add a blank row after each date
for date in df['Date'].unique():
    date_data = df[df['Date'] == date]
    date_data = date_data[date_data['Name'].isin(employees)]
    # Add a blank row after each date
    blank_row = pd.DataFrame([['', '', '', '']], columns=['Date', 'Name', 'Clock In', 'Clock Out'])
    date_data = pd.concat([date_data, blank_row], ignore_index=True)
    final_data = pd.concat([final_data, date_data], ignore_index=True)

# Replace subsequent occurrences of each date with an empty string
final_data['Date'] = final_data['Date'].mask(final_data['Date'].duplicated(), '')

# Format the 'Date' column

final_data['Date'] = final_data['Date'].apply(lambda x: x if isinstance(x, str) else x.strftime('%d, %A, %B, %Y'))

# Create a new DataFrame for the summary report
summary_report = pd.DataFrame(columns=['Name', 'Clock In Before 8:00', 'No Clock In/Clock Out'])

# Calculate the required statistics for each employee
for employee in employees:
    employee_data = final_data[final_data['Name'] == employee]
    clock_in_between_8_and_11 = employee_data[(employee_data['Clock In'] >= '08:00') & (employee_data['Clock In'] <= '11:00')].shape[0]
    no_clock_in_or_out = employee_data[(employee_data['Clock In'].isna()) | (employee_data['Clock Out'].isna())].shape[0]
    new_row = pd.DataFrame([[employee, clock_in_between_8_and_11, no_clock_in_or_out]], columns=summary_report.columns)
    summary_report = pd.concat([summary_report, new_row], ignore_index=True)

# Add a new row with the heading 'USIU SUMMARY REPORT' in the 'Name' column
heading = pd.DataFrame([['', 'USIU SUMMARY REPORT', '', '']], columns=['Date', 'Name', 'Clock In', 'Clock Out'])
final_data = pd.concat([final_data, heading], ignore_index=True)

# Append the summary report to the main data
summary_report = summary_report.rename(columns={'Name': 'Date', 'Clock In Before 8:00': 'Name', 'No Clock In/Clock Out': 'Clock In'})
summary_report['Clock Out'] = ''
final_data = pd.concat([final_data, summary_report], ignore_index=True)

# Specify the path to the output Excel file
output_path = "/Users/nyagaderrick/Developer/Report_automation/JAN_2024_USIU_report.xlsx"

# Save the DataFrame to an Excel file
final_data.to_excel(output_path, index=False)