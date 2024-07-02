import pandas as pd
import numpy as np
import datetime

# Load the data and skip the first row
file_path = 'data/june/USIU_JUNE_24_REPORT.xlsx'
data = pd.read_excel(file_path, engine='openpyxl')

# Rename columns to standard names for easier processing
data.columns = ['Date', 'Name', 'CHECK IN TIME', 'CHECK OUT TIME'] + list(data.columns[4:])

# Define holidays
holidays = [datetime.date(2024, 6, 1), datetime.date(2024, 6, 17)]

# Convert 'Date' and 'Clock In'/'Clock Out' to datetime format
data['Date'] = pd.to_datetime(data['Date']).dt.date
data['CHECK IN TIME'] = pd.to_datetime(data['CHECK IN TIME'], format='%H:%M:%S', errors='coerce').dt.time
data['CHECK OUT TIME'] = pd.to_datetime(data['CHECK OUT TIME'], format='%H:%M:%S', errors='coerce').dt.time

# Initialize a list to hold processed records
processed_records = []

# Get a list of all unique employees
employees = data['Name'].unique()

# Create a record for each employee for each day in June, marking holidays and missing records
for date in pd.date_range(start='2024-06-01', end='2024-06-30').date:
    if date.weekday() == 6:  # Skip Sundays
        continue
    for employee in employees:
        if date in holidays:
            processed_records.append({'Date': date, 'Name': employee, 'CHECK IN TIME': 'HOLIDAY', 'CHECK OUT TIME': 'HOLIDAY'})
        else:
            employee_records = data[(data['Name'] == employee) & (data['Date'] == date)]
            if not employee_records.empty:
                check_in_time = employee_records['CHECK IN TIME'].min().strftime('%H:%M') if pd.notna(employee_records['CHECK IN TIME'].min()) else '_'
                check_out_time = employee_records['CHECK OUT TIME'].max().strftime('%H:%M') if pd.notna(employee_records['CHECK OUT TIME'].max()) else '_'
                processed_records.append({'Date': date, 'Name': employee, 'CHECK IN TIME': check_in_time, 'CHECK OUT TIME': check_out_time})
            else:
                processed_records.append({'Date': date, 'Name': employee, 'CHECK IN TIME': '_', 'CHECK OUT TIME': '_'})

# Convert processed records to a DataFrame
processed_df = pd.DataFrame(processed_records)

# Add summary for missing check-ins and check-outs
summary_df = processed_df[~processed_df['CHECK IN TIME'].isin(['HOLIDAY'])].copy()
summary_df['Late'] = summary_df['CHECK IN TIME'].apply(lambda x: x > '08:30' if x != '_' else False)
summary_df['No Record'] = summary_df['CHECK IN TIME'] == '_'

summary = summary_df.groupby('Name').agg({'Late': 'sum', 'No Record': 'sum'}).reset_index()
summary.columns = ['Name', 'Late', 'No Record']

# Add summary to the processed DataFrame
processed_df = pd.concat([processed_df, pd.DataFrame({'Date': [''], 'Name': [''], 'CHECK IN TIME': ['USIU JUNE ATTENDANCE'], 'CHECK OUT TIME': ['']})], ignore_index=True)
processed_df = pd.concat([processed_df, pd.DataFrame({'Date': [''], 'Name': ['NAME'], 'CHECK IN TIME': ['CHECK-IN AFTER 8:30 A.M'], 'CHECK OUT TIME': ['NO CHECK IN/OUT RECORDS']})], ignore_index=True)
summary['CHECK IN TIME'] = summary['Late']
summary['CHECK OUT TIME'] = summary['No Record']
processed_df = pd.concat([processed_df, summary[['Name', 'CHECK IN TIME', 'CHECK OUT TIME']]], ignore_index=True)

# Save the final report as an Excel file
output_path = 'report/june/USIU_JUNE_24_REPORT_PROCESSED.xlsx'
processed_df.to_excel(output_path, index=False)

print(f"Report saved to {output_path}")