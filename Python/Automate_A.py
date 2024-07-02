import pandas as pd
import numpy as np
import datetime

# Load the data and skip the first row
file_path = 'report/March/ADAMS_report.xlsx'
data = pd.read_excel(file_path, engine='openpyxl')

# Rename columns to standard names for easier processing
data.columns = ['Date', 'Name', 'CHECK IN TIME', 'CHECK OUT TIME'] + list(data.columns[4:])

# Define holidays
holidays = [datetime.date(2024, 6, 1), datetime.date(2024, 6, 17)]

# Convert 'Date' to datetime format and keep 'Clock In'/'Clock Out' as is
data['Date'] = pd.to_datetime(data['Date'], errors='coerce').dt.date

# Extract day from the original dates, handle NaN values
data['Day'] = data['Date'].apply(lambda x: x.day if pd.notna(x) else np.nan)

# Create new Date column for June 2024 using the extracted days
data['New Date'] = data['Day'].apply(lambda x: datetime.date(2024, 6, int(x)) if pd.notna(x) and x <= 30 else None)

# Replace 1st and 17th with 'HOLIDAY'
data.loc[data['New Date'].isin(holidays), ['CHECK IN TIME', 'CHECK OUT TIME']] = 'HOLIDAY'

# Update the Date column to the new dates in June
data['Date'] = data['New Date']

# Drop the helper columns
data.drop(columns=['Day', 'New Date'], inplace=True)

# Reorder the columns to place Date at the beginning
columns_order = ['Date', 'Name', 'CHECK IN TIME', 'CHECK OUT TIME'] + list(data.columns[4:])
data = data[columns_order]

# Save the final report as an Excel file
output_path = 'report/june/ADAMS_JUNE_REPORT_PROCESSED.xlsx'
data.to_excel(output_path, index=False)

print(f"Report saved to {output_path}")