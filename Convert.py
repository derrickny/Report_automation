from tabula.io import read_pdf
import pandas as pd

# List of PDF files
pdf_files = [
 "data/ATT FEB 24.pdf"
]

# Convert each PDF to Excel
for pdf_file in pdf_files:
    # Read PDF file
    tables = read_pdf(pdf_file, pages="all")

    # Concatenate all tables into a single DataFrame
    all_tables = pd.concat(tables, ignore_index=True)

    # Write the DataFrame to an Excel file
    all_tables.to_excel(pdf_file.replace(".pdf", ".xlsx"), index=False)
    

    