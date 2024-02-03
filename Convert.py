from tabula.io import read_pdf
import pandas as pd

# List of PDF files
pdf_files = [
    "/Users/nyagaderrick/Developer/Report_automation/report._adams.pdf",
    "/Users/nyagaderrick/Developer/Report_automation/JAN ATTE_kencom.pdf",
    "/Users/nyagaderrick/Developer/Report_automation/ATT JAN 24_ruaraka.pdf",
]

# Convert each PDF to Excel
for pdf_file in pdf_files:
    # Read PDF file
    tables = read_pdf(pdf_file, pages="all")

    # Concatenate all tables into a single DataFrame
    all_tables = pd.concat(tables, ignore_index=True)

    # Write the DataFrame to an Excel file
    all_tables.to_excel(pdf_file.replace(".pdf", ".xlsx"), index=False)