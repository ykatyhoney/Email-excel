import pandas as pd
import re

# Replace 'your_file.xlsx' with the path to your Excel file
file_path = 'sample1.xlsx'

# Load the Excel file
xls = pd.ExcelFile(file_path)

# Iterate over each sheet
for sheet_name in xls.sheet_names:
    data = pd.read_excel(xls, sheet_name)
    print(f"Sheet: {sheet_name}")
    # Iterate over each column and identify potential email addresses
    for column in data.columns:
        potential_emails = []
        for cell in data[column]:
            if re.match(r"[^@]+@[^@]+\.[^@]+", str(cell)):
                potential_emails.append(cell)
        if potential_emails:
            print(potential_emails)
