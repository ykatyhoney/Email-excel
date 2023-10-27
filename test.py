import pandas as pd
import re

# Replace 'your_file.xlsx' with the path to your Excel file
file_path = 'sample1.xlsx'
output_file_path = 'output.xlsx'

# Load the Excel file
xls = pd.ExcelFile(file_path)
output_data = []

# Iterate over each sheet
for sheet_name in xls.sheet_names:
    data = pd.read_excel(xls, sheet_name)
    for column in data.columns:
        potential_emails = []
        for cell in data[column]:
            if re.match(r"[^@]+@[^@]+\.[^@]+", str(cell)):
                potential_emails.append(cell)
        if potential_emails:
            output_data.append({'Sheet': sheet_name, 'Column': column, 'Emails': ', '.join(potential_emails)})

# Create a new Excel file with the output data
output_df = pd.DataFrame(output_data)
output_df.to_excel(output_file_path, index=False)
print(f"Emails saved to {output_file_path}")