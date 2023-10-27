import pandas as pd
import re

# Replace 'your_file.xlsx' with the path to your Excel file
file_path = 'sample1.csv'
output_file_path = 'output.xlsx'

# Load the Excel file
xls = pd.ExcelFile(file_path)
output_data = []
seen_emails = set()

# Iterate over each sheet
for sheet_name in xls.sheet_names:
    data = pd.read_excel(xls, sheet_name)
    for column in data.columns:
        for cell in data[column]:
            if re.match(r"[^@]+@[^@]+\.[^@]+", str(cell)) and cell not in seen_emails:
                output_data.append({'Sheet': sheet_name, 'Type': 'Email', 'Value': cell})
                seen_emails.add(cell)

# Create a new Excel file with the output data
output_df = pd.DataFrame(output_data)
output_df.to_excel(output_file_path, index=False)
print(f"Data saved to {output_file_path}")