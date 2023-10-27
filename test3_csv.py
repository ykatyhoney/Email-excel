import pandas as pd
import re
from verify_email import verify_email

# Replace 'your_file.csv' with the path to your CSV file
file_path = 'sample1.csv'
output_file_path = 'output.xlsx'

# Load the CSV file, specifying all columns as string type to avoid DtypeWarning
data = pd.read_csv(file_path, dtype=str)

output_data = []
seen_emails = set()

# Iterate over each column
for column in data.columns:
    for cell in data[column]:
        if re.match(r"[^@]+@[^@]+\.[^@]+", str(cell)):
            if cell not in seen_emails:
                if verify_email(cell):
                    output_data.append({'Type': 'Email', 'Value': cell})
                    seen_emails.add(cell)

# Create a new Excel file with the output data
output_df = pd.DataFrame(output_data)
output_df.to_excel(output_file_path, index=False)
print(f"Data saved to {output_file_path}")
