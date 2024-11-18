import pandas as pd

# Load the Excel file
file_path = 'D:\crosspoint\working file 5336 Digikey.xlsx'  # Replace with your file path
sheet_name = 'My Lists Worksheet'  # Replace with your sheet name if needed
df = pd.read_excel(file_path, sheet_name=sheet_name)
df['Price'] = df[['Availability', 'Requested Quantity 1']].min(axis=1) * df['Unit Price 1']
# Display the entire DataFrame without truncation
for value in df['Price']:
    print(value)
