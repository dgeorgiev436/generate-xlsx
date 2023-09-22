import json
import pandas as pd
from datetime import datetime

# Specify the JSON file path
json_file_path = 'PreRegisterData.json'

# Read JSON data from the file
with open(json_file_path, 'r', encoding='utf-8') as json_file:
    data = json.load(json_file)

# Convert milliseconds timestamp to readable date
for item in data:
    timestamp_ms = item["createdAt"]
    timestamp_seconds = timestamp_ms / 1000.0  # Convert to seconds
    item["createdAt"] = datetime.utcfromtimestamp(timestamp_seconds).strftime('%Y-%m-%d %H:%M:%S')

# Create a DataFrame from the JSON data
df = pd.DataFrame(data)

# Specify the Excel file name
excel_filename = 'pre-registrations.xlsx'

# Save the DataFrame to an Excel file
df.to_excel(excel_filename, index=False)

print(f"Data has been converted and saved to {excel_filename}.")
