import pandas as pd

# Load the Excel files
dell1 = pd.read_excel('Dell1.xlsx')
dell2 = pd.read_excel('Dell2.xlsx')

# Print the column names to verify
print("Columns in Dell1.xlsx:", dell1.columns)
print("Columns in Dell2.xlsx:", dell2.columns)

# Remove any trailing spaces from column names
dell1.columns = dell1.columns.str.strip()
dell2.columns = dell2.columns.str.strip()

# Extract the package names and statuses
dell1['Package Name'] = dell1['Package Name'].str.extract(r'([^/]+)$')
dell2['Package Name'] = dell2['Package Name'].str.extract(r'([^/]+)$')

# Create a dictionary from dell1 for quick lookup
status_dict = dict(zip(dell1['Package Name'], dell1['Column C Status']))

# Update the status in dell2 based on matching package names
dell2['Column C Status'] = dell2['Package Name'].map(status_dict).fillna(dell2['Column C Status'])

# Save the updated dell2 to a new Excel file
dell2.to_excel('Updated_Dell2.xlsx', index=False)

print("The status in Dell2.xlsx has been updated based on matching package names from Dell1.xlsx and saved as Updated_Dell2.xlsx.")
