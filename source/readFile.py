import pandas as pd

"""
    - how use pandas to read the data of each sheet
    - store the all the sheets data of the excel file in a structured manner

"""
# Update with your actual file path
EXCEL_FILE = "5.EastAfricaPipelinev0.5.xlsx"

# Load all sheet names
xls = pd.ExcelFile(EXCEL_FILE)

# Print sheet names
print("Available Sheets:")
for sheet in xls.sheet_names:
    print(f"- {sheet}")

# Print columns for each sheet
for sheet in xls.sheet_names:
    df = pd.read_excel(EXCEL_FILE, sheet_name=sheet)  # Read only first 5 rows to speed up
    print(f"\nSheet: {sheet}")
    print(df.columns.tolist())  # Print column names
