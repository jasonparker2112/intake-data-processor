import os
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
import re
import logging

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
log = logging.getLogger(__name__)

# Define valid labels and other necessary mappings
valid_labels = [
    "Families", "Individuals", "Refused", "Female H.H.", "Male", "Female", "Indiv. Disabled", "Veteran",
    "New To System", "Impact Avenues", "EAS", "Wellsky", "DCF/TANF", "Employment", "SSI/SSD", "VA Disability",
    "Social Security", "Unemp./Workers Comp.", "Pensions (VA, RR, ETC.)", "Child Support", "No Income", "Other Income",
    "Food Stamps", "WIC", "Kancare", "Health Access", "Budget Problems", "Car Repairs", "Check Late",
    "Crime Victim", "Domestic Problems", "Homeless/Evicted", "Illness/Death", "In Treatment", "Lack of Effort",
    "Lack of Work", "Laid Off/Fired", "Moving Expenses", "Released from Jail", "School", "DCF Pending",
    "SSI/SSD Pending", "Utilities High/Off", "Food", "Utilities", "Medical - RX", "Medical - Dental", "Rent",
    "Transportation - Bus", "Transportation - Gas", "TRSNPT - Auto Services", "Miscellaneous", "Education/Childcare",
    "Phone Bill", "Outreach", "Christmas", "Water Share", "Food (lbs)", "Under 5", "5 to 9", "10 to 14", "15 to 19",
    "20 to 24", "25 to 29", "30 to 34", "35 to 40", "41 to 44", "45 to 54", "55 to 64", "65 to 74", "Over 75",
    "At or Below Poverty", "Near Poverty", "Other Income Level", "Unknown Income Level", "Black", "White", "Hispanic", "Indian", "Other Race",
    "Homeless", "66602", "66603", "66604", "66605", "66606", "66607", "66608", "66609", "66610", "66611", "66612",
    "66613", "66614", "66615", "66616", "66617", "66618", "66619", "Outside of Topeka", "Unknown Zip"
]

def extract_date_from_filename(filename, sheet_name):
    match = re.search(r'Intake (\w+) (\d{4})', filename)
    if match:
        month_str, year = match.groups()
        try:
            date = datetime.strptime(f"{sheet_name} {month_str} {year}", "%d %b %Y").date()
            return date
        except ValueError:
            return None
    return None

def process_workbook(filepath):
    wb = load_workbook(filepath, data_only=True)
    file_name = os.path.basename(filepath)
    all_data = []

    for sheet_name in wb.sheetnames:
        if sheet_name.lower() in ["monthend", "compatibility report"]:
            continue
        ws = wb[sheet_name]
        date = extract_date_from_filename(file_name, sheet_name)
        if not date:
            continue

        # Find the column with "Total" in the second row
        total_col = None
        for col_idx, cell in enumerate(ws[2], start=1):  # Check the second row
            if cell.value and "total" in str(cell.value).lower():
                total_col = col_idx
                break

        if total_col is None:
            log.warning(f"⚠️ 'Total' column not found in sheet: {sheet_name}. Skipping this sheet.")
            continue

        # Process each row starting from the second row (skip header row)
        for row in ws.iter_rows(min_row=3, values_only=True):
            label = row[0]
            value = row[total_col - 1]  # Extract value from the "Total" column
            if label and str(label).strip() in valid_labels:
                all_data.append({
                    'label': str(label).strip(),
                    'value': value,
                    'sheet_date': date
                })

    return all_data

# List of directories to check
base_dirs = ['c:/Users/folderpath/']

# Load all Excel files and insert into DB
all_records = []

for dir_path in base_dirs:
    for file in os.listdir(dir_path):
        if file.endswith('.xlsx'):
            filepath = os.path.join(dir_path, file)
            log.info(f"Processing: {filepath}")
            records = process_workbook(filepath)
            all_records.extend(records)

# Convert to DataFrame
df_final = pd.DataFrame(all_records)
log.info(f"Final DataFrame: {df_final.head()}")

from sqlalchemy import create_engine

# Update with your PostgreSQL connection details
db_user = 'db_username'
db_password = 'db_password'
db_host = 'db_host'
db_port = 'db_port'
db_name = 'db_name'
table_name = 'db_table_name'  # <- Your target table name

# Create connection string
engine = create_engine(f'postgresql+psycopg2://{db_user}:{db_password}@{db_host}:{db_port}/{db_name}')

# Pivot the DataFrame: rows = date, columns = label, values = value
df_pivot = df_final.pivot_table(index='sheet_date', columns='label', values='value', aggfunc='sum').reset_index()

# Ensure the column order matches the valid labels (if there are any additional ones, they will be appended)
ordered_columns = ['sheet_date'] + [label for label in valid_labels if label in df_pivot.columns]
df_pivot = df_pivot[ordered_columns]

# Fill NaN values and ensure correct dtype inference
df_pivot = df_pivot.fillna(0).infer_objects()

# Send to PostgreSQL
df_pivot.to_sql(table_name, engine, if_exists='replace', index=False)

print(f"✅ Data successfully written to the '{table_name}' table.")
