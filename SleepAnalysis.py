import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter

# Define the paths and sheet names
input_file_path = "C:\\Users\\adiaz\\Downloads\\SP24_rawSleepFormatted.xlsx"
consent_file_path = "C:\\Users\\adiaz\\Downloads\\SP24_Consent_SleepMatch.xlsx"
sheet_name = 'TotalMinutesAsleep'
output_file_path = "C:\\Users\\adiaz\\Downloads\\SP24_rawSleepFormatted(Analysis).xlsx"

# Load the specific sheets from the Excel files
df = pd.read_excel(input_file_path, sheet_name=sheet_name)
consent_df = pd.read_excel(consent_file_path)

# Filter out participants based on the consent
# Assuming the first row of the consent_df has column titles and the IDs are in the first column,
# and the consent status is in the seventh column.
consented_ids = consent_df[consent_df.iloc[:, 6] != 'FALSE'].iloc[:, 0].astype(str)

# Adjust DataFrame to only consider valid data columns, assuming first two columns are not participant data
df = df.iloc[:, 2:]

# Ensure the columns of df are strings to match consented_ids
df.columns = df.columns.astype(str)

# Filter the df to include only those columns (participant IDs) that are in consented_ids
df = df.filter(items=consented_ids.tolist())

# Prepare dictionaries to hold the means, standard deviations, and days of recorded data
first_5_day_means = {}
first_5_day_stds = {}
last_5_day_means = {}
last_5_day_stds = {}
days_of_recorded_data = []

# Iterate over each participant column to calculate the mean, std
for participant_id in df.columns:
    column_data = df[participant_id].dropna()
    recorded_data = column_data[column_data != 0]

    days_of_recorded_data.append(len(recorded_data))  # Count of non-zero, non-empty cells

    if not column_data.empty:
        first_valid_days = column_data.head(5)
        last_valid_days = column_data.tail(5)
        first_5_day_means[participant_id] = first_valid_days.mean() if len(first_valid_days) >= 5 else None
        first_5_day_stds[participant_id] = first_valid_days.std() if len(first_valid_days) >= 5 else None
        last_5_day_means[participant_id] = last_valid_days.mean() if len(last_valid_days) >= 5 else None
        last_5_day_stds[participant_id] = last_valid_days.std() if len(last_valid_days) >= 5 else None

# Create a new DataFrame with Participant IDs and their corresponding data
new_df = pd.DataFrame({
    'Participant ID': df.columns,
    '#DaysOfRecordedData': days_of_recorded_data,
    'First 5 Consecutive Day Mean': list(first_5_day_means.values()),
    'First 5 Consecutive Day Std Dev': list(first_5_day_stds.values()),
    'Last Consecutive 5 day mean': list(last_5_day_means.values()),
    'Last Consecutive 5 day Std Dev': list(last_5_day_stds.values())
})

# Save the results to a new Excel file
new_df.to_excel(output_file_path, index=False)

# Adjust column widths for better readability
def auto_adjust_columns(filepath):
    wb = openpyxl.load_workbook(filepath)
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        for col in ws.columns:
            max_length = max((len(str(cell.value)) if cell.value is not None else 0) for cell in col)
            adjusted_width = (max_length + 2)
            ws.column_dimensions[get_column_letter(col[0].column)].width = adjusted_width
    wb.save(filepath)

auto_adjust_columns(output_file_path)

print("New file created with comprehensive participant data:", output_file_path)
