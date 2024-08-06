import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter

# Define the paths
input_file_path = "C:\\Users\\adiaz\\Downloads\\SP24_rawSleepFormatted.xlsx"
consent_file_path = "C:\\Users\\adiaz\\Downloads\\SP24_Consent_SleepMatch.xlsx"
output_file_path = "C:\\Users\\adiaz\\Downloads\\SP24_rawSleepFormatted(Analysis).xlsx"

# Load the consent data
consent_df = pd.read_excel(consent_file_path)
consented_ids = consent_df[consent_df.iloc[:, 6].astype(str).str.upper() != 'FALSE'].iloc[:, 0].astype(str)
tier_info = consent_df[consent_df.iloc[:, 6] != 'FALSE'].iloc[:, [0, 3]]
tier_info.columns = ['Participant ID', 'Intervention Group']
tier_info['Participant ID'] = tier_info['Participant ID'].astype(str)

# Functions to find first and last consecutive days with valid data
def find_consecutive_days(data, num_days=5, first=True):
    valid_indices = data.dropna().where(data != 0).dropna().index.tolist()
    if len(valid_indices) < num_days:
        return None
    if first:
        for i in range(len(valid_indices) - num_days + 1):
            if all(data.index[j] - data.index[i] == j - i for j in range(i, i + num_days)):
                return data.iloc[i:i + num_days]
    else:
        for i in range(len(valid_indices) - num_days, -1, -1):
            if all(data.index[j] - data.index[i] == j - i for j in range(i, i + num_days)):
                return data.iloc[i:i + num_days]
    return None

# Initialize ExcelWriter
with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
    # Load the Excel file
    xlsx = pd.ExcelFile(input_file_path)
    for sheet_name in xlsx.sheet_names:
        df = pd.read_excel(xlsx, sheet_name=sheet_name)
        df = df.iloc[:, 2:]  # Adjust DataFrame
        df.columns = df.columns.astype(str)  # Ensure columns are strings
        df = df.filter(items=consented_ids.tolist())  # Filter by consented IDs

        # Prepare results for this sheet
        results = { 'Participant ID': df.columns }
        for participant_id in df.columns:
            column_data = df[participant_id].dropna()
            recorded_data = column_data[column_data != 0]
            
            first_valid_days = find_consecutive_days(recorded_data, first=True)
            last_valid_days = find_consecutive_days(recorded_data, first=False)

            results.setdefault('First 5 Consecutive Day Mean', []).append(first_valid_days.mean() if first_valid_days is not None else None)
            results.setdefault('First 5 Consecutive Day Std Dev', []).append(first_valid_days.std() if first_valid_days is not None else None)
            results.setdefault('Last Consecutive 5 day mean', []).append(last_valid_days.mean() if last_valid_days is not None else None)
            results.setdefault('Last Consecutive 5 day Std Dev', []).append(last_valid_days.std() if last_valid_days is not None else None)
            results.setdefault('#DaysOfRecordedData', []).append(len(recorded_data))

        # Create a DataFrame and write to the corresponding sheet
        new_df = pd.DataFrame(results)
        new_df = new_df.merge(tier_info, on='Participant ID', how='left')
        new_df.to_excel(writer, sheet_name=sheet_name, index=False)

        # Optionally adjust column widths
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        for col in worksheet.columns:
            max_length = max((len(str(cell.value)) if cell.value is not None else 0) for cell in col)
            adjusted_width = (max_length + 2)
            worksheet.column_dimensions[get_column_letter(col[0].column)].width = adjusted_width

print("New file created with comprehensive participant data:", output_file_path)
