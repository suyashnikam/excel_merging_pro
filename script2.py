import pandas as pd

# Read sheet-A
sheet_A = pd.read_excel('sheet-A.xlsx')  # Adjust path and sheet name

# List of other sheets
sheet_names = ['sheet-B.xlsx', 'sheet-C.xlsx', 'sheet-E.xlsx']  # Adjust to your sheet names

# Clean up the 'standard' and 'standard number' columns to remove any leading/trailing spaces
sheet_A['standard'] = sheet_A['standard'].str.strip()
sheet_A['standard number'] = sheet_A['standard number'].str.strip()

merged_data = sheet_A  # Start with sheet-A as the base

# Loop through the other sheets to merge their data
for sheet_name in sheet_names:
    # Read the current sheet
    df = pd.read_excel(sheet_name)  # Adjust the file path and sheet name
    
    # Clean up the 'standard' and 'standard number' columns in the other sheets
    df['standard'] = df['standard'].str.strip()
    df['standard number'] = df['standard number'].str.strip()

    # Check if 'standard' and 'standard number' columns exist in both sheets
    if 'standard' in sheet_A.columns and 'standard number' in sheet_A.columns:
        if 'standard' in df.columns and 'standard number' in df.columns:
            # Drop the 'merge_status' column if it exists
            if 'merge_status' in merged_data.columns:
                merged_data = merged_data.drop(columns=['merge_status'])

            # Merge data on 'standard' and 'standard number' with a custom merge status column name
            merged_data = pd.merge(merged_data, df, on=['standard', 'standard number'], how='left', suffixes=('', f'_{sheet_name.split(".")[0]}'), indicator='merge_status')
        else:
            print(f"Columns 'standard' and 'standard number' not found in {sheet_name}")
    else:
        print("Columns 'standard' and 'standard number' not found in sheet-A")

# Drop the 'merge_status' column before saving
merged_data = merged_data.drop(columns=['merge_status'])

# Save the final merged data to a new file
merged_data.to_excel('final_merged_data.xlsx', index=False)

print("Merge complete. Final data saved to 'final_merged_data.xlsx'.")
