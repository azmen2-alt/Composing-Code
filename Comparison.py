import pandas as pd

# --- SETUP: Update these file names if needed ---
sharepoint_file = "sharepoint_upload_ready.csv"
excel_file = "1.xlsx"
output_file = "Content_Mismatches.xlsx"
# -------------------------------------------------

# Load the files
file_sharepoint = pd.read_csv(sharepoint_file, low_memory=False)
file_1 = pd.read_excel(excel_file)

# Clean column names to make sure they match
file_sharepoint.columns = file_sharepoint.columns.str.replace('cree9_', '', regex=False).str.lower()
file_1.columns = file_1.columns.str.replace('cree9_', '', regex=False).str.lower()

# Clean the 'applicationid' column in both files
for df in [file_sharepoint, file_1]:
    if 'applicationid' in df.columns:
        df['applicationid'] = df['applicationid'].fillna('').astype(str).str.strip().str.lower()
        df['applicationid'] = df['applicationid'].str.replace(r'\.0$', '', regex=True)

# Find only the columns that exist in BOTH files
shared_cols = list(set(file_sharepoint.columns).intersection(set(file_1.columns)))
if 'applicationid' in shared_cols:
    shared_cols.remove('applicationid')

# Clean the data in the shared columns so we don't get false mismatch errors
for col in shared_cols:
    file_sharepoint[col] = file_sharepoint[col].fillna('').astype(str).str.strip().str.lower()
    file_sharepoint[col] = file_sharepoint[col].str.replace(r'\.0$', '', regex=True).replace('nan', '')
    
    file_1[col] = file_1[col].fillna('').astype(str).str.strip().str.lower()
    file_1[col] = file_1[col].str.replace(r'\.0$', '', regex=True).replace('nan', '')

# Remove duplicates and set applicationid as the index
f_sp = file_sharepoint[['applicationid'] + shared_cols].drop_duplicates(subset=['applicationid']).set_index('applicationid')
f_1 = file_1[['applicationid'] + shared_cols].drop_duplicates(subset=['applicationid']).set_index('applicationid')

# Find only the IDs that exist in BOTH files
shared_ids = f_sp.index.intersection(f_1.index)

# Filter both files to only include the shared IDs and shared columns
f_sp_shared = f_sp.loc[shared_ids]
f_1_shared = f_1.loc[shared_ids]

# Compare the cells to find exact differences
differences = f_sp_shared.compare(f_1_shared, align_axis=0).rename(index={'self': 'SharePoint File', 'other': 'File 1.xlsx'})

# Save the differences to an Excel file
with pd.ExcelWriter(output_file) as writer:
    if not differences.empty:
        differences.to_excel(writer, sheet_name='Cell Differences')
    else:
        # Create an empty sheet with a success message if everything matches perfectly
        pd.DataFrame({'Result': ['All overlapping cells match perfectly!']}).to_excel(writer, sheet_name='Perfect Match', index=False)

print("Comparison complete!")
print(f"- Overlapping rows checked: {len(shared_ids)}")
print(f"- Overlapping columns checked: {len(shared_cols)}")
print(f"- Rows with different cell data: {len(differences) // 2}")
print(f"\nYour Excel file is ready and saved as: {output_file}")