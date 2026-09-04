# Composing-Code
Compare two excel or CSV files
Here is a clear and simple README file you can use for GitHub.

---

# Data Comparison Tool

This Python script compares data between a SharePoint CSV export and an Excel file. It checks overlapping columns and rows to find exact cell differences.

## What It Does

* Matches rows using the `applicationid` column.
* Only compares columns and rows that exist in both files.
* Ignores extra or missing rows and columns.
* Outputs a new Excel file highlighting any mismatched data.

## Requirements

You need Python installed on your computer. You also need two Python libraries: `pandas` and `openpyxl`.

To install the required libraries, open your terminal or command prompt and run:

```bash
pip install pandas openpyxl

```

## How to Use

1. Place the Python script, your SharePoint CSV file (`sharepoint_upload_ready.csv`), and your Excel file (`1.xlsx`) in the exact same folder.
2. Open your terminal or command prompt and navigate to that folder.
3. Run the script:

```bash
python script_name.py

```

*(Note: If your file names are different, open the Python script and change the names in the "SETUP" section at the top).*

## Output

The script creates a new file called `Content_Mismatches.xlsx` in the same folder.

* If there are differences, it will list the exact cell changes, showing the SharePoint value next to the Excel value.
* If everything matches perfectly, it will create a sheet saying "All overlapping cells match perfectly!".
