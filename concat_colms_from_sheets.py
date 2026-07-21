from pathlib import Path
import pandas as pd

file_path = Path(r"C:\Users\hkildani002\Downloads\OneDrive_2026-07-21\210726_Consolidated_Catalog.xlsx")
output_file = Path(r"C:\Users\hkildani002\Downloads\OneDrive_2026-07-21\210726_Consolidated_Catalog3.xlsx")

cols_to_keep = ["Entity","#","Attribute ID (English)","Attribute ID (Arabic)","Attribute Description (English)","Attribute Description (Arabic)","Data Asset ID","Created Date","Last Update Date","Nullable (Y/N)","Data Type","Data Length","CDE Flag (Y/N)","Primary Key (Y/N)","Foreign Key Data Asset","Data Classification","Is Calculated? (Y/N)","Calculation Logic","Business Validation Logic","Data Completeness","Data Completeness Target","Missing Data Handling (Y/N)","Missing Data Handling Process","Duplicate Data Handling (Y/N)","Duplicate Data Handling Process","Duplicate (Y/N)","Data Quality Issues","Uniqueness Target","Data Glossary","Data Owner","Data Steward"]
#use this excel formula to concat the column names =""""&TEXTJOIN(""",""",TRUE,A1:E1)&""""
all_sheets = pd.read_excel(file_path, sheet_name=None)

consolidated = []

for sheet_name, df in all_sheets.items():
    # Clean column names
    df.columns = df.columns.astype(str).str.strip()

    # Identify which required columns exist in this sheet
    present_cols = [col for col in cols_to_keep if col in df.columns]

    # Skip only if none of the requested columns exist
    if not present_cols:
        continue

    # Keep present columns only
    temp = df[present_cols].copy()

    # Add missing columns as blank
    for col in cols_to_keep:
        if col not in temp.columns:
            temp[col] = ""

    # Reorder columns consistently
    temp = temp[cols_to_keep]

    # Optional traceability
    temp["Source Sheet"] = sheet_name

    consolidated.append(temp)

final_df = pd.concat(consolidated, ignore_index=True)

with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    final_df.to_excel(writer, sheet_name="Consolidated", index=False)

print(f"Created consolidated file: {output_file}")