from pathlib import Path
import re
import pandas as pd

input_folder = Path(r"C:\Users\hkildani002\Downloads\OneDrive_2026-07-21\210726_July Catalogs")
output_file = Path(r"C:\Users\hkildani002\Downloads\OneDrive_2026-07-21\210726_Consolidated_Catalog.xlsx")

def clean_sheet_name(name, existing_names):
    # Remove Excel-invalid characters
    sheet_name = re.sub(r'[\[\]\:\*\?\/\\]', "_", name).strip()

    if not sheet_name:
        sheet_name = "Sheet"

    # Truncate to Excel's 31-character sheet limit
    sheet_name = sheet_name[:31]

    # Ensure sheet name is unique
    original = sheet_name
    counter = 1

    while sheet_name in existing_names:
        suffix = f"_{counter}"
        sheet_name = original[:31 - len(suffix)] + suffix
        counter += 1

    existing_names.add(sheet_name)
    return sheet_name

existing_sheet_names = set()

with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    for file in input_folder.glob("*.xlsx"):
        if file.resolve() == output_file.resolve():
            continue

        # Read all tabs from each workbook
        workbook_sheets = pd.read_excel(file, sheet_name=None)

        for tab_name, df in workbook_sheets.items():
            new_sheet_name = clean_sheet_name(
                f"{file.stem}_{tab_name}",
                existing_sheet_names
            )

            df.to_excel(writer, sheet_name=new_sheet_name, index=False)

print(f"Consolidated file created: {output_file}")