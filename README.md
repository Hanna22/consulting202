# Consulting 202

Small Python tools that take repetitive Excel work off a consultant's plate.

You do **not** need to know how to code to use this project. Follow the steps below exactly. Your original Excel files are not added to this repository, and you should never upload client data to GitHub.

## What is included

| Tool | What it does | What it creates |
| --- | --- | --- |
| `Consolidate_Excels.py` | Copies every worksheet from every `.xlsx` file in a folder into one workbook | One workbook containing many worksheets |
| `concat_colms_from_sheets.py` | Takes a selected list of columns from every worksheet and stacks the rows together | One workbook with a `Consolidated` worksheet |
| `split_excels.py` | Splits the first worksheet into a separate workbook for every value in a chosen column | A folder containing multiple workbooks |

## Before you start (Windows)

### 1. Install Python

1. Go to [python.org/downloads](https://www.python.org/downloads/).
2. Download Python 3.11 or newer.
3. Open the installer.
4. Tick **Add python.exe to PATH**.
5. Select **Install Now**.

### 2. Download this project

1. On this GitHub page, select the green **Code** button.
2. Select **Download ZIP**.
3. Open your Downloads folder.
4. Right-click the ZIP file and select **Extract All**.
5. Open the extracted `consulting202` folder.

### 3. Open a command window in the project folder

1. Click the folder's address bar in File Explorer.
2. Type `powershell`.
3. Press Enter.

A blue or black command window will open. You will paste commands into that window.

### 4. Create a private Python environment

Paste these commands one at a time, pressing Enter after each one:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

If the second command is blocked, run this command once and then try it again:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

The environment is ready when you see `(.venv)` at the start of the command line.

## Tool 1: combine workbooks and keep every worksheet

Use `Consolidate_Excels.py` when you have a folder of Excel files and want every worksheet copied into one workbook.

1. Right-click `Consolidate_Excels.py` and open it with Notepad.
2. Find the line beginning with `input_folder =`.
3. Replace the path inside the quotation marks with the folder containing your Excel files.
4. Find the line beginning with `output_file =`.
5. Replace that path with the full name and location for the new workbook.
6. Save and close Notepad.
7. In PowerShell, run:

```powershell
python .\Consolidate_Excels.py
```

The tool reads `.xlsx` files directly inside the chosen folder. It does not search subfolders. The resulting worksheet names may be shortened because Excel limits names to 31 characters.

## Tool 2: stack selected columns from every worksheet

Use `concat_colms_from_sheets.py` when worksheets have similar columns and you want all their rows in one table.

1. Right-click `concat_colms_from_sheets.py` and open it with Notepad.
2. Change `file_path` to the workbook you want to read.
3. Change `output_file` to the new workbook you want to create.
4. Review `cols_to_keep`. These names must match the Excel column headings. You may add, remove, or rename items inside the square brackets.
5. Save and close Notepad.
6. In PowerShell, run:

```powershell
python .\concat_colms_from_sheets.py
```

The result contains one worksheet named `Consolidated`. A `Source Sheet` column records where each row came from. Worksheets with none of the requested columns are skipped.

## Tool 3: split one workbook into many workbooks

Use `split_excels.py` when one worksheet contains a category, department, owner, country, or similar column and you need a separate file for every value.

For example, imagine `input.xlsx` has a column named `Country`. Put that workbook in this project folder and run:

```powershell
python .\split_excels.py .\input.xlsx --key-column "Country" --output-dir .\country_files
```

Replace:

- `input.xlsx` with your workbook's filename.
- `Country` with the exact column heading you want to split by.
- `country_files` with the folder name you want for the results.

The tool uses the first worksheet only. It keeps all columns in their original order, places blank key values in a `blank` group, avoids overwriting existing output files, and prints detailed progress messages.

Optional: add text to the start of every output filename:

```powershell
python .\split_excels.py .\input.xlsx --key-column "Country" --output-dir .\country_files --prefix "country_"
```

To see every available option:

```powershell
python .\split_excels.py --help
```

## Common problems

### `python` is not recognized

Python is not installed or was not added to PATH. Reinstall Python and make sure **Add python.exe to PATH** is selected.

### `No module named pandas` or `No module named openpyxl`

Open PowerShell in the project folder, activate the environment, and install the requirements again:

```powershell
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

### Permission denied

Close the Excel workbook if it is open, then run the tool again. Excel often locks open files.

### A column cannot be found

Check spelling and spaces in the first row of the worksheet. For `split_excels.py`, the name is not case-sensitive unless you add `--case-sensitive-column`.

### The output is not where expected

Read the final message in PowerShell. Each tool prints the output location when it finishes.

## Handling client data safely

- Work on approved devices and approved storage locations only.
- Keep client workbooks outside this project folder whenever possible.
- Never commit or upload client workbooks, credentials, or confidential information to GitHub.
- Check the output before sharing it. Automation reduces repetitive work but does not replace professional review.
- Keep an untouched copy of the original workbook until you have verified the result.

## For contributors

Read `AGENTS.md` before changing the tools. Keep changes small, document them in plain language, use synthetic test data, and make focused Git commits with detailed explanations.
