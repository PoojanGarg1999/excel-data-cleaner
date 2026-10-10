
# Excel Data Cleaner 7.1

A Python-based Excel and CSV data cleaning automation tool with a Windows desktop application.

## Features

- Combine multiple Excel and CSV files
- Remove duplicate customer records
- Normalize email addresses
- Detect missing and invalid emails
- Validate Indian mobile numbers
- Recognize common column name variations
- Generate professionally formatted Excel reports
- Choose input files and output location using file dialogs
- Run as a standalone Windows EXE

## Output Sheets

1. Clean Data
2. Missing Emails
3. Invalid Emails
4. Invalid Phones
5. Summary
6. File Errors

## Technologies Used

- Python
- OpenPyXL
- Tkinter
- CSV
- PyInstaller

## How to Run

Install the dependency:

```bash
pip install openpyxl
```

Run the Python script:

```bash
python bulk_excel_cleaner.py
```

Select your Excel or CSV files, choose where to save the report, and wait for the completion message.

## Build Windows Application

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name ExcelDataCleaner bulk_excel_cleaner.py
```

The executable will be generated in the `dist` folder.

## Sample Test Results

- Files processed: 2
- Total records: 9
- Duplicates removed: 2
- Missing emails: 1
- Clean customers: 6

## Limitations

- Input files must have compatible columns.
- Email validation checks formatting, not whether an address exists.
- Phone validation is designed for Indian mobile numbers.
- Excel processing uses the active worksheet.

## Privacy

Do not upload real customer data or confidential spreadsheets to public repositories.

## Project Status

Version 7.1 — Tested on Windows.
## Demo Screenshot

![Excel Data Cleaner Demo](excel-cleaner-demo%201.png)
