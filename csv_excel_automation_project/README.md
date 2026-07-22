# CSV / Excel Automation Script

A small Python automation project that cleans a CSV file, removes duplicate rows, validates emails, calculates basic order metrics, and generates a formatted Excel report.

## What this project does

- Reads raw CSV order/customer data
- Removes duplicate records
- Cleans names, emails, phone numbers and dates
- Detects invalid rows
- Calculates total revenue and order statistics
- Creates a clean Excel report with:
  - Dashboard
  - Clean Data
  - Invalid Rows / Errors

## Tech Stack

- Python
- pandas
- openpyxl
- Excel automation
- CSV data processing

## How to run

1. Install dependencies:

```bash
pip install -r requirements.txt
```

2. Run the script:

```bash
python csv_excel_automation.py
```

3. Output file will be created here:

```text
output/clean_sales_report.xlsx
```

## Example use cases

This script can be adapted for:

- cleaning customer lead lists
- preparing sales reports
- removing duplicate contacts
- converting CSV files into Excel reports
- preparing data for Google Sheets or CRM import

## Portfolio description

I built a Python automation script that takes messy CSV data, cleans and validates it, removes duplicates, and generates a structured Excel report with a dashboard and error sheet. This helps businesses save time on repetitive spreadsheet cleanup and reporting tasks.