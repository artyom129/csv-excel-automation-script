# CSV / Excel Automation Script

[![Tests](https://github.com/artyom129/csv-excel-automation-script/actions/workflows/tests.yml/badge.svg)](https://github.com/artyom129/csv-excel-automation-script/actions/workflows/tests.yml)

Python automation for cleaning messy CSV order data and generating a formatted Excel report with a dashboard, clean records, and validation errors.

## Business problem

Small businesses often receive spreadsheets with duplicate orders, inconsistent names, invalid emails, broken dates, and incorrect numeric values. Manually cleaning those files is repetitive and error-prone.

## Features

- Reads raw CSV order and customer data
- Normalizes names, emails, phone numbers, products, and dates
- Removes duplicate orders
- Validates required fields, email format, quantities, and prices
- Separates valid and invalid records
- Calculates revenue and order metrics
- Generates a formatted Excel workbook with:
  - Dashboard
  - Clean Data
  - Invalid Rows
  - Bar chart and summary metrics

## Tech stack

Python, Pandas, OpenPyXL, CSV processing, Excel automation, data validation.

## Quick start

```bash
python -m venv .venv
```

Activate the environment:

```bash
# Windows
.venv\Scripts\activate

# macOS / Linux
source .venv/bin/activate
```

Install dependencies and run:

```bash
pip install -r requirements.txt
python csv_excel_automation.py
```

The generated workbook will appear at:

```text
output/clean_sales_report.xlsx
```

## Example use cases

- Cleaning customer and lead lists
- Preparing sales reports
- Removing duplicate CRM records
- Validating CSV files before import
- Converting raw CSV data into management-ready Excel reports

## Portfolio positioning

This is a personal demonstration project built to showcase practical Python automation, spreadsheet cleanup, validation, and Excel reporting.
