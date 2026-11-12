# CSV / Excel Automation Script

Python automation for cleaning messy CSV order data and generating a formatted Excel report with a dashboard, clean records, validation errors, and a per-product revenue summary.

## Business problem

Small businesses often receive spreadsheets with duplicate orders, inconsistent names, invalid emails, broken dates, and incorrect numeric values. Manually cleaning those files is repetitive and error-prone.

## Features

- Reads raw CSV order and customer data
- Normalizes names, emails, phone numbers (`+1 (555) 111-222` becomes `+1555111222`), products, and dates
- Removes duplicate orders (first occurrence wins) and reports how many were dropped
- Validates order id, customer name, email format, product, date, quantity, and price
- Lists every problem per row in an `error` column instead of silently dropping data
- Separates valid and invalid records
- Generates a formatted Excel workbook with:
  - **Dashboard**: valid orders, invalid rows, duplicates removed, unique customers, total revenue, average order value, and a chart
  - **Clean Data**: validated orders with real Excel dates, currency formatting, and filters
  - **Invalid Rows**: rejected rows with the reason highlighted
  - **By Product**: orders, units, and revenue per product with a revenue chart
- Command-line interface; works from any directory
- Covered by an automated test suite

## Tech stack

Python, Pandas, OpenPyXL, CSV processing, Excel automation, data validation, pytest.

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

Install dependencies and run on the bundled sample:

```bash
pip install -r requirements.txt
python csv_excel_automation.py
```

The generated workbook will appear at:

```text
output/clean_sales_report.xlsx
```

## Usage

```bash
python csv_excel_automation.py                      # sample_orders.csv -> output/clean_sales_report.xlsx
python csv_excel_automation.py orders.csv           # your own file
python csv_excel_automation.py orders.csv -o report.xlsx
python csv_excel_automation.py --help
```

Example output for the bundled sample:

```text
Done. Report created: output/clean_sales_report.xlsx
Valid rows:         4
Invalid rows:       3
Duplicates removed: 1
Total revenue:      191.96
```

The exit code is `0` on success and `1` if the file is missing or required columns are absent, so the script can be used in scheduled jobs.

### Expected CSV columns

```text
order_id, date, customer_name, email, phone, product, quantity, price
```

## Tests

```bash
python -m pytest -q
```

The suite covers normalization, duplicate handling, every validation rule, empty input, the product summary, an end-to-end CSV to Excel run (the workbook is re-opened and checked), and the CLI.

## Example use cases

- Cleaning customer and lead lists
- Preparing sales reports
- Removing duplicate CRM records
- Validating CSV files before import
- Converting raw CSV data into management-ready Excel reports

## Portfolio positioning

This is a personal demonstration project built to showcase practical Python automation, spreadsheet cleanup, validation, and Excel reporting.
