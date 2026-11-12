<div align="center">

# 🧹 CSV → Excel Automation Script

**Turn messy CSV order data into a clean, formatted Excel report in one command.**

[![Tests](https://github.com/artyom129/csv-excel-automation-script/actions/workflows/tests.yml/badge.svg)](https://github.com/artyom129/csv-excel-automation-script/actions/workflows/tests.yml)
![Python](https://img.shields.io/badge/python-3.11%20%7C%203.12-3776AB?logo=python&logoColor=white)
![pandas](https://img.shields.io/badge/pandas-3.x-150458?logo=pandas&logoColor=white)
![openpyxl](https://img.shields.io/badge/openpyxl-3.1-217346?logo=microsoftexcel&logoColor=white)

**🌐 Language:** **English** · [Русский](README.ru.md) · [Қазақша](README.kk.md)

</div>

---

## 📌 The problem

Small businesses often receive spreadsheets with duplicate orders, inconsistent names, invalid emails, broken dates, and incorrect numbers. Cleaning those files by hand is slow, repetitive, and error-prone.

## ✨ What it does

| | Feature |
|---|---|
| 🔤 | **Normalizes** names, emails, phone numbers (`+1 (555) 111-222` → `+1555111222`), products, and dates |
| 🗑️ | **Removes duplicate orders** (first occurrence wins) and reports how many were dropped |
| ✅ | **Validates** order id, customer name, email, product, date, quantity, and price |
| 🔍 | **Explains every rejection**: all problems of a row are listed in the `error` column, nothing is dropped silently |
| 📊 | **Builds a formatted Excel workbook** with a dashboard, charts, filters, and currency formats |
| 💻 | **Command-line interface** that works from any directory and returns a proper exit code |
| 🧪 | **Automated tests** (pytest) run on every push by GitHub Actions |

## 📑 The Excel report

| Sheet | Contents |
|---|---|
| **Dashboard** | Valid orders, invalid rows, duplicates removed, unique customers, total revenue, average order value, and a *Rows by outcome* chart |
| **Clean Data** | Validated orders with real Excel dates, currency formatting, and filters |
| **Invalid Rows** | Rejected rows with the reason highlighted in red |
| **By Product** | Orders, units, and revenue per product, with a *Revenue by product* chart |

## 🚀 Quick start

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

Install the dependencies and run on the bundled sample:

```bash
pip install -r requirements.txt
python csv_excel_automation.py
```

The report appears at `output/clean_sales_report.xlsx`.

## 🛠️ Usage

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

> **Exit codes:** `0` on success, `1` if the file is missing or required columns are absent. This makes the script safe to use in scheduled jobs.

### Expected CSV columns

```text
order_id, date, customer_name, email, phone, product, quantity, price
```

## 🧪 Tests

```bash
python -m pytest -q
```

The suite covers normalization, duplicate handling, every validation rule, empty input, the product summary, an end-to-end CSV → Excel run (the workbook is re-opened and verified), and the CLI.

## 📂 Project structure

```text
.
├── csv_excel_automation.py     # cleaning, validation, Excel export, CLI
├── sample_orders.csv           # demo input
├── tests/test_automation.py    # pytest suite
├── requirements.txt
└── .github/workflows/tests.yml # CI: Python 3.11 and 3.12
```

## 💼 Example use cases

- Cleaning customer and lead lists
- Preparing sales reports
- Removing duplicate CRM records
- Validating CSV files before import
- Converting raw CSV data into management-ready Excel reports

## 🧰 Tech stack

Python · pandas · openpyxl · pytest · GitHub Actions

---

<div align="center">

*A personal demonstration project showcasing practical Python automation, spreadsheet cleanup, validation, and Excel reporting.*

</div>
