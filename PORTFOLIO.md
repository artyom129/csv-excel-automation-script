# Upwork Portfolio Entry

## Title

CSV / Excel Data Cleaning Automation

## Description

Built a Python automation script that cleans messy CSV order data, removes duplicate records, validates order ids, names, emails, dates, quantities, and prices, and generates a structured Excel report.

The workbook includes a dashboard with charts, a clean-data sheet, an invalid-records sheet that states the reason for every rejected row, a per-product revenue summary, and calculated revenue metrics.

The script runs from the command line on any CSV file, returns a proper exit code for scheduled jobs, and is covered by an automated pytest suite.

## Technologies

Python, Pandas, OpenPyXL, pytest, CSV processing, Excel automation, data validation.

## Client value

The automation reduces repetitive spreadsheet cleanup, improves data quality, and produces an accounting- or CRM-ready report. Nothing is dropped silently: every rejected row is kept with the reason it failed.
