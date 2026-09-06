<a id="english"></a>

<div align="center">

**🇬🇧 English** · [🇷🇺 Русский](#russian)

</div>

# CSV / Excel Automation Script

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

---

<a id="russian"></a>

<div align="center">

[🇬🇧 English](#english) · **🇷🇺 Русский**

</div>

# CSV / Excel Automation Script — Русская версия

Python-скрипт для автоматической очистки сырых CSV-данных и формирования готового Excel-отчёта.

## Что решает проект

В реальных таблицах часто встречаются дубли заказов, некорректные email, разные форматы имён и дат, пустые поля и ошибочные числовые значения. Скрипт автоматически очищает такие данные и формирует понятный отчёт вместо ручной обработки.

## Возможности

- чтение CSV с заказами и клиентами;
- нормализация имён, email, телефонов, товаров и дат;
- удаление дубликатов заказов;
- проверка обязательных полей, email, количества и цены;
- разделение корректных и ошибочных строк;
- расчёт выручки и основных показателей;
- генерация Excel-файла с листами Dashboard, Clean Data и Invalid Rows;
- диаграмма и сводные метрики.

## Стек

Python, Pandas, OpenPyXL, CSV, Excel automation, data validation.

## Запуск

```bash
python -m venv .venv
```

Windows:

```bat
.venv\Scripts\activate
pip install -r requirements.txt
python csv_excel_automation.py
```

macOS / Linux:

```bash
source .venv/bin/activate
pip install -r requirements.txt
python csv_excel_automation.py
```

Результат появится в:

```text
output/clean_sales_report.xlsx
```

## Где можно применять

- очистка клиентских и лид-листов;
- подготовка отчётов по продажам;
- удаление дублей перед импортом в CRM;
- проверка CSV перед загрузкой в другие системы;
- автоматическое преобразование сырых данных в готовый Excel-отчёт.

Проект создан как демонстрация практической автоматизации обработки таблиц на Python.
