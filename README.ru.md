<div align="center">

# 🧹 CSV → Excel Automation Script

**Превращает «грязный» CSV с заказами в чистый отформатированный Excel-отчёт одной командой.**

[![Tests](https://github.com/artyom129/csv-excel-automation-script/actions/workflows/tests.yml/badge.svg)](https://github.com/artyom129/csv-excel-automation-script/actions/workflows/tests.yml)
![Python](https://img.shields.io/badge/python-3.11%20%7C%203.12-3776AB?logo=python&logoColor=white)
![pandas](https://img.shields.io/badge/pandas-3.x-150458?logo=pandas&logoColor=white)
![openpyxl](https://img.shields.io/badge/openpyxl-3.1-217346?logo=microsoftexcel&logoColor=white)

**🌐 Язык:** [English](README.md) · **Русский** · [Қазақша](README.kk.md)

</div>

---

## 📌 Проблема

Небольшие компании часто получают таблицы с дубликатами заказов, неоднородными именами, некорректными email, битыми датами и неверными числами. Чистить такие файлы вручную долго, однообразно и легко ошибиться.

## ✨ Что делает скрипт

| | Возможность |
|---|---|
| 🔤 | **Нормализует** имена, email, телефоны (`+1 (555) 111-222` → `+1555111222`), товары и даты |
| 🗑️ | **Удаляет дубликаты заказов** (остаётся первая запись) и сообщает, сколько их было |
| ✅ | **Проверяет** номер заказа, имя клиента, email, товар, дату, количество и цену |
| 🔍 | **Объясняет каждый отказ**: все проблемы строки перечислены в колонке `error`, ничего не пропадает молча |
| 📊 | **Строит отформатированный Excel-файл** с дашбордом, графиками, фильтрами и денежными форматами |
| 💻 | **Работает из командной строки** из любой папки и возвращает корректный код выхода |
| 🧪 | **Автотесты** (pytest) запускаются в GitHub Actions при каждом пуше |

## 📑 Excel-отчёт

| Лист | Содержимое |
|---|---|
| **Dashboard** | Валидные заказы, невалидные строки, удалённые дубликаты, уникальные клиенты, общая выручка, средний чек и график *Rows by outcome* |
| **Clean Data** | Проверенные заказы с настоящими датами Excel, денежным форматом и фильтрами |
| **Invalid Rows** | Отклонённые строки, причина выделена красным |
| **By Product** | Заказы, штуки и выручка по каждому товару, график *Revenue by product* |

## 🚀 Быстрый старт

```bash
python -m venv .venv
```

Активируйте окружение:

```bash
# Windows
.venv\Scripts\activate

# macOS / Linux
source .venv/bin/activate
```

Установите зависимости и запустите на встроенном примере:

```bash
pip install -r requirements.txt
python csv_excel_automation.py
```

Отчёт появится в `output/clean_sales_report.xlsx`.

## 🛠️ Использование

```bash
python csv_excel_automation.py                      # sample_orders.csv -> output/clean_sales_report.xlsx
python csv_excel_automation.py orders.csv           # ваш файл
python csv_excel_automation.py orders.csv -o report.xlsx
python csv_excel_automation.py --help
```

Пример вывода для встроенного примера:

```text
Done. Report created: output/clean_sales_report.xlsx
Valid rows:         4
Invalid rows:       3
Duplicates removed: 1
Total revenue:      191.96
```

> **Коды выхода:** `0` при успехе, `1`, если файл не найден или не хватает обязательных колонок. Поэтому скрипт безопасно использовать в заданиях по расписанию.

### Ожидаемые колонки CSV

```text
order_id, date, customer_name, email, phone, product, quantity, price
```

## 🧪 Тесты

```bash
python -m pytest -q
```

Тесты покрывают нормализацию, обработку дубликатов, каждое правило валидации, пустой ввод, сводку по товарам, полный прогон CSV → Excel (книга открывается заново и проверяется) и CLI.

## 📂 Структура проекта

```text
.
├── csv_excel_automation.py     # очистка, валидация, экспорт в Excel, CLI
├── sample_orders.csv           # демонстрационные данные
├── tests/test_automation.py    # тесты pytest
├── requirements.txt
└── .github/workflows/tests.yml # CI: Python 3.11 и 3.12
```

## 💼 Примеры применения

- Очистка списков клиентов и лидов
- Подготовка отчётов по продажам
- Удаление дублей в CRM
- Проверка CSV-файлов перед импортом
- Превращение сырых CSV-данных в отчёты для руководства

## 🧰 Технологии

Python · pandas · openpyxl · pytest · GitHub Actions

---

<div align="center">

*Личный демонстрационный проект: практическая автоматизация на Python, очистка таблиц, валидация и Excel-отчётность.*

</div>
