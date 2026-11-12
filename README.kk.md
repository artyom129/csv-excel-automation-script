<div align="center">

# 🧹 CSV → Excel Automation Script

**Тапсырыстары бар «лас» CSV файлын бір пәрменмен таза, пішімделген Excel есебіне айналдырады.**

[![Tests](https://github.com/artyom129/csv-excel-automation-script/actions/workflows/tests.yml/badge.svg)](https://github.com/artyom129/csv-excel-automation-script/actions/workflows/tests.yml)
![Python](https://img.shields.io/badge/python-3.11%20%7C%203.12-3776AB?logo=python&logoColor=white)
![pandas](https://img.shields.io/badge/pandas-3.x-150458?logo=pandas&logoColor=white)
![openpyxl](https://img.shields.io/badge/openpyxl-3.1-217346?logo=microsoftexcel&logoColor=white)

**🌐 Тіл:** [English](README.md) · [Русский](README.ru.md) · **Қазақша**

</div>

---

## 📌 Мәселе

Шағын компаниялар жиі қайталанған тапсырыстары, біркелкі емес аты-жөндері, қате email мекенжайлары, бұзылған күндері және дұрыс емес сандары бар кестелер алады. Мұндай файлдарды қолмен тазалау ұзақ, жалықтырарлық және қателікке толы.

## ✨ Скрипт не істейді

| | Мүмкіндік |
|---|---|
| 🔤 | Аттарды, email, телефондарды (`+1 (555) 111-222` → `+1555111222`), тауарларды және күндерді **қалыпқа келтіреді** |
| 🗑️ | **Қайталанған тапсырыстарды жояды** (бірінші жазба қалады) және олардың санын хабарлайды |
| ✅ | Тапсырыс нөмірін, клиент атын, email, тауарды, күнді, санды және бағаны **тексереді** |
| 🔍 | **Әр бас тартуды түсіндіреді**: жолдың барлық мәселесі `error` бағанында көрсетіледі, ештеңе үнсіз жоғалмайды |
| 📊 | Дашборды, диаграммалары, сүзгілері және ақша пішімі бар **пішімделген Excel файлын жасайды** |
| 💻 | Кез келген қалтадан **командалық жолмен жұмыс істейді** және дұрыс шығу кодын қайтарады |
| 🧪 | **Автотесттер** (pytest) әр push кезінде GitHub Actions арқылы іске қосылады |

## 📑 Excel есебі

| Парақ | Мазмұны |
|---|---|
| **Dashboard** | Жарамды тапсырыстар, жарамсыз жолдар, жойылған қайталанулар, бірегей клиенттер, жалпы түсім, орташа чек және *Rows by outcome* диаграммасы |
| **Clean Data** | Excel-дің нақты күндері, ақша пішімі және сүзгілері бар тексерілген тапсырыстар |
| **Invalid Rows** | Қабылданбаған жолдар, себебі қызыл түспен белгіленеді |
| **By Product** | Әр тауар бойынша тапсырыстар, дана саны және түсім, *Revenue by product* диаграммасы |

## 🚀 Жылдам бастау

```bash
python -m venv .venv
```

Ортаны іске қосыңыз:

```bash
# Windows
.venv\Scripts\activate

# macOS / Linux
source .venv/bin/activate
```

Тәуелділіктерді орнатып, кірістірілген үлгіде іске қосыңыз:

```bash
pip install -r requirements.txt
python csv_excel_automation.py
```

Есеп `output/clean_sales_report.xlsx` файлында пайда болады.

## 🛠️ Қолдану

```bash
python csv_excel_automation.py                      # sample_orders.csv -> output/clean_sales_report.xlsx
python csv_excel_automation.py orders.csv           # өз файлыңыз
python csv_excel_automation.py orders.csv -o report.xlsx
python csv_excel_automation.py --help
```

Кірістірілген үлгі үшін шығыс мысалы:

```text
Done. Report created: output/clean_sales_report.xlsx
Valid rows:         4
Invalid rows:       3
Duplicates removed: 1
Total revenue:      191.96
```

> **Шығу кодтары:** сәтті болса `0`, файл табылмаса немесе міндетті бағандар жетіспесе `1`. Сондықтан скриптті кесте бойынша жұмыстарда қауіпсіз қолдануға болады.

### Күтілетін CSV бағандары

```text
order_id, date, customer_name, email, phone, product, quantity, price
```

## 🧪 Тесттер

```bash
python -m pytest -q
```

Тесттер қалыпқа келтіруді, қайталануларды өңдеуді, әр тексеру ережесін, бос енгізуді, тауарлар бойынша жиынтықты, CSV → Excel толық жұмысын (кітап қайта ашылып тексеріледі) және CLI-ды қамтиды.

## 📂 Жоба құрылымы

```text
.
├── csv_excel_automation.py     # тазалау, тексеру, Excel-ге экспорт, CLI
├── sample_orders.csv           # демонстрациялық деректер
├── tests/test_automation.py    # pytest тесттері
├── requirements.txt
└── .github/workflows/tests.yml # CI: Python 3.11 және 3.12
```

## 💼 Қолдану мысалдары

- Клиенттер мен лидтер тізімін тазалау
- Сатылым есептерін дайындау
- CRM-дегі қайталанған жазбаларды жою
- Импорттаудан бұрын CSV файлдарын тексеру
- Шикі CSV деректерін басшылыққа арналған есептерге айналдыру

## 🧰 Технологиялар

Python · pandas · openpyxl · pytest · GitHub Actions

---

<div align="center">

*Жеке демонстрациялық жоба: Python-да практикалық автоматтандыру, кестелерді тазалау, тексеру және Excel есептілігі.*

</div>
