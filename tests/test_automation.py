import pandas as pd

from csv_excel_automation import build_dashboard, clean_data


def test_clean_data_removes_duplicates_and_separates_invalid_rows():
    raw = pd.DataFrame([
        {
            "order_id": 1,
            "date": "2026-07-01",
            "customer_name": " john smith ",
            "email": "JOHN@example.com",
            "phone": "123",
            "product": " Keyboard ",
            "quantity": 2,
            "price": 10,
        },
        {
            "order_id": 1,
            "date": "2026-07-01",
            "customer_name": "Duplicate",
            "email": "duplicate@example.com",
            "phone": "456",
            "product": "Mouse",
            "quantity": 1,
            "price": 5,
        },
        {
            "order_id": 2,
            "date": "bad-date",
            "customer_name": "",
            "email": "bad-email",
            "phone": "789",
            "product": "Monitor",
            "quantity": 0,
            "price": -1,
        },
    ])

    clean, invalid = clean_data(raw)

    assert len(clean) == 1
    assert len(invalid) == 1
    assert clean.iloc[0]["customer_name"] == "John Smith"
    assert clean.iloc[0]["email"] == "john@example.com"
    assert clean.iloc[0]["total"] == 20
    assert "Invalid email" in invalid.iloc[0]["error"]


def test_dashboard_metrics():
    clean = pd.DataFrame([
        {"email": "a@example.com", "total": 10},
        {"email": "b@example.com", "total": 30},
    ])
    invalid = pd.DataFrame([{"error": "Invalid email"}])

    dashboard = build_dashboard(clean, invalid)
    metrics = dict(zip(dashboard["Metric"], dashboard["Value"]))

    assert metrics["Valid orders"] == 2
    assert metrics["Invalid rows"] == 1
    assert metrics["Unique customers"] == 2
    assert metrics["Total revenue"] == 40
    assert metrics["Average order value"] == 20


# --------------------------------------------------------------------------
# Extended coverage
# --------------------------------------------------------------------------

import pytest
from openpyxl import load_workbook

from csv_excel_automation import (
    build_product_summary,
    count_duplicates,
    load_data,
    main,
    normalize_phone,
    run,
)

HEADER = "order_id,date,customer_name,email,phone,product,quantity,price\n"


def write_csv(path, rows):
    path.write_text(HEADER + "\n".join(rows) + "\n", encoding="utf-8")
    return path


@pytest.mark.parametrize("raw, expected", [
    ("+1 (555) 111-222", "+1555111222"),
    ("555-999", "555999"),
    ("  123 ", "123"),
    ("", ""),
    (None, ""),
    ("n/a", ""),
])
def test_normalize_phone(raw, expected):
    assert normalize_phone(raw) == expected


def test_errors_are_listed_per_row_without_trailing_separator():
    raw = pd.DataFrame([{
        "order_id": "7", "date": "nope", "customer_name": "", "email": "x",
        "phone": "1", "product": "", "quantity": 0, "price": -1,
    }])
    _, invalid = clean_data(raw)
    error = invalid.iloc[0]["error"]

    for message in ("Missing customer name", "Invalid email", "Missing product",
                    "Invalid date", "Invalid quantity", "Invalid price"):
        assert message in error
    assert not error.endswith(";") and not error.endswith(" ")


def test_rows_without_order_id_are_reported_not_merged():
    raw = pd.DataFrame([
        {"order_id": None, "date": "2026-01-01", "customer_name": "A B",
         "email": "a@x.io", "phone": "1", "product": "P", "quantity": 1, "price": 1},
        {"order_id": None, "date": "2026-01-02", "customer_name": "C D",
         "email": "c@x.io", "phone": "2", "product": "P", "quantity": 1, "price": 1},
    ])
    clean, invalid = clean_data(raw)

    assert clean.empty
    assert len(invalid) == 2
    assert invalid["error"].str.contains("Missing order id").all()


def test_count_duplicates_ignores_empty_ids():
    df = pd.DataFrame({"order_id": ["1", "1", "2", None, None, "1"]})
    assert count_duplicates(df) == 2


def test_dashboard_handles_empty_input():
    dashboard = build_dashboard(pd.DataFrame(columns=["email", "total"]),
                                pd.DataFrame(columns=["error"]))
    metrics = dict(zip(dashboard["Metric"], dashboard["Value"]))

    assert metrics["Valid orders"] == 0
    assert metrics["Total revenue"] == 0
    assert metrics["Average order value"] == 0


def test_product_summary_sorted_by_revenue():
    clean = pd.DataFrame([
        {"order_id": "1", "product": "Mouse", "quantity": 1, "total": 20.0},
        {"order_id": "2", "product": "Monitor", "quantity": 1, "total": 180.0},
        {"order_id": "3", "product": "Mouse", "quantity": 2, "total": 40.0},
    ])
    summary = build_product_summary(clean)

    assert list(summary["Product"]) == ["Monitor", "Mouse"]
    mouse = summary[summary["Product"] == "Mouse"].iloc[0]
    assert (mouse["Orders"], mouse["Units"], mouse["Revenue"]) == (2, 3, 60.0)


def test_load_data_reports_missing_file_and_columns(tmp_path):
    with pytest.raises(FileNotFoundError):
        load_data(tmp_path / "nope.csv")

    bad = tmp_path / "bad.csv"
    bad.write_text("order_id,date\n1,2026-01-01\n", encoding="utf-8")
    with pytest.raises(ValueError, match="Missing columns"):
        load_data(bad)


def test_end_to_end_report(tmp_path):
    source = write_csv(tmp_path / "orders.csv", [
        "1001,2026-07-01, john smith ,JOHN@example.com,+1 555 111 222,Keyboard,2,45.50",
        "1001,2026-07-01,John Smith,john@example.com,+1 555 111 222,Keyboard,2,45.50",
        "1002,2026-07-02,Anna Brown,anna@example.com,555-333,Mouse,1,20",
        "1003,bad-date,Mike Wilson,bad-email,555,Monitor,1,180",
    ])
    output = tmp_path / "nested" / "report.xlsx"

    result = run(source, output)

    assert result["valid"] == 2
    assert result["invalid"] == 1
    assert result["duplicates"] == 1
    assert result["revenue"] == 111.0
    assert output.exists()

    wb = load_workbook(output)
    assert wb.sheetnames == ["Dashboard", "Clean Data", "Invalid Rows", "By Product"]

    dashboard = {row[0].value: row[1].value
                 for row in wb["Dashboard"].iter_rows(min_row=3, max_col=2)}
    assert dashboard["Valid orders"] == 2
    assert dashboard["Duplicates removed"] == 1
    assert dashboard["Total revenue"] == 111.0

    clean = wb["Clean Data"]
    header = [c.value for c in clean[1]]
    first = dict(zip(header, [c.value for c in clean[2]]))
    assert first["customer_name"] == "John Smith"
    assert first["phone"] == "+1555111222"
    assert first["total"] == 91.0

    invalid = wb["Invalid Rows"]
    assert invalid.max_row == 2
    assert len(wb["Dashboard"]._charts) == 1
    assert len(wb["By Product"]._charts) == 1


def test_cli_success_and_failure(tmp_path, capsys):
    source = write_csv(tmp_path / "orders.csv", [
        "1,2026-01-01,Ann Lee,ann@example.com,1,Pen,2,1.5",
    ])
    output = tmp_path / "out.xlsx"

    assert main([str(source), "-o", str(output)]) == 0
    assert output.exists()
    assert "Valid rows:" in capsys.readouterr().out

    assert main([str(tmp_path / "missing.csv"), "-o", str(output)]) == 1
    assert "Input file not found" in capsys.readouterr().err
