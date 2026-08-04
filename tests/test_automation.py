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
