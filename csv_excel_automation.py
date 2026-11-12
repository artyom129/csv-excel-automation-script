"""Clean a messy CSV of orders and build a formatted Excel report.

Usage:
    python csv_excel_automation.py [input.csv] [-o report.xlsx]
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_INPUT = BASE_DIR / "sample_orders.csv"
DEFAULT_OUTPUT = BASE_DIR / "output" / "clean_sales_report.xlsx"

REQUIRED_COLUMNS = {
    "order_id", "date", "customer_name", "email",
    "phone", "product", "quantity", "price",
}

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

MONEY_FORMAT = "#,##0.00"
DATE_FORMAT = "yyyy-mm-dd"
MONEY_HEADERS = {"price", "total", "Revenue"}
MONEY_METRICS = {"Total revenue", "Average order value"}


# --------------------------------------------------------------------------
# Normalization helpers
# --------------------------------------------------------------------------

def normalize_name(value) -> str:
    if pd.isna(value):
        return ""
    return " ".join(str(value).split()).title()


def normalize_email(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip().lower()


def normalize_phone(value) -> str:
    """Keep digits and a leading '+', drop spaces, dashes and brackets."""
    if pd.isna(value):
        return ""
    raw = str(value).strip()
    digits = re.sub(r"\D", "", raw)
    if not digits:
        return ""
    return ("+" if raw.startswith("+") else "") + digits


def normalize_text(value) -> str:
    if pd.isna(value):
        return ""
    return " ".join(str(value).split())


# --------------------------------------------------------------------------
# Loading and cleaning
# --------------------------------------------------------------------------

def load_data(path: Path) -> pd.DataFrame:
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"Input file not found: {path}")

    # Read ids and phones as text so '1001' and '+1 555...' are not mangled.
    df = pd.read_csv(path, dtype={"order_id": str, "phone": str})
    df.columns = [str(c).strip() for c in df.columns]

    missing = REQUIRED_COLUMNS - set(df.columns)
    if missing:
        raise ValueError(f"Missing columns: {', '.join(sorted(missing))}")

    return df


def count_duplicates(df: pd.DataFrame) -> int:
    """Number of rows that repeat an already seen, non-empty order_id."""
    order_ids = df["order_id"].map(normalize_text)
    return int((order_ids.ne("") & order_ids.duplicated()).sum())


def clean_data(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    df = df.copy()

    df["order_id"] = df["order_id"].map(normalize_text)
    df["customer_name"] = df["customer_name"].map(normalize_name)
    df["email"] = df["email"].map(normalize_email)
    df["phone"] = df["phone"].map(normalize_phone)
    df["product"] = df["product"].map(normalize_text)

    df["quantity"] = pd.to_numeric(df["quantity"], errors="coerce")
    df["price"] = pd.to_numeric(df["price"], errors="coerce")
    df["date"] = pd.to_datetime(df["date"], errors="coerce")

    # Repeated order ids are dropped (first one wins); empty ids are kept so
    # they can be reported as errors instead of silently collapsing together.
    is_repeat = df["order_id"].ne("") & df["order_id"].duplicated()
    df = df[~is_repeat].copy()

    df["error"] = ""

    def flag(mask: pd.Series, message: str) -> None:
        df.loc[mask, "error"] += message + "; "

    flag(df["order_id"].eq(""), "Missing order id")
    flag(df["customer_name"].eq(""), "Missing customer name")
    flag(~df["email"].map(lambda x: bool(EMAIL_RE.match(x))), "Invalid email")
    flag(df["product"].eq(""), "Missing product")
    flag(df["date"].isna(), "Invalid date")
    flag(df["quantity"].isna() | (df["quantity"] <= 0), "Invalid quantity")
    flag(df["price"].isna() | (df["price"] <= 0), "Invalid price")

    df["error"] = df["error"].str.rstrip("; ")
    df["total"] = (df["quantity"] * df["price"]).round(2)

    invalid_df = df[df["error"].ne("")].copy()
    clean_df = df[df["error"].eq("")].drop(columns=["error"]).copy()

    for frame in (clean_df, invalid_df):
        frame["date"] = frame["date"].dt.strftime("%Y-%m-%d")

    return clean_df, invalid_df


# --------------------------------------------------------------------------
# Summaries
# --------------------------------------------------------------------------

def build_dashboard(
    clean_df: pd.DataFrame,
    invalid_df: pd.DataFrame,
    duplicates: int = 0,
) -> pd.DataFrame:
    total_orders = len(clean_df)
    total_revenue = clean_df["total"].sum() if total_orders else 0
    avg_order_value = clean_df["total"].mean() if total_orders else 0
    unique_customers = clean_df["email"].nunique() if total_orders else 0

    return pd.DataFrame({
        "Metric": [
            "Valid orders",
            "Invalid rows",
            "Duplicates removed",
            "Unique customers",
            "Total revenue",
            "Average order value",
        ],
        "Value": [
            total_orders,
            len(invalid_df),
            duplicates,
            unique_customers,
            round(float(total_revenue), 2),
            round(float(avg_order_value), 2),
        ],
    })


def build_product_summary(clean_df: pd.DataFrame) -> pd.DataFrame:
    columns = ["Product", "Orders", "Units", "Revenue"]
    if clean_df.empty:
        return pd.DataFrame(columns=columns)

    summary = (
        clean_df.groupby("product")
        .agg(Orders=("order_id", "count"),
             Units=("quantity", "sum"),
             Revenue=("total", "sum"))
        .reset_index()
        .rename(columns={"product": "Product"})
        .sort_values(["Revenue", "Product"], ascending=[False, True])
    )
    summary["Revenue"] = summary["Revenue"].round(2)
    return summary[columns].reset_index(drop=True)


# --------------------------------------------------------------------------
# Excel export
# --------------------------------------------------------------------------

def export_excel(
    clean_df: pd.DataFrame,
    invalid_df: pd.DataFrame,
    duplicates: int = 0,
    output: Path = DEFAULT_OUTPUT,
) -> Path:
    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)

    dashboard_df = build_dashboard(clean_df, invalid_df, duplicates)
    product_df = build_product_summary(clean_df)

    clean_out = clean_df.copy()
    clean_out["date"] = pd.to_datetime(clean_out["date"])
    invalid_out = invalid_df.copy()
    invalid_out["date"] = pd.to_datetime(invalid_out["date"])

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # startrow=1 leaves row 1 free for the dashboard title.
        dashboard_df.to_excel(writer, sheet_name="Dashboard", index=False, startrow=1)
        clean_out.to_excel(writer, sheet_name="Clean Data", index=False)
        invalid_out.to_excel(writer, sheet_name="Invalid Rows", index=False)
        product_df.to_excel(writer, sheet_name="By Product", index=False)

    format_workbook(output)
    return output


def format_workbook(path: Path) -> None:
    wb = load_workbook(path)

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    title_font = Font(size=16, bold=True, color="1F4E78")
    error_font = Font(color="C00000")
    thin = Side(border_style="thin", color="D9E2F3")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for ws in wb.worksheets:
        header_row = 2 if ws.title == "Dashboard" else 1
        headers = {cell.column: str(cell.value) for cell in ws[header_row]}

        ws.freeze_panes = f"A{header_row + 1}"
        if ws.title != "Dashboard" and ws.max_row > 1:
            ws.auto_filter.ref = ws.dimensions

        for cell in ws[header_row]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")
            cell.border = border

        for row in ws.iter_rows(min_row=header_row + 1):
            for cell in row:
                cell.border = border
                cell.alignment = Alignment(vertical="center")
                header = headers.get(cell.column, "")
                if header in MONEY_HEADERS:
                    cell.number_format = MONEY_FORMAT
                elif header == "date":
                    cell.number_format = DATE_FORMAT
                elif header == "error":
                    cell.font = error_font

        for column_cells in ws.iter_cols(min_row=header_row):
            longest = max(len("" if c.value is None else str(c.value))
                          for c in column_cells)
            letter = column_cells[0].column_letter
            ws.column_dimensions[letter].width = min(longest + 3, 40)

    dashboard = wb["Dashboard"]
    dashboard["A1"] = "CSV / Excel Automation Report"
    dashboard["A1"].font = title_font
    dashboard["A1"].alignment = Alignment(horizontal="center")
    dashboard.merge_cells("A1:B1")

    for row in dashboard.iter_rows(min_row=3, max_col=2):
        metric, value = row
        if metric.value in MONEY_METRICS:
            value.number_format = MONEY_FORMAT

    # Dashboard chart: data-quality counts only, so the scale stays readable.
    counts = BarChart()
    counts.title = "Rows by outcome"
    counts.legend = None
    counts.add_data(Reference(dashboard, min_col=2, min_row=3, max_row=5),
                    titles_from_data=False)
    counts.set_categories(Reference(dashboard, min_col=1, min_row=3, max_row=5))
    counts.height, counts.width = 7, 14
    dashboard.add_chart(counts, "D2")

    products = wb["By Product"]
    if products.max_row > 1:
        revenue = BarChart()
        revenue.title = "Revenue by product"
        revenue.legend = None
        revenue.y_axis.title = "Revenue"
        revenue.add_data(Reference(products, min_col=4, min_row=2,
                                   max_row=products.max_row),
                         titles_from_data=False)
        revenue.set_categories(Reference(products, min_col=1, min_row=2,
                                         max_row=products.max_row))
        revenue.height, revenue.width = 8, 16
        products.add_chart(revenue, "F2")

    wb.save(path)


# --------------------------------------------------------------------------
# Entry point
# --------------------------------------------------------------------------

def run(input_path: Path = DEFAULT_INPUT, output: Path = DEFAULT_OUTPUT) -> dict:
    raw_df = load_data(input_path)
    duplicates = count_duplicates(raw_df)
    clean_df, invalid_df = clean_data(raw_df)
    report = export_excel(clean_df, invalid_df, duplicates, output)

    return {
        "report": report,
        "valid": len(clean_df),
        "invalid": len(invalid_df),
        "duplicates": duplicates,
        "revenue": round(float(clean_df["total"].sum()), 2) if len(clean_df) else 0.0,
    }


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Clean a CSV of orders and generate a formatted Excel report.")
    parser.add_argument("input", nargs="?", type=Path, default=DEFAULT_INPUT,
                        help="CSV file to process (default: sample_orders.csv)")
    parser.add_argument("-o", "--output", type=Path, default=DEFAULT_OUTPUT,
                        help="Excel file to create (default: output/clean_sales_report.xlsx)")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        result = run(args.input, args.output)
    except (FileNotFoundError, ValueError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    print(f"Done. Report created: {result['report']}")
    print(f"Valid rows:         {result['valid']}")
    print(f"Invalid rows:       {result['invalid']}")
    print(f"Duplicates removed: {result['duplicates']}")
    print(f"Total revenue:      {result['revenue']:,.2f}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
