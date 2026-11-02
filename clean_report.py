"""
Task 4: Data Cleaning & Reporting Automation

What this script does automatically:
1. Loads a messy dataset (creates a sample one if messy_data.csv is missing)
2. Cleans it: duplicates, missing values, inconsistent text/dates/numbers
3. Saves the cleaned data to cleaned_data.csv
4. Generates an Excel report (report.xlsx) and visual summaries (PNG charts)
5. Writes a short text summary (summary.txt)

Run:  python clean_report.py
Use your own file:  python clean_report.py your_file.csv
"""

import os
import sys

import matplotlib
matplotlib.use("Agg")  # works without a display
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

RAW_FILE = "messy_data.csv"


# ------------------------------------------------------- SAMPLE MESSY DATA
def create_messy_data(path):
    """Create a deliberately messy sales dataset."""
    rng = np.random.default_rng(7)
    n = 120
    regions = ["North", "south", "SOUTH ", "East", "west", "West", "north ", "EAST"]
    products = ["Laptop", "Phone", "Tablet", "laptop", "PHONE", "Headphones", "tablet "]
    dates = pd.date_range("2025-01-01", periods=n, freq="3D").strftime("%d-%m-%Y").tolist()
    df = pd.DataFrame(
        {
            "order_id": np.arange(1001, 1001 + n),
            "order_date": dates,
            "region": rng.choice(regions, n),
            "product": rng.choice(products, n),
            "quantity": rng.integers(1, 10, n).astype(float),
            "unit_price": rng.integers(500, 60000, n).astype(float),
        }
    )
    # add problems: missing values, duplicates, bad numbers
    df.loc[rng.choice(n, 10, replace=False), "quantity"] = np.nan
    df.loc[rng.choice(n, 8, replace=False), "unit_price"] = np.nan
    df.loc[rng.choice(n, 5, replace=False), "region"] = np.nan
    df.loc[3, "order_date"] = "not available"
    df = pd.concat([df, df.sample(8, random_state=1)], ignore_index=True)  # duplicates
    df.to_csv(path, index=False)
    print(f"Sample messy dataset created: {path}")


# ---------------------------------------------------------------- CLEANING
def clean_data(df):
    log = {"rows_before": len(df)}
    df = df.copy()

    # 1. standardize column names
    df.columns = [c.strip().lower().replace(" ", "_") for c in df.columns]

    # 2. remove exact duplicates
    before = len(df)
    df = df.drop_duplicates()
    log["duplicates_removed"] = before - len(df)

    # 3. fix inconsistent text (spaces, upper/lower case)
    text_cols = [
        c for c in df.columns
        if c != "order_date" and not pd.api.types.is_numeric_dtype(df[c])
    ]
    for col in text_cols:
        df[col] = df[col].astype("string").str.strip().str.title()

    # 4. convert dates (invalid dates become NaT)
    df["order_date"] = pd.to_datetime(df["order_date"], format="%d-%m-%Y", errors="coerce")

    # 5. missing values
    log["missing_before"] = int(df.isna().sum().sum())
    df = df.dropna(subset=["order_date"])  # a row without a date is unusable
    for col in ["quantity", "unit_price"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")
        df[col] = df[col].fillna(df[col].median())  # numeric -> median
    df["region"] = df["region"].fillna("Unknown")  # text -> "Unknown"
    df["product"] = df["product"].fillna("Unknown")
    log["missing_after"] = int(df.isna().sum().sum())

    # 6. final types and a new useful column
    df["quantity"] = df["quantity"].astype(int)
    df["total_sales"] = df["quantity"] * df["unit_price"]
    df["month"] = df["order_date"].dt.to_period("M").astype(str)

    log["rows_after"] = len(df)
    return df.reset_index(drop=True), log


# --------------------------------------------------------------- REPORTING
def make_charts(df):
    by_region = df.groupby("region")["total_sales"].sum().sort_values(ascending=False)
    by_product = df.groupby("product")["total_sales"].sum().sort_values(ascending=False)
    by_month = df.groupby("month")["total_sales"].sum()

    fig, ax = plt.subplots(figsize=(7, 4))
    by_region.plot(kind="bar", ax=ax, color="steelblue")
    ax.set_title("Total Sales by Region")
    ax.set_ylabel("Sales")
    plt.xticks(rotation=0)
    plt.tight_layout()
    plt.savefig("chart_sales_by_region.png", dpi=150)
    plt.close()

    fig, ax = plt.subplots(figsize=(7, 4))
    by_product.plot(kind="bar", ax=ax, color="darkorange")
    ax.set_title("Total Sales by Product")
    ax.set_ylabel("Sales")
    plt.xticks(rotation=0)
    plt.tight_layout()
    plt.savefig("chart_sales_by_product.png", dpi=150)
    plt.close()

    fig, ax = plt.subplots(figsize=(8, 4))
    by_month.plot(kind="line", marker="o", ax=ax, color="green")
    ax.set_title("Monthly Sales Trend")
    ax.set_ylabel("Sales")
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig("chart_monthly_trend.png", dpi=150)
    plt.close()

    return by_region, by_product, by_month


def make_excel_report(df, log, by_region, by_product, by_month):
    with pd.ExcelWriter("report.xlsx", engine="openpyxl") as writer:
        summary = pd.DataFrame(
            {
                "Metric": [
                    "Rows before cleaning",
                    "Duplicates removed",
                    "Rows after cleaning",
                    "Missing values before",
                    "Missing values after",
                    "Total sales",
                    "Average order value",
                ],
                "Value": [
                    log["rows_before"],
                    log["duplicates_removed"],
                    log["rows_after"],
                    log["missing_before"],
                    log["missing_after"],
                    round(df["total_sales"].sum(), 2),
                    round(df["total_sales"].mean(), 2),
                ],
            }
        )
        summary.to_excel(writer, sheet_name="Summary", index=False)
        by_region.reset_index().to_excel(writer, sheet_name="By Region", index=False)
        by_product.reset_index().to_excel(writer, sheet_name="By Product", index=False)
        by_month.reset_index().to_excel(writer, sheet_name="By Month", index=False)
        df.to_excel(writer, sheet_name="Cleaned Data", index=False)
        for ws in writer.book.worksheets:  # auto column width
            for col in ws.columns:
                width = max(len(str(c.value)) if c.value is not None else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(width + 3, 40)


def write_summary(df, log, by_region, by_product):
    text = f"""DATA CLEANING & REPORTING SUMMARY
=================================
Rows before cleaning : {log['rows_before']}
Duplicates removed   : {log['duplicates_removed']}
Rows after cleaning  : {log['rows_after']}
Missing values       : {log['missing_before']} -> {log['missing_after']}

Total sales          : {df['total_sales'].sum():,.2f}
Average order value  : {df['total_sales'].mean():,.2f}
Top region           : {by_region.index[0]} ({by_region.iloc[0]:,.2f})
Top product          : {by_product.index[0]} ({by_product.iloc[0]:,.2f})
"""
    with open("summary.txt", "w") as f:
        f.write(text)
    print(text)


# -------------------------------------------------------------------- MAIN
def main():
    path = sys.argv[1] if len(sys.argv) > 1 else RAW_FILE
    if not os.path.exists(path):
        create_messy_data(path)

    raw = pd.read_csv(path)
    print(f"Loaded {len(raw)} rows from {path}")

    df, log = clean_data(raw)
    df.to_csv("cleaned_data.csv", index=False)

    by_region, by_product, by_month = make_charts(df)
    make_excel_report(df, log, by_region, by_product, by_month)
    write_summary(df, log, by_region, by_product)
    print("Done. Created: cleaned_data.csv, report.xlsx, summary.txt and 3 chart PNG files")


if __name__ == "__main__":
    main()
