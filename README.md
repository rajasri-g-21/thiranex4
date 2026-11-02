# Data Cleaning & Reporting Automation

A Python script that automatically cleans a messy dataset and generates reports.

## What it does
- Standardizes column names and text (extra spaces, upper/lower case)
- Removes duplicate rows
- Handles missing values (median for numbers, "Unknown" for text, drops rows without a valid date)
- Fixes inconsistent date and number formats
- Generates an Excel report (`report.xlsx`) with Summary, By Region, By Product, By Month and Cleaned Data sheets
- Creates visual summaries (3 PNG charts) and a text summary (`summary.txt`)

## Project Structure
```
clean_report.py                 main script
requirements.txt                dependencies
messy_data.csv                  sample raw data (auto-generated if missing)
cleaned_data.csv                output: cleaned data
report.xlsx                     output: Excel report
summary.txt                     output: short text summary
chart_sales_by_region.png       output: chart
chart_sales_by_product.png      output: chart
chart_monthly_trend.png         output: chart
```

## How to Run
```bash
pip install -r requirements.txt
python clean_report.py
```
To use your own CSV file (needs columns order_id, order_date, region, product, quantity, unit_price):
```bash
python clean_report.py your_file.csv
```

## Sample Result
- 128 rows before, 119 after cleaning (8 duplicates and 1 invalid date row removed)
- 24 missing values fixed, 0 left

## Tech Stack
Python, pandas, NumPy, Matplotlib, openpyxl
