# Excel Master Database System
### Nexa Parts Supply - Automated Vendor Data Pipeline

> **Turning five messy vendor spreadsheets into one professional, audit-ready Excel master database - fully automated with Python.**

---

> ⚠️ **DISCLAIMER**
> This is a portfolio demo built with synthetic data. It is designed to reflect realistic business workflows without using real customer, vendor, company, or confidential data. All company names, vendor names, product names, SKUs, prices, and contact details are entirely fictional and generated for demonstration purposes only.

---

## 🧩 Business Problem

Industrial supply companies like Nexa Parts Supply receive product catalogues and pricing updates from dozens of vendors. Each vendor uses different column names, unit labels, category names, and data formats. Manually consolidating these into a master product database is slow, error-prone, and hard to audit.

This system automates the entire pipeline:
- Ingest messy vendor files in any format (CSV / XLSX)
- Standardise columns, categories, units, prices, and statuses
- Detect duplicate SKUs across vendors and flag data quality issues
- Produce a professional, audit-ready Excel workbook with one click

---

## 🚀 What the System Does

| Step | Script | What Happens |
|------|--------|-------------|
| 1 | `generate_vendor_data.py` | Creates 5 realistic messy vendor files (synthetic data) |
| 2 | `clean_and_merge.py` | Cleans, standardises, deduplicates, and validates all records |
| 3 | `build_workbook.py` | Builds the professional Excel master database workbook |

---

## 🛠 Tech Stack

| Tool | Purpose |
|------|---------|
| Python 3.x | Core scripting language |
| pandas | Data loading, cleaning, merging, aggregation |
| openpyxl | Excel workbook creation with formatting and validation |

No GUI. No macros. No manual steps. Pure Python automation.

---

## 📁 Project Structure

```
excel-master-database-system/
├── README.md
├── case-study.md
├── requirements.txt
├── .gitignore
│
├── config/
│   ├── column_map.json       # Raw column name -> canonical name mapping
│   ├── category_map.json     # Raw category -> canonical category mapping
│   └── unit_map.json         # Raw unit label -> canonical UoM mapping
│
├── data/
│   ├── raw/                  # Input: 5 messy vendor files (CSV/XLSX)
│   ├── processed/            # Output: master_products.csv, vendor_imports.csv, workbook
│   └── reports/              # Output: duplicate_review.csv, data_issues.csv, import_log.csv
│
├── scripts/
│   ├── generate_vendor_data.py   # Step 1: Generate synthetic vendor files
│   ├── clean_and_merge.py        # Step 2: Clean, merge, validate
│   └── build_workbook.py         # Step 3: Build the Excel workbook
│
└── screenshots/
    └── README.md             # Screenshot capture guide
```

---

## ⚡ Quick Start (Windows PowerShell)

### 1. Clone and set up

```powershell
# Navigate to project
cd excel-master-database-system

# Install dependencies
pip install -r requirements.txt
```

### 2. Generate synthetic vendor files

```powershell
python scripts/generate_vendor_data.py --seed 42
```

### 3. Run the cleaning pipeline

```powershell
python scripts/clean_and_merge.py
```

### 4. Build the Excel workbook

```powershell
python scripts/build_workbook.py
```

### 5. Open the result

```powershell
Start-Process "data\processed\nexa_parts_master_database.xlsx"
```

---

## 📂 Output Files

| File | Location | Description |
|------|----------|-------------|
| `master_products.csv` | `data/processed/` | Clean, deduplicated product master |
| `vendor_imports.csv` | `data/processed/` | All raw import records (canonical format) |
| `duplicate_review.csv` | `data/reports/` | Duplicate SKU groups for review |
| `data_issues.csv` | `data/reports/` | All detected data quality issues |
| `import_log.csv` | `data/reports/` | Per-vendor import run summary |
| `nexa_parts_master_database.xlsx` | `data/processed/` | **The final professional workbook** |

---

## 📊 Workbook Sheets

| Sheet | Description |
|-------|-------------|
| **README** | In-workbook guide: sheet descriptions, quality rules, disclaimer |
| **Master Products** | Cleaned master with SKU, name, category, UoM, cost, price, margin, stock, status |
| **Vendor Imports** | Every raw import record with source tracking and canonical columns |
| **Duplicate Review** | Duplicate SKU groups across vendors; SELECTED vs DUPLICATE highlighted |
| **Data Issues** | All detected quality issues with HIGH/MEDIUM/LOW severity ratings |
| **Import Log** | Per-vendor run summary: row counts, duplicate counts, issue counts |
| **Category Summary** | Aggregated tables: by category, by vendor, status breakdown, top 10 by value |
| **Inventory Dashboard** | KPI cards + summary tables: total products, active, discontinued, inventory value |

---

## ✅ Data Quality Rules Applied

| Rule | Detail |
|------|--------|
| **Column Standardisation** | 50+ raw column name variants mapped to 9 canonical names via `column_map.json` |
| **SKU Cleaning** | Whitespace removed, uppercased, blanks flagged as HIGH severity issues |
| **Price Cleaning** | Currency symbols (`$`, `USD`) stripped; blanks, zeros, negatives detected |
| **Quantity Cleaning** | Non-numeric values (`unknown`, `out`, `five`, `TBD`) detected and flagged |
| **Status Normalisation** | 15+ status variants (`active`, `Y`, `1`, `In Stock`, etc.) -> `Active`/`Discontinued` |
| **Category Mapping** | 20+ raw category variants -> 12 canonical categories via `category_map.json` |
| **Unit Mapping** | 10+ unit variants (`pcs`, `piece`, `pieces`, `ea`, `each`) -> 9 canonical UoMs |
| **Cost > Sell Detection** | Rows where cost price exceeds selling price flagged as HIGH severity |
| **Duplicate Detection** | Cross-vendor SKU deduplication; best record selected by completeness score |
| **Issue Severity** | HIGH = blocking data error, MEDIUM = data gap, LOW = informational |

---

## 🖼 Screenshots

See the [`screenshots/README.md`](screenshots/README.md) for capture instructions.

Add your screenshots to the `screenshots/` directory and link them here.

---

## 💼 Skills Demonstrated

- **Python automation** - end-to-end ETL pipeline without manual steps
- **pandas** - data loading, column mapping, cleaning, aggregation, deduplication
- **openpyxl** - professional workbook creation with styles, conditional formatting, data validation
- **Data quality engineering** - issue detection, severity classification, audit trails
- **Config-driven design** - JSON mapping files decouple logic from data
- **Business process thinking** - workflow mirrors real procurement/supply chain operations
- **Documentation** - polished README, in-workbook guide, case study, and screenshot guide

---

*This is a portfolio demo project. All data is synthetic.*
