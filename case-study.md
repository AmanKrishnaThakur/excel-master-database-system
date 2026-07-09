# Case Study: Automated Vendor Data Consolidation
### Nexa Parts Supply - Excel Master Database System

> ⚠️ **DISCLAIMER**
> This is a portfolio demo built with synthetic data. It is designed to reflect realistic business workflows without using real customer, vendor, company, or confidential data. All company names, vendor names, product names, SKUs, prices, and contact details are entirely fictional and generated for demonstration purposes only. This is not paid client work.

---

## The Problem

Industrial supply companies receive product data from multiple vendors, and no two vendors format their catalogues the same way.

**Nexa Parts Supply** (fictional company) sources industrial parts from five vendor suppliers. Each quarter, they receive updated price lists and product catalogues - but each vendor uses different:

- Column names (`sku` vs `part_number` vs `article_no` vs `ref_no`)
- Category labels (`Fasteners` vs `Bolts & Nuts` vs `Hardware Fasteners`)
- Unit labels (`ea` vs `each` vs `pcs` vs `pieces`)
- Status values (`Active` vs `In Stock` vs `Y` vs `1`)
- Price formats (`$12.50` vs `USD 12.50` vs `12.5` vs ` 12.50 `)

On top of inconsistent formatting, the raw files contained:
- Blank SKUs and missing product names
- Zero and negative prices
- Cost prices exceeding selling prices (margin inversion)
- Non-numeric stock quantities (`unknown`, `out`, `five`)
- Negative stock quantities (`-5`, `-1`)
- Duplicate SKUs across multiple vendors
- Discontinued products mixed with active ones
- Junk columns (`internal_ref`, `buyer_notes`, `legacy_code`, etc.)

Without automation, consolidating these files took hours of manual Excel work - and was still error-prone.

---

## The Solution

A three-script Python pipeline that:

1. **Ingests** all vendor files (CSV and XLSX)
2. **Standardises** columns, categories, units, and statuses using JSON config maps
3. **Detects and flags** duplicate SKUs, data quality issues with severity ratings
4. **Produces** a professional Excel workbook with 8 structured sheets

The entire pipeline runs in under 30 seconds.

---

## Technical Approach

### Config-Driven Column Mapping

Rather than hard-coding per-vendor logic, the system uses `config/column_map.json` to map 50+ raw column name variants to 9 canonical names:

```
"sku"          <- sku, part_number, article_no, ref_no, item_no, item_code, ...
"product_name" <- name, description, item_desc, product_title, item_name, ...
"cost_price"   <- cost, unit_cost, purchase_price, buy_price, wholesale_price, ...
```

This means adding a new vendor requires only a config update - no code changes.

### Category and Unit Normalisation

`config/category_map.json` maps 20+ raw category labels to 12 canonical categories.
`config/unit_map.json` maps 10+ unit variants to 9 canonical UoM values.

Examples:
- `"Bolts & Nuts"`, `"Hardware Fasteners"`, `"Fixings"` -> `"Fasteners"`
- `"pcs"`, `"piece"`, `"pieces"`, `"ea"`, `"each"` -> `"EA"`

### Price and Quantity Cleaning

Prices are cleaned with a regex that strips currency symbols and whitespace, then parsed to float. Blanks, zeros, and negatives are detected as issues.

Stock quantities handle text values like `"unknown"`, `"out"`, `"five"`, `"TBD"` - all flagged as issues with severity MEDIUM.

### Duplicate Detection

SKUs appearing in more than one vendor file are collected into duplicate groups. For each group, the system selects the "best" record using a completeness score:
- Active status preferred over Discontinued
- Non-null valid selling price: +10 points
- Non-null valid cost price: +5 points
- Non-null valid stock qty: +2 points

The full duplicate group is written to `duplicate_review.csv` and the **Duplicate Review** sheet for human review.

### Issue Severity Classification

| Severity | Examples |
|----------|---------|
| **HIGH** | Blank SKU, zero/negative price, cost > selling price, negative qty |
| **MEDIUM** | Blank cost price, blank selling price, non-numeric stock qty |
| **LOW** | Unknown status, uncategorised product |

### Excel Workbook Features

The workbook is built programmatically using `openpyxl` with no manual steps:

- **Styled headers** - navy header rows with white bold text
- **Alternating row fills** - light blue/white for readability
- **Freeze panes** - header rows stay visible while scrolling
- **Auto-filters** - all data sheets are filterable
- **Conditional formatting**:
  - Discontinued products -> light purple italic
  - Low-stock items -> orange highlight
  - HIGH issues -> red background
  - MEDIUM issues -> amber background
  - LOW issues -> light green background
  - SELECTED duplicates -> green highlight
  - DUPLICATE rows -> orange highlight
- **Dropdown validations**: `status`, `unit_of_measure`, `resolved` (for issue tracking)
- **KPI Dashboard** - 8 KPI cards with colour-coded metrics
- **Summary tables** - by category, by vendor, status breakdown, top 10 by stock value

---

## Results (Seed 42 Run)

The pipeline was run with `--seed 42` to produce a fully reproducible synthetic dataset:

| Metric | Value |
|--------|-------|
| Raw vendor files processed | 5 |
| Total raw rows imported | 324 |
| Master products (after dedup) | 211 |
| Active products | 192 |
| Discontinued products | 19 |
| Duplicate groups detected | 54 |
| Total duplicate rows reviewed | 167 |
| Total data issues found | 89 |
| HIGH severity issues | 34 |
| MEDIUM severity issues | 55 |
| LOW severity issues | 0 |
| Low-stock products flagged | 27 |
| Estimated inventory value | $1,189,025 |

---

## Vendor Files Summary

| Vendor | File Format | Primary Categories | Rows (seed 42) |
|--------|-------------|-------------------|-------------|
| AlphaPro Parts | CSV | Fasteners, Filtration, Tools | 87 |
| DeltaForge Supply | XLSX | Bearings, Drive Components, Structural | 69 |
| Meridian Components | CSV | Seals & Gaskets, Bearings, Lubrication | 53 |
| Stratos Industrial | XLSX | Electrical Components, Couplings | 61 |
| Zenitech Catalog | CSV | Hydraulics, Pneumatics | 54 |

Each file intentionally uses different column names, category labels, unit formats, and contains its own pattern of data issues.

---

## Skills Demonstrated

| Skill | How |
|-------|-----|
| Python ETL automation | 3-script pipeline processes raw -> clean -> workbook end-to-end |
| pandas | Load mixed formats, column mapping, cleaning, deduplication, aggregation |
| openpyxl | Programmatic workbook with styles, conditional formatting, data validation |
| Config-driven design | JSON maps decouple business rules from code |
| Data quality engineering | Issue detection, severity classification, audit trail |
| Business process thinking | Mirrors real procurement/supply chain workflows |
| Reproducible outputs | `--seed` parameter ensures identical runs |
| Documentation | README, case study, in-workbook guide, screenshot guide |

---

## What This Would Look Like in Production

In a real deployment, this pipeline could be extended to:
- Pull vendor files automatically from email attachments, SFTP, or SharePoint
- Run on a scheduled trigger (weekly, monthly) without manual intervention
- Push the master database to a SQL database or ERP system
- Send an email summary report with issue highlights
- Track historical price changes across vendor updates

The foundation built here - config-driven column mapping, quality scoring, audit trail, structured reporting - would translate directly to a production data pipeline.

---

*This project is a portfolio demo. All company names, vendors, products, SKUs, and prices are entirely fictional.*
