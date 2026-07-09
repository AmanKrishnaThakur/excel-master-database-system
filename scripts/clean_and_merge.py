"""
clean_and_merge.py
==================
Nexa Parts Supply - Excel Master Database System
Portfolio Demo | Synthetic Data Only

Reads all raw vendor files, standardises columns, cleans data,
detects duplicates and issues, then outputs processed CSV files.

Usage:
    python scripts/clean_and_merge.py
"""

import json
import re
import sys
import warnings
from pathlib import Path
from datetime import datetime

import pandas as pd

warnings.filterwarnings("ignore")

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
ROOT = Path(__file__).resolve().parent.parent
RAW_DIR    = ROOT / "data" / "raw"
PROC_DIR   = ROOT / "data" / "processed"
RPT_DIR    = ROOT / "data" / "reports"
CFG_DIR    = ROOT / "config"

for d in [PROC_DIR, RPT_DIR]:
    d.mkdir(parents=True, exist_ok=True)

# ---------------------------------------------------------------------------
# Load config maps
# ---------------------------------------------------------------------------

def load_json(path):
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    # Remove comment keys
    return {k: v for k, v in data.items() if not k.startswith("_")}


COLUMN_MAP   = load_json(CFG_DIR / "column_map.json")
CATEGORY_MAP = load_json(CFG_DIR / "category_map.json")
UNIT_MAP     = load_json(CFG_DIR / "unit_map.json")

# Build reverse lookup: raw_value -> canonical
def build_reverse_map(mapping):
    rev = {}
    for canonical, variants in mapping.items():
        for v in variants:
            rev[v.strip().lower()] = canonical
    return rev

CATEGORY_LOOKUP = build_reverse_map(CATEGORY_MAP)
UNIT_LOOKUP     = build_reverse_map(UNIT_MAP)

# Build column map: lower raw col name -> canonical col name
COLUMN_LOOKUP = {}
for canonical, variants in COLUMN_MAP.items():
    for v in variants:
        COLUMN_LOOKUP[v.strip().lower()] = canonical


# ---------------------------------------------------------------------------
# Canonical columns
# ---------------------------------------------------------------------------
CANONICAL_COLS = [
    "sku", "product_name", "category", "unit_of_measure",
    "cost_price", "selling_price", "stock_qty", "status", "vendor_name",
]

# ---------------------------------------------------------------------------
# Vendor file registry
# ---------------------------------------------------------------------------
VENDOR_FILES = [
    (RAW_DIR / "vendor_alphapro_parts.csv",     "csv",  "AlphaPro Parts"),
    (RAW_DIR / "vendor_deltaforge_supply.xlsx",  "xlsx", "DeltaForge Supply"),
    (RAW_DIR / "vendor_meridian_components.csv", "csv",  "Meridian Components"),
    (RAW_DIR / "vendor_stratos_industrial.xlsx", "xlsx", "Stratos Industrial"),
    (RAW_DIR / "vendor_zenitech_catalog.csv",    "csv",  "Zenitech Catalog"),
]

# ---------------------------------------------------------------------------
# Helper: clean price string → float or NaN
# ---------------------------------------------------------------------------
PRICE_RE = re.compile(r"[^\d.\-]")

def parse_price(val):
    if pd.isna(val) or str(val).strip() == "":
        return float("nan")
    cleaned = PRICE_RE.sub("", str(val).strip())
    if cleaned in ("", "."):
        return float("nan")
    try:
        return float(cleaned)
    except ValueError:
        return float("nan")


# ---------------------------------------------------------------------------
# Helper: clean quantity → int or NaN
# ---------------------------------------------------------------------------
def parse_qty(val):
    if pd.isna(val) or str(val).strip() == "":
        return float("nan")
    s = str(val).strip().lower()
    if s in ("unknown", "out", "tbd", "n/a", "none", ""):
        return float("nan")
    try:
        f = float(s)
        return int(f) if f == int(f) else float("nan")
    except ValueError:
        return float("nan")


# ---------------------------------------------------------------------------
# Helper: normalise status → Active / Discontinued / Unknown
# ---------------------------------------------------------------------------
ACTIVE_VALS = {"active", "in stock", "available", "1", "y", "yes"}
DISC_VALS   = {"discontinued", "inactive", "obsolete", "0", "n", "no"}

def parse_status(val):
    if pd.isna(val) or str(val).strip() == "":
        return "Unknown"
    s = str(val).strip().lower()
    if s in ACTIVE_VALS:
        return "Active"
    if s in DISC_VALS:
        return "Discontinued"
    return "Unknown"


# ---------------------------------------------------------------------------
# Helper: normalise SKU
# ---------------------------------------------------------------------------
SKU_RE = re.compile(r"\s+")

def clean_sku(val):
    if pd.isna(val) or str(val).strip() == "":
        return ""
    return SKU_RE.sub("", str(val).strip()).upper()


# ---------------------------------------------------------------------------
# Helper: clean vendor name
# ---------------------------------------------------------------------------
VENDOR_CANONICAL = {
    "alphapro":  "AlphaPro Parts",
    "deltaforge": "DeltaForge Supply",
    "meridian":  "Meridian Components",
    "stratos":   "Stratos Industrial",
    "zenitech":  "Zenitech Catalog",
}

def clean_vendor(val, default):
    if pd.isna(val) or str(val).strip() == "":
        return default
    s = str(val).strip().lower()
    for key, canonical in VENDOR_CANONICAL.items():
        if key in s:
            return canonical
    return str(val).strip().title()


# ---------------------------------------------------------------------------
# Helper: clean product name
# ---------------------------------------------------------------------------
def clean_name(val):
    if pd.isna(val) or str(val).strip() == "":
        return ""
    return " ".join(str(val).strip().split())  # collapse whitespace + strip


# ---------------------------------------------------------------------------
# Step 1: Load & map columns
# ---------------------------------------------------------------------------

def load_vendor_file(path, fmt, vendor_label):
    if fmt == "csv":
        df = pd.read_csv(path, dtype=str, keep_default_na=False)
    else:
        df = pd.read_excel(path, dtype=str, keep_default_na=False)

    # Normalise column names → canonical
    rename = {}
    for col in df.columns:
        mapped = COLUMN_LOOKUP.get(col.strip().lower())
        if mapped:
            rename[col] = mapped
    df.rename(columns=rename, inplace=True)

    # Drop columns not in canonical set (junk columns)
    keep = [c for c in df.columns if c in CANONICAL_COLS]
    df = df[keep].copy()

    # Ensure all canonical cols present
    for col in CANONICAL_COLS:
        if col not in df.columns:
            df[col] = ""

    df["_source_vendor"] = vendor_label
    df["_source_file"]   = path.name
    df["_row_index"]     = range(1, len(df) + 1)

    return df[CANONICAL_COLS + ["_source_vendor", "_source_file", "_row_index"]]


# ---------------------------------------------------------------------------
# Step 2: Clean each column
# ---------------------------------------------------------------------------

def clean_dataframe(df):
    df = df.copy()

    df["sku"]           = df["sku"].apply(clean_sku)
    df["product_name"]  = df["product_name"].apply(clean_name)

    # Category normalisation
    df["category"] = df["category"].apply(
        lambda v: CATEGORY_LOOKUP.get(str(v).strip().lower(), str(v).strip().title() if str(v).strip() else "Uncategorised")
    )

    # Unit normalisation
    df["unit_of_measure"] = df["unit_of_measure"].apply(
        lambda v: UNIT_LOOKUP.get(str(v).strip().lower(), str(v).strip().upper() if str(v).strip() else "EA")
    )

    df["cost_price"]    = df["cost_price"].apply(parse_price)
    df["selling_price"] = df["selling_price"].apply(parse_price)
    df["stock_qty"]     = df["stock_qty"].apply(parse_qty)
    df["status"]        = df["status"].apply(parse_status)
    df["vendor_name"]   = df.apply(
        lambda r: clean_vendor(r["vendor_name"], r["_source_vendor"]), axis=1
    )

    return df


# ---------------------------------------------------------------------------
# Step 3: Detect data issues
# ---------------------------------------------------------------------------

def detect_issues(df):
    """
    Returns a DataFrame of data quality issues with columns:
    source_file, row_index, sku, product_name, issue_type, issue_detail, severity
    """
    issues = []

    for _, row in df.iterrows():
        src  = row["_source_file"]
        ridx = row["_row_index"]
        sku  = row["sku"]
        name = row["product_name"]

        def add(issue_type, detail, severity):
            issues.append({
                "source_file":   src,
                "row_index":     ridx,
                "sku":           sku,
                "product_name":  name,
                "issue_type":    issue_type,
                "issue_detail":  detail,
                "severity":      severity,
            })

        # Missing SKU
        if not sku:
            add("Missing SKU", "SKU is blank or missing", "HIGH")

        # Missing product name
        if not name:
            add("Missing Product Name", "Product name is blank or missing", "HIGH")

        # Missing/invalid cost price
        cp = row["cost_price"]
        if pd.isna(cp):
            add("Missing Cost Price", "Cost price is blank or unparseable", "MEDIUM")
        elif cp <= 0:
            add("Invalid Cost Price", f"Cost price is {cp} (zero or negative)", "HIGH")

        # Missing/invalid selling price
        sp = row["selling_price"]
        if pd.isna(sp):
            add("Missing Selling Price", "Selling price is blank or unparseable", "MEDIUM")
        elif sp <= 0:
            add("Invalid Selling Price", f"Selling price is {sp} (zero or negative)", "HIGH")

        # Cost > Selling price
        if not pd.isna(cp) and not pd.isna(sp) and cp > 0 and sp > 0 and cp > sp:
            add("Cost Exceeds Selling Price", f"Cost={cp:.2f} > Selling={sp:.2f}", "HIGH")

        # Missing/invalid stock qty
        qty = row["stock_qty"]
        if pd.isna(qty):
            add("Invalid Stock Quantity", "Stock qty is blank or non-numeric", "MEDIUM")
        elif qty < 0:
            add("Negative Stock Quantity", f"Stock qty is {int(qty)}", "HIGH")

        # Unknown status
        if row["status"] == "Unknown":
            add("Unknown Status", "Status could not be mapped to Active/Discontinued", "LOW")

        # Uncategorised
        if row["category"] == "Uncategorised":
            add("Uncategorised Product", "Category could not be mapped", "LOW")

    return pd.DataFrame(issues)


# ---------------------------------------------------------------------------
# Step 4: Deduplicate & flag duplicates
# ---------------------------------------------------------------------------

def detect_duplicates(df):
    """
    Returns:
    - master_df: one row per SKU (best record chosen)
    - dup_review_df: all duplicate groups with review notes
    """
    # Only consider rows with valid SKU
    has_sku  = df[df["sku"] != ""].copy()
    no_sku   = df[df["sku"] == ""].copy()

    dup_groups = []
    master_rows = []
    seen_skus   = set()

    grouped = has_sku.groupby("sku", sort=False)

    for sku, group in grouped:
        if len(group) == 1:
            master_rows.append(group.iloc[0])
            seen_skus.add(sku)
        else:
            # Multiple records for same SKU — duplicate group
            seen_skus.add(sku)

            # Pick "best" record: Active preferred, then highest selling price
            active = group[group["status"] == "Active"]
            pool   = active if len(active) > 0 else group

            # Score: prefer non-null selling price, non-null cost, highest qty
            def score(row):
                s = 0
                if not pd.isna(row["selling_price"]) and row["selling_price"] > 0:
                    s += 10
                if not pd.isna(row["cost_price"]) and row["cost_price"] > 0:
                    s += 5
                if not pd.isna(row["stock_qty"]) and row["stock_qty"] >= 0:
                    s += 2
                return s

            pool = pool.copy()
            pool["_score"] = pool.apply(score, axis=1)
            best = pool.sort_values("_score", ascending=False).iloc[0]
            master_rows.append(best)

            # Build duplicate review rows
            for i, (_, r) in enumerate(group.iterrows()):
                dup_groups.append({
                    "sku":             sku,
                    "duplicate_group": f"DUP-{sku}",
                    "record_number":   i + 1,
                    "source_file":     r["_source_file"],
                    "vendor_name":     r["vendor_name"],
                    "product_name":    r["product_name"],
                    "category":        r["category"],
                    "cost_price":      r["cost_price"],
                    "selling_price":   r["selling_price"],
                    "stock_qty":       r["stock_qty"],
                    "status":          r["status"],
                    "review_note":     "SELECTED — kept as master" if i == 0 else "DUPLICATE — review required",
                })

    # Rows with no SKU — include but flag
    for _, r in no_sku.iterrows():
        master_rows.append(r)

    master_df    = pd.DataFrame(master_rows).reset_index(drop=True)
    dup_review_df = pd.DataFrame(dup_groups)

    return master_df, dup_review_df


# ---------------------------------------------------------------------------
# Step 5: Build master_products and vendor_imports
# ---------------------------------------------------------------------------

def build_master_products(df):
    """
    Final clean master product table.
    Assigns a sequential master_id, computes margin, value.
    """
    df = df.copy()

    # Drop internal tracking cols
    df.drop(columns=["_source_vendor", "_source_file", "_row_index", "_score"],
            errors="ignore", inplace=True)

    # Assign master IDs
    df.insert(0, "master_id", [f"MPD-{i:04d}" for i in range(1, len(df) + 1)])

    # Compute derived fields
    df["margin_pct"] = df.apply(
        lambda r: round((r["selling_price"] - r["cost_price"]) / r["selling_price"] * 100, 1)
        if not pd.isna(r["selling_price"]) and not pd.isna(r["cost_price"])
           and r["selling_price"] > 0 and r["cost_price"] > 0
        else float("nan"),
        axis=1,
    )

    df["stock_value"] = df.apply(
        lambda r: round(r["selling_price"] * r["stock_qty"], 2)
        if not pd.isna(r["selling_price"]) and not pd.isna(r["stock_qty"])
           and r["selling_price"] > 0 and r["stock_qty"] >= 0
        else float("nan"),
        axis=1,
    )

    df["low_stock_flag"] = df["stock_qty"].apply(
        lambda q: "YES" if not pd.isna(q) and q < 20 else "NO"
    )

    return df


def build_vendor_imports(all_raw):
    """
    One row per raw import record with clean canonical cols + source tracking.
    """
    df = all_raw.copy()
    df.drop(columns=["_score"], errors="ignore", inplace=True)
    df.insert(0, "import_id", [f"IMP-{i:04d}" for i in range(1, len(df) + 1)])
    df.rename(columns={"_source_file": "source_file", "_source_vendor": "source_vendor",
                        "_row_index": "source_row"}, inplace=True)
    return df


# ---------------------------------------------------------------------------
# Step 6: Build import log
# ---------------------------------------------------------------------------

def build_import_log(all_frames_meta):
    """
    all_frames_meta: list of (vendor_label, filename, raw_row_count, clean_row_count, issue_count)
    """
    rows = []
    run_ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for vendor, filename, raw_cnt, clean_cnt, issue_cnt, dup_cnt in all_frames_meta:
        rows.append({
            "run_timestamp":    run_ts,
            "vendor_name":      vendor,
            "source_file":      filename,
            "raw_row_count":    raw_cnt,
            "clean_row_count":  clean_cnt,
            "duplicate_rows":   dup_cnt,
            "issue_count":      issue_cnt,
            "status":           "Processed",
            "notes":            "",
        })
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# Main pipeline
# ---------------------------------------------------------------------------

def main():
    print("=" * 62)
    print("  Nexa Parts Supply - Data Clean & Merge Pipeline")
    print("  Portfolio Demo | Synthetic Data Only")
    print("=" * 62)

    all_dfs   = []
    meta_list = []

    # --- Load & map ---
    print("\n  [1/5]  Loading vendor files...")
    for path, fmt, label in VENDOR_FILES:
        if not path.exists():
            print(f"  [ERROR] {path.name} not found - run generate_vendor_data.py first")
            sys.exit(1)
        df_raw = load_vendor_file(path, fmt, label)
        df_clean = clean_dataframe(df_raw)
        all_dfs.append((label, path.name, df_clean))
        print(f"     [OK]  {path.name:<42} {len(df_clean):>4} rows loaded")

    # --- Combine ---
    print("\n  [2/5]  Combining all records...")
    combined = pd.concat([d for _, _, d in all_dfs], ignore_index=True)
    print(f"         Total combined rows: {len(combined)}")

    # --- Detect issues (pre-dedup) ---
    print("\n  [3/5]  Detecting data quality issues...")
    issues_df = detect_issues(combined)
    print(f"         Issues found: {len(issues_df)}")

    # --- Deduplicate ---
    print("\n  [4/5]  Detecting duplicates & selecting master records...")
    master_df_raw, dup_review_df = detect_duplicates(combined)
    master_df = build_master_products(master_df_raw)
    vendor_imports_df = build_vendor_imports(combined)

    dup_count_total = len(dup_review_df)
    dup_sku_groups  = dup_review_df["sku"].nunique() if len(dup_review_df) > 0 else 0
    print(f"         Duplicate groups: {dup_sku_groups}")
    print(f"         Total duplicate rows: {dup_count_total}")
    print(f"         Master products after dedup: {len(master_df)}")

    # Build per-vendor issue/dup counts for log
    for label, fname, df_clean in all_dfs:
        file_issues = issues_df[issues_df["source_file"] == fname]
        file_dups   = dup_review_df[dup_review_df["source_file"] == fname] if len(dup_review_df) > 0 else pd.DataFrame()
        meta_list.append((label, fname, len(df_clean), len(df_clean), len(file_issues), len(file_dups)))

    import_log_df = build_import_log(meta_list)

    # --- Save outputs ---
    print("\n  [5/5]  Saving processed files...")

    out_map = [
        (PROC_DIR / "master_products.csv",  master_df,          "Master Products"),
        (PROC_DIR / "vendor_imports.csv",   vendor_imports_df,  "Vendor Imports"),
        (RPT_DIR  / "duplicate_review.csv", dup_review_df,      "Duplicate Review"),
        (RPT_DIR  / "data_issues.csv",      issues_df,          "Data Issues"),
        (RPT_DIR  / "import_log.csv",       import_log_df,      "Import Log"),
    ]

    for path, df, label in out_map:
        df.to_csv(path, index=False)
        print(f"     [OK]  {path.name:<40} {len(df):>4} rows  ->  {label}")

    # --- Summary ---
    active_ct  = (master_df["status"] == "Active").sum()
    disc_ct    = (master_df["status"] == "Discontinued").sum()
    low_stock  = (master_df["low_stock_flag"] == "YES").sum()
    high_sev   = (issues_df["severity"] == "HIGH").sum() if len(issues_df) > 0 else 0
    med_sev    = (issues_df["severity"] == "MEDIUM").sum() if len(issues_df) > 0 else 0
    low_sev    = (issues_df["severity"] == "LOW").sum() if len(issues_df) > 0 else 0
    total_val  = master_df["stock_value"].sum()

    print(f"\n  " + "-" * 58)
    print(f"  PIPELINE SUMMARY")
    print(f"  " + "-" * 58)
    print(f"  Raw rows imported       : {len(combined)}")
    print(f"  Master products         : {len(master_df)}")
    print(f"    - Active              : {active_ct}")
    print(f"    - Discontinued        : {disc_ct}")
    print(f"  Duplicate groups        : {dup_sku_groups}")
    print(f"  Data issues             : {len(issues_df)}")
    print(f"    - HIGH severity       : {high_sev}")
    print(f"    - MEDIUM severity     : {med_sev}")
    print(f"    - LOW severity        : {low_sev}")
    print(f"  Low-stock products      : {low_stock}")
    print(f"  Est. inventory value    : ${total_val:,.2f}")
    print(f"  " + "-" * 58)
    print("  All outputs saved successfully.")
    print("=" * 62)


if __name__ == "__main__":
    main()
