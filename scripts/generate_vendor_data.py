"""
generate_vendor_data.py
=======================
Nexa Parts Supply - Excel Master Database System
Portfolio Demo | Synthetic Data Only

Generates 5 messy vendor product/pricing files into data/raw/.
All company names, vendor names, product names, SKUs, prices, and
contact details are entirely fictional and generated for demonstration purposes only.

Usage:
    python scripts/generate_vendor_data.py --seed 42
"""

import argparse
import os
import random
import sys
from pathlib import Path

import pandas as pd

# ---------------------------------------------------------------------------
# Path setup
# ---------------------------------------------------------------------------
ROOT = Path(__file__).resolve().parent.parent
RAW_DIR = ROOT / "data" / "raw"
RAW_DIR.mkdir(parents=True, exist_ok=True)


# ---------------------------------------------------------------------------
# Product Universe
# ---------------------------------------------------------------------------

PRODUCT_UNIVERSE = [
    # Fasteners
    ("NXP-F001", "Hex Bolt M8 x 40mm Grade 8.8", "Fasteners", "EA", 0.38, 0.85, "AlphaPro Parts"),
    ("NXP-F002", "Hex Bolt M10 x 50mm Grade 8.8", "Fasteners", "EA", 0.55, 1.20, "AlphaPro Parts"),
    ("NXP-F003", "Hex Bolt M12 x 60mm Grade 10.9", "Fasteners", "EA", 0.72, 1.60, "AlphaPro Parts"),
    ("NXP-F004", "Hex Nut M8 DIN 934", "Fasteners", "EA", 0.12, 0.28, "AlphaPro Parts"),
    ("NXP-F005", "Hex Nut M10 DIN 934", "Fasteners", "EA", 0.18, 0.40, "AlphaPro Parts"),
    ("NXP-F006", "Hex Nut M12 DIN 934", "Fasteners", "EA", 0.25, 0.55, "AlphaPro Parts"),
    ("NXP-F007", "Socket Head Cap Screw M6 x 25mm", "Fasteners", "EA", 0.22, 0.50, "AlphaPro Parts"),
    ("NXP-F008", "Socket Head Cap Screw M8 x 30mm", "Fasteners", "EA", 0.35, 0.78, "AlphaPro Parts"),
    ("NXP-F009", "Flat Washer M8 DIN 125", "Fasteners", "EA", 0.05, 0.12, "AlphaPro Parts"),
    ("NXP-F010", "Flat Washer M10 DIN 125", "Fasteners", "EA", 0.08, 0.18, "AlphaPro Parts"),
    ("NXP-F011", "Spring Washer M8 DIN 127", "Fasteners", "EA", 0.07, 0.16, "AlphaPro Parts"),
    ("NXP-F012", "Stainless Steel Screw M5 x 20mm", "Fasteners", "EA", 0.30, 0.68, "AlphaPro Parts"),
    ("NXP-F013", "Self-Tapping Screw 4.2 x 19mm", "Fasteners", "BOX", 4.50, 9.80, "AlphaPro Parts"),
    ("NXP-F014", "Threaded Rod M10 x 1000mm", "Fasteners", "EA", 3.20, 7.00, "DeltaForge Supply"),
    ("NXP-F015", "Nylon Locking Nut M8 DIN 985", "Fasteners", "EA", 0.15, 0.35, "AlphaPro Parts"),
    # Bearings
    ("NXP-B001", "Deep Groove Ball Bearing 6205-2RS", "Bearings", "EA", 3.80, 8.50, "DeltaForge Supply"),
    ("NXP-B002", "Deep Groove Ball Bearing 6206-2RS", "Bearings", "EA", 4.50, 9.90, "DeltaForge Supply"),
    ("NXP-B003", "Deep Groove Ball Bearing 6208-2RS", "Bearings", "EA", 6.20, 13.50, "DeltaForge Supply"),
    ("NXP-B004", "Cylindrical Roller Bearing NU205", "Bearings", "EA", 12.50, 27.00, "DeltaForge Supply"),
    ("NXP-B005", "Taper Roller Bearing 30205", "Bearings", "EA", 9.80, 21.50, "DeltaForge Supply"),
    ("NXP-B006", "Thrust Ball Bearing 51105", "Bearings", "EA", 5.60, 12.20, "DeltaForge Supply"),
    ("NXP-B007", "Angular Contact Ball Bearing 7205B", "Bearings", "EA", 11.20, 24.50, "DeltaForge Supply"),
    ("NXP-B008", "Needle Roller Bearing HK2012", "Bearings", "EA", 2.90, 6.40, "Meridian Components"),
    ("NXP-B009", "Self-Aligning Ball Bearing 1205K", "Bearings", "EA", 8.40, 18.50, "Meridian Components"),
    ("NXP-B010", "Pillow Block Bearing UCP205", "Bearings", "EA", 14.20, 31.00, "DeltaForge Supply"),
    # Seals & Gaskets
    ("NXP-S001", "Oil Seal 30 x 52 x 10mm", "Seals & Gaskets", "EA", 1.80, 4.00, "Meridian Components"),
    ("NXP-S002", "Oil Seal 40 x 62 x 10mm", "Seals & Gaskets", "EA", 2.20, 4.80, "Meridian Components"),
    ("NXP-S003", "O-Ring 50 x 3mm NBR 70", "Seals & Gaskets", "EA", 0.45, 1.00, "Meridian Components"),
    ("NXP-S004", "O-Ring 60 x 4mm NBR 70", "Seals & Gaskets", "EA", 0.65, 1.45, "Meridian Components"),
    ("NXP-S005", "Rubber Gasket Sheet 300 x 300mm", "Seals & Gaskets", "EA", 4.80, 10.50, "Meridian Components"),
    ("NXP-S006", "PTFE Gasket Sheet 200 x 200mm", "Seals & Gaskets", "EA", 6.20, 13.50, "Meridian Components"),
    ("NXP-S007", "Shaft Seal 35 x 55 x 8mm", "Seals & Gaskets", "EA", 1.60, 3.50, "Meridian Components"),
    ("NXP-S008", "Hydraulic Piston Seal 50mm", "Seals & Gaskets", "EA", 5.40, 11.80, "Meridian Components"),
    # Electrical Components
    ("NXP-E001", "Circuit Breaker 16A Single Pole", "Electrical Components", "EA", 4.20, 9.20, "Stratos Industrial"),
    ("NXP-E002", "Circuit Breaker 32A Three Pole", "Electrical Components", "EA", 12.50, 27.50, "Stratos Industrial"),
    ("NXP-E003", "Contactor 25A 24VDC Coil", "Electrical Components", "EA", 18.40, 40.00, "Stratos Industrial"),
    ("NXP-E004", "Push Button Switch N/O Green", "Electrical Components", "EA", 2.80, 6.20, "Stratos Industrial"),
    ("NXP-E005", "Push Button Switch N/C Red", "Electrical Components", "EA", 2.80, 6.20, "Stratos Industrial"),
    ("NXP-E006", "Indicator Light LED Red 24V", "Electrical Components", "EA", 1.50, 3.30, "Stratos Industrial"),
    ("NXP-E007", "Indicator Light LED Green 24V", "Electrical Components", "EA", 1.50, 3.30, "Stratos Industrial"),
    ("NXP-E008", "Terminal Block 10A Din Rail", "Electrical Components", "EA", 0.85, 1.90, "Stratos Industrial"),
    ("NXP-E009", "Cable Duct 40 x 40mm 2m", "Electrical Components", "EA", 3.60, 7.90, "Stratos Industrial"),
    ("NXP-E010", "Din Rail 35mm x 1000mm", "Electrical Components", "EA", 2.10, 4.60, "Stratos Industrial"),
    ("NXP-E011", "Power Supply 24VDC 5A", "Electrical Components", "EA", 32.00, 69.50, "Stratos Industrial"),
    ("NXP-E012", "PLC Module Digital Input 16ch", "Electrical Components", "EA", 85.00, 185.00, "Stratos Industrial"),
    # Hydraulics
    ("NXP-H001", "Hydraulic Cylinder 50mm Bore 200mm Stroke", "Hydraulics", "EA", 68.00, 148.00, "Zenitech Catalog"),
    ("NXP-H002", "Hydraulic Pump Gear 16cc/rev", "Hydraulics", "EA", 145.00, 315.00, "Zenitech Catalog"),
    ("NXP-H003", "Hydraulic Control Valve 4/3 24VDC", "Hydraulics", "EA", 92.00, 200.00, "Zenitech Catalog"),
    ("NXP-H004", "Hydraulic Hose 1/2 inch x 3m SAE 100R2", "Hydraulics", "EA", 18.50, 40.50, "Zenitech Catalog"),
    ("NXP-H005", "Hydraulic Fitting BSP 1/2 Male", "Hydraulics", "EA", 2.40, 5.30, "Zenitech Catalog"),
    ("NXP-H006", "Hydraulic Filter Element 10 Micron", "Hydraulics", "EA", 12.80, 28.00, "Zenitech Catalog"),
    ("NXP-H007", "Pressure Gauge 0-250 Bar 1/4 BSP", "Hydraulics", "EA", 8.50, 18.50, "Zenitech Catalog"),
    ("NXP-H008", "Hydraulic Tank 20L with Fittings", "Hydraulics", "EA", 48.00, 105.00, "Zenitech Catalog"),
    # Pneumatics
    ("NXP-P001", "Pneumatic Cylinder 32mm Bore 100mm Stroke", "Pneumatics", "EA", 22.00, 48.00, "Zenitech Catalog"),
    ("NXP-P002", "Pneumatic Cylinder 50mm Bore 200mm Stroke", "Pneumatics", "EA", 38.00, 83.00, "Zenitech Catalog"),
    ("NXP-P003", "Air Filter Regulator 1/4 BSP", "Pneumatics", "EA", 14.50, 31.50, "Zenitech Catalog"),
    ("NXP-P004", "Solenoid Valve 5/2 24VDC 1/4 BSP", "Pneumatics", "EA", 16.80, 36.50, "Zenitech Catalog"),
    ("NXP-P005", "Push-in Fitting Straight 8mm x 1/8 BSP", "Pneumatics", "EA", 1.20, 2.65, "Zenitech Catalog"),
    ("NXP-P006", "Pneumatic Tubing 8mm OD x 100m", "Pneumatics", "ROLL", 28.00, 61.00, "Zenitech Catalog"),
    ("NXP-P007", "Compressed Air Dryer 150L/min", "Pneumatics", "EA", 185.00, 400.00, "Zenitech Catalog"),
    # Filtration
    ("NXP-FL001", "Oil Filter Spin-On HF6553", "Filtration", "EA", 4.20, 9.20, "AlphaPro Parts"),
    ("NXP-FL002", "Air Filter Element AF25708", "Filtration", "EA", 6.80, 14.80, "AlphaPro Parts"),
    ("NXP-FL003", "Fuel Filter FF5052", "Filtration", "EA", 5.60, 12.20, "AlphaPro Parts"),
    ("NXP-FL004", "Hydraulic Filter HF35490", "Filtration", "EA", 9.40, 20.50, "AlphaPro Parts"),
    ("NXP-FL005", "Coolant Filter WF2071", "Filtration", "EA", 7.20, 15.70, "AlphaPro Parts"),
    # Drive Components
    ("NXP-D001", "V-Belt A42 Rubber", "Drive Components", "EA", 4.80, 10.50, "DeltaForge Supply"),
    ("NXP-D002", "V-Belt B55 Rubber", "Drive Components", "EA", 6.50, 14.20, "DeltaForge Supply"),
    ("NXP-D003", "Timing Belt 480-5M-15", "Drive Components", "EA", 8.20, 18.00, "DeltaForge Supply"),
    ("NXP-D004", "Sprocket 25B16 Teeth 1 inch Bore", "Drive Components", "EA", 12.40, 27.00, "DeltaForge Supply"),
    ("NXP-D005", "Sprocket 25B24 Teeth 1 inch Bore", "Drive Components", "EA", 16.80, 36.50, "DeltaForge Supply"),
    ("NXP-D006", "Roller Chain 25B x 3m", "Drive Components", "EA", 14.20, 31.00, "DeltaForge Supply"),
    ("NXP-D007", "Shaft Coupling Flexible 19mm", "Drive Components", "EA", 9.60, 21.00, "DeltaForge Supply"),
    ("NXP-D008", "Gear Reducer 10:1 Ratio Size 60", "Drive Components", "EA", 96.00, 210.00, "DeltaForge Supply"),
    # Lubrication
    ("NXP-L001", "Lithium Grease NLGI 2 500g", "Lubrication", "EA", 4.80, 10.50, "Meridian Components"),
    ("NXP-L002", "Hydraulic Oil ISO 46 5L", "Lubrication", "L", 12.50, 27.50, "Meridian Components"),
    ("NXP-L003", "Gear Oil 90W GL-4 5L", "Lubrication", "L", 11.80, 25.80, "Meridian Components"),
    ("NXP-L004", "Chain Lubricant Spray 400ml", "Lubrication", "EA", 3.60, 7.90, "Meridian Components"),
    ("NXP-L005", "Copper Anti-Seize Paste 500g", "Lubrication", "EA", 6.20, 13.50, "Meridian Components"),
    # Couplings & Fittings
    ("NXP-C001", "BSP to NPT Adapter 1/2 inch", "Couplings & Fittings", "EA", 3.20, 7.00, "Stratos Industrial"),
    ("NXP-C002", "Equal Tee BSP 1/2 inch", "Couplings & Fittings", "EA", 2.80, 6.20, "Stratos Industrial"),
    ("NXP-C003", "Elbow 90 Degree BSP 1/2 inch", "Couplings & Fittings", "EA", 2.40, 5.30, "Stratos Industrial"),
    ("NXP-C004", "Threaded Coupling BSP 3/4 inch", "Couplings & Fittings", "EA", 3.80, 8.30, "Stratos Industrial"),
    ("NXP-C005", "Flexible Hose 1/2 inch x 500mm SS", "Couplings & Fittings", "EA", 12.50, 27.00, "Stratos Industrial"),
    ("NXP-C006", "Quick Release Coupling 1/2 inch Male", "Couplings & Fittings", "EA", 5.60, 12.20, "Stratos Industrial"),
    # Structural Parts
    ("NXP-ST001", "Steel Angle Bracket 90 Degree 50mm", "Structural Parts", "EA", 1.80, 4.00, "DeltaForge Supply"),
    ("NXP-ST002", "Steel Angle Bracket 90 Degree 100mm", "Structural Parts", "EA", 3.20, 7.00, "DeltaForge Supply"),
    ("NXP-ST003", "Aluminium Profile 40 x 40mm x 1m", "Structural Parts", "EA", 8.50, 18.50, "DeltaForge Supply"),
    ("NXP-ST004", "Aluminium Profile 40 x 80mm x 1m", "Structural Parts", "EA", 14.20, 31.00, "DeltaForge Supply"),
    ("NXP-ST005", "Steel Plate 200 x 200 x 5mm", "Structural Parts", "EA", 6.80, 14.80, "DeltaForge Supply"),
    # Tools & Accessories
    ("NXP-T001", "Hex Key Set 9-Piece Metric", "Tools & Accessories", "SET", 8.40, 18.50, "AlphaPro Parts"),
    ("NXP-T002", "Torque Wrench 1/2 Drive 40-200Nm", "Tools & Accessories", "EA", 38.00, 83.00, "AlphaPro Parts"),
    ("NXP-T003", "Cable Tie 200mm x 3.6mm Black Pack 100", "Tools & Accessories", "PACK", 2.80, 6.20, "AlphaPro Parts"),
    ("NXP-T004", "Heat Shrink Tubing 3mm x 1m Black", "Tools & Accessories", "M", 0.45, 1.00, "AlphaPro Parts"),
    ("NXP-T005", "Safety Gloves Nitrile Coated Size L", "Tools & Accessories", "PAIR", 3.20, 7.00, "AlphaPro Parts"),
    ("NXP-T006", "Multimeter Digital Auto-Range", "Tools & Accessories", "EA", 28.00, 61.00, "Stratos Industrial"),
]

# Programmatically expand PRODUCT_UNIVERSE to include Heavy Duty and Light Duty versions
_expanded_universe = []
for item in PRODUCT_UNIVERSE:
    _expanded_universe.append(item)
    sku, name, cat, unit, cost, sell, vendor = item
    # Heavy Duty variant
    _expanded_universe.append((
        f"{sku}-HD",
        f"{name} - Heavy Duty",
        cat,
        unit,
        round(cost * 1.45, 2),
        round(sell * 1.45, 2),
        vendor
    ))
    # Light Duty variant for Fasteners
    if cat == "Fasteners":
        _expanded_universe.append((
            f"{sku}-LT",
            f"{name} - Light Duty",
            cat,
            unit,
            round(cost * 0.70, 2),
            round(sell * 0.70, 2),
            vendor
        ))
PRODUCT_UNIVERSE = _expanded_universe

# Define a duplicate candidate pool to control the number of unique duplicate groups
_base_dups = [p for p in PRODUCT_UNIVERSE if not p[0].endswith("-HD") and not p[0].endswith("-LT")][:30]
DUPLICATE_POOL = _base_dups + [p for p in PRODUCT_UNIVERSE if any(p[0] == f"{x[0]}-HD" for x in _base_dups)]

# Discontinued products (subset of above SKUs, including their HD and LT variants)
DISCONTINUED_SKUS = {
    "NXP-F014", "NXP-B009", "NXP-S006", "NXP-E012", "NXP-H008",
    "NXP-D008", "NXP-L003", "NXP-ST004", "NXP-ST005"
}
DISCONTINUED_SKUS.update({f"{sku}-HD" for sku in DISCONTINUED_SKUS})
DISCONTINUED_SKUS.update({f"{sku}-LT" for sku in DISCONTINUED_SKUS})

# Stock quantities per SKU
def make_stock(sku, rng):
    if sku in DISCONTINUED_SKUS:
        return rng.choice([0, 0, rng.randint(1, 15)])
    return rng.randint(10, 500)


# ---------------------------------------------------------------------------
# Vendor-specific noise helpers
# ---------------------------------------------------------------------------

UNIT_VARIANTS = {
    "EA":   ["ea", "each", "pcs", "piece", "pieces", "Ea", "EA", "Each", "EACH", "Piece"],
    "BOX":  ["box", "boxes", "Box", "Boxes", "BOX", "bx"],
    "PACK": ["pack", "packs", "Pack", "PACK", "pkg"],
    "ROLL": ["roll", "Roll", "ROLL", "spool"],
    "SET":  ["set", "Set", "SET", "sets"],
    "KIT":  ["kit", "Kit", "KIT"],
    "L":    ["L", "l", "ltr", "litre", "Litre"],
    "KG":   ["kg", "KG", "kgs", "Kg"],
    "M":    ["m", "M", "meter", "meters"],
    "PAIR": ["pair", "Pair", "PAIR", "pairs"],
}

CATEGORY_VARIANTS = {
    "Fasteners":              ["Fasteners", "Fastener & Hardware", "Hardware Fasteners", "Bolts & Nuts", "Fixings"],
    "Bearings":               ["Bearings", "Ball Bearings", "Precision Bearings", "Bearing Components"],
    "Seals & Gaskets":        ["Seals & Gaskets", "Gaskets & Seals", "Sealing Components", "O-Rings"],
    "Electrical Components":  ["Electrical Components", "Electrical", "Electronics", "Elec Components"],
    "Hydraulics":             ["Hydraulics", "Hydraulic Components", "Hydraulic Parts", "Hydro Components"],
    "Pneumatics":             ["Pneumatics", "Pneumatic Components", "Pneumatic Parts", "Air Components"],
    "Filtration":             ["Filtration", "Filters", "Filter Elements", "Filters & Filtration"],
    "Drive Components":       ["Drive Components", "Drive Train", "Transmission", "Power Transmission"],
    "Lubrication":            ["Lubrication", "Lubricants", "Lube Products", "Oils & Greases"],
    "Tools & Accessories":    ["Tools & Accessories", "Tools", "Hand Tools", "Accessories", "Workshop Tools"],
    "Couplings & Fittings":   ["Couplings & Fittings", "Fittings", "Pipe Fittings", "Connectors"],
    "Structural Parts":       ["Structural Parts", "Metal Parts", "Steel Parts", "Fabrication"],
}

PRICE_FORMATS = [
    lambda p: f"{p:.2f}",
    lambda p: f"${p:.2f}",
    lambda p: f"USD {p:.2f}",
    lambda p: f" {p:.2f} ",
    lambda p: str(round(p, 2)),
]

STATUS_VARIANTS = {
    "Active":       ["Active", "active", "ACTIVE", "In Stock", "Available", "1", "Y", "Yes"],
    "Discontinued": ["Discontinued", "discontinued", "DISCONTINUED", "Inactive", "inactive", "Obsolete", "0", "N"],
}

VENDOR_NAME_VARIANTS = {
    "AlphaPro Parts":    ["AlphaPro Parts", "ALPHAPRO PARTS", "Alpha Pro Parts", "AlphaPro", "Alpha-Pro Parts Ltd"],
    "DeltaForge Supply": ["DeltaForge Supply", "DELTAFORGE SUPPLY", "Delta Forge Supply", "DeltaForge", "Delta Forge Supplies Co"],
    "Meridian Components": ["Meridian Components", "MERIDIAN COMPONENTS", "Meridian Comps", "Meridian", "Meridian Components Ltd"],
    "Stratos Industrial":  ["Stratos Industrial", "STRATOS INDUSTRIAL", "Stratos Ind.", "Stratos", "Stratos Industrial Supplies"],
    "Zenitech Catalog":    ["Zenitech Catalog", "ZENITECH CATALOG", "Zenitech", "Zeni-Tech Catalog", "Zenitech Products"],
}

JUNK_COLUMNS = [
    ("internal_ref", lambda rng: f"IR-{rng.randint(1000, 9999)}"),
    ("data_source", lambda rng: rng.choice(["ERP", "CRM", "Manual", "Web Portal", "EDI"])),
    ("last_updated", lambda rng: f"2024-{rng.randint(1,12):02d}-{rng.randint(1,28):02d}"),
    ("buyer_notes", lambda rng: rng.choice(["", "", "", "check with buyer", "price TBC", "to be confirmed", ""])),
    ("legacy_code", lambda rng: rng.choice(["", "", f"LC{rng.randint(100,999)}", ""])),
    ("weight_kg", lambda rng: round(rng.uniform(0.01, 12.0), 3)),
    ("location_bin", lambda rng: rng.choice(["A1", "B2", "C3", "D4", "", "TBC", "N/A"])),
    ("currency", lambda rng: rng.choice(["USD", "USD", "USD", "usd", "US$", ""])),
    ("min_order_qty", lambda rng: rng.choice([1, 5, 10, 25, 50, ""])),
    ("lead_time_days", lambda rng: rng.choice([7, 14, 21, 30, "TBD", ""])),
]


def corrupt_qty(qty, rng):
    """Randomly corrupt stock qty with messy values."""
    roll = rng.random()
    if roll < 0.05:
        return rng.choice([-5, -1, -10])
    if roll < 0.09:
        return rng.choice(["unknown", "out", "five", "TBD", "N/A"])
    if roll < 0.11:
        return ""
    return qty


def corrupt_price(price, rng, fmt_fn):
    """Randomly corrupt a price value."""
    roll = rng.random()
    if roll < 0.04:
        return ""          # blank
    if roll < 0.06:
        return 0           # zero
    if roll < 0.07:
        return -abs(price) # negative
    return fmt_fn(round(price * rng.uniform(0.95, 1.05), 2))


def maybe_blank_sku(sku, rng, prob=0.05):
    return "" if rng.random() < prob else sku


def corrupt_name(name, rng):
    """Minor name corruption: extra spaces, case changes, abbreviations."""
    roll = rng.random()
    if roll < 0.1:
        return name.upper()
    if roll < 0.2:
        return name.lower()
    if roll < 0.25:
        return "  " + name + "  "
    if roll < 0.30:
        return name.replace("-", " ").replace("mm", "MM")
    return name


# ---------------------------------------------------------------------------
# Vendor file builders
# ---------------------------------------------------------------------------

def build_alphapro(products, rng):
    """
    vendor_alphapro_parts.csv
    Columns: sku, item_name, product_type, qty_unit, unit_cost, list_price, qty, item_status, vendor
    """
    rows = []
    pool = [p for p in products if p[6] == "AlphaPro Parts"]

    for p in pool:
        sku, name, cat, unit, cost, sell, vendor = p
        stock = make_stock(sku, rng)
        status = "Discontinued" if sku in DISCONTINUED_SKUS else "Active"

        # Add junk columns
        junk = {jc[0]: jc[1](rng) for jc in rng.sample(JUNK_COLUMNS, k=rng.randint(2, 4))}

        row = {
            "sku": maybe_blank_sku(sku, rng, 0.04),
            "item_name": corrupt_name(name, rng),
            "product_type": rng.choice(CATEGORY_VARIANTS[cat]),
            "qty_unit": rng.choice(UNIT_VARIANTS.get(unit, [unit])),
            "unit_cost": corrupt_price(cost, rng, rng.choice(PRICE_FORMATS)),
            "list_price": corrupt_price(sell, rng, rng.choice(PRICE_FORMATS)),
            "qty": corrupt_qty(stock, rng),
            "item_status": rng.choice(STATUS_VARIANTS[status]),
            "vendor": rng.choice(VENDOR_NAME_VARIANTS["AlphaPro Parts"]),
            **junk,
        }
        rows.append(row)

    # Add duplicates from other vendors
    extras = rng.sample([p for p in DUPLICATE_POOL if p[6] != "AlphaPro Parts"], k=rng.randint(22, 26))
    for p in extras:
        sku, name, cat, unit, cost, sell, vendor = p
        stock = make_stock(sku, rng)
        status = "Discontinued" if sku in DISCONTINUED_SKUS else "Active"
        junk = {jc[0]: jc[1](rng) for jc in rng.sample(JUNK_COLUMNS, k=rng.randint(2, 4))}
        row = {
            "sku": maybe_blank_sku(sku, rng, 0.06),
            "item_name": corrupt_name(name, rng),
            "product_type": rng.choice(CATEGORY_VARIANTS[cat]),
            "qty_unit": rng.choice(UNIT_VARIANTS.get(unit, [unit])),
            "unit_cost": corrupt_price(cost * rng.uniform(0.9, 1.1), rng, rng.choice(PRICE_FORMATS)),
            "list_price": corrupt_price(sell * rng.uniform(0.9, 1.1), rng, rng.choice(PRICE_FORMATS)),
            "qty": corrupt_qty(stock, rng),
            "item_status": rng.choice(STATUS_VARIANTS[status]),
            "vendor": rng.choice(VENDOR_NAME_VARIANTS["AlphaPro Parts"]),
            **junk,
        }
        rows.append(row)

    rng.shuffle(rows)
    return pd.DataFrame(rows)


def build_deltaforge(products, rng):
    """
    vendor_deltaforge_supply.xlsx
    Columns: part_number, description, division, sell_unit, purchase_price, retail_price,
             stock_on_hand, availability, supplier
    """
    rows = []
    pool = [p for p in products if p[6] == "DeltaForge Supply"]

    for p in pool:
        sku, name, cat, unit, cost, sell, vendor = p
        stock = make_stock(sku, rng)
        status = "Discontinued" if sku in DISCONTINUED_SKUS else "Active"
        junk = {jc[0]: jc[1](rng) for jc in rng.sample(JUNK_COLUMNS, k=rng.randint(1, 3))}

        row = {
            "part_number": maybe_blank_sku(sku, rng, 0.03),
            "description": corrupt_name(name, rng),
            "division": rng.choice(CATEGORY_VARIANTS[cat]),
            "sell_unit": rng.choice(UNIT_VARIANTS.get(unit, [unit])),
            "purchase_price": corrupt_price(cost, rng, rng.choice(PRICE_FORMATS)),
            "retail_price": corrupt_price(sell, rng, rng.choice(PRICE_FORMATS)),
            "stock_on_hand": corrupt_qty(stock, rng),
            "availability": rng.choice(STATUS_VARIANTS[status]),
            "supplier": rng.choice(VENDOR_NAME_VARIANTS["DeltaForge Supply"]),
            **junk,
        }
        rows.append(row)

    # Extra cross-vendor duplicates
    extras = rng.sample([p for p in DUPLICATE_POOL if p[6] != "DeltaForge Supply"], k=rng.randint(22, 26))
    for p in extras:
        sku, name, cat, unit, cost, sell, vendor = p
        stock = make_stock(sku, rng)
        status = "Discontinued" if sku in DISCONTINUED_SKUS else "Active"
        junk = {jc[0]: jc[1](rng) for jc in rng.sample(JUNK_COLUMNS, k=rng.randint(1, 3))}
        row = {
            "part_number": maybe_blank_sku(sku, rng, 0.04),
            "description": corrupt_name(name, rng),
            "division": rng.choice(CATEGORY_VARIANTS[cat]),
            "sell_unit": rng.choice(UNIT_VARIANTS.get(unit, [unit])),
            "purchase_price": corrupt_price(cost * rng.uniform(0.88, 1.12), rng, rng.choice(PRICE_FORMATS)),
            "retail_price": corrupt_price(sell * rng.uniform(0.88, 1.12), rng, rng.choice(PRICE_FORMATS)),
            "stock_on_hand": corrupt_qty(stock, rng),
            "availability": rng.choice(STATUS_VARIANTS[status]),
            "supplier": rng.choice(VENDOR_NAME_VARIANTS["DeltaForge Supply"]),
            **junk,
        }
        rows.append(row)

    rng.shuffle(rows)
    return pd.DataFrame(rows)


def build_meridian(products, rng):
    """
    vendor_meridian_components.csv
    Columns: article_no, product_title, item_category, measure, cost_usd, price_usd,
             available_qty, product_state, company
    """
    rows = []
    pool = [p for p in products if p[6] == "Meridian Components"]

    for p in pool:
        sku, name, cat, unit, cost, sell, vendor = p
        stock = make_stock(sku, rng)
        status = "Discontinued" if sku in DISCONTINUED_SKUS else "Active"
        junk = {jc[0]: jc[1](rng) for jc in rng.sample(JUNK_COLUMNS, k=rng.randint(2, 5))}

        # Occasionally swap cost and sell to create cost > sell issues
        if rng.random() < 0.06:
            cost, sell = sell, cost

        row = {
            "article_no": maybe_blank_sku(sku, rng, 0.06),
            "product_title": corrupt_name(name, rng),
            "item_category": rng.choice(CATEGORY_VARIANTS[cat]),
            "measure": rng.choice(UNIT_VARIANTS.get(unit, [unit])),
            "cost_usd": corrupt_price(cost, rng, rng.choice(PRICE_FORMATS)),
            "price_usd": corrupt_price(sell, rng, rng.choice(PRICE_FORMATS)),
            "available_qty": corrupt_qty(stock, rng),
            "product_state": rng.choice(STATUS_VARIANTS[status]),
            "company": rng.choice(VENDOR_NAME_VARIANTS["Meridian Components"]),
            **junk,
        }
        rows.append(row)

    extras = rng.sample([p for p in DUPLICATE_POOL if p[6] != "Meridian Components"], k=rng.randint(20, 24))
    for p in extras:
        sku, name, cat, unit, cost, sell, vendor = p
        stock = make_stock(sku, rng)
        status = "Discontinued" if sku in DISCONTINUED_SKUS else "Active"
        junk = {jc[0]: jc[1](rng) for jc in rng.sample(JUNK_COLUMNS, k=rng.randint(2, 5))}
        row = {
            "article_no": maybe_blank_sku(sku, rng, 0.05),
            "product_title": corrupt_name(name, rng),
            "item_category": rng.choice(CATEGORY_VARIANTS[cat]),
            "measure": rng.choice(UNIT_VARIANTS.get(unit, [unit])),
            "cost_usd": corrupt_price(cost * rng.uniform(0.92, 1.08), rng, rng.choice(PRICE_FORMATS)),
            "price_usd": corrupt_price(sell * rng.uniform(0.92, 1.08), rng, rng.choice(PRICE_FORMATS)),
            "available_qty": corrupt_qty(stock, rng),
            "product_state": rng.choice(STATUS_VARIANTS[status]),
            "company": rng.choice(VENDOR_NAME_VARIANTS["Meridian Components"]),
            **junk,
        }
        rows.append(row)

    rng.shuffle(rows)
    return pd.DataFrame(rows)


def build_stratos(products, rng):
    """
    vendor_stratos_industrial.xlsx
    Columns: item_no, item_desc, product_group, unit_type, buy_price, sell_price,
             inv_qty, status, distributor
    """
    rows = []
    pool = [p for p in products if p[6] == "Stratos Industrial"]

    for p in pool:
        sku, name, cat, unit, cost, sell, vendor = p
        stock = make_stock(sku, rng)
        status = "Discontinued" if sku in DISCONTINUED_SKUS else "Active"
        junk = {jc[0]: jc[1](rng) for jc in rng.sample(JUNK_COLUMNS, k=rng.randint(1, 4))}

        row = {
            "item_no": maybe_blank_sku(sku, rng, 0.04),
            "item_desc": corrupt_name(name, rng),
            "product_group": rng.choice(CATEGORY_VARIANTS[cat]),
            "unit_type": rng.choice(UNIT_VARIANTS.get(unit, [unit])),
            "buy_price": corrupt_price(cost, rng, rng.choice(PRICE_FORMATS)),
            "sell_price": corrupt_price(sell, rng, rng.choice(PRICE_FORMATS)),
            "inv_qty": corrupt_qty(stock, rng),
            "status": rng.choice(STATUS_VARIANTS[status]),
            "distributor": rng.choice(VENDOR_NAME_VARIANTS["Stratos Industrial"]),
            **junk,
        }
        rows.append(row)

    extras = rng.sample([p for p in DUPLICATE_POOL if p[6] != "Stratos Industrial"], k=rng.randint(20, 24))
    for p in extras:
        sku, name, cat, unit, cost, sell, vendor = p
        stock = make_stock(sku, rng)
        status = "Discontinued" if sku in DISCONTINUED_SKUS else "Active"
        junk = {jc[0]: jc[1](rng) for jc in rng.sample(JUNK_COLUMNS, k=rng.randint(1, 4))}
        row = {
            "item_no": maybe_blank_sku(sku, rng, 0.05),
            "item_desc": corrupt_name(name, rng),
            "product_group": rng.choice(CATEGORY_VARIANTS[cat]),
            "unit_type": rng.choice(UNIT_VARIANTS.get(unit, [unit])),
            "buy_price": corrupt_price(cost * rng.uniform(0.90, 1.10), rng, rng.choice(PRICE_FORMATS)),
            "sell_price": corrupt_price(sell * rng.uniform(0.90, 1.10), rng, rng.choice(PRICE_FORMATS)),
            "inv_qty": corrupt_qty(stock, rng),
            "status": rng.choice(STATUS_VARIANTS[status]),
            "distributor": rng.choice(VENDOR_NAME_VARIANTS["Stratos Industrial"]),
            **junk,
        }
        rows.append(row)

    rng.shuffle(rows)
    return pd.DataFrame(rows)


def build_zenitech(products, rng):
    """
    vendor_zenitech_catalog.csv
    Columns: ref_no, name, family, units, wholesale_price, customer_price,
             units_on_hand, availability_status, brand
    """
    rows = []
    pool = [p for p in products if p[6] == "Zenitech Catalog"]

    for p in pool:
        sku, name, cat, unit, cost, sell, vendor = p
        stock = make_stock(sku, rng)
        status = "Discontinued" if sku in DISCONTINUED_SKUS else "Active"
        junk = {jc[0]: jc[1](rng) for jc in rng.sample(JUNK_COLUMNS, k=rng.randint(2, 5))}

        row = {
            "ref_no": maybe_blank_sku(sku, rng, 0.05),
            "name": corrupt_name(name, rng),
            "family": rng.choice(CATEGORY_VARIANTS[cat]),
            "units": rng.choice(UNIT_VARIANTS.get(unit, [unit])),
            "wholesale_price": corrupt_price(cost, rng, rng.choice(PRICE_FORMATS)),
            "customer_price": corrupt_price(sell, rng, rng.choice(PRICE_FORMATS)),
            "units_on_hand": corrupt_qty(stock, rng),
            "availability_status": rng.choice(STATUS_VARIANTS[status]),
            "brand": rng.choice(VENDOR_NAME_VARIANTS["Zenitech Catalog"]),
            **junk,
        }
        rows.append(row)

    extras = rng.sample([p for p in DUPLICATE_POOL if p[6] != "Zenitech Catalog"], k=rng.randint(20, 24))
    for p in extras:
        sku, name, cat, unit, cost, sell, vendor = p
        stock = make_stock(sku, rng)
        status = "Discontinued" if sku in DISCONTINUED_SKUS else "Active"
        junk = {jc[0]: jc[1](rng) for jc in rng.sample(JUNK_COLUMNS, k=rng.randint(2, 5))}
        row = {
            "ref_no": maybe_blank_sku(sku, rng, 0.06),
            "name": corrupt_name(name, rng),
            "family": rng.choice(CATEGORY_VARIANTS[cat]),
            "units": rng.choice(UNIT_VARIANTS.get(unit, [unit])),
            "wholesale_price": corrupt_price(cost * rng.uniform(0.85, 1.15), rng, rng.choice(PRICE_FORMATS)),
            "customer_price": corrupt_price(sell * rng.uniform(0.85, 1.15), rng, rng.choice(PRICE_FORMATS)),
            "units_on_hand": corrupt_qty(stock, rng),
            "availability_status": rng.choice(STATUS_VARIANTS[status]),
            "brand": rng.choice(VENDOR_NAME_VARIANTS["Zenitech Catalog"]),
            **junk,
        }
        rows.append(row)

    rng.shuffle(rows)
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Generate synthetic messy vendor data files for Nexa Parts Supply demo."
    )
    parser.add_argument("--seed", type=int, default=42, help="Random seed for reproducibility (default: 42)")
    args = parser.parse_args()

    rng = random.Random(args.seed)

    print("=" * 60)
    print("  Nexa Parts Supply - Vendor Data Generator")
    print("  Portfolio Demo | Synthetic Data Only")
    print("=" * 60)
    print(f"\n  Seed : {args.seed}")
    print(f"  Output: {RAW_DIR}\n")

    products = PRODUCT_UNIVERSE[:]

    files = [
        ("vendor_alphapro_parts.csv",     build_alphapro,  "csv",  "AlphaPro Parts"),
        ("vendor_deltaforge_supply.xlsx",  build_deltaforge,"xlsx", "DeltaForge Supply"),
        ("vendor_meridian_components.csv", build_meridian,  "csv",  "Meridian Components"),
        ("vendor_stratos_industrial.xlsx", build_stratos,   "xlsx", "Stratos Industrial"),
        ("vendor_zenitech_catalog.csv",    build_zenitech,  "csv",  "Zenitech Catalog"),
    ]

    total_rows = 0
    for filename, builder_fn, fmt, label in files:
        df = builder_fn(products, rng)
        out_path = RAW_DIR / filename
        if fmt == "csv":
            df.to_csv(out_path, index=False)
        else:
            df.to_excel(out_path, index=False)
        total_rows += len(df)
        print(f"  [OK] {filename:<40} {len(df):>4} rows  |  {len(df.columns)} columns")

    print(f"\n  " + "-" * 49)
    print(f"  Total raw rows generated : {total_rows}")
    print(f"  Product universe size    : {len(PRODUCT_UNIVERSE)}")
    print(f"  Discontinued SKUs        : {len(DISCONTINUED_SKUS)}")
    print(f"\n  All files saved to: {RAW_DIR}")
    print("=" * 60)
    print("  DISCLAIMER: All data is entirely synthetic.")
    print("  No real company, vendor, or product data is used.")
    print("=" * 60)


if __name__ == "__main__":
    main()
