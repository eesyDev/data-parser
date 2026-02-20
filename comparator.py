"""
Data comparison engine for multiple sources.
Comparison by SKU (primary key), fuzzy name matching, and attribute comparison.
"""

import os
import json
import re
import pandas as pd
from thefuzz import fuzz
import config


def merge_by_sku(sources):
    """
    Merge data from all sources by SKU.
    Returns a DataFrame with columns for each source.
    """
    source_names = list(sources.keys())
    if not source_names:
        return pd.DataFrame()

    # Collect all unique SKUs
    all_skus = set()
    for name, df in sources.items():
        skus = df[df["sku"].str.len() > 0]["sku"].unique()
        all_skus.update(skus)

    rows = []
    for sku in sorted(all_skus):
        row = {"sku": sku}
        for name, df in sources.items():
            match = df[df["sku"] == sku]
            if not match.empty:
                first = match.iloc[0]
                row[f"name_{name}"] = first["product_name"]
                row[f"price_{name}"] = first["price"]
                row[f"sale_price_{name}"] = first.get("sale_price")
                row[f"category_{name}"] = first["category"]
                row[f"status_{name}"] = first.get("status", "")
                row[f"in_{name}"] = True
            else:
                row[f"name_{name}"] = None
                row[f"price_{name}"] = None
                row[f"sale_price_{name}"] = None
                row[f"category_{name}"] = None
                row[f"status_{name}"] = None
                row[f"in_{name}"] = False
        rows.append(row)

    return pd.DataFrame(rows)


def find_missing(merged, sources):
    """
    Find products missing from one or more sources.
    """
    source_names = list(sources.keys())
    missing_rows = []

    for _, row in merged.iterrows():
        present_in = [s for s in source_names if row.get(f"in_{s}", False)]
        absent_from = [s for s in source_names if not row.get(f"in_{s}", False)]

        if absent_from:
            # Get name from the first available source
            name = None
            for s in present_in:
                if row.get(f"name_{s}"):
                    name = row[f"name_{s}"]
                    break

            # Get Zoho status if available
            zoho_status = row.get("status_zoho", "") or ""

            missing_rows.append({
                "sku": row["sku"],
                "product_name": name or "",
                "present_in": ", ".join(present_in),
                "absent_from": ", ".join(absent_from),
                "zoho_status": zoho_status,
            })

    return pd.DataFrame(missing_rows)


def find_price_differences(merged, sources):
    """
    Find products with different prices across sources.
    """
    source_names = list(sources.keys())
    diff_rows = []

    for _, row in merged.iterrows():
        prices = {}
        for s in source_names:
            if row.get(f"in_{s}", False) and pd.notna(row.get(f"price_{s}")):
                prices[s] = row[f"price_{s}"]

        if len(prices) < 2:
            continue

        price_values = list(prices.values())
        max_price = max(price_values)
        min_price = min(price_values)
        abs_diff = max_price - min_price

        # Check acceptable tolerance
        if abs_diff <= config.PRICE_TOLERANCE:
            continue

        if config.PRICE_TOLERANCE_PERCENT > 0 and min_price > 0:
            pct_diff = (abs_diff / min_price) * 100
            if pct_diff <= config.PRICE_TOLERANCE_PERCENT:
                continue

        # Get name from the first available source
        name = None
        for s in source_names:
            if row.get(f"name_{s}"):
                name = row[f"name_{s}"]
                break

        diff_row = {
            "sku": row["sku"],
            "product_name": name or "",
            "price_diff": round(abs_diff, 2),
        }
        for s in source_names:
            diff_row[f"price_{s}"] = prices.get(s)

        diff_rows.append(diff_row)

    return pd.DataFrame(diff_rows)


def find_name_differences(merged, sources):
    """
    Find products with different names (fuzzy matching).
    """
    source_names = list(sources.keys())
    diff_rows = []

    for _, row in merged.iterrows():
        names = {}
        for s in source_names:
            if row.get(f"in_{s}", False) and row.get(f"name_{s}"):
                names[s] = row[f"name_{s}"]

        if len(names) < 2:
            continue

        # Compare all pairs
        name_list = list(names.items())
        has_diff = False
        for i in range(len(name_list)):
            for j in range(i + 1, len(name_list)):
                s1, n1 = name_list[i]
                s2, n2 = name_list[j]
                ratio = fuzz.ratio(n1.lower(), n2.lower())
                if ratio < 100 and ratio >= config.FUZZY_MATCH_THRESHOLD:
                    has_diff = True

        if has_diff:
            diff_row = {"sku": row["sku"]}
            for s in source_names:
                diff_row[f"name_{s}"] = names.get(s, "")
            diff_rows.append(diff_row)

    return pd.DataFrame(diff_rows)


def _parse_weight(val):
    """Extract numeric weight value from string like '170.23 lb', '110.23', etc."""
    if not val or (isinstance(val, float) and pd.isna(val)):
        return None
    s = str(val).strip().lower().replace(",", "")
    m = re.search(r"([\d.]+)", s)
    if m:
        try:
            w = float(m.group(1))
            # Convert kg to lb if needed
            if "kg" in s:
                w = round(w * 2.20462, 2)
            return round(w, 2)
        except ValueError:
            pass
    return None


def _parse_dimensions(val):
    """Extract LxWxH dimensions as tuple of floats."""
    if not val or (isinstance(val, float) and pd.isna(val)):
        return None
    s = str(val).strip().lower().replace('"', '').replace("'", "")
    # Match patterns like "48 in x 40 in x 30 in" or "48 x 40 x 30"
    nums = re.findall(r"([\d.]+)", s)
    if len(nums) >= 3:
        try:
            return tuple(round(float(n), 2) for n in nums[:3])
        except ValueError:
            pass
    return None


def find_attribute_differences(merged, sources):
    """
    Compare attributes (weight, dimensions, category) across sources.
    Uses raw data from Zoho API cache, website CSV, and Google Sheets.
    """
    # Load raw attribute data from each source
    zoho_attrs = _load_zoho_raw_attrs()
    web_attrs = _load_website_raw_attrs()
    google_attrs = _load_google_raw_attrs()

    source_names = list(sources.keys())
    diff_rows = []

    for _, row in merged.iterrows():
        sku = row["sku"]

        # Count how many sources have this SKU
        present = [s for s in source_names if row.get(f"in_{s}", False)]
        if len(present) < 2:
            continue

        diffs = []

        # --- Weight comparison ---
        weights = {}
        if "zoho" in present and sku in zoho_attrs:
            w = _parse_weight(zoho_attrs[sku].get("weight"))
            if w:
                weights["zoho"] = w
        if "website" in present and sku in web_attrs:
            w = _parse_weight(web_attrs[sku].get("weight"))
            if w:
                weights["website"] = w
        if "google" in present and sku in google_attrs:
            # Google has both Weight (lb) and Shipping Weight
            w = _parse_weight(google_attrs[sku].get("weight_lb"))
            if w:
                weights["google"] = w

        if len(weights) >= 2:
            vals = list(weights.values())
            if max(vals) - min(vals) > 0.5:  # >0.5 lb tolerance
                diffs.append({
                    "attribute": "Weight (lb)",
                    "values": weights,
                    "diff": round(max(vals) - min(vals), 2),
                })

        # --- Shipping Weight comparison ---
        ship_weights = {}
        if "zoho" in present and sku in zoho_attrs:
            w = _parse_weight(zoho_attrs[sku].get("shipping_weight"))
            if w:
                ship_weights["zoho"] = w
        if "google" in present and sku in google_attrs:
            w = _parse_weight(google_attrs[sku].get("shipping_weight_lb"))
            if w:
                ship_weights["google"] = w

        if len(ship_weights) >= 2:
            vals = list(ship_weights.values())
            if max(vals) - min(vals) > 0.5:
                diffs.append({
                    "attribute": "Shipping Weight (lb)",
                    "values": ship_weights,
                    "diff": round(max(vals) - min(vals), 2),
                })

        # --- Dimensions comparison ---
        dims = {}
        if "zoho" in present and sku in zoho_attrs:
            d = _parse_dimensions(zoho_attrs[sku].get("dimensions"))
            if d:
                dims["zoho"] = d
        if "google" in present and sku in google_attrs:
            d = _parse_dimensions(google_attrs[sku].get("dimensions"))
            if d:
                dims["google"] = d

        if len(dims) >= 2:
            # Compare each dimension
            for i, label in enumerate(["Length", "Width", "Height"]):
                vals = {s: d[i] for s, d in dims.items()}
                v = list(vals.values())
                if max(v) - min(v) > 0.5:
                    diffs.append({
                        "attribute": f"Shipping {label} (in)",
                        "values": vals,
                        "diff": round(max(v) - min(v), 2),
                    })

        if diffs:
            # Get name from first available source
            name = ""
            for s in present:
                if row.get(f"name_{s}"):
                    name = row[f"name_{s}"]
                    break

            for d in diffs:
                diff_row = {
                    "sku": sku,
                    "product_name": name,
                    "attribute": d["attribute"],
                    "diff": d["diff"],
                }
                for s in source_names:
                    diff_row[f"value_{s}"] = d["values"].get(s, "")
                diff_rows.append(diff_row)

    return pd.DataFrame(diff_rows)


def _load_zoho_raw_attrs():
    """Load raw attributes from Zoho API cache."""
    cache_path = os.path.join(config.DATA_DIR, "zoho_api_cache.json")
    if not os.path.exists(cache_path):
        return {}
    with open(cache_path, "r", encoding="utf-8") as f:
        items = json.load(f)
    result = {}
    for item in items:
        sku = str(item.get("sku", "")).strip().upper()
        if not sku:
            continue
        result[sku] = {
            "weight": item.get("weight_with_unit", ""),
            "shipping_weight": item.get("weight_with_unit", ""),  # Zoho uses same field
            "dimensions": item.get("dimensions_with_unit", ""),
        }
    return result


def _load_website_raw_attrs():
    """Load raw attributes from website CSV."""
    if not os.path.exists(config.WEBSITE_CSV):
        return {}
    df = pd.read_csv(config.WEBSITE_CSV, dtype=str, low_memory=False)
    result = {}
    for _, row in df.iterrows():
        sku = str(row.get("SKU", "")).strip().upper()
        if not sku:
            continue
        result[sku] = {
            "weight": row.get("Weight (lbs)", ""),
        }
    return result


def _load_google_raw_attrs():
    """Load raw attributes from Google Sheets (all sheets)."""
    path = config.GOOGLE_XLSX
    if not os.path.exists(path):
        return {}
    result = {}
    xls = pd.ExcelFile(path)
    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet, dtype=str)
        df.columns = [c.strip() for c in df.columns]
        sku_col = "SKU" if "SKU" in df.columns else None
        if not sku_col:
            for col in df.columns:
                if col.strip().upper() == "SKU" or col.strip() == "":
                    sku_col = col
                    break
        if not sku_col:
            continue
        for _, row in df.iterrows():
            sku = str(row.get(sku_col, "")).strip().upper()
            if not sku or sku == "NAN":
                continue
            result[sku] = {
                "weight_lb": row.get("Weight (lb)", ""),
                "shipping_weight_lb": row.get("Shipping Weight (lb)/NA",
                                     row.get("Shipping Weight (lb)", "")),
                "dimensions": f"{row.get('Shipping Length (in)/NA', row.get('Shipping Length (in)', ''))} x "
                             f"{row.get('Shipping Width (in)/NA', row.get('Shipping Width (in)', ''))} x "
                             f"{row.get('Shipping Height (in)/NA', row.get('Shipping Height (in)', ''))}",
            }
    return result


def compare_all(sources):
    """
    Full comparison of all sources.
    Returns a dict with results.
    """
    print("\n2. Comparing data...")
    merged = merge_by_sku(sources)
    print(f"   Total unique SKUs: {len(merged)}")

    missing = find_missing(merged, sources)
    print(f"   Products with gaps: {len(missing)}")

    price_diff = find_price_differences(merged, sources)
    print(f"   Price discrepancies: {len(price_diff)}")

    name_diff = find_name_differences(merged, sources)
    print(f"   Name discrepancies: {len(name_diff)}")

    attr_diff = find_attribute_differences(merged, sources)
    print(f"   Attribute discrepancies: {len(attr_diff)}")

    return {
        "merged": merged,
        "missing": missing,
        "price_differences": price_diff,
        "name_differences": name_diff,
        "attribute_differences": attr_diff,
        "sources": sources,
    }
