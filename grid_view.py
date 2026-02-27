"""
Full attribute grid: ALL products grouped by category.
For each product shows attribute values from all 3 sources (Zoho name, Website, Google).
Highlights what's present, missing, or mismatched.
"""

import os
import json
import re
import pandas as pd
import config
import name_parser


# Key attributes to show per category group.
# "common" applies to most attachment categories.
COMMON_ATTRS = [
    "Bucket Size",
    "Front Pin Diameter",
    "Pin Size",
    "Front Pin Length (mm)",
    "Rear Pin Length (mm)",
    "Carrier Weight Class",
    "Machine Type",
    "Head Style",
    "Shipping Weight (lb)",
]

# Category-specific extra attributes
CATEGORY_ATTRS = {
    "Bucket Tooth": ["Tooth Style", "Tooth Pin Part Number", "Serie"],
    "Bucket Pin": ["Diameter (mm)", "Length (mm)"],
    "Bucket Shims": ["Outer Diameter", "Interior Diameter", "Thickness"],
    "Auger Bits": ["Auger Bit Size", "Hex Size"],
    "Rock Auger Bits": ["Auger Bit Size", "Hex Size"],
    "Hydraulic Hammer": ["Energy Class (J)"],
    "Hydraulic Hammers": ["Energy Class (J)"],
    "Post Driver Hammer": ["Energy Class (J)"],
    "Post Driver Hammers": ["Energy Class (J)"],
    "Plate Compactor": ["Compaction Width (in)", "Compaction Width (mm)", "Impulse Force (lbs)", "Impulse Force (tons)"],
    "Plate Compactors": ["Compaction Width (in)", "Compaction Width (mm)", "Impulse Force (lbs)", "Impulse Force (tons)"],
    "Hammer Chisel Bits": ["Diameter (mm)"],
    "Hammer Moil Chisel Bits": ["Diameter (mm)"],
    "Hammer Wedge Chisel Bits": ["Diameter (mm)"],
    "Mechanical Thumb": ["Thumb Width (in)", "Product Length (in)"],
    "Mechanical Grapple": ["Grapple Width (in)", "Grapple Width (mm)"],
    "Mechanical Grapples": ["Grapple Width (in)", "Grapple Width (mm)"],
    "Rotating Grapple": ["Grapple Width (in)", "Grapple Width (mm)"],
    "Rotating Hydraulic Grapple": ["Grapple Width (in)", "Grapple Width (mm)"],
    "Rotating Hydraulic Grapples": ["Grapple Width (in)", "Grapple Width (mm)"],
    "Vibratory Roller": ["Compaction Width (in)", "Compaction Width (mm)"],
    "Vibratory Rollers": ["Compaction Width (in)", "Compaction Width (mm)"],
    "Compaction Wheel": ["Compaction Width (in)", "Compaction Width (mm)"],
    "Compaction Wheels": ["Compaction Width (in)", "Compaction Width (mm)"],
}

# Categories where "Bucket Size" doesn't apply
NO_BUCKET_SIZE = {
    "Hydraulic Hammer", "Hydraulic Hammers", "Hammer Chisel Bits",
    "Hammer Moil Chisel Bits", "Hammer Wedge Chisel Bits",
    "Hammer Plate Head", "Hammer Plate Heads",
    "Post Driver Hammer", "Post Driver Hammers",
    "Hydraulic Shear", "Hydraulic Shears",
    "Concrete Pulverizer", "Concrete Pulverizers",
    "Bucket Tooth", "Bucket Tooth Adapter", "Bucket Tooth Adapters",
    "Bucket Tooth Pin", "Bucket Tooth Pins",
    "Bucket Tooth Retainer", "Bucket Tooth Retainers",
    "Bucket Pin", "Bucket Pins", "Bucket Shims",
    "Bucket Side Cutters", "Bucket Side Bar Protector", "Bucket Side Bar Protectors",
    "Ripper Tooth", "Ripper Tooth - Shank", "Ripper Tooth Shank",
    "Ripper Tooth - Tooth Pin", "Ripper Tooth Pin",
    "Ripper Tooth - Tooth Replacement", "Ripper Tooth Replacement",
    "Ripper Shank Protectors",
    "Skid Steer Wear Parts",
    "Mechanical Grapple", "Mechanical Grapples",
    "Rotating Grapple", "Rotating Hydraulic Grapple", "Rotating Hydraulic Grapples",
    "Heavy Duty Grapple Buckets", "Heavy Duty Grapple Skeleton Buckets",
    "Mechanical Thumb", "Hydraulic Thumb",
    "Main Pin Hydraulic Thumb", "Main Pin Hydraulic Progressive Thumb",
    "QC Main Pin Hydraulic Thumb",
    "Bolt-On Mount", "Auger Bits", "Rock Auger Bits",
    "Plate Compactor", "Plate Compactors",
    "Vibratory Roller", "Vibratory Rollers",
    "Compaction Wheel", "Compaction Wheels",
    "Pallet Forks - Excavators", "Pallet Forks - Skid Steer",
    "Universal Skid Steer Loader Adapter",
    "Flat Face Hydraulic Quick Connector",
    "Aux Hydraulic Piping Kits", "GearBox For Brush Cutter",
    "Hydraulic Motor Pump For Roller", "Vibratory Roller Parts Connection",
    "Blade Set for Brush Cutter", "Wheel & Plate Compactors Bolt-On",
    "Universal Weld-On Head Plates Set",
}

# Attributes that are NEVER in the product name (data-only from sources).
# Don't flag as PARTIAL just because parsed-name value is missing.
DATA_ONLY_ATTRS = {
    "Front Pin Diameter",
    "Center to Center",
    "Drain Holes",
    "Product Weight (lbs)",
    "Product Weight (kg)",
    "Capacity (m³)",
    "Add-on included",
    "Add-ons Included",
    "Front Ear to Ear",
    "Rear Ear to Ear",
    "Front Pin Length (mm)",
    "Rear Pin Length (mm)",
    "Shipping Length (in)",
    "Shipping Width (in)",
    "Shipping Height (in)",
    "Shipping Weight (lb)",
    "Shipping Weight (kg)",
    "Impulse Force (lbs)",
    "Impulse Force (tons)",
}


def build_grid():
    """
    Build full attribute grid for all products, grouped by category.
    Returns dict: {category: DataFrame with all products and attribute columns}.
    """
    print("  Building attribute grid...")

    # Load data
    zoho_titles = name_parser._load_zoho_website_titles()
    zoho_attrs = name_parser._load_zoho_attrs()
    web_attrs = name_parser._load_website_attrs()
    google_attrs = name_parser._load_google_attrs()

    # Load zoho cache for price/status — exclude inactive items
    zoho_cache = {}
    cache_path = os.path.join(config.DATA_DIR, "zoho_api_cache.json")
    if os.path.exists(cache_path):
        with open(cache_path, "r", encoding="utf-8") as f:
            for item in json.load(f):
                if str(item.get("status", "")).lower() == "inactive":
                    continue
                sku = str(item.get("sku", "")).strip().upper()
                if sku:
                    zoho_cache[sku] = item

    # Load website names/prices
    web_products = {}
    if os.path.exists(config.WEBSITE_CSV):
        df = pd.read_csv(config.WEBSITE_CSV, dtype=str, low_memory=False)
        for _, row in df.iterrows():
            sku = str(row.get("SKU", "")).strip().upper()
            if sku:
                web_products[sku] = {
                    "name": row.get("Name", ""),
                    "price": row.get("Regular price", ""),
                }

    # Group ALL Zoho products by category (not just those with website title)
    categories = {}
    for sku, item in zoho_cache.items():
        cat = item.get("category_name", "") or "Unknown"
        if cat not in categories:
            categories[cat] = []
        categories[cat].append(sku)

    # Build grid per category
    grids = {}
    for cat in sorted(categories.keys()):
        skus = sorted(categories[cat])
        if not skus:
            continue

        # Determine relevant attributes for this category
        attr_keys = [a for a in COMMON_ATTRS if not (
            a in ("Bucket Size", "Front Pin Length (mm)", "Rear Pin Length (mm)")
            and cat in NO_BUCKET_SIZE
        )]
        if cat in CATEGORY_ATTRS:
            attr_keys = CATEGORY_ATTRS[cat] + attr_keys

        # Collect all unique attribute keys from all sources for this category
        all_zoho_keys = set()
        all_web_keys = set()
        all_google_keys = set()
        for sku in skus:
            if sku in zoho_attrs:
                all_zoho_keys.update(zoho_attrs[sku].keys())
            if sku in web_attrs:
                all_web_keys.update(web_attrs[sku].keys())
            if sku in google_attrs:
                all_google_keys.update(google_attrs[sku].keys())

        # Add source-specific attrs that aren't already covered
        extra_zoho = _filter_important_attrs(all_zoho_keys, attr_keys)
        extra_web = _filter_important_attrs(all_web_keys, attr_keys)
        extra_google = _filter_important_attrs(all_google_keys, attr_keys)

        # Build unified list of attribute columns — deduplicate by base name (including substrings)
        required_attrs = set(attr_keys)  # COMMON_ATTRS + CATEGORY_ATTRS — always shown if any data
        all_attr_names = list(dict.fromkeys(attr_keys))  # preserve order, unique
        covered_bases = {_attr_base(a) for a in all_attr_names}
        for a in extra_zoho + extra_web + extra_google:
            base = _attr_base(a)
            if a not in all_attr_names and not any(
                cb == base or cb in base or base in cb
                for cb in covered_bases
            ):
                all_attr_names.append(a)
                covered_bases.add(base)

        rows = []
        for sku in skus:
            zoho_data = zoho_titles.get(sku, {})
            title = zoho_data.get("website_title", "")
            if not title or not title.strip():
                continue
            parsed = name_parser.parse_name(title, cat)

            zoho_item = zoho_cache.get(sku, {})
            web_prod = web_products.get(sku, {})

            # Google product name
            google_name = ""
            g_attrs = google_attrs.get(sku, {})
            for key in ["Variation Name", "Variation Name/NA", "Product Name", "Name"]:
                if key in g_attrs:
                    google_name = g_attrs[key]
                    break

            zoho_status = str(zoho_item.get("status", "")).capitalize()

            row = {
                "SKU": sku,
                "Zoho Status": zoho_status,
                "Zoho Name": str(zoho_item.get("name", ""))[:80],
                "Zoho Title": title[:80] if title else "",
                "Website Name": str(web_prod.get("name", ""))[:80],
                "Google Name": str(google_name)[:80],
                "Zoho Price": zoho_item.get("rate", ""),
                "Website Price": web_prod.get("price", ""),
                "In Website": "Yes" if sku in web_attrs else "No",
                "In Google": "Yes" if sku in google_attrs else "No",
            }

            for attr_name in all_attr_names:
                # Value from parsed name (skip data-only attrs — never extracted from title)
                name_val = "" if attr_name in DATA_ONLY_ATTRS else _find_parsed_value(parsed, attr_name)

                # Value from Zoho attributes
                zoho_val = _find_source_value(zoho_attrs.get(sku, {}), attr_name)

                # Value from website
                web_val = _find_source_value(web_attrs.get(sku, {}), attr_name)

                # Value from google
                google_val = _find_source_value(google_attrs.get(sku, {}), attr_name)

                row[f"{attr_name} [Name]"] = name_val
                row[f"{attr_name} [Zoho]"] = zoho_val
                row[f"{attr_name} [Web]"] = web_val
                row[f"{attr_name} [Google]"] = google_val

                # Status — for data-only attrs, ignore name value in evaluation
                if attr_name in DATA_ONLY_ATTRS:
                    vals = [v for v in [zoho_val, web_val, google_val] if v]
                else:
                    vals = [v for v in [name_val, zoho_val, web_val, google_val] if v]
                if not vals:
                    row[f"{attr_name} [Status]"] = ""
                elif len(vals) == 1:
                    row[f"{attr_name} [Status]"] = "PARTIAL"
                else:
                    # Check if all non-empty values match
                    match = _all_match(vals)
                    row[f"{attr_name} [Status]"] = "OK" if match else "MISMATCH"

            rows.append(row)

        df = pd.DataFrame(rows)

        # Drop attribute columns that are empty or too sparse.
        # Required attrs (COMMON_ATTRS + CATEGORY_ATTRS): keep if any row has data.
        # Extra (dynamically discovered) attrs: keep only if ≥25% of rows have data.
        min_coverage = max(2, int(0.25 * len(df))) if len(df) > 0 else 1
        non_empty_attrs = []
        for attr in all_attr_names:
            rows_with_data = 0
            for suffix in ["[Name]", "[Zoho]", "[Web]", "[Google]"]:
                col = f"{attr} {suffix}"
                if col in df.columns:
                    rows_with_data = max(
                        rows_with_data,
                        df[col].apply(
                            lambda x: bool(str(x).strip()) and str(x).lower() not in ("", "nan", "none")
                        ).sum()
                    )
            if attr in required_attrs:
                # Required: show if any single row has data
                if rows_with_data >= 1:
                    non_empty_attrs.append(attr)
            else:
                # Extra: show only if enough rows have data
                if rows_with_data >= min_coverage:
                    non_empty_attrs.append(attr)

        grids[cat] = {
            "data": df,
            "attr_names": non_empty_attrs,
        }

    return grids


def _attr_base(key):
    """Strip trailing (unit) and normalize for dedup comparison.
    Capacity columns keep their unit so yd³ and m³ stay as separate columns.
    """
    if re.search(r'(?i)capacity|grapple width|compaction width|impulse force', key):
        return key.strip().lower()
    return re.sub(r'\s*\([^)]*\)\s*$', '', key).strip().lower()


def _filter_important_attrs(all_keys, already_covered):
    """Filter to important/interesting attributes, skip noise."""
    skip_patterns = [
        "Variation Name", "Category", "Handling Unit", "Unit",
        "Weight (lb)", "Weight (kg)",
        "Product Name", "Name",
        # Handled via Bucket Size alias — avoid duplicate column
        "Product Width (in)", "Product Width (mm)",
        # Handled via Compaction Width (in)/(mm) — avoid duplicate column
        "Compaction Width",
        # Handled via Impulse Force (lbs)/(tons) — avoid duplicate column
        "Impulse Force",
        # Duplicates — already covered by Head Style column
        "Coupler Head Type", "Coupler Type",
        # Duplicate — already covered by Product Type
        "Bucket Type",
        # Handled via Diameter (mm) — avoid duplicate column
        "Chisel Bit Size", "Bit Diameter",
    ]
    result = []
    for key in sorted(all_keys):
        if any(p.lower() in key.lower() for p in skip_patterns):
            continue
        base = _attr_base(key)
        # Check if it's already covered by our common attrs or result so far
        covered = False
        for existing in list(already_covered) + result:
            existing_base = _attr_base(existing)
            if existing_base == base or existing_base in base or base in existing_base:
                covered = True
                break
        if not covered:
            result.append(key)
    return result[:15]  # Limit to avoid too many columns


def _find_parsed_value(parsed, attr_name):
    """Find value in parsed dict, matching by name or ATTR_MAP reverse lookup."""
    if attr_name in parsed:
        return parsed[attr_name]
    # Fuzzy key match
    for k, v in parsed.items():
        if k.lower() in attr_name.lower() or attr_name.lower() in k.lower():
            return v
    # Reverse ATTR_MAP lookup: if attr_name appears in a parsed key's alias list,
    # return that parsed key's value.
    # e.g. grid column "Front Pin Diameter (mm)" → ATTR_MAP["Pin Size"] contains it → use parsed["Pin Size"]
    for parsed_key, alias_list in name_parser.ATTR_MAP.items():
        if attr_name in alias_list and parsed_key in parsed:
            return parsed[parsed_key]
    return ""


def _find_source_value(attrs, attr_name):
    """Find value in source attributes dict."""
    if not attrs:
        return ""
    if attr_name in attrs:
        return attrs[attr_name]
    # For capacity/grapple/compaction/impulse/weight columns use strict matching only — no partial across units
    if re.search(r'(?i)capacity|grapple width|compaction width|impulse force|shipping weight|product weight', attr_name):
        pass  # skip partial match, go straight to aliases below
    else:
        # Try partial match
        attr_lower = attr_name.lower()
        for k, v in attrs.items():
            if attr_lower in k.lower() or k.lower() in attr_lower:
                return v
    # Map common aliases
    aliases = {
        "Bucket Size": ["Bucket Size (in)", "Bucket Size (in)/Filter",
                        "Product Width (in)",
                        "Rake Size", "Rake Width (in)", "Width (in)",
                        "Grapple Width (in)", "Grapple Size", "Grapple Width",
                        "Broom Width (in)", "Broom Size", "Brush Size",
                        "Compaction Width", "Fork Size", "Saw Length"],
        "Front Pin Diameter": ["Front Pin Diameter (mm)", "Front Pin Diameter (mm)/Filter",
                               "Pin Diameter (mm)", "Front Pin Size"],
        "Pin Size": ["Front Pin Diameter (mm)", "Front Pin Diameter",
                     "Pin Diameter (mm)", "Pin size", "Front Pin Size"],
        "Front Pin Diameter (mm)": ["Front Pin Diameter", "Pin Size",
                                     "Pin Diameter (mm)", "Pin size", "Front Pin Size"],
        "Rear Pin Diameter (mm)": ["Rear Pin Diameter", "Rear Pin Diameter (mm)/Filter",
                                    "Rear Pin Size (mm)", "Back Pin Size (mm)"],
        "Rear Ear to Ear": ["Rear Ear to Ear (mm)", "Rear Ear to Ear Distance",
                             "Back Ear to Ear"],
        "Carrier Weight Class": ["Carrier Weight Class (tn)", "Carrier Weight Class "],
        "Head Style": ["Coupler Head Type", "Head Style", "Coupler Type", "Coupler Type/Filter",
                       "Head Type"],
        "Machine Type": ["Machine Type", "Attachment Types", "Attachment Types/NA"],
        "Fits To": ["Fits To"],
        "Front Ear to Ear": ["Front Ear to Ear (mm)", "Front Ear to Ear Distance",
                              "Front Ear to Ear Distance (mm)"],
        # Capacity — strict per-unit aliases, no cross-unit lookup
        "Capacity (yd\u00b3)": ["Capacity (yd\u00b3)/Filter", "Capacity ($yd^3$)", "Capacity"],
        "Capacity (m\u00b3)":  ["Capacity ($m^3$)", "Capacity (m3)"],
        "Attachment Types": ["Attachment Types/NA"],
        "Bucket Type": ["Category", "Category/NA"],
        "Category": ["Bucket Type"],
        "Bucket Width (mm)": ["Product Width (mm)", "Bucket Width (mm)"],
        "Grapple Width (in)": ["Grapple Width (in)", "Grapple Width", "Grapple Size", "Product Width (in)"],
        "Grapple Width (mm)": ["Grapple Width (mm)"],
        "Compaction Width (in)": ["Compaction Width (in)", "Compaction Width", "Width (in)", "Plate Width (in)", "Product Width (in)"],
        "Compaction Width (mm)": ["Compaction Width (mm)", "Width (mm)", "Product Width (mm)"],
        # "Impluse" is a typo in the Google Sheet source — keep both spellings
        "Impulse Force (lbs)": ["Impulse Force (lbs)", "Impulse Force (lb)",
                                 "Impluse Force (lb)", "Impluse Force (lb)/Filter"],
        "Impulse Force (tons)": ["Impulse Force (tons)", "Impulse Force (tn)", "Impulse Force (t)"],
        "Product Weight (lbs)": ["Weight (lb)", "Weight (lbs)", "Rake Weight (lb)"],
        "Product Weight (kg)": ["Weight (kg)", "Rake Weight (kg)"],
        "Shipping Length (in)": ["Shipping Length (in)/NA", "Shipping Length (in)"],
        "Shipping Width (in)":  ["Shipping Width (in)/NA",  "Shipping Width (in)"],
        "Shipping Height (in)": ["Shipping Height (in)/NA", "Shipping Height (in)"],
        "Shipping Weight (lb)": ["Shipping Weight (lb)/NA", "Shipping Weight (lb)", "Shipping Weight (lbs)", "Shipping Weight (lbs)/NA"],
        "Shipping Weight (kg)": ["Shipping Weight (kg)/NA", "Shipping Weight (kg)"],
        "Front Pin Length (mm)": ["Front Pin Length (mm)/NA", "Front Pin Length (mm)", "Front Pin Length"],
        "Rear Pin Length (mm)":  ["Rear Pin Length (mm)/NA",  "Rear Pin Length (mm)",  "Rear Pin Length"],
        "Diameter (mm)": ["Chisel Bit Size", "Chisel Bit Diameter (mm)",
                          "Bit Diameter (mm)", "Pin Diameter (mm)"],
        "Outer Diameter": ["Outer Diameter (mm)/Filter", "Outer Diameter (mm)"],
        "Interior Diameter": ["Interior Diameter (mm)/Filter", "Interior Diameter (mm)"],
        "Thickness": ["Height (mm)/Filter", "Height (mm)", "Thickness (mm)"],
    }
    for alias in aliases.get(attr_name, []):
        if alias in attrs:
            return attrs[alias]
    return ""


def _all_match(values):
    """Check if all non-empty values semantically match."""
    if len(values) < 2:
        return True
    cleaned = [v for v in values if v and str(v).strip()]
    if len(cleaned) < 2:
        return True
    # Check all pairs against each other
    for i in range(len(cleaned)):
        for j in range(i + 1, len(cleaned)):
            if not name_parser._match_value(cleaned[i], cleaned[j]):
                return False
    return True


_GRID_SUFFIXES = [
    ("Name",   "sub-name",   "val-name"),
    ("Zoho",   "sub-zoho",   "val-zoho"),
    ("Web",    "sub-web",    "val-web"),
    ("Google", "sub-google", "val-google"),
]
_GRID_SUB_LABELS = {"Name": "From Title", "Zoho": "Zoho", "Web": "Web", "Google": "Google"}


def _render_grid_table(df_group, attr_names):
    """Render grid-scroll + table HTML for one group. Hides sub-columns with no data."""
    if df_group is None or len(df_group) == 0:
        return ''

    vcls = {s: vc for s, _, vc in _GRID_SUFFIXES}
    scls = {s: sc for s, sc, _ in _GRID_SUFFIXES}

    # Which sub-columns have any data per attr?
    active_subs = {}
    for attr in attr_names:
        subs = []
        for s, _, _ in _GRID_SUFFIXES:
            col = f"{attr} [{s}]"
            if col in df_group.columns:
                if df_group[col].apply(
                    lambda x: bool(str(x).strip()) and str(x).lower() not in ("", "nan", "none")
                ).any():
                    subs.append(s)
        if subs:
            active_subs[attr] = subs

    t = '      <div class="grid-scroll">\n        <table>\n'

    # Header row 1
    t += '          <thead>\n          <tr>\n'
    t += '            <th rowspan="2" class="sticky-col sticky-col-0">SKU</th>\n'
    t += '            <th rowspan="2" class="sticky-col sticky-col-1">Zoho Title (website)</th>\n'
    t += '            <th rowspan="2" class="done-col">&#10003;</th>\n'
    t += '            <th rowspan="2">Status</th>\n'
    t += '            <th rowspan="2">Website Name</th>\n'
    t += '            <th rowspan="2">Google Name</th>\n'
    t += '            <th rowspan="2">In Web</th>\n'
    t += '            <th rowspan="2">In Google</th>\n'
    for attr in attr_names:
        subs = active_subs.get(attr)
        if not subs:
            continue
        t += f'            <th class="attr-group" colspan="{len(subs) + 1}">{_esc(attr)}</th>\n'
    t += '          </tr>\n'

    # Header row 2
    t += '          <tr>\n'
    for attr in attr_names:
        subs = active_subs.get(attr)
        if not subs:
            continue
        for s in subs:
            t += f'            <th class="sub {scls[s]}">{_GRID_SUB_LABELS[s]}</th>\n'
        t += '            <th class="sub">St</th>\n'
    t += '          </tr>\n          </thead>\n'

    # Body
    t += '          <tbody>\n'
    for _, row in df_group.iterrows():
        has_mismatch = any(
            str(row.get(f"{a} [Status]", "")) == "MISMATCH"
            for a in attr_names if active_subs.get(a)
        )
        _sku = _esc(str(row.get("SKU", "")))
        t += f'          <tr data-has-mismatch="{1 if has_mismatch else 0}" data-sku="{_sku}">\n'
        t += f'            <td class="sticky-col sticky-col-0"><strong>{_sku}</strong></td>\n'
        t += f'            <td class="sticky-col sticky-col-1">{_esc(str(row.get("Zoho Title", "")))}</td>\n'
        t += f'            <td class="done-check"><input type="checkbox" class="done-cb" data-sku="{_sku}" onchange="toggleRow(this)"></td>\n'
        zoho_st = str(row.get("Zoho Status", ""))
        t += f'            <td class="{"tag-no" if zoho_st.lower() == "inactive" else ""}">{_esc(zoho_st)}</td>\n'
        t += f'            <td class="val-web">{_esc(str(row.get("Website Name", "")))}</td>\n'
        t += f'            <td class="val-google">{_esc(str(row.get("Google Name", "")))}</td>\n'
        in_web = str(row.get("In Website", ""))
        in_google = str(row.get("In Google", ""))
        t += f'            <td class="{"tag-yes" if in_web == "Yes" else "tag-no"}">{in_web}</td>\n'
        t += f'            <td class="{"tag-yes" if in_google == "Yes" else "tag-no"}">{in_google}</td>\n'
        for attr in attr_names:
            subs = active_subs.get(attr)
            if not subs:
                continue
            status = str(row.get(f"{attr} [Status]", "") or "")
            cell_cls = {"MISMATCH": "cell-mismatch", "PARTIAL": "cell-partial", "OK": "cell-ok"}.get(status, "")
            for s in subs:
                v = str(row.get(f"{attr} [{s}]", "") or "")
                if v and v.lower() not in ("nan", "none"):
                    t += f'            <td class="{cell_cls} {vcls[s]}">{_esc(v)}</td>\n'
                else:
                    t += f'            <td class="{cell_cls} cell-empty">&mdash;</td>\n'
            st_cls = {"OK": "tag-ok", "MISMATCH": "tag-mis", "PARTIAL": "tag-part"}.get(status, "cell-empty")
            st_icon = {"OK": "&#10003;", "MISMATCH": "&#10007;", "PARTIAL": "~"}.get(status, "")
            t += f'            <td class="{cell_cls} {st_cls}">{st_icon}</td>\n'
        t += '          </tr>\n'
    t += '          </tbody>\n        </table>\n      </div>\n'
    return t


def generate_grid_html(grids, output_path=None):
    """Generate HTML page with attribute grids grouped by category."""
    output_path = output_path or os.path.join(config.OUTPUT_DIR, "attribute_grid.html")
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    # Count totals
    total_products = sum(len(g["data"]) for g in grids.values())
    total_mismatches = 0
    total_partial = 0
    for g in grids.values():
        for col in g["data"].columns:
            if col.endswith("[Status]"):
                total_mismatches += (g["data"][col] == "MISMATCH").sum()
                total_partial += (g["data"][col] == "PARTIAL").sum()

    cat_list = sorted(grids.keys(), key=lambda c: -len(grids[c]["data"]))

    # Pin-related attributes to hide in "without pins" groups
    _PIN_ATTRS = {"Front Pin Diameter", "Pin Size", "Rear Pin Diameter (mm)",
                  "Front Pin Length (mm)", "Rear Pin Length (mm)"}

    def _row_has_fpd(row, fpd_cols):
        for c in fpd_cols:
            v = str(row.get(c, "") or "").strip()
            if v and v.lower() not in ("nan", "none"):
                return True
        return False

    display_cats = []
    for _cat in cat_list:
        _df = grids[_cat]["data"]
        _an = grids[_cat]["attr_names"]
        _fpd_cols = [f"Front Pin Diameter [{s}]" for s in ["Zoho", "Web", "Google"]]
        _avail = [c for c in _fpd_cols if c in _df.columns]
        if _avail:
            _mask = _df.apply(lambda r: _row_has_fpd(r, _avail), axis=1)
            _wp, _wop = _df[_mask], _df[~_mask]
            if len(_wp) > 0 and len(_wop) > 0:
                _sb = re.sub(r'[^a-zA-Z0-9]', '', _cat)
                display_cats.append({"label": f"{_cat} with pins",
                                     "safe_id": _sb + "withpins",
                                     "df": _wp, "attr_names": _an})
                display_cats.append({"label": f"{_cat} without pins",
                                     "safe_id": _sb + "withoutpins",
                                     "df": _wop,
                                     "attr_names": [a for a in _an if a not in _PIN_ATTRS]})
                continue
        display_cats.append({"label": _cat,
                             "safe_id": re.sub(r'[^a-zA-Z0-9]', '', _cat),
                             "df": _df, "attr_names": _an})

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Attribute Grid — JM Attachments</title>
<style>
  :root {{
    --bg: #0f172a; --card: #1e293b; --border: #334155;
    --text: #e2e8f0; --muted: #94a3b8; --accent: #3b82f6;
    --green: #22c55e; --red: #ef4444; --orange: #f59e0b; --purple: #a78bfa;
  }}
  * {{ margin:0; padding:0; box-sizing:border-box; }}
  body {{
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', system-ui, sans-serif;
    background: var(--bg); color: var(--text); line-height: 1.4; padding: 20px;
    font-size: 13px; overflow-x: hidden;
  }}
  .container {{ max-width: 100%; margin: 0 auto; }}
  /* Page navigation */
  .page-nav {{
    display:flex; gap:8px; margin-bottom:20px; flex-wrap:wrap;
  }}
  .page-nav a {{
    padding:7px 16px; border-radius:8px; font-size:0.85rem; font-weight:500;
    text-decoration:none; border:1px solid var(--border); color:var(--muted);
    background:var(--card); transition:all 0.15s;
  }}
  .page-nav a:hover {{ color:var(--text); border-color:var(--accent); }}
  .page-nav a.current {{ color:var(--accent); border-color:var(--accent); background:rgba(59,130,246,0.1); }}
  h1 {{ font-size: 1.5rem; margin-bottom: 4px; }}
  .subtitle {{ color: var(--muted); margin-bottom: 16px; }}
  .stats {{ display:flex; gap:12px; margin-bottom:20px; flex-wrap:wrap; }}
  .stat {{ background:var(--card); border:1px solid var(--border); border-radius:10px; padding:12px 16px; }}
  .stat .num {{ font-size:1.5rem; font-weight:700; }}
  .stat .label {{ color:var(--muted); font-size:0.8rem; }}

  /* Category nav */
  .cat-nav {{
    display:flex; flex-wrap:wrap; gap:6px; margin-bottom:20px;
  }}
  .cat-btn {{
    padding:6px 12px; border-radius:8px; cursor:pointer;
    background:var(--card); border:1px solid var(--border); color:var(--muted);
    font-size:0.8rem; text-decoration:none; transition:all 0.15s;
  }}
  .cat-btn:hover {{ color:var(--text); border-color:var(--accent); }}
  .cat-btn .cnt {{ color:var(--accent); margin-left:4px; }}

  /* Category sections */
  .cat-section {{ margin-bottom:32px; }}
  .cat-header {{
    font-size:1.1rem; font-weight:600; margin-bottom:8px;
    display:flex; align-items:center; gap:8px;
    padding-top:8px;
  }}
  .cat-header .cnt {{
    background:rgba(59,130,246,0.15); color:var(--accent);
    padding:2px 8px; border-radius:6px; font-size:0.8rem;
  }}

  /* Grid table */
  .grid-wrap {{
    background:var(--card); border:1px solid var(--border); border-radius:10px;
    overflow:hidden; max-width: calc(100vw - 40px);
  }}
  .grid-scroll {{ overflow-x:auto; }}
  table {{ border-collapse:separate; border-spacing:0; font-size:0.78rem; white-space:nowrap; }}
  th {{
    background:#0d1a2e; padding:6px 8px; text-align:left;
    font-weight:600; color:var(--muted); text-transform:uppercase; font-size:0.7rem;
    letter-spacing:0.03em; border-bottom:1px solid var(--border);
    position:sticky; top:0; z-index:10;
  }}
  th.sub {{ top:28px; }}
  td {{ padding:5px 8px; border-bottom:1px solid rgba(51,65,85,0.5); }}
  tr:hover {{ background:rgba(255,255,255,0.03); }}

  /* Sticky left columns */
  .sticky-col {{ position:sticky; z-index:5; background:var(--card); }}
  .sticky-col-0 {{ left:0; min-width:80px; }}
  .sticky-col-1 {{ left:80px; min-width:180px; border-right:1px solid var(--border); }}
  th.sticky-col {{ z-index:15; background:#0d1a2e; }}
  tr:hover .sticky-col {{ background:#253349; }}

  /* Attribute group headers */
  th.attr-group {{
    background:#162035; color:var(--accent);
    text-align:center; border-left:2px solid var(--accent);
    font-size:0.7rem;
  }}
  th.sub {{ font-size:0.65rem; color:var(--muted); text-align:center; }}
  th.sub-name {{ color:#60a5fa; }}
  th.sub-zoho {{ color:#fbbf24; }}
  th.sub-web {{ color:#4ade80; }}
  th.sub-google {{ color:#c084fc; }}

  /* Cell styles */
  .cell-ok {{ background:rgba(34,197,94,0.08); }}
  .cell-mismatch {{ background:rgba(239,68,68,0.12); }}
  .cell-partial {{ background:rgba(245,158,11,0.08); }}
  .cell-empty {{ color:var(--muted); }}
  .val-name {{ color:#93c5fd; }}
  .val-zoho {{ color:#fbbf24; }}
  .val-web {{ color:#86efac; }}
  .val-google {{ color:#d8b4fe; }}
  .tag-ok {{ color:var(--green); font-weight:600; }}
  .tag-mis {{ color:var(--red); font-weight:600; }}
  .tag-part {{ color:var(--orange); }}
  .tag-yes {{ color:var(--green); }}
  .tag-no {{ color:var(--red); }}

  .search {{ padding:8px; border-bottom:1px solid var(--border); display:flex; gap:10px; align-items:center; flex-wrap:wrap; }}
  .search input {{
    padding:6px 10px; border-radius:6px; border:1px solid var(--border);
    background:var(--bg); color:var(--text); font-size:0.85rem; width:300px;
  }}
  .filter-btn {{
    padding:5px 12px; border-radius:6px; cursor:pointer; font-size:0.8rem;
    border:1px solid var(--border); background:var(--card); color:var(--muted);
    transition:all 0.15s;
  }}
  .filter-btn:hover {{ border-color:var(--accent); color:var(--text); }}
  .filter-btn.active {{ background:var(--red); color:#fff; border-color:var(--red); }}

  /* Collapsible categories */
  .cat-header {{ cursor:pointer; user-select:none; }}
  .cat-header .toggle {{ font-size:0.8rem; color:var(--muted); margin-left:6px; }}
  .cat-body {{ display:none; }}
  .cat-section.expanded .cat-body {{ display:block; }}

  /* Done checkbox per row */
  th.done-col {{ width:28px; text-align:center; padding:0 4px; }}
  td.done-check {{ width:28px; text-align:center; padding:0 4px; }}
  td.done-check input {{ cursor:pointer; accent-color:var(--green); width:14px; height:14px; }}
  tr.row-done td:not(.done-check) {{ opacity:0.38; text-decoration:line-through; }}
</style>
</head>
<body>
<div class="container">
  <nav class="page-nav">
    <a href="comparison_report.html">Comparison Report</a>
    <a href="attribute_grid.html" class="current">Attribute Grid</a>
    <a href="head_type_report.html">Head Type (JD / Bobcat / Kubota)</a>
    <a href="all_products.html">All Products</a>
  </nav>
  <h1>Attribute Grid — JM Attachments</h1>
  <p class="subtitle">All products grouped by category. Attributes from: <span class="val-name">From Title (extracted from Zoho website title)</span> | <span class="val-zoho">Zoho (actual fields)</span> | <span class="val-web">Website (WooCommerce)</span> | <span class="val-google">Google Sheets</span></p>

  <div class="stats">
    <div class="stat">
      <div class="num" style="color:var(--accent)">{total_products}</div>
      <div class="label">Products</div>
    </div>
    <div class="stat">
      <div class="num" style="color:var(--accent)">{len(display_cats)}</div>
      <div class="label">Categories</div>
    </div>
    <div class="stat">
      <div class="num" style="color:var(--red)">{total_mismatches}</div>
      <div class="label">Mismatches</div>
    </div>
    <div class="stat">
      <div class="num" style="color:var(--orange)">{total_partial}</div>
      <div class="label">Partial (only 1 source)</div>
    </div>
    <div class="stat" style="display:flex;align-items:center;gap:6px;flex-wrap:wrap;">
      <button class="filter-btn" onclick="toggleGlobalMismatch(this)">Mismatches Only</button>
      <button class="filter-btn" onclick="expandAll()">Expand All</button>
      <button class="filter-btn" onclick="collapseAll()">Collapse All</button>
    </div>
    <div class="stat" style="display:flex;align-items:center;gap:6px;flex-wrap:wrap;">
      <span style="color:var(--muted);font-size:0.8rem;">Progress:</span>
      <button class="filter-btn" onclick="exportProgress()" title="Download progress.json">&#8595; Export</button>
      <label class="filter-btn" style="cursor:pointer;" title="Load progress.json">&#8593; Import<input type="file" accept=".json" style="display:none" onchange="importProgress(event)"></label>
      <button class="filter-btn" onclick="clearProgress()" title="Uncheck all">&#10005; Clear all</button>
      <span id="progress-count" style="color:var(--green);font-size:0.8rem;"></span>
    </div>
  </div>

  <div class="cat-nav">
"""

    for dc in display_cats:
        html += f'    <a class="cat-btn" href="#{dc["safe_id"]}" onclick="expandCat(\'{dc["safe_id"]}\')">{_esc(dc["label"])}<span class="cnt">{len(dc["df"])}</span></a>\n'

    html += '  </div>\n'

    # Category sections
    for dc in display_cats:
        label = dc["label"]
        safe_id = dc["safe_id"]
        df = dc["df"]
        attr_names = dc["attr_names"]

        cat_mismatches = sum(
            (df[f"{a} [Status]"] == "MISMATCH").sum()
            for a in attr_names if f"{a} [Status]" in df.columns
        )

        html += f'\n  <div class="cat-section" id="{safe_id}">\n'
        html += f'    <div class="cat-header" onclick="toggleCat(\'{safe_id}\')">'
        html += f'      <span>{_esc(label)} <span class="cnt">{len(df)} products</span>'
        if cat_mismatches:
            html += f' <span style="color:var(--red);font-size:0.8rem;">{cat_mismatches} mismatches</span>'
        html += f'</span>\n'
        html += f'      <span class="toggle">▶ expand</span>\n'
        html += f'    </div>\n'
        html += '    <div class="cat-body">\n'
        html += f'    <div class="search"><input type="text" placeholder="Search in {_esc(label)}..." onkeyup="filterCat(this,\'{safe_id}\')">'
        html += f' <button class="filter-btn" onclick="toggleMismatch(this,\'{safe_id}\')">Mismatch Only</button>'
        html += f'</div>\n'
        html += '    <div class="grid-wrap">\n'
        html += _render_grid_table(df, attr_names)
        html += '    </div>\n'
        html += '    </div>\n'
        html += '  </div>\n'

    html += """
</div>
<script>
function filterCat(input, catId) {
  const filter = input.value.toLowerCase();
  const section = document.getElementById(catId);
  const rows = section.querySelectorAll('tbody tr');
  const mismatchOnly = section.querySelector('.filter-btn.active');
  rows.forEach(row => {
    const textMatch = !filter || row.textContent.toLowerCase().includes(filter);
    const mismatchMatch = !mismatchOnly || row.dataset.hasMismatch === '1';
    row.style.display = (textMatch && mismatchMatch) ? '' : 'none';
  });
}

function toggleMismatch(btn, catId) {
  btn.classList.toggle('active');
  const section = document.getElementById(catId);
  const input = section.querySelector('.search input');
  filterCat(input, catId);
}

function toggleCat(catId) {
  const section = document.getElementById(catId);
  const toggle = section.querySelector('.toggle');
  section.classList.toggle('expanded');
  toggle.textContent = section.classList.contains('expanded') ? '▼ collapse' : '▶ expand';
}

function expandCat(catId) {
  const section = document.getElementById(catId);
  if (!section.classList.contains('expanded')) {
    section.classList.add('expanded');
    const t = section.querySelector('.toggle');
    if (t) t.textContent = '▼ collapse';
  }
}

function expandAll() {
  document.querySelectorAll('.cat-section').forEach(s => {
    s.classList.add('expanded');
    const t = s.querySelector('.toggle');
    if (t) t.textContent = '▼ collapse';
  });
}

function collapseAll() {
  document.querySelectorAll('.cat-section').forEach(s => {
    s.classList.remove('expanded');
    const t = s.querySelector('.toggle');
    if (t) t.textContent = '▶ expand';
  });
}

// Global mismatch filter
function toggleGlobalMismatch(btn) {
  btn.classList.toggle('active');
  const active = btn.classList.contains('active');
  document.querySelectorAll('.cat-section').forEach(section => {
    const rows = section.querySelectorAll('tbody tr');
    const localBtn = section.querySelector('.filter-btn');
    if (active) {
      localBtn.classList.add('active');
    } else {
      localBtn.classList.remove('active');
    }
    const input = section.querySelector('.search input');
    filterCat(input, section.id);
  });
}

// ── Progress checkboxes (per row / SKU) ─────────────────────────
const LS_KEY = 'jma_progress_grid';

function _loadProgress() {
  try { return JSON.parse(localStorage.getItem(LS_KEY) || '{}'); } catch { return {}; }
}
function _saveProgress(p) {
  localStorage.setItem(LS_KEY, JSON.stringify(p));
  _updateCount(p);
}
function _updateCount(p) {
  const done = Object.keys(p).filter(k => p[k]).length;
  const total = document.querySelectorAll('.done-cb[data-sku]').length;
  const el = document.getElementById('progress-count');
  if (el) el.textContent = done ? `${done} / ${total} done` : '';
}
function _applyProgress(p) {
  document.querySelectorAll('.done-cb[data-sku]').forEach(cb => {
    if (p[cb.dataset.sku]) {
      cb.checked = true;
      cb.closest('tr').classList.add('row-done');
    }
  });
  _updateCount(p);
}

function toggleRow(cb) {
  const sku = cb.dataset.sku;
  cb.closest('tr').classList.toggle('row-done', cb.checked);
  const p = _loadProgress();
  if (cb.checked) p[sku] = true; else delete p[sku];
  _saveProgress(p);
}

function exportProgress() {
  const p = _loadProgress();
  const blob = new Blob([JSON.stringify(p, null, 2)], {type: 'application/json'});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'progress.json';
  a.click();
}

function importProgress(e) {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = ev => {
    try {
      const p = JSON.parse(ev.target.result);
      _saveProgress(p);
      _applyProgress(p);
    } catch { alert('Invalid JSON file'); }
  };
  reader.readAsText(file);
  e.target.value = '';
}

function clearProgress() {
  if (!confirm('Uncheck all products?')) return;
  localStorage.removeItem(LS_KEY);
  document.querySelectorAll('.done-cb').forEach(cb => cb.checked = false);
  document.querySelectorAll('tr.row-done').forEach(tr => tr.classList.remove('row-done'));
  _updateCount({});
}

document.addEventListener('DOMContentLoaded', () => _applyProgress(_loadProgress()));
</script>
</body>
</html>"""

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"  Attribute grid: {output_path}")
    return output_path


def _esc(text):
    """Escape HTML."""
    if not text:
        return ""
    return (str(text)
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;"))


# ─── All Products flat view ──────────────────────────────────────────────────

# (display_name, attr_key) — exact order from the spec
_FLAT_COLUMNS = [
    # "Attachment Types" (Google) covered via alias in Machine Type
    ("Machine Type",           "Machine Type"),
    ("Head Style",             "Head Style"),
    ("Bucket Size",            "Bucket Size"),
    ("Carrier Weight Class",   "Carrier Weight Class"),
    ("Pin Size",               "Pin Size"),
    ("Front Pin Diameter",     "Front Pin Diameter"),
    ("Front Pin Length (mm)",  "Front Pin Length (mm)"),
    ("Rear Pin Diameter (mm)", "Rear Pin Diameter (mm)"),
    ("Rear Pin Length (mm)",   "Rear Pin Length (mm)"),
    ("Front Ear to Ear",       "Front Ear to Ear"),
    ("Rear Ear to Ear",        "Rear Ear to Ear"),
    ("Center to Center",       "Center to Center"),
    ("Shipping Height (in)",   "Shipping Height (in)"),
    ("Shipping Length (in)",   "Shipping Length (in)"),
    ("Shipping Width (in)",    "Shipping Width (in)"),
    ("Shipping Weight (lb)",   "Shipping Weight (lb)"),
]


def build_all_products_flat():
    """Build flat list of all active Zoho products with per-source attribute values."""
    print("  Building all-products flat table...")

    zoho_titles  = name_parser._load_zoho_website_titles()
    zoho_attrs   = name_parser._load_zoho_attrs()
    web_attrs    = name_parser._load_website_attrs()
    google_attrs = name_parser._load_google_attrs()

    zoho_cache = {}
    cache_path = os.path.join(config.DATA_DIR, "zoho_api_cache.json")
    if os.path.exists(cache_path):
        with open(cache_path, "r", encoding="utf-8") as f:
            for item in json.load(f):
                if str(item.get("status", "")).lower() == "inactive":
                    continue
                sku = str(item.get("sku", "")).strip().upper()
                if sku:
                    zoho_cache[sku] = item

    web_products = {}
    if os.path.exists(config.WEBSITE_CSV):
        df = pd.read_csv(config.WEBSITE_CSV, dtype=str, low_memory=False)
        for _, row in df.iterrows():
            sku = str(row.get("SKU", "")).strip().upper()
            if sku:
                web_products[sku] = {"name": row.get("Name", ""), "price": row.get("Regular price", "")}

    rows = []
    for sku in sorted(zoho_cache.keys()):
        item = zoho_cache[sku]
        title = (zoho_titles.get(sku, {}).get("website_title", "")
                 or str(item.get("name", "")))
        cat = item.get("category_name", "")
        z_attrs = zoho_attrs.get(sku, {})
        w_attrs = web_attrs.get(sku, {})
        g_attrs = google_attrs.get(sku, {})
        web_prod = web_products.get(sku, {})

        parsed = name_parser.parse_name(title, cat) if title else {}

        google_name = ""
        for key in ["Variation Name", "Variation Name/NA", "Product Name", "Name"]:
            if key in g_attrs:
                google_name = g_attrs[key]
                break

        row = {
            "SKU": sku,
            "Zoho Title": title[:100],
            "Status": str(item.get("status", "")).capitalize(),
            "Category": cat,
            "Website Name": str(web_prod.get("name", ""))[:100],
            "Google Name": str(google_name)[:100],
            "In Web": "Yes" if sku in web_attrs else "No",
            "In Google": "Yes" if sku in google_attrs else "No",
        }

        for display_name, attr_key in _FLAT_COLUMNS:
            data_only = display_name in DATA_ONLY_ATTRS or attr_key in DATA_ONLY_ATTRS
            nv = "" if data_only else _find_parsed_value(parsed, attr_key)
            zv = _find_source_value(z_attrs, attr_key)
            wv = _find_source_value(w_attrs, attr_key)
            gv = _find_source_value(g_attrs, attr_key)
            row[f"{display_name} [Name]"]   = nv
            row[f"{display_name} [Zoho]"]   = zv
            row[f"{display_name} [Web]"]    = wv
            row[f"{display_name} [Google]"] = gv
            src_vals = [zv, wv, gv] if data_only else [nv, zv, wv, gv]
            vals = [v for v in src_vals if v and str(v).strip() not in ("", "nan", "none")]
            if not vals:
                st = ""
            elif len(vals) == 1:
                st = "PARTIAL"
            else:
                st = "OK" if _all_match(vals) else "MISMATCH"
            row[f"{display_name} [Status]"] = st

        rows.append(row)

    print(f"    {len(rows)} products")
    return rows


_NAV_ALL = """  <nav class="page-nav">
    <a href="comparison_report.html">Comparison Report</a>
    <a href="attribute_grid.html">Attribute Grid</a>
    <a href="head_type_report.html">Head Type (JD / Bobcat / Kubota)</a>
    <a href="all_products.html">All Products</a>
  </nav>
"""


def generate_all_products_html(rows, output_path=None):
    """Generate all-products HTML table with per-source comparison columns."""
    output_path = output_path or os.path.join(config.OUTPUT_DIR, "all_products.html")
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    total = len(rows)
    mismatches = sum(
        1 for r in rows
        for col, _ in _FLAT_COLUMNS
        if r.get(f"{col} [Status]") == "MISMATCH"
    )

    SUB = 'font-size:0.65rem;background:rgba(0,0,0,0.2);text-align:center'

    def fmt_status(s):
        if not s or s.lower() in ("", "nan", "none"):
            return '<span style="color:var(--muted)">—</span>'
        cls = "tag-active" if s.lower() == "active" else "tag-inactive"
        return f'<span class="{cls}">{_esc(s)}</span>'

    def val_td(val, extra_cls=""):
        v = str(val).strip() if val else ""
        if v and v.lower() not in ("nan", "none"):
            return f'<td class="{extra_cls}">{_esc(v)}</td>'
        return f'<td class="{extra_cls}" style="color:var(--muted)">—</td>'

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>All Products — JM Attachments</title>
<style>
  :root {{
    --bg:#0f172a; --card:#1e293b; --border:#334155;
    --text:#e2e8f0; --muted:#94a3b8; --accent:#3b82f6;
    --green:#22c55e; --red:#ef4444; --orange:#f59e0b; --purple:#a78bfa;
  }}
  * {{ margin:0; padding:0; box-sizing:border-box; }}
  body {{
    font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',system-ui,sans-serif;
    background:var(--bg); color:var(--text); line-height:1.4; padding:20px;
    font-size:13px; overflow-x:hidden;
  }}
  h1 {{ font-size:1.5rem; margin-bottom:4px; }}
  .subtitle {{ color:var(--muted); margin-bottom:16px; }}
  .page-nav {{ display:flex; gap:8px; margin-bottom:20px; flex-wrap:wrap; }}
  .page-nav a {{
    padding:7px 16px; border-radius:8px; font-size:0.85rem; font-weight:500;
    text-decoration:none; border:1px solid var(--border); color:var(--muted);
    background:var(--card); transition:all 0.15s;
  }}
  .page-nav a:hover {{ color:var(--text); border-color:var(--accent); }}
  .page-nav a.current {{ color:var(--accent); border-color:var(--accent); background:rgba(59,130,246,0.1); }}
  .stats {{ display:flex; gap:12px; margin-bottom:20px; flex-wrap:wrap; }}
  .stat {{ background:var(--card); border:1px solid var(--border); border-radius:10px; padding:12px 18px; }}
  .stat .num {{ font-size:1.6rem; font-weight:700; }}
  .stat .label {{ color:var(--muted); font-size:0.8rem; }}
  .grid-wrap {{ background:var(--card); border:1px solid var(--border); border-radius:10px; overflow:hidden; max-width:calc(100vw - 40px); }}
  .search {{ padding:8px; border-bottom:1px solid var(--border); display:flex; gap:10px; align-items:center; flex-wrap:wrap; }}
  .search input {{
    padding:6px 10px; border-radius:6px; border:1px solid var(--border);
    background:var(--bg); color:var(--text); font-size:0.85rem; width:340px;
  }}
  .filter-btn {{
    padding:5px 12px; border-radius:6px; cursor:pointer; font-size:0.8rem;
    border:1px solid var(--border); background:var(--card); color:var(--muted); transition:all 0.15s;
  }}
  .filter-btn:hover {{ border-color:var(--accent); color:var(--text); }}
  .filter-btn.active {{ background:var(--red); color:#fff; border-color:var(--red); }}
  .grid-scroll {{ overflow-x:auto; }}
  table {{ border-collapse:separate; border-spacing:0; font-size:0.78rem; white-space:nowrap; }}
  th {{
    background:#0d1a2e; padding:6px 8px; text-align:left;
    font-weight:600; color:var(--muted); text-transform:uppercase; font-size:0.7rem;
    letter-spacing:0.03em; border-bottom:1px solid var(--border);
    position:sticky; top:0; z-index:10;
  }}
  th.sub {{ top:28px; }}
  td {{ padding:5px 8px; border-bottom:1px solid rgba(51,65,85,0.5); }}
  tr:hover {{ background:rgba(255,255,255,0.03); }}
  .sc {{ position:sticky; z-index:5; background:var(--card); }}
  .sc0 {{ left:0; min-width:80px; }}
  .sc1 {{ left:80px; min-width:220px; border-right:1px solid var(--border); }}
  th.sc {{ z-index:15; background:#0d1a2e; }}
  tr:hover .sc {{ background:#253349; }}
  th.attr-group {{
    background:#162035; color:var(--accent);
    text-align:center; border-left:2px solid var(--accent); font-size:0.7rem;
  }}
  th.sub {{ font-size:0.65rem; color:var(--muted); text-align:center; }}
  th.sub-name {{ color:#60a5fa; }} th.sub-zoho {{ color:#fbbf24; }}
  th.sub-web {{ color:#4ade80; }} th.sub-google {{ color:#c084fc; }}
  .cell-ok  {{ background:rgba(34,197,94,0.08); }}
  .cell-mismatch {{ background:rgba(239,68,68,0.12); }}
  .cell-partial  {{ background:rgba(245,158,11,0.08); }}
  .cell-empty {{ color:var(--muted); }}
  td.group-start, th.group-start {{ border-left:2px solid var(--accent); }}
  .val-name {{ color:#93c5fd; }} .val-zoho {{ color:#fbbf24; }}
  .val-web  {{ color:#86efac; }} .val-google {{ color:#d8b4fe; }}
  .tag-ok  {{ color:var(--green); font-weight:600; }}
  .tag-mis {{ color:var(--red);   font-weight:600; }}
  .tag-part {{ color:var(--orange); }}
  .tag-yes {{ color:var(--green); }} .tag-no {{ color:var(--red); }}
  .tag-active {{ color:var(--green); }} .tag-inactive {{ color:var(--red); }}
  .cat-nav {{ display:flex; flex-wrap:wrap; gap:6px; margin-bottom:20px; }}
  .cat-btn {{ padding:5px 12px; border-radius:8px; cursor:pointer; background:var(--card); border:1px solid var(--border); color:var(--muted); font-size:0.8rem; text-decoration:none; transition:all 0.15s; }}
  .cat-btn:hover {{ color:var(--text); border-color:var(--accent); }}
  .cat-btn .cnt {{ color:var(--accent); margin-left:4px; }}
  .cat-section {{ margin-bottom:28px; }}
  .cat-header {{ font-size:1.05rem; font-weight:600; margin-bottom:8px; display:flex; align-items:center; gap:8px; padding-top:8px; cursor:pointer; user-select:none; }}
  .cat-header .cnt {{ background:rgba(59,130,246,0.15); color:var(--accent); padding:2px 8px; border-radius:6px; font-size:0.8rem; }}
  .cat-header .toggle {{ font-size:0.8rem; color:var(--muted); margin-left:6px; }}
  .cat-body {{ display:none; }}
  .cat-section.expanded .cat-body {{ display:block; }}
  /* Done checkbox per row */
  th.done-col {{ width:28px; text-align:center; padding:0 4px; }}
  td.done-check {{ width:28px; text-align:center; padding:0 4px; }}
  td.done-check input {{ cursor:pointer; accent-color:var(--green); width:14px; height:14px; }}
  tr.row-done td:not(.done-check) {{ opacity:0.38; text-decoration:line-through; }}
  th.cell-check-col, td.cell-check-col {{ width:22px; text-align:center; padding:0 3px; }}
  .cell-cb {{ cursor:pointer; accent-color:var(--green); width:12px; height:12px; }}
  td[data-attr].cell-done {{ text-decoration:line-through; opacity:0.45; }}
  tr.row-done td[data-attr].cell-done {{ opacity:0.45; }}
  td.cell-check-col.cell-done {{ background:rgba(74,222,128,0.2) !important; opacity:1 !important; }}
</style>
</head>
<body>
{_NAV_ALL.replace('href="all_products.html"', 'href="all_products.html" class="current"')}
  <h1>All Products — JM Attachments</h1>
  <p class="subtitle">All active products · <span style="color:#93c5fd">From Title</span> | <span style="color:#fbbf24">Zoho</span> | <span style="color:#4ade80">Website</span> | <span style="color:#c084fc">Google Sheets</span></p>

"""

    # Group by category, sort by size desc
    from collections import defaultdict
    by_cat = defaultdict(list)
    for r in rows:
        by_cat[r["Category"]].append(r)
    cat_list = sorted(by_cat.keys(), key=lambda c: -len(by_cat[c]))

    _PIN_FLAT_NAMES = {"Front Pin Diameter", "Pin Size", "Rear Pin Diameter (mm)",
                       "Front Pin Length (mm)", "Rear Pin Length (mm)"}
    _FLAT_NO_PINS = [(col, key) for col, key in _FLAT_COLUMNS if col not in _PIN_FLAT_NAMES]

    def _flat_r_has_fpd(r):
        for suf in ["[Zoho]", "[Web]", "[Google]"]:
            v = str(r.get(f"Front Pin Diameter {suf}", "") or "").strip()
            if v and v.lower() not in ("nan", "none"):
                return True
        return False

    display_cats_flat = []
    for _cat in cat_list:
        _rows = by_cat[_cat]
        _wp = [r for r in _rows if _flat_r_has_fpd(r)]
        _wop = [r for r in _rows if not _flat_r_has_fpd(r)]
        if len(_wp) > 0 and len(_wop) > 0:
            _sb = re.sub(r'[^a-zA-Z0-9]', '', _cat)
            display_cats_flat.append({"label": f"{_cat} with pins",
                                      "safe_id": _sb + "withpins",
                                      "rows": _wp, "flat_cols": _FLAT_COLUMNS})
            display_cats_flat.append({"label": f"{_cat} without pins",
                                      "safe_id": _sb + "withoutpins",
                                      "rows": _wop, "flat_cols": _FLAT_NO_PINS})
        else:
            display_cats_flat.append({"label": _cat,
                                      "safe_id": re.sub(r'[^a-zA-Z0-9]', '', _cat),
                                      "rows": _rows, "flat_cols": _FLAT_COLUMNS})

    total_mis = sum(
        1 for r in rows
        for col, _ in _FLAT_COLUMNS
        if r.get(f"{col} [Status]") == "MISMATCH"
    )

    html += f"""  <div class="stats">
    <div class="stat"><div class="num" style="color:var(--accent)">{total}</div><div class="label">Active products</div></div>
    <div class="stat"><div class="num" style="color:var(--accent)">{len(display_cats_flat)}</div><div class="label">Categories</div></div>
    <div class="stat"><div class="num" style="color:var(--red)">{total_mis}</div><div class="label">Mismatches</div></div>
    <div class="stat" style="display:flex;align-items:center;gap:6px;flex-wrap:wrap;">
      <button class="filter-btn" onclick="toggleGlobalMismatch(this)">Mismatches Only</button>
      <button class="filter-btn" onclick="expandAll()">Expand All</button>
      <button class="filter-btn" onclick="collapseAll()">Collapse All</button>
    </div>
    <div class="stat" style="display:flex;align-items:center;gap:6px;flex-wrap:wrap;">
      <span style="color:var(--muted);font-size:0.8rem;">Progress:</span>
      <button class="filter-btn" onclick="exportProgress()" title="Download progress.json">&#8595; Export</button>
      <label class="filter-btn" style="cursor:pointer;" title="Load progress.json">&#8593; Import<input type="file" accept=".json" style="display:none" onchange="importProgress(event)"></label>
      <button class="filter-btn" onclick="clearProgress()" title="Uncheck all">&#10005; Clear all</button>
      <span id="progress-count" style="color:var(--green);font-size:0.8rem;"></span>
    </div>
  </div>

  <div class="cat-nav">
"""
    for dc in display_cats_flat:
        html += f'    <a class="cat-btn" href="#{dc["safe_id"]}" onclick="expandCat(\'{dc["safe_id"]}\')">{_esc(dc["label"])}<span class="cnt">{len(dc["rows"])}</span></a>\n'
    html += '  </div>\n'

    SUFFIXES = [("[Name]", "sub-name", "val-name"),
                ("[Zoho]", "sub-zoho", "val-zoho"),
                ("[Web]",  "sub-web",  "val-web"),
                ("[Google]","sub-google","val-google")]
    SUB_LABELS = {"[Name]": "From Title", "[Zoho]": "Zoho",
                  "[Web]": "Web", "[Google]": "Google"}

    def _has_any(cat_rows, col, suffix):
        return any(
            str(r.get(f"{col} {suffix}", "") or "").strip()
            not in ("", "nan", "none")
            for r in cat_rows
        )

    def _thead_cat(active_map):
        t = '        <thead>\n          <tr>\n'
        t += '            <th rowspan="2" class="sc sc0">SKU</th>\n'
        t += '            <th rowspan="2" class="sc sc1">Zoho Title (website)</th>\n'
        t += '            <th rowspan="2" class="done-col">&#10003;</th>\n'
        t += '            <th rowspan="2">Status</th>\n'
        t += '            <th rowspan="2">Website Name</th>\n'
        t += '            <th rowspan="2">Google Name</th>\n'
        t += '            <th rowspan="2">In Web</th>\n'
        t += '            <th rowspan="2">In Google</th>\n'
        for col, _ in _FLAT_COLUMNS:
            suf_list = active_map.get(col, [])
            if not suf_list:
                continue
            t += f'            <th class="attr-group" colspan="{len(suf_list)+1}">{_esc(col)}</th>\n'
        t += '          </tr>\n          <tr>\n'
        for col, _ in _FLAT_COLUMNS:
            suf_list = active_map.get(col, [])
            for i, (suf, sub_cls, _) in enumerate(suf_list):
                first_cls = " group-start" if i == 0 else ""
                t += f'            <th class="sub {sub_cls}{first_cls}">{SUB_LABELS[suf]}</th>\n'
            if suf_list:
                t += '            <th class="sub">St</th>\n'
                t += '            <th class="sub cell-check-col">&#9744;</th>\n'
        t += '          </tr>\n        </thead>\n'
        return t

    def _render_flat_table(rows_group, flat_cols):
        grp_active_map = {}
        for col, _ in flat_cols:
            active_suf = [(suf, sc, vc) for suf, sc, vc in SUFFIXES
                          if _has_any(rows_group, col, suf)]
            if active_suf:
                grp_active_map[col] = active_suf
        t = '      <div class="grid-scroll">\n        <table>\n'
        t += _thead_cat(grp_active_map)
        t += '          <tbody>\n'
        for r in rows_group:
            in_web = r["In Web"]
            in_google = r["In Google"]
            has_mis = any(r.get(f"{col} [Status]") == "MISMATCH" for col in grp_active_map)
            _sku = _esc(r["SKU"])
            t += f'          <tr data-has-mismatch="{1 if has_mis else 0}" data-sku="{_sku}">\n'
            t += f'            <td class="sc sc0"><strong>{_sku}</strong></td>\n'
            t += f'            <td class="sc sc1">{_esc(r["Zoho Title"])}</td>\n'
            t += f'            <td class="done-check"><input type="checkbox" class="done-cb" data-sku="{_sku}" onchange="toggleRow(this)"></td>\n'
            zoho_st = r["Status"]
            st_cls = "tag-inactive" if zoho_st.lower() == "inactive" else ""
            t += f'            <td class="{st_cls}">{_esc(zoho_st)}</td>\n'
            t += f'            <td class="val-web">{_esc(r["Website Name"])}</td>\n'
            t += f'            <td class="val-google">{_esc(r["Google Name"])}</td>\n'
            t += f'            <td class="{"tag-yes" if in_web=="Yes" else "tag-no"}">{in_web}</td>\n'
            t += f'            <td class="{"tag-yes" if in_google=="Yes" else "tag-no"}">{in_google}</td>\n'
            for col, _ in flat_cols:
                suf_list = grp_active_map.get(col)
                if not suf_list:
                    continue
                st = r.get(f"{col} [Status]", "")
                ccls = {"MISMATCH": "cell-mismatch", "PARTIAL": "cell-partial", "OK": "cell-ok"}.get(st, "")
                scls = {"MISMATCH": "tag-mis", "PARTIAL": "tag-part", "OK": "tag-ok"}.get(st, "cell-empty")
                icon = {"MISMATCH": "&#10007;", "PARTIAL": "~", "OK": "&#10003;"}.get(st, "")
                for i, (suf, _, vcls) in enumerate(suf_list):
                    first_cls = " group-start" if i == 0 else ""
                    v = str(r.get(f"{col} {suf}", "") or "")
                    if v and v.lower() not in ("nan", "none"):
                        t += f'            <td class="{ccls} {vcls}{first_cls}" data-attr="{_esc(col)}">{_esc(v)}</td>\n'
                    else:
                        t += f'            <td class="{ccls} cell-empty{first_cls}" data-attr="{_esc(col)}">&mdash;</td>\n'
                t += f'            <td class="{ccls} {scls}" data-attr="{_esc(col)}">{icon}</td>\n'
                t += f'            <td class="cell-check-col"><input type="checkbox" class="cell-cb" data-sku="{_sku}" data-col="{_esc(col)}" onchange="toggleCell(this)"></td>\n'
            t += '          </tr>\n'
        t += '          </tbody>\n        </table>\n      </div>\n'
        return t

    for dc in display_cats_flat:
        label = dc["label"]
        safe_id = dc["safe_id"]
        cat_rows = dc["rows"]
        flat_cols = dc["flat_cols"]

        cat_mis = sum(1 for r in cat_rows for col, _ in flat_cols
                      if r.get(f"{col} [Status]") == "MISMATCH")

        html += f'\n  <div class="cat-section" id="{safe_id}">\n'
        html += f'    <div class="cat-header" onclick="toggleCat(\'{safe_id}\')">'
        html += f'      <span>{_esc(label)} <span class="cnt">{len(cat_rows)} products</span>'
        if cat_mis:
            html += f' <span style="color:var(--red);font-size:0.8rem;">{cat_mis} mismatches</span>'
        html += f'</span>\n'
        html += f'      <span class="toggle">&#9654; expand</span>\n'
        html += f'    </div>\n'
        html += '    <div class="cat-body">\n'
        html += f'    <div class="search"><input type="text" placeholder="Search in {_esc(label)}..." onkeyup="filterCat(this,\'{safe_id}\')">'
        html += f' <button class="filter-btn" onclick="toggleMismatch(this,\'{safe_id}\')">Mismatch Only</button></div>\n'
        html += '    <div class="grid-wrap">\n'
        html += _render_flat_table(cat_rows, flat_cols)
        html += '    </div>\n'
        html += '    </div>\n  </div>\n'

    html += """
</div>
<script>
function filterCat(input, catId) {
  const filter = input.value.toLowerCase();
  const section = document.getElementById(catId);
  const mismatchOnly = section.querySelector('.filter-btn.active');
  section.querySelectorAll('tbody tr').forEach(row => {
    const textMatch = !filter || row.textContent.toLowerCase().includes(filter);
    const misMatch  = !mismatchOnly || row.dataset.hasMismatch === '1';
    row.style.display = (textMatch && misMatch) ? '' : 'none';
  });
}
function toggleMismatch(btn, catId) {
  btn.classList.toggle('active');
  filterCat(btn.closest('.grid-wrap').querySelector('input'), catId);
}
function toggleCat(catId) {
  const s = document.getElementById(catId);
  const t = s.querySelector('.toggle');
  s.classList.toggle('expanded');
  t.innerHTML = s.classList.contains('expanded') ? '&#9660; collapse' : '&#9654; expand';
}
function expandCat(catId) {
  const s = document.getElementById(catId);
  if (!s.classList.contains('expanded')) {
    s.classList.add('expanded');
    const t = s.querySelector('.toggle');
    if (t) t.innerHTML = '&#9660; collapse';
  }
}
function expandAll() {
  document.querySelectorAll('.cat-section').forEach(s => {
    s.classList.add('expanded');
    const t = s.querySelector('.toggle');
    if (t) t.innerHTML = '&#9660; collapse';
  });
}
function collapseAll() {
  document.querySelectorAll('.cat-section').forEach(s => {
    s.classList.remove('expanded');
    const t = s.querySelector('.toggle');
    if (t) t.innerHTML = '&#9654; expand';
  });
}
function toggleGlobalMismatch(btn) {
  btn.classList.toggle('active');
  const active = btn.classList.contains('active');
  document.querySelectorAll('.cat-section').forEach(section => {
    const lb = section.querySelector('.filter-btn');
    active ? lb.classList.add('active') : lb.classList.remove('active');
    filterCat(section.querySelector('.search input'), section.id);
  });
}

// ── Progress checkboxes (per row / SKU) + per-cell ───────────────
const LS_KEY = 'jma_progress_flat';

function _loadProgress() {
  try { return JSON.parse(localStorage.getItem(LS_KEY) || '{}'); } catch { return {}; }
}
function _saveProgress(p) {
  localStorage.setItem(LS_KEY, JSON.stringify(p));
  _updateCount(p);
}
function _updateCount(p) {
  const rows  = Object.keys(p).filter(k => p[k] && !k.includes('::')).length;
  const cells = Object.keys(p).filter(k => p[k] &&  k.includes('::')).length;
  const el = document.getElementById('progress-count');
  if (!el) return;
  if (!rows && !cells) { el.textContent = ''; return; }
  const parts = [];
  if (rows)  parts.push(`${rows} rows`);
  if (cells) parts.push(`${cells} cells`);
  el.textContent = parts.join(', ') + ' done';
}
function _applyProgress(p) {
  document.querySelectorAll('.done-cb[data-sku]').forEach(cb => {
    if (p[cb.dataset.sku]) {
      cb.checked = true;
      cb.closest('tr').classList.add('row-done');
    }
  });
  document.querySelectorAll('.cell-cb[data-col]').forEach(cb => {
    const key = cb.dataset.sku + '::' + cb.dataset.col;
    if (p[key]) {
      cb.checked = true;
      cb.closest('td').classList.add('cell-done');
      const col = cb.dataset.col;
      cb.closest('tr').querySelectorAll('td[data-attr]').forEach(td => {
        if (td.dataset.attr === col) td.classList.add('cell-done');
      });
    }
  });
  _updateCount(p);
}

function toggleCell(cb) {
  const key = cb.dataset.sku + '::' + cb.dataset.col;
  const on = cb.checked;
  const col = cb.dataset.col;
  cb.closest('td').classList.toggle('cell-done', on);
  cb.closest('tr').querySelectorAll(`td[data-attr]`).forEach(td => {
    if (td.dataset.attr === col) td.classList.toggle('cell-done', on);
  });
  const p = _loadProgress();
  if (on) p[key] = true; else delete p[key];
  _saveProgress(p);
}

function toggleRow(cb) {
  const sku = cb.dataset.sku;
  cb.closest('tr').classList.toggle('row-done', cb.checked);
  const p = _loadProgress();
  if (cb.checked) p[sku] = true; else delete p[sku];
  _saveProgress(p);
}

function exportProgress() {
  const p = _loadProgress();
  const blob = new Blob([JSON.stringify(p, null, 2)], {type: 'application/json'});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'progress.json';
  a.click();
}

function importProgress(e) {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = ev => {
    try {
      const p = JSON.parse(ev.target.result);
      _saveProgress(p);
      _applyProgress(p);
    } catch { alert('Invalid JSON file'); }
  };
  reader.readAsText(file);
  e.target.value = '';
}

function clearProgress() {
  if (!confirm('Uncheck all?')) return;
  localStorage.removeItem(LS_KEY);
  document.querySelectorAll('.done-cb').forEach(cb => cb.checked = false);
  document.querySelectorAll('tr.row-done').forEach(tr => tr.classList.remove('row-done'));
  document.querySelectorAll('.cell-cb').forEach(cb => cb.checked = false);
  document.querySelectorAll('td.cell-done').forEach(td => td.classList.remove('cell-done'));
  _updateCount({});
}

document.addEventListener('DOMContentLoaded', () => _applyProgress(_loadProgress()));
</script>
</body>
</html>"""

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"  All products table: {output_path}")
    return output_path


# ─── Head Type filtered page ────────────────────────────────────────────────

HEAD_TYPE_BRANDS = ["John Deere", "Bobcat", "Kubota"]

# All attribute keys that may carry head type / coupler type information
_HEAD_KEYS = [
    "Head Type", "Head Style", "Coupler Head Type",
    "Coupler Type", "Coupler Type/Filter",
]

# Extra columns to show in the table — mirrors COMMON_ATTRS (Head Style already shown as Head Type)
_HT_EXTRA_ATTRS = [
    "Bucket Size",
    "Pin Size",
    "Front Pin Length (mm)",
    "Rear Pin Length (mm)",
    "Carrier Weight Class",
    "Machine Type",
    "Shipping Weight (lb)",
    "Front Ear to Ear",
    "Rear Ear to Ear",
    "Fits To",
    "Capacity (yd³)",
]


def _get_head_type(attrs):
    """Return head type value from an attribute dict (any matching key)."""
    for key in _HEAD_KEYS:
        val = attrs.get(key, "")
        if val and str(val).strip():
            return str(val).strip()
    # Partial match fallback
    for k, v in attrs.items():
        if "head" in k.lower() or "coupler" in k.lower():
            if v and str(v).strip():
                return str(v).strip()
    return ""


def _brand_match(val):
    """Return the matched brand name or '' if none of the brands found."""
    if not val:
        return ""
    val_lower = val.lower()
    for brand in HEAD_TYPE_BRANDS:
        if brand.lower() in val_lower:
            return brand
    return ""


def build_head_type_grid():
    """
    Build list of products whose Head Type is John Deere, Bobcat, or Kubota.
    Returns list of dicts ready for HTML generation.
    """
    print("  Building Head Type grid (John Deere / Bobcat / Kubota)...")

    zoho_titles = name_parser._load_zoho_website_titles()
    zoho_attrs  = name_parser._load_zoho_attrs()
    web_attrs   = name_parser._load_website_attrs()
    google_attrs = name_parser._load_google_attrs()

    # Load zoho cache — skip inactive
    zoho_cache = {}
    cache_path = os.path.join(config.DATA_DIR, "zoho_api_cache.json")
    if os.path.exists(cache_path):
        with open(cache_path, "r", encoding="utf-8") as f:
            for item in json.load(f):
                if str(item.get("status", "")).lower() == "inactive":
                    continue
                sku = str(item.get("sku", "")).strip().upper()
                if sku:
                    zoho_cache[sku] = item

    # Load website names/prices
    web_products = {}
    if os.path.exists(config.WEBSITE_CSV):
        df = pd.read_csv(config.WEBSITE_CSV, dtype=str, low_memory=False)
        for _, row in df.iterrows():
            sku = str(row.get("SKU", "")).strip().upper()
            if sku:
                web_products[sku] = {
                    "name": row.get("Name", ""),
                    "price": row.get("Regular price", ""),
                }

    rows = []
    for sku in sorted(zoho_cache.keys()):
        item = zoho_cache[sku]
        zoho_data = zoho_titles.get(sku, {})
        title = zoho_data.get("website_title", "") or str(item.get("name", ""))

        z_attrs = zoho_attrs.get(sku, {})
        w_attrs = web_attrs.get(sku, {})
        g_attrs = google_attrs.get(sku, {})

        ht_zoho   = _get_head_type(z_attrs)
        ht_web    = _get_head_type(w_attrs)
        ht_google = _get_head_type(g_attrs)

        # Check if any source has a matching brand
        brand = (_brand_match(ht_zoho)
                 or _brand_match(ht_web)
                 or _brand_match(ht_google))
        if not brand:
            continue

        web_prod = web_products.get(sku, {})

        # Google product name
        google_name = ""
        for key in ["Variation Name", "Variation Name/NA", "Product Name", "Name"]:
            if key in g_attrs:
                google_name = g_attrs[key]
                break

        # Compute Head Type status
        ht_vals = [v for v in [ht_zoho, ht_web, ht_google] if v and str(v).strip() not in ("", "nan")]
        if not ht_vals:
            ht_status = ""
        elif len(ht_vals) == 1:
            ht_status = "PARTIAL"
        else:
            ht_status = "OK" if _all_match(ht_vals) else "MISMATCH"

        row = {
            "SKU": sku,
            "Brand": brand,
            "Category": item.get("category_name", ""),
            "Zoho Title": title[:80],
            "Website Name": str(web_prod.get("name", ""))[:80],
            "Google Name": str(google_name)[:80],
            "Zoho Price": item.get("rate", ""),
            "Website Price": web_prod.get("price", ""),
            "In Website": "Yes" if sku in web_attrs else "No",
            "In Google":  "Yes" if sku in google_attrs else "No",
            "Head Type [Zoho]":   ht_zoho,
            "Head Type [Web]":    ht_web,
            "Head Type [Google]": ht_google,
            "Head Type [Status]": ht_status,
        }

        # Extra attribute columns with status
        for attr in _HT_EXTRA_ATTRS:
            z_val = _find_source_value(z_attrs, attr)
            w_val = _find_source_value(w_attrs, attr)
            g_val = _find_source_value(g_attrs, attr)
            row[f"{attr} [Zoho]"]   = z_val
            row[f"{attr} [Web]"]    = w_val
            row[f"{attr} [Google]"] = g_val
            vals = [v for v in [z_val, w_val, g_val] if v and str(v).strip() not in ("", "nan")]
            if not vals:
                row[f"{attr} [Status]"] = ""
            elif len(vals) == 1:
                row[f"{attr} [Status]"] = "PARTIAL"
            else:
                row[f"{attr} [Status]"] = "OK" if _all_match(vals) else "MISMATCH"

        rows.append(row)

    print(f"    Found {len(rows)} products with Head Type John Deere / Bobcat / Kubota")
    return rows


def generate_head_type_html(rows, output_path=None):
    """Generate HTML page for products filtered by Head Type brand."""
    output_path = output_path or os.path.join(config.OUTPUT_DIR, "head_type_report.html")
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    by_brand = {b: [r for r in rows if r["Brand"] == b] for b in HEAD_TYPE_BRANDS}
    total = len(rows)

    brand_colors = {
        "John Deere": ("#166534", "#4ade80"),   # green
        "Bobcat":     ("#7c3aed", "#c084fc"),   # purple
        "Kubota":     ("#b45309", "#fbbf24"),   # orange
    }

    def fmt_price(val):
        try:
            return f"${float(val):,.2f}"
        except (ValueError, TypeError):
            return '<span style="color:var(--muted)">—</span>'

    def ht_cell(val, brand):
        if not val or str(val).strip() in ("", "nan", "none"):
            return '<span style="color:var(--muted)">—</span>'
        matched = _brand_match(val)
        if matched:
            _, color = brand_colors.get(matched, ("#1e293b", "#e2e8f0"))
            return f'<span style="color:{color};font-weight:600">{_esc(val)}</span>'
        return _esc(val)

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Head Type Products — John Deere / Bobcat / Kubota</title>
<style>
  :root {{
    --bg:#0f172a; --card:#1e293b; --border:#334155;
    --text:#e2e8f0; --muted:#94a3b8; --accent:#3b82f6;
    --green:#22c55e; --red:#ef4444; --orange:#f59e0b; --purple:#a78bfa;
  }}
  * {{ margin:0; padding:0; box-sizing:border-box; }}
  body {{
    font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',system-ui,sans-serif;
    background:var(--bg); color:var(--text); line-height:1.4; padding:20px;
    font-size:13px; overflow-x:hidden;
  }}
  .page-nav {{
    display:flex; gap:8px; margin-bottom:20px; flex-wrap:wrap;
  }}
  .page-nav a {{
    padding:7px 16px; border-radius:8px; font-size:0.85rem; font-weight:500;
    text-decoration:none; border:1px solid var(--border); color:var(--muted);
    background:var(--card); transition:all 0.15s;
  }}
  .page-nav a:hover {{ color:var(--text); border-color:var(--accent); }}
  .page-nav a.current {{ color:var(--accent); border-color:var(--accent); background:rgba(59,130,246,0.1); }}
  h1 {{ font-size:1.5rem; margin-bottom:4px; }}
  .subtitle {{ color:var(--muted); margin-bottom:20px; }}
  .stats {{ display:flex; gap:12px; margin-bottom:20px; flex-wrap:wrap; }}
  .stat {{ background:var(--card); border:1px solid var(--border); border-radius:10px; padding:12px 18px; }}
  .stat .num {{ font-size:1.6rem; font-weight:700; }}
  .stat .label {{ color:var(--muted); font-size:0.8rem; }}
  /* Sticky columns */
  .sc {{ position:sticky; z-index:5; background:var(--card); }}
  .sc0 {{ left:0; min-width:80px; }}
  .sc1 {{ left:80px; min-width:220px; border-right:1px solid var(--border); }}
  th.sc {{ z-index:15; background:#0d1823; }}
  tr:hover .sc {{ background:#253349; }}
  /* Tabs */
  .tabs {{ display:flex; gap:4px; border-bottom:2px solid var(--border); margin-bottom:0; }}
  .tab {{
    padding:10px 20px; cursor:pointer; border-radius:8px 8px 0 0;
    background:transparent; color:var(--muted); border:1px solid transparent;
    border-bottom:none; font-size:0.88rem; font-weight:500;
    transition:all 0.15s; position:relative; bottom:-2px;
  }}
  .tab:hover {{ color:var(--text); background:var(--card); }}
  .tab.active {{ color:var(--accent); background:var(--card); border-color:var(--border); }}
  .tab .badge {{
    background:var(--border); color:var(--muted);
    padding:2px 8px; border-radius:10px; font-size:0.72rem; margin-left:6px;
  }}
  .tab.active .badge {{ background:rgba(59,130,246,0.2); color:var(--accent); }}
  .tab-content {{ display:none; }}
  .tab-content.active {{ display:block; }}
  /* Table */
  .table-wrap {{
    background:var(--card); border:1px solid var(--border);
    border-radius:0 12px 12px 12px; overflow:hidden; margin-bottom:32px;
  }}
  .search-bar {{ padding:10px 14px; border-bottom:1px solid var(--border); }}
  .search-bar input {{
    padding:7px 11px; border-radius:7px; border:1px solid var(--border);
    background:var(--bg); color:var(--text); font-size:0.85rem; width:340px;
  }}
  .search-bar input::placeholder {{ color:var(--muted); }}
  .scrollable {{ overflow-x:auto; }}
  table {{ width:100%; border-collapse:collapse; font-size:0.8rem; white-space:nowrap; }}
  thead th {{
    background:rgba(0,0,0,0.25); padding:8px 10px; text-align:left;
    font-weight:600; color:var(--muted); text-transform:uppercase;
    font-size:0.68rem; letter-spacing:0.04em; border-bottom:1px solid var(--border);
    position:sticky; top:0; z-index:10;
  }}
  th.group-head {{
    color:var(--accent); text-align:center; border-left:2px solid var(--accent);
    background:#0d1823; font-size:0.68rem;
  }}
  td {{ padding:6px 10px; border-bottom:1px solid rgba(51,65,85,0.5); vertical-align:top; }}
  tr:hover {{ background:rgba(255,255,255,0.03); }}
  .tag-yes {{ color:var(--green); }}
  .tag-no  {{ color:var(--red); }}
  .price   {{ font-family:'SF Mono',Monaco,monospace; }}
  .cell-ok      {{ background:rgba(34,197,94,0.08); }}
  .cell-mismatch{{ background:rgba(239,68,68,0.12); }}
  .cell-partial {{ background:rgba(245,158,11,0.08); }}
  .tag-ok  {{ color:var(--green); font-weight:600; font-size:0.75rem; }}
  .tag-mis {{ color:var(--red);   font-weight:600; font-size:0.75rem; }}
  .tag-part{{ color:var(--orange);font-size:0.75rem; }}
  .val-zoho   {{ color:#fbbf24; }}
  .val-web    {{ color:#4ade80; }}
  .val-google {{ color:#c084fc; }}
</style>
</head>
<body>
<div style="max-width:100%;margin:0 auto">
  <nav class="page-nav">
    <a href="comparison_report.html">Comparison Report</a>
    <a href="attribute_grid.html">Attribute Grid</a>
    <a href="head_type_report.html" class="current">Head Type (JD / Bobcat / Kubota)</a>
    <a href="all_products.html">All Products</a>
  </nav>
  <h1>Head Type Products — John Deere / Bobcat / Kubota</h1>
  <p class="subtitle">Products filtered by Coupler Head Type from all sources (Zoho · Website · Google Sheets)</p>

  <div class="stats">
    <div class="stat"><div class="num" style="color:var(--accent)">{total}</div><div class="label">Total products</div></div>
"""
    for brand in HEAD_TYPE_BRANDS:
        _, color = brand_colors[brand]
        cnt = len(by_brand[brand])
        html += f'    <div class="stat"><div class="num" style="color:{color}">{cnt}</div><div class="label">{_esc(brand)}</div></div>\n'

    html += '  </div>\n\n  <div class="tabs">\n'
    html += f'    <div class="tab active" onclick="showTab(\'all\',this)">All Brands<span class="badge">{total}</span></div>\n'
    for brand in HEAD_TYPE_BRANDS:
        tab_id = brand.replace(" ", "")
        _, color = brand_colors[brand]
        html += f'    <div class="tab" onclick="showTab(\'{tab_id}\',this)" style="--tc:{color}">{_esc(brand)}<span class="badge">{len(by_brand[brand])}</span></div>\n'
    html += '  </div>\n'

    def _has_data(rows, key):
        return any(
            r.get(key, "") and str(r.get(key, "")).strip() not in ("", "nan", "none")
            for r in rows
        )

    def _val_cell(val):
        if val and str(val).strip() not in ("", "nan", "none"):
            return f'<td>{_esc(str(val))}</td>'
        return '<td style="color:var(--muted)">—</td>'

    SRC_COLORS = {"Zoho": "#fbbf24", "Web": "#4ade80", "Google": "#c084fc"}
    SUB_STYLE = 'font-size:0.65rem;background:rgba(0,0,0,0.2);text-align:center'

    def _build_table(tab_rows, table_id):
        # Pre-compute which sub-columns actually have data
        ht_srcs = [s for s in ["Zoho", "Web", "Google"]
                   if _has_data(tab_rows, f"Head Type [{s}]")]

        attr_srcs = {}  # attr -> list of sources with data
        visible_attrs = []
        for attr in _HT_EXTRA_ATTRS:
            srcs = [s for s in ["Zoho", "Web", "Google"]
                    if _has_data(tab_rows, f"{attr} [{s}]")]
            if srcs:
                attr_srcs[attr] = srcs
                visible_attrs.append(attr)

        # Head Type colspan = sources + 1 status column
        ht_colspan = len(ht_srcs) + 1

        out = f"""
  <div class="table-wrap">
    <div class="search-bar"><input type="text" placeholder="Search by SKU, name, category..." onkeyup="filterTable(this,'{table_id}')"></div>
    <div class="scrollable">
      <table id="{table_id}">
        <thead>
          <tr>
            <th rowspan="2" class="sc sc0">SKU</th>
            <th rowspan="2" class="sc sc1">Zoho Title</th>
            <th rowspan="2">Category</th>
            <th rowspan="2">Brand</th>
            <th rowspan="2">Website Name</th>
            <th rowspan="2">In Web</th>
            <th rowspan="2">In Google</th>
            <th rowspan="2" class="price">Zoho Price</th>
            <th rowspan="2" class="price">Web Price</th>
            <th class="group-head" colspan="{ht_colspan}">Head Type</th>
"""
        for attr in visible_attrs:
            attr_colspan = len(attr_srcs[attr]) + 1  # +1 for status
            out += f'            <th class="group-head" colspan="{attr_colspan}">{_esc(attr)}</th>\n'
        out += '          </tr>\n          <tr>\n'
        for src in ht_srcs:
            out += f'            <th style="{SUB_STYLE};color:{SRC_COLORS[src]}">{src}</th>\n'
        out += f'            <th style="{SUB_STYLE};color:var(--muted)">St</th>\n'
        for attr in visible_attrs:
            for src in attr_srcs[attr]:
                out += f'            <th style="{SUB_STYLE};color:{SRC_COLORS[src]}">{src}</th>\n'
            out += f'            <th style="{SUB_STYLE};color:var(--muted)">St</th>\n'
        out += '          </tr>\n        </thead>\n        <tbody>\n'

        STATUS_CLS = {"OK": "tag-ok", "MISMATCH": "tag-mis", "PARTIAL": "tag-part"}
        STATUS_ICO = {"OK": "&#10003;", "MISMATCH": "&#10007;", "PARTIAL": "~"}
        CELL_CLS   = {"OK": "cell-ok", "MISMATCH": "cell-mismatch", "PARTIAL": "cell-partial"}

        def _src_color_cell(val, src, status):
            cell_bg = CELL_CLS.get(status, "")
            color = SRC_COLORS.get(src, "")
            color_style = f"color:{color};" if color else ""
            if val and str(val).strip() not in ("", "nan", "none"):
                return f'<td class="{cell_bg}" style="{color_style}">{_esc(str(val))}</td>'
            return f'<td class="{cell_bg}" style="color:var(--muted)">—</td>'

        def _status_cell(status):
            cls = STATUS_CLS.get(status, "cell-empty")
            ico = STATUS_ICO.get(status, "")
            bg  = CELL_CLS.get(status, "")
            return f'<td class="{bg} {cls}">{ico}</td>'

        for r in tab_rows:
            brand = r["Brand"]
            _, color = brand_colors.get(brand, ("#1e293b", "#e2e8f0"))
            in_web = r["In Website"]
            in_google = r["In Google"]
            out += '          <tr>\n'
            out += f'            <td class="sc sc0"><strong>{_esc(r["SKU"])}</strong></td>\n'
            out += f'            <td class="sc sc1">{_esc(r["Zoho Title"])}</td>\n'
            out += f'            <td>{_esc(r["Category"])}</td>\n'
            out += f'            <td style="color:{color};font-weight:600">{_esc(brand)}</td>\n'
            out += f'            <td>{_esc(r["Website Name"])}</td>\n'
            out += f'            <td class="{"tag-yes" if in_web=="Yes" else "tag-no"}">{in_web}</td>\n'
            out += f'            <td class="{"tag-yes" if in_google=="Yes" else "tag-no"}">{in_google}</td>\n'
            out += f'            <td class="price">{fmt_price(r["Zoho Price"])}</td>\n'
            out += f'            <td class="price">{fmt_price(r["Website Price"])}</td>\n'
            ht_status = r.get("Head Type [Status]", "")
            for src in ht_srcs:
                out += f'            {_src_color_cell(r["Head Type [" + src + "]"], src, ht_status)}\n'
            out += f'            {_status_cell(ht_status)}\n'
            for attr in visible_attrs:
                attr_status = r.get(f"{attr} [Status]", "")
                for src in attr_srcs[attr]:
                    out += f'            {_src_color_cell(r.get(f"{attr} [{src}]", ""), src, attr_status)}\n'
                out += f'            {_status_cell(attr_status)}\n'
            out += '          </tr>\n'

        out += '        </tbody>\n      </table>\n    </div>\n  </div>\n'
        return out

    # All tab
    html += '\n  <div class="tab-content active" id="tab-all">'
    html += _build_table(rows, "tbl-all")
    html += '  </div>\n'

    # Per-brand tabs
    for brand in HEAD_TYPE_BRANDS:
        tab_id = brand.replace(" ", "")
        html += f'\n  <div class="tab-content" id="tab-{tab_id}">'
        html += _build_table(by_brand[brand], f"tbl-{tab_id}")
        html += '  </div>\n'

    html += """
<script>
function showTab(name, el) {
  document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  document.getElementById('tab-' + name).classList.add('active');
  el.classList.add('active');
}
function filterTable(input, tableId) {
  const filter = input.value.toLowerCase();
  document.querySelectorAll('#' + tableId + ' tbody tr').forEach(row => {
    row.style.display = row.textContent.toLowerCase().includes(filter) ? '' : 'none';
  });
}
</script>
</div>
</body>
</html>"""

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"  Head Type report: {output_path}")
    return output_path
