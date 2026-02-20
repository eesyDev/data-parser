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
    "Pin Size",
    "Carrier Weight Class",
    "Machine Type",
    "Head Style",
]

# Category-specific extra attributes
CATEGORY_ATTRS = {
    "Bucket Tooth": ["Tooth Style", "Tooth Pin Part Number", "Serie"],
    "Bucket Pin": ["Diameter (mm)", "Length (mm)"],
    "Bucket Shims": ["Outer Diameter", "Interior Diameter", "Thickness"],
    "Auger Bits": ["Hex Size"],
    "Hydraulic Hammer": ["Energy Class (J)"],
    "Hammer Chisel Bits": ["Diameter (mm)"],
    "Hammer Moil Chisel Bits": ["Diameter (mm)"],
    "Hammer Wedge Chisel Bits": ["Diameter (mm)"],
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
    "Center to Center",
    "Drain Holes",
    "Product Weight (lbs)",
    "Product Weight (kg)",
    "Capacity (m³)",
    "Add-on included",
    "Add-ons Included",
    "Front Ear to Ear",
    "Rear Ear to Ear",
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
        attr_keys = [a for a in COMMON_ATTRS if not (a == "Bucket Size" and cat in NO_BUCKET_SIZE)]
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
                # Value from parsed name
                name_val = _find_parsed_value(parsed, attr_name)

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

        # Drop attribute columns where ALL 4 value sub-columns are empty across every row
        non_empty_attrs = []
        for attr in all_attr_names:
            has_data = any(
                df[f"{attr} {suffix}"].apply(
                    lambda x: bool(str(x).strip()) and str(x) not in ("", "nan")
                ).any()
                for suffix in ["[Name]", "[Zoho]", "[Web]", "[Google]"]
                if f"{attr} {suffix}" in df.columns
            )
            if has_data:
                non_empty_attrs.append(attr)

        grids[cat] = {
            "data": df,
            "attr_names": non_empty_attrs,
        }

    return grids


def _attr_base(key):
    """Strip trailing (unit) and normalize for dedup comparison."""
    return re.sub(r'\s*\([^)]*\)\s*$', '', key).strip().lower()


def _filter_important_attrs(all_keys, already_covered):
    """Filter to important/interesting attributes, skip noise."""
    skip_patterns = [
        "Variation Name", "Category", "Handling Unit", "Unit",
        "Shipping Length", "Shipping Width", "Shipping Height",
        "Shipping Weight", "Weight (lb)", "Weight (kg)",
        "Product Name", "Name",
        # Handled via Bucket Size alias — avoid duplicate column
        "Product Width (in)", "Product Width (mm)",
        # Handled via Capacity aliases
        "Capacity (yd", "Capacity ($yd",
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
    """Find value in parsed dict, matching by name."""
    if attr_name in parsed:
        return parsed[attr_name]
    # Fuzzy key match
    for k, v in parsed.items():
        if k.lower() in attr_name.lower() or attr_name.lower() in k.lower():
            return v
    return ""


def _find_source_value(attrs, attr_name):
    """Find value in source attributes dict."""
    if not attrs:
        return ""
    if attr_name in attrs:
        return attrs[attr_name]
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
        "Machine Type": ["Machine Type"],
        "Product Capacity (yds)": ["Capacity (yd³)", "Capacity (yd³)/Filter",
                                    "Capacity ($yd^3$)", "Capacity (yds)"],
        "Capacity (m³)": ["Capacity (m³)", "Capacity ($m^3$)", "Capacity (m3)"],
        "Attachment Types": ["Attachment Types/NA"],
        "Bucket Type": ["Category", "Category/NA"],
        "Category": ["Bucket Type"],
        "Product Weight (lbs)": ["Weight (lb)", "Weight (lbs)", "Rake Weight (lb)"],
        "Product Weight (kg)": ["Weight (kg)", "Rake Weight (kg)"],
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
  .container {{ max-width: 100%; margin: 0 auto; overflow-x: hidden; }}
  h1 {{ font-size: 1.5rem; margin-bottom: 4px; }}
  .subtitle {{ color: var(--muted); margin-bottom: 16px; }}
  .stats {{ display:flex; gap:12px; margin-bottom:20px; flex-wrap:wrap; }}
  .stat {{ background:var(--card); border:1px solid var(--border); border-radius:10px; padding:12px 16px; }}
  .stat .num {{ font-size:1.5rem; font-weight:700; }}
  .stat .label {{ color:var(--muted); font-size:0.8rem; }}

  /* Category nav */
  .cat-nav {{
    display:flex; flex-wrap:wrap; gap:6px; margin-bottom:20px;
    position:sticky; top:0; background:var(--bg); padding:8px 0; z-index:100;
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
  .grid-scroll {{ overflow-x:auto; max-height:600px; overflow-y:auto; }}
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
</style>
</head>
<body>
<div class="container">
  <h1>Attribute Grid — JM Attachments</h1>
  <p class="subtitle">All products grouped by category. Attributes from: <span class="val-name">Parsed (extracted from Zoho Title)</span> | <span class="val-zoho">Zoho (actual fields)</span> | <span class="val-web">Website (WooCommerce)</span> | <span class="val-google">Google Sheets</span></p>

  <div class="stats">
    <div class="stat">
      <div class="num" style="color:var(--accent)">{total_products}</div>
      <div class="label">Products</div>
    </div>
    <div class="stat">
      <div class="num" style="color:var(--accent)">{len(cat_list)}</div>
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
  </div>

  <div class="cat-nav">
"""

    for cat in cat_list:
        cnt = len(grids[cat]["data"])
        safe_id = re.sub(r'[^a-zA-Z0-9]', '', cat)
        html += f'    <a class="cat-btn" href="#{safe_id}" onclick="expandCat(\'{safe_id}\')">{_esc(cat)}<span class="cnt">{cnt}</span></a>\n'

    html += '  </div>\n'

    # Category sections
    for cat in cat_list:
        grid = grids[cat]
        df = grid["data"]
        attr_names = grid["attr_names"]
        safe_id = re.sub(r'[^a-zA-Z0-9]', '', cat)

        # Count issues in this category
        cat_mismatches = 0
        for a in attr_names:
            col = f"{a} [Status]"
            if col in df.columns:
                cat_mismatches += (df[col] == "MISMATCH").sum()

        html += f'\n  <div class="cat-section" id="{safe_id}">\n'
        html += f'    <div class="cat-header" onclick="toggleCat(\'{safe_id}\')">{_esc(cat)} <span class="cnt">{len(df)} products</span>'
        if cat_mismatches:
            html += f' <span style="color:var(--red);font-size:0.8rem;">{cat_mismatches} mismatches</span>'
        html += ' <span class="toggle">▶ expand</span></div>\n'
        html += '    <div class="cat-body">\n'
        html += '    <div class="grid-wrap">\n'
        html += f'      <div class="search"><input type="text" placeholder="Search in {_esc(cat)}..." onkeyup="filterCat(this,\'{safe_id}\')">'
        html += f' <button class="filter-btn" onclick="toggleMismatch(this,\'{safe_id}\')">Mismatch Only</button>'
        html += f'</div>\n'
        html += '      <div class="grid-scroll">\n'
        html += '        <table>\n'

        # Header row 1: attribute group names
        html += '          <thead>\n'
        html += '          <tr>\n'
        html += '            <th rowspan="2" class="sticky-col sticky-col-0">SKU</th>\n'
        html += '            <th rowspan="2" class="sticky-col sticky-col-1">Zoho Title (website)</th>\n'
        html += '            <th rowspan="2">Status</th>\n'
        html += '            <th rowspan="2">Website Name</th>\n'
        html += '            <th rowspan="2">Google Name</th>\n'
        html += '            <th rowspan="2">In Web</th>\n'
        html += '            <th rowspan="2">In Google</th>\n'

        for attr in attr_names:
            html += f'            <th class="attr-group" colspan="5">{_esc(attr)}</th>\n'

        html += '          </tr>\n'

        # Header row 2: sub-columns (Name, Zoho, Web, Google, Status)
        html += '          <tr>\n'
        for attr in attr_names:
            html += '            <th class="sub sub-name">Parsed</th>\n'
            html += '            <th class="sub sub-zoho">Zoho</th>\n'
            html += '            <th class="sub sub-web">Web</th>\n'
            html += '            <th class="sub sub-google">Google</th>\n'
            html += '            <th class="sub">St</th>\n'
        html += '          </tr>\n'
        html += '          </thead>\n'

        # Data rows
        html += '          <tbody>\n'
        for _, row in df.iterrows():
            has_mismatch = any(
                str(row.get(f"{a} [Status]", "")) == "MISMATCH"
                for a in attr_names
            )
            html += f'          <tr data-has-mismatch="{1 if has_mismatch else 0}">\n'
            html += f'            <td class="sticky-col sticky-col-0"><strong>{_esc(str(row.get("SKU","")))}</strong></td>\n'
            html += f'            <td class="sticky-col sticky-col-1">{_esc(str(row.get("Zoho Title","")))}</td>\n'
            zoho_st = str(row.get("Zoho Status", ""))
            st_color = "tag-no" if zoho_st.lower() == "inactive" else ""
            html += f'            <td class="{st_color}">{_esc(zoho_st)}</td>\n'
            html += f'            <td class="val-web">{_esc(str(row.get("Website Name","")))}</td>\n'
            html += f'            <td class="val-google">{_esc(str(row.get("Google Name","")))}</td>\n'

            in_web = str(row.get("In Website", ""))
            in_google = str(row.get("In Google", ""))
            html += f'            <td class="{"tag-yes" if in_web=="Yes" else "tag-no"}">{in_web}</td>\n'
            html += f'            <td class="{"tag-yes" if in_google=="Yes" else "tag-no"}">{in_google}</td>\n'

            for attr in attr_names:
                name_v = str(row.get(f"{attr} [Name]", "") or "")
                zoho_v = str(row.get(f"{attr} [Zoho]", "") or "")
                web_v = str(row.get(f"{attr} [Web]", "") or "")
                google_v = str(row.get(f"{attr} [Google]", "") or "")
                status = str(row.get(f"{attr} [Status]", "") or "")

                cell_cls = ""
                if status == "MISMATCH":
                    cell_cls = "cell-mismatch"
                elif status == "PARTIAL":
                    cell_cls = "cell-partial"
                elif status == "OK":
                    cell_cls = "cell-ok"

                html += f'            <td class="{cell_cls} val-name">{_esc(name_v) or "&mdash;"}</td>\n'
                html += f'            <td class="{cell_cls} val-zoho">{_esc(zoho_v) or "&mdash;"}</td>\n'
                html += f'            <td class="{cell_cls} val-web">{_esc(web_v) or "&mdash;"}</td>\n'
                html += f'            <td class="{cell_cls} val-google">{_esc(google_v) or "&mdash;"}</td>\n'

                st_cls = {"OK": "tag-ok", "MISMATCH": "tag-mis", "PARTIAL": "tag-part"}.get(status, "cell-empty")
                st_icon = {"OK": "&#10003;", "MISMATCH": "&#10007;", "PARTIAL": "~"}.get(status, "")
                html += f'            <td class="{cell_cls} {st_cls}">{st_icon}</td>\n'

            html += '          </tr>\n'

        html += '          </tbody>\n'
        html += '        </table>\n'
        html += '      </div>\n'
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
