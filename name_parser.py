"""
Parse product attributes from website title (cf_website_title_only in Zoho).
Extract: size, category, weight class, machine type, pin size, coupler type, etc.
Then compare extracted attributes with actual attributes from Website/Google Sheets.
"""

import re
import json
import os
import pandas as pd
import config


# ============================================================
# ATTRIBUTE EXTRACTION FROM PRODUCT NAMES
# ============================================================

def parse_name(title, category=""):
    """
    Extract structured attributes from a product website title.
    Returns dict of extracted attributes.
    """
    if not title or not isinstance(title, str):
        return {}

    title = title.strip()
    attrs = {}

    # --- Size (inches) ---
    # "42" Ditching Bucket", "75" 4 in 1 Bucket", "24" Digging Bucket"
    # Also handle smart quotes: " (\u201c), " (\u201d), ″ (\u2033)
    # Quote is REQUIRED to avoid matching "800 Joules..." as bucket size
    m = re.match(r'^(\d+(?:\.\d+)?)["\u201c\u201d\u2033]\s', title)
    if m:
        attrs["Bucket Size"] = f'{m.group(1)}"'

    # --- Size for pins/bits (mm diameter) ---
    # "60 mm Diameter Bucket Pin", "30" Heavy Duty Rock Auger Bit"
    m_diam = re.match(r'^(\d+)\s*mm\s+Diameter', title)
    if m_diam:
        attrs["Diameter (mm)"] = f"{m_diam.group(1)} mm"

    # --- Pin Size ---
    # "45mm | 38mm Pins", "45mm / 38mm Pins", "45mm Pins"
    m_dual_pin = re.search(r'(\d+)\s*mm\s*[|/]\s*(\d+)\s*mm\s*[Pp]ins?', title)
    if m_dual_pin:
        attrs["Pin Size"] = f"{m_dual_pin.group(1)}mm | {m_dual_pin.group(2)}mm"
    else:
        m_pin = re.search(r'(\d+)\s*mm\s*[Pp]ins?', title)
        if m_pin:
            attrs["Pin Size"] = f"{m_pin.group(1)}mm"

    # --- Carrier Weight Class ---
    # "for 3 - 4.5 Tons Mini Excavators", "for 16 – 25 Tons Excavators"
    # Also handle en-dash (–), em-dash (—)
    m_weight = re.search(
        r'for\s+([\d.]+ ?[-\u2013\u2014] ?[\d.]+)\s*[Tt]ons?\s*(Mini\s+)?(\w+)',
        title
    )
    if m_weight:
        tons = m_weight.group(1).replace(" ", "").replace("\u2013", "-").replace("\u2014", "-")
        attrs["Carrier Weight Class"] = f"{tons} tons"
        machine = m_weight.group(3).strip()
        if m_weight.group(2):
            attrs["Machine Type"] = f"Mini {machine}"
        else:
            attrs["Machine Type"] = machine

    # --- Coupler Type ---
    coupler_patterns = [
        (r'(?i)John\s+Deere\s+Wedge\s+Lock\s+Coupler', "John Deere Wedge Lock Coupler"),
        (r'(?i)Kubota\s+Wedge\s+(?:Lock\s+)?(?:Coupler\s+)?Style', "Kubota Wedge Lock Coupler"),
        (r'(?i)Bobcat\s+X-?Change\s+Coupler', "Bobcat X-Change Coupler"),
        (r'(?i)Cat(?:erpillar)?\s+Pin\s+Grabber', "Cat Pin Grabber Coupler"),
        (r'(?i)Dual\s+Lock\s+Hydraulic\s+Quick\s+Coupler', "Dual Lock Hydraulic Quick Coupler"),
        (r'(?i)Spring\s+Manual\s+Quick\s+Coupler', "Spring Manual Quick Coupler"),
        (r'(?i)Pin\s+On\s+Style', "Pin On Style"),
    ]
    for pattern, coupler_name in coupler_patterns:
        if re.search(pattern, title):
            attrs["Coupler Type"] = coupler_name
            break

    # If no coupler detected and has "Pins" → likely Pin On
    if "Coupler Type" not in attrs and "Pin Size" in attrs:
        attrs["Coupler Type"] = "Pin On"

    # --- Product Type / Category (from name) ---
    type_patterns = [
        (r'(?i)Ditching Bucket', "Ditching Bucket"),
        (r'(?i)Digging Bucket', "Digging Bucket"),
        (r'(?i)Trenching Bucket', "Trenching Bucket"),
        (r'(?i)Banana Bucket', "Banana Bucket"),
        (r'(?i)Claw Bucket', "Claw Bucket"),
        (r'(?i)Tilt Bucket', "Tilt Bucket"),
        (r'(?i)Severe Duty (?:Skeleton )?Bucket', "Severe Duty Bucket"),
        (r'(?i)Heavy Duty (?:Digging |General Purpose )?Bucket', "Heavy Duty Bucket"),
        (r'(?i)Skeleton Bucket', "Skeleton Bucket"),
        (r'(?i)4 in 1 Bucket', "4 in 1 Bucket"),
        (r'(?i)Grapple Bucket', "Grapple Bucket"),
        (r'(?i)V-?Bottom.*Bucket', "V-Bottom Bucket"),
        (r'(?i)Ripper Tooth', "Ripper Tooth"),
        (r'(?i)Hydraulic Hammer', "Hydraulic Hammer"),
        (r'(?i)Post Driver Hammer', "Post Driver Hammer"),
        (r'(?i)Mechanical Thumb', "Mechanical Thumb"),
        (r'(?i)(?:QC )?Main Pin Hydraulic (?:Progressive )?Thumb', "Hydraulic Thumb"),
        (r'(?i)Mechanical Grapple', "Mechanical Grapple"),
        (r'(?i)Rotating Hydraulic Grapple', "Rotating Grapple"),
        (r'(?i)Hydraulic Quick Coupler', "Hydraulic Quick Coupler"),
        (r'(?i)Manual Quick Coupler', "Manual Quick Coupler"),
        (r'(?i)Brush Rake', "Brush Rake"),
        (r'(?i)Root Rake', "Root Rake"),
        (r'(?i)Bucket Rake', "Bucket Rake"),
        (r'(?i)Concrete Pulverizer', "Concrete Pulverizer"),
        (r'(?i)Hydraulic Shear', "Hydraulic Shear"),
        (r'(?i)Compaction Wheel', "Compaction Wheel"),
        (r'(?i)Plate Compactor', "Plate Compactor"),
        (r'(?i)Vibratory Roller', "Vibratory Roller"),
        (r'(?i)Pallet Fork', "Pallet Fork"),
        (r'(?i)Angle Broom', "Angle Broom"),
        (r'(?i)Sweeper Broom', "Sweeper Broom"),
        (r'(?i)Brush Cutter', "Brush Cutter"),
        (r'(?i)Auger (?:Drive )?Bit', "Auger Bit"),
        (r'(?i)Rock Auger Bit', "Rock Auger Bit"),
        (r'(?i)Bolt-?On Mount', "Bolt-On Mount"),
        (r'(?i)Aux Hydraulic Piping', "Aux Hydraulic Piping Kit"),
        (r'(?i)Bucket Pin', "Bucket Pin"),
        (r'(?i)Bucket Shim', "Bucket Shim"),
        (r'(?i)Bucket Tooth', "Bucket Tooth"),
        (r'(?i)Side Cutter', "Side Cutter"),
        (r'(?i)Tooth Retainer', "Tooth Retainer"),
        (r'(?i)Tooth Adapter', "Tooth Adapter"),
        (r'(?i)Tooth Pin', "Tooth Pin"),
        (r'(?i)Trencher', "Trencher"),
    ]
    for pattern, type_name in type_patterns:
        if re.search(pattern, title):
            attrs["Product Type"] = type_name
            break

    # --- Hex size for auger bits ---
    m_hex = re.search(r'(\d+)["\u201c\u201d\u2033]\s*Hex', title)
    if m_hex:
        attrs["Hex Size"] = f'{m_hex.group(1)}"'

    # --- Shim dimensions ---
    m_shim = re.match(r'^(\d+)\s*x\s*(\d+)\s*x\s*(\d+)\s*mm', title)
    if m_shim:
        attrs["Dimensions (mm)"] = f"{m_shim.group(1)} x {m_shim.group(2)} x {m_shim.group(3)} mm"

    # --- Bucket Pin length ---
    m_pin_len = re.search(r'with\s+(\d+)\s*mm\s+[Ll]ength', title)
    if m_pin_len:
        attrs["Length (mm)"] = f"{m_pin_len.group(1)} mm"

    # --- Fits To (Skid Steer) ---
    if re.search(r'(?i)Skid Steer', title):
        attrs["Fits To"] = "Skid Steer"

    # --- Attachment Types (derived from title context) ---
    if re.search(r'(?i)Skid\s+Steer', title):
        attrs["Attachment Types"] = "Skid Steer Attachment"
    elif re.search(r'(?i)Mini\s+Excavator', title):
        attrs["Attachment Types"] = "Mini Excavator Attachment"
    elif re.search(r'(?i)Excavator', title):
        attrs["Attachment Types"] = "Excavator Attachment"
    elif re.search(r'(?i)Backhoe', title):
        attrs["Attachment Types"] = "Backhoe Attachment"

    return attrs


# ============================================================
# LOAD ACTUAL ATTRIBUTES FROM SOURCES
# ============================================================

def _load_zoho_website_titles():
    """Load cf_website_title_only and category from Zoho cache."""
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
            "website_title": item.get("cf_website_title_only", "") or "",
            "category": item.get("category_name", "") or "",
        }
    return result


def _load_zoho_attrs():
    """Load actual attributes from Zoho cache (attribute_name/option + cf_ fields)."""
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
        attrs = {}
        # attribute_name1..3 / attribute_option_name1..3
        for i in range(1, 4):
            name = item.get(f"attribute_name{i}", "")
            val = item.get(f"attribute_option_name{i}", "")
            if name and val:
                attrs[name.strip()] = val.strip()
        # cf_ custom fields
        cf_map = {
            "cf_front_pin_size": "Front Pin Diameter (mm)",
            "cf_back_pin_size": "Rear Pin Diameter (mm)",
            "cf_coupler_head_type": "Coupler Head Type",
            "cf_product_weight": "Product Weight (kg)",
            "cf_product_weight_lbs": "Product Weight (lbs)",
            "cf_product_width_mm": "Product Width (mm)",
            "cf_product_width_in": "Product Width (in)",
            "cf_capacity_yds": "Product Capacity (yds)",
            "cf_product_capacity_m3": "Capacity (m³)",
            "cf_teeth_type": "Teeth Type",
            "cf_center_to_center": "Center to Center",
            "cf_front_ear_to_ear": "Front Ear to Ear",
            "cf_back_ear_to_ear": "Rear Ear to Ear",
            "cf_drain_holes": "Drain Holes",
            "cf_add_ons": "Add-on included",
            # cf_model_number intentionally omitted — not used
        }
        for cf_key, display_name in cf_map.items():
            val = item.get(cf_key, "")
            if val and str(val).strip() and str(val).lower() not in ("false", ""):
                attrs[display_name] = str(val).strip()
        result[sku] = attrs
    return result


def _load_website_attrs():
    """Load actual attributes from WooCommerce CSV."""
    if not os.path.exists(config.WEBSITE_CSV):
        return {}
    df = pd.read_csv(config.WEBSITE_CSV, dtype=str, low_memory=False)
    result = {}
    for _, row in df.iterrows():
        sku = str(row.get("SKU", "")).strip().upper()
        if not sku:
            continue
        attrs = {}
        for i in range(1, 24):
            name_col = f"Attribute {i} name"
            val_col = f"Attribute {i} value(s)"
            if name_col in df.columns and pd.notna(row.get(name_col)):
                attrs[row[name_col].strip()] = str(row.get(val_col, "")).strip()
        result[sku] = attrs
    return result


def _load_google_attrs():
    """Load actual attributes from Google Sheets."""
    path = config.GOOGLE_XLSX
    if not os.path.exists(path):
        return {}
    result = {}
    xls = pd.ExcelFile(path)
    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet, dtype=str)
        df.columns = [c.strip() for c in df.columns]
        sku_col = None
        if "SKU" in df.columns:
            sku_col = "SKU"
        else:
            for col in df.columns:
                if col.strip().upper() == "SKU":
                    sku_col = col
                    break
        if sku_col is None and len(df.columns) > 0:
            # Fallback: first column is always SKU if no SKU column found
            sku_col = df.columns[0]
        if sku_col is None:
            continue
        for _, row in df.iterrows():
            sku = str(row.get(sku_col, "")).strip().upper()
            if not sku or sku == "NAN":
                continue
            attrs = {}
            for col in df.columns:
                if col == sku_col:
                    continue
                val = row.get(col, "")
                if pd.notna(val) and str(val).strip() and str(val).lower() != "nan":
                    clean = col.replace("/NA", "").replace("/Filter", "").strip()
                    attrs[clean] = str(val).strip()
            result[sku] = attrs
    return result


# ============================================================
# NORMALIZE ATTRIBUTE VALUES FOR COMPARISON
# ============================================================

def _normalize_value(val):
    """Normalize attribute value for comparison."""
    if not val:
        return ""
    s = str(val).strip().lower()
    # Normalize smart quotes and primes to standard
    s = s.replace('\u201c', '"').replace('\u201d', '"').replace('\u2033', '"')
    s = s.replace('\u2018', "'").replace('\u2019', "'").replace('\u2032', "'")
    # Normalize dashes
    s = s.replace('\u2013', '-').replace('\u2014', '-')
    # Remove trailing units for comparison
    s = re.sub(r'\s*(mm|tons?|lbs?|in|inches|")\s*$', '', s)
    # Collapse spaces around mm (40mm == 40 mm)
    s = re.sub(r'(\d)\s*mm', r'\1mm', s)
    s = re.sub(r'\s+', ' ', s)
    # Fix escaped commas from WooCommerce CSV (\, -> ,)
    s = s.replace('\\,', ',')
    s = s.replace(',', '').replace('"', '').replace("'", '')
    s = s.rstrip('/')
    return s


# Mapping: parsed attribute name → possible actual attribute names in Website/Google
ATTR_MAP = {
    "Bucket Size": ["Bucket Size", "Bucket Size (in)", "Bucket Size (in)/Filter",
                    "Rake Width (in)", "Width (in)", "Bucket Width (in)",
                    "Grapple Width (in)", "Product Width (in)"],
    "Pin Size": ["Front Pin Diameter (mm)", "Front Pin Diameter", "Pin Size",
                 "Pin Diameter (mm)", "Front Pin Size"],
    "Carrier Weight Class": ["Carrier Weight Class", "Carrier Weight Class (tn)"],
    "Machine Type": ["Machine Type"],
    "Coupler Type": ["Coupler Head Type", "Coupler Type", "Coupler Type/Filter"],
    "Product Type": ["Bucket Type", "Category", "Attachment Types"],
    "Fits To": ["Fits To"],
    "Diameter (mm)": ["Pin Diameter (mm)", "Chisel Bit Diameter (mm)",
                       "Auger Bit Width (mm)"],
    "Length (mm)": ["Length (mm)", "Pin Length (mm)"],
    "Hex Size": ["Hex Size"],
    "Product Capacity (yds)": ["Capacity (yd³)", "Capacity (yd³)/Filter",
                                "Capacity ($yd^3$)", "Capacity (yds)"],
    "Capacity (m³)": ["Capacity (m³)", "Capacity ($m^3$)", "Capacity (m3)"],
    "Rear Pin Diameter (mm)": ["Rear Pin Diameter", "Rear Pin Diameter (mm)/Filter",
                                "Rear Pin Size (mm)", "Back Pin Size (mm)"],
    "Attachment Types": ["Attachment Types", "Attachment Types/NA"],
    "Product Weight (lbs)": ["Weight (lb)", "Weight (lbs)", "Rake Weight (lb)"],
    "Product Type": ["Bucket Type", "Category", "Attachment Types"],
}


def _match_value(parsed_val, actual_val):
    """Check if parsed value matches actual value (fuzzy)."""
    p = _normalize_value(parsed_val)
    a = _normalize_value(actual_val)
    if not p or not a:
        return None  # Can't compare
    # Direct match
    if p == a:
        return True
    # Dual value match: "45 | 38" should match if actual is either "45" or "38"
    if "|" in p:
        parts = [_normalize_value(x) for x in parsed_val.split("|")]
        for part in parts:
            if part == a:
                return True
            # Also check numbers
            p_nums = re.findall(r'[\d.]+', part)
            a_nums = re.findall(r'[\d.]+', a)
            if p_nums and a_nums and p_nums == a_nums:
                return True
        return False
    # Synonyms / equivalent terms
    synonyms = [
        ({"excavators", "excavator", "excavator attachment"}, True),
        ({"mini excavators", "mini excavator", "mini excavator attachment"}, True),
        ({"skid steer", "skid steer attachment", "skid steer loader"}, True),
        ({"pin on", "pin on style", "pin on coupler", "pin on style coupler",
          "no quick coupler pin on coupler", "backhoe pin on style",
          "backhoe pin on coupler", "bolt-on adapter"}, True),
        ({"bobcat x-change coupler", "bobcat x-change", "bobcat style", "bobcat x-change style"}, True),
        ({"john deere wedge lock coupler", "john deere wedge lock", "john deere style",
          "jd wedge lock", "deere wedge lock coupler", "deere style"}, True),
        ({"kubota wedge lock coupler", "kubota wedge lock", "kubota style",
          "kubota wedge style", "kubota wedge lock style"}, True),
        ({"cat pin grabber coupler", "cat pin grabber", "cat pin grabber style"}, True),
        ({"dual lock hydraulic quick coupler", "dual lock hqc", "dual lock coupler"}, True),
    ]
    for syn_set, result in synonyms:
        if p in syn_set and a in syn_set:
            return result
    # If actual contains multiple values (comma-separated), check each
    if "," in a:
        parts = [x.strip() for x in a.split(",")]
        for part in parts:
            if _normalize_value(part) == p:
                return True
            for syn_set, result in synonyms:
                if p in syn_set and _normalize_value(part) in syn_set:
                    return result
    # Number extraction and compare
    p_nums = re.findall(r'[\d.]+', p)
    a_nums = re.findall(r'[\d.]+', a)
    if p_nums and a_nums:
        # For ranges like "3-4.5" vs "3 - 4.5"
        if p_nums == a_nums:
            return True
        # Single number comparison
        if len(p_nums) == 1 and len(a_nums) == 1:
            try:
                return abs(float(p_nums[0]) - float(a_nums[0])) < 0.1
            except ValueError:
                pass
    # Contains check
    if p in a or a in p:
        return True
    return False


# ============================================================
# MAIN: COMPARE PARSED NAME ATTRIBUTES VS ACTUAL
# ============================================================

def compare_name_vs_attributes():
    """
    For each product:
    1. Parse Zoho website title → extracted attributes
    2. Compare with actual attributes from Website and Google Sheets
    3. Return DataFrame of mismatches grouped by category
    """
    print("  Loading data for name-vs-attribute comparison...")
    zoho_titles = _load_zoho_website_titles()
    web_attrs = _load_website_attrs()
    google_attrs = _load_google_attrs()

    rows = []
    for sku, zoho_data in zoho_titles.items():
        title = zoho_data["website_title"]
        category = zoho_data["category"]
        if not title:
            continue

        parsed = parse_name(title, category)
        if not parsed:
            continue

        web = web_attrs.get(sku, {})
        google = google_attrs.get(sku, {})

        for parsed_attr, parsed_val in parsed.items():
            # Find matching actual attribute names
            possible_names = ATTR_MAP.get(parsed_attr, [parsed_attr])

            # Check Website
            web_val = ""
            for name in possible_names:
                if name in web:
                    web_val = web[name]
                    break

            # Check Google
            google_val = ""
            for name in possible_names:
                if name in google:
                    google_val = google[name]
                    break

            # Compare
            web_match = _match_value(parsed_val, web_val) if web_val else None
            google_match = _match_value(parsed_val, google_val) if google_val else None

            # Only report mismatches or missing
            has_issue = False
            status = ""
            if web_match is False or google_match is False:
                has_issue = True
                status = "MISMATCH"
            elif web_match is None and google_match is None:
                has_issue = True
                status = "NOT FOUND"

            if has_issue:
                rows.append({
                    "sku": sku,
                    "category": category,
                    "product_title": title[:80],
                    "attribute": parsed_attr,
                    "from_name": parsed_val,
                    "website_value": web_val or "—",
                    "google_value": google_val or "—",
                    "status": status,
                    "web_match": "Yes" if web_match else ("No" if web_match is False else "—"),
                    "google_match": "Yes" if google_match else ("No" if google_match is False else "—"),
                })

    df = pd.DataFrame(rows)
    if not df.empty:
        df = df.sort_values(["category", "attribute", "sku"]).reset_index(drop=True)
    return df
