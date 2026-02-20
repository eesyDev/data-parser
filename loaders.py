"""
Loading and normalization of data from each source.
Unified format: product_name, sku, price, sale_price, category, source
"""

import os
import json
import pandas as pd
import re
import requests
import config


def _clean_price(value):
    """Clean price: remove currency symbols, spaces, commas."""
    if pd.isna(value):
        return None
    s = str(value).strip()
    s = re.sub(r"[^\d.]", "", s)
    try:
        return round(float(s), 2)
    except ValueError:
        return None


def _clean_text(value):
    """Clean text: remove extra spaces."""
    if pd.isna(value):
        return ""
    return str(value).strip()


def _normalize(df, column_map, source_name):
    """General DataFrame normalization."""
    available = {k: v for k, v in column_map.items() if k in df.columns}
    if not available:
        raise ValueError(
            f"None of the expected columns found in {source_name}. "
            f"Expected: {list(column_map.keys())}. "
            f"Found: {list(df.columns)}"
        )

    df = df.rename(columns=available)

    for col in config.UNIFIED_COLUMNS:
        if col not in df.columns and col != "source":
            df[col] = ""

    df["source"] = source_name
    df["product_name"] = df["product_name"].apply(_clean_text)
    df["sku"] = df["sku"].apply(lambda x: _clean_text(x).upper() if pd.notna(x) else "")
    df["price"] = df["price"].apply(_clean_price)
    df["sale_price"] = df["sale_price"].apply(_clean_price)
    df["category"] = df["category"].apply(_clean_text)
    if "status" in df.columns:
        df["status"] = df["status"].apply(_clean_text)

    df = df[df["product_name"].str.len() > 0]
    return df[config.UNIFIED_COLUMNS].reset_index(drop=True)


# ============================================================
# ZOHO INVENTORY (API)
# ============================================================

def _zoho_refresh_token():
    """Refresh access token using refresh token."""
    if not all([config.ZOHO_REFRESH_TOKEN, config.ZOHO_CLIENT_ID, config.ZOHO_CLIENT_SECRET]):
        return None

    resp = requests.post(
        "https://accounts.zoho.com/oauth/v2/token",
        params={
            "refresh_token": config.ZOHO_REFRESH_TOKEN,
            "client_id": config.ZOHO_CLIENT_ID,
            "client_secret": config.ZOHO_CLIENT_SECRET,
            "grant_type": "refresh_token",
        },
    )
    resp.raise_for_status()
    data = resp.json()
    if "access_token" in data:
        return data["access_token"]
    raise ValueError(f"Failed to refresh Zoho token: {data}")


def _zoho_get_access_token():
    """Get access token: from config or refresh via refresh token."""
    if config.ZOHO_ACCESS_TOKEN:
        return config.ZOHO_ACCESS_TOKEN
    return _zoho_refresh_token()


def load_zoho():
    """Load all products from Zoho Inventory API (with pagination)."""
    token = _zoho_get_access_token()
    if not token:
        raise ValueError(
            "No Zoho access token. Set ZOHO_ACCESS_TOKEN "
            "or ZOHO_REFRESH_TOKEN + CLIENT_ID + CLIENT_SECRET in config.py"
        )

    headers = {"Authorization": f"Zoho-oauthtoken {token}"}
    all_items = []
    page = 1
    per_page = 200

    while True:
        url = (
            f"{config.ZOHO_API_BASE}/items"
            f"?organization_id={config.ZOHO_ORGANIZATION_ID}"
            f"&page={page}&per_page={per_page}"
        )
        resp = requests.get(url, headers=headers)
        resp.raise_for_status()
        data = resp.json()

        if data.get("code") != 0:
            raise ValueError(f"Zoho API error: {data.get('message', data)}")

        items = data.get("items", [])
        if not items:
            break

        all_items.extend(items)
        print(f"    Zoho API: page {page}, fetched {len(items)} products")

        if not data.get("page_context", {}).get("has_more_page", False):
            break
        page += 1

    # Cache JSON for debugging
    cache_path = os.path.join(config.DATA_DIR, "zoho_api_cache.json")
    with open(cache_path, "w", encoding="utf-8") as f:
        json.dump(all_items, f, ensure_ascii=False, indent=2)

    # Convert to DataFrame — skip inactive items
    rows = []
    for item in all_items:
        if str(item.get("status", "")).lower() == "inactive":
            continue
        rows.append({
            "name": item.get("name", ""),
            "sku": item.get("sku", ""),
            "rate": item.get("rate", ""),
            "category_name": item.get("category_name", ""),
        })

    df = pd.DataFrame(rows)
    return _normalize(df, config.ZOHO_API_FIELDS, "zoho")


def load_zoho_csv(path=None):
    """Load data from Zoho CSV export (fallback when API is unavailable)."""
    path = path or config.ZOHO_CSV
    df = pd.read_csv(path, dtype=str, low_memory=False)
    # Filter inactive items
    if "Status" in df.columns:
        df = df[df["Status"].str.lower() != "inactive"]
    return _normalize(df, config.ZOHO_CSV_COLUMNS, "zoho")


# ============================================================
# WEBSITE (WooCommerce CSV)
# ============================================================

def load_website(path=None):
    """Load data from WooCommerce CSV export."""
    path = path or config.WEBSITE_CSV
    df = pd.read_csv(path, dtype=str, low_memory=False)

    if config.WEBSITE_FILTER_PUBLISHED and "Published" in df.columns:
        df = df[df["Published"].astype(str) == "1"]

    if config.WEBSITE_FILTER_TYPES and "Type" in df.columns:
        df = df[df["Type"].isin(config.WEBSITE_FILTER_TYPES)]

    return _normalize(df, config.WEBSITE_COLUMNS, "website")


# ============================================================
# GOOGLE SHEETS (Excel/CSV)
# ============================================================

def load_google(path=None):
    """Load data from Google Sheets (Excel with multiple sheets)."""
    path = path or config.GOOGLE_XLSX

    if path.endswith(".xlsx") or path.endswith(".xls"):
        xls = pd.ExcelFile(path)
        all_dfs = []
        for sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet, dtype=str)
            # Normalize column names: remove extra spaces
            df.columns = [c.strip() for c in df.columns]
            # Some sheets have the SKU column with or without spaces
            if "SKU" not in df.columns:
                found = False
                for col in df.columns:
                    if col.strip().upper() == "SKU" or col.strip() in ("", " "):
                        df = df.rename(columns={col: "SKU"})
                        found = True
                        break
                # Fallback: first column is always SKU
                if not found and len(df.columns) > 0:
                    df = df.rename(columns={df.columns[0]: "SKU"})
            if "SKU" in df.columns:
                all_dfs.append(df)
        if not all_dfs:
            raise ValueError("No sheets with SKU column found")
        # Merge — use common columns for normalization, others for attributes
        # Normalize each sheet separately and combine
        normalized = []
        for df in all_dfs:
            try:
                normalized.append(_normalize(df, config.GOOGLE_COLUMNS, "google"))
            except ValueError:
                continue  # Sheet missing required columns
        if not normalized:
            raise ValueError("Failed to normalize any sheet")
        result = pd.concat(normalized, ignore_index=True)
        # Remove duplicates by SKU
        result = result.drop_duplicates(subset="sku", keep="first")
        return result.reset_index(drop=True)
    else:
        df = pd.read_csv(path, dtype=str)
        return _normalize(df, config.GOOGLE_COLUMNS, "google")


# ============================================================
# LOAD ALL SOURCES
# ============================================================

def load_all():
    """Load all available sources."""
    sources = {}

    # Zoho — via API, fallback to CSV
    try:
        sources["zoho"] = load_zoho()
        print(f"  [OK] zoho: loaded {len(sources['zoho'])} products (API)")
    except Exception as e:
        print(f"  [WARN] zoho API: {e}")
        if os.path.exists(config.ZOHO_CSV):
            try:
                sources["zoho"] = load_zoho_csv()
                print(f"  [OK] zoho: loaded {len(sources['zoho'])} products (CSV)")
            except Exception as e2:
                print(f"  [ERROR] zoho CSV: {e2}")
        else:
            print(f"  [SKIP] zoho: no CSV fallback ({config.ZOHO_CSV})")

    # Website — CSV file
    if os.path.exists(config.WEBSITE_CSV):
        try:
            sources["website"] = load_website()
            print(f"  [OK] website: loaded {len(sources['website'])} products")
        except Exception as e:
            print(f"  [ERROR] website: {e}")
    else:
        print(f"  [SKIP] website: file not found ({config.WEBSITE_CSV})")

    # Google Sheets — Excel/CSV file
    if os.path.exists(config.GOOGLE_XLSX):
        try:
            sources["google"] = load_google()
            print(f"  [OK] google: loaded {len(sources['google'])} products")
        except Exception as e:
            print(f"  [ERROR] google: {e}")
    else:
        print(f"  [SKIP] google: file not found ({config.GOOGLE_XLSX})")

    return sources
