"""
Product data comparison parser settings for 3 sources.
Copy this file to config.py and fill in your credentials.
"""

import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

# ============================================================
# ZOHO INVENTORY API
# ============================================================
ZOHO_ORGANIZATION_ID = "YOUR_ORG_ID"
ZOHO_ACCESS_TOKEN = "YOUR_ACCESS_TOKEN"
# For automatic token refresh:
ZOHO_REFRESH_TOKEN = ""
ZOHO_CLIENT_ID = ""
ZOHO_CLIENT_SECRET = ""

ZOHO_API_BASE = "https://www.zohoapis.com/inventory/v1"
ZOHO_ITEMS_URL = f"{ZOHO_API_BASE}/items?organization_id={ZOHO_ORGANIZATION_ID}"

# Zoho API field mapping -> unified format
ZOHO_API_FIELDS = {
    "name": "product_name",
    "sku": "sku",
    "rate": "price",         # rate = selling price in Zoho Inventory
    "category_name": "category",
}

# Zoho CSV export (fallback when API is unavailable)
ZOHO_CSV = "/path/to/your/zoho_export.csv"

ZOHO_CSV_COLUMNS = {
    "Item Name": "product_name",
    "SKU": "sku",
    "Selling Price": "price",
    "Category Name": "category",
    "Status": "status",
}

# ============================================================
# WEBSITE (WooCommerce CSV export)
# ============================================================
WEBSITE_CSV = "/path/to/your/woocommerce_export.csv"

WEBSITE_COLUMNS = {
    "Name": "product_name",
    "SKU": "sku",
    "Regular price": "price",
    "Sale price": "sale_price",
    "Categories": "category",
    "Type": "type",
    "Published": "published",
}

WEBSITE_FILTER_PUBLISHED = True
WEBSITE_FILTER_TYPES = ["simple", "variable"]

# ============================================================
# GOOGLE SHEETS (Excel/CSV export)
# ============================================================
GOOGLE_XLSX = os.path.join(DATA_DIR, "google_products.xlsx")

GOOGLE_COLUMNS = {
    "Variation Name/NA": "product_name",
    "SKU": "sku",
    "Category/NA": "category",
}

# ============================================================
# GENERAL SETTINGS
# ============================================================
UNIFIED_COLUMNS = ["product_name", "sku", "price", "sale_price", "category", "status", "source"]

PRICE_TOLERANCE = 0.01
PRICE_TOLERANCE_PERCENT = 0.0
FUZZY_MATCH_THRESHOLD = 85

REPORT_FILE = os.path.join(OUTPUT_DIR, "comparison_report.xlsx")
