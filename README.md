# Product Data Comparison Parser

Tool for comparing product data across three sources: **Zoho Inventory**, **WooCommerce website**, and **Google Sheets**. Detects price discrepancies, SKU mismatches, and attribute inconsistencies, then generates interactive HTML reports.

## Reports (GitHub Pages)

- [Comparison Report](output/comparison_report.html) — side-by-side product comparison across all three sources
- [Attribute Grid](output/attribute_grid.html) — attribute completeness grid by category

## Features

- Loads products from Zoho Inventory API, WooCommerce CSV export, and Google Sheets (XLSX)
- Fuzzy name matching (Levenshtein) to detect duplicates and cross-source matches
- Compares prices, SKUs, categories, and product names
- Parses product names and checks against actual attributes for consistency
- Generates Excel and interactive HTML reports

## Setup

**1. Clone the repository**
```bash
git clone https://github.com/YOUR_USERNAME/data_parsing.git
cd data_parsing
```

**2. Install dependencies**
```bash
pip install -r requirements.txt
```

**3. Configure**
```bash
cp config.example.py config.py
# Edit config.py and fill in your Zoho credentials and file paths
```

**4. Add data files**

Place your source files in the `data/` folder:
- `data/google_products.xlsx` — Google Sheets export
- Zoho data is loaded via API (configured in `config.py`)
- WooCommerce CSV path is set in `config.py`

**5. Run**
```bash
python main.py
```

Reports are saved to `output/`.

## Requirements

- Python 3.8+
- pandas, openpyxl, thefuzz, python-Levenshtein
