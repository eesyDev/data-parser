"""
Generation of a polished HTML product comparison report.
"""

import os
import json
import pandas as pd
import config


def _load_zoho_attributes():
    """Load attributes from Zoho API cache."""
    cache_path = os.path.join(config.DATA_DIR, "zoho_api_cache.json")
    if not os.path.exists(cache_path):
        return {}

    with open(cache_path, "r", encoding="utf-8") as f:
        items = json.load(f)

    attrs = {}
    for item in items:
        sku = str(item.get("sku", "")).strip().upper()
        if not sku:
            continue
        attrs[sku] = {
            "brand": item.get("brand", ""),
            "description": (item.get("description") or "")[:200],
            "stock": item.get("stock_on_hand", ""),
            "weight": item.get("weight_with_unit", ""),
            "dimensions": item.get("dimensions_with_unit", ""),
            "status": item.get("status", ""),
            "cf_item_status": item.get("cf_item_status", ""),
        }
        # Zoho dynamic attributes
        for i in range(1, 4):
            name_key = f"attribute_name{i}"
            val_key = f"attribute_option_name{i}"
            aname = item.get(name_key, "")
            aval = item.get(val_key, "")
            if aname and aval:
                attrs[sku][aname] = aval
    return attrs


def _load_website_attributes():
    """Load attributes from WooCommerce CSV."""
    df = pd.read_csv(config.WEBSITE_CSV, dtype=str, low_memory=False)
    attrs = {}

    for _, row in df.iterrows():
        sku = str(row.get("SKU", "")).strip().upper()
        if not sku:
            continue

        item_attrs = {
            "sale_price": row.get("Sale price", ""),
            "short_description": str(row.get("Short description", ""))[:200],
            "weight": row.get("Weight (lbs)", ""),
        }

        # Dynamic attributes
        for i in range(1, 24):
            name_col = f"Attribute {i} name"
            val_col = f"Attribute {i} value(s)"
            if name_col in df.columns and pd.notna(row.get(name_col)):
                item_attrs[row[name_col]] = row.get(val_col, "")

        attrs[sku] = item_attrs
    return attrs


def _load_google_attributes():
    """Load attributes from Google Sheets (all sheets)."""
    path = config.GOOGLE_XLSX
    if not os.path.exists(path):
        return {}

    attrs = {}

    if path.endswith(".xlsx") or path.endswith(".xls"):
        xls = pd.ExcelFile(path)
        for sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet, dtype=str)
            df.columns = [c.strip() for c in df.columns]
            # Find SKU column
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
                item_attrs = {"Sheet": sheet}
                for col in df.columns:
                    if col == sku_col:
                        continue
                    val = row.get(col, "")
                    if pd.notna(val) and str(val).strip() and str(val).strip().lower() != "nan":
                        # Clean column name: remove /NA, /Filter suffixes for display
                        clean_name = col.replace("/NA", "").replace("/Filter", "").strip()
                        item_attrs[clean_name] = str(val).strip()
                attrs[sku] = item_attrs
    else:
        df = pd.read_csv(path, dtype=str)
        for _, row in df.iterrows():
            sku = str(row.get("SKU", "")).strip().upper()
            if not sku:
                continue
            item_attrs = {}
            for col in df.columns:
                val = row.get(col, "")
                if pd.notna(val) and str(val).strip():
                    item_attrs[col] = str(val).strip()
            attrs[sku] = item_attrs

    return attrs


def generate_html(results, output_path=None):
    """Generate HTML report."""
    output_path = output_path or os.path.join(config.OUTPUT_DIR, "comparison_report.html")
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    sources = results["sources"]
    merged = results["merged"]
    missing = results["missing"]
    price_diff = results["price_differences"]
    name_diff = results["name_differences"]
    attr_diff = results.get("attribute_differences", pd.DataFrame())
    name_vs_attrs = results.get("name_vs_attributes", pd.DataFrame())
    source_names = list(sources.keys())

    print("  Loading attributes...")
    zoho_attrs = _load_zoho_attributes()
    web_attrs = _load_website_attributes()
    google_attrs = _load_google_attributes()

    # Statistics
    total_skus = len(merged)
    in_all = merged
    for s in source_names:
        in_all = in_all[in_all[f"in_{s}"] == True]
    in_all_count = len(in_all)

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Product Comparison — JM Attachments</title>
<style>
  :root {{
    --bg: #0f172a;
    --card: #1e293b;
    --border: #334155;
    --text: #e2e8f0;
    --muted: #94a3b8;
    --accent: #3b82f6;
    --green: #22c55e;
    --red: #ef4444;
    --orange: #f59e0b;
    --purple: #a78bfa;
  }}
  * {{ margin:0; padding:0; box-sizing:border-box; }}
  body {{
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', system-ui, sans-serif;
    background: var(--bg);
    color: var(--text);
    line-height: 1.5;
    padding: 20px;
  }}
  .container {{ max-width: 1600px; margin: 0 auto; }}
  h1 {{
    font-size: 1.75rem;
    font-weight: 700;
    margin-bottom: 8px;
  }}
  .subtitle {{ color: var(--muted); margin-bottom: 24px; }}

  /* Stats cards */
  .stats {{
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
    gap: 12px;
    margin-bottom: 32px;
  }}
  .stat-card {{
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 16px 20px;
  }}
  .stat-card .num {{
    font-size: 2rem;
    font-weight: 700;
    line-height: 1.2;
  }}
  .stat-card .label {{ color: var(--muted); font-size: 0.85rem; }}
  .stat-card.green .num {{ color: var(--green); }}
  .stat-card.red .num {{ color: var(--red); }}
  .stat-card.orange .num {{ color: var(--orange); }}
  .stat-card.blue .num {{ color: var(--accent); }}
  .stat-card.purple .num {{ color: var(--purple); }}

  /* Tabs */
  .tabs {{
    display: flex;
    gap: 4px;
    margin-bottom: 0;
    border-bottom: 2px solid var(--border);
    padding-bottom: 0;
  }}
  .tab {{
    padding: 10px 20px;
    cursor: pointer;
    border-radius: 8px 8px 0 0;
    background: transparent;
    color: var(--muted);
    border: 1px solid transparent;
    border-bottom: none;
    font-size: 0.9rem;
    font-weight: 500;
    transition: all 0.15s;
    position: relative;
    bottom: -2px;
  }}
  .tab:hover {{ color: var(--text); background: var(--card); }}
  .tab.active {{
    color: var(--accent);
    background: var(--card);
    border-color: var(--border);
  }}
  .tab .badge {{
    background: var(--border);
    color: var(--muted);
    padding: 2px 8px;
    border-radius: 10px;
    font-size: 0.75rem;
    margin-left: 6px;
  }}
  .tab.active .badge {{ background: rgba(59,130,246,0.2); color: var(--accent); }}

  .tab-content {{ display: none; }}
  .tab-content.active {{ display: block; }}

  /* Tables */
  .table-wrap {{
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 0 12px 12px 12px;
    overflow: hidden;
    margin-bottom: 32px;
  }}
  .search-bar {{
    padding: 12px 16px;
    border-bottom: 1px solid var(--border);
  }}
  .search-bar input {{
    width: 100%;
    max-width: 400px;
    padding: 8px 12px;
    border-radius: 8px;
    border: 1px solid var(--border);
    background: var(--bg);
    color: var(--text);
    font-size: 0.9rem;
  }}
  .search-bar input::placeholder {{ color: var(--muted); }}
  table {{
    width: 100%;
    border-collapse: collapse;
    font-size: 0.85rem;
  }}
  thead th {{
    background: rgba(0,0,0,0.2);
    padding: 10px 12px;
    text-align: left;
    font-weight: 600;
    color: var(--muted);
    text-transform: uppercase;
    font-size: 0.75rem;
    letter-spacing: 0.05em;
    border-bottom: 1px solid var(--border);
    position: sticky;
    top: 0;
    white-space: nowrap;
  }}
  td {{
    padding: 8px 12px;
    border-bottom: 1px solid var(--border);
    vertical-align: top;
  }}
  tr:hover {{ background: rgba(255,255,255,0.03); }}
  .scrollable {{ max-height: 70vh; overflow-y: auto; }}

  /* Tags */
  .tag {{
    display: inline-block;
    padding: 2px 8px;
    border-radius: 6px;
    font-size: 0.75rem;
    font-weight: 500;
  }}
  .tag-zoho {{ background: rgba(59,130,246,0.15); color: #60a5fa; }}
  .tag-website {{ background: rgba(34,197,94,0.15); color: #4ade80; }}
  .tag-google {{ background: rgba(168,85,247,0.15); color: #c084fc; }}
  .tag-yes {{ background: rgba(34,197,94,0.15); color: #4ade80; }}
  .tag-no {{ background: rgba(239,68,68,0.15); color: #f87171; }}

  .price {{ font-family: 'SF Mono', Monaco, monospace; }}
  .price-diff {{ color: var(--red); font-weight: 600; }}
  .price-match {{ color: var(--green); }}
  .text-muted {{ color: var(--muted); }}
  .text-warn {{ color: var(--orange); }}

  /* Attribute comparison */
  .attr-table {{ margin-top: 4px; }}
  .attr-table td {{
    padding: 2px 8px;
    border: none;
    font-size: 0.8rem;
  }}
  .attr-label {{ color: var(--muted); white-space: nowrap; }}
  .attr-diff {{ background: rgba(239,68,68,0.1); border-radius: 4px; }}

  /* Expand row */
  .expand-btn {{
    cursor: pointer;
    color: var(--accent);
    font-size: 0.8rem;
    text-decoration: none;
    user-select: none;
  }}
  .expand-btn:hover {{ text-decoration: underline; }}
  .detail-row {{ display: none; }}
  .detail-row.open {{ display: table-row; }}
  .detail-cell {{
    padding: 12px 24px;
    background: rgba(0,0,0,0.15);
  }}
  .detail-grid {{
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
    gap: 16px;
  }}
  .detail-source {{
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 12px;
  }}
  .detail-source h4 {{
    font-size: 0.85rem;
    margin-bottom: 8px;
    display: flex;
    align-items: center;
    gap: 8px;
  }}
  .detail-source table {{ font-size: 0.8rem; }}
  .detail-source td {{ padding: 3px 8px; border: none; }}
</style>
</head>
<body>
<div class="container">
  <h1>Product Comparison — JM Attachments</h1>
  <p class="subtitle">Zoho Inventory vs Website (WooCommerce){' vs Google Sheets' if 'google' in source_names else ''}</p>

  <div class="stats">
    <div class="stat-card blue">
      <div class="num">{total_skus}</div>
      <div class="label">Unique SKUs</div>
    </div>
    <div class="stat-card green">
      <div class="num">{in_all_count}</div>
      <div class="label">Present in All Sources</div>
    </div>
    <div class="stat-card orange">
      <div class="num">{len(missing)}</div>
      <div class="label">Missing Products</div>
    </div>
    <div class="stat-card red">
      <div class="num">{len(price_diff)}</div>
      <div class="label">Price Discrepancies</div>
    </div>
    <div class="stat-card purple">
      <div class="num">{len(name_diff)}</div>
      <div class="label">Name Discrepancies</div>
    </div>
    <div class="stat-card">
      <div class="num" style="color: #fb923c;">{len(attr_diff)}</div>
      <div class="label">Attribute Discrepancies</div>
    </div>
    <div class="stat-card">
      <div class="num" style="color: #f472b6;">{len(name_vs_attrs)}</div>
      <div class="label">Name vs Attributes</div>
    </div>
"""

    for name, df in sources.items():
        tag = name
        html += f"""    <div class="stat-card">
      <div class="num">{len(df)}</div>
      <div class="label"><span class="tag tag-{tag}">{name.upper()}</span></div>
    </div>
"""

    html += """  </div>

  <!-- Tabs -->
  <div class="tabs">
    <div class="tab active" onclick="showTab('prices')">Price Discrepancies<span class="badge">""" + str(len(price_diff)) + """</span></div>
    <div class="tab" onclick="showTab('missing')">Missing Products<span class="badge">""" + str(len(missing)) + """</span></div>
    <div class="tab" onclick="showTab('names')">Name Discrepancies<span class="badge">""" + str(len(name_diff)) + """</span></div>
    <div class="tab" onclick="showTab('attrs')">Attribute Discrepancies<span class="badge">""" + str(len(attr_diff)) + """</span></div>
    <div class="tab" onclick="showTab('namecheck')">Name vs Attributes<span class="badge">""" + str(len(name_vs_attrs)) + """</span></div>
    <div class="tab" onclick="showTab('all')">All Products<span class="badge">""" + str(total_skus) + """</span></div>
  </div>
"""

    # === TAB: Price Differences ===
    html += """
  <div class="tab-content active" id="tab-prices">
    <div class="table-wrap">
      <div class="search-bar"><input type="text" placeholder="Search by SKU or product name..." onkeyup="filterTable(this, 'price-table')"></div>
      <div class="scrollable">
        <table id="price-table">
          <thead><tr>
            <th>SKU</th>
            <th>Product Name</th>
"""
    for s in source_names:
        html += f'            <th>Price {s.upper()}</th>\n'
    html += """            <th>Zoho Status</th>
            <th>Difference</th>
            <th>Details</th>
          </tr></thead>
          <tbody>
"""

    if not price_diff.empty:
        for idx, row in price_diff.iterrows():
            sku = str(row.get("sku", ""))
            name = str(row.get("product_name", ""))[:80]
            diff = row.get("price_diff", 0)
            zoho_status = _get_zoho_status(sku, merged)

            html += f'          <tr>\n            <td><strong>{_esc(sku)}</strong></td>\n            <td>{_esc(name)}</td>\n'
            for s in source_names:
                p = row.get(f"price_{s}")
                html += f'            <td class="price">{_fmt_price(p)}</td>\n'
            html += f'            <td>{_fmt_status(zoho_status)}</td>\n'
            html += f'            <td class="price price-diff">${diff:,.2f}</td>\n'
            html += f'            <td><span class="expand-btn" onclick="toggleDetail(\'pd-{idx}\')">&#9660; attributes</span></td>\n'
            html += '          </tr>\n'

            # Detail row
            html += _build_detail_row(f"pd-{idx}", sku, source_names, zoho_attrs, web_attrs, google_attrs)

    html += """          </tbody>
        </table>
      </div>
    </div>
  </div>
"""

    # === TAB: Missing ===
    html += """
  <div class="tab-content" id="tab-missing">
    <div class="table-wrap">
      <div class="search-bar"><input type="text" placeholder="Search..." onkeyup="filterTable(this, 'missing-table')"></div>
      <div class="scrollable">
        <table id="missing-table">
          <thead><tr>
            <th>SKU</th>
            <th>Product Name</th>
"""
    for s in source_names:
        html += f'            <th>{s.upper()}</th>\n'
    html += """            <th>Zoho Status</th>
            <th>Details</th>
          </tr></thead>
          <tbody>
"""

    if not missing.empty:
        for idx, row in missing.iterrows():
            sku = str(row.get("sku", ""))
            name = str(row.get("product_name", ""))[:80]
            present = str(row.get("present_in", "")).split(", ")
            zoho_status = str(row.get("zoho_status", "")) if pd.notna(row.get("zoho_status")) else ""

            html += f'          <tr>\n            <td><strong>{_esc(sku)}</strong></td>\n            <td>{_esc(name)}</td>\n'
            for s in source_names:
                if s in present:
                    html += '            <td><span class="tag tag-yes">Yes</span></td>\n'
                else:
                    html += '            <td><span class="tag tag-no">No</span></td>\n'
            html += f'            <td>{_fmt_status(zoho_status)}</td>\n'
            html += f'            <td><span class="expand-btn" onclick="toggleDetail(\'ms-{idx}\')">&#9660; attributes</span></td>\n'
            html += '          </tr>\n'
            html += _build_detail_row(f"ms-{idx}", sku, source_names, zoho_attrs, web_attrs, google_attrs)

    html += """          </tbody>
        </table>
      </div>
    </div>
  </div>
"""

    # === TAB: Name Differences ===
    html += """
  <div class="tab-content" id="tab-names">
    <div class="table-wrap">
      <div class="scrollable">
        <table>
          <thead><tr>
            <th>SKU</th>
"""
    for s in source_names:
        html += f'            <th>Name in {s.upper()}</th>\n'
    html += """          </tr></thead>
          <tbody>
"""

    if not name_diff.empty:
        for _, row in name_diff.iterrows():
            html += f'          <tr>\n            <td><strong>{_esc(str(row.get("sku", "")))}</strong></td>\n'
            for s in source_names:
                html += f'            <td>{_esc(str(row.get(f"name_{s}", "")))}</td>\n'
            html += '          </tr>\n'

    html += """          </tbody>
        </table>
      </div>
    </div>
  </div>
"""

    # === TAB: Attribute Differences ===
    html += """
  <div class="tab-content" id="tab-attrs">
    <div class="table-wrap">
      <div class="search-bar"><input type="text" placeholder="Search by SKU, name or attribute..." onkeyup="filterTable(this, 'attr-table')"></div>
      <div class="scrollable">
        <table id="attr-table">
          <thead><tr>
            <th>SKU</th>
            <th>Product Name</th>
            <th>Attribute</th>
"""
    for s in source_names:
        html += f'            <th>Value {s.upper()}</th>\n'
    html += """            <th>Difference</th>
          </tr></thead>
          <tbody>
"""

    if not attr_diff.empty:
        for _, row in attr_diff.iterrows():
            sku = str(row.get("sku", ""))
            name = str(row.get("product_name", ""))[:60]
            attr = str(row.get("attribute", ""))
            diff = row.get("diff", 0)

            html += f'          <tr>\n            <td><strong>{_esc(sku)}</strong></td>\n            <td>{_esc(name)}</td>\n'
            html += f'            <td>{_esc(attr)}</td>\n'
            for s in source_names:
                v = row.get(f"value_{s}", "")
                if v and str(v) != "" and str(v) != "nan":
                    html += f'            <td class="price">{_esc(str(v))}</td>\n'
                else:
                    html += '            <td class="text-muted">&mdash;</td>\n'
            html += f'            <td class="price price-diff">{diff}</td>\n'
            html += '          </tr>\n'

    html += """          </tbody>
        </table>
      </div>
    </div>
  </div>
"""

    # === TAB: Name vs Attributes ===
    html += """
  <div class="tab-content" id="tab-namecheck">
    <div class="table-wrap">
      <div class="search-bar">
        <input type="text" placeholder="Search by SKU, category, attribute..." onkeyup="filterTable(this, 'namecheck-table')" style="margin-right:12px;">
        <select onchange="filterByStatus(this, 'namecheck-table')" style="padding:8px 12px;border-radius:8px;border:1px solid var(--border);background:var(--bg);color:var(--text);font-size:0.9rem;">
          <option value="">All statuses</option>
          <option value="MISMATCH">MISMATCH only</option>
          <option value="NOT FOUND">NOT FOUND only</option>
        </select>
      </div>
      <div class="scrollable">
        <table id="namecheck-table">
          <thead><tr>
            <th>Category</th>
            <th>SKU</th>
            <th>Product Title (Zoho)</th>
            <th>Attribute</th>
            <th>From Name</th>
            <th>Website Value</th>
            <th>Google Value</th>
            <th>Status</th>
          </tr></thead>
          <tbody>
"""

    if not name_vs_attrs.empty:
        for _, row in name_vs_attrs.iterrows():
            status = str(row.get("status", ""))
            status_cls = "price-diff" if status == "MISMATCH" else "text-warn"
            status_tag = f'<span class="tag tag-no">{status}</span>' if status == "MISMATCH" else f'<span class="tag" style="background:rgba(245,158,11,0.15);color:#fbbf24;">{status}</span>'

            html += '          <tr>\n'
            html += f'            <td>{_esc(str(row.get("category", "")))}</td>\n'
            html += f'            <td><strong>{_esc(str(row.get("sku", "")))}</strong></td>\n'
            html += f'            <td>{_esc(str(row.get("product_title", "")))}</td>\n'
            html += f'            <td><strong>{_esc(str(row.get("attribute", "")))}</strong></td>\n'
            html += f'            <td style="color:var(--accent);">{_esc(str(row.get("from_name", "")))}</td>\n'

            wv = str(row.get("website_value", "—"))
            gv = str(row.get("google_value", "—"))
            wm = str(row.get("web_match", "—"))
            gm = str(row.get("google_match", "—"))

            w_style = 'color:var(--red);' if wm == "No" else ('color:var(--green);' if wm == "Yes" else 'color:var(--muted);')
            g_style = 'color:var(--red);' if gm == "No" else ('color:var(--green);' if gm == "Yes" else 'color:var(--muted);')

            html += f'            <td style="{w_style}">{_esc(wv)}</td>\n'
            html += f'            <td style="{g_style}">{_esc(gv)}</td>\n'
            html += f'            <td>{status_tag}</td>\n'
            html += '          </tr>\n'

    html += """          </tbody>
        </table>
      </div>
    </div>
  </div>
"""

    # === TAB: All Products ===
    html += """
  <div class="tab-content" id="tab-all">
    <div class="table-wrap">
      <div class="search-bar"><input type="text" placeholder="Search by SKU or product name..." onkeyup="filterTable(this, 'all-table')"></div>
      <div class="scrollable">
        <table id="all-table">
          <thead><tr>
            <th>SKU</th>
            <th>Product Name</th>
"""
    for s in source_names:
        html += f'            <th>Price {s.upper()}</th>\n'
    for s in source_names:
        html += f'            <th>In {s.upper()}</th>\n'
    html += """            <th>Zoho Status</th>
            <th>Details</th>
          </tr></thead>
          <tbody>
"""

    for idx, row in merged.iterrows():
        sku = str(row.get("sku", ""))
        name = ""
        for s in source_names:
            if row.get(f"name_{s}"):
                name = str(row[f"name_{s}"])[:80]
                break
        zoho_status = str(row.get("status_zoho", "")) if pd.notna(row.get("status_zoho")) else ""

        html += f'          <tr>\n            <td><strong>{_esc(sku)}</strong></td>\n            <td>{_esc(name)}</td>\n'
        for s in source_names:
            p = row.get(f"price_{s}")
            html += f'            <td class="price">{_fmt_price(p)}</td>\n'
        for s in source_names:
            in_s = row.get(f"in_{s}", False)
            if in_s:
                html += '            <td><span class="tag tag-yes">&#10003;</span></td>\n'
            else:
                html += '            <td><span class="tag tag-no">&#10007;</span></td>\n'
        html += f'            <td>{_fmt_status(zoho_status)}</td>\n'
        html += f'            <td><span class="expand-btn" onclick="toggleDetail(\'al-{idx}\')">&#9660;</span></td>\n'
        html += '          </tr>\n'
        html += _build_detail_row(f"al-{idx}", sku, source_names, zoho_attrs, web_attrs, google_attrs)

    html += """          </tbody>
        </table>
      </div>
    </div>
  </div>

</div>

<script>
function showTab(name) {
  document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
  document.querySelectorAll('.tab').forEach(el => el.classList.remove('active'));
  document.getElementById('tab-' + name).classList.add('active');
  event.currentTarget.classList.add('active');
}

function toggleDetail(id) {
  const row = document.getElementById(id);
  if (!row) return;
  row.classList.toggle('open');
  row.style.display = row.classList.contains('open') ? 'table-row' : '';
}

function filterTable(input, tableId) {
  const filter = input.value.toLowerCase();
  const table = document.getElementById(tableId);
  const rows = table.querySelectorAll('tbody tr:not(.detail-row)');
  rows.forEach(row => {
    const text = row.textContent.toLowerCase();
    const show = text.includes(filter);
    row.style.display = show ? '' : 'none';
    const next = row.nextElementSibling;
    if (next && next.classList.contains('detail-row')) {
      next.classList.remove('open');
      next.style.display = '';
    }
  });
}

function filterByStatus(select, tableId) {
  const filter = select.value.toUpperCase();
  const table = document.getElementById(tableId);
  const rows = table.querySelectorAll('tbody tr');
  rows.forEach(row => {
    if (!filter) { row.style.display = ''; return; }
    const text = row.textContent.toUpperCase();
    row.style.display = text.includes(filter) ? '' : 'none';
  });
}
</script>
</body>
</html>"""

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"  HTML report: {output_path}")
    return output_path


def _esc(text):
    """Escape HTML."""
    return (str(text)
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;"))


def _fmt_price(val):
    """Format price."""
    if pd.isna(val) or val is None:
        return '<span class="text-muted">—</span>'
    return f"${float(val):,.2f}"


def _fmt_status(status):
    """Format Zoho active/inactive status."""
    if not status or str(status).strip().lower() in ("", "nan", "none"):
        return '<span class="text-muted">—</span>'
    s = str(status).strip()
    if s.lower() == "active":
        return '<span class="tag tag-yes">Active</span>'
    elif s.lower() == "inactive":
        return '<span class="tag tag-no">Inactive</span>'
    return _esc(s)


def _get_zoho_status(sku, merged):
    """Get Zoho status from merged data."""
    match = merged[merged["sku"] == sku]
    if not match.empty:
        val = match.iloc[0].get("status_zoho", "")
        if pd.notna(val):
            return str(val)
    return ""


def _build_detail_row(row_id, sku, source_names, zoho_attrs, web_attrs, google_attrs):
    """Build expandable detail row with attributes from all sources."""
    sku_upper = str(sku).strip().upper()

    attr_sources = {
        "zoho": zoho_attrs.get(sku_upper, {}),
        "website": web_attrs.get(sku_upper, {}),
        "google": google_attrs.get(sku_upper, {}),
    }

    html = f'          <tr class="detail-row" id="{row_id}"><td colspan="20" class="detail-cell">\n'
    html += '            <div class="detail-grid">\n'

    for s in source_names:
        attrs = attr_sources.get(s, {})
        tag_class = f"tag-{s}"
        html += f'              <div class="detail-source">\n'
        html += f'                <h4><span class="tag {tag_class}">{s.upper()}</span> Attributes</h4>\n'

        if attrs:
            html += '                <table class="attr-table">\n'
            for k, v in attrs.items():
                if v and str(v).strip() and str(v).strip().lower() not in ("nan", "none"):
                    html += f'                  <tr><td class="attr-label">{_esc(k)}</td><td>{_esc(str(v))}</td></tr>\n'
            html += '                </table>\n'
        else:
            html += '                <p class="text-muted">No data available</p>\n'

        html += '              </div>\n'

    html += '            </div>\n'
    html += '          </td></tr>\n'
    return html
