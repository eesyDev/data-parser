"""
Report generation: console output + Excel with tabs.
"""

import os
import pandas as pd
import config


def print_summary(results):
    """Print summary to console."""
    sources = results["sources"]
    merged = results["merged"]
    missing = results["missing"]
    price_diff = results["price_differences"]
    name_diff = results["name_differences"]
    attr_diff = results.get("attribute_differences", pd.DataFrame())

    print("\n" + "=" * 60)
    print("COMPARISON SUMMARY")
    print("=" * 60)

    print(f"\nSources:")
    for name, df in sources.items():
        print(f"  - {name}: {len(df)} products")

    print(f"\nTotal unique SKUs: {len(merged)}")

    # Count products present in all sources
    source_names = list(sources.keys())
    if source_names:
        in_all = merged
        for s in source_names:
            in_all = in_all[in_all[f"in_{s}"] == True]
        print(f"Products in all sources: {len(in_all)}")

    print(f"\nDiscrepancies:")
    print(f"  - Missing products: {len(missing)}")
    print(f"  - Price differences: {len(price_diff)}")
    print(f"  - Name differences: {len(name_diff)}")
    if not attr_diff.empty:
        unique_skus = attr_diff["sku"].nunique()
        print(f"  - Attribute differences: {len(attr_diff)} ({unique_skus} products)")

    if not price_diff.empty:
        print(f"\nTop 5 largest price discrepancies:")
        top = price_diff.nlargest(5, "price_diff")
        for _, row in top.iterrows():
            prices = " | ".join(
                f"{s}: ${row.get(f'price_{s}', 'N/A')}"
                for s in source_names
                if pd.notna(row.get(f"price_{s}"))
            )
            print(f"  SKU: {row['sku']} — {row['product_name'][:50]}")
            print(f"    {prices}  (difference: ${row['price_diff']})")

    print("=" * 60)


def export_excel(results, output_path=None):
    """Export detailed report to Excel."""
    output_path = output_path or config.REPORT_FILE
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    sources = results["sources"]
    merged = results["merged"]
    missing = results["missing"]
    price_diff = results["price_differences"]
    name_diff = results["name_differences"]
    attr_diff = results.get("attribute_differences", pd.DataFrame())

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        # "Summary" tab
        source_names = list(sources.keys())
        summary_data = {
            "Metric": [
                "Total unique SKUs",
                "Missing products",
                "Price discrepancies",
                "Name discrepancies",
                "Attribute discrepancies",
            ],
            "Value": [
                len(merged),
                len(missing),
                len(price_diff),
                len(name_diff),
                len(attr_diff),
            ],
        }
        for name, df in sources.items():
            summary_data["Metric"].append(f"Products in {name}")
            summary_data["Value"].append(len(df))

        pd.DataFrame(summary_data).to_excel(writer, sheet_name="Summary", index=False)

        # "All Products" tab
        if not merged.empty:
            merged.to_excel(writer, sheet_name="All Products", index=False)

        # "Price Discrepancies" tab
        if not price_diff.empty:
            price_diff.to_excel(writer, sheet_name="Price Discrepancies", index=False)

        # "Missing Products" tab
        if not missing.empty:
            missing.to_excel(writer, sheet_name="Missing Products", index=False)

        # "Name Discrepancies" tab
        if not name_diff.empty:
            name_diff.to_excel(writer, sheet_name="Name Discrepancies", index=False)

        # "Attribute Discrepancies" tab
        if not attr_diff.empty:
            attr_diff.to_excel(writer, sheet_name="Attr Discrepancies", index=False)

        # "Name vs Attributes" tab
        name_vs_attrs = results.get("name_vs_attributes", pd.DataFrame())
        if not name_vs_attrs.empty:
            name_vs_attrs.to_excel(writer, sheet_name="Name vs Attributes", index=False)

        # Tabs with raw data from each source
        for name, df in sources.items():
            sheet_name = f"Source_{name}"[:31]  # Excel sheet name length limit
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"\nReport saved: {output_path}")


def generate_report(results):
    """Generate full report (console + Excel)."""
    print("\n3. Generating report...")
    print_summary(results)
    export_excel(results)
