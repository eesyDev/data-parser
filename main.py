"""
Entry point: Load -> Normalize -> Compare -> Report
Run: python main.py
"""

import loaders
import comparator
import report
import report_html
import name_parser
import grid_view


def main():
    print("=" * 60)
    print("PRODUCT DATA COMPARISON PARSER")
    print("=" * 60)

    # 1. Load data
    print("\n1. Loading data...")
    sources = loaders.load_all()

    if len(sources) < 2:
        print("\nAt least 2 sources required for comparison.")
        print("Configure paths in config.py")
        return

    # 2. Compare
    results = comparator.compare_all(sources)

    # 3. Name vs Attributes comparison
    print("\n3. Parsing product names vs actual attributes...")
    name_vs_attrs = name_parser.compare_name_vs_attributes()
    results["name_vs_attributes"] = name_vs_attrs
    if not name_vs_attrs.empty:
        mismatches = name_vs_attrs[name_vs_attrs["status"] == "MISMATCH"]
        not_found = name_vs_attrs[name_vs_attrs["status"] == "NOT FOUND"]
        print(f"   Mismatches (name says X, attribute says Y): {len(mismatches)}")
        print(f"   Not found (name has info, but no attribute set): {len(not_found)}")
        print(f"   Total issues: {len(name_vs_attrs)} across {name_vs_attrs['sku'].nunique()} products")

    # 4. Reports
    report.generate_report(results)
    print("\n5. HTML report with attributes...")
    html_path = report_html.generate_html(results)

    # 6. Attribute grid
    print("\n6. Attribute grid by category...")
    grids = grid_view.build_grid()
    grid_path = grid_view.generate_grid_html(grids)

    print("\nDone!")


if __name__ == "__main__":
    main()
