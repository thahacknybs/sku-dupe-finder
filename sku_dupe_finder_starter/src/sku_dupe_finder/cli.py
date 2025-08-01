from __future__ import annotations
import argparse, sys, os
from .core import find_excel_files, analyze, write_report

def build_parser():
    p = argparse.ArgumentParser(description="Find SKUs that appear in more than one Excel workbook.")
    p.add_argument("--inputs", nargs="+", required=True,
                   help="One or more paths to .xlsx files or directories to scan.")
    p.add_argument("--recursive", action="store_true",
                   help="Recurse into subfolders when input is a directory.")
    p.add_argument("--sku-columns", nargs="*", default=None,
                   help="Explicit column names to treat as SKU columns (case-insensitive, exact match).")
    p.add_argument("--sku-col-patterns", nargs="*", default=None,
                   help="Regex patterns to detect SKU columns (override defaults).")
    p.add_argument("--out", default="sku_crossworkbook_duplicates.xlsx",
                   help="Path to write the Excel report.")
    p.add_argument("--include-within-workbook-dupes", action="store_true",
                   help="Also include duplicates within the same workbook (by default we focus on cross-workbook only).")
    return p

def main(argv=None):
    argv = argv or sys.argv[1:]
    args = build_parser().parse_args(argv)

    files = find_excel_files(args.inputs, recursive=args.recursive)
    if not files:
        print("No .xlsx files found. Check the paths or use --recursive for folders.", file=sys.stderr)
        return 2

    details_df, presence_counts, presence_bool, read_errors, sku_col_map = analyze(
        files,
        sku_cols=args.sku_columns,
        patterns=args.sku_col_patterns,
        include_within_workbook_dupes=args.include_within_workbook_dupes,
    )

    write_report(
        out_path=args.out,
        details_df=details_df,
        presence_counts=presence_counts,
        presence_bool=presence_bool,
        read_errors=read_errors,
        sku_col_map=sku_col_map,
        only_across_workbooks=not args.include_within_workbook_dupes,
    )

    print(f"Wrote report to: {args.out}")
    if read_errors:
        print("Some files had issues:")
        for f, e in read_errors.items():
            print(f" - {f}: {e}")
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
