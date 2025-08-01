# SKU Dupe Finder (Python)
Find SKUs that appear **in more than one Excel workbook** (cross-workbook duplicates).

## Features
- Scans one or more Excel workbooks (all sheets).
- Auto-detects SKU-like columns (configurable) or use explicit column names.
- Normalizes SKUs (trims, uppercases, handles `123.0` → `123`).
- Exports an Excel report with:
  - Presence matrix by workbook (+ count)
  - Detailed occurrences (file, sheet, row, column)
  - Detected SKU columns per sheet
  - Read issues (if any)
- **CLI** tool for automation and **Streamlit UI** for simple use.

## Quick start (CLI)
```bash
# 1) Create & activate a virtual env (recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
# source .venv/bin/activate

# 2) Install
pip install -U pip
pip install -e .  # editable install for local dev

# 3) Run (examples)
python -m sku_dupe_finder \  --inputs "C:\path\to\Inventory_Adjustment chemtok.xlsx" \           "C:\path\to\Inventory_Adjustment non std fevisa.xlsx" \           "C:\path\to\Inventory_Adjustment sepco (non std & std).xlsx" \           "C:\path\to\Inventory_Adjustment std fevisa.xlsx" \  --out "C:\path\to\sku_crossworkbook_duplicates.xlsx"

# Or scan a folder (recursively)
python -m sku_dupe_finder --inputs "C:\path\to\folder" --recursive --out report.xlsx

# Use explicit SKU column names (exact match, case-insensitive)
python -m sku_dupe_finder --inputs "C:\files" --sku-columns "SKU" "Item Code"
```

## Streamlit app (optional GUI)
```bash
pip install -e .[app]
streamlit run app_streamlit.py
```
Upload the Excel files and click **Run analysis** to generate and download the report.

## Notes
- Supported formats: `.xlsx`. (Old `.xls` is not processed by default.)
- If a workbook has no explicit SKU column names, the tool tries to detect likely SKU columns.
- By default, duplicates **across different workbooks** are reported. You can also include within-workbook dupes.

## License
MIT
