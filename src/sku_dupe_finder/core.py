from __future__ import annotations
import os
import re
from typing import Dict, List, Tuple, Iterable
import pandas as pd

DEFAULT_SKU_COL_PATTERNS = [
    r"\bsku\b",
    r"\bitem\s*code\b",
    r"\bitem\b(?!\s*name)",
    r"\bpart\s*(?:no|number)?\b",
    r"\bmaterial\s*code\b",
    r"\bproduct\s*code\b",
    r"\bstock\s*code\b",
    r"\bstyle\s*code\b",
]

def normalize_sku(val):
    """Normalize SKU values to comparable strings."""
    if pd.isna(val):
        return None
    try:
        if isinstance(val, (int, float)) or (isinstance(val, str) and val.strip().replace('.', '', 1).isdigit()):
            f = float(val)
            val = str(int(f)) if f.is_integer() else str(f)
    except Exception:
        pass
    s = str(val).strip()
    s = re.sub(r"\s+", " ", s)
    s = s.strip().strip('"').strip("'").upper()
    if s in {"", "N/A", "NA", "NONE", "NULL", "-"}:
        return None
    return s

def _compile_patterns(patterns: Iterable[str] | None):
    if not patterns:
        patterns = DEFAULT_SKU_COL_PATTERNS
    return [re.compile(p, flags=re.IGNORECASE) for p in patterns]

def find_sku_columns(df: pd.DataFrame, explicit_cols: List[str] | None = None, patterns: Iterable[str] | None = None) -> List[str]:
    """Return a list of columns that likely contain SKUs (or explicit ones if provided)."""
    cols = [str(c).strip() for c in df.columns]
    if explicit_cols:
        wanted = set(c.strip().lower() for c in explicit_cols)
        return [c for c in cols if c.lower() in wanted]

    regexes = _compile_patterns(patterns)
    candidates = []
    for col in cols:
        for rx in regexes:
            if rx.search(col):
                candidates.append(col)
                break
    if not candidates:
        # Fallback: a first column heuristic if looks alphanumeric
        first_col = cols[0]
        sample_series = df[first_col].dropna().astype(str).head(50)
        alnum_ratio = (sample_series.str.contains(r"[A-Za-z0-9]").mean()) if len(sample_series) else 0
        if alnum_ratio > 0.5:
            candidates = [first_col]
    # de-duplicate preserving order
    seen = set()
    out = []
    for c in candidates:
        lc = c.lower()
        if lc not in seen:
            out.append(c); seen.add(lc)
    return out

def find_excel_files(inputs: List[str], recursive: bool = False) -> List[str]:
    files = []
    for p in inputs:
        if os.path.isdir(p):
            for root, dirs, fs in os.walk(p):
                for name in fs:
                    if name.lower().endswith(".xlsx"):
                        files.append(os.path.join(root, name))
                if not recursive:
                    break
        else:
            if os.path.isfile(p) and p.lower().endswith(".xlsx"):
                files.append(p)
    # unique preserve order
    seen = set()
    ordered = []
    for f in files:
        if f not in seen:
            ordered.append(f); seen.add(f)
    return ordered

def analyze(files: List[str], sku_cols: List[str] | None = None, patterns: Iterable[str] | None = None,
            include_within_workbook_dupes: bool = False):
    """Core analysis. Returns (details_df, presence_counts, presence_bool_with_count, read_errors, sku_col_map)."""
    details_rows = []
    sku_col_map: Dict[Tuple[str, str], List[str]] = {}
    read_errors: Dict[str, str] = {}

    for fp in files:
        try:
            xls = pd.read_excel(fp, sheet_name=None, dtype=str, engine="openpyxl")
        except Exception as e:
            read_errors[fp] = f"Failed to read: {e}"
            continue
        basename = os.path.basename(fp)
        for sheet_name, df in (xls or {}).items():
            if df is None or df.empty:
                continue
            df.columns = [str(c).strip() for c in df.columns]
            cols = find_sku_columns(df, explicit_cols=sku_cols, patterns=patterns)
            if not cols:
                continue
            sku_col_map[(basename, sheet_name)] = cols
            for col in cols:
                series = df[col].map(normalize_sku)
                for idx, sku in series.items():
                    if sku:
                        details_rows.append({
                            "SKU": sku,
                            "File": basename,
                            "Sheet": sheet_name,
                            "Column": col,
                            "RowNumber": int(idx) + 2 if isinstance(idx, (int, float)) else None
                        })

    details_df = pd.DataFrame(details_rows)
    if details_df.empty:
        presence_counts = pd.DataFrame()
        presence_bool = pd.DataFrame()
        return details_df, presence_counts, presence_bool, read_errors, sku_col_map

    # Cross-workbook presence matrix
    presence_counts = pd.crosstab(details_df["SKU"], details_df["File"])

    presence_bool = presence_counts > 0
    presence_bool["WorkbooksCount"] = presence_bool.sum(axis=1)

    if not include_within_workbook_dupes:
        # We keep the matrix as-is; filtering to dupes happens when writing the report
        pass

    return details_df, presence_counts, presence_bool, read_errors, sku_col_map

def write_report(out_path: str,
                 details_df: pd.DataFrame,
                 presence_counts: pd.DataFrame,
                 presence_bool: pd.DataFrame,
                 read_errors: Dict[str, str],
                 sku_col_map: Dict[Tuple[str, str], List[str]],
                 only_across_workbooks: bool = True):
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)

    if details_df.empty:
        with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
            pd.DataFrame([{"Message": "No SKU-like data found in the provided files."}]).to_excel(writer, sheet_name="Summary", index=False)
            if read_errors:
                pd.DataFrame([{"File": os.path.basename(k), "Issue": v} for k, v in read_errors.items()]).to_excel(writer, sheet_name="Read_Issues", index=False)
        return

    # Determine dup SKUs
    if "WorkbooksCount" in presence_bool.columns:
        if only_across_workbooks:
            dup_index = presence_bool.index[presence_bool["WorkbooksCount"] > 1]
        else:
            # Any duplicated SKU anywhere (across or within) – treat WorkbooksCount>=1 and check within-file dupes separately if needed
            dup_index = presence_bool.index
    else:
        dup_index = presence_bool.index

    details_dups = details_df[details_df["SKU"].isin(dup_index)].sort_values(["SKU", "File", "Sheet", "RowNumber"])

    with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
        if not presence_counts.empty:
            presence_counts.loc[sorted(set(details_dups["SKU"]))].to_excel(writer, sheet_name="Counts_by_File")
            presence_bool.loc[sorted(set(details_dups["SKU"]))].to_excel(writer, sheet_name="Presence_by_File")
        details_dups.to_excel(writer, sheet_name="Details", index=False)

        sku_map_records = [{
            "File": k[0],
            "Sheet": k[1],
            "Detected_SKU_Columns": ", ".join(v)
        } for k, v in sku_col_map.items()]
        pd.DataFrame(sku_map_records).to_excel(writer, sheet_name="Detected_Columns", index=False)

        if read_errors:
            pd.DataFrame([{"File": os.path.basename(k), "Issue": v} for k, v in read_errors.items()]).to_excel(writer, sheet_name="Read_Issues", index=False)
