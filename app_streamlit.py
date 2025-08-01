import sys, os
sys.path.append(os.path.join(os.path.dirname(__file__), "src"))
import streamlit as st
import pandas as pd
from io import BytesIO
from sku_dupe_finder.core import analyze, write_report

st.set_page_config(page_title="SKU Cross-Workbook Duplicates", layout="centered")

st.title("SKU Cross-Workbook Duplicates")
st.write("Upload multiple **.xlsx** files and find SKUs that appear in more than one workbook.")

uploads = st.file_uploader("Upload Excel files", type=["xlsx"], accept_multiple_files=True)
explicit_cols = st.text_input("Explicit SKU column names (comma-separated, optional)", value="")
patterns_text = st.text_input("Custom regex patterns for SKU columns (comma-separated, optional)", value="")
include_within = st.checkbox("Include duplicates within the same workbook (default OFF)", value=False)

run = st.button("Run analysis")

if run:
    if not uploads:
        st.warning("Please upload at least two Excel files.")
    else:
        files = []
        for up in uploads:
            # Save to temp in memory
            files.append(up)

        # Read into pandas via in-memory buffers; we will adapt analyze() by saving temp files
        tmp_paths = []
        import tempfile, os
        for up in uploads:
            suffix = ".xlsx"
            fd, tmp = tempfile.mkstemp(suffix=suffix)
            os.close(fd)
            with open(tmp, "wb") as f:
                f.write(up.getbuffer())
            tmp_paths.append(tmp)

        # Run core
        details_df, presence_counts, presence_bool, read_errors, sku_col_map = analyze(
            tmp_paths,
            sku_cols=[c.strip() for c in explicit_cols.split(",") if c.strip()] or None,
            patterns=[p.strip() for p in patterns_text.split(",") if p.strip()] or None,
            include_within_workbook_dupes=include_within,
        )

        # Prepare report to download
        output = BytesIO()
        write_report(
            out_path=tmp_paths[0] + "_report.xlsx",  # write to disk then read back
            details_df=details_df,
            presence_counts=presence_counts,
            presence_bool=presence_bool,
            read_errors=read_errors,
            sku_col_map=sku_col_map,
            only_across_workbooks=not include_within,
        )
        # Read back
        with open(tmp_paths[0] + "_report.xlsx", "rb") as f:
            data = f.read()
        st.download_button("Download Excel Report", data=data, file_name="sku_crossworkbook_duplicates.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        if not details_df.empty:
            st.subheader("Preview (first 200 rows of details)")
            st.dataframe(details_df.head(200))
        if read_errors:
            st.subheader("Read issues")
            st.json(read_errors)

