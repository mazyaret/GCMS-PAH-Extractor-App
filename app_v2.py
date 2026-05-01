import streamlit as st
import tempfile
import subprocess
import sys
from pathlib import Path

import pandas as pd
from openpyxl import Workbook

from GCMS_PAH_Extractor_NoTableFix_v2 import extract_gcms, load_df_overrides, norm_name

st.set_page_config(page_title="GC/MS PAH Extractor", page_icon="🧪", layout="centered")

st.title("GC/MS PAH Extraction Platform")
st.write("Upload the original GC/MS Excel file, review/edit dilution factors, and download the processed workbook.")

st.markdown("### 1) Upload files")
gcms_file = st.file_uploader("Upload original GC/MS Excel file (.xlsx)", type=["xlsx"])
df_file = st.file_uploader("Optional: Upload lab submission form or DF template (.xlsx)", type=["xlsx"])

st.caption("The optional file may be the lab submission form with 'Sample Name (Unique)' and 'Dilution factor' columns, or a simple DF template with 'Sample Name' and 'Dilution Factor'.")


def build_df_table(gcms_upload, df_upload):
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        gcms_path = tmp / gcms_upload.name
        gcms_path.write_bytes(gcms_upload.getbuffer())

        overrides = {}
        if df_upload is not None:
            df_path = tmp / df_upload.name
            df_path.write_bytes(df_upload.getbuffer())
            overrides = load_df_overrides(df_path)

        samples, meta = extract_gcms(gcms_path, df_overrides=None)
        rows = []
        for s in samples:
            sample_name = s["sample_name"]
            df = s["dilution_factor"]
            source = s["df_source"]
            if str(sample_name).strip() in overrides:
                df = overrides[str(sample_name).strip()]
                source = "Uploaded submission/DF file"
            elif norm_name(sample_name) in overrides:
                df = overrides[norm_name(sample_name)]
                source = "Uploaded submission/DF file"
            rows.append({"Sample Name": sample_name, "Dilution Factor": float(df), "DF Source": source})
        return pd.DataFrame(rows), meta, len(overrides) // 2 if overrides else 0


def write_manual_df_workbook(df_table, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Dilution_Factors"
    ws.append(["Sample Name", "Dilution Factor", "DF Source"])
    for _, row in df_table.iterrows():
        ws.append([str(row["Sample Name"]), float(row["Dilution Factor"]), str(row.get("DF Source", "Manual in web app"))])
    wb.save(path)

if gcms_file is not None:
    try:
        df_table, meta, override_count = build_df_table(gcms_file, df_file)
        st.success(f"Detected {len(df_table)} sample(s). Uploaded DF matches found: {override_count}.")
        st.markdown("### 2) Review or edit dilution factors")
        st.write("Edit the Dilution Factor column if needed. These values will override any DF parsed from the sample name.")

        edited_df = st.data_editor(
            df_table,
            use_container_width=True,
            hide_index=True,
            disabled=["Sample Name", "DF Source"],
            column_config={
                "Dilution Factor": st.column_config.NumberColumn("Dilution Factor", min_value=0.0, step=1.0, format="%.4f")
            },
            key="df_editor"
        )

        st.markdown("### 3) Process and download")
        if st.button("Process GC/MS file", type="primary"):
            with st.spinner("Processing file..."):
                with tempfile.TemporaryDirectory() as tmp:
                    tmp = Path(tmp)
                    input_path = tmp / gcms_file.name
                    input_path.write_bytes(gcms_file.getbuffer())
                    manual_df_path = tmp / "manual_df_from_web_app.xlsx"
                    write_manual_df_workbook(edited_df, manual_df_path)
                    output_name = Path(gcms_file.name).stem + "_processed.xlsx"
                    output_path = tmp / output_name
                    script_path = Path(__file__).parent / "GCMS_PAH_Extractor_NoTableFix_v2.py"
                    cmd = [sys.executable, str(script_path), str(input_path), "--df-template", str(manual_df_path), "--output", str(output_path)]
                    result = subprocess.run(cmd, capture_output=True, text=True)
                    if result.returncode != 0:
                        st.error("Processing failed.")
                        st.code(result.stderr or result.stdout)
                    elif not output_path.exists():
                        st.error("Processing finished, but no output file was created.")
                        st.code(result.stdout + "\n" + result.stderr)
                    else:
                        st.success("Done! Download the processed workbook below.")
                        st.download_button(
                            "Download processed Excel file",
                            data=output_path.read_bytes(),
                            file_name=output_name,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        with st.expander("Processing log"):
                            st.code(result.stdout or "No additional log.")
    except Exception as e:
        st.error("The app could not prepare the DF table. Please check the uploaded files.")
        st.code(str(e))
else:
    st.info("Upload the original GC/MS Excel file to begin.")

st.divider()
st.caption("GC/MS PAH Extractor v2 — upload, review DF, process, download.")
