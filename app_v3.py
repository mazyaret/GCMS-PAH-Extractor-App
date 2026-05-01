import math
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import Workbook

from GCMS_PAH_Extractor_NoTableFix_v3 import extract_gcms, load_sample_metadata, write_processed_workbook

st.set_page_config(page_title="GC/MS PAH Extractor v3", page_icon="🧪", layout="wide")

st.title("GC/MS PAH Extraction Platform v3")
st.write("Upload the original GC/MS Excel file, upload the submission form, review/edit dilution and normalization inputs, and download the processed workbook.")

st.markdown("### 1) Upload files")
gcms_file = st.file_uploader("Upload original GC/MS Excel file (.xlsx)", type=["xlsx"])
metadata_file = st.file_uploader("Optional: Upload lab submission form / metadata template (.xlsx)", type=["xlsx"])

st.caption("The optional file may include columns such as 'Sample Name (Unique)', 'Dilution factor', 'Sample Mass or Volume (ml)', 'Fabric size (cm^2)', 'fabric mass (g)', and 'fuel mass loss (g)'.")


def value_or_blank(value):
    if value is None:
        return None
    if isinstance(value, float) and math.isnan(value):
        return None
    return value


def build_metadata_table(gcms_upload, metadata_upload):
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        gcms_path = tmp / gcms_upload.name
        gcms_path.write_bytes(gcms_upload.getbuffer())
        metadata = {}
        if metadata_upload is not None:
            metadata_path = tmp / metadata_upload.name
            metadata_path.write_bytes(metadata_upload.getbuffer())
            metadata = load_sample_metadata(metadata_path)
        samples, meta = extract_gcms(gcms_path, sample_metadata=metadata)
        rows = []
        match_count = 0
        for s in samples:
            source = s.get("parameter_source", "Default/blank")
            if source == "Uploaded submission/metadata file":
                match_count += 1
            rows.append({
                "Sample Name": s["sample_name"],
                "Data File": s.get("data_file"),
                "Dilution Factor": float(s.get("dilution_factor") or 1.0),
                "Extract Volume (mL)": float(s.get("extract_volume_ml") or 1.0),
                "Fabric Size (cm^2)": value_or_blank(s.get("fabric_size_cm2")),
                "Fabric Mass (g)": value_or_blank(s.get("fabric_mass_g")),
                "Fuel Mass Loss (g)": value_or_blank(s.get("fuel_mass_loss_g")),
                "DF Source": s.get("df_source"),
                "Parameter Source": source,
            })
        return pd.DataFrame(rows), meta, match_count


def write_manual_metadata_workbook(df_table, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sample_Metadata"
    ws.append(["Sample Name", "Dilution Factor", "Extract Volume (mL)", "Fabric size (cm^2)", "fabric mass (g)", "fuel mass loss (g)", "DF Source", "Parameter Source"])
    for _, row in df_table.iterrows():
        ws.append([
            str(row["Sample Name"]),
            value_or_blank(row.get("Dilution Factor")),
            value_or_blank(row.get("Extract Volume (mL)")),
            value_or_blank(row.get("Fabric Size (cm^2)")),
            value_or_blank(row.get("Fabric Mass (g)")),
            value_or_blank(row.get("Fuel Mass Loss (g)")),
            str(row.get("DF Source", "Manual in web app")),
            str(row.get("Parameter Source", "Manual in web app")),
        ])
    wb.save(path)


if gcms_file is not None:
    try:
        metadata_table, meta, match_count = build_metadata_table(gcms_file, metadata_file)
        st.success(f"Detected {len(metadata_table)} sample(s). Uploaded metadata matches found: {match_count}.")
        st.markdown("### 2) Review or edit dilution and normalization inputs")
        st.write("Edit any values if needed. These values will be used for corrected concentrations, surface load, ug/g fabric, ug/g fuel mass loss, totals, and TEQ calculations.")
        edited_df = st.data_editor(
            metadata_table,
            use_container_width=True,
            hide_index=True,
            disabled=["Sample Name", "Data File", "DF Source", "Parameter Source"],
            column_config={
                "Dilution Factor": st.column_config.NumberColumn("Dilution Factor", min_value=0.0, step=1.0, format="%.4f"),
                "Extract Volume (mL)": st.column_config.NumberColumn("Extract Volume (mL)", min_value=0.0, step=0.1, format="%.4f"),
                "Fabric Size (cm^2)": st.column_config.NumberColumn("Fabric Size (cm^2)", min_value=0.0, step=1.0, format="%.4f"),
                "Fabric Mass (g)": st.column_config.NumberColumn("Fabric Mass (g)", min_value=0.0, step=0.1, format="%.4f"),
                "Fuel Mass Loss (g)": st.column_config.NumberColumn("Fuel Mass Loss (g)", min_value=0.0, step=1.0, format="%.4f"),
            },
            key="metadata_editor",
        )
        st.markdown("### 3) Process and download")
        if st.button("Process GC/MS file", type="primary"):
            with st.spinner("Processing file..."):
                with tempfile.TemporaryDirectory() as tmp:
                    tmp = Path(tmp)
                    input_path = tmp / gcms_file.name
                    input_path.write_bytes(gcms_file.getbuffer())
                    manual_metadata_path = tmp / "manual_metadata_from_web_app.xlsx"
                    write_manual_metadata_workbook(edited_df, manual_metadata_path)
                    sample_metadata = load_sample_metadata(manual_metadata_path)
                    samples, run_meta = extract_gcms(input_path, sample_metadata=sample_metadata)
                    output_name = Path(gcms_file.name).stem + "_processed_v3.xlsx"
                    output_path = tmp / output_name
                    write_processed_workbook(samples, run_meta, output_path)
                    st.success("Done! Download the processed workbook below.")
                    st.download_button(
                        "Download processed Excel file",
                        data=output_path.read_bytes(),
                        file_name=output_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
    except Exception as e:
        st.error("The app could not prepare the metadata table. Please check the uploaded files.")
        st.code(str(e))
else:
    st.info("Upload the original GC/MS Excel file to begin.")

st.divider()
st.caption("GC/MS PAH Extractor v3 — DF, normalization parameters, PAH totals, LMW/HMW totals, and TEQ metrics.")
