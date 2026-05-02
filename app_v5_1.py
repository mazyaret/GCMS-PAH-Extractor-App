import math
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import Workbook

from GCMS_PAH_Extractor_NoTableFix_v5_1 import (
    extract_gcms,
    load_sample_metadata,
    write_processed_workbook,
)
from GCMS_PAH_Extractor_UIowa_v3_1 import (
    extract_uiowa,
    load_submission_metadata,
    metadata_from_rows,
    write_processed_workbook_uiowa,
)

st.set_page_config(page_title="GC/MS PAH Extractor", page_icon="🧪", layout="wide")

st.title("GC/MS PAH Extraction Platform")
st.write(
    "Choose the GC/MS facility format, upload the required file(s), review/edit dilution and "
    "normalization inputs, and download one processed Excel workbook. Version 5 keeps the previous "
    "DF-corrected results and adds separate surrogate recovery-corrected columns."
)


def value_or_blank(value):
    if value is None:
        return None
    if isinstance(value, float) and math.isnan(value):
        return None
    return value


def download_excel_button(output_path: Path, output_name: str):
    st.success("Done! Download the processed workbook below.")
    st.download_button(
        "Download processed Excel file",
        data=output_path.read_bytes(),
        file_name=output_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def run_isu_v3_tab():
    st.subheader("ISU / previous GC/MS format")
    st.caption(
        "Use this option for the earlier format where all analytes are in one GC/MS Excel output "
        "with analyte blocks and Final Conc. columns."
    )

    gcms_file = st.file_uploader("Upload original GC/MS Excel file (.xlsx)", type=["xlsx"], key="isu_gcms")
    metadata_file = st.file_uploader(
        "Optional: Upload lab submission form / metadata template (.xlsx)",
        type=["xlsx"],
        key="isu_metadata",
    )
    st.caption(
        "The optional file may include columns such as Sample Name, Dilution factor, Extract Volume, "
        "Fabric size, Fabric mass, and Fuel mass loss."
    )

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
        ws.append([
            "Sample Name", "Dilution Factor", "Extract Volume (mL)", "Fabric size (cm^2)",
            "fabric mass (g)", "fuel mass loss (g)", "DF Source", "Parameter Source"
        ])
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

    if gcms_file is None:
        st.info("Upload the original GC/MS Excel file to begin.")
        return

    try:
        metadata_table, meta, match_count = build_metadata_table(gcms_file, metadata_file)
        st.success(f"Detected {len(metadata_table)} sample(s). Uploaded metadata matches found: {match_count}.")
        st.markdown("### Review or edit dilution and normalization inputs")
        edited_df = st.data_editor(
            metadata_table,
            use_container_width=True,
            hide_index=True,
            disabled=["Sample Name", "Data File", "DF Source", "Parameter Source"],
            column_config={
                "Dilution Factor": st.column_config.NumberColumn("Dilution Factor", min_value=0.0, step=1.0, format="%.4f"),
                "Extract Volume (mL)": st.column_config.NumberColumn("Extract Volume (mL)", min_value=0.0, step=0.1, format="%.4f"),
                "Fabric Size (cm^2)": st.column_config.NumberColumn("Fabric Size (cm^2)", min_value=0.0, step=1.0, format="%.4f"),
                "Fabric Mass (g)": st.column_config.NumberColumn("Fabric Mass (g)", min_value=0.0, step=0.01, format="%.6f"),
                "Fuel Mass Loss (g)": st.column_config.NumberColumn("Fuel Mass Loss (g)", min_value=0.0, step=0.01, format="%.6f"),
            },
            key="isu_metadata_editor",
        )
        if st.button("Process ISU-format GC/MS file", type="primary", key="isu_process"):
            with st.spinner("Processing ISU-format file..."):
                with tempfile.TemporaryDirectory() as tmp:
                    tmp = Path(tmp)
                    input_path = tmp / gcms_file.name
                    input_path.write_bytes(gcms_file.getbuffer())
                    manual_metadata_path = tmp / "manual_metadata_from_web_app.xlsx"
                    write_manual_metadata_workbook(edited_df, manual_metadata_path)
                    metadata = load_sample_metadata(manual_metadata_path)
                    samples, meta = extract_gcms(input_path, sample_metadata=metadata)
                    output_name = Path(gcms_file.name).stem + "_processed_v5_ISU.xlsx"
                    output_path = tmp / output_name
                    write_processed_workbook(samples, meta, output_path)
                    download_excel_button(output_path, output_name)
    except Exception as e:
        st.error("The app could not process the ISU-format file.")
        st.code(str(e))


def run_uiowa_tab():
    st.subheader("University of Iowa GC/MS format")
    st.caption(
        "Use this option for University of Iowa files: one submission form, one 16-PAH results workbook, "
        "and one surrogate results workbook."
    )

    submission_file = st.file_uploader("1) Upload University of Iowa sample submission file (.xlsx)", type=["xlsx"], key="uiowa_submission")
    target_file = st.file_uploader("2) Upload University of Iowa 16-PAH results file (.xlsx)", type=["xlsx"], key="uiowa_targets")
    surrogate_file = st.file_uploader("3) Upload University of Iowa surrogate results file (.xlsx)", type=["xlsx"], key="uiowa_surrogates")

    st.caption(
        "Expected UIowa result format: one worksheet per analyte. The app uses rows with Type = Sample "
        "and reads Calc. Conc. from the Calc. Conc. column, usually column G. Calibration and QC rows are ignored."
    )

    if not (submission_file and target_file and surrogate_file):
        st.info("Upload all three University of Iowa files to begin.")
        return

    try:
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            submission_path = tmp / submission_file.name
            target_path = tmp / target_file.name
            surrogate_path = tmp / surrogate_file.name
            submission_path.write_bytes(submission_file.getbuffer())
            target_path.write_bytes(target_file.getbuffer())
            surrogate_path.write_bytes(surrogate_file.getbuffer())

            metadata = load_submission_metadata(submission_path)
            samples, meta = extract_uiowa(target_path, surrogate_path, sample_metadata=metadata)

            rows = []
            match_count = 0
            for s in samples:
                if s.get("parameter_source") == "Uploaded UIowa submission file":
                    match_count += 1
                rows.append({
                    "Sample Name": s["sample_name"],
                    "UIowa Sample Key": s.get("uiowa_sample_key"),
                    "Data File": s.get("data_file"),
                    "Dilution Factor": float(s.get("dilution_factor") or 1.0),
                    "Extract Volume (mL)": float(s.get("extract_volume_ml") or 1.0),
                    "Fabric Size (cm^2)": value_or_blank(s.get("fabric_size_cm2")),
                    "Fabric Mass (g)": value_or_blank(s.get("fabric_mass_g")),
                    "Fuel Mass Loss (g)": value_or_blank(s.get("fuel_mass_loss_g")),
                    "Surrogate Spike (µg/mL)": float(s.get("surrogate_spike_ug_ml") or 0.05),
                    "DF Source": s.get("df_source"),
                    "Parameter Source": s.get("parameter_source"),
                })
            metadata_table = pd.DataFrame(rows)

        st.success(f"Detected {len(metadata_table)} sample(s). Submission-form matches found: {match_count}.")
        st.markdown("### Review or edit dilution and normalization inputs")
        st.write(
            "For UIowa output, Raw Calc. Conc. is read as ng/mL, but Corrected Conc. is reported as µg/mL. "
            "Normalization uses Corrected Conc. (µg/mL) × Extract Volume (mL)."
        )
        edited_df = st.data_editor(
            metadata_table,
            use_container_width=True,
            hide_index=True,
            disabled=["UIowa Sample Key", "Data File", "DF Source", "Parameter Source"],
            column_config={
                "Dilution Factor": st.column_config.NumberColumn("Dilution Factor", min_value=0.0, step=1.0, format="%.4f"),
                "Extract Volume (mL)": st.column_config.NumberColumn("Extract Volume (mL)", min_value=0.0, step=0.1, format="%.4f"),
                "Fabric Size (cm^2)": st.column_config.NumberColumn("Fabric Size (cm^2)", min_value=0.0, step=1.0, format="%.4f"),
                "Fabric Mass (g)": st.column_config.NumberColumn("Fabric Mass (g)", min_value=0.0, step=0.01, format="%.6f"),
                "Fuel Mass Loss (g)": st.column_config.NumberColumn("Fuel Mass Loss (g)", min_value=0.0, step=0.01, format="%.6f"),
                "Surrogate Spike (µg/mL)": st.column_config.NumberColumn("Surrogate Spike (µg/mL)", min_value=0.0, step=0.1, format="%.4f"),
            },
            key="uiowa_metadata_editor",
        )

        if st.button("Process University of Iowa files", type="primary", key="uiowa_process"):
            with st.spinner("Processing University of Iowa files..."):
                with tempfile.TemporaryDirectory() as tmp:
                    tmp = Path(tmp)
                    submission_path = tmp / submission_file.name
                    target_path = tmp / target_file.name
                    surrogate_path = tmp / surrogate_file.name
                    submission_path.write_bytes(submission_file.getbuffer())
                    target_path.write_bytes(target_file.getbuffer())
                    surrogate_path.write_bytes(surrogate_file.getbuffer())

                    reviewed_metadata = metadata_from_rows(edited_df.to_dict(orient="records"))
                    samples, meta = extract_uiowa(target_path, surrogate_path, sample_metadata=reviewed_metadata)
                    output_name = Path(target_file.name).stem + "_processed_v5_UIowa.xlsx"
                    output_path = tmp / output_name
                    write_processed_workbook_uiowa(samples, meta, output_path)
                    download_excel_button(output_path, output_name)

    except Exception as e:
        st.error("The app could not prepare or process the University of Iowa files.")
        st.code(str(e))


tab1, tab2 = st.tabs(["ISU / previous format", "University of Iowa format"])

with tab1:
    run_isu_v3_tab()

with tab2:
    run_uiowa_tab()

st.divider()
st.caption("GC/MS PAH Extractor v5 — unified platform for ISU-format and University of Iowa-format GC/MS outputs.")
