import streamlit as st
import tempfile
import subprocess
import sys
from pathlib import Path

st.set_page_config(
    page_title="GC/MS PAH Extractor",
    page_icon="🧪",
    layout="centered"
)

st.title("GC/MS PAH Extraction Platform")
st.write(
    "Upload the original GC/MS Excel file. The app will extract Final Conc., "
    "apply dilution factors, calculate surrogate recovery, and return a processed Excel workbook."
)

st.info(
    "For the first version, only the original GC/MS Excel file is required. "
    "A dilution-factor template is optional."
)

gcms_file = st.file_uploader(
    "1) Upload original GC/MS Excel file (.xlsx)",
    type=["xlsx"]
)

df_template = st.file_uploader(
    "2) Optional: Upload dilution factor template (.xlsx)",
    type=["xlsx"]
)

process = st.button("Process GC/MS file")

if process:
    if gcms_file is None:
        st.error("Please upload the original GC/MS Excel file first.")
    else:
        with st.spinner("Processing file..."):
            try:
                with tempfile.TemporaryDirectory() as tmpdir:
                    tmpdir = Path(tmpdir)

                    input_path = tmpdir / gcms_file.name
                    input_path.write_bytes(gcms_file.getbuffer())

                    output_name = Path(gcms_file.name).stem + "_processed.xlsx"
                    output_path = tmpdir / output_name

                    script_path = Path(__file__).parent / "GCMS_PAH_Extractor_NoTableFix.py"

                    cmd = [
                        sys.executable,
                        str(script_path),
                        str(input_path),
                        "--output",
                        str(output_path)
                    ]

                    if df_template is not None:
                        df_path = tmpdir / df_template.name
                        df_path.write_bytes(df_template.getbuffer())
                        cmd.extend(["--df-template", str(df_path)])

                    result = subprocess.run(
                        cmd,
                        capture_output=True,
                        text=True
                    )

                    if result.returncode != 0:
                        st.error("Processing failed. Please check the error message below.")
                        st.code(result.stderr or result.stdout)
                    elif not output_path.exists():
                        st.error("The script finished, but no output file was created.")
                        st.code(result.stdout + "\n" + result.stderr)
                    else:
                        processed_bytes = output_path.read_bytes()
                        st.success("Done! Download your processed Excel file below.")

                        st.download_button(
                            label="Download processed Excel file",
                            data=processed_bytes,
                            file_name=output_name,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )

                        with st.expander("Processing log"):
                            st.code(result.stdout or "No additional log.")
            except Exception as e:
                st.error("Unexpected error.")
                st.code(str(e))

st.divider()
st.caption("GC/MS PAH Extractor — first-step upload/process/download version.")
