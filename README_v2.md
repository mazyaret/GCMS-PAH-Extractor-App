# GC/MS PAH Streamlit Web App v2

User workflow:
1. Upload the original GC/MS Excel file.
2. Optional: upload the lab submission form or a DF template.
3. Review or manually edit dilution factors in the web table.
4. Click **Process GC/MS file**.
5. Download the processed Excel workbook.

The optional DF file may be:
- a simple template with headers `Sample Name` and `Dilution Factor`, or
- the lab sample submission form with headers such as `Sample Name (Unique)` and `Dilution factor`.

For deployment on Streamlit Cloud:
- Rename `app_v2.py` to `app.py`, or set main file path to `app_v2.py`.
- Rename `requirements_v2.txt` to `requirements.txt`.
- Upload `GCMS_PAH_Extractor_NoTableFix_v2.py`.
