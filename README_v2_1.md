# GC/MS PAH Streamlit Web App v2.1

This version keeps v1 and v2 available, but improves dilution-factor matching.

User workflow:
1. Upload the original GC/MS Excel file.
2. Optional: upload the lab submission form or a simple DF template.
3. The app reads the DF from the uploaded form/template and matches samples even if the GC/MS sample label has the `8X` removed.
4. Review or manually edit dilution factors in the web table.
5. Click **Process GC/MS file**.
6. Download the processed Excel workbook.

The optional DF file may be:
- a simple template with headers `Sample Name` and `Dilution Factor`, or
- the lab sample submission form with headers such as `Sample Name (Unique)` and `Dilution factor`.

Main deployment settings on Streamlit Cloud:
- Repository: `mazyaret/GCMS-PAH-Extractor-App`
- Branch: `main`
- Main file path: `app_v2_1.py`

Required files:
- `app_v2_1.py`
- `GCMS_PAH_Extractor_NoTableFix_v2_1.py`
- `requirements.txt` containing `streamlit`, `openpyxl`, and `pandas`
