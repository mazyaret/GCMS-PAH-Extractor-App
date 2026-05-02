# GC/MS PAH Streamlit Web App v4.1

Unified platform for:
- ISU / previous GC/MS format
- University of Iowa per-analyte worksheet format

## v4.1 update
For the University of Iowa format, raw `Calc. Conc.` values are read as **ng/mL**, but all output `Corrected Conc.` columns are now reported as **µg/mL**.

Formula:

```text
Corrected Conc. (µg/mL) = Raw Calc. Conc. (ng/mL) × Dilution Factor ÷ 1000
Mass (µg) = Corrected Conc. (µg/mL) × Extract Volume (mL)
```

Surrogate spike is also handled in **µg/mL** in the web editor and output QC sheet. Default surrogate spike is 0.5 µg/mL.

## Files for deployment
- `app_v4_1.py`
- `GCMS_PAH_Extractor_NoTableFix_v3.py`
- `GCMS_PAH_Extractor_UIowa_v2.py`
- `requirements.txt`
- `README_v4_1.md`

## Streamlit Cloud deployment
Use:

```text
Repository: mazyaret/GCMS-PAH-Extractor-App
Branch: main
Main file path: app_v4_1.py
```

Suggested app URL:

```text
gcms-pah-extractor-v4-1
```
