# GC/MS PAH Streamlit Web App v4

This is the unified platform for two GC/MS output formats.

## Option 1: ISU / previous format
Use this tab for the earlier GC/MS output format where one workbook contains the analyte result blocks and Final Conc. columns.

Required:
- Original GC/MS Excel output

Optional:
- Submission/metadata form with dilution factor, extract volume, fabric size, fabric mass, and fuel mass loss.

## Option 2: University of Iowa format
Use this tab for University of Iowa output files.

Required:
1. University of Iowa sample submission form
2. University of Iowa 16-PAH results workbook
3. University of Iowa surrogate results workbook

The app:
- uses Type = Sample rows only,
- reads Calc. Conc. values,
- ignores calibration and QC rows,
- applies dilution factor from the submission form,
- allows manual review/editing before processing,
- generates one processed workbook with raw/corrected values, normalized outputs, Total/LMW/HMW PAHs, TEQ, and surrogate recovery.

## Deployment
Upload these files to GitHub:
- app_v4.py
- GCMS_PAH_Extractor_NoTableFix_v3.py
- GCMS_PAH_Extractor_UIowa_v1.py
- requirements.txt
- README_v4.md

Then deploy a new Streamlit app with:
- Main file path: app_v4.py
