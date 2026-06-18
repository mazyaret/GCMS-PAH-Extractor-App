# GC/MS PAH Streamlit Web App v5.2

Unified platform with two tabs:

1. **ISU / previous format**
   - Upload one GC/MS Excel output file.
   - Optional: upload submission/metadata file.
   - Review dilution factor, extract volume, fabric size, fabric mass, and fuel mass loss.

2. **University of Iowa format**
   - Upload submission file.
   - Upload 16-PAH results workbook.
   - Upload surrogate results workbook.
   - Review dilution factor and normalization inputs.

## New in v5

The output keeps all previous columns and adds **surrogate recovery-corrected** columns.

Formula:

```
Recovery-Corrected Conc. = Corrected Conc. × 100 / Assigned Surrogate Recovery %
```

These are added in addition to the original DF-corrected columns, not replacing them.

## ISU recovery surrogate mapping

- Naphthalene, Acenaphthylene, Acenaphthene → 2-fluorobiphenyl
- Fluorene, Phenanthrene, Anthracene → Fluorene-d10
- Fluoranthene, Pyrene, Benz[a]anthracene, Chrysene → Pyrene-d10
- Benzo[b]fluoranthene, Benzo[k]fluoranthene, Benzo[a]pyrene, Indeno[1,2,3-cd]pyrene, Dibenz[a,h]anthracene, Benzo[g,h,i]perylene → Benzo[a]pyrene-d12

## UIowa recovery surrogate mapping

- LMW PAHs → 2-Fluorobiphenyl
- HMW PAHs → p-Terphenyl-d14

## Deployment on Streamlit Cloud

Upload these files to GitHub:

- `app_v5.py`
- `GCMS_PAH_Extractor_NoTableFix_v5.py`
- `GCMS_PAH_Extractor_UIowa_v3.py`
- `requirements.txt`
- `README_v5.md`

Create a new Streamlit app with:

- Repository: `mazyaret/GCMS-PAH-Extractor-App`
- Branch: `main`
- Main file path: `app_v5.py`

Suggested app URL: `gcms-pah-extractor-v5`


## v5.1 changes
- TEF for Dibenz[a,h]anthracene changed from 5 to 1 for both ISU and University of Iowa workflows.
- University of Iowa default surrogate spike concentration changed to 0.05 µg/mL.
- ISU surrogate spike concentration remains 0.5 µg/mL.

Recommended Streamlit main file path: `app_v5_1.py` if renamed, or use `app_v5.py` if replacing the previous v5 app code.


## v5.2 changes
- Fixed decimal entry in the web editor for normalization fields such as Fabric Size (cm²).
- Editable numeric fields are now entered as text and converted safely to numbers during processing.
- Decimal dot and decimal comma are both accepted: `112.5` and `112,5`.
- The original output calculations and recovery-corrected columns remain unchanged.

Recommended Streamlit main file path: `app_v5_2.py`.


V5.4 update:
- The app now automatically recognizes the ISU submission-form column labeled `Surrogate (ug/ml)` as the sample-level surrogate spike concentration.
- It still also recognizes labels such as `Surrogate Spike (µg/mL)`, `Surrogate Spike Conc.`, and similar variants.
