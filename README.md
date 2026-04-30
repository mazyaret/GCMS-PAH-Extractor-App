GC/MS PAH Streamlit Web App
===========================

This is the first-step web version of the GC/MS PAH extractor.

User workflow:
1. Open the Streamlit app.
2. Upload the original GC/MS Excel file.
3. Optional: upload a dilution-factor template.
4. Click "Process GC/MS file".
5. Download the processed Excel workbook.

Files:
- app.py
- GCMS_PAH_Extractor_NoTableFix.py
- requirements.txt

Local test:
    pip install -r requirements.txt
    streamlit run app.py

Deployment:
Upload these files to a GitHub repository and deploy app.py on Streamlit Community Cloud.
