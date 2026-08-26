# Concrete Test Dashboard

A memory-conscious Streamlit app that processes concrete-test PDFs one at a time, writes each extracted row directly into Excel, and releases the PDF content before the next report is opened.

## Accepted filenames

Both naming prefixes are accepted, and the Sample ID may contain 3 or 4 digits:

- `W02229_Concrete Test_Sample ###_YYYY-MM-DD_Report ######-##_Compressive Strength of Concrete_######.pdf`
- `W02229_Test_Sample ###_YYYY-MM-DD_Report ######-##_Compressive Strength of Concrete_######.pdf`
- `W02229_Concrete Test_Sample ####_YYYY-MM-DD_Report ######-##_Compressive Strength of Concrete_######.pdf`
- `W02229_Test_Sample ####_YYYY-MM-DD_Report ######-##_Compressive Strength of Concrete_######.pdf`

Example accepted filename:

`W02229_Test_Sample 922_2025-03-13_Report 000020-01_Compressive Strength of Concrete_17359783.pdf`

The last filename identifier may contain 6 or 8 digits.

## Memory behavior

- Direct PDFs and ZIP members are processed sequentially.
- Each PDF is opened once for text and lab-table extraction.
- Each extracted row is appended directly to a temporary Excel workbook.
- PDF bytes, extracted full text, and row-level temporary variables are released before the next PDF is read.
- Only the compact final Excel table is loaded for the dashboard and selected chart.
- Only one Plotly chart is built at a time.
- A bottom-page **Clear uploaded files and reset memory** button resets the uploader, Streamlit caches, the current table, workbook bytes, and chart data.

## Repository files

- `app.py`
- `requirements.txt`
- `.gitignore`
- `.streamlit/config.toml`

## Run locally

```bash
python -m pip install -r requirements.txt
streamlit run app.py
```
