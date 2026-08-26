# Concrete Test Dashboard

This build performs a forced startup wipe once per browser session before PDF libraries are loaded.

## Startup wipe

- Clears Streamlit data and resource caches
- Deletes `/tmp/concrete_test_dashboard`
- Recreates a clean temporary directory
- Runs garbage collection

## Stability changes

- Process ZIP members one at a time
- Open each PDF only once
- Run table detection only on pages whose extracted text contains `strength`
- Import pandas, openpyxl, NumPy, and Plotly only after PDF processing succeeds
- Display one selected chart
- Limit the first stable build to 20 PDFs, 25 MB total, and 8 MB per PDF

## Accepted example

`W02229_Test_Sample 922_2025-03-13_Report 000020-01_Compressive Strength of Concrete_17359783.pdf`
