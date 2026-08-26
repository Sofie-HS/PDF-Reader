# Concrete Test Dashboard

A memory-conscious Streamlit app that removes uploaded PDF objects immediately after processing.

## Accepted filename example

`W02229_Test_Sample 922_2025-03-13_Report 000020-01_Compressive Strength of Concrete_17359783.pdf`

The app accepts:

- `W02229_Test_` or `W02229_Concrete Test_`
- 3-digit or 4-digit Sample IDs
- 6-digit or 8-digit final filename identifiers

## Memory design

1. A user uploads PDFs or ZIP folders.
2. The user selects **Process reports and build Excel**.
3. The app reads one PDF at a time.
4. One extracted row is appended directly to a temporary Excel workbook.
5. PDF bytes, full extracted text, tables, and row variables are deleted.
6. After the batch is complete, the app changes the uploader widget key and reruns automatically.
7. The automatic rerun removes uploaded PDF and ZIP objects from the active uploader session.
8. Only the temporary workbook path and compact final table remain.
9. The workbook itself is read into bytes only when the download button is selected.
10. **Delete workbook and clear memory** removes the workbook, table, chart state, uploader state, and caches.

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
