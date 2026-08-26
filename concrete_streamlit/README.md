# Concrete Test Dashboard

This folder restores the original simple Streamlit design that worked before the later session-state, automatic-rerun, temporary-file, cache-clearing, and memory-button changes.

## Included behavior

- Upload individual PDFs, multiple PDFs, or ZIP folders
- Accept `Test` and `Concrete Test` filename variants
- Accept 3-digit and 4-digit Sample IDs
- Extract Location Details and field measurements
- Calculate 7-day and 28-day average strengths
- Display the master table
- Download a formatted Excel workbook
- Display separate charts with linear best-fit lines

## Accepted example

`W02229_Test_Sample 922_2025-03-13_Report 000020-01_Compressive Strength of Concrete_17359783.pdf`

## Run locally

```bash
python -m pip install -r requirements.txt
streamlit run app.py
```
