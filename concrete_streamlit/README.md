# Concrete Test Dashboard

This is the low-memory build.

## Key change

The app no longer uses `pdfplumber` table extraction. It uses `pypdf` for plain text extraction and reads the printed 7-day and 28-day average-strength values from the PDF text. Table detection was removed because it was the highest-memory operation.

## Processing order

1. Read one PDF.
2. Extract one row.
3. Append that row to a temporary CSV.
4. Delete PDF text and buffers.
5. Repeat.
6. After every PDF is released, create the formatted Excel workbook once.
7. Reset the upload widget.
8. Display the compact CSV table and one selected chart.

## Accepted filename

`W02229_Test_Sample 922_2025-03-13_Report 000020-01_Compressive Strength of Concrete_17359783.pdf`

## Limits

- 50 PDFs per batch
- 50 MB total upload
- 10 MB per PDF
