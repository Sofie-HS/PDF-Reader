# Concrete Test Dashboard

This build addresses a runtime crash that could occur after processing started.

## Important fixes

- The app no longer calls `.close()` on Streamlit's original `UploadedFile` objects before a rerun. Streamlit owns those objects, and the uploader is instead cleared by changing its widget key after processing.
- ZIP members are still opened one at a time, written to the workbook, closed, deleted, and garbage-collected.
- The final workbook is held in memory only when the results page needs the download button.
- Package versions are pinned for repeatable Streamlit Community Cloud deployments.
- Safer Community Cloud limits are used: 60 PDFs, 75 MB total upload, and 15 MB per PDF.

## Accepted example

`W02229_Test_Sample 922_2025-03-13_Report 000020-01_Compressive Strength of Concrete_17359783.pdf`

## Deploy

Place all repository files at the GitHub repository root, push, then reboot the Streamlit app.
