# Concrete Test Dashboard

This Streamlit app converts a batch of concrete-test PDF reports into a reviewable dashboard and downloadable Excel workbook.

## User workflow

1. Keep the original PDF filenames.
2. Upload individual PDFs, multiple PDFs, or a ZIP file containing a full folder and nested subfolders of PDFs.
3. Review the processing status and Concrete Test Log.
4. Download the formatted `Concrete_Test_Log.xlsx` workbook.
5. Review interactive charts for Air Temp, Concrete Temp, Slump, Air Content, Min Temp, and Max Temp.
6. Compare 7-day and 28-day strength results using separate second-order polynomial best-fit curves.

## Files

- `app.py`: Streamlit interface, PDF extraction, Excel generation, and chart creation
- `requirements.txt`: Python dependencies
- `.gitignore`: local files excluded from Git

## Run locally

```bash
python -m pip install -r requirements.txt
streamlit run app.py
```

## Deploy from GitHub

1. Put `app.py`, `requirements.txt`, and `.gitignore` in the repository root.
2. Push the files to GitHub.
3. Create an app in Streamlit Community Cloud.
4. Select the repository and branch.
5. Set the main file path to `app.py`.
6. Deploy.

## Upload notes

The uploader accepts `.pdf` and `.zip`. A ZIP is the reliable browser-friendly option for uploading a complete folder while preserving nested folder contents. The app reads ZIP contents in memory and does not add a skipped-files worksheet to the Excel download.
