# Concrete Test Dashboard

This architecture keeps large PDF processing out of the browser.

## Why

Streamlit Community Cloud has limited RAM, and uploaded files are held in server memory. A source folder larger than 1 GB should therefore be processed locally rather than uploaded into the Community Cloud app.

## Step 1: Process the PDF folder locally

Install local dependencies:

```powershell
python -m pip install -r requirements-local.txt
```

Run with the helper:

```text
run_processor.bat
```

Or run directly:

```powershell
python processor.py "C:\path\to\pdf-folder" --output "Concrete_Test_Log.xlsx"
```

The local processor:

- Accepts `Test` and `Concrete Test`
- Accepts 3-digit and 4-digit Sample IDs
- Processes nested folders
- Reads one PDF at a time
- Writes each result to Excel
- Produces `Skipped_PDFs.txt` when needed

## Step 2: Run the Streamlit dashboard

```powershell
python -m pip install -r requirements.txt
streamlit run app.py
```

Upload only `Concrete_Test_Log.xlsx`. The app does not display the full Excel table. It shows summary counts, provides the workbook download, and displays one tightly ranged quadratic chart at a time.

## Deploy

Deploy `app.py` with `requirements.txt` to Streamlit Community Cloud. `processor.py` and `requirements-local.txt` remain available in the GitHub repository for local folder processing.
