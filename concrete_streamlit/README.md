# Concrete Test Dashboard

A Streamlit app that:

- Accepts multiple concrete-test PDF uploads
- Extracts the 4-digit Sample ID, report date, Location Details, field measurements, and 7-day/28-day strength averages
- Displays the combined Concrete Test Log
- Generates an Excel workbook for download
- Displays separate interactive scatter plots with 7-day and 28-day linear best-fit lines

## Run locally

```bash
python -m pip install -r requirements.txt
streamlit run app.py
```

## Deploy from GitHub with Streamlit Community Cloud

1. Push `app.py`, `requirements.txt`, and `.gitignore` to the repository root.
2. In Streamlit Community Cloud, create an app from the repository.
3. Set the main file path to `app.py`.
4. Deploy.

The app reads PDFs uploaded in the browser, so it does not use a local Windows folder path.
