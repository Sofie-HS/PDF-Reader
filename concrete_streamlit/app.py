import gc
import io
import os
import re
import shutil
import tempfile
import zipfile
from datetime import datetime
from pathlib import Path

import streamlit as st


# ============================================================
# STARTUP WIPE
# ============================================================

# This runs once for each browser session before any PDF libraries are loaded.
# It clears Streamlit caches and removes files left by a previous failed run.
if "startup_wipe_complete" not in st.session_state:
    st.cache_data.clear()
    st.cache_resource.clear()

    cleanup_directory = Path(tempfile.gettempdir()) / "concrete_test_dashboard"
    shutil.rmtree(cleanup_directory, ignore_errors=True)
    cleanup_directory.mkdir(parents=True, exist_ok=True)

    st.session_state.clear()
    st.session_state["startup_wipe_complete"] = True
    gc.collect()


# ============================================================
# APP SETTINGS
# ============================================================

st.set_page_config(
    page_title="Concrete Test Dashboard",
    page_icon="📊",
    layout="wide",
)

# Deliberately small limits for the first stable cloud build.
MAX_PDF_FILES = 20
MAX_TOTAL_UPLOAD_MB = 25
MAX_SINGLE_PDF_MB = 8

FILENAME_PATTERN = re.compile(
    r"^W02229_(?:Concrete Test|Test)_"
    r"Sample (?P<sample_id>\d{3,4})_"
    r"(?P<report_date>\d{4}-\d{2}-\d{2})_"
    r"Report (?P<report_id>\d{6}-\d{2})_"
    r"Compressive Strength of Concrete_"
    r"(?:\d{6}|\d{8})\.pdf$",
    flags=re.IGNORECASE,
)

MASTER_COLUMNS = [
    "Report Date",
    "Sample ID",
    "Location Details",
    "Air Temp",
    "Concrete Temp",
    "Slump",
    "Air Content",
    "Min Temp",
    "Max Temp",
    "7 Day Test Average",
    "28 Day Test Average",
]

CHART_FIELDS = [
    "Air Temp",
    "Concrete Temp",
    "Slump",
    "Air Content",
    "Min Temp",
    "Max Temp",
]


# ============================================================
# SMALL HELPERS
# ============================================================


def clean_value(value):
    if value is None:
        return ""
    value = str(value).replace("\r", " ").replace("\n", " ")
    return re.sub(r"\s+", " ", value).strip()


def clean_number(value):
    if value is None:
        return None
    match = re.search(
        r"-?\d+(?:\.\d+)?",
        clean_value(value).replace(",", ""),
    )
    if not match:
        return None
    number = float(match.group(0))
    return int(number) if number.is_integer() else number


def extract_pattern(text, patterns):
    for pattern in patterns:
        match = re.search(
            pattern,
            text,
            flags=re.IGNORECASE | re.MULTILINE,
        )
        if match:
            return clean_value(match.group(1))
    return ""


def parse_mixed_number(value):
    text = clean_value(value)
    mixed = re.search(r"(\d+)\s*[- ]\s*(\d+)\s*/\s*(\d+)", text)
    if mixed:
        whole, numerator, denominator = map(int, mixed.groups())
        if denominator:
            return round(whole + numerator / denominator, 3)
    return clean_number(text)


def normalize_header(value):
    return re.sub(
        r"[^a-z0-9]+",
        " ",
        clean_value(value).lower(),
    ).strip()


def upload_size_bytes(uploaded_file):
    size = getattr(uploaded_file, "size", None)
    if size is not None:
        return size
    current = uploaded_file.tell()
    uploaded_file.seek(0, io.SEEK_END)
    size = uploaded_file.tell()
    uploaded_file.seek(current)
    return size


# ============================================================
# PDF EXTRACTION, ONE OPEN PER PDF
# ============================================================


def find_lab_header_row(table):
    for index, row in enumerate(table):
        combined = " | ".join(
            normalize_header(cell)
            for cell in (row or [])
        )
        if "test age days" in combined and "strength psi" in combined:
            return index
    return None


def extract_pdf_text_and_strengths(pdf_source):
    """Open one PDF once and return text plus completed strengths."""
    import pdfplumber

    pdf_source.seek(0)
    text_parts = []
    strengths_by_age = {}

    with pdfplumber.open(pdf_source) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                text_parts.append(text)

            # Use table extraction only on pages whose text mentions lab results.
            # This avoids scanning unrelated pages for tables.
            if not text or "strength" not in text.lower():
                continue

            for table in page.extract_tables():
                if not table:
                    continue

                header_index = find_lab_header_row(table)
                if header_index is None:
                    continue

                header_map = {}
                for column_index, heading in enumerate(table[header_index]):
                    normalized = normalize_header(heading)
                    if "test age days" in normalized:
                        header_map["age"] = column_index
                    elif "strength psi" in normalized:
                        header_map["strength"] = column_index

                if "age" not in header_map or "strength" not in header_map:
                    continue

                for row in table[header_index + 1:]:
                    if not row:
                        continue
                    age_index = header_map["age"]
                    strength_index = header_map["strength"]
                    age = clean_number(
                        row[age_index] if age_index < len(row) else ""
                    )
                    strength = clean_number(
                        row[strength_index] if strength_index < len(row) else ""
                    )
                    if age is not None and strength is not None:
                        strengths_by_age.setdefault(age, []).append(strength)

    pdf_source.seek(0)
    return "\n".join(text_parts), strengths_by_age


def extract_location_details(text):
    location = extract_pattern(
        text,
        [
            r"Location\s+Details\s*:\s*(.+?)(?=\n|$)",
            r"Location\s+Details\s*\n\s*(.+?)(?=\n|$)",
            r"Pour\s+Description\s*:\s*(.+?)(?=\n|$)",
        ],
    )
    return clean_value(location)


def extract_field_measurements(text):
    air_temp = extract_pattern(text, [
        r"Air\s+Temperature\s*(?:\(F\))?\s*:\s*(-?\d+(?:\.\d+)?)",
        r"Air\s+Temp(?:erature)?\s*(?:\(F\))?\s*:\s*(-?\d+(?:\.\d+)?)",
    ])
    concrete_temp = extract_pattern(text, [
        r"Concrete\s+Temp(?:erature)?\s*(?:\(F\))?\s*:\s*(-?\d+(?:\.\d+)?)",
    ])
    slump = extract_pattern(text, [
        r"Slump\s*(?:\(in\))?\s*:\s*(\d+\s*[- ]\s*\d+\s*/\s*\d+)",
        r"Slump\s*(?:\(in\))?\s*:\s*(\d+(?:\.\d+)?)",
    ])
    air_content = extract_pattern(text, [
        r"Air\s+Content\s*(?:\(%\))?\s*:\s*(-?\d+(?:\.\d+)?)",
    ])
    min_max = extract_pattern(text, [
        r"Min\s*/\s*Max\s+Temp\s*(?:\(F\))?\s*:\s*"
        r"(-?\d+(?:\.\d+)?\s*/\s*-?\d+(?:\.\d+)?)",
    ])

    min_temp = None
    max_temp = None
    if min_max:
        pieces = re.split(r"\s*/\s*", min_max, maxsplit=1)
        if len(pieces) == 2:
            min_temp = clean_number(pieces[0])
            max_temp = clean_number(pieces[1])

    return {
        "Air Temp": clean_number(air_temp),
        "Concrete Temp": clean_number(concrete_temp),
        "Slump": parse_mixed_number(slump),
        "Air Content": clean_number(air_content),
        "Min Temp": min_temp,
        "Max Temp": max_temp,
    }


def reported_average(text, age):
    return clean_number(extract_pattern(text, [
        rf"{age}\s*Day\s*-\s*([\d,]+(?:\.\d+)?)",
        rf"{age}\s*Day\s*Average\s*:?\s*([\d,]+(?:\.\d+)?)",
    ]))


def average_for_age(strengths_by_age, text, age):
    strengths = strengths_by_age.get(age, [])
    if strengths:
        return round(sum(strengths) / len(strengths))
    return reported_average(text, age)


def process_pdf(pdf_source, filename):
    match = FILENAME_PATTERN.fullmatch(filename)
    if not match:
        return None, "Filename does not match the accepted convention."

    info = match.groupdict()
    try:
        datetime.strptime(info["report_date"], "%Y-%m-%d")
    except ValueError:
        return None, "Filename contains an invalid date."

    text, strengths = extract_pdf_text_and_strengths(pdf_source)
    measurements = extract_field_measurements(text)

    record = {
        "Report Date": info["report_date"],
        "Sample ID": info["sample_id"],
        "Location Details": extract_location_details(text),
        **measurements,
        "7 Day Test Average": average_for_age(strengths, text, 7),
        "28 Day Test Average": average_for_age(strengths, text, 28),
    }

    del text
    del strengths
    del measurements
    gc.collect()
    return record, None


# ============================================================
# SEQUENTIAL PDF AND ZIP PROCESSING
# ============================================================


def count_pdfs(uploaded_items):
    total = 0
    notes = []

    for item in uploaded_items:
        if item.name.lower().endswith(".pdf"):
            total += 1
            continue

        if not item.name.lower().endswith(".zip"):
            notes.append((item.name, "Unsupported file type."))
            continue

        try:
            item.seek(0)
            with zipfile.ZipFile(item) as archive:
                count = sum(
                    1
                    for member in archive.infolist()
                    if not member.is_dir()
                    and member.filename.lower().endswith(".pdf")
                )
            total += count
            notes.append((item.name, f"Found {count} PDF file(s)."))
        except zipfile.BadZipFile:
            notes.append((item.name, "Could not read this ZIP file."))
        finally:
            item.seek(0)

    return total, notes


def process_uploads(uploaded_items, total_count, progress):
    records = []
    errors = []
    completed = 0

    for item in uploaded_items:
        lower_name = item.name.lower()

        if lower_name.endswith(".pdf"):
            size_mb = upload_size_bytes(item) / 1024 / 1024
            if size_mb > MAX_SINGLE_PDF_MB:
                errors.append((item.name, f"Skipped: {size_mb:.1f} MB."))
            else:
                try:
                    record, error = process_pdf(item, item.name)
                    if record is not None:
                        records.append(record)
                    else:
                        errors.append((item.name, error))
                except Exception as exc:
                    errors.append((item.name, f"{type(exc).__name__}: {exc}"))

            completed += 1
            progress.progress(
                completed / max(total_count, 1),
                text=f"Processed {completed} of {total_count} PDFs",
            )
            gc.collect()
            continue

        if not lower_name.endswith(".zip"):
            continue

        try:
            item.seek(0)
            with zipfile.ZipFile(item) as archive:
                for member in archive.infolist():
                    if member.is_dir() or not member.filename.lower().endswith(".pdf"):
                        continue

                    filename = Path(member.filename).name
                    size_mb = member.file_size / 1024 / 1024

                    if size_mb > MAX_SINGLE_PDF_MB:
                        errors.append((filename, f"Skipped: {size_mb:.1f} MB."))
                    else:
                        pdf_buffer = None
                        try:
                            pdf_buffer = io.BytesIO(archive.read(member))
                            record, error = process_pdf(pdf_buffer, filename)
                            if record is not None:
                                records.append(record)
                            else:
                                errors.append((filename, error))
                        except Exception as exc:
                            errors.append((filename, f"{type(exc).__name__}: {exc}"))
                        finally:
                            if pdf_buffer is not None:
                                pdf_buffer.close()
                            del pdf_buffer

                    completed += 1
                    progress.progress(
                        completed / max(total_count, 1),
                        text=f"Processed {completed} of {total_count} PDFs",
                    )
                    gc.collect()
        except zipfile.BadZipFile:
            errors.append((item.name, "Could not read this ZIP file."))
        finally:
            item.seek(0)

    return records, errors


# ============================================================
# EXCEL AND ONE SELECTED CHART
# ============================================================


def dataframe_to_excel(dataframe):
    import pandas as pd
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        dataframe.to_excel(
            writer,
            sheet_name="Concrete Test Log",
            index=False,
        )
        worksheet = writer.sheets["Concrete Test Log"]
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions

        fill = PatternFill(fill_type="solid", fgColor="1F4E78")
        font = Font(color="FFFFFF", bold=True)
        for cell in worksheet[1]:
            cell.fill = fill
            cell.font = font
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True,
            )

        widths = [15, 14, 52, 12, 16, 12, 14, 12, 12, 22, 24]
        for number, width in enumerate(widths, start=1):
            worksheet.column_dimensions[get_column_letter(number)].width = width

    output.seek(0)
    return output.getvalue()


def add_linear_series(figure, dataframe, x_field, y_field, name, color, symbol):
    import numpy as np
    import plotly.graph_objects as go

    plot_data = dataframe[
        [x_field, y_field, "Sample ID", "Location Details"]
    ].dropna()

    if plot_data.empty:
        return

    figure.add_trace(go.Scatter(
        x=plot_data[x_field],
        y=plot_data[y_field],
        mode="markers",
        name=name,
        marker={"color": color, "size": 9, "symbol": symbol},
    ))

    if len(plot_data) >= 2 and plot_data[x_field].nunique() >= 2:
        x_values = plot_data[x_field].astype(float).to_numpy()
        y_values = plot_data[y_field].astype(float).to_numpy()
        slope, intercept = np.polyfit(x_values, y_values, 1)
        x_line = np.linspace(x_values.min(), x_values.max(), 80)
        y_line = slope * x_line + intercept
        figure.add_trace(go.Scatter(
            x=x_line,
            y=y_line,
            mode="lines",
            name=f"{name} best fit",
            line={"color": color, "width": 3},
        ))


def build_chart(dataframe, x_field):
    import plotly.graph_objects as go

    figure = go.Figure()
    add_linear_series(
        figure, dataframe, x_field, "7 Day Test Average",
        "7 Day Average", "#4472C4", "circle",
    )
    add_linear_series(
        figure, dataframe, x_field, "28 Day Test Average",
        "28 Day Average", "#ED7D31", "diamond",
    )
    figure.update_layout(
        title=f"{x_field} vs Concrete Compressive Strength",
        template="plotly_white",
        xaxis_title=x_field,
        yaxis_title="Concrete Compressive Strength Average (PSI)",
        height=520,
    )
    return figure


# ============================================================
# PAGE
# ============================================================

st.title("Concrete Test Dashboard")
st.caption(
    "Startup-wipe build. Old temporary files and Streamlit caches are cleared "
    "once when the browser session begins."
)

with st.expander("How to use this tool", expanded=True):
    st.markdown(
        """
        1. Start with 2 to 5 PDFs to verify the deployment.
        2. Upload PDFs or one ZIP folder.
        3. The app processes each ZIP member one at a time.
        4. Review the table, download Excel, and choose one chart.
        5. If the app has previously exceeded limits, reboot the app once after deploying this version.
        """
    )

uploaded_items = st.file_uploader(
    "Upload concrete test PDFs or a ZIP folder",
    type=["pdf", "zip"],
    accept_multiple_files=True,
)

if not uploaded_items:
    st.info("Upload a small test batch to begin.")
    st.stop()

total_upload_mb = sum(
    upload_size_bytes(item)
    for item in uploaded_items
) / 1024 / 1024

if total_upload_mb > MAX_TOTAL_UPLOAD_MB:
    st.error(
        f"This upload is {total_upload_mb:.1f} MB. "
        f"Please keep the test batch below {MAX_TOTAL_UPLOAD_MB} MB."
    )
    st.stop()

total_count, upload_notes = count_pdfs(uploaded_items)

if total_count > MAX_PDF_FILES:
    st.error(
        f"This upload contains {total_count} PDFs. "
        f"Please keep the test batch below {MAX_PDF_FILES} PDFs."
    )
    st.stop()

progress = st.progress(0, text="Processing reports...")
records, processing_errors = process_uploads(
    uploaded_items,
    total_count,
    progress,
)
progress.empty()

if not records:
    st.error("No valid report data was extracted.")
    with st.expander("Processing details"):
        for name, message in upload_notes + processing_errors:
            st.write(f"{name}: {message}")
    st.stop()

# Heavy libraries are imported only after PDF processing succeeds.
import pandas as pd

master_df = pd.DataFrame(records, columns=MASTER_COLUMNS)
master_df["Report Date"] = pd.to_datetime(
    master_df["Report Date"],
    errors="coerce",
)
master_df.sort_values(
    ["Report Date", "Sample ID"],
    inplace=True,
    na_position="last",
)
master_df.reset_index(drop=True, inplace=True)

st.subheader("Concrete Test Log")
st.dataframe(master_df, use_container_width=True, hide_index=True)

excel_bytes = dataframe_to_excel(master_df)
st.download_button(
    "Download Concrete Test Log (.xlsx)",
    data=excel_bytes,
    file_name="Concrete_Test_Log.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)

if upload_notes or processing_errors:
    with st.expander("Processing details"):
        for name, message in upload_notes + processing_errors:
            st.write(f"{name}: {message}")

st.subheader("Strength chart")
selected_chart = st.selectbox(
    "Choose the field measurement",
    CHART_FIELDS,
)
chart = build_chart(master_df, selected_chart)
st.plotly_chart(chart, use_container_width=True)
