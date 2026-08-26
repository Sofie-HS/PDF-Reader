import csv
import gc
import io
import re
import zipfile
from datetime import datetime
from pathlib import Path
from uuid import uuid4

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from pypdf import PdfReader


# ============================================================
# 1. APP AND RESOURCE SETTINGS
# ============================================================

st.set_page_config(
    page_title="Concrete Test Dashboard",
    page_icon="📊",
    layout="wide",
)

# Conservative limits for Streamlit Community Cloud.
MAX_PDF_FILES = 50
MAX_TOTAL_UPLOAD_MB = 50
MAX_SINGLE_PDF_MB = 10

RESULT_DIRECTORY = Path("/tmp/concrete_test_dashboard")
RESULT_DIRECTORY.mkdir(parents=True, exist_ok=True)

# Accepts Test or Concrete Test, 3-digit or 4-digit Sample IDs,
# and a 6-digit or 8-digit final identifier.
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
# 2. SMALL PARSING HELPERS
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
    fraction = re.fullmatch(r"\s*(\d+)\s*/\s*(\d+)\s*", text)
    if fraction:
        numerator, denominator = map(int, fraction.groups())
        if denominator:
            return round(numerator / denominator, 3)
    return clean_number(text)


# ============================================================
# 3. LOW-MEMORY PDF TEXT EXTRACTION
# ============================================================


def extract_pdf_text(pdf_source):
    """
    Extract plain text only.

    This build intentionally does not run table detection, which was the
    highest-memory operation in the prior build. The required 7-day and
    28-day results are read from the report's printed average-strength text.
    """
    pdf_source.seek(0)
    reader = PdfReader(pdf_source, strict=False)
    page_text_parts = []

    for page in reader.pages:
        text = page.extract_text()
        if text:
            page_text_parts.append(text)

    result = "\n".join(page_text_parts)
    del reader
    pdf_source.seek(0)
    return result


def extract_location_details(text):
    location = extract_pattern(
        text,
        [
            r"Location\s+Details\s*:\s*(.+?)(?=\n|$)",
            r"Location\s+Details\s*\n\s*(.+?)(?=\n|$)",
            r"Location\s+Detail\s*:\s*(.+?)(?=\n|$)",
            r"Pour\s+Description\s*:\s*(.+?)(?=\n|$)",
        ],
    )

    for label in [
        "Mix Design",
        "Supplier",
        "Technician",
        "Weather",
        "Field Measurements",
        "On-Site Admixtures",
    ]:
        location = re.split(
            rf"\s+{re.escape(label)}\s*:",
            location,
            maxsplit=1,
            flags=re.IGNORECASE,
        )[0]

    return clean_value(location)


def extract_field_measurements(text):
    air_temp = extract_pattern(
        text,
        [
            r"Air\s+Temperature\s*(?:\(F\))?\s*:\s*(-?\d+(?:\.\d+)?)",
            r"Air\s+Temp(?:erature)?\s*(?:\(F\))?\s*:\s*(-?\d+(?:\.\d+)?)",
        ],
    )

    concrete_temp = extract_pattern(
        text,
        [
            r"Concrete\s+Temp(?:erature)?\s*(?:\(F\))?\s*:\s*(-?\d+(?:\.\d+)?)",
        ],
    )

    slump = extract_pattern(
        text,
        [
            r"Slump\s*(?:\(in\))?\s*:\s*(\d+\s*[- ]\s*\d+\s*/\s*\d+)",
            r"Slump\s*(?:\(in\))?\s*:\s*(\d+\s*/\s*\d+)",
            r"Slump\s*(?:\(in\))?\s*:\s*(\d+(?:\.\d+)?)",
        ],
    )

    air_content = extract_pattern(
        text,
        [
            r"Air\s+Content\s*(?:\(%\))?\s*:\s*(-?\d+(?:\.\d+)?)",
        ],
    )

    min_max = extract_pattern(
        text,
        [
            r"Min\s*/\s*Max\s+Temp\s*(?:\(F\))?\s*:\s*"
            r"(-?\d+(?:\.\d+)?\s*/\s*-?\d+(?:\.\d+)?)",
        ],
    )

    min_temp = None
    max_temp = None
    if min_max:
        parts = re.split(r"\s*/\s*", min_max, maxsplit=1)
        if len(parts) == 2:
            min_temp = clean_number(parts[0])
            max_temp = clean_number(parts[1])

    return {
        "Air Temp": clean_number(air_temp),
        "Concrete Temp": clean_number(concrete_temp),
        "Slump": parse_mixed_number(slump),
        "Air Content": clean_number(air_content),
        "Min Temp": min_temp,
        "Max Temp": max_temp,
    }


def extract_test_average(text, test_age):
    value = extract_pattern(
        text,
        [
            rf"{test_age}\s*Day\s*-\s*([\d,]+(?:\.\d+)?)",
            rf"{test_age}\s*Day\s*Average\s*:?\s*([\d,]+(?:\.\d+)?)",
            rf"{test_age}[-\s]*Day\s*Average\s*:?\s*([\d,]+(?:\.\d+)?)",
        ],
    )
    return clean_number(value)


def process_pdf_to_row(pdf_source, filename):
    match = FILENAME_PATTERN.fullmatch(filename)
    if not match:
        return None, "Filename does not match an accepted convention."

    info = match.groupdict()
    try:
        datetime.strptime(info["report_date"], "%Y-%m-%d")
    except ValueError:
        return None, "Filename contains an invalid report date."

    text = extract_pdf_text(pdf_source)
    measurements = extract_field_measurements(text)

    row = [
        info["report_date"],
        info["sample_id"],
        extract_location_details(text),
        measurements["Air Temp"],
        measurements["Concrete Temp"],
        measurements["Slump"],
        measurements["Air Content"],
        measurements["Min Temp"],
        measurements["Max Temp"],
        extract_test_average(text, 7),
        extract_test_average(text, 28),
    ]

    del measurements
    del text
    gc.collect()
    return row, None


# ============================================================
# 4. TEMP CSV AND FINAL XLSX
# ============================================================


def create_csv(csv_path):
    with open(csv_path, "w", newline="", encoding="utf-8") as csv_file:
        writer = csv.writer(csv_file)
        writer.writerow(MASTER_COLUMNS)


def append_csv_row(csv_path, row):
    with open(csv_path, "a", newline="", encoding="utf-8") as csv_file:
        writer = csv.writer(csv_file)
        writer.writerow(row)


def csv_to_formatted_xlsx(csv_path, xlsx_path):
    """Create the final XLSX once, after all PDFs are gone."""
    workbook = Workbook(write_only=False)
    worksheet = workbook.active
    worksheet.title = "Concrete Test Log"

    with open(csv_path, "r", newline="", encoding="utf-8") as csv_file:
        for row in csv.reader(csv_file):
            worksheet.append(row)

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions
    worksheet.sheet_view.showGridLines = False

    header_fill = PatternFill(fill_type="solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)

    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )

    worksheet.row_dimensions[1].height = 38
    widths = [15, 14, 52, 12, 16, 12, 14, 12, 12, 22, 24]

    for column_number, width in enumerate(widths, start=1):
        worksheet.column_dimensions[get_column_letter(column_number)].width = width

    for row_number in range(2, worksheet.max_row + 1):
        date_text = worksheet[f"A{row_number}"].value
        if date_text:
            try:
                worksheet[f"A{row_number}"].value = datetime.strptime(
                    date_text,
                    "%Y-%m-%d",
                )
                worksheet[f"A{row_number}"].number_format = "yyyy-mm-dd"
            except ValueError:
                pass

        for column_letter in ["D", "E", "F", "G", "H", "I", "J", "K"]:
            cell = worksheet[f"{column_letter}{row_number}"]
            if cell.value not in (None, ""):
                cell.value = clean_number(cell.value)

        for column_letter in ["D", "E", "F", "G", "H", "I"]:
            worksheet[f"{column_letter}{row_number}"].number_format = "0.00"
        worksheet[f"J{row_number}"].number_format = "#,##0"
        worksheet[f"K{row_number}"].number_format = "#,##0"

    workbook.save(xlsx_path)
    workbook.close()


def load_result_dataframe(csv_path):
    dataframe = pd.read_csv(csv_path)
    dataframe["Report Date"] = pd.to_datetime(
        dataframe["Report Date"],
        errors="coerce",
    )
    dataframe.sort_values(
        ["Report Date", "Sample ID"],
        inplace=True,
        na_position="last",
    )
    dataframe.reset_index(drop=True, inplace=True)
    return dataframe


# ============================================================
# 5. SEQUENTIAL UPLOAD PROCESSING
# ============================================================


def upload_size_bytes(uploaded_file):
    size = getattr(uploaded_file, "size", None)
    if size is not None:
        return size
    current = uploaded_file.tell()
    uploaded_file.seek(0, io.SEEK_END)
    size = uploaded_file.tell()
    uploaded_file.seek(current)
    return size


def count_pdfs(uploaded_items):
    count = 0
    notes = []

    for item in uploaded_items:
        lower_name = item.name.lower()
        if lower_name.endswith(".pdf"):
            count += 1
            continue
        if not lower_name.endswith(".zip"):
            notes.append((item.name, "Unsupported file type."))
            continue
        try:
            item.seek(0)
            with zipfile.ZipFile(item) as archive:
                item_count = sum(
                    1
                    for member in archive.infolist()
                    if not member.is_dir()
                    and member.filename.lower().endswith(".pdf")
                )
            count += item_count
            notes.append((item.name, f"Found {item_count} PDF file(s)."))
        except zipfile.BadZipFile:
            notes.append((item.name, "Could not read this ZIP file."))
        finally:
            item.seek(0)

    return count, notes


def process_uploads(uploaded_items, csv_path, total_count, progress):
    processed = 0
    successful = 0
    messages = []

    for item in uploaded_items:
        lower_name = item.name.lower()

        if lower_name.endswith(".pdf"):
            size_mb = upload_size_bytes(item) / 1024 / 1024
            if size_mb > MAX_SINGLE_PDF_MB:
                messages.append((item.name, f"Skipped: {size_mb:.1f} MB is too large."))
            else:
                try:
                    row, error = process_pdf_to_row(item, item.name)
                    if row is not None:
                        append_csv_row(csv_path, row)
                        successful += 1
                        del row
                    else:
                        messages.append((item.name, error))
                except Exception as exc:
                    messages.append((item.name, f"{type(exc).__name__}: {exc}"))

            processed += 1
            progress.progress(
                processed / max(total_count, 1),
                text=f"Processed {processed} of {total_count} PDFs",
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
                    pdf_buffer = None

                    if size_mb > MAX_SINGLE_PDF_MB:
                        messages.append((filename, f"Skipped: {size_mb:.1f} MB is too large."))
                    else:
                        try:
                            pdf_buffer = io.BytesIO(archive.read(member))
                            row, error = process_pdf_to_row(pdf_buffer, filename)
                            if row is not None:
                                append_csv_row(csv_path, row)
                                successful += 1
                                del row
                            else:
                                messages.append((filename, error))
                        except Exception as exc:
                            messages.append((filename, f"{type(exc).__name__}: {exc}"))
                        finally:
                            if pdf_buffer is not None:
                                pdf_buffer.close()
                            del pdf_buffer

                    processed += 1
                    progress.progress(
                        processed / max(total_count, 1),
                        text=f"Processed {processed} of {total_count} PDFs",
                    )
                    gc.collect()
        except zipfile.BadZipFile:
            messages.append((item.name, "Could not read this ZIP file."))
        finally:
            item.seek(0)

    return successful, messages


# ============================================================
# 6. QUADRATIC CHART
# ============================================================


def add_quadratic_series(figure, dataframe, x_field, y_field, name, color, symbol):
    plot_data = dataframe[
        [x_field, y_field, "Sample ID", "Location Details"]
    ].dropna()

    if plot_data.empty:
        return

    figure.add_trace(
        go.Scatter(
            x=plot_data[x_field],
            y=plot_data[y_field],
            mode="markers",
            name=name,
            marker={"color": color, "size": 9, "symbol": symbol},
            customdata=plot_data[["Sample ID", "Location Details"]],
            hovertemplate=(
                "Sample ID: %{customdata[0]}<br>"
                "Location: %{customdata[1]}<br>"
                f"{x_field}: %{{x}}<br>"
                f"{y_field}: %{{y:,.0f}} PSI<extra></extra>"
            ),
        )
    )

    if len(plot_data) < 3 or plot_data[x_field].nunique() < 3:
        return

    x_values = plot_data[x_field].astype(float).to_numpy()
    y_values = plot_data[y_field].astype(float).to_numpy()
    a_value, b_value, c_value = np.polyfit(x_values, y_values, 2)

    x_line = np.linspace(x_values.min(), x_values.max(), 100)
    y_line = a_value * x_line ** 2 + b_value * x_line + c_value
    y_predicted = a_value * x_values ** 2 + b_value * x_values + c_value

    residual = np.sum((y_values - y_predicted) ** 2)
    total = np.sum((y_values - np.mean(y_values)) ** 2)
    r_squared = 1 - residual / total if total else np.nan

    equation = f"y={a_value:.3f}x²+{b_value:.2f}x+{c_value:.1f}"
    r_text = f"R²={r_squared:.3f}" if not np.isnan(r_squared) else "R² unavailable"

    figure.add_trace(
        go.Scatter(
            x=x_line,
            y=y_line,
            mode="lines",
            name=f"{name} quadratic ({equation}; {r_text})",
            line={"color": color, "width": 3},
            hoverinfo="skip",
        )
    )


def build_chart(dataframe, x_field):
    figure = go.Figure()
    add_quadratic_series(
        figure, dataframe, x_field, "7 Day Test Average",
        "7 Day Average", "#4472C4", "circle",
    )
    add_quadratic_series(
        figure, dataframe, x_field, "28 Day Test Average",
        "28 Day Average", "#ED7D31", "diamond",
    )
    figure.update_layout(
        title=f"{x_field} vs Concrete Compressive Strength",
        template="plotly_white",
        height=540,
        legend_title="Test Age and Quadratic Curve",
        xaxis={
            "title": f"{x_field} (numeric field measurement)",
            "showgrid": True,
            "tickformat": ".2f",
            "nticks": 10,
        },
        yaxis={
            "title": "Average Concrete Compressive Strength (PSI)",
            "showgrid": True,
            "tickformat": ",.0f",
            "nticks": 10,
            "rangemode": "tozero",
        },
    )
    return figure


# ============================================================
# 7. SESSION AND RESULT HELPERS
# ============================================================


def delete_path(path_value):
    if not path_value:
        return
    try:
        path = Path(path_value)
        if path.exists():
            path.unlink()
    except OSError:
        pass


def clear_result_state():
    delete_path(st.session_state.get("csv_path"))
    delete_path(st.session_state.get("xlsx_path"))
    generation = st.session_state.get("uploader_generation", 0) + 1
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.session_state["uploader_generation"] = generation
    st.cache_data.clear()
    st.cache_resource.clear()
    gc.collect()


# ============================================================
# 8. USER INTERFACE
# ============================================================

if "uploader_generation" not in st.session_state:
    st.session_state["uploader_generation"] = 0
if "csv_path" not in st.session_state:
    st.session_state["csv_path"] = ""
if "xlsx_path" not in st.session_state:
    st.session_state["xlsx_path"] = ""
if "processing_messages" not in st.session_state:
    st.session_state["processing_messages"] = []

st.title("Concrete Test Dashboard")
st.caption(
    "Low-memory build: plain-text PDF extraction, one PDF at a time, "
    "temporary CSV rows, and one selected chart."
)

with st.expander("How to use this tool", expanded=True):
    st.markdown(
        """
        1. Keep the original filenames. `Test` and `Concrete Test` are accepted, and Sample IDs may have 3 or 4 digits.
        2. Upload individual PDFs or a ZIP folder.
        3. Select **Process reports**.
        4. Each PDF is read once, one row is written to a temporary CSV, and the PDF content is released.
        5. After all PDFs are gone, the app creates the formatted Excel workbook once.
        6. Download the workbook and use **Delete files and reset memory** when finished.
        """
    )

uploaded_items = st.file_uploader(
    "Upload concrete-test PDFs or ZIP folders",
    type=["pdf", "zip"],
    accept_multiple_files=True,
    key=f"upload_{st.session_state['uploader_generation']}",
)

if st.button(
    "Process reports",
    type="primary",
    disabled=not uploaded_items,
    use_container_width=True,
):
    delete_path(st.session_state.get("csv_path"))
    delete_path(st.session_state.get("xlsx_path"))
    st.session_state["csv_path"] = ""
    st.session_state["xlsx_path"] = ""
    st.session_state["processing_messages"] = []
    gc.collect()

    total_mb = sum(upload_size_bytes(item) for item in uploaded_items) / 1024 / 1024
    if total_mb > MAX_TOTAL_UPLOAD_MB:
        st.error(
            f"This upload is {total_mb:.1f} MB. "
            f"Please keep each batch below {MAX_TOTAL_UPLOAD_MB} MB."
        )
        st.stop()

    total_count, count_notes = count_pdfs(uploaded_items)
    if total_count > MAX_PDF_FILES:
        st.error(
            f"This upload contains {total_count} PDFs. "
            f"Please upload no more than {MAX_PDF_FILES} at one time."
        )
        st.stop()
    if total_count == 0:
        st.error("No PDF files were found.")
        st.stop()

    run_id = uuid4().hex
    csv_path = RESULT_DIRECTORY / f"Concrete_Test_Log_{run_id}.csv"
    xlsx_path = RESULT_DIRECTORY / f"Concrete_Test_Log_{run_id}.xlsx"
    create_csv(csv_path)

    progress = st.progress(0, text="Processing reports...")
    successful, process_messages = process_uploads(
        uploaded_items,
        csv_path,
        total_count,
        progress,
    )
    progress.empty()

    if successful == 0:
        delete_path(csv_path)
        st.session_state["processing_messages"] = count_notes + process_messages
        st.session_state["uploader_generation"] += 1
        gc.collect()
        st.rerun()

    # PDF processing is finished before openpyxl is used.
    del uploaded_items
    gc.collect()
    csv_to_formatted_xlsx(csv_path, xlsx_path)

    st.session_state["csv_path"] = str(csv_path)
    st.session_state["xlsx_path"] = str(xlsx_path)
    st.session_state["processing_messages"] = count_notes + process_messages
    st.session_state["uploader_generation"] += 1
    gc.collect()
    st.rerun()

csv_path = st.session_state.get("csv_path", "")
xlsx_path = st.session_state.get("xlsx_path", "")
results_ready = (
    csv_path
    and xlsx_path
    and Path(csv_path).exists()
    and Path(xlsx_path).exists()
)

if not results_ready:
    st.info("Upload a smaller batch and select Process reports.")
    if st.session_state.get("processing_messages"):
        with st.expander("Processing details"):
            for name, message in st.session_state["processing_messages"]:
                st.write(f"{name}: {message}")
    st.stop()

master_df = load_result_dataframe(csv_path)

st.success(
    "Processing complete. PDF extraction finished before the Excel workbook "
    "was created, and the uploader was reset."
)

st.subheader("Concrete Test Log")
st.dataframe(
    master_df,
    use_container_width=True,
    hide_index=True,
    column_config={
        "Report Date": st.column_config.DateColumn("Report Date", format="YYYY-MM-DD"),
        "7 Day Test Average": st.column_config.NumberColumn("7 Day Test Average", format="%d"),
        "28 Day Test Average": st.column_config.NumberColumn("28 Day Test Average", format="%d"),
    },
)

st.download_button(
    "Download Concrete Test Log (.xlsx)",
    data=Path(xlsx_path).read_bytes(),
    file_name="Concrete_Test_Log.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)

if st.session_state.get("processing_messages"):
    with st.expander("Processing details"):
        for name, message in st.session_state["processing_messages"]:
            st.write(f"{name}: {message}")

st.subheader("Strength chart")
selected_chart = st.selectbox(
    "Choose the field measurement for the horizontal axis",
    CHART_FIELDS,
)
figure = build_chart(master_df, selected_chart)
st.plotly_chart(figure, use_container_width=True)
del figure
gc.collect()

st.divider()
st.subheader("Memory control")
st.warning(
    "After downloading the workbook, select the button below to delete the "
    "temporary CSV and Excel files and clear the remaining session state."
)

if st.button(
    "Delete files and reset memory",
    use_container_width=True,
):
    clear_result_state()
    st.rerun()
