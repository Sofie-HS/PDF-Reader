import io
import re
import zipfile
from datetime import datetime

import numpy as np
import pandas as pd
import pdfplumber
import plotly.graph_objects as go
import streamlit as st
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


# ============================================================
# APP SETUP
# ============================================================

st.set_page_config(
    page_title="Concrete Test Dashboard",
    page_icon="📊",
    layout="wide",
)

# Accepted filename examples:
# W02229_Concrete Test_Sample 1826_2025-10-02_Report 000082-01_Compressive Strength of Concrete_19330805.pdf
# W02229_Test_Sample 922_2025-03-13_Report 000020-01_Compressive Strength of Concrete_17359783.pdf
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
# GENERAL HELPERS
# ============================================================


def clean_value(value):
    """Remove line breaks and repeated spaces."""
    if value is None:
        return ""

    value = str(value).replace("\r", " ").replace("\n", " ")
    return re.sub(r"\s+", " ", value).strip()


def clean_number(value):
    """Return the first number found in a value."""
    if value is None:
        return None

    text = clean_value(value).replace(",", "")
    match = re.search(r"-?\d+(?:\.\d+)?", text)

    if not match:
        return None

    number = float(match.group(0))
    return int(number) if number.is_integer() else number


def extract_pattern(text, patterns):
    """Return the first captured value matching any supplied pattern."""
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
    """Convert slump values such as 7-1/2 into 7.5."""
    text = clean_value(value)

    mixed_match = re.search(
        r"(\d+)\s*[- ]\s*(\d+)\s*/\s*(\d+)",
        text,
    )

    if mixed_match:
        whole, numerator, denominator = map(int, mixed_match.groups())

        if denominator:
            return round(whole + numerator / denominator, 3)

    fraction_match = re.fullmatch(
        r"\s*(\d+)\s*/\s*(\d+)\s*",
        text,
    )

    if fraction_match:
        numerator, denominator = map(int, fraction_match.groups())

        if denominator:
            return round(numerator / denominator, 3)

    return clean_number(text)


def normalize_header(value):
    """Normalize PDF table headings for matching."""
    value = clean_value(value).lower()
    return re.sub(r"[^a-z0-9]+", " ", value).strip()


# ============================================================
# UPLOAD HELPERS
# ============================================================


class NamedBytesIO(io.BytesIO):
    """An in-memory PDF that retains its original filename."""

    def __init__(self, data, name):
        super().__init__(data)
        self.name = name


def expand_uploads(uploaded_items):
    """Return individual PDF objects from PDFs and ZIP uploads."""
    pdf_files = []
    notes = []

    for item in uploaded_items:
        lower_name = item.name.lower()

        if lower_name.endswith(".pdf"):
            pdf_files.append(item)
            continue

        if not lower_name.endswith(".zip"):
            notes.append((item.name, "Unsupported file type."))
            continue

        try:
            item.seek(0)

            with zipfile.ZipFile(item) as archive:
                pdf_names = [
                    name
                    for name in archive.namelist()
                    if not name.endswith("/")
                    and name.lower().endswith(".pdf")
                ]

                for archived_name in pdf_names:
                    pdf_files.append(
                        NamedBytesIO(
                            archive.read(archived_name),
                            archived_name.split("/")[-1],
                        )
                    )

                notes.append(
                    (item.name, f"Loaded {len(pdf_names)} PDF file(s).")
                )

        except zipfile.BadZipFile:
            notes.append((item.name, "Could not read this ZIP file."))
        finally:
            item.seek(0)

    return pdf_files, notes


# ============================================================
# PDF TEXT AND FIELD EXTRACTION
# ============================================================


def extract_pdf_text(pdf_file):
    """Extract selectable text from every page."""
    pdf_file.seek(0)
    parts = []

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()

            if text:
                parts.append(text)

    pdf_file.seek(0)
    return "\n".join(parts)


def extract_location_details(text):
    """Extract the pour description shown under Location Details."""
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
    """Extract the requested field measurements."""
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


# ============================================================
# BREAK RESULT EXTRACTION
# ============================================================


def find_lab_header_row(table):
    """Find the lab table row containing test age and strength."""
    for row_index, row in enumerate(table):
        combined = " | ".join(
            normalize_header(cell)
            for cell in (row or [])
        )

        if "test age days" in combined and "strength psi" in combined:
            return row_index

    return None


def extract_strengths_by_age(pdf_file):
    """Extract completed strengths from the lab table."""
    pdf_file.seek(0)
    strengths_by_age = {}

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
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

    pdf_file.seek(0)
    return strengths_by_age


def extract_reported_average(text, test_age):
    """Fallback to the average printed in the report text."""
    value = extract_pattern(
        text,
        [
            rf"{test_age}\s*Day\s*-\s*([\d,]+(?:\.\d+)?)",
            rf"{test_age}\s*Day\s*Average\s*:?\s*([\d,]+(?:\.\d+)?)",
            rf"{test_age}[-\s]*Day\s*Average\s*:?\s*([\d,]+(?:\.\d+)?)",
        ],
    )

    return clean_number(value)


def calculate_average(strengths_by_age, text, test_age):
    """Average all completed break strengths for one test age."""
    strengths = strengths_by_age.get(test_age, [])

    if strengths:
        return round(sum(strengths) / len(strengths))

    return extract_reported_average(text, test_age)


def process_pdf(pdf_file):
    """Convert one uploaded PDF into one log record."""
    match = FILENAME_PATTERN.fullmatch(pdf_file.name)

    if not match:
        return None, "Filename does not match the accepted convention."

    info = match.groupdict()

    try:
        datetime.strptime(info["report_date"], "%Y-%m-%d")
    except ValueError:
        return None, "Filename contains an invalid date."

    text = extract_pdf_text(pdf_file)
    measurements = extract_field_measurements(text)
    strengths_by_age = extract_strengths_by_age(pdf_file)

    record = {
        "Report Date": info["report_date"],
        "Sample ID": info["sample_id"],
        "Location Details": extract_location_details(text),
        **measurements,
        "7 Day Test Average": calculate_average(
            strengths_by_age,
            text,
            7,
        ),
        "28 Day Test Average": calculate_average(
            strengths_by_age,
            text,
            28,
        ),
    }

    return record, None


# ============================================================
# EXCEL DOWNLOAD
# ============================================================


def dataframe_to_excel(dataframe):
    """Create the formatted Excel workbook in memory."""
    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl",
        date_format="yyyy-mm-dd",
        datetime_format="yyyy-mm-dd",
    ) as writer:
        dataframe.to_excel(
            writer,
            sheet_name="Concrete Test Log",
            index=False,
        )

        worksheet = writer.sheets["Concrete Test Log"]
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions
        worksheet.sheet_view.showGridLines = False

        header_fill = PatternFill(
            fill_type="solid",
            fgColor="1F4E78",
        )
        header_font = Font(
            color="FFFFFF",
            bold=True,
        )

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
            worksheet.column_dimensions[
                get_column_letter(column_number)
            ].width = width

        for row_number in range(2, worksheet.max_row + 1):
            worksheet[f"A{row_number}"].number_format = "yyyy-mm-dd"

            for column_letter in ["D", "E", "F", "G", "H", "I"]:
                worksheet[
                    f"{column_letter}{row_number}"
                ].number_format = "0.00"

            worksheet[f"J{row_number}"].number_format = "#,##0"
            worksheet[f"K{row_number}"].number_format = "#,##0"

    output.seek(0)
    return output.getvalue()


# ============================================================
# LINEAR CHARTS, RESTORED TO THE ORIGINAL WORKING APPROACH
# ============================================================


def add_linear_series(
    figure,
    dataframe,
    x_field,
    y_field,
    name,
    color,
    symbol,
):
    """Add scatter points and a linear best-fit line."""
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

    if len(plot_data) < 2 or plot_data[x_field].nunique() < 2:
        return

    x_values = plot_data[x_field].astype(float).to_numpy()
    y_values = plot_data[y_field].astype(float).to_numpy()
    slope, intercept = np.polyfit(x_values, y_values, 1)

    x_line = np.linspace(x_values.min(), x_values.max(), 100)
    y_line = slope * x_line + intercept

    figure.add_trace(
        go.Scatter(
            x=x_line,
            y=y_line,
            mode="lines",
            name=f"{name} best fit",
            line={"color": color, "width": 3},
            hoverinfo="skip",
        )
    )


def build_chart(dataframe, x_field):
    """Build one chart for one field measurement."""
    figure = go.Figure()

    add_linear_series(
        figure,
        dataframe,
        x_field,
        "7 Day Test Average",
        "7 Day Average",
        "#4472C4",
        "circle",
    )

    add_linear_series(
        figure,
        dataframe,
        x_field,
        "28 Day Test Average",
        "28 Day Average",
        "#ED7D31",
        "diamond",
    )

    figure.update_layout(
        title=f"{x_field} vs Concrete Compressive Strength",
        template="plotly_white",
        height=540,
        legend_title="Test Age and Best-Fit Line",
        xaxis_title=f"{x_field} (Field Measurement Value)",
        yaxis_title="Concrete Compressive Strength Average (PSI)",
        hovermode="closest",
    )

    return figure


# ============================================================
# STREAMLIT PAGE
# ============================================================

st.title("Concrete Test Dashboard")
st.caption(
    "Restored simple version: upload reports, build one master table, "
    "download Excel, and review separate charts."
)

with st.expander("How to use this tool", expanded=True):
    st.markdown(
        """
        1. Keep the original report filenames.
        2. Upload individual PDFs, multiple PDFs, or a ZIP folder.
        3. The app extracts the report date, Sample ID, Location Details, field measurements, and 7-day/28-day strength averages.
        4. Review the table and download the Excel workbook.
        5. Scroll through the separate strength charts below the table.
        """
    )

uploaded_items = st.file_uploader(
    "Upload concrete test PDFs or ZIP folders",
    type=["pdf", "zip"],
    accept_multiple_files=True,
)

if not uploaded_items:
    st.info("Upload one or more PDFs or ZIP folders to begin.")
    st.stop()

pdf_files, upload_notes = expand_uploads(uploaded_items)

if not pdf_files:
    st.error("No PDF files were found in the upload.")
    st.stop()

records = []
processing_errors = []
progress = st.progress(0, text="Reading concrete test reports...")

for index, pdf_file in enumerate(pdf_files, start=1):
    try:
        record, error = process_pdf(pdf_file)

        if record is not None:
            records.append(record)
        else:
            processing_errors.append((pdf_file.name, error))

    except Exception as exc:
        processing_errors.append(
            (pdf_file.name, f"{type(exc).__name__}: {exc}")
        )

    progress.progress(
        index / len(pdf_files),
        text=f"Processed {index} of {len(pdf_files)} PDFs",
    )

progress.empty()

if not records:
    st.error("No valid report data was extracted.")

    with st.expander("Processing details"):
        for name, message in upload_notes + processing_errors:
            st.write(f"{name}: {message}")

    st.stop()

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

metric_one, metric_two, metric_three = st.columns(3)
metric_one.metric("Reports", len(master_df))
metric_two.metric(
    "7-Day Results",
    int(master_df["7 Day Test Average"].notna().sum()),
)
metric_three.metric(
    "28-Day Results",
    int(master_df["28 Day Test Average"].notna().sum()),
)

st.dataframe(
    master_df,
    use_container_width=True,
    hide_index=True,
    column_config={
        "Report Date": st.column_config.DateColumn(
            "Report Date",
            format="YYYY-MM-DD",
        ),
        "7 Day Test Average": st.column_config.NumberColumn(
            "7 Day Test Average",
            format="%d",
        ),
        "28 Day Test Average": st.column_config.NumberColumn(
            "28 Day Test Average",
            format="%d",
        ),
    },
)

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

st.divider()
st.header("Strength Relationship Charts")
st.caption(
    "Blue shows 7-day averages. Orange shows 28-day averages."
)

for field_name in CHART_FIELDS:
    st.subheader(f"{field_name} vs Strength")
    chart = build_chart(master_df, field_name)
    st.plotly_chart(chart, use_container_width=True)
