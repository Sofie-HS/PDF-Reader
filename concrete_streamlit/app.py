import io
import re
from datetime import datetime

import numpy as np
import pandas as pd
import pdfplumber
import plotly.graph_objects as go
import streamlit as st
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Concrete Test Dashboard", page_icon="📊", layout="wide")

FILENAME_PATTERN = re.compile(
    r"^W02229_Concrete Test_"
    r"Sample (?P<sample_id>\d{4})_"
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


def clean_value(value):
    if value is None:
        return ""
    value = str(value).replace("\r", " ").replace("\n", " ")
    return re.sub(r"\s+", " ", value).strip()


def clean_number(value):
    if value is None:
        return None
    match = re.search(r"-?\d+(?:\.\d+)?", clean_value(value).replace(",", ""))
    if not match:
        return None
    number = float(match.group(0))
    return int(number) if number.is_integer() else number


def extract_pattern(text, patterns):
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE | re.MULTILINE)
        if match:
            return clean_value(match.group(1))
    return ""


def parse_mixed_number(value):
    value = clean_value(value)
    mixed = re.search(r"(\d+)\s*[- ]\s*(\d+)\s*/\s*(\d+)", value)
    if mixed:
        whole, numerator, denominator = map(int, mixed.groups())
        if denominator:
            return round(whole + numerator / denominator, 3)
    fraction = re.fullmatch(r"\s*(\d+)\s*/\s*(\d+)\s*", value)
    if fraction:
        numerator, denominator = map(int, fraction.groups())
        if denominator:
            return round(numerator / denominator, 3)
    return clean_number(value)


def extract_pdf_text(uploaded_file):
    uploaded_file.seek(0)
    parts = []
    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                parts.append(text)
    uploaded_file.seek(0)
    return "\n".join(parts)


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
    air_temp = extract_pattern(text, [
        r"Air\s+Temperature\s*(?:\(F\))?\s*:\s*(-?\d+(?:\.\d+)?)",
        r"Air\s+Temp(?:erature)?\s*(?:\(F\))?\s*:\s*(-?\d+(?:\.\d+)?)",
    ])
    concrete_temp = extract_pattern(text, [
        r"Concrete\s+Temp(?:erature)?\s*(?:\(F\))?\s*:\s*(-?\d+(?:\.\d+)?)",
    ])
    slump = extract_pattern(text, [
        r"Slump\s*(?:\(in\))?\s*:\s*(\d+\s*[- ]\s*\d+\s*/\s*\d+)",
        r"Slump\s*(?:\(in\))?\s*:\s*(\d+\s*/\s*\d+)",
        r"Slump\s*(?:\(in\))?\s*:\s*(\d+(?:\.\d+)?)",
    ])
    air_content = extract_pattern(text, [
        r"Air\s+Content\s*(?:\(%\))?\s*:\s*(-?\d+(?:\.\d+)?)",
    ])
    min_max = extract_pattern(text, [
        r"Min\s*/\s*Max\s+Temp\s*(?:\(F\))?\s*:\s*"
        r"(-?\d+(?:\.\d+)?\s*/\s*-?\d+(?:\.\d+)?)",
    ])
    min_temp = max_temp = None
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


def normalize_header(value):
    return re.sub(r"[^a-z0-9]+", " ", clean_value(value).lower()).strip()


def find_lab_header_row(table):
    for index, row in enumerate(table):
        combined = " | ".join(normalize_header(cell) for cell in (row or []))
        if "test age days" in combined and "strength psi" in combined:
            return index
    return None


def extract_strengths_by_age(uploaded_file):
    uploaded_file.seek(0)
    strengths_by_age = {}
    with pdfplumber.open(uploaded_file) as pdf:
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
                    age = clean_number(row[age_index] if age_index < len(row) else "")
                    strength = clean_number(row[strength_index] if strength_index < len(row) else "")
                    if age is not None and strength is not None:
                        strengths_by_age.setdefault(age, []).append(strength)
    uploaded_file.seek(0)
    return strengths_by_age


def extract_reported_average(text, age):
    value = extract_pattern(text, [
        rf"{age}\s*Day\s*-\s*([\d,]+(?:\.\d+)?)",
        rf"{age}\s*Day\s*Average\s*:?\s*([\d,]+(?:\.\d+)?)",
        rf"{age}[-\s]*Day\s*Average\s*:?\s*([\d,]+(?:\.\d+)?)",
    ])
    return clean_number(value)


def calculate_average(strengths_by_age, text, age):
    strengths = strengths_by_age.get(age, [])
    if strengths:
        return round(sum(strengths) / len(strengths))
    return extract_reported_average(text, age)


def process_pdf(uploaded_file):
    match = FILENAME_PATTERN.fullmatch(uploaded_file.name)
    if not match:
        return None, "Filename does not match the required convention."
    info = match.groupdict()
    try:
        datetime.strptime(info["report_date"], "%Y-%m-%d")
    except ValueError:
        return None, "Filename contains an invalid date."

    text = extract_pdf_text(uploaded_file)
    measurements = extract_field_measurements(text)
    strengths_by_age = extract_strengths_by_age(uploaded_file)

    record = {
        "Report Date": info["report_date"],
        "Sample ID": info["sample_id"],
        "Location Details": extract_location_details(text),
        **measurements,
        "7 Day Test Average": calculate_average(strengths_by_age, text, 7),
        "28 Day Test Average": calculate_average(strengths_by_age, text, 28),
    }
    return record, None


def dataframe_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(
        output,
        engine="openpyxl",
        date_format="yyyy-mm-dd",
        datetime_format="yyyy-mm-dd",
    ) as writer:
        df.to_excel(writer, sheet_name="Concrete Test Log", index=False)
        worksheet = writer.sheets["Concrete Test Log"]
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions
        worksheet.sheet_view.showGridLines = False

        fill = PatternFill(fill_type="solid", fgColor="1F4E78")
        font = Font(color="FFFFFF", bold=True)
        for cell in worksheet[1]:
            cell.fill = fill
            cell.font = font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        worksheet.row_dimensions[1].height = 38

        widths = [15, 14, 52, 12, 16, 12, 14, 12, 12, 22, 24]
        for number, width in enumerate(widths, start=1):
            worksheet.column_dimensions[get_column_letter(number)].width = width

        for row_number in range(2, worksheet.max_row + 1):
            worksheet[f"A{row_number}"].number_format = "yyyy-mm-dd"
            for column in ["D", "E", "F", "G", "H", "I"]:
                worksheet[f"{column}{row_number}"].number_format = "0.00"
            worksheet[f"J{row_number}"].number_format = "#,##0"
            worksheet[f"K{row_number}"].number_format = "#,##0"

    output.seek(0)
    return output.getvalue()


def add_series_with_trendline(fig, data, x_field, y_field, name, color, symbol):
    plot_data = data[[x_field, y_field, "Sample ID", "Location Details"]].dropna()
    if plot_data.empty:
        return

    fig.add_trace(
        go.Scatter(
            x=plot_data[x_field],
            y=plot_data[y_field],
            mode="markers",
            name=name,
            marker=dict(color=color, size=9, symbol=symbol),
            customdata=plot_data[["Sample ID", "Location Details"]],
            hovertemplate=(
                "Sample ID: %{customdata[0]}<br>"
                "Location: %{customdata[1]}<br>"
                f"{x_field}: %{{x}}<br>"
                f"{y_field}: %{{y:,.0f}} PSI<extra></extra>"
            ),
        )
    )

    if len(plot_data) >= 2 and plot_data[x_field].nunique() >= 2:
        x = plot_data[x_field].astype(float).to_numpy()
        y = plot_data[y_field].astype(float).to_numpy()
        slope, intercept = np.polyfit(x, y, 1)
        x_line = np.linspace(x.min(), x.max(), 100)
        y_line = slope * x_line + intercept
        y_pred = slope * x + intercept
        ss_res = np.sum((y - y_pred) ** 2)
        ss_tot = np.sum((y - np.mean(y)) ** 2)
        r_squared = 1 - ss_res / ss_tot if ss_tot else np.nan
        equation = f"y = {slope:.2f}x + {intercept:.1f}"
        r_text = f"R² = {r_squared:.3f}" if not np.isnan(r_squared) else "R² unavailable"

        fig.add_trace(
            go.Scatter(
                x=x_line,
                y=y_line,
                mode="lines",
                name=f"{name} best fit ({equation}; {r_text})",
                line=dict(color=color, width=3),
                hoverinfo="skip",
            )
        )


def build_chart(data, x_field):
    fig = go.Figure()
    add_series_with_trendline(
        fig, data, x_field, "7 Day Test Average", "7 Day Average", "#4472C4", "circle"
    )
    add_series_with_trendline(
        fig, data, x_field, "28 Day Test Average", "28 Day Average", "#ED7D31", "diamond"
    )
    fig.update_layout(
        title=f"{x_field} vs Concrete Compressive Strength",
        xaxis_title=f"{x_field} (Field Measurement Value)",
        yaxis_title="Concrete Compressive Strength Average (PSI)",
        legend_title="Test Age / Best-Fit Line",
        template="plotly_white",
        height=560,
        hovermode="closest",
    )
    return fig


st.title("Concrete Test Dashboard")
st.caption(
    "Upload concrete test PDFs to build the master log, download the Excel workbook, "
    "and review 7-day and 28-day strength relationships."
)

uploaded_files = st.file_uploader(
    "Upload concrete test PDFs",
    type=["pdf"],
    accept_multiple_files=True,
    help="Files must follow the established W02229 concrete-test filename convention.",
)

if not uploaded_files:
    st.info("Upload one or more PDFs to generate the dashboard.")
    st.stop()

records = []
skipped = []
progress = st.progress(0, text="Reading PDFs...")

for index, uploaded_file in enumerate(uploaded_files, start=1):
    try:
        record, error = process_pdf(uploaded_file)
        if record is not None:
            records.append(record)
        else:
            skipped.append((uploaded_file.name, error))
    except Exception as exc:
        skipped.append((uploaded_file.name, f"{type(exc).__name__}: {exc}"))
    progress.progress(index / len(uploaded_files), text=f"Processed {index} of {len(uploaded_files)} PDFs")

progress.empty()

if not records:
    st.error("No valid PDF data was extracted. Check the filename convention and PDF text quality.")
    if skipped:
        with st.expander("Processing details"):
            for name, reason in skipped:
                st.write(f"{name}: {reason}")
    st.stop()

master_df = pd.DataFrame(records, columns=MASTER_COLUMNS)
master_df["Report Date"] = pd.to_datetime(master_df["Report Date"], errors="coerce")
master_df.sort_values(["Report Date", "Sample ID"], inplace=True, na_position="last")
master_df.reset_index(drop=True, inplace=True)

st.subheader("Concrete Test Log")
metric_one, metric_two, metric_three = st.columns(3)
metric_one.metric("Reports", len(master_df))
metric_two.metric("7-Day Results", int(master_df["7 Day Test Average"].notna().sum()))
metric_three.metric("28-Day Results", int(master_df["28 Day Test Average"].notna().sum()))

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

excel_bytes = dataframe_to_excel(master_df)
st.download_button(
    "Download Concrete Test Log (.xlsx)",
    data=excel_bytes,
    file_name="Concrete_Test_Log.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)

if skipped:
    st.warning(f"{len(skipped)} file(s) were skipped. Expand Processing details to review them.")
    with st.expander("Processing details"):
        for name, reason in skipped:
            st.write(f"{name}: {reason}")

st.divider()
st.header("Strength Relationship Charts")
st.caption(
    "Blue circles and blue best-fit lines represent 7-day averages. "
    "Orange diamonds and orange best-fit lines represent 28-day averages."
)

for field in CHART_FIELDS:
    st.subheader(f"{field} vs Strength")
    st.plotly_chart(build_chart(master_df, field), use_container_width=True)
