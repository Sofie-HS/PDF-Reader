"""Streamlit dashboard for a locally generated Concrete_Test_Log.xlsx file."""

import hashlib
import io

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(
    page_title="Concrete Test Dashboard",
    page_icon="📊",
    layout="wide",
)

REQUIRED_COLUMNS = [
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


@st.cache_data(ttl=1800, max_entries=2, show_spinner="Loading dashboard data...")
def load_excel(file_hash, file_bytes):
    """Load the prepared workbook once and reuse it across widget reruns."""
    del file_hash
    dataframe = pd.read_excel(
        io.BytesIO(file_bytes),
        sheet_name="Concrete Test Log",
        engine="openpyxl",
    )
    dataframe["Report Date"] = pd.to_datetime(
        dataframe["Report Date"],
        errors="coerce",
    )
    return dataframe


def padded_range(values, include_zero=False):
    """Return a tight axis range with a small data-dependent margin."""
    numeric = pd.to_numeric(values, errors="coerce").dropna()
    if numeric.empty:
        return None

    minimum = float(numeric.min())
    maximum = float(numeric.max())

    if minimum == maximum:
        margin = max(abs(minimum) * 0.08, 1.0)
    else:
        margin = (maximum - minimum) * 0.08

    low = minimum - margin
    high = maximum + margin

    if include_zero and minimum >= 0 and low < 0:
        low = 0

    return [low, high]


def add_quadratic_series(figure, dataframe, x_field, y_field, name, color, symbol):
    plot_data = dataframe[[
        x_field,
        y_field,
        "Sample ID",
        "Location Details",
    ]].copy()

    plot_data[x_field] = pd.to_numeric(plot_data[x_field], errors="coerce")
    plot_data[y_field] = pd.to_numeric(plot_data[y_field], errors="coerce")
    plot_data.dropna(subset=[x_field, y_field], inplace=True)

    if plot_data.empty:
        return

    figure.add_trace(go.Scatter(
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
    ))

    if len(plot_data) < 3 or plot_data[x_field].nunique() < 3:
        return

    x_values = plot_data[x_field].to_numpy(dtype=float)
    y_values = plot_data[y_field].to_numpy(dtype=float)
    a_value, b_value, c_value = np.polyfit(x_values, y_values, 2)

    x_line = np.linspace(x_values.min(), x_values.max(), 120)
    y_line = a_value * x_line ** 2 + b_value * x_line + c_value

    figure.add_trace(go.Scatter(
        x=x_line,
        y=y_line,
        mode="lines",
        name=f"{name} quadratic fit",
        line={"color": color, "width": 3},
        hoverinfo="skip",
    ))


def build_chart(dataframe, x_field):
    figure = go.Figure()

    add_quadratic_series(
        figure,
        dataframe,
        x_field,
        "7 Day Test Average",
        "7 Day Average",
        "#4472C4",
        "circle",
    )
    add_quadratic_series(
        figure,
        dataframe,
        x_field,
        "28 Day Test Average",
        "28 Day Average",
        "#ED7D31",
        "diamond",
    )

    x_range = padded_range(dataframe[x_field])
    combined_y = pd.concat([
        pd.to_numeric(dataframe["7 Day Test Average"], errors="coerce"),
        pd.to_numeric(dataframe["28 Day Test Average"], errors="coerce"),
    ])
    y_range = padded_range(combined_y, include_zero=False)

    figure.update_layout(
        title=f"{x_field} vs Concrete Compressive Strength",
        template="plotly_white",
        height=560,
        legend_title="Test Age and Quadratic Fit",
        hovermode="closest",
        xaxis={
            "title": f"{x_field} (field measurement)",
            "range": x_range,
            "showgrid": True,
            "tickformat": ".2f",
            "nticks": 10,
        },
        yaxis={
            "title": "Average Concrete Compressive Strength (PSI)",
            "range": y_range,
            "showgrid": True,
            "tickformat": ",.0f",
            "nticks": 10,
        },
    )

    return figure


st.title("Concrete Test Dashboard")
st.caption(
    "The large PDF folder is processed locally. This browser app receives only "
    "the compact Excel result, which keeps the dashboard below cloud memory limits."
)

with st.expander("How this tool works", expanded=True):
    st.markdown(
        """
        **1. Build the workbook outside the browser**
        - Run `processor.py` on the computer that stores the concrete-test PDF folder.
        - The processor accepts both `W02229_Test_...` and `W02229_Concrete Test_...` names.
        - The processor accepts 3-digit and 4-digit Sample IDs.
        - The processor reads one PDF at a time, so the source folder can exceed 1 GB without uploading that folder to Streamlit.

        **2. Upload only the Excel result**
        - Upload `Concrete_Test_Log.xlsx` below.
        - The browser never receives the source PDFs.

        **3. Review the dashboard**
        - The Excel data is not displayed as a large table.
        - Summary counts confirm the workbook loaded.
        - Select one field measurement to show one focused graph at a time.

        **4. Download the workbook**
        - The same prepared Excel file remains available for download.
        """
    )

uploaded_excel = st.file_uploader(
    "Upload the locally generated Concrete_Test_Log.xlsx",
    type=["xlsx"],
    accept_multiple_files=False,
)

if uploaded_excel is None:
    st.info(
        "Run processor.py on the PDF folder first, then upload the generated Excel workbook."
    )
    st.stop()

excel_bytes = uploaded_excel.getvalue()
file_hash = hashlib.sha256(excel_bytes).hexdigest()

try:
    master_df = load_excel(file_hash, excel_bytes)
except Exception as exc:
    st.error("The uploaded workbook could not be opened.")
    st.exception(exc)
    st.stop()

missing_columns = [
    column
    for column in REQUIRED_COLUMNS
    if column not in master_df.columns
]

if missing_columns:
    st.error(
        "The workbook is missing required columns: "
        + ", ".join(missing_columns)
    )
    st.stop()

st.success("Workbook loaded successfully.")

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

st.download_button(
    "Download Concrete Test Log (.xlsx)",
    data=excel_bytes,
    file_name="Concrete_Test_Log.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)

st.divider()
st.header("Strength Relationship Charts")
selected_chart = st.selectbox(
    "Choose the field measurement",
    CHART_FIELDS,
)

chart = build_chart(master_df, selected_chart)
st.plotly_chart(chart, use_container_width=True)
