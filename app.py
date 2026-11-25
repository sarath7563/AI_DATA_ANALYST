from modules.question_answering import answer_question
import io
import re
import html
import zipfile
import sqlite3
import tempfile
import os
import json
import hashlib
import uuid
import sys
import shutil
import importlib.util
from pathlib import Path
from datetime import date, datetime, timezone

# Stage 7.12 — optional PostgreSQL connectivity for cloud migration.
try:
    import psycopg
    from psycopg.rows import dict_row
    PSYCOPG_AVAILABLE = True
except ImportError:
    psycopg = None
    dict_row = None
    PSYCOPG_AVAILABLE = False

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

# Static chart images are used inside PDF and PowerPoint exports.
try:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    MATPLOTLIB_AVAILABLE = True
except ImportError:
    MATPLOTLIB_AVAILABLE = False

# Optional report libraries
try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image as ReportImage
    )
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.enum.text import PP_ALIGN
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    PPTX_AVAILABLE = True
except ImportError:
    PPTX_AVAILABLE = False


# =========================================================
# PAGE CONFIGURATION + DESIGN
# =========================================================
st.set_page_config(
    page_title="DataWise AI | Smart Analytics",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
    html, body, [class*="css"] { font-family: Inter, sans-serif; }
    .stApp { background: #f4f7fb; }
    [data-testid="stSidebar"] { background: #101827; }
    [data-testid="stSidebar"] * { color: #f8fafc; }
    .brand { font-size: 27px; font-weight: 800; letter-spacing: -.7px; }
    .brand-sub { color: #a8b4c7; font-size: 11px; letter-spacing: 1.5px; margin-bottom: 24px; }
    .hero {
        padding: 26px 30px; border-radius: 18px; color: white;
        background: linear-gradient(120deg, #172554, #2563eb);
        margin: 5px 0 20px 0;
    }
    .hero h1 { color: white; font-size: 29px; margin: 0 0 7px 0; }
    .hero p { color: #dbeafe; margin: 0; font-size: 14px; }
    .small-label { color: #64748b; font-size: 12px; font-weight: 700; text-transform: uppercase; }
    .insight-card {
        background: white; border: 1px solid #e2e8f0; border-radius: 13px;
        padding: 16px 18px; margin-bottom: 10px;
    }
    .insight-card h4 { margin: 0 0 7px 0; color: #0f172a; font-size: 15px; }
    .insight-card p { margin: 0; color: #475569; font-size: 13px; }
    .muted { color: #64748b; font-size: 13px; }
    div[data-testid="stMetric"] {
        background: white; border: 1px solid #e2e8f0; border-radius: 13px; padding: 15px;
    }
    div[data-testid="stFileUploader"] {
        background: white; border: 1px dashed #93c5fd; border-radius: 12px; padding: 10px;
    }
    .footer { text-align: center; color: #94a3b8; font-size: 11px; padding: 25px 0 8px; }

    /* Stage 7.2 — polished workspace components */
    .block-container { padding-top: 1.35rem; padding-bottom: 2.5rem; max-width: 1500px; }
    [data-testid="stSidebar"] { border-right: 1px solid rgba(148,163,184,.18); }
    [data-testid="stSidebar"] [data-testid="stRadio"] label { padding: 5px 8px; border-radius: 8px; }
    [data-testid="stSidebar"] [data-testid="stRadio"] label:hover { background: rgba(148,163,184,.14); }
    .hero { position: relative; overflow: hidden; box-shadow: 0 12px 30px rgba(37,99,235,.13); }
    .hero:after { content: ""; position: absolute; width: 230px; height: 230px; right: -65px; top: -105px; border: 1px solid rgba(255,255,255,.18); border-radius: 50%; box-shadow: 0 0 0 28px rgba(255,255,255,.045), 0 0 0 58px rgba(255,255,255,.035); }
    h2, h3 { letter-spacing: -.35px; }
    div[data-testid="stMetric"] { box-shadow: 0 3px 12px rgba(15,23,42,.035); transition: transform .15s ease, box-shadow .15s ease; }
    div[data-testid="stMetric"]:hover { transform: translateY(-2px); box-shadow: 0 8px 20px rgba(15,23,42,.07); }
    div[data-testid="stPlotlyChart"] { background: white; border: 1px solid #e2e8f0; border-radius: 13px; padding: 8px; }
    div[data-testid="stDataFrame"] { border: 1px solid #e2e8f0; border-radius: 12px; overflow: hidden; }
    div[data-testid="stButton"] button, div[data-testid="stDownloadButton"] button { border-radius: 9px; font-weight: 600; }
    div[data-testid="stTabs"] button { font-weight: 600; }
    .section-kicker { color: #2563eb; font-size: 11px; font-weight: 800; letter-spacing: 1.3px; text-transform: uppercase; margin-bottom: 3px; }
    .snapshot-card { background: white; border: 1px solid #e2e8f0; border-radius: 12px; padding: 13px 15px; min-height: 82px; }
    .snapshot-label { color: #64748b; font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: .6px; }
    .snapshot-value { color: #0f172a; font-size: 18px; font-weight: 750; margin-top: 5px; overflow-wrap: anywhere; }
    .status-pill { display: inline-block; border-radius: 20px; padding: 4px 9px; background: #dcfce7; color: #166534; font-size: 11px; font-weight: 700; }
    @media (max-width: 768px) { .block-container { padding-left: 1rem; padding-right: 1rem; } .hero { padding: 20px; } .hero h1 { font-size: 23px; } }
    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# HELPERS
# =========================================================
def read_uploaded_file(uploaded_file, sheet_name=None):
    """Read supported tabular formats into a DataFrame."""
    suffix = Path(uploaded_file.name).suffix.lower()
    uploaded_file.seek(0)
    if suffix == ".csv":
        try:
            return pd.read_csv(uploaded_file)
        except UnicodeDecodeError:
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file, encoding="latin-1")
    if suffix in (".xlsx", ".xls"):
        return pd.read_excel(uploaded_file, sheet_name=sheet_name or 0)
    if suffix == ".json":
        return pd.read_json(uploaded_file)
    raise ValueError("Unsupported file format.")


def load_sqlite_table(database_bytes, table_name):
    """Load one table from an uploaded SQLite database safely."""
    if not database_bytes:
        raise ValueError("The SQLite database file is empty.")
    with tempfile.NamedTemporaryFile(suffix=".sqlite", delete=True) as temp_file:
        temp_file.write(database_bytes)
        temp_file.flush()
        # Open the temporary database in read-only mode.
        uri = Path(temp_file.name).resolve().as_uri() + "?mode=ro"
        connection = sqlite3.connect(uri, uri=True)
        try:
            valid_tables = pd.read_sql_query(
                "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'",
                connection,
            )["name"].tolist()
            if table_name not in valid_tables:
                raise ValueError("The selected table was not found in this database.")
            quoted_name = '"' + table_name.replace('"', '""') + '"'
            return pd.read_sql_query(f"SELECT * FROM {quoted_name}", connection)
        finally:
            connection.close()


def load_sqlalchemy_table(connection_url, table_name):
    """Load a table from a SQLAlchemy-supported database using a read query."""
    try:
        from sqlalchemy import create_engine, inspect, text
    except ImportError as exc:
        raise ImportError(
            "SQLAlchemy is not installed. Run: python -m pip install sqlalchemy"
        ) from exc

    if not connection_url.strip():
        raise ValueError("Enter a database connection URL.")
    engine = create_engine(connection_url.strip(), pool_pre_ping=True)
    try:
        inspector = inspect(engine)
        tables = inspector.get_table_names()
        if table_name not in tables:
            raise ValueError("The selected table was not found in this database.")
        quoted_table = engine.dialect.identifier_preparer.quote(table_name)
        with engine.connect() as connection:
            return pd.read_sql_query(text(f"SELECT * FROM {quoted_table}"), connection)
    finally:
        engine.dispose()


def safe_text(value, max_chars=180):
    text = str(value)
    return text if len(text) <= max_chars else text[: max_chars - 3] + "..."


def numeric_columns(frame):
    return frame.select_dtypes(include=np.number).columns.tolist()


def date_columns(frame):
    found = []
    for col in frame.columns:
        name = str(col).lower()
        if pd.api.types.is_datetime64_any_dtype(frame[col]):
            found.append(col)
        elif any(word in name for word in ("date", "time", "month", "year")):
            converted = pd.to_datetime(frame[col], errors="coerce")
            if len(frame) and converted.notna().mean() >= 0.7:
                found.append(col)
    return list(dict.fromkeys(found))


def clean_column_names(frame):
    """Normalize column names and make duplicate names unique."""
    result = frame.copy()
    cleaned = []
    used = set()

    for i, col in enumerate(result.columns):
        name = re.sub(r"\s+", "_", str(col).strip())
        name = name or f"column_{i + 1}"
        original = name
        counter = 1

        while name.casefold() in used:
            counter += 1
            name = f"{original}_{counter}"

        used.add(name.casefold())
        cleaned.append(name)

    result.columns = cleaned
    return result


def validate_dataframe(frame):
    """Return blocking errors and non-blocking data-quality warnings."""
    errors = []
    warnings = []

    if not isinstance(frame, pd.DataFrame):
        return ["The uploaded file could not be converted into a table."], warnings

    rows, columns = frame.shape

    if columns == 0:
        errors.append("The dataset contains no columns.")
    if rows == 0:
        errors.append("The dataset contains no data rows.")

    if errors:
        return errors, warnings

    empty_columns = frame.columns[frame.isna().all()].tolist()
    if empty_columns:
        preview = ", ".join(map(str, empty_columns[:5]))
        suffix = " ..." if len(empty_columns) > 5 else ""
        warnings.append(
            f"{len(empty_columns)} completely empty column(s): {preview}{suffix}."
        )

    empty_rows = int(frame.isna().all(axis=1).sum())
    if empty_rows:
        warnings.append(f"{empty_rows:,} completely empty row(s) detected.")

    missing = int(frame.isna().sum().sum())
    if missing:
        total_cells = rows * columns
        missing_pct = (missing / total_cells * 100) if total_cells else 0
        warnings.append(
            f"{missing:,} missing cell(s) detected ({missing_pct:.2f}% of the dataset)."
        )

    duplicates = int(frame.duplicated().sum())
    if duplicates:
        warnings.append(f"{duplicates:,} exact duplicate row(s) detected.")

    constant_columns = [
        str(col) for col in frame.columns
        if frame[col].nunique(dropna=True) <= 1
    ]
    if constant_columns:
        preview = ", ".join(constant_columns[:5])
        suffix = " ..." if len(constant_columns) > 5 else ""
        warnings.append(
            f"{len(constant_columns)} column(s) contain only one unique value or no usable values: {preview}{suffix}."
        )

    if rows > 1_000_000:
        warnings.append(
            "This dataset contains more than one million rows; some operations may take longer."
        )
    if columns > 200:
        warnings.append(
            "This dataset contains more than 200 columns; consider removing unnecessary columns."
        )

    return errors, warnings


def display_validation_report(errors, warnings):
    """Show validation results in a compact Streamlit panel."""
    if errors:
        st.error("Dataset validation failed. Please correct the file and upload it again.")
        for message in errors:
            st.write(f"❌ {message}")
    elif warnings:
        st.warning("The dataset loaded, but some data-quality issues need review.")
        for message in warnings:
            st.write(f"⚠️ {message}")
    else:
        st.success("Dataset validation passed. No common data-quality issues were detected.")


def data_quality(frame):
    total_cells = frame.shape[0] * frame.shape[1]
    missing = int(frame.isna().sum().sum())
    duplicates = int(frame.duplicated().sum())
    completeness = 100.0 if total_cells == 0 else (1 - missing / total_cells) * 100
    return missing, duplicates, max(0.0, completeness)


def make_insights(frame):
    """Create transparent, locally calculated business findings."""
    insights = []
    rows, cols = frame.shape
    missing, duplicates, completeness = data_quality(frame)
    insights.append(("Dataset overview", f"The dataset contains {rows:,} rows and {cols:,} columns."))
    insights.append(("Data completeness", f"{completeness:.1f}% of cells are populated; {missing:,} missing values were detected."))
    insights.append(("Duplicate records", f"{duplicates:,} exact duplicate rows were detected." if duplicates else "No exact duplicate rows were detected."))

    nums = numeric_columns(frame)
    dcols = date_columns(frame)
    cats = [c for c in frame.columns if c not in nums and c not in dcols and 1 < frame[c].nunique(dropna=True) <= min(100, max(20, len(frame) // 2))]

    for col in nums[:4]:
        series = pd.to_numeric(frame[col], errors="coerce").dropna()
        if not series.empty:
            insights.append((f"{col}: numeric summary", f"Total {series.sum():,.2f}; average {series.mean():,.2f}; minimum {series.min():,.2f}; maximum {series.max():,.2f}."))
            q1, q3 = series.quantile([0.25, 0.75])
            iqr = q3 - q1
            if iqr > 0:
                outliers = int(((series < q1 - 1.5 * iqr) | (series > q3 + 1.5 * iqr)).sum())
                if outliers:
                    insights.append((f"Potential outliers in {col}", f"{outliers:,} value(s) fall outside the 1.5×IQR screening range. Review them; this rule flags candidates, not confirmed errors."))

    if cats and nums:
        # Prefer a non-date category and a likely business measure when names support it.
        metric = next((c for c in nums if any(k in str(c).lower() for k in ("sales", "revenue", "amount", "profit", "cost", "total"))), nums[0])
        cat = next((c for c in cats if any(k in str(c).lower() for k in ("product", "region", "category", "department", "segment"))), cats[0])
        temp = frame[[cat, metric]].copy()
        temp[metric] = pd.to_numeric(temp[metric], errors="coerce")
        grouped = temp.groupby(cat, dropna=False)[metric].sum(min_count=1).dropna().sort_values(ascending=False)
        if not grouped.empty:
            insights.append((f"Top {cat} by {metric}", f"'{safe_text(grouped.index[0], 70)}' has the highest total {metric} ({grouped.iloc[0]:,.2f}) among {len(grouped):,} groups."))
            if len(grouped) > 1 and grouped.sum() != 0:
                share = grouped.iloc[0] / grouped.sum() * 100
                insights.append(("Top-category contribution", f"The leading {cat} accounts for approximately {share:.1f}% of the summed {metric} across these groups."))

    if dcols and nums:
        dc = dcols[0]
        temp = frame[[dc, nums[0]]].copy()
        temp[dc] = pd.to_datetime(temp[dc], errors="coerce")
        temp[nums[0]] = pd.to_numeric(temp[nums[0]], errors="coerce")
        temp = temp.dropna()
        if not temp.empty:
            totals = temp.groupby(temp[dc].dt.date)[nums[0]].sum().sort_values(ascending=False)
            if not totals.empty:
                insights.append((f"Highest {nums[0]} date", f"{totals.index[0]} has the highest daily total for {nums[0]} ({totals.iloc[0]:,.2f})."))
    return insights[:12]



def detect_iqr_outliers(frame, columns=None):
    """Return a row-level table of potential IQR outliers; flags are review candidates, not errors."""
    columns = columns or numeric_columns(frame)
    records = []
    for col in columns:
        values = pd.to_numeric(frame[col], errors="coerce")
        valid = values.dropna()
        if len(valid) < 4:
            continue
        q1, q3 = valid.quantile([0.25, 0.75])
        iqr = q3 - q1
        if pd.isna(iqr) or iqr == 0:
            continue
        lower, upper = q1 - 1.5 * iqr, q3 + 1.5 * iqr
        mask = values.notna() & ((values < lower) | (values > upper))
        records.append({
            "Column": str(col),
            "Valid values": int(valid.size),
            "Outlier candidates": int(mask.sum()),
            "Candidate %": round(float(mask.sum()) / len(valid) * 100, 2),
            "Lower screening bound": float(lower),
            "Upper screening bound": float(upper),
            "Minimum flagged": float(values[mask].min()) if mask.any() else np.nan,
            "Maximum flagged": float(values[mask].max()) if mask.any() else np.nan,
        })
    return pd.DataFrame(records)


def make_correlation_matrix(frame, columns):
    """Calculate correlations only for numeric columns with more than one valid value."""
    if len(columns) < 2:
        return pd.DataFrame()
    numeric = frame[columns].apply(pd.to_numeric, errors="coerce")
    usable = [col for col in columns if numeric[col].notna().sum() >= 2 and numeric[col].nunique(dropna=True) > 1]
    if len(usable) < 2:
        return pd.DataFrame()
    return numeric[usable].corr().round(3)


def summarize_category_performance(frame, category_col, metric_col, aggregation="Sum"):
    """Build a category performance table with metric, record count, and share."""
    temp = frame[[category_col, metric_col]].copy()
    temp[metric_col] = pd.to_numeric(temp[metric_col], errors="coerce")
    if aggregation == "Average":
        grouped = temp.groupby(category_col, dropna=False)[metric_col].mean()
    elif aggregation == "Maximum":
        grouped = temp.groupby(category_col, dropna=False)[metric_col].max()
    elif aggregation == "Minimum":
        grouped = temp.groupby(category_col, dropna=False)[metric_col].min()
    elif aggregation == "Count":
        grouped = temp.groupby(category_col, dropna=False)[metric_col].count()
    else:
        grouped = temp.groupby(category_col, dropna=False)[metric_col].sum(min_count=1)
    counts = temp.groupby(category_col, dropna=False)[metric_col].count()
    result = pd.DataFrame({str(category_col): grouped.index.astype(str), aggregation: grouped.values, "Records with numeric value": counts.reindex(grouped.index).values})
    total = result[aggregation].sum()
    result["Share of total (%)"] = (result[aggregation] / total * 100).round(2) if aggregation == "Sum" and pd.notna(total) and total != 0 else np.nan
    return result.sort_values(aggregation, ascending=False, na_position="last").reset_index(drop=True)

def choose_default_category(frame, nums):
    candidates = [
        c for c in frame.columns
        if c not in nums and 1 < frame[c].nunique(dropna=True) <= 30
    ]
    return candidates[0] if candidates else frame.columns[0]


def aggregate_data(frame, x_col, y_col, aggregation):
    temp = frame[[x_col, y_col]].copy()
    temp[y_col] = pd.to_numeric(temp[y_col], errors="coerce")
    temp = temp.dropna(subset=[y_col])
    if aggregation == "Count":
        grouped = temp.groupby(x_col, dropna=False).size().rename("Count").reset_index()
        value_col = "Count"
    else:
        functions = {
            "Sum": "sum", "Average": "mean", "Minimum": "min",
            "Maximum": "max", "Count": "count"
        }
        grouped = temp.groupby(x_col, dropna=False)[y_col].agg(functions[aggregation]).reset_index()
        value_col = y_col
    return grouped, value_col


def create_chart(frame, x_col, y_col, chart_type, aggregation="Sum", title=None, top_n=20):
    if chart_type == "Histogram":
        fig = px.histogram(
            frame, x=y_col, title=title or f"Distribution of {y_col}",
            template="plotly_white", color_discrete_sequence=["#2563eb"]
        )
    elif chart_type == "Scatter":
        fig = px.scatter(
            frame, x=x_col, y=y_col, title=title or f"{y_col} vs {x_col}",
            template="plotly_white", color_discrete_sequence=["#2563eb"]
        )
    else:
        grouped, value_col = aggregate_data(frame, x_col, y_col, aggregation)
        if chart_type == "Bar chart":
            grouped = grouped.sort_values(value_col, ascending=False).head(top_n)
            fig = px.bar(
                grouped, x=x_col, y=value_col,
                title=title or f"{aggregation} of {y_col} by {x_col}",
                template="plotly_white", color_discrete_sequence=["#2563eb"]
            )
        elif chart_type == "Line chart":
            fig = px.line(
                grouped, x=x_col, y=value_col, markers=True,
                title=title or f"{aggregation} of {y_col} by {x_col}",
                template="plotly_white", color_discrete_sequence=["#2563eb"]
            )
        elif chart_type == "Pie chart":
            grouped = grouped.sort_values(value_col, ascending=False).head(8)
            fig = px.pie(
                grouped, names=x_col, values=value_col,
                title=title or f"{y_col} by {x_col}",
                template="plotly_white"
            )
        else:
            raise ValueError("Unsupported chart type.")
    fig.update_layout(
        title_x=0.02, margin=dict(l=18, r=18, t=65, b=22),
        font=dict(family="Inter, sans-serif", color="#334155"),
        paper_bgcolor="white", plot_bgcolor="white"
    )
    return fig


def make_report_charts(frame, max_charts=4):
    """Create report-ready PNG charts with sensible date/category detection."""
    if not MATPLOTLIB_AVAILABLE or frame.empty:
        return []
    charts = []
    nums = numeric_columns(frame)
    dcols = date_columns(frame)
    cats = [c for c in frame.columns if c not in nums and c not in dcols and 1 < frame[c].nunique(dropna=True) <= 40]

    def save_figure(fig, title):
        output = io.BytesIO()
        fig.tight_layout()
        fig.savefig(output, format="png", dpi=170, bbox_inches="tight", facecolor="white")
        plt.close(fig)
        output.seek(0)
        charts.append((title, output.getvalue()))

    if cats and nums and len(charts) < max_charts:
        cat = next((c for c in cats if any(k in str(c).lower() for k in ("product", "region", "category", "segment"))), cats[0])
        num = next((c for c in nums if any(k in str(c).lower() for k in ("sales", "revenue", "amount", "profit", "cost"))), nums[0])
        temp = frame[[cat, num]].copy()
        temp[num] = pd.to_numeric(temp[num], errors="coerce")
        grouped = temp.groupby(cat, dropna=False)[num].sum(min_count=1).dropna().sort_values(ascending=False).head(12)
        if not grouped.empty:
            fig, ax = plt.subplots(figsize=(8.5, 4.5))
            grouped.sort_values().plot(kind="barh", ax=ax, color="#2563eb")
            ax.set_title(f"Total {num} by {cat}", loc="left", fontweight="bold")
            ax.set_xlabel(f"Total {num}")
            ax.set_ylabel(str(cat))
            ax.grid(axis="x", alpha=0.22)
            ax.spines[["top", "right", "left"]].set_visible(False)
            save_figure(fig, f"Total {num} by {cat}")

    # A second categorical chart when there is another meaningful category.
    if len(cats) > 1 and nums and len(charts) < max_charts:
        cat = cats[1]
        num = next((c for c in nums if any(k in str(c).lower() for k in ("profit", "sales", "revenue", "amount"))), nums[-1])
        temp = frame[[cat, num]].copy()
        temp[num] = pd.to_numeric(temp[num], errors="coerce")
        grouped = temp.groupby(cat, dropna=False)[num].sum(min_count=1).dropna().sort_values(ascending=False).head(12)
        if not grouped.empty:
            fig, ax = plt.subplots(figsize=(8.5, 4.5))
            grouped.sort_values().plot(kind="barh", ax=ax, color="#0f766e")
            ax.set_title(f"Total {num} by {cat}", loc="left", fontweight="bold")
            ax.set_xlabel(f"Total {num}")
            ax.set_ylabel(str(cat))
            ax.grid(axis="x", alpha=0.22)
            ax.spines[["top", "right", "left"]].set_visible(False)
            save_figure(fig, f"Total {num} by {cat}")

    if dcols and nums and len(charts) < max_charts:
        dc = dcols[0]
        num = next((c for c in nums if any(k in str(c).lower() for k in ("sales", "revenue", "amount", "profit"))), nums[0])
        temp = frame[[dc, num]].copy()
        temp[dc] = pd.to_datetime(temp[dc], errors="coerce")
        temp[num] = pd.to_numeric(temp[num], errors="coerce")
        temp = temp.dropna()
        if not temp.empty:
            temp["_period"] = temp[dc].dt.strftime("%Y-%m-%d") if temp[dc].nunique() <= 35 else temp[dc].dt.to_period("M").astype(str)
            grouped = temp.groupby("_period", as_index=False)[num].sum()
            fig, ax = plt.subplots(figsize=(8.5, 4.5))
            ax.plot(grouped["_period"], grouped[num], marker="o", linewidth=2, color="#7c3aed")
            ax.set_title(f"{num} over time", loc="left", fontweight="bold")
            ax.set_xlabel("Date / period")
            ax.set_ylabel(f"Total {num}")
            ax.tick_params(axis="x", rotation=35)
            ax.grid(alpha=0.22)
            ax.spines[["top", "right"]].set_visible(False)
            save_figure(fig, f"{num} over time")

    if len(nums) >= 2 and len(charts) < max_charts:
        x_col = next((c for c in nums if "sales" in str(c).lower() or "revenue" in str(c).lower()), nums[0])
        y_col = next((c for c in nums if c != x_col and ("profit" in str(c).lower() or "margin" in str(c).lower())), next((c for c in nums if c != x_col), None))
        if y_col:
            temp = frame[[x_col, y_col]].apply(pd.to_numeric, errors="coerce").dropna()
            if len(temp) > 1:
                fig, ax = plt.subplots(figsize=(8.5, 4.5))
                ax.scatter(temp[x_col], temp[y_col], alpha=0.72, color="#ea580c", edgecolors="white", linewidths=0.4)
                ax.set_title(f"Relationship: {x_col} vs {y_col}", loc="left", fontweight="bold")
                ax.set_xlabel(str(x_col)); ax.set_ylabel(str(y_col))
                ax.grid(alpha=0.22)
                ax.spines[["top", "right"]].set_visible(False)
                save_figure(fig, f"Relationship: {x_col} vs {y_col}")
    return charts


def make_pdf(frame, filename, insight_items):
    if not REPORTLAB_AVAILABLE:
        raise RuntimeError("ReportLab is not installed. Run: python -m pip install reportlab")
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=16*mm, leftMargin=16*mm,
                           topMargin=15*mm, bottomMargin=15*mm, title="DataWise AI Business Report",
                           author="DataWise AI")
    styles = getSampleStyleSheet()
    navy = colors.HexColor("#173b7a")
    blue = colors.HexColor("#2563eb")
    styles.add(ParagraphStyle(name="DWTitle", parent=styles["Title"], textColor=navy,
                              fontName="Helvetica-Bold", fontSize=25, leading=30, alignment=0, spaceAfter=8))
    styles.add(ParagraphStyle(name="DWSub", parent=styles["Normal"], textColor=colors.HexColor("#64748b"),
                              fontSize=10, leading=14, spaceAfter=8))
    styles.add(ParagraphStyle(name="DWSection", parent=styles["Heading2"], textColor=navy,
                              fontName="Helvetica-Bold", fontSize=15, leading=19, spaceBefore=12, spaceAfter=7))
    styles.add(ParagraphStyle(name="DWBody", parent=styles["BodyText"], fontSize=9, leading=13, spaceAfter=5))
    story = [Paragraph("DataWise AI", styles["DWTitle"]),
             Paragraph("BUSINESS DATA ANALYSIS REPORT", styles["Heading3"]),
             Paragraph(f"Source file: {html.escape(safe_text(filename))} &nbsp;&nbsp; | &nbsp;&nbsp; Generated: {date.today().isoformat()}", styles["DWSub"]),
             Spacer(1, 3*mm)]
    missing, duplicates, completeness = data_quality(frame)
    nums = numeric_columns(frame)
    kpis = [("RECORDS", f"{len(frame):,}"), ("COLUMNS", f"{len(frame.columns):,}"),
            ("MISSING CELLS", f"{missing:,}"), ("COMPLETENESS", f"{completeness:.1f}%")]
    kpi_table = Table([[item[0] for item in kpis], [item[1] for item in kpis]], colWidths=[41*mm]*4)
    kpi_table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), colors.HexColor("#eff6ff")),
        ("TEXTCOLOR", (0,0), (-1,-1), navy), ("FONTNAME", (0,0), (-1,-1), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,-1), 10), ("ALIGN", (0,0), (-1,-1), "CENTER"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"), ("BOX", (0,0), (-1,-1), 0.6, colors.HexColor("#bfdbfe")),
        ("INNERGRID", (0,0), (-1,-1), 0.4, colors.HexColor("#bfdbfe")), ("TOPPADDING", (0,0), (-1,-1), 12),
        ("BOTTOMPADDING", (0,0), (-1,-1), 12)]))
    story.extend([kpi_table, Paragraph("Executive summary", styles["DWSection"]),
                  Paragraph(f"This report summarizes {len(frame):,} records across {len(frame.columns):,} columns. It found {missing:,} missing cells and {duplicates:,} exact duplicate rows. Results below are calculated from the selected dataset.", styles["DWBody"])])

    if nums:
        story.append(Paragraph("Numeric performance summary", styles["DWSection"]))
        numeric_table = [["Measure", "Total", "Average", "Minimum", "Maximum"]]
        for col in nums[:10]:
            series = pd.to_numeric(frame[col], errors="coerce").dropna()
            if not series.empty:
                numeric_table.append([safe_text(col, 28), f"{series.sum():,.2f}", f"{series.mean():,.2f}", f"{series.min():,.2f}", f"{series.max():,.2f}"])
        t = Table(numeric_table, repeatRows=1, colWidths=[42*mm, 32*mm, 32*mm, 28*mm, 28*mm])
        t.setStyle(TableStyle([("BACKGROUND", (0,0), (-1,0), navy), ("TEXTCOLOR", (0,0), (-1,0), colors.white),
                               ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"), ("FONTSIZE", (0,0), (-1,-1), 8),
                               ("GRID", (0,0), (-1,-1), 0.35, colors.HexColor("#d8e0eb")),
                               ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#f8fafc")]),
                               ("VALIGN", (0,0), (-1,-1), "MIDDLE"), ("PADDING", (0,0), (-1,-1), 5)]))
        story.append(t)

    story.append(Paragraph("Key findings", styles["DWSection"]))
    for title, body in insight_items:
        story.append(Paragraph(f"<b>{html.escape(safe_text(title))}</b> — {html.escape(safe_text(body))}", styles["DWBody"]))

    report_charts = make_report_charts(frame)
    if report_charts:
        story.append(Paragraph("Charts and patterns", styles["DWSection"]))
        for title, image_bytes in report_charts:
            story.append(Paragraph(html.escape(safe_text(title)), styles["Heading3"]))
            story.append(ReportImage(io.BytesIO(image_bytes), width=165*mm, height=82*mm, kind="proportional"))
            story.append(Spacer(1, 2*mm))

    story.append(Paragraph("Data quality and column profile", styles["DWSection"]))
    profile = [["Column", "Type", "Missing", "Unique values"]]
    for col in frame.columns[:35]:
        profile.append([safe_text(col, 35), str(frame[col].dtype), f"{int(frame[col].isna().sum()):,}", f"{int(frame[col].nunique(dropna=True)):,}"])
    pt = Table(profile, repeatRows=1, colWidths=[58*mm, 34*mm, 25*mm, 35*mm])
    pt.setStyle(TableStyle([("BACKGROUND", (0,0), (-1,0), navy), ("TEXTCOLOR", (0,0), (-1,0), colors.white),
                            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"), ("GRID", (0,0), (-1,-1), 0.3, colors.HexColor("#d8e0eb")),
                            ("FONTSIZE", (0,0), (-1,-1), 8), ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#f8fafc")]),
                            ("PADDING", (0,0), (-1,-1), 5)]))
    story.extend([pt, Spacer(1, 5*mm), Paragraph("Method note: findings are generated by local Python rules and descriptive statistics. Potential outliers are screening flags, not confirmed errors; this report does not claim causal explanations.", styles["DWSub"])])

    def footer(canvas, document):
        canvas.saveState()
        canvas.setStrokeColor(colors.HexColor("#dbe4f0"))
        canvas.line(16*mm, 11*mm, A4[0]-16*mm, 11*mm)
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(colors.HexColor("#64748b"))
        canvas.drawString(16*mm, 7*mm, "DataWise AI • Generated analysis")
        canvas.drawRightString(A4[0]-16*mm, 7*mm, f"Page {document.page}")
        canvas.restoreState()
    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    return buffer.getvalue()



def make_custom_pdf_report(frame, filename, title, company, report_date, reporting_period,
                          sections, selected_metrics=None, category_col=None, category_metric=None,
                          automated_findings=None, theme_name="Professional Blue"):
    """Build a customized, branded PDF from the selected Report Generator sections."""
    if not REPORTLAB_AVAILABLE:
        raise RuntimeError("ReportLab is not installed. Run: python -m pip install reportlab")

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=A4, rightMargin=16*mm, leftMargin=16*mm,
        topMargin=18*mm, bottomMargin=18*mm,
        title=safe_text(title or "DataWise AI Business Report"),
        author=safe_text(company or "DataWise AI"),
    )
    styles = getSampleStyleSheet()
    theme_palettes = {
        "Professional Blue": ("#173b7a", "#2563eb", "#eff6ff"),
        "Executive Green": ("#14532d", "#16a34a", "#f0fdf4"),
        "Modern Purple": ("#4c1d95", "#7c3aed", "#f5f3ff"),
        "Minimal Gray": ("#334155", "#64748b", "#f1f5f9"),
    }
    navy_hex, blue_hex, pale_hex = theme_palettes.get(theme_name, theme_palettes["Professional Blue"])
    navy = colors.HexColor(navy_hex)
    blue = colors.HexColor(blue_hex)
    pale = colors.HexColor(pale_hex)
    styles.add(ParagraphStyle(name="CustomDWTitle", parent=styles["Title"], textColor=navy,
                              fontName="Helvetica-Bold", fontSize=24, leading=29,
                              alignment=0, spaceAfter=8))
    styles.add(ParagraphStyle(name="CustomDWSubtitle", parent=styles["Normal"],
                              textColor=colors.HexColor("#64748b"), fontSize=9, leading=13,
                              spaceAfter=5))
    styles.add(ParagraphStyle(name="CustomDWSection", parent=styles["Heading2"],
                              textColor=navy, fontName="Helvetica-Bold", fontSize=14,
                              leading=18, spaceBefore=12, spaceAfter=7))
    styles.add(ParagraphStyle(name="CustomDWBody", parent=styles["BodyText"],
                              fontSize=9, leading=13, spaceAfter=5))
    story = [
        Paragraph("DATAWISE AI", styles["CustomDWSubtitle"]),
        Paragraph(html.escape(safe_text(title or "Business Data Analysis Report")), styles["CustomDWTitle"]),
    ]
    if company and str(company).strip():
        story.append(Paragraph(f"<b>{html.escape(safe_text(company))}</b>", styles["Heading3"]))
    metadata = [f"Report date: {html.escape(str(report_date))}",
                f"Source file: {html.escape(safe_text(filename))}"]
    if reporting_period and str(reporting_period).strip():
        metadata.append(f"Reporting period: {html.escape(safe_text(reporting_period))}")
    story.append(Paragraph(" &nbsp; | &nbsp; ".join(metadata), styles["CustomDWSubtitle"]))
    story.append(Spacer(1, 4*mm))

    missing, duplicates, completeness = data_quality(frame)
    if "overview" in sections:
        story.append(Paragraph("Dataset overview and quality", styles["CustomDWSection"]))
        summary = [
            ["Records", "Columns", "Missing cells", "Completeness"],
            [f"{len(frame):,}", f"{len(frame.columns):,}", f"{missing:,}", f"{completeness:.1f}%"],
        ]
        table = Table(summary, colWidths=[41*mm]*4)
        table.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), navy), ("TEXTCOLOR", (0,0), (-1,0), colors.white),
            ("BACKGROUND", (0,1), (-1,1), pale), ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("FONTNAME", (0,1), (-1,1), "Helvetica-Bold"), ("FONTSIZE", (0,0), (-1,-1), 9),
            ("ALIGN", (0,0), (-1,-1), "CENTER"), ("GRID", (0,0), (-1,-1), .4, colors.HexColor("#bfdbfe")),
            ("TOPPADDING", (0,0), (-1,-1), 9), ("BOTTOMPADDING", (0,0), (-1,-1), 9),
        ]))
        story.extend([table, Spacer(1, 2*mm),
                      Paragraph(f"The filtered dataset contains {len(frame):,} records and {len(frame.columns):,} columns, with {duplicates:,} exact duplicate rows.", styles["CustomDWBody"])])

    nums = numeric_columns(frame)
    if "kpis" in sections:
        metrics = [c for c in (selected_metrics or nums[:4]) if c in nums]
        if not metrics:
            metrics = nums[:4]
        story.append(Paragraph("Numeric KPIs", styles["CustomDWSection"]))
        if metrics:
            rows = [["Metric", "Total", "Average", "Minimum", "Maximum", "Valid values"]]
            for col in metrics[:12]:
                vals = pd.to_numeric(frame[col], errors="coerce").dropna()
                if not vals.empty:
                    rows.append([safe_text(col, 24), f"{vals.sum():,.2f}", f"{vals.mean():,.2f}",
                                 f"{vals.min():,.2f}", f"{vals.max():,.2f}", f"{len(vals):,}"])
            if len(rows) > 1:
                widths = [34*mm, 29*mm, 27*mm, 25*mm, 25*mm, 22*mm]
                t = Table(rows, repeatRows=1, colWidths=widths)
                t.setStyle(TableStyle([
                    ("BACKGROUND", (0,0), (-1,0), navy), ("TEXTCOLOR", (0,0), (-1,0), colors.white),
                    ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"), ("FONTSIZE", (0,0), (-1,-1), 7),
                    ("GRID", (0,0), (-1,-1), .35, colors.HexColor("#d8e0eb")),
                    ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#f8fafc")]),
                    ("VALIGN", (0,0), (-1,-1), "MIDDLE"), ("PADDING", (0,0), (-1,-1), 5),
                ]))
                story.append(t)
            else:
                story.append(Paragraph("No valid numeric values were found for the selected metrics.", styles["CustomDWBody"]))
        else:
            story.append(Paragraph("No numeric columns are available in this dataset.", styles["CustomDWBody"]))

    if "insights" in sections:
        story.append(Paragraph("Key findings", styles["CustomDWSection"]))
        for heading, body in make_insights(frame):
            story.append(Paragraph(f"<b>{html.escape(safe_text(heading))}</b> — {html.escape(safe_text(body))}", styles["CustomDWBody"]))
        for heading, body in (automated_findings or []):
            story.append(Paragraph(f"<b>{html.escape(safe_text(heading))}</b> — {html.escape(safe_text(body))}", styles["CustomDWBody"]))

    if "categories" in sections:
        cat = category_col if category_col in frame.columns else None
        metric = category_metric if category_metric in nums else (nums[0] if nums else None)
        if cat and metric:
            temp = frame[[cat, metric]].copy()
            temp[metric] = pd.to_numeric(temp[metric], errors="coerce")
            grouped = temp.groupby(cat, dropna=False)[metric].agg(Total="sum", Average="mean", Records="count").reset_index()
            grouped = grouped.sort_values("Total", ascending=False, na_position="last").head(15)
            story.append(Paragraph(f"Category performance: {safe_text(cat)} by {safe_text(metric)}", styles["CustomDWSection"]))
            rows = [[safe_text(cat, 25), "Total", "Average", "Records"]]
            for _, r in grouped.iterrows():
                rows.append([safe_text(r[cat], 35),
                             f"{r['Total']:,.2f}" if pd.notna(r["Total"]) else "—",
                             f"{r['Average']:,.2f}" if pd.notna(r["Average"]) else "—",
                             f"{int(r['Records']):,}"])
            t = Table(rows, repeatRows=1, colWidths=[65*mm, 38*mm, 38*mm, 22*mm])
            t.setStyle(TableStyle([
                ("BACKGROUND", (0,0), (-1,0), navy), ("TEXTCOLOR", (0,0), (-1,0), colors.white),
                ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"), ("FONTSIZE", (0,0), (-1,-1), 8),
                ("GRID", (0,0), (-1,-1), .35, colors.HexColor("#d8e0eb")),
                ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#f8fafc")]),
                ("PADDING", (0,0), (-1,-1), 5),
            ]))
            story.append(t)
        else:
            story.append(Paragraph("Category analysis was not included because a suitable category and numeric measure were not selected.", styles["CustomDWBody"]))

    if "charts" in sections:
        charts = make_report_charts(frame, max_charts=4) if MATPLOTLIB_AVAILABLE else []
        if charts:
            story.append(Paragraph("Charts and trends", styles["CustomDWSection"]))
            for chart_title, image_bytes in charts:
                story.append(Paragraph(html.escape(safe_text(chart_title)), styles["Heading3"]))
                story.append(ReportImage(io.BytesIO(image_bytes), width=165*mm, height=82*mm, kind="proportional"))
                story.append(Spacer(1, 2*mm))
        else:
            story.append(Paragraph("No suitable charts were available, or Matplotlib is not installed.", styles["CustomDWBody"]))

    if "recommendations" in sections:
        story.append(Paragraph("Recommendations and next steps", styles["CustomDWSection"]))
        recommendations = []
        if missing:
            recommendations.append("Review missing values and decide whether to fill, correct, or retain them.")
        if duplicates:
            recommendations.append("Review exact duplicate rows before using this dataset for final decisions.")
        if nums:
            recommendations.append("Validate extreme values in important numeric measures against the original source.")
        if not missing and not duplicates:
            recommendations.append("No missing cells or exact duplicate rows were detected; continue with business-specific validation.")
        for item in recommendations:
            story.append(Paragraph("• " + html.escape(item), styles["CustomDWBody"]))
        story.append(Paragraph("Findings are based on descriptive statistics and local rules. They are not proof of causation; potential outliers should be reviewed before action.", styles["CustomDWSubtitle"]))

    if "overview" in sections:
        story.append(Paragraph("Column profile", styles["CustomDWSection"]))
        profile = [["Column", "Data type", "Missing", "Unique"]]
        for col in frame.columns[:35]:
            profile.append([safe_text(col, 32), safe_text(frame[col].dtype), f"{int(frame[col].isna().sum()):,}",
                            f"{int(frame[col].nunique(dropna=True)):,}"])
        t = Table(profile, repeatRows=1, colWidths=[62*mm, 35*mm, 25*mm, 25*mm])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), navy), ("TEXTCOLOR", (0,0), (-1,0), colors.white),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"), ("FONTSIZE", (0,0), (-1,-1), 8),
            ("GRID", (0,0), (-1,-1), .3, colors.HexColor("#d8e0eb")),
            ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#f8fafc")]),
            ("PADDING", (0,0), (-1,-1), 5),
        ]))
        story.append(t)

    def footer(canvas, document):
        canvas.saveState()
        canvas.setStrokeColor(colors.HexColor("#dbe4f0"))
        canvas.line(16*mm, 12*mm, A4[0]-16*mm, 12*mm)
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(colors.HexColor("#64748b"))
        canvas.drawString(16*mm, 7*mm, "DataWise AI • Business report")
        canvas.drawRightString(A4[0]-16*mm, 7*mm, f"Page {document.page}")
        canvas.restoreState()

    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    return buffer.getvalue()


def make_custom_pptx_report(frame, source_name, report_title, company_name, report_date, reporting_period,
                            sections, selected_metrics=None, category_col=None, category_metric=None,
                            automated_findings=None, theme_name="Professional Blue"):
    """Create a customized, branded PowerPoint presentation from the selected dataset."""
    if not PPTX_AVAILABLE:
        raise RuntimeError("python-pptx is not installed. Run: python -m pip install python-pptx")
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    theme_palettes = {
        "Professional Blue": ((23, 59, 122), (37, 99, 235), (239, 246, 255)),
        "Executive Green": ((20, 83, 45), (22, 163, 74), (240, 253, 244)),
        "Modern Purple": ((76, 29, 149), (124, 58, 237), (245, 243, 255)),
        "Minimal Gray": ((51, 65, 85), (100, 116, 139), (241, 245, 249)),
    }
    navy_rgb, blue_rgb, pale_rgb = theme_palettes.get(theme_name, theme_palettes["Professional Blue"])
    navy = RGBColor(*navy_rgb)
    blue = RGBColor(*blue_rgb)
    pale = RGBColor(*pale_rgb)
    dark = RGBColor(31, 41, 55)
    muted = RGBColor(100, 116, 139)
    white = RGBColor(255, 255, 255)

    def textbox(slide, x, y, w, h, text, size=16, bold=False, color=dark, font='Aptos'):
        shape = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = shape.text_frame
        tf.clear(); tf.word_wrap = True
        for i, line in enumerate(str(text).split('\n')):
            para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            para.text = line
            para.font.name = font; para.font.size = Pt(size); para.font.bold = bold; para.font.color.rgb = color
        return shape

    def base_slide(title_text, subtitle=None):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.background.fill.solid(); slide.background.fill.fore_color.rgb = white
        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(0.13))
        bar.fill.solid(); bar.fill.fore_color.rgb = blue; bar.line.fill.background()
        textbox(slide, 0.65, 0.35, 12, 0.55, title_text, 27, True, navy, 'Aptos Display')
        if subtitle:
            textbox(slide, 0.68, 0.95, 11.8, 0.35, subtitle, 10, False, muted)
        textbox(slide, 0.65, 7.12, 11.8, 0.2, f"DATAWISE AI  •  {report_date}  •  {len(prs.slides)}", 8, False, muted)
        return slide

    def add_table(slide, rows, x=0.65, y=1.55, w=12.0, h=4.9, font_size=11):
        if not rows: return
        ncols = max(len(r) for r in rows)
        table = slide.shapes.add_table(len(rows), ncols, Inches(x), Inches(y), Inches(w), Inches(h)).table
        for ri, row in enumerate(rows):
            for ci in range(ncols):
                val = row[ci] if ci < len(row) else ''
                cell = table.cell(ri, ci); cell.text = safe_text(val, 90)
                cell.margin_left = Inches(0.08); cell.margin_right = Inches(0.06)
                cell.margin_top = Inches(0.04); cell.margin_bottom = Inches(0.03)
                cell.fill.solid(); cell.fill.fore_color.rgb = navy if ri == 0 else (pale if ri % 2 == 0 else white)
                for para in cell.text_frame.paragraphs:
                    para.font.name = 'Aptos'; para.font.size = Pt(font_size if ri else font_size+1)
                    para.font.bold = ri == 0; para.font.color.rgb = white if ri == 0 else dark

    # Cover slide
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid(); slide.background.fill.fore_color.rgb = navy
    band = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(5.85), prs.slide_width, Inches(1.65))
    band.fill.solid(); band.fill.fore_color.rgb = blue; band.line.fill.background()
    textbox(slide, 0.9, 1.15, 11.5, 0.8, company_name.strip() or 'BUSINESS ANALYTICS', 17, True, RGBColor(191,219,254))
    textbox(slide, 0.9, 2.0, 11.5, 1.5, report_title or 'Business Data Analysis Report', 34, True, white, 'Aptos Display')
    meta = [f"Source dataset: {source_name}", f"Report date: {report_date}"]
    if reporting_period.strip(): meta.append(f"Reporting period: {reporting_period.strip()}")
    textbox(slide, 0.95, 3.75, 11.3, 1.2, '\n'.join(meta), 15, False, RGBColor(226,232,240))
    textbox(slide, 0.95, 6.35, 11.4, 0.5, 'DATAWISE AI  |  SMART ANALYTICS WORKSPACE', 12, True, white)

    missing, duplicates, completeness = data_quality(frame)
    nums = numeric_columns(frame)
    insights = make_insights(frame)
    if 'overview' in sections:
        slide = base_slide('Executive summary', 'A high-level overview of the selected dataset')
        cards = [('RECORDS', f'{len(frame):,}'), ('COLUMNS', f'{len(frame.columns):,}'), ('MISSING CELLS', f'{missing:,}'), ('COMPLETENESS', f'{completeness:.1f}%')]
        for i, (label, value) in enumerate(cards):
            x = 0.7 + i*3.15
            sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(1.65), Inches(2.8), Inches(1.3))
            sh.fill.solid(); sh.fill.fore_color.rgb = pale; sh.line.color.rgb = RGBColor(191,219,254)
            textbox(slide, x+0.15, 1.88, 2.5, 0.25, label, 10, True, muted)
            textbox(slide, x+0.15, 2.2, 2.5, 0.5, value, 25, True, navy)
        textbox(slide, 0.9, 3.45, 11.3, 1.65,
                f'The dataset contains {len(frame):,} records and {len(frame.columns):,} columns. It has {missing:,} missing cells, {duplicates:,} exact duplicate rows, and {completeness:.1f}% cell completeness. Findings are descriptive and should be validated against business context.', 18, False, dark)

    if 'kpis' in sections:
        metrics = [m for m in (selected_metrics or nums[:4]) if m in nums]
        if metrics:
            slide = base_slide('Key performance indicators', 'Descriptive statistics calculated from the current filtered dataset')
            rows = [['Metric', 'Total', 'Average', 'Median', 'Minimum', 'Maximum']]
            for col in metrics[:12]:
                vals = pd.to_numeric(frame[col], errors='coerce').dropna()
                if vals.empty: rows.append([col, '—', '—', '—', '—', '—'])
                else: rows.append([col, f'{vals.sum():,.2f}', f'{vals.mean():,.2f}', f'{vals.median():,.2f}', f'{vals.min():,.2f}', f'{vals.max():,.2f}'])
            add_table(slide, rows, h=min(4.8, 0.55*len(rows)+0.3), font_size=10)

    if 'insights' in sections:
        slide = base_slide('Key findings', 'Rule-based findings generated from the selected data')
        lines = []
        for head, body in insights[:7]: lines.append(f'• {head}: {body}')
        for head, body in (automated_findings or [])[:5]: lines.append(f'• {head}: {body}')
        if not lines: lines = ['No findings could be generated from the current dataset.']
        textbox(slide, 0.8, 1.5, 11.7, 5.25, '\n\n'.join(lines), 15, False, dark)

    if 'categories' in sections and category_col in frame.columns and category_metric in frame.columns:
        temp = frame[[category_col, category_metric]].copy()
        temp[category_metric] = pd.to_numeric(temp[category_metric], errors='coerce')
        summary = temp.groupby(category_col, dropna=False)[category_metric].sum(min_count=1).dropna().sort_values(ascending=False).head(12)
        if not summary.empty:
            slide = base_slide('Category performance', f'Top categories by total {category_metric}')
            rows = [[str(category_col), f'Total {category_metric}']]
            rows.extend([[safe_text(k, 65), f'{v:,.2f}'] for k, v in summary.items()])
            add_table(slide, rows, h=min(5.2, 0.45*len(rows)+0.3), font_size=11)

    if 'charts' in sections:
        charts = make_report_charts(frame, max_charts=3) if MATPLOTLIB_AVAILABLE else []
        for chart_title, chart_bytes in charts[:3]:
            slide = base_slide(safe_text(chart_title, 100), 'Chart generated from the selected dataset')
            try:
                slide.shapes.add_picture(io.BytesIO(chart_bytes), Inches(0.8), Inches(1.45), width=Inches(11.7), height=Inches(5.3))
            except Exception:
                textbox(slide, 0.8, 1.6, 11.5, 1.0, 'This chart could not be embedded in the presentation.', 14)

    if 'recommendations' in sections:
        slide = base_slide('Recommendations and next steps', 'Practical checks based on data quality and summary statistics')
        recs = []
        if missing: recs.append(f'Review {missing:,} missing cell(s) and decide whether to correct, fill, or retain them.')
        if duplicates: recs.append(f'Inspect {duplicates:,} exact duplicate row(s) before final reporting.')
        if nums: recs.append('Validate the largest and smallest values in key numeric measures against the source system.')
        if not missing and not duplicates: recs.append('Basic missing-value and exact-duplicate checks found no issues; continue with domain-specific validation.')
        recs.append('Treat trends and category comparisons as descriptive; they do not establish causation.')
        textbox(slide, 0.9, 1.55, 11.5, 4.8, '\n\n'.join(f'• {r}' for r in recs), 17, False, dark)

    if len(prs.slides) == 1:
        slide = base_slide('Report contents', 'No optional sections were selected')
        textbox(slide, 0.9, 1.6, 11.5, 1.2, 'Select at least one report section in the app to populate the presentation.', 18)
    output = io.BytesIO()
    prs.save(output)
    return output.getvalue()

def make_pptx(frame, filename, insight_items):
    if not PPTX_AVAILABLE:
        raise RuntimeError("python-pptx is not installed. Run: python -m pip install python-pptx")
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    navy = RGBColor(23, 59, 122)
    blue = RGBColor(37, 99, 235)
    light = RGBColor(239, 246, 255)
    dark = RGBColor(30, 41, 59)
    muted = RGBColor(100, 116, 139)

    def set_bg(slide, rgb=RGBColor(255, 255, 255)):
        fill = slide.background.fill
        fill.solid(); fill.fore_color.rgb = rgb

    def title(slide, text, subtitle=None):
        box = slide.shapes.add_textbox(Inches(0.65), Inches(0.35), Inches(12), Inches(0.65))
        p = box.text_frame.paragraphs[0]; p.text = text; p.font.name = "Aptos Display"; p.font.size = Pt(28); p.font.bold = True; p.font.color.rgb = navy
        if subtitle:
            sb = slide.shapes.add_textbox(Inches(0.68), Inches(1.02), Inches(11.8), Inches(0.35))
            p2 = sb.text_frame.paragraphs[0]; p2.text = subtitle; p2.font.size = Pt(11); p2.font.color.rgb = muted

    def footer(slide, number):
        box = slide.shapes.add_textbox(Inches(0.65), Inches(7.12), Inches(12), Inches(0.2))
        p = box.text_frame.paragraphs[0]; p.text = f"DATAWISE AI  •  {date.today().isoformat()}                                             {number}"
        p.font.size = Pt(8); p.font.color.rgb = muted

    # Cover slide
    slide = prs.slides.add_slide(prs.slide_layouts[6]); set_bg(slide, navy)
    band = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(5.85), prs.slide_width, Inches(1.65))
    band.fill.solid(); band.fill.fore_color.rgb = blue; band.line.fill.background()
    box = slide.shapes.add_textbox(Inches(0.9), Inches(1.4), Inches(11.4), Inches(1.1))
    p = box.text_frame.paragraphs[0]; p.text = "DATAWISE AI"; p.font.name = "Aptos Display"; p.font.size = Pt(42); p.font.bold = True; p.font.color.rgb = RGBColor(255,255,255)
    box2 = slide.shapes.add_textbox(Inches(0.95), Inches(2.65), Inches(11.2), Inches(1.0))
    p = box2.text_frame.paragraphs[0]; p.text = "Sales & Data Performance Report"; p.font.size = Pt(25); p.font.color.rgb = RGBColor(219,234,254)
    box3 = slide.shapes.add_textbox(Inches(0.95), Inches(3.65), Inches(11.2), Inches(0.8))
    p = box3.text_frame.paragraphs[0]; p.text = f"Source: {safe_text(filename, 120)}\nGenerated: {date.today().isoformat()}"; p.font.size = Pt(14); p.font.color.rgb = RGBColor(255,255,255)

    missing, duplicates, completeness = data_quality(frame)
    # Executive summary with KPI cards
    slide = prs.slides.add_slide(prs.slide_layouts[6]); set_bg(slide); title(slide, "Executive summary", "A high-level view of the selected dataset")
    nums = numeric_columns(frame)
    kpis = [("RECORDS", f"{len(frame):,}"), ("COLUMNS", f"{len(frame.columns):,}"), ("MISSING CELLS", f"{missing:,}"), ("COMPLETENESS", f"{completeness:.1f}%")]
    for i, (label, value) in enumerate(kpis):
        x = Inches(0.7 + i*3.15)
        shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, Inches(1.65), Inches(2.8), Inches(1.35))
        shape.fill.solid(); shape.fill.fore_color.rgb = light; shape.line.color.rgb = RGBColor(191,219,254)
        tb = slide.shapes.add_textbox(x+Inches(0.15), Inches(1.88), Inches(2.5), Inches(0.25))
        p = tb.text_frame.paragraphs[0]; p.text = label; p.font.size = Pt(10); p.font.bold = True; p.font.color.rgb = muted
        tb2 = slide.shapes.add_textbox(x+Inches(0.15), Inches(2.18), Inches(2.5), Inches(0.55))
        p = tb2.text_frame.paragraphs[0]; p.text = value; p.font.size = Pt(26); p.font.bold = True; p.font.color.rgb = navy
    summary = slide.shapes.add_textbox(Inches(0.9), Inches(3.55), Inches(11.4), Inches(2.2))
    tf = summary.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.text = f"This dataset contains {len(frame):,} records and {len(frame.columns):,} columns. It has {missing:,} missing cells and {duplicates:,} exact duplicate rows. The following slides summarize numeric measures, category performance, time patterns where available, and data quality."; p.font.size = Pt(19); p.font.color.rgb = dark
    footer(slide, len(prs.slides))

    # Numeric measures slide
    if nums:
        slide = prs.slides.add_slide(prs.slide_layouts[6]); set_bg(slide); title(slide, "Numeric performance", "Descriptive statistics calculated from the selected data")
        rows = min(len(nums), 10) + 1
        table = slide.shapes.add_table(rows, 5, Inches(0.7), Inches(1.55), Inches(12), Inches(4.7)).table
        headers = ["Measure", "Total", "Average", "Minimum", "Maximum"]
        for j, h in enumerate(headers):
            table.cell(0,j).text = h
        for i, col in enumerate(nums[:10], start=1):
            series = pd.to_numeric(frame[col], errors="coerce").dropna()
            vals = [str(col), f"{series.sum():,.2f}", f"{series.mean():,.2f}", f"{series.min():,.2f}", f"{series.max():,.2f}"] if not series.empty else [str(col), "—", "—", "—", "—"]
            for j, val in enumerate(vals): table.cell(i,j).text = val
        for j in range(5):
            cell = table.cell(0,j); cell.fill.solid(); cell.fill.fore_color.rgb = navy
            for para in cell.text_frame.paragraphs:
                para.font.bold = True; para.font.color.rgb = RGBColor(255,255,255); para.font.size = Pt(12)
        for i in range(1, rows):
            for j in range(5):
                cell = table.cell(i,j); cell.fill.solid(); cell.fill.fore_color.rgb = light if i % 2 == 0 else RGBColor(255,255,255)
                for para in cell.text_frame.paragraphs: para.font.size = Pt(11); para.font.color.rgb = dark
        footer(slide, len(prs.slides))

    # Insight slides (group up to two per slide)
    for index in range(0, len(insight_items), 2):
        slide = prs.slides.add_slide(prs.slide_layouts[6]); set_bg(slide); title(slide, "Business findings", "Rule-based observations from the data; review context before making decisions")
        for j, (head, body) in enumerate(insight_items[index:index+2]):
            y = 1.55 + j*2.25
            card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(y), Inches(11.7), Inches(1.75))
            card.fill.solid(); card.fill.fore_color.rgb = light; card.line.color.rgb = RGBColor(219,234,254)
            tb = slide.shapes.add_textbox(Inches(1.05), Inches(y+0.18), Inches(11.1), Inches(0.4))
            p = tb.text_frame.paragraphs[0]; p.text = safe_text(head, 100); p.font.size = Pt(17); p.font.bold = True; p.font.color.rgb = navy
            tb2 = slide.shapes.add_textbox(Inches(1.05), Inches(y+0.65), Inches(11.1), Inches(0.85))
            p = tb2.text_frame.paragraphs[0]; p.text = safe_text(body, 400); p.font.size = Pt(14); p.font.color.rgb = dark
        footer(slide, len(prs.slides))

    # Chart slides
    for chart_title, image_bytes in make_report_charts(frame):
        slide = prs.slides.add_slide(prs.slide_layouts[6]); set_bg(slide); title(slide, chart_title, "Chart generated from the current filtered dataset")
        slide.shapes.add_picture(io.BytesIO(image_bytes), Inches(1.05), Inches(1.45), width=Inches(11.2), height=Inches(5.3))
        footer(slide, len(prs.slides))

    # Data quality + profile slide
    slide = prs.slides.add_slide(prs.slide_layouts[6]); set_bg(slide); title(slide, "Data quality & column profile", "Use this profile to identify areas that may need cleaning")
    profile_cols = list(frame.columns[:14]); rows = len(profile_cols)+1
    table = slide.shapes.add_table(rows, 4, Inches(0.7), Inches(1.45), Inches(12), Inches(5.2)).table
    for j, h in enumerate(["Column", "Data type", "Missing", "Unique values"]): table.cell(0,j).text = h
    for i, col in enumerate(profile_cols, start=1):
        values = [safe_text(col, 50), str(frame[col].dtype), str(int(frame[col].isna().sum())), str(int(frame[col].nunique(dropna=True)))]
        for j, val in enumerate(values): table.cell(i,j).text = val
    for i in range(rows):
        for j in range(4):
            cell = table.cell(i,j); cell.fill.solid(); cell.fill.fore_color.rgb = navy if i == 0 else (light if i % 2 == 0 else RGBColor(255,255,255))
            for para in cell.text_frame.paragraphs:
                para.font.size = Pt(10); para.font.color.rgb = RGBColor(255,255,255) if i == 0 else dark
                if i == 0: para.font.bold = True
    footer(slide, len(prs.slides))

    slide = prs.slides.add_slide(prs.slide_layouts[6]); set_bg(slide); title(slide, "Conclusion & next steps", "Recommendations are prompts for further analysis, not guarantees")
    recs = [f"Review the {missing:,} missing cell(s) and {duplicates:,} duplicate row(s), if present.",
            "Compare product and region performance to identify areas worth investigating, where those fields exist.",
            "Validate potential outliers against source systems before correcting or removing values.",
            "Use additional context and business knowledge before acting on statistical patterns."]
    for i, textline in enumerate(recs):
        tb = slide.shapes.add_textbox(Inches(1.0), Inches(1.55+i*1.05), Inches(11.2), Inches(0.7))
        p = tb.text_frame.paragraphs[0]; p.text = f"{i+1}.  {textline}"; p.font.size = Pt(17); p.font.color.rgb = dark
    footer(slide, len(prs.slides))

    output = io.BytesIO(); prs.save(output); return output.getvalue()


def make_excel_report(frame, filename, insight_items):
    """Build a formatted multi-sheet Excel report with native charts."""
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
        from openpyxl.chart import BarChart, Reference
        from openpyxl.chart.label import DataLabelList
        from openpyxl.utils import get_column_letter
    except ImportError as exc:
        raise RuntimeError("openpyxl is required for Excel export. Install it with pip.") from exc

    wb = Workbook()
    ws = wb.active; ws.title = "Executive Summary"
    navy = "173B7A"; blue = "2563EB"; pale = "EFF6FF"; light = "F8FAFC"; white = "FFFFFF"; dark = "1E293B"
    thin_blue = Side(style="thin", color="BFDBFE")
    missing, duplicates, completeness = data_quality(frame)
    ws.merge_cells("A1:F1"); ws["A1"] = "DATAWISE AI | BUSINESS REPORT"
    ws["A1"].font = Font(name="Aptos Display", size=18, bold=True, color=white)
    ws["A1"].fill = PatternFill("solid", fgColor=navy); ws["A1"].alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 34
    ws.merge_cells("A2:F2"); ws["A2"] = f"Source: {filename} | Generated: {date.today().isoformat()}"
    ws["A2"].font = Font(italic=True, color="64748B", size=10)
    kpis = [("Records", len(frame)), ("Columns", len(frame.columns)), ("Missing cells", missing), ("Duplicate rows", duplicates), ("Completeness", f"{completeness:.1f}%")]
    for idx, (label, value) in enumerate(kpis, start=1):
        col = idx
        ws.cell(4,col,label); ws.cell(5,col,value)
        ws.cell(4,col).font = Font(bold=True, color=navy, size=10)
        ws.cell(4,col).fill = PatternFill("solid", fgColor=pale)
        ws.cell(5,col).font = Font(bold=True, color=blue, size=16)
        ws.cell(5,col).fill = PatternFill("solid", fgColor=pale)
        ws.cell(4,col).alignment = Alignment(horizontal="center")
        ws.cell(5,col).alignment = Alignment(horizontal="center")
        ws.cell(4,col).border = ws.cell(5,col).border = Border(left=thin_blue,right=thin_blue,top=thin_blue,bottom=thin_blue)
    row = 8; ws.cell(row,1,"Key findings")
    ws.cell(row,1).font = Font(bold=True, size=14, color=navy)
    for i, (head, body) in enumerate(insight_items, start=row+1):
        ws.cell(i,1,head); ws.cell(i,1).font = Font(bold=True, color=navy)
        ws.cell(i,2,body); ws.cell(i,2).alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=i,start_column=2,end_row=i,end_column=6)
        ws.row_dimensions[i].height = 30
    for col in range(1,7): ws.column_dimensions[get_column_letter(col)].width = [25,24,18,18,18,18][col-1]
    ws.freeze_panes = "A4"; ws.sheet_view.showGridLines = False

    data_ws = wb.create_sheet("Data")
    for j, col in enumerate(frame.columns, start=1):
        cell = data_ws.cell(1,j,str(col)); cell.font = Font(bold=True,color=white); cell.fill = PatternFill("solid",fgColor=navy); cell.alignment = Alignment(wrap_text=True)
    max_rows = min(len(frame), 50000)
    for i, rowvals in enumerate(frame.head(max_rows).itertuples(index=False, name=None), start=2):
        for j, value in enumerate(rowvals, start=1):
            if pd.isna(value): value = None
            elif isinstance(value, (np.integer,)): value = int(value)
            elif isinstance(value, (np.floating,)): value = float(value)
            elif isinstance(value, (pd.Timestamp, datetime, date)): value = value.isoformat()
            elif isinstance(value, np.bool_): value = bool(value)
            elif not isinstance(value, (str, int, float, bool, type(None))): value = str(value)
            data_ws.cell(i,j,value)
    data_ws.freeze_panes = "A2"; data_ws.auto_filter.ref = data_ws.dimensions; data_ws.sheet_view.showGridLines = False
    for j, col in enumerate(frame.columns, start=1):
        max_len = max([len(str(col))] + [len(str(v)) for v in frame[col].head(200).fillna("").astype(str).tolist()])
        data_ws.column_dimensions[get_column_letter(j)].width = min(max(max_len+2, 12), 28)

    profile_ws = wb.create_sheet("Column Profile")
    profile_headers = ["Column", "Data type", "Missing values", "Unique values", "Missing %"]
    profile_ws.append(profile_headers)
    for cell in profile_ws[1]: cell.font=Font(bold=True,color=white); cell.fill=PatternFill("solid",fgColor=navy)
    for col in frame.columns:
        miss = int(frame[col].isna().sum())
        profile_ws.append([str(col),str(frame[col].dtype),miss,int(frame[col].nunique(dropna=True)), round(miss/len(frame)*100,2) if len(frame) else 0])
    profile_ws.freeze_panes="A2"; profile_ws.auto_filter.ref=profile_ws.dimensions; profile_ws.sheet_view.showGridLines=False
    for col, width in zip("ABCDE", [30,20,18,18,15]): profile_ws.column_dimensions[col].width=width
    for rowvals in profile_ws.iter_rows(min_row=2):
        for cell in rowvals:
            if cell.row % 2 == 0: cell.fill=PatternFill("solid",fgColor=light)

    insights_ws = wb.create_sheet("Insights")
    insights_ws.append(["Finding", "Observation"])
    for cell in insights_ws[1]: cell.font=Font(bold=True,color=white); cell.fill=PatternFill("solid",fgColor=navy)
    for item in insight_items: insights_ws.append(list(item))
    insights_ws.column_dimensions["A"].width=34; insights_ws.column_dimensions["B"].width=100
    insights_ws.freeze_panes="A2"; insights_ws.sheet_view.showGridLines=False
    for rowvals in insights_ws.iter_rows(min_row=2):
        rowvals[1].alignment=Alignment(wrap_text=True,vertical="top")
        insights_ws.row_dimensions[rowvals[0].row].height=32

    # Add group summary sheet and embedded native bar chart when a usable category/measure exists.
    nums = numeric_columns(frame); dcols = date_columns(frame)
    cats = [c for c in frame.columns if c not in nums and c not in dcols and 1 < frame[c].nunique(dropna=True) <= 100]
    if cats and nums:
        cat = next((c for c in cats if any(k in str(c).lower() for k in ("product", "region", "category", "segment"))), cats[0])
        metric = next((c for c in nums if any(k in str(c).lower() for k in ("sales", "revenue", "amount", "profit"))), nums[0])
        temp = frame[[cat,metric]].copy(); temp[metric]=pd.to_numeric(temp[metric],errors="coerce")
        grouped = temp.groupby(cat,dropna=False)[metric].sum(min_count=1).dropna().sort_values(ascending=False).head(30)
        summary_ws = wb.create_sheet("Category Analysis")
        summary_ws.append([str(cat), f"Total {metric}"])
        for cell in summary_ws[1]: cell.font=Font(bold=True,color=white); cell.fill=PatternFill("solid",fgColor=navy)
        for key, val in grouped.items(): summary_ws.append([str(key),float(val)])
        summary_ws.column_dimensions["A"].width=28; summary_ws.column_dimensions["B"].width=22
        if len(grouped):
            chart = BarChart(); chart.type = "bar"; chart.style = 10; chart.title = f"Total {metric} by {cat}"; chart.y_axis.title = str(cat); chart.x_axis.title = f"Total {metric}"
            chart.height = 8; chart.width = 16
            data = Reference(summary_ws, min_col=2, min_row=1, max_row=len(grouped)+1)
            cats_ref = Reference(summary_ws, min_col=1, min_row=2, max_row=len(grouped)+1)
            chart.add_data(data, titles_from_data=True); chart.set_categories(cats_ref); chart.legend = None
            summary_ws.add_chart(chart, "D2")

    out = io.BytesIO(); wb.save(out); return out.getvalue()


# =========================================================
# STAGE 7.13 — DATABASE BACKEND (POSTGRESQL / SQLITE)
# PostgreSQL is selected whenever DATABASE_URL is configured.
# SQLite remains available for local fallback and uploaded SQLite files.
# =========================================================
APP_DB_PATH = Path(os.environ.get("DATAWISE_DB_PATH", str(Path(__file__).resolve().parent / "datawise_app.db")))
APP_DB_BACKEND = "postgresql" if os.environ.get("DATABASE_URL", "").strip() else "sqlite"


class _HybridRow(dict):
    """Dictionary-style row that also supports SQLite-style numeric indexes."""
    def __getitem__(self, key):
        if isinstance(key, int):
            return list(self.values())[key]
        return super().__getitem__(key)


def _hybrid_row_factory(cursor):
    # PostgreSQL DDL and INSERT/UPDATE statements without RETURNING have no
    # result-set description. psycopg still initializes the row factory for
    # these statements, so handle description=None safely.
    description = cursor.description or ()
    names = [column.name for column in description]
    return lambda values: _HybridRow(zip(names, values))


class _PostgresConnectionAdapter:
    """Small compatibility layer for existing app DB helper functions."""
    def __init__(self, raw_connection):
        self.raw = raw_connection

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        try:
            if exc_type is None:
                self.raw.commit()
            else:
                self.raw.rollback()
        finally:
            self.raw.close()
        return False

    def execute(self, sql, params=()):
        # Existing application helper SQL uses SQLite's '?' placeholders.
        sql = sql.replace(" COLLATE NOCASE", "")
        sql = sql.replace("?", "%s")
        return self.raw.execute(sql, params or ())


def app_db_connect():
    if APP_DB_BACKEND == "postgresql":
        if not PSYCOPG_AVAILABLE:
            raise RuntimeError("PostgreSQL is selected but psycopg is not installed.")
        raw = psycopg.connect(
            os.environ["DATABASE_URL"],
            connect_timeout=12,
            row_factory=_hybrid_row_factory,
        )
        return _PostgresConnectionAdapter(raw)

    APP_DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(APP_DB_PATH), timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute("PRAGMA busy_timeout = 30000")
    return conn


def _postgres_init_database():
    statements = [
        """CREATE TABLE IF NOT EXISTS dataset_history (
            id BIGSERIAL PRIMARY KEY,
            dataset_signature TEXT NOT NULL UNIQUE,
            session_id TEXT NOT NULL,
            owner_user_id BIGINT,
            source_name TEXT NOT NULL,
            source_type TEXT NOT NULL,
            row_count INTEGER NOT NULL,
            column_count INTEGER NOT NULL,
            columns_json TEXT NOT NULL,
            uploaded_at TEXT NOT NULL
        )""",
        """CREATE TABLE IF NOT EXISTS report_history (
            id BIGSERIAL PRIMARY KEY,
            session_id TEXT NOT NULL,
            owner_user_id BIGINT,
            dataset_signature TEXT,
            report_title TEXT NOT NULL,
            source_name TEXT NOT NULL,
            formats_json TEXT NOT NULL,
            row_count INTEGER NOT NULL,
            column_count INTEGER NOT NULL,
            created_at TEXT NOT NULL
        )""",
        """CREATE TABLE IF NOT EXISTS activity_log (
            id BIGSERIAL PRIMARY KEY,
            session_id TEXT NOT NULL,
            event_type TEXT NOT NULL,
            description TEXT NOT NULL,
            created_at TEXT NOT NULL
        )""",
        """CREATE TABLE IF NOT EXISTS app_settings (
            setting_key TEXT PRIMARY KEY,
            setting_value TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )""",
        """CREATE TABLE IF NOT EXISTS app_users (
            id BIGSERIAL PRIMARY KEY,
            email TEXT NOT NULL UNIQUE,
            display_name TEXT NOT NULL,
            password_hash TEXT NOT NULL,
            password_salt TEXT NOT NULL,
            role TEXT NOT NULL CHECK(role IN ('admin','analyst','viewer')),
            is_active BOOLEAN NOT NULL DEFAULT TRUE,
            created_at TEXT NOT NULL,
            last_login_at TEXT
        )""",
        "CREATE INDEX IF NOT EXISTS idx_app_users_email ON app_users(LOWER(email))",
        "CREATE INDEX IF NOT EXISTS idx_dataset_uploaded ON dataset_history(uploaded_at DESC)",
        "CREATE INDEX IF NOT EXISTS idx_report_created ON report_history(created_at DESC)",
        "CREATE INDEX IF NOT EXISTS idx_activity_created ON activity_log(created_at DESC)",
        "CREATE INDEX IF NOT EXISTS idx_dataset_owner ON dataset_history(owner_user_id, uploaded_at DESC)",
        "CREATE INDEX IF NOT EXISTS idx_report_owner ON report_history(owner_user_id, created_at DESC)",
    ]
    with app_db_connect() as conn:
        for statement in statements:
            conn.execute(statement)
        # The migration copied explicit IDs. Advance SERIAL sequences so new rows
        # do not collide with the imported IDs.
        for table_name in ("app_users", "dataset_history", "report_history", "activity_log"):
            sequence_row = conn.execute(
                "SELECT pg_get_serial_sequence(?, 'id') AS sequence_name", (table_name,)
            ).fetchone()
            if sequence_row and sequence_row["sequence_name"]:
                conn.execute(
                    f"SELECT setval(%s::regclass, COALESCE((SELECT MAX(id) FROM {table_name}), 1), "
                    f"EXISTS (SELECT 1 FROM {table_name}))",
                    (sequence_row["sequence_name"],),
                )


def init_app_database():
    if APP_DB_BACKEND == "postgresql":
        _postgres_init_database()
        return

    with app_db_connect() as conn:
        conn.execute("PRAGMA journal_mode = WAL")
        conn.executescript("""
        CREATE TABLE IF NOT EXISTS dataset_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            dataset_signature TEXT NOT NULL UNIQUE,
            session_id TEXT NOT NULL,
            owner_user_id INTEGER,
            source_name TEXT NOT NULL,
            source_type TEXT NOT NULL,
            row_count INTEGER NOT NULL,
            column_count INTEGER NOT NULL,
            columns_json TEXT NOT NULL,
            uploaded_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS report_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            session_id TEXT NOT NULL,
            owner_user_id INTEGER,
            dataset_signature TEXT,
            report_title TEXT NOT NULL,
            source_name TEXT NOT NULL,
            formats_json TEXT NOT NULL,
            row_count INTEGER NOT NULL,
            column_count INTEGER NOT NULL,
            created_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS activity_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            session_id TEXT NOT NULL,
            event_type TEXT NOT NULL,
            description TEXT NOT NULL,
            created_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS app_settings (
            setting_key TEXT PRIMARY KEY,
            setting_value TEXT NOT NULL,
            updated_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS app_users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            email TEXT NOT NULL UNIQUE COLLATE NOCASE,
            display_name TEXT NOT NULL,
            password_hash TEXT NOT NULL,
            password_salt TEXT NOT NULL,
            role TEXT NOT NULL CHECK(role IN ('admin','analyst','viewer')),
            is_active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL,
            last_login_at TEXT
        );
        CREATE INDEX IF NOT EXISTS idx_app_users_email ON app_users(email);
        CREATE INDEX IF NOT EXISTS idx_dataset_uploaded ON dataset_history(uploaded_at DESC);
        CREATE INDEX IF NOT EXISTS idx_report_created ON report_history(created_at DESC);
        CREATE INDEX IF NOT EXISTS idx_activity_created ON activity_log(created_at DESC);
        """)
        dataset_cols = {row[1] for row in conn.execute("PRAGMA table_info(dataset_history)")}
        report_cols = {row[1] for row in conn.execute("PRAGMA table_info(report_history)")}
        if "owner_user_id" not in dataset_cols:
            conn.execute("ALTER TABLE dataset_history ADD COLUMN owner_user_id INTEGER")
        if "owner_user_id" not in report_cols:
            conn.execute("ALTER TABLE report_history ADD COLUMN owner_user_id INTEGER")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_dataset_owner ON dataset_history(owner_user_id, uploaded_at DESC)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_report_owner ON report_history(owner_user_id, created_at DESC)")


def _query_dataframe(conn, query, params=()):
    cursor = conn.execute(query, params)
    rows = cursor.fetchall()
    columns = [item.name for item in cursor.description] if APP_DB_BACKEND == "postgresql" else [item[0] for item in cursor.description]
    return pd.DataFrame([dict(row) for row in rows], columns=columns)


def db_log_event(session_id, event_type, description):
    with app_db_connect() as conn:
        conn.execute(
            "INSERT INTO activity_log(session_id,event_type,description,created_at) VALUES(?,?,?,?)",
            (session_id, event_type[:60], description[:500], datetime.now(timezone.utc).isoformat(timespec="seconds")),
        )


def db_record_dataset(session_id, signature, source_name, source_type, frame, owner_user_id=None):
    now = datetime.now(timezone.utc).isoformat(timespec="seconds")
    with app_db_connect() as conn:
        cur = conn.execute("""
            INSERT INTO dataset_history
            (dataset_signature,session_id,owner_user_id,source_name,source_type,row_count,column_count,columns_json,uploaded_at)
            VALUES(?,?,?,?,?,?,?,?,?)
            ON CONFLICT(dataset_signature) DO NOTHING
        """, (signature, session_id, owner_user_id, source_name, source_type, int(len(frame)), int(len(frame.columns)),
              json.dumps([str(c) for c in frame.columns], ensure_ascii=False), now))
        inserted = cur.rowcount > 0
    if inserted:
        db_log_event(session_id, "dataset_loaded", f"Loaded {source_name} ({len(frame):,} rows × {len(frame.columns):,} columns)")


def db_record_report(session_id, signature, title, source_name, formats, frame, owner_user_id=None):
    now = datetime.now(timezone.utc).isoformat(timespec="seconds")
    with app_db_connect() as conn:
        conn.execute("""
            INSERT INTO report_history
            (session_id,owner_user_id,dataset_signature,report_title,source_name,formats_json,row_count,column_count,created_at)
            VALUES(?,?,?,?,?,?,?,?,?)
        """, (session_id, owner_user_id, signature, title, source_name, json.dumps(formats), int(len(frame)), int(len(frame.columns)), now))
    db_log_event(session_id, "report_generated", f"Generated {', '.join(formats)} report: {title}")


def db_read_table(table_name, limit=500):
    allowed = {"dataset_history", "report_history", "activity_log"}
    if table_name not in allowed:
        raise ValueError("Unknown history table")
    with app_db_connect() as conn:
        return _query_dataframe(conn, f"SELECT * FROM {table_name} ORDER BY id DESC LIMIT ?", (int(limit),))


def db_read_user_history(table_name, user_id, is_admin=False, limit=500):
    allowed = {"dataset_history", "report_history"}
    if table_name not in allowed:
        raise ValueError("Unsupported user history table")
    with app_db_connect() as conn:
        if is_admin:
            return _query_dataframe(conn, f"SELECT * FROM {table_name} ORDER BY id DESC LIMIT ?", (int(limit),))
        return _query_dataframe(
            conn, f"SELECT * FROM {table_name} WHERE owner_user_id=? ORDER BY id DESC LIMIT ?",
            (int(user_id), int(limit))
        )


def db_get_setting(key, default=""):
    with app_db_connect() as conn:
        row = conn.execute("SELECT setting_value FROM app_settings WHERE setting_key=?", (key,)).fetchone()
    return row[0] if row else default


def db_set_setting(key, value):
    with app_db_connect() as conn:
        conn.execute("""INSERT INTO app_settings(setting_key,setting_value,updated_at) VALUES(?,?,?)
        ON CONFLICT(setting_key) DO UPDATE SET setting_value=excluded.setting_value, updated_at=excluded.updated_at""",
                     (key, str(value), datetime.now(timezone.utc).isoformat(timespec="seconds")))


def db_validate_backup(file_bytes):
    with tempfile.NamedTemporaryFile(suffix=".db", delete=True) as tmp:
        tmp.write(file_bytes)
        tmp.flush()
        check = sqlite3.connect(f"file:{Path(tmp.name).resolve()}?mode=ro", uri=True)
        try:
            integrity = check.execute("PRAGMA integrity_check").fetchone()[0]
            tables = {r[0] for r in check.execute("SELECT name FROM sqlite_master WHERE type='table'")}
            required = {"dataset_history", "report_history", "activity_log", "app_settings", "app_users"}
            return integrity == "ok" and required.issubset(tables)
        finally:
            check.close()


# =========================================================
# STAGE 7.7 — AUTHENTICATION, USERS & ROLE-BASED ACCESS
# Passwords are stored as PBKDF2 hashes (never plaintext).
# =========================================================
import hmac
import secrets

AUTH_ITERATIONS = 310_000
VALID_ROLES = {"admin", "analyst", "viewer"}


def auth_hash_password(password, salt_hex=None):
    salt = bytes.fromhex(salt_hex) if salt_hex else secrets.token_bytes(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, AUTH_ITERATIONS)
    return digest.hex(), salt.hex()


def auth_verify_password(password, stored_hash, salt_hex):
    calculated, _ = auth_hash_password(password, salt_hex)
    return hmac.compare_digest(calculated, stored_hash)


def auth_user_count():
    with app_db_connect() as conn:
        return int(conn.execute("SELECT COUNT(*) FROM app_users").fetchone()[0])


def auth_get_user_by_email(email):
    with app_db_connect() as conn:
        row = conn.execute(
            "SELECT id,email,display_name,password_hash,password_salt,role,is_active,created_at,last_login_at "
            "FROM app_users WHERE LOWER(email)=LOWER(?)", (email.strip(),)
        ).fetchone()
    return dict(row) if row else None


def _database_active_value(conn, active):
    if APP_DB_BACKEND == "postgresql":
        info = conn.execute(
            "SELECT data_type FROM information_schema.columns "
            "WHERE table_schema=current_schema() AND table_name='app_users' AND column_name='is_active'"
        ).fetchone()
        if info and str(info[0]).lower() == "boolean":
            return bool(active)
    return int(bool(active))


def auth_create_user(email, display_name, password, role="analyst"):
    email = email.strip().lower()
    display_name = display_name.strip()
    if not email or "@" not in email or len(email) > 254:
        raise ValueError("Enter a valid email address.")
    if not display_name:
        raise ValueError("Enter the user's name.")
    if len(password) < 12:
        raise ValueError("Use a password with at least 12 characters.")
    if role not in VALID_ROLES:
        raise ValueError("Select a valid role.")
    password_hash, salt = auth_hash_password(password)
    try:
        with app_db_connect() as conn:
            conn.execute(
                "INSERT INTO app_users(email,display_name,password_hash,password_salt,role,is_active,created_at) "
                "VALUES(?,?,?,?,?,?,?)",
                (email, display_name, password_hash, salt, role, _database_active_value(conn, True),
                 datetime.now(timezone.utc).isoformat(timespec="seconds"))
            )
    except (sqlite3.IntegrityError, psycopg.IntegrityError if PSYCOPG_AVAILABLE else sqlite3.IntegrityError) as exc:
        raise ValueError("An account with this email already exists.") from exc


def auth_login(email, password):
    user = auth_get_user_by_email(email)
    if not user or not user["is_active"]:
        return None
    if not auth_verify_password(password, user["password_hash"], user["password_salt"]):
        return None
    now = datetime.now(timezone.utc).isoformat(timespec="seconds")
    with app_db_connect() as conn:
        conn.execute("UPDATE app_users SET last_login_at=? WHERE id=?", (now, user["id"]))
    return {"id": user["id"], "email": user["email"],
            "display_name": user["display_name"], "role": user["role"]}


def auth_list_users():
    with app_db_connect() as conn:
        return _query_dataframe(conn, "SELECT id,email,display_name,role,is_active,created_at,last_login_at FROM app_users ORDER BY id")


def auth_update_user(user_id, role, is_active):
    if role not in VALID_ROLES:
        raise ValueError("Invalid role.")
    with app_db_connect() as conn:
        current = conn.execute("SELECT role,is_active FROM app_users WHERE id=?", (int(user_id),)).fetchone()
        if not current:
            raise ValueError("User not found.")
        if current["role"] == "admin" and (role != "admin" or not is_active):
            active_admins = conn.execute(
                "SELECT COUNT(*) FROM app_users WHERE role='admin' AND is_active=?"
            , (_database_active_value(conn, True),)).fetchone()[0]
            if active_admins <= 1:
                raise ValueError("You cannot deactivate or demote the last active administrator.")
        conn.execute("UPDATE app_users SET role=?,is_active=? WHERE id=?",
                     (role, _database_active_value(conn, is_active), int(user_id)))


def auth_change_password(user_id, current_password, new_password):
    with app_db_connect() as conn:
        row = conn.execute("SELECT password_hash,password_salt FROM app_users WHERE id=?", (int(user_id),)).fetchone()
        if not row or not auth_verify_password(current_password, row["password_hash"], row["password_salt"]):
            raise ValueError("Current password is incorrect.")
        if len(new_password) < 12:
            raise ValueError("New password must contain at least 12 characters.")
        new_hash, new_salt = auth_hash_password(new_password)
        conn.execute("UPDATE app_users SET password_hash=?,password_salt=? WHERE id=?",
                     (new_hash, new_salt, int(user_id)))


LOGIN_MAX_FAILURES = 5
LOGIN_LOCK_MINUTES = 5


def auth_login_lock_status(email):
    """Return remaining lock seconds for an email address, using persistent DB state."""
    normalized_email = str(email or "").strip().lower()[:254]
    if not normalized_email:
        return 0
    with app_db_connect() as conn:
        conn.execute("""CREATE TABLE IF NOT EXISTS login_attempts (
            email TEXT PRIMARY KEY,
            failed_attempts INTEGER NOT NULL DEFAULT 0,
            locked_until TEXT
        )""")
        row = conn.execute(
            "SELECT failed_attempts, locked_until FROM login_attempts WHERE email=?",
            (normalized_email,),
        ).fetchone()
    if not row or not row["locked_until"]:
        return 0
    try:
        unlock_at = datetime.fromisoformat(row["locked_until"])
        remaining = int((unlock_at - datetime.now(timezone.utc)).total_seconds())
    except (TypeError, ValueError):
        remaining = 0
    if remaining <= 0:
        with app_db_connect() as conn:
            conn.execute("DELETE FROM login_attempts WHERE email=?", (normalized_email,))
        return 0
    return remaining


def auth_record_failed_login(email):
    """Persist failed attempts and apply a temporary account-level lockout."""
    normalized_email = str(email or "").strip().lower()[:254]
    if not normalized_email:
        return 0
    now = datetime.now(timezone.utc)
    with app_db_connect() as conn:
        conn.execute("""CREATE TABLE IF NOT EXISTS login_attempts (
            email TEXT PRIMARY KEY,
            failed_attempts INTEGER NOT NULL DEFAULT 0,
            locked_until TEXT
        )""")
        row = conn.execute(
            "SELECT failed_attempts, locked_until FROM login_attempts WHERE email=?",
            (normalized_email,),
        ).fetchone()
        count = int(row["failed_attempts"]) if row else 0
        locked_until = row["locked_until"] if row else None
        if locked_until:
            try:
                if datetime.fromisoformat(locked_until) > now:
                    return LOGIN_LOCK_MINUTES * 60
            except ValueError:
                pass
        count += 1
        new_lock = (now + pd.Timedelta(minutes=LOGIN_LOCK_MINUTES)).isoformat(timespec="seconds") if count >= LOGIN_MAX_FAILURES else None
        conn.execute(
            """INSERT INTO login_attempts(email, failed_attempts, locked_until) VALUES(?,?,?)
               ON CONFLICT(email) DO UPDATE SET failed_attempts=excluded.failed_attempts,
               locked_until=excluded.locked_until""",
            (normalized_email, count, new_lock),
        )
    return LOGIN_LOCK_MINUTES * 60 if new_lock else 0


def auth_clear_failed_logins(email):
    normalized_email = str(email or "").strip().lower()[:254]
    if normalized_email:
        with app_db_connect() as conn:
            conn.execute("DELETE FROM login_attempts WHERE email=?", (normalized_email,))


def auth_render_gate():
    """Render first-admin setup or login, and return the authenticated user."""
    st.markdown('<div class="hero"><h1>DataWise AI</h1><p>Secure analytics workspace · Stage 7.7</p></div>',
                unsafe_allow_html=True)
    left, center, right = st.columns([1, 1.2, 1])
    with center:
        if auth_user_count() == 0:
            st.subheader("🔐 Create the first administrator")
            st.info("This one-time setup creates the owner account. Keep its password private.")
            with st.form("stage77_first_admin_form"):
                name = st.text_input("Administrator name")
                email = st.text_input("Administrator email")
                password = st.text_input("Password (minimum 12 characters)", type="password")
                confirm = st.text_input("Confirm password", type="password")
                submitted = st.form_submit_button("Create administrator", type="primary", use_container_width=True)
            if submitted:
                if password != confirm:
                    st.error("Passwords do not match.")
                else:
                    try:
                        auth_create_user(email, name, password, "admin")
                        st.success("Administrator account created. Please sign in.")
                        st.rerun()
                    except ValueError as exc:
                        st.error(str(exc))
            return None

        st.subheader("Sign in to DataWise AI")
        with st.form("stage77_login_form"):
            email = st.text_input("Email address")
            password = st.text_input("Password", type="password")
            submitted = st.form_submit_button("Sign in", type="primary", use_container_width=True)
        if submitted:
            remaining_lock = auth_login_lock_status(email)
            if remaining_lock > 0:
                st.error(f"Too many unsuccessful attempts. Try again in about {max(1, (remaining_lock + 59) // 60)} minute(s).")
            else:
                user = auth_login(email, password)
                if user:
                    auth_clear_failed_logins(email)
                    db_log_event("login-" + str(uuid.uuid4()), "login_success", f"Successful sign-in: {user['email']}")
                    st.session_state.datawise_user = user
                    st.session_state.datawise_session_id = str(uuid.uuid4())
                    st.session_state.datawise_login_at = datetime.now(timezone.utc).timestamp()
                    st.session_state.datawise_last_activity = datetime.now(timezone.utc).timestamp()
                    st.session_state.datawise_login_failures = 0
                    st.rerun()
                else:
                    lock_seconds = auth_record_failed_login(email)
                    db_log_event("login-" + str(uuid.uuid4()), "login_failed", f"Failed sign-in attempt for: {email.strip().lower()[:254]}")
                    st.session_state.datawise_login_failures = st.session_state.get("datawise_login_failures", 0) + 1
                    if lock_seconds > 0:
                        st.error("Too many unsuccessful attempts. This account is temporarily locked for 5 minutes.")
                    else:
                        st.error("Sign-in failed. Check your email/password or contact an administrator.")
        st.caption("Accounts are created and managed by an administrator. Contact your workspace administrator if you need access.")
    return None


try:
    init_app_database()
    APP_DB_READY = True
except Exception as _db_init_error:
    APP_DB_READY = False
    st.error(f"Application database could not be initialized: {_db_init_error}")

if not APP_DB_READY:
    st.stop()

if "datawise_user" not in st.session_state:
    auth_render_gate()
    st.stop()

# Stage 7.9: expire idle sessions after 30 minutes and absolute sessions after 12 hours.
_now_ts = datetime.now(timezone.utc).timestamp()
_idle_ts = float(st.session_state.get("datawise_last_activity", _now_ts))
_login_ts = float(st.session_state.get("datawise_login_at", _now_ts))
if (_now_ts - _idle_ts) > 30 * 60 or (_now_ts - _login_ts) > 12 * 60 * 60:
    _expired_email = st.session_state.get("datawise_user", {}).get("email", "unknown")
    _expired_session = st.session_state.get("datawise_session_id", "unknown")
    db_log_event(str(_expired_session), "session_expired", f"Session expired due to inactivity or maximum lifetime: {_expired_email}")
    for _key in ("datawise_user", "datawise_session_id", "datawise_login_at", "datawise_last_activity", "raw_df", "cleaned_df", "source_name", "dataset_signature"):
        st.session_state.pop(_key, None)
    st.warning("Your session expired for security. Please sign in again.")
    auth_render_gate()
    st.stop()
st.session_state.datawise_last_activity = _now_ts

# Refresh role/activity status from the database on each rerun.
_authenticated_user = auth_get_user_by_email(st.session_state.datawise_user.get("email", ""))
if not _authenticated_user or not _authenticated_user["is_active"]:
    st.session_state.pop("datawise_user", None)
    st.error("This account is unavailable or has been deactivated. Please sign in again.")
    st.stop()
st.session_state.datawise_user = {
    "id": _authenticated_user["id"], "email": _authenticated_user["email"],
    "display_name": _authenticated_user["display_name"], "role": _authenticated_user["role"]
}

# =========================================================
# SESSION STATE
# =========================================================
if "raw_df" not in st.session_state:
    st.session_state.raw_df = None
if "source_name" not in st.session_state:
    st.session_state.source_name = None
if "cleaned_df" not in st.session_state:
    st.session_state.cleaned_df = None
if "datawise_session_id" not in st.session_state:
    st.session_state.datawise_session_id = str(uuid.uuid4())
if "datawise_login_at" not in st.session_state:
    st.session_state.datawise_login_at = datetime.now(timezone.utc).timestamp()
if "datawise_last_activity" not in st.session_state:
    st.session_state.datawise_last_activity = datetime.now(timezone.utc).timestamp()


# =========================================================
# SIDEBAR
# =========================================================
with st.sidebar:
    _workspace_brand = html.escape(
        safe_text(db_get_setting("workspace_label", "DataWise AI"), 60),
        quote=True,
    )
    st.markdown(f'<div class="brand">{_workspace_brand}</div><div class="brand-sub">SMART ANALYTICS WORKSPACE</div>', unsafe_allow_html=True)
    current_user = st.session_state.datawise_user
    st.markdown(f"**{safe_text(current_user['display_name'], 80)}**")
    st.caption(f"{safe_text(current_user['email'], 120)} · {current_user['role'].upper()}")
    all_pages = ["🏠 Dashboard", "📂 Data Explorer", "🧠 Advanced Analysis", "✨ Business Insights", "🛠️ Chart Builder", "🧹 Data Cleaning", "💬 Data Assistant", "🧾 Report Generator", "📑 Reports", "🗄️ Database Import", "📚 My History", "🗃️ App History", "👥 User Management", "🛡️ Admin Monitor", "🚀 Deployment & Settings", "🔐 Account Settings"]
    if current_user["role"] == "admin":
        allowed_pages = all_pages
    elif current_user["role"] == "analyst":
        allowed_pages = [item for item in all_pages if item not in ("🗃️ App History", "👥 User Management", "🛡️ Admin Monitor", "🚀 Deployment & Settings")]
    else:
        allowed_pages = ["🏠 Dashboard", "📂 Data Explorer", "🧠 Advanced Analysis", "✨ Business Insights", "🛠️ Chart Builder", "💬 Data Assistant", "📑 Reports", "📚 My History", "🔐 Account Settings"]
    page = st.radio("WORKSPACE", allowed_pages, label_visibility="visible")
    if st.button("↪ Sign out", use_container_width=True, key="stage77_logout"):
        db_log_event(str(st.session_state.get("datawise_session_id", "unknown")), "logout", f"Signed out: {current_user['email']}")
        for _key in ("datawise_user", "datawise_session_id", "datawise_login_at", "datawise_last_activity", "raw_df", "cleaned_df", "source_name", "dataset_signature"):
            st.session_state.pop(_key, None)
        st.rerun()
    st.divider()
    st.markdown("### ✨ Platform")
    st.caption("Turn raw data into clear, useful insights.")
    st.caption("Stage 7.10 adds workspace customization and deployment diagnostics.")
    st.markdown("### 🔐 Privacy")
    st.caption("This version analyzes data locally in the app. Avoid uploading confidential data to a public deployment until access controls and storage policies are configured.")
    st.markdown("### 🧠 AI mode")
    st.caption("Free-first analytics: numerical results are calculated with Python. Insights are rule-based, not generated by a hosted AI model.")


# =========================================================
# HEADER + UPLOAD
# =========================================================
st.markdown(
    '<div class="hero"><h1>AI Data Analyst</h1><p>Stage 7.10 · Workspace customization and deployment readiness.</p></div>',
    unsafe_allow_html=True,
)

with st.container(border=True):
    if st.session_state.datawise_user["role"] == "viewer":
        uploaded_file = None
        st.info("You are signed in with Viewer access. Dataset uploads, imports, cleaning and report generation are disabled for this role.")
    else:
        left, right = st.columns([3, 2])
        with left:
            uploaded_file = st.file_uploader(
                "Upload a CSV, Excel, or JSON dataset",
                type=["csv", "xlsx", "xls", "json"],
                help="For Excel workbooks, choose the worksheet after uploading.",
            )
        with right:
            st.markdown("**Quick start**")
            st.caption("Upload your file, review the automatic dashboard, then customize charts or export a report.")

if uploaded_file is not None:
    suffix = Path(uploaded_file.name).suffix.lower()
    allowed_extensions = {".csv", ".xlsx", ".xls", ".json"}
    max_upload_bytes = 100 * 1024 * 1024  # 100 MB

    if suffix not in allowed_extensions:
        st.error("Unsupported file format. Please upload CSV, XLSX, XLS, or JSON.")

    elif uploaded_file.size > max_upload_bytes:
        st.error("The uploaded file is larger than the 100 MB limit. Please upload a smaller file.")

    else:
        try:
            selected_sheet = None

            if suffix in (".xlsx", ".xls"):
                uploaded_file.seek(0)
                workbook = pd.ExcelFile(uploaded_file)
                if not workbook.sheet_names:
                    raise ValueError("No worksheets were found in this Excel workbook.")

                selected_sheet = st.selectbox(
                    "Excel worksheet",
                    workbook.sheet_names,
                    key=f"sheet_{uploaded_file.name}",
                )
                uploaded_file.seek(0)
                loaded = read_uploaded_file(uploaded_file, sheet_name=selected_sheet)
            else:
                loaded = read_uploaded_file(uploaded_file)

            if not isinstance(loaded, pd.DataFrame):
                raise ValueError("The uploaded file did not produce a valid table.")

            if loaded.empty:
                st.error("The uploaded file contains no data rows. Please check the file and try again.")
            else:
                # Detect duplicate headers before normalizing them.
                duplicate_header_count = int(loaded.columns.duplicated().sum())
                loaded = clean_column_names(loaded)
                errors, warnings = validate_dataframe(loaded)

                if duplicate_header_count:
                    warnings.insert(
                        0,
                        f"{duplicate_header_count} duplicate column header(s) were renamed automatically to keep column names unique."
                    )

                display_validation_report(errors, warnings)

                if not errors:
                    # Include file bytes in the signature so replacing a file with
                    # different content but the same name/shape is detected.
                    import hashlib
                    uploaded_file.seek(0)
                    file_digest = hashlib.sha256(uploaded_file.getvalue()).hexdigest()
                    signature = f"{uploaded_file.name}|{selected_sheet or ''}|{file_digest}"

                    if st.session_state.get("dataset_signature") != signature:
                        st.session_state.raw_df = loaded.copy()
                        st.session_state.cleaned_df = None
                        st.session_state.source_name = uploaded_file.name
                        st.session_state.dataset_signature = signature
                        if APP_DB_READY:
                            db_record_dataset(st.session_state.datawise_session_id, signature, uploaded_file.name,
                                              suffix.lstrip("."), loaded, st.session_state.datawise_user["id"])

                    st.success(
                        f"Loaded {uploaded_file.name} — "
                        f"{len(loaded):,} rows × {len(loaded.columns):,} columns."
                    )

        except Exception as exc:
            st.error(
                "Could not read this file. Check that it is a valid CSV, Excel, or JSON dataset."
            )
            st.caption(f"Technical details: {exc}")


# =========================================================
# DATABASE IMPORT (STAGE 7.5)
# =========================================================
if page == "🗄️ Database Import":
    st.markdown("## 🗄️ Database import")
    st.caption("Connect to a SQLite database file or a SQLAlchemy-supported database. Data is imported into this session for analysis; this app does not modify database records.")
    db_tab_sqlite, db_tab_server = st.tabs(["SQLite database file", "PostgreSQL / MySQL / other SQL"])

    with db_tab_sqlite:
        sqlite_file = st.file_uploader(
            "Upload a SQLite database", type=["db", "sqlite", "sqlite3"],
            help="Upload a .db, .sqlite, or .sqlite3 file (maximum 100 MB).",
            key="stage75_sqlite_file",
        )
        if sqlite_file is not None:
            if sqlite_file.size > 100 * 1024 * 1024:
                st.error("The database file exceeds the 100 MB limit.")
            else:
                try:
                    sqlite_bytes = sqlite_file.getvalue()
                    with tempfile.NamedTemporaryFile(suffix=".sqlite", delete=True) as temp_db:
                        temp_db.write(sqlite_bytes)
                        temp_db.flush()
                        test_conn = sqlite3.connect(Path(temp_db.name).resolve().as_uri() + "?mode=ro", uri=True)
                        try:
                            sqlite_tables = pd.read_sql_query(
                                "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name",
                                test_conn,
                            )["name"].tolist()
                        finally:
                            test_conn.close()
                    if not sqlite_tables:
                        st.warning("No user tables were found in this SQLite database.")
                    else:
                        selected_sqlite_table = st.selectbox("Choose a table", sqlite_tables, key="stage75_sqlite_table")
                        if st.button("Load SQLite table", type="primary", key="stage75_load_sqlite"):
                            imported = load_sqlite_table(sqlite_bytes, selected_sqlite_table)
                            if imported.empty:
                                st.warning("The selected table contains no rows.")
                            else:
                                imported = clean_column_names(imported)
                                errors, warnings = validate_dataframe(imported)
                                display_validation_report(errors, warnings)
                                if not errors:
                                    st.session_state.raw_df = imported.copy()
                                    st.session_state.cleaned_df = None
                                    st.session_state.source_name = f"{sqlite_file.name} · {selected_sqlite_table}"
                                    st.session_state.dataset_signature = f"sqlite|{sqlite_file.name}|{selected_sqlite_table}|{hashlib.sha256(sqlite_bytes).hexdigest()}"
                                    if APP_DB_READY:
                                        db_record_dataset(st.session_state.datawise_session_id, st.session_state.dataset_signature,
                                                          st.session_state.source_name, "sqlite", imported, st.session_state.datawise_user["id"])
                                    st.success(f"Imported {len(imported):,} rows × {len(imported.columns):,} columns. Your analytics workspace is ready.")
                                    st.rerun()
                except Exception as exc:
                    st.error(f"Could not read the SQLite database: {exc}")

    with db_tab_server:
        st.info("Install SQLAlchemy and the database driver for your database. Use a read-only database account where possible.")
        connection_url = st.text_input(
            "Database connection URL", type="password", key="stage75_connection_url",
            placeholder="postgresql+psycopg://user:password@host:5432/database",
            help="Examples: postgresql+psycopg://… or mysql+pymysql://… . Keep this URL private.",
        )
        st.caption("Install the matching driver too. PostgreSQL example: python -m pip install sqlalchemy psycopg. MySQL example: python -m pip install sqlalchemy pymysql.")
        table_to_load = st.text_input("Table name", key="stage75_server_table", placeholder="sales")
        if st.button("Connect and import table", type="primary", key="stage75_load_server"):
            if not connection_url.strip() or not table_to_load.strip():
                st.warning("Enter both the connection URL and table name.")
            else:
                try:
                    with st.spinner("Connecting and importing table..."):
                        imported = load_sqlalchemy_table(connection_url, table_to_load.strip())
                    if imported.empty:
                        st.warning("The selected table contains no rows.")
                    else:
                        imported = clean_column_names(imported)
                        errors, warnings = validate_dataframe(imported)
                        display_validation_report(errors, warnings)
                        if not errors:
                            st.session_state.raw_df = imported.copy()
                            st.session_state.cleaned_df = None
                            st.session_state.source_name = f"SQL database · {table_to_load.strip()}"
                            st.session_state.dataset_signature = f"sql|{table_to_load.strip()}|{len(imported)}|{hashlib.sha256(imported.to_csv(index=False).encode('utf-8')).hexdigest()}"
                            if APP_DB_READY:
                                db_record_dataset(st.session_state.datawise_session_id, st.session_state.dataset_signature,
                                                  st.session_state.source_name, "sql", imported, st.session_state.datawise_user["id"])
                            st.success(f"Imported {len(imported):,} rows × {len(imported.columns):,} columns. Your analytics workspace is ready.")
                            st.rerun()
                except Exception as exc:
                    st.error(f"Database connection or import failed: {exc}")


base_df = st.session_state.cleaned_df if st.session_state.cleaned_df is not None else st.session_state.raw_df

if base_df is None:
    st.markdown("### Get started")
    st.info("To import SQL data, open **Database Import** in the sidebar. You can also upload a CSV, Excel, or JSON file above.")
    c1, c2, c3 = st.columns(3)
    c1.info("**1. Upload**\n\nBring a CSV, Excel, or JSON file.")
    c2.info("**2. Explore**\n\nSee KPIs, charts, and data-quality findings.")
    c3.info("**3. Export**\n\nDownload a PDF report, PowerPoint, Excel, or CSV.")
    st.markdown('<div class="footer">DATAWISE AI · Built with Python and Streamlit</div>', unsafe_allow_html=True)
    st.stop()


# =========================================================
# FILTERS (APPLY TO ANALYTICS)
# =========================================================
st.markdown('<div class="section-kicker">Explore your data</div>', unsafe_allow_html=True)
st.markdown("### 🔎 Dashboard filters")
filtered_df = base_df.copy()
filter_cols = [c for c in base_df.columns if 1 < base_df[c].nunique(dropna=True) <= 100]
date_cols = date_columns(base_df)
filter_candidates = list(dict.fromkeys(date_cols + filter_cols))[:3]

if filter_candidates:
    filter_boxes = st.columns(min(3, len(filter_candidates)))
    for idx, col in enumerate(filter_candidates):
        with filter_boxes[idx]:
            if col in date_cols:
                converted = pd.to_datetime(base_df[col], errors="coerce")
                valid_dates = converted.dropna()
                if not valid_dates.empty:
                    min_date = valid_dates.min().date()
                    max_date = valid_dates.max().date()
                    chosen = st.date_input(
                        f"Date range: {col}",
                        value=(min_date, max_date),
                        min_value=min_date,
                        max_value=max_date,
                        key=f"filter_date_{col}",
                    )
                    if isinstance(chosen, tuple) and len(chosen) == 2:
                        mask = converted.dt.date.between(chosen[0], chosen[1])
                        filtered_df = filtered_df.loc[mask]
            else:
                values = base_df[col].dropna().astype(str).unique().tolist()
                values = sorted(values, key=str)[:100]
                selected = st.multiselect(
                    f"Filter: {col}",
                    options=values,
                    default=[],
                    key=f"filter_cat_{col}",
                    placeholder="All values",
                )
                if selected:
                    filtered_df = filtered_df[filtered_df[col].astype(str).isin(selected)]

if filtered_df.empty:
    st.warning("The selected filters return no rows. Change or clear the filters.")
    st.stop()


# =========================================================
# DASHBOARD
# =========================================================
if page == "🏠 Dashboard":
    st.markdown("## 📊 Executive dashboard")
    st.caption(f"Showing {len(filtered_df):,} of {len(base_df):,} rows from **{st.session_state.source_name or 'uploaded dataset'}**.")

    # Stage 7.2 dataset snapshot: gives users an immediate, scannable view of the file.
    snapshot_cols = st.columns(4)
    numeric_count = len(numeric_columns(filtered_df))
    category_count = len([c for c in filtered_df.columns if c not in numeric_columns(filtered_df) and not pd.api.types.is_datetime64_any_dtype(filtered_df[c])])
    date_count = len(date_columns(filtered_df))
    snapshot_items = [
        ("DATASET", st.session_state.source_name or "Uploaded data"),
        ("NUMERIC FIELDS", f"{numeric_count:,}"),
        ("CATEGORY FIELDS", f"{category_count:,}"),
        ("DATE FIELDS", f"{date_count:,}"),
    ]
    for box, (label, value) in zip(snapshot_cols, snapshot_items):
        box.markdown(f'<div class="snapshot-card"><div class="snapshot-label">{html.escape(label)}</div><div class="snapshot-value">{html.escape(str(value))}</div></div>', unsafe_allow_html=True)
    export_col, status_col = st.columns([1, 4])
    with export_col:
        st.download_button("⬇️ Export filtered data", data=filtered_df.to_csv(index=False).encode("utf-8-sig"), file_name="datawise_filtered_data.csv", mime="text/csv", key="stage72_filtered_export", use_container_width=True)
    with status_col:
        st.markdown('<span class="status-pill">● Dataset loaded and ready</span>', unsafe_allow_html=True)

    missing, duplicates, completeness = data_quality(filtered_df)
    nums = numeric_columns(filtered_df)
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Total records", f"{len(filtered_df):,}")
    k2.metric("Columns", f"{len(filtered_df.columns):,}")
    k3.metric("Missing cells", f"{missing:,}")
    k4.metric("Completeness", f"{completeness:.1f}%")

    if nums:
        metric_col = st.selectbox("Choose a key metric for additional KPIs", nums, key="dashboard_metric")
        series = pd.to_numeric(filtered_df[metric_col], errors="coerce").dropna()
        a, b, c = st.columns(3)
        a.metric(f"Total {metric_col}", f"{series.sum():,.2f}")
        b.metric(f"Average {metric_col}", f"{series.mean():,.2f}" if not series.empty else "N/A")
        c.metric(f"Maximum {metric_col}", f"{series.max():,.2f}" if not series.empty else "N/A")

    # =====================================================
    # STAGE 5.7 — INTERACTIVE KPI MONITORING
    # =====================================================
    st.markdown("### 🎯 Stage 5.7 — Interactive KPI monitoring")
    st.caption("Set a target for the selected metric, monitor performance, and export a KPI snapshot. Targets are user-defined benchmarks, not forecasts.")
    if nums:
        monitor_col = st.selectbox("KPI to monitor", nums, key="stage57_monitor_metric")
        monitor_values = pd.to_numeric(filtered_df[monitor_col], errors="coerce").dropna()
        if not monitor_values.empty:
            m_total = float(monitor_values.sum())
            m_average = float(monitor_values.mean())
            m_count = int(monitor_values.size)
            cfg1, cfg2, cfg3 = st.columns([1.3, 1.2, 1.5])
            with cfg1:
                target_basis = st.selectbox("Target compares against", ["Total", "Average", "Maximum", "Minimum"], key="stage57_target_basis")
            actual_map = {"Total": m_total, "Average": m_average, "Maximum": float(monitor_values.max()), "Minimum": float(monitor_values.min())}
            actual_value = actual_map[target_basis]
            with cfg2:
                target_value = st.number_input("Target value", value=float(actual_value), key="stage57_target_value", help="Enter the business target you want to monitor.")
            with cfg3:
                target_direction = st.selectbox("Success direction", ["Higher is better", "Lower is better"], key="stage57_target_direction")

            if target_direction == "Higher is better":
                target_met = actual_value >= target_value
            else:
                target_met = actual_value <= target_value
            variance = actual_value - target_value
            variance_pct = ((variance / abs(target_value)) * 100) if target_value != 0 else None
            status = "On target" if target_met else "Needs attention"
            s1, s2, s3, s4 = st.columns(4)
            s1.metric(f"Actual {target_basis.lower()}", f"{actual_value:,.2f}")
            s2.metric("Target", f"{target_value:,.2f}")
            s3.metric("Variance", f"{variance:+,.2f}", delta=(f"{variance_pct:+.1f}% vs target" if variance_pct is not None else "Target is zero"))
            s4.metric("KPI status", "✅ On target" if target_met else "⚠️ Needs attention")
            if target_met:
                st.success(f"{monitor_col}: current {target_basis.lower()} is meeting the target.")
            else:
                st.warning(f"{monitor_col}: current {target_basis.lower()} is not meeting the target. Review the trend and contributing categories below.")

            snapshot = pd.DataFrame([{
                "Source file": st.session_state.source_name or "uploaded_dataset",
                "Rows after filters": int(len(filtered_df)),
                "KPI metric": str(monitor_col),
                "Target basis": target_basis,
                "Actual value": actual_value,
                "Target value": float(target_value),
                "Success direction": target_direction,
                "Variance": variance,
                "Variance percent": variance_pct,
                "Status": status,
                "Valid numeric values": m_count,
                "Generated date": date.today().isoformat(),
            }])
            st.download_button(
                "⬇️ Download KPI snapshot (CSV)",
                data=snapshot.to_csv(index=False).encode("utf-8-sig"),
                file_name="datawise_kpi_snapshot.csv",
                mime="text/csv",
                key="stage57_kpi_download",
            )

            # Drill-down: show which categories contribute most to the selected KPI.
            category_options = [c for c in filtered_df.columns if c not in nums and c not in date_columns(filtered_df) and 1 < filtered_df[c].nunique(dropna=True) <= 100]
            if category_options:
                st.markdown("#### 🔎 KPI category drill-down")
                drill_col = st.selectbox("Break this KPI down by", category_options, key="stage57_drill_category")
                drill_frame = filtered_df[[drill_col, monitor_col]].copy()
                drill_frame[monitor_col] = pd.to_numeric(drill_frame[monitor_col], errors="coerce")
                drill_frame = drill_frame.dropna(subset=[monitor_col])
                if not drill_frame.empty:
                    drill_summary = drill_frame.groupby(drill_col, dropna=False)[monitor_col].agg(Total="sum", Average="mean", Records="count").reset_index()
                    drill_summary = drill_summary.sort_values("Total", ascending=False)
                    st.dataframe(drill_summary, use_container_width=True, hide_index=True)
                    top_n = min(15, len(drill_summary))
                    chart_summary = drill_summary.head(top_n).copy()
                    chart_summary[drill_col] = chart_summary[drill_col].astype(str)
                    drill_fig = px.bar(chart_summary, x=drill_col, y="Total", title=f"Top {top_n} {drill_col} groups by total {monitor_col}", template="plotly_white")
                    drill_fig.update_layout(title_x=0.02, margin=dict(l=18, r=18, t=65, b=22))
                    st.plotly_chart(drill_fig, use_container_width=True)
                    st.download_button(
                        "⬇️ Download KPI category breakdown (CSV)",
                        data=drill_summary.to_csv(index=False).encode("utf-8-sig"),
                        file_name="datawise_kpi_category_breakdown.csv",
                        mime="text/csv",
                        key="stage57_drill_download",
                    )
                else:
                    st.info("No valid numeric values are available for this category breakdown.")
        else:
            st.info("The selected KPI column has no valid numeric values after filtering.")
    else:
        st.info("Upload data containing at least one numeric column to use KPI monitoring.")

    # =====================================================
    # STAGE 5.8 — PERIOD-OVER-PERIOD COMPARISON
    # =====================================================
    st.markdown("### 🗓️ Stage 5.8 — Period-over-period comparison")
    st.caption("Compare two user-selected date ranges using the same metric and aggregation. Percentage changes describe the selected data; they do not establish causation.")
    comparison_dates = date_columns(filtered_df)
    comparison_nums = numeric_columns(filtered_df)
    if comparison_dates and comparison_nums:
        compare_date_col = st.selectbox("Date column for comparison", comparison_dates, key="stage58_date_col")
        compare_metric = st.selectbox("Metric to compare", comparison_nums, key="stage58_metric")
        compare_aggregation = st.selectbox("Aggregation", ["Sum", "Average", "Record count"], key="stage58_aggregation")
        dates_parsed = pd.to_datetime(filtered_df[compare_date_col], errors="coerce")
        valid_dates = dates_parsed.dropna()
        if not valid_dates.empty:
            earliest = valid_dates.min().date()
            latest = valid_dates.max().date()
            p1, p2 = st.columns(2)
            with p1:
                st.markdown("**Period A (baseline)**")
                range_a = st.date_input("Choose baseline date range", value=(earliest, latest), min_value=earliest, max_value=latest, key="stage58_range_a")
            with p2:
                st.markdown("**Period B (comparison)**")
                range_b = st.date_input("Choose comparison date range", value=(earliest, latest), min_value=earliest, max_value=latest, key="stage58_range_b")

            def _stage58_value(mask):
                frame = filtered_df.loc[mask, [compare_metric]].copy()
                frame[compare_metric] = pd.to_numeric(frame[compare_metric], errors="coerce")
                values = frame[compare_metric].dropna()
                if compare_aggregation == "Record count":
                    return float(len(values)), int(len(values))
                if values.empty:
                    return None, 0
                if compare_aggregation == "Sum":
                    return float(values.sum()), int(len(values))
                return float(values.mean()), int(len(values))

            if isinstance(range_a, tuple) and len(range_a) == 2 and isinstance(range_b, tuple) and len(range_b) == 2:
                if range_a[0] > range_a[1] or range_b[0] > range_b[1]:
                    st.warning("For each period, choose a start date that is on or before its end date.")
                else:
                    dates_only = dates_parsed.dt.date
                    mask_a = dates_only.between(range_a[0], range_a[1])
                    mask_b = dates_only.between(range_b[0], range_b[1])
                    value_a, count_a = _stage58_value(mask_a)
                    value_b, count_b = _stage58_value(mask_b)
                    label_a = f"{range_a[0]} to {range_a[1]}"
                    label_b = f"{range_b[0]} to {range_b[1]}"
                    if value_a is None or value_b is None:
                        st.warning("One of the selected periods has no valid values for this metric. Adjust the date ranges or metric.")
                    else:
                        difference = value_b - value_a
                        change_pct = (difference / abs(value_a) * 100) if value_a != 0 else None
                        ca, cb, cc, cd = st.columns(4)
                        ca.metric("Period A", f"{value_a:,.2f}", help=label_a)
                        cb.metric("Period B", f"{value_b:,.2f}", help=label_b)
                        cc.metric("Absolute change", f"{difference:+,.2f}")
                        cd.metric("Percentage change", f"{change_pct:+.2f}%" if change_pct is not None else "N/A", help="Not defined when the baseline value is zero.")
                        if change_pct is not None:
                            if difference > 0:
                                st.success(f"{compare_metric} {compare_aggregation.lower()} increased by {change_pct:.2f}% from Period A to Period B.")
                            elif difference < 0:
                                st.info(f"{compare_metric} {compare_aggregation.lower()} decreased by {abs(change_pct):.2f}% from Period A to Period B.")
                            else:
                                st.info(f"{compare_metric} {compare_aggregation.lower()} was unchanged between the selected periods.")
                        comparison_summary = pd.DataFrame([
                            {"Period": "A (baseline)", "Start date": str(range_a[0]), "End date": str(range_a[1]), "Metric": compare_metric, "Aggregation": compare_aggregation, "Value": value_a, "Valid records": count_a},
                            {"Period": "B (comparison)", "Start date": str(range_b[0]), "End date": str(range_b[1]), "Metric": compare_metric, "Aggregation": compare_aggregation, "Value": value_b, "Valid records": count_b},
                        ])
                        st.dataframe(comparison_summary, use_container_width=True, hide_index=True)
                        st.download_button("⬇️ Download period comparison (CSV)", data=comparison_summary.to_csv(index=False).encode("utf-8-sig"), file_name="datawise_period_comparison.csv", mime="text/csv", key="stage58_comparison_download")

                        # Optional category-level comparison helps locate where the change came from.
                        category_options_58 = [c for c in filtered_df.columns if c not in comparison_nums and c not in comparison_dates and 1 < filtered_df[c].nunique(dropna=True) <= 100]
                        if category_options_58:
                            compare_category = st.selectbox("Optional category breakdown", category_options_58, key="stage58_category")
                            cat_frame = filtered_df[[compare_category, compare_metric]].copy()
                            cat_frame[compare_metric] = pd.to_numeric(cat_frame[compare_metric], errors="coerce")
                            cat_frame["_date"] = dates_parsed
                            cat_frame = cat_frame.dropna(subset=[compare_metric, "_date"])
                            cat_a = cat_frame[cat_frame["_date"].dt.date.between(range_a[0], range_a[1])].groupby(compare_category)[compare_metric].agg("sum" if compare_aggregation != "Average" else "mean").rename("Period A")
                            cat_b = cat_frame[cat_frame["_date"].dt.date.between(range_b[0], range_b[1])].groupby(compare_category)[compare_metric].agg("sum" if compare_aggregation != "Average" else "mean").rename("Period B")
                            cat_compare = pd.concat([cat_a, cat_b], axis=1).fillna(0).reset_index()
                            cat_compare["Change"] = cat_compare["Period B"] - cat_compare["Period A"]
                            cat_compare["Change %"] = cat_compare.apply(lambda row: (row["Change"] / abs(row["Period A"]) * 100) if row["Period A"] != 0 else None, axis=1)
                            cat_compare = cat_compare.sort_values("Change", ascending=False)
                            st.markdown("#### Category contribution comparison")
                            st.dataframe(cat_compare, use_container_width=True, hide_index=True)
                            st.download_button("⬇️ Download category comparison (CSV)", data=cat_compare.to_csv(index=False).encode("utf-8-sig"), file_name="datawise_period_category_comparison.csv", mime="text/csv", key="stage58_category_download")
        else:
            st.info("The selected date column has no valid dates to compare.")
    else:
        st.info("Stage 5.8 requires at least one detected date column and one numeric metric.")

    # =====================================================
    # STAGE 5.9 — ADVANCED DATA INSIGHTS & ANOMALY DETECTION
    # =====================================================
    st.markdown("### 🚨 Stage 5.9 — Advanced data insights & anomaly detection")
    st.caption("Flags statistical outliers and unusually large period-to-period changes for investigation. These are screening signals, not proof of an error or cause.")

    anomaly_nums = numeric_columns(filtered_df)
    anomaly_dates = date_columns(filtered_df)
    if anomaly_nums:
        anomaly_metric = st.selectbox("Numeric metric to inspect", anomaly_nums, key="stage59_metric")
        anomaly_method = st.selectbox("Outlier detection method", ["IQR (robust default)", "Z-score"], key="stage59_method")
        anomaly_sensitivity = st.slider("Detection sensitivity", min_value=1.0, max_value=3.5, value=1.5 if anomaly_method.startswith("IQR") else 2.5, step=0.5, key="stage59_sensitivity", help="IQR uses this as the multiplier around the interquartile range. Z-score uses this as the absolute score threshold.")
        values = pd.to_numeric(filtered_df[anomaly_metric], errors="coerce")
        valid = values.dropna()
        anomaly_rows = pd.DataFrame()
        lower_bound = upper_bound = np.nan
        if len(valid) >= 4:
            if anomaly_method.startswith("IQR"):
                q1, q3 = valid.quantile([0.25, 0.75])
                iqr = q3 - q1
                lower_bound = float(q1 - anomaly_sensitivity * iqr)
                upper_bound = float(q3 + anomaly_sensitivity * iqr)
                mask_outlier = values.lt(lower_bound) | values.gt(upper_bound)
                score_text = "Outside IQR bounds"
            else:
                mean_v = float(valid.mean())
                std_v = float(valid.std(ddof=0))
                zscores = (values - mean_v) / std_v if std_v > 0 else pd.Series(np.nan, index=values.index)
                mask_outlier = zscores.abs().gt(anomaly_sensitivity)
                lower_bound, upper_bound = mean_v - anomaly_sensitivity * std_v, mean_v + anomaly_sensitivity * std_v
                score_text = "Absolute z-score above threshold"
            if mask_outlier.any():
                anomaly_rows = filtered_df.loc[mask_outlier].copy()
                anomaly_rows.insert(0, "Original row index", anomaly_rows.index.astype(str))
                anomaly_rows["Detected metric"] = anomaly_metric
                anomaly_rows["Detected value"] = values.loc[mask_outlier].values
                anomaly_rows["Detection method"] = anomaly_method
                anomaly_rows["Detection rule"] = score_text
                anomaly_rows["Lower reference bound"] = lower_bound
                anomaly_rows["Upper reference bound"] = upper_bound
            total_valid = int(len(valid))
            outlier_count = int(mask_outlier.sum())
            k1, k2, k3 = st.columns(3)
            k1.metric("Valid metric values", f"{total_valid:,}")
            k2.metric("Potential outliers", f"{outlier_count:,}")
            k3.metric("Potential outlier rate", f"{(outlier_count / total_valid * 100):.2f}%" if total_valid else "N/A")
            st.caption(f"Reference bounds: {lower_bound:,.4g} to {upper_bound:,.4g}. Values outside the bounds are flagged for review.")
            if not anomaly_rows.empty:
                st.dataframe(anomaly_rows, use_container_width=True, hide_index=True)
                st.download_button("⬇️ Download potential outliers (CSV)", data=anomaly_rows.to_csv(index=False).encode("utf-8-sig"), file_name="datawise_potential_outliers.csv", mime="text/csv", key="stage59_outlier_download")
            else:
                st.success("No values crossed the selected outlier threshold.")
        else:
            st.info("At least four valid numeric values are needed for outlier screening.")

        # Periodic spike/drop screening, only when a date column is available.
        if anomaly_dates:
            st.markdown("#### 📉 Unusual changes over time")
            anomaly_date_col = st.selectbox("Date column for trend screening", anomaly_dates, key="stage59_date_col")
            trend_frequency = st.selectbox("Time aggregation", ["Day", "Week", "Month"], index=2, key="stage59_frequency")
            change_threshold = st.slider("Flag absolute change of at least (%)", min_value=10, max_value=300, value=50, step=10, key="stage59_change_threshold")
            trend = pd.DataFrame({"_date": pd.to_datetime(filtered_df[anomaly_date_col], errors="coerce"), "_value": values})
            trend = trend.dropna(subset=["_date", "_value"])
            if not trend.empty:
                trend = trend.set_index("_date").sort_index()
                rule = {"Day": "D", "Week": "W", "Month": "MS"}[trend_frequency]
                trend_summary = trend["_value"].resample(rule).sum(min_count=1).dropna().rename("Value").to_frame()
                trend_summary["Previous period value"] = trend_summary["Value"].shift(1)
                trend_summary["Absolute change"] = trend_summary["Value"] - trend_summary["Previous period value"]
                trend_summary["Change %"] = trend_summary.apply(lambda r: (r["Absolute change"] / abs(r["Previous period value"]) * 100) if pd.notna(r["Previous period value"]) and r["Previous period value"] != 0 else np.nan, axis=1)
                trend_summary["Anomaly"] = trend_summary["Change %"].abs().ge(change_threshold)
                trend_summary = trend_summary.reset_index().rename(columns={"_date": "Period", "index": "Period"})
                trend_summary["Period"] = pd.to_datetime(trend_summary["Period"], errors="coerce").dt.strftime("%Y-%m-%d")
                flagged_trend = trend_summary[trend_summary["Anomaly"]].copy()
                st.dataframe(trend_summary, use_container_width=True, hide_index=True)
                if not flagged_trend.empty:
                    st.warning(f"Found {len(flagged_trend)} period(s) with an absolute change of at least {change_threshold}%.")
                    st.dataframe(flagged_trend, use_container_width=True, hide_index=True)
                else:
                    st.success("No period-to-period changes crossed the selected threshold.")
                st.download_button("⬇️ Download time anomaly analysis (CSV)", data=trend_summary.to_csv(index=False).encode("utf-8-sig"), file_name="datawise_time_anomaly_analysis.csv", mime="text/csv", key="stage59_trend_download")
            else:
                st.info("No valid date-and-metric rows are available for time anomaly screening.")
    else:
        st.info("Stage 5.9 requires at least one numeric column. Upload a dataset with numeric measures to detect potential anomalies.")

    st.markdown("### 📈 Automatic visualizations")
    if nums:
        category_candidates = [
            col for col in filtered_df.columns
            if col not in nums and 1 < filtered_df[col].nunique(dropna=True) <= 100
        ]
        if category_candidates:
            category = category_candidates[0]
            metric = st.selectbox("Metric used for automatic charts", nums, key="auto_metric")
            chart1, chart2 = st.columns(2)
            try:
                with chart1:
                    fig = create_chart(filtered_df, category, metric, "Bar chart", "Sum", f"Top categories by {metric}")
                    st.plotly_chart(fig, use_container_width=True)
                with chart2:
                    fig = create_chart(filtered_df, category, metric, "Pie chart", "Sum", f"{metric} share by {category}")
                    st.plotly_chart(fig, use_container_width=True)
            except Exception as exc:
                st.info(f"Could not create category charts for these columns: {exc}")
        else:
            st.info("No suitable categorical column was found for category-comparison charts.")

        dcols = date_columns(filtered_df)
        if dcols:
            date_col = dcols[0]
            trend_metric = st.selectbox("Metric for trend chart", nums, key="trend_metric")
            trend_frame = filtered_df[[date_col, trend_metric]].copy()
            trend_frame[date_col] = pd.to_datetime(trend_frame[date_col], errors="coerce")
            trend_frame[trend_metric] = pd.to_numeric(trend_frame[trend_metric], errors="coerce")
            trend_frame = trend_frame.dropna()
            if not trend_frame.empty:
                trend_frame["_period"] = trend_frame[date_col].dt.to_period("M").astype(str)
                trend_frame = trend_frame.groupby("_period", as_index=False)[trend_metric].sum()
                fig = px.line(trend_frame, x="_period", y=trend_metric, markers=True, title=f"Monthly trend: {trend_metric}", template="plotly_white")
                fig.update_layout(title_x=0.02, margin=dict(l=18, r=18, t=65, b=22))
                st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No date column was confidently detected. Add a date field to enable trend analysis.")
    else:
        st.info("No numeric columns were detected. You can still explore categories and data quality.")

    st.markdown("### 🧠 Key findings")
    for title, body in make_insights(filtered_df):
        safe_title = html.escape(safe_text(title), quote=True)
        safe_body = html.escape(safe_text(body), quote=True)
        st.markdown(
            f'<div class="insight-card"><h4>{safe_title}</h4><p>{safe_body}</p></div>',
            unsafe_allow_html=True,
        )
    st.caption("These are locally calculated, rule-based findings. They are not generated by a hosted AI model.")


# =========================================================
# DATABASE IMPORT WORKSPACE (STAGE 7.5)
# =========================================================
elif page == "🗄️ Database Import":
    st.markdown("## 🗄️ Database import workspace")
    st.success("Database import is available above. Once a table is loaded, all existing DataWise analytics and report tools can use it.")
    if st.session_state.source_name:
        st.metric("Current data source", safe_text(st.session_state.source_name, 80))
        st.caption(f"{len(filtered_df):,} rows currently available after dashboard filters · {len(filtered_df.columns):,} columns")
        st.dataframe(filtered_df.head(100), use_container_width=True, height=360)
        st.download_button(
            "⬇️ Download imported table as CSV",
            filtered_df.to_csv(index=False).encode("utf-8"),
            file_name="database_import.csv", mime="text/csv",
        )
    st.markdown("### Supported connection options")
    st.markdown("""
    - **SQLite:** upload `.db`, `.sqlite`, or `.sqlite3` files.
    - **PostgreSQL:** SQLAlchemy URL with `psycopg` driver.
    - **MySQL:** SQLAlchemy URL with `pymysql` driver.
    - **Other SQL databases:** supported when SQLAlchemy and the appropriate driver are installed.
    """)
    st.warning("Use read-only database credentials. Connection details are used only for the import action and are not written into exported reports. Avoid exposing database URLs or passwords in screenshots or shared code.")


# =========================================================
# DATA EXPLORER
# =========================================================
elif page == "📂 Data Explorer":
    st.markdown("## 📂 Data explorer")
    st.caption(f"{len(filtered_df):,} rows · {len(filtered_df.columns):,} columns after filters")
    st.dataframe(filtered_df, use_container_width=True, height=480)
    with st.expander("Column profile"):
        profile = pd.DataFrame({
            "Column": filtered_df.columns.astype(str),
            "Data type": filtered_df.dtypes.astype(str).values,
            "Missing values": filtered_df.isna().sum().values,
            "Unique values": filtered_df.nunique(dropna=True).values,
        })
        st.dataframe(profile, use_container_width=True)
    st.download_button(
        "⬇️ Download filtered data (CSV)",
        filtered_df.to_csv(index=False).encode("utf-8"),
        file_name="filtered_dataset.csv",
        mime="text/csv",
    )



# =========================================================
# ADVANCED AUTOMATED ANALYSIS (STAGE 2)
# =========================================================
elif page == "🧠 Advanced Analysis":
    st.markdown("## 🧠 Advanced automated analysis")
    st.caption("Explore relationships, category performance, trends, distributions, and potential outliers. Results are calculated locally from the current filtered dataset.")

    nums = numeric_columns(filtered_df)
    dcols = date_columns(filtered_df)
    cats = [c for c in filtered_df.columns if c not in nums and c not in dcols and 1 < filtered_df[c].nunique(dropna=True) <= 500]

    st.markdown("### 1. Data-quality and outlier screening")
    outlier_summary = detect_iqr_outliers(filtered_df, nums)
    if outlier_summary.empty:
        st.info("Not enough numeric variation to run the IQR outlier screen. It generally needs at least four valid values and non-zero spread.")
    else:
        flagged_total = int(outlier_summary["Outlier candidates"].sum())
        o1, o2 = st.columns(2)
        o1.metric("Numeric columns screened", len(outlier_summary))
        o2.metric("Potential outlier flags", flagged_total)
        st.dataframe(outlier_summary, use_container_width=True, hide_index=True)
        st.download_button(
            "⬇️ Download outlier screening (CSV)",
            outlier_summary.to_csv(index=False).encode("utf-8"),
            file_name="datawise_outlier_screening.csv",
            mime="text/csv",
            key="download_outliers_stage2",
        )
        st.caption("IQR flags are screening candidates only. Unusual values may be valid business events; verify them before changing or deleting records.")

    st.divider()
    st.markdown("### 2. Numeric relationships")
    if len(nums) >= 2:
        corr = make_correlation_matrix(filtered_df, nums[:20])
        if not corr.empty:
            fig_corr = px.imshow(
                corr, text_auto=True, aspect="auto", color_continuous_scale="RdBu_r",
                zmin=-1, zmax=1, origin="lower", title="Numeric correlation matrix (Pearson)"
            )
            fig_corr.update_layout(title_x=0.02, margin=dict(l=15, r=15, t=65, b=20))
            st.plotly_chart(fig_corr, use_container_width=True)
            st.caption("Correlation ranges from -1 to +1. Correlation does not prove that one variable causes another.")
            pairs = []
            cols = corr.columns.tolist()
            for i in range(len(cols)):
                for j in range(i + 1, len(cols)):
                    pairs.append((cols[i], cols[j], float(corr.iloc[i, j])))
            pairs.sort(key=lambda item: abs(item[2]), reverse=True)
            if pairs:
                best = pairs[0]
                st.info(f"Strongest absolute pair in the displayed matrix: **{best[0]}** and **{best[1]}** (correlation {best[2]:.3f}). This describes association, not causation.")
        else:
            st.info("The selected numeric columns do not have enough variation to calculate correlations.")
    else:
        st.info("Upload at least two numeric columns to explore correlations.")

    st.divider()
    st.markdown("### 3. Category performance comparison")
    if cats and nums:
        cat_col, metric_col, agg_col = st.columns([1.2, 1.2, 1])
        with cat_col:
            selected_category = st.selectbox("Category", cats, key="advanced_category")
        with metric_col:
            selected_metric = st.selectbox("Measure", nums, key="advanced_metric")
        with agg_col:
            selected_aggregation = st.selectbox("Calculation", ["Sum", "Average", "Count", "Maximum", "Minimum"], key="advanced_aggregation")
        category_result = summarize_category_performance(filtered_df, selected_category, selected_metric, selected_aggregation)
        if not category_result.empty:
            shown = category_result.head(25)
            fig_cat = px.bar(
                shown.sort_values(selected_aggregation, ascending=True),
                x=selected_aggregation, y=str(selected_category), orientation="h",
                title=f"{selected_aggregation} of {selected_metric} by {selected_category}",
                template="plotly_white", text_auto=".2s", color_discrete_sequence=["#2563eb"]
            )
            fig_cat.update_layout(title_x=0.02, yaxis_title=str(selected_category), xaxis_title=f"{selected_aggregation} of {selected_metric}")
            st.plotly_chart(fig_cat, use_container_width=True)
            st.dataframe(category_result, use_container_width=True, hide_index=True)
            st.download_button(
                "⬇️ Download category analysis (CSV)",
                category_result.to_csv(index=False).encode("utf-8"),
                file_name="datawise_category_analysis.csv",
                mime="text/csv",
                key="download_category_stage2",
            )
            top = category_result.iloc[0]
            st.success(f"Highest {selected_aggregation.lower()} group: **{top[str(selected_category)]}** — {top[selected_aggregation]:,.2f} for {selected_metric}.")
    else:
        st.info("A numeric measure and a categorical column are needed for category performance analysis.")

    st.divider()
    st.markdown("### 4. Trend analysis")
    if dcols and nums:
        trend_controls = st.columns(4)
        with trend_controls[0]:
            trend_date_col = st.selectbox("Date field", dcols, key="advanced_trend_date")
        with trend_controls[1]:
            trend_metric = st.selectbox("Trend measure", nums, key="advanced_trend_metric")
        with trend_controls[2]:
            trend_frequency = st.selectbox("Time grouping", ["Day", "Week", "Month", "Quarter", "Year"], index=2, key="advanced_trend_frequency")
        with trend_controls[3]:
            trend_aggregation = st.selectbox("Trend calculation", ["Sum", "Average", "Count"], key="advanced_trend_aggregation")
        trend_data = filtered_df[[trend_date_col, trend_metric]].copy()
        trend_data[trend_date_col] = pd.to_datetime(trend_data[trend_date_col], errors="coerce")
        trend_data[trend_metric] = pd.to_numeric(trend_data[trend_metric], errors="coerce")
        trend_data = trend_data.dropna(subset=[trend_date_col])
        if not trend_data.empty:
            trend_data = trend_data.set_index(trend_date_col)
            if trend_frequency == "Day": rule = "D"
            elif trend_frequency == "Week": rule = "W"
            elif trend_frequency == "Month": rule = "MS"
            elif trend_frequency == "Quarter": rule = "QS"
            else: rule = "YS"
            grouped_trend = trend_data[trend_metric].resample(rule)
            if trend_aggregation == "Average": trend_result = grouped_trend.mean()
            elif trend_aggregation == "Count": trend_result = grouped_trend.count()
            else: trend_result = grouped_trend.sum(min_count=1)
            trend_result = trend_result.dropna().rename("Value").reset_index()
            if not trend_result.empty:
                fig_trend = px.line(trend_result, x=trend_date_col, y="Value", markers=True, title=f"{trend_aggregation} of {trend_metric} by {trend_frequency.lower()}", template="plotly_white")
                fig_trend.update_layout(title_x=0.02, xaxis_title=trend_frequency, yaxis_title=f"{trend_aggregation} of {trend_metric}")
                st.plotly_chart(fig_trend, use_container_width=True)
                st.dataframe(trend_result, use_container_width=True, hide_index=True)
                st.download_button(
                    "⬇️ Download trend analysis (CSV)",
                    trend_result.to_csv(index=False).encode("utf-8"),
                    file_name="datawise_trend_analysis.csv",
                    mime="text/csv",
                    key="download_trend_stage2",
                )
            else:
                st.info("No usable numeric observations remain after date and value cleaning.")
        else:
            st.info("No valid dates were found in the selected date column.")
    else:
        st.info("A recognizable date field and a numeric measure are needed for trend analysis.")

    st.divider()
    st.markdown("### 5. Numeric distribution explorer")
    if nums:
        distribution_col = st.selectbox("Choose a numeric column", nums, key="advanced_distribution_col")
        distribution_values = pd.to_numeric(filtered_df[distribution_col], errors="coerce").dropna()
        if not distribution_values.empty:
            d1, d2, d3, d4 = st.columns(4)
            d1.metric("Mean", f"{distribution_values.mean():,.2f}")
            d2.metric("Median", f"{distribution_values.median():,.2f}")
            d3.metric("Standard deviation", f"{distribution_values.std():,.2f}" if len(distribution_values) > 1 else "N/A")
            d4.metric("Distinct values", f"{distribution_values.nunique():,}")
            hist_col, box_col = st.columns(2)
            with hist_col:
                hist = px.histogram(filtered_df, x=distribution_col, nbins=30, title=f"Distribution of {distribution_col}", template="plotly_white", color_discrete_sequence=["#2563eb"])
                st.plotly_chart(hist, use_container_width=True)
            with box_col:
                box = px.box(filtered_df, y=distribution_col, points="outliers", title=f"Box plot of {distribution_col}", template="plotly_white", color_discrete_sequence=["#0f766e"])
                st.plotly_chart(box, use_container_width=True)
        else:
            st.info("No valid numeric values are available for this column.")
    else:
        st.info("No numeric columns are available for distribution analysis.")

    st.divider()
    st.markdown("### 6. Business Intelligence Studio — Stage 7.4")
    st.caption("Build a compact KPI scorecard, compare the latest periods, and identify category concentration. All calculations use the filtered data currently shown in this app.")

    if nums:
        bi_metric = st.selectbox("Primary business KPI", nums, key="stage74_bi_metric")
        bi_values = pd.to_numeric(filtered_df[bi_metric], errors="coerce").dropna()
        if not bi_values.empty:
            total_value = float(bi_values.sum())
            average_value = float(bi_values.mean())
            median_value = float(bi_values.median())
            maximum_value = float(bi_values.max())
            minimum_value = float(bi_values.min())
            bi_cards = st.columns(5)
            bi_cards[0].metric("Total", f"{total_value:,.2f}")
            bi_cards[1].metric("Average", f"{average_value:,.2f}")
            bi_cards[2].metric("Median", f"{median_value:,.2f}")
            bi_cards[3].metric("Highest", f"{maximum_value:,.2f}")
            bi_cards[4].metric("Lowest", f"{minimum_value:,.2f}")

            bi_summary = pd.DataFrame([
                {"KPI": "Total", "Metric": bi_metric, "Value": total_value},
                {"KPI": "Average", "Metric": bi_metric, "Value": average_value},
                {"KPI": "Median", "Metric": bi_metric, "Value": median_value},
                {"KPI": "Highest", "Metric": bi_metric, "Value": maximum_value},
                {"KPI": "Lowest", "Metric": bi_metric, "Value": minimum_value},
            ])

            if dcols:
                st.markdown("#### Period-over-period performance")
                period_controls = st.columns(3)
                with period_controls[0]:
                    bi_date_col = st.selectbox("Date field", dcols, key="stage74_bi_date")
                with period_controls[1]:
                    bi_period = st.selectbox("Period", ["Month", "Quarter", "Year"], key="stage74_bi_period")
                with period_controls[2]:
                    bi_aggregation = st.selectbox("KPI calculation", ["Sum", "Average", "Count"], key="stage74_bi_aggregation")

                period_data = filtered_df[[bi_date_col, bi_metric]].copy()
                period_data[bi_date_col] = pd.to_datetime(period_data[bi_date_col], errors="coerce")
                period_data[bi_metric] = pd.to_numeric(period_data[bi_metric], errors="coerce")
                period_data = period_data.dropna(subset=[bi_date_col])
                if not period_data.empty:
                    period_data = period_data.set_index(bi_date_col).sort_index()
                    period_rule = {"Month": "MS", "Quarter": "QS", "Year": "YS"}[bi_period]
                    period_group = period_data[bi_metric].resample(period_rule)
                    if bi_aggregation == "Average":
                        period_values = period_group.mean()
                    elif bi_aggregation == "Count":
                        period_values = period_group.count()
                    else:
                        period_values = period_group.sum(min_count=1)
                    period_table = period_values.dropna().rename("KPI value").reset_index()
                    if not period_table.empty:
                        period_table["Period"] = period_table[bi_date_col].dt.strftime("%Y-%m-%d")
                        if len(period_table) >= 2:
                            previous = float(period_table["KPI value"].iloc[-2])
                            latest = float(period_table["KPI value"].iloc[-1])
                            if previous != 0:
                                change_pct = (latest - previous) / abs(previous) * 100
                                change_label = f"{change_pct:+.2f}%"
                            else:
                                change_label = "N/A (previous period is zero)"
                            p1, p2, p3 = st.columns(3)
                            p1.metric("Latest period", f"{latest:,.2f}")
                            p2.metric("Previous period", f"{previous:,.2f}")
                            p3.metric("Change vs previous", change_label)
                            st.caption("The percentage compares the latest available period with the previous available period; it is not an annualized growth rate.")
                        else:
                            st.info("At least two periods with valid KPI values are needed for a period comparison.")
                        bi_trend_fig = px.bar(period_table, x="Period", y="KPI value", title=f"{bi_aggregation} {bi_metric} by {bi_period.lower()}", template="plotly_white", text_auto=".2s")
                        bi_trend_fig.update_layout(title_x=0.02, xaxis_title=bi_period, yaxis_title=f"{bi_aggregation} of {bi_metric}")
                        st.plotly_chart(bi_trend_fig, use_container_width=True)
                        st.dataframe(period_table.drop(columns=[bi_date_col]), use_container_width=True, hide_index=True)
                        st.download_button("⬇️ Download period performance (CSV)", period_table.to_csv(index=False).encode("utf-8-sig"), file_name="datawise_period_performance.csv", mime="text/csv", key="stage74_period_download")
                    else:
                        st.info("No usable period values were found for this date and KPI combination.")
                else:
                    st.info("No valid dates are available for period comparison.")

            if cats:
                st.markdown("#### Category contribution and concentration")
                bi_category = st.selectbox("Segment / category", cats, key="stage74_bi_category")
                category_metric_data = filtered_df[[bi_category, bi_metric]].copy()
                category_metric_data[bi_metric] = pd.to_numeric(category_metric_data[bi_metric], errors="coerce")
                contribution = category_metric_data.groupby(bi_category, dropna=False)[bi_metric].sum(min_count=1).dropna().sort_values(ascending=False).reset_index()
                if not contribution.empty:
                    contribution.columns = [str(bi_category), "KPI total"]
                    contribution["Contribution (%)"] = (contribution["KPI total"] / contribution["KPI total"].sum() * 100) if contribution["KPI total"].sum() != 0 else np.nan
                    contribution["Cumulative contribution (%)"] = contribution["Contribution (%)"].cumsum()
                    top_share = contribution["Contribution (%)"].head(3).sum()
                    b1, b2 = st.columns(2)
                    b1.metric("Number of segments", f"{len(contribution):,}")
                    b2.metric("Top 3 segment share", f"{top_share:.1f}%" if pd.notna(top_share) else "N/A")
                    pareto_fig = px.bar(contribution.head(20), x=str(bi_category), y="KPI total", title=f"Top segments by {bi_metric}", template="plotly_white", text_auto=".2s")
                    pareto_fig.update_layout(title_x=0.02, xaxis_title=str(bi_category), yaxis_title=f"Total {bi_metric}", xaxis_tickangle=-25)
                    st.plotly_chart(pareto_fig, use_container_width=True)
                    st.dataframe(contribution, use_container_width=True, hide_index=True)
                    st.download_button("⬇️ Download category contribution (CSV)", contribution.to_csv(index=False).encode("utf-8-sig"), file_name="datawise_category_contribution.csv", mime="text/csv", key="stage74_contribution_download")
                    if pd.notna(top_share) and top_share >= 80:
                        st.warning(f"The top three {bi_category} segments contribute {top_share:.1f}% of the selected KPI. This indicates concentration in the current filtered dataset.")
                    else:
                        st.info(f"The top three {bi_category} segments contribute {top_share:.1f}% of the selected KPI. Review the table to understand the distribution.")
                else:
                    st.info("No valid numeric category totals are available.")
        else:
            st.info("The selected KPI does not contain usable numeric values.")
    else:
        st.info("Upload a dataset with at least one numeric field to use the Business Intelligence Studio.")

    st.caption("Stage 7.4 provides descriptive business intelligence. Period comparisons and category contribution describe the available data; they do not establish causation or guarantee future performance.")

    st.caption("Stage 2 uses descriptive statistics and screening rules only. It does not prove causation or provide guaranteed forecasts.")

# =========================================================
# CUSTOM CHART BUILDER
# =========================================================
elif page == "✨ Business Insights":
    st.markdown("## ✨ Automatic Business Insights")
    st.caption("Findings are calculated from the currently filtered dataset. They are rule-based, not generated by a hosted AI model.")

    overview_rows = len(filtered_df)
    overview_cols = len(filtered_df.columns)
    missing_cells = int(filtered_df.isna().sum().sum())
    duplicate_rows = int(filtered_df.duplicated().sum())
    numeric_cols = numeric_columns(filtered_df)
    categorical_cols = [c for c in filtered_df.columns if c not in numeric_cols]

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Records", f"{overview_rows:,}")
    m2.metric("Columns", f"{overview_cols:,}")
    m3.metric("Missing cells", f"{missing_cells:,}")
    m4.metric("Duplicate rows", f"{duplicate_rows:,}")

    st.markdown("### 📝 Executive summary")
    for title, body in make_insights(filtered_df):
        st.markdown(f"**{title}:** {body}")

    # Stage 6.1: build a transparent, evidence-based narrative from calculated values.
    st.markdown("### 🤖 Stage 6.1 — Evidence-based insight narrative")
    st.caption(
        "This narrative is generated from calculations on the currently filtered dataset. "
        "It describes patterns but does not claim that one factor caused another."
    )
    narrative_lines = []
    narrative_lines.append(f"Dataset contains {overview_rows:,} rows and {overview_cols:,} columns.")
    total_cells = overview_rows * overview_cols
    missing_pct = (missing_cells / total_cells * 100) if total_cells else 0.0
    narrative_lines.append(
        f"Data quality: {missing_cells:,} missing cells ({missing_pct:.2f}% of all cells) "
        f"and {duplicate_rows:,} exact duplicate rows."
    )

    narrative_metric = None
    narrative_category = None
    if numeric_cols:
        likely_metrics = [c for c in numeric_cols if any(
            token in str(c).lower() for token in ("sales", "revenue", "profit", "amount", "total", "cost", "quantity", "price")
        )]
        metric_options = likely_metrics + [c for c in numeric_cols if c not in likely_metrics]
        narrative_metric = st.selectbox(
            "Metric for the narrative", metric_options, key="stage61_narrative_metric"
        )
        metric_series = pd.to_numeric(filtered_df[narrative_metric], errors="coerce").dropna()
        if metric_series.empty:
            narrative_lines.append(f"No valid numeric values were available for {narrative_metric}.")
        else:
            narrative_lines.append(
                f"{narrative_metric}: total {metric_series.sum():,.2f}; average {metric_series.mean():,.2f}; "
                f"median {metric_series.median():,.2f}; minimum {metric_series.min():,.2f}; "
                f"maximum {metric_series.max():,.2f} across {len(metric_series):,} valid values."
            )
            if len(metric_series) >= 2 and metric_series.mean() != 0:
                gap_pct = abs(metric_series.mean() - metric_series.median()) / abs(metric_series.mean()) * 100
                if gap_pct >= 25:
                    narrative_lines.append(
                        f"The mean and median of {narrative_metric} differ by {gap_pct:.1f}% of the absolute mean; "
                        "this may indicate a skewed distribution and is worth reviewing."
                    )
            q1, q3 = metric_series.quantile([0.25, 0.75])
            iqr = q3 - q1
            if iqr > 0:
                outlier_count = int(((metric_series < q1 - 1.5 * iqr) | (metric_series > q3 + 1.5 * iqr)).sum())
                narrative_lines.append(
                    f"IQR screening flagged {outlier_count:,} potential outlier(s) in {narrative_metric}; "
                    "these are candidates for review, not confirmed errors."
                )

        narrative_cats = [c for c in categorical_cols if 1 < filtered_df[c].nunique(dropna=True) <= 500]
        if narrative_cats:
            default_cat = next((c for c in narrative_cats if any(
                token in str(c).lower() for token in ("product", "region", "category", "department", "segment", "customer")
            )), narrative_cats[0])
            narrative_category = st.selectbox(
                "Optional category breakdown", narrative_cats,
                index=narrative_cats.index(default_cat), key="stage61_narrative_category"
            )
            if narrative_metric is not None:
                grouped_source = filtered_df[[narrative_category, narrative_metric]].copy()
                grouped_source[narrative_metric] = pd.to_numeric(grouped_source[narrative_metric], errors="coerce")
                grouped = grouped_source.groupby(narrative_category, dropna=False)[narrative_metric].sum(min_count=1).dropna().sort_values(ascending=False)
                if not grouped.empty:
                    top_name = "(missing)" if pd.isna(grouped.index[0]) else str(grouped.index[0])
                    bottom_name = "(missing)" if pd.isna(grouped.index[-1]) else str(grouped.index[-1])
                    narrative_lines.append(
                        f"By {narrative_category}, the highest total {narrative_metric} is {top_name} "
                        f"({grouped.iloc[0]:,.2f}); the lowest is {bottom_name} ({grouped.iloc[-1]:,.2f}). "
                        "This is a descriptive comparison and may reflect different group sizes."
                    )
    else:
        st.info("A numeric column is required for metric-level narrative insights.")

    for line in narrative_lines:
        st.markdown(f"- {line}")
    narrative_text = "DATAWISE AI — EVIDENCE-BASED INSIGHT NARRATIVE\n\n" + "\n".join(f"• {line}" for line in narrative_lines)
    st.download_button(
        "⬇️ Download insight narrative (TXT)",
        data=narrative_text.encode("utf-8"),
        file_name="datawise_insight_narrative.txt",
        mime="text/plain",
        key="stage61_narrative_txt_download",
    )
    narrative_df = pd.DataFrame({"Finding": narrative_lines})
    st.download_button(
        "⬇️ Download insight findings (CSV)",
        data=narrative_df.to_csv(index=False).encode("utf-8"),
        file_name="datawise_insight_findings.csv",
        mime="text/csv",
        key="stage61_narrative_csv_download",
    )

    st.markdown("### 📊 Numeric performance")
    if numeric_cols:
        numeric_summary = filtered_df[numeric_cols].describe().T
        numeric_summary["Missing values"] = filtered_df[numeric_cols].isna().sum()
        numeric_summary["Unique values"] = filtered_df[numeric_cols].nunique(dropna=True)
        st.dataframe(numeric_summary.round(2), use_container_width=True)
        st.download_button(
            "⬇️ Download numeric summary (CSV)",
            numeric_summary.round(2).to_csv(index=True).encode("utf-8"),
            file_name="datawise_numeric_summary.csv",
            mime="text/csv",
            key="business_numeric_summary_download",
        )
    else:
        st.info("No numeric columns were detected in the current dataset.")

    st.markdown("### 🏆 Top-performing category")
    if categorical_cols and numeric_cols:
        col_a, col_b = st.columns(2)
        with col_a:
            insight_category = st.selectbox("Category", categorical_cols, key="business_insight_category")
        with col_b:
            insight_metric = st.selectbox("Measure", numeric_cols, key="business_insight_metric")
        insight_aggregation = st.selectbox("Calculation", ["Sum", "Average", "Count", "Maximum", "Minimum"], key="business_insight_aggregation")

        work = filtered_df[[insight_category, insight_metric]].copy()
        work[insight_metric] = pd.to_numeric(work[insight_metric], errors="coerce")
        work = work.dropna(subset=[insight_category])
        grouped = work.groupby(insight_category, dropna=False)[insight_metric]
        if insight_aggregation == "Sum":
            ranked = grouped.sum(min_count=1)
        elif insight_aggregation == "Average":
            ranked = grouped.mean()
        elif insight_aggregation == "Count":
            ranked = grouped.count()
        elif insight_aggregation == "Maximum":
            ranked = grouped.max()
        else:
            ranked = grouped.min()
        ranked = ranked.dropna().sort_values(ascending=False).head(10)
        if not ranked.empty:
            ranked_df = ranked.rename(f"{insight_aggregation} of {insight_metric}").reset_index()
            st.dataframe(ranked_df, use_container_width=True)
            st.plotly_chart(px.bar(ranked_df, x=insight_category, y=f"{insight_aggregation} of {insight_metric}", title=f"Top categories by {insight_aggregation.lower()} of {insight_metric}", text_auto=".2s"), use_container_width=True)
            st.download_button(
                "⬇️ Download category findings (CSV)",
                ranked_df.to_csv(index=False).encode("utf-8"),
                file_name="datawise_category_insights.csv",
                mime="text/csv",
                key="business_category_download",
            )
        else:
            st.info("There are no valid numeric values to rank for this selection.")
    else:
        st.info("Category ranking requires at least one categorical column and one numeric column.")

    st.markdown("### ⚠️ Data-quality checks")
    quality_df = pd.DataFrame({
        "Check": ["Missing cells", "Exact duplicate rows", "Rows with any missing value"],
        "Count": [missing_cells, duplicate_rows, int(filtered_df.isna().any(axis=1).sum())],
    })
    st.dataframe(quality_df, use_container_width=True)

    outlier_rows = []
    for col in numeric_cols:
        values = pd.to_numeric(filtered_df[col], errors="coerce").dropna()
        if len(values) < 4:
            continue
        q1, q3 = values.quantile([0.25, 0.75])
        iqr = q3 - q1
        if pd.isna(iqr) or iqr == 0:
            continue
        low, high = q1 - 1.5 * iqr, q3 + 1.5 * iqr
        count = int(((values < low) | (values > high)).sum())
        if count:
            outlier_rows.append({"Column": col, "Potential outliers": count, "Lower IQR bound": round(float(low), 2), "Upper IQR bound": round(float(high), 2)})
    if outlier_rows:
        st.dataframe(pd.DataFrame(outlier_rows), use_container_width=True)
        st.caption("IQR flags are review candidates, not proof of incorrect data.")
    else:
        st.success("No potential numeric outliers were flagged by the IQR screening rule.")

    insights_export = pd.DataFrame([{"Finding": f"{title}: {body}"} for title, body in make_insights(filtered_df)])
    st.download_button(
        "⬇️ Download all business insights (CSV)",
        insights_export.to_csv(index=False).encode("utf-8"),
        file_name="datawise_business_insights.csv",
        mime="text/csv",
        key="business_insights_download",
    )

elif page == "🛠️ Chart Builder":
    st.markdown("## 🛠️ Chart builder & recommendations")
    st.write("Get data-driven chart suggestions, then customize and export your chart.")
    nums = numeric_columns(filtered_df)

    # Stage 6.3: recommend charts using column types, cardinality, and date-like fields.
    st.markdown("### 💡 Stage 6.3 — Automatic chart recommendations")
    if filtered_df.empty:
        st.info("Upload a dataset with rows to receive chart recommendations.")
    elif not nums:
        st.warning("No numeric columns are available. Try a dataset with numeric metrics such as sales, revenue, profit, or quantity.")
    else:
        all_cols = filtered_df.columns.tolist()
        cat_cols = [
            c for c in all_cols
            if c not in nums and 1 < filtered_df[c].nunique(dropna=True) <= 50
        ]
        date_cols = [
            c for c in all_cols
            if (
                pd.api.types.is_datetime64_any_dtype(filtered_df[c])
                or any(token in str(c).lower() for token in ["date", "time", "month", "year"])
            )
        ]

        recommendations = []
        primary_metric = nums[0]
        if date_cols:
            recommendations.append({
                "Chart": "Line chart",
                "Suggested X-axis": date_cols[0],
                "Suggested metric": primary_metric,
                "Why this chart": "Useful for viewing change over time when the field contains valid dates or time periods.",
                "Caution": "Check that the date/time field is parsed and ordered correctly."
            })
        if cat_cols:
            recommendations.append({
                "Chart": "Bar chart",
                "Suggested X-axis": cat_cols[0],
                "Suggested metric": primary_metric,
                "Why this chart": "Useful for comparing a numeric metric across a manageable number of categories.",
                "Caution": "High-cardinality categories may need filtering or a top-N view."
            })
            if filtered_df[cat_cols[0]].nunique(dropna=True) <= 8:
                recommendations.append({
                    "Chart": "Pie chart",
                    "Suggested X-axis": cat_cols[0],
                    "Suggested metric": primary_metric,
                    "Why this chart": "Can show parts of a whole when there are only a few categories and values represent a meaningful total.",
                    "Caution": "Use a bar chart when precise comparisons or many categories matter."
                })
        if len(nums) >= 2:
            recommendations.append({
                "Chart": "Scatter plot",
                "Suggested X-axis": nums[1],
                "Suggested metric": nums[0],
                "Why this chart": "Useful for exploring the relationship between two numeric variables.",
                "Caution": "An association does not establish causation."
            })
        recommendations.append({
            "Chart": "Histogram",
            "Suggested X-axis": primary_metric,
            "Suggested metric": primary_metric,
            "Why this chart": "Useful for inspecting the distribution and spread of a numeric metric.",
            "Caution": "The appearance depends on binning and data quality."
        })

        recommendation_df = pd.DataFrame(recommendations)
        st.dataframe(recommendation_df, use_container_width=True, hide_index=True)
        st.download_button(
            "⬇️ Download chart recommendations (CSV)",
            data=recommendation_df.to_csv(index=False).encode("utf-8-sig"),
            file_name="datawise_chart_recommendations.csv",
            mime="text/csv",
            key="stage_6_3_chart_recommendations_download",
        )
        st.caption("Recommendations are heuristic suggestions based on column types and category counts, not proof that a chart is statistically appropriate.")

    st.divider()
    st.markdown("### 🛠️ Customize your chart")
    if not nums:
        st.warning("This dataset has no numeric columns for the current chart builder.")
    else:
        left, right = st.columns(2)
        with left:
            x_col = st.selectbox("Category / X-axis", filtered_df.columns.tolist(), key="custom_x")
            chart_type = st.selectbox("Chart type", ["Bar chart", "Line chart", "Pie chart", "Scatter", "Histogram"], key="custom_type")
            aggregation = st.selectbox("Aggregation", ["Sum", "Average", "Count", "Minimum", "Maximum"], key="custom_aggregation")
        with right:
            y_col = st.selectbox("Numeric / Y-axis", nums, key="custom_y")
            chart_title = st.text_input("Chart title (optional)", placeholder="e.g. Revenue by region", key="custom_title")
            top_n = st.slider("Categories to show", 5, 50, 15, key="custom_topn")
        try:
            fig = create_chart(filtered_df, x_col, y_col, chart_type, aggregation, chart_title or None, top_n)
            st.plotly_chart(fig, use_container_width=True)
            st.download_button(
                "⬇️ Download chart as interactive HTML",
                fig.to_html(include_plotlyjs="cdn", full_html=True).encode("utf-8"),
                file_name="custom_chart.html",
                mime="text/html",
            )
        except Exception as exc:
            st.error(f"Unable to create this chart. Try a different column or chart type. Details: {exc}")


# =========================================================
# DATA CLEANING
# =========================================================
elif page == "🧹 Data Cleaning":
    st.markdown("## 🧹 Data cleaning")
    st.write("Review issues first. Cleaning actions create a cleaned copy and preserve the original dataset in memory.")
    missing, duplicates, completeness = data_quality(base_df)
    a, b, c = st.columns(3)
    a.metric("Missing cells", f"{missing:,}")
    b.metric("Duplicate rows", f"{duplicates:,}")
    c.metric("Completeness", f"{completeness:.1f}%")
    st.markdown("### Preview missing values")
    missing_by_col = base_df.isna().sum()
    st.dataframe(missing_by_col[missing_by_col > 0].rename("Missing values").to_frame(), use_container_width=True)
    st.markdown("### Choose cleaning actions")
    fill_missing = st.checkbox("Fill missing numeric values with each column's median")
    drop_duplicates = st.checkbox("Remove exact duplicate rows")
    trim_text = st.checkbox("Trim leading/trailing spaces in text columns")
    standardize_names = st.checkbox("Standardize column names (spaces become underscores)")
    if st.button("Apply selected cleaning actions", type="primary"):
        cleaned = base_df.copy()
        if fill_missing:
            for col in numeric_columns(cleaned):
                if cleaned[col].isna().any():
                    cleaned[col] = cleaned[col].fillna(cleaned[col].median())
        if drop_duplicates:
            cleaned = cleaned.drop_duplicates().copy()
        if trim_text:
            for col in cleaned.select_dtypes(include=["object", "string"]).columns:
                cleaned[col] = cleaned[col].map(lambda x: x.strip() if isinstance(x, str) else x)
        if standardize_names:
            cleaned = clean_column_names(cleaned)
        st.session_state.cleaned_df = cleaned
        st.success(f"Cleaning applied to a copy. New shape: {cleaned.shape[0]:,} rows × {cleaned.shape[1]:,} columns.")
        st.dataframe(cleaned.head(20), use_container_width=True)
    if st.session_state.cleaned_df is not None:
        st.download_button(
            "⬇️ Download cleaned dataset (CSV)",
            st.session_state.cleaned_df.to_csv(index=False).encode("utf-8"),
            file_name="cleaned_dataset.csv",
            mime="text/csv",
        )
        if st.button("Discard cleaned copy and restore original"):
            st.session_state.cleaned_df = None
            st.rerun()

# =========================================================
# DATAWISE AI — STAGE 4.4
# ADVANCED NATURAL-LANGUAGE DATA ASSISTANT
# =========================================================
elif page == "💬 Data Assistant":
    st.markdown("## 💬 Data Assistant")
    st.caption(
        "Ask questions in plain English. Data is analyzed locally with Python "
        "and pandas; no paid AI API is required."
    )
    st.markdown("### 🚀 Stage 6.2 — Advanced natural-language queries")
    st.info(
        "Ask a business question, review the calculated result, and revisit recent questions. "
        "Numbers are calculated from the currently filtered dataset; the assistant does not "
        "infer causes that the data cannot establish."
    )

    # Use the current dashboard-filtered data so answers match selected filters.
    assistant_df = filtered_df.copy()
    info1, info2, info3 = st.columns(3)
    info1.metric("Rows available", f"{len(assistant_df):,}")
    info2.metric("Columns available", f"{len(assistant_df.columns):,}")
    info3.metric("Missing cells", f"{int(assistant_df.isna().sum().sum()):,}")

    with st.expander("Preview data and available columns"):
        st.write([str(c) for c in assistant_df.columns])
        st.dataframe(assistant_df.head(10), use_container_width=True, hide_index=True)

    question = st.text_input(
        "Ask a question about your data",
        placeholder="Example: Show the top 5 products by sales",
        key="data_assistant_question_stage44",
    )
    st.caption(
        "Try: total sales · top 5 products by sales · sales above 50000 and region is North · "
        "how many unique customers · correlation between sales and profit · monthly sales trends"
    )

    if st.button("🔍 Analyze question", type="primary", key="data_assistant_analyze_stage44"):
        import re

        if not question.strip():
            st.warning("Please enter a question first.")
        elif assistant_df.empty:
            st.warning("There is no data available to analyze.")
        else:
            q = question.strip().lower()
            nums = numeric_columns(assistant_df)
            dcols = date_columns(assistant_df)
            cats = [c for c in assistant_df.columns if c not in nums and c not in dcols and assistant_df[c].nunique(dropna=True) > 1]
            answer_text = None
            result_table = None
            chart_type = None
            chart_title = None

            def column_mentioned(columns, text):
                # Match complete column names, allowing underscores as spaces.
                for col in sorted(columns, key=lambda x: len(str(x)), reverse=True):
                    name = str(col).strip().lower()
                    variants = {name, name.replace("_", " "), name.replace(" ", "_")}
                    for variant in variants:
                        if variant and re.search(r"(?<!\w)" + re.escape(variant) + r"(?!\w)", text):
                            return col
                return None

            metric_col = column_mentioned(nums, q)
            category_col = column_mentioned(cats, q)

            # Recognize common business terms when a matching numeric column exists.
            if metric_col is None:
                aliases = {
                    "sales": ("sales", "sale", "revenue", "turnover"),
                    "profit": ("profit", "earnings", "margin"),
                    "quantity": ("quantity", "units", "items sold"),
                    "price": ("price", "amount", "cost"),
                }
                for col in nums:
                    col_name = str(col).lower().replace("_", " ")
                    if any(any(term in q for term in terms) and any(term in col_name for term in terms) for terms in aliases.values()):
                        metric_col = col
                        break

            # A. Dataset quality and shape questions.
            if any(term in q for term in ("how many rows", "number of rows", "total rows", "row count", "how many records", "record count")):
                answer_text = f"The dataset contains {len(assistant_df):,} rows."
            elif any(term in q for term in ("how many columns", "number of columns", "column count", "total columns")):
                answer_text = f"The dataset contains {len(assistant_df.columns):,} columns."
            elif "duplicate" in q:
                answer_text = f"The dataset contains {int(assistant_df.duplicated().sum()):,} exact duplicate rows."
            elif any(term in q for term in ("missing", "null values", "null count")):
                result_table = assistant_df.isna().sum().rename("Missing Values").to_frame().query("`Missing Values` > 0").reset_index(names="Column")
                answer_text = "No missing values were found." if result_table.empty else "Missing values by column are shown below."

            # B. Stage 6.2: distinct-value questions.
            elif any(term in q for term in ("unique", "distinct", "different values", "how many values")):
                all_cols = list(assistant_df.columns)
                chosen_col = column_mentioned(all_cols, q)
                if chosen_col is None:
                    answer_text = "Please name a column, e.g. 'How many unique customers are there?'."
                else:
                    distinct_values = assistant_df[chosen_col].dropna().drop_duplicates()
                    answer_text = f"Column {chosen_col} contains {len(distinct_values):,} distinct non-missing values."
                    if any(term in q for term in ("show", "list", "what are", "which")):
                        result_table = distinct_values.astype("string").sort_values().rename("Distinct values").to_frame().reset_index(drop=True)

            # C. Stage 6.2: correlation between two named numeric columns.
            elif "correlation" in q or "correlate" in q:
                named_metrics = [col for col in sorted(nums, key=lambda x: len(str(x)), reverse=True)
                                 if re.search(r"(?<!\w)" + re.escape(str(col).lower().replace("_", " ")) + r"(?!\w)", q)
                                 or re.search(r"(?<!\w)" + re.escape(str(col).lower()) + r"(?!\w)", q)]
                named_metrics = list(dict.fromkeys(named_metrics))
                if len(named_metrics) < 2:
                    answer_text = "Please name two numeric columns, e.g. 'Correlation between Sales and Profit'."
                else:
                    left, right = named_metrics[:2]
                    pair = assistant_df[[left, right]].apply(pd.to_numeric, errors="coerce").dropna()
                    if len(pair) < 2:
                        answer_text = "At least two rows with valid values in both columns are needed to calculate correlation."
                    else:
                        corr = pair[left].corr(pair[right])
                        answer_text = f"Pearson correlation between {left} and {right}: {corr:.3f} (based on {len(pair):,} complete rows). Correlation does not prove causation."
                        result_table = pair.corr().rename_axis("Metric").reset_index()

            # D. Stage 6.2: multiple conditions joined with AND.
            elif " and " in q and metric_col is not None and re.search(r"\b(above|greater than|more than|over|below|less than|under|equal to|equals)\b", q):
                work = assistant_df.copy()
                applied_conditions = []
                # Apply each numeric comparison only when its column name occurs in the relevant query.
                condition_pattern = re.compile(r"(?P<col>[a-zA-Z_][\w ]*?)\s+(?P<op>above|greater than|more than|over|below|less than|under|equal to|equals)\s+(?P<num>-?\d[\d,]*(?:\.\d+)?)", re.I)
                for match in condition_pattern.finditer(question.strip()):
                    raw_col = match.group("col").strip(" ,.?'")
                    col = column_mentioned(nums, raw_col.lower()) or column_mentioned(nums, q)
                    # Prefer a metric explicitly mentioned in the local condition text.
                    local_col = column_mentioned(nums, raw_col.lower())
                    if local_col is not None:
                        col = local_col
                    if col is None:
                        continue
                    op_word = match.group("op").lower()
                    threshold = float(match.group("num").replace(",", ""))
                    series = pd.to_numeric(work[col], errors="coerce")
                    if op_word in ("above", "greater than", "more than", "over"):
                        work = work.loc[series > threshold]
                        symbol = ">"
                    elif op_word in ("below", "less than", "under"):
                        work = work.loc[series < threshold]
                        symbol = "<"
                    else:
                        work = work.loc[series == threshold]
                        symbol = "="
                    applied_conditions.append(f"{col} {symbol} {threshold:g}")
                # Apply an exact category value if one appears in the question.
                for cat_col in cats:
                    for cat_value in assistant_df[cat_col].dropna().unique():
                        value_text = str(cat_value).strip()
                        if value_text and re.search(r"(?<!\w)" + re.escape(value_text.lower()) + r"(?!\w)", q):
                            work = work[work[cat_col].astype("string").str.strip().str.lower() == value_text.lower()]
                            applied_conditions.append(f"{cat_col} = {value_text}")
                            break
                if len(applied_conditions) >= 2:
                    result_table = work
                    answer_text = f"Found {len(result_table):,} records matching all detected conditions: " + "; ".join(applied_conditions) + "."
                else:
                    answer_text = "I detected a multi-condition question but couldn't identify at least two clear conditions. Try naming each column and value explicitly."

            # E. Numeric threshold filtering.
            elif metric_col is not None and re.search(r"\b(above|greater than|more than|over|below|less than|under|equal to|equals)\b", q):
                number_match = re.search(r"-?\d[\d,]*(?:\.\d+)?", q)
                if not number_match:
                    answer_text = "Please include a numeric threshold, e.g. 'Show sales above 50000'."
                else:
                    threshold = float(number_match.group().replace(",", ""))
                    if re.search(r"\b(above|greater than|more than|over)\b", q):
                        mask = pd.to_numeric(assistant_df[metric_col], errors="coerce") > threshold
                        op = ">"
                    elif re.search(r"\b(below|less than|under)\b", q):
                        mask = pd.to_numeric(assistant_df[metric_col], errors="coerce") < threshold
                        op = "<"
                    else:
                        mask = pd.to_numeric(assistant_df[metric_col], errors="coerce") == threshold
                        op = "="
                    result_table = assistant_df.loc[mask].copy()
                    answer_text = f"Found {len(result_table):,} records where {metric_col} {op} {threshold:g}."

            # C. Top / bottom N, or highest/lowest individual record.
            elif metric_col is not None and any(term in q for term in ("top ", "bottom ", "highest", "lowest", "best", "worst")):
                n_match = re.search(r"\b(top|bottom)\s+(\d+)\b", q)
                n = max(1, min(int(n_match.group(2)), 100)) if n_match else 1
                ascending = any(term in q for term in ("bottom", "lowest", "worst"))
                result_table = assistant_df.assign(**{str(metric_col): pd.to_numeric(assistant_df[metric_col], errors="coerce")}).sort_values(metric_col, ascending=ascending, na_position="last").head(n)
                direction = "lowest" if ascending else "highest"
                answer_text = f"Showing {len(result_table):,} record(s) with the {direction} {metric_col}."

            # D. Grouping and comparison, e.g. total sales by product.
            elif metric_col is not None and category_col is not None and any(term in q for term in (" by ", " each ", " per ", "compare", "group")):
                if "average" in q or "mean" in q:
                    grouped = assistant_df.groupby(category_col, dropna=False)[metric_col].mean()
                    label = f"Average {metric_col}"
                elif "count" in q:
                    grouped = assistant_df.groupby(category_col, dropna=False)[metric_col].count()
                    label = f"Count of {metric_col}"
                elif "minimum" in q or "lowest" in q:
                    grouped = assistant_df.groupby(category_col, dropna=False)[metric_col].min()
                    label = f"Minimum {metric_col}"
                elif "maximum" in q or "highest" in q:
                    grouped = assistant_df.groupby(category_col, dropna=False)[metric_col].max()
                    label = f"Maximum {metric_col}"
                else:
                    grouped = assistant_df.groupby(category_col, dropna=False)[metric_col].sum(min_count=1)
                    label = f"Total {metric_col}"
                result_table = grouped.sort_values(ascending=False, na_position="last").rename(label).reset_index()
                result_table[category_col] = result_table[category_col].astype("string").fillna("(missing)")
                chart_type = "bar"
                chart_title = f"{label} by {category_col}"
                answer_text = f"Calculated {label.lower()} for each {category_col}."

            # E. Monthly trend by date and measure.
            elif ("monthly" in q or "month" in q or "trend" in q) and dcols:
                if metric_col is None and len(nums) == 1:
                    metric_col = nums[0]
                if metric_col is None:
                    answer_text = "Please specify a numeric metric, e.g. 'Show monthly sales trends'."
                else:
                    date_col = dcols[0]
                    temp = assistant_df[[date_col, metric_col]].copy()
                    temp[date_col] = pd.to_datetime(temp[date_col], errors="coerce")
                    temp[metric_col] = pd.to_numeric(temp[metric_col], errors="coerce")
                    temp = temp.dropna(subset=[date_col])
                    temp["Month"] = temp[date_col].dt.to_period("M").astype(str)
                    result_table = temp.groupby("Month", as_index=False)[metric_col].sum(min_count=1)
                    chart_type = "line"
                    chart_title = f"Monthly trend: {metric_col}"
                    answer_text = f"Monthly {metric_col} trend calculated using {date_col}."

            # F. Totals, averages, medians, minima, maxima.
            elif any(term in q for term in ("total", "sum", "average", "mean", "median", "maximum", "minimum", "max value", "min value")):
                if metric_col is None:
                    if len(nums) == 1:
                        metric_col = nums[0]
                    elif len(nums) > 1:
                        answer_text = "Several numeric columns are available. Please specify one, e.g. 'Calculate total Sales'."
                    else:
                        answer_text = "No numeric columns are available for calculations."
                if metric_col is not None and answer_text is None:
                    series = pd.to_numeric(assistant_df[metric_col], errors="coerce").dropna()
                    if series.empty:
                        answer_text = f"No valid numeric values were found in {metric_col}."
                    elif "average" in q or "mean" in q:
                        answer_text = f"Average {metric_col}: {series.mean():,.2f}"
                    elif "median" in q:
                        answer_text = f"Median {metric_col}: {series.median():,.2f}"
                    elif "maximum" in q or "max value" in q:
                        answer_text = f"Maximum {metric_col}: {series.max():,.2f}"
                    elif "minimum" in q or "min value" in q:
                        answer_text = f"Minimum {metric_col}: {series.min():,.2f}"
                    else:
                        answer_text = f"Total {metric_col}: {series.sum():,.2f}"

            # G. Category filtering. Match whole values, not partial words.
            elif cats:
                chosen_col = category_col
                chosen_value = None
                search_cols = [chosen_col] if chosen_col is not None else cats
                for col in search_cols:
                    for value in assistant_df[col].dropna().unique():
                        value_text = str(value).strip()
                        if value_text and re.search(r"(?<!\w)" + re.escape(value_text.lower()) + r"(?!\w)", q):
                            chosen_col, chosen_value = col, value
                            break
                    if chosen_value is not None:
                        break
                if chosen_value is not None:
                    result_table = assistant_df[assistant_df[chosen_col].astype("string").str.strip().str.lower() == str(chosen_value).strip().lower()].copy()
                    answer_text = f"Found {len(result_table):,} records where {chosen_col} is {chosen_value}."
                else:
                    # Keep existing question-answering module as a fallback for supported questions.
                    try:
                        fallback = answer_question(assistant_df, question.strip())
                        if isinstance(fallback, dict):
                            answer_text = str(fallback.get("answer", "I couldn't confidently interpret that question."))
                            if isinstance(fallback.get("table"), pd.DataFrame):
                                result_table = fallback["table"]
                        else:
                            answer_text = str(fallback)
                    except Exception:
                        answer_text = "I couldn't confidently interpret that question. Try one of the suggested examples below."

            else:
                try:
                    fallback = answer_question(assistant_df, question.strip())
                    if isinstance(fallback, dict):
                        answer_text = str(fallback.get("answer", "I couldn't interpret that question."))
                        if isinstance(fallback.get("table"), pd.DataFrame):
                            result_table = fallback["table"]
                    else:
                        answer_text = str(fallback)
                except Exception:
                    answer_text = "I couldn't confidently interpret that question. Please try one of the suggested examples below."

            st.divider()
            st.markdown("### 📊 Analysis results")
            if answer_text:
                st.success(answer_text)
            if isinstance(result_table, pd.DataFrame):
                if result_table.empty:
                    st.info("No matching records or result rows were found.")
                else:
                    st.dataframe(result_table, use_container_width=True, hide_index=True)
                    st.download_button(
                        "⬇️ Download analysis results (CSV)",
                        data=result_table.to_csv(index=False).encode("utf-8"),
                        file_name="datawise_analysis_results.csv",
                        mime="text/csv",
                        key="data_assistant_download_stage44",
                    )
                    if chart_type in ("bar", "line") and len(result_table.columns) >= 2:
                        numeric_result_cols = result_table.select_dtypes(include=np.number).columns.tolist()
                        category_result_cols = [c for c in result_table.columns if c not in numeric_result_cols]
                        if numeric_result_cols and category_result_cols:
                            try:
                                if chart_type == "line":
                                    fig = px.line(result_table, x=category_result_cols[0], y=numeric_result_cols[0], markers=True, title=chart_title, template="plotly_white")
                                else:
                                    fig = px.bar(result_table, x=category_result_cols[0], y=numeric_result_cols[0], title=chart_title, template="plotly_white", text_auto=".2s", color_discrete_sequence=["#2563eb"])
                                fig.update_layout(title_x=0.02, margin=dict(l=18, r=18, t=65, b=22))
                                st.plotly_chart(fig, use_container_width=True)
                            except Exception as chart_exc:
                                st.caption(f"Chart unavailable: {chart_exc}")

        # Stage 6.0: keep a short, session-only history of completed questions.
        if answer_text:
            if "datawise_question_history" not in st.session_state:
                st.session_state.datawise_question_history = []
            history_entry = {
                "Question": question.strip(),
                "Answer": str(answer_text),
                "Rows in result": len(result_table) if isinstance(result_table, pd.DataFrame) else None,
            }
            history = st.session_state.datawise_question_history
            if not history or history[-1]["Question"] != history_entry["Question"]:
                history.append(history_entry)
            st.session_state.datawise_question_history = history[-20:]

    st.divider()
    with st.expander("Suggested questions", expanded=True):
        st.markdown(
            """
            **Dataset and quality**
            - How many rows are there?
            - How many columns are there?
            - Show missing values.
            - How many duplicate rows exist?

            **Calculations**
            - Calculate total sales.
            - Calculate average profit.
            - What is the maximum sales value?
            - What is the minimum profit?

            **Category and ranking**
            - Show the top 5 records by sales.
            - Which product has the highest sales?
            - Show total sales by Product.
            - Compare average profit by Region.
            - Show Laptop records.
            - Show sales above 50000.

            **Trends**
            - Show monthly sales trends.

            **Advanced queries (Stage 6.2)**
            - Show sales above 50000 and Region North.
            - How many unique customers are there?
            - Correlation between Sales and Profit.
            """
        )
    with st.expander("🕘 Recent questions in this session", expanded=False):
        question_history = st.session_state.get("datawise_question_history", [])
        if question_history:
            history_df = pd.DataFrame(question_history[::-1])
            st.dataframe(history_df, use_container_width=True, hide_index=True)
            st.download_button(
                "⬇️ Download question history (CSV)",
                data=history_df.to_csv(index=False).encode("utf-8"),
                file_name="datawise_question_history.csv",
                mime="text/csv",
                key="datawise_question_history_download",
            )
            if st.button("Clear question history", key="datawise_clear_question_history"):
                st.session_state.datawise_question_history = []
                st.rerun()
        else:
            st.caption("Completed questions will appear here during this app session.")

    st.caption("Free-first mode: calculations are performed locally with Python. The assistant supports common question patterns; it is not a general-purpose language model.")

# =========================================================
# STAGE 5.1 — REPORT GENERATOR WORKSPACE
# =========================================================
elif page == "🧾 Report Generator":
    st.markdown("## 🧾 Report Generator")
    st.caption("Configure a business report, review the live dataset summary, and export it. Calculations are based on the current filtered dataset.")

    source_name = st.session_state.source_name or "uploaded_dataset"
    default_title = Path(source_name).stem.replace("_", " ").replace("-", " ").title() + " — Business Report"

    left, right = st.columns([3, 2])
    with left:
        report_title = st.text_input("Report title", value=default_title, key="report_generator_title")
        company_name = st.text_input("Company / organization name (optional)", value="", key="report_generator_company")
    with right:
        report_date = st.date_input("Report date", value=date.today(), key="report_generator_date")
        reporting_period = st.text_input("Reporting period (optional)", placeholder="e.g. Jan–Mar 2026", key="report_generator_period")

    st.markdown("### 🎨 Stage 7.3 — Report appearance and branding")
    theme_name = st.selectbox(
        "Choose a report theme",
        ["Professional Blue", "Executive Green", "Modern Purple", "Minimal Gray"],
        index=0,
        help="The selected theme is applied to generated PDF and PowerPoint reports.",
        key="rg_report_theme",
    )
    st.caption(f"Selected design: **{theme_name}**. Your report title, organization name, date, period, and chosen sections will be included in exports.")

    st.markdown("### Choose report sections")
    section_cols = st.columns(3)
    with section_cols[0]:
        include_overview = st.checkbox("Dataset overview and quality", value=True, key="rg_overview")
        include_kpis = st.checkbox("Numeric KPIs", value=True, key="rg_kpis")
    with section_cols[1]:
        include_insights = st.checkbox("Key findings", value=True, key="rg_insights")
        include_charts = st.checkbox("Charts and trends", value=True, key="rg_charts")
    with section_cols[2]:
        include_categories = st.checkbox("Category performance", value=True, key="rg_categories")
        include_recommendations = st.checkbox("Recommendations / next steps", value=True, key="rg_recommendations")

    missing, duplicates, completeness = data_quality(filtered_df)
    nums = numeric_columns(filtered_df)
    insights = make_insights(filtered_df)

    st.markdown("### Live report preview")
    st.markdown(f"# {report_title.strip() or 'Untitled business report'}")
    if company_name.strip():
        st.markdown(f"**{company_name.strip()}**")
    meta_parts = [f"Report date: {report_date.isoformat()}", f"Source: {source_name}"]
    if reporting_period.strip():
        meta_parts.append(f"Period: {reporting_period.strip()}")
    st.caption(" · ".join(meta_parts))

    if include_overview:
        st.markdown("#### Dataset overview and quality")
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Records", f"{len(filtered_df):,}")
        k2.metric("Columns", f"{len(filtered_df.columns):,}")
        k3.metric("Missing cells", f"{missing:,}")
        k4.metric("Completeness", f"{completeness:.1f}%")
        st.caption(f"Exact duplicate rows: {duplicates:,}. Preview uses the dataset after dashboard filters.")

    if include_kpis:
        st.markdown("#### Numeric KPIs")
        if nums:
            selected_metrics = st.multiselect("Choose numeric metrics", nums, default=nums[:min(4, len(nums))], key="rg_metrics")
            if selected_metrics:
                kpi_rows = []
                for col in selected_metrics:
                    values = pd.to_numeric(filtered_df[col], errors="coerce").dropna()
                    kpi_rows.append({"Metric": str(col), "Total": values.sum(), "Average": values.mean(), "Minimum": values.min(), "Maximum": values.max(), "Valid values": int(values.size)})
                st.dataframe(pd.DataFrame(kpi_rows), use_container_width=True, hide_index=True)
        else:
            st.info("No numeric columns are available for KPI calculations.")

    if include_categories:
        st.markdown("#### Category performance")
        cats = [c for c in filtered_df.columns if c not in nums and c not in date_columns(filtered_df) and 1 < filtered_df[c].nunique(dropna=True) <= 100]
        if cats and nums:
            cat_col = st.selectbox("Category", cats, key="rg_category")
            metric_col = st.selectbox("Measure", nums, key="rg_category_metric")
            temp = filtered_df[[cat_col, metric_col]].copy()
            temp[metric_col] = pd.to_numeric(temp[metric_col], errors="coerce")
            summary = temp.groupby(cat_col, dropna=False)[metric_col].sum(min_count=1).dropna().sort_values(ascending=False).head(10).rename("Total").reset_index()
            st.dataframe(summary, use_container_width=True, hide_index=True)
            if not summary.empty:
                st.plotly_chart(px.bar(summary, x=cat_col, y="Total", title=f"Top {cat_col} by {metric_col}", template="plotly_white"), use_container_width=True)
        else:
            st.info("Category performance needs at least one numeric column and one suitable category column.")

    if include_insights:
        st.markdown("#### Key findings")
        for heading, body in insights:
            st.markdown(f"**{heading}:** {body}")

    if include_charts:
        st.markdown("#### Chart preview")
        chart_items = make_report_charts(filtered_df, max_charts=4) if MATPLOTLIB_AVAILABLE else []
        if chart_items:
            for chart_title, chart_bytes in chart_items:
                st.markdown(f"**{chart_title}**")
                st.image(chart_bytes, use_container_width=True)
        else:
            st.info("No suitable report charts were detected, or Matplotlib is not installed.")

    if include_recommendations:
        st.markdown("#### Recommended next steps")
        if missing:
            st.write("- Review missing values and decide whether to fill, correct, or retain them.")
        if duplicates:
            st.write("- Review duplicate rows before using the dataset for final business decisions.")
        if nums:
            st.write("- Validate the largest and smallest values in the key numeric measures against the original source.")
        if not missing and not duplicates:
            st.write("- Data-quality checks found no missing cells or exact duplicate rows; continue with business-specific validation.")
        st.caption("Recommendations are rule-based prompts for review, not guaranteed business conclusions.")

    # =========================================================
    # STAGE 5.2 — AUTOMATED BUSINESS ANALYSIS ENGINE
    # =========================================================
    st.divider()
    st.markdown("### 🤖 Automated business analysis")
    st.caption("Automatically select useful fields, calculate comparisons and trends, and create evidence-based findings from the current filtered data.")

    auto_nums = numeric_columns(filtered_df)
    auto_dates = date_columns(filtered_df)
    auto_cats = [c for c in filtered_df.columns if c not in auto_nums and c not in auto_dates and 1 < filtered_df[c].nunique(dropna=True) <= 200]

    def _likely_column(columns, keywords, fallback=None):
        return next((c for c in columns if any(word in str(c).lower() for word in keywords)), fallback if fallback in columns else (columns[0] if columns else None))

    auto_metric = _likely_column(auto_nums, ("sales", "revenue", "amount", "profit", "total", "cost"))
    auto_category = _likely_column(auto_cats, ("product", "region", "category", "segment", "department", "customer"))
    auto_date = auto_dates[0] if auto_dates else None
    automated_findings = []
    automated_tables = []

    if not auto_nums:
        st.info("Automatic business analysis needs at least one numeric column. Dataset quality findings are still available above.")
    else:
        # Let the user override sensible automatic field choices.
        a1, a2, a3 = st.columns(3)
        with a1:
            auto_metric = st.selectbox("Primary business metric", auto_nums, index=auto_nums.index(auto_metric) if auto_metric in auto_nums else 0, key="rg_auto_metric")
        with a2:
            category_options = ["(No category)"] + auto_cats
            chosen_cat = st.selectbox("Compare by category", category_options, index=(category_options.index(auto_category) if auto_category in category_options else 0), key="rg_auto_category")
            auto_category = None if chosen_cat == "(No category)" else chosen_cat
        with a3:
            date_options = ["(No date trend)"] + auto_dates
            chosen_date = st.selectbox("Trend date column", date_options, index=(date_options.index(auto_date) if auto_date in date_options else 0), key="rg_auto_date")
            auto_date = None if chosen_date == "(No date trend)" else chosen_date

        metric_values = pd.to_numeric(filtered_df[auto_metric], errors="coerce").dropna()
        if metric_values.empty:
            st.warning(f"The selected metric '{auto_metric}' contains no valid numeric values.")
        else:
            total_value = float(metric_values.sum())
            average_value = float(metric_values.mean())
            median_value = float(metric_values.median())
            min_value = float(metric_values.min())
            max_value = float(metric_values.max())
            q1, q3 = metric_values.quantile([0.25, 0.75])
            iqr = q3 - q1
            outlier_count = int(((metric_values < q1 - 1.5 * iqr) | (metric_values > q3 + 1.5 * iqr)).sum()) if pd.notna(iqr) and iqr > 0 else 0

            m1, m2, m3, m4 = st.columns(4)
            m1.metric(f"Total {auto_metric}", f"{total_value:,.2f}")
            m2.metric(f"Average {auto_metric}", f"{average_value:,.2f}")
            m3.metric("Median", f"{median_value:,.2f}")
            m4.metric("Potential outliers", f"{outlier_count:,}")

            automated_findings.append(("Primary metric", f"{auto_metric}: total {total_value:,.2f}; average {average_value:,.2f}; median {median_value:,.2f}; range {min_value:,.2f} to {max_value:,.2f}."))
            if outlier_count:
                automated_findings.append(("Value review", f"{outlier_count:,} value(s) fall outside the 1.5×IQR screening range for {auto_metric}. These are candidates for review, not confirmed errors."))

            if auto_category:
                cat_temp = filtered_df[[auto_category, auto_metric]].copy()
                cat_temp[auto_metric] = pd.to_numeric(cat_temp[auto_metric], errors="coerce")
                category_summary = cat_temp.groupby(auto_category, dropna=False)[auto_metric].agg(Total="sum", Average="mean", Records="count").reset_index()
                category_summary = category_summary.sort_values("Total", ascending=False, na_position="last")
                automated_tables.append(("Category performance", category_summary))
                valid_groups = category_summary.dropna(subset=["Total"])
                if not valid_groups.empty:
                    leader = valid_groups.iloc[0]
                    automated_findings.append(("Leading category", f"{auto_category} '{safe_text(leader[auto_category], 80)}' has the highest total {auto_metric} ({leader['Total']:,.2f}) among the groups shown."))
                    st.markdown(f"#### Performance by {auto_category}")
                    st.dataframe(category_summary, use_container_width=True, hide_index=True)
                    st.plotly_chart(px.bar(category_summary.head(15), x=auto_category, y="Total", title=f"Total {auto_metric} by {auto_category}", template="plotly_white"), use_container_width=True)

            if auto_date:
                trend_temp = filtered_df[[auto_date, auto_metric]].copy()
                trend_temp[auto_date] = pd.to_datetime(trend_temp[auto_date], errors="coerce")
                trend_temp[auto_metric] = pd.to_numeric(trend_temp[auto_metric], errors="coerce")
                trend_temp = trend_temp.dropna()
                if not trend_temp.empty:
                    trend_temp["Period"] = trend_temp[auto_date].dt.to_period("M").astype(str)
                    trend_summary = trend_temp.groupby("Period", as_index=False)[auto_metric].sum().sort_values("Period")
                    automated_tables.append(("Time trend", trend_summary))
                    st.markdown(f"#### Monthly trend for {auto_metric}")
                    st.plotly_chart(px.line(trend_summary, x="Period", y=auto_metric, markers=True, title=f"Monthly {auto_metric}", template="plotly_white"), use_container_width=True)
                    if len(trend_summary) >= 2:
                        first = float(trend_summary[auto_metric].iloc[0])
                        last = float(trend_summary[auto_metric].iloc[-1])
                        change = last - first
                        pct = (change / abs(first) * 100) if first != 0 else None
                        pct_text = f" ({pct:+.1f}%)" if pct is not None else " (percentage change unavailable because the first period is zero)"
                        direction = "increased" if change > 0 else "decreased" if change < 0 else "was unchanged"
                        automated_findings.append(("Period comparison", f"From the first to the last displayed month, {auto_metric} {direction} by {abs(change):,.2f}{pct_text}. This compares endpoints, not every period."))
                else:
                    st.info("The selected date column has no valid date/metric pairs for trend analysis.")

            st.markdown("#### Automatically generated findings")
            for finding_title, finding_body in automated_findings:
                safe_title = html.escape(safe_text(finding_title), quote=True)
                safe_body = html.escape(safe_text(finding_body, 500), quote=True)
                st.markdown(
                    f'<div class="insight-card"><h4>{safe_title}</h4><p>{safe_body}</p></div>',
                    unsafe_allow_html=True,
                )

            if automated_tables:
                export_rows = []
                for table_title, table_frame in automated_tables:
                    for _, record in table_frame.iterrows():
                        export_rows.append({"Analysis": table_title, **{str(k): (None if pd.isna(v) else v) for k, v in record.to_dict().items()}})
                if export_rows:
                    analysis_csv = pd.DataFrame(export_rows).to_csv(index=False).encode("utf-8")
                    st.download_button("⬇️ Download automated analysis tables (CSV)", analysis_csv, file_name="datawise_automated_analysis.csv", mime="text/csv", key="rg_auto_analysis_csv")

    # =========================================================
    # STAGE 5.6 — EVIDENCE-BASED BUSINESS RECOMMENDATIONS
    # =========================================================
    st.divider()
    st.markdown("### 💡 Stage 5.6 — AI-style business recommendations")
    st.caption("Creates explainable, rule-based recommendations from the current filtered dataset. These are decision-support prompts, not causal claims or a substitute for business context.")

    recommendation_rows = []
    recommendation_rows.append({
        "Priority": "High" if missing > 0 else "Low",
        "Area": "Data quality",
        "Observation": f"{missing:,} missing cell(s) across {len(filtered_df):,} records." if missing else "No missing cells detected in the current filtered dataset.",
        "Recommended action": "Measure missingness by column, confirm why values are absent, then document a consistent handling rule before analysis." if missing else "Keep the current completeness check in the reporting workflow and validate business-specific rules such as valid ranges and required fields.",
        "Evidence / limitation": "Based on the current filtered dataset; missing values may be expected for some fields."
    })
    if duplicates:
        recommendation_rows.append({
            "Priority": "Medium",
            "Area": "Data quality",
            "Observation": f"{duplicates:,} exact duplicate row(s) detected.",
            "Recommended action": "Review duplicate rows against the source system and decide whether they represent repeated events or accidental duplicates before removing anything.",
            "Evidence / limitation": "Exact duplicates are a review flag, not automatic proof of bad data."
        })

    if auto_nums:
        chosen_metric = st.session_state.get("rg_auto_metric", auto_metric)
        if chosen_metric in filtered_df.columns:
            rec_values = pd.to_numeric(filtered_df[chosen_metric], errors="coerce").dropna()
            if not rec_values.empty:
                rec_q1, rec_q3 = rec_values.quantile([0.25, 0.75])
                rec_iqr = rec_q3 - rec_q1
                rec_outliers = int(((rec_values < rec_q1 - 1.5 * rec_iqr) | (rec_values > rec_q3 + 1.5 * rec_iqr)).sum()) if pd.notna(rec_iqr) and rec_iqr > 0 else 0
                if rec_outliers:
                    recommendation_rows.append({
                        "Priority": "Medium",
                        "Area": "Metric validation",
                        "Observation": f"{rec_outliers:,} value(s) in '{chosen_metric}' fall outside the 1.5×IQR screening range.",
                        "Recommended action": "Inspect flagged records, units, source entries, and relevant business events; correct only when verified against source data.",
                        "Evidence / limitation": "IQR is a screening heuristic; legitimate rare events can also be flagged."
                    })
                if len(rec_values) >= 2 and rec_values.nunique() > 1:
                    rec_mean, rec_median = float(rec_values.mean()), float(rec_values.median())
                    if abs(rec_mean - rec_median) / max(abs(rec_median), 1e-9) >= 0.25:
                        recommendation_rows.append({
                            "Priority": "Medium",
                            "Area": "Metric distribution",
                            "Observation": f"Mean ({rec_mean:,.2f}) differs from median ({rec_median:,.2f}) for '{chosen_metric}'.",
                            "Recommended action": "Review the distribution and segment the metric by relevant categories or time periods before relying on the average alone.",
                            "Evidence / limitation": "A mean–median gap can reflect skew or outliers; this check does not identify the cause."
                        })

    # Use existing automated category/trend tables as evidence where available.
    for table_title, table_frame in automated_tables:
        if table_title == "Category performance" and not table_frame.empty and "Total" in table_frame.columns:
            usable = table_frame.dropna(subset=["Total"]).sort_values("Total", ascending=False)
            if len(usable) >= 2:
                leader = usable.iloc[0]
                lagger = usable.iloc[-1]
                lead_value, low_value = float(leader["Total"]), float(lagger["Total"])
                cat_name = str(table_frame.columns[0])
                if lead_value != low_value:
                    recommendation_rows.append({
                        "Priority": "Medium",
                        "Area": "Category performance",
                        "Observation": f"'{safe_text(leader.iloc[0], 80)}' leads and '{safe_text(lagger.iloc[0], 80)}' is lowest by total {auto_metric} in the groups shown.",
                        "Recommended action": "Compare group sizes, margins or other relevant measures, and operational context before reallocating effort; investigate what differs in the lowest-total group.",
                        "Evidence / limitation": "Totals can be affected by record volume and do not alone measure profitability or efficiency."
                    })
        elif table_title == "Time trend" and not table_frame.empty and len(table_frame) >= 2 and auto_metric in table_frame.columns:
            vals = pd.to_numeric(table_frame[auto_metric], errors="coerce").dropna()
            if len(vals) >= 2:
                first_val, last_val = float(vals.iloc[0]), float(vals.iloc[-1])
                if first_val != last_val:
                    direction_text = "increased" if last_val > first_val else "decreased"
                    recommendation_rows.append({
                        "Priority": "Medium",
                        "Area": "Trend monitoring",
                        "Observation": f"Monthly '{auto_metric}' total {direction_text} from {first_val:,.2f} to {last_val:,.2f} between the first and last available periods.",
                        "Recommended action": "Review month-by-month movements, seasonality, record counts, and business events; monitor additional periods before treating the endpoint change as a persistent trend.",
                        "Evidence / limitation": "Endpoint comparison only; it does not establish the cause of the change."
                    })

    if not recommendation_rows:
        recommendation_rows.append({"Priority": "Low", "Area": "General", "Observation": "No specific rule-based flags were triggered.", "Recommended action": "Confirm the analysis covers the intended period and fields, then review findings with a domain expert.", "Evidence / limitation": "A lack of flags does not guarantee data quality or business performance."})

    # Stable de-duplication and a useful minimum set of recommendations.
    _rec_seen = set()
    _unique_recommendations = []
    for _row in recommendation_rows:
        _key = (_row["Area"], _row["Observation"])
        if _key not in _rec_seen:
            _rec_seen.add(_key)
            _unique_recommendations.append(_row)
    recommendation_rows = _unique_recommendations

    rec_df = pd.DataFrame(recommendation_rows)
    priority_order = {"High": 0, "Medium": 1, "Low": 2}
    rec_df["_priority_order"] = rec_df["Priority"].map(priority_order).fillna(3)
    rec_df = rec_df.sort_values(["_priority_order", "Area"]).drop(columns="_priority_order").reset_index(drop=True)
    rec_filter = st.multiselect("Filter recommendations by priority", ["High", "Medium", "Low"], default=["High", "Medium", "Low"], key="rg_rec_priority_filter")
    visible_rec_df = rec_df[rec_df["Priority"].isin(rec_filter)]
    st.dataframe(visible_rec_df, use_container_width=True, hide_index=True)
    st.download_button(
        "⬇️ Download recommendations (CSV)",
        data=visible_rec_df.to_csv(index=False).encode("utf-8-sig"),
        file_name="datawise_business_recommendations.csv",
        mime="text/csv",
        key="rg_download_recommendations_csv",
    )
    st.caption("Priority indicates a suggested review order based on simple rules, not business risk measured by a validated model.")


# =========================================================
    # STAGE 6.4 — AUTOMATED EXECUTIVE SUMMARY
    # =========================================================
    st.divider()
    st.markdown("### 🧾 Stage 6.4 — Automated executive summary")
    st.caption(
        "Build a concise management-ready summary from calculated dataset metrics, "
        "category comparisons, time trends, and review recommendations."
    )

    _summary_source = st.session_state.source_name or "uploaded dataset"
    _summary_lines = [
        f"EXECUTIVE SUMMARY — {report_title.strip() or 'Untitled business report'}",
        f"Source: {_summary_source}",
        f"Report date: {report_date.isoformat()}",
        f"Records analyzed: {len(filtered_df):,}",
        f"Columns analyzed: {len(filtered_df.columns):,}",
        "",
        "1. DATA QUALITY",
        f"Missing cells: {missing:,}",
        f"Exact duplicate rows: {duplicates:,}",
        f"Completeness: {completeness:.1f}%",
    ]
    if reporting_period.strip():
        _summary_lines.insert(3, f"Reporting period: {reporting_period.strip()}")

    _summary_findings = []
    _summary_findings.append(
        f"The analysis covers {len(filtered_df):,} records and {len(filtered_df.columns):,} columns."
    )
    _summary_findings.append(
        f"Data completeness is {completeness:.1f}%, with {missing:,} missing cell(s) "
        f"and {duplicates:,} exact duplicate row(s) detected."
    )

    _summary_metric_name = st.session_state.get("rg_auto_metric", None)
    if _summary_metric_name in filtered_df.columns:
        _summary_values = pd.to_numeric(filtered_df[_summary_metric_name], errors="coerce").dropna()
        if not _summary_values.empty:
            _summary_total = float(_summary_values.sum())
            _summary_mean = float(_summary_values.mean())
            _summary_median = float(_summary_values.median())
            _summary_min = float(_summary_values.min())
            _summary_max = float(_summary_values.max())
            _summary_lines.extend([
                "",
                "2. PRIMARY METRIC",
                f"Metric: {_summary_metric_name}",
                f"Total: {_summary_total:,.2f}",
                f"Average: {_summary_mean:,.2f}",
                f"Median: {_summary_median:,.2f}",
                f"Minimum: {_summary_min:,.2f}",
                f"Maximum: {_summary_max:,.2f}",
            ])
            _summary_findings.append(
                f"{_summary_metric_name} totals {_summary_total:,.2f}; its average is "
                f"{_summary_mean:,.2f} and median is {_summary_median:,.2f}."
            )

    _summary_category = st.session_state.get("rg_auto_category", None)
    if _summary_category in filtered_df.columns and _summary_metric_name in filtered_df.columns:
        _summary_cat_frame = filtered_df[[_summary_category, _summary_metric_name]].copy()
        _summary_cat_frame[_summary_metric_name] = pd.to_numeric(
            _summary_cat_frame[_summary_metric_name], errors="coerce"
        )
        _summary_cat_table = (
            _summary_cat_frame.groupby(_summary_category, dropna=False)[_summary_metric_name]
            .sum(min_count=1).dropna().sort_values(ascending=False)
        )
        if not _summary_cat_table.empty:
            _summary_leader = _summary_cat_table.index[0]
            _summary_lines.extend([
                "",
                "3. CATEGORY COMPARISON",
                f"Category field: {_summary_category}",
                f"Highest total: {safe_text(_summary_leader, 100)} "
                f"({_summary_cat_table.iloc[0]:,.2f})",
                f"Groups compared: {len(_summary_cat_table):,}",
            ])
            _summary_findings.append(
                f"The highest-total {_summary_category} group is "
                f"'{safe_text(_summary_leader, 100)}' at {_summary_cat_table.iloc[0]:,.2f} "
                f"for {_summary_metric_name}."
            )

    _summary_date = st.session_state.get("rg_auto_date", None)
    if _summary_date in filtered_df.columns and _summary_metric_name in filtered_df.columns:
        _summary_trend = filtered_df[[_summary_date, _summary_metric_name]].copy()
        _summary_trend[_summary_date] = pd.to_datetime(_summary_trend[_summary_date], errors="coerce")
        _summary_trend[_summary_metric_name] = pd.to_numeric(
            _summary_trend[_summary_metric_name], errors="coerce"
        )
        _summary_trend = _summary_trend.dropna()
        if not _summary_trend.empty:
            _summary_trend["Period"] = _summary_trend[_summary_date].dt.to_period("M").astype(str)
            _summary_months = _summary_trend.groupby("Period")[_summary_metric_name].sum().sort_index()
            _summary_lines.extend([
                "",
                "4. TIME TREND",
                f"Date field: {_summary_date}",
                f"Monthly periods available: {len(_summary_months):,}",
            ])
            if len(_summary_months) >= 2:
                _first, _last = float(_summary_months.iloc[0]), float(_summary_months.iloc[-1])
                _delta = _last - _first
                _direction = "increased" if _delta > 0 else "decreased" if _delta < 0 else "remained unchanged"
                _summary_lines.append(
                    f"{_summary_metric_name} { _direction } from {_first:,.2f} "
                    f"to {_last:,.2f} between the first and last months."
                )
                _summary_findings.append(
                    f"Across the first and last available months, {_summary_metric_name} "
                    f"{_direction} from {_first:,.2f} to {_last:,.2f}."
                )

    _summary_lines.extend(["", "5. KEY FINDINGS"])
    _summary_lines.extend([f"- {item}" for item in _summary_findings])
    _summary_lines.extend(["", "6. RECOMMENDED REVIEW ACTIONS"])
    for _, _recommendation in visible_rec_df.iterrows():
        _summary_lines.append(
            f"- [{_recommendation['Priority']}] {_recommendation['Area']}: "
            f"{_recommendation['Recommended action']}"
        )
    _summary_lines.extend([
        "",
        "INTERPRETATION NOTE",
        "This summary is generated from descriptive statistics and transparent rules. "
        "It does not establish causation, forecast future outcomes, or replace domain review.",
    ])

    _summary_text = "\\n".join(_summary_lines)
    with st.expander("Preview executive summary", expanded=True):
        st.text(_summary_text)

    _summary_col1, _summary_col2 = st.columns(2)
    with _summary_col1:
        st.download_button(
            "⬇️ Download executive summary (TXT)",
            data=_summary_text.encode("utf-8"),
            file_name="datawise_executive_summary.txt",
            mime="text/plain",
            use_container_width=True,
            key="stage64_summary_txt",
        )
    with _summary_col2:
        _summary_findings_df = pd.DataFrame(
            [{"Finding": finding} for finding in _summary_findings]
        )
        st.download_button(
            "⬇️ Download executive findings (CSV)",
            data=_summary_findings_df.to_csv(index=False).encode("utf-8-sig"),
            file_name="datawise_executive_findings.csv",
            mime="text/csv",
            use_container_width=True,
            key="stage64_summary_csv",
        )
    st.success("Executive summary is ready. It reflects the current dashboard filters and selected analysis fields.")


    # Make the recommendations available to existing PDF / PowerPoint report builders.
    _recommendation_findings = []
    for _, _row in visible_rec_df.iterrows():
        _recommendation_findings.append((
            f"Recommendation — {_row['Area']} ({_row['Priority']})",
            f"Observation: {_row['Observation']} Recommended action: {_row['Recommended action']} Evidence / limitation: {_row['Evidence / limitation']}"
        ))
    automated_findings.extend(_recommendation_findings)

    st.divider()
    st.markdown("### 📄 Stage 5.3 — Professional PDF export")
    st.write("Create a polished PDF using the report title, company, reporting period, selected sections, KPI fields, category comparison, and available charts.")

    pdf_sections = set()
    if include_overview:
        pdf_sections.add("overview")
    if include_kpis:
        pdf_sections.add("kpis")
    if include_insights:
        pdf_sections.add("insights")
    if include_charts:
        pdf_sections.add("charts")
    if include_categories:
        pdf_sections.add("categories")
    if include_recommendations:
        pdf_sections.add("recommendations")

    pdf_selected_metrics = st.session_state.get("rg_metrics", nums[:min(4, len(nums))]) if include_kpis else []
    pdf_category_col = st.session_state.get("rg_category") if include_categories else None
    pdf_category_metric = st.session_state.get("rg_category_metric") if include_categories else None

    if not REPORTLAB_AVAILABLE:
        st.warning("PDF generation needs ReportLab. Install it in your project terminal: python -m pip install reportlab")
    elif not pdf_sections:
        st.warning("Select at least one report section before creating a PDF.")
    else:
        try:
            pdf_bytes = make_custom_pdf_report(
                filtered_df,
                source_name,
                report_title.strip() or "Business Data Analysis Report",
                company_name.strip(),
                report_date.isoformat(),
                reporting_period.strip(),
                pdf_sections,
                selected_metrics=pdf_selected_metrics,
                category_col=pdf_category_col,
                category_metric=pdf_category_metric,
                automated_findings=automated_findings if "automated_findings" in locals() else [],
                theme_name=theme_name,
            )
            safe_file_title = re.sub(r"[^A-Za-z0-9_-]+", "_", (report_title.strip() or "datawise_report")).strip("_")
            st.success("Your customized PDF report is ready.")
            st.download_button(
                "⬇️ Download professional PDF report",
                data=pdf_bytes,
                file_name=f"{safe_file_title or 'datawise_report'}.pdf",
                mime="application/pdf",
                use_container_width=True,
                key="rg_stage_53_pdf_download",
            )
        except Exception as exc:
            st.error(f"PDF creation failed: {exc}")

    st.caption("PDF findings are calculated locally from the current filtered dataset. Potential outliers and category/trend comparisons are descriptive indicators, not proof of business causes.")

    st.divider()
    st.markdown("### 📊 Stage 5.4 — Professional PowerPoint generator")
    st.write("Build a customized slide deck using your report title, company details, selected sections, KPIs, category analysis, and charts.")
    ppt_sections = set()
    if include_overview: ppt_sections.add("overview")
    if include_kpis: ppt_sections.add("kpis")
    if include_insights: ppt_sections.add("insights")
    if include_charts: ppt_sections.add("charts")
    if include_categories: ppt_sections.add("categories")
    if include_recommendations: ppt_sections.add("recommendations")
    ppt_selected_metrics = st.session_state.get("rg_metrics", nums[:min(4, len(nums))]) if include_kpis else []
    ppt_category_col = st.session_state.get("rg_category") if include_categories else None
    ppt_category_metric = st.session_state.get("rg_category_metric") if include_categories else None
    if not PPTX_AVAILABLE:
        st.warning("PowerPoint generation needs python-pptx. Install it using: python -m pip install python-pptx")
    elif not ppt_sections:
        st.warning("Select at least one report section before creating a PowerPoint.")
    else:
        try:
            pptx_custom_bytes = make_custom_pptx_report(
                filtered_df, source_name, report_title.strip() or "Business Data Analysis Report",
                company_name.strip(), report_date.isoformat(), reporting_period.strip(), ppt_sections,
                selected_metrics=ppt_selected_metrics, category_col=ppt_category_col,
                category_metric=ppt_category_metric,
                automated_findings=automated_findings if "automated_findings" in locals() else [],
                theme_name=theme_name,
            )
            safe_ppt_title = re.sub(r"[^A-Za-z0-9_-]+", "_", (report_title.strip() or "datawise_presentation")).strip("_")
            st.success("Your customized PowerPoint presentation is ready.")
            st.download_button(
                "⬇️ Download professional PowerPoint (.pptx)",
                data=pptx_custom_bytes,
                file_name=f"{safe_ppt_title or 'datawise_presentation'}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
                key="rg_stage_54_pptx_download",
            )
        except Exception as exc:
            st.error(f"PowerPoint creation failed: {exc}")

    st.caption("Presentation content is generated locally from the current filtered dataset. Rule-based findings and descriptive comparisons are not proof of causation.")

    # =========================================================
    # STAGE 5.5 — MULTI-FORMAT REPORT AUTOMATION
    # =========================================================
    st.divider()
    st.markdown("### 📦 Stage 5.5 — Multi-format report automation")
    st.write("Generate a single ZIP package containing several report formats from the same filtered dataset.")
    st.caption("The package uses the current report title, company, reporting period, selected sections, metrics, and category settings where supported.")

    pack_cols = st.columns(4)
    with pack_cols[0]:
        pack_pdf = st.checkbox("Include PDF", value=True, key="rg_pack_pdf")
    with pack_cols[1]:
        pack_pptx = st.checkbox("Include PowerPoint", value=True, key="rg_pack_pptx")
    with pack_cols[2]:
        pack_excel = st.checkbox("Include Excel workbook", value=True, key="rg_pack_excel")
    with pack_cols[3]:
        pack_csv = st.checkbox("Include filtered data CSV", value=True, key="rg_pack_csv")

    pack_analysis_csv = st.checkbox("Include automated analysis tables (when available)", value=True, key="rg_pack_analysis_csv")
    pack_sections = set()
    if include_overview: pack_sections.add("overview")
    if include_kpis: pack_sections.add("kpis")
    if include_insights: pack_sections.add("insights")
    if include_charts: pack_sections.add("charts")
    if include_categories: pack_sections.add("categories")
    if include_recommendations: pack_sections.add("recommendations")

    pack_metrics = st.session_state.get("rg_metrics", nums[:min(4, len(nums))]) if include_kpis else []
    pack_category_col = st.session_state.get("rg_category") if include_categories else None
    pack_category_metric = st.session_state.get("rg_category_metric") if include_categories else None
    pack_findings = automated_findings if "automated_findings" in locals() else []
    safe_pack_title = re.sub(r"[^A-Za-z0-9_-]+", "_", (report_title.strip() or "datawise_report")).strip("_") or "datawise_report"

    if not any([pack_pdf, pack_pptx, pack_excel, pack_csv, pack_analysis_csv]):
        st.warning("Select at least one output format to create a package.")
    elif not pack_sections and (pack_pdf or pack_pptx):
        st.warning("Select at least one report section to include PDF or PowerPoint in the package.")

    if st.button("⚙️ Build all selected formats", type="primary", key="rg_build_multiformat"):
        try:
            package_buffer = io.BytesIO()
            created_files = []
            with zipfile.ZipFile(package_buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as archive:
                if pack_pdf:
                    if not REPORTLAB_AVAILABLE:
                        st.warning("PDF skipped: install ReportLab with `python -m pip install reportlab`.")
                    elif not pack_sections:
                        st.warning("PDF skipped because no report sections are selected.")
                    else:
                        content = make_custom_pdf_report(
                            filtered_df, source_name, report_title.strip() or "Business Data Analysis Report",
                            company_name.strip(), report_date.isoformat(), reporting_period.strip(), pack_sections,
                            selected_metrics=pack_metrics, category_col=pack_category_col,
                            category_metric=pack_category_metric, automated_findings=pack_findings,
                            theme_name=theme_name,
                        )
                        archive.writestr(f"{safe_pack_title}.pdf", content); created_files.append("PDF")
                if pack_pptx:
                    if not PPTX_AVAILABLE:
                        st.warning("PowerPoint skipped: install python-pptx with `python -m pip install python-pptx`.")
                    elif not pack_sections:
                        st.warning("PowerPoint skipped because no report sections are selected.")
                    else:
                        content = make_custom_pptx_report(
                            filtered_df, source_name, report_title.strip() or "Business Data Analysis Report",
                            company_name.strip(), report_date.isoformat(), reporting_period.strip(), pack_sections,
                            selected_metrics=pack_metrics, category_col=pack_category_col,
                            category_metric=pack_category_metric, automated_findings=pack_findings,
                            theme_name=theme_name,
                        )
                        archive.writestr(f"{safe_pack_title}.pptx", content); created_files.append("PowerPoint")
                if pack_excel:
                    content = make_excel_report(filtered_df, source_name, insights)
                    archive.writestr(f"{safe_pack_title}.xlsx", content); created_files.append("Excel")
                if pack_csv:
                    archive.writestr(f"{safe_pack_title}_filtered_data.csv", filtered_df.to_csv(index=False).encode("utf-8-sig")); created_files.append("Filtered data CSV")
                if pack_analysis_csv and "automated_tables" in locals() and automated_tables:
                    export_rows = []
                    for table_title, table_frame in automated_tables:
                        for _, record in table_frame.iterrows():
                            export_rows.append({"Analysis": table_title, **{str(k): (None if pd.isna(v) else v) for k, v in record.to_dict().items()}})
                    if export_rows:
                        archive.writestr(f"{safe_pack_title}_analysis_tables.csv", pd.DataFrame(export_rows).to_csv(index=False).encode("utf-8-sig"))
                        created_files.append("Analysis tables CSV")
                manifest = [
                    "DataWise AI — Multi-format report package",
                    f"Title: {report_title.strip() or 'Business Data Analysis Report'}",
                    f"Company: {company_name.strip() or 'Not provided'}",
                    f"Report date: {report_date.isoformat()}",
                    f"Reporting period: {reporting_period.strip() or 'Not specified'}",
                    f"Source dataset: {source_name}",
                    f"Filtered records: {len(filtered_df)}",
                    f"Columns: {len(filtered_df.columns)}",
                    "", "Files included:", *[f"- {name}" for name in created_files],
                    "", "Note: findings are descriptive and rule-based; they do not establish causation."
                ]
                archive.writestr("README_report_details.txt", "\n".join(manifest).encode("utf-8"))
            if not created_files:
                st.error("No report files were generated. Check the selected formats and installed dependencies.")
                st.session_state.pop("rg_multiformat_zip", None)
            else:
                st.session_state["rg_multiformat_zip"] = package_buffer.getvalue()
                st.session_state["rg_multiformat_filename"] = f"{safe_pack_title}_report_package.zip"
                if APP_DB_READY:
                    db_record_report(st.session_state.datawise_session_id, st.session_state.get("dataset_signature"),
                                     report_title.strip() or "Business Data Analysis Report", source_name,
                                     created_files, filtered_df, st.session_state.datawise_user["id"])
                st.success("Package ready: " + ", ".join(created_files))
        except Exception as exc:
            st.error(f"Could not build the multi-format package: {exc}")

    if st.session_state.get("rg_multiformat_zip"):
        st.download_button(
            "⬇️ Download complete report package (.zip)",
            data=st.session_state["rg_multiformat_zip"],
            file_name=st.session_state.get("rg_multiformat_filename", "datawise_report_package.zip"),
            mime="application/zip", use_container_width=True, key="rg_download_multiformat_zip",
        )

    st.caption("Each file in the ZIP is generated locally from the current filtered dataset. If an optional library is unavailable, that format may be skipped; review the messages above the download button.")


# =========================================================
# REPORT CENTER
# =========================================================
elif page == "📑 Reports":
    st.markdown("## 📑 Report center")
    st.write("Generate a dataset summary and download it for sharing.")
    missing, duplicates, completeness = data_quality(filtered_df)
    report_table = pd.DataFrame({
        "Metric": ["Rows", "Columns", "Missing cells", "Duplicate rows", "Cell completeness"],
        "Value": [len(filtered_df), len(filtered_df.columns), missing, duplicates, f"{completeness:.1f}%"],
    })
    st.dataframe(report_table, use_container_width=True)
    insights = make_insights(filtered_df)
    st.markdown("### Key findings")
    for title, body in insights:
        st.markdown(f"**{title}:** {body}")

    col1, col2, col3 = st.columns(3)
    with col1:
        st.download_button(
            "⬇️ Export report data (CSV)",
            report_table.to_csv(index=False).encode("utf-8"),
            file_name="datawise_summary.csv",
            mime="text/csv",
            use_container_width=True,
        )
    with col2:
        try:
            excel_bytes = make_excel_report(filtered_df, st.session_state.source_name or "dataset", insights)
            st.download_button(
                "⬇️ Export professional Excel report",
                excel_bytes,
                file_name="datawise_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        except Exception as exc:
            st.error(f"Excel creation failed: {exc}")
    with col3:
        if REPORTLAB_AVAILABLE:
            try:
                pdf_bytes = make_pdf(filtered_df, st.session_state.source_name or "dataset", insights)
                st.download_button(
                    "⬇️ Download PDF report",
                    pdf_bytes,
                    file_name="datawise_report.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )
            except Exception as exc:
                st.error(f"PDF creation failed: {exc}")
        else:
            st.warning("Install ReportLab to enable PDF export.")

    if PPTX_AVAILABLE:
        try:
            pptx_bytes = make_pptx(filtered_df, st.session_state.source_name or "dataset", insights)
            st.download_button(
                "⬇️ Download PowerPoint",
                pptx_bytes,
                file_name="datawise_presentation.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )
        except Exception as exc:
            st.error(f"PowerPoint creation failed: {exc}")
    else:
        st.warning("Install python-pptx to enable PowerPoint export.")

    st.caption("PDF, PowerPoint, and Excel exports include formatted summaries and locally calculated findings. Suitable charts are embedded in the PDF and PowerPoint, and the Excel workbook includes a category analysis chart when the dataset supports it. Findings are rule-based; no hosted AI model is connected in this free-first version.")


# =========================================================
# STAGE 7.9 — PERSONAL DATA & REPORT HISTORY
# =========================================================
elif page == "📚 My History":
    st.markdown("## 📚 My workspace history")
    st.caption("Only records owned by your account are shown here. Original dataset files and report binaries are not stored in this history.")
    current_user = st.session_state.datawise_user
    my_datasets = db_read_user_history("dataset_history", current_user["id"], is_admin=False, limit=1000)
    my_reports = db_read_user_history("report_history", current_user["id"], is_admin=False, limit=1000)
    a, b = st.columns(2)
    a.metric("Your datasets", f"{len(my_datasets):,}")
    b.metric("Your generated reports", f"{len(my_reports):,}")
    tab_my_data, tab_my_reports = st.tabs(["📁 My datasets", "📄 My reports"])
    with tab_my_data:
        if my_datasets.empty:
            st.info("You haven't uploaded or imported any datasets yet.")
        else:
            display_data = my_datasets.drop(columns=["dataset_signature", "columns_json", "session_id"], errors="ignore")
            st.dataframe(display_data, use_container_width=True, hide_index=True)
            st.download_button("⬇️ Export my dataset history", my_datasets.to_csv(index=False).encode("utf-8-sig"),
                               "my_datawise_datasets.csv", "text/csv", key="stage79_my_datasets_export")
    with tab_my_reports:
        if my_reports.empty:
            st.info("You haven't generated any report packages yet.")
        else:
            display_reports = my_reports.drop(columns=["session_id"], errors="ignore").copy()
            display_reports["formats_json"] = display_reports["formats_json"].apply(
                lambda value: ", ".join(json.loads(value)) if value else ""
            )
            st.dataframe(display_reports, use_container_width=True, hide_index=True)
            st.download_button("⬇️ Export my report history", my_reports.to_csv(index=False).encode("utf-8-sig"),
                               "my_datawise_reports.csv", "text/csv", key="stage79_my_reports_export")


# =========================================================
# STAGE 7.6 — APPLICATION HISTORY, SETTINGS & BACKUP
# =========================================================
elif page == "🗃️ App History":
    st.markdown("## 🗃️ Application database & history")
    st.caption("This page shows metadata stored by DataWise AI. Raw uploaded datasets and generated report files are not saved in the database.")
    if not APP_DB_READY:
        st.error("The application database is unavailable. Check the startup error and database folder permissions.")
    else:
        datasets = db_read_user_history("dataset_history", st.session_state.datawise_user["id"], is_admin=True)
        reports = db_read_user_history("report_history", st.session_state.datawise_user["id"], is_admin=True)
        activities = db_read_table("activity_log")
        m1, m2, m3 = st.columns(3)
        m1.metric("Datasets recorded", f"{len(datasets):,}")
        m2.metric("Reports generated", f"{len(reports):,}")
        m3.metric("Activity events", f"{len(activities):,}")

        tab_datasets, tab_reports, tab_activity, tab_settings = st.tabs(
            ["📁 Dataset history", "📄 Report history", "🧾 Activity log", "⚙️ Settings & backup"]
        )
        with tab_datasets:
            st.markdown("### Previously loaded datasets")
            if datasets.empty:
                st.info("No dataset uploads have been recorded yet.")
            else:
                view = datasets.drop(columns=["dataset_signature", "columns_json"], errors="ignore")
                st.dataframe(view, use_container_width=True, hide_index=True)
                st.download_button("⬇️ Download dataset history CSV", datasets.to_csv(index=False).encode("utf-8-sig"),
                                   "datawise_dataset_history.csv", "text/csv", key="stage76_dataset_csv")
        with tab_reports:
            st.markdown("### Generated report packages")
            if reports.empty:
                st.info("No report packages have been generated yet. Create one in Report Generator.")
            else:
                report_view = reports.copy()
                report_view["formats_json"] = report_view["formats_json"].apply(
                    lambda value: ", ".join(json.loads(value)) if value else ""
                )
                st.dataframe(report_view, use_container_width=True, hide_index=True)
                st.download_button("⬇️ Download report history CSV", reports.to_csv(index=False).encode("utf-8-sig"),
                                   "datawise_report_history.csv", "text/csv", key="stage76_report_csv")
        with tab_activity:
            st.markdown("### Recent application activity")
            if activities.empty:
                st.info("No activity has been recorded yet.")
            else:
                st.dataframe(activities, use_container_width=True, hide_index=True)
                st.download_button("⬇️ Download activity log CSV", activities.to_csv(index=False).encode("utf-8-sig"),
                                   "datawise_activity_log.csv", "text/csv", key="stage76_activity_csv")
        with tab_settings:
            st.markdown("### Persistent application settings")
            workspace_label = st.text_input("Workspace label", value=db_get_setting("workspace_label", "DataWise AI"),
                                             key="stage76_workspace_label")
            if st.button("Save workspace settings", key="stage76_save_settings"):
                db_set_setting("workspace_label", workspace_label.strip() or "DataWise AI")
                db_log_event(st.session_state.datawise_session_id, "settings_updated", "Updated workspace label")
                st.success("Settings saved in the application database.")

            st.divider()
            if APP_DB_BACKEND == "postgresql":
                st.markdown("### PostgreSQL database backup")
                st.success("DataWise AI is using Neon PostgreSQL for application records.")
                st.info(
                    "SQLite file backup and restore are disabled while PostgreSQL is active. "
                    "Use your Neon dashboard's backup/restore and branching tools for managed database recovery. "
                    "Keep your original SQLite backup safely until production verification is complete."
                )
                st.caption("Uploaded SQLite databases remain separate data sources and are not affected by this setting.")
            else:
                st.markdown("### SQLite database backup")
                st.caption(f"Database location: `{APP_DB_PATH}`")
                try:
                    backup_path = None
                    with tempfile.NamedTemporaryFile(suffix=".db", delete=False) as tmp_backup:
                        backup_path = tmp_backup.name
                    src = sqlite3.connect(str(APP_DB_PATH))
                    dst = sqlite3.connect(backup_path)
                    try:
                        src.backup(dst)
                    finally:
                        dst.close()
                        src.close()
                    backup_bytes = Path(backup_path).read_bytes()
                    Path(backup_path).unlink(missing_ok=True)
                    st.download_button(
                        "⬇️ Download database backup", backup_bytes,
                        f"datawise_backup_{date.today().isoformat()}.db",
                        "application/octet-stream", key="stage76_backup_download"
                    )
                except Exception as backup_error:
                    st.error(f"Could not create backup: {backup_error}")

                st.warning("Restore replaces the current application history and settings. Download a backup first.")
                restore_file = st.file_uploader(
                    "Select DataWise AI database backup", type=["db", "sqlite"], key="stage76_restore_file"
                )
                if restore_file is not None and st.button(
                    "Restore database backup", type="secondary", key="stage76_restore_button"
                ):
                    try:
                        backup_bytes = restore_file.getvalue()
                        if not db_validate_backup(backup_bytes):
                            st.error("This file is not a valid DataWise AI database backup.")
                        else:
                            restore_tmp = APP_DB_PATH.with_suffix(".restore.tmp")
                            restore_tmp.write_bytes(backup_bytes)
                            with sqlite3.connect(str(restore_tmp)) as check_conn:
                                integrity = check_conn.execute("PRAGMA integrity_check").fetchone()[0]
                                if integrity != "ok":
                                    raise ValueError("Backup integrity check failed")
                            for sidecar in (Path(str(APP_DB_PATH) + "-wal"), Path(str(APP_DB_PATH) + "-shm")):
                                sidecar.unlink(missing_ok=True)
                            os.replace(restore_tmp, APP_DB_PATH)
                            st.success("Database restored successfully. Refreshing the app...")
                            st.rerun()
                    except Exception as restore_error:
                        st.error(f"Database restore failed: {restore_error}")

    st.divider()
    st.info("Privacy note: this history is stored locally in the app database. On a cloud host with ephemeral storage, the database may reset when the app restarts or redeploys. Use a persistent disk or managed database for production. This page is not protected by user authentication; do not expose it publicly with sensitive data.")


# =========================================================
# STAGE 7.8 — ADMIN DASHBOARD & ACTIVITY MONITORING
# =========================================================
elif page == "🛡️ Admin Monitor":
    if st.session_state.datawise_user["role"] != "admin":
        st.error("Administrator permission is required to view monitoring data.")
        st.stop()

    st.markdown("## 🛡️ Administrator monitoring")
    st.caption("Workspace health, account activity, sign-in events and usage trends.")

    users_monitor = auth_list_users()
    datasets_monitor = db_read_table("dataset_history", limit=10000)
    reports_monitor = db_read_table("report_history", limit=10000)
    activity_monitor = db_read_table("activity_log", limit=10000)

    total_users = len(users_monitor)
    active_users = int(users_monitor["is_active"].astype(bool).sum()) if total_users else 0
    admin_users = int((users_monitor["role"] == "admin").sum()) if total_users else 0
    failed_logins = int((activity_monitor["event_type"] == "login_failed").sum()) if not activity_monitor.empty else 0
    successful_logins = int((activity_monitor["event_type"] == "login_success").sum()) if not activity_monitor.empty else 0
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Total accounts", f"{total_users:,}")
    c2.metric("Active accounts", f"{active_users:,}")
    c3.metric("Administrators", f"{admin_users:,}")
    c4.metric("Datasets recorded", f"{len(datasets_monitor):,}")
    c5.metric("Reports recorded", f"{len(reports_monitor):,}")

    st.divider()
    tab_overview, tab_signins, tab_users, tab_events = st.tabs(
        ["📈 Overview", "🔑 Sign-in activity", "👥 Accounts", "🧾 Audit events"]
    )
    with tab_overview:
        st.markdown("### Workspace activity overview")
        if activity_monitor.empty:
            st.info("No activity has been recorded yet.")
        else:
            event_counts = (activity_monitor.groupby("event_type", dropna=False)
                            .size().reset_index(name="events").sort_values("events", ascending=False))
            left, right = st.columns(2)
            with left:
                st.markdown("**Events by type**")
                st.plotly_chart(px.bar(event_counts, x="event_type", y="events", color="event_type",
                                       title="Recorded events"), use_container_width=True)
            with right:
                st.markdown("**Recent activity**")
                recent = activity_monitor.head(12).copy()
                recent["created_at"] = pd.to_datetime(recent["created_at"], errors="coerce", utc=True)
                daily = (recent.dropna(subset=["created_at"]).assign(day=lambda x: x["created_at"].dt.strftime("%Y-%m-%d"))
                         .groupby("day").size().reset_index(name="events").sort_values("day"))
                if not daily.empty:
                    st.plotly_chart(px.line(daily, x="day", y="events", markers=True,
                                            title="Recent event volume"), use_container_width=True)
                else:
                    st.caption("Not enough dated events to plot a trend.")
            st.caption(f"Recorded sign-ins: {successful_logins:,} successful · {failed_logins:,} failed. These are event counts, not a count of currently online users.")
        st.markdown("### Latest account sign-ins")
        if users_monitor.empty:
            st.info("No accounts exist yet.")
        else:
            latest = users_monitor[["display_name", "email", "role", "is_active", "last_login_at"]].copy()
            latest["is_active"] = latest["is_active"].map({1: "Active", 0: "Inactive", True: "Active", False: "Inactive"})
            st.dataframe(latest.sort_values("last_login_at", ascending=False, na_position="last"),
                         use_container_width=True, hide_index=True)

    with tab_signins:
        st.markdown("### Authentication events")
        if activity_monitor.empty:
            signins = pd.DataFrame()
        else:
            signins = activity_monitor[activity_monitor["event_type"].isin(["login_success", "login_failed", "logout"])].copy()
        if signins.empty:
            st.info("No sign-in events are available yet. New successful and failed sign-ins will be recorded from this version onward.")
        else:
            status_filter = st.multiselect("Event types", ["login_success", "login_failed", "logout"],
                                           default=["login_success", "login_failed", "logout"], key="stage78_signin_filter")
            filtered_signins = signins[signins["event_type"].isin(status_filter)]
            st.dataframe(filtered_signins, use_container_width=True, hide_index=True)
            st.download_button("⬇️ Export sign-in events", filtered_signins.to_csv(index=False).encode("utf-8-sig"),
                               "datawise_signin_events.csv", "text/csv", key="stage78_signin_export")

    with tab_users:
        st.markdown("### Account directory")
        if users_monitor.empty:
            st.info("No users have been created.")
        else:
            role_filter = st.multiselect("Filter by role", ["admin", "analyst", "viewer"],
                                         default=["admin", "analyst", "viewer"], key="stage78_role_filter")
            status_filter_users = st.selectbox("Account status", ["All", "Active", "Inactive"], key="stage78_status_filter")
            filtered_users = users_monitor[users_monitor["role"].isin(role_filter)].copy()
            if status_filter_users != "All":
                filtered_users = filtered_users[filtered_users["is_active"].astype(bool) == (status_filter_users == "Active")]
            st.dataframe(filtered_users, use_container_width=True, hide_index=True)
            st.download_button("⬇️ Export account directory", filtered_users.to_csv(index=False).encode("utf-8-sig"),
                               "datawise_accounts.csv", "text/csv", key="stage78_users_export")

    with tab_events:
        st.markdown("### Audit event explorer")
        if activity_monitor.empty:
            st.info("No audit events have been recorded.")
        else:
            event_types = sorted(activity_monitor["event_type"].dropna().astype(str).unique().tolist())
            selected_types = st.multiselect("Filter event types", event_types, default=event_types,
                                            key="stage78_event_filter")
            search_text = st.text_input("Search event descriptions", key="stage78_event_search").strip().lower()
            filtered_events = activity_monitor[activity_monitor["event_type"].isin(selected_types)].copy()
            if search_text:
                filtered_events = filtered_events[filtered_events["description"].astype(str).str.lower().str.contains(search_text, regex=False)]
            st.dataframe(filtered_events, use_container_width=True, hide_index=True)
            st.download_button("⬇️ Export filtered audit events", filtered_events.to_csv(index=False).encode("utf-8-sig"),
                               "datawise_audit_events.csv", "text/csv", key="stage78_audit_export")

    st.divider()
    st.warning("Monitoring is based on events stored by this application. It does not show live online presence, device identity or IP addresses. Historical events from before Stage 7.8 cannot be reconstructed. For public deployment, use HTTPS, persistent storage, backups and a managed authentication/security review.")


# =========================================================
# STAGE 7.7 — ADMIN USER MANAGEMENT
# =========================================================
elif page == "👥 User Management":
    if st.session_state.datawise_user["role"] != "admin":
        st.error("Administrator permission is required to manage accounts.")
        st.stop()
    st.markdown("## 👥 User management")
    st.caption("Create accounts, assign roles, and activate or deactivate workspace access.")
    with st.expander("➕ Create a user account", expanded=True):
        with st.form("stage77_create_user_form"):
            new_name = st.text_input("Full name", key="stage77_new_name")
            new_email = st.text_input("Email address", key="stage77_new_email")
            new_password = st.text_input("Temporary password (minimum 12 characters)", type="password", key="stage77_new_password")
            new_role = st.selectbox("Role", ["analyst", "viewer", "admin"], index=0, key="stage77_new_role")
            create_submitted = st.form_submit_button("Create account", type="primary")
        if create_submitted:
            try:
                auth_create_user(new_email, new_name, new_password, new_role)
                db_log_event(str(st.session_state.datawise_session_id), "user_created", f"Created {new_role} account: {new_email.strip().lower()}")
                st.success("Account created successfully.")
                st.rerun()
            except ValueError as exc:
                st.error(str(exc))

    st.markdown("### Existing accounts")
    users_df = auth_list_users()
    st.dataframe(users_df, use_container_width=True, hide_index=True)
    if not users_df.empty:
        selected_user_id = st.selectbox(
            "Select account to update", users_df["id"].tolist(),
            format_func=lambda uid: f"{users_df.loc[users_df['id'] == uid, 'email'].iloc[0]} (ID {uid})",
            key="stage77_selected_user"
        )
        selected_row = users_df[users_df["id"] == selected_user_id].iloc[0]
        with st.form("stage77_update_user_form"):
            updated_role = st.selectbox("Role", ["admin", "analyst", "viewer"],
                                        index=["admin", "analyst", "viewer"].index(selected_row["role"]))
            updated_active = st.checkbox("Account active", value=bool(selected_row["is_active"]))
            update_submitted = st.form_submit_button("Save account changes")
        if update_submitted:
            try:
                auth_update_user(selected_user_id, updated_role, updated_active)
                db_log_event(str(st.session_state.datawise_session_id), "user_updated", f"Updated account ID {selected_user_id}")
                st.success("Account updated.")
                st.rerun()
            except ValueError as exc:
                st.error(str(exc))

    st.info("Role permissions: Admin — full access and account/database management; Analyst — data analysis, cleaning, imports and reports; Viewer — read-only analysis and report viewing. This is application-level access control, not a substitute for HTTPS, secure hosting, or managed identity controls.")


elif page == "🔐 Account Settings":
    st.markdown("## 🔐 Account settings")
    user = st.session_state.datawise_user
    st.write(f"**Name:** {safe_text(user['display_name'], 80)}")
    st.write(f"**Email:** {safe_text(user['email'], 120)}")
    st.write(f"**Role:** {user['role'].upper()}")
    st.divider()
    st.markdown("### 🛡️ Session security")
    st.info("For security, sessions expire after 30 minutes of inactivity or 12 hours total. Signing out clears the active in-memory dataset from this browser session.")
    st.markdown("### Change your password")
    with st.form("stage77_change_password_form"):
        current_pw = st.text_input("Current password", type="password")
        next_pw = st.text_input("New password (minimum 12 characters)", type="password")
        confirm_pw = st.text_input("Confirm new password", type="password")
        change_submitted = st.form_submit_button("Change password", type="primary")
    if change_submitted:
        if next_pw != confirm_pw:
            st.error("New passwords do not match.")
        else:
            try:
                auth_change_password(user["id"], current_pw, next_pw)
                db_log_event(str(st.session_state.datawise_session_id), "password_changed", f"{user['role']} user changed their password")
                st.success("Password changed successfully. Your current session remains active.")
            except ValueError as exc:
                st.error(str(exc))


# =========================================================
# STAGE 7.10 — WORKSPACE CUSTOMIZATION & DEPLOYMENT READINESS
# =========================================================
elif page == "🚀 Deployment & Settings":
    if st.session_state.datawise_user["role"] != "admin":
        st.error("Administrator permission is required to change workspace and deployment settings.")
        st.stop()

    st.markdown("## 🚀 Workspace & deployment centre")
    st.caption("Customize the workspace and review the environment before deploying DataWise AI.")
    tab_brand, tab_health, tab_deploy, tab_cloud, tab_postgres = st.tabs(["🎨 Workspace branding", "🩺 System health", "☁️ Deployment checklist", "🌐 Cloud preflight", "🐘 PostgreSQL"])

    with tab_brand:
        st.markdown("### Workspace identity")
        current_label = db_get_setting("workspace_label", "DataWise AI")
        new_label = st.text_input("Workspace name", value=current_label, max_chars=60, key="stage710_workspace_name")
        st.caption("This name appears in the sidebar. Keep it short for a clean layout.")
        if st.button("Save workspace name", type="primary", key="stage710_save_brand"):
            cleaned_label = re.sub(r"\s+", " ", new_label).strip()
            if not cleaned_label:
                st.error("Workspace name cannot be empty.")
            else:
                db_set_setting("workspace_label", cleaned_label[:60])
                db_log_event(str(st.session_state.datawise_session_id), "workspace_brand_updated", "Updated workspace display name")
                st.success("Workspace name saved. The sidebar will update on the next rerun.")
                st.rerun()
        st.info("Branding changes affect the display name only; they do not change account permissions or data ownership.")

    with tab_health:
        st.markdown("### Runtime health checks")
        checks = []
        checks.append(("Python runtime", sys.version.split()[0], sys.version_info >= (3, 10)))
        checks.append(("Streamlit", getattr(st, "__version__", "Unknown"), True))
        checks.append(("Application database", APP_DB_BACKEND.upper(), APP_DB_READY))
        checks.append(("Database folder writable", str(APP_DB_PATH.parent), os.access(APP_DB_PATH.parent, os.W_OK)) if APP_DB_BACKEND == "sqlite" else ("Local SQLite fallback", "Not active; using PostgreSQL", True))
        checks.append(("PDF export library", "reportlab installed" if REPORTLAB_AVAILABLE else "not installed", REPORTLAB_AVAILABLE))
        checks.append(("PowerPoint export library", "python-pptx installed" if PPTX_AVAILABLE else "not installed", PPTX_AVAILABLE))
        checks.append(("Matplotlib", "installed" if MATPLOTLIB_AVAILABLE else "not installed", MATPLOTLIB_AVAILABLE))
        if APP_DB_READY:
            try:
                with app_db_connect() as _health_conn:
                    _health_conn.execute("SELECT 1").fetchone()
                checks.append((f"{APP_DB_BACKEND.title()} connection", "Connection successful", True))
            except Exception as _health_error:
                checks.append((f"{APP_DB_BACKEND.title()} connection", str(_health_error)[:120], False))
        for _label, _detail, _ok in checks:
            (_c1, _c2, _c3) = st.columns([2, 4, 1])
            _c1.write(_label)
            _c2.caption(_detail)
            _c3.success("OK") if _ok else _c3.warning("Check")
        try:
            _disk = shutil.disk_usage(str(APP_DB_PATH.parent))
            _d1, _d2, _d3 = st.columns(3)
            _d1.metric("Disk total", f"{_disk.total / (1024**3):.2f} GB")
            _d2.metric("Disk free", f"{_disk.free / (1024**3):.2f} GB")
            _d3.metric("Database size", f"{APP_DB_PATH.stat().st_size / (1024**2):.2f} MB" if APP_DB_PATH.exists() else "0 MB")
        except Exception as _disk_error:
            st.caption(f"Disk usage unavailable: {_disk_error}")

    with tab_deploy:
        st.markdown("### Production deployment checklist")
        st.warning("A successful local health check does not mean the application is production-secure. Complete the items below before inviting real users.")
        deployment_items = [
            "Use HTTPS through your hosting provider or reverse proxy.",
            "Use persistent storage or a managed database; local SQLite files may be lost on redeploy.",
            "Set DATAWISE_DB_PATH to a persistent writable location when using SQLite.",
            "Configure Streamlit server upload limits in .streamlit/config.toml.",
            "Keep credentials and API keys in hosting secrets/environment variables, never in source code.",
            "Restrict database backups and protect downloaded backup files.",
            "Test user roles, ownership boundaries, login failure handling and session expiry.",
            "Create a backup and test restoring it before production use.",
        ]
        for item in deployment_items:
            st.markdown(f"- [ ] {item}")
        st.markdown("#### Suggested Streamlit configuration")
        st.code("[server]\\nmaxUploadSize = 200\\nenableCORS = false\\nenableXsrfProtection = true\\n\\n[browser]\\ngatherUsageStats = false", language="toml")
        st.caption("Save this as .streamlit/config.toml in your project. Adjust maxUploadSize to your hosting and memory limits.")
        st.markdown("#### Database storage configuration")
        st.code("# Windows PowerShell example\\n$env:DATAWISE_DB_PATH = \"D:\\\\DataWiseData\\\\datawise_app.db\"\\nstreamlit run app.py", language="powershell")
        st.caption("The database path must point to a persistent folder that the app process can write to. Do not expose the database file publicly.")

    with tab_cloud:
        st.markdown("### Cloud deployment preflight")
        st.info("This stage prepares a Streamlit deployment package. It does not provision hosting or convert SQLite into a managed cloud database.")
        _is_cloud = any(os.environ.get(k) for k in ("STREAMLIT_SERVER_HEADLESS", "DYNO", "RAILWAY_ENVIRONMENT", "RENDER", "K_SERVICE"))
        _cloud_checks = [
            ("Application entry point", Path(__file__).name, Path(__file__).exists()),
            ("Required question-answering module", "modules/question_answering.py", (Path(__file__).resolve().parent / "modules" / "question_answering.py").exists()),
            ("Database path configured", str(APP_DB_PATH), bool(os.environ.get("DATAWISE_DB_PATH"))),
            ("SQLite database writable", str(APP_DB_PATH.parent), os.access(APP_DB_PATH.parent, os.W_OK)),
            ("Running in detected cloud environment", "Detected" if _is_cloud else "Not detected (may be local)", True),
        ]
        for _label, _detail, _ok in _cloud_checks:
            _a, _b, _c = st.columns([2, 4, 1])
            _a.write(_label); _b.caption(_detail)
            _c.success("OK") if _ok else _c.warning("Action")
        st.markdown("#### Deployment boundaries")
        st.markdown("""- SQLite is suitable for local testing or a single persistent instance with a durable disk.
- Do not run multiple app replicas against a copied/local SQLite file.
- For multi-user production hosting, move application records to a managed database such as PostgreSQL in a later database migration stage.
- This app currently has application-level login; it is not a substitute for an external identity provider or a professional security audit.""")
        st.markdown("#### Before pushing to a public Git repository")
        st.markdown("""- Keep `datawise_app.db`, `.env`, `.streamlit/secrets.toml`, backups and uploaded customer files out of Git.
- Add the `modules` directory from your working project; the app imports `modules.question_answering`.
- Add the generated `requirements.txt` and `.streamlit/config.toml` from the Stage 7.11 deployment bundle.""")



    with tab_postgres:
        st.markdown("### 🐘 PostgreSQL cloud database")
        st.info(
            "Stage 7.13 connects the application runtime to PostgreSQL whenever DATABASE_URL is configured. "
            "The migrated Neon records are used directly for accounts, history, activity and settings."
        )
        st.markdown("#### Configure connection")
        st.caption("Set DATABASE_URL in your hosting provider's secrets/environment settings. Do not paste production credentials into source code.")
        _pg_url = os.environ.get("DATABASE_URL", "")
        _pg_ready = bool(_pg_url)
        _pg_col1, _pg_col2 = st.columns(2)
        _pg_col1.metric("PostgreSQL driver", "Installed" if PSYCOPG_AVAILABLE else "Not installed")
        _pg_col2.metric("DATABASE_URL", "Configured" if _pg_ready else "Not configured")
        if st.button("Test PostgreSQL connection", key="stage712_test_pg", disabled=not (_pg_ready and PSYCOPG_AVAILABLE)):
            try:
                with psycopg.connect(_pg_url, connect_timeout=8, row_factory=dict_row) as _pg_conn:
                    _pg_row = _pg_conn.execute("SELECT version() AS version, current_database() AS db_name").fetchone()
                st.success(f"Connected to {_pg_row['db_name']}.")
                st.caption(str(_pg_row["version"])[:220])
                db_log_event(str(st.session_state.datawise_session_id), "postgres_connection_test", "PostgreSQL connection test succeeded")
            except Exception as _pg_exc:
                st.error(f"PostgreSQL connection failed: {str(_pg_exc)[:300]}")
                db_log_event(str(st.session_state.datawise_session_id), "postgres_connection_test_failed", "PostgreSQL connection test failed")
        if not PSYCOPG_AVAILABLE:
            st.warning("Install the PostgreSQL driver with: python -m pip install 'psycopg[binary]>=3.2,<4'")
        elif not _pg_ready:
            st.warning("DATABASE_URL is not configured. Add it to your local environment or cloud secrets.")
        st.markdown("#### Migration procedure")
        st.markdown("""
1. Create a managed PostgreSQL database with your hosting provider.
2. Set `DATABASE_URL` as a private environment variable.
3. Install the PostgreSQL driver from `requirements.txt`.
4. Stop the app and make a fresh SQLite backup.
5. Run `python migrate_sqlite_to_postgres.py --sqlite datawise_app.db`.
6. Verify the migrated row counts before switching application reads/writes in the next stage.
        """)
        st.warning("Do not delete the SQLite database after migration. Keep a protected backup until the PostgreSQL-backed application has passed its verification tests.")
        st.markdown("#### Supported data migrated")
        st.write("User accounts, dataset history, report history, activity log and workspace settings. Password hashes are copied as hashes; plaintext passwords are never exported.")


st.markdown('<div class="footer">DATAWISE AI · Smart Analytics Workspace · Stage 7.13 — PostgreSQL runtime · Cloud deployment preparation</div>', unsafe_allow_html=True)
