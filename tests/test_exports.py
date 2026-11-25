"""Stage 3.2: validate PDF, PowerPoint, Excel, and CSV exports from app.py."""
import ast
import io
import zipfile
from datetime import date, datetime
from pathlib import Path
import html

import numpy as np
import pandas as pd
import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[1]
APP_FILE = PROJECT_ROOT / "app.py"

EXPORT_FUNCTIONS = {
    "safe_text", "numeric_columns", "date_columns", "data_quality",
    "make_insights", "make_report_charts", "make_pdf", "make_pptx",
    "make_excel_report",
}


def _load_export_functions():
    """Load export helpers from app.py without executing Streamlit UI code."""
    source = APP_FILE.read_text(encoding="utf-8")
    tree = ast.parse(source, filename=str(APP_FILE))
    selected = [node for node in tree.body if isinstance(node, ast.FunctionDef) and node.name in EXPORT_FUNCTIONS]
    found = {node.name for node in selected}
    missing = EXPORT_FUNCTIONS - found
    if missing:
        raise RuntimeError(
            "The app.py file is missing expected Stage 2 export helpers: " + ", ".join(sorted(missing))
        )

    namespace = {
        "io": io, "html": html, "np": np, "pd": pd,
        "date": date, "datetime": datetime,
    }

    # Match the optional export libraries used by app.py.
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    namespace.update({"matplotlib": matplotlib, "plt": plt, "MATPLOTLIB_AVAILABLE": True})

    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image as ReportImage
    namespace.update({
        "colors": colors, "A4": A4, "landscape": landscape,
        "getSampleStyleSheet": getSampleStyleSheet, "ParagraphStyle": ParagraphStyle,
        "mm": mm, "SimpleDocTemplate": SimpleDocTemplate, "Paragraph": Paragraph,
        "Spacer": Spacer, "Table": Table, "TableStyle": TableStyle,
        "PageBreak": PageBreak, "ReportImage": ReportImage, "REPORTLAB_AVAILABLE": True,
    })

    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.enum.text import PP_ALIGN
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    namespace.update({
        "Presentation": Presentation, "Inches": Inches, "Pt": Pt,
        "PP_ALIGN": PP_ALIGN, "RGBColor": RGBColor, "MSO_SHAPE": MSO_SHAPE,
        "PPTX_AVAILABLE": True,
    })

    exec(compile(ast.Module(body=selected, type_ignores=[]), str(APP_FILE), "exec"), namespace)
    return namespace


@pytest.fixture(scope="module")
def exports():
    if not APP_FILE.exists():
        pytest.fail(f"Could not find {APP_FILE}. Keep the tests folder beside your app.py file.")
    return _load_export_functions()


@pytest.fixture
def sales_data():
    return pd.DataFrame({
        "Product": ["Laptop", "Phone", "Tablet", "Laptop"],
        "Region": ["East", "West", "East", "North"],
        "Sales": [1000, 500, 300, 1200],
        "Profit": [200, 80, 45, 260],
        "Order Date": pd.to_datetime(["2026-01-03", "2026-01-04", "2026-02-10", "2026-02-12"]),
    })


def test_pdf_export_is_valid_and_contains_report_text(exports, sales_data):
    insights = exports["make_insights"](sales_data)
    pdf_bytes = exports["make_pdf"](sales_data, "sales_test.csv", insights)
    assert isinstance(pdf_bytes, bytes)
    assert pdf_bytes.startswith(b"%PDF-")
    assert len(pdf_bytes) > 1000


def test_powerpoint_export_is_valid_and_has_expected_slides(exports, sales_data):
    insights = exports["make_insights"](sales_data)
    pptx_bytes = exports["make_pptx"](sales_data, "sales_test.csv", insights)
    assert isinstance(pptx_bytes, bytes)
    assert zipfile.is_zipfile(io.BytesIO(pptx_bytes))
    presentation = exports["Presentation"](io.BytesIO(pptx_bytes))
    assert len(presentation.slides) >= 4
    all_text = " ".join(shape.text for slide in presentation.slides for shape in slide.shapes if shape.has_text_frame)
    assert "DATAWISE AI" in all_text
    assert "Numeric performance" in all_text


def test_excel_export_has_expected_sheets_and_source_values(exports, sales_data):
    from openpyxl import load_workbook
    insights = exports["make_insights"](sales_data)
    xlsx_bytes = exports["make_excel_report"](sales_data, "sales_test.csv", insights)
    assert isinstance(xlsx_bytes, bytes)
    assert zipfile.is_zipfile(io.BytesIO(xlsx_bytes))
    workbook = load_workbook(io.BytesIO(xlsx_bytes), data_only=True)
    assert "Executive Summary" in workbook.sheetnames
    assert "Data" in workbook.sheetnames
    assert workbook["Executive Summary"]["A1"].value == "DATAWISE AI | BUSINESS REPORT"
    data_sheet = workbook["Data"]
    assert data_sheet["A1"].value == "Product"
    assert data_sheet["C1"].value == "Sales"
    assert data_sheet["C2"].value == 1000


def test_csv_export_round_trip_preserves_rows_and_values(sales_data):
    csv_bytes = sales_data.to_csv(index=False).encode("utf-8")
    restored = pd.read_csv(io.BytesIO(csv_bytes), parse_dates=["Order Date"])
    assert len(restored) == len(sales_data)
    assert restored["Sales"].sum() == 3000
    assert restored["Profit"].sum() == 585
    assert restored.loc[0, "Product"] == "Laptop"
