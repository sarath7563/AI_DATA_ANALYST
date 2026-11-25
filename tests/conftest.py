"""Load pure helper functions from app.py without starting the Streamlit UI."""
import ast
from pathlib import Path
import io
import re
import numpy as np
import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parents[1]
APP_FILE = PROJECT_ROOT / "app.py"

FUNCTIONS_TO_LOAD = {
    "safe_text",
    "numeric_columns",
    "date_columns",
    "clean_column_names",
    "data_quality",
    "make_insights",
    "detect_iqr_outliers",
    "make_correlation_matrix",
    "summarize_category_performance",
    "aggregate_data",
    "read_uploaded_file",
}

def pytest_sessionstart(session):
    if not APP_FILE.exists():
        raise FileNotFoundError(
            f"Could not find {APP_FILE}. Put the tests folder beside your app.py file."
        )

# Extract only selected helper definitions so importing app.py does not launch Streamlit.
source = APP_FILE.read_text(encoding="utf-8")
tree = ast.parse(source, filename=str(APP_FILE))
selected = [
    node for node in tree.body
    if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef))
    and node.name in FUNCTIONS_TO_LOAD
]
loaded_names = {node.name for node in selected}
missing = FUNCTIONS_TO_LOAD - loaded_names
if missing:
    raise RuntimeError(
        "The app.py version does not contain expected helper functions: "
        + ", ".join(sorted(missing))
        + ". Make sure you are using the Stage 2 app.py."
    )

namespace = {"np": np, "pd": pd, "re": re, "Path": Path, "io": io}
exec(compile(ast.Module(body=selected, type_ignores=[]), str(APP_FILE), "exec"), namespace)

# Expose helper functions to all test modules.
for _name in FUNCTIONS_TO_LOAD:
    globals()[_name] = namespace[_name]
