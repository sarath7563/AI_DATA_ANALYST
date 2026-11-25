import pandas as pd
from conftest import data_quality, clean_column_names, numeric_columns, detect_iqr_outliers


def test_data_quality_counts_missing_duplicates_and_completeness():
    df = pd.DataFrame({"Product": ["A", "B", "B"], "Sales": [100, None, None]})
    missing, duplicates, completeness = data_quality(df)
    assert missing == 2
    assert duplicates == 1
    assert round(completeness, 2) == 66.67


def test_clean_column_names_removes_spaces_and_empty_names():
    df = pd.DataFrame([[1, 2, 3]], columns=[" Product Name ", "Sales Amount", "   "])
    cleaned = clean_column_names(df)
    assert cleaned.columns.tolist() == ["Product_Name", "Sales_Amount", "column_3"]


def test_numeric_columns_detects_numeric_fields_only():
    df = pd.DataFrame({"Region": ["East", "West"], "Sales": [10, 20], "Profit": [2.5, 3.5]})
    assert numeric_columns(df) == ["Sales", "Profit"]


def test_outlier_screening_flags_extreme_candidate():
    df = pd.DataFrame({"Sales": [10, 11, 12, 13, 14, 1000]})
    result = detect_iqr_outliers(df, ["Sales"])
    assert not result.empty
    assert result.loc[0, "Outlier candidates"] == 1
    assert result.loc[0, "Maximum flagged"] == 1000


def test_outlier_screening_handles_constant_column():
    df = pd.DataFrame({"Sales": [5, 5, 5, 5, 5]})
    assert detect_iqr_outliers(df, ["Sales"]).empty
