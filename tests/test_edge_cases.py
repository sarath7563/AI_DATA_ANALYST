import pandas as pd
import numpy as np
import pytest


def test_empty_dataset():
    """An empty DataFrame should contain no rows."""
    df = pd.DataFrame()
    assert df.empty


def test_dataset_with_missing_values():
    """Missing values should be identifiable."""
    df = pd.DataFrame({
        "Product": ["Laptop", "Phone", "Tablet"],
        "Sales": [1000, np.nan, 3000],
    })

    assert df["Sales"].isna().sum() == 1


def test_duplicate_rows():
    """Duplicate rows should be detectable."""
    df = pd.DataFrame({
        "Product": ["Laptop", "Phone", "Laptop"],
        "Sales": [1000, 2000, 1000],
    })

    assert df.duplicated().sum() == 1


def test_mixed_data_types():
    """Text in a numeric column should be identifiable."""
    df = pd.DataFrame({
        "Sales": ["1000", "2000", "invalid", "4000"]
    })

    converted = pd.to_numeric(df["Sales"], errors="coerce")

    assert converted.isna().sum() == 1


def test_large_dataset():
    """A dataset with 10,000 rows should be created correctly."""
    df = pd.DataFrame({
        "Sales": range(1, 10001),
        "Profit": range(101, 10101),
    })

    assert len(df) == 10000
    assert df["Sales"].sum() == 50005000


def test_all_values_missing():
    """A column containing only missing values should be detected."""
    df = pd.DataFrame({
        "Sales": [np.nan, np.nan, np.nan]
    })

    assert df["Sales"].isna().all()


def test_zero_sales():
    """Zero sales should be preserved as a valid numeric value."""
    df = pd.DataFrame({
        "Sales": [0, 1000, 2000]
    })

    assert df["Sales"].iloc[0] == 0


def test_negative_values():
    """Negative values should be detectable for further investigation."""
    df = pd.DataFrame({
        "Profit": [100, -50, 200]
    })

    assert (df["Profit"] < 0).sum() == 1


def test_empty_strings():
    """Empty strings should be identifiable."""
    df = pd.DataFrame({
        "Product": ["Laptop", "", "Phone"]
    })

    assert (df["Product"] == "").sum() == 1


def test_missing_column():
    """A missing expected column should be detected."""
    df = pd.DataFrame({
        "Product": ["Laptop", "Phone"]
    })

    assert "Sales" not in df.columns
