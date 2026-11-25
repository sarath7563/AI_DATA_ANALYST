import pandas as pd
import pytest

from modules.smart_filtering import answer_filtered_question


@pytest.fixture
def sales_data():
    return pd.DataFrame({
        "Product": [
            "Laptop",
            "Phone",
            "Laptop",
            "Tablet",
        ],
        "Region": [
            "South",
            "North",
            "East",
            "South",
        ],
        "Sales": [
            2000,
            500,
            1500,
            700,
        ],
        "Profit": [
            200,
            100,
            300,
            140,
        ],
    })


def test_sales_greater_than(sales_data):
    result = answer_filtered_question(
        sales_data,
        "Show sales greater than 1000",
    )

    assert result["value"] == 2
    assert len(result["table"]) == 2


def test_sales_less_than(sales_data):
    result = answer_filtered_question(
        sales_data,
        "Show sales less than 1000",
    )

    assert result["value"] == 2
    assert len(result["table"]) == 2


def test_filter_by_product(sales_data):
    result = answer_filtered_question(
        sales_data,
        "Show sales for product Laptop",
    )

    assert result["value"] == 2
    assert len(result["table"]) == 2
    assert set(result["table"]["Product"]) == {"Laptop"}


def test_total_sales_for_laptop(sales_data):
    result = answer_filtered_question(
        sales_data,
        "Calculate total sales for product Laptop",
    )

    assert result["value"] == 3500


def test_average_profit_for_laptop(sales_data):
    result = answer_filtered_question(
        sales_data,
        "Calculate average profit for product Laptop",
    )

    assert result["value"] == 250


def test_count_south_records(sales_data):
    result = answer_filtered_question(
        sales_data,
        "How many records are from South region?",
    )

    assert result["value"] == 2


def test_no_matching_records(sales_data):
    result = answer_filtered_question(
        sales_data,
        "Show sales for product Camera",
    )

    assert result["value"] == 0
    assert result["table"].empty


def test_invalid_dataframe():
    result = answer_filtered_question(
        None,
        "Show sales greater than 1000",
    )

    assert "valid pandas DataFrame" in result["answer"]


def test_empty_question(sales_data):
    result = answer_filtered_question(
        sales_data,
        "",
    )

    assert "enter a question" in result["answer"].lower()
