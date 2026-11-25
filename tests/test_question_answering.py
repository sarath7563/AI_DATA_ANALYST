import pandas as pd
import pytest

from modules.question_answering import answer_question


@pytest.fixture
def sample_data():
    return pd.DataFrame({
        "Product": ["Laptop", "Phone", "Laptop", "Tablet"],
        "Region": ["South", "North", "East", "South"],
        "Sales": [2000, 500, 1500, 700],
        "Profit": [200, 100, 300, 140],
    })


def test_total_sales(sample_data):
    result = answer_question(sample_data, "What is the total sales?")

    assert result["value"] == 4700
    assert "4,700" in result["answer"]


def test_average_profit(sample_data):
    result = answer_question(sample_data, "What is the average profit?")

    assert result["value"] == 185
    assert "185" in result["answer"]


def test_highest_sales_product(sample_data):
    result = answer_question(
        sample_data,
        "Which product has the highest sales?",
    )

    assert "Laptop" in result["answer"]
    assert result["value"] == 3500


def test_lowest_sales_product(sample_data):
    result = answer_question(
        sample_data,
        "Which product has the lowest sales?",
    )

    assert "Phone" in result["answer"]
    assert result["value"] == 500


def test_missing_values(sample_data):
    result = answer_question(
        sample_data,
        "How many missing values are there?",
    )

    assert result["value"] == 0


def test_duplicate_rows(sample_data):
    result = answer_question(
        sample_data,
        "How many duplicate rows exist?",
    )

    assert result["value"] == 0


def test_empty_question(sample_data):
    result = answer_question(sample_data, "")

    assert "enter a question" in result["answer"].lower()


def test_empty_dataframe():
    result = answer_question(
        pd.DataFrame(),
        "What is the total sales?",
    )

    assert "empty" in result["answer"].lower()