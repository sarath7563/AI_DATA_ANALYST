import pandas as pd
from conftest import make_correlation_matrix, summarize_category_performance, aggregate_data, make_insights


def test_correlation_matrix_has_expected_relationship():
    df = pd.DataFrame({"Sales": [1, 2, 3, 4], "Profit": [2, 4, 6, 8], "RegionCode": [1, 1, 2, 2]})
    corr = make_correlation_matrix(df, ["Sales", "Profit", "RegionCode"])
    assert corr.loc["Sales", "Profit"] == 1.0
    assert "Sales" in corr.columns


def test_correlation_returns_empty_when_only_one_usable_column():
    df = pd.DataFrame({"Sales": [1, 2, 3], "Constant": [5, 5, 5]})
    assert make_correlation_matrix(df, ["Sales", "Constant"]).empty


def test_category_sum_orders_categories_and_calculates_share():
    df = pd.DataFrame({"Product": ["Laptop", "Phone", "Laptop"], "Sales": [100, 50, 150]})
    result = summarize_category_performance(df, "Product", "Sales", "Sum")
    assert result.iloc[0]["Product"] == "Laptop"
    assert result.iloc[0]["Sum"] == 250
    assert result.iloc[0]["Share of total (%)"] == 83.33


def test_category_average_aggregation():
    df = pd.DataFrame({"Product": ["Laptop", "Laptop", "Phone"], "Sales": [100, 300, 200]})
    result = summarize_category_performance(df, "Product", "Sales", "Average")
    laptop = result.loc[result["Product"] == "Laptop", "Average"].iloc[0]
    assert laptop == 200


def test_aggregate_data_sum():
    df = pd.DataFrame({"Region": ["East", "East", "West"], "Sales": [10, 15, 7]})
    grouped, value_col = aggregate_data(df, "Region", "Sales", "Sum")
    totals = dict(zip(grouped["Region"], grouped[value_col]))
    assert totals == {"East": 25, "West": 7}


def test_insights_include_dataset_overview_and_missingness():
    df = pd.DataFrame({"Product": ["Laptop", "Phone", "Phone"], "Sales": [100, None, 50]})
    insights = make_insights(df)
    titles = [title for title, _ in insights]
    assert "Dataset overview" in titles
    assert "Data completeness" in titles
