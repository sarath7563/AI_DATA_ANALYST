import pandas as pd
import numpy as np


def get_dataset_overview(df):
    """Return a general overview of an uploaded dataset."""
    if df is None or df.empty:
        return {
            "rows": 0,
            "columns": 0,
            "missing_cells": 0,
            "duplicate_rows": 0,
            "memory_usage_mb": 0,
        }

    return {
        "rows": int(df.shape[0]),
        "columns": int(df.shape[1]),
        "missing_cells": int(df.isna().sum().sum()),
        "duplicate_rows": int(df.duplicated().sum()),
        "memory_usage_mb": round(
            df.memory_usage(deep=True).sum() / (1024 ** 2), 2
        ),
    }


def get_numeric_summary(df):
    """Calculate descriptive statistics for numeric columns."""
    if df is None or df.empty:
        return pd.DataFrame()

    numeric_df = df.select_dtypes(include="number")

    if numeric_df.empty:
        return pd.DataFrame()

    summary = numeric_df.describe().T
    summary["missing_values"] = numeric_df.isna().sum()
    summary["unique_values"] = numeric_df.nunique()

    return summary.round(2)


def find_top_categories(df, category_column, value_column, top_n=5):
    """Rank categories by the sum of a numeric measure."""
    if df is None or df.empty:
        return pd.DataFrame()

    if category_column not in df.columns or value_column not in df.columns:
        return pd.DataFrame()

    values = pd.to_numeric(df[value_column], errors="coerce")

    working_df = pd.DataFrame({
        "Category": df[category_column],
        "Value": values,
    }).dropna()

    if working_df.empty:
        return pd.DataFrame()

    result = (
        working_df.groupby("Category", dropna=False)["Value"]
        .sum()
        .sort_values(ascending=False)
        .head(max(1, int(top_n)))
        .reset_index()
    )

    result.columns = [category_column, f"Total {value_column}"]

    return result


def detect_numeric_outliers(df):
    """Flag potential outliers using the IQR rule."""
    results = []

    if df is None or df.empty:
        return pd.DataFrame(
            columns=["Column", "Outlier Count", "Lower Bound", "Upper Bound"]
        )

    for column in df.select_dtypes(include="number").columns:
        values = df[column].dropna()

        if len(values) < 4:
            continue

        q1 = values.quantile(0.25)
        q3 = values.quantile(0.75)
        iqr = q3 - q1

        if pd.isna(iqr) or iqr == 0:
            continue

        lower_bound = q1 - 1.5 * iqr
        upper_bound = q3 + 1.5 * iqr

        outlier_count = int(
            ((values < lower_bound) | (values > upper_bound)).sum()
        )

        results.append({
            "Column": column,
            "Outlier Count": outlier_count,
            "Lower Bound": round(float(lower_bound), 2),
            "Upper Bound": round(float(upper_bound), 2),
        })

    return pd.DataFrame(results)


def generate_business_insights(df):
    """Generate deterministic, data-backed business findings."""
    insights = []

    if df is None or df.empty:
        return ["Upload a dataset containing records to generate insights."]

    overview = get_dataset_overview(df)

    insights.append(
        f"The dataset contains {overview['rows']:,} rows "
        f"and {overview['columns']:,} columns."
    )

    if overview["missing_cells"] > 0:
        insights.append(
            f"There are {overview['missing_cells']:,} missing cells. "
            "Review them before drawing conclusions."
        )
    else:
        insights.append("No missing cells were detected.")

    if overview["duplicate_rows"] > 0:
        insights.append(
            f"{overview['duplicate_rows']:,} duplicate rows were detected. "
            "Check whether they represent repeated transactions."
        )
    else:
        insights.append("No duplicate rows were detected.")

    numeric_df = df.select_dtypes(include="number")

    if numeric_df.empty:
        insights.append(
            "No numeric columns were detected, so numeric performance "
            "calculations are unavailable."
        )
    else:
        for column in numeric_df.columns:
            values = numeric_df[column].dropna()

            if values.empty:
                continue

            total = values.sum()
            average = values.mean()

            insights.append(
                f"{column}: total {total:,.2f}; "
                f"average {average:,.2f} across {len(values):,} "
                "non-missing numeric records."
            )

            if (values < 0).any():
                insights.append(
                    f"{column} contains negative values. "
                    "Check whether these are valid for this measure."
                )

    outliers = detect_numeric_outliers(df)

    if not outliers.empty:
        flagged = outliers[outliers["Outlier Count"] > 0]

        for _, row in flagged.iterrows():
            insights.append(
                f"{row['Column']} has {int(row['Outlier Count'])} "
                "potential outlier(s) under the IQR rule. "
                "These are candidates for review, not necessarily errors."
            )

    return insights


def calculate_profit_margin(total_sales, total_profit):
    """Calculate profit margin as a percentage when sales are nonzero."""
    try:
        sales = float(total_sales)
        profit = float(total_profit)

        if not np.isfinite(sales) or not np.isfinite(profit) or sales == 0:
            return None

        return round((profit / sales) * 100, 2)

    except (TypeError, ValueError):
        return None
