"""
DataWise AI - Advanced Rule-Based Question Answering

Stage 4.3A:
- Flexible question wording
- Better column-name matching
- Common business terminology
- Numeric calculations
- Highest/lowest category analysis
- Dataset quality questions

No API key required.
"""

import re

import pandas as pd


# ==================================================
# 1. TEXT UTILITIES
# ==================================================

def _normalise(text):
    """
    Convert text to lowercase.
    Treat spaces, underscores, and hyphens consistently.
    """
    text = str(text).strip().lower()
    text = re.sub(r"[_\-]+", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text


def _format_number(value):
    """Format numeric results for display."""

    if pd.isna(value):
        return "N/A"

    value = float(value)

    if value.is_integer():
        return f"{int(value):,}"

    return f"{value:,.2f}"


def _make_result(
    answer,
    table=None,
    metric=None,
    value=None,
    chart_type=None,
    explanation="",
):
    """Return a consistent result dictionary."""

    if table is None:
        table = pd.DataFrame()

    return {
        "answer": answer,
        "table": table,
        "metric": metric,
        "value": value,
        "chart_type": chart_type,
        "explanation": explanation,
    }


def _contains_phrase(text, phrase):
    """Check whether a complete word or phrase occurs."""

    text = _normalise(text)
    phrase = _normalise(phrase)

    if not phrase:
        return False

    return bool(
        re.search(
            r"(?<!\w)" + re.escape(phrase) + r"(?!\w)",
            text,
        )
    )


# ==================================================
# 2. COLUMN NAME RECOGNITION
# ==================================================

COLUMN_ALIASES = {
    "sales": [
        "revenue",
        "turnover",
        "sales amount",
        "total sales",
        "sales value",
    ],
    "revenue": [
        "sales",
        "turnover",
        "sales amount",
    ],
    "profit": [
        "earnings",
        "net profit",
        "profit amount",
    ],
    "quantity": [
        "units",
        "units sold",
        "quantity sold",
    ],
    "price": [
        "unit price",
        "selling price",
        "product price",
    ],
    "cost": [
        "expense",
        "expenses",
        "cost amount",
    ],
    "customer": [
        "client",
        "buyer",
        "customer name",
    ],
    "product": [
        "item",
        "product name",
    ],
    "region": [
        "area",
        "territory",
        "location",
    ],
}


def _find_column(df, question):
    """
    Find a column mentioned in the question.

    Matching order:
    1. Exact column-name phrase.
    2. Business terminology aliases.

    Longer column names are checked first.
    """

    question_normalised = _normalise(question)

    columns = sorted(
        df.columns,
        key=lambda col: len(_normalise(col)),
        reverse=True,
    )

    # First: match actual column names.
    for column in columns:
        column_name = _normalise(column)

        if _contains_phrase(question_normalised, column_name):
            return column

    # Second: match aliases.
    for column in columns:
        column_name = _normalise(column)

        aliases = COLUMN_ALIASES.get(column_name, [])

        for alias in aliases:
            if _contains_phrase(question_normalised, alias):
                return column

    return None


# ==================================================
# 3. QUESTION INTENT RECOGNITION
# ==================================================

def _detect_operation(question):
    """
    Identify the requested calculation.

    Returns:
        tuple: (operation, display_label)
    """

    q = _normalise(question)

    # Check specific operations before general ones.
    if (
        re.search(r"\b(average|avg|mean)\b", q)
        or "on average" in q
    ):
        return "mean", "Average"

    if re.search(r"\bmedian\b", q):
        return "median", "Median"

    if re.search(
        r"\b(maximum|max|highest|greatest|largest)\b",
        q,
    ):
        return "max", "Maximum"

    if re.search(
        r"\b(minimum|min|lowest|smallest)\b",
        q,
    ):
        return "min", "Minimum"

    if re.search(
        r"\b(count|number of|how many)\b",
        q,
    ):
        return "count", "Count"

    if re.search(
        r"\b(total|sum|overall|combined)\b",
        q,
    ):
        return "sum", "Total"

    # Additional wording for totals.
    if re.search(
        r"\bhow much\b.*\b(generate|generated|earn|earned|make|made)\b",
        q,
    ):
        return "sum", "Total"

    return None, None


# ==================================================
# 4. MAIN QUESTION-ANSWERING FUNCTION
# ==================================================

def answer_question(df, question):
    """
    Answer supported questions using a pandas DataFrame.

    Supported categories:
    - Missing values
    - Duplicate rows
    - Row and column counts
    - Numeric sum, average, median, maximum,
      minimum, and count
    - Highest/lowest category rankings

    Returns a dictionary containing:
        answer, table, metric, value, chart_type, explanation
    """

    # --------------------------------------------------
    # Validate input
    # --------------------------------------------------

    if not isinstance(df, pd.DataFrame):
        return _make_result(
            "Please provide a valid pandas DataFrame."
        )

    if not isinstance(question, str) or not question.strip():
        return _make_result(
            "Please enter a question about your dataset."
        )

    if df.empty or len(df.columns) == 0:
        return _make_result(
            "The dataset is empty. Please upload a dataset containing data."
        )

    q = _normalise(question)

    # --------------------------------------------------
    # A. DATASET-LEVEL QUESTIONS
    # --------------------------------------------------

    if re.search(
        r"\b(missing|null|empty cells|blank cells)\b",
        q,
    ):
        missing_by_column = df.isna().sum()
        total_missing = int(missing_by_column.sum())

        table = (
            missing_by_column[missing_by_column > 0]
            .rename("Missing Values")
            .rename_axis("Column")
            .reset_index()
        )

        return _make_result(
            f"The dataset contains {total_missing:,} missing values.",
            table=table,
            metric="Missing values",
            value=total_missing,
            explanation="Counted missing cells across all columns.",
        )

    if re.search(r"\bduplicate(s)?\b", q):
        count = int(df.duplicated().sum())

        return _make_result(
            f"The dataset contains {count:,} duplicate rows.",
            metric="Duplicate rows",
            value=count,
            explanation="Counted rows that duplicate an earlier row.",
        )

    if (
        "how many rows" in q
        or "number of rows" in q
        or "row count" in q
        or "total rows" in q
    ):
        return _make_result(
            f"The dataset contains {len(df):,} rows.",
            metric="Rows",
            value=len(df),
        )

    if (
        "how many columns" in q
        or "number of columns" in q
        or "column count" in q
        or "total columns" in q
    ):
        return _make_result(
            f"The dataset contains {len(df.columns):,} columns.",
            metric="Columns",
            value=len(df.columns),
        )

    # --------------------------------------------------
    # B. DETECT RANKING QUESTIONS
    # --------------------------------------------------

    ranking_words = {
        "highest",
        "lowest",
        "best",
        "worst",
        "top",
        "bottom",
        "most",
        "least",
        "largest",
        "smallest",
        "greatest",
    }

    asks_for_ranking = any(
        word in q.split()
        for word in ranking_words
    )

    numeric_columns = df.select_dtypes(
        include="number"
    ).columns.tolist()

    category_columns = [
        column
        for column in df.columns
        if column not in numeric_columns
    ]

    if asks_for_ranking:
        category_column = None

        # Example:
        # Which product has the highest sales?
        for column in sorted(
            category_columns,
            key=lambda col: len(_normalise(col)),
            reverse=True,
        ):
            if re.search(
                r"\bwhich\s+"
                + re.escape(_normalise(column))
                + r"\b",
                q,
            ):
                category_column = column
                break

        # Example:
        # Highest sales by product
        if category_column is None:
            for column in sorted(
                category_columns,
                key=lambda col: len(_normalise(col)),
                reverse=True,
            ):
                if (
                    _contains_phrase(q, column)
                    and re.search(r"\b(by|per|among)\b", q)
                ):
                    category_column = column
                    break

        measure_column = None

        for column in sorted(
            numeric_columns,
            key=lambda col: len(_normalise(col)),
            reverse=True,
        ):
            if _contains_phrase(q, column):
                measure_column = column
                break

        if measure_column is None:
            measure_column = _find_column(df, q)

        if category_column is not None and measure_column is not None:
            grouped = df.groupby(
                category_column,
                dropna=False,
            )[measure_column].sum()

            if grouped.empty:
                return _make_result(
                    "There is no data available for this comparison."
                )

            lowest_requested = bool(
                re.search(
                    r"\b(lowest|worst|least|bottom|smallest)\b",
                    q,
                )
            )

            grouped = grouped.sort_values(
                ascending=lowest_requested
            )

            selected_category = grouped.index[0]
            selected_value = float(grouped.iloc[0])

            ranking_label = (
                "lowest" if lowest_requested else "highest"
            )

            display_category = (
                "Missing category"
                if pd.isna(selected_category)
                else str(selected_category)
            )

            result_table = (
                grouped.rename(measure_column)
                .reset_index()
            )

            return _make_result(
                (
                    f"{display_category} has the {ranking_label} "
                    f"total {measure_column}: "
                    f"{_format_number(selected_value)}."
                ),
                table=result_table,
                metric=f"{ranking_label.title()} {measure_column}",
                value=selected_value,
                chart_type="bar",
                explanation=(
                    f"Grouped '{measure_column}' by "
                    f"'{category_column}' and summed each group."
                ),
            )

        if category_column is not None and measure_column is None:
            return _make_result(
                "I identified the category, but couldn't identify "
                "the numeric measure. Include a column name, such "
                "as 'Which product has the highest sales?'"
            )

    # --------------------------------------------------
    # C. NUMERIC AGGREGATIONS
    # --------------------------------------------------

    measure_column = _find_column(df, q)

    if measure_column is None:
        return _make_result(
            "I couldn't identify the requested column. Try asking "
            "for the total sales, average profit, or highest sales "
            "by product. You can also mention an exact column name."
        )

    if not pd.api.types.is_numeric_dtype(df[measure_column]):
        return _make_result(
            f"The column '{measure_column}' is not numeric. "
            "Please ask about a numeric column."
        )

    values = df[measure_column].dropna()

    if values.empty:
        return _make_result(
            f"The column '{measure_column}' has no numeric values."
        )

    operation, label = _detect_operation(q)

    if operation is None:
        return _make_result(
            f"I found the column '{measure_column}', but couldn't "
            "identify the calculation. Try total, average, median, "
            "maximum, minimum, or count."
        )

    if operation == "sum":
        result_value = float(values.sum())

    elif operation == "mean":
        result_value = float(values.mean())

    elif operation == "median":
        result_value = float(values.median())

    elif operation == "max":
        result_value = float(values.max())

    elif operation == "min":
        result_value = float(values.min())

    elif operation == "count":
        result_value = int(values.count())

    else:
        return _make_result(
            "I couldn't process that calculation."
        )

    return _make_result(
        (
            f"{label} {measure_column}: "
            f"{_format_number(result_value)}."
        ),
        metric=f"{label} {measure_column}",
        value=result_value,
        explanation=(
            f"Calculated {label.lower()} using the "
            f"'{measure_column}' column after excluding missing values."
        ),
    )
