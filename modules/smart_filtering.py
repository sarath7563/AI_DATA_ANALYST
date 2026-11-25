import re
import pandas as pd


def _result(answer, filtered_df=None, success=True):
    return {
        "success": success,
        "answer": answer,
        "filtered_data": (
            filtered_df.copy()
            if isinstance(filtered_df, pd.DataFrame)
            else pd.DataFrame()
        ),
    }


def _normalise(text):
    return re.sub(r"\s+", " ", str(text).strip().lower())


def _find_column(df, question):
    q = _normalise(question)

    for column in df.columns:
        name = _normalise(column)

        if name and re.search(
            r"(?<!\w)" + re.escape(name) + r"(?!\w)", q
        ):
            return column

    aliases = {
        "sales": ["revenue", "turnover"],
        "profit": ["earnings", "net profit"],
        "product": ["item", "product name"],
        "category": ["product category"],
        "quantity": ["units sold"],
        "price": ["selling price", "unit price"],
        "customer": ["client", "buyer"],
        "region": ["area", "territory"],
    }

    for column in df.columns:
        name = _normalise(column)

        for original, alternatives in aliases.items():
            if name == original:
                for alias in alternatives:
                    if re.search(
                        r"(?<!\w)" + re.escape(alias) + r"(?!\w)", q
                    ):
                        return column

    return None


def _find_category_filter(df, question, target_column=None):
    import re
    import pandas as pd

    question_text = str(question).strip().lower()

    for column in df.columns:
        if not (
            pd.api.types.is_object_dtype(df[column])
            or isinstance(df[column].dtype, pd.CategoricalDtype)
            or pd.api.types.is_string_dtype(df[column])
        ):
            continue

        unique_values = df[column].dropna().unique()

        for value in unique_values:
            value_text = str(value).strip()

            if not value_text:
                continue

            pattern = (
                r"(?<!\w)"
                + re.escape(value_text.lower())
                + r"(?!\w)"
            )

            if re.search(pattern, question_text):
                return column, value

    return None, None


def answer_filtered_question(df, question):
    """Answer questions about a dataset and return filtered rows."""

    if not isinstance(df, pd.DataFrame):
        return _result("Please provide a valid DataFrame.", success=False)

    if df.empty:
        return _result("The dataset is empty.", success=False)

    if not isinstance(question, str) or not question.strip():
        return _result("Please enter a question.", success=False)

    q = _normalise(question)

    # Detect the operation.
    if any(word in q for word in ["show", "display", "list", "filter"]):
        operation = "show"
    elif any(word in q for word in ["average", "mean"]):
        operation = "average"
    elif any(word in q for word in ["highest", "maximum", "largest"]):
        operation = "maximum"
    elif any(word in q for word in ["lowest", "minimum", "smallest"]):
        operation = "minimum"
    elif any(word in q for word in ["total", "sum"]):
        operation = "sum"
    elif any(word in q for word in ["count", "how many"]):
        operation = "count"
    else:
        return _result(
            "Try asking for show, total, average, highest, lowest, or count.",
            success=False,
        )

    target_column = _find_column(df, question)

    # Apply category filtering.
    filter_column, filter_value = _find_category_filter(
        df, question, target_column
    )

    filtered_df = df.copy()

    if filter_column is not None:
        filtered_df = df[
            df[filter_column].astype(str).str.strip().str.lower()
            == str(filter_value).strip().lower()
        ].copy()

    if filtered_df.empty:
        return _result("No matching records were found.", filtered_df)

    if operation == "show":
        return _result(
            f"Found {len(filtered_df):,} matching records.",
            filtered_df,
        )

    if operation == "count":
        if target_column is None:
            return _result(
                f"Total records: {len(filtered_df):,}.",
                filtered_df,
            )

        return _result(
            f"Count: {filtered_df[target_column].count():,}.",
            filtered_df,
        )

    if target_column is None:
        numeric_columns = filtered_df.select_dtypes(
            include="number"
        ).columns.tolist()

        if len(numeric_columns) != 1:
            return _result(
                "Please specify a numeric column, such as Sales or Profit.",
                filtered_df,
                success=False,
            )

        target_column = numeric_columns[0]

    numbers = pd.to_numeric(
        filtered_df[target_column], errors="coerce"
    ).dropna()

    if numbers.empty:
        return _result(
            f"No numeric values found in {target_column}.",
            filtered_df,
            success=False,
        )

    if operation == "sum":
        answer = f"Total {target_column}: {numbers.sum():,.2f}."
    elif operation == "average":
        answer = f"Average {target_column}: {numbers.mean():,.2f}."
    elif operation == "maximum":
        answer = f"Highest {target_column}: {numbers.max():,.2f}."
    else:
        answer = f"Lowest {target_column}: {numbers.min():,.2f}."

    return _result(answer, filtered_df)

