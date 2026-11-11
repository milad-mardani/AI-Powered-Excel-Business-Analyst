# & "C:\Users\Asus\anaconda3\envs\Gen\python.exe" -m streamlit run app.py
from io import BytesIO
import re
import unicodedata

import pandas as pd
import streamlit as st

try:
    import ollama
except ImportError:
    ollama = None

from openpyxl.styles import (
    Font,
    PatternFill,
    Border,
    Side,
    Alignment,
)
from openpyxl.utils import get_column_letter


# =========================================================
# CONFIGURATION
# =========================================================

MAX_FILE_SIZE = 10 * 1024 * 1024
OLLAMA_MODEL = "qwen2.5:7b"

REQUIRED_COLUMNS = [
    "OrderID",
    "Product",
    "Region",
    "Quantity",
    "UnitPrice",
    "Status",
]

st.set_page_config(
    page_title="AI-Powered Excel Business Analyst",
    page_icon="📊",
    layout="wide",
)


# =========================================================
# 1. DATA LOADING
# =========================================================

def load_data(uploaded_file):

    file_bytes = uploaded_file.getvalue()
    file_name = uploaded_file.name.lower()

    if not file_bytes:
        raise ValueError("The uploaded file is empty.")

    if len(file_bytes) > MAX_FILE_SIZE:
        raise ValueError("Maximum file size is 10 MB.")

    if file_name.endswith(".csv"):

        try:
            df = pd.read_csv(
                BytesIO(file_bytes),
                encoding="utf-8-sig",
            )

        except UnicodeDecodeError:
            df = pd.read_csv(
                BytesIO(file_bytes),
                encoding="cp1252",
            )

    elif file_name.endswith(".xlsx"):

        df = pd.read_excel(
            BytesIO(file_bytes),
            engine="openpyxl",
        )

    else:
        raise ValueError("Please upload CSV or XLSX.")

    if df.empty:
        raise ValueError("The uploaded file contains no rows.")

    df.columns = [
        str(column).strip()
        for column in df.columns
    ]

    if len(df.columns) != len(set(df.columns)):
        raise ValueError("Duplicate column names detected.")

    if any(not column for column in df.columns):
        raise ValueError("An empty column name was detected.")

    return df


# =========================================================
# 2. DATA PREPARATION
# =========================================================

def prepare_sales_data(df):

    df = df.copy()

    df.columns = [
        str(column).strip()
        for column in df.columns
    ]

    missing_columns = [
        column
        for column in REQUIRED_COLUMNS
        if column not in df.columns
    ]

    if missing_columns:
        raise ValueError(
            "Missing required columns: "
            + ", ".join(missing_columns)
        )

    for column in [
        "OrderID",
        "Product",
        "Region",
        "Status",
    ]:

        df[column] = (
            df[column]
            .astype("string")
            .str.strip()
        )

    for column in [
        "Quantity",
        "UnitPrice",
    ]:

        df[column] = pd.to_numeric(
            df[column],
            errors="coerce",
        )

    valid_numeric = (
        df["Quantity"].notna()
        & df["UnitPrice"].notna()
        & df["Quantity"].ge(0)
        & df["UnitPrice"].ge(0)
    )

    df["Revenue"] = (
        df["Quantity"] * df["UnitPrice"]
    ).where(valid_numeric)

    return df


# =========================================================
# 3. ANALYTICAL ENGINE
# =========================================================

def calculate_sales_metrics(df):

    df = df.copy()

    df["Status_Normalized"] = (
        df["Status"]
        .fillna("Unknown")
        .astype("string")
        .str.strip()
        .str.casefold()
        .replace("", "unknown")
    )

    status_display_map = {
        "completed": "Completed",
        "complete": "Completed",
        "pending": "Pending",
        "cancelled": "Cancelled",
        "canceled": "Cancelled",
        "unknown": "Unknown",
    }

    df["Status_Display"] = (
        df["Status_Normalized"]
        .map(status_display_map)
        .fillna(
            df["Status"]
            .fillna("Unknown")
            .astype("string")
            .str.strip()
            .replace("", "Unknown")
        )
    )

    completed = df[
        df["Status_Normalized"].isin(
            ["completed", "complete"]
        )
    ].copy()

    valid_revenue = completed[
        completed["Revenue"].notna()
    ].copy()

    total_revenue = float(
        valid_revenue["Revenue"].sum()
    )

    completed_order_count = int(
        valid_revenue["OrderID"]
        .dropna()
        .nunique()
    )

    order_revenue = (
        valid_revenue
        .dropna(subset=["OrderID"])
        .groupby("OrderID")["Revenue"]
        .sum()
    )

    average_order_value = (
        float(order_revenue.mean())
        if not order_revenue.empty
        else 0.0
    )

    product_summary = (
        valid_revenue
        .groupby("Product", dropna=False)
        .agg(
            Revenue=("Revenue", "sum"),
            **{
                "Quantity Sold": ("Quantity", "sum")
            },
        )
        .reset_index()
        .sort_values(
            "Revenue",
            ascending=False,
        )
        .reset_index(drop=True)
    )

    region_summary = (
        valid_revenue
        .groupby("Region", dropna=False)["Revenue"]
        .sum()
        .reset_index()
        .sort_values(
            "Revenue",
            ascending=False,
        )
        .reset_index(drop=True)
    )

    status_summary = (
        df["Status_Display"]
        .value_counts()
        .rename_axis("Status")
        .reset_index(name="Order Rows")
    )

    status_order_counts = (
        df
        .dropna(subset=["OrderID"])
        .groupby("Status_Normalized")["OrderID"]
        .nunique()
        .to_dict()
    )

    status_order_counts["cancelled"] = (
        status_order_counts.get("cancelled", 0)
        + status_order_counts.get("canceled", 0)
    )

    status_order_counts["completed"] = (
        status_order_counts.get("completed", 0)
        + status_order_counts.get("complete", 0)
    )

    return {
        "total_revenue": total_revenue,
        "completed_order_count": completed_order_count,
        "average_order_value": average_order_value,
        "product_summary": product_summary,
        "region_summary": region_summary,
        "status_summary": status_summary,
        "status_order_counts": status_order_counts,
        "valid_revenue": valid_revenue,
    }


# =========================================================
# 4. QUESTION ROUTER
# =========================================================

def normalize_question(question):

    question = unicodedata.normalize(
        "NFKC",
        question,
    ).casefold()

    question = re.sub(
        r"[^\w\s]",
        " ",
        question,
    )

    return " ".join(question.split())


def find_extreme_rows(data, column, highest=True):

    if data.empty:
        return data

    target = (
        data[column].max()
        if highest
        else data[column].min()
    )

    return data[
        data[column].eq(target)
    ]


def answer_business_question(question, metrics, df):

    q = normalize_question(question)

    products = metrics["product_summary"]
    regions = metrics["region_summary"]

    # 1. Highest revenue product
    highest_product_phrases = [
        "highest revenue product",
        "most revenue product",
        "top product",
        "best selling product",
        "best performing product",
        "which product made the most money",
        "which product generated the most",
        "which product earned the most",
        "which product has the highest revenue",
    ]

    if any(phrase in q for phrase in highest_product_phrases):

        result = find_extreme_rows(
            products,
            "Revenue",
            highest=True,
        )

        return (
            "Highest Revenue Product:\n\n"
            + result.to_string(index=False)
        )

    # 2. Lowest revenue product
    lowest_product_phrases = [
        "lowest revenue product",
        "least revenue product",
        "worst selling product",
        "which product made the least money",
        "which product generated the least",
        "which product earned the least",
        "which product has the lowest revenue",
    ]

    if any(phrase in q for phrase in lowest_product_phrases):

        result = find_extreme_rows(
            products,
            "Revenue",
            highest=False,
        )

        return (
            "Lowest Revenue Product:\n\n"
            + result.to_string(index=False)
        )

    # 3. Highest revenue region
    highest_region_phrases = [
        "highest revenue region",
        "best region",
        "top region",
        "best performing region",
        "which region performed best",
        "which region generated the most",
        "which region made the most money",
        "which region has the highest revenue",
    ]

    if any(phrase in q for phrase in highest_region_phrases):

        result = find_extreme_rows(
            regions,
            "Revenue",
            highest=True,
        )

        return (
            "Highest Revenue Region:\n\n"
            + result.to_string(index=False)
        )

    # 4. Lowest revenue region
    lowest_region_phrases = [
        "lowest revenue region",
        "worst region",
        "worst performing region",
        "which region performed worst",
        "which region generated the least",
        "which region made the least money",
        "which region has the lowest revenue",
    ]

    if any(phrase in q for phrase in lowest_region_phrases):

        result = find_extreme_rows(
            regions,
            "Revenue",
            highest=False,
        )

        return (
            "Lowest Revenue Region:\n\n"
            + result.to_string(index=False)
        )

    # 5. Product comparison
    product_compare_phrases = [
        "revenue by product",
        "compare products",
        "compare product",
        "compare sales by product",
        "compare sales across products",
        "sales across different products",
        "all products",
        "products by revenue",
        "show product comparison",
    ]

    if any(phrase in q for phrase in product_compare_phrases):

        return (
            "Revenue by Product:\n\n"
            + products.to_string(index=False)
        )

    # 6. Region comparison
    region_compare_phrases = [
        "revenue by region",
        "compare regions",
        "compare region",
        "compare sales across regions",
        "compare sales across different regions",
        "sales across regions",
        "sales across different regions",
        "compare regional sales",
        "all regions",
        "regions by revenue",
        "show region comparison",
    ]

    if any(phrase in q for phrase in region_compare_phrases):

        return (
            "Revenue by Region:\n\n"
            + regions.to_string(index=False)
        )

    # 7. Specific product
    matched_product = None

    for product in products["Product"].dropna().unique():

        product_name = str(product).casefold()

        if product_name and product_name in q:
            matched_product = product
            break

    if matched_product is not None:

        result = products[
            products["Product"].eq(matched_product)
        ]

        return (
            f"Revenue Details for {matched_product}:\n\n"
            + result.to_string(index=False)
        )

    # 8. Total revenue
    total_revenue_phrases = [
        "total revenue",
        "overall revenue",
        "how much revenue",
        "total sales",
        "overall sales",
    ]

    if any(phrase in q for phrase in total_revenue_phrases):

        return (
            f"Total Completed Revenue: "
            f"{metrics['total_revenue']:,.2f}"
        )

    # 9. Average Order Value
    aov_phrases = [
        "average order value",
        "aov",
        "average value per order",
    ]

    if any(phrase in q for phrase in aov_phrases):

        return (
            f"Average Order Value: "
            f"{metrics['average_order_value']:,.2f}"
        )

    # 10. Order status counts
    status_map = {
        "completed": ["completed", "complete"],
        "pending": ["pending"],
        "cancelled": ["cancelled", "canceled"],
    }

    for status, keywords in status_map.items():

        if any(keyword in q for keyword in keywords):

            count = metrics[
                "status_order_counts"
            ].get(status, 0)

            return (
                f"Number of {status.title()} Orders: "
                f"{count:,}"
            )

    # 11. Order status summary
    status_summary_phrases = [
        "order status",
        "all status",
        "status summary",
        "show statuses",
        "show order statuses",
    ]

    if any(phrase in q for phrase in status_summary_phrases):

        return (
            "Order Status Summary:\n\n"
            + metrics["status_summary"].to_string(index=False)
        )

    # 12. Number of different products
    product_count_phrases = [
        "different products",
        "number of products",
        "how many products",
        "count products",
    ]

    if any(phrase in q for phrase in product_count_phrases):

        count = int(
            df["Product"].nunique(dropna=True)
        )

        return (
            f"Number of Different Products: {count:,}"
        )

    # 13. Ollama fallback
    if ollama is None:

        return (
            "This question is not supported by the current "
            "Question Router. The Ollama Python package "
            "is unavailable."
        )

    prompt = f"""
You are an AI Business Analyst.

Dataset columns:
{list(df.columns)}

User question:
{question}

Explain what analysis is required to answer this question.

Important rules:
- Do not invent numerical results.
- Do not claim that you calculated anything.
- Do not generate Python or SQL code.
- Do not make unsupported assumptions.
- Keep the explanation concise.
"""

    response = ollama.chat(
        model=OLLAMA_MODEL,
        messages=[
            {
                "role": "system",
                "content": (
                    "You interpret business questions. "
                    "You do not calculate or invent numerical results."
                ),
            },
            {
                "role": "user",
                "content": prompt,
            },
        ],
        options={
            "temperature": 0,
        },
    )

    return response["message"]["content"]


# =========================================================
# 5. EXCEL REPORT FORMATTING
# =========================================================

def format_excel_workbook(writer):

    # Professional color palette
    navy = "17365D"
    blue = "D9EAF7"
    light_gray = "F4F6F8"
    white = "FFFFFF"
    dark_text = "243447"
    border_color = "D6DEE8"
    warning_fill = "FFF2CC"
    warning_text = "9C6500"

    header_fill = PatternFill(
        fill_type="solid",
        fgColor=navy,
    )

    alternate_fill = PatternFill(
        fill_type="solid",
        fgColor=light_gray,
    )

    summary_fill = PatternFill(
        fill_type="solid",
        fgColor=blue,
    )

    missing_fill = PatternFill(
        fill_type="solid",
        fgColor=warning_fill,
    )

    thin_border = Border(
        bottom=Side(
            style="thin",
            color=border_color,
        )
    )

    # -----------------------------------------------------
    # General formatting for all worksheets
    # -----------------------------------------------------

    for worksheet in writer.sheets.values():

        worksheet.sheet_view.showGridLines = False
        worksheet.freeze_panes = "A2"

        worksheet.auto_filter.ref = (
            worksheet.dimensions
        )

        worksheet.sheet_properties.pageSetUpPr.fitToPage = True
        worksheet.page_setup.fitToWidth = 1
        worksheet.page_setup.fitToHeight = 0
        worksheet.page_setup.orientation = "landscape"

        # Header formatting
        for cell in worksheet[1]:

            cell.fill = header_fill

            cell.font = Font(
                name="Aptos",
                size=11,
                bold=True,
                color=white,
            )

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True,
            )

            cell.border = thin_border

        worksheet.row_dimensions[1].height = 28

        # Data row formatting
        for row_index in range(2, worksheet.max_row + 1):

            worksheet.row_dimensions[row_index].height = 22

            for cell in worksheet[row_index]:

                cell.font = Font(
                    name="Aptos",
                    size=10,
                    color=dark_text,
                )

                cell.alignment = Alignment(
                    vertical="center",
                )

                cell.border = thin_border

                if row_index % 2 == 0:
                    cell.fill = alternate_fill

        # Automatic column widths
        for column_index in range(1, worksheet.max_column + 1):

            column_letter = get_column_letter(column_index)

            max_length = 0

            for cell in worksheet[column_letter]:

                value = (
                    ""
                    if cell.value is None
                    else str(cell.value)
                )

                max_length = max(
                    max_length,
                    len(value),
                )

            worksheet.column_dimensions[
                column_letter
            ].width = min(
                max(max_length + 4, 14),
                40,
            )

        # Default alignment for numeric values
        for row in worksheet.iter_rows(
            min_row=2,
            max_row=worksheet.max_row,
        ):

            for cell in row:

                if isinstance(cell.value, (int, float)):
                    cell.alignment = Alignment(
                        horizontal="right",
                        vertical="center",
                    )

    # -----------------------------------------------------
    # Summary sheet
    # -----------------------------------------------------

    ws = writer.sheets["Summary"]

    ws.sheet_properties.tabColor = "4472C4"

    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 22

    for row_index in range(2, ws.max_row + 1):

        metric_name = ws.cell(
            row=row_index,
            column=1,
        ).value

        value_cell = ws.cell(
            row=row_index,
            column=2,
        )

        value_cell.font = Font(
            name="Aptos",
            size=11,
            bold=True,
            color=navy,
        )

        value_cell.fill = summary_fill

        if metric_name in [
            "Total Completed Revenue",
            "Average Order Value",
        ]:

            value_cell.number_format = '#,##0.00'

        elif metric_name == "Completed Orders":

            value_cell.number_format = '#,##0'

    # -----------------------------------------------------
    # Data Quality sheet
    # -----------------------------------------------------

    ws = writer.sheets["Data Quality"]

    ws.sheet_properties.tabColor = "70AD47"

    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 20

    for row_index in range(2, ws.max_row + 1):

        ws.cell(
            row=row_index,
            column=2,
        ).number_format = '#,##0'

    # -----------------------------------------------------
    # Product Analysis sheet
    # -----------------------------------------------------

    ws = writer.sheets["Product Analysis"]

    ws.sheet_properties.tabColor = "5B9BD5"

    for row_index in range(2, ws.max_row + 1):

        ws.cell(
            row=row_index,
            column=2,
        ).number_format = '#,##0.00'

        ws.cell(
            row=row_index,
            column=3,
        ).number_format = '#,##0.00'

    # -----------------------------------------------------
    # Region Analysis sheet
    # -----------------------------------------------------

    ws = writer.sheets["Region Analysis"]

    ws.sheet_properties.tabColor = "ED7D31"

    for row_index in range(2, ws.max_row + 1):

        ws.cell(
            row=row_index,
            column=2,
        ).number_format = '#,##0.00'

    # -----------------------------------------------------
    # Order Status sheet
    # -----------------------------------------------------

    ws = writer.sheets["Order Status"]

    ws.sheet_properties.tabColor = "A5A5A5"

    for row_index in range(2, ws.max_row + 1):

        ws.cell(
            row=row_index,
            column=2,
        ).number_format = '#,##0'

    # -----------------------------------------------------
    # Prepared Data sheet
    # -----------------------------------------------------

    ws = writer.sheets["Prepared Data"]

    ws.sheet_properties.tabColor = "FFC000"

    header_positions = {
        cell.value: cell.column
        for cell in ws[1]
    }

    # Format numeric columns
    for column_name in ["UnitPrice", "Revenue"]:

        column_index = header_positions.get(column_name)

        if column_index is not None:

            for row_index in range(2, ws.max_row + 1):

                ws.cell(
                    row=row_index,
                    column=column_index,
                ).number_format = '#,##0.00'

    quantity_column = header_positions.get("Quantity")

    if quantity_column is not None:

        for row_index in range(2, ws.max_row + 1):

            ws.cell(
                row=row_index,
                column=quantity_column,
            ).number_format = '#,##0.00'

    # Highlight blank cells in Prepared Data
    for row in ws.iter_rows(
        min_row=2,
        max_row=ws.max_row,
    ):

        for cell in row:

            if cell.value is None:

                cell.fill = missing_fill

                cell.font = Font(
                    name="Aptos",
                    size=10,
                    bold=True,
                    color=warning_text,
                )

    # Set workbook metadata
    writer.book.properties.title = (
        "AI Business Analysis Report"
    )

    writer.book.properties.subject = (
        "Sales analysis and data quality report"
    )

    writer.book.properties.creator = (
        "AI-Powered Excel Business Analyst"
    )


# =========================================================
# 6. EXCEL REPORT EXPORT
# =========================================================

def create_excel_report(
    original_df,
    prepared_df,
    metrics,
    duplicate_count,
    invalid_numeric_count,
):

    output = BytesIO()

    summary_df = pd.DataFrame({
        "Metric": [
            "Total Completed Revenue",
            "Completed Orders",
            "Average Order Value",
        ],
        "Value": [
            metrics["total_revenue"],
            metrics["completed_order_count"],
            metrics["average_order_value"],
        ],
    })

    data_quality_df = pd.DataFrame({
        "Metric": [
            "Input Rows",
            "Input Columns",
            "Exact Duplicate Rows",
            "Rows Included in Analysis",
            "Rows with Invalid Numeric Values",
        ],
        "Value": [
            len(original_df),
            len(original_df.columns),
            duplicate_count,
            len(prepared_df),
            invalid_numeric_count,
        ],
    })

    with pd.ExcelWriter(
        output,
        engine="openpyxl",
    ) as writer:

        summary_df.to_excel(
            writer,
            sheet_name="Summary",
            index=False,
        )

        data_quality_df.to_excel(
            writer,
            sheet_name="Data Quality",
            index=False,
        )

        metrics["product_summary"].to_excel(
            writer,
            sheet_name="Product Analysis",
            index=False,
        )

        metrics["region_summary"].to_excel(
            writer,
            sheet_name="Region Analysis",
            index=False,
        )

        metrics["status_summary"].to_excel(
            writer,
            sheet_name="Order Status",
            index=False,
        )

        prepared_df.to_excel(
            writer,
            sheet_name="Prepared Data",
            index=False,
        )

        format_excel_workbook(writer)

    output.seek(0)

    return output.getvalue()


# =========================================================
# 7. STREAMLIT USER INTERFACE
# =========================================================

st.title(
    "📊 AI-Powered Excel Business Analyst"
)

st.write(
    "Upload your sales data to explore business metrics "
    "and ask questions using natural language."
)

st.divider()

uploaded_file = st.file_uploader(
    "Upload your CSV or Excel file",
    type=["csv", "xlsx"],
)


if uploaded_file is None:

    st.info(
        "Upload a CSV or XLSX file to start your analysis."
    )

    st.markdown(
        """
        **Available features**

        - Data Quality Analysis
        - Duplicate and Missing Value Detection
        - Sales KPIs
        - Product Revenue Analysis
        - Regional Revenue Analysis
        - Order Status Analysis
        - Natural Language Question Answering
        - Excel Report Export
        """
    )

else:

    try:

        # -------------------------------------------------
        # LOAD FILE
        # -------------------------------------------------

        original_df = load_data(
            uploaded_file
        )

        duplicate_count = int(
            original_df.duplicated().sum()
        )

        st.success(
            "File loaded successfully."
        )

        st.subheader(
            "1. Data Quality"
        )

        col1, col2, col3 = st.columns(3)

        col1.metric(
            "Input Rows",
            f"{len(original_df):,}",
        )

        col2.metric(
            "Columns",
            f"{len(original_df.columns):,}",
        )

        col3.metric(
            "Exact Duplicate Rows",
            f"{duplicate_count:,}",
        )

        exclude_duplicates = st.checkbox(
            "Exclude exact duplicate rows from analysis",
            value=True,
        )

        if exclude_duplicates:

            analysis_df = (
                original_df
                .drop_duplicates()
                .copy()
            )

        else:

            analysis_df = original_df.copy()

        # -------------------------------------------------
        # PREPARE DATA
        # -------------------------------------------------

        prepared_df = prepare_sales_data(
            analysis_df
        )

        st.caption(
            f"Rows included in analysis: "
            f"{len(prepared_df):,}"
        )

        # -------------------------------------------------
        # MISSING VALUE REPORT
        # -------------------------------------------------

        missing_report = pd.DataFrame({
            "Column": prepared_df.columns.astype(str),
            "Missing Values": (
                prepared_df.isna().sum().values
            ),
            "Missing Percentage": (
                prepared_df.isna().mean().values * 100
            ).round(2),
        })

        st.dataframe(
            missing_report,
            use_container_width=True,
            hide_index=True,
        )

        # -------------------------------------------------
        # INVALID NUMERIC VALUES
        # -------------------------------------------------

        invalid_numeric_count = int(
            (
                prepared_df["Quantity"].isna()
                | prepared_df["UnitPrice"].isna()
                | prepared_df["Quantity"].lt(0).fillna(False)
                | prepared_df["UnitPrice"].lt(0).fillna(False)
            ).sum()
        )

        if invalid_numeric_count:

            st.warning(
                f"{invalid_numeric_count} rows have missing "
                "or invalid Quantity/UnitPrice values. "
                "Their revenue is excluded."
            )

        # -------------------------------------------------
        # CALCULATE METRICS
        # -------------------------------------------------

        metrics = calculate_sales_metrics(
            prepared_df
        )

        st.divider()

        st.subheader(
            "2. Sales KPIs"
        )

        kpi1, kpi2, kpi3 = st.columns(3)

        kpi1.metric(
            "Completed Revenue",
            f"{metrics['total_revenue']:,.2f}",
        )

        kpi2.metric(
            "Completed Orders",
            f"{metrics['completed_order_count']:,}",
        )

        kpi3.metric(
            "Average Order Value",
            f"{metrics['average_order_value']:,.2f}",
        )

        st.caption(
            "Revenue includes Completed records with valid "
            "non-negative quantities and prices. "
            "Currency is not inferred."
        )

        # -------------------------------------------------
        # PRODUCT ANALYSIS
        # -------------------------------------------------

        st.divider()

        st.subheader(
            "3. Product Revenue Analysis"
        )

        product_summary = metrics[
            "product_summary"
        ]

        st.dataframe(
            product_summary,
            use_container_width=True,
            hide_index=True,
        )

        if not product_summary.empty:

            st.bar_chart(
                product_summary.set_index(
                    "Product"
                )["Revenue"]
            )

        # -------------------------------------------------
        # REGION ANALYSIS
        # -------------------------------------------------

        st.divider()

        st.subheader(
            "4. Revenue by Region"
        )

        region_summary = metrics[
            "region_summary"
        ]

        st.dataframe(
            region_summary,
            use_container_width=True,
            hide_index=True,
        )

        if not region_summary.empty:

            st.bar_chart(
                region_summary.set_index(
                    "Region"
                )["Revenue"]
            )

        # -------------------------------------------------
        # ORDER STATUS
        # -------------------------------------------------

        st.divider()

        st.subheader(
            "5. Order Status"
        )

        st.dataframe(
            metrics["status_summary"],
            use_container_width=True,
            hide_index=True,
        )

        st.caption(
            "Order Rows counts data rows. "
            "Question answers count unique OrderIDs."
        )

        # -------------------------------------------------
        # DATA PREVIEW
        # -------------------------------------------------

        st.divider()

        st.subheader(
            "6. Prepared Data Preview"
        )

        st.dataframe(
            prepared_df.head(20),
            use_container_width=True,
            hide_index=True,
        )

        st.info(
            "All numerical calculations are performed using Pandas. "
            "The uploaded file is not modified on disk."
        )

        # -------------------------------------------------
        # EXCEL REPORT DOWNLOAD
        # -------------------------------------------------

        st.divider()

        st.header(
            "📥 Export Excel Report"
        )

        excel_report = create_excel_report(
            original_df=original_df,
            prepared_df=prepared_df,
            metrics=metrics,
            duplicate_count=duplicate_count,
            invalid_numeric_count=invalid_numeric_count,
        )

        st.write(
            "Download a complete Excel report containing "
            "your KPIs, data quality results, product analysis, "
            "regional analysis, order status, and prepared data."
        )

        st.download_button(
            label="Download Excel Report",
            data=excel_report,
            file_name="AI_Business_Analysis_Report.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
        )

        # -------------------------------------------------
        # AI BUSINESS ANALYST
        # -------------------------------------------------

        st.divider()

        st.header(
            "🤖 Ask the AI Business Analyst"
        )

        with st.form(
            "business_question_form"
        ):

            question = st.text_input(
                "Ask a question about your sales data",
                placeholder=(
                    "Which product generated the highest revenue?"
                ),
            )

            submitted = st.form_submit_button(
                "Analyze Question"
            )

        if submitted:

            if not question.strip():

                st.warning(
                    "Please enter a question."
                )

            else:

                try:

                    with st.spinner(
                        "Analyzing your question..."
                    ):

                        answer = answer_business_question(
                            question,
                            metrics,
                            prepared_df,
                        )

                    st.subheader(
                        "Analysis Result"
                    )

                    st.write(answer)

                    st.caption(
                        "Numerical results are calculated by Pandas. "
                        "Ollama is used only for unsupported questions."
                    )

                except Exception as error:

                    st.error(
                        "Could not complete the analysis."
                    )

                    st.exception(error)

    except ValueError as error:

        st.error(
            f"Data validation error: {error}"
        )

    except Exception as error:

        st.error(
            "An unexpected error occurred."
        )

        st.exception(error)