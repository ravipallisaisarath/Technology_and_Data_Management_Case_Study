
import streamlit as st
import pandas as pd
from io import BytesIO
from pathlib import Path

st.set_page_config(
    page_title="Technology & Data Management",
    page_icon="📊",
    layout="wide"
)

st.title("Technology & Data Management Case Study")
st.write("Inventory aging, SKU classification, and customer recommendations.")

WORKBOOK = "Case_Data_Technology_and_Data_Management_All_Sheets.xlsx"

uploaded = st.sidebar.file_uploader(
    "Upload the Excel workbook",
    type=["xlsx"]
)


@st.cache_data
def load_workbook(file_bytes):
    excel = pd.ExcelFile(BytesIO(file_bytes))

    required_sheets = [
        "Item_Master",
        "Stock_Aging",
        "Sales_Lines",
        "Customers"
    ]

    missing = [
        sheet for sheet in required_sheets
        if sheet not in excel.sheet_names
    ]

    if missing:
        raise ValueError(
            "Missing sheets: " + ", ".join(missing)
        )

    return {
        sheet: pd.read_excel(excel, sheet_name=sheet)
        for sheet in required_sheets
    }


if uploaded is not None:
    try:
        data = load_workbook(uploaded.getvalue())
    except Exception as error:
        st.error(f"Could not load Excel workbook: {error}")
        st.stop()
elif Path(WORKBOOK).exists():
    try:
        data = load_workbook(Path(WORKBOOK).read_bytes())
    except Exception as error:
        st.error(f"Could not load Excel workbook: {error}")
        st.stop()
else:
    st.info(
        "Upload your Excel workbook using the sidebar. "
        "Alternatively, add the workbook to the same "
        "GitHub repository as app.py."
    )
    st.stop()


items = data["Item_Master"].copy()
stock = data["Stock_Aging"].copy()
sales = data["Sales_Lines"].copy()
customers = data["Customers"].copy()

for df in [items, stock, sales, customers]:
    df.columns = df.columns.astype(str).str.strip()

required_columns = {
    "Item_Master": [
        "SKU", "Item Name", "Category", "Master Pack Qty",
        "Unit Cost (Rs)", "Alternate Item (name)",
        "Alternate Flag", "Status"
    ],
    "Stock_Aging": [
        "SKU", "Qty 0-30 days", "Qty 31-60 days",
        "Qty 61-90 days", "Qty 91-180 days", "Qty 180+ days"
    ],
    "Sales_Lines": [
        "Month", "CustomerID", "CustomerName", "Industry",
        "SKU", "Qty", "Value (Rs)"
    ],
    "Customers": [
        "CustomerID", "CustomerName", "Industry", "Area"
    ]
}

frames = {
    "Item_Master": items,
    "Stock_Aging": stock,
    "Sales_Lines": sales,
    "Customers": customers
}

for sheet, columns in required_columns.items():
    missing = [
        col for col in columns
        if col not in frames[sheet].columns
    ]
    if missing:
        st.error(
            f"{sheet} is missing columns: {', '.join(missing)}"
        )
        st.stop()


# Convert numeric columns safely.
numeric_columns = [
    (items, ["Master Pack Qty", "Unit Cost (Rs)"]),
    (stock, [
        "Qty 0-30 days", "Qty 31-60 days", "Qty 61-90 days",
        "Qty 91-180 days", "Qty 180+ days"
    ]),
    (sales, ["Qty", "Value (Rs)"])
]

for df, columns in numeric_columns:
    for col in columns:
        df[col] = pd.to_numeric(
            df[col], errors="coerce"
        ).fillna(0)

for df in [items, stock, sales]:
    df["SKU"] = df["SKU"].astype(str).str.strip()

items["Status"] = items["Status"].fillna("").astype(str).str.strip()
items["Alternate Item (name)"] = (
    items["Alternate Item (name)"].fillna("").astype(str).str.strip()
)

# Map old SKUs to replacement SKUs using the alternate item name.
name_to_sku = {
    str(row["Item Name"]).strip().casefold(): row["SKU"]
    for _, row in items.iterrows()
    if pd.notna(row["Item Name"])
}

switch_map = {}

for _, row in items.iterrows():
    old_sku = row["SKU"]
    alternate_name = row["Alternate Item (name)"]
    flag = str(row["Alternate Flag"]).strip().casefold()

    if (
        alternate_name
        and alternate_name.casefold() in name_to_sku
        and flag not in ["", "no", "n", "false", "0", "nan"]
    ):
        switch_map[old_sku] = name_to_sku[
            alternate_name.casefold()
        ]


def canonical_sku(sku):
    seen = set()

    while sku in switch_map and sku not in seen:
        seen.add(sku)
        sku = switch_map[sku]

    return sku


items["Canonical SKU"] = items["SKU"].map(canonical_sku)
stock["Canonical SKU"] = stock["SKU"].map(canonical_sku)
sales["Canonical SKU"] = sales["SKU"].map(canonical_sku)

# R3: Calculate total stock and aged stock.
age_columns = [
    "Qty 0-30 days", "Qty 31-60 days", "Qty 61-90 days",
    "Qty 91-180 days", "Qty 180+ days"
]

stock["Total Stock"] = stock[age_columns].sum(axis=1)
stock["Aged Qty"] = stock[
    ["Qty 91-180 days", "Qty 180+ days"]
].sum(axis=1)

stock_summary = stock.groupby(
    "Canonical SKU", as_index=False
).agg({
    "Total Stock": "sum",
    "Qty 91-180 days": "sum",
    "Qty 180+ days": "sum",
    "Aged Qty": "sum"
})

# Use the latest six months of sales in the workbook.
sales["Month"] = pd.to_datetime(
    sales["Month"], errors="coerce"
)

sales_6m = sales.copy()

if sales["Month"].notna().any():
    latest_month = sales["Month"].max().to_period("M")
    first_month = latest_month - 5

    sales_6m = sales[
        sales["Month"].dt.to_period("M").between(
            first_month, latest_month
        )
    ].copy()

sales_summary = sales_6m.groupby(
    "Canonical SKU", as_index=False
).agg({
    "Qty": "sum",
    "Value (Rs)": "sum"
}).rename(columns={
    "Qty": "Six-Month Sold Qty",
    "Value (Rs)": "Six-Month Sales Value (Rs)"
})

# R2: Exclude inactive products from the main analysis.
inactive_statuses = ["inactive", "in-active", "disabled"]

active_items = items[
    ~items["Status"].str.casefold().isin(inactive_statuses)
].copy()

# Use the replacement SKU's master details for switched products.
master = active_items.drop_duplicates(
    "Canonical SKU", keep="last"
)

analysis = master.merge(
    stock_summary, on="Canonical SKU", how="left"
).merge(
    sales_summary, on="Canonical SKU", how="left"
)

for col in [
    "Total Stock", "Qty 91-180 days", "Qty 180+ days",
    "Aged Qty", "Six-Month Sold Qty",
    "Six-Month Sales Value (Rs)"
]:
    analysis[col] = analysis[col].fillna(0)

# R3 and R4.
analysis["Aged Value (Rs)"] = (
    analysis["Aged Qty"] * analysis["Unit Cost (Rs)"]
)

analysis["Average Monthly Sales Qty"] = (
    analysis["Six-Month Sold Qty"] / 6
)


def months_to_clear(row):
    average = row["Average Monthly Sales Qty"]

    if average <= 0:
        return float("inf") if row["Total Stock"] > 0 else 0

    return row["Total Stock"] / average


analysis["Months to Clear"] = analysis.apply(
    months_to_clear, axis=1
)

# R5: Classification rules.
def classify(row):
    if row["Six-Month Sold Qty"] < row["Master Pack Qty"]:
        return "Red List"

    if (
        row["Aged Qty"] <= 0
        or row["Months to Clear"] < 3
    ):
        return "No Action"

    return "Push"


analysis["Classification"] = analysis.apply(
    classify, axis=1
)

# Flag inactive items that have stock.
inactive_items = items[
    items["Status"].str.casefold().isin(inactive_statuses)
]

inactive_stock = stock.merge(
    inactive_items[["SKU", "Item Name"]],
    on="SKU",
    how="inner"
)

inactive_stock = inactive_stock[
    inactive_stock["Total Stock"] > 0
]

# Dashboard summary.
counts = analysis["Classification"].value_counts()

col1, col2, col3, col4 = st.columns(4)

col1.metric("SKUs analysed", len(analysis))
col2.metric("Red List", int(counts.get("Red List", 0)))
col3.metric("Push", int(counts.get("Push", 0)))
col4.metric("No Action", int(counts.get("No Action", 0)))

overview, sku_tab, quality_tab, download_tab = st.tabs([
    "Overview",
    "SKU Analysis",
    "Data Quality",
    "Download"
])

with overview:
    st.subheader("Classification summary")

    st.bar_chart(
        counts.rename_axis("Classification").to_frame("SKU count")
    )

    push_aged_value = analysis.loc[
        analysis["Classification"] == "Push",
        "Aged Value (Rs)"
    ].sum()

    st.metric(
        "Total aged value for Push",
        f"₹{push_aged_value:,.2f}"
    )

    st.subheader("Top 10 Push SKUs by aged value")

    top10 = analysis[
        analysis["Classification"] == "Push"
    ].sort_values(
        "Aged Value (Rs)", ascending=False
    ).head(10)

    st.dataframe(top10, use_container_width=True, hide_index=True)


with sku_tab:
    st.subheader("SKU analysis")

    selected_classes = st.multiselect(
        "Filter by classification",
        ["Red List", "Push", "No Action"],
        default=["Red List", "Push", "No Action"]
    )

    filtered = analysis[
        analysis["Classification"].isin(selected_classes)
    ]

    st.dataframe(
        filtered, use_container_width=True, hide_index=True
    )

    st.subheader("Requested SKU classifications")

    for sku in ["SKU1001", "SKU1006", "SKU1019"]:
        match = analysis[
            analysis["Canonical SKU"].astype(str) == sku
        ]

        if match.empty:
            st.write(f"{sku}: not found as a canonical SKU")
        else:
            st.write(
                f"{sku}: {match.iloc[0]['Classification']}"
            )


with quality_tab:
    st.subheader("Inactive items with stock")

    if inactive_stock.empty:
        st.success("No inactive items with stock were found.")
    else:
        st.dataframe(
            inactive_stock, use_container_width=True, hide_index=True
        )

    st.subheader("Basic data-quality checks")

    quality = pd.DataFrame([
        {
            "Sheet": name,
            "Rows": len(df),
            "Duplicate rows": int(df.duplicated().sum()),
            "Missing cells": int(df.isna().sum().sum())
        }
        for name, df in [
            ("Item_Master", items),
            ("Stock_Aging", stock),
            ("Sales_Lines", sales),
            ("Customers", customers)
        ]
    ])

    st.dataframe(quality, use_container_width=True, hide_index=True)


with download_tab:
    st.subheader("Download the analysis workbook")

    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        analysis.to_excel(
            writer, sheet_name="SKU_Analysis", index=False
        )
        inactive_stock.to_excel(
            writer, sheet_name="Inactive_Stock", index=False
        )
        quality.to_excel(
            writer, sheet_name="Data_Quality", index=False
        )

    st.download_button(
        label="Download Excel analysis",
        data=output.getvalue(),
        file_name="Technology_Data_Management_Analysis_Output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.sidebar.success("Workbook loaded successfully.")
st.sidebar.caption("Sales period: latest six months found in the workbook.")
