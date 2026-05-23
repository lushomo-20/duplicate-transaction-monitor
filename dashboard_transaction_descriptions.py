from __future__ import annotations

from pathlib import Path

import pandas as pd
import streamlit as st


DEFAULT_REPORT = Path(r"C:\duplicate checker\transaction_description_duplicate_report_june_2025_to_date.xlsx")


st.set_page_config(page_title="Transaction Duplicate Review", layout="wide")
st.title("Transaction Duplicate Review")


@st.cache_data(show_spinner=False)
def read_sheet(path: str, sheet_name: str) -> pd.DataFrame:
    try:
        return pd.read_excel(path, sheet_name=sheet_name)
    except ValueError:
        return pd.DataFrame()


def fmt(value: float | int) -> str:
    return f"{value:,.0f}"


report_path = st.sidebar.text_input("Excel report", str(DEFAULT_REPORT))
path = Path(report_path)
if not path.exists():
    st.error(f"Report not found: {path}")
    st.stop()

summary = read_sheet(str(path), "Summary")
categories = read_sheet(str(path), "Summary by description")
actual = read_sheet(str(path), "Actual duplicates")
possible = read_sheet(str(path), "Possible duplicates")
by_id = read_sheet(str(path), "Duplicated IDs by sheet")
accounts = read_sheet(str(path), "Top duplicated accounts")
devices = read_sheet(str(path), "Top duplicate devices")
amounts = read_sheet(str(path), "Top duplicate amounts")
trend = read_sheet(str(path), "Duplicate trend by day")

if actual.empty:
    st.warning("No actual duplicates were found for the selected transaction descriptions.")
    st.stop()

actual["transaction_date"] = pd.to_datetime(actual["transaction_date"], errors="coerce")
min_date = actual["transaction_date"].min().date()
max_date = actual["transaction_date"].max().date()

st.sidebar.header("Filters")
date_range = st.sidebar.date_input("Date range", (min_date, max_date), min_value=min_date, max_value=max_date)
category_filter = st.sidebar.multiselect(
    "Description type",
    sorted(actual["description_category"].dropna().unique().tolist()),
    default=sorted(actual["description_category"].dropna().unique().tolist()),
)

filtered = actual.copy()
if isinstance(date_range, tuple) and len(date_range) == 2:
    filtered = filtered[filtered["transaction_date"].between(pd.Timestamp(date_range[0]), pd.Timestamp(date_range[1]))]
if category_filter:
    filtered = filtered[filtered["description_category"].isin(category_filter)]

st.caption("This dashboard uses TransactionList only. It checks transactions whose description looks like FNB OB PMT, Smart CIT, or Cash deposit partner.")

col1, col2, col3, col4 = st.columns(4)
col1.metric("Duplicate rows", fmt(len(filtered)))
col2.metric("Duplicate Track/VODS IDs", fmt(filtered["track_vods_id"].nunique()))
col3.metric("Accounts affected", fmt(filtered["account_key"].nunique()))
col4.metric("Devices involved", fmt(filtered["device"].nunique()))

st.subheader("Plain summary by description type")
st.dataframe(categories, use_container_width=True, hide_index=True)

left, right = st.columns(2)
with left:
    st.subheader("Where the duplicates came from")
    st.dataframe(
        by_id[["track_vods_id", "duplicate_rows", "source_sheets", "description_categories", "accounts", "devices"]].head(100),
        use_container_width=True,
        hide_index=True,
    )
with right:
    st.subheader("Highest risk accounts")
    st.dataframe(accounts.head(25), use_container_width=True, hide_index=True)

left, right = st.columns(2)
with left:
    st.subheader("Top devices")
    st.caption("Device comes from TransactionList. A high count can point to a posting source, terminal, or process to investigate.")
    st.dataframe(devices.head(25), use_container_width=True, hide_index=True)
with right:
    st.subheader("Repeated amounts")
    st.caption("Repeated amounts help spot repeated posting patterns.")
    st.dataframe(amounts.head(25), use_container_width=True, hide_index=True)

st.subheader("Daily duplicate count")
if not trend.empty:
    pivot = trend.pivot_table(index="transaction_date", columns="description_category", values="duplicate_rows", aggfunc="sum", fill_value=0)
    st.line_chart(pivot)

st.subheader("Actual duplicate details")
detail_cols = [
    "source_sheet",
    "row_number",
    "description_category",
    "account_name",
    "track_vods_id",
    "transaction_date",
    "transaction_time",
    "amount",
    "description",
    "device",
    "ref_no",
    "duplicate_count",
]
st.dataframe(filtered[[col for col in detail_cols if col in filtered.columns]], use_container_width=True, hide_index=True)

st.subheader("Possible duplicates")
st.caption("These match by account, date, time, amount, and description, but not by the same Track/VODS ID.")
st.dataframe(possible, use_container_width=True, hide_index=True)
