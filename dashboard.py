from __future__ import annotations

from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st


DEFAULT_REPORT = Path(r"C:\duplicate checker\duplicate_report_june_2025_to_date.xlsx")


st.set_page_config(page_title="Duplicate Transaction Monitor", layout="wide")


@st.cache_data(show_spinner=False)
def read_sheet(path: str, sheet_name: str) -> pd.DataFrame:
    try:
        return pd.read_excel(path, sheet_name=sheet_name)
    except ValueError:
        return pd.DataFrame()


def number(value: float | int) -> str:
    return f"{value:,.0f}"


st.title("Duplicate Transaction Monitor")

report_path = st.sidebar.text_input("Report workbook", str(DEFAULT_REPORT))
path = Path(report_path)

if not path.exists():
    st.error(f"Report not found: {path}")
    st.stop()

summary = read_sheet(str(path), "Summary")
actual = read_sheet(str(path), "Actual duplicates")
possible = read_sheet(str(path), "Possible duplicates")
accounts = read_sheet(str(path), "Top duplicated accounts")
devices = read_sheet(str(path), "Top duplicate devices")
amounts = read_sheet(str(path), "Top duplicate amounts")
trend = read_sheet(str(path), "Duplicate trend by day")
agents = read_sheet(str(path), "SmartCIT agent ranking")
smartcit_possible = read_sheet(str(path), "SmartCIT possible duplicates")
matched_agents = read_sheet(str(path), "Matched SmartCIT agents")

if actual.empty:
    st.warning("No actual duplicates were found in the report.")
    st.stop()

actual["transaction_date"] = pd.to_datetime(actual["transaction_date"], errors="coerce")
if not trend.empty:
    trend["transaction_date"] = pd.to_datetime(trend["transaction_date"], errors="coerce")

min_date = actual["transaction_date"].min().date()
max_date = actual["transaction_date"].max().date()
date_range = st.sidebar.date_input("Date range", (min_date, max_date), min_value=min_date, max_value=max_date)

filtered = actual.copy()
if isinstance(date_range, tuple) and len(date_range) == 2:
    start, end = pd.Timestamp(date_range[0]), pd.Timestamp(date_range[1])
    filtered = filtered[filtered["transaction_date"].between(start, end)]

account_filter = st.sidebar.multiselect(
    "Accounts",
    sorted(filtered["account_key"].dropna().unique().tolist()),
    max_selections=25,
)
if account_filter:
    filtered = filtered[filtered["account_key"].isin(account_filter)]

device_filter = st.sidebar.multiselect(
    "Devices",
    sorted(filtered["device"].dropna().astype(str).unique().tolist()),
    max_selections=25,
)
if device_filter:
    filtered = filtered[filtered["device"].astype(str).isin(device_filter)]

total_duplicate_rows = len(filtered)
duplicate_trace_ids = filtered["vods_id"].nunique()
possible_rows = len(possible)
top_account_count = filtered["account_key"].nunique()

col1, col2, col3, col4 = st.columns(4)
col1.metric("Total duplicates", number(total_duplicate_rows), help="Overall issue size")
col2.metric("Duplicate Track/VODS IDs", number(duplicate_trace_ids), help="Critical fraud/system issue")
col3.metric("Possible duplicates", number(possible_rows), help="Same date/time/account/amount/description with different TRACE_ID")
col4.metric("Risk accounts", number(top_account_count), help="Accounts present in actual duplicates")

left, right = st.columns(2)

with left:
    st.subheader("Top duplicated accounts")
    account_rank = (
        filtered.groupby("account_key", dropna=False)
        .agg(duplicate_rows=("vods_id", "size"), duplicate_trace_ids=("vods_id", "nunique"), total_amount=("amount", "sum"))
        .reset_index()
        .sort_values("duplicate_rows", ascending=False)
        .head(15)
    )
    fig = px.bar(account_rank, x="duplicate_rows", y="account_key", orientation="h", hover_data=["duplicate_trace_ids", "total_amount"])
    fig.update_layout(yaxis={"categoryorder": "total ascending"}, height=460, margin=dict(l=10, r=10, t=10, b=10))
    st.plotly_chart(fig, use_container_width=True)

with right:
    st.subheader("Top duplicate devices")
    device_rank = (
        filtered.groupby("device", dropna=False)
        .agg(duplicate_rows=("vods_id", "size"), duplicate_trace_ids=("vods_id", "nunique"), total_amount=("amount", "sum"))
        .reset_index()
        .sort_values("duplicate_rows", ascending=False)
        .head(15)
    )
    fig = px.bar(device_rank, x="duplicate_rows", y="device", orientation="h", hover_data=["duplicate_trace_ids", "total_amount"])
    fig.update_layout(yaxis={"categoryorder": "total ascending"}, height=460, margin=dict(l=10, r=10, t=10, b=10))
    st.plotly_chart(fig, use_container_width=True)

left, right = st.columns(2)

with left:
    st.subheader("Top duplicate amounts")
    amount_rank = (
        filtered.groupby("amount", dropna=False)
        .agg(duplicate_rows=("vods_id", "size"), duplicate_trace_ids=("vods_id", "nunique"), accounts=("account_key", "nunique"))
        .reset_index()
        .sort_values("duplicate_rows", ascending=False)
        .head(20)
    )
    st.dataframe(amount_rank, use_container_width=True, hide_index=True)

with right:
    st.subheader("Duplicate trend by day")
    filtered_trend = (
        filtered.groupby("transaction_date", dropna=False)
        .agg(duplicate_rows=("vods_id", "size"), duplicate_trace_ids=("vods_id", "nunique"))
        .reset_index()
        .sort_values("transaction_date")
    )
    fig = px.line(filtered_trend, x="transaction_date", y="duplicate_rows", markers=True, hover_data=["duplicate_trace_ids"])
    fig.update_layout(height=390, margin=dict(l=10, r=10, t=10, b=10))
    st.plotly_chart(fig, use_container_width=True)

st.subheader("SmartCIT agent ranking")
if matched_agents.empty or matched_agents["agent_name"].dropna().empty:
    st.caption("No direct Track/VODS ID matches were found between actual transaction duplicates and the uploaded SmartCIT report. Ranking below uses SmartCIT possible duplicates: same client, same date, same amount.")
    if agents.empty:
        st.info("No SmartCIT possible duplicate agent ranking is available.")
    else:
        st.dataframe(agents, use_container_width=True, hide_index=True)
else:
    agent_source = matched_agents[matched_agents["vods_id"].isin(filtered["vods_id"].unique())]
    agent_rank = (
        agent_source.groupby(["agent_name", "agent_mobile"], dropna=False)
        .agg(
            duplicate_rows=("vods_id", "size"),
            duplicate_trace_ids=("vods_id", "nunique"),
            accounts=("account_key", "nunique"),
            devices=("device", "nunique"),
            total_amount=("amount", "sum"),
        )
        .reset_index()
        .sort_values(["duplicate_rows", "duplicate_trace_ids"], ascending=False)
    )
    st.dataframe(agent_rank, use_container_width=True, hide_index=True)

if not smartcit_possible.empty:
    st.subheader("SmartCIT possible duplicate details")
    st.dataframe(smartcit_possible, use_container_width=True, hide_index=True)

st.subheader("Duplicate details")
detail_columns = [
    "source_sheet",
    "row_number",
    "account_name",
    "vods_id",
    "transaction_date",
    "transaction_time",
    "amount",
    "description",
    "device",
    "ref_no",
    "duplicate_count",
]
st.dataframe(filtered[[col for col in detail_columns if col in filtered.columns]], use_container_width=True, hide_index=True)

st.download_button(
    "Download filtered duplicates as CSV",
    filtered.to_csv(index=False).encode("utf-8"),
    "filtered_duplicates.csv",
    "text/csv",
)
