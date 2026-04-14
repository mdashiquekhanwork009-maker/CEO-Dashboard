import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import datetime

st.set_page_config(page_title="Recruitment Dashboard", layout="wide")

# =========================
# LOAD DATA (UPLOAD OR FILE)
# =========================
@st.cache_data
def load_data():
    df = pd.read_excel("your_data.xlsx")  # 🔁 replace with your file
    df.columns = df.columns.str.strip()
    df["display_date"] = pd.to_datetime(df["display_date"], dayfirst=True, errors="coerce")
    return df

df = load_data()

# =========================
# SIDEBAR FILTERS
# =========================
st.sidebar.title("Filters")

years = sorted(df["display_date"].dt.year.dropna().unique())
months = sorted(df["display_date"].dt.month.dropna().unique())

selected_year = st.sidebar.selectbox("Year", years)
selected_month = st.sidebar.selectbox("Month", months)

clients = st.sidebar.multiselect("Client", df["company_name"].unique())
status = st.sidebar.multiselect("Status", df["Status"].unique())

# RESET BUTTON
if st.sidebar.button("Reset"):
    st.session_state.clear()
    st.rerun()

# =========================
# FILTER DATA
# =========================
filtered = df.copy()

filtered = filtered[
    (filtered["display_date"].dt.year == selected_year) &
    (filtered["display_date"].dt.month == selected_month)
]

if clients:
    filtered = filtered[filtered["company_name"].isin(clients)]

if status:
    filtered = filtered[filtered["Status"].isin(status)]

# =========================
# KPI CALCULATIONS
# =========================
dem = len(filtered)
sub = len(filtered[filtered["Status"] == "Submission"])
l1 = len(filtered[filtered["Status"] == "L1"])
sel = len(filtered[filtered["Status"] == "Selected"])
ob = len(filtered[filtered["Status"] == "Onboarded"])
ex = len(filtered[filtered["Status"] == "Exit"])

active = len(filtered[filtered["Status"] == "Active"])

po = filtered["p_o_value"].fillna(0).sum()
margin = filtered["margin"].fillna(0).sum()

# =========================
# KPI UI
# =========================
col1, col2, col3, col4 = st.columns(4)

col1.metric("Demands", dem)
col2.metric("Submissions", sub)
col3.metric("Selections", sel)
col4.metric("Onboarded", ob)

col5, col6, col7 = st.columns(3)

col5.metric("Active HC", active)
col6.metric("PO Value", f"₹{round(po,2)}")
col7.metric("Margin", f"₹{round(margin,2)}")

# =========================
# TREND DATA
# =========================
trend = filtered.groupby(filtered["display_date"].dt.date).size().reset_index(name="count")

# =========================
# CHART
# =========================
fig = go.Figure()

fig.add_trace(go.Scatter(
    x=trend["display_date"],
    y=trend["count"],
    mode="lines+markers",
))

fig.update_layout(
    title="Daily Trend",
    xaxis_title="Date",
    yaxis_title="Count"
)

st.plotly_chart(fig, width="stretch")

# =========================
# TABLE
# =========================
st.dataframe(filtered, width="stretch")