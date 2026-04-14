# =========================================
# J2W CEO DASHBOARD (Single File Version)
# =========================================

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import datetime, timedelta
import calendar

st.set_page_config(page_title="CEO Dashboard", layout="wide")

FILE_PATH = "your_data.xlsx"

# =========================================
# LOAD DATA
# =========================================
@st.cache_data
def load_data():
    df = pd.read_excel(FILE_PATH)
    df.columns = df.columns.str.strip()

    df["date"] = pd.to_datetime(df["display_date"], dayfirst=True, errors="coerce")

    df["p_o_value"] = pd.to_numeric(df["p_o_value"], errors="coerce").fillna(0)
    df["margin"] = pd.to_numeric(df["margin"], errors="coerce").fillna(0)

    return df

df = load_data()

# =========================================
# FILTERS
# =========================================
st.sidebar.title("🔎 Filters")

years = sorted(df["date"].dt.year.dropna().unique())
months = list(range(1, 13))

selected_year = st.sidebar.selectbox("Year", years, index=len(years)-1)
selected_month = st.sidebar.selectbox("Month", months, index=datetime.now().month-1)

clients = st.sidebar.multiselect("Client", df["company_name"].dropna().unique())
designations = st.sidebar.multiselect("Designation", df["designation"].dropna().unique())

if st.sidebar.button("🔄 Reset Filters"):
    st.session_state.clear()
    st.rerun()

# =========================================
# FILTER FUNCTION (SINGLE SOURCE)
# =========================================
def apply_filters(data, year, month):
    data = data[
        (data["date"].dt.year == year) &
        (data["date"].dt.month == month)
    ]

    if clients:
        data = data[data["company_name"].isin(clients)]

    if designations:
        data = data[data["designation"].isin(designations)]

    return data

data = apply_filters(df.copy(), selected_year, selected_month)

# =========================================
#