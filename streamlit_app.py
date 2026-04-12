import pandas as pd
import streamlit as st
import plotly.express as px
from datetime import datetime, timedelta
import calendar

from dashboard import (
    compute_all_cached,
    daily_trends_cached,
    freeze_filter,
    get_client_catalog,
    get_periods_cached,
    grand_total,
    round_m,
)

# =========================
# CONFIG
# =========================
st.set_page_config(page_title="J2W Dashboard", layout="wide")

# =========================
# CSS (BLACK TEXT FIX)
# =========================
st.markdown("""
<style>
body, .stApp {
    color: #000000 !important;
}
.kpi-card {
    background:#ffffff;
    border-radius:12px;
    padding:18px;
    height:140px;
    box-shadow:0 2px 8px rgba(0,0,0,0.05);
}
.kpi-title {color:#555;font-size:14px;}
.kpi-value {font-size:30px;font-weight:700;color:#000;}

.red{border-top:4px solid #e74c3c;}
.blue{border-top:4px solid #3498db;}
.green{border-top:4px solid #2ecc71;}
.orange{border-top:4px solid #f39c12;}

div[data-testid="column"] {padding:0 6px;}

div[role="radiogroup"] > label {
    background-color: #f1f3f5;
    padding: 6px 14px;
    border-radius: 20px;
    margin-right: 6px;
}
div[role="radiogroup"] > label:has(input:checked) {
    background-color: #e74c3c !important;
    color: white !important;
}
</style>
""", unsafe_allow_html=True)

# =========================
# FUNCTIONS
# =========================
def calculate_mom(current, previous):
    if previous == 0:
        return 0
    return ((current - previous) / previous) * 100

def series_to_df(series, label):
    return pd.DataFrame({
        "Period": [i["d"] for i in series],
        label: [i["v"] for i in series],
    }) if series else pd.DataFrame()

# =========================
# LOAD DATA
# =========================
@st.cache_data
def init():
    periods = get_periods_cached()
    client_meta = list(get_client_catalog())
    years = sorted({str(p[0]) for p in periods}, reverse=True)
    months = sorted({int(p[1]) for p in periods})
    return client_meta, years, months

client_meta, year_options, month_options = init()

# =========================
# TOP FILTERS
# =========================
month_map = {i: calendar.month_abbr[i] for i in range(1,13)}

col1, col2, col3 = st.columns(3)

with col1:
    years = st.multiselect("Year", year_options, default=[year_options[0]])

with col2:
    month_names = [month_map[m] for m in month_options]
    selected_month_names = st.multiselect("Month", month_names, default=[month_map[month_options[-1]]])
    months = [k for k,v in month_map.items() if v in selected_month_names]

with col3:
    clients = st.multiselect("Clients", [c["name"] for c in client_meta])

# =========================
# DATA
# =========================
@st.cache_data
def get_data(y, m, c):
    res = compute_all_cached(
        freeze_filter(set(map(int, y))),
        freeze_filter(set(m)),
        freeze_filter(set(c)),
    )
    return round_m(grand_total(res))

grand = get_data(years, months, clients)

# =========================
# KPI
# =========================
def card(title, value, color, icon):
    return f"""
    <div class='kpi-card {color}'>
        <div style='display:flex;justify-content:space-between'>
            <div class='kpi-title'>{title}</div>
            <div>{icon}</div>
        </div>
        <div class='kpi-value'>{value}</div>
    </div>
    """

cols = st.columns(5)
data = [
    ("Demands", grand.get("dem",0),"red","📋"),
    ("Submissions", grand.get("sub",0),"blue","📤"),
    ("Interviews", grand.get("l1",0),"orange","🎤"),
    ("Selections", grand.get("sel",0),"green","✅"),
    ("Active HC", grand.get("active_hc",0),"blue","👥"),
]

for col,(t,v,c,i) in zip(cols,data):
    col.markdown(card(t,f"{v:,}",c,i), unsafe_allow_html=True)

# =========================
# DAY FILTERS (NEW)
# =========================
st.markdown("## Day-on-Day Trends")

col1, col2 = st.columns(2)
from_day = col1.date_input("From Date", key="d1")
to_day = col2.date_input("To Date", key="d2")

metric_map = {"Demands":"dem","Submissions":"sub","Interviews":"intv","Selections":"sel"}
metric = st.radio("", list(metric_map.keys()), horizontal=True)

day_data = daily_trends_cached(None,None,None,"day")

filtered_day = []
for i in day_data.get(metric_map[metric], []):
    d = pd.to_datetime(i["d"]).date()
    if (not from_day or d >= from_day) and (not to_day or d <= to_day):
        filtered_day.append(i)

df = series_to_df(filtered_day, metric)

fig = px.line(df, x="Period", y=metric, markers=True)
fig.update_traces(text=df[metric], textposition="top center")
st.plotly_chart(fig, use_container_width=True)

# =========================
# MONTH FILTERS (NEW)
# =========================
st.markdown("## Month-on-Month Trends")

col1, col2 = st.columns(2)
from_m = col1.date_input("From Month", key="m1")
to_m = col2.date_input("To Month", key="m2")

metric2 = st.radio("", list(metric_map.keys()), horizontal=True, key="m_radio")

month_data = daily_trends_cached(None,None,None,"month")

filtered_month = []
for i in month_data.get(metric_map[metric2], []):
    d = pd.to_datetime(i["d"]).date()
    if (not from_m or d >= from_m) and (not to_m or d <= to_m):
        filtered_month.append(i)

df2 = series_to_df(filtered_month, metric2)

fig2 = px.line(df2, x="Period", y=metric2, markers=True)
fig2.update_traces(text=df2[metric2], textposition="top center")
st.plotly_chart(fig2, use_container_width=True)

# =========================
# CEO ANALYSIS (🔥 NEW)
# =========================
st.markdown("## 📊 CEO Insights")

growth = calculate_mom(grand.get("dem",0), grand.get("sub",0))

st.markdown(f"""
- 📌 **Demand vs Supply Gap:** {'High' if grand.get("sub",0) < grand.get("dem",0) else 'Healthy'}
- 📈 **Hiring Efficiency:** {'Strong' if grand.get("sel",0) > 0 else 'Needs Attention'}
- 🚀 **Net HC Movement:** {grand.get("net_hc",0)}
- 💰 **Revenue Trend (PO):** ₹{grand.get("net_po",0):.2f}L
- ⚠️ **Risk Indicator:** {'High exits' if grand.get("ex_hc",0) > grand.get("ob_hc",0) else 'Stable'}
""")