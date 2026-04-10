import pandas as pd
import streamlit as st
import plotly.express as px
from datetime import datetime, timedelta

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
# FUNCTIONS
# =========================
def calculate_mom(current, previous):
    if previous == 0:
        return 0
    return ((current - previous) / previous) * 100

def get_previous_month(year, month):
    return (year - 1, 12) if month == 1 else (year, month - 1)

def series_to_df(series, label):
    if not series:
        return pd.DataFrame(columns=["Period", label])
    return pd.DataFrame({
        "Period": [i["d"] for i in series],
        label: [i["v"] for i in series],
    })

# =========================
# CONFIG
# =========================
st.set_page_config(page_title="J2W Dashboard", layout="wide")

# =========================
# CSS (PILLS + CARDS)
# =========================
st.markdown("""
<style>
.kpi-card {
    background:#ffffff;
    border-radius:12px;
    padding:18px;
    height:140px;
    box-shadow:0 2px 8px rgba(0,0,0,0.05);
}
.kpi-title {color:#6c757d;font-size:14px;}
.kpi-value {font-size:30px;font-weight:700;color:#2c3e50;}

.red{border-top:4px solid #e74c3c;}
.blue{border-top:4px solid #3498db;}
.green{border-top:4px solid #2ecc71;}
.orange{border-top:4px solid #f39c12;}

div[data-testid="column"] {padding:0 6px;}

/* Pill buttons */
div[role="radiogroup"] > label {
    background-color: #f1f3f5;
    padding: 6px 14px;
    border-radius: 20px;
    margin-right: 6px;
    border: 1px solid #e0e0e0;
    cursor: pointer;
}
div[role="radiogroup"] > label:has(input:checked) {
    background-color: #e74c3c !important;
    color: white !important;
}
</style>
""", unsafe_allow_html=True)

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
# SIDEBAR
# =========================
st.title("J2W Dashboard")

with st.sidebar:
    years = st.multiselect("Year", year_options, default=[year_options[0]])
    months = st.multiselect("Month", month_options, default=[month_options[-1]])
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
# KPI CARDS WITH ICONS
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
# MOM STRIP
# =========================
cy, cm = int(years[0]), int(months[0])
py, pm = get_previous_month(cy, cm)
prev = get_data([py],[pm],clients)

def mom(label,cur,pr):
    ch = calculate_mom(cur,pr)
    arrow = "▲" if ch>=0 else "▼"
    color = "#27ae60" if ch>=0 else "#e74c3c"
    return f"<span><b>{label}</b> {cur:,} <span style='color:{color}'>{arrow} {abs(ch):.0f}%</span></span>"

st.markdown(f"""
<div style="background:#f8f9fb;color:#2c3e50;padding:12px;border-radius:10px;margin-top:10px;display:flex;gap:20px;">
<b>VS {pm}</b>
{mom("Dem",grand.get("dem",0),prev.get("dem",0))}
{mom("Sub",grand.get("sub",0),prev.get("sub",0))}
</div>
""", unsafe_allow_html=True)

# =========================
# DAY TRENDS
# =========================
st.markdown("### Day-on-Day Trends")

range_day = st.radio("",["Last 7 Days","Last 15 Days","Last 30 Days"], horizontal=True)

col1,col2 = st.columns(2)
from_day = col1.date_input("From Date")
to_day = col2.date_input("To Date")

metric_map = {"Demands":"dem","Submissions":"sub","Interviews":"intv","Selections":"sel"}
metric_day = st.radio("", list(metric_map.keys()), horizontal=True)

today = datetime.now().date()
start_day = today - timedelta(days=7 if range_day=="Last 7 Days" else 15 if range_day=="Last 15 Days" else 30)
end_day = to_day if to_day else today

day_data = daily_trends_cached(None,None,None,"day")

filtered_day = [
    i for i in day_data.get(metric_map[metric_day],[])
    if start_day <= pd.to_datetime(i["d"]).date() <= end_day
]

df_day = series_to_df(filtered_day, metric_day)

fig = px.line(df_day, x="Period", y=metric_day, markers=True)
fig.update_traces(text=df_day[metric_day], textposition="top center")

st.plotly_chart(fig, use_container_width=True)

# =========================
# MONTH TRENDS (NEW SECTION)
# =========================
st.markdown("### Month-on-Month Trends")

range_month = st.radio("",["Last 6 Months","Last 12 Months"], horizontal=True)

metric_month = st.radio("", list(metric_map.keys()), horizontal=True, key="month")

month_data = daily_trends_cached(None,None,None,"month")

month_filtered = month_data.get(metric_map[metric_month], [])[-6:] if range_month=="Last 6 Months" else month_data.get(metric_map[metric_month], [])[-12:]

df_month = series_to_df(month_filtered, metric_month)

fig2 = px.line(df_month, x="Period", y=metric_month, markers=True)
fig2.update_traces(text=df_month[metric_month], textposition="top center")

st.plotly_chart(fig2, use_container_width=True)