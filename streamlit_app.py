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
# CSS
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
# TOP FILTER BAR
# =========================
col1, col2, col3, col4, col5, col6, col7 = st.columns([1,1,1,1,1,0.5,0.5])

month_map = {i: calendar.month_abbr[i] for i in range(1,13)}

with col1:
    years = st.multiselect("Year", year_options, default=[year_options[0]])

with col2:
    month_names = [month_map[m] for m in month_options]
    selected_month_names = st.multiselect("Month", month_names, default=[month_map[month_options[-1]]])
    months = [k for k,v in month_map.items() if v in selected_month_names]

with col3:
    clients = st.multiselect("Clients", [c["name"] for c in client_meta])

with col4:
    domains = st.multiselect("Domain", ["Tech","Non-Tech","Infra"])

with col5:
    bhs = st.multiselect("BH", ["BH1","BH2"])

with col6:
    if st.button("Reset"):
        st.session_state.clear()
        st.rerun()

with col7:
    if st.button("🔄"):
        st.rerun()

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
# KPI CARDS (FULL)
# =========================
def card(title, value, color, icon, sub=""):
    return f"""
    <div class='kpi-card {color}'>
        <div style='display:flex;justify-content:space-between'>
            <div class='kpi-title'>{title}</div>
            <div>{icon}</div>
        </div>
        <div class='kpi-value'>{value}</div>
        <div style='font-size:12px'>{sub}</div>
    </div>
    """

row1 = st.columns(6)
data1 = [
    ("Demands", grand.get("dem",0),"red","📋",""),
    ("Submissions", grand.get("sub",0),"blue","📤",""),
    ("Interviews", grand.get("l1",0),"orange","🎤",""),
    ("Selections", grand.get("sel",0),"green","✅",""),
    ("Selection Pipeline", grand.get("sp_hc",0),"blue","📊",""),
    ("Active HC", grand.get("active_hc",0),"blue","👥",""),
]

for col,(t,v,c,i,s) in zip(row1,data1):
    col.markdown(card(t,f"{v:,}",c,i,s), unsafe_allow_html=True)

row2 = st.columns(5)
data2 = [
    ("Onboarded", grand.get("ob_hc",0),"green","🚀",""),
    ("Exits", grand.get("ex_hc",0),"red","🚪",""),
    ("Net HC", grand.get("net_hc",0),"green","📈",""),
    ("Net PO", f"₹{grand.get('net_po',0):.2f}L","green","💰",""),
    ("Net Margin", f"₹{grand.get('net_mg',0):.2f}L","green","📊",""),
]

for col,(t,v,c,i,s) in zip(row2,data2):
    col.markdown(card(t,v,c,i,s), unsafe_allow_html=True)

# =========================
# COMPARISONS
# =========================
cy, cm = int(years[0]), int(months[0])
py, pm = get_previous_month(cy, cm)
prev = get_data([py],[pm],clients)

def comp(label,cur,pr):
    ch = calculate_mom(cur,pr)
    arrow = "▲" if ch>=0 else "▼"
    color = "#27ae60" if ch>=0 else "#e74c3c"
    return f"<span><b>{label}</b> {cur:,} <span style='color:{color}'>{arrow} {abs(ch):.0f}%</span></span>"

# LMTD
st.markdown(f"""
<div style="background:#ffffff;padding:12px;border-radius:10px;margin-top:10px;display:flex;gap:20px;">
<b>VS LMTD</b>
{comp("Dem",grand.get("dem",0),prev.get("dem",0))}
{comp("Sub",grand.get("sub",0),prev.get("sub",0))}
</div>
""", unsafe_allow_html=True)

# MOM
st.markdown(f"""
<div style="background:#f8f9fb;padding:12px;border-radius:10px;margin-top:10px;display:flex;gap:20px;">
<b>VS {pm}</b>
{comp("Dem",grand.get("dem",0),prev.get("dem",0))}
{comp("Sub",grand.get("sub",0),prev.get("sub",0))}
</div>
""", unsafe_allow_html=True)

# =========================
# DAY CHART
# =========================
st.markdown("### Day-on-Day Trends")

metric_map = {"Demands":"dem","Submissions":"sub","Interviews":"intv","Selections":"sel"}
metric = st.radio("", list(metric_map.keys()), horizontal=True)

data_day = daily_trends_cached(None,None,None,"day")
df = series_to_df(data_day.get(metric_map[metric],[]), metric)

fig = px.line(df, x="Period", y=metric, markers=True)
fig.update_traces(text=df[metric], textposition="top center")
st.plotly_chart(fig, use_container_width=True)

# =========================
# MONTH CHART
# =========================
st.markdown("### Month-on-Month Trends")

metric2 = st.radio("", list(metric_map.keys()), horizontal=True, key="m2")
data_month = daily_trends_cached(None,None,None,"month")

df2 = series_to_df(data_month.get(metric_map[metric2],[]), metric2)

fig2 = px.line(df2, x="Period", y=metric2, markers=True)
fig2.update_traces(text=df2[metric2], textposition="top center")
st.plotly_chart(fig2, use_container_width=True)