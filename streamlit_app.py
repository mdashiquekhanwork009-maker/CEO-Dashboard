import pandas as pd
import streamlit as st
from datetime import datetime, timedelta

from dashboard import (
    RAW_DATASET_CONFIG,
    CAPTIVE_SUFFIX,
    compute_all_cached,
    daily_trends_cached,
    freeze_filter,
    get_client_catalog,
    get_mapping_context,
    get_mapped_client_name,
    get_periods_cached,
    get_raw_dataset_frame,
    grand_total,
    resolve_client_filter_cached,
    round_m,
)

# =========================
# BASIC FUNCTIONS
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
        "Period": [item["d"] for item in series],
        label: [item["v"] for item in series],
    })

# =========================
# PAGE CONFIG
# =========================
st.set_page_config(page_title="J2W Dashboard", layout="wide")

# =========================
# CSS
# =========================
st.markdown("""
<style>
.kpi-card {
    background: #ffffff;
    border-radius: 14px;
    padding: 18px;
    height: 150px;
    display: flex;
    flex-direction: column;
    justify-content: space-between;
    box-shadow: 0 2px 10px rgba(0,0,0,0.05);
    margin-bottom: 10px;
}
.kpi-title { font-size: 14px; color: #6c757d; font-weight: 600; }
.kpi-value { font-size: 32px; font-weight: 700; color: #2c3e50; }
.kpi-sub { font-size: 12px; color: #7f8c8d; }

.red { border-top: 4px solid #e74c3c; }
.green { border-top: 4px solid #2ecc71; }
.blue { border-top: 4px solid #3498db; }
.orange { border-top: 4px solid #f39c12; }

.badge {
    padding: 4px 10px;
    border-radius: 20px;
    font-size: 12px;
    background: #f8d7da;
    color: #721c24;
}

div[data-testid="column"] {
    padding: 0 6px;
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
# SIDEBAR FILTERS
# =========================
st.title("J2W Dashboard")

with st.sidebar:
    selected_years = st.multiselect("Year", year_options, default=[year_options[0]])
    selected_months = st.multiselect("Month", month_options, default=[month_options[-1]])
    selected_clients = st.multiselect("Clients", [c["name"] for c in client_meta])

# =========================
# DATA COMPUTE
# =========================
@st.cache_data
def get_data(years, months, clients):
    result = compute_all_cached(
        freeze_filter(set(map(int, years))),
        freeze_filter(set(months)),
        freeze_filter(set(clients)),
    )
    return round_m(grand_total(result))

grand = get_data(selected_years, selected_months, selected_clients)

# =========================
# KPI FUNCTION
# =========================
def kpi_card(title, value, color):
    return f"""
    <div class="kpi-card {color}">
        <div class="kpi-title">{title}</div>
        <div class="kpi-value">{value}</div>
    </div>
    """

# =========================
# KPI UI
# =========================
cols = st.columns(5)
metrics = [
    ("Demands", grand.get("dem", 0), "red"),
    ("Submissions", grand.get("sub", 0), "blue"),
    ("Interviews", grand.get("l1", 0), "orange"),
    ("Selections", grand.get("sel", 0), "green"),
    ("Active HC", grand.get("active_hc", 0), "blue"),
]

for col, (t, v, c) in zip(cols, metrics):
    col.markdown(kpi_card(t, f"{v:,}", c), unsafe_allow_html=True)

# =========================
# MOM STRIP
# =========================
curr_year = int(selected_years[0])
curr_month = int(selected_months[0])
prev_year, prev_month = get_previous_month(curr_year, curr_month)

prev = get_data([prev_year], [prev_month], selected_clients)

def mom(label, c, p):
    change = calculate_mom(c, p)
    arrow = "▲" if change >= 0 else "▼"
    color = "#27ae60" if change >= 0 else "#e74c3c"
    return f"<span><b>{label}</b> {c:,} <span style='color:{color}'>{arrow} {abs(change):.0f}%</span></span>"

st.markdown(f"""
<div style="background:#fff;padding:12px;border-radius:10px;">
<b>VS {prev_month}</b>
{mom("Dem", grand.get("dem",0), prev.get("dem",0))}
{mom("Sub", grand.get("sub",0), prev.get("sub",0))}
</div>
""", unsafe_allow_html=True)

# =========================
# CHART FILTERS
# =========================
st.markdown("### Day-on-Day Trends")

range_option = st.radio("", ["Last 7 Days","Last 15 Days","Last 30 Days"], horizontal=True)

col1, col2 = st.columns(2)
from_date = col1.date_input("From")
to_date = col2.date_input("To")

metric_map = {
    "Demands": "dem",
    "Submissions": "sub",
    "Interviews": "intv",
    "Selections": "sel",
}

selected_metric = st.radio("", list(metric_map.keys()), horizontal=True)

# =========================
# DATE LOGIC
# =========================
today = datetime.now().date()

if range_option == "Last 7 Days":
    start = today - timedelta(days=7)
elif range_option == "Last 15 Days":
    start = today - timedelta(days=15)
else:
    start = today - timedelta(days=30)

end = to_date if to_date else today

# =========================
# TRENDS
# =========================
day_trends = daily_trends_cached(None, None, None, "day")

filtered = []
for i in day_trends.get(metric_map[selected_metric], []):
    d = pd.to_datetime(i["d"]).date()
    if start <= d <= end:
        filtered.append(i)

# =========================
# CHARTS
# =========================
c1, c2 = st.columns(2)

with c1:
    st.subheader("Daily")
    st.line_chart(series_to_df(filtered, selected_metric).set_index("Period"))

with c2:
    st.subheader("Monthly")
    month = daily_trends_cached(None, None, None, "month")
    st.line_chart(series_to_df(month.get(metric_map[selected_metric], []), selected_metric).set_index("Period"))