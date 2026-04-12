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

/* Range buttons */
.range-btn-row {display:flex;gap:8px;margin-bottom:10px;}
.range-btn {
    padding:6px 18px;border-radius:20px;border:1.5px solid #dee2e6;
    background:white;font-size:13px;font-weight:600;color:#495057;cursor:pointer;
}
.range-btn.active {background:#e74c3c;color:white;border-color:#e74c3c;}

/* Metric pill buttons */
.metric-pill-row {display:flex;gap:6px;flex-wrap:wrap;margin:10px 0;}
.metric-pill {
    padding:5px 14px;border-radius:20px;border:1.5px solid #dee2e6;
    background:white;font-size:12px;font-weight:600;color:#495057;cursor:pointer;
    display:inline-flex;align-items:center;gap:5px;
}
.metric-pill.active {background:#e74c3c;color:white;border-color:#e74c3c;}

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
        freeze_filter(set(map(int, y))) if y else None,
        freeze_filter(set(m)) if m else None,
        freeze_filter(set(c)) if c else None,
    )
    return round_m(grand_total(res))

grand = get_data(years, months, clients)

# =========================
# KPI CARDS
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
    ("Demands",           grand.get("dem",0),       "red",    "📋", ""),
    ("Submissions",       grand.get("sub",0),        "blue",   "📤", ""),
    ("Interviews",        grand.get("l1",0),         "orange", "🎤", ""),
    ("Selections",        grand.get("sel",0),        "green",  "✅", ""),
    ("Selection Pipeline",grand.get("sp_hc",0),      "blue",   "📊", ""),
    ("Active HC",         grand.get("active_hc",0),  "blue",   "👥", ""),
]
for col,(t,v,c,i,s) in zip(row1,data1):
    col.markdown(card(t,f"{v:,}",c,i,s), unsafe_allow_html=True)

row2 = st.columns(5)
data2 = [
    ("Onboarded",  grand.get("ob_hc",0),                        "green","🚀",""),
    ("Exits",      grand.get("ex_hc",0),                        "red",  "🚪",""),
    ("Net HC",     grand.get("net_hc",0),                       "green","📈",""),
    ("Net PO",     f"₹{grand.get('net_po',0):.2f}L",            "green","💰",""),
    ("Net Margin", f"₹{grand.get('net_mg',0):.2f}L",            "green","📊",""),
]
for col,(t,v,c,i,s) in zip(row2,data2):
    col.markdown(card(t,v,c,i,s), unsafe_allow_html=True)

# =========================
# COMPARISONS
# =========================
cy = int(years[0]) if years else datetime.now().year
cm = int(months[0]) if months else datetime.now().month
py, pm = get_previous_month(cy, cm)
prev = get_data([str(py)],[pm],clients)

def comp(label, cur, pr):
    ch = calculate_mom(cur, pr)
    arrow = "▲" if ch >= 0 else "▼"
    color = "#27ae60" if ch >= 0 else "#e74c3c"
    return f"<span><b>{label}</b> {cur:,} <span style='color:{color}'>{arrow} {abs(ch):.0f}%</span></span>"

st.markdown(f"""
<div style="background:#ffffff;padding:12px;border-radius:10px;margin-top:10px;display:flex;gap:20px;flex-wrap:wrap;">
<b>VS LMTD</b>
{comp("Dem", grand.get("dem",0), prev.get("dem",0))}
{comp("Sub", grand.get("sub",0), prev.get("sub",0))}
{comp("L1",  grand.get("l1",0),  prev.get("l1",0))}
{comp("Sel", grand.get("sel",0), prev.get("sel",0))}
{comp("Ob",  grand.get("ob_hc",0), prev.get("ob_hc",0))}
{comp("Ex",  grand.get("ex_hc",0), prev.get("ex_hc",0))}
</div>
""", unsafe_allow_html=True)

st.markdown(f"""
<div style="background:#f8f9fb;padding:12px;border-radius:10px;margin-top:8px;display:flex;gap:20px;flex-wrap:wrap;">
<b>VS {calendar.month_abbr[pm]} {py}</b>
{comp("Dem", grand.get("dem",0), prev.get("dem",0))}
{comp("Sub", grand.get("sub",0), prev.get("sub",0))}
{comp("L1",  grand.get("l1",0),  prev.get("l1",0))}
{comp("Sel", grand.get("sel",0), prev.get("sel",0))}
{comp("Ob",  grand.get("ob_hc",0), prev.get("ob_hc",0))}
{comp("Ex",  grand.get("ex_hc",0), prev.get("ex_hc",0))}
</div>
""", unsafe_allow_html=True)

# =========================
# DAY-ON-DAY CHART
# =========================
st.markdown("### Day-on-Day Trends")

# --- Range selector ---
dod_range_options = ["Last 7 Days", "Last 15 Days", "Last 30 Days"]
if "dod_range" not in st.session_state:
    st.session_state["dod_range"] = "Last 7 Days"

dod_range_cols = st.columns([1.1, 1.2, 1.2, 5.5])
for i, label in enumerate(dod_range_options):
    with dod_range_cols[i]:
        if st.button(label, key=f"dod_range_{i}", use_container_width=True):
            st.session_state["dod_range"] = label
            st.rerun()

# --- Custom date range ---
dod_date_col1, _, dod_date_col2, _ = st.columns([1.5, 0.1, 1.5, 5])
with dod_date_col1:
    dod_from = st.date_input("From", value=None, format="DD-MM-YYYY", key="dod_from_input")
with dod_date_col2:
    dod_to = st.date_input("To", value=None, format="DD-MM-YYYY", key="dod_to_input")

# --- Metric pill selector ---
dod_metric_options = [
    ("📋", "Demand",             "dem"),
    ("",   "Unserviced Demands", "dem_u"),
    ("📤", "Submission",         "sub"),
    ("",   "Feedback Pending",   "sub_fp"),
    ("🎯", "Interview",          "intv"),
    ("✅", "Selection",          "sel"),
    ("🚀", "Onboarding",         "ob"),
    ("🚪", "Exit",               "ex"),
]

if "dod_metric" not in st.session_state:
    st.session_state["dod_metric"] = "dem"

dod_pill_cols = st.columns(len(dod_metric_options))
for i, (icon, label, key) in enumerate(dod_metric_options):
    with dod_pill_cols[i]:
        btn_label = f"{icon} {label}" if icon else label
        if st.button(btn_label, key=f"dod_pill_{key}", use_container_width=True):
            st.session_state["dod_metric"] = key
            st.rerun()

# --- Build chart data ---
dod_active_metric = st.session_state["dod_metric"]
dod_active_label  = next((label for _, label, k in dod_metric_options if k == dod_active_metric), dod_active_metric)

data_day = daily_trends_cached(None, None, None, "day")
raw_series_day = data_day.get(dod_active_metric, [])
df_day = series_to_df(raw_series_day, dod_active_label)

# Apply range filter
today = datetime.today()
if dod_from and dod_to:
    df_day["_date"] = pd.to_datetime(df_day["Period"], errors="coerce")
    df_day = df_day[(df_day["_date"] >= pd.Timestamp(dod_from)) & (df_day["_date"] <= pd.Timestamp(dod_to))]
    df_day = df_day.drop(columns=["_date"])
elif not df_day.empty:
    n_days = {"Last 7 Days": 7, "Last 15 Days": 15, "Last 30 Days": 30}.get(
        st.session_state["dod_range"], 7
    )
    cutoff = today - timedelta(days=n_days)
    df_day["_date"] = pd.to_datetime(df_day["Period"], errors="coerce")
    df_day = df_day[df_day["_date"] >= cutoff]
    df_day = df_day.drop(columns=["_date"])

y_col_day = dod_active_label
if not df_day.empty and y_col_day in df_day.columns:
    fig_dod = px.line(df_day, x="Period", y=y_col_day, markers=True,
                      color_discrete_sequence=["#e74c3c"])
    fig_dod.update_traces(
        mode="lines+markers+text",
        text=df_day[y_col_day],
        textposition="top center",
        line=dict(width=2.5),
        marker=dict(size=7),
    )
    fig_dod.update_layout(
        plot_bgcolor="white",
        paper_bgcolor="white",
        margin=dict(t=20, b=20, l=10, r=10),
        xaxis=dict(showgrid=False, title=""),
        yaxis=dict(showgrid=True, gridcolor="#f0f0f0", title=""),
        font=dict(family="sans-serif", size=12),
    )
    st.plotly_chart(fig_dod, use_container_width=True)
else:
    st.info("No data available for the selected range and metric.")

# =========================
# MONTH-ON-MONTH CHART
# =========================
st.markdown("### Month-on-Month Trends")

# --- Range selector ---
range_options = ["Last 3 Months", "Last 6 Months", "Last 12 Months"]
if "mom_range" not in st.session_state:
    st.session_state["mom_range"] = "Last 12 Months"

range_cols = st.columns([1.2, 1.2, 1.5, 4])
for i, label in enumerate(range_options):
    is_active = st.session_state["mom_range"] == label
    btn_style = (
        "background:#e74c3c;color:white;border:1.5px solid #e74c3c;"
        if is_active else
        "background:white;color:#495057;border:1.5px solid #dee2e6;"
    )
    with range_cols[i]:
        if st.button(
            label,
            key=f"range_{i}",
            use_container_width=True,
        ):
            st.session_state["mom_range"] = label
            st.session_state.pop("mom_from", None)
            st.session_state.pop("mom_to", None)
            st.rerun()

# --- Custom date range ---
date_col1, _, date_col2, _ = st.columns([1.5, 0.1, 1.5, 5])
with date_col1:
    mom_from = st.date_input("From", value=None, format="DD-MM-YYYY", key="mom_from_input")
with date_col2:
    mom_to = st.date_input("To", value=None, format="DD-MM-YYYY", key="mom_to_input")

# --- Metric pill selector ---
mom_metric_options = [
    ("📋", "Demand",            "dem"),
    ("",   "Unserviced Demands","dem_u"),
    ("📤", "Submission",        "sub"),
    ("",   "Feedback Pending",  "sub_fp"),
    ("🎯", "Interview",         "intv"),
    ("✅", "Selection",         "sel"),
    ("🚀", "Onboarding",        "ob"),
    ("👥", "Headcount",         "hc"),
    ("🚪", "Exit",              "ex"),
]

if "mom_metric" not in st.session_state:
    st.session_state["mom_metric"] = "dem"

pill_cols = st.columns(len(mom_metric_options))
for i, (icon, label, key) in enumerate(mom_metric_options):
    is_active = st.session_state["mom_metric"] == key
    with pill_cols[i]:
        btn_label = f"{icon} {label}" if icon else label
        if st.button(btn_label, key=f"pill_{key}", use_container_width=True):
            st.session_state["mom_metric"] = key
            st.rerun()

# Style active pill via custom CSS injected dynamically
active_metric = st.session_state["mom_metric"]
active_label  = next((label for _, label, k in mom_metric_options if k == active_metric), "")

st.markdown(f"""
<style>
/* Highlight active pill button */
div[data-testid="stButton"] button[kind="secondary"] {{
    border-radius: 20px !important;
    font-size: 12px !important;
    padding: 4px 10px !important;
}}
</style>
""", unsafe_allow_html=True)

# --- Build chart data ---
data_month = daily_trends_cached(None, None, None, "month")
raw_series = data_month.get(active_metric, [])
df_month = series_to_df(raw_series, active_label or active_metric)

# Apply range filter
today = datetime.today()
if mom_from and mom_to:
    # Custom range overrides buttons
    from_ts = pd.Timestamp(mom_from)
    to_ts   = pd.Timestamp(mom_to)
    df_month["_date"] = pd.to_datetime(df_month["Period"], errors="coerce")
    df_month = df_month[(df_month["_date"] >= from_ts) & (df_month["_date"] <= to_ts)]
    df_month = df_month.drop(columns=["_date"])
elif not df_month.empty:
    n_months = {"Last 3 Months": 3, "Last 6 Months": 6, "Last 12 Months": 12}.get(
        st.session_state["mom_range"], 12
    )
    cutoff = today - timedelta(days=n_months * 30)
    df_month["_date"] = pd.to_datetime(df_month["Period"], errors="coerce")
    df_month = df_month[df_month["_date"] >= cutoff]
    df_month = df_month.drop(columns=["_date"])

y_col = active_label or active_metric
if not df_month.empty and y_col in df_month.columns:
    fig_mom = px.line(df_month, x="Period", y=y_col, markers=True,
                      color_discrete_sequence=["#e74c3c"])
    fig_mom.update_traces(
        mode="lines+markers+text",
        text=df_month[y_col],
        textposition="top center",
        line=dict(width=2.5),
        marker=dict(size=7),
    )
    fig_mom.update_layout(
        plot_bgcolor="white",
        paper_bgcolor="white",
        margin=dict(t=20, b=20, l=10, r=10),
        xaxis=dict(showgrid=False, title=""),
        yaxis=dict(showgrid=True, gridcolor="#f0f0f0", title=""),
        font=dict(family="sans-serif", size=12),
    )
    st.plotly_chart(fig_mom, use_container_width=True)
else:
    st.info("No data available for the selected range and metric.")