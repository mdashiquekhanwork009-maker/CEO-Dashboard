import pandas as pd
import streamlit as st
from datetime import datetime

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
    load_data_cached,
    resolve_client_filter_cached,
    round_m,
)


def calculate_mom(current, previous):
    if previous == 0:
        return 0
    return ((current - previous) / previous) * 100


def get_previous_month(year, month):
    if month == 1:
        return year - 1, 12
    return year, month - 1

st.set_page_config(page_title="J2W Dashboard", layout="wide")

st.markdown("""
<style>
.kpi-card {
    background: #ffffff;
    border-radius: 14px;
    padding: 18px;
    height: 150px;                 /* 🔥 Equal height */
    display: flex;
    flex-direction: column;
    justify-content: space-between;  /* 🔥 Balanced spacing */
    box-shadow: 0 2px 10px rgba(0,0,0,0.05);
    margin-bottom: 10px;
}
.kpi-title {
    font-size: 14px;
    color: #6c757d;
    font-weight: 600;
}

.kpi-value {
    font-size: 32px;
    font-weight: 700;
    color: #2c3e50;
    margin: 4px 0;
}

.kpi-sub {
    font-size: 12px;
    color: #7f8c8d;
}
            
.red { border-top: 4px solid #e74c3c; }
.green { border-top: 4px solid #2ecc71; }
.blue { border-top: 4px solid #3498db; }
.orange { border-top: 4px solid #f39c12; }
.badge {
    display: inline-block;
    padding: 4px 10px;
    border-radius: 20px;
    font-size: 12px;
    background-color: #f8d7da;
    color: #721c24;   /* already good */
    margin-top: 8px;
}
            
</style>
            div[data-testid="column"] {
    padding: 0 6px;   /* 🔥 equal gap between cards */
}

""", unsafe_allow_html=True)



@st.cache_data(show_spinner=False)
def get_streamlit_init_data():
    periods = get_periods_cached()
    client_meta = list(get_client_catalog())
    years = sorted({str(p[0]) for p in periods}, reverse=True)
    months = sorted({int(p[1]) for p in periods})
    return periods, client_meta, years, months


@st.cache_data(show_spinner=False)
def get_visible_rows(year_values, month_values, client_values, domain_values, bh_values):
    years = {int(y) for y in year_values} if year_values else None
    months = {int(m) for m in month_values} if month_values else None
    clients = set(client_values) if client_values else None
    domains = set(domain_values) if domain_values else None
    bhs = set(bh_values) if bh_values else None

    resolved_clients = resolve_client_filter_cached(
        freeze_filter(clients),
        freeze_filter(domains),
        freeze_filter(bhs),
    )

    result = compute_all_cached(
        freeze_filter(years),
        freeze_filter(months),
        freeze_filter(resolved_clients),
    )
    client_to_domain, client_to_bh, client_lookup, _, _ = get_mapping_context()

    for client in list(result.keys()):
        if CAPTIVE_SUFFIX in client:
            base = client.replace(CAPTIVE_SUFFIX, "").strip()
            client_to_domain[client] = "Captive"
            client_to_bh[client] = client_to_bh.get(base, "")

    visible = []
    for client, metrics in sorted(result.items(), key=lambda item: item[0].lower()):
        mapped = get_mapped_client_name(client, client_lookup)
        domain = client_to_domain.get(mapped, "")
        bh = client_to_bh.get(mapped, "")
        if domains and domain not in domains:
            continue
        if bhs and bh not in bhs:
            continue
        visible.append(
            {
                "Client": client,
                "Domain": domain or "Unmapped",
                "Business Head": bh or "Unassigned",
                **round_m(metrics),
            }
        )

    grand_source = {
        row["Client"]: {
            key: value
            for key, value in row.items()
            if key not in {"Client", "Domain", "Business Head"}
        }
        for row in visible
    }
    grand = grand_total(grand_source)
    return visible, round_m(grand), list(resolved_clients or [])


@st.cache_data(show_spinner=False)
def get_trend_frames(client_values):
    clients = set(client_values) if client_values else None
    day = daily_trends_cached(freeze_filter(clients), None, None, "day")
    month = daily_trends_cached(freeze_filter(clients), None, None, "month")
    return day, month


def series_to_df(series, label):
    if not series:
        return pd.DataFrame(columns=["Period", label])
    return pd.DataFrame(
        {
            "Period": [item["d"] for item in series],
            label: [item["v"] for item in series],
        }
    )


_, client_meta, year_options, month_options = get_streamlit_init_data()
all_clients = [entry["name"] for entry in client_meta]
domain_options = sorted({entry["domain"] for entry in client_meta if entry["domain"]})
bh_options = sorted({entry["bh"] for entry in client_meta if entry["bh"]})

today = datetime.now()

# Year default
current_year_default = [str(today.year)] if str(today.year) in year_options else [year_options[0]]

# Month default
current_month_default = [today.month] if today.month in month_options else [month_options[-1]]

st.title("J2W Dashboard")
st.caption("")

with st.sidebar:
    st.header("Filters")
    selected_years = st.multiselect("Year", year_options, default=current_year_default)
    selected_months = st.multiselect("Month", month_options, default=current_month_default)
    selected_domains = st.multiselect("Domain", domain_options)
    selected_bhs = st.multiselect("Business Head", bh_options)
    selected_clients = st.multiselect("Clients", all_clients)

visible_rows, grand, resolved_clients = get_visible_rows(
    tuple(selected_years),
    tuple(selected_months),
    tuple(selected_clients),
    tuple(selected_domains),
    tuple(selected_bhs),
)

# Get selected month/year
if selected_years and selected_months:
    curr_year = int(selected_years[0])
    curr_month = int(selected_months[0])
else:
    curr_year = datetime.now().year
    curr_month = datetime.now().month

prev_year, prev_month = get_previous_month(curr_year, curr_month)

# Get previous month data
_, prev_grand, _ = get_visible_rows(
    (prev_year,),   # ✅ FIXED
    (prev_month,), 
    tuple(selected_clients),
    tuple(selected_domains),
    tuple(selected_bhs),
)

def kpi_card(title, value, subtext="", color="blue", badge=None):
    return f"""
    <div class="kpi-card {color}">
        <div class="kpi-title">{title}</div>
        <div class="kpi-value">{value}</div>
        <div class="kpi-sub">{subtext}</div>
        {f'<div class="badge">{badge}</div>' if badge else ''}
    </div>
    """

def mom_item(label, curr, prev):
    change = calculate_mom(curr, prev)
    arrow = "▲" if change >= 0 else "▼"
    color = "#27ae60" if change >= 0 else "#e74c3c"

    return f"<div style='min-width:120px; display:inline-block;'>\
<div style='font-size:12px; color:#7f8c8d;'>{label}</div>\
<div style='font-weight:600;'>{curr:,} \
<span style='color:{color}; margin-left:6px;'>{arrow} {abs(change):.0f}%</span>\
</div></div>"

st.markdown("###")
st.markdown("<br>", unsafe_allow_html=True)

mom_html = f"""
<div style="
    background:#ffffff;
    color:#2c3e50;
    padding:14px 18px;
    border-radius:14px;
    font-size:13px;
    display:flex;
    flex-wrap:wrap;
    align-items:center;
    gap:18px;   /* 🔥 spacing between items */
    box-shadow: 0 2px 8px rgba(0,0,0,0.05);
">
<b style="margin-right:10px;">VS {pd.Timestamp(prev_year, prev_month, 1).strftime('%b %Y')}</b>

{mom_item("Demands", grand.get('dem',0), prev_grand.get('dem',0))}
{mom_item("Submissions", grand.get('sub',0), prev_grand.get('sub',0))}
{mom_item("L1", grand.get('l1',0), prev_grand.get('l1',0))}
{mom_item("L2", grand.get('l2',0), prev_grand.get('l2',0))}
{mom_item("Selections", grand.get('sel',0), prev_grand.get('sel',0))}
{mom_item("Onboarded", grand.get('ob_hc',0), prev_grand.get('ob_hc',0))}
{mom_item("Exits", grand.get('ex_hc',0), prev_grand.get('ex_hc',0))}
{mom_item("Net HC", grand.get('net_hc',0), prev_grand.get('net_hc',0))}
{mom_item("Net PO", grand.get('net_po',0), prev_grand.get('net_po',0))}
{mom_item("Net Margin", grand.get('net_mg',0), prev_grand.get('net_mg',0))}

</div>
"""

col1, col2, col3, col4, col5, col6 = st.columns(6)

with col1:
    st.markdown(kpi_card(
        "Demands",
        f"{int(grand.get('dem',0)):,}",
        f"{int(grand.get('dem_open',0))} openings",
        "red",
        f"{int(grand.get('dem_u',0))} pending"
    ), unsafe_allow_html=True)

with col2:
    st.markdown(kpi_card(
        "Submissions",
        f"{int(grand.get('sub',0)):,}",
        f"{int(grand.get('sub_fp',0))} pending",
        "blue"
    ), unsafe_allow_html=True)

with col3:
    st.markdown(kpi_card(
        "Interviews",
        f"{int(grand.get('l1',0)+grand.get('l2',0)+grand.get('l3',0)):,}",
        f"L1 {int(grand.get('l1',0))} | L2 {int(grand.get('l2',0))}",
        "orange"
    ), unsafe_allow_html=True)

with col4:
    st.markdown(kpi_card(
        "Selections",
        f"{int(grand.get('sel',0)):,}",
        "Confirmed candidates",
        "green"
    ), unsafe_allow_html=True)

with col5:
    st.markdown(kpi_card(
        "Selection Pipeline",
        f"{int(grand.get('sp_hc',0)):,}",
        f"₹{grand.get('sp_po',0):.2f}L PO",
        "blue"
    ), unsafe_allow_html=True)

with col6:
    st.markdown(kpi_card(
        "Active HC",
        f"{int(grand.get('active_hc',0)):,}",
        "Current active",
        "blue"
    ), unsafe_allow_html=True)

col1, col2, col3, col4, col5 = st.columns(5)

with col1:
    st.markdown(kpi_card(
        "Onboarded",
        f"{int(grand.get('ob_hc',0)):,}",
        f"₹{grand.get('ob_po',0):.2f}L",
        "green"
    ), unsafe_allow_html=True)

with col2:
    st.markdown(kpi_card(
        "Exits",
        f"{int(grand.get('ex_hc',0)):,}",
        f"₹{grand.get('ex_po',0):.2f}L",
        "red"
    ), unsafe_allow_html=True)

with col3:
    st.markdown(kpi_card(
        "Net HC",
        f"{int(grand.get('net_hc',0)):,}",
        "Movement",
        "red" if grand.get('net_hc',0) < 0 else "green"
    ), unsafe_allow_html=True)

with col4:
    st.markdown(kpi_card(
        "Net PO",
        f"₹{grand.get('net_po',0):.2f}L",
        "",
        "red" if grand.get('net_po',0) < 0 else "green"
    ), unsafe_allow_html=True)

with col5:
    st.markdown(kpi_card(
        "Net Margin",
        f"₹{grand.get('net_mg',0):.2f}L",
        "",
        "red" if grand.get('net_mg',0) < 0 else "green"
    ), unsafe_allow_html=True)

st.markdown(mom_html, unsafe_allow_html=True)


day_trends, month_trends = get_trend_frames(tuple(resolved_clients))

trend_col1, trend_col2 = st.columns(2)
with trend_col1:
    st.subheader("Daily Trends")
    day_metric = st.selectbox(
        "Daily metric",
        [
            ("dem", "Demands"),
            ("dem_u", "Unserviced Demands"),
            ("sub", "Submissions"),
            ("sub_fp", "Feedback Pending"),
            ("intv", "Interviews"),
            ("sel", "Selections"),
            ("ob", "Onboardings"),
            ("ex", "Exits"),
        ],
        format_func=lambda item: item[1],
        key="day_metric",
    )
    st.line_chart(series_to_df(day_trends.get(day_metric[0], []), day_metric[1]).set_index("Period"))

with trend_col2:
    st.subheader("Month-on-Month Trends")
    month_metric = st.selectbox(
        "Monthly metric",
        [
            ("dem", "Demands"),
            ("dem_u", "Unserviced Demands"),
            ("sub", "Submissions"),
            ("sub_fp", "Feedback Pending"),
            ("intv", "Interviews"),
            ("sel", "Selections"),
            ("ob", "Onboardings"),
            ("hc", "Headcount Movement"),
            ("ex", "Exits"),
        ],
        format_func=lambda item: item[1],
        key="month_metric",
    )
    st.line_chart(series_to_df(month_trends.get(month_metric[0], []), month_metric[1]).set_index("Period"))

with st.expander("Recruitment Pipeline", expanded=False):
    pipeline_cols = ["Client", "dem", "sub", "l1", "l2", "l3", "sel", "ob_hc", "Domain", "Business Head"]
    pipeline_df = pd.DataFrame(visible_rows)
    if not pipeline_df.empty:
        st.dataframe(
            pipeline_df[[col for col in pipeline_cols if col in pipeline_df.columns]].rename(
                columns={
                    "dem": "Demands",
                    "sub": "Submissions",
                    "l1": "L1",
                    "l2": "L2",
                    "l3": "L3",
                    "sel": "Selections",
                    "ob_hc": "Onboarded",
                }
            ),
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.info("No pipeline activity found for the current filter.")

with st.expander("Client Breakdown - MTD", expanded=False):
    table_df = pd.DataFrame(visible_rows)
    if not table_df.empty:
        preferred_columns = [
            "Client", "Domain", "Business Head", "dem", "dem_u", "sub", "sub_fp", "l1", "l2", "l3",
            "sel", "sp_hc", "sp_po", "sp_mg", "ob_hc", "ob_po", "ob_mg", "ex_hc", "ex_po", "net_hc", "net_po", "net_mg",
        ]
        st.dataframe(table_df[[c for c in preferred_columns if c in table_df.columns]], use_container_width=True, hide_index=True)
    else:
        st.info("No client breakdown rows found for the current filter.")

with st.expander("Raw Data Explorer", expanded=False):
    raw_dataset = st.selectbox(
        "Dataset",
        options=list(RAW_DATASET_CONFIG.keys()),
        format_func=lambda key: RAW_DATASET_CONFIG[key]["label"],
    )
    raw_col1, raw_col2 = st.columns(2)
    with raw_col1:
        raw_month = st.date_input("Month picker", value=None, format="YYYY-MM-DD")
        raw_from = st.date_input("From Date", value=None, format="YYYY-MM-DD")
    with raw_col2:
        raw_to = st.date_input("To Date", value=None, format="YYYY-MM-DD")
        demand_status = st.radio("Demand Filter", ["all", "unserviced", "serviced"], horizontal=True)

    raw_clients = st.multiselect("Raw data clients", all_clients)

    month_filter = None
    year_filter = None
    if raw_month:
        raw_month_ts = pd.Timestamp(raw_month)
        year_filter = {int(raw_month_ts.year)}
        month_filter = {int(raw_month_ts.month)}

    raw_df = get_raw_dataset_frame(
        raw_dataset,
        year_filter=year_filter,
        month_filter=month_filter,
        client_filter=set(raw_clients) if raw_clients else None,
        from_date=pd.Timestamp(raw_from) if raw_from else None,
        to_date=pd.Timestamp(raw_to) if raw_to else None,
        demand_status=demand_status,
    )
    visible_columns = [col for col in raw_df.columns if not col.startswith("_")]
    st.caption(f"{RAW_DATASET_CONFIG[raw_dataset]['label']}: {len(raw_df):,} row(s)")
    if visible_columns:
        st.dataframe(raw_df[visible_columns], use_container_width=True, hide_index=True)
    else:
        st.info("No raw records found for the selected filters.")