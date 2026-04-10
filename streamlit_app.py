import pandas as pd
import streamlit as st

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


st.set_page_config(page_title="Recruitment Dashboard", layout="wide")


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

current_year_default = [year_options[0]] if year_options else []
current_month_default = [month_options[-1]] if month_options else []

st.title("Recruitment Dashboard")
st.caption("Starter Streamlit version powered by your existing dashboard logic.")

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

metric_cols = st.columns(5)
metric_cols[0].metric("Demands", f"{int(grand.get('dem', 0)):,}")
metric_cols[1].metric("Unserviced", f"{int(grand.get('dem_u', 0)):,}")
metric_cols[2].metric("Submissions", f"{int(grand.get('sub', 0)):,}")
metric_cols[3].metric("Feedback Pending", f"{int(grand.get('sub_fp', 0)):,}")
metric_cols[4].metric("Onboarded", f"{int(grand.get('ob_hc', 0)):,}")

metric_cols2 = st.columns(5)
metric_cols2[0].metric("Selections", f"{int(grand.get('sel', 0)):,}")
metric_cols2[1].metric("Exits", f"{int(grand.get('ex_hc', 0)):,}")
metric_cols2[2].metric("Net HC", f"{int(grand.get('net_hc', 0)):,}")
metric_cols2[3].metric("Net PO (L)", f"{grand.get('net_po', 0):,.2f}")
metric_cols2[4].metric("Net Margin (L)", f"{grand.get('net_mg', 0):,.2f}")

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