"""
J2W Recruitment Dashboard — Streamlit
Mirrors dashboard.py layout, KPI formulas, filter logic, and trend sections exactly.
"""
import calendar
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from datetime import datetime, timedelta
from dashboard import RAW_DATASET_CONFIG, get_raw_dataset_frame
from dashboard import (
    CAPTIVE_SUFFIX,
    compute_all_cached,
    freeze_filter,
    get_client_catalog,
    get_mapping_context,
    get_mapped_client_name,
    get_periods_cached,
    grand_total,
    load_mapping,
    normalize_bh_label,
    resolve_client_filter_cached,
    round_m,
)
# ─── PAGE CONFIG ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="J2W Recruitment Dashboard",
    layout="wide",
    page_icon="📊",
    initial_sidebar_state="expanded"
)

# ─── CSS ──────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
/* Background */
[data-testid="stSidebar"] {
    display: block !important;
}
[data-testid="stMainBlockContainer"] { padding-top: 12px; }

/* KPI Cards */
.kpi {
    background: var(--background-color);
    border: 1px solid var(--secondary-background-color);
    padding: 16px;
    border-radius: 12px;
    min-height: 140px;
    display: flex;
    flex-direction: column;
    justify-content: space-between;
}
            
.kpi::after { content:''; position:absolute; top:0; left:0; right:0;
              height:4px; border-radius:14px 14px 0 0; }
.kpi.red::after    { background:linear-gradient(90deg,#e8453c,#ff7043); }
.kpi.green::after  { background:linear-gradient(90deg,#1db85a,#27ae60); }
.kpi.grey::after   { background:linear-gradient(90deg,#5b6578,#99a3b6); }
.kpi.blue::after   { background:linear-gradient(90deg,#3498db,#69b8ff); }

.kpi-val { font-size:36px; font-weight:900; line-height:1; margin-bottom:6px; letter-spacing:-.02em; }
.kpi-val.red   { color:#e8453c; }
.kpi-val.green { color:#1db85a; }
.kpi-val.blue  { color:#3498db; }
.kpi-sub strong { color: var(--text-color); }
.kpi-lbl { font-size:10px; font-weight:800; text-transform:uppercase;
           letter-spacing:.1em; color: var(--text-color); opacity: 0.7; margin-bottom:4px; }
.kpi-sub { font-size:12px; color: var(--text-color); opacity: 0.8;
           line-height:1.45; min-height:32px; }
.kpi-val.grey  { color: var(--text-color); }
.ktag { display:inline-flex; align-items:center; gap:3px; font-size:10px;
        font-weight:700; padding:3px 9px; border-radius:20px; margin-top:6px; }
.tg  { background:rgba(29,184,90,.12);  color:#17a050; border:1px solid rgba(29,184,90,.28); }
.tr  { background:rgba(232,69,60,.10);  color:#d63b33; border:1px solid rgba(232,69,60,.28); }
.tb  { background:rgba(52,152,219,.12); color:#2c6ea4; border:1px solid rgba(52,152,219,.28); }
.tgr { background:rgba(107,115,133,.10); color:#5b6578; border:1px solid rgba(107,115,133,.22); }

/* Section title */
.sec {
    font-size:13px;
    font-weight:700;
    text-transform:none;
    letter-spacing:.03em;
    color: var(--text-color);
    margin:25px 0 12px;
}
            
.sec::after { content:''; flex:1; height:1px; background:rgba(0,0,0,0.08); }

/* Comparison bar */
.prev-bar { background:#f5f6fa; border:1px solid rgba(0,0,0,0.08); border-radius:10px;
            padding:10px 16px; margin-bottom:14px; display:flex; flex-wrap:wrap;
            align-items:center; gap:6px 18px; }
.prev-bar-title { font-size:10px; font-weight:800; text-transform:uppercase;
                  letter-spacing:.07em; color:#8a91a0; margin-right:4px; }
.pm  { display:flex; align-items:center; gap:5px; }
.pml { font-size:11px; color:#52586a; }
.pmv { font-size:13px; font-weight:700; color:#1a1d23; }
.pmc { font-size:10px; font-weight:700; }
.sep { width:1px; height:18px; background:rgba(0,0,0,0.1); }

/* Trend pill buttons */
.pill-row { display:flex; flex-wrap:wrap; gap:6px; margin-bottom:10px; }
.pill-btn { padding:6px 14px; border-radius:20px; border:1.5px solid #dee2e6;
            background: var(--secondary-background-color); font-size:12px; font-weight:600; color: var(--text-color);
            cursor:pointer; }
.pill-btn.on { background:#e74c3c; color:white; border-color:#e74c3c; }

/* Range buttons */
.range-btn { padding:6px 16px; border-radius:20px; border:1.5px solid #dee2e6;
             background: var(--secondary-background-color); font-size:12px; font-weight:600; color: var(--text-color); cursor:pointer; }
.range-btn.on { background:#e74c3c; color:white; border-color:#e74c3c; }

/* MRR grid */
.mrr-grid { display:grid; grid-template-columns:repeat(3,1fr); gap:8px; }
.mrr-c { background: var(--secondary-background-color); border:1px solid rgba(0,0,0,0.07); border-radius:10px;
         padding:12px; text-align:center; }
.mrr-v { font-size:20px; font-weight:800; margin-bottom:2px; }
.mrr-l { font-size:10px; color: var(--text-color); opacity: 0.7; text-transform:uppercase; letter-spacing:.04em; }

/* Stage funnel */
.stage-wrap { display:flex; gap:4px; align-items:stretch; overflow-x:auto; }
.stage { flex:1; min-width:72px; background: var(--secondary-background-color); border:1px solid rgba(0,0,0,0.07);
         border-radius:10px; padding:12px 8px; text-align:center; }
.sv { font-size:20px; font-weight:800; line-height:1; margin-bottom:2px; }
.sl { font-size:9px; color: var(--text-color); opacity: 0.7; text-transform:uppercase; letter-spacing:.04em; }
.sc { font-size:10px; font-weight:700; margin-top:4px; }
.sarrow { display:flex; align-items:center; justify-content:center; color: var(--text-color); opacity: 0.7;
          font-size:20px; padding:0 1px; flex-shrink:0; }

/* Volume funnel bars */
.fn { display:flex; align-items:center; gap:10px; margin-bottom:8px; }
.fnl { font-size:11px; color: var(--text-color); opacity: 0.8; width:120px; flex-shrink:0; }
.fnt { flex:1; height:18px; background: var(--secondary-background-color); border-radius:4px; overflow:hidden; }
.fnf { height:100%; border-radius:3px; }
.fnn { font-size:12px; font-weight:700; min-width:50px; text-align:right; }
.fnp { font-size:10px; color: var(--text-color); opacity: 0.7; min-width:34px; text-align:right; }

/* 🔥 Button Styling (ADD HERE) */
button[kind="secondary"] {
    border-radius: 8px !important;
    border: 1px solid var(--secondary-background-color) !important;
}

button[kind="primary"] {
    border-radius: 8px !important;
    background: #e8453c !important;
    color: white !important;
}
            
</style>
""", unsafe_allow_html=True)

MON = ['','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']

# ─── DATA INIT ────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def init_data():
    periods     = get_periods_cached()
    client_meta = list(get_client_catalog())
    years       = sorted({str(p[0]) for p in periods}, reverse=True)
    months      = sorted({int(p[1]) for p in periods})
    _, _, domains, business_heads = load_mapping()
    return client_meta, years, months, domains, business_heads

client_meta, year_options, month_options, domain_options, bh_options = init_data()
all_clients = [c["name"] for c in client_meta]
month_map   = {i: calendar.month_abbr[i] for i in range(1, 13)}

# ─── SIDEBAR FILTERS ──────────────────────────────────────────────────────────
now = datetime.now()

with st.sidebar:

    st.markdown("## 🔎 Filters")

    # YEAR
    current_year = str(now.year)
    selected_years = st.multiselect(
        "YEAR",
        year_options,
        default=[current_year] if current_year in year_options else [year_options[0]]
    )

    # MONTH
    current_month_name = month_map[now.month]
    month_names_list = [month_map[m] for m in month_options]

    if current_month_name in month_names_list:
        default_month = [current_month_name]
    else:
        default_month = [month_map[month_options[-1]]] if month_options else []

    selected_month_names = st.multiselect(
        "MONTH",
        month_names_list,
        default=default_month
    )

    selected_months = [k for k, v in month_map.items() if v in selected_month_names]

    # CLIENT
    selected_clients = st.multiselect("CLIENTS", all_clients)

    # DOMAIN
    selected_domains = st.multiselect("DOMAIN", domain_options)

    # BH
    selected_bhs = st.multiselect("BH", bh_options)

    # RESET
    if st.button("🔄 Reset Filters"):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()
# =========================
# APPLY UI FILTERS (ADD THIS)
# =========================
# ─── DATA FETCH ───────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def get_grand(y_tup, m_tup, c_tup, d_tup, b_tup, from_date=None, to_date=None):
    resolved = resolve_client_filter_cached(
        freeze_filter(set(c_tup)) if c_tup else None,
        freeze_filter(set(d_tup)) if d_tup else None,
        freeze_filter(set(b_tup)) if b_tup else None,
    )
    res = compute_all_cached(
        freeze_filter({int(y) for y in y_tup}) if y_tup else None,
        freeze_filter({int(m) for m in m_tup}) if m_tup else None,
        resolved,
        from_date,
        to_date
)
    # Build visible rows with domain/BH
    client_to_domain, client_to_bh, client_lookup, _, _ = get_mapping_context()
    for cl in list(res.keys()):
        if CAPTIVE_SUFFIX in cl:
            base = cl.replace(CAPTIVE_SUFFIX, "").strip()
            client_to_domain[cl] = "Captive"
            client_to_bh[cl] = client_to_bh.get(base, "")
    dom_set = set(d_tup) if d_tup else None
    bh_set  = set(b_tup) if b_tup else None
    visible = {}
    rows = []
    for cl, m in sorted(res.items(), key=lambda x: x[0].lower()):
        mapped = get_mapped_client_name(cl, client_lookup)
        domain = client_to_domain.get(mapped, "")
        bh     = normalize_bh_label(client_to_bh.get(mapped, ""))
        if dom_set and domain not in dom_set: continue
        if bh_set  and bh     not in bh_set:  continue
        visible[cl] = m
        rows.append({"label": cl, "metrics": m, "domain": domain or "Unmapped", "bh": bh or "Unassigned"})
    g = round_m(grand_total(visible))
    return g, rows

def get_previous_month(year, month):
    return (year - 1, 12) if month == 1 else (year, month - 1)

def pct(a, b): return round(a / b * 100) if b else 0
def L(v):      return f"{abs(float(v)):.2f}"
def sign(v):   return "+" if float(v) >= 0 else "−"
def clr_val(v): return "green" if float(v) >= 0 else "red"

def tag_g(s): return f'<span class="ktag tg">{s}</span>'
def tag_r(s): return f'<span class="ktag tr">{s}</span>'
def tag_b(s): return f'<span class="ktag tb">{s}</span>'
def tag_gr(s):return f'<span class="ktag tgr">{s}</span>'

def kpi_card(color, icon, label, value, value_color, sub, badge):
    return f"""
    <div class="kpi {color}">
      <div style="font-size:22px;margin-bottom:4px">{icon}</div>
      <div class="kpi-lbl">{label}</div>
      <div class="kpi-val {value_color}">{value}</div>
      <div class="kpi-sub">{sub}</div>
      {badge}
    </div>"""

def comp_bar(title, cur, prev):
    metrics = [
        ("Demands",   cur.get("dem",0),    prev.get("dem",0),    False),
        ("Subs",      cur.get("sub",0),    prev.get("sub",0),    False),
        ("L1",        cur.get("l1",0),     prev.get("l1",0),     False),
        ("L2",        cur.get("l2",0),     prev.get("l2",0),     False),
        ("L3",        cur.get("l3",0),     prev.get("l3",0),     False),
        ("Sel",       cur.get("sel",0),    prev.get("sel",0),    False),
        ("Onboarded", cur.get("ob_hc",0),  prev.get("ob_hc",0),  False),
        ("Exits",     cur.get("ex_hc",0),  prev.get("ex_hc",0),  False),
        ("Net HC",    cur.get("net_hc",0), prev.get("net_hc",0), False),
        ("Net PO",    cur.get("net_po",0), prev.get("net_po",0), True),
        ("Net Mgn",   cur.get("net_mg",0), prev.get("net_mg",0), True),
    ]
    items = ""
    for i, (lbl, c, p, is_l) in enumerate(metrics):
        diff = float(c) - float(p)
        pct_chg = round(abs(diff) / abs(float(p)) * 100) if float(p) != 0 else 0
        clr = "#1db85a" if diff >= 0 else "#e8453c"
        arrow = ("▲ " if diff > 0 else "▼ ") + str(pct_chg) + "%" if diff != 0 else "—"
        val_str = f"₹{L(abs(float(p)))}L" if is_l else f"{round(float(p)):,}"
        sep = '<span class="sep"></span>' if i > 0 else ""
        items += f"""{sep}<div class="pm">
            <span class="pml">{lbl}</span>&nbsp;
            <span class="pmv">{val_str}</span>&nbsp;
            <span class="pmc" style="color:{clr}">{arrow}</span>
        </div>"""
    return f'<div class="prev-bar"><span class="prev-bar-title">{title}</span>{items}</div>'

def hex_to_rgba(hex_color, alpha=0.13):
    """Convert hex color to rgba string for Plotly fillcolor."""
    hex_color = hex_color.lstrip("#")
    r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
    return f"rgba({r},{g},{b},{alpha})"

def trend_chart(series, label, color="#e74c3c"):
    if not series:
        return go.Figure()
    df = pd.DataFrame({"Period": [x["d"] for x in series], "Value": [x["v"] for x in series]})
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=df["Period"], y=df["Value"],
        mode="lines+markers+text",
        text=df["Value"],
        textposition="top center",
        textfont=dict(size=11, color="#52586a"),
        line=dict(color=color, width=2.5),
        marker=dict(size=7, color="#fff", line=dict(color=color, width=2)),
        fill="tozeroy",
        fillcolor=hex_to_rgba(color),
    ))
    fig.update_layout(
        height=260, margin=dict(t=20, b=20, l=10, r=10),
        plot_bgcolor="white", paper_bgcolor="white",
        xaxis=dict(showgrid=False, tickfont=dict(size=10, color="#52586a")),
        yaxis=dict(showgrid=True, gridcolor="#f0f0f0", tickfont=dict(size=10, color="#52586a"), rangemode="tozero"),
        showlegend=False,
    )
    return fig
def generate_ceo_insights(current, previous):
    insights = []

    def pct_change(cur, prev):
        if prev == 0:
            return 0
        return round((cur - prev) / prev * 100)

    # Demand servicing
    dem = current.get("dem", 0)
    dem_u = current.get("dem_u", 0)

    if dem > 0:
        unserv_pct = round(dem_u / dem * 100)
        if unserv_pct > 20:
            insights.append(f"⚠️ {unserv_pct}% demands are unserviced — risk to delivery")
        else:
            insights.append(f"✅ Demand servicing healthy ({100 - unserv_pct}% covered)")

    # Submissions
    sub_change = pct_change(current.get("sub", 0), previous.get("sub", 0))
    if sub_change < -10:
        insights.append(f"📉 Submissions dropped {abs(sub_change)}% vs last period")
    elif sub_change > 10:
        insights.append(f"🚀 Submissions increased {sub_change}%")

    # Onboarding
    ob_change = pct_change(current.get("ob_hc", 0), previous.get("ob_hc", 0))
    if ob_change > 10:
        insights.append(f"🚀 Onboarding improved {ob_change}%")
    elif ob_change < -10:
        insights.append(f"📉 Onboarding declined {abs(ob_change)}%")

    # Conversion
    l1 = current.get("l1", 0)
    sel = current.get("sel", 0)
    if l1 > 0:
        conv = round(sel / l1 * 100)
        if conv < 40:
            insights.append(f"⚠️ Low L1→Selection conversion ({conv}%)")
        else:
            insights.append(f"✅ Strong L1→Selection conversion ({conv}%)")

    # Net HC
    net_hc = current.get("net_hc", 0)
    if net_hc < 0:
        insights.append(f"📉 Net HC declined by {abs(net_hc)}")
    elif net_hc > 0:
        insights.append(f"📈 Net HC grew by {net_hc}")

    return insights
def get_lmtd_range(year, month):
    today = datetime.now()

    # Last month logic
    if month == 1:
        lm_year = year - 1
        lm_month = 12
    else:
        lm_year = year
        lm_month = month - 1

    # Same day cutoff
    day = min(today.day, 28)  # safe for Feb

    from_date = datetime(lm_year, lm_month, 1)
    to_date   = datetime(lm_year, lm_month, day)

    return from_date, to_date


# ─── FETCH GRAND DATA ─────────────────────────────────────────────────────────
grand, rows = get_grand(
    tuple(selected_years), 
    tuple(selected_months),
    tuple(selected_clients), 
    tuple(selected_domains), 
    tuple(selected_bhs),
)

# 🔥 BUILD DATE RANGE FROM FILTERS
if selected_years and selected_months:
    year = int(selected_years[0])

    from_date = datetime(year, min(selected_months), 1)

    # safer end date
    last_month = max(selected_months)
    if last_month == 12:
        to_date = datetime(year, 12, 31)
    else:
        next_month = datetime(year, last_month + 1, 1)
        to_date = next_month - timedelta(days=1)
else:
    from_date = None
    to_date = None
# 🔥 LMTD CALCULATION
cy = int(selected_years[0]) if selected_years else datetime.now().year
cm = int(selected_months[0]) if selected_months else datetime.now().month

lm_from, lm_to = get_lmtd_range(cy, cm)

# ⚠️ IMPORTANT: pass same filters
lmtd_grand, _ = get_grand(
    (str(lm_from.year),),
    (lm_from.month,),
    tuple(selected_clients),
    tuple(selected_domains),
    tuple(selected_bhs),
    lm_from,
    lm_to
)

# Previous month for comparison bar
cy = int(selected_years[0]) if selected_years else now.year
cm = int(selected_months[0]) if selected_months else now.month
py, pm = get_previous_month(cy, cm)
prev_grand, _ = get_grand(
    (str(py),), (pm,),
    tuple(selected_clients), tuple(selected_domains), tuple(selected_bhs),
)

# ─── KPI VALUES (matching dashboard.py formulas exactly) ─────────────────────
dem       = int(grand.get("dem", 0))
dem_open  = int(grand.get("dem_open", dem))
dem_u     = int(grand.get("dem_u", 0))
sub       = int(grand.get("sub", 0))
sub_fp    = int(grand.get("sub_fp", 0))
l1        = int(grand.get("l1", 0))
l2        = int(grand.get("l2", 0))
l3        = int(grand.get("l3", 0))
ti        = l1 + l2 + l3
sel       = int(grand.get("sel", 0))
sp_hc     = int(grand.get("sp_hc", 0))
sp_po     = grand.get("sp_po", 0.0)
sp_mg     = grand.get("sp_mg", 0.0)
ob_hc     = int(grand.get("ob_hc", 0))
ob_po     = grand.get("ob_po", 0.0)
ob_mg     = grand.get("ob_mg", 0.0)
active_hc = int(grand.get("active_hc", 0))
ex_hc     = int(grand.get("ex_hc", 0))
ex_po     = grand.get("ex_po", 0.0)
ex_mg     = grand.get("ex_mg", 0.0)
ex_pipe_hc= int(grand.get("ex_pipe_hc", 0))
ex_pipe_po= grand.get("ex_pipe_po", 0.0)
ex_pipe_mg= grand.get("ex_pipe_mg", 0.0)
net_hc    = int(grand.get("net_hc", 0))
net_po    = grand.get("net_po", 0.0)
net_mg    = grand.get("net_mg", 0.0)

# Conversion rates — same as dashboard.py pc() function
dem_cov = pct(sub, dem)
sub_l1  = pct(l1, sub)
l1_sel  = pct(sel, l1)
sel_ob  = pct(ob_hc, sel)
sp_yet  = max(0, round((sel or 0) - (ob_hc or 0) - (sp_hc or 0)))

# ─── ROW 1 KPI CARDS ──────────────────────────────────────────────────────────
st.markdown('<div class="sec">📊 Recruitment Overview</div>', unsafe_allow_html=True)   
cols = st.columns(6)

with cols[0]:
    st.markdown(kpi_card("red","📋","Demands", f"{dem:,}", "red",
        f"<strong>{dem_open:,}</strong> openings · <strong>{dem_u:,}</strong> roles are still waiting to be serviced",
        tag_r(f"⚠ {dem_u:,} pending") if dem_u > 0 else tag_g("All serviced")),
        unsafe_allow_html=True)

with cols[1]:
    st.markdown(kpi_card("grey","📤","Submissions", f"{sub:,}", "grey",
        f"<strong>{sub_fp:,}</strong> profiles are awaiting client feedback",
        tag_gr(f"{dem_cov}% demand coverage")),
        unsafe_allow_html=True)

with cols[2]:
    st.markdown(kpi_card("grey","🎙","Interviews", f"{ti:,}", "grey",
        f"L1 <strong>{l1:,}</strong> · L2 <strong>{l2:,}</strong> · L3 <strong>{l3:,}</strong>",
        tag_gr(f"{sub_l1}% sub→L1")),
        unsafe_allow_html=True)

with cols[3]:
    st.markdown(kpi_card("grey","✅","Selections", f"{sel:,}", "grey",
        "Confirmed selected candidates in the current period",
        tag_gr(f"{l1_sel}% L1→selected")),
        unsafe_allow_html=True)

with cols[4]:
    st.markdown(kpi_card("blue","📑","Selection Pipeline", f"{sp_hc:,}", "blue",
        f"<strong>₹{L(sp_po)}L</strong> PO value · <strong>₹{L(sp_mg)}L</strong> margin",
        tag_b(f"{sp_yet} Yet to be Onboarded")),
        unsafe_allow_html=True)

with cols[5]:
    active_po = grand.get("active_po", 0.0)
    active_mg = grand.get("active_mg", 0.0)
    st.markdown(kpi_card("blue","👥","Active Headcount", f"{active_hc:,}", "blue",
        f"<strong>{active_hc:,}</strong> active (filter-adjusted)<br>"
        f"<strong>₹{L(active_po)}L</strong> PO value · <strong>₹{L(active_mg)}L</strong> margin",
        tag_b("Selected-period onboarding excluded")),
        unsafe_allow_html=True)
    
# ─── ROW 2 KPI CARDS ──────────────────────────────────────────────────────────
st.markdown("<div style='margin-top:10px'></div>", unsafe_allow_html=True)
cols2 = st.columns(5)

with cols2[0]:
    st.markdown(kpi_card("green","🚀","Onboarded", f"{ob_hc:,}", "green",
        f"<strong>₹{L(ob_po)}L</strong> PO value · <strong>₹{L(ob_mg)}L</strong> margin",
        tag_g(f"{sel_ob}% sel→joined")),
        unsafe_allow_html=True)

with cols2[1]:
    st.markdown(kpi_card("red","🚪","Exits", f"{ex_hc:,}", "red",
        f"<strong>₹{L(ex_po)}L</strong> PO value · <strong>₹{L(ex_mg)}L</strong> margin<br>"
        f"Pipeline <strong>{ex_pipe_hc}</strong> HC · <strong>₹{L(ex_pipe_po)}L</strong> PO · <strong>₹{L(ex_pipe_mg)}L</strong> margin",
        tag_r(f"{ex_hc} headcount lost")),
        unsafe_allow_html=True)

with cols2[2]:
    net_hc_clr = "green" if net_hc >= 0 else "red"
    net_hc_str = f"{'+'if net_hc>=0 else '−'}{abs(net_hc):,}"
    st.markdown(kpi_card(net_hc_clr,"📈","Net HC", net_hc_str, net_hc_clr,
        f"<strong>₹{L(abs(net_po))}L</strong> net PO movement",
        tag_g("▲ Growth") if net_hc >= 0 else tag_r("▼ Decline")),
        unsafe_allow_html=True)

with cols2[3]:
    net_po_clr = "green" if net_po >= 0 else "red"
    net_po_str = f"{'+'if net_po>=0 else '−'}₹{L(abs(net_po))}L"
    st.markdown(kpi_card(net_po_clr,"💰","Net PO", net_po_str, net_po_clr,
        f"Ob <strong>₹{L(ob_po)}L</strong> – Ex <strong>₹{L(ex_po)}L</strong>",
        tag_g("Positive") if net_po >= 0 else tag_r("Negative")),
        unsafe_allow_html=True)

with cols2[4]:
    net_mg_clr = "green" if net_mg >= 0 else "red"
    net_mg_str = f"{'+'if net_mg>=0 else '−'}₹{L(abs(net_mg))}L"
    st.markdown(kpi_card(net_mg_clr,"🎯","Net Margin", net_mg_str, net_mg_clr,
        f"Ob <strong>₹{L(ob_mg)}L</strong> – Ex <strong>₹{L(ex_mg)}L</strong>",
        tag_g("Positive") if net_mg >= 0 else tag_r("Negative")),
        unsafe_allow_html=True)

# ─── COMPARISON BARS ──────────────────────────────────────────────────────────
if selected_years and selected_months:
    # 🔥 LMTD BAR (ADD FIRST)
    st.markdown(
        comp_bar(
            f"VS LMTD ({MON[lm_from.month]} {lm_from.year} till {lm_to.day})",
            grand,
            lmtd_grand
        ),
        unsafe_allow_html=True
    )

    # 🔥 EXISTING COMPARISON
    st.markdown(
        comp_bar("VS Previous Month", grand, prev_grand),
        unsafe_allow_html=True
    )
# ─── MRR BREAKDOWN ────────────────────────────────────────────────────────────
st.markdown('<div class="sec">MRR Breakdown</div>', unsafe_allow_html=True)
mrr1, mrr2, mrr3 = st.columns(3)

with mrr1:
    st.markdown(f"""<div class="kpi green" style="min-height:auto">
      <div class="kpi-lbl" style="color:#1db85a">🚀 Onboarding</div>
      <div class="mrr-grid" style="margin-top:8px">
        <div class="mrr-c"><div class="mrr-v" style="color:#1db85a">{ob_hc}</div><div class="mrr-l">HC</div></div>
        <div class="mrr-c"><div class="mrr-v" style="color:#17a050">₹{L(ob_po)}L</div><div class="mrr-l">PO Value</div></div>
        <div class="mrr-c"><div class="mrr-v" style="color:#17a050">₹{L(ob_mg)}L</div><div class="mrr-l">Margin</div></div>
      </div></div>""", unsafe_allow_html=True)

with mrr2:
    st.markdown(f"""<div class="kpi red" style="min-height:auto">
      <div class="kpi-lbl" style="color:#e8453c">🚪 Exits</div>
      <div class="mrr-grid" style="margin-top:8px">
        <div class="mrr-c"><div class="mrr-v" style="color:#e8453c">{ex_hc}</div><div class="mrr-l">HC</div></div>
        <div class="mrr-c"><div class="mrr-v" style="color:#e8453c">₹{L(ex_po)}L</div><div class="mrr-l">PO Value</div></div>
        <div class="mrr-c"><div class="mrr-v" style="color:#e8453c">₹{L(ex_mg)}L</div><div class="mrr-l">Margin</div></div>
      </div></div>""", unsafe_allow_html=True)

with mrr3:
    nc = "#1db85a" if net_hc >= 0 else "#e8453c"
    npc = "#1db85a" if net_po >= 0 else "#e8453c"
    nmc = "#1db85a" if net_mg >= 0 else "#e8453c"
    st.markdown(f"""<div class="kpi {net_hc_clr}" style="min-height:auto">
      <div class="kpi-lbl" style="color:{nc}">📊 Net MRR</div>
      <div class="mrr-grid" style="margin-top:8px">
        <div class="mrr-c"><div class="mrr-v" style="color:{nc}">{'+'if net_hc>=0 else '−'}{abs(net_hc)}</div><div class="mrr-l">Net HC</div></div>
        <div class="mrr-c"><div class="mrr-v" style="color:{npc}">{'+'if net_po>=0 else '−'}₹{L(abs(net_po))}L</div><div class="mrr-l">Net PO</div></div>
        <div class="mrr-c"><div class="mrr-v" style="color:{nmc}">{'+'if net_mg>=0 else '−'}₹{L(abs(net_mg))}L</div><div class="mrr-l">Net Margin</div></div>
      </div></div>""", unsafe_allow_html=True)

# ─── PIPELINE FUNNEL ──────────────────────────────────────────────────────────
st.markdown('<div class="sec">Recruitment Pipeline</div>', unsafe_allow_html=True)
with st.expander("Stage Snapshot & Volume Funnel", expanded=False):
    stages = [
        ("Demands",   dem,   "#e8453c"),
        ("Submitted", sub,   "#8892a4"),
        ("L1",        l1,    "#2ecc71"),
        ("L2",        l2,    "#27ae60"),
        ("L3",        l3,    "#1e8449"),
        ("Selected",  sel,   "#8892a4"),
        ("Onboarded", ob_hc, "#2ecc71"),
    ]

    # Stage snapshot using Streamlit columns (no HTML)
    stage_cols = []
    for i in range(len(stages) * 2 - 1):
        stage_cols.append(st.columns([1 if i % 2 == 0 else 0.15])[0] if len(stages) * 2 - 1 == 1 else None)

    # Use plotly for stage funnel instead of raw HTML
    fig_funnel = go.Figure(go.Funnel(
        y=[s[0] for s in stages],
        x=[s[1] for s in stages],
        textinfo="value+percent previous",
        marker=dict(color=[s[2] for s in stages]),
    ))
    fig_funnel.update_layout(height=300, margin=dict(t=10, b=10, l=10, r=10),
                              plot_bgcolor="white", paper_bgcolor="white")
    st.plotly_chart(fig_funnel,width="stretch", key="funnel_chart")

    # Volume bars using Streamlit progress bars
    st.markdown("**Volume Funnel**")
    mx = max(dem, sub, ti, sel, ob_hc, ex_hc, 1)
    fn_items = [
        ("Demands",       dem,   "#e8453c", None),
        ("Submissions",   sub,   "#8892a4", dem),
        ("L1 Interviews", l1,    "#2ecc71", sub),
        ("L2 Interviews", l2,    "#27ae60", l1),
        ("L3 Interviews", l3,    "#1e8449", l2),
        ("Selections",    sel,   "#8892a4", l1),
        ("Onboarded",     ob_hc, "#2ecc71", sel),
        ("Exits",         ex_hc, "#e8453c", None),
    ]
    for name, val, color, base in fn_items:
        c1, c2, c3, c4 = st.columns([2, 5, 1, 0.8])
        c1.caption(name)
        c2.progress(int(val / mx * 100) if mx > 0 else 0)
        c3.markdown(f"**{val:,}**")
        if base:
            c4.caption(f"{pct(val,base)}%")

# =========================================================
# 📋 RAW DATA EXPLORER (SINGLE FILE VERSION)
# =========================================================

# =========================================================
# 📋 RAW DATA EXPLORER (ADVANCED BI VERSION)
# =========================================================

st.markdown('<div class="sec">📋 Raw Data Explorer</div>', unsafe_allow_html=True)

with st.expander("View Raw Data (Filtered / Custom)", expanded=False):

    ss = st.session_state

    # -----------------------------
    # DEFAULT STATE
    # -----------------------------
    if "raw_dataset" not in ss:
        ss["raw_dataset"] = "All Data"

    # -----------------------------
    # DATASET BUTTONS
    # -----------------------------
    dataset_options = ["All Data", "Active HC", "Exited", "High Margin"]

    cols = st.columns(len(dataset_options))

    for col, option in zip(cols, dataset_options):
        with col:
            if st.button(option, key=f"raw_{option}", width="stretch"):
                ss["raw_dataset"] = option
                st.rerun()

    st.markdown("---")

    # -----------------------------
    # FILTERS
    # -----------------------------
    f1, f2, f3 = st.columns(3)

    with f1:
        raw_client = st.multiselect("Client", df["company_name"].dropna().unique())

    with f2:
        raw_from = st.date_input("From Date", value=None)

    with f3:
        raw_to = st.date_input("To Date", value=None)

    # -----------------------------
    # APPLY BASE FILTERS
    # -----------------------------
    raw_df = df.copy()

    if ss["raw_dataset"] == "Active HC":
        raw_df = raw_df[raw_df["Status"] == "Active"]

    elif ss["raw_dataset"] == "Exited":
        raw_df = raw_df[raw_df["Status"] == "Exit"]

    elif ss["raw_dataset"] == "High Margin":
        raw_df = raw_df[raw_df["margin"] > raw_df["margin"].median()]

    if raw_client:
        raw_df = raw_df[raw_df["company_name"].isin(raw_client)]

    if raw_from:
        raw_df = raw_df[raw_df["date"] >= pd.Timestamp(raw_from)]

    if raw_to:
        raw_df = raw_df[raw_df["date"] <= pd.Timestamp(raw_to)]

    # -----------------------------
    # 🔍 GLOBAL SEARCH
    # -----------------------------
    search_term = st.text_input("🔍 Search (any column)")

    if search_term:
        search_term = str(search_term).lower()
        raw_df = raw_df[
            raw_df.astype(str)
                  .apply(lambda row: row.str.lower().str.contains(search_term).any(), axis=1)
        ]

    # -----------------------------
    # 📊 COLUMN SELECTOR
    # -----------------------------
    all_cols = list(raw_df.columns)

    selected_cols = st.multiselect(
        "Select Columns",
        all_cols,
        default=all_cols[:10]  # show first 10 by default
    )

    if selected_cols:
        display_df = raw_df[selected_cols]
    else:
        display_df = raw_df

    # -----------------------------
    # 📥 DOWNLOAD BUTTON
    # -----------------------------
    def convert_to_excel(df):
        return df.to_csv(index=False).encode("utf-8")

    st.download_button(
        label="⬇ Download Data",
        data=convert_to_excel(display_df),
        file_name="raw_data.csv",
        mime="text/csv"
    )

    # -----------------------------
    # OUTPUT
    # -----------------------------
    st.caption(f"Rows: {len(display_df):,}")

    st.dataframe(
        display_df,
        width="stretch"
    )