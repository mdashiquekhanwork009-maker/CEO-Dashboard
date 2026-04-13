"""J2W  —  CEO Dashboard  v5
─────────────────────────────────
Run:   py -3.11 dashboard.py
Open:  http://localhost:5050

Requirements: pip install flask pandas openpyxl
"""
from functools import lru_cache
from io import BytesIO
import os
import re
import socket
import pandas as pd
from flask import Flask, Response, render_template_string, jsonify, request, send_from_directory
from plotly import data

app = Flask(__name__)

APP_ROOT = app.root_path
REPO_DATA_FOLDER = os.path.join(APP_ROOT, "data")
WINDOWS_LOGO_FOLDER = r"C:\Services Dashboard"
WINDOWS_DATA_FOLDER = r"C:\Data"

from datetime import datetime

LAST_REFRESHED_AT = datetime.now()

@app.route("/api/last_updated")
def last_updated():
    return jsonify({"time": LAST_REFRESHED_AT.strftime("%d %b %Y %H:%M:%S")})

@app.route('/logo.jpg')
def serve_logo():
    folder = os.path.dirname(LOGO_FILE)
    return send_from_directory(folder, os.path.basename(LOGO_FILE))

# ─────────────────────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────────────────────

FILE_PATHS = {
    "demand":   "demand_data.csv",
    "sub":      "submission.csv",
    "intv":     "interview.csv",
    "sel":      "selection.csv",
    "selpipe":  "selection_pipeline.csv",
    "ob":       "onboarding_data.csv",
    "activehc": "activeheadcount.csv",
    "exit":     "exit_data.csv",
    "exitpipe": "exit_pipeline_data.csv",
}


def _existing_path(candidates, must_exist=True):
    for candidate in candidates:
        if candidate and (not must_exist or os.path.exists(candidate)):
            return candidate
    return candidates[-1] if candidates else ""


def resolve_logo_file():
    return _existing_path([
        os.path.join(APP_ROOT, "logo.jpg"),
        os.path.join(REPO_DATA_FOLDER, "logo.jpg"),
        os.path.join(WINDOWS_LOGO_FOLDER, "logo.jpg"),
    ])


def resolve_data_folder():
    candidate_folders = [
        os.environ.get("DASHBOARD_DATA_DIR", "").strip(),
        REPO_DATA_FOLDER,
        APP_ROOT,
        WINDOWS_DATA_FOLDER,
    ]
    best_folder = APP_ROOT
    best_matches = -1
    for folder in candidate_folders:
        if not folder or not os.path.isdir(folder):
            continue
        matches = sum(
            os.path.exists(os.path.join(folder, filename))
            for filename in FILE_PATHS.values()
        )
        if matches > best_matches:
            best_folder = folder
            best_matches = matches
    return best_folder


def resolve_mapping_file(data_folder):
    return _existing_path([
        os.environ.get("DASHBOARD_MAPPING_FILE", "").strip(),
        os.path.join(APP_ROOT, "Org mapping.xlsx"),
        os.path.join(REPO_DATA_FOLDER, "Org mapping.xlsx"),
        os.path.join(data_folder, "Org mapping.xlsx"),
    ])


DATA_FOLDER = resolve_data_folder()
MAPPING_FILE = resolve_mapping_file(DATA_FOLDER)
LOGO_FILE = resolve_logo_file()

DATE_COLS = {
    "demand":   "Created_at",
    "sub":      "date",
    "intv":     "Interview_date",
    "sel":      "selection_date",
    "selpipe":  "display_date",
    "ob":       "display_date",
    "activehc": "display_date",
    "exit":     "last_work_day",
    "exitpipe": "tentative_exit_date",
}

CLIENT_COLS = {
    "demand":   "Company_name",
    "sub":      "client",
    "intv":     "company_name",
    "sel":      "company_name",
    "selpipe":  "company_name",
    "ob":       "company_name",
    "activehc": "company_name",
    "exit":     "company_name",
    "exitpipe": "company_name",
}

ID_COL_CANDIDATES = ["id", "ID", "job_id", "Job_ID", "job_ID"]
OPENING_COL_CANDIDATES = [
    "no.of opening", "No.of opening", "No.Of Opening",
    "no_of_openings", "No_of_openings", "no_of_opening", "No_of_opening",
    "openings", "Openings", "opening_count", "Opening_count",
    "number_of_openings", "Number_of_openings",
    "no_of_positions", "No_of_positions", "positions", "Positions",
    "vacancies", "Vacancies",
]
PO_COL_KEYS = {"selpipe", "ob", "exit", "exitpipe"}
RAW_DATASET_CONFIG = {
    "demand": {"label": "Demands"},
    "sub": {"label": "Submissions"},
    "intv": {"label": "Interviews"},
    "sel": {"label": "Selections"},
    "selpipe": {"label": "Selection Pipeline"},
    "ob": {"label": "Onboardings"},
    "exitpipe": {"label": "Exit Pipeline"},
    "exit": {"label": "Exit"},
}


def _prepare_frame(df, file_key):
    if df.empty:
        df = df.copy()
        df["_date"] = pd.NaT
        df["_year"] = pd.Series(dtype="Int64")
        df["_month"] = pd.Series(dtype="Int64")
        df["_po_end_date"] = pd.NaT
        if file_key in PO_COL_KEYS:
            df["_po"] = pd.Series(dtype="float64")
            df["_mg"] = pd.Series(dtype="float64")
        return df

    df = df.copy()
    df.columns = df.columns.str.strip()

    client_col = CLIENT_COLS.get(file_key)
    if client_col and client_col in df.columns:
        df[client_col] = df[client_col].str.strip()

    date_col = DATE_COLS.get(file_key)
    if file_key == "selpipe":
        parsed = pd.Series(pd.NaT, index=df.index)
        # Selection pipeline metrics should follow the display date for dashboard filters.
        for candidate in ["display_date", "offer_created_date", "Selection_date"]:
            if candidate in df.columns:
                candidate_parsed = pd.to_datetime(df[candidate], errors="coerce")
                if getattr(candidate_parsed.dt, "tz", None) is not None:
                    candidate_parsed = candidate_parsed.dt.tz_localize(None)
                parsed = parsed.where(parsed.notna(), candidate_parsed)
    elif date_col and date_col in df.columns:
        parsed = pd.to_datetime(df[date_col], errors="coerce")
        if getattr(parsed.dt, "tz", None) is not None:
            parsed = parsed.dt.tz_localize(None)
    else:
        parsed = None

    if parsed is not None:
        df["_date"] = parsed
        df["_year"] = parsed.dt.year.astype("Int64")
        df["_month"] = parsed.dt.month.astype("Int64")
    else:
        df["_date"] = pd.NaT
        df["_year"] = pd.Series(pd.NA, index=df.index, dtype="Int64")
        df["_month"] = pd.Series(pd.NA, index=df.index, dtype="Int64")

    if "po_end_date" in df.columns:
        po_end_parsed = pd.to_datetime(df["po_end_date"], errors="coerce")
        if getattr(po_end_parsed.dt, "tz", None) is not None:
            po_end_parsed = po_end_parsed.dt.tz_localize(None)
        df["_po_end_date"] = po_end_parsed
    else:
        df["_po_end_date"] = pd.NaT

    if file_key in PO_COL_KEYS:
        df["_po"] = _flt(df["p_o_value"]) if "p_o_value" in df.columns else pd.Series(0.0, index=df.index)
        df["_mg"] = _flt(df["margin"]) if "margin" in df.columns else pd.Series(0.0, index=df.index)

    if file_key == "demand":
        opening_col = next((c for c in OPENING_COL_CANDIDATES if c in df.columns), None)
        if opening_col:
            df["_openings"] = _flt(df[opening_col])
        else:
            df["_openings"] = pd.Series(1.0, index=df.index)

    if file_key == "sub" and "status" in df.columns:
        df["_st"] = df["status"].str.strip()
        df["_stl"] = df["_st"].str.lower()

    if file_key == "intv" and "workflow_step" in df.columns:
        df["_ws"] = df["workflow_step"].str.lower()

    id_col = next((c for c in ID_COL_CANDIDATES if c in df.columns), None)
    if id_col:
        df["_id_norm"] = df[id_col].astype(str).str.strip()

    return df

# ─────────────────────────────────────────────────────────────────────────────
# DATA LOADING
# ─────────────────────────────────────────────────────────────────────────────

def load_data():
    data = {}
    for key, filename in FILE_PATHS.items():
        path = os.path.join(DATA_FOLDER, filename)
        if os.path.exists(path):
            try:
                if filename.endswith(".xlsx"):
                    df = pd.read_excel(path, dtype=str).fillna("")
                else:
                    df = pd.read_csv(path, dtype=str).fillna("")
                data[key] = _prepare_frame(df, key)
            except Exception as e:
                print(f"  x {filename}: {e}")
                data[key] = pd.DataFrame()
        else:
            data[key] = pd.DataFrame()
    return data

@lru_cache(maxsize=1)
def load_data_cached():
    return load_data()

# ─────────────────────────────────────────────────────────────────────────────
# CLIENT MAPPING
# ─────────────────────────────────────────────────────────────────────────────

@lru_cache(maxsize=1)
def load_mapping():
    client_to_domain = {}
    client_to_bh = {}

    if os.path.exists(MAPPING_FILE):
        try:
            if MAPPING_FILE.endswith(".xlsx"):
                df = pd.read_excel(MAPPING_FILE, dtype=str).fillna("")
            else:
                df = pd.read_csv(MAPPING_FILE, dtype=str).fillna("")

            # Normalize column names
            df.columns = df.columns.str.strip().str.lower().str.replace(" ", "_", regex=False)

            # Flexible column detection
            cl_col  = next((c for c in df.columns if c in ["client","clients","client_name","company","company_name"]), None)
            dom_col = next((c for c in df.columns if "domain" in c or "category" in c), None)
            bh_col  = next((c for c in df.columns if "business_head" in c or c == "bh" or "head" in c), None)

            print(f"[DEBUG] Detected -> client_col={cl_col}  domain_col={dom_col}  bh_col={bh_col}")

            if cl_col is None:
                print("[DEBUG] WARNING: Could not find client column in mapping file!")
            else:
                clients = df[cl_col].astype(str).str.strip()
                valid = clients.ne("") & clients.str.lower().ne("nan")

                if dom_col:
                    domains = df[dom_col].astype(str).map(normalize_domain_label)
                    domain_valid = valid & domains.ne("") & domains.str.lower().ne("nan")
                    client_to_domain = dict(zip(clients[domain_valid], domains[domain_valid]))

                if bh_col:
                    bhs = df[bh_col].astype(str).map(normalize_bh_label)
                    bh_valid = valid & bhs.ne("") & bhs.str.lower().ne("nan")
                    client_to_bh = dict(zip(clients[bh_valid], bhs[bh_valid]))

        except Exception as e:
            print("Mapping error:", e)
            import traceback; traceback.print_exc()

    domains        = sorted(set(client_to_domain.values()), key=str.lower)
    business_heads = sorted(set(client_to_bh.values()),     key=str.lower)

    return client_to_domain, client_to_bh, domains, business_heads


def normalize_domain_label(value):
    text = str(value or "").strip()
    key = re.sub(r"[\s_]+", " ", text).strip().lower()
    if key in {"captive", "sow captive"}:
        return "Captive"
    return text.title()


def normalize_bh_label(value):
    text = str(value or "").strip()
    if not text:
        return ""
    return re.sub(r"\s+non[\s-]*it\s*$", "", text, flags=re.IGNORECASE).strip()


def normalize_client_key(value):
    value = str(value or "").strip().lower()
    return re.sub(r"[^a-z0-9]+", "", value)


PO_CAPTIVE_CLIENTS = {normalize_client_key("Flipkart")}
PO_CAPTIVE_THRESHOLD = 100000
CAPTIVE_SUFFIX = " [Captive]"


def is_po_captive_client(client_name):
    return normalize_client_key(client_name) in PO_CAPTIVE_CLIENTS


def build_client_lookup(*mapping_dicts):
    lookup = {}
    for mapping in mapping_dicts:
        for client in mapping.keys():
            norm = normalize_client_key(client)
            if norm and norm not in lookup:
                lookup[norm] = client
    return lookup


def get_mapped_client_name(client_name, client_lookup):
    return client_lookup.get(normalize_client_key(client_name), client_name)


def parse_csv_filter_arg(arg_name):
    raw = request.args.get(arg_name, default="")
    return set(v.strip() for v in raw.split(",") if v.strip()) if raw else None


def parse_int_csv_filter_arg(arg_name):
    values = parse_csv_filter_arg(arg_name)
    if not values:
        return None
    parsed = set()
    for value in values:
        try:
            parsed.add(int(value))
        except (TypeError, ValueError):
            continue
    return parsed or None


def freeze_filter(values):
    if not values:
        return None
    return tuple(sorted(values))


def thaw_filter(values):
    return set(values) if values else None


def freeze_date(value):
    if value is None or pd.isna(value):
        return None
    ts = pd.Timestamp(value)
    if ts.tzinfo is not None:
        ts = ts.tz_localize(None)
    return ts.isoformat()


def thaw_date(value):
    return pd.Timestamp(value) if value else None


def get_resolved_filters():
    cl_filter = parse_csv_filter_arg("clients")
    dom_filter = parse_csv_filter_arg("domains")
    bh_filter = parse_csv_filter_arg("bhs")
    resolved = resolve_client_filter_cached(
        freeze_filter(cl_filter),
        freeze_filter(dom_filter),
        freeze_filter(bh_filter),
    )
    return thaw_filter(resolved), dom_filter, bh_filter


def get_period_filters():
    years = parse_int_csv_filter_arg("years")
    months = parse_int_csv_filter_arg("months")
    if years is None:
        single_year = request.args.get("year", type=int)
        years = {single_year} if single_year else None
    if months is None:
        single_month = request.args.get("month", type=int)
        months = {single_month} if single_month else None
    return years, months


def get_mapping_context():
    client_to_domain, client_to_bh, domains, business_heads = load_mapping()
    client_to_domain = dict(client_to_domain)
    client_to_bh = dict(client_to_bh)
    client_lookup = build_client_lookup(client_to_domain, client_to_bh)
    return client_to_domain, client_to_bh, client_lookup, domains, business_heads


# ─────────────────────────────────────────────────────────────────────────────
# DATE HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def add_ym(df, file_key):
    if "_year" not in df.columns or "_month" not in df.columns:
        return _prepare_frame(df, file_key)
    return df


def filter_ym(df, sel_year, sel_month):
    if sel_year and "_year" in df.columns:
        year_values = sel_year if isinstance(sel_year, (set, tuple, list)) else {sel_year}
        df = df[df["_year"].isin(year_values)]
    if sel_month and "_month" in df.columns:
        month_values = sel_month if isinstance(sel_month, (set, tuple, list)) else {sel_month}
        df = df[df["_month"].isin(month_values)]
    return df


def filter_date_range(df, from_date=None, to_date=None):
    if df.empty or "_date" not in df.columns:
        return df

    if from_date is not None:
        from_ts = pd.Timestamp(from_date)
        if from_ts.tzinfo is not None:
            from_ts = from_ts.tz_localize(None)
        df = df[df["_date"] >= from_ts]

    if to_date is not None:
        to_ts = pd.Timestamp(to_date)
        if to_ts.tzinfo is not None:
            to_ts = to_ts.tz_localize(None)
        df = df[df["_date"] < (to_ts + pd.Timedelta(days=1))]

    return df


def get_raw_period_filters():
    year = request.args.get("year", type=int)
    month = request.args.get("month", type=int)
    years = {year} if year else None
    months = {month} if month else None
    return years, months


def filter_clients(df, file_key, client_filter=None):
    if df.empty or not client_filter:
        return df
    client_col = CLIENT_COLS.get(file_key)
    if client_col and client_col in df.columns:
        return df[df[client_col].isin(client_filter)]
    return df

def filter_unserviced_demands(df):
    if df.empty:
        return df
    sub_df = add_ym(load_data_cached().get("sub", pd.DataFrame()), "sub")
    submitted_ids = set()
    if not sub_df.empty:
        if "_id_norm" in sub_df.columns:
            submitted_ids = set(sub_df["_id_norm"].dropna().astype(str).str.strip())
        else:
            id_col_sub = next((c for c in ID_COL_CANDIDATES if c in sub_df.columns), None)
            if id_col_sub:
                submitted_ids = set(sub_df[id_col_sub].dropna().astype(str).str.strip())

    id_col_dem = "_id_norm" if "_id_norm" in df.columns else next((c for c in ID_COL_CANDIDATES if c in df.columns), None)
    if not id_col_dem:
        return df
    demand_ids = df[id_col_dem].dropna().astype(str).str.strip()
    unsubmitted = df.loc[~demand_ids.isin(submitted_ids)]

    # Only treat as unserviced if id_status == "0" (when column exists)
    if "id_status" in unsubmitted.columns:
        unsubmitted = unsubmitted[
            unsubmitted["id_status"].astype(str).str.strip() == "0"
        ]
    return unsubmitted


def get_raw_dataset_frame(dataset_key, year_filter=None, month_filter=None, client_filter=None, from_date=None, to_date=None, demand_status="all"):
    data = load_data_cached()
    df = add_ym(data.get(dataset_key, pd.DataFrame()), dataset_key).copy()
    if df.empty:
        return df
    df = filter_ym(df, year_filter, month_filter)
    df = filter_date_range(df, from_date, to_date)
    df = filter_clients(df, dataset_key, client_filter)
    if dataset_key == "demand":
        if demand_status == "unserviced":
            df = filter_unserviced_demands(df)
        elif demand_status == "serviced":
            unserviced_df = filter_unserviced_demands(df)
            if unserviced_df.empty:
                df = df.iloc[0:0] if df.empty else df
            else:
                remaining_index = df.index.difference(unserviced_df.index)
                df = df.loc[remaining_index]
    if "_date" in df.columns:
        df = df.sort_values(by="_date", ascending=False, na_position="last")
    return df


def serialize_raw_dataset(df):
    if df.empty:
        return {"columns": [], "rows": [], "row_count": 0}

    visible_columns = [col for col in df.columns if not col.startswith("_")]
    safe_df = df[visible_columns].copy()
    safe_df = safe_df.fillna("")
    rows = []
    for row in safe_df.to_dict(orient="records"):
        clean_row = {}
        for key, value in row.items():
            if pd.isna(value):
                clean_row[key] = ""
            else:
                clean_row[key] = str(value)
        rows.append(clean_row)
    return {"columns": visible_columns, "rows": rows, "row_count": len(rows)}


def get_raw_request_context():
    dataset = request.args.get("dataset", default="demand", type=str)
    year_filter, month_filter = get_raw_period_filters()
    client_filter = parse_csv_filter_arg("raw_clients")
    demand_status = request.args.get("demand_status", default="", type=str).strip().lower()
    if demand_status not in {"all", "unserviced", "serviced"}:
        legacy_unserviced_only = request.args.get("unserviced_only", default="0", type=str) in {"1", "true", "True", "yes", "on"}
        demand_status = "unserviced" if legacy_unserviced_only else "all"
    from_str = request.args.get("from")
    to_str = request.args.get("to")
    from_date = pd.to_datetime(from_str, errors="coerce") if from_str else None
    to_date = pd.to_datetime(to_str, errors="coerce") if to_str else None
    if from_date is not None and pd.isna(from_date):
        from_date = None
    if to_date is not None and pd.isna(to_date):
        to_date = None
    return dataset, year_filter, month_filter, client_filter, from_date, to_date, demand_status


def get_export_columns(df, column_mode="visible"):
    if column_mode == "all":
        return list(df.columns)
    return [col for col in df.columns if not col.startswith("_")]


def sanitize_filename_part(value):
    clean = re.sub(r"[^A-Za-z0-9._-]+", "_", str(value).strip())
    clean = re.sub(r"_+", "_", clean).strip("._-")
    return clean or "all"


def build_raw_export_filename(dataset, client_filter=None, year_filter=None, month_filter=None, from_date=None, to_date=None, column_mode="visible", file_ext="csv", demand_status="all"):
    parts = [sanitize_filename_part(RAW_DATASET_CONFIG.get(dataset, {}).get("label", dataset).lower())]
    if dataset == "demand" and demand_status in {"unserviced", "serviced"}:
        parts.append(demand_status)

    if year_filter and month_filter and len(year_filter) == 1 and len(month_filter) == 1:
        year = next(iter(year_filter))
        month = next(iter(month_filter))
        parts.append(f"{year}-{int(month):02d}")
    elif from_date is not None or to_date is not None:
        if from_date is not None:
            parts.append(f"from_{pd.Timestamp(from_date).strftime('%Y-%m-%d')}")
        if to_date is not None:
            parts.append(f"to_{pd.Timestamp(to_date).strftime('%Y-%m-%d')}")
    else:
        parts.append("all_periods")

    if client_filter:
        sorted_clients = sorted(client_filter, key=str.lower)
        if len(sorted_clients) == 1:
            parts.append(sanitize_filename_part(sorted_clients[0]))
        elif len(sorted_clients) <= 3:
            parts.append(sanitize_filename_part("_".join(sorted_clients)))
        else:
            parts.append(f"{len(sorted_clients)}_clients")
    else:
        parts.append("all_clients")

    parts.append("all_columns" if column_mode == "all" else "visible_data")
    return f"{'_'.join(parts)}.{file_ext}"


def _flt(series):
    return pd.to_numeric(
        series.astype(str).str.replace(",", "", regex=False),
        errors="coerce"
    ).fillna(0.0)


def collect_periods(data):
    ym = set()
    for df in data.values():
        if df.empty or "_year" not in df.columns or "_month" not in df.columns:
            continue
        pairs = (
            df[["_year", "_month"]]
            .dropna()
            .drop_duplicates()
            .itertuples(index=False, name=None)
        )
        ym.update((int(year), int(month)) for year, month in pairs)
    return sorted(ym, reverse=True)


def collect_all_clients(data):
    clients = set()
    for key, df in data.items():
        col = CLIENT_COLS.get(key)
        if col and col in df.columns:
            vals = df[col]
            clients.update(vals[vals != ""].unique().tolist())
    return sorted(clients, key=str.lower)


@lru_cache(maxsize=1)
def get_periods_cached():
    return tuple(collect_periods(load_data_cached()))


@lru_cache(maxsize=1)
def get_client_catalog():
    data = load_data_cached()
    clients = collect_all_clients(data)
    client_to_domain, client_to_bh, _, _ = load_mapping()
    client_lookup = build_client_lookup(client_to_domain, client_to_bh)

    catalog = []
    for client in clients:
        mapped = get_mapped_client_name(client, client_lookup)
        catalog.append({
            "name": client,
            "domain": client_to_domain.get(mapped, ""),
            "bh": normalize_bh_label(client_to_bh.get(mapped, "")),
        })
    return tuple(catalog)


def get_shareable_urls(port):
    ips = set()

    try:
        hostname = socket.gethostname()
        for info in socket.getaddrinfo(hostname, None, family=socket.AF_INET):
            ip = info[4][0]
            if ip.startswith(("10.", "172.", "192.168.")):
                ips.add(ip)
    except OSError:
        pass

    try:
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as sock:
            sock.connect(("8.8.8.8", 80))
            ip = sock.getsockname()[0]
            if ip.startswith(("10.", "172.", "192.168.")):
                ips.add(ip)
    except OSError:
        pass

    urls = [f"http://{ip}:{port}" for ip in sorted(ips)]
    if not urls:
        urls.append(f"http://localhost:{port}")
    return urls


# ─────────────────────────────────────────────────────────────────────────────
# METRICS ENGINE
# ─────────────────────────────────────────────────────────────────────────────

ZERO = dict(
    dem=0, dem_open=0, dem_u=0, serviced_dem=0,
    sub=0, sub_fp=0,
    l1=0, l1_fp=0, l2=0, l2_fp=0, l3=0, l3_fp=0,
    sel=0, sel_pure=0, sp_hc=0, sp_po=0.0, sp_mg=0.0,
    ob_hc=0, ob_po=0.0, ob_mg=0.0,
    active_hc=0, active_po=0.0, active_mg=0.0,
    ex_hc=0, ex_po=0.0, ex_mg=0.0,
    ex_pipe_hc=0, ex_pipe_po=0.0, ex_pipe_mg=0.0,
    net_hc=0, net_po=0.0, net_mg=0.0,
)


def compute_all(data, sel_year, sel_month, client_filter=None, from_date=None, to_date=None):
    frames = {
        k: filter_date_range(filter_ym(add_ym(df, k), sel_year, sel_month), from_date, to_date)
        for k, df in data.items()
    }
    res = {}

    def ensure(name):
        if name not in res:
            res[name] = dict(ZERO)

    def add_po_metrics(grouped_df, client_name, hc_key, po_key, mg_key):
        """
        Keep captive PO splitting consistent across onboarding and exits.
        """
        ensure(client_name)
        if is_po_captive_client(client_name):
            captive_key = f"{client_name}{CAPTIVE_SUFFIX}"
            high = grouped_df[grouped_df["_po"] > PO_CAPTIVE_THRESHOLD]
            normal = grouped_df[grouped_df["_po"] <= PO_CAPTIVE_THRESHOLD]
            if not high.empty:
                ensure(captive_key)
                res[captive_key][hc_key] += len(high)
                res[captive_key][po_key] += float(high["_po"].sum())
                res[captive_key][mg_key] += float(high["_mg"].sum())
            if not normal.empty:
                res[client_name][hc_key] += len(normal)
                res[client_name][po_key] += float(normal["_po"].sum())
                res[client_name][mg_key] += float(normal["_mg"].sum())
            return

        res[client_name][hc_key] += len(grouped_df)
        res[client_name][po_key] += float(grouped_df["_po"].sum())
        res[client_name][mg_key] += float(grouped_df["_mg"].sum())

    # DEMAND
    df     = frames["demand"]
    sub_df = frames["sub"]

    submitted_ids = set()
    if not sub_df.empty:
        if "_id_norm" in sub_df.columns:
            submitted_ids = set(sub_df["_id_norm"].unique())
        else:
            id_col_sub = next((c for c in ID_COL_CANDIDATES if c in sub_df.columns), None)
            if id_col_sub:
                submitted_ids = set(sub_df[id_col_sub].astype(str).str.strip().unique())

    cl_col = None
    if not df.empty:
        cl_col = next((c for c in ["Company_name", "company_name", "client", "Client"] if c in df.columns), None)

    if not df.empty and cl_col:
        if client_filter is not None:
            df = df[df[cl_col].isin(client_filter)]
        unserviced_counts = {}

        if "_id_norm" in df.columns:
            unserviced_df = df.loc[~df["_id_norm"].isin(submitted_ids)]

            if "id_status" in unserviced_df.columns:
                unserviced_df = unserviced_df[
                    unserviced_df["id_status"].astype(str).str.strip() == "0"
                ]

            unserviced_counts = unserviced_df.groupby(cl_col).size().to_dict()
            unserviced_counts = unserviced_df.groupby(cl_col).size().to_dict()
        for cl, g in df.groupby(cl_col):
            ensure(cl)

            total_dem = len(g)
            unserv = int(unserviced_counts.get(cl, 0))

            res[cl]["dem"] += total_dem
            res[cl]["dem_open"] += float(g["_openings"].sum()) if "_openings" in g.columns else total_dem
            res[cl]["dem_u"] += unserv

    # ✅ ADD THIS
            res[cl]["serviced_dem"] += (total_dem - unserv)

    # SUBMISSION
    df = frames["sub"]
    cl_col = next((c for c in ["client", "Client"] if c in df.columns), None)
    if not df.empty and cl_col and "_st" in df.columns and "_stl" in df.columns:
        if client_filter is not None:
            df = df[df[cl_col].isin(client_filter)]
        for cl, g in df.groupby(cl_col):
            ensure(cl)
            res[cl]["sub"]    += len(g)
            res[cl]["sub_fp"] += int((g["_st"] == "Client Submit").sum())
            res[cl]["l1_fp"]  += int(g["_stl"].str.contains("schedule l1 interview", na=False).sum())
            res[cl]["l2_fp"]  += int(g["_stl"].str.contains("schedule l2 interview", na=False).sum())
            res[cl]["l3_fp"]  += int(g["_stl"].str.contains("schedule l3 interview", na=False).sum())

    # INTERVIEW — includes Schedule + Reschedule
    df = frames["intv"]
    cl_col = next((c for c in ["Company_name", "company_name", "client", "Client"] if c in df.columns), None)
    if not df.empty and cl_col and "_ws" in df.columns:
        if client_filter is not None:
            df = df[df[cl_col].isin(client_filter)]
        for cl, g in df.groupby(cl_col):
            ensure(cl)
            res[cl]["l1"] += int(g["_ws"].str.contains("l1 interview", na=False).sum())
            res[cl]["l2"] += int(g["_ws"].str.contains("l2 interview", na=False).sum())
            res[cl]["l3"] += int(g["_ws"].str.contains("l3 interview", na=False).sum())

    # SELECTION
    df = frames["sel"]
    cl_col = next((c for c in ["Company_name", "company_name", "client", "Client"] if c in df.columns), None)
    if not df.empty and cl_col:
        if client_filter is not None:
            df = df[df[cl_col].isin(client_filter)]
        for cl, g in df.groupby(cl_col):
            ensure(cl)
            res[cl]["sel"]      += len(g)
            res[cl]["sel_pure"] += len(g)  # same as sel — kept separate for future use

    # SELECTION PIPELINE
    df = frames["selpipe"]
    cl_col = None
    if not df.empty:
        cl_col = next((c for c in ["Company_name", "company_name", "client", "Client"] if c in df.columns), None)
    if not df.empty and cl_col:
        if client_filter is not None:
            df = df[df[cl_col].isin(client_filter)]
        for cl, g in df.groupby(cl_col):
            ensure(cl)
            res[cl]["sp_hc"] += len(g)
            res[cl]["sp_po"] += float(g["_po"].sum())
            res[cl]["sp_mg"] += float(g["_mg"].sum())

    # ONBOARDING
    # Rows where p_o_value > 100000 for Flipkart are counted under a
    # separate "Flipkart [Captive]" key so they appear under Captive domain.
    df = frames["ob"]
    cl_col = None
    if not df.empty:
        cl_col = next((c for c in ["Company_name", "company_name", "client", "Client"] if c in df.columns), None)
    if not df.empty and cl_col:
        if client_filter is not None:
            df = df[df[cl_col].isin(client_filter)]
        for cl, g in df.groupby(cl_col):
            add_po_metrics(g, cl, "ob_hc", "ob_po", "ob_mg")

    # ACTIVE HEADCOUNT
    # Count all candidate rows present in the active headcount sheet per client.
    
    df = frames["activehc"]
    cl_col = None
    if not df.empty:
        cl_col = next((c for c in ["company_name", "Company_name", "client", "Client"] if c in df.columns), None)

    if not df.empty and cl_col:
        if client_filter is not None:
            df = df[df[cl_col].isin(client_filter)]

    # 🔥 Ensure PO & Margin columns exist
        def clean_numeric(series):
          return pd.to_numeric(
              series.astype(str)
              .str.replace(",", "", regex=False)
              .str.replace("₹", "", regex=False)
              .str.strip(),
              errors="coerce"
          ).fillna(0)

# 🔥 FORCE overwrite (IMPORTANT)
        if "p_o_value" in df.columns:
            df["_po"] = clean_numeric(df["p_o_value"])

        if "margin" in df.columns:
            df["_mg"] = clean_numeric(df["margin"])

        for cl, g in df.groupby(cl_col):
            ensure(cl)
            res[cl]["active_hc"] += len(g)
            res[cl]["active_po"] += float(g["_po"].sum())
            res[cl]["active_mg"] += float(g["_mg"].sum())

    # For filtered views, show opening active HC for the selected period by
    

    # EXIT
    df = frames["exit"]
    cl_col = None
    if not df.empty:
        cl_col = next((c for c in ["Company_name", "company_name", "client", "Client"] if c in df.columns), None)
    if not df.empty and cl_col:
        if client_filter is not None:
            df = df[df[cl_col].isin(client_filter)]
        for cl, g in df.groupby(cl_col):
            add_po_metrics(g, cl, "ex_hc", "ex_po", "ex_mg")

    # EXIT PIPELINE
    df = frames["exitpipe"]
    cl_col = None
    if not df.empty:
        cl_col = next((c for c in ["Company_name", "company_name", "client", "Client"] if c in df.columns), None)
    if not df.empty and cl_col:
        if client_filter is not None:
            df = df[df[cl_col].isin(client_filter)]
        for cl, g in df.groupby(cl_col):
            add_po_metrics(g, cl, "ex_pipe_hc", "ex_pipe_po", "ex_pipe_mg")

    for cl in res:
        m = res[cl]

        for k in [
            "sp_po", "sp_mg", "ob_po", "ob_mg",
            "ex_po", "ex_mg", "ex_pipe_po", "ex_pipe_mg",
            "active_po", "active_mg"
        ]:
            value = m[k] / 1e5
            m[k] = value if value >= 0.01 else (m[k] / 1e3)

        m["net_hc"] = m["ob_hc"] - m["ex_hc"]
        m["net_po"] = m["ob_po"] - m["ex_po"]
        m["net_mg"] = m["ob_mg"] - m["ex_mg"]
    return res


@lru_cache(maxsize=256)
def compute_all_cached(sel_year, sel_month, client_filter_key=None, from_date_key=None, to_date_key=None):
    return compute_all(
        load_data_cached(),
        sel_year,
        sel_month,
        thaw_filter(client_filter_key),
        thaw_date(from_date_key),
        thaw_date(to_date_key),
    )


def grand_total(res):
    g = dict(ZERO)
    for m in res.values():
        for k in g:
            g[k] += m.get(k, 0)
    g["net_hc"] = g["ob_hc"] - g["ex_hc"]
    g["net_po"] = g["ob_po"] - g["ex_po"]
    g["net_mg"] = g["ob_mg"] - g["ex_mg"]
    g["sel_pure"] = g["sel"]  # grand total pure sel = grand total sel
    # Keep legacy key aligned to selection pipeline headcount for the KPI card.
    g["sel_not_ob"] = g["sp_hc"]
    return g


def round_m(m):
    return {k: round(v, 4) if isinstance(v, float) else v for k, v in m.items()}


# ─────────────────────────────────────────────────────────────────────────────
# CLIENT FILTER RESOLVER
# ─────────────────────────────────────────────────────────────────────────────

def resolve_client_filter(cl_filter, dom_filter, bh_filter):
    """Resolve domain/BH filters into a client set with normalized client-name matching."""
    if not dom_filter and not bh_filter:
        return cl_filter

    bh_filter = {normalize_bh_label(v) for v in bh_filter} if bh_filter else bh_filter
    allowed = set()
    for entry in get_client_catalog():
        actual_cl = entry["name"]
        dom_ok = (not dom_filter) or (entry["domain"] in dom_filter)
        if not dom_ok and dom_filter and "Captive" in dom_filter and is_po_captive_client(actual_cl):
            dom_ok = True
        bh_ok = (not bh_filter) or (entry["bh"] in bh_filter)
        if dom_ok and bh_ok:
            allowed.add(actual_cl)

    print(f"[DEBUG] resolve_client_filter: dom={dom_filter} bh={bh_filter} -> {len(allowed)} clients matched")
    return (cl_filter & allowed) if cl_filter else allowed


@lru_cache(maxsize=256)
def resolve_client_filter_cached(cl_filter_key=None, dom_filter_key=None, bh_filter_key=None):
    return freeze_filter(
        resolve_client_filter(
            thaw_filter(cl_filter_key),
            thaw_filter(dom_filter_key),
            thaw_filter(bh_filter_key),
        )
    )


# ─────────────────────────────────────────────────────────────────────────────
# MONTH-ON-MONTH TRENDS
# ─────────────────────────────────────────────────────────────────────────────

MON_NAMES_PY = ["","Jan","Feb","Mar","Apr","May","Jun",
                "Jul","Aug","Sep","Oct","Nov","Dec"]

def mom_trends(data, client_filter=None):
    def include(name):
        return client_filter is None or name in client_filter

    trends = {"dem":{}, "sub":{}, "sub_fp":{}, "l1":{}, "l2":{}, "l3":{}}

    def period_str(year, month):
        return f"{MON_NAMES_PY[month]}'{str(year)[2:]}"

    df = add_ym(data.get("demand", pd.DataFrame()), "demand")
    if not df.empty and "Company_name" in df.columns:
        df = df[df["Company_name"].apply(include)]
        for (yr, mo), g in df.groupby(["_year","_month"]):
            if pd.isna(yr) or pd.isna(mo): continue
            p = period_str(int(yr), int(mo))
            trends["dem"][p] = trends["dem"].get(p, 0) + len(g)

    df = add_ym(data.get("sub", pd.DataFrame()), "sub")
    cl_col = next((c for c in ["client","Client"] if c in df.columns), None)
    if not df.empty and cl_col and "status" in df.columns:
        df = df[df[cl_col].apply(include)].copy()
        df["_st"] = df["status"].str.strip()
        for (yr, mo), g in df.groupby(["_year","_month"]):
            if pd.isna(yr) or pd.isna(mo): continue
            p = period_str(int(yr), int(mo))
            trends["sub"][p]    = trends["sub"].get(p, 0)    + len(g)
            trends["sub_fp"][p] = trends["sub_fp"].get(p, 0) + int((g["_st"] == "Client Submit").sum())

    df = add_ym(data.get("intv", pd.DataFrame()), "intv")
    if not df.empty and "company_name" in df.columns and "workflow_step" in df.columns:
        df = df[df["company_name"].apply(include)].copy()
        df["_ws"] = df["workflow_step"].str.lower()
        for (yr, mo), g in df.groupby(["_year","_month"]):
            if pd.isna(yr) or pd.isna(mo): continue
            p = period_str(int(yr), int(mo))
            trends["l1"][p] = trends["l1"].get(p, 0) + int(g["_ws"].str.contains("l1 interview", na=False).sum())
            trends["l2"][p] = trends["l2"].get(p, 0) + int(g["_ws"].str.contains("l2 interview", na=False).sum())
            trends["l3"][p] = trends["l3"].get(p, 0) + int(g["_ws"].str.contains("l3 interview", na=False).sum())

    def sort_key(p):
        mon_map = {v: k for k, v in enumerate(MON_NAMES_PY)}
        import re
        m = re.match(r"([A-Za-z]+)'(\d+)", p)
        return int(m.group(2)) * 100 + mon_map.get(m.group(1), 0) if m else 0

    all_periods = sorted(set().union(*[set(d.keys()) for d in trends.values()]), key=sort_key)
    return {metric: [{"p": p, "v": d.get(p, 0)} for p in all_periods] for metric, d in trends.items()}


@lru_cache(maxsize=256)
def mom_trends_cached(client_filter_key=None):
    return mom_trends(load_data_cached(), thaw_filter(client_filter_key))


# ─────────────────────────────────────────────────────────────────────────────
# DAILY TRENDS
# ─────────────────────────────────────────────────────────────────────────────

def daily_trends(data, client_filter=None, from_date=None, to_date=None, grain="day"):
    client_filter = set(client_filter) if client_filter else None

    def norm(dt):
        if dt is None: return None
        dt = pd.Timestamp(dt)
        return dt.tz_localize(None) if dt.tzinfo else dt

    from_ts = norm(from_date)
    to_ts   = norm(to_date)

    def period_key(series):
        if grain == "month":
            return series.dt.to_period("M").astype(str)
        return series.dt.strftime("%d %b %Y")

    def count_by_date(df, date_col, cl_col):
        if df.empty:
            return {}
        if "_date" in df.columns:
            df = df[df["_date"].notna()]
        elif date_col in df.columns:
            df = df.copy()
            dates = pd.to_datetime(df[date_col], errors="coerce")
            if dates.dt.tz is not None:
                dates = dates.dt.tz_localize(None)
            df["_date"] = dates
            df = df.dropna(subset=["_date"])
        else:
            return {}
        if client_filter is not None and cl_col and cl_col in df.columns:
            df = df[df[cl_col].isin(client_filter)]
        if from_ts is not None:
            df = df[df["_date"] >= from_ts]
        if to_ts is not None:
            df = df[df["_date"] < (to_ts + pd.Timedelta(days=1))]
        if df.empty:
            return {}
        df = df.assign(__ds=period_key(df["_date"]))
        return df.groupby("__ds").size().to_dict()

    def cl(df, options):
        return next((c for c in options if c in df.columns), None)

    dem_df  = data.get("demand", pd.DataFrame())
    sub_df  = data.get("sub",    pd.DataFrame())
    intv_df = data.get("intv",   pd.DataFrame())
    sel_df  = data.get("sel",    pd.DataFrame())
    ob_df   = data.get("ob",     pd.DataFrame())
    hc_df   = data.get("activehc", pd.DataFrame())
    ex_df   = data.get("exit",   pd.DataFrame())

    dem_counts  = count_by_date(dem_df,  "Created_at",     cl(dem_df,  ["Company_name","company_name","client","Client"]))
    sub_counts  = count_by_date(sub_df,  "date",           cl(sub_df,  ["client","Client"]))
    intv_counts = count_by_date(intv_df, "Interview_date", cl(intv_df, ["company_name","Company_name","client","Client"]))
    sel_counts  = count_by_date(sel_df,  "selection_date", cl(sel_df,  ["company_name","Company_name","client","Client"]))
    ob_counts   = count_by_date(ob_df,   "display_date",   cl(ob_df,   ["company_name","Company_name","client","Client"]))
    hc_counts   = count_by_date(hc_df,   "display_date",   cl(hc_df,   ["company_name","Company_name","client","Client"]))

    # For month grain, convert total headcount snapshots into month-over-month
    # movement so the chart shows additions (+) and reductions (-).
    hc_movement_counts = hc_counts
    if grain == "month" and hc_counts:
        ordered_months = sorted(
            hc_counts.keys(),
            key=lambda value: pd.Period(value, freq="M").to_timestamp()
        )
        hc_movement_counts = {}
        prev_total = None
        for month_key in ordered_months:
            total = hc_counts.get(month_key, 0)
            hc_movement_counts[month_key] = 0 if prev_total is None else total - prev_total
            prev_total = total
    ex_counts   = count_by_date(ex_df,   "last_work_day",  cl(ex_df,   ["company_name","Company_name","client","Client"]))

    sub_fp_counts = {}
    if not sub_df.empty:
        work_sub = sub_df.copy()
        if "_date" not in work_sub.columns:
            if "date" in work_sub.columns:
                work_sub["_date"] = pd.to_datetime(work_sub["date"], errors="coerce")
            else:
                work_sub["_date"] = pd.NaT
        work_sub = work_sub[work_sub["_date"].notna()]
        sub_client_col = cl(work_sub, ["client","Client"])
        if client_filter is not None and sub_client_col and sub_client_col in work_sub.columns:
            work_sub = work_sub[work_sub[sub_client_col].isin(client_filter)]
        if from_ts is not None:
            work_sub = work_sub[work_sub["_date"] >= from_ts]
        if to_ts is not None:
            work_sub = work_sub[work_sub["_date"] < (to_ts + pd.Timedelta(days=1))]
        if "_st" not in work_sub.columns and "status" in work_sub.columns:
            work_sub["_st"] = work_sub["status"].astype(str).str.strip()
        if "_st" in work_sub.columns:
            work_sub = work_sub[work_sub["_st"] == "Client Submit"]
        else:
            work_sub = work_sub.iloc[0:0]
        if not work_sub.empty:
            work_sub = work_sub.assign(__ds=period_key(work_sub["_date"]))
            sub_fp_counts = work_sub.groupby("__ds").size().to_dict()

    dem_u_counts = {}
    if not dem_df.empty:
        work_dem = dem_df.copy()
        if "_date" not in work_dem.columns:
            if "Created_at" in work_dem.columns:
                work_dem["_date"] = pd.to_datetime(work_dem["Created_at"], errors="coerce")
            else:
                work_dem["_date"] = pd.NaT
        work_dem = work_dem[work_dem["_date"].notna()]

        demand_client_col = cl(work_dem, ["Company_name","company_name","client","Client"])
        if client_filter is not None and demand_client_col and demand_client_col in work_dem.columns:
            work_dem = work_dem[work_dem[demand_client_col].isin(client_filter)]

        submitted_ids = set()
        if not sub_df.empty:
            if client_filter is not None:
                sub_client_col = cl(sub_df, ["client","Client"])
                if sub_client_col and sub_client_col in sub_df.columns:
                    sub_scope = sub_df[sub_df[sub_client_col].isin(client_filter)]
                else:
                    sub_scope = sub_df
            else:
                sub_scope = sub_df

            if "_id_norm" in sub_scope.columns:
                submitted_ids = set(sub_scope["_id_norm"].dropna().astype(str).str.strip())
            else:
                id_col_sub = next((c for c in ID_COL_CANDIDATES if c in sub_scope.columns), None)
                if id_col_sub:
                    submitted_ids = set(sub_scope[id_col_sub].dropna().astype(str).str.strip())

        if from_ts is not None:
            work_dem = work_dem[work_dem["_date"] >= from_ts]
        if to_ts is not None:
            work_dem = work_dem[work_dem["_date"] < (to_ts + pd.Timedelta(days=1))]

        id_col_dem = "_id_norm" if "_id_norm" in work_dem.columns else next((c for c in ID_COL_CANDIDATES if c in work_dem.columns), None)
        if id_col_dem:
            demand_ids = work_dem[id_col_dem].astype(str).str.strip()
            work_dem = work_dem[~demand_ids.isin(submitted_ids)]
            # Also require id_status == "0"
            if "id_status" in work_dem.columns:
                work_dem = work_dem[
                    work_dem["id_status"].astype(str).str.strip() == "0"
                ]

        if not work_dem.empty:
            work_dem = work_dem.assign(__ds=period_key(work_dem["_date"]))
            dem_u_counts = work_dem.groupby("__ds").size().to_dict()

    print(f"[DEBUG] daily_trends from={from_ts} to={to_ts} | dem={sum(dem_counts.values())} sub={sum(sub_counts.values())} intv={sum(intv_counts.values())} sel={sum(sel_counts.values())} ob={sum(ob_counts.values())} ex={sum(ex_counts.values())}")
    if not dem_counts and not dem_df.empty and "Created_at" in dem_df.columns:
        sample = pd.to_datetime(dem_df["Created_at"], errors="coerce").dropna()
        if not sample.empty:
            print(f"[DEBUG] demand CSV date range: {sample.min()} -> {sample.max()}")

    all_counts = {"dem": dem_counts, "dem_u": dem_u_counts, "sub": sub_counts, "sub_fp": sub_fp_counts, "intv": intv_counts, "sel": sel_counts, "ob": ob_counts, "hc": hc_movement_counts, "ex": ex_counts}
    def sort_period(value):
        if grain == "month":
            return pd.Period(value, freq="M").to_timestamp()
        return pd.to_datetime(value, format="%d %b %Y", errors="coerce")
    all_dates = sorted(
        set().union(*[set(d.keys()) for d in all_counts.values()]),
        key=sort_period
    )
    return {
        k: [{"d": date, "v": all_counts[k].get(date, 0)} for date in all_dates]
        for k in all_counts
    }


@lru_cache(maxsize=256)
def daily_trends_cached(client_filter_key=None, from_date_key=None, to_date_key=None, grain="day"):
    return daily_trends(
        load_data_cached(),
        thaw_filter(client_filter_key),
        thaw_date(from_date_key),
        thaw_date(to_date_key),
        grain=grain,
    )


# ─────────────────────────────────────────────────────────────────────────────
# ROUTES
# ─────────────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template_string(HTML)


@app.route("/api/init")
def api_init():
    periods = get_periods_cached()
    client_meta = list(get_client_catalog())
    clients = [entry["name"] for entry in client_meta]
    years = sorted({p[0] for p in periods}, reverse=True)
    months = sorted({p[1] for p in periods})
    _, _, domains, business_heads = load_mapping()
    return jsonify({
        "years": years, "months": months, "clients": clients,
        "domains": domains, "business_heads": business_heads,
        "client_meta": client_meta,
    })


@app.route("/api/daily_trends")
def api_daily_trends():
    from_str = request.args.get("from")
    to_str   = request.args.get("to")
    grain    = request.args.get("grain", default="day", type=str)
    from_date = pd.to_datetime(from_str, errors="coerce") if from_str else None
    to_date   = pd.to_datetime(to_str,   errors="coerce") if to_str   else None
    if from_date is not None and hasattr(from_date, 'tzinfo') and from_date.tzinfo:
        from_date = from_date.tz_localize(None)
    if to_date is not None and hasattr(to_date, 'tzinfo') and to_date.tzinfo:
        to_date = to_date.tz_localize(None)

    cl_filter, _, _ = get_resolved_filters()

    result = daily_trends_cached(
        freeze_filter(cl_filter),
        freeze_date(from_date),
        freeze_date(to_date),
        grain,
    )

    # Fallback to full dataset if date filter returns nothing
    total_all = sum(sum(x['v'] for x in v) for v in result.values())
    if total_all == 0 and (from_date is not None or to_date is not None):
        print("[DEBUG] No data matched date filter - falling back to full dataset")
        result = daily_trends_cached(freeze_filter(cl_filter), None, None, grain)

    return jsonify(result)


@app.route("/api/months")
def api_months():
    periods  = get_periods_cached()
    sel_years = parse_int_csv_filter_arg("years")
    if sel_years is None:
        single_year = request.args.get("year", type=int)
        sel_years = {single_year} if single_year else None
    months   = sorted({p[1] for p in periods if not sel_years or p[0] in sel_years})
    return jsonify({"months": months})


@app.route("/api/refresh")
def api_refresh():
    global LAST_REFRESHED_AT
    load_data_cached.cache_clear()
    load_mapping.cache_clear()
    get_periods_cached.cache_clear()
    get_client_catalog.cache_clear()
    resolve_client_filter_cached.cache_clear()
    compute_all_cached.cache_clear()
    mom_trends_cached.cache_clear()
    daily_trends_cached.cache_clear()
    LAST_REFRESHED_AT = datetime.now()
    return jsonify({"status": "refreshed"})


@app.route("/api/debug_context")
def api_debug_context():
    sel_years, sel_months = get_period_filters()
    cl_filter, dom_filter, bh_filter = get_resolved_filters()

    res = compute_all_cached(freeze_filter(sel_years), freeze_filter(sel_months), freeze_filter(cl_filter))
    client_to_domain, client_to_bh, client_lookup, _, _ = get_mapping_context()

    for cl in list(res.keys()):
        if CAPTIVE_SUFFIX in cl:
            base = cl.replace(CAPTIVE_SUFFIX, "").strip()
            client_to_domain[cl] = "Captive"
            client_to_bh[cl] = client_to_bh.get(base, "")

    visible = {}
    for cl, metrics in res.items():
        mapped = get_mapped_client_name(cl, client_lookup)
        domain = client_to_domain.get(mapped, "")
        bh = normalize_bh_label(client_to_bh.get(mapped, ""))
        if dom_filter and domain not in dom_filter:
            continue
        if bh_filter and bh not in bh_filter:
            continue
        visible[cl] = metrics

    g = grand_total(visible)
    data = load_data_cached()

    return jsonify({
        "data_folder": DATA_FOLDER,
        "selected_years": sorted(sel_years) if sel_years else [],
        "selected_months": sorted(sel_months) if sel_months else [],
        "exit_source_rows": len(data.get("exit", pd.DataFrame())),
        "visible_clients": len(visible),
        "exits_hc": int(round(g["ex_hc"])),
        "exits_po_l": round(float(g["ex_po"]), 2),
        "exits_margin_l": round(float(g["ex_mg"]), 2),
    })


@app.route("/api/data")
def api_data():
    sel_years, sel_months = get_period_filters()
    cl_filter, dom_filter, bh_filter = get_resolved_filters()

    res = compute_all_cached(freeze_filter(sel_years), freeze_filter(sel_months), freeze_filter(cl_filter))

    client_to_domain, client_to_bh, client_lookup, _, _ = get_mapping_context()

    # Assign domain/BH for any split captive keys (e.g. "Flipkart [Captive]")
    for cl in list(res.keys()):
        if CAPTIVE_SUFFIX in cl:
            base = cl.replace(CAPTIVE_SUFFIX, "").strip()
            client_to_domain[cl] = "Captive"
            client_to_bh[cl]     = client_to_bh.get(base, "")

    visible = []
    for cl, m in sorted(res.items(), key=lambda x: x[0].lower()):
        mapped = get_mapped_client_name(cl, client_lookup)
        domain = client_to_domain.get(mapped, "")
        bh = normalize_bh_label(client_to_bh.get(mapped, ""))
        if dom_filter and domain not in dom_filter:
            continue
        if bh_filter and bh not in bh_filter:
            continue
        visible.append((cl, m, domain, bh))

    g = grand_total({cl: m for cl, m, _, _ in visible})
    rows = [{"label": cl, "metrics": round_m(m), "domain": domain, "bh": bh}
            for cl, m, domain, bh in visible]
    return jsonify({"rows": rows, "grand": round_m(g)})


@app.route("/api/trends")
def api_trends():
    cl_filter, _, _ = get_resolved_filters()
    return jsonify(mom_trends_cached(freeze_filter(cl_filter)))


@app.route("/api/raw_data")
def api_raw_data():
    dataset, year_filter, month_filter, client_filter, from_date, to_date, demand_status = get_raw_request_context()
    if dataset not in RAW_DATASET_CONFIG:
        return jsonify({"error": "Invalid dataset"}), 400

    df = get_raw_dataset_frame(dataset, year_filter, month_filter, client_filter, from_date, to_date, demand_status)
    payload = serialize_raw_dataset(df)
    payload["dataset"] = dataset
    if dataset == "demand" and demand_status in {"unserviced", "serviced"}:
        payload["label"] = RAW_DATASET_CONFIG[dataset]["label"] + f" - {demand_status.title()}"
    else:
        payload["label"] = RAW_DATASET_CONFIG[dataset]["label"]
    return jsonify(payload)


@app.route("/api/raw_data_export")
def api_raw_data_export():
    dataset, year_filter, month_filter, client_filter, from_date, to_date, demand_status = get_raw_request_context()
    if dataset not in RAW_DATASET_CONFIG:
        return jsonify({"error": "Invalid dataset"}), 400

    file_format = request.args.get("format", default="csv", type=str).lower()
    column_mode = request.args.get("columns", default="visible", type=str).lower()
    if file_format not in {"csv", "xlsx"}:
        return jsonify({"error": "Invalid export format"}), 400
    if column_mode not in {"visible", "all"}:
        return jsonify({"error": "Invalid column mode"}), 400

    df = get_raw_dataset_frame(dataset, year_filter, month_filter, client_filter, from_date, to_date, demand_status)
    export_columns = get_export_columns(df, column_mode)
    export_df = df[export_columns].fillna("") if export_columns else pd.DataFrame()
    filename = build_raw_export_filename(
        dataset,
        client_filter=client_filter,
        year_filter=year_filter,
        month_filter=month_filter,
        from_date=from_date,
        to_date=to_date,
        column_mode=column_mode,
        file_ext=file_format,
        demand_status=demand_status,
    )

    if file_format == "xlsx":
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            export_df.to_excel(writer, index=False, sheet_name="Raw Data")
        output.seek(0)
        return Response(
            output.getvalue(),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'},
        )

    csv_text = export_df.to_csv(index=False) if export_columns else ""
    return Response(
        csv_text,
        mimetype="text/csv",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.route("/api/lmtd")
def api_lmtd():
    sel_years, sel_months = get_period_filters()
    cl_filter, dom_filter, bh_filter = get_resolved_filters()
    sel_year = next(iter(sel_years)) if sel_years and len(sel_years) == 1 else None
    sel_month = next(iter(sel_months)) if sel_months and len(sel_months) == 1 else None

    today = pd.Timestamp.now().normalize()
    if not sel_year or not sel_month or sel_year != today.year or sel_month != today.month:
        return jsonify({"enabled": False})

    current_start = today.replace(day=1)
    prev_month_end = current_start - pd.Timedelta(days=1)
    prev_start = prev_month_end.replace(day=1)
    prev_end_day = min(today.day, prev_month_end.day)
    prev_end = prev_start.replace(day=prev_end_day)

    current_res = compute_all_cached(
        sel_year,
        sel_month,
        freeze_filter(cl_filter),
        freeze_date(current_start),
        freeze_date(today),
    )
    prev_res = compute_all_cached(
        prev_start.year,
        prev_start.month,
        freeze_filter(cl_filter),
        freeze_date(prev_start),
        freeze_date(prev_end),
    )

    client_to_domain, client_to_bh, client_lookup, _, _ = get_mapping_context()

    for res in (current_res, prev_res):
        for cl in list(res.keys()):
            if CAPTIVE_SUFFIX in cl:
                base = cl.replace(CAPTIVE_SUFFIX, "").strip()
                client_to_domain[cl] = "Captive"
                client_to_bh[cl] = client_to_bh.get(base, "")

    def visible_metrics(res):
        visible = {}
        for cl, m in res.items():
            mapped = get_mapped_client_name(cl, client_lookup)
            domain = client_to_domain.get(mapped, "")
            bh = normalize_bh_label(client_to_bh.get(mapped, ""))
            if dom_filter and domain not in dom_filter:
                continue
            if bh_filter and bh not in bh_filter:
                continue
            visible[cl] = m
        return visible

    current_visible = visible_metrics(current_res)
    prev_visible = visible_metrics(prev_res)

    title = f"LMTD vs {prev_end.strftime('%d %b %Y')}"
    return jsonify({
        "enabled": True,
        "title": title,
        "grand": round_m(grand_total(current_visible)),
        "previous_grand": round_m(grand_total(prev_visible)),
    })


# ─────────────────────────────────────────────────────────────────────────────
# CEO ANALYTICS — NEW ROUTES
# ─────────────────────────────────────────────────────────────────────────────

@app.route("/api/aging_demands")
def api_aging_demands():
    """Unserviced demands bucketed by age: <30, 30-60, 60-90, 90+ days."""
    data = load_data_cached()
    dem_df = data.get("demand", pd.DataFrame())
    sub_df = data.get("sub",    pd.DataFrame())

    if dem_df.empty or "Created_at" not in dem_df.columns:
        return jsonify({"buckets": [], "top_clients": []})

    submitted_ids = set()
    id_col = next((c for c in ["job_ID","job_id","Job_ID","ID","id"] if c in sub_df.columns), None)
    if id_col:
        submitted_ids = set(sub_df[id_col].astype(str).str.strip().unique())

    dem_df = dem_df.copy()
    dem_df["_date"] = pd.to_datetime(dem_df["Created_at"], errors="coerce")
    today = pd.Timestamp.now().normalize()

    id_col_d = next((c for c in ["id","ID","job_id","Job_ID"] if c in dem_df.columns), None)
    if id_col_d:
        unserviced = dem_df[~dem_df[id_col_d].astype(str).str.strip().isin(submitted_ids)].copy()
    else:
        unserviced = dem_df.copy()

    unserviced = unserviced.dropna(subset=["_date"])
    unserviced["_age"] = (today - unserviced["_date"]).dt.days

    buckets = {
        "<30 days":  int((unserviced["_age"] <  30).sum()),
        "30-60 days":int(((unserviced["_age"] >= 30) & (unserviced["_age"] < 60)).sum()),
        "60-90 days":int(((unserviced["_age"] >= 60) & (unserviced["_age"] < 90)).sum()),
        "90+ days":  int((unserviced["_age"] >= 90).sum()),
    }

    cl_col = next((c for c in ["Company_name","company_name","client","Client"] if c in unserviced.columns), None)
    top_clients = []
    if cl_col:
        tc = unserviced.groupby(cl_col).size().sort_values(ascending=False).head(5)
        top_clients = [{"client": k, "count": int(v)} for k, v in tc.items()]

    return jsonify({"buckets": [{"label": k, "v": v} for k, v in buckets.items()], "top_clients": top_clients})


@app.route("/api/exit_reasons")
def api_exit_reasons():
    """Exit breakdown by exit_type column."""
    sel_years, sel_months = get_period_filters()
    cl_filter, _, _ = get_resolved_filters()

    data   = load_data_cached()
    ex_df  = data.get("exit", pd.DataFrame())
    if ex_df.empty or "exit_type" not in ex_df.columns:
        return jsonify({"reasons": []})

    ex_df = add_ym(ex_df.copy(), "exit")
    ex_df = filter_ym(ex_df, sel_years, sel_months)

    cl_col = next((c for c in ["company_name","Company_name","client","Client"] if c in ex_df.columns), None)
    if cl_filter is not None and cl_col:
        ex_df = ex_df[ex_df[cl_col].isin(cl_filter)]

    counts = ex_df["exit_type"].str.strip().value_counts()
    total  = int(counts.sum()) or 1
    reasons = [{"label": k, "v": int(v), "pct": round(int(v)/total*100, 1)}
               for k, v in counts.items() if k]
    return jsonify({"reasons": reasons, "total": total})


@app.route("/api/bh_conversion")
def api_bh_conversion():
    """Conversion rates by Business Head: demand→sub, sub→L1, L1→sel, sel→ob."""
    sel_years, sel_months = get_period_filters()
    cl_filter, _, _ = get_resolved_filters()

    _, client_to_bh, client_lookup, _, _ = get_mapping_context()
    res = compute_all_cached(freeze_filter(sel_years), freeze_filter(sel_months), freeze_filter(cl_filter))

    # Aggregate metrics per BH
    bh_metrics = {}
    for cl, m in res.items():
        bh = normalize_bh_label(client_to_bh.get(get_mapped_client_name(cl, client_lookup), "")) or "Unassigned"
        if bh not in bh_metrics:
            bh_metrics[bh] = dict(dem=0, sub=0, l1=0, sel=0, ob_hc=0)
        for k in bh_metrics[bh]:
            bh_metrics[bh][k] += m.get(k, 0)

    def pct(a, b): return round(a/b*100, 1) if b else 0

    rows = []
    for bh, m in sorted(bh_metrics.items()):
        rows.append({
            "bh":       bh,
            "dem":      m["dem"],
            "sub":      m["sub"],
            "l1":       m["l1"],
            "sel":      m["sel"],
            "ob":       m["ob_hc"],
            "dem_sub":  pct(m["sub"],   m["dem"]),
            "sub_l1":   pct(m["l1"],    m["sub"]),
            "l1_sel":   pct(m["sel"],   m["l1"]),
            "sel_ob":   pct(m["ob_hc"], m["sel"]),
        })
    return jsonify({"rows": rows})


@app.route("/api/revenue_trend")
def api_revenue_trend():
    """12-month rolling net PO and margin trend."""
    cl_filter, _, _ = get_resolved_filters()

    data  = load_data_cached()
    ob_df = data.get("ob",   pd.DataFrame())
    ex_df = data.get("exit", pd.DataFrame())

    def monthly_stats(df, date_col, client_filter):
        if df.empty:
            return {}
        if "_date" not in df.columns and date_col not in df.columns:
            return {}
        cl_col = next((c for c in ["company_name","Company_name","client","Client"] if c in df.columns), None)
        if client_filter is not None and cl_col:
            df = df[df[cl_col].isin(client_filter)]
        if "_date" in df.columns:
            df = df[df["_date"].notna()]
        else:
            df = df.copy()
            df["_date"] = pd.to_datetime(df[date_col], errors="coerce")
            df = df.dropna(subset=["_date"])
        if df.empty:
            return {}
        df = df.assign(_ym=df["_date"].dt.to_period("M"))
        grp = df.groupby("_ym").agg(
            hc=("_date", "size"),
            po=("_po", "sum"),
            mg=("_mg", "sum"),
        )
        return {
            str(k): {
                "hc": int(v["hc"]),
                "po": round(float(v["po"]) / 1e5, 2),
                "mg": round(float(v["mg"]) / 1e5, 2),
            }
            for k, v in grp.iterrows()
        }

    ob_monthly = monthly_stats(ob_df, "display_date",  cl_filter)
    ex_monthly = monthly_stats(ex_df, "last_work_day", cl_filter)

    # Last 12 months
    today = pd.Timestamp.now()
    periods = [(today - pd.DateOffset(months=i)).to_period("M") for i in range(11, -1, -1)]

    trend = []
    for p in periods:
        ps   = str(p)
        ob_hc = ob_monthly.get(ps, {}).get("hc", 0)
        ex_hc = ex_monthly.get(ps, {}).get("hc", 0)
        ob_po = ob_monthly.get(ps, {}).get("po", 0)
        ob_mg = ob_monthly.get(ps, {}).get("mg", 0)
        ex_po = ex_monthly.get(ps, {}).get("po", 0)
        ex_mg = ex_monthly.get(ps, {}).get("mg", 0)
        trend.append({
            "period":  p.strftime("%b'%y"),
            "ob_hc":   ob_hc,
            "ex_hc":   ex_hc,
            "net_hc":  int(ob_hc - ex_hc),
            "ob_po":   ob_po,
            "ex_po":   ex_po,
            "net_po":  round(ob_po - ex_po, 2),
            "ob_mg":   ob_mg,
            "ex_mg":   ex_mg,
            "net_mg":  round(ob_mg - ex_mg, 2),
        })
    return jsonify({"trend": trend})



# ─────────────────────────────────────────────────────────────────────────────
# HTML
# ─────────────────────────────────────────────────────────────────────────────
HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Joules to Watts — Recruitment Dashboard</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js"></script>
<script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0/dist/chartjs-plugin-datalabels.min.js"></script>
<style>
:root{
  --bg:#f0f2f5;--surface:#ffffff;--s2:#f5f6fa;--s3:#eaecf0;--s4:#dde1e8;
  --border:rgba(0,0,0,0.08);--border2:rgba(0,0,0,0.14);
  --text:#1a1d23;--t2:#52586a;--t3:#8a91a0;
  --red:#e8453c;--red2:#d63b33;--red-lt:rgba(232,69,60,0.10);--red-bd:rgba(232,69,60,0.28);
  --green:#1db85a;--green2:#17a050;--green-lt:rgba(29,184,90,0.10);--green-bd:rgba(29,184,90,0.28);
  --grey:#6b7385;--grey-lt:rgba(107,115,133,0.10);--grey-bd:rgba(107,115,133,0.22);
  --r:10px;--rsm:6px;--rlg:14px;
}
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:'Segoe UI',system-ui,sans-serif;font-size:13px;background:var(--bg);color:var(--text);min-height:100vh;}
.nav{background:rgba(255,255,255,0.97);border-bottom:1px solid var(--border);padding:0 22px;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:0;position:sticky;top:0;z-index:100;backdrop-filter:blur(14px);min-height:54px;}
.brand{display:flex;align-items:center;gap:12px;padding:8px 0;}
.pill{font-size:11px;font-weight:600;padding:3px 10px;border-radius:20px;background:var(--red-lt);color:var(--red2);border:1px solid var(--red-bd);}
.nav-r{display:flex;align-items:center;gap:6px;flex-wrap:wrap;padding:10px 0;}
.nlbl{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:var(--t3);margin-left:4px;}
.nav-r select{padding:5px 9px;border:1px solid var(--border2);border-radius:var(--rsm);background:var(--s2);color:var(--text);font-size:12px;cursor:pointer;outline:none;}
.nav-r select:focus{border-color:var(--red);}
.btn-g{padding:5px 12px;border:1px solid var(--border2);border-radius:var(--rsm);background:transparent;color:var(--t2);font-size:12px;cursor:pointer;}
.btn-g:hover{background:var(--s2);color:var(--text);}
.btn-r{padding:5px 14px;border:none;border-radius:var(--rsm);background:var(--red);color:#fff;font-size:12px;font-weight:700;cursor:pointer;}
.btn-r:hover{background:#c0392b;}
.cl-wrap{position:relative;}
.cl-btn{padding:5px 10px;border:1px solid var(--border2);border-radius:var(--rsm);background:var(--s2);color:var(--text);font-size:12px;cursor:pointer;display:flex;align-items:center;gap:6px;white-space:nowrap;}
.cl-ct{background:var(--red);color:#fff;font-size:10px;font-weight:700;padding:1px 6px;border-radius:10px;display:none;}
.cl-panel{position:absolute;top:calc(100% + 6px);right:0;width:270px;background:var(--surface);border:1px solid var(--border2);border-radius:var(--r);box-shadow:0 8px 32px rgba(0,0,0,.15);z-index:200;display:none;}
.cl-panel.open{display:block;}
.cl-search{padding:10px 10px 6px;}
.cl-search input{width:100%;padding:6px 10px;border:1px solid var(--border2);border-radius:var(--rsm);background:var(--s2);color:var(--text);font-size:12px;outline:none;}
.cl-search input:focus{border-color:var(--red);}
.cl-actions{display:flex;gap:6px;padding:0 10px 6px;}
.cl-actions button{flex:1;padding:4px;border:1px solid var(--border2);border-radius:var(--rsm);background:transparent;color:var(--t2);font-size:11px;cursor:pointer;}
.cl-actions button:hover{background:var(--s2);color:var(--text);}
.cl-list{max-height:260px;overflow-y:auto;padding:4px 0 8px;}
.cl-item{display:flex;align-items:center;gap:8px;padding:5px 12px;cursor:pointer;}
.cl-item:hover{background:var(--s2);}
.cl-item input{accent-color:var(--red);width:13px;height:13px;flex-shrink:0;cursor:pointer;}
.cl-item label{font-size:12px;color:var(--text);cursor:pointer;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.cl-apply{padding:8px 10px;border-top:1px solid var(--border);}
.cl-apply button{width:100%;padding:7px;border:none;border-radius:var(--rsm);background:var(--red);color:#fff;font-size:12px;font-weight:700;cursor:pointer;}
.cl-apply button:hover{background:#c0392b;}
.body{max-width:1600px;margin:0 auto;padding:20px 18px 60px;}
.sec{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.09em;color:var(--t3);margin:22px 0 10px;display:flex;align-items:center;gap:10px;}
.sec::after{content:'';flex:1;height:1px;background:var(--border);}
.kpi-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(205px,1fr));gap:14px;margin-bottom:16px;}
.kpi{background:
  radial-gradient(circle at top right,rgba(255,255,255,.95),transparent 34%),
  linear-gradient(180deg,rgba(255,255,255,.96),rgba(247,249,252,.96));
  border:1px solid rgba(255,255,255,.8);border-radius:18px;padding:18px 18px 16px;position:relative;overflow:hidden;
  box-shadow:0 12px 28px rgba(16,24,40,.08), inset 0 1px 0 rgba(255,255,255,.85);
  transition:transform .22s ease, box-shadow .22s ease, border-color .22s ease;
  min-height:162px;
}
.kpi:hover{transform:translateY(-4px);box-shadow:0 18px 34px rgba(16,24,40,.14), inset 0 1px 0 rgba(255,255,255,.9);border-color:rgba(255,255,255,.95);}
.kpi::before{content:'';position:absolute;inset:auto -28% 58% auto;width:140px;height:140px;border-radius:50%;filter:blur(10px);opacity:.18;}
.kpi::after{content:'';position:absolute;top:0;left:16px;right:16px;height:4px;border-radius:999px;}
.kpi.red::before{background:radial-gradient(circle,var(--red),transparent 68%);}
.kpi.green::before{background:radial-gradient(circle,var(--green),transparent 68%);}
.kpi.grey::before{background:radial-gradient(circle,#667085,transparent 68%);}
.kpi.blue::before{background:radial-gradient(circle,#3498db,transparent 68%);}
.kpi.red::after{background:linear-gradient(90deg,var(--red),#ff7a70);}
.kpi.green::after{background:linear-gradient(90deg,var(--green),#2ed38d);}
.kpi.grey::after{background:linear-gradient(90deg,#5b6578,#99a3b6);}
.kpi.blue::after{background:linear-gradient(90deg,#3498db,#69b8ff);}
.kpi-head{display:flex;align-items:center;justify-content:space-between;gap:10px;margin-bottom:12px;}
.kpi-ico{width:42px;height:42px;border-radius:14px;display:flex;align-items:center;justify-content:center;font-size:21px;
  background:rgba(255,255,255,.88);border:1px solid rgba(255,255,255,.95);box-shadow:0 8px 20px rgba(15,23,42,.07);}
.kpi.red .kpi-ico{background:linear-gradient(180deg,rgba(232,69,60,.16),rgba(255,255,255,.95));}
.kpi.green .kpi-ico{background:linear-gradient(180deg,rgba(29,184,90,.16),rgba(255,255,255,.95));}
.kpi.grey .kpi-ico{background:linear-gradient(180deg,rgba(107,115,133,.16),rgba(255,255,255,.95));}
.kpi.blue .kpi-ico{background:linear-gradient(180deg,rgba(52,152,219,.16),rgba(255,255,255,.95));}
.kpi-lbl{font-size:10px;font-weight:800;text-transform:uppercase;letter-spacing:.12em;color:var(--t3);margin-bottom:2px;}
.kpi-val{font-size:40px;font-weight:900;line-height:.95;letter-spacing:-.03em;margin-bottom:8px;}
.kpi-val.animating{opacity:.88;transform:translateY(2px);}
.kpi-sub{font-size:12px;color:var(--t2);line-height:1.55;min-height:38px;}
.kpi-sub strong{font-size:13px;color:var(--text);}
.kpi .ktag{margin-top:10px;padding:5px 10px;font-size:10px;border-radius:999px;box-shadow:inset 0 1px 0 rgba(255,255,255,.7);}
.ktag{display:inline-flex;align-items:center;gap:3px;font-size:10px;font-weight:700;padding:2px 8px;border-radius:20px;margin-top:5px;}
.tg{background:var(--green-lt);color:var(--green2);border:1px solid var(--green-bd);}
.tr{background:var(--red-lt);color:var(--red2);border:1px solid var(--red-bd);}
.tb{background:rgba(52,152,219,.12);color:#2c6ea4;border:1px solid rgba(52,152,219,.28);}
.tgr{background:var(--grey-lt);color:var(--grey);border:1px solid var(--grey-bd);}
.prev-bar{background:var(--s2);border:1px solid var(--border);border-radius:var(--r);padding:10px 16px;margin-bottom:20px;display:flex;align-items:center;flex-wrap:wrap;gap:6px 18px;}
.prev-bar-title{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.07em;color:var(--t3);margin-right:4px;white-space:nowrap;}
.prev-metric{display:flex;align-items:center;gap:5px;}
.prev-metric-lbl{font-size:11px;color:var(--t2);}
.prev-metric-val{font-size:13px;font-weight:700;color:var(--text);}
.prev-metric-chg{font-size:10px;font-weight:700;}
.prev-sep{width:1px;height:20px;background:var(--border2);}
.r2{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:14px;}
.r3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:14px;margin-bottom:14px;}
@media(max-width:900px){.r2,.r3{grid-template-columns:1fr;}}
.card{background:var(--surface);border:1px solid var(--border);border-radius:var(--rlg);padding:16px 18px;}
.card-t{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:var(--t3);margin-bottom:14px;}
.trend-chart-wrap{position:relative;height:260px;}
.trend-avg{
  position:absolute;top:0;right:0;z-index:2;
  display:flex;align-items:center;gap:6px;
  padding:6px 10px;border-radius:999px;
  background:rgba(255,255,255,.92);border:1px solid var(--border);
  box-shadow:0 10px 24px rgba(15,23,42,.08);
  color:var(--t2);font-size:11px;font-weight:700;
  backdrop-filter:blur(6px);
}
.trend-avg strong{color:var(--text);font-size:12px;}
.stage-wrap{display:flex;gap:4px;align-items:stretch;overflow-x:auto;padding-bottom:2px;}
.stage{flex:1;min-width:78px;background:var(--s2);border:1px solid var(--border);border-radius:var(--r);padding:12px 8px;text-align:center;}
.sarrow{display:flex;align-items:center;justify-content:center;color:var(--t3);font-size:18px;padding:0 1px;flex-shrink:0;}
.sv{font-size:20px;font-weight:800;line-height:1;margin-bottom:3px;}
.sl{font-size:9px;color:var(--t3);text-transform:uppercase;letter-spacing:.05em;}
.sc{font-size:10px;font-weight:700;margin-top:4px;}
.fn{display:flex;align-items:center;gap:10px;margin-bottom:9px;}
.fnl{font-size:11px;color:var(--t2);width:122px;flex-shrink:0;}
.fnt{flex:1;height:20px;background:var(--s2);border-radius:4px;overflow:hidden;border:1px solid var(--border);}
.fnf{height:100%;border-radius:3px;transition:width .6s cubic-bezier(.4,0,.2,1);}
.fnn{font-size:12px;font-weight:700;min-width:55px;text-align:right;flex-shrink:0;}
.fnp{font-size:10px;color:var(--t3);min-width:38px;text-align:right;flex-shrink:0;}
.mrr-r{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;}
.mrr-c{background:var(--s2);border:1px solid var(--border);border-radius:var(--r);padding:12px;text-align:center;}
.mrr-v{font-size:20px;font-weight:800;margin-bottom:3px;}
.mrr-l{font-size:10px;color:var(--t3);text-transform:uppercase;letter-spacing:.04em;}
.trend-tabs{display:flex;gap:8px;margin-bottom:14px;flex-wrap:wrap;align-items:center;}
.ttab{
  position:relative;overflow:hidden;isolation:isolate;
  padding:7px 15px;border:1px solid rgba(0,0,0,0.10);border-radius:999px;
  font-size:11px;font-weight:700;letter-spacing:.01em;cursor:pointer;
  background:linear-gradient(180deg,rgba(255,255,255,.92),rgba(244,246,250,.96));
  color:var(--t2);
  box-shadow:0 4px 14px rgba(15,23,42,.04), inset 0 1px 0 rgba(255,255,255,.9);
  transition:transform .22s ease, box-shadow .22s ease, border-color .22s ease, color .22s ease, background .22s ease;
}
.ttab::before{
  content:'';position:absolute;inset:0;z-index:-2;border-radius:inherit;
  background:linear-gradient(135deg,rgba(255,255,255,.96),rgba(235,238,244,.92));
}
.ttab::after{
  content:'';position:absolute;top:-120%;left:-40%;width:42%;height:300%;z-index:-1;
  background:linear-gradient(90deg,transparent,rgba(255,255,255,.55),transparent);
  transform:rotate(18deg) translateX(-220%);
  transition:transform .7s ease;
}
.ttab:hover{
  transform:translateY(-1px);
  box-shadow:0 8px 18px rgba(15,23,42,.08), inset 0 1px 0 rgba(255,255,255,.95);
}
.ttab:hover::after{transform:rotate(18deg) translateX(420%);}
.ttab.on{
  color:#fff;border-color:rgba(232,69,60,.82);
  background:linear-gradient(135deg,#ff6a5f 0%, #e8453c 52%, #cf3129 100%);
  box-shadow:0 12px 26px rgba(232,69,60,.22), inset 0 1px 0 rgba(255,255,255,.18);
  transform:translateY(-1px) scale(1.01);
}
.ttab.on::before{
  background:
    radial-gradient(circle at 18% 20%,rgba(255,255,255,.24),transparent 34%),
    linear-gradient(135deg,#ff6a5f 0%, #e8453c 52%, #cf3129 100%);
}
.ttab.on::after{
  transform:rotate(18deg) translateX(280%);
  animation:tabSweep 2.8s ease-in-out infinite;
}
.ttab:hover:not(.on){color:var(--text);border-color:rgba(0,0,0,0.14);}
.ttab:active{transform:translateY(0) scale(.985);}
@keyframes tabSweep{
  0%, 100%{transform:rotate(18deg) translateX(-220%);}
  45%, 55%{transform:rotate(18deg) translateX(360%);}
}
.sec-toggle{width:100%;display:flex;align-items:center;justify-content:space-between;gap:12px;padding:0;border:none;background:transparent;cursor:pointer;text-align:left;}
.sec-toggle .card-t{margin-bottom:0;}
.sec-toggle-meta{font-size:11px;color:var(--t3);font-weight:600;white-space:nowrap;}
.sec-chevron{font-size:18px;color:var(--t2);transition:transform .18s ease;}
.sec-toggle.collapsed .sec-chevron{transform:rotate(-90deg);}
.pipeline-shell{margin-bottom:14px;}
.pipeline-body{margin-top:14px;}
.pipeline-body.collapsed{display:none;}
.pipe-client-head{display:flex;align-items:center;justify-content:space-between;gap:10px;flex-wrap:wrap;margin:16px 0 10px;}
.pipe-client-meta{font-size:11px;color:var(--t3);}
.pipe-client-wrap{border:1px solid var(--border);border-radius:12px;overflow:hidden;background:#fff;}
.pipe-client-scroll{max-height:320px;overflow:auto;}
.pipe-table{width:100%;border-collapse:collapse;font-size:12px;}
.pipe-table thead th{position:sticky;top:0;background:#f8f9fc;z-index:1;text-align:right;padding:10px 12px;border-bottom:1px solid var(--border2);font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:var(--t3);white-space:nowrap;}
.pipe-table thead th:first-child,.pipe-table tbody td:first-child{text-align:left;}
.pipe-table tbody td{padding:9px 12px;border-bottom:1px solid var(--border);text-align:right;white-space:nowrap;}
.pipe-table tbody tr:hover{background:#fbfbfd;}
.pipe-client-name{font-weight:700;color:var(--text);}
.pipe-sub{display:block;font-size:10px;color:var(--t3);margin-top:2px;}
.pipe-rate{font-weight:700;}
.pipe-empty{padding:16px;color:var(--t3);font-size:12px;text-align:center;}
.raw-shell{margin-top:14px;}
.raw-toolbar{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:14px;align-items:stretch;margin-bottom:14px;}
.raw-filter{
  display:flex;flex-direction:column;gap:8px;min-width:0;
  padding:12px;border:1px solid var(--border);border-radius:14px;
  background:linear-gradient(180deg,rgba(255,255,255,.95),rgba(245,247,250,.96));
  box-shadow:0 8px 22px rgba(15,23,42,.04);
}
.raw-filter label{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:var(--t3);}
.raw-inline-check{
  display:flex;align-items:center;gap:8px;min-height:38px;
  padding:7px 10px;border:1px solid var(--border2);border-radius:8px;background:var(--s2);
}
.raw-inline-check input{accent-color:var(--red);width:14px;height:14px;}
.raw-inline-check span{font-size:12px;color:var(--text);font-weight:600;}
.raw-filter input{
  padding:7px 10px;border:1px solid var(--border2);border-radius:8px;
  background:var(--s2);color:var(--text);font-size:12px;outline:none;
}
.raw-filter select{
  padding:7px 10px;border:1px solid var(--border2);border-radius:8px;
  background:var(--s2);color:var(--text);font-size:12px;outline:none;
}
.raw-filter input:focus{border-color:var(--red);}
.raw-filter select:focus{border-color:var(--red);}
.raw-filter-note{font-size:11px;color:var(--t3);line-height:1.35;}
.raw-time-stack{display:flex;flex-direction:column;gap:12px;}
.raw-date-stack{display:flex;flex-direction:column;gap:10px;}
.raw-date-field{display:flex;flex-direction:column;gap:6px;}
.raw-date-field span{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:var(--t3);}
.raw-range-presets{display:flex;gap:6px;flex-wrap:wrap;align-items:center;margin-top:4px;}
.raw-range-presets button{padding:6px 10px;font-size:11px;}
.raw-demand-pills{display:flex;gap:8px;flex-wrap:wrap;}
.raw-demand-pill{
  padding:8px 12px;border:1px solid var(--border2);border-radius:999px;
  background:var(--s2);color:var(--t2);font-size:12px;font-weight:700;cursor:pointer;
  transition:background .18s ease,border-color .18s ease,color .18s ease,transform .18s ease,box-shadow .18s ease;
}
.raw-demand-pill:hover{transform:translateY(-1px);color:var(--text);border-color:rgba(0,0,0,0.18);}
.raw-demand-pill.active{
  background:linear-gradient(135deg,#ff6a5f 0%, #e8453c 55%, #cf3129 100%);
  border-color:rgba(232,69,60,.78);color:#fff;
  box-shadow:0 10px 20px rgba(232,69,60,.18);
}
.raw-filter-wide{grid-column:span 2;}
.raw-client-panel{
  min-width:280px;max-width:360px;flex:1;
  background:var(--s2);border:1px solid var(--border);border-radius:12px;padding:10px;
}
.raw-client-filter{min-width:220px;max-width:320px;flex:1;}
.raw-client-filter .cl-btn{width:100%;min-height:38px;justify-content:space-between;}
.raw-client-filter .cl-panel{width:100%;min-width:280px;}
.raw-client-filter .cl-apply{display:flex;flex-direction:column;gap:6px;}
.raw-client-filter,.raw-filter.raw-client-filter{max-width:none;}
.raw-client-actions{display:flex;gap:8px;margin:8px 0;}
.raw-client-actions button{flex:1;}
.raw-client-list{
  max-height:150px;overflow:auto;background:#fff;border:1px solid var(--border);
  border-radius:10px;padding:6px;
}
.raw-client-meta{font-size:11px;color:var(--t3);margin-top:8px;}
.raw-action-bar{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:14px;}
.raw-summary{font-size:12px;color:var(--t2);font-weight:600;}
.raw-download-group{display:flex;gap:8px;align-items:flex-end;flex-wrap:wrap;margin-left:auto;}
@media (max-width: 900px){
  .raw-filter-wide{grid-column:span 1;}
}
.raw-table-wrap{
  border:1px solid var(--border);border-radius:12px;overflow:auto;max-height:520px;background:#fff;
}
.raw-table{width:100%;border-collapse:collapse;font-size:12px;}
.raw-table thead th{
  position:sticky;top:0;z-index:1;background:#f8f9fc;color:var(--t3);
  text-transform:uppercase;letter-spacing:.05em;font-size:10px;font-weight:700;
  text-align:left;padding:10px 12px;border-bottom:1px solid var(--border2);white-space:nowrap;
}
.raw-table tbody td{
  padding:9px 12px;border-bottom:1px solid var(--border);white-space:nowrap;vertical-align:top;
}
.raw-table tbody tr:hover td{background:#fbfbfd!important;}
.raw-empty{padding:28px 16px;text-align:center;color:var(--t3);font-size:12px;}
.tbl-w{overflow-x:auto;border-radius:var(--rlg);border:1px solid var(--border);}
table{width:100%;border-collapse:collapse;font-size:11.5px;}
.tg2{background:var(--s3);font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:var(--t3);text-align:center;padding:8px 6px;border-bottom:1px solid var(--border2);border-left:1px solid var(--border);}
.tg2:first-child{border-left:none;}
thead .sh th{font-size:10px;font-weight:600;color:var(--t3);text-align:right;padding:7px 8px;border-bottom:2px solid var(--border2);background:var(--s2);white-space:nowrap;}
thead .sh th:first-child{text-align:left;min-width:185px;padding-left:14px;}
.bl{border-left:1px solid var(--border)!important;}
.tr-r td{background:var(--surface);}
.tr-r td:first-child{padding-left:14px;}
.tr-g td{background:linear-gradient(90deg,rgba(232,69,60,.07),var(--surface));font-weight:800;font-size:12px;border-top:2px solid var(--red-bd);}
.tr-g td:first-child{padding-left:14px;}
tbody td{padding:8px;border-bottom:1px solid var(--border);text-align:right;white-space:nowrap;}
tbody td:first-child{text-align:left;}
tbody tr:hover td{background:var(--s2)!important;}
.pos{color:var(--green2);font-weight:700;}
.neg{color:var(--red2);font-weight:700;}
.dim{color:var(--t3);}
.loading{text-align:center;padding:100px 20px;color:var(--t3);}
.loading-shell{display:grid;gap:16px;animation:fadeLift .28s ease;}
.loading-hero{
  display:flex;align-items:center;justify-content:center;gap:12px;
  min-height:120px;background:linear-gradient(180deg,rgba(255,255,255,.92),rgba(247,249,252,.94));
  border:1px solid var(--border);border-radius:18px;color:var(--t3);font-size:14px;
  box-shadow:0 10px 28px rgba(16,24,40,.06);
}
.loading-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(205px,1fr));gap:14px;}
.loading-card,.loading-bar,.loading-panel{
  position:relative;overflow:hidden;border-radius:18px;border:1px solid var(--border);
  background:linear-gradient(180deg,rgba(255,255,255,.92),rgba(245,247,250,.96));
}
.loading-card{height:162px;}
.loading-bar{height:86px;border-radius:14px;}
.loading-panel{height:320px;}
.loading-card::before,.loading-bar::before,.loading-panel::before{
  content:'';position:absolute;inset:0;
  background:
    linear-gradient(90deg,transparent,rgba(255,255,255,.75),transparent),
    linear-gradient(180deg,rgba(255,255,255,.55),rgba(235,239,244,.72));
  transform:translateX(-100%);
  animation:skeletonSweep 1.4s ease-in-out infinite;
}
.loading-card::after{
  content:'';position:absolute;left:18px;right:18px;top:18px;bottom:18px;border-radius:14px;
  background:
    linear-gradient(#e9edf3,#e9edf3) left top/38% 12px no-repeat,
    linear-gradient(#eef2f6,#eef2f6) left 30px/52% 36px no-repeat,
    linear-gradient(#eef2f6,#eef2f6) left 78px/78% 10px no-repeat,
    linear-gradient(#eef2f6,#eef2f6) left 98px/64% 10px no-repeat,
    linear-gradient(#eef2f6,#eef2f6) left bottom/42% 24px no-repeat;
}
.loading-bar::after{
  content:'';position:absolute;left:16px;right:16px;top:16px;bottom:16px;border-radius:12px;
  background:
    linear-gradient(#e9edf3,#e9edf3) left top/120px 10px no-repeat,
    linear-gradient(#eef2f6,#eef2f6) left 26px/100% 26px no-repeat;
}
.loading-panel::after{
  content:'';position:absolute;left:18px;right:18px;top:18px;bottom:18px;border-radius:14px;
  background:
    linear-gradient(#e9edf3,#e9edf3) left top/180px 12px no-repeat,
    linear-gradient(#eef2f6,#eef2f6) left 30px/100% 220px no-repeat,
    linear-gradient(#eef2f6,#eef2f6) left bottom/72% 12px no-repeat;
}
@keyframes skeletonSweep{
  0%{transform:translateX(-100%);}
  100%{transform:translateX(100%);}
}
.render-fade{animation:fadeLift .34s ease;}
@keyframes fadeLift{from{opacity:.6;transform:translateY(6px);}to{opacity:1;transform:translateY(0);}}
.pill.syncing{
  position:relative;padding-left:28px;background:rgba(255,255,255,.92);color:var(--t2);
  border:1px solid var(--border2);box-shadow:0 8px 20px rgba(16,24,40,.08);display:none;
}
.pill.syncing.on{display:inline-flex;align-items:center;}
.pill.syncing::before{
  content:'';position:absolute;left:10px;top:50%;width:12px;height:12px;border-radius:50%;
  border:2px solid rgba(232,69,60,.18);border-top-color:var(--red);transform:translateY(-50%);
  animation:spin .7s linear infinite;
}
.sp{display:inline-block;width:22px;height:22px;border:2px solid var(--border2);border-top-color:var(--red);border-radius:50%;animation:spin .7s linear infinite;margin-right:8px;vertical-align:middle;}
@keyframes spin{to{transform:rotate(360deg);}}
::-webkit-scrollbar{width:6px;height:6px;}
::-webkit-scrollbar-track{background:var(--bg);}
::-webkit-scrollbar-thumb{background:var(--s4);border-radius:3px;}
::-webkit-scrollbar-thumb:hover{background:var(--grey);}
.tr-d td{background:linear-gradient(90deg,rgba(232,69,60,.08),var(--surface));color:var(--red2);font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.08em;padding:6px 14px;border-bottom:1px solid var(--border);}
.tr-t td{background:var(--s2);font-weight:800;border-top:1px solid var(--border2);}
.tr-t td:first-child{padding-left:14px;color:var(--grey);}
</style>
</head>
<body>

<nav class="nav">
  <div class="brand">
    <img src="logo.jpg" style="height:38px;width:auto;object-fit:contain;">
    <span class="pill" id="pbadge">Loading…</span>
    <span class="pill syncing" id="syncStatus">Updating...</span>
    <span class="pill" id="lastUpdated" style="margin-left:8px;background:var(--grey-lt);color:var(--grey)">--</span>
  </div>
  <div class="nav-r">
    <span class="nlbl">Year</span>
    <div class="cl-wrap" id="yearWrap">
      <button class="cl-btn" onclick="toggleYearPanel(event)">Select Years&nbsp;<span class="cl-ct" id="yearCt">0</span>&nbsp;&#9662;</button>
      <div class="cl-panel" id="yearPanel">
        <div class="cl-search"><input id="yearQ" placeholder="Search years..." oninput="filterYearPanel()"></div>
        <div class="cl-actions"><button onclick="selectAllYears()">Select All</button><button onclick="clearAllYears()">Clear All</button></div>
        <div class="cl-list" id="yearList"></div>
        <div class="cl-apply"><button onclick="applyYear()">Apply</button></div>
      </div>
    </div>
    <span class="nlbl">Month</span>
    <div class="cl-wrap" id="monthWrap">
      <button class="cl-btn" onclick="toggleMonthPanel(event)">Select Months&nbsp;<span class="cl-ct" id="monthCt">0</span>&nbsp;&#9662;</button>
      <div class="cl-panel" id="monthPanel">
        <div class="cl-search"><input id="monthQ" placeholder="Search months..." oninput="filterMonthPanel()"></div>
        <div class="cl-actions"><button onclick="selectAllMonths()">Select All</button><button onclick="clearAllMonths()">Clear All</button></div>
        <div class="cl-list" id="monthList"></div>
        <div class="cl-apply"><button onclick="applyMonth()">Apply</button></div>
      </div>
    </div>
    <span class="nlbl">Clients</span>
    <div class="cl-wrap" id="clWrap">
      <button class="cl-btn" onclick="togglePanel(event)">Select Clients&nbsp;<span class="cl-ct" id="clCt">0</span>&nbsp;&#9662;</button>
      <div class="cl-panel" id="clPanel">
        <div class="cl-search"><input id="clQ" placeholder="Search clients…" oninput="filterCl()"></div>
        <div class="cl-actions"><button onclick="selectAll()">Select All</button><button onclick="clearAll()">Clear All</button></div>
        <div class="cl-list" id="clList"></div>
        <div class="cl-apply"><button onclick="applyCl()">Apply</button></div>
      </div>
    </div>
    <span class="nlbl">Domain</span>
    <div class="cl-wrap" id="domWrap">
      <button class="cl-btn" onclick="toggleDomPanel(event)">Select Domain&nbsp;<span class="cl-ct" id="domCt">0</span>&nbsp;&#9662;</button>
      <div class="cl-panel" id="domPanel">
        <div class="cl-search"><input id="domQ" placeholder="Search domain…" oninput="filterDomPanel()"></div>
        <div class="cl-actions"><button onclick="selectAllDom()">Select All</button><button onclick="clearAllDom()">Clear All</button></div>
        <div class="cl-list" id="domList"></div>
        <div class="cl-apply"><button onclick="applyDom()">Apply</button></div>
      </div>
    </div>
    <span class="nlbl">BH</span>
    <div class="cl-wrap" id="bhWrap">
      <button class="cl-btn" onclick="toggleBhPanel(event)">Select BH&nbsp;<span class="cl-ct" id="bhCt">0</span>&nbsp;&#9662;</button>
      <div class="cl-panel" id="bhPanel">
        <div class="cl-search"><input id="bhQ" placeholder="Search BH…" oninput="filterBhPanel()"></div>
        <div class="cl-actions"><button onclick="selectAllBh()">Select All</button><button onclick="clearAllBh()">Clear All</button></div>
        <div class="cl-list" id="bhList"></div>
        <div class="cl-apply"><button onclick="applyBh()">Apply</button></div>
      </div>
    </div>
    <button class="btn-g" onclick="resetAll()">Reset</button>
    <button class="btn-r" onclick="refresh()">&#8635; Refresh</button>
  </div>
</nav>

<div class="body" id="body">
  <div class="loading-shell">
    <div class="loading-hero"><span class="sp"></span><span>Loading dashboard...</span></div>
    <div class="loading-grid">
      <div class="loading-card"></div>
      <div class="loading-card"></div>
      <div class="loading-card"></div>
      <div class="loading-card"></div>
      <div class="loading-card"></div>
    </div>
    <div class="loading-bar"></div>
    <div class="loading-panel"></div>
  </div>
</div>

<script>
const MON=['','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
let charts={};
let allClients=[];
let allClientMeta=[];
let selectedClients=new Set();
let selectedDomains=new Set();
let selectedBHs=new Set();
let selectedYears=new Set();
let selectedMonths=new Set();
let rawSelectedClients=new Set();
let _rawDataset='demand';

let _dayFromDate = '';
let _dayToDate   = '';
let _dayRangePreset = 7;
let _monthFromDate = '';
let _monthToDate   = '';
let _monthRangePreset = 12;
let _dailyTab = 'dem';
let _trendGrain = 'day';
let _dayTrendTab = 'dem';
let _monthTrendTab = 'dem';
let _pipelineCollapsed = true;
let _clientBreakdownCollapsed = true;
let _rawExplorerCollapsed = true;
(function initDates(){
  const today = new Date();
  const dayPast  = new Date();
  dayPast.setDate(today.getDate() - 7);
  _dayFromDate = dayPast.toISOString().split('T')[0];
  _dayToDate   = today.toISOString().split('T')[0];

  const monthPast = new Date(today.getFullYear(), today.getMonth() - 11, 1);
  _monthFromDate = monthPast.toISOString().split('T')[0];
  _monthToDate   = today.toISOString().split('T')[0];
})();

async function init(){
  try{
    const r=await fetch('/api/init');
    if(!r.ok) throw new Error('Failed to load dashboard filters');
    const {years,months,clients,domains,business_heads,client_meta}=await r.json();
    allClients=clients;
    allClientMeta=client_meta||[];
    const yearValues=(years||[]).map(y=>String(y));
    buildList('yearList', yearValues, selectedYears, 'year');
    refreshClientList();
    buildList('domList', domains||[], selectedDomains, 'dom');
    buildList('bhList',  business_heads||[], selectedBHs, 'bh');
    if(yearValues.length){
      const now=new Date();
      const currentYear=String(now.getFullYear());
      selectedYears=new Set([yearValues.includes(currentYear) ? currentYear : yearValues[0]]);
      syncFilterChecks('yearList', selectedYears);
      updateFilterCount('yearCt', selectedYears);
      await populateMonths(true);
    } else {
      await loadAll();
    }
  }catch(err){
    console.error(err);
    document.getElementById('body').innerHTML='<div class="loading">Unable to load dashboard. Please refresh and check the Flask console for API errors.</div>';
  }
}

async function populateMonths(setLatest){
  try{
    const queryYears=[...selectedYears];
    const r=await fetch('/api/months'+(queryYears.length?'?years='+queryYears.join(','):''));
    if(!r.ok) throw new Error('Failed to load months');
    const {months}=await r.json();
    const monthValues=(months||[]).map(m=>String(m));
    const prev=[...selectedMonths];
    buildMonthList(monthValues);
    if(setLatest&&monthValues.length){
      const now=new Date();
      const currentYear=String(now.getFullYear());
      const currentMonth=String(now.getMonth()+1);
      if(selectedYears.size===1 && selectedYears.has(currentYear) && monthValues.includes(currentMonth)) selectedMonths=new Set([currentMonth]);
      else selectedMonths=new Set([monthValues[monthValues.length-1]]);
    }
    else selectedMonths=new Set(prev.filter(m=>monthValues.includes(m)));
    syncFilterChecks('monthList', selectedMonths);
    updateFilterCount('monthCt', selectedMonths);
    await loadAll();
  }catch(err){
    console.error(err);
    document.getElementById('body').innerHTML='<div class="loading">Unable to load months. Please refresh and check the Flask console for API errors.</div>';
  }
}

function updateFilterCount(countId, selectedSet){
  const ct=document.getElementById(countId);
  if(!ct) return;
  if(selectedSet.size>0){
    ct.textContent=selectedSet.size;
    ct.style.display='inline';
  }else{
    ct.style.display='none';
  }
}

function syncFilterChecks(listId, selectedSet){
  document.querySelectorAll(`#${listId} input`).forEach(cb=>{cb.checked=selectedSet.has(cb.value);});
}

function buildMonthList(items){
  document.getElementById('monthList').innerHTML=items.map(v=>`
    <div class="cl-item" data-n="${MON[parseInt(v,10)].toLowerCase()} ${v}">
      <input type="checkbox" id="month_${encodeURIComponent(v)}" value="${v}" ${selectedMonths.has(v)?'checked':''}>
      <label for="month_${encodeURIComponent(v)}">${MON[parseInt(v,10)]}</label>
    </div>`).join('');
}

function buildCl(clients){
  document.getElementById('clList').innerHTML=clients.map(c=>`
    <div class="cl-item" data-n="${c.toLowerCase()}">
      <input type="checkbox" id="cb_${encodeURIComponent(c)}" value="${c}" ${selectedClients.has(c)?'checked':''}>
      <label for="cb_${encodeURIComponent(c)}">${c}</label>
    </div>`).join('');
}

function refreshClientList(){
  const visibleClients=(allClientMeta.length?allClientMeta:allClients.map(c=>({name:c,domain:'',bh:''})))
    .filter(c=>{
      const domOk=selectedDomains.size===0 || selectedDomains.has(c.domain||'');
      const bhOk=selectedBHs.size===0 || selectedBHs.has(c.bh||'');
      return domOk && bhOk;
    })
    .map(c=>c.name)
    .sort((a,b)=>a.localeCompare(b, undefined, {sensitivity:'base'}));

  selectedClients=new Set([...selectedClients].filter(c=>visibleClients.includes(c)));
  updateFilterCount('clCt', selectedClients);
  buildCl(visibleClients);
}

function buildList(listId, items, selectedSet, prefix){
  document.getElementById(listId).innerHTML=items.map(v=>`
    <div class="cl-item" data-n="${v.toLowerCase()}">
      <input type="checkbox" id="${prefix}_${encodeURIComponent(v)}" value="${v}" ${selectedSet.has(v)?'checked':''}>
      <label for="${prefix}_${encodeURIComponent(v)}">${v}</label>
    </div>`).join('');
}

function esc(value){
  return String(value ?? '')
    .replace(/&/g,'&amp;')
    .replace(/</g,'&lt;')
    .replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;')
    .replace(/'/g,'&#39;');
}

function getRawClientPool(){
  return (allClientMeta.length ? allClientMeta.map(c=>c.name) : allClients)
    .slice()
    .sort((a,b)=>a.localeCompare(b, undefined, {sensitivity:'base'}));
}

function buildRawClientList(){
  const list=document.getElementById('rawClientList');
  if(!list) return;
  const items=getRawClientPool();
  list.innerHTML=items.map(name=>`
    <div class="cl-item" data-n="${esc(name.toLowerCase())}">
      <input type="checkbox" id="raw_${encodeURIComponent(name)}" value="${esc(name)}" ${rawSelectedClients.has(name)?'checked':''}>
      <label for="raw_${encodeURIComponent(name)}">${esc(name)}</label>
    </div>`).join('') || '<div class="raw-empty">No clients match this search.</div>';
  updateRawClientMeta();
}

function filterRawClientList(){
  const q=(document.getElementById('rawClientSearch')?.value || '').toLowerCase();
  document.querySelectorAll('#rawClientList .cl-item').forEach(el=>{
    el.style.display=el.dataset.n.includes(q)?'flex':'none';
  });
}

function updateRawClientMeta(){
  const meta=document.getElementById('rawClientMeta');
  if(!meta) return;
  meta.textContent = rawSelectedClients.size
    ? `${rawSelectedClients.size} client${rawSelectedClients.size===1?'':'s'} selected`
    : 'All clients included';
  updateFilterCount('rawClientCt', rawSelectedClients);
}

function collectRawClientSelection(){
  const checked=[...document.querySelectorAll('#rawClientList input:checked')].map(cb=>cb.value);
  rawSelectedClients=new Set(checked);
  updateRawClientMeta();
  updateRawDownloadLinks();
}

function selectAllRawClients(){
  document.querySelectorAll('#rawClientList input').forEach(cb=>cb.checked=true);
  collectRawClientSelection();
}

function clearRawClients(){
  document.querySelectorAll('#rawClientList input').forEach(cb=>cb.checked=false);
  collectRawClientSelection();
}

function toggleRawClientPanel(e){
  e.stopPropagation();
  document.getElementById('rawClientPanel').classList.toggle('open');
}

function applyRawClients(){
  collectRawClientSelection();
  document.getElementById('rawClientPanel').classList.remove('open');
  loadRawData();
}

function syncRawMonthWithDates(){
  const monthEl=document.getElementById('rawMonthFilter');
  const fromEl=document.getElementById('rawFromDate');
  const toEl=document.getElementById('rawToDate');
  if(!monthEl || !fromEl || !toEl) return;
  const fromDate=fromEl.value;
  const toDate=toEl.value;
  if(!fromDate || !toDate){
    monthEl.value='';
    return;
  }
  const fromParts=fromDate.split('-').map(Number);
  const toParts=toDate.split('-').map(Number);
  if(fromParts.length!==3 || toParts.length!==3){
    monthEl.value='';
    return;
  }
  const [fromYear, fromMonth, fromDay]=fromParts;
  const [toYear, toMonth, toDay]=toParts;
  const lastDay=new Date(fromYear, fromMonth, 0).getDate();
  monthEl.value=(fromYear===toYear && fromMonth===toMonth && fromDay===1 && toDay===lastDay)
    ? `${fromYear}-${String(fromMonth).padStart(2,'0')}`
    : '';
}

function syncRawDemandPills(){
  const activeStatus=document.getElementById('rawDemandStatus')?.value || 'all';
  document.querySelectorAll('.raw-demand-pill').forEach(btn=>{
    btn.classList.toggle('active', btn.dataset.value===activeStatus);
  });
}

function setRawMonthDates(){
  const monthEl=document.getElementById('rawMonthFilter');
  const fromEl=document.getElementById('rawFromDate');
  const toEl=document.getElementById('rawToDate');
  if(!monthEl || !fromEl || !toEl || !monthEl.value) return;
  const [year, month]=monthEl.value.split('-').map(Number);
  const monthStart=`${year}-${String(month).padStart(2,'0')}-01`;
  const monthEnd=new Date(year, month, 0).toISOString().split('T')[0];
  fromEl.value=monthStart;
  toEl.value=monthEnd;
  updateRawDownloadLinks();
}

function handleRawDateChange(){
  const fromEl=document.getElementById('rawFromDate');
  const toEl=document.getElementById('rawToDate');
  if(fromEl && toEl && fromEl.value && toEl.value && fromEl.value>toEl.value){
    toEl.value=fromEl.value;
  }
  syncRawMonthWithDates();
  loadRawData();
}

function setRawCurrentMonth(){
  const today=new Date();
  const monthValue=`${today.getFullYear()}-${String(today.getMonth()+1).padStart(2,'0')}`;
  const monthEl=document.getElementById('rawMonthFilter');
  if(monthEl) monthEl.value=monthValue;
  setRawMonthDates();
  loadRawData();
}

function setRawDatePreset(days){
  const toEl=document.getElementById('rawToDate');
  const fromEl=document.getElementById('rawFromDate');
  if(!fromEl || !toEl) return;
  const today=new Date();
  const start=new Date();
  start.setDate(today.getDate()-(Number(days)-1));
  fromEl.value=start.toISOString().split('T')[0];
  toEl.value=today.toISOString().split('T')[0];
  syncRawMonthWithDates();
  loadRawData();
}

function setRawDemandStatus(status){
  const demandStatusEl=document.getElementById('rawDemandStatus');
  if(!demandStatusEl) return;
  demandStatusEl.value=status || 'all';
  syncRawDemandPills();
  loadRawData();
}

function resetRawFilters(){
  const monthEl=document.getElementById('rawMonthFilter');
  const fromEl=document.getElementById('rawFromDate');
  const toEl=document.getElementById('rawToDate');
  const demandStatusEl=document.getElementById('rawDemandStatus');
  const searchEl=document.getElementById('rawClientSearch');
  if(monthEl) monthEl.value='';
  if(fromEl) fromEl.value='';
  if(toEl) toEl.value='';
  if(demandStatusEl) demandStatusEl.value='all';
  rawSelectedClients=new Set();
  if(searchEl) searchEl.value='';
  buildRawClientList();
  filterRawClientList();
  updateRawClientMeta();
  syncRawDemandPills();
  updateRawDownloadLinks();
  loadRawData();
}

function getRawParams(){
  const p=new URLSearchParams();
  p.set('dataset', _rawDataset);
  const monthValue=document.getElementById('rawMonthFilter')?.value || '';
  if(monthValue){
    const [year, month]=monthValue.split('-');
    if(year) p.set('year', year);
    if(month) p.set('month', String(parseInt(month,10)));
  }
  const fromDate=document.getElementById('rawFromDate')?.value || '';
  const toDate=document.getElementById('rawToDate')?.value || '';
  if(fromDate) p.set('from', fromDate);
  if(toDate) p.set('to', toDate);
  const demandStatus=document.getElementById('rawDemandStatus')?.value || 'all';
  if(_rawDataset==='demand' && demandStatus!=='all') p.set('demand_status', demandStatus);
  if(rawSelectedClients.size) p.set('raw_clients', [...rawSelectedClients].join(','));
  return p.toString();
}

function getRawExportParams(format){
  const p=new URLSearchParams(getRawParams());
  const columnMode=document.getElementById('rawColumnMode')?.value || 'visible';
  p.set('columns', columnMode);
  p.set('format', format || 'csv');
  return p.toString();
}

function updateRawDownloadLinks(){
  const csvBtn=document.getElementById('rawDownloadCsvBtn');
  const xlsxBtn=document.getElementById('rawDownloadXlsxBtn');
  if(csvBtn) csvBtn.href='/api/raw_data_export?'+getRawExportParams('csv');
  if(xlsxBtn) xlsxBtn.href='/api/raw_data_export?'+getRawExportParams('xlsx');
}

function setRawDataset(dataset, btn){
  _rawDataset=dataset;
  const tabWrap=document.getElementById('rawDatasetTabs');
  if(tabWrap){
    tabWrap.querySelectorAll('.ttab').forEach(tab=>tab.classList.remove('on'));
  }
  if(btn) btn.classList.add('on');
  const demandOnlyWrap=document.getElementById('rawDemandOnlyWrap');
  if(demandOnlyWrap) demandOnlyWrap.style.display = dataset==='demand' ? 'flex' : 'none';
  syncRawDemandPills();
  loadRawData();
}

function renderRawTable(columns, rows){
  const mount=document.getElementById('rawDataTable');
  if(!mount) return;
  if(!rows.length){
    mount.innerHTML='<div class="raw-empty">No raw records found for the selected filters.</div>';
    return;
  }
  const head=columns.map(col=>`<th>${esc(col)}</th>`).join('');
  const body=rows.map(row=>`<tr>${columns.map(col=>`<td>${esc(row[col] || '')}</td>`).join('')}</tr>`).join('');
  mount.innerHTML=`<div class="raw-table-wrap"><table class="raw-table"><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table></div>`;
}

function hydrateRawSection(){
  if(!document.getElementById('rawClientList')) return;
  buildRawClientList();
  updateRawClientMeta();
  const demandOnlyWrap=document.getElementById('rawDemandOnlyWrap');
  if(demandOnlyWrap) demandOnlyWrap.style.display = _rawDataset==='demand' ? 'flex' : 'none';
  syncRawDemandPills();
  updateRawDownloadLinks();
}

async function loadRawData(){
  const mount=document.getElementById('rawDataTable');
  if(!mount) return;
  const summary=document.getElementById('rawSummary');
  mount.innerHTML='<div class="raw-empty">Loading raw data...</div>';
  if(summary) summary.textContent='Fetching records...';
  try{
    const res=await fetch('/api/raw_data?'+getRawParams());
    if(!res.ok) throw new Error('Failed to load raw data');
    const data=await res.json();
    renderRawTable(data.columns||[], data.rows||[]);
    if(summary) summary.textContent=`${data.label || 'Raw Data'}: ${Number(data.row_count||0).toLocaleString()} row(s)`;
    updateRawDownloadLinks();
  }catch(err){
    console.error(err);
    if(summary) summary.textContent='Raw data could not be loaded';
    mount.innerHTML='<div class="raw-empty">Raw data could not be loaded. Please try again.</div>';
  }
}

function toggleYearPanel(e){e.stopPropagation();document.getElementById('yearPanel').classList.toggle('open');}
function filterYearPanel(){const q=document.getElementById('yearQ').value.toLowerCase();document.querySelectorAll('#yearList .cl-item').forEach(el=>{el.style.display=el.dataset.n.includes(q)?'flex':'none';});}
function selectAllYears(){document.querySelectorAll('#yearList input').forEach(cb=>cb.checked=true);}
function clearAllYears(){document.querySelectorAll('#yearList input').forEach(cb=>cb.checked=false);}
async function applyYear(){
  selectedYears=new Set([...document.querySelectorAll('#yearList input:checked')].map(cb=>cb.value));
  updateFilterCount('yearCt', selectedYears);
  document.getElementById('yearPanel').classList.remove('open');
  await populateMonths(false);
}

function toggleMonthPanel(e){e.stopPropagation();document.getElementById('monthPanel').classList.toggle('open');}
function filterMonthPanel(){const q=document.getElementById('monthQ').value.toLowerCase();document.querySelectorAll('#monthList .cl-item').forEach(el=>{el.style.display=el.dataset.n.includes(q)?'flex':'none';});}
function selectAllMonths(){document.querySelectorAll('#monthList input').forEach(cb=>cb.checked=true);}
function clearAllMonths(){document.querySelectorAll('#monthList input').forEach(cb=>cb.checked=false);}
function applyMonth(){
  selectedMonths=new Set([...document.querySelectorAll('#monthList input:checked')].map(cb=>cb.value));
  updateFilterCount('monthCt', selectedMonths);
  document.getElementById('monthPanel').classList.remove('open');
  loadAll();
}

function toggleDomPanel(e){e.stopPropagation();document.getElementById('domPanel').classList.toggle('open');}
function filterDomPanel(){const q=document.getElementById('domQ').value.toLowerCase();document.querySelectorAll('#domList .cl-item').forEach(el=>{el.style.display=el.dataset.n.includes(q)?'flex':'none';});}
function selectAllDom(){document.querySelectorAll('#domList input').forEach(cb=>cb.checked=true);}
function clearAllDom(){document.querySelectorAll('#domList input').forEach(cb=>cb.checked=false);}
function applyDom(){
  selectedDomains=new Set([...document.querySelectorAll('#domList input:checked')].map(cb=>cb.value));
  updateFilterCount('domCt', selectedDomains);
  document.getElementById('domPanel').classList.remove('open');
  refreshClientList();
  loadAll();
}
function toggleBhPanel(e){e.stopPropagation();document.getElementById('bhPanel').classList.toggle('open');}
function filterBhPanel(){const q=document.getElementById('bhQ').value.toLowerCase();document.querySelectorAll('#bhList .cl-item').forEach(el=>{el.style.display=el.dataset.n.includes(q)?'flex':'none';});}
function selectAllBh(){document.querySelectorAll('#bhList input').forEach(cb=>cb.checked=true);}
function clearAllBh(){document.querySelectorAll('#bhList input').forEach(cb=>cb.checked=false);}
function applyBh(){
  selectedBHs=new Set([...document.querySelectorAll('#bhList input:checked')].map(cb=>cb.value));
  updateFilterCount('bhCt', selectedBHs);
  document.getElementById('bhPanel').classList.remove('open');
  refreshClientList();
  loadAll();
}
function filterCl(){const q=document.getElementById('clQ').value.toLowerCase();document.querySelectorAll('#clList .cl-item').forEach(el=>{el.style.display=el.dataset.n.includes(q)?'flex':'none';});}
function togglePanel(e){e.stopPropagation();document.getElementById('clPanel').classList.toggle('open');}
function selectAll(){document.querySelectorAll('#clList input').forEach(cb=>cb.checked=true);}
function clearAll(){document.querySelectorAll('#clList input').forEach(cb=>cb.checked=false);}
function applyCl(){
  selectedClients=new Set([...document.querySelectorAll('#clList input:checked')].map(cb=>cb.value));
  updateFilterCount('clCt', selectedClients);
  document.getElementById('clPanel').classList.remove('open');
  loadAll();
}
document.addEventListener('click',e=>{
  if(!document.getElementById('yearWrap').contains(e.target)) document.getElementById('yearPanel').classList.remove('open');
  if(!document.getElementById('monthWrap').contains(e.target)) document.getElementById('monthPanel').classList.remove('open');
  if(!document.getElementById('clWrap').contains(e.target))  document.getElementById('clPanel').classList.remove('open');
  if(!document.getElementById('domWrap').contains(e.target)) document.getElementById('domPanel').classList.remove('open');
  if(!document.getElementById('bhWrap').contains(e.target))  document.getElementById('bhPanel').classList.remove('open');
  if(document.getElementById('rawClientWrap') && !document.getElementById('rawClientWrap').contains(e.target)) document.getElementById('rawClientPanel').classList.remove('open');
});
function resetAll(){
  selectedYears=new Set(); updateFilterCount('yearCt', selectedYears); document.querySelectorAll('#yearList input').forEach(cb=>cb.checked=false);
  selectedMonths=new Set(); updateFilterCount('monthCt', selectedMonths); document.querySelectorAll('#monthList input').forEach(cb=>cb.checked=false);
  selectedClients=new Set(); updateFilterCount('clCt', selectedClients); document.querySelectorAll('#clList input').forEach(cb=>cb.checked=false);
  selectedDomains=new Set(); updateFilterCount('domCt', selectedDomains); document.querySelectorAll('#domList input').forEach(cb=>cb.checked=false);
  selectedBHs=new Set(); updateFilterCount('bhCt', selectedBHs);  document.querySelectorAll('#bhList input').forEach(cb=>cb.checked=false);
  refreshClientList();
  loadAll();
}

async function refresh(){
  await fetch('/api/refresh');
  document.getElementById('body').innerHTML='<div class="loading"><span class="sp"></span>Refreshing data…</div>';
  loadAll();
}

function getActiveTrendFromDate(mode){
  mode = mode || _trendGrain;
  return mode==='month' ? _monthFromDate : _dayFromDate;
}

function getActiveTrendToDate(mode){
  mode = mode || _trendGrain;
  return mode==='month' ? _monthToDate : _dayToDate;
}

function getActiveTrendPreset(mode){
  mode = mode || _trendGrain;
  return mode==='month' ? _monthRangePreset : _dayRangePreset;
}

function setActiveTrendDates(mode, fromDate, toDate){
  if (toDate === undefined) {
    toDate = fromDate;
    fromDate = mode;
    mode = _trendGrain;
  }
  if(mode==='month'){
    _monthFromDate = fromDate;
    _monthToDate = toDate;
  }else{
    _dayFromDate = fromDate;
    _dayToDate = toDate;
  }
}

function setActiveTrendPreset(mode, value){
  if (value === undefined) {
    value = mode;
    mode = _trendGrain;
  }
  if(mode==='month'){
    _monthRangePreset = value;
  }else{
    _dayRangePreset = value;
  }
}

function clearActiveTrendPreset(mode){
  mode = mode || _trendGrain;
  setActiveTrendPreset(mode, null);
}

function getParams(yr, mo){
  const p=new URLSearchParams();
  const years=yr ? [yr] : [...selectedYears];
  const months=mo ? [mo] : [...selectedMonths];
  if(years.length)p.set('years',years.join(','));
  if(months.length)p.set('months',months.join(','));
  if(selectedClients.size>0)p.set('clients',[...selectedClients].join(','));
  if(selectedDomains.size>0)p.set('domains',[...selectedDomains].join(','));
  if(selectedBHs.size>0)p.set('bhs',[...selectedBHs].join(','));
  return p;
}

function trendParams(mode){
  const p=new URLSearchParams();
  if(selectedClients.size>0)p.set('clients',[...selectedClients].join(','));
  if(selectedDomains.size>0)p.set('domains',[...selectedDomains].join(','));
  if(selectedBHs.size>0)p.set('bhs',[...selectedBHs].join(','));
  const fromDate=getActiveTrendFromDate(mode);
  const toDate=getActiveTrendToDate(mode);
  if(fromDate) p.set('from', fromDate);
  if(toDate)   p.set('to',   toDate);
  p.set('grain', mode);
  return p;
}

async function loadAll(){
  try{
    showBodyLoading();
    const dayFromEl = document.getElementById('fromDateDay');
    const dayToEl   = document.getElementById('toDateDay');
    if(dayFromEl && dayToEl){
      const currentFrom=getActiveTrendFromDate('day');
      const currentTo=getActiveTrendToDate('day');
      if(dayFromEl.value && dayFromEl.value!==currentFrom) setActiveTrendDates('day', dayFromEl.value, currentTo);
      if(dayToEl.value && dayToEl.value!==currentTo) setActiveTrendDates('day', getActiveTrendFromDate('day'), dayToEl.value);
      dayFromEl.value = getActiveTrendFromDate('day');
      dayToEl.value   = getActiveTrendToDate('day');
    }
    const monthFromEl = document.getElementById('fromDateMonth');
    const monthToEl   = document.getElementById('toDateMonth');
    if(monthFromEl && monthToEl){
      const currentFrom=getActiveTrendFromDate('month');
      const currentTo=getActiveTrendToDate('month');
      if(monthFromEl.value && monthFromEl.value!==currentFrom) setActiveTrendDates('month', monthFromEl.value, currentTo);
      if(monthToEl.value && monthToEl.value!==currentTo) setActiveTrendDates('month', getActiveTrendFromDate('month'), monthToEl.value);
      monthFromEl.value = getActiveTrendFromDate('month');
      monthToEl.value   = getActiveTrendToDate('month');
    }
    updateLastUpdated();
    const years=[...selectedYears].sort();
    const months=[...selectedMonths].sort((a,b)=>parseInt(a,10)-parseInt(b,10));
    const singleYear=years.length===1?years[0]:'';
    const singleMonth=months.length===1?months[0]:'';
    const lbl=singleYear&&singleMonth
      ? MON[parseInt(singleMonth,10)]+' '+singleYear
      : singleYear
        ? singleYear
        : singleMonth
          ? MON[parseInt(singleMonth,10)]
          : years.length||months.length
            ? `${years.length||'All'}Y / ${months.length||'All'}M`
            : 'All Periods';
    document.getElementById('pbadge').textContent=lbl;

    let pyr='', pmo='';
    if(singleYear && singleMonth){
      let pm=parseInt(singleMonth)-1, py=parseInt(singleYear);
      if(pm<1){pm=12;py=py-1;}
      pmo=String(pm); pyr=String(py);
    }

    const [dRes, tRes, pRes, dailyDayRes, dailyMonthRes, lmtdRes]=await Promise.all([
      fetch('/api/data?'+getParams(singleYear,singleMonth)),
      fetch('/api/trends?'+getParams()),
      pyr&&pmo ? fetch('/api/data?'+getParams(pyr,pmo)) : Promise.resolve(null),
      fetch('/api/daily_trends?'+trendParams('day')),
      fetch('/api/daily_trends?'+trendParams('month')),
      singleYear&&singleMonth ? fetch('/api/lmtd?'+getParams(singleYear,singleMonth)) : Promise.resolve(null)
    ]);
    if(!dRes.ok || !tRes.ok || !dailyDayRes.ok || !dailyMonthRes.ok || (pRes && !pRes.ok) || (lmtdRes && !lmtdRes.ok)){
      throw new Error('One or more dashboard API requests failed');
    }

    const {rows,grand:g}=await dRes.json();
    const trends=await tRes.json();
    const dailyDay=await dailyDayRes.json();
    const dailyMonth=await dailyMonthRes.json();
    const prevData=pRes ? await pRes.json() : null;
    const pg=prevData ? prevData.grand : null;
    const lmtdData=lmtdRes ? await lmtdRes.json() : null;

    render(rows, g, trends, pg, pmo, pyr, dailyDay, dailyMonth, singleYear, singleMonth, lmtdData);
  }catch(err){
    console.error(err);
    document.getElementById('body').innerHTML='<div class="loading">Dashboard data could not be loaded. Please refresh and check the Flask console for API errors.</div>';
  }finally{
    hideBodyLoading();
  }
}

const L=v=>parseFloat(v).toFixed(2);
const fN=v=>v===0?'<span class="dim">—</span>':parseInt(v).toLocaleString();
const f2=v=>parseFloat(v)===0?'<span class="dim">—</span>':parseFloat(v).toFixed(2);
const pn=(v,isL)=>{const neg=parseFloat(v)<0,a=Math.abs(parseFloat(v));const s=isL?'&#8377;'+L(a)+'L':Math.round(a).toLocaleString();return`<span class="${neg?'neg':'pos'}">${neg?'&minus;':'+'}${s}</span>`;};
const pc=(a,b)=>b?Math.round(a/b*100):0;
const tagG=s=>`<span class="ktag tg">${s}</span>`;
const tagR=s=>`<span class="ktag tr">${s}</span>`;
const tagBl=s=>`<span class="ktag tb">${s}</span>`;
const tagGr=s=>`<span class="ktag tgr">${s}</span>`;
function dc(id){if(charts[id]){charts[id].destroy();delete charts[id];}}
function showBodyLoading(){
  const status=document.getElementById('syncStatus');
  if(status) status.classList.add('on');
}
function hideBodyLoading(){
  const status=document.getElementById('syncStatus');
  if(status) status.classList.remove('on');
}
function formatAnimatedValue(value, format, signed){
  const n=Number(value)||0;
  if(format==='currency'){
    const abs=Math.abs(n).toFixed(2);
    const sign=signed?(n>=0?'+':'-'):'';
    return `${sign}₹${abs}L`;
  }
  const rounded=Math.round(Math.abs(n)).toLocaleString();
  if(signed){
    return `${n>=0?'+':'-'}${rounded}`;
  }
  return Math.round(n).toLocaleString();
}
function animatedKpiValue(value, color, format='number', signed=false, fontSize=''){
  const style=`color:${color}${fontSize?`;font-size:${fontSize}`:''}`;
  return `<div class="kpi-val" style="${style}" data-animate="true" data-target="${Number(value)||0}" data-format="${format}" data-signed="${signed?'true':'false'}">${formatAnimatedValue(value, format, signed)}</div>`;
}
function animateKpiValues(){
  const els=[...document.querySelectorAll('.kpi-val[data-animate="true"]')];
  if(!els.length) return;
  if(window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches){
    els.forEach(el=>{el.textContent=formatAnimatedValue(el.dataset.target, el.dataset.format, el.dataset.signed==='true');});
    return;
  }
  const duration=700;
  const ease=t=>1-Math.pow(1-t,3);
  els.forEach(el=>{
    const target=Number(el.dataset.target)||0;
    const format=el.dataset.format||'number';
    const signed=el.dataset.signed==='true';
    const start=performance.now();
    el.classList.add('animating');
    const tick=now=>{
      const p=Math.min((now-start)/duration,1);
      const current=target*ease(p);
      el.textContent=formatAnimatedValue(current, format, signed);
      if(p<1){
        requestAnimationFrame(tick);
      }else{
        el.textContent=formatAnimatedValue(target, format, signed);
        el.classList.remove('animating');
      }
    };
    requestAnimationFrame(tick);
  });
}
function togglePipelineSection(){
  _pipelineCollapsed=!_pipelineCollapsed;
  const body=document.getElementById('pipelineBody');
  const btn=document.getElementById('pipelineToggle');
  if(body) body.classList.toggle('collapsed', _pipelineCollapsed);
  if(btn) btn.classList.toggle('collapsed', _pipelineCollapsed);
}

function toggleClientBreakdownSection(){
  _clientBreakdownCollapsed=!_clientBreakdownCollapsed;
  const body=document.getElementById('clientBreakdownBody');
  const btn=document.getElementById('clientBreakdownToggle');
  if(body) body.classList.toggle('collapsed', _clientBreakdownCollapsed);
  if(btn) btn.classList.toggle('collapsed', _clientBreakdownCollapsed);
}

function toggleRawExplorerSection(){
  _rawExplorerCollapsed=!_rawExplorerCollapsed;
  const body=document.getElementById('rawExplorerBody');
  const btn=document.getElementById('rawExplorerToggle');
  if(body) body.classList.toggle('collapsed', _rawExplorerCollapsed);
  if(btn) btn.classList.toggle('collapsed', _rawExplorerCollapsed);
}

function buildCompareBar(title, curMetrics, prevMetrics){
  const metrics=[
    ['Demands',curMetrics.dem,prevMetrics.dem,false],['Submissions',curMetrics.sub,prevMetrics.sub,false],
    ['L1',curMetrics.l1,prevMetrics.l1,false],['L2',curMetrics.l2,prevMetrics.l2,false],['L3',curMetrics.l3,prevMetrics.l3,false],
    ['Selections',curMetrics.sel,prevMetrics.sel,false],['Onboarded',curMetrics.ob_hc,prevMetrics.ob_hc,false],
    ['Exits',curMetrics.ex_hc,prevMetrics.ex_hc,false],['Net HC',curMetrics.net_hc,prevMetrics.net_hc,false],
    ['Net PO',curMetrics.net_po,prevMetrics.net_po,true],['Net Margin',curMetrics.net_mg,prevMetrics.net_mg,true],
  ];
  const items=metrics.map(([lbl,cur,prev,isL],i)=>{
    const diff=parseFloat(cur)-parseFloat(prev);
    const pct=parseFloat(prev)!==0?Math.round(Math.abs(diff)/Math.abs(parseFloat(prev))*100):0;
    const clr=diff===0?'var(--t3)':diff>0?'#1db85a':'#e8453c';
    const arrow=diff===0?'&mdash;':(diff>0?'&#9650; ':'&#9660; ')+pct+'%';
    const valStr=isL?'&#8377;'+L(Math.abs(parseFloat(prev)))+'L':Math.round(prev).toLocaleString();
    return `${i>0?'<div class="prev-sep"></div>':''}
      <div class="prev-metric">
        <span class="prev-metric-lbl">${lbl}</span>&nbsp;
        <span class="prev-metric-val">${valStr}</span>&nbsp;
        <span class="prev-metric-chg" style="color:${clr}">${arrow}</span>
      </div>`;
  }).join('');
  return `<div class="prev-bar"><span class="prev-bar-title">${title}</span>${items}</div>`;
}

function render(rows, g, trends, pg, pmo, pyr, dailyDay, dailyMonth, yr, mo, lmtdData){
  const ti=g.l1+g.l2+g.l3;
  const mxF=Math.max(g.dem,g.sub,ti,g.sel,g.ob_hc,g.ex_hc,1);

  const stg=[
    {l:'Demands',v:g.dem,c:'#e8453c'},{l:'Submitted',v:g.sub,c:'#8892a4'},
    {l:'L1',v:g.l1,c:'#2ecc71'},{l:'L2',v:g.l2,c:'#27ae60'},{l:'L3',v:g.l3,c:'#1e8449'},
    {l:'Selected',v:g.sel,c:'#8892a4'},{l:'Onboarded',v:g.ob_hc,c:'#2ecc71'},
  ];
  let sh='<div class="stage-wrap">';
  stg.forEach((s,i)=>{
    const cv=i>0?pc(s.v,stg[i-1].v):100;
    sh+=`<div class="stage"><div class="sv" style="color:${s.c}">${Math.round(s.v).toLocaleString()}</div><div class="sl">${s.l}</div>${i>0?`<div class="sc" style="color:${cv>=50?'#2ecc71':'#e8453c'}">${cv}%</div>`:''}</div>`;
    if(i<stg.length-1)sh+=`<div class="sarrow">&#8250;</div>`;
  });
  sh+='</div>';

  const fn=(lb,v,col,base)=>`<div class="fn"><div class="fnl">${lb}</div><div class="fnt"><div class="fnf" style="width:${Math.round(v/mxF*100)}%;background:${col}"></div></div><div class="fnn">${Math.round(v).toLocaleString()}</div><div class="fnp">${base?pc(v,base)+'%':''}</div></div>`;

  let lmtdBarHTML='';
  if(lmtdData && lmtdData.enabled && lmtdData.previous_grand){
    lmtdBarHTML=buildCompareBar(lmtdData.title || 'LMTD', g, lmtdData.previous_grand);
  }

  let prevBarHTML='';
  if(pg && pmo && pyr){
    const metrics=[
      ['Demands',g.dem,pg.dem,false],['Submissions',g.sub,pg.sub,false],
      ['L1',g.l1,pg.l1,false],['L2',g.l2,pg.l2,false],['L3',g.l3,pg.l3,false],
      ['Selections',g.sel,pg.sel,false],['Onboarded',g.ob_hc,pg.ob_hc,false],
      ['Exits',g.ex_hc,pg.ex_hc,false],['Net HC',g.net_hc,pg.net_hc,false],
      ['Net PO',g.net_po,pg.net_po,true],['Net Margin',g.net_mg,pg.net_mg,true],
    ];
    const items=metrics.map(([lbl,cur,prev,isL],i)=>{
      const diff=parseFloat(cur)-parseFloat(prev);
      const pct=parseFloat(prev)!==0?Math.round(Math.abs(diff)/Math.abs(parseFloat(prev))*100):0;
      const clr=diff===0?'var(--t3)':diff>0?'#1db85a':'#e8453c';
      const arrow=diff===0?'—':(diff>0?'▲ ':'▼ ')+pct+'%';
      const valStr=isL?'&#8377;'+L(Math.abs(parseFloat(prev)))+'L':Math.round(prev).toLocaleString();
      return `${i>0?'<div class="prev-sep"></div>':''}
        <div class="prev-metric">
          <span class="prev-metric-lbl">${lbl}</span>&nbsp;
          <span class="prev-metric-val">${valStr}</span>&nbsp;
          <span class="prev-metric-chg" style="color:${clr}">${arrow}</span>
        </div>`;
    }).join('');
    prevBarHTML=`<div class="prev-bar"><span class="prev-bar-title">vs ${MON[parseInt(pmo)]} ${pyr}</span>${items}</div>`;
  }

  const activityKeys=['dem','dem_u','sub','sub_fp','l1','l1_fp','l2','l2_fp','l3','l3_fp','sel','sp_hc','ob_hc','active_hc','ex_hc','ex_pipe_hc','net_hc'];
  const visibleRows=[...rows].filter(r=>{
    const m=r.metrics||{};
    return activityKeys.some(k=>(parseFloat(m[k])||0)!==0);
  }).sort((a,b)=>{
    const am=a.metrics||{}, bm=b.metrics||{};
    const aScore=activityKeys.reduce((sum,k)=>sum+Math.abs(parseFloat(am[k])||0),0);
    const bScore=activityKeys.reduce((sum,k)=>sum+Math.abs(parseFloat(bm[k])||0),0);
    return bScore-aScore || a.label.localeCompare(b.label);
  });

  const pipelineRows=[...rows].filter(r=>{
    const m=r.metrics||{};
    return ['dem','sub','l1','l2','l3','sel','ob_hc'].some(k=>(parseFloat(m[k])||0)>0);
  }).sort((a,b)=>{
    const am=a.metrics||{}, bm=b.metrics||{};
    return (parseFloat(bm.ob_hc)||0)-(parseFloat(am.ob_hc)||0)
      || (parseFloat(bm.sel)||0)-(parseFloat(am.sel)||0)
      || (parseFloat(bm.sub)||0)-(parseFloat(am.sub)||0)
      || a.label.localeCompare(b.label);
  });

  const pipeRate=(a,b)=>b?Math.round(a/b*100):0;
  const pipelineMeta=`${pipelineRows.length} client${pipelineRows.length===1?'':'s'} in current filter`;
  const pipelineClientHTML=pipelineRows.length
    ? `<div class="pipe-client-wrap">
         <div class="pipe-client-scroll">
           <table class="pipe-table">
             <thead>
               <tr>
                 <th>Client</th>
                 <th>Dem</th>
                 <th>Sub</th>
                 <th>L1</th>
                 <th>L2</th>
                 <th>L3</th>
                 <th>Sel</th>
                 <th>Ob</th>
                 <th>Sub→L1</th>
                 <th>L1→Sel</th>
                 <th>Sel→Ob</th>
               </tr>
             </thead>
             <tbody>
               ${pipelineRows.map(r=>{
                  const m=r.metrics||{};
                  const sub=parseFloat(m.sub)||0;
                  const l1=parseFloat(m.l1)||0;
                  const sel=parseFloat(m.sel)||0;
                  const ob=parseFloat(m.ob_hc)||0;
                  return `<tr>
                    <td>
                      <span class="pipe-client-name">${r.label}</span>
                      <span class="pipe-sub">${r.domain||'Unmapped'}${r.bh?` · ${r.bh}`:''}</span>
                    </td>
                    <td>${fN(m.dem)}</td>
                    <td>${fN(sub)}</td>
                    <td>${fN(l1)}</td>
                    <td>${fN(m.l2)}</td>
                    <td>${fN(m.l3)}</td>
                    <td>${fN(sel)}</td>
                    <td>${fN(ob)}</td>
                    <td class="pipe-rate" style="color:${pipeRate(l1,sub)>=50?'#1db85a':'#e67e22'}">${pipeRate(l1,sub)}%</td>
                    <td class="pipe-rate" style="color:${pipeRate(sel,l1)>=30?'#1db85a':'#e67e22'}">${pipeRate(sel,l1)}%</td>
                    <td class="pipe-rate" style="color:${pipeRate(ob,sel)>=50?'#1db85a':'#e67e22'}">${pipeRate(ob,sel)}%</td>
                  </tr>`;
               }).join('')}
             </tbody>
           </table>
         </div>
       </div>`
    : `<div class="pipe-empty">No client pipeline activity found for the current filter.</div>`;

  const grouped={};
  visibleRows.forEach(r=>{
    const cat=r.domain||'Unmapped';
    if(!grouped[cat]) grouped[cat]=[];
    grouped[cat].push(r);
  });
  const catOrder=['Services','Captive','ITES','Unmapped'];
  const sortedCats=[...catOrder.filter(c=>grouped[c]),...Object.keys(grouped).filter(c=>!catOrder.includes(c))];
  const clientBreakdownMeta=`${visibleRows.length} client${visibleRows.length===1?'':'s'} in current filter`;
  const tableGrand={};
  const keys=['dem','dem_u','sub','sub_fp','l1','l1_fp','l2','l2_fp','l3','l3_fp','sel','sp_hc','sp_po','sp_mg','ob_hc','ob_po','ob_mg','ex_hc','ex_po','net_hc','net_po','net_mg'];
  keys.forEach(k=>tableGrand[k]=visibleRows.reduce((a,r)=>a+(parseFloat(r.metrics[k])||0),0));

  let th='';
  if(!visibleRows.length){
    th=`<tr class="tr-r"><td colspan="26" style="text-align:center;color:var(--t3);padding:18px">No clients with demands, submissions, interviews, selections, onboardings, exits, or other activity for the current filter.</td></tr>`;
  }
  sortedCats.forEach(cat=>{
    const catRows=grouped[cat];
    th+=`<tr class="tr-d"><td colspan="26">${cat}</td></tr>`;
    catRows.forEach(r=>{
      const m=r.metrics;
      th+=`<tr class="tr-r">
        <td>${r.label}</td>
        <td>${fN(m.dem)}</td><td>${fN(m.dem_u)}</td>
        <td class="bl">${fN(m.sub)}</td><td>${fN(m.sub_fp)}</td>
        <td class="bl">${fN(m.l1)}</td><td>${fN(m.l1_fp)}</td>
        <td>${fN(m.l2)}</td><td>${fN(m.l2_fp)}</td>
        <td>${fN(m.l3)}</td><td>${fN(m.l3_fp)}</td>
        <td class="bl">${fN(m.sel)}</td>
        <td class="bl">${fN(m.sp_hc)}</td><td>${f2(m.sp_po)}</td><td>${f2(m.sp_mg)}</td>
        <td class="bl">${fN(m.ob_hc)}</td><td>${f2(m.ob_po)}</td><td>${f2(m.ob_mg)}</td>
        <td class="bl">${fN(m.ex_hc)}</td><td>${f2(m.ex_po)}</td>
        <td class="bl">${pn(m.net_hc,false)}</td><td>${pn(m.net_po,true)}</td><td>${pn(m.net_mg,true)}</td>
        <td class="bl" style="color:var(--t2)">${r.domain||'<span class="dim">—</span>'}</td>
        <td style="color:var(--t2)">${r.bh||'<span class="dim">—</span>'}</td>
      </tr>`;
    });
    const ct={};
    keys.forEach(k=>ct[k]=catRows.reduce((a,r)=>a+(parseFloat(r.metrics[k])||0),0));
    th+=`<tr class="tr-t">
      <td>${cat} Total</td>
      <td>${fN(ct.dem)}</td><td>${fN(ct.dem_u)}</td>
      <td class="bl">${fN(ct.sub)}</td><td>${fN(ct.sub_fp)}</td>
      <td class="bl">${fN(ct.l1)}</td><td>${fN(ct.l1_fp)}</td>
      <td>${fN(ct.l2)}</td><td>${fN(ct.l2_fp)}</td>
      <td>${fN(ct.l3)}</td><td>${fN(ct.l3_fp)}</td>
      <td class="bl">${fN(ct.sel)}</td>
      <td class="bl">${fN(ct.sp_hc)}</td><td>${f2(ct.sp_po)}</td><td>${f2(ct.sp_mg)}</td>
      <td class="bl">${fN(ct.ob_hc)}</td><td>${f2(ct.ob_po)}</td><td>${f2(ct.ob_mg)}</td>
      <td class="bl">${fN(ct.ex_hc)}</td><td>${f2(ct.ex_po)}</td>
      <td class="bl">${pn(ct.net_hc,false)}</td><td>${pn(ct.net_po,true)}</td><td>${pn(ct.net_mg,true)}</td>
      <td class="bl"></td><td></td>
    </tr>`;
  });
  th+=`<tr class="tr-g">
    <td>Grand Total</td>
    <td>${fN(tableGrand.dem)}</td><td>${fN(tableGrand.dem_u)}</td>
    <td class="bl">${fN(tableGrand.sub)}</td><td>${fN(tableGrand.sub_fp)}</td>
    <td class="bl">${fN(tableGrand.l1)}</td><td>${fN(tableGrand.l1_fp)}</td>
    <td>${fN(tableGrand.l2)}</td><td>${fN(tableGrand.l2_fp)}</td>
    <td>${fN(tableGrand.l3)}</td><td>${fN(tableGrand.l3_fp)}</td>
    <td class="bl">${fN(tableGrand.sel)}</td>
    <td class="bl">${fN(tableGrand.sp_hc)}</td><td>${f2(tableGrand.sp_po)}</td><td>${f2(tableGrand.sp_mg)}</td>
    <td class="bl">${fN(tableGrand.ob_hc)}</td><td>${f2(tableGrand.ob_po)}</td><td>${f2(tableGrand.ob_mg)}</td>
    <td class="bl">${fN(tableGrand.ex_hc)}</td><td>${f2(tableGrand.ex_po)}</td>
    <td class="bl">${pn(tableGrand.net_hc,false)}</td><td>${pn(tableGrand.net_po,true)}</td><td>${pn(tableGrand.net_mg,true)}</td>
    <td class="bl"></td><td></td>
  </tr>`;

  document.getElementById('body').innerHTML=`
  <div class="render-fade">
  <div class="kpi-grid">
    <div class="kpi red"><div class="kpi-head"><div class="kpi-ico">&#128203;</div></div><div class="kpi-lbl">Demands</div>${animatedKpiValue(g.dem,'#e8453c')}<div class="kpi-sub"><strong>${Math.round(g.dem_open || g.dem)}</strong> openings &middot; <strong>${Math.round(g.dem_u)}</strong> roles are still waiting to be serviced</div>${g.dem_u>0?tagR('&#9888; '+Math.round(g.dem_u)+' pending'):tagG('All serviced')}</div>
    <div class="kpi grey"><div class="kpi-head"><div class="kpi-ico">&#128228;</div></div><div class="kpi-lbl">Submissions</div>${animatedKpiValue(g.sub,'#5f6b7f')}<div class="kpi-sub"><strong>${Math.round(g.sub_fp)}</strong> profiles are awaiting client feedback</div>${tagGr(pc(g.sub,g.dem)+'% demand coverage')}</div>
    <div class="kpi grey"><div class="kpi-head"><div class="kpi-ico">&#127897;</div></div><div class="kpi-lbl">Interviews</div>${animatedKpiValue(ti,'#5f6b7f')}<div class="kpi-sub">L1 <strong>${Math.round(g.l1)}</strong> &middot; L2 <strong>${Math.round(g.l2)}</strong> &middot; L3 <strong>${Math.round(g.l3)}</strong></div>${tagGr(pc(g.l1,g.sub)+'% sub&#8594;L1')}</div>
    <div class="kpi grey"><div class="kpi-head"><div class="kpi-ico">&#9989;</div></div><div class="kpi-lbl">Selections</div>${animatedKpiValue(g.sel,'#5f6b7f')}<div class="kpi-sub">Confirmed selected candidates in the current period</div>${tagGr(pc(g.sel,g.l1)+'% L1&#8594;selected')}</div>
    <div class="kpi blue"><div class="kpi-head"><div class="kpi-ico">&#128203;</div></div><div class="kpi-lbl">Selection Pipeline</div>${animatedKpiValue(g.sp_hc,'#3498db')}<div class="kpi-sub"><strong>&#8377;${L(g.sp_po)}L</strong> PO value &middot; <strong>&#8377;${L(g.sp_mg)}L</strong> margin</div>${tagBl(Math.max(0, Math.round((g.sel||0)-(g.ob_hc||0)-(g.sp_hc||0)))+' Yet to be Onboarded')}</div>
    <div class="kpi blue"><div class="kpi-head"><div class="kpi-ico">&#128101;</div></div><div class="kpi-lbl">Active Headcount</div>${animatedKpiValue(g.active_hc,'#1f7ae0')}<div class="kpi-sub"><span id="active_hc_label"></span></div>${tagBl('Selected-period onboarding excluded')}</div>
    <div class="kpi green"><div class="kpi-head"><div class="kpi-ico">&#128640;</div></div><div class="kpi-lbl">Onboarded</div>${animatedKpiValue(g.ob_hc,'#1db85a')}<div class="kpi-sub"><strong>&#8377;${L(g.ob_po)}L</strong> PO value &middot; <strong>&#8377;${L(g.ob_mg)}L</strong> margin</div>${tagG(pc(g.ob_hc,g.sel)+'% sel&#8594;joined')}</div>
    <div class="kpi red"><div class="kpi-head"><div class="kpi-ico">&#128682;</div></div><div class="kpi-lbl">Exits</div>${animatedKpiValue(g.ex_hc,'#e8453c')}<div class="kpi-sub"><strong>&#8377;${L(g.ex_po)}L</strong> PO value &middot; <strong>&#8377;${L(g.ex_mg)}L</strong> margin<br>Pipeline <strong>${Math.round(g.ex_pipe_hc)}</strong> HC &middot; <strong>&#8377;${L(g.ex_pipe_po)}L</strong> PO &middot; <strong>&#8377;${L(g.ex_pipe_mg)}L</strong> margin</div>${tagR(Math.round(g.ex_hc)+' headcount lost')}</div>
    <div class="kpi ${g.net_hc>=0?'green':'red'}"><div class="kpi-head"><div class="kpi-ico">&#128202;</div></div><div class="kpi-lbl">Net HC</div>${animatedKpiValue(g.net_hc,g.net_hc>=0?'#1db85a':'#e8453c','number',true)}<div class="kpi-sub"><strong>&#8377;${L(Math.abs(g.net_po))}L</strong> net PO movement</div>${g.net_hc>=0?tagG('&#9650; Growth'):tagR('&#9660; Decline')}</div>
    <div class="kpi ${g.net_po>=0?'green':'red'}"><div class="kpi-head"><div class="kpi-ico">&#128176;</div></div><div class="kpi-lbl">Net PO</div>${animatedKpiValue(g.net_po,g.net_po>=0?'#1db85a':'#e8453c','currency',true,'30px')}<div class="kpi-sub">Ob <strong>&#8377;${L(g.ob_po)}L</strong> &minus; Ex <strong>&#8377;${L(g.ex_po)}L</strong></div>${g.net_po>=0?tagG('Positive'):tagR('Negative')}</div>
    <div class="kpi ${g.net_mg>=0?'green':'red'}"><div class="kpi-head"><div class="kpi-ico">&#128181;</div></div><div class="kpi-lbl">Net Margin</div>${animatedKpiValue(g.net_mg,g.net_mg>=0?'#1db85a':'#e8453c','currency',true,'30px')}<div class="kpi-sub">Ob <strong>&#8377;${L(g.ob_mg)}L</strong> &minus; Ex <strong>&#8377;${L(g.ex_mg)}L</strong></div>${g.net_mg>=0?tagG('Positive'):tagR('Negative')}</div>
  </div>
  ${lmtdBarHTML}
  ${prevBarHTML}
  <div class="sec">Recruitment Pipeline</div>
  <div class="card pipeline-shell">
    <button id="pipelineToggle" class="sec-toggle ${_pipelineCollapsed?'collapsed':''}" onclick="togglePipelineSection()">
      <div>
        <div class="card-t">End-to-end funnel — conversion at each stage</div>
        <div class="sec-toggle-meta">${pipelineMeta}</div>
      </div>
      <span class="sec-chevron">&#9662;</span>
    </button>
    <div id="pipelineBody" class="pipeline-body ${_pipelineCollapsed?'collapsed':''}">
  <div class="card" style="margin-bottom:14px">
    <div class="card-t">Stage Snapshot</div>${sh}
  </div>
  <div class="r2">
    <div class="card">
      <div class="card-t">Volume Funnel</div>
      ${fn('Demands',g.dem,'#e8453c',null)}
      ${fn('Submissions',g.sub,'#8892a4',g.dem)}
      ${fn('L1 Interviews',g.l1,'#2ecc71',g.sub)}
      ${fn('L2 Interviews',g.l2,'#27ae60',g.l1)}
      ${fn('L3 Interviews',g.l3,'#1e8449',g.l2)}
      ${fn('Selections',g.sel,'#8892a4',g.l1)}
      ${fn('Onboarded',g.ob_hc,'#2ecc71',g.sel)}
      ${fn('Exits',g.ex_hc,'#e8453c',null)}
    </div>
  </div>
      <div class="pipe-client-head">
        <div class="card-t" style="margin-bottom:0;">Client-wise Pipeline Breakdown</div>
        <div class="pipe-client-meta">Stage counts and core conversion ratios per client</div>
      </div>
      ${pipelineClientHTML}
    </div>
  </div>
  <div class="sec">MRR Breakdown</div>
  <div class="r3">
    <div class="card" style="border-color:var(--green-bd)">
      <div class="card-t" style="color:#1db85a">&#128640; Onboarding</div>
      <div class="mrr-r">
        <div class="mrr-c"><div class="mrr-v" style="color:#1db85a">${Math.round(g.ob_hc)}</div><div class="mrr-l">HC</div></div>
        <div class="mrr-c"><div class="mrr-v" style="color:#17a050">&#8377;${L(g.ob_po)}L</div><div class="mrr-l">PO Value</div></div>
        <div class="mrr-c"><div class="mrr-v" style="color:#17a050">&#8377;${L(g.ob_mg)}L</div><div class="mrr-l">Margin</div></div>
      </div>
    </div>
    <div class="card" style="border-color:var(--red-bd)">
      <div class="card-t" style="color:#e8453c">&#128682; Exits</div>
      <div class="mrr-r">
        <div class="mrr-c"><div class="mrr-v" style="color:#e8453c">${Math.round(g.ex_hc)}</div><div class="mrr-l">HC</div></div>
        <div class="mrr-c"><div class="mrr-v" style="color:#e8453c">&#8377;${L(g.ex_po)}L</div><div class="mrr-l">PO Value</div></div>
        <div class="mrr-c"><div class="mrr-v" style="color:#e8453c">&#8377;${L(g.ex_mg)}L</div><div class="mrr-l">Margin</div></div>
      </div>
    </div>
    <div class="card" style="border-color:${g.net_hc>=0?'var(--green-bd)':'var(--red-bd)'}">
      <div class="card-t" style="color:${g.net_hc>=0?'#1db85a':'#e8453c'}">&#128202; Net MRR</div>
      <div class="mrr-r">
        <div class="mrr-c"><div class="mrr-v" style="color:${g.net_hc>=0?'#1db85a':'#e8453c'}">${g.net_hc>=0?'+':''}${Math.round(g.net_hc)}</div><div class="mrr-l">Net HC</div></div>
        <div class="mrr-c"><div class="mrr-v" style="color:${g.net_po>=0?'#1db85a':'#e8453c'}">${g.net_po>=0?'+':'&minus;'}&#8377;${L(Math.abs(g.net_po))}L</div><div class="mrr-l">Net PO</div></div>
        <div class="mrr-c"><div class="mrr-v" style="color:${g.net_mg>=0?'#1db85a':'#e8453c'}">${g.net_mg>=0?'+':'&minus;'}&#8377;${L(Math.abs(g.net_mg))}L</div><div class="mrr-l">Net Margin</div></div>
      </div>
    </div>
  </div>
  <div class="sec">Day-on-Day Trends</div>
  <div style="margin-bottom:10px;display:flex;gap:6px;">
    <button class="ttab ${getActiveTrendPreset('day')===7?'on':''}" id="dayBtn7" onclick="setRange('day',7,this)">Last 7 Days</button>
    <button class="ttab ${getActiveTrendPreset('day')===15?'on':''}" id="dayBtn15" onclick="setRange('day',15,this)">Last 15 Days</button>
    <button class="ttab ${getActiveTrendPreset('day')===30?'on':''}" id="dayBtn30" onclick="setRange('day',30,this)">Last 30 Days</button>
  </div>
  <div style="display:flex;gap:10px;margin-bottom:10px;align-items:center;">
    <span style="font-size:11px;color:var(--t3)">From</span>
    <input type="date" id="fromDateDay" onchange="clearActiveTrendPreset('day');setActiveTrendDates('day',this.value,getActiveTrendToDate('day'));loadAll()" style="padding:4px;background:var(--s2);color:var(--text);border:1px solid var(--border);border-radius:6px;">
    <span style="font-size:11px;color:var(--t3)">To</span>
    <input type="date" id="toDateDay" onchange="clearActiveTrendPreset('day');setActiveTrendDates('day',getActiveTrendFromDate('day'),this.value);loadAll()" style="padding:4px;background:var(--s2);color:var(--text);border:1px solid var(--border);border-radius:6px;">
  </div>
  <div class="trend-tabs" id="dayTrendTabs">
    <button class="ttab ${_dayTrendTab==='dem'?'on':''}" onclick="switchTrendMetric('day','dem',this)">📄 Demand</button>
    <button class="ttab ${_dayTrendTab==='dem_u'?'on':''}" onclick="switchTrendMetric('day','dem_u',this)">Unserviced Demands</button>
    <button class="ttab ${_dayTrendTab==='sub'?'on':''}" onclick="switchTrendMetric('day','sub',this)">📤 Submission</button>
    <button class="ttab ${_dayTrendTab==='sub_fp'?'on':''}" onclick="switchTrendMetric('day','sub_fp',this)">Feedback Pending</button>
    <button class="ttab ${_dayTrendTab==='intv'?'on':''}" onclick="switchTrendMetric('day','intv',this)">🎯 Interview</button>
    <button class="ttab ${_dayTrendTab==='sel'?'on':''}" onclick="switchTrendMetric('day','sel',this)">✅ Selection</button>
    <button class="ttab ${_dayTrendTab==='ob'?'on':''}" onclick="switchTrendMetric('day','ob',this)">🚀 Onboarding</button>
    <button class="ttab ${_dayTrendTab==='ex'?'on':''}" onclick="switchTrendMetric('day','ex',this)">🚪 Exit</button>
  </div>
  <div class="card" style="margin-bottom:16px">
    <div class="card-t" id="trendTitleDay">Demand — Daily Trend</div>
    <div class="trend-chart-wrap">
      <div class="trend-avg" id="trendAvgDay">Average: <strong>-</strong></div>
      <canvas id="trendChartDay"></canvas>
    </div>
  </div>
  <div class="sec">Month-on-Month Trends</div>
  <div style="margin-bottom:10px;display:flex;gap:6px;">
    <button class="ttab ${getActiveTrendPreset('month')===3?'on':''}" id="monthBtn3" onclick="setRange('month',3,this)">Last 3 Months</button>
    <button class="ttab ${getActiveTrendPreset('month')===6?'on':''}" id="monthBtn6" onclick="setRange('month',6,this)">Last 6 Months</button>
    <button class="ttab ${getActiveTrendPreset('month')===12?'on':''}" id="monthBtn12" onclick="setRange('month',12,this)">Last 12 Months</button>
  </div>
  <div style="display:flex;gap:10px;margin-bottom:10px;align-items:center;">
    <span style="font-size:11px;color:var(--t3)">From</span>
    <input type="date" id="fromDateMonth" onchange="clearActiveTrendPreset('month');setActiveTrendDates('month',this.value,getActiveTrendToDate('month'));loadAll()" style="padding:4px;background:var(--s2);color:var(--text);border:1px solid var(--border);border-radius:6px;">
    <span style="font-size:11px;color:var(--t3)">To</span>
    <input type="date" id="toDateMonth" onchange="clearActiveTrendPreset('month');setActiveTrendDates('month',getActiveTrendFromDate('month'),this.value);loadAll()" style="padding:4px;background:var(--s2);color:var(--text);border:1px solid var(--border);border-radius:6px;">
  </div>
  <div class="trend-tabs" id="monthTrendTabs">
    <button class="ttab ${_monthTrendTab==='dem'?'on':''}" onclick="switchTrendMetric('month','dem',this)">📄 Demand</button>
    <button class="ttab ${_monthTrendTab==='dem_u'?'on':''}" onclick="switchTrendMetric('month','dem_u',this)">Unserviced Demands</button>
    <button class="ttab ${_monthTrendTab==='sub'?'on':''}" onclick="switchTrendMetric('month','sub',this)">📤 Submission</button>
    <button class="ttab ${_monthTrendTab==='sub_fp'?'on':''}" onclick="switchTrendMetric('month','sub_fp',this)">Feedback Pending</button>
    <button class="ttab ${_monthTrendTab==='intv'?'on':''}" onclick="switchTrendMetric('month','intv',this)">🎯 Interview</button>
    <button class="ttab ${_monthTrendTab==='sel'?'on':''}" onclick="switchTrendMetric('month','sel',this)">✅ Selection</button>
    <button class="ttab ${_monthTrendTab==='ob'?'on':''}" onclick="switchTrendMetric('month','ob',this)">🚀 Onboarding</button>
    <button class="ttab ${_monthTrendTab==='hc'?'on':''}" onclick="switchTrendMetric('month','hc',this)">👥 Headcount</button>
    <button class="ttab ${_monthTrendTab==='ex'?'on':''}" onclick="switchTrendMetric('month','ex',this)">🚪 Exit</button>
  </div>
  <div class="card">
    <div class="card-t" id="trendTitleMonth">Demand — Month-on-Month Trend</div>
    <div class="trend-chart-wrap">
      <div class="trend-avg" id="trendAvgMonth">Average: <strong>-</strong></div>
      <canvas id="trendChartMonth"></canvas>
    </div>
  </div>
  <div id="ceo-revenue"></div>
  <div id="ceo-exit"></div>
  <div id="ceo-bh"></div>
  <div class="sec">Client Breakdown - MTD</div>
  <div class="card pipeline-shell">
    <button id="clientBreakdownToggle" class="sec-toggle ${_clientBreakdownCollapsed?'collapsed':''}" onclick="toggleClientBreakdownSection()">
      <div>
        <div class="card-t">Client Breakdown - MTD</div>
        <div class="sec-toggle-meta">${clientBreakdownMeta}</div>
      </div>
      <span class="sec-chevron">&#9662;</span>
    </button>
    <div id="clientBreakdownBody" class="pipeline-body ${_clientBreakdownCollapsed?'collapsed':''}">
      <table>
        <thead>
          <tr>
            <th class="tg2" rowspan="2" style="text-align:left;border-left:none">Client</th>
            <th class="tg2" colspan="2">Demands</th><th class="tg2" colspan="2">Submissions</th>
            <th class="tg2" colspan="6">Interviews</th><th class="tg2" colspan="1">Selections</th>
            <th class="tg2" colspan="3">Sel. Pipeline</th><th class="tg2" colspan="3">Onboarding</th>
            <th class="tg2" colspan="2">Exit</th><th class="tg2" colspan="3">Net MRR</th>
            <th class="tg2" colspan="1">Domain</th><th class="tg2" colspan="1">Business Head</th>
          </tr>
          <tr class="sh">
            <th>Actual</th><th>Unsvc</th>
            <th class="bl">Actual</th><th>F/B Pend</th>
            <th class="bl">L1</th><th>L1 Pend</th><th>L2</th><th>L2 Pend</th><th>L3</th><th>L3 Pend</th>
            <th class="bl">Actual</th>
            <th class="bl">HC</th><th>PO(L)</th><th>Mgn(L)</th>
            <th class="bl">HC</th><th>PO(L)</th><th>Mgn(L)</th>
            <th class="bl">HC</th><th>PO(L)</th>
            <th class="bl">Net HC</th><th>Net PO</th><th>Net Mgn</th>
            <th class="bl">Domain</th><th>BH</th>
          </tr>
        </thead>
        <tbody>${th}</tbody>
      </table>
    </div>
  </div>
  <div class="sec">Raw Data Explorer</div>
  <div class="card raw-shell pipeline-shell">
    <button id="rawExplorerToggle" class="sec-toggle ${_rawExplorerCollapsed?'collapsed':''}" onclick="toggleRawExplorerSection()">
      <div>
        <div class="card-t">Raw Data Explorer</div>
        <div class="sec-toggle-meta">Filtered source rows for each pipeline stage</div>
      </div>
      <span class="sec-chevron">&#9662;</span>
    </button>
    <div id="rawExplorerBody" class="pipeline-body ${_rawExplorerCollapsed?'collapsed':''}">
    <div class="trend-tabs" id="rawDatasetTabs">
      <button class="ttab ${_rawDataset==='demand'?'on':''}" onclick="setRawDataset('demand',this)">Demands</button>
      <button class="ttab ${_rawDataset==='sub'?'on':''}" onclick="setRawDataset('sub',this)">Submissions</button>
      <button class="ttab ${_rawDataset==='intv'?'on':''}" onclick="setRawDataset('intv',this)">Interviews</button>
      <button class="ttab ${_rawDataset==='sel'?'on':''}" onclick="setRawDataset('sel',this)">Selections</button>
      <button class="ttab ${_rawDataset==='selpipe'?'on':''}" onclick="setRawDataset('selpipe',this)">Selection Pipeline</button>
      <button class="ttab ${_rawDataset==='ob'?'on':''}" onclick="setRawDataset('ob',this)">Onboardings</button>
      <button class="ttab ${_rawDataset==='exitpipe'?'on':''}" onclick="setRawDataset('exitpipe',this)">Exit Pipeline</button>
      <button class="ttab ${_rawDataset==='exit'?'on':''}" onclick="setRawDataset('exit',this)">Exit</button>
    </div>
    <div class="raw-toolbar">
      <div class="raw-filter">
        <label>Time Filter</label>
        <div class="raw-time-stack">
          <div class="raw-date-field">
            <span>Month</span>
            <input type="month" id="rawMonthFilter" onchange="setRawMonthDates();loadRawData()">
          </div>
          <div class="raw-date-stack">
            <div class="raw-date-field">
              <span>From Date</span>
              <input type="date" id="rawFromDate" onchange="handleRawDateChange()">
            </div>
            <div class="raw-date-field">
              <span>To Date</span>
              <input type="date" id="rawToDate" onchange="handleRawDateChange()">
            </div>
          </div>
        </div>
        <div class="raw-filter-note">Pick a month to auto-fill the full date range, or use custom dates below.</div>
      </div>
      <div class="raw-filter raw-filter-wide">
        <label>Quick Range</label>
        <div class="raw-range-presets">
          <button class="btn-g" onclick="setRawCurrentMonth()">This Month</button>
          <button class="btn-g" onclick="setRawDatePreset(7)">Last 7 Days</button>
          <button class="btn-g" onclick="setRawDatePreset(30)">Last 30 Days</button>
        </div>
        <div class="raw-filter-note">Use quick presets for common ranges, or override with custom dates.</div>
      </div>
      <div class="raw-filter" id="rawDemandOnlyWrap">
        <label>Demand Filter</label>
        <select id="rawDemandStatus" onchange="syncRawDemandPills();loadRawData()" style="display:none;">
          <option value="all">All Demands</option>
          <option value="unserviced">Unserviced</option>
          <option value="serviced">Serviced</option>
        </select>
        <div class="raw-demand-pills">
          <button type="button" class="raw-demand-pill active" data-value="all" onclick="setRawDemandStatus('all')">All</button>
          <button type="button" class="raw-demand-pill" data-value="unserviced" onclick="setRawDemandStatus('unserviced')">Unserviced</button>
          <button type="button" class="raw-demand-pill" data-value="serviced" onclick="setRawDemandStatus('serviced')">Serviced</button>
        </div>
        <div class="raw-filter-note">Choose whether to show all demand rows, only unserviced ones, or only serviced ones.</div>
      </div>
      <div class="raw-filter raw-client-filter raw-filter-wide">
        <label>Client Filter</label>
        <div class="cl-wrap" id="rawClientWrap">
          <button class="cl-btn" onclick="toggleRawClientPanel(event)">Select Clients <span class="cl-ct" id="rawClientCt">0</span> <span>&#9662;</span></button>
          <div class="cl-panel" id="rawClientPanel">
            <div class="cl-search"><input type="text" id="rawClientSearch" placeholder="Search clients..." oninput="filterRawClientList()"></div>
            <div class="raw-client-actions">
              <button class="btn-g" onclick="selectAllRawClients()">Select All</button>
              <button class="btn-g" onclick="clearRawClients()">Clear All</button>
            </div>
            <div class="raw-client-list" id="rawClientList"></div>
            <div class="cl-apply">
              <div class="raw-client-meta" id="rawClientMeta">All clients included</div>
              <button class="btn-r" onclick="applyRawClients()">Apply</button>
            </div>
          </div>
        </div>
      </div>
    </div>
    <div class="raw-action-bar">
      <div class="raw-summary" id="rawSummary">Waiting for raw data...</div>
      <button class="btn-g" onclick="resetRawFilters()">Reset Filters</button>
      <div class="raw-download-group">
        <div class="raw-filter" style="min-width:180px;">
          <label for="rawColumnMode">Download Content</label>
          <select id="rawColumnMode" onchange="updateRawDownloadLinks()">
            <option value="visible">Visible Data</option>
            <option value="all">All Columns</option>
          </select>
        </div>
        <a id="rawDownloadCsvBtn" class="btn-g" href="/api/raw_data_export?dataset=demand&format=csv&columns=visible" download>Download CSV</a>
        <a id="rawDownloadXlsxBtn" class="btn-r" href="/api/raw_data_export?dataset=demand&format=xlsx&columns=visible" download>Download Excel</a>
      </div>
    </div>
    <div id="rawDataTable"><div class="raw-empty">Loading raw data...</div></div>
    </div>
  </div>
  </div>
  </div>`;

  window._trends=trends;
  window._dailyDay=dailyDay;
  window._dailyMonth=dailyMonth;

  renderTrend('demand');
  renderTrendSection('day', _dayTrendTab);
  renderTrendSection('month', _monthTrendTab);
  animateKpiValues();
  hydrateRawSection();
  loadRawData();

  // Load CEO analytics sections
  loadExitReasons(yr, mo);
  loadBHConversion(yr, mo);
  loadRevenueTrend();
}

function setRange(modeOrDays, daysOrBtn, maybeBtn){
  let mode=_trendGrain, days=modeOrDays, btn=daysOrBtn;
  if(typeof modeOrDays==='string'){
    mode=modeOrDays;
    days=daysOrBtn;
    btn=maybeBtn;
  }
  _trendGrain=mode;
  setActiveTrendPreset(mode, days);
  ['btn7','btn15','btn30','dayBtn7','dayBtn15','dayBtn30','monthBtn3','monthBtn6','monthBtn12'].forEach(id=>{const el=document.getElementById(id);if(el)el.classList.remove('on');});
  if(btn) btn.classList.add('on');
  const today=new Date();
  const past=new Date(today);
  if(mode==='month'){
    past.setMonth(today.getMonth()-(days-1), 1);
  }else{
    past.setDate(today.getDate()-days);
  }
  setActiveTrendDates(mode, past.toISOString().split('T')[0], today.toISOString().split('T')[0]);
  const fromEl=document.getElementById(mode==='month'?'fromDateMonth':'fromDate') || document.getElementById('fromDateDay');
  const toEl=document.getElementById(mode==='month'?'toDateMonth':'toDate') || document.getElementById('toDateDay');
  if(fromEl) fromEl.value=getActiveTrendFromDate(mode);
  if(toEl)   toEl.value=getActiveTrendToDate(mode);
  loadAll();
}

function setTrendGrain(grain, btn){
  _trendGrain=grain;
  document.querySelectorAll('[onclick^="setTrendGrain"]').forEach(t=>t.classList.remove('on'));
  if(btn) btn.classList.add('on');
  const fromEl=document.getElementById('fromDate');
  const toEl=document.getElementById('toDate');
  if(fromEl) fromEl.value=getActiveTrendFromDate();
  if(toEl)   toEl.value=getActiveTrendToDate();
  loadAll();
}

function renderTrend(type){
  const t=window._trends; if(!t) return;
  let labels,vals1,vals2,color1,color2;
  if(type==='demand'){
    const pts=t.dem||[];
    labels=pts.map(p=>p.p); vals1=pts.map(p=>p.v);
    vals2=vals1.reduce((acc,v,i)=>{acc.push((acc[i-1]||0)+v);return acc;},[]);
    color1='rgba(232,69,60,0.85)';color2='rgba(232,69,60,0.4)';
  } else if(type==='submission'){
    const pts=t.sub||[],pts2=t.sub_fp||[];
    labels=pts.map(p=>p.p);vals1=pts.map(p=>p.v);vals2=pts2.map(p=>p.v);
    color1='rgba(107,115,133,0.85)';color2='rgba(232,69,60,0.7)';
  } else {
    const l1=t.l1||[],l2=t.l2||[],l3=t.l3||[];
    labels=l1.map(p=>p.p);vals1=l1.map(p=>p.v);vals2=l2.map(p=>p.v);
    const vals3=l3.map(p=>p.v);
    dc('t1');dc('t2');
    const c1=document.getElementById('trend1'),c2=document.getElementById('trend2');
    if(c1) charts['t1']=new Chart(c1,{type:'bar',data:{labels,datasets:[{label:'L1',data:vals1,backgroundColor:'rgba(29,184,90,0.85)',borderRadius:4}]},options:trendOpts()});
    if(c2) charts['t2']=new Chart(c2,{type:'bar',data:{labels,datasets:[{label:'L2',data:vals2,backgroundColor:'rgba(23,160,80,0.7)',borderRadius:4},{label:'L3',data:vals3,backgroundColor:'rgba(15,120,60,0.7)',borderRadius:4}]},options:trendOpts()});
    return;
  }
  dc('t1');dc('t2');
  const c1=document.getElementById('trend1'),c2=document.getElementById('trend2');
  if(c1) charts['t1']=new Chart(c1,{type:'bar',data:{labels,datasets:[{label:'',data:vals1,backgroundColor:color1,borderRadius:4}]},options:trendOpts()});
  if(c2) charts['t2']=new Chart(c2,{type:'line',data:{labels,datasets:[{label:'',data:vals2,borderColor:color2,backgroundColor:color2.replace('0.','0.1'),tension:.35,fill:true,pointRadius:3}]},options:trendOpts()});
}

function switchDaily(type, btn){
  _dailyTab=type;
  document.querySelectorAll('.trend-tabs .ttab').forEach(t=>t.classList.remove('on'));
  btn.classList.add('on');
  renderDaily(type);
}

function switchTrendMetric(mode, type, btn){
  if(mode==='month'){
    _monthTrendTab=type;
  }else{
    _dayTrendTab=type;
    _dailyTab=type;
  }
  const tabs=document.getElementById(mode==='month'?'monthTrendTabs':'dayTrendTabs');
  if(tabs){
    tabs.querySelectorAll('.ttab').forEach(t=>t.classList.remove('on'));
    if(btn) btn.classList.add('on');
    renderTrendSection(mode, type);
    return;
  }
  switchDaily(type, btn);
}

function formatTrendAverage(values){
  if(!values || !values.length) return '-';
  const avg=values.reduce((sum,val)=>sum+Number(val||0),0)/values.length;
  return Number.isInteger(avg) ? String(avg) : avg.toFixed(1);
}

function setTrendAverage(mode, values){
  const avgEl=document.getElementById(mode==='month'?'trendAvgMonth':'trendAvgDay');
  if(!avgEl) return;
  avgEl.innerHTML=`Average: <strong>${formatTrendAverage(values)}</strong>`;
}

function renderTrendSection(mode, type){
  const dataSource = mode==='month' ? window._dailyMonth : window._dailyDay;
  const titleEl = document.getElementById(mode==='month'?'trendTitleMonth':'trendTitleDay');
  const chartEl = document.getElementById(mode==='month'?'trendChartMonth':'trendChartDay');
  if(!dataSource || !titleEl || !chartEl){
    _trendGrain = mode;
    renderDaily(type);
    return;
  }

  let data, title, color;
  const suffix=mode==='month'?'Month-on-Month Trend':'Daily Trend';
  if(type==='dem'){data=dataSource.dem;title=`Demand — ${suffix}`;color='#e8453c';}
  else if(type==='dem_u'){data=dataSource.dem_u;title=`Unserviced Demands — ${suffix}`;color='#c0392b';}
  else if(type==='sub'){data=dataSource.sub;title=`Submission — ${suffix}`;color='#6b7385';}
  else if(type==='sub_fp'){data=dataSource.sub_fp;title=`Feedback Pending — ${suffix}`;color='#8e6e53';}
  else if(type==='intv'){data=dataSource.intv;title=`Interview — ${suffix}`;color='#1db85a';}
  else if(type==='sel'){data=dataSource.sel;title=`Selection — ${suffix}`;color='#f39c12';}
  else if(type==='ob'){data=dataSource.ob;title=`Onboarding — ${suffix}`;color='#1abc9c';}
  else if(type==='hc'){data=dataSource.hc;title=`Headcount Movement (+/-) — ${suffix}`;color='#1f7ae0';}
  else{data=dataSource.ex;title=`Exit — ${suffix}`;color='#e8453c';}

  const labels=(data||[]).map(x=>{
    if(mode==='month'){
      const [year, month]=String(x.d).split('-');
      const dt=new Date(Number(year), Number(month)-1, 1);
      return dt.toLocaleDateString('en-GB',{month:'short', year:'2-digit'});
    }
    const dt=new Date(x.d);
    return dt.toLocaleDateString('en-GB',{day:'2-digit',month:'short'});
  });
  const values=(data||[]).map(x=>x.v);
  const chartKey=mode==='month'?'trendMonth':'trendDay';
  dc(chartKey);
  titleEl.textContent=title;
  setTrendAverage(mode, values);
  charts[chartKey]=new Chart(chartEl,{
    type:'line',
    plugins:[ChartDataLabels],
    data:{labels,datasets:[{label:title,data:values,borderColor:color,backgroundColor:color+'22',tension:0.3,fill:true,pointBackgroundColor:'#fff',pointBorderColor:color,pointBorderWidth:2,pointRadius:5}]},
    options:dailyTrendOpts()
  });
}

function renderDaily(type){
  const d=window._daily;
  if(!d||!d[type]) return;
  let data, title, color;
  const suffix=_trendGrain==='month'?'Month-on-Month Trend':'Daily Trend';
  if(type==='dem'){data=d.dem;title=`Demand — ${suffix}`;color='#e8453c';}
  else if(type==='dem_u'){data=d.dem_u;title=`Unserviced Demands — ${suffix}`;color='#c0392b';}
  else if(type==='sub'){data=d.sub;title=`Submission — ${suffix}`;color='#6b7385';}
  else if(type==='sub_fp'){data=d.sub_fp;title=`Feedback Pending — ${suffix}`;color='#8e6e53';}
  else if(type==='intv'){data=d.intv;title=`Interview — ${suffix}`;color='#1db85a';}
  else if(type==='sel'){data=d.sel;title=`Selection — ${suffix}`;color='#f39c12';}
  else if(type==='ob'){data=d.ob;title=`Onboarding — ${suffix}`;color='#1abc9c';}
  else if(type==='hc'){data=d.hc;title=`Headcount Movement (+/-) — ${suffix}`;color='#1f7ae0';}
  else{data=d.ex;title=`Exit — ${suffix}`;color='#e8453c';}

  const labels=data.map(x=>{
    if(_trendGrain==='month'){
      const [year, month]=String(x.d).split('-');
      const dt=new Date(Number(year), Number(month)-1, 1);
      return dt.toLocaleDateString('en-GB',{month:'short', year:'2-digit'});
    }
    const dt=new Date(x.d);
    return dt.toLocaleDateString('en-GB',{day:'2-digit',month:'short'});
  });
  const values=data.map(x=>x.v);
  dc('daily');
  document.getElementById('dailyTitle').textContent=title;
  charts['daily']=new Chart(document.getElementById('dailyChart'),{
    type:'line',
    plugins:[ChartDataLabels],
    data:{labels,datasets:[{label:title,data:values,borderColor:color,backgroundColor:color+'22',tension:0.3,fill:true,pointBackgroundColor:'#fff',pointBorderColor:color,pointBorderWidth:2,pointRadius:5}]},
    options:dailyTrendOpts()
  });
}

function trendOpts(){
  return{responsive:true,maintainAspectRatio:false,
    plugins:{legend:{display:false},tooltip:{enabled:true}},
    scales:{
      x:{grid:{color:'rgba(0,0,0,0.05)'},ticks:{color:'#52586a',font:{size:10},maxRotation:40}},
      y:{grid:{color:'rgba(0,0,0,0.05)'},ticks:{color:'#52586a',font:{size:11}},beginAtZero:true,grace:'12%'}
    }};
}

function dailyTrendOpts(){
  return{responsive:true,maintainAspectRatio:false,
    plugins:{
      legend:{display:false},
      tooltip:{enabled:true},
      datalabels:{display:true,align:'top',anchor:'end',color:'#52586a',font:{size:11,weight:'700'},formatter:v=>v===0?'':v}
    },
    scales:{
      x:{grid:{color:'rgba(0,0,0,0.05)'},ticks:{color:'#52586a',font:{size:10},maxRotation:40}},
      y:{grid:{color:'rgba(0,0,0,0.05)'},ticks:{color:'#52586a',font:{size:11}},beginAtZero:true,grace:'15%'}
    }};
}

// ── CEO ANALYTICS

function ceoParams(yr, mo){
  const p = new URLSearchParams();
  const years=yr ? [yr] : [...selectedYears];
  const months=mo ? [mo] : [...selectedMonths];
  if(years.length) p.set('years', years.join(','));
  if(months.length) p.set('months', months.join(','));
  if(selectedClients.size>0) p.set('clients',[...selectedClients].join(','));
  if(selectedDomains.size>0) p.set('domains',[...selectedDomains].join(','));
  if(selectedBHs.size>0)     p.set('bhs',[...selectedBHs].join(','));
  return p;
}

async function loadRevenueTrend(){
  const r = await fetch('/api/revenue_trend?'+ceoParams());
  const {trend} = await r.json();
  if(!trend || !trend.length) return;
  const labels  = trend.map(t=>t.period);
  const net_hc  = trend.map(t=>t.net_hc);
  const net_po  = trend.map(t=>t.net_po);
  const net_mg  = trend.map(t=>t.net_mg);
  const ob_po   = trend.map(t=>t.ob_po);
  const ex_po   = trend.map(t=>t.ex_po);

  document.getElementById('ceo-revenue').innerHTML=`
    <div class="sec">📈 12-Month Revenue Trend</div>
    <div class="r3">
      <div class="card">
        <div class="card-t">Net PO Value — Last 12 Months (₹ Lakhs)</div>
        <div style="position:relative;height:220px"><canvas id="revPOChart"></canvas></div>
      </div>
      <div class="card">
        <div class="card-t">Onboarding vs Exit PO — Last 12 Months (₹ Lakhs)</div>
        <div style="position:relative;height:220px"><canvas id="revObExChart"></canvas></div>
      </div>
      <div class="card">
        <div class="card-t">Net MRR Movement â€” Last 12 Months</div>
        <div style="position:relative;height:220px"><canvas id="revMRRChart"></canvas></div>
      </div>
    </div>`;

  dc('revPO'); dc('revObEx'); dc('revMRR');
  charts['revPO'] = new Chart(document.getElementById('revPOChart'),{
    type:'bar',
    data:{labels, datasets:[
      {label:'Net PO',data:net_po,backgroundColor:net_po.map(v=>v>=0?'rgba(29,184,90,0.8)':'rgba(232,69,60,0.8)'),borderRadius:4},
      {label:'Net Margin',data:net_mg,type:'line',borderColor:'#f39c12',backgroundColor:'rgba(243,156,18,0.1)',tension:0.3,fill:true,pointRadius:3,yAxisID:'y'}
    ]},
    options:{...trendOpts(), plugins:{legend:{display:true,labels:{color:'#52586a',font:{size:11}}}}}
  });
  charts['revObEx'] = new Chart(document.getElementById('revObExChart'),{
    type:'bar',
    data:{labels, datasets:[
      {label:'Onboarding PO',data:ob_po,backgroundColor:'rgba(29,184,90,0.75)',borderRadius:4},
      {label:'Exit PO',data:ex_po.map(v=>-v),backgroundColor:'rgba(232,69,60,0.75)',borderRadius:4}
    ]},
    options:{...trendOpts(), plugins:{legend:{display:true,labels:{color:'#52586a',font:{size:11}}}},
      scales:{...trendOpts().scales, y:{...trendOpts().scales.y, stacked:false}}}
  });
  charts['revMRR'] = new Chart(document.getElementById('revMRRChart'),{
    type:'bar',
    data:{labels, datasets:[
      {label:'Net HC',data:net_hc,backgroundColor:net_hc.map(v=>v>=0?'rgba(29,184,90,0.82)':'rgba(232,69,60,0.82)'),borderRadius:4,yAxisID:'y'},
      {label:'Net PO',data:net_po,type:'line',borderColor:'#3498db',backgroundColor:'rgba(52,152,219,0.10)',tension:0.3,fill:false,pointRadius:3,yAxisID:'y1'},
      {label:'Net Margin',data:net_mg,type:'line',borderColor:'#f39c12',backgroundColor:'rgba(243,156,18,0.10)',tension:0.3,fill:false,pointRadius:3,yAxisID:'y1'}
    ]},
    options:{...trendOpts(), plugins:{legend:{display:true,labels:{color:'#52586a',font:{size:11}}}},
      scales:{
        ...trendOpts().scales,
        y:{...trendOpts().scales.y,title:{display:true,text:'Net HC'}},
        y1:{position:'right',grid:{drawOnChartArea:false},ticks:{color:'#52586a',font:{size:11}},title:{display:true,text:'₹ Lakhs'}}
      }}
  });
}

async function loadExitReasons(yr, mo){
  const r = await fetch('/api/exit_reasons?'+ceoParams(yr,mo));
  const {reasons, total} = await r.json();
  if(!reasons || !reasons.length){
    document.getElementById('ceo-exit').innerHTML='';
    return;
  }
  const palette=['#e8453c','#f39c12','#1db85a','#3498db','#9b59b6','#1abc9c','#e67e22','#e91e63'];
  const maxV = Math.max(...reasons.map(r=>r.v), 1);

  const bars = reasons.map((r,i)=>`
    <div style="margin-bottom:10px;">
      <div style="display:flex;justify-content:space-between;margin-bottom:3px;">
        <span style="font-size:12px;color:var(--text)">${r.label}</span>
        <span style="font-size:12px;font-weight:700;color:${palette[i%palette.length]}">${r.v} <span style="color:var(--t3);font-weight:400">(${r.pct}%)</span></span>
      </div>
      <div style="height:8px;background:var(--s2);border-radius:4px;overflow:hidden;">
        <div style="height:100%;width:${Math.round(r.v/maxV*100)}%;background:${palette[i%palette.length]};border-radius:4px;transition:width .6s"></div>
      </div>
    </div>`).join('');

  document.getElementById('ceo-exit').innerHTML=`
    <div class="sec">🚪 Exit Reason Breakdown</div>
    <div class="r2">
      <div class="card">
        <div class="card-t">Exit by Type — ${total} total exits</div>
        ${bars}
      </div>
      <div class="card">
        <div class="card-t">Exit Type Distribution</div>
        <div style="position:relative;height:220px"><canvas id="exitPieChart"></canvas></div>
      </div>
    </div>`;

  dc('exitPie');
  charts['exitPie'] = new Chart(document.getElementById('exitPieChart'),{
    type:'doughnut',
    data:{
      labels: reasons.map(r=>r.label),
      datasets:[{data:reasons.map(r=>r.v), backgroundColor:palette.slice(0,reasons.length), borderWidth:2, borderColor:'#fff'}]
    },
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{position:'right',labels:{color:'#52586a',font:{size:11},boxWidth:12}},tooltip:{enabled:true}}}
  });
}

async function loadBHConversion(yr, mo){
  const r = await fetch('/api/bh_conversion?'+ceoParams(yr,mo));
  const {rows} = await r.json();
  if(!rows || !rows.length){
    document.getElementById('ceo-bh').innerHTML='';
    return;
  }
  const pctColor = v => v>=60?'#1db85a':v>=30?'#f39c12':'#e8453c';
  const pctBadge = v => `<span style="font-size:11px;font-weight:700;color:${pctColor(v)}">${v}%</span>`;

  const tableRows = rows.map(r=>`
    <tr class="tr-r">
      <td style="padding-left:14px;font-weight:600">${r.bh}</td>
      <td>${r.dem.toLocaleString()}</td>
      <td>${r.sub.toLocaleString()}</td>
      <td>${r.l1.toLocaleString()}</td>
      <td>${r.sel.toLocaleString()}</td>
      <td>${r.ob.toLocaleString()}</td>
      <td>${pctBadge(r.dem_sub)}</td>
      <td>${pctBadge(r.sub_l1)}</td>
      <td>${pctBadge(r.l1_sel)}</td>
      <td>${pctBadge(r.sel_ob)}</td>
    </tr>`).join('');

  document.getElementById('ceo-bh').innerHTML=`
    <div class="sec">👤 Conversion Rates by Business Head</div>
    <div class="tbl-w" style="margin-bottom:14px">
      <table>
        <thead>
          <tr>
            <th class="tg2" style="text-align:left;border-left:none">Business Head</th>
            <th class="tg2">Demands</th><th class="tg2">Submissions</th>
            <th class="tg2">L1</th><th class="tg2">Selections</th><th class="tg2">Onboarded</th>
            <th class="tg2">Dem→Sub%</th><th class="tg2">Sub→L1%</th>
            <th class="tg2">L1→Sel%</th><th class="tg2">Sel→Ob%</th>
          </tr>
          <tr class="sh">
            <th style="text-align:left;padding-left:14px">BH</th>
            <th>Count</th><th>Count</th><th>Count</th><th>Count</th><th>Count</th>
            <th>Rate</th><th>Rate</th><th>Rate</th><th>Rate</th>
          </tr>
        </thead>
        <tbody>${tableRows}</tbody>
      </table>
    </div>`;
}

async function updateLastUpdated(){
  const r=await fetch('/api/last_updated');
  const d=await r.json();
  document.getElementById('lastUpdated').textContent='Updated: '+d.time;
}

init();
</script>
</body>
</html>"""

# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    import webbrowser, threading, time
    port = 5050
    shareable_urls = get_shareable_urls(port)
    print("\n" + "="*50)
    print("  Joules to Watts  —  Recruitment Dashboard")
    print("="*50)
    print(f"  Data : {DATA_FOLDER}")
    print(f"  Local: http://localhost:{port}")
    for url in shareable_urls:
        if "localhost" not in url:
            print(f"  Share: {url}")
    print("  Ctrl+C to stop\n")
    def _open():
        time.sleep(1.2)
        webbrowser.open(f"http://localhost:{port}")
    threading.Thread(target=_open, daemon=True).start()
    app.run(host="0.0.0.0", port=port, debug=False)


