from functools import lru_cache
from io import BytesIO
import os
import re
import socket
import warnings
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
    "exit":     "exit.csv",
    "exitpipe": "exit_pipeline_data.csv",
}

_LAST_CACHE_SIGNATURE = None

COLUMN_MAPPING = {
    "selpipe": {
        "po": "p_o_value",
        "margin": "margin"
    },
    "ob": {
        "po": "p_o_value",
        "margin": "margin"
    },
    "exit": {
        "po": "p_o_value",
        "margin": "margin"
    },
    "exitpipe": {
        "po": "p_o_value",
        "margin": "margin"
    },
    "activehc": {
        "po": "p_o_value",
        "margin": "margin"
    },
    "sel": {
        "po": "po",
        "margin": "margin"
    }
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
        os.environ.get("DASHBOARDDATAFOLDER", "").strip(),
        REPO_DATA_FOLDER,
        os.path.join(os.path.dirname(__file__), "data"),
        APP_ROOT,
        WINDOWS_DATA_FOLDER,
    ]
    best_folder = os.path.join(os.path.dirname(__file__), "data")
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
    "demand":   "created_at",
    "sub":      "date",
    "intv":     "interview_date",
    "sel":      "selection_date",
    "selpipe":  "display_date",
    "ob":       "display_date",
    "activehc": "display_date",
    "exit":     "last_work_day",
    "exitpipe": "tentative_exit_date",
}

CLIENT_COLS = {
    "demand":   "company_name",
    "sub":      "client",
    "intv":     "company_name",
    "sel":      "company_name",
    "selpipe":  "company_name",
    "ob":       "company_name",
    "activehc": "company_name",
    "exit":     "company_name",
    "exitpipe": "company_name",
}


def _file_signature(path):
    if not path:
        return ("", False, 0, 0)
    abs_path = os.path.abspath(path)
    try:
        stat = os.stat(abs_path)
        return (abs_path, True, int(stat.st_mtime_ns), int(stat.st_size))
    except OSError:
        return (abs_path, False, 0, 0)


def get_runtime_cache_signature():
    data_folder = resolve_data_folder()
    mapping_file = resolve_mapping_file(data_folder)
    data_signature = (
        os.path.abspath(data_folder),
        tuple(
            (key, filename, *_file_signature(os.path.join(data_folder, filename)))
            for key, filename in sorted(FILE_PATHS.items())
        ),
    )
    mapping_signature = ("mapping", *_file_signature(mapping_file))
    return (data_signature, mapping_signature)


def clear_runtime_caches():
    global _LAST_CACHE_SIGNATURE, LAST_REFRESHED_AT

    for cached_func_name in (
        "_load_data_cached",
        "_load_mapping_cached",
        "_get_periods_cached",
        "_get_client_catalog_cached",
        "_resolve_client_filter_cached",
        "_compute_all_cached",
        "_mom_trends_cached",
        "_daily_trends_cached",
    ):
        cached_func = globals().get(cached_func_name)
        if cached_func is not None:
            cached_func.cache_clear()

    _LAST_CACHE_SIGNATURE = None
    LAST_REFRESHED_AT = datetime.now()


def refresh_runtime_caches_if_needed(force=False):
    global _LAST_CACHE_SIGNATURE, DATA_FOLDER, MAPPING_FILE, LAST_REFRESHED_AT

    current_signature = get_runtime_cache_signature()
    if force or current_signature != _LAST_CACHE_SIGNATURE:
        clear_runtime_caches()
        DATA_FOLDER = resolve_data_folder()
        MAPPING_FILE = resolve_mapping_file(DATA_FOLDER)
        _LAST_CACHE_SIGNATURE = current_signature
        LAST_REFRESHED_AT = datetime.now()
    return current_signature


def _current_data_signature():
    return refresh_runtime_caches_if_needed()[0]


def _current_mapping_signature():
    return refresh_runtime_caches_if_needed()[1]

ID_COL_CANDIDATES = ["id", "ID", "job_id", "Job_ID", "job_ID"]
OPENING_COL_CANDIDATES = [
    "no.of opening", "No.of opening", "No.Of Opening",
    "no_of_openings", "No_of_openings", "no_of_opening", "No_of_opening",
    "openings", "Openings", "opening_count", "Opening_count",
    "number_of_openings", "Number_of_openings",
    "no_of_positions", "No_of_positions", "positions", "Positions",
    "vacancies", "Vacancies",
]
PO_COL_KEYS = {"selpipe", "ob", "exit", "exitpipe", "activehc"}
RAW_DATASET_CONFIG = {
    "demand": {"label": "Demands"},
    "sub": {"label": "Submissions"},
    "intv": {"label": "Interviews"},
    "sel": {"label": "Selections"},
    "selpipe": {"label": "Selection Pipeline"},
    "ob": {"label": "Onboardings"},
    "activehc": {"label": "Active Headcount"},
    "exitpipe": {"label": "Exit Pipeline"},
    "exit": {"label": "Exit"},
}


DATE_PARSE_FORMATS = (
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%d %I:%M %p",
    "%Y-%m-%d",
    "%d-%m-%Y %H:%M:%S",
    "%d-%m-%Y %I:%M %p",
    "%d-%m-%Y",
    "%d/%m/%Y %H:%M:%S",
    "%d/%m/%Y %I:%M %p",
    "%d/%m/%Y",
)


def parse_datetime_series(values, dayfirst=False):
    series = pd.Series(values)
    parsed = pd.Series(pd.NaT, index=series.index, dtype="datetime64[ns]")
    remaining = series.notna() & series.astype(str).str.strip().ne("")

    for date_format in DATE_PARSE_FORMATS:
        if not remaining.any():
            break
        candidates = pd.to_datetime(series[remaining], errors="coerce", format=date_format)
        matched = candidates.notna()
        if matched.any():
            parsed.loc[candidates[matched].index] = candidates[matched]
            remaining.loc[candidates[matched].index] = False

    if remaining.any():
        try:
            candidates = pd.to_datetime(
                series[remaining],
                errors="coerce",
                format="mixed",
                dayfirst=dayfirst,
            )
        except (TypeError, ValueError):
            with warnings.catch_warnings():
                warnings.simplefilter("ignore", UserWarning)
                candidates = pd.to_datetime(series[remaining], errors="coerce", dayfirst=dayfirst)

        matched = candidates.notna()
        if matched.any():
            parsed.loc[candidates[matched].index] = candidates[matched]

    if getattr(parsed.dt, "tz", None) is not None:
        parsed = parsed.dt.tz_localize(None)
    return parsed


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
    df.columns = df.columns.str.lower()

    client_col = CLIENT_COLS.get(file_key)
    if client_col and client_col in df.columns:
        df[client_col] = df[client_col].str.strip()

    date_col = DATE_COLS.get(file_key)
    if file_key == "selpipe":
        parsed = pd.Series(pd.NaT, index=df.index)
        # Selection pipeline metrics should follow the display date for dashboard filters.
        for candidate in ["display_date", "offer_created_date", "selection_date"]:
            if candidate in df.columns:
                candidate_parsed = parse_datetime_series(df[candidate])
                parsed = parsed.where(parsed.notna(), candidate_parsed)
    elif date_col and date_col in df.columns:
        parsed = parse_datetime_series(df[date_col])
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
        po_end_parsed = parse_datetime_series(df["po_end_date"])
        df["_po_end_date"] = po_end_parsed
    else:
        df["_po_end_date"] = pd.NaT

    if file_key in PO_COL_KEYS:
        mapping = COLUMN_MAPPING.get(file_key, {})

        po_col = mapping.get("po")
        mg_col = mapping.get("margin")

        df["_po"] = _flt(df[po_col]) if po_col in df.columns else pd.Series(0.0, index=df.index)
        df["_mg"] = _flt(df[mg_col]) if mg_col in df.columns else pd.Series(0.0, index=df.index)
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
    global DATA_FOLDER, MAPPING_FILE
    DATA_FOLDER = resolve_data_folder()
    MAPPING_FILE = resolve_mapping_file(DATA_FOLDER)
    data = {}
    for key, filename in FILE_PATHS.items():
        path = os.path.join(DATA_FOLDER, filename)
        if os.path.exists(path):
            try:
                if filename.endswith(".xlsx"):
                    df = pd.read_excel(path, dtype=str).fillna("")
                else:
                    try:
                        df = pd.read_csv(path, dtype=str).fillna("")
                    except UnicodeDecodeError:
                        df = pd.read_csv(path, dtype=str, encoding="latin1").fillna("")
                data[key] = _prepare_frame(df, key)
                print(f"[OK] Loaded {filename} from {path} with {len(data[key])} rows")
            except Exception as e:
                print(f"[ERROR] {filename}: {e}")
                data[key] = pd.DataFrame()
        else:
            print(f"[MISSING] {path}")
            data[key] = pd.DataFrame()
    return data

@lru_cache(maxsize=4)
def _load_data_cached(data_signature):
    return load_data()


def load_data_cached():
    return _load_data_cached(_current_data_signature())


load_data_cached.cache_clear = clear_runtime_caches

# ─────────────────────────────────────────────────────────────────────────────
# CLIENT MAPPING
# ─────────────────────────────────────────────────────────────────────────────

def _mapping_column(df, candidates):
    normalized = {
        re.sub(r"[^a-z0-9]+", "", str(column).strip().lower()): column
        for column in df.columns
    }
    for candidate in candidates:
        match = normalized.get(re.sub(r"[^a-z0-9]+", "", candidate.strip().lower()))
        if match:
            return match
    return None


def _load_mapping_impl():
    client_to_domain = {}
    client_to_bh = {}

    if not MAPPING_FILE or not os.path.exists(MAPPING_FILE):
        return client_to_domain, client_to_bh, [], []

    try:
        mapping_df = pd.read_excel(MAPPING_FILE, dtype=str)
    except Exception as exc:
        print(f"[WARN] Unable to read mapping file '{MAPPING_FILE}': {exc}")
        return client_to_domain, client_to_bh, [], []

    if mapping_df.empty:
        return client_to_domain, client_to_bh, [], []

    mapping_df = mapping_df.copy()
    mapping_df.columns = [str(col).strip() for col in mapping_df.columns]

    client_col = _mapping_column(mapping_df, ["client", "client name", "company", "company name"])
    domain_col = _mapping_column(mapping_df, ["domain", "client domain"])
    bh_col = _mapping_column(mapping_df, ["bh tag", "bh", "business head", "business head tag"])

    if not client_col:
        return client_to_domain, client_to_bh, [], []

    for _, row in mapping_df.iterrows():
        client_name = str(row.get(client_col, "") or "").strip()
        if not client_name or client_name.lower() in {"nan", "none"}:
            continue

        domain = normalize_domain_label(row.get(domain_col, "")) if domain_col else ""
        bh = normalize_bh_label(row.get(bh_col, "")) if bh_col else ""

        if domain.lower() == "none":
            domain = ""
        if bh.lower() == "none":
            bh = ""

        if client_name not in client_to_domain or domain:
            client_to_domain[client_name] = domain
        if client_name not in client_to_bh or bh:
            client_to_bh[client_name] = bh

    domains = sorted({value for value in client_to_domain.values() if value}, key=str.lower)
    business_heads = sorted({value for value in client_to_bh.values() if value}, key=str.lower)
    return client_to_domain, client_to_bh, domains, business_heads


@lru_cache(maxsize=4)
def _load_mapping_cached(mapping_signature):
    return _load_mapping_impl()


def load_mapping():
    return _load_mapping_cached(_current_mapping_signature())


load_mapping.cache_clear = clear_runtime_caches


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
    if value is None or value == "":
        return None
    return pd.Timestamp(value)

def get_period_filters():
    from datetime import datetime

    today = datetime.today()

    years = parse_int_csv_filter_arg("years")
    months = parse_int_csv_filter_arg("months")

    # ✅ DEFAULT YEAR
    if years is None:
        single_year = request.args.get("year", type=int)
        years = {single_year} if single_year else {today.year}

    # ✅ DEFAULT MONTH
    if months is None:
        single_month = request.args.get("month", type=int)
        months = {single_month} if single_month else {today.month}

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


def filter_id_status_zero(df):
    if df.empty or "id_status" not in df.columns:
        return df
    return df[df["id_status"].astype(str).str.strip() == "0"]


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

    return filter_id_status_zero(unsubmitted)


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
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("₹", "", regex=False)
        .str.replace("INR", "", regex=False)
        .str.replace("L", "00000", regex=False)   # handles "2L"
        .str.replace(" ", "", regex=False)
        .replace({"": None, "nan": None, "None": None}),
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


@lru_cache(maxsize=4)
def _get_periods_cached(data_signature):
    return tuple(collect_periods(load_data_cached()))


def get_periods_cached():
    return _get_periods_cached(_current_data_signature())


get_periods_cached.cache_clear = clear_runtime_caches


@lru_cache(maxsize=4)
def _get_client_catalog_cached(data_signature, mapping_signature):
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


def get_client_catalog():
    return _get_client_catalog_cached(_current_data_signature(), _current_mapping_signature())


get_client_catalog.cache_clear = clear_runtime_caches


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
    overdue_hc=0,overdue_po=0.0,overdue_mg=0.0
)
def apply_date_filter(df, from_date, to_date):
    if df is None or df.empty:
        return df
    # ✅ PRIORITIZE CORRECT DATE COLUMN
    if "display_date" in df.columns:
        col = "display_date"
    elif "created_date" in df.columns:
        col = "created_date"
    elif "date" in df.columns:
        col = "date"
    else:
        return df
    df[col] = pd.to_datetime(df[col], dayfirst=True, errors="coerce")
    if from_date:
        df = df[df[col] >= pd.Timestamp(from_date)]
    if to_date:
        df = df[df[col] <= pd.Timestamp(to_date)]
    return df


def resolve_activehc_cutoff(sel_year, sel_month, from_date=None, to_date=None):
    if to_date is not None:
        cutoff = pd.Timestamp(to_date)
        return cutoff.tz_localize(None) if cutoff.tzinfo else cutoff

    years = set(sel_year) if isinstance(sel_year, (set, tuple, list)) else ({sel_year} if sel_year else set())
    months = set(sel_month) if isinstance(sel_month, (set, tuple, list)) else ({sel_month} if sel_month else set())

    if years and months:
        return max(
            pd.Timestamp(year=int(year), month=int(month), day=1) + pd.offsets.MonthEnd(1)
            for year in years
            for month in months
        )
    if years:
        return pd.Timestamp(year=max(int(year) for year in years), month=12, day=31)
    if from_date is not None:
        cutoff = pd.Timestamp(from_date)
        return cutoff.tz_localize(None) if cutoff.tzinfo else cutoff
    return None


def cumulative_activehc_counts(df, date_col, cl_col, client_filter=None, from_date=None, to_date=None, grain="day"):
    if df.empty:
        return {}
    if "_date" in df.columns:
        df = df[df["_date"].notna()].copy()
    elif date_col in df.columns:
        df = df.copy()
        dates = pd.to_datetime(df[date_col], errors="coerce")
        if getattr(dates.dt, "tz", None) is not None:
            dates = dates.dt.tz_localize(None)
        df["_date"] = dates
        df = df[df["_date"].notna()]
    else:
        return {}

    if df.empty:
        return {}
    if client_filter is not None and cl_col and cl_col in df.columns:
        df = df[df[cl_col].isin(client_filter)]
    if df.empty:
        return {}

    date_series = df["_date"].dt.normalize()
    if to_date is not None:
        upper = pd.Timestamp(to_date)
        upper = upper.tz_localize(None) if upper.tzinfo else upper
        date_series = date_series[date_series <= upper.normalize()]
    if date_series.empty:
        return {}

    if grain == "month":
        start_period = (pd.Timestamp(from_date) if from_date is not None else date_series.min()).to_period("M")
        end_period = (pd.Timestamp(to_date) if to_date is not None else date_series.max()).to_period("M")
        periods = pd.period_range(start=start_period, end=end_period, freq="M")
        additions = date_series.dt.to_period("M").value_counts().reindex(periods, fill_value=0).sort_index()
        base_total = int((df["_date"].dt.normalize() < start_period.to_timestamp()).sum())
        cumulative = additions.cumsum() + base_total
        return {str(period): int(value) for period, value in cumulative.items()}

    start_day = (pd.Timestamp(from_date) if from_date is not None else date_series.min()).normalize()
    end_day = (pd.Timestamp(to_date) if to_date is not None else date_series.max()).normalize()
    days = pd.date_range(start=start_day, end=end_day, freq="D")
    additions = date_series.value_counts().reindex(days, fill_value=0).sort_index()
    base_total = int((df["_date"].dt.normalize() < start_day).sum())
    cumulative = additions.cumsum() + base_total
    return {day.strftime("%d %b %Y"): int(value) for day, value in cumulative.items()}

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

        id_col_dem = "_id_norm" if "_id_norm" in df.columns else next((c for c in ID_COL_CANDIDATES if c in df.columns), None)
        if id_col_dem:
            demand_ids = df[id_col_dem].dropna().astype(str).str.strip()
            unserviced_df = df.loc[~demand_ids.isin(submitted_ids)]
            unserviced_df = filter_id_status_zero(unserviced_df)
            unserviced_counts = unserviced_df.groupby(cl_col).size().to_dict()
        for cl, g in df.groupby(cl_col):
            ensure(cl)
            total_dem = len(g)
            unserv = int(unserviced_counts.get(cl, 0))
            res[cl]["dem"] += total_dem
            res[cl]["dem_open"] += float(g["_openings"].sum()) if "_openings" in g.columns else total_dem
            res[cl]["dem_u"] += unserv
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

    # SELECTION PIPEPLINE
    df = frames["selpipe"]

    cl_col = None
    if not df.empty:
        cl_col = next((c for c in ["Company_name", "company_name", "client", "Client"] if c in df.columns), None)

    if not df.empty and cl_col:

        if client_filter is not None:
            df = df[df[cl_col].isin(client_filter)]

        today = pd.Timestamp.now().normalize()

        for cl, g in df.groupby(cl_col):
            ensure(cl)

            # Pipeline metrics
            res[cl]["sp_hc"] += len(g)
            res[cl]["sp_po"] += float(g["_po"].sum())
            res[cl]["sp_mg"] += float(g["_mg"].sum())

            # 🔥 FIXED OVERDUE LOGIC
            g["_date"] = pd.to_datetime(g["_date"], errors="coerce")

            overdue_df = g[
                (g["_date"].notna()) &
                (g["_date"].dt.normalize() < today)
            ]

            res[cl]["overdue_hc"] += len(overdue_df)
            res[cl]["overdue_po"] += float(overdue_df["_po"].sum())
            res[cl]["overdue_mg"] += float(overdue_df["_mg"].sum())
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
    # Count cumulatively by display_date only. Active HC ignores lower-bound
    # period filters and uses the selected period only as an upper cutoff.
    df = data["activehc"]
    activehc_cutoff = resolve_activehc_cutoff(sel_year, sel_month, from_date, to_date)

    cl_col = None
    if not df.empty:
        cl_col = next((c for c in ["company_name", "Company_name", "client", "Client"] if c in df.columns), None)

    if not df.empty and cl_col:
        df = df[df["_date"].notna()].copy()
        if activehc_cutoff is not None:
            df = df[df["_date"] <= activehc_cutoff]
        if client_filter is not None:
            df = df[df[cl_col].isin(client_filter)]

        for cl, g in df.groupby(cl_col):
            ensure(cl)
            res[cl]["active_hc"] += len(g)
            if "_po" in g.columns:
                res[cl]["active_po"] = res[cl].get("active_po", 0) + float(g["_po"].sum())
            if "_mg" in g.columns:
                res[cl]["active_mg"] = res[cl].get("active_mg", 0) + float(g["_mg"].sum())
    

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
            "active_po", "active_mg", "overdue_po", "overdue_mg"
        ]:
            m[k] = m[k] / 1e5  # Always convert to Lakhs

        m["net_hc"] = m["ob_hc"] - m["ex_hc"]
        m["net_po"] = m["ob_po"] - m["ex_po"]
        m["net_mg"] = m["ob_mg"] - m["ex_mg"]

    return res


@lru_cache(maxsize=256)
def _compute_all_cached(data_signature, sel_year, sel_month, client_filter_key=None, from_date_key=None, to_date_key=None):
    return compute_all(
        load_data_cached(),
        sel_year,
        sel_month,
        thaw_filter(client_filter_key),
        thaw_date(from_date_key),
        thaw_date(to_date_key),
    )


def compute_all_cached(sel_year, sel_month, client_filter_key=None, from_date_key=None, to_date_key=None):
    return _compute_all_cached(
        _current_data_signature(),
        sel_year,
        sel_month,
        client_filter_key,
        freeze_date(from_date_key) if not isinstance(from_date_key, str) else from_date_key,
        freeze_date(to_date_key) if not isinstance(to_date_key, str) else to_date_key,
    )


compute_all_cached.cache_clear = clear_runtime_caches


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

    return (cl_filter & allowed) if cl_filter else allowed


@lru_cache(maxsize=256)
def _resolve_client_filter_cached(data_signature, mapping_signature, cl_filter_key=None, dom_filter_key=None, bh_filter_key=None):
    return freeze_filter(
        resolve_client_filter(
            thaw_filter(cl_filter_key),
            thaw_filter(dom_filter_key),
            thaw_filter(bh_filter_key),
        )
    )


def resolve_client_filter_cached(cl_filter_key=None, dom_filter_key=None, bh_filter_key=None):
    return _resolve_client_filter_cached(
        _current_data_signature(),
        _current_mapping_signature(),
        cl_filter_key,
        dom_filter_key,
        bh_filter_key,
    )


resolve_client_filter_cached.cache_clear = clear_runtime_caches


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
    if not df.empty and "company_name" in df.columns:
        df = df[df["company_name"].apply(include)]
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
def _mom_trends_cached(data_signature, client_filter_key=None):
    return mom_trends(load_data_cached(), thaw_filter(client_filter_key))


def mom_trends_cached(client_filter_key=None):
    return _mom_trends_cached(_current_data_signature(), client_filter_key)


mom_trends_cached.cache_clear = clear_runtime_caches


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

    dem_counts  = count_by_date(dem_df,  "created_at",     cl(dem_df,  ["company_name","client"]))
    sub_counts  = count_by_date(sub_df,  "date",           cl(sub_df,  ["client","Client"]))
    intv_counts = count_by_date(intv_df, "interview_date", cl(intv_df, ["company_name","client"]))
    sel_counts  = count_by_date(sel_df,  "selection_date", cl(sel_df,  ["company_name","client"]))
    ob_counts   = count_by_date(ob_df,   "display_date",   cl(ob_df,   ["company_name","client"]))
    hc_counts   = cumulative_activehc_counts(
        hc_df,
        "display_date",
        cl(hc_df, ["company_name", "client"]),
        client_filter=client_filter,
        from_date=from_ts,
        to_date=to_ts,
        grain=grain,
    )
    ex_counts   = count_by_date(ex_df,   "last_work_day",  cl(ex_df,   ["company_name","client"]))

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
            if "created_at" in work_dem.columns:
                work_dem["_date"] = pd.to_datetime(work_dem["created_at"], errors="coerce")
            else:
                work_dem["_date"] = pd.NaT
        work_dem = work_dem[work_dem["_date"].notna()]

        demand_client_col = cl(work_dem, ["company_name","client"])
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
            work_dem = filter_id_status_zero(work_dem)

        if not work_dem.empty:
            work_dem = work_dem.assign(__ds=period_key(work_dem["_date"]))
            dem_u_counts = work_dem.groupby("__ds").size().to_dict()

    all_counts = {"dem": dem_counts, "dem_u": dem_u_counts, "sub": sub_counts, "sub_fp": sub_fp_counts, "intv": intv_counts, "sel": sel_counts, "ob": ob_counts, "hc": hc_counts, "ex": ex_counts}
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
def _daily_trends_cached(data_signature, mapping_signature, client_filter_key=None, domain_filter_key=None, bh_filter_key=None, from_date_key=None, to_date_key=None, grain="day"):
    resolved_clients = resolve_client_filter_cached(
        client_filter_key,
        domain_filter_key,
        bh_filter_key
    )

    return daily_trends(
        load_data_cached(),
        thaw_filter(resolved_clients),
        thaw_date(from_date_key),
        thaw_date(to_date_key),
        grain=grain,
    )


def daily_trends_cached(client_filter_key=None, domain_filter_key=None, bh_filter_key=None, from_date_key=None, to_date_key=None, grain="day"):
    return _daily_trends_cached(
        _current_data_signature(),
        _current_mapping_signature(),
        client_filter_key,
        domain_filter_key,
        bh_filter_key,
        freeze_date(from_date_key) if not isinstance(from_date_key, str) else from_date_key,
        freeze_date(to_date_key) if not isinstance(to_date_key, str) else to_date_key,
        grain,
    )


daily_trends_cached.cache_clear = clear_runtime_caches


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

    cl_filter, dom_filter, bh_filter = get_resolved_filters()

    result = daily_trends_cached(
        freeze_filter(cl_filter),
        freeze_filter(dom_filter),
        freeze_filter(bh_filter),
        freeze_date(from_date),
        freeze_date(to_date),
        grain,
    )

    # Fallback to full dataset if date filter returns nothing
    total_all = sum(sum(x['v'] for x in v) for v in result.values())
    if total_all == 0 and (from_date is not None or to_date is not None):
        result = daily_trends_cached(
            freeze_filter(cl_filter),
            freeze_filter(dom_filter),
            freeze_filter(bh_filter),
            grain=grain,
        )

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
        freeze_filter({sel_year}),
        freeze_filter({sel_month}),
        freeze_filter(cl_filter),
        freeze_date(current_start),
        freeze_date(today),
        )
    prev_res = compute_all_cached(
        freeze_filter({prev_start.year}),
        freeze_filter({prev_start.month}),
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

    if dem_df.empty or "created_at" not in dem_df.columns:
        return jsonify({"buckets": [], "top_clients": []})

    submitted_ids = set()
    id_col = next((c for c in ["job_ID","job_id","Job_ID","ID","id"] if c in sub_df.columns), None)
    if id_col:
        submitted_ids = set(sub_df[id_col].astype(str).str.strip().unique())

    dem_df = dem_df.copy()
    dem_df["_date"] = pd.to_datetime(dem_df["created_at"], errors="coerce")
    today = pd.Timestamp.now().normalize()

    id_col_d = next((c for c in ["id","ID","job_id","Job_ID"] if c in dem_df.columns), None)
    if id_col_d:
        unserviced = dem_df[~dem_df[id_col_d].astype(str).str.strip().isin(submitted_ids)].copy()
    else:
        unserviced = dem_df.copy()

    unserviced = filter_id_status_zero(unserviced)
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


