import pandas as pd

# =========================
# ZERO TEMPLATE
# =========================
ZERO = dict(
    dem=0, dem_u=0,
    sub=0, sub_fp=0,
    l1=0, l2=0, l3=0,
    sel=0,
    ob_hc=0, ob_po=0.0, ob_mg=0.0,
    ex_hc=0, ex_po=0.0, ex_mg=0.0,
    active_hc=0, active_po=0.0, active_mg=0.0,
    net_hc=0, net_po=0.0, net_mg=0.0
)

# =========================
# MAIN FUNCTION
# =========================
def compute_all(data, year=None, month=None, client_filter=None):

    res = {}

    def ensure(cl):
        if cl not in res:
            res[cl] = ZERO.copy()

    # =========================
    # DEMAND
    # =========================
    df = data["demand"]

    if year:
        df = df[df["_date"].dt.year == year]
    if month:
        df = df[df["_date"].dt.month == month]

    for cl, g in df.groupby("company_name"):
        ensure(cl)
        res[cl]["dem"] += len(g)
        res[cl]["dem_u"] += (g["Status"] == 0).sum()

    # =========================
    # SUBMISSION
    # =========================
    df = data["submission"]

    if year:
        df = df[df["_date"].dt.year == year]
    if month:
        df = df[df["_date"].dt.month == month]

    for cl, g in df.groupby("company_name"):
        ensure(cl)
        res[cl]["sub"] += len(g)
        res[cl]["sub_fp"] += (g["feedback_status"] == "Pending").sum()

    # =========================
    # INTERVIEW
    # =========================
    df = data["interview"]

    if year:
        df = df[df["_date"].dt.year == year]
    if month:
        df = df[df["_date"].dt.month == month]

    for cl, g in df.groupby("company_name"):
        ensure(cl)
        res[cl]["l1"] += (g["round"] == "L1").sum()
        res[cl]["l2"] += (g["round"] == "L2").sum()
        res[cl]["l3"] += (g["round"] == "L3").sum()

    # =========================
    # SELECTION
    # =========================
    df = data["selection"]

    if year:
        df = df[df["_date"].dt.year == year]
    if month:
        df = df[df["_date"].dt.month == month]

    for cl, g in df.groupby("company_name"):
        ensure(cl)
        res[cl]["sel"] += len(g)

    # =========================
    # ONBOARDING
    # =========================
    df = data["onboarding"]

    if year:
        df = df[df["_date"].dt.year == year]
    if month:
        df = df[df["_date"].dt.month == month]

    for cl, g in df.groupby("company_name"):
        ensure(cl)
        res[cl]["ob_hc"] += len(g)
        res[cl]["ob_po"] += g["p_o_value"].sum()
        res[cl]["ob_mg"] += g["margin"].sum()

    # =========================
    # EXIT
    # =========================
    df = data["exit"]

    if year:
        df = df[df["_date"].dt.year == year]
    if month:
        df = df[df["_date"].dt.month == month]

    for cl, g in df.groupby("company_name"):
        ensure(cl)
        res[cl]["ex_hc"] += len(g)
        res[cl]["ex_po"] += g["p_o_value"].sum()
        res[cl]["ex_mg"] += g["margin"].sum()

    # =========================
    # ACTIVE HC
    # =========================
    df = data["active"]

    for cl, g in df.groupby("company_name"):
        ensure(cl)
        res[cl]["active_hc"] += len(g)
        res[cl]["active_po"] += g["p_o_value"].sum()
        res[cl]["active_mg"] += g["margin"].sum()

    # =========================
    # NET CALCULATION
    # =========================
    for cl in res:
        res[cl]["net_hc"] = res[cl]["ob_hc"] - res[cl]["ex_hc"]
        res[cl]["net_po"] = res[cl]["ob_po"] - res[cl]["ex_po"]
        res[cl]["net_mg"] = res[cl]["ob_mg"] - res[cl]["ex_mg"]

    return res