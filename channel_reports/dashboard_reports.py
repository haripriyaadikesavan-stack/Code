# %% [markdown]
# # Channel Analytics — Dashboard reports in Excel
#
# Writes the analysis of the Retail and HS dashboards as six Excel workbooks:
#
# | File | Scope | Values |
# |---|---|---|
# | `Retail_Analysis_Gross_<Mon>_<year>.xlsx` | Retail | Gross sales |
# | `Retail_Analysis_Net_<Mon>_<year>.xlsx` | Retail | Net sales (your rebates) |
# | `HS_with_FD_Analysis_Gross_…` / `…_Net_…` | HS (Retail + Project), FD included | Gross / Net |
# | `HS_without_FD_Analysis_Gross_…` / `…_Net_…` | HS (Retail + Project), FD left out | Gross / Net |
#
# **How to run:** set the rebates and the data source in cell 1, then *Run All*.
# The data are the dashboard input files written by `Channel_Analytics_v9.ipynb`
# (Retail_Sales_OB_Backlog_Combined.xlsx and H&D_Overall_Sales_OB_Backlog_Combined_ppcfd removed.xlsx),
# the Material Master, the Data Dictionary and the target files — the same files the dashboards read.
# Each workbook starts with an **About** sheet (period, basis, rebates, rules) and an **Overview** sheet.

# %%
# ════════════════════════════════════════════════════════════════════
# 0 · SETUP — installs anything missing (run once per machine)
# ════════════════════════════════════════════════════════════════════
import importlib
import subprocess
import sys

for pip_name, import_name in {"pandas": "pandas", "numpy": "numpy", "openpyxl": "openpyxl", "xlsxwriter": "xlsxwriter",
                              "pyarrow": "pyarrow", "msal": "msal", "requests": "requests", "python-dotenv": "dotenv",
                              "python-calamine": "python_calamine"}.items():
    try:
        importlib.import_module(import_name)
    except ImportError:
        print(f"  ⏳ installing {pip_name} ...", end=" ", flush=True)
        subprocess.check_call([sys.executable, "-m", "pip", "install", pip_name, "--quiet"])
        print("done")
print("✅ Packages ready")

# %%
# ════════════════════════════════════════════════════════════════════
# 1 · SETTINGS — the only cell you normally edit
# ════════════════════════════════════════════════════════════════════
import io
import os
import re
import time
from datetime import datetime, timedelta

import numpy as np
import pandas as pd

# --- Rebates for the Net files: Net = Gross × (1 − rebate), by Retail / Project and year ------------
REBATES = {
    "Retail": {2026: 0.135, 2025: 0.122, 2024: 0.145, 2023: 0.14, 2022: 0.13},
    "Project": {2026: 0.051, 2025: 0.045, 2024: 0.045, 2023: 0.045, 2022: 0.045},
}

# --- Period -------------------------------------------------------------------------------------------
AS_OF = None                 # None = the last invoice date in the data; or a date in quotes, e.g. "2026-10-31"

# --- Where the input files are -----------------------------------------------------------------------
SOURCE = "sharepoint"        # "sharepoint" (needs the .env with SP_TENANT_ID / SP_CLIENT_ID / SP_CLIENT_SECRET)
                             # or "local" (files in LOCAL_DIR, e.g. your OneDrive-synced folder)
ENV_FILE_PATH = ""           # .env location if it is not next to this notebook
LOCAL_DIR = r""              # folder that holds the files below when SOURCE = "local"

SITE_MAIN = "schneiderelectric.sharepoint.com:/sites/HDIndiaTeamGroup"
_ROOT = "Test - haripriya"
_MASTER = f"{_ROOT}/Data Downloads - Input Files/Master data"
_DASH = f"{_ROOT}/Dashboard Input Files/H&D Channel"
INPUTS = {   # key: (SharePoint path, local file name)
    "retail": (f"{_DASH}/Retail_Sales_OB_Backlog_Combined.xlsx", "Retail_Sales_OB_Backlog_Combined.xlsx"),
    "hs": (f"{_DASH}/H&D_Overall_Sales_OB_Backlog_Combined_ppcfd removed.xlsx",
           "H&D_Overall_Sales_OB_Backlog_Combined_ppcfd removed.xlsx"),
    "material": (f"{_MASTER}/Material Master Data.xlsx", "Material Master Data.xlsx"),
    "dictionary": (f"{_MASTER}/Data Dictionary.xlsx", "Data Dictionary.xlsx"),
    "targets": (f"{_MASTER}/Master Targets File.xlsx", "Master Targets File.xlsx"),
    "qtod": (f"{_MASTER}/Targets File/QTOD and ATOD file.xlsx", "QTOD and ATOD file.xlsx"),
    "hs_targets": (f"{_MASTER}/Master Targets File (Retail+Project).xlsx", "Master Targets File (Retail+Project).xlsx"),
}

# --- Output -------------------------------------------------------------------------------------------
OUTPUT_DIR = "./Dashboard Reports"
UPLOAD_TO_SHAREPOINT = False
SP_OUTPUT_FOLDER = f"{_ROOT}/Final Reports/Dashboard Reports"
CACHE_DIR = "./local_cache_reports"     # parsed input files are cached here (re-read only when a file changes)

# --- Rules (as the dashboards) --------------------------------------------------------------------------
FD_HELIOS = ["HDCFD", "PPCFD", "HDFDA", "HDFDB", "HDFDS"]
MANUAL_OLD_PARTNER_IDS = {"500073231"}   # always "Old" in the policy tracker
FLAT_NEW_PARTNER_TARGET = 2_500_000      # policy tracker: assigned target of a new partner (gross)
BACKLOG_OLD_DAYS = 50                     # backlog older than this is flagged
UNIT, UNIT_LABEL = 1e6, "Mn"              # values in the reports: INR Million
FOCUS_GROUPS = ["Wiser", "Miluz Zeta", "Unica Pure", "Livia", "AvatarOn", "Easy9 6kA", "Easy9 10kA", "Easy9 RCCB", "Easy9 DB"]

for _d in (OUTPUT_DIR, CACHE_DIR):
    os.makedirs(_d, exist_ok=True)
print(f"✅ Settings | source: {SOURCE} | Net rebates {REBATES['Retail']} (Retail), {REBATES['Project']} (Project)")

# %%
# ════════════════════════════════════════════════════════════════════
# 2 · SHAREPOINT (only when SOURCE = "sharepoint") — same client as Channel_Analytics_v9
# ════════════════════════════════════════════════════════════════════
#@@SHAREPOINT@@

def _load_env():
    from dotenv import find_dotenv, load_dotenv
    path = ENV_FILE_PATH or find_dotenv(usecwd=True)
    if path:
        load_dotenv(path, override=False)
    missing = [k for k in ("SP_TENANT_ID", "SP_CLIENT_ID", "SP_CLIENT_SECRET") if not os.environ.get(k)]
    if missing:
        raise EnvironmentError(f"SharePoint keys missing: {', '.join(missing)} — put them in the .env file "
                               "(set ENV_FILE_PATH) or use SOURCE = 'local'")


sp = None
if SOURCE == "sharepoint" or UPLOAD_TO_SHAREPOINT:
    _load_env()
    sp = SharePoint()
    print("✅ SharePoint client ready")

# %%
# ════════════════════════════════════════════════════════════════════
# 3 · REPORT STYLE — same look as the Channel Analytics reports
# ════════════════════════════════════════════════════════════════════
#@@STYLE@@

# %%
# ════════════════════════════════════════════════════════════════════
# 4 · TARGET RULES — organize_targets.py (the rules behind Channel_Targets_<year>.xlsx)
# ════════════════════════════════════════════════════════════════════
import types

#@@TARGET_TOOLS@@

TS = _target_split_tools()
print("✅ Target rules ready")

# %%
# ════════════════════════════════════════════════════════════════════
# 5 · READ THE INPUT FILES (cached locally; re-read only when a file changes)
# ════════════════════════════════════════════════════════════════════
import hashlib


def _file_bytes(key):
    sp_path, local_name = INPUTS[key]
    if SOURCE == "local":
        with open(os.path.join(LOCAL_DIR, local_name), "rb") as fh:
            return fh.read()
    return sp.download(sp_path).getvalue()


def read_book(key, sheet_name=None):
    """All sheets (or `sheet_name`) of an input file as {sheet: DataFrame}, cached as pickle by content."""
    data = _file_bytes(key)
    digest = hashlib.md5(data).hexdigest()[:12]
    cache = os.path.join(CACHE_DIR, f"{key}_{digest}.pkl")
    if os.path.exists(cache):
        return pd.read_pickle(cache)
    t0 = time.time()
    try:
        book = pd.read_excel(io.BytesIO(data), sheet_name=sheet_name, engine="calamine")
    except Exception:
        book = pd.read_excel(io.BytesIO(data), sheet_name=sheet_name)
    if not isinstance(book, dict):
        book = {sheet_name: book}
    for name, df in book.items():
        df.columns = [str(c).strip() for c in df.columns]
    for old in [f for f in os.listdir(CACHE_DIR) if f.startswith(f"{key}_")]:
        os.remove(os.path.join(CACHE_DIR, old))
    pd.to_pickle(book, cache)
    print(f"   📥 {INPUTS[key][1]}: {sum(len(d) for d in book.values()):,} rows in {len(book)} sheet(s), {time.time() - t0:.0f}s")
    return book


def _stack(book, prefix):
    parts = [df for name, df in book.items() if str(name).lower().startswith(prefix.lower()) and len(df)]
    return pd.concat(parts, ignore_index=True) if parts else pd.DataFrame()

# %%
# ════════════════════════════════════════════════════════════════════
# 6 · PREPARE — one table per data type with the same columns for Retail and HS
#
#   Date · GTM (Retail / Project) · Branch · Region · Activity · Product Group · Family
#   Partner · Partner name · City · Partner type (RD / KRO / Project) · Tier · Gross · Qty
# Branch: the target branch (Retail: partner's branch; Project: sales-office branch);
# Region: the Branch Targets zone for Retail branches, else the input file's region.
# ════════════════════════════════════════════════════════════════════
def _clean_code(s):
    return s.astype(str).str.strip().str.upper().str.replace(r"\.0$", "", regex=True)


def load_masters():
    mm = read_book("material", "Sheet1")["Sheet1"]
    family = dict(zip(_clean_code(mm["Material"]), mm["Family Actual"]))
    d = read_book("dictionary")
    pm = d.get("Partner Master Data", pd.DataFrame())
    geo = d.get("Geography mapping", pd.DataFrame())
    ssa = d.get("SSA Mapping", pd.DataFrame())
    if len(pm):
        pm = pm.copy()
        pm["Partner"] = _clean_code(pm["Sold-to Party"])
        pm["City"] = pm["City"].astype(str).str.strip()
        if len(geo):
            g = geo[["City", "Branch"]].copy()
            g["City"] = g["City"].astype(str).str.strip()
            pm = pm.merge(g.drop_duplicates("City").rename(columns={"Branch": "Geo branch"}), on="City", how="left")
        pm = pm.drop_duplicates("Partner").set_index("Partner")
    ssa_map = {}
    if len(ssa) and {"Branch", "SSA"} <= set(ssa.columns):
        ssa_map = dict(zip(ssa["Branch"].astype(str).str.strip().str.upper(), ssa["SSA"]))
    return {"family": family, "partners": pm, "ssa": ssa_map}


def standardise(df, kind, M, region_of, gtm=None):
    """Input rows → the standard columns. kind = 'sales' | 'ob' | 'backlog'."""
    if df is None or df.empty:
        return pd.DataFrame(columns=["Date", "GTM", "Branch", "Region", "Activity", "Product Group", "Family", "Partner",
                                     "Partner name", "City", "Partner type", "Tier", "Gross", "Qty", "Order date"])
    out = pd.DataFrame(index=df.index)
    date_col = {"sales": "InvoiceDate", "ob": "Calendar Day", "backlog": None}[kind]
    out["Date"] = pd.to_datetime(df[date_col], errors="coerce") if date_col in df else pd.NaT
    out["Order date"] = pd.to_datetime(df["OrderCreationDate"], errors="coerce") if "OrderCreationDate" in df else pd.NaT
    out["GTM"] = gtm if gtm else (df["Target GTM"] if "Target GTM" in df else df.get("Classification", "Retail"))
    partner = _clean_code(df["SoldToPartyID"])
    out["Partner"] = partner
    pm = M["partners"]
    retail = out["GTM"] == "Retail"
    # branch: the target key column when the notebook wrote it, else partner master (Retail) / Branch (HS file)
    if "Target Branch" in df:
        branch = df["Target Branch"]
    else:
        branch = df["Branch"] if "Branch" in df else pd.Series(np.nan, index=df.index)
        if len(pm) and "Geo branch" in pm:
            branch = branch.where(~retail, partner.map(pm["Geo branch"]))
    names = {b: TS.canon_branch(b) for b in pd.Series(branch[retail]).dropna().unique()}
    out["Branch"] = branch.where(~retail, branch.map(names)).fillna("Unmapped")
    reg_in = df["Region"] if "Region" in df else pd.Series(np.nan, index=df.index)
    out["Region"] = out["Branch"].map(region_of).where(retail, reg_in).fillna(reg_in).fillna("Unmapped")
    helios = df["HeliosCode"].astype(str).str.strip().str.upper()
    out["Helios"] = helios
    out["Activity"] = helios.map(TS.ACTIVITY).fillna("Other")
    fam = _clean_code(df["MaterialCode"]).map(M["family"])
    out["Family"] = fam.fillna("Unmapped")
    pairs = {(a, f): TS.product_group(a, f) for a, f in set(zip(out["Activity"], fam))}
    out["Product Group"] = [pairs[(a, f)] or "Other" for a, f in zip(out["Activity"], fam)]
    out["Partner name"] = partner.map(pm["Patner"]) if len(pm) and "Patner" in pm else np.nan
    out["Partner name"] = out["Partner name"].fillna(df.get("SoldToPartyName")).fillna(partner)
    out["City"] = (partner.map(pm["City"]) if len(pm) else np.nan)
    out["City"] = out["City"].where(retail).fillna(df.get("PartnerCity")).astype(str).str.strip().str.title()
    l3 = df["PricingClassL3"].astype(str).str.strip() if "PricingClassL3" in df else pd.Series("", index=df.index)
    ptype = l3.map({"ERW": "RD", "EKR": "KRO"})
    if len(pm) and "RD/KRO" in pm:
        ptype = partner.map(pm["RD/KRO"].astype(str).str.strip().str.upper().map({"RD": "RD", "KRO": "KRO"})).fillna(ptype)
    out["Partner type"] = np.where(~retail, "Project", ptype.fillna("Other Retail"))
    tier_col = next((c for c in pm.columns if "tier" in c.lower()), None) if len(pm) else None
    out["Tier"] = partner.map(pm[tier_col]).fillna("Unmapped") if tier_col else "Unmapped"
    out["Gross"] = pd.to_numeric(df["Gross Sales"], errors="coerce").fillna(0.0)
    out["Qty"] = pd.to_numeric(df.get("BillingQuantity", 0), errors="coerce").fillna(0.0)
    out["Material"] = df.get("MaterialCode")
    out["Order"] = df.get("SalesOrderHeader")
    return out


def load_data():
    M = load_masters()
    book = read_book("targets")
    tmp = TS.Issues()
    _, _, _, _, _, region_of = TS.retail_targets(book, pd.Timestamp.today().year, tmp)
    r = read_book("retail")
    h = read_book("hs")
    D = {
        "retail_sales": standardise(_stack(r, "Sales"), "sales", M, region_of, "Retail"),
        "retail_ob": standardise(_stack(r, "OB"), "ob", M, region_of, "Retail"),
        "retail_bl": standardise(_stack(r, "Backlog"), "backlog", M, region_of, "Retail"),
        "hs_sales": standardise(_stack(h, "Sales"), "sales", M, region_of),
        "hs_ob": standardise(_stack(h, "OB"), "ob", M, region_of),
        "hs_bl": standardise(_stack(h, "Backlog"), "backlog", M, region_of),
    }
    D["masters"] = M
    D["books"] = {"targets": book, "qtod": next(iter(read_book("qtod").values())), "hs_targets": read_book("hs_targets")}
    for k in ("retail_sales", "hs_sales"):
        print(f"   {k}: {len(D[k]):,} rows, {D[k]['Date'].min():%d %b %Y} – {D[k]['Date'].max():%d %b %Y}")
    return D

# %%
# ════════════════════════════════════════════════════════════════════
# 7 · PERIOD, VALUES AND TARGETS
#
# Net = Gross × (1 − rebate of the row's GTM and year) — REBATES in cell 1.
# Targets in the target files are net: the Gross files show them grossed up with the same rebate
# (target ÷ (1 − rebate)); QTOD partner targets are gross and are netted the same way for the Net files.
# ════════════════════════════════════════════════════════════════════
def period(D):
    last = D["retail_sales"]["Date"].max()
    as_of = pd.Timestamp(AS_OF) if AS_OF else last
    as_of = min(as_of, max(last, D["hs_sales"]["Date"].max())).normalize()
    y = as_of.year

    def ly(d):
        try:
            return d.replace(year=d.year - 1)
        except ValueError:
            return d.replace(year=d.year - 1, day=28)
    return {"as_of": as_of, "year": y, "m_start": as_of.replace(day=1), "y_start": pd.Timestamp(y, 1, 1),
            "ly_as_of": ly(as_of), "ly_m_start": ly(as_of.replace(day=1)), "ly_y_start": pd.Timestamp(y - 1, 1, 1),
            "ly_end": pd.Timestamp(y - 1, 12, 31), "months_left": 12 - as_of.month + (0 if as_of.is_month_end else 1),
            "label": as_of.strftime("%d %b %Y"), "mon": as_of.strftime("%b")}


def rebate(gtm, year):
    table = REBATES.get(gtm, REBATES["Retail"])
    return table.get(int(year), table[max(table)]) if pd.notna(year) else table[max(table)]


def with_values(df, basis, P):
    """Copy of df with 'Value' = Gross or Net (rebate by GTM and year; backlog at this year's rate)."""
    out = df.copy()
    if basis == "Gross" or out.empty:
        out["Value"] = out["Gross"]
        return out
    year = out["Date"].dt.year.fillna(P["year"]).astype(int)
    rates = {(g, yy): rebate(g, yy) for g, yy in set(zip(out["GTM"], year))}
    out["Value"] = out["Gross"] * (1 - pd.Series([rates[k] for k in zip(out["GTM"], year)], index=out.index))
    return out


def targets(D, P):
    """All targets of the year (net, as in the files)."""
    y, issues = P["year"], TS.Issues()
    book, hs_book = D["books"]["targets"], D["books"]["hs_targets"]
    _, _, prod_m, yearly, _, region_of = TS.retail_targets(book, y, issues)
    mt = book["Monthly Targets"].iloc[:, :4].copy()
    mt.columns = ["Zone", "Branch", "Month", "Target"]
    mt["Branch"] = mt["Branch"].map(TS.canon_branch)
    mt["Month"] = pd.to_datetime(mt["Month"], errors="coerce")
    mt["Target"] = pd.to_numeric(mt["Target"], errors="coerce")
    monthly = mt[mt["Month"].dt.year == y].dropna(subset=["Target"]).groupby(["Branch", "Month"])["Target"].sum()
    hs = D["hs_sales"]
    sales = TS.prepare_sales(pd.DataFrame({"Year": hs["Date"].dt.year, "Month": hs["Date"].dt.month, "GTM": hs["GTM"],
                                           "Branch": hs["Branch"], "Region": hs["Region"], "HeliosCode": hs["Helios"],
                                           "Family": hs["Family"], "Partner": hs["Partner"], "Gross": hs["Gross"]})
                             .dropna(subset=["Year"]))
    py = TS.project_targets(hs_book, sales, y, issues)
    com = TS.commits(hs_book, y, issues)
    if len(com) and len(py):
        names = {b.upper(): b for b in py["Branch"].unique()}
        p = com["GTM"] == "Project"
        com.loc[p, "Branch"] = com.loc[p, "Branch"].map(lambda b: names.get(str(b).upper(), b))
    pm_ = TS.project_monthly(com, y, issues, py)
    pm = D["masters"]["partners"]
    info = {k: {"Branch": TS.canon_branch(v), "RD/KRO": t} for k, v, t in
            zip(pm.index, pm.get("Geo branch", pd.Series(index=pm.index, dtype=object)), pm.get("RD/KRO", pd.Series(index=pm.index, dtype=object)))} if len(pm) else {}
    pt, _ = TS.partner_targets(D["books"]["qtod"], book, sales, y, info, region_of, issues)
    return {"yearly": yearly, "monthly": monthly, "prod_m": prod_m, "region_of": region_of,
            "project_yearly": py, "commits": com, "project_monthly": pm_, "partners": pt,
            "ssp": TS.ssp_targets(book, y, issues), "issues": pd.DataFrame(issues)}


def to_basis(v, gtm, basis, year):
    """Net target → the report's basis."""
    return v if basis == "Net" else v / (1 - rebate(gtm, year))

# %%
# ════════════════════════════════════════════════════════════════════
# 8 · ANALYSIS BUILDING BLOCKS
# ════════════════════════════════════════════════════════════════════
M_ = f" ({UNIT_LABEL})"


def _sum(df, start=None, end=None, col="Value"):
    if df is None or df.empty:
        return 0.0
    if start is None:
        return df[col].sum()
    d = df["Date"]
    return df.loc[(d >= start) & (d <= end), col].sum()


def _pct(a, b):
    return a / b - 1 if b else np.nan


def _ach(a, t):
    return a / t if pd.notna(t) and t else np.nan


def scorecard(df, keys, P, ob=None, bl=None, month_target=None, year_target=None, qty=False):
    """MTD / YTD / last year by `keys`, growth, targets (already on the report basis), YTD OB, open backlog."""
    def by(start, end, col="Value", frame=df):
        if frame is None or frame.empty:
            return pd.Series(dtype=float)
        d = frame["Date"]
        return frame[(d >= start) & (d <= end)].groupby(keys)[col].sum()
    t = pd.DataFrame({
        "MTD Sales": by(P["m_start"], P["as_of"]), "MTD Sales LY": by(P["ly_m_start"], P["ly_as_of"]),
        "YTD Sales": by(P["y_start"], P["as_of"]), "YTD Sales LY": by(P["ly_y_start"], P["ly_as_of"]),
        f"FY {P['year'] - 1} Sales": by(P["ly_y_start"], P["ly_end"])}).fillna(0.0)
    t["MTD Growth %"] = t["MTD Sales"] / t["MTD Sales LY"].where(t["MTD Sales LY"] > 0) - 1
    t["YTD Growth %"] = t["YTD Sales"] / t["YTD Sales LY"].where(t["YTD Sales LY"] > 0) - 1
    if month_target is not None:
        t.insert(0, "Month Target", pd.to_numeric(month_target.reindex(t.index), errors="coerce"))
        t["Month Ach %"] = t["MTD Sales"] / t["Month Target"].where(t["Month Target"] > 0)
    if year_target is not None:
        t["Yearly Target"] = pd.to_numeric(year_target.reindex(t.index), errors="coerce")
        t["YTD Ach %"] = t["YTD Sales"] / t["Yearly Target"].where(t["Yearly Target"] > 0)
        t["Balance to target"] = (t["Yearly Target"] - t["YTD Sales"]).clip(lower=0)
        t["Needed per month"] = t["Balance to target"] / max(P["months_left"], 1)
    if qty:
        t["YTD Qty"] = by(P["y_start"], P["as_of"], "Qty").reindex(t.index).fillna(0)
        t["YTD Qty LY"] = by(P["ly_y_start"], P["ly_as_of"], "Qty").reindex(t.index).fillna(0)
        t["Qty Growth %"] = t["YTD Qty"] / t["YTD Qty LY"].where(t["YTD Qty LY"] > 0) - 1
    if ob is not None and len(ob):
        t["YTD OB"] = by(P["y_start"], P["as_of"], frame=ob).reindex(t.index).fillna(0.0)
    if bl is not None and len(bl):
        t["Open Backlog"] = bl.groupby(keys)["Value"].sum().reindex(t.index).fillna(0.0)
    return t


def with_total(t, label="Total"):
    """Adds a total row; ratios are recomputed from the totals."""
    if t.empty:
        return t
    s = t.select_dtypes("number").sum(min_count=1)
    for pct_col, a, b in (("MTD Growth %", "MTD Sales", "MTD Sales LY"), ("YTD Growth %", "YTD Sales", "YTD Sales LY"),
                          ("Qty Growth %", "YTD Qty", "YTD Qty LY")):
        if pct_col in t:
            s[pct_col] = s[a] / s[b] - 1 if s.get(b) else np.nan
    for pct_col, a, b in (("Month Ach %", "MTD Sales", "Month Target"), ("YTD Ach %", "YTD Sales", "Yearly Target")):
        if pct_col in t:
            s[pct_col] = s[a] / s[b] if pd.notna(s.get(b)) and s.get(b) else np.nan
    if "Needed per month" in t and "Balance to target" in t:
        s["Needed per month"] = s["Balance to target"] / max(PERIOD["months_left"], 1)
    row = pd.DataFrame([s], index=pd.Index([label] if t.index.nlevels == 1 else [tuple([label] + [""] * (t.index.nlevels - 1))],
                                           name=t.index.name))
    if t.index.nlevels > 1:
        row.index = pd.MultiIndex.from_tuples(row.index, names=t.index.names)
    return pd.concat([t, row])


def month_trend(df, P, targets_by_month=None, year_target=None):
    """Each month of the year: sales this year, last year, the year before, growth; targets and YTD progress."""
    y = P["year"]
    rows, cum = [], 0.0
    for m in range(1, 13):
        a = pd.Timestamp(y, m, 1)
        b = a + pd.offsets.MonthEnd(0)
        cur = _sum(df, a, min(b, P["as_of"])) if a <= P["as_of"] else np.nan
        ly = _sum(df, a.replace(year=y - 1), b.replace(year=y - 1))
        ly2 = _sum(df, a.replace(year=y - 2), (a.replace(year=y - 2) + pd.offsets.MonthEnd(0)))
        cum = cum + (cur if pd.notna(cur) else 0)
        r = {"Month": a.strftime("%b"), f"Sales {y}": cur, f"Sales {y - 1}": ly, f"Sales {y - 2}": ly2,
             "Growth % vs LY": _pct(cur, ly) if pd.notna(cur) else np.nan}
        if targets_by_month is not None:
            t = targets_by_month.get(a, np.nan)
            r["Month Target"] = t
            r["Month Ach %"] = _ach(cur, t)
        if year_target is not None:
            r["YTD Sales"] = cum if pd.notna(cur) else np.nan
            r["YTD Ach % of yearly target"] = _ach(cum, year_target) if pd.notna(cur) else np.nan
            r["Year elapsed %"] = m / 12
        rows.append(r)
    return pd.DataFrame(rows).set_index("Month")


def crosstab(df, rows, cols, P, col="Value"):
    """rows × cols: YTD this year, YTD last year and growth, side by side."""
    def piv(start, end):
        d = df[(df["Date"] >= start) & (df["Date"] <= end)]
        return d.assign(**{col: d[col].astype(float)}).pivot_table(index=rows, columns=cols, values=col, aggfunc="sum", fill_value=0.0)
    ty, ly = piv(P["y_start"], P["as_of"]), piv(P["ly_y_start"], P["ly_as_of"])
    keys = sorted(set(ty.columns) | set(ly.columns), key=str)
    ty, ly = ty.reindex(columns=keys, fill_value=0.0), ly.reindex(index=ty.index.union(ly.index), columns=keys, fill_value=0.0)
    ty = ty.reindex(ly.index, fill_value=0.0)
    out = {}
    for k in keys:
        out[(k, "This year")] = ty[k]
        out[(k, "Last year")] = ly[k]
        out[(k, "Growth %")] = ty[k] / ly[k].where(ly[k] > 0) - 1
    o = pd.DataFrame(out)
    o.columns = [f"{a} — {b}" + (M_ if b != "Growth %" and col == "Value" else "") for a, b in o.columns]
    return o


def partner_table(df, ob, bl, P, extra=None):
    """One row per partner: who, where, sales, growth, growth band, OB, backlog, last billing."""
    t = scorecard(df, ["Partner"], P, ob, bl)
    info = df.sort_values("Date").groupby("Partner").agg(**{"Partner name": ("Partner name", "last"), "Branch": ("Branch", "last"),
                                                          "Region": ("Region", "last"), "City": ("City", "last"),
                                                          "Partner type": ("Partner type", "last"), "Tier": ("Tier", "last"),
                                                          "GTM": ("GTM", "last"), "Last billing": ("Date", "max")})
    t = info.join(t, how="right")
    yl, yy = t["YTD Sales LY"], t["YTD Sales"]
    t["Growth band"] = np.select(
        [(yl <= 0) & (yy > 0), yy <= 0, t["YTD Growth %"] < -0.2, t["YTD Growth %"] < 0, t["YTD Growth %"] < 0.2,
         t["YTD Growth %"] < 0.5], ["New / no sales last year", "No sales this year", "Decline > 20%", "Decline 0–20%",
                                     "Growth 0–20%", "Growth 20–50%"], default="Growth > 50%")
    if extra is not None:
        t = t.join(extra, how="left")
    return t.sort_values("YTD Sales", ascending=False)


def policy_categories(df, P):
    """Partner policy categories as the Partner Analysis policy tracker (needs 3 years of Retail sales)."""
    y = P["year"]
    s = df[df["Date"] <= P["as_of"]]
    sm = s.groupby(["Partner", s["Date"].dt.year])["Gross"].sum().unstack(fill_value=0.0)
    g = lambda yy: sm[yy] if yy in sm.columns else pd.Series(0.0, index=sm.index)
    first_ly = s[s["Date"].dt.year == y - 1].groupby("Partner")["Date"].min().reindex(sm.index)
    new_ty = (g(y) > 0) & (g(y - 1) == 0) & (g(y - 2) == 0)
    new_ly = (g(y - 1) > 0) & (g(y - 2) == 0) & (g(y - 3) == 0)
    q4 = new_ly & first_ly.dt.month.isin([10, 11, 12])
    big = g(y - 1) >= 500000
    cat = pd.Series(np.select([new_ty, q4, new_ly & ~q4 & big, new_ly & ~q4 & ~big],
                              [f"New Partner {y}", f"Q4'{y - 1} Onboarded", f">5L {y - 1} Onboarded (Handholding)",
                               f"<5L {y - 1} Onboarded"], default="Old Partner"), index=sm.index)
    cat[cat.index.isin({str(x) for x in MANUAL_OLD_PARTNER_IDS})] = "Old Partner"
    order = [f"New Partner {y}", f"Q4'{y - 1} Onboarded", f"<5L {y - 1} Onboarded", f">5L {y - 1} Onboarded (Handholding)",
             "Old Partner"]
    bucket = cat.map(lambda c: "New partner" if c in order[:3] else "Handholding" if c == order[3] else "Old partner")
    return cat, bucket, order


def new_geo(df, P):
    """Cities with sales this year and none in the two years before."""
    y = P["year"]
    s = df[(df["Date"] <= P["as_of"]) & df["City"].notna() & (df["City"] != "Nan")]
    cm = s.groupby(["City", s["Date"].dt.year])["Gross"].sum().unstack(fill_value=0.0)
    g = lambda yy: cm[yy] if yy in cm.columns else pd.Series(0.0, index=cm.index)
    cities = cm.index[(g(y) > 0) & (g(y - 1) == 0) & (g(y - 2) == 0)]
    d = s[s["City"].isin(cities) & (s["Date"].dt.year == y)]
    return (d.groupby("City").agg(**{"Region": ("Region", "first"), "Branch": ("Branch", "first"),
                                     "Partners": ("Partner", "nunique"), "First sale": ("Date", "min"), "YTD Sales": ("Value", "sum")})
            .sort_values("YTD Sales", ascending=False))


def backlog_ageing(bl, keys, P):
    if bl is None or bl.empty:
        return pd.DataFrame()
    age = (P["as_of"] - bl["Order date"]).dt.days
    b = bl.assign(Band=pd.cut(age, [-1e9, 30, BACKLOG_OLD_DAYS, 90, 1e9],
                              labels=["0–30 days", f"31–{BACKLOG_OLD_DAYS} days", f"{BACKLOG_OLD_DAYS + 1}–90 days", "Over 90 days"]))
    t = b.pivot_table(index=keys, columns="Band", values="Value", aggfunc="sum", fill_value=0.0, observed=False)
    t.columns = [f"{c}{M_}" for c in t.columns]
    t[f"Open Backlog{M_}"] = t.sum(axis=1)
    t[f"Older than {BACKLOG_OLD_DAYS} days %"] = (b[age > BACKLOG_OLD_DAYS].groupby(keys)["Value"].sum()
                                                  .reindex(t.index).fillna(0) / t[f"Open Backlog{M_}"].where(t[f"Open Backlog{M_}"] > 0))
    return t.sort_values(f"Open Backlog{M_}", ascending=False)


def kpis(rows):
    """(measure, value, note) → a table with the value as text: Mn with 1 decimal, % or a count."""
    def show(measure, v):
        if v is None or (isinstance(v, float) and not np.isfinite(v)):
            return "–"
        if "%" in measure:
            return f"{v:.1%}"
        if measure.endswith(M_):
            return f"{v:,.1f}"
        return f"{v:,.0f}"
    return pd.DataFrame([(m, show(m, v), n) for m, v, n in rows], columns=["Measure", "Value", "Note"]).set_index("Measure")

# %%
# ════════════════════════════════════════════════════════════════════
# 9 · WORKBOOKS
# ════════════════════════════════════════════════════════════════════
MONEY_WORDS = ("Sales", "Target", "OB", "Backlog", "Commit", "Needed", "Balance", "Value")


def _for_excel(df):
    """Money columns ÷ UNIT with '(Mn)' in the name; the index as columns; ±inf → blank."""
    out = df.copy()
    ren = {}
    for c in out.columns:
        name = str(c)
        if not pd.api.types.is_numeric_dtype(out[c]) or pd.api.types.is_bool_dtype(out[c]):
            continue
        if name.endswith(M_):
            out[c] = out[c] / UNIT
        elif "%" not in name and not any(w in name for w in ("Qty", "Partners", "partners", "Lines")) \
                and any(w in name for w in MONEY_WORDS):
            out[c] = out[c] / UNIT
            ren[c] = f"{name}{M_}"
    out = out.rename(columns=ren).replace([np.inf, -np.inf], np.nan)
    if not isinstance(out.index, pd.RangeIndex):
        out = out.reset_index()
    return out


def write_book(filename, sheets, P):
    """sheets: list of (sheet name, title, DataFrame, note)."""
    def build(writer):
        for name, title, df, note in sheets:
            if df is None or len(df) == 0:
                continue
            out = _for_excel(df)
            kinds = {c: "int" for c in out.columns if any(w in str(c) for w in ("Qty", "Partners", "Lines", "Months", "Cities"))
                     and pd.api.types.is_numeric_dtype(out[c])}
            kinds.update({c: "mn" for c in out.columns if str(c).endswith(M_)})
            kinds.update({c: "pct" for c in out.columns if "%" in str(c)})
            write_table(writer, out, name[:31], title=title, subtitle=note, tab=THEME["tab_summary"],
                        freeze_cols=min(2, len(out.columns) - 1), kinds=kinds)
    return save_report(build, filename, local_dir=OUTPUT_DIR, sp_folder=SP_OUTPUT_FOLDER if UPLOAD_TO_SHAREPOINT else None)


def about_sheet(scope, basis, P, T):
    rows = [("Report", scope), ("Values", f"{basis} sales, INR {UNIT_LABEL}"),
            ("Period", f"YTD 1 Jan – {P['label']}; MTD {P['m_start']:%d %b} – {P['label']}; last year = the same days a year earlier"),
            ("Rebates (Net)", "; ".join(f"{g}: " + ", ".join(f"{y} {r:.1%}" for y, r in sorted(v.items(), reverse=True))
                                         for g, v in REBATES.items())),
            ("Targets", "Branch yearly and monthly targets, product (Zeta / Wiser) and SSP targets are net in the target files; "
                        + ("the Gross file shows them ÷ (1 − rebate)." if basis == "Gross" else "shown as they are.")
                        + " QTOD partner targets are gross" + (" (netted with the Retail rebate here)." if basis == "Net" else ".")),
            ("Created", datetime.now().strftime("%d %b %Y %H:%M"))]
    t = pd.DataFrame(rows, columns=["Item", "Detail"]).set_index("Item")
    return t


def retail_book(D, T, P, basis):
    y, mon = P["year"], P["mon"]
    S, O, B = (with_values(D[k], basis, P) for k in ("retail_sales", "retail_ob", "retail_bl"))
    conv = lambda s: s.map(lambda v: to_basis(v, "Retail", basis, y)) if len(s) else s
    yt = conv(T["yearly"])
    mt_all = T["monthly"].map(lambda v: to_basis(v, "Retail", basis, y))
    m_key = P["m_start"]
    mt = mt_all.xs(m_key, level="Month") if m_key in mt_all.index.get_level_values(1) else pd.Series(dtype=float)
    region_of = T["region_of"]
    rg_t = lambda s: s.groupby(s.index.map(lambda b: region_of.get(b, "Unmapped"))).sum()
    month_tot = mt_all.groupby(level="Month").sum().to_dict()

    # --- partners and policy
    cat, bucket, order = policy_categories(S, P)
    pt = T["partners"]
    pt = pt[pt["Family"] == "All"] if len(pt) else pt
    p_rate = 1.0 if basis == "Gross" else (1 - rebate("Retail", y))
    q_now = (P["as_of"].month - 1) // 3 + 1
    if len(pt):
        tq = (pt["Target 100%"] * p_rate).groupby([pt["Partner Code"], pt["Quarter"]]).sum().unstack(fill_value=0.0)
        q_start = pd.Timestamp(y, 3 * (q_now - 1) + 1, 1)
        q_sales = S[(S["Date"] >= q_start) & (S["Date"] <= P["as_of"])].groupby("Partner")["Value"].sum()
        extra = pd.DataFrame({"Annual Target (QTOD file)": tq.sum(axis=1), f"Q{q_now} Target": tq.get(f"Q{q_now}"),
                              f"Q{q_now} Sales": q_sales})
    else:
        extra = pd.DataFrame()
    extra = extra.reindex(extra.index.union(cat.index)) if len(extra) else pd.DataFrame(index=cat.index)
    extra["Policy category"] = cat
    extra["Policy group"] = bucket
    flat = extra["Policy group"].eq("New partner") & extra.get("Annual Target (QTOD file)", pd.Series(dtype=float)).reindex(extra.index).isna()
    extra["Assigned Target (policy)"] = extra.get("Annual Target (QTOD file)").where(~flat, FLAT_NEW_PARTNER_TARGET * p_rate) \
        if "Annual Target (QTOD file)" in extra else np.where(extra["Policy group"].eq("New partner"), FLAT_NEW_PARTNER_TARGET * p_rate, np.nan)
    parts = partner_table(S, O, B, P, extra)
    parts["ATOD Ach %"] = parts["YTD Sales"] / parts["Assigned Target (policy)"].where(parts["Assigned Target (policy)"] > 0)
    if f"Q{q_now} Target" in parts:
        parts[f"Q{q_now} QTOD Ach %"] = parts[f"Q{q_now} Sales"] / parts[f"Q{q_now} Target"].where(parts[f"Q{q_now} Target"] > 0)
    S2 = S.assign(Policy=S["Partner"].map(cat).fillna("Old Partner"))
    pol = scorecard(S2, ["Policy"], P, O.assign(Policy=O["Partner"].map(cat).fillna("Old Partner")),
                    B.assign(Policy=B["Partner"].map(cat).fillna("Old Partner"))).reindex(order)
    billed = S2[(S2["Date"] >= P["y_start"]) & (S2["Gross"] > 0)].groupby("Policy")["Partner"].nunique()
    pol.insert(0, "Partners billed YTD", billed.reindex(order).fillna(0))
    pol = with_total(pol)

    # --- QTOD by partner and quarter
    qt = pd.DataFrame()
    if len(pt):
        rows = []
        for q in range(1, 5):
            a = pd.Timestamp(y, 3 * (q - 1) + 1, 1)
            b = min(a + pd.offsets.QuarterEnd(0), P["as_of"])
            qs = S[(S["Date"] >= a) & (S["Date"] <= b)].groupby("Partner")["Value"].sum() if a <= P["as_of"] else pd.Series(dtype=float)
            qtgt = (pt[pt["Quarter"] == f"Q{q}"].set_index("Partner Code")["Target 100%"] * p_rate)
            rows.append(pd.DataFrame({f"Q{q} Target": qtgt, f"Q{q} Sales": qs}))
        qt = pd.concat(rows, axis=1)
        qt = qt.reindex(pt["Partner Code"].unique())
        for q in range(1, 5):
            qt[f"Q{q} Ach %"] = qt[f"Q{q} Sales"] / qt[f"Q{q} Target"].where(qt[f"Q{q} Target"] > 0)
        fam = T["partners"]
        for f_, grp in (("Miluz Zeta", "Miluz Zeta"), ("Wiser", "Wiser")):
            ft = fam[fam["Family"] == f_].groupby("Partner Code")["Target 100%"].sum() * p_rate
            fs = S[(S["Product Group"] == grp) & (S["Date"] >= P["y_start"]) & (S["Date"] <= P["as_of"])].groupby("Partner")["Value"].sum()
            qt[f"{f_} Annual Target"], qt[f"{f_} YTD Sales"] = ft, fs.reindex(qt.index)
            qt[f"{f_} Ach %"] = qt[f"{f_} YTD Sales"] / qt[f"{f_} Annual Target"].where(qt[f"{f_} Annual Target"] > 0)
        info = pt.drop_duplicates("Partner Code").set_index("Partner Code")[["Partner", "Branch", "Region", "Partner type"]]
        qt = info.join(qt)
        qt.index.name = "Partner Code"

    # --- SSP
    ssp_t = T["ssp"]
    ssp = pd.DataFrame()
    if len(ssp_t):
        st = ssp_t[ssp_t["Year"] == y].copy()
        st["Target"] = st["Target"].map(lambda v: to_basis(v, "Retail", basis, y))
        bs = scorecard(S, ["Branch"], P)
        ssp = st.groupby(["SSP", "Branch"])["Target"].sum().to_frame("Yearly Target")
        ssp["YTD Sales"] = [bs["YTD Sales"].get(b, 0.0) for _, b in ssp.index]
        ssp["YTD Sales LY"] = [bs["YTD Sales LY"].get(b, 0.0) for _, b in ssp.index]
        ssp["YTD Growth %"] = ssp["YTD Sales"] / ssp["YTD Sales LY"].where(ssp["YTD Sales LY"] > 0) - 1
        ssp["YTD Ach %"] = ssp["YTD Sales"] / ssp["Yearly Target"].where(ssp["Yearly Target"] > 0)
        ssp = ssp.reset_index().set_index(["SSP", "Branch"])
    ssa = D["masters"]["ssa"]
    with_ssp = pd.DataFrame()
    if ssa:
        S3 = S.assign(**{"SSP coverage": S["Branch"].astype(str).str.upper().map(ssa).fillna("No SSP")})
        S3["SSP coverage"] = np.where(S3["SSP coverage"].astype(str).str.lower().isin(["no ssa", "no ssp", "nan", ""]), "Without SSP", "With SSP")
        with_ssp = with_total(scorecard(S3, ["SSP coverage"], P))

    # --- MTD daily
    d = S[(S["Date"] >= P["m_start"]) & (S["Date"] <= P["as_of"])]
    daily = d.groupby(d["Date"].dt.normalize()).agg(**{"Sales": ("Value", "sum"), "Partners billed": ("Partner", "nunique"),
                                               "Lines": ("Value", "size")})
    daily.index.name = "Date"
    daily["Cumulative Sales"] = daily["Sales"].cumsum()

    # --- orders
    ob_trend = month_trend(O, P)
    ob_trend.columns = [c.replace("Sales", "OB") for c in ob_trend.columns]

    yearly_total = yt.sum() if len(yt) else np.nan
    pg = with_total(scorecard(S, ["Product Group"], P, O, B, qty=True))
    pm = T["prod_m"]
    if len(pm):
        pmt = pm.assign(Target=pm["Target"].map(lambda v: to_basis(v, "Retail", basis, y)))
        pmt["Month"] = pd.to_datetime(pmt["Month"])
        ytd_t = pmt[pmt["Month"] <= P["as_of"]].groupby("Product group")["Target"].sum()
        pg["YTD Product Target (Zeta / Wiser)"] = ytd_t.reindex(pg.index)
        pg["YTD Product Ach %"] = pg["YTD Sales"] / pg["YTD Product Target (Zeta / Wiser)"].where(pg["YTD Product Target (Zeta / Wiser)"] > 0)
    tot = scorecard(S.assign(All="Retail"), ["All"], P, O.assign(All="Retail"), B.assign(All="Retail"))
    r0 = tot.iloc[0] if len(tot) else pd.Series(dtype=float)
    mtd_t = mt.sum() if len(mt) else np.nan
    active = lambda a, b: S[(S["Date"] >= a) & (S["Date"] <= b) & (S["Gross"] > 0)]["Partner"].nunique()
    geo = new_geo(S, P)
    overview = kpis([
        (f"YTD Sales{M_}", r0.get("YTD Sales", 0) / UNIT, f"1 Jan – {P['label']}"),
        (f"YTD Sales last year{M_}", r0.get("YTD Sales LY", 0) / UNIT, "same days last year"),
        ("YTD Growth %", r0.get("YTD Growth %"), ""),
        (f"Yearly Target{M_}", yearly_total / UNIT, "Branch Targets"),
        ("YTD Ach % of yearly target", _ach(r0.get("YTD Sales", 0), yearly_total), f"{(12 - P['months_left']) / 12:.0%} of the year gone"),
        (f"Needed per month{M_}", max(yearly_total - r0.get("YTD Sales", 0), 0) / max(P["months_left"], 1) / UNIT, f"{P['months_left']} month(s) left"),
        (f"{mon} Sales (MTD){M_}", r0.get("MTD Sales", 0) / UNIT, f"{P['m_start']:%d %b} – {P['label']}"),
        (f"{mon} Target{M_}", mtd_t / UNIT, "Monthly Targets"),
        (f"{mon} Ach %", _ach(r0.get("MTD Sales", 0), mtd_t), ""),
        (f"{mon} Growth %", r0.get("MTD Growth %"), "vs the same days last year"),
        (f"YTD OB{M_}", r0.get("YTD OB", 0) / UNIT, ""),
        (f"Open Backlog{M_}", r0.get("Open Backlog", 0) / UNIT, ""),
        ("Active partners YTD", active(P["y_start"], P["as_of"]), f"last year {active(P['ly_y_start'], P['ly_as_of'])}"),
        (f"New partners (policy) — {order[0]}", int((cat == order[0]).sum()), f"new-partner group: {int(bucket.eq('New partner').sum())}"),
        ("Handholding partners", int(bucket.eq("Handholding").sum()), order[3]),
        ("New geo cities", len(geo), "sales this year, none in the two years before"),
    ])
    overview.columns = ["Value", "Note"]

    rg = with_total(scorecard(S, ["Region"], P, O, B, rg_t(mt) if len(mt) else None, rg_t(yt) if len(yt) else None))
    br = scorecard(S, ["Branch"], P, O, B, mt, yt)
    br.insert(0, "Region", br.index.map(lambda b: region_of.get(b, "Unmapped")))
    br.insert(1, "Active partners YTD", S[(S["Date"] >= P["y_start"]) & (S["Gross"] > 0)].groupby("Branch")["Partner"].nunique().reindex(br.index))
    br = with_total(br.sort_values(["Region", "YTD Sales"], ascending=[True, False]))
    bm = []
    for (b, m), tv in mt_all.items():
        if m > P["as_of"]:
            continue
        e = min(m + pd.offsets.MonthEnd(0), P["as_of"])
        sv = _sum(S[S["Branch"] == b], m, e)
        lv = _sum(S[S["Branch"] == b], m.replace(year=y - 1), (m + pd.offsets.MonthEnd(0)).replace(year=y - 1))
        bm.append({"Branch": b, "Region": region_of.get(b, "Unmapped"), "Month": m.strftime("%b %Y"), "Month Target": tv,
                   "Sales": sv, "Ach %": _ach(sv, tv), "Sales LY": lv, "Growth %": _pct(sv, lv)})
    bm = pd.DataFrame(bm).set_index(["Branch", "Region", "Month"]) if bm else pd.DataFrame()

    sheets = [
        ("About", f"Retail analysis — {basis}", about_sheet("Retail", basis, P, T), None),
        ("Overview", f"Retail — overview ({basis}, INR {UNIT_LABEL})", overview, f"As of {P['label']}"),
        ("Monthly trend", "Retail — sales by month vs last years, monthly target and yearly target progress",
         month_trend(S, P, month_tot, yearly_total), "Month Target = Monthly Targets (commit); YTD Ach % = YTD ÷ yearly target"),
        ("Regions", "Retail — regions", rg, "Region of each branch as in Branch Targets"),
        ("Branches", "Retail — branches", br, None),
        ("Branch x month", "Retail — monthly target vs sales by branch", bm, "Monthly Targets as given; months without a target are not shown"),
        ("Products", "Retail — product groups (value and quantity)", pg, "Product targets only where given (Zeta / Wiser monthly)"),
        ("Families", "Retail — families", with_total(scorecard(S, ["Family"], P, O, B, qty=True).sort_values("YTD Sales", ascending=False)), None),
        ("Branch x family", "Retail — YTD by branch and product group", crosstab(S, "Branch", "Product Group", P), None),
        ("Branch x family qty", "Retail — YTD quantity by branch and product group", crosstab(S, "Branch", "Product Group", P, col="Qty"), None),
        ("Activities", "Retail — activities", with_total(scorecard(S, ["Activity"], P, O, B)), None),
        ("RD vs KRO", "Retail — partner type", with_total(scorecard(S, ["Partner type"], P, O, B)), "RD = pricing class ERW, KRO = EKR"),
        ("City tier", "Retail — city tier", with_total(scorecard(S, ["Tier"], P, O, B)), "Tier from the Partner Master"),
        ("Partners", "Retail — all partners", parts, "Policy category and assigned target as the Partner Analysis policy tracker"),
        ("Partner targets", "Retail — QTOD / ATOD by partner and quarter", qt, "Targets from the QTOD and ATOD file"),
        ("New partners", "Retail — partners by policy category", pol, "New Partner = sales this year, none in the two years before"),
        ("New partner list", "Retail — new and onboarded partners",
         parts[parts["Policy group"].isin(["New partner", "Handholding"])] if "Policy group" in parts else pd.DataFrame(), None),
        ("New geo", "Retail — new cities", geo, "Cities with sales this year and none in the two years before"),
        ("SSP", "Retail — SSP target vs achievement", ssp, "Achievement = sales of the SSP's branch"),
        ("With vs without SSP", "Retail — sales in branches with / without an SSP", with_ssp, "SSP Mapping of the Data Dictionary"),
        ("MTD", f"Retail — {mon} by branch", scorecard(S, ["Branch"], P, None, None, mt).sort_values("MTD Sales", ascending=False), None),
        ("MTD daily", f"Retail — {mon} day by day", daily, None),
        ("Order book", "Retail — order booking by month", ob_trend, None),
        ("Backlog", "Retail — open backlog by branch and age", backlog_ageing(B, ["Branch"], P), "Age = days since order creation"),
        ("Backlog by partner", "Retail — open backlog by partner (top 100)",
         backlog_ageing(B, ["Partner", "Partner name"], P).head(100), None),
        ("Target issues", "Problems found in the target files", T["issues"].set_index("Source sheet") if len(T["issues"]) else None, None),
    ]
    return sheets


def hs_book(D, T, P, basis, with_fd):
    y, mon = P["year"], P["mon"]
    keep = (lambda d: d) if with_fd else (lambda d: d[d["Activity"] != "FD"])
    S, O, B = (keep(with_values(D[k], basis, P)) for k in ("hs_sales", "hs_ob", "hs_bl"))
    tag = "with FD" if with_fd else "without FD"
    region_of = T["region_of"]
    # targets: Retail yearly / monthly (with FD only), Project yearly by activity, commits by activity
    r_yt = T["yearly"].map(lambda v: to_basis(v, "Retail", basis, y))
    r_mt_all = T["monthly"].map(lambda v: to_basis(v, "Retail", basis, y))
    py = T["project_yearly"].copy()
    if len(py):
        py["Target"] = py["Target"].map(lambda v: to_basis(v, "Project", basis, y))
    com = T["commits"].copy()
    if len(com):
        com["Commit"] = [to_basis(v, g, basis, y) for v, g in zip(com["Commit"], com["GTM"])]
        com = com[(com["GTM"] == "Retail") | (com["Activity"] != "FD")]
        if not with_fd:
            com = com[com["Activity"] != "FD"]
    m_key = P["m_start"]
    r_mt = r_mt_all.xs(m_key, level="Month") if m_key in r_mt_all.index.get_level_values(1) else pd.Series(dtype=float)
    p_yt = py.groupby("Branch")["Target"].sum() if len(py) else pd.Series(dtype=float)
    p_mt = (com.loc[(com["GTM"] == "Project") & (pd.to_datetime(com["Month"]) == m_key), "Commit"].sum()
            if len(com) else np.nan) or np.nan                     # Project month target = this month's commits

    # Retail / Project / overall
    def seg(sel, yt=np.nan, mt=np.nan):
        s, o, b = S[sel(S)], O[sel(O)], B[sel(B)]
        t = scorecard(s.assign(All=1), ["All"], P, o.assign(All=1), b.assign(All=1)).iloc[0] if len(s) else pd.Series(dtype=float)
        t = t.copy()
        t["Month Target"], t["Yearly Target"] = mt, yt
        t["Month Ach %"], t["YTD Ach %"] = _ach(t.get("MTD Sales", 0), mt), _ach(t.get("YTD Sales", 0), yt)
        return t
    ret, proj, fd = (lambda d: d["GTM"] == "Retail"), (lambda d: d["GTM"] == "Project"), (lambda d: d["Activity"] == "FD")
    r_y, r_m = (r_yt.sum(), r_mt.sum()) if with_fd else (np.nan, np.nan)
    rows = {}
    if with_fd:
        rows["Retail – without FD"] = seg(lambda d: ret(d) & ~fd(d))
        rows["Retail – FD only"] = seg(lambda d: ret(d) & fd(d))
    rows["Retail – total" if with_fd else "Retail (without FD)"] = seg(ret, r_y, r_m)
    rows["HS Project"] = seg(proj, p_yt.sum() if len(p_yt) else np.nan, p_mt)
    rows[f"HS overall (Retail + Project, {tag})"] = seg(lambda d: pd.Series(True, index=d.index),
                                                        (r_y + p_yt.sum()) if with_fd and len(p_yt) else np.nan,
                                                        r_m + p_mt if with_fd else np.nan)
    gt = pd.DataFrame(rows).T
    gt.index.name = "Segment"
    gt = gt[[c for c in ["Month Target", "MTD Sales", "Month Ach %", "MTD Growth %", "Yearly Target", "YTD Sales", "YTD Ach %",
                         "YTD Growth %", "YTD Sales LY", f"FY {y - 1} Sales", "YTD OB", "Open Backlog"] if c in gt]]

    # branches: Retail branches (Retail targets) and Project branches (Project targets)
    br = scorecard(S, ["GTM", "Branch"], P, O, B)
    yt_b = pd.Series({("Retail", b): v for b, v in r_yt.items()} if with_fd else {}, dtype=float)
    yt_b = pd.concat([yt_b, pd.Series({("Project", b): v for b, v in p_yt.items()}, dtype=float)])
    br["Yearly Target"] = yt_b.reindex(br.index)
    br["YTD Ach %"] = br["YTD Sales"] / br["Yearly Target"].where(br["Yearly Target"] > 0)
    br["Needed per month"] = (br["Yearly Target"] - br["YTD Sales"]).clip(lower=0) / max(P["months_left"], 1)
    br = with_total(br.sort_values("YTD Sales", ascending=False))
    rg = with_total(scorecard(S, ["Region"], P, O, B))
    act = scorecard(S, ["Activity", "GTM"], P, O, B)
    if len(py):
        act["Yearly Target"] = py.groupby(["Activity", "GTM"])["Target"].sum().reindex(act.index)
        act["YTD Ach %"] = act["YTD Sales"] / act["Yearly Target"].where(act["Yearly Target"] > 0)
    act = with_total(act)
    # commits vs sales by month and activity
    cv = pd.DataFrame()
    if len(com):
        c = com.assign(Month=pd.to_datetime(com["Month"]))
        cs = c.groupby(["Month", "GTM", "Activity"])["Commit"].sum()
        rows = []
        for (m, g, a), v in cs.items():
            e = min(m + pd.offsets.MonthEnd(0), P["as_of"])
            sv = _sum(S[(S["GTM"] == g) & (S["Activity"] == a)], m, e) if m <= P["as_of"] else np.nan
            rows.append({"Month": m.strftime("%b %Y"), "GTM": g, "Activity": a, "Commit": v, "Sales": sv,
                         "Commit Ach %": _ach(sv, v) if pd.notna(v) else np.nan})
        cv = pd.DataFrame(rows).set_index(["Month", "GTM", "Activity"])
    month_tot = r_mt_all.groupby(level="Month").sum().to_dict() if with_fd else None
    trend = month_trend(S, P, month_tot, (r_y + p_yt.sum()) if with_fd and len(p_yt) else None)
    if with_fd:
        trend = trend.rename(columns={"Month Target": "Month Target (Retail only)", "Month Ach %": "Month Ach % (Retail only)"})
    tr_g = pd.concat({g: month_trend(S[S["GTM"] == g], P)[[f"Sales {y}", f"Sales {y - 1}", "Growth % vs LY"]]
                      for g in ("Retail", "Project")}, axis=1)
    tr_g.columns = [f"{g} {c}" for g, c in tr_g.columns]
    parts = partner_table(S, O, B, P)
    wis = lambda d: d[d["Product Group"] == "Wiser"]
    wiser = with_total(scorecard(wis(S), ["GTM"], P, wis(O), wis(B), qty=True), "Wiser total")
    wiser_br = with_total(scorecard(wis(S), ["GTM", "Branch"], P, wis(O), wis(B), qty=True).sort_values("YTD Sales", ascending=False))
    d = S[(S["Date"] >= P["m_start"]) & (S["Date"] <= P["as_of"])]
    daily = d.pivot_table(index=d["Date"].dt.normalize(), columns="GTM", values="Value", aggfunc="sum", fill_value=0.0)
    daily.columns = [f"{c} Sales" for c in daily.columns]
    daily["Total Sales"] = daily.sum(axis=1)
    daily.index.name = "Date"
    ob_trend = month_trend(O, P)
    ob_trend.columns = [c.replace("Sales", "OB") for c in ob_trend.columns]
    tot = gt.iloc[-1]
    overview = kpis([
        (f"YTD Sales{M_}", tot.get("YTD Sales", 0) / UNIT, f"Retail + Project, {tag}, 1 Jan – {P['label']}"),
        ("YTD Growth %", tot.get("YTD Growth %"), "vs the same days last year"),
        (f"Yearly Target{M_}", tot.get("Yearly Target", np.nan) / UNIT,
         "Retail Branch Targets + Project targets" if with_fd else "no target without FD (Retail targets include FD)"),
        ("YTD Ach %", tot.get("YTD Ach %"), ""),
        (f"Retail YTD Sales{M_}", gt.iloc[-3 if with_fd else 0].get("YTD Sales", 0) / UNIT if len(gt) > 2 else np.nan, ""),
        (f"Project YTD Sales{M_}", gt.loc["HS Project", "YTD Sales"] / UNIT if "HS Project" in gt.index else np.nan, ""),
        ("Project YTD Ach %", gt.loc["HS Project", "YTD Ach %"] if "HS Project" in gt.index else np.nan, "Project yearly targets (HS file)"),
        (f"{mon} Sales (MTD){M_}", tot.get("MTD Sales", 0) / UNIT, ""),
        (f"{mon} Growth %", tot.get("MTD Growth %"), ""),
        (f"YTD OB{M_}", tot.get("YTD OB", 0) / UNIT, ""),
        (f"Open Backlog{M_}", tot.get("Open Backlog", 0) / UNIT, ""),
        ("Retail share of YTD Sales %", _ach(S.loc[(S["GTM"] == "Retail") & (S["Date"] >= P["y_start"]), "Value"].sum(),
                                             S.loc[S["Date"] >= P["y_start"], "Value"].sum()), ""),
        ("Active partners YTD", S[(S["Date"] >= P["y_start"]) & (S["Gross"] > 0)]["Partner"].nunique(), ""),
    ])
    overview.columns = ["Value", "Note"]
    scope = f"HS (Retail + Project) {tag}"
    return [
        ("About", f"{scope} — {basis}", about_sheet(scope, basis, P, T), None),
        ("Overview", f"{scope} — overview ({basis}, INR {UNIT_LABEL})", overview, f"As of {P['label']}"),
        ("Retail vs Project", f"{scope} — Retail and Project", gt,
         "Month Target: Retail = Monthly Targets, Project = this month's commits (HS targets file)"
         + ("" if with_fd else "; Retail targets include FD, so they are not shown without FD")),
        ("Monthly trend", f"{scope} — sales by month vs last years", trend,
         "Month Target = Retail Monthly Targets only (Project has no monthly target)" if with_fd else None),
        ("Monthly trend by GTM", f"{scope} — Retail and Project by month", tr_g, None),
        ("Regions", f"{scope} — regions", rg, "Retail rows: region of the branch; Project rows: sales-office region"),
        ("Branches", f"{scope} — branches (Retail branch / Project sales office)", br, None),
        ("Activities", f"{scope} — activities", act, "Yearly target = Project targets by activity (HS file)"),
        ("Commit vs sales", f"{scope} — monthly commits vs sales by activity", cv,
         "Commits as in the HS targets file (Project FD commits left out); months without a commit are not shown"),
        ("Products", f"{scope} — product groups", with_total(scorecard(S, ["Product Group"], P, O, B, qty=True).sort_values("YTD Sales", ascending=False)), None),
        ("Families", f"{scope} — families", with_total(scorecard(S, ["Family"], P, O, B, qty=True).sort_values("YTD Sales", ascending=False)), None),
        ("Branch x family", f"{scope} — YTD by branch and product group", crosstab(S, "Branch", "Product Group", P), None),
        ("GTM x family", f"{scope} — YTD by Retail / Project and product group", crosstab(S, "Product Group", "GTM", P), None),
        ("Partners", f"{scope} — all partners", parts, None),
        ("Wiser", f"{scope} — Wiser", wiser, None),
        ("Wiser by branch", f"{scope} — Wiser by branch", wiser_br, None),
        ("MTD", f"{scope} — {mon} by branch", scorecard(S, ["GTM", "Branch"], P).sort_values("MTD Sales", ascending=False), None),
        ("MTD daily", f"{scope} — {mon} day by day", daily, None),
        ("Order book", f"{scope} — order booking by month", ob_trend, None),
        ("Backlog", f"{scope} — open backlog by branch and age", backlog_ageing(B, ["GTM", "Branch"], P), "Age = days since order creation"),
        ("Target issues", "Problems found in the target files", T["issues"].set_index("Source sheet") if len(T["issues"]) else None, None),
    ]

# %%
# ════════════════════════════════════════════════════════════════════
# 10 · RUN — reads the data once and writes the six workbooks
# ════════════════════════════════════════════════════════════════════
t0 = time.time()
DATA = load_data()
PERIOD = period(DATA)
TARGETS = targets(DATA, PERIOD)
stamp = f"{PERIOD['mon']}_{PERIOD['year']}"
written = []
for basis in ("Gross", "Net"):
    written.append(write_book(f"Retail_Analysis_{basis}_{stamp}.xlsx", retail_book(DATA, TARGETS, PERIOD, basis), PERIOD))
    for with_fd in (True, False):
        tag = "with_FD" if with_fd else "without_FD"
        written.append(write_book(f"HS_{tag}_Analysis_{basis}_{stamp}.xlsx", hs_book(DATA, TARGETS, PERIOD, basis, with_fd), PERIOD))
print(f"\n✅ {len(written)} workbooks as of {PERIOD['label']} in {time.time() - t0:.0f}s:")
for w in written:
    print("   ", w)
