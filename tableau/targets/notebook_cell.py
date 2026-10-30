"""Write the Channel Analytics notebook cells generated from organize_targets.py, so the notebook runs exactly the
tested code:
  keys   cell 2.5  TARGET KEYS — the module (wrapped in one function, so no global names clash with the notebook)
                   + add_target_keys(), used by the Retail and HS dashboard input files
  split  cell 10   TARGET PLANNING — sales contribution + Channel_Targets_<year>.xlsx

usage: python notebook_cell.py keys|split > cell.py
"""
import sys
import pathlib
import re

HERE = pathlib.Path(__file__).parent

HEADER = '''# ════════════════════════════════════════════════════════════════════
# 10 · TARGET PLANNING — last year's sales contribution + one organised targets file
#
# Reads the Retail targets (Master Targets File + QTOD and ATOD file) and the HS
# targets file, and writes to OUTPUT_DIR:
#   Sales_Contribution_<last year>.xlsx  how last year's sales split by branch,
#       month, activity, product group, Retail vs Project and partner
#   Channel_Targets_<year>.xlsx  every target in one long-format workbook for the
#       Retail, HS and Wiser dashboards, as given:
#       Yearly Targets (year × branch × GTM) · Monthly Targets (the monthly commit,
#       month × branch × GTM; separate from the yearly target, months not set stay
#       empty) · Product Targets Monthly / Yearly (only products with a target;
#       region product targets shared out to branches by last year's sales)
#       · Partner Targets (quarters, All / Zeta / Wiser) · SSP Targets · Region
#       Targets · Issues (problems found in the source files) · Check
# Retail targets are the same rows for the Retail and the HS dashboard.
# Uses TS (cell 2.5). Source: tableau/targets/organize_targets.py.
# ════════════════════════════════════════════════════════════════════
TARGET_SPLIT_YEAR = TARGET_YEAR          # year of the targets; last year's sales are the basis
HS_TARGETS_PATH = f"{MASTER}/Master Targets File (Retail+Project).xlsx"   # HS dashboard's targets ("" = none)
TARGET_SPLIT_UPLOAD = False              # True = also upload both files to Final Reports on SharePoint


def target_split_sales():
    """Retail + Project sales as the HS dashboard sees them (branch, region, Retail / Project) + family."""
    base = build_hs_dataset(df_Sales_final, "InvoiceDate", **_dash_args)
    d = pd.to_datetime(base["InvoiceDate"], errors="coerce")
    fam = family_map.drop_duplicates("Material").set_index("Material")["Family Actual"]
    return pd.DataFrame({
        "Year": d.dt.year, "Month": d.dt.month, "GTM": base["Classification"], "Branch": base["Branch"],
        "Region": base["Region"], "HeliosCode": base["HeliosCode"],
        "Family": base["MaterialCode"].astype(str).str.strip().str.upper().map(fam),
        "Partner": base["SoldToPartyID"], "Gross": base["Gross Sales"]}).dropna(subset=["Year"])


def _read_book(path):
    return pd.read_excel(io.BytesIO(sp.download(path).getvalue()), sheet_name=None)


def report_target_split():
    year = TARGET_SPLIT_YEAR
    sales = target_split_sales()
    master_book = _read_book(PATHS["targets_master"])
    qtod = next(iter(_read_book(PATHS["targets_qtod"]).values()))
    hs = None
    if HS_TARGETS_PATH:
        try:
            hs = _read_book(HS_TARGETS_PATH)
        except Exception as e:
            print(f"   ⚠️ HS targets file not read ({e}) — Project targets left out")
    meta = Partner_Metadata.copy()
    meta["SoldToPartyID"] = meta["SoldToPartyID"].astype(str).str.strip()
    partner_info = {k: {"Branch": b, "RD/KRO": t} for k, b, t in zip(meta["SoldToPartyID"], meta["Branch"], meta["RD/KRO"])}

    tables, issues, given = TS.organize(master_book, qtod, sales, year, hs, partner_info)
    print(f"   🎯 {len(issues)} issue(s) in the target files — see the 'Issues' sheet")
    targets_file = f"Channel_Targets_{year}.xlsx"
    path = os.path.join(OUTPUT_DIR, targets_file)
    TS.write_targets(path, tables, issues, given, year, year - 1)
    print(f"   💾 {path}")
    if TARGET_SPLIT_UPLOAD:
        with open(path, "rb") as fh:
            sp.upload(f"{PATHS['final_reports']}/{targets_file}", fh.read())

    contribution = TS.contribution(TS.prepare_sales(sales), year - 1)

    def build(writer):
        for name, df in contribution.items():
            write_table(writer, df, name[:31], tab=THEME["tab_data"], freeze_cols=2)
    saved = save_report(build, f"Sales_Contribution_{year - 1}.xlsx", local_dir=OUTPUT_DIR,
                        sp_folder=PATHS["final_reports"] if TARGET_SPLIT_UPLOAD else None)
    return f"{targets_file} + {saved}"


run_report("target_split", report_target_split)
'''


KEYS_HEADER = '''# ════════════════════════════════════════════════════════════════════
# 2.5 TARGET KEYS — columns that link sales to Channel_Targets_<year>.xlsx
#
# add_target_keys(df) adds, to the Retail and HS dashboard input files:
#   Target Branch   Retail: the partner's branch (Partner Master city → Geography
#                   mapping, as the Retail dashboard); Project: the sales office
#                   branch (NS Branches)
#   Target GTM      Retail / Project
#   Target Month    1st of the invoice month (Sales) or calendar-day month (OB)
#   Target Year · Target Quarter (Q1–Q4)
#   Target Activity WD / FD / COS / SPE / EVC (from the Helios code)
#   Product Group   Unica Pure, Livia, AvatarOn, Miluz Zeta, WD- All others,
#                   Easy9 6kA / 10kA / RCCB / DB, FD- All others, Wiser,
#                   COS- All others, SPE, EVC
#   Partner key     partner code as text
# The same rules split the targets (section 10), so a sale and its target always
# meet. In Tableau, relate each targets sheet on these columns (no calculations).
# Source: tableau/targets/organize_targets.py (this cell is generated from it).
# ════════════════════════════════════════════════════════════════════
import types


def _target_split_tools():
    """organize_targets.py, wrapped so its names stay out of the notebook's globals."""
'''

KEYS_FOOTER = '''

TS = _target_split_tools()


def _retail_target_branch():
    """Partner code → branch: Partner Master city → Geography mapping (as the Retail dashboard and HS datasets)."""
    p = master("partner_retail")
    p["City"] = p["City"].astype(str).str.strip()
    g = master("geo_retail")[["City", "Branch"]].copy()
    g["City"] = g["City"].astype(str).str.strip()
    p = p.merge(g.drop_duplicates("City"), on="City", how="left")
    code = p["Sold-to Party"].astype(str).str.strip().str.upper().str.replace(r"\\.0$", "", regex=True)
    return dict(zip(code, p["Branch"]))


RETAIL_TARGET_BRANCH = _retail_target_branch()
TARGET_KEY_KINDS = {"Target Year": "plain"}      # write_table would format a column named "Target …" as rupees
_FAMILY = family_map.drop_duplicates("Material").set_index("Material")["Family Actual"]


def add_target_keys(df, gtm=None):
    """df + the target key columns. gtm="Retail" for Retail-only data (branch from the partner);
    otherwise Retail / Project come from 'Classification' and the branch from 'Branch' (HS datasets)."""
    if df is None or df.empty:
        return df
    out = df.copy()
    partner = out["SoldToPartyID"].astype(str).str.strip().str.upper().str.replace(r"\\.0$", "", regex=True)
    out["Target GTM"] = gtm if gtm else out["Classification"]
    branch = partner.map(RETAIL_TARGET_BRANCH) if gtm == "Retail" else out["Branch"]
    retail = out["Target GTM"] == "Retail"
    names = {b: TS.canon_branch(b) for b in branch[retail].dropna().unique()}
    out["Target Branch"] = branch.where(~retail, branch.map(names))
    date_col = next((c for c in ("InvoiceDate", "Calendar Day") if c in out.columns), None)
    d = pd.to_datetime(out[date_col], errors="coerce") if date_col else pd.Series(pd.NaT, index=out.index)
    out["Target Month"] = d.dt.to_period("M").dt.to_timestamp()
    out["Target Year"] = d.dt.year.astype("Int64")
    out["Target Quarter"] = ("Q" + d.dt.quarter.astype("Int64").astype(str)).where(d.notna())
    out["Target Activity"] = out["HeliosCode"].astype(str).str.strip().str.upper().map(TS.ACTIVITY)
    family = out["MaterialCode"].astype(str).str.strip().str.upper().map(_FAMILY)
    pairs = pd.DataFrame({"a": out["Target Activity"], "f": family}).drop_duplicates()
    group = {(a, f): TS.product_group(a, f) for a, f in zip(pairs["a"], pairs["f"])}
    out["Product Group"] = [group[k] for k in zip(out["Target Activity"], family)]
    out["Partner key"] = partner
    missing = out["Target Branch"].isna() & out["Target GTM"].eq("Retail")
    if missing.any():
        gross = pd.to_numeric(out.loc[missing, "Gross Sales"], errors="coerce").sum() if "Gross Sales" in out else 0
        print(f"   ⚠️ {int(missing.sum()):,} Retail rows ({gross / 1e5:,.1f} L) have no Target Branch — "
              f"partner not in the Partner Master or city not in the Geography mapping")
    return out


print(f"✅ Target keys ready | {len(RETAIL_TARGET_BRANCH):,} Retail partners mapped to a branch")
'''


def keys_source():
    src = (HERE / "organize_targets.py").read_text()
    body = src[src.index("\nSALES_COLUMNS = ") + 1:src.index("\ndef main():")].rstrip()
    body = re.sub(r"(?<![\w.])dt\.", "_dt.", body)       # the datetime module, not pandas .dt
    lines = ["import datetime as _dt", "import re", "import numpy as np", "import pandas as pd", ""] + body.split("\n")
    inner = "\n".join(("    " + ln) if ln.strip() else "" for ln in lines)
    return (KEYS_HEADER + inner + "\n\n    return types.SimpleNamespace(**{k: v for k, v in locals().items() if not k.startswith('__')})\n"
            + KEYS_FOOTER)


def split_source():
    return HEADER


if __name__ == "__main__":
    print({"keys": keys_source, "split": split_source}[sys.argv[1]]())
