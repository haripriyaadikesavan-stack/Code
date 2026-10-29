"""Write the Channel Analytics notebook chunk (section 10) from organize_targets.py, so the notebook runs exactly the
tested code. The module's functions are wrapped in one function (no global names clash with the notebook).

usage: python notebook_cell.py > cell.py      (then paste, or let a build script insert it)
"""
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
#       Retail, HS and Wiser dashboards, split by last year's sales:
#       Branch Targets (month × branch × GTM × activity) · Family Targets (product
#       groups; Zeta / Wiser as given) · Partner Targets (quarters, All / Zeta /
#       Wiser) · Partner Product Split · Commits · SSP Targets · Region Targets
#       · Issues (problems found in the source files) · Check (totals agree)
# Retail targets are the same rows for the Retail and the HS dashboard.
# Source: tableau/targets/organize_targets.py (this cell is generated from it).
# ════════════════════════════════════════════════════════════════════
import types

TARGET_SPLIT_YEAR = TARGET_YEAR          # year of the targets; last year's sales are the basis
HS_TARGETS_PATH = f"{MASTER}/Master Targets File (Retail+Project).xlsx"   # HS dashboard's targets ("" = none)
TARGET_SPLIT_UPLOAD = False              # True = also upload both files to Final Reports on SharePoint


def _target_split_tools():
    """organize_targets.py, wrapped so its names stay out of the notebook's globals."""
'''

FOOTER = '''

TS = _target_split_tools()


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

    tables, issues = TS.organize(master_book, qtod, sales, year, hs, partner_info)
    print(f"   🎯 {len(issues)} issue(s) in the target files — see the 'Issues' sheet")
    targets_file = f"Channel_Targets_{year}.xlsx"
    path = os.path.join(OUTPUT_DIR, targets_file)
    TS.write_targets(path, tables, issues, year, year - 1)
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


def cell_source():
    src = (HERE / "organize_targets.py").read_text()
    body = src[src.index("\nSALES_COLUMNS = ") + 1:src.index("\ndef main():")].rstrip()
    body = re.sub(r"(?<![\w.])dt\.", "_dt.", body)       # the datetime module, not pandas .dt
    lines = ["import datetime as _dt", "import re", "import numpy as np", "import pandas as pd", ""] + body.split("\n")
    inner = "\n".join(("    " + ln) if ln.strip() else "" for ln in lines)
    return HEADER + inner + "\n\n    return types.SimpleNamespace(**{k: v for k, v in locals().items() if not k.startswith('__')})\n" + FOOTER


if __name__ == "__main__":
    print(cell_source())
