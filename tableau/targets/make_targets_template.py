"""Sample targets workbook for the HS and Wiser dashboards (one file for both).

usage: python make_targets_template.py "HS_Wiser_Targets_2026.xlsx" [--year 2026] [--empty]

Design (fixes the lags of the current target sheets):
  * Long tables: one row per target (year, month, branch, GTM, activity / family). No pivoting in Tableau,
    a new month or branch is just more rows, and every sheet joins the sales data on the same keys.
  * Every value row says its basis (Net / Gross). The dashboards turn each row into both
    (gross = net / 0.83 for targets, net / 0.87 for commits), so nobody has to remember which sheet is which.
  * GTM is a column (Retail / Project), never a duplicated row: Wiser / Zeta targets are entered once.
  * Yearly targets are the sum of the months, so yearly and monthly targets can never disagree.
  * One branch list ('Branches') with the name each data source uses (partner map for Retail,
    NS Branches / sales office for Project), so 'MUM' and 'Mum' meet on one target branch.
  * Quarter targets are numbers (the Retail QTOD sheet stores Q2-Q4 as text).
  * Dropdowns on every key column and a 'Check' sheet (totals, unknown branches, duplicate keys).
--empty writes the same workbook without the sample rows.
"""
import argparse
import datetime as dt
import random

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

GREEN, LIGHT = "1F7A4D", "E8F3EC"
MAX_ROWS = 20000          # rows covered by dropdowns
CHECK_ROWS = 6000         # rows covered by the Check sheet (duplicate checks are slow on more)

# target branch, region, branch name in Retail data (partner map), in Project data (NS Branches), Retail share
BRANCHES = [
    ("AP/TG", "South", "AP", "AP/TG", 0.1), ("TEL", "South", "TEL", "", 1), ("KAR", "South", "KAR", "Kar", 0.6),
    ("KER", "South", "KER", "", 1), ("TN-1", "South", "TN-1", "", 1), ("TN-2", "South", "TN-2", "", 1),
    ("N TN", "South", "", "N TN", 0), ("S TN", "South", "", "S TN", 0),
    ("CHD", "North", "CHD", "", 1), ("Del/NCR", "North", "", "Del/NCR", 0), ("NCR", "North", "NCR", "", 1),
    ("NOIDA-GHZ", "North", "NOIDA-GHZ", "", 1), ("HR", "North", "HR", "", 1), ("J&K", "North", "J&K", "", 1),
    ("UP-1", "North", "UP-1", "", 1), ("UP-2", "North", "UP-2", "", 1), ("U North", "North", "", "U North", 0),
    ("UP/UK", "North", "", "UP/UK", 0),
    ("WB", "East", "WB", "", 1), ("GAU", "East", "GAU", "", 1), ("JKD", "East", "JKD", "", 1), ("OBJ", "East", "", "OBJ", 0),
    ("MUM", "West1", "MUM", "Mum", 0.4), ("PUNE", "West1", "PUNE", "Pune", 0.15), ("PCMC", "West1", "PCMC", "", 1),
    ("NASHIK", "West1", "NASHIK", "", 1), ("INDORE", "West1", "INDORE", "", 1),
    ("GUJ", "West2", "GUJ", "Ahm", 0.8), ("Vad", "West2", "", "Vad", 0), ("RAJ", "West2", "RAJ", "Raj", 0.9),
]
GTM = ["Retail", "Project"]
ACTIVITIES = [("WD", "Wiring Devices", "HDCWD, HSCWD"), ("FD", "Final Distribution", "HDCFD, PPCFD"),
              ("COS", "COS (incl. Wiser)", "HDCOS, HSCOS"), ("SPE", "SPE", "HDSPE, HSSPE"), ("EVC", "EV Charging", "PREVC")]
ACTIVITY_MIX = {"WD": 0.45, "FD": 0.25, "COS": 0.12, "SPE": 0.13, "EVC": 0.05}
FAMILIES = ["Wiser", "Miluz Zeta", "Livia", "Unica Pure", "AvatarOn", "Easy9"]
PARTNER_TYPES = ["RD", "KRO", "Project"]
BASIS = ["Net", "Gross"]
SEASON = [0.07, 0.07, 0.10, 0.07, 0.08, 0.08, 0.08, 0.08, 0.09, 0.09, 0.09, 0.10]   # Jan..Dec, sums to 1


def header(ws, cols, widths):
    for i, (c, w) in enumerate(zip(cols, widths), 1):
        cell = ws.cell(1, i, c)
        cell.font, cell.fill = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor=GREEN)
        cell.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 32
    ws.freeze_panes = "A2"


def table(ws, name, rows, cols, widths, formats=None):
    header(ws, cols, widths)
    for r in rows:
        ws.append(list(r))
    last = max(2, ws.max_row)
    t = Table(displayName=name, ref=f"A1:{get_column_letter(len(cols))}{last}")
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
    ws.add_table(t)
    for col, fmt in (formats or {}).items():
        j = cols.index(col) + 1
        for r in range(2, last + 200):
            ws.cell(r, j).number_format = fmt


def dropdown(ws, cols, col, source, prompt):
    j = get_column_letter(cols.index(col) + 1)
    dv = DataValidation(type="list", formula1=source, allow_blank=False, showErrorMessage=True,
                        errorTitle="Not in the list", error=prompt, promptTitle=col, prompt=prompt, showInputMessage=True)
    dv.add(f"{j}2:{j}{MAX_ROWS}")
    ws.add_data_validation(dv)


def list_range(ws_name, col, n):
    return f"'{ws_name}'!${col}$2:${col}${n + 1}"


def build(path, year, empty):
    rnd = random.Random(year)
    wb = Workbook()

    # ------------------------------------------------------------------ Read me
    ws = wb.active
    ws.title = "Read me"
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 110
    notes = [
        ("HS & Wiser Targets", f"Targets for the HS and Wiser dashboards, year {year}." +
         ("" if empty else "  SAMPLE VALUES — replace them with the real targets.")),
        ("", ""),
        ("How to fill", "One row per target. Add rows at the bottom of a table (it grows by itself); never add columns, "
                        "merge cells, subtotal rows or blank rows. Pick keys from the dropdowns."),
        ("Amounts", "Rupees (not lakh / crore), whole numbers. 'Basis' says whether the amount is Net or Gross; "
                    "the dashboards show both (targets: gross = net / 0.83; commits: gross = net / 0.87)."),
        ("Month", "First day of the month (e.g. 01-Apr-2026). Yearly targets = sum of the 12 months, so there is no "
                  "separate yearly sheet that could disagree with the months."),
        ("GTM", "Retail or Project, as the sales data classifies rows (Retail = pricing class L2 ED1/ED2/RE3 and "
                "L3 ERW/EKR/R36/ZD1). FD is always Retail, so FD has no Project targets or commits."),
        ("", ""),
        ("Branches", "The one branch list. 'Retail data branch' is the name in the partner map, 'Project data branch' "
                     "the name in NS Branches (sales office). Both point to one target branch, e.g. MUM and Mum. "
                     "Every branch used in the other sheets must be here."),
        ("Branch Targets", "Branch × month × GTM × activity (WD, FD, COS, SPE, EVC). Used for monthly, MTD, YTD and "
                           "yearly achievement in the HS dashboard; the COS rows include Wiser."),
        ("Commits", "Branch × month × GTM × activity commitments, same keys as Branch Targets."),
        ("Family Targets", "Focus families (Wiser, Miluz Zeta, ...) by branch × month × GTM, value and optional "
                           "quantity. Entered once — not repeated on a Retail and a Project row. Used by the Wiser "
                           "dashboard and the Wiser / Zeta views of the HS and Retail dashboards."),
        ("Partner Targets", "Partner × quarter, optional per family ('All' = the partner's total). Target 100% and the "
                            "stretch Target 108% are numbers. Annual = sum of the quarters."),
        ("Wiser Service", "Wiser services: installations / demos per branch and month and the expected turnaround days."),
        ("Lists", "Allowed values for the dropdowns. Add a new family or activity here first."),
        ("Check", "Totals per branch, rows whose branch is not in 'Branches', and duplicate keys. "
                  "All checks should show 0 before the file is published."),
        ("", ""),
        ("Keys", "Branch Targets / Commits: Year + Month + Branch + GTM + Activity.  Family Targets: Year + Month + "
                 "Branch + GTM + Family.  Partner Targets: Year + Quarter + Partner Code + Family.  Each key once."),
    ]
    for r, (a, b) in enumerate(notes, 1):
        ws.cell(r, 1, a).font = Font(bold=True, size=14 if r == 1 else 11, color=GREEN if r == 1 else "000000")
        c = ws.cell(r, 2, b)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if r == 1 and not empty:
            c.font = Font(bold=True, color="C00000")

    # ------------------------------------------------------------------ Lists
    wl = wb.create_sheet("Lists")
    lists = {"A": ("GTM", GTM), "B": ("Activity", [a for a, _, _ in ACTIVITIES]),
             "C": ("Activity name", [n for _, n, _ in ACTIVITIES]), "D": ("Helios codes", [h for _, _, h in ACTIVITIES]),
             "E": ("Family", FAMILIES), "F": ("Family (partner)", ["All"] + FAMILIES), "G": ("Basis", BASIS),
             "H": ("Partner type", PARTNER_TYPES), "I": ("Quarter", ["Q1", "Q2", "Q3", "Q4"]),
             "J": ("Region", ["North", "South", "East", "West1", "West2"]), "K": ("Family GTM", GTM + ["All"])}
    header(wl, [v[0] for v in lists.values()], [10, 10, 20, 16, 14, 16, 8, 13, 9, 10, 11])
    for col, (_, values) in lists.items():
        for i, v in enumerate(values, 2):
            wl[f"{col}{i}"] = v
    L = {k: list_range("Lists", col, len(v[1])) for k, (col, v) in zip(
        ["gtm", "act", "_n", "_h", "fam", "pfam", "basis", "ptype", "qtr", "region", "fgtm"], lists.items())}

    # ------------------------------------------------------------------ Branches
    wbr = wb.create_sheet("Branches")
    bcols = ["Branch", "Region", "Retail data branch", "Project data branch", "Active"]
    table(wbr, "Branches", [(b, r, rd, pd_, "Y") for b, r, rd, pd_, _ in BRANCHES], bcols, [14, 10, 20, 20, 8])
    dropdown(wbr, bcols, "Region", L["region"], "North, South, East, West1 or West2")
    branch_src = f"'Branches'!$A$2:$A${len(BRANCHES) + 50}"

    # ------------------------------------------------------------------ Branch Targets / Commits
    months = [dt.date(year, m, 1) for m in range(1, 13)]
    region = {b: r for b, r, *_ in BRANCHES}
    size = {b: rnd.uniform(4, 30) * 1e7 for b, *_ in BRANCHES}         # yearly net target, rupees
    tcols = ["Year", "Month", "Region", "Branch", "GTM", "Activity", "Target", "Basis"]
    ccols = ["Year", "Month", "Region", "Branch", "GTM", "Activity", "Commit", "Basis"]
    trows, crows = [], []
    if not empty:
        for b, _, _, _, share in BRANCHES:
            for g in GTM:
                gshare = share if g == "Retail" else 1 - share
                if gshare <= 0:
                    continue
                for a, mix in ACTIVITY_MIX.items():
                    if a == "FD" and g == "Project":
                        continue                                   # FD is always Retail
                    for m, s in zip(months, SEASON):
                        v = round(size[b] * gshare * mix * s * rnd.uniform(0.9, 1.1), -3)
                        trows.append((year, m, region[b], b, g, a, v, "Net"))
                        crows.append((year, m, region[b], b, g, a, round(v * rnd.uniform(0.85, 1.05), -3), "Net"))
    wt = wb.create_sheet("Branch Targets")
    table(wt, "BranchTargets", trows, tcols, [7, 11, 9, 13, 9, 9, 15, 7], {"Month": "mmm-yy", "Target": "#,##0"})
    wc = wb.create_sheet("Commits")
    table(wc, "Commits", crows, ccols, [7, 11, 9, 13, 9, 9, 15, 7], {"Month": "mmm-yy", "Commit": "#,##0"})
    for w, cols in ((wt, tcols), (wc, ccols)):
        dropdown(w, cols, "Branch", branch_src, "A branch from the 'Branches' sheet")
        dropdown(w, cols, "Region", L["region"], "North, South, East, West1 or West2")
        dropdown(w, cols, "GTM", L["gtm"], "Retail or Project")
        dropdown(w, cols, "Activity", L["act"], "WD, FD, COS, SPE or EVC")
        dropdown(w, cols, "Basis", L["basis"], "Net or Gross")

    # ------------------------------------------------------------------ Family Targets
    fcols = ["Year", "Month", "Region", "Branch", "GTM", "Family", "Target", "Target quantity", "Basis"]
    frows = []
    if not empty:
        for b, _, _, _, share in BRANCHES:
            for fam, part, gtms in (("Wiser", 0.04, GTM), ("Miluz Zeta", 0.05, ["Retail"])):
                for g in gtms:
                    gshare = share if g == "Retail" else 1 - share
                    if gshare <= 0:
                        continue
                    for m, s in zip(months, SEASON):
                        v = round(size[b] * gshare * part * s * rnd.uniform(0.85, 1.15), -3)
                        frows.append((year, m, region[b], b, g, fam, v, int(v / (3200 if fam == "Wiser" else 450)), "Net"))
    wf = wb.create_sheet("Family Targets")
    table(wf, "FamilyTargets", frows, fcols, [7, 11, 9, 13, 9, 13, 15, 11, 7],
          {"Month": "mmm-yy", "Target": "#,##0", "Target quantity": "#,##0"})
    dropdown(wf, fcols, "Branch", branch_src, "A branch from the 'Branches' sheet")
    dropdown(wf, fcols, "Region", L["region"], "North, South, East, West1 or West2")
    dropdown(wf, fcols, "GTM", L["fgtm"], "Retail, Project, or All when the target is not split")
    dropdown(wf, fcols, "Family", L["fam"], "A family from the 'Lists' sheet")
    dropdown(wf, fcols, "Basis", L["basis"], "Net or Gross")

    # ------------------------------------------------------------------ Partner Targets
    pcols = ["Year", "Quarter", "Partner Code", "Partner", "Branch", "Region", "GTM", "Partner type", "Family",
             "Target 100%", "Target 108%", "Basis"]
    prows = []
    if not empty:
        n = 0
        for b, _, _, _, share in BRANCHES:
            for g in GTM:
                if (share if g == "Retail" else 1 - share) <= 0:
                    continue
                for _ in range(2):
                    n += 1
                    code = str(500100000 + n)
                    ptype = rnd.choice(["RD", "KRO"]) if g == "Retail" else "Project"
                    annual = rnd.uniform(0.3, 3) * 1e7
                    for fam, part in (("All", 1), ("Wiser", 0.03)):
                        for q, qs in zip(["Q1", "Q2", "Q3", "Q4"], [0.22, 0.24, 0.26, 0.28]):
                            v = round(annual * part * qs, -3)
                            prows.append((year, q, code, f"Sample Partner {n}", b, region[b], g, ptype, fam,
                                          v, round(v * 1.08, -3), "Gross"))
    wp = wb.create_sheet("Partner Targets")
    table(wp, "PartnerTargets", prows, pcols, [7, 8, 13, 22, 13, 9, 9, 10, 12, 15, 15, 7],
          {"Target 100%": "#,##0", "Target 108%": "#,##0", "Partner Code": "@"})
    dropdown(wp, pcols, "Quarter", L["qtr"], "Q1, Q2, Q3 or Q4")
    dropdown(wp, pcols, "Branch", branch_src, "A branch from the 'Branches' sheet")
    dropdown(wp, pcols, "Region", L["region"], "North, South, East, West1 or West2")
    dropdown(wp, pcols, "GTM", L["gtm"], "Retail or Project")
    dropdown(wp, pcols, "Partner type", L["ptype"], "RD, KRO or Project")
    dropdown(wp, pcols, "Family", L["pfam"], "All (partner total) or a family")
    dropdown(wp, pcols, "Basis", L["basis"], "Net or Gross")

    # ------------------------------------------------------------------ Wiser Service
    scols = ["Year", "Month", "Region", "Branch", "Installations target", "Demos target", "Expected TAT days"]
    srows = []
    if not empty:
        for b, *_ in BRANCHES:
            for m in months:
                srows.append((year, m, region[b], b, rnd.randint(5, 40), rnd.randint(2, 15), 3))
    wsv = wb.create_sheet("Wiser Service")
    table(wsv, "WiserService", srows, scols, [7, 11, 9, 13, 13, 12, 12], {"Month": "mmm-yy"})
    dropdown(wsv, scols, "Branch", branch_src, "A branch from the 'Branches' sheet")
    dropdown(wsv, scols, "Region", L["region"], "North, South, East, West1 or West2")

    # ------------------------------------------------------------------ Check
    wk = wb.create_sheet("Check", 1)
    def rng(sheet, col):
        return f"'{sheet}'!${col}$2:${col}${CHECK_ROWS}"

    kcols = ["Branch", "Region", "Retail target", "Project target", "Total target", "Total commit", "Commit vs target",
             "Wiser target", "Zeta target", "Partner targets (100%, All)"]
    header(wk, kcols, [13, 9, 15, 15, 15, 15, 11, 14, 14, 18])
    for i in range(2, len(BRANCHES) + 2):
        wk[f"A{i}"] = f"=Branches!A{i}"
        wk[f"B{i}"] = f"=Branches!B{i}"
        wk[f"C{i}"] = f"=SUMIFS({rng('Branch Targets', 'G')},{rng('Branch Targets', 'D')},A{i},{rng('Branch Targets', 'E')},\"Retail\")"
        wk[f"D{i}"] = f"=SUMIFS({rng('Branch Targets', 'G')},{rng('Branch Targets', 'D')},A{i},{rng('Branch Targets', 'E')},\"Project\")"
        wk[f"E{i}"] = f"=C{i}+D{i}"
        wk[f"F{i}"] = f"=SUMIFS({rng('Commits', 'G')},{rng('Commits', 'D')},A{i})"
        wk[f"G{i}"] = f"=IF(E{i}=0,\"\",F{i}/E{i})"
        wk[f"H{i}"] = f"=SUMIFS({rng('Family Targets', 'G')},{rng('Family Targets', 'D')},A{i},{rng('Family Targets', 'F')},\"Wiser\")"
        wk[f"I{i}"] = f"=SUMIFS({rng('Family Targets', 'G')},{rng('Family Targets', 'D')},A{i},{rng('Family Targets', 'F')},\"Miluz Zeta\")"
        wk[f"J{i}"] = f"=SUMIFS({rng('Partner Targets', 'J')},{rng('Partner Targets', 'E')},A{i},{rng('Partner Targets', 'I')},\"All\")"
        for c in "CDEFHIJ":
            wk[f"{c}{i}"].number_format = "#,##0"
        wk[f"G{i}"].number_format = "0%"
    tot = len(BRANCHES) + 2
    wk[f"A{tot}"] = "Total"
    wk[f"A{tot}"].font = Font(bold=True)
    for c in "CDEFHIJ":
        wk[f"{c}{tot}"] = f"=SUM({c}2:{c}{tot - 1})"
        wk[f"{c}{tot}"].number_format = "#,##0"
        wk[f"{c}{tot}"].font = Font(bold=True)
    wk[f"G{tot}"] = f"=IF(E{tot}=0,\"\",F{tot}/E{tot})"
    wk[f"G{tot}"].number_format = "0%"

    # problems: unknown branches, duplicate keys, wrong basis, FD Project rows
    n = CHECK_ROWS
    br = f"Branches!$A$2:$A${len(BRANCHES) + 50}"

    def unknown(sheet, col):
        r = f"'{sheet}'!${col}$2:${col}${n}"
        return f"=SUMPRODUCT(({r}<>\"\")*(COUNTIF({br},{r})=0))"

    def dups(sheet, cols):
        rs = [f"'{sheet}'!${c}$2:${c}${n}" for c in cols]
        first = rs[0]
        args = ",".join(f"{r},{r}" for r in rs)
        return f"=SUMPRODUCT(({first}<>\"\")*(COUNTIFS({args})>1))"

    problems = [
        ("Branch Targets: branch not in 'Branches'", unknown("Branch Targets", "D")),
        ("Commits: branch not in 'Branches'", unknown("Commits", "D")),
        ("Family Targets: branch not in 'Branches'", unknown("Family Targets", "D")),
        ("Partner Targets: branch not in 'Branches'", unknown("Partner Targets", "E")),
        ("Wiser Service: branch not in 'Branches'", unknown("Wiser Service", "D")),
        ("Branch Targets: duplicate Year + Month + Branch + GTM + Activity", dups("Branch Targets", "BDEF")),
        ("Commits: duplicate Year + Month + Branch + GTM + Activity", dups("Commits", "BDEF")),
        ("Family Targets: duplicate Year + Month + Branch + GTM + Family", dups("Family Targets", "BDEF")),
        ("Partner Targets: duplicate Year + Quarter + Partner + Family", dups("Partner Targets", "BCI")),
        ("FD rows marked Project (FD is always Retail)",
         f"=COUNTIFS('Branch Targets'!$E$2:$E${n},\"Project\",'Branch Targets'!$F$2:$F${n},\"FD\")"
         f"+COUNTIFS('Commits'!$E$2:$E${n},\"Project\",'Commits'!$F$2:$F${n},\"FD\")"),
        ("Rows without a Basis",
         f"=COUNTIFS('Branch Targets'!$A$2:$A${n},\"<>\",'Branch Targets'!$H$2:$H${n},\"\")"
         f"+COUNTIFS('Family Targets'!$A$2:$A${n},\"<>\",'Family Targets'!$I$2:$I${n},\"\")"
         f"+COUNTIFS('Partner Targets'!$A$2:$A${n},\"<>\",'Partner Targets'!$L$2:$L${n},\"\")"),
    ]
    top = tot + 2
    wk.cell(top, 1, "Checks (all should be 0)").font = Font(bold=True, color=GREEN, size=12)
    for k, (label, formula) in enumerate(problems, top + 1):
        wk.cell(k, 1, label)
        wk.cell(k, 6, formula).number_format = "#,##0"
    rng_checks = f"F{top + 1}:F{top + len(problems)}"
    wk.conditional_formatting.add(rng_checks, CellIsRule(operator="greaterThan", formula=["0"],
                                                         fill=PatternFill("solid", fgColor="F8CBAD")))
    wk.conditional_formatting.add(rng_checks, CellIsRule(operator="equal", formula=["0"],
                                                         fill=PatternFill("solid", fgColor=LIGHT)))
    wk.freeze_panes = "B2"

    for w in wb.worksheets:
        w.sheet_properties.tabColor = GREEN if w.title in ("Read me", "Check") else "8FBC9F"
    wb.save(path)
    return {"Branch Targets": len(trows), "Commits": len(crows), "Family Targets": len(frows),
            "Partner Targets": len(prows), "Wiser Service": len(srows)}


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("out")
    ap.add_argument("--year", type=int, default=dt.date.today().year)
    ap.add_argument("--empty", action="store_true", help="no sample rows")
    a = ap.parse_args()
    print(a.out, build(a.out, a.year, a.empty))
