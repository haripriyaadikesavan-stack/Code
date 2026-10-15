"""Sheets and dashboards of the Retail dashboard."""
from core import C_DARK, C_MID, C_MUTED, C_PY, C_TARGET, C_TEXT
from sheets import MN, MV, Sheet, kpi, nav, run, title_runs

PAGES = ["Home", "Executive", "Branches", "Products", "Partners", "Partner Targets", "Order Book", "Explore"]
PARTNER_PAGES = {"Partners", "Partner Targets", "Explore"}     # row-level security applied here


# ---------------------------------------------------------------- generic builders
def has_value(R, measures):
    """Show '(Mn)' etc. in the title only when the sheet shows money values."""
    return any(R[m].fmt == "n#,##0.0;-#,##0.0" for m in measures)


def table(R, name, title, dims, measures, sub=None, total=True, sort_by=None, top=None, nonnull="Has activity (agg)",
          widths=None, color_by=None):
    s = Sheet(R, name, title_runs(title, sub, param_units=has_value(R, measures)), "table")
    drefs = []
    for d in dims:
        f = R[d] if d in R.by_caption else R.fields[d]
        drefs.append(s.use(f, "dim"))
    s.measures(measures)
    s.rows = " / ".join(drefs)
    s.rows_total = total
    s.cols = MN
    s.encodings.append(("text", MV))
    if color_by:
        s.encodings.append(("color", s.use(color_by, "usrn")))
    if nonnull:
        s.nonnull(nonnull)
    if top:
        n, by = top
        f = R[dims[0]] if dims[0] in R.by_caption else R.fields[dims[0]]
        s.top_n(drefs[0], f, n, by)
    if sort_by:
        s.sort_by(drefs[-1], s.use(sort_by, "usr"))
    s.widths.append((MN, 78))
    for i, w in enumerate(widths or []):
        s.widths.append((drefs[i], w))
    s.cell_align = "right"
    return s


def bars(R, name, title, dim, measures, sub=None, color_by=None, sort_by=None, top=None, horizontal=True,
         nonnull="Has activity (agg)", exclude_null=False, color=None):
    """Bar chart; several measures -> side by side bars coloured by Measure Names."""
    s = Sheet(R, name, title_runs(title, sub, param_units=has_value(R, measures)), "chart")
    f = R[dim] if dim in R.by_caption else R.fields[dim]
    dref = s.use(f, "dim")
    if len(measures) > 1:
        s.measures(measures)
        val, inner = MV, MN
        s.encodings.append(("color", MN))
    else:
        val, inner = s.use(measures[0], "usr"), None
        if color_by:
            s.encodings.append(("color", s.use(color_by, "usrn")))
        else:
            s.mark_color = color or C_MID
    s.encodings.append(("text", val))
    if horizontal:
        s.rows = f"({dref} / {inner})" if inner else dref
        s.cols = val
    else:
        s.cols = f"({dref} / {inner})" if inner else dref
        s.rows = val
    s.mark = "Bar"
    if nonnull:
        s.nonnull(nonnull)
    if exclude_null:
        s.exclude_null(dref, f)
    if top:
        n, by = top
        s.top_n(dref, f, n, by)
    if sort_by:
        s.sort_by(dref, s.use(sort_by, "usr"), shelf="rows" if horizontal else "cols")
    return s


def trend(R, name, title, measures, sub=None, mark="Line", dim="Month"):
    s = Sheet(R, name, title_runs(title, sub, param_units=has_value(R, measures)), "chart")
    dref = s.use(dim, "dim")
    s.measures(measures)
    s.encodings.append(("color", MN))
    s.cols = dref
    s.rows = MV
    s.mark = mark
    s.nonnull("Has activity (agg)")
    s.exclude_null(dref, R[dim])
    if mark == "Bar":
        s.cols = f"({dref} / {MN})"
        s.encodings.append(("text", MV))
    return s


def text_sheet(R, name, runs):
    """A sheet that only shows formatted text (used for highlights)."""
    s = Sheet(R, name, "", "kpi")
    s.label = runs
    s.use("Nav count", "usr")
    s.encodings.append(("text", s.use("Nav count", "usr")))
    return s


# ---------------------------------------------------------------- the sheets
def build_sheets(R):
    S = {}

    def add(s):
        S[s.name] = s
        return s

    for p in PAGES:
        add(nav(R, p))

    # ---- KPI cards (shared across pages)
    add(kpi(R, "K | Sales", "Sales", "Sales", [("Last year  ", "Sales PY", None), (None, "Growth text", None)]))
    add(kpi(R, "K | Achievement", "Achievement vs target", "Ach %",
            [("Target  ", "Target", None), ("Gap  ", "Gap to target", None)], units=False))
    add(kpi(R, "K | Annual target", "Annual target", "Annual target",
            [("Achieved  ", "Annual Ach %", None), ("Balance  ", "Balance to annual target", None)]))
    add(kpi(R, "K | Run-rate", "Monthly run-rate", "Current run-rate / month",
            [("Required to hit annual target  ", "Required run-rate / month", None)]))
    add(kpi(R, "K | Order book", "Order book (OB)", "OB", [("Last year  ", "OB PY", None), ("Growth  ", "OB growth %", None)]))
    add(kpi(R, "K | Backlog", "Open backlog", "Backlog", [("Older than 60 days  ", "Backlog over 60 days", None)]))
    add(kpi(R, "K | Partners", "Active partners", "Active partners",
            [("Last year  ", "Active partners PY", None), ("New / revival  ", "New / revival partners", None)], units=False))
    add(kpi(R, "K | Per partner", "Sales per active partner", "Sales per partner",
            [("New / revival partner sales  ", "New partner sales", None)]))
    add(kpi(R, "K | Data", "Data as of", "Data as of", [("Period  ", "Period text", None)], units=False))
    add(kpi(R, "K | Sales to OB", "Sales to OB", "Sales to OB %", [("Sales  ", "Sales", None), ("OB  ", "OB", None)], units=False))
    add(kpi(R, "K | Partner target", "Partner annual target (QTOD/ATOD)", "Partner annual target",
            [("Achieved  ", "Partner Ach %", None), ("Balance  ", "Partner balance", None)]))

    # ---- Executive
    add(trend(R, "E | Monthly trend", "Monthly sales - this year vs last year", ["Sales", "Sales PY"],
              "Green = selected period, grey = same period last year"))
    add(trend(R, "E | Sales vs target", "Monthly sales vs target", ["Sales", "Target"],
              "Target = monthly branch targets, pro-rated for the current month", mark="Bar"))
    add(table(R, "E | Region table", "Region performance", ["Region (clean)"],
              ["Sales", "Sales PY", "Growth %", "Target", "Ach %", "Annual target", "Annual Ach %", "OB", "Backlog", "Active partners"],
              widths=[90]))
    add(bars(R, "E | Branch achievement", "Branch achievement vs target", "Branch (clean)", ["Ach %"],
             "Sorted by achievement; colour = achievement band", color_by="Ach band", sort_by="Ach %", nonnull="Ach %"))
    add(bars(R, "E | Branch growth", "Branch growth vs last year", "Branch (clean)", ["Growth %"],
             "Sorted by growth; colour = growing / declining", color_by="Growth band", sort_by="Growth %", nonnull="Has sales (agg)"))

    # ---- Branches
    add(table(R, "B | Branch table", "Branch scorecard", ["Region (clean)", "Branch (clean)"],
              ["Sales", "Sales PY", "Growth %", "Target", "Ach %", "Gap to target", "Annual target", "Annual Ach %",
               "OB", "Backlog", "Active partners", "New / revival partners"], widths=[70, 90]))
    add(table(R, "B | Focus families", "Focus families by branch - Wiser and Miluz Zeta", ["Region (clean)", "Branch (clean)"],
              ["Wiser sales", "Wiser target", "Wiser Ach %", "Zeta sales", "Zeta target", "Zeta Ach %"], widths=[70, 90]))
    add(table(R, "B | ASM table", "ASM performance", ["Region (clean)", "ASM"],
              ["Sales", "Sales PY", "Growth %", "Share of sales %", "Active partners", "New / revival partners", "Sales per partner"],
              widths=[70, 150], nonnull="Has sales (agg)"))
    s = add(Sheet(R, "B | Branch by month", title_runs("Branch sales by month"), "table"))
    br = s.use("Branch (clean)", "dim")
    mo = s.use("Month", "dim")
    sv = s.use("Sales", "usr")
    s.rows, s.cols, s.rows_total = br, mo, True
    s.encodings.append(("text", sv))
    s.nonnull("Sales")
    s.cell_align = "right"
    s.exclude_null(mo, R["Month"])

    # ---- Products
    add(bars(R, "P | Family bars", "Sales by family - this year vs last year", "Family (clean)", ["Sales", "Sales PY"],
             sort_by="Sales", nonnull="Has sales (agg)"))
    add(table(R, "P | Family table", "Family performance", ["Family (clean)"],
              ["Sales", "Sales PY", "Growth %", "Share of sales %", "Quantity"], sort_by="Sales", widths=[150],
              nonnull="Has sales (agg)"))
    add(bars(R, "P | Activity", "Sales by activity (Helios code)", "Helios Code", ["Sales", "Sales PY"],
             sort_by="Sales", nonnull="Has sales (agg)", exclude_null=True))
    add(bars(R, "P | Classification", "Sales by product classification", "Classification (Sheet1)", ["Sales", "Sales PY"],
             sort_by="Sales", nonnull="Has sales (agg)", exclude_null=True))
    add(table(R, "P | Top SKUs", "Top 25 materials", ["MaterialDescription"],
              ["Sales", "Sales PY", "Growth %", "Share of sales %", "Quantity"], sort_by="Sales",
              top=(25, "Sales CY (row)"), widths=[260], nonnull="Has sales (agg)", total=False))
    add(table(R, "P | Focus by region", "Focus families by region", ["Region (clean)"],
              ["Wiser sales", "Wiser target", "Wiser Ach %", "Zeta sales", "Zeta target", "Zeta Ach %"], widths=[90]))

    # ---- Partners (row-level security)
    add(bars(R, "Pa | Slabs", "Partners by sales slab", "Partner slab", ["Partners (count)"],
             "Sales in the selected period (INR)", nonnull="Partners (count)", exclude_null=True, horizontal=True))
    S["Pa | Slabs"].title = title_runs("Partners by sales slab", "Partner sales in the selected period", param_units=False)
    add(bars(R, "Pa | Tier", "Sales by tier", "Tier Category", ["Sales", "Sales PY"], sort_by="Sales",
             nonnull="Has sales (agg)", exclude_null=True))
    add(bars(R, "Pa | Partner type", "Sales by partner type", "RD/KRO", ["Sales", "Sales PY"], sort_by="Sales",
             nonnull="Has sales (agg)", exclude_null=True))
    s = add(trend(R, "Pa | New partners", "New / revival partners by month", ["New / revival partners"],
                  "First sale this year (no sales in the previous two years)", mark="Bar"))
    s.title = title_runs("New / revival partners by month", "First sale this year (no sales in the previous two years)", param_units=False)
    add(table(R, "Pa | Top partners", "Top 25 partners", ["Partner (sales)", "Branch (clean)", "Partner status"],
              ["Sales", "Sales PY", "Growth %", "Share of sales %"], sort_by="Sales", top=(25, "Sales CY (row)"),
              widths=[240, 80, 90], total=False, nonnull="Has sales (agg)"))
    add(table(R, "Pa | At-risk partners", "Largest last-year partners now declining", ["Partner (sales)", "Branch (clean)"],
              ["Sales PY", "Sales", "Growth", "Growth %"], sort_by="Growth", top=(25, "Sales PY (row)"),
              widths=[240, 80], total=False, nonnull="Has sales (agg)", color_by="Growth band"))
    add(table(R, "Pa | SSA", "With vs without SSA", ["SSA group", "SSA"],
              ["Sales", "Sales PY", "Growth %", "Active partners", "Sales per partner"], widths=[90, 160]))

    # ---- Partner targets (QTOD / ATOD)
    add(table(R, "T | Annual", "Partner target vs achievement - annual", ["Target partner branch", "Target partner"],
              ["Partner annual target", "Sales", "Partner Ach %", "Partner balance", "Partner Wiser target", "Wiser sales",
               "Partner Wiser Ach %", "Partner Zeta target", "Zeta sales", "Partner Zeta Ach %"],
              nonnull="Has partner target", widths=[70, 240], color_by="Partner Ach band"))
    add(table(R, "T | Quarterly", "Partner target vs achievement - by quarter", ["Target partner branch", "Target partner"],
              ["Q1 target", "Q1 sales", "Q1 Ach %", "Q2 target", "Q2 sales", "Q2 Ach %",
               "Q3 target", "Q3 sales", "Q3 Ach %", "Q4 target", "Q4 sales", "Q4 Ach %"],
              nonnull="Has partner target", widths=[70, 240]))

    # ---- Order book & backlog
    add(trend(R, "O | OB trend", "Monthly order book - this year vs last year", ["OB", "OB PY"]))
    add(table(R, "O | OB by branch", "Order book and backlog by branch", ["Region (clean)", "Branch (clean)"],
              ["OB", "OB PY", "OB growth %", "Sales", "Sales to OB %", "Backlog", "Backlog over 60 days"], widths=[70, 90]))
    add(bars(R, "O | Backlog ageing", "Open backlog by age", "Backlog ageing", ["Backlog"], nonnull="Backlog",
             exclude_null=True, color=C_TARGET))
    add(table(R, "O | Backlog partners", "Top 20 partners by open backlog", ["Partner (sales)", "Branch (clean)"],
              ["Backlog", "Backlog over 60 days", "Sales"], sort_by="Backlog", top=(20, "Backlog (row)"),
              widths=[240, 80], total=False, nonnull="Backlog"))

    # ---- Explore
    add(table(R, "X | Explore table", "Explore", ["Explore by"],
              ["Sales", "Sales PY", "Growth %", "Share of sales %", "OB", "Backlog", "Active partners", "Quantity"],
              sort_by="Sales", widths=[240]))
    S["X | Explore table"].title = (run("Explore by ", C_DARK, 11, True) + run("<[Parameters].[p Dim]>", C_DARK, 11, True)
                                    + run(" (", C_DARK, 11, True) + run("<[Parameters].[p Units]>", C_DARK, 11, True) + run(")", C_DARK, 11, True))
    add(bars(R, "X | Explore bars", "Top 30", "Explore by", ["Sales", "Sales PY"], sort_by="Sales",
             top=(30, "Sales CY (row)"), nonnull="Has sales (agg)"))

    # row-level security on partner-level sheets
    for name, s in S.items():
        if name.startswith(("Pa |", "X |", "O | Backlog partners")):
            s.bool_filter("Branch access (sales)", context=True)
        if name.startswith("T |") or name == "K | Partner target":
            s.bool_filter("Branch access (partner targets)", context=True)
    return S


# ---------------------------------------------------------------- dashboards
SIDEBAR_PARAMS = [("p Period", "compact"), ("p Start", "datetime"), ("p End", "datetime"), ("p Measure", "compact"),
                  ("p Units", "compact"), ("p Region", "compact"), ("p Branch", "compact"), ("p Family", "compact"),
                  ("p Channel", "compact"), ("p Basis", "compact")]

NOTES = {
    "Home": "Use the buttons above to move between pages. Filters on the left apply to every page.",
    "Executive": "Are we on track? Sales vs last year and vs target, by region and branch.",
    "Branches": "Branch and ASM scorecards. Achievement uses monthly branch targets (net; gross = net / 0.83).",
    "Products": "Family, activity and SKU performance, and the focus families Wiser and Miluz Zeta.",
    "Partners": "Partner base, slabs, new / revival partners and the partners to watch.",
    "Partner Targets": "QTOD / ATOD partner targets vs achievement (targets are gross; net = gross x 0.83).",
    "Order Book": "Order intake (OB) against last year, conversion to sales and the open backlog.",
    "Explore": "Pick any dimension in 'Explore by' to slice sales, OB, backlog and partners.",
}

LAYOUT = {
    # page: (height, [(sheet, x, y, w, h)])  -- content area starts at x=230, width 1150
    "Home": (980, [
        ("K | Sales", 230, 112, 280, 120), ("K | Achievement", 520, 112, 280, 120), ("K | Annual target", 810, 112, 280, 120), ("K | Run-rate", 1100, 112, 280, 120),
        ("K | Order book", 230, 242, 280, 120), ("K | Backlog", 520, 242, 280, 120), ("K | Partners", 810, 242, 280, 120), ("K | Data", 1100, 242, 280, 120),
        ("E | Region table", 230, 372, 1150, 230),
        ("E | Monthly trend", 230, 612, 1150, 350)]),
    "Executive": (1560, [
        ("K | Sales", 230, 112, 280, 120), ("K | Achievement", 520, 112, 280, 120), ("K | Annual target", 810, 112, 280, 120), ("K | Run-rate", 1100, 112, 280, 120),
        ("K | Order book", 230, 242, 280, 120), ("K | Backlog", 520, 242, 280, 120), ("K | Partners", 810, 242, 280, 120), ("K | Data", 1100, 242, 280, 120),
        ("E | Monthly trend", 230, 372, 570, 330), ("E | Sales vs target", 810, 372, 570, 330),
        ("E | Region table", 230, 712, 1150, 230),
        ("E | Branch achievement", 230, 952, 570, 590), ("E | Branch growth", 810, 952, 570, 590)]),
    "Branches": (1980, [
        ("K | Sales", 230, 112, 280, 120), ("K | Achievement", 520, 112, 280, 120), ("K | Annual target", 810, 112, 280, 120), ("K | Run-rate", 1100, 112, 280, 120),
        ("B | Branch table", 230, 242, 1150, 640),
        ("B | Focus families", 230, 892, 1150, 520),
        ("B | ASM table", 230, 1422, 1150, 540)]),
    "Products": (1900, [
        ("K | Sales", 230, 112, 370, 120), ("K | Per partner", 620, 112, 370, 120), ("K | Data", 1010, 112, 370, 120),
        ("P | Family bars", 230, 242, 570, 620), ("P | Family table", 810, 242, 570, 620),
        ("P | Activity", 230, 872, 570, 240), ("P | Classification", 810, 872, 570, 460),
        ("P | Focus by region", 230, 1122, 570, 210),
        ("P | Top SKUs", 230, 1342, 1150, 540)]),
    "Partners": (1670, [
        ("K | Partners", 230, 112, 370, 120), ("K | Per partner", 620, 112, 370, 120), ("K | Sales", 1010, 112, 370, 120),
        ("Pa | Slabs", 230, 242, 370, 260), ("Pa | Tier", 620, 242, 370, 260), ("Pa | Partner type", 1010, 242, 370, 260),
        ("Pa | New partners", 230, 512, 1150, 260),
        ("Pa | Top partners", 230, 782, 570, 600), ("Pa | At-risk partners", 810, 782, 570, 600),
        ("Pa | SSA", 230, 1392, 1150, 260)]),
    "Partner Targets": (1700, [
        ("K | Partner target", 230, 112, 570, 120), ("K | Data", 810, 112, 570, 120),
        ("T | Annual", 230, 242, 1150, 720),
        ("T | Quarterly", 230, 972, 1150, 710)]),
    "Order Book": (1500, [
        ("K | Order book", 230, 112, 280, 120), ("K | Sales to OB", 520, 112, 280, 120), ("K | Backlog", 810, 112, 280, 120), ("K | Data", 1100, 112, 280, 120),
        ("O | OB trend", 230, 242, 760, 320), ("O | Backlog ageing", 1000, 242, 380, 320),
        ("O | OB by branch", 230, 572, 1150, 560),
        ("O | Backlog partners", 230, 1142, 1150, 340)]),
    "Explore": (1300, [
        ("X | Explore table", 230, 112, 640, 1170), ("X | Explore bars", 880, 112, 500, 1170)]),
}
