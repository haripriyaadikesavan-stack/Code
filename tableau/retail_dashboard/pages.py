"""Sheets and dashboards of the Retail dashboard."""
from core import C_MID, C_TARGET, MERGE, OLD_CUR, OLD_END, OLD_GN, OLD_START, PDS
from sheets import ACH_SCALE, GROWTH_SCALE, MN, MV, UNITS_TOKEN, Sheet, kpi, nav, run, title_runs, is_value

REGIONS = ["East", "North", "South", "West1", "West2"]
# top navigation: (button label, dashboard name)
NAV = [("Home", "Home"), ("MTD", "MTD"), ("Executive", "Executive"), ("Regions", "Region - East"),
       ("Branches", "Branches"), ("Products", "Products"), ("Partners", "Partners"),
       ("Target vs Ach", "Target vs Achievement"), ("SSP", "SSP"), ("Pavilion", "Pavilion"),
       ("Order Book", "Order Book"), ("Explore", "Explore")]
REGION_DASH = [f"Region - {r}" for r in REGIONS]
DASHBOARDS = [d for _, d in NAV if not d.startswith("Region")] + REGION_DASH
if MERGE:
    # analysis pages added to the existing workbook; "Home" goes back to its Retail Overall Dashboard
    NAV = [("Home", "Retail Overall Dashboard"), ("Executive", "Executive"), ("Branches", "Branches"),
           ("Products", "Products"), ("Partners", "Partners"), ("Target vs Ach", "Target vs Achievement"),
           ("Order Book", "Order Book"), ("Explore", "Explore")]
    DASHBOARDS = [d for _, d in NAV[1:]]


# ---------------------------------------------------------------- generic builders
def F(R, x):
    return x if not isinstance(x, str) else (R[x] if x in R.by_caption else R.fields[x])


def table(R, name, title, dims, measures, sub=None, total=True, sort_by=None, top=None, nonnull="Has activity (agg)",
          widths=None, scale=None, units=None, sort_dims=None):
    caps = [m if isinstance(m, str) else None for m in measures]
    if units is None:
        units = any(c and is_value(R, c) for c in caps)
    s = Sheet(R, name, title_runs(title, sub, param_units=units), "table")
    drefs = [s.use(F(R, d), "dim") for d in dims]
    s.measures(measures)
    s.rows = " / ".join(drefs)
    s.rows_total = total
    s.cols = s.mn
    s.encodings.append(("text", s.mv))
    if scale:
        cap, sc = scale
        ref = s.use(cap, "usr")
        s.encodings.append(("color", ref))
        s.scales.append((ref, sc))
    if nonnull:
        s.nonnull(nonnull)
    if top:
        n, by = top
        s.top_n(drefs[0], F(R, dims[0]), n, by)
    if sort_by:
        cap, *direction = sort_by if isinstance(sort_by, tuple) else (sort_by,)
        for i in (sort_dims if sort_dims is not None else [len(drefs) - 1]):
            s.sort_by(drefs[i], s.use(cap, "usr"), direction=(direction[0] if direction else "DESC"))
    for i, w in enumerate(widths or []):
        s.widths.append((drefs[i], w))
    s.cell_align = "right"
    s.fit = "fit-width"
    return s


def hbar(R, name, title, dim, measure, sub=None, scale=None, top=None, nonnull="Has sales (agg)", sort=True,
         exclude_null=False, color=None, vertical=False):
    """Single-measure bar chart, sorted, labelled; value axis hidden.
    scale: ('Ach %' | 'Growth %', SCALE) colours bars by that measure."""
    s = Sheet(R, name, title_runs(title, sub, param_units=is_value(R, measure)), "chart")
    f = F(R, dim)
    dref = s.use(f, "dim")
    val = s.use(measure, "usr")
    s.encodings.append(("text", val))
    if scale:
        cap, sc = scale
        cref = s.use(cap, "usr")
        s.encodings.append(("color", cref))
        s.scales.append((cref, sc))
    else:
        s.mark_color = color or C_MID
    if vertical:
        s.cols, s.rows = dref, val
        s.hide_axis.append((val, "rows"))
    else:
        s.rows, s.cols = dref, val
        s.hide_axis.append((val, "cols"))
    s.mark = "Bar"
    if nonnull:
        s.nonnull(nonnull)
    if exclude_null:
        s.exclude_null(dref, f)
    if top:
        n, by = top
        s.top_n(dref, f, n, by)
    if sort:
        s.sort_by(dref, val, shelf="cols" if vertical else "rows")
    return s


def trend(R, name, title, measures, sub=None, dim="Month", nonnull="Has activity (agg)"):
    """Two lines (e.g. this year vs last year) by month or day; value axis hidden."""
    s = Sheet(R, name, title_runs(title, sub, param_units=is_value(R, measures[0])), "chart")
    dref = s.use(dim, "dim")
    s.measures(measures)
    s.encodings.append(("color", s.mn))
    s.cols, s.rows = dref, s.mv
    s.mark = "Line"
    s.hide_axis.append((s.mv, "rows"))
    s.nonnull(nonnull)
    s.exclude_null(dref, R[dim])
    return s


def heat(R, name, title, rowdim, coldim, measure, scale, sub=None, nonnull="Has sales (agg)"):
    s = Sheet(R, name, title_runs(title, sub, param_units=False), "table")
    r = s.use(F(R, rowdim), "dim")
    c = s.use(F(R, coldim), "dim")
    v = s.use(measure, "usr")
    s.rows, s.cols = r, c
    s.mark = "Square"
    s.encodings += [("color", v), ("text", v)]
    s.scales.append((v, scale))
    s.nonnull(nonnull)
    s.exclude_null(c, F(R, coldim))
    s.cell_align = "center"
    s.fit = "fit-width"
    return s


def kpi_set(R, prefix):
    """The standard KPI cards (two values max per card)."""
    return [
        kpi(R, f"{prefix}Sales", "Sales", "Sales", ("", "Growth text")),
        kpi(R, f"{prefix}Achievement", "Achievement vs target", "Ach %", ("Target  ", "Target")),
        kpi(R, f"{prefix}Yearly target", "Yearly target", "Yearly target", ("Achieved  ", "Yearly Ach %")),
        kpi(R, f"{prefix}Run-rate", "Monthly run-rate", "Current run-rate / month", ("Required  ", "Required run-rate / month")),
        kpi(R, f"{prefix}Order book", "Order book (OB)", "OB", ("", "OB growth text")),
        kpi(R, f"{prefix}Backlog", "Open backlog", "Backlog", ("Older than 60 days  ", "Backlog over 60 days")),
        kpi(R, f"{prefix}Partners", "Active partners", "Active partners", ("New / revival  ", "New / revival partners")),
        kpi(R, f"{prefix}Data", "Data as of", "Data as of", ("Period  ", "Period text")),
    ]


# ---------------------------------------------------------------- the sheets
def build_sheets(R):
    """Returns {name: Sheet}; each sheet gets .scope: 'P' / 'Q' / 'S' / 'PV' / 'nav' / 'region:X' and .rls."""
    S = {}

    def add(s, scope="P", rls=None):
        s.scope, s.rls = scope, rls
        S[s.name] = s
        return s

    for label, dash in NAV:
        add(nav(R, dash, label), "nav")
    for r in REGIONS:
        add(nav(R, f"Region - {r}", f"Region {r}"), "nav")

    # ---- shared KPIs
    for k in kpi_set(R, "K | "):
        add(k)
    add(kpi(R, "K | Wiser", "Wiser sales", "Wiser sales", ("Achieved  ", "Wiser Ach %")))
    add(kpi(R, "K | Zeta", "Miluz Zeta sales", "Zeta sales", ("Achieved  ", "Zeta Ach %")))
    add(kpi(R, "K | Per partner", "Sales per active partner", "Sales per partner", ("New / revival partner sales  ", "New partner sales")))
    add(kpi(R, "K | New partners", "New / revival partners", "New / revival partners", ("Their sales  ", "New partner sales"), units=False))
    add(kpi(R, "K | Sales to OB", "Sales to OB", "Sales to OB %", ("Sales  ", "Sales")))
    add(kpi(R, "K | SSP sales", "Sales through SSPs", "SSP sales", ("Growth  ", "SSP growth %")))
    add(kpi(R, "K | SSP target", "SSP target", "SSP target", ("Achieved  ", "SSP Ach %")))
    add(kpi(R, "K | Partner target", "Partner target (QTOD / ATOD)", "Partner target", ("Achieved  ", "Partner Ach %")), "Q", "Q")

    # ---- Home / Executive
    add(trend(R, "E | Monthly trend", "Monthly sales - this year vs last year", ["Sales", "Sales PY"],
              "Green = this year, grey = same period last year"))
    add(trend(R, "E | Sales vs target", "Monthly sales vs target", ["Sales", "Target"],
              "Green = sales, amber = monthly branch target", nonnull="Has sales (agg)"))
    add(table(R, "E | Region table", "Region performance", ["Region"],
              ["Sales", "Sales PY", "Growth %", "Target", "Ach %", "Yearly target", "Yearly Ach %", "OB", "Active partners"],
              widths=[90]))
    add(hbar(R, "E | Branch achievement", "Branch achievement vs target", "Branch", "Ach %",
             "Red < 75%, amber 75-100%, green above 100%", scale=("Ach %", ACH_SCALE), nonnull="Ach %"))
    add(hbar(R, "E | Branch growth", "Branch growth vs last year", "Branch", "Growth %",
             "Green = growing, red = declining", scale=("Growth %", GROWTH_SCALE)))

    # ---- MTD
    add(kpi(R, "M | Sales", "MTD sales", "MTD sales", ("", "MTD growth text")))
    add(kpi(R, "M | Achievement", "MTD achievement", "MTD Ach %", ("Month target  ", "MTD target")))
    add(kpi(R, "M | OB", "MTD order book", "MTD OB", ("Yesterday  ", "OB yesterday")))
    add(kpi(R, "M | Yesterday", "Yesterday's sales", "Sales yesterday", ("Month  ", "Month text")))
    add(trend(R, "M | Daily trend", "Day-wise sales - this month vs same month last year", ["MTD sales", "MTD sales LY"],
              dim="Day", nonnull="Has MTD activity (agg)"))
    add(table(R, "M | Branch table", "MTD branch performance", ["Region", "Branch"],
              ["MTD sales", "MTD sales LY", "MTD growth %", "MTD target", "MTD Ach %", "MTD OB",
               "MTD Wiser Ach %", "MTD Zeta Ach %", "MTD active partners"],
              widths=[70, 90], nonnull="Has MTD activity (agg)"))
    add(table(R, "M | ASM table", "MTD ASM performance", ["Region", "ASM"],
              ["MTD sales", "MTD sales LY", "MTD growth %", "MTD OB", "MTD active partners"],
              widths=[70, 180], nonnull="Has MTD activity (agg)", sort_by="MTD sales"))
    add(hbar(R, "M | Family", "MTD sales by family", "Family", "MTD sales", top=(15, "Sales MTD (row)"),
             nonnull="Has MTD activity (agg)"))
    add(table(R, "M | Partners", "MTD top 20 partners", ["Partner", "Branch"], ["MTD sales", "MTD OB"],
              sort_by="MTD sales", top=(20, "Sales MTD (row)"), widths=[260, 80], total=False,
              nonnull="Has MTD activity (agg)"), "P", "P")

    # ---- Branches
    add(table(R, "B | Branch table", "Branch scorecard", ["Region", "Branch"],
              ["Sales", "Sales PY", "Growth %", "Target", "Ach %", "Gap to target", "Yearly target", "Yearly Ach %",
               "OB", "Active partners"], widths=[70, 90]))
    add(table(R, "B | Focus families", "Focus families by branch - Wiser and Miluz Zeta", ["Region", "Branch"],
              ["Wiser sales", "Wiser target", "Wiser Ach %", "Zeta sales", "Zeta target", "Zeta Ach %"], widths=[70, 90]))
    add(table(R, "B | ASM table", "ASM performance", ["Region", "ASM"],
              ["Sales", "Sales PY", "Growth %", "Share of sales %", "Active partners", "New / revival partners", "Sales per partner"],
              widths=[70, 180], nonnull="Has sales (agg)", sort_by="Sales"))

    # ---- Products
    add(hbar(R, "P | Family bars", "Sales by family", "Family", "Sales", "Top 15; green = growing, red = declining vs last year",
             scale=("Growth %", GROWTH_SCALE), top=(15, "Sales CY (row)")))
    add(table(R, "P | Family table", "Family performance", ["Family"],
              ["Sales", "Sales PY", "Growth %", "Share of sales %", "Quantity"], sort_by="Sales", widths=[170],
              nonnull="Has sales (agg)"))
    add(hbar(R, "P | Activity", "Sales by activity (Helios code)", "Helios Code", "Sales", scale=("Growth %", GROWTH_SCALE),
             exclude_null=True))
    add(hbar(R, "P | Classification", "Sales by product classification", "Classification (Sheet1)", "Sales",
             scale=("Growth %", GROWTH_SCALE), exclude_null=True))
    add(table(R, "P | Focus by region", "Focus families by region", ["Region"],
              ["Wiser sales", "Wiser target", "Wiser Ach %", "Zeta sales", "Zeta target", "Zeta Ach %"], widths=[90]))
    add(table(R, "P | Top SKUs", "Top 25 materials", ["MaterialDescription"],
              ["Sales", "Sales PY", "Growth %", "Share of sales %", "Quantity"], sort_by="Sales",
              top=(25, "Sales CY (row)"), widths=[320], nonnull="Has sales (agg)", total=False))

    # ---- Partners (row-level security)
    s = add(hbar(R, "Pa | Slabs", "Partners by sales slab", "Partner slab", "Partners", "Partner sales in the selected period",
                 nonnull="Partners", exclude_null=True, sort=False), "P", "P")
    add(hbar(R, "Pa | Tier", "Sales by tier", "Tier Category", "Sales", scale=("Growth %", GROWTH_SCALE), exclude_null=True), "P", "P")
    add(hbar(R, "Pa | Partner type", "Sales by partner type", "RD/KRO", "Sales", scale=("Growth %", GROWTH_SCALE), exclude_null=True), "P", "P")
    add(hbar(R, "Pa | New partners", "New / revival partners by month", "Month", "New / revival partners",
             "First sale this year (no sales in the previous two years)", nonnull="New / revival partners",
             exclude_null=True, sort=False, vertical=True), "P", "P")
    add(hbar(R, "Pa | Onboarded", "Partners by onboarded year", "Onboarded Year (Partner Master Data)", "Onboarded partners",
             "From the Partner Master", nonnull=None, exclude_null=True, sort=False, vertical=True), "P", "P")
    add(table(R, "Pa | Top partners", "Top 25 partners", ["Partner", "Branch", "Partner status"],
              ["Sales", "Sales PY", "Growth %", "Share of sales %"], sort_by="Sales", top=(25, "Sales CY (row)"),
              widths=[300, 90, 110], total=False, nonnull="Has sales (agg)"), "P", "P")
    add(table(R, "Pa | At-risk partners", "Largest last-year partners now declining", ["Partner", "Branch"],
              ["Sales PY", "Sales", "Growth", "Growth %"], sort_by=("Growth", "ASC"), top=(25, "Sales PY (row)"),
              widths=[300, 90], total=False, nonnull="Has sales (agg)", scale=("Growth %", GROWTH_SCALE)), "P", "P")

    # ---- Target vs Achievement
    add(table(R, "T | Branch", "Branch target vs achievement", ["Region", "Branch"],
              ["Target", "Sales", "Ach %", "Gap to target", "Yearly target", "Yearly Ach %", "Balance to yearly target",
               "Wiser Ach %", "Zeta Ach %"], widths=[70, 90], nonnull="Has sales (agg)"))
    add(heat(R, "T | Month heat", "Achievement by branch and month", "Branch", "Month", "Ach %", ACH_SCALE,
             "Red < 75%, amber 75-100%, green above 100%"))
    add(table(R, "T | Partner branches", "Partner targets by branch", ["Target partner branch"],
              ["Partner target", "Partner achieved", "Partner Ach %", "Partner balance"],
              sub="View selected in 'Target view' (annual 100% / 108%, quarter, Wiser, Miluz Zeta)",
              nonnull="Has partner target", widths=[90], scale=("Partner Ach %", ACH_SCALE)), "Q", "Q")
    add(table(R, "T | Partners", "Partner target vs achievement", ["Target partner branch", "Target partner"],
              ["Partner target", "Partner achieved", "Partner Ach %", "Partner balance"],
              sub="Partner targets are gross (QTOD / ATOD file); Net = gross x 0.83",
              nonnull="Has partner target", widths=[80, 320], scale=("Partner Ach %", ACH_SCALE),
              sort_by="Partner target"), "Q", "Q")

    # ---- SSP
    add(table(R, "S | SSP table", "SSP target vs achievement", ["SSP", "Branch (SSA Targets)"],
              ["SSP target", "SSP sales", "SSP Ach %", "SSP sales PY", "SSP growth %"],
              sub="Sales of partners mapped to the SSP; SSP targets are used as they are",
              nonnull="SSP target", widths=[200, 90], scale=("SSP Ach %", ACH_SCALE), total=False), "S")
    add(table(R, "S | With vs without", "With vs without SSP", ["SSA group"],
              ["Sales", "Sales PY", "Growth %", "Share of sales %", "Active partners", "Sales per partner"], widths=[120]))
    add(trend(R, "S | Trend", "Monthly sales through SSPs - this year vs last year", ["SSP sales", "SSP sales PY"]))
    s = add(table(R, "S | Partners", "Top 25 SSP partners", ["SSA", "Partner", "Branch"],
                  ["Sales", "Sales PY", "Growth %"], sort_by="Sales", top=(25, "SSA sales CY (row)"),
                  widths=[160, 300, 80], total=False, nonnull="Has sales (agg)"), "P", "P")
    s.member_filter("SSA group", "With SSP")

    # ---- Order book
    add(trend(R, "O | OB trend", "Monthly order book - this year vs last year", ["OB", "OB PY"]))
    add(hbar(R, "O | Backlog ageing", "Open backlog by age", "Backlog ageing", "Backlog", nonnull="Backlog",
             exclude_null=True, sort=False, color=C_TARGET))
    add(table(R, "O | OB by branch", "Order book and backlog by branch", ["Region", "Branch"],
              ["OB", "OB PY", "OB growth %", "Sales", "Sales to OB %", "Backlog", "Backlog over 60 days"], widths=[70, 90]))
    add(table(R, "O | Backlog partners", "Top 20 partners by open backlog", ["Partner", "Branch"],
              ["Backlog", "Backlog over 60 days", "Sales"], sort_by="Backlog", top=(20, "Backlog (row)"),
              widths=[300, 90], total=False, nonnull="Backlog"), "P", "P")

    # ---- Explore
    s = add(table(R, "X | Explore table", "Explore", ["Explore by"],
                  ["Sales", "Sales PY", "Growth %", "Share of sales %", "OB", "Backlog", "Active partners", "Quantity"],
                  sort_by="Sales", widths=[300]), "P", "P")
    s.title = (run("Explore by ", "#0b4d33", 11, True) + run("<[Parameters].[p Dim]>", "#0b4d33", 11, True)
               + run(" (", "#0b4d33", 11, True) + run(UNITS_TOKEN, "#0b4d33", 11, True) + run(")", "#0b4d33", 11, True))
    add(hbar(R, "X | Explore bars", "Top 30 by sales", "Explore by", "Sales", scale=("Growth %", GROWTH_SCALE),
             top=(30, "Sales CY (row)")), "P", "P")

    # ---- Pavilion (published summary tables)
    pf = lambda n: R.fields[f"{PDS}|{n}"]
    MS, PT, RA, SC, TT, WS = ("(Minimum!Slab!Achievement)", "(Pavilion!Type!Summary)", "(Retailer!App!TnC!Acceptance)",
                              "(Sales!Classification!Summary)", "(Total!Target!Achievement)", "(Wiser!Slabs!Unpivoted)")
    s = add(table(R, "PV | Type", "Pavilions by type", [pf("Pavilion Type"), pf(f"Sum of YTD Slab 2025 {PT}"), pf("Sum of YTD Slab 2026")],
                  [(pf("Count of 2026 Classification"), "sum")], total=False, nonnull=None, widths=[110, 120, 120], units=False,
                  sort_by=("Type row order", "ASC"), sort_dims=[0]), "PV")
    s = add(table(R, "PV | Classification", "Pavilions by classification",
                  [pf("Classification"), pf("Sum of YTD Slab 2025"), pf(f"Sum of YTD Slab 2026 {SC}")],
                  [(pf(f"Count of 2026 Classification {SC}"), "sum"), (pf(f"2025-26 Growth % {SC}"), "sum")],
                  total=False, nonnull=None, widths=[110, 120, 120], units=False, sort_by=("Class row order", "ASC"), sort_dims=[0]), "PV")
    s = add(table(R, "PV | Min slab", "Minimum slab achievement by region", [pf(f"Zone/Region {MS}")],
                  [(pf(f"No. of KRO {MS}"), "sum"), (pf(f"No. of RO {MS}"), "sum"), (pf("No of KRO Achieved"), "sum"),
                   (pf("No. of RO Achieved"), "sum"), "Min slab qualifier %"], nonnull=None, widths=[100], units=False), "PV")
    s.member_filter(pf(f"Row Type {MS}"), "Region")
    s = add(table(R, "PV | Total target", "Total target achievement by region", [pf(f"Zone/Region {TT}")],
                  [(pf(f"No. of KRO {TT}"), "sum"), (pf(f"No. of RO {TT}"), "sum"), (pf(f"No of KRO Achieved {TT}"), "sum"),
                   (pf(f"No. of RO Achieved {TT}"), "sum"), "Total target qualifier %"], nonnull=None, widths=[100], units=False), "PV")
    s.member_filter(pf(f"Row Type {TT}"), "Region")
    s = add(table(R, "PV | App T&C", "Retailer app T&C acceptance", [pf(f"Zone/Region {RA}")],
                  [(pf("Total no. of RO Pavilions"), "sum"), (pf("Accepted"), "sum"), (pf("Pending"), "sum"), "Acceptance %"],
                  nonnull=None, widths=[100], units=False), "PV")
    s.member_filter(pf(f"Row Type {RA}"), "Region")
    s = add(table(R, "PV | Sales slabs", "Sales slabs by ASM - 2025 vs 2026",
                  [pf("Zone/Region"), pf("Sales Manager / Entity"), pf("Branch"), pf("2025 Sales Slab"), pf("2026 Sales Slab")],
                  [(pf("2025-26 Growth %"), "sum")], total=False, nonnull=None, widths=[80, 200, 70, 120, 120], units=False,
                  sort_by=("Slab row order", "ASC"), sort_dims=[0, 1]), "PV")
    s = add(table(R, "PV | Wiser slabs", "Wiser slab achievement by ASM",
                  [pf(f"Zone/Region {WS}"), pf(f"Sales Manager / Entity {WS}"), pf(f"Branch {WS}")],
                  [(pf("No. of KRO"), "sum"), (pf("No. of RO"), "sum"), (pf("No of KRO > 1 lac"), "sum"),
                   (pf("No. of RO > 25K"), "sum"), "Wiser qualifier %"], nonnull=None, widths=[90, 200, 70], units=False), "PV")
    s.member_filter(pf(f"Row Type {WS}"), "ASM")

    # ---- Region dashboards (fixed region filter, like the old regional dashboards)
    for reg in REGIONS:
        p = f"{reg} | "
        for k in kpi_set(R, p)[:5]:
            add(k, f"region:{reg}")
        add(trend(R, p + "Trend", f"{reg} - monthly sales, this year vs last year", ["Sales", "Sales PY"]), f"region:{reg}")
        add(table(R, p + "Branches", f"{reg} - branch performance", ["Branch"],
                  ["Sales", "Sales PY", "Growth %", "Target", "Ach %", "Yearly target", "Yearly Ach %", "OB", "Active partners"],
                  widths=[100]), f"region:{reg}")
        add(hbar(R, p + "Families", f"{reg} - sales by family", "Family", "Sales", "Top 12; green = growing, red = declining",
                 scale=("Growth %", GROWTH_SCALE), top=(12, "Sales CY (row)")), f"region:{reg}")
        add(table(R, p + "Focus", f"{reg} - Wiser and Miluz Zeta by branch", ["Branch"],
                  ["Wiser sales", "Wiser target", "Wiser Ach %", "Zeta sales", "Zeta target", "Zeta Ach %"], widths=[100]), f"region:{reg}")
        add(table(R, p + "ASM", f"{reg} - ASM performance", ["ASM"],
                  ["Sales", "Sales PY", "Growth %", "Active partners", "New / revival partners"], widths=[200],
                  nonnull="Has sales (agg)", sort_by="Sales"), f"region:{reg}")
        add(table(R, p + "Top partners", f"{reg} - top 10 partners", ["Partner", "Branch"],
                  ["Sales", "Sales PY", "Growth %", "Share of sales %"], sort_by="Sales", top=(10, "Sales CY (row)"),
                  widths=[300, 90], total=False, nonnull="Has sales (agg)"), f"region:{reg}")

    # ---- filters: region selection, fixed regions, row-level security
    for name, s in S.items():
        if s.scope in ("P", "S"):
            s.bool_filter("Region selected", context=True)
        elif s.scope == "Q":
            s.bool_filter("Target partner region selected", context=True)
        elif s.scope.startswith("region:"):
            s.member_filter(R.fields["Zone"], s.scope.split(":")[1])
        if s.rls == "P":
            s.bool_filter("Branch access (sales)", context=True)
        elif s.rls == "Q":
            s.bool_filter("Branch access (partner targets)", context=True)
    return S


# ---------------------------------------------------------------- dashboards
BASE_PARAMS = [("p Period", "compact"), ("p Start", "datetime"), ("p End", "datetime"), ("p Measure", "compact"),
               ("p Units", "compact"), ("p Region", "compact"), ("p Branch", "compact"), ("p Family", "compact"),
               ("p Channel", "compact"), ("p Basis", "compact")]
if MERGE:
    BASE_PARAMS = [(OLD_START, "datetime"), (OLD_END, "datetime"), (OLD_GN, "compact"), (OLD_CUR, "compact"),
                   ("p Region", "compact"), ("p Branch", "compact"), ("p Family", "compact"), ("p Channel", "compact"),
                   ("p Basis", "compact")]
NO_PERIOD = [p for p in BASE_PARAMS if p[0] not in ("p Period", "p Start", "p End", "p Basis")]


def params_for(dash):
    if dash == "MTD":
        return NO_PERIOD
    if dash == "Pavilion":
        return []
    if dash.startswith("Region - "):
        return [p for p in BASE_PARAMS if p[0] != "p Region"]
    if dash == "Target vs Achievement":
        return BASE_PARAMS + [("p View", "compact")]
    if dash == "Explore":
        return BASE_PARAMS + [("p Dim", "compact")]
    return BASE_PARAMS


NOTES = {
    "Home": "Headline numbers. Use the buttons to move between pages; the filters on the left apply to every page.",
    "MTD": "Month to date (1st of this month to today) vs the same days last year, and this month's target.",
    "Executive": "Are we on track? Sales vs last year and vs target, by region and branch.",
    "Branches": "Branch and ASM scorecards.",
    "Products": "Family, activity, classification and SKU performance; focus families Wiser and Miluz Zeta.",
    "Partners": "Partner base, slabs, new / revival partners and the partners to watch.",
    "Target vs Achievement": "Monthly branch targets (net), yearly branch targets (net) and partner targets (gross).",
    "SSP": "Sales through SSPs vs SSP targets, and with vs without SSP.",
    "Pavilion": "Pavilion slabs and qualification, as published in Pavilion_Unpivoted.xlsx.",
    "Order Book": "Order intake (OB) vs last year, conversion to sales and the open backlog.",
    "Explore": "Pick any dimension in 'Explore by' to slice sales, OB, backlog and partners.",
}
for r in REGIONS:
    NOTES[f"Region - {r}"] = f"{r} region - branches, families, ASMs and top partners."

SIDE_NOTE = ("Gross / Net: monthly and yearly targets are net (gross = net / 0.83); partner targets are gross (net = gross x 0.83).\n"
             "Growth compares the same dates last year.\n"
             "Monthly target basis: 'Full months' counts each month's full target; 'Pro-rated to date' counts the current month only up to today.")

K4 = lambda names, y=112: [(n, 230 + i * 290, y, 280, 110) for i, n in enumerate(names)]

LAYOUT = {
    "Home": (1000, K4(["K | Sales", "K | Achievement", "K | Yearly target", "K | Run-rate"])
             + K4(["K | Order book", "K | Backlog", "K | Partners", "K | Data"], 232)
             + [("E | Region table", 230, 352, 1150, 250), ("E | Monthly trend", 230, 612, 1150, 370)]),
    "MTD": (1640, K4(["M | Sales", "M | Achievement", "M | OB", "M | Yesterday"])
            + [("M | Daily trend", 230, 232, 1150, 300), ("M | Branch table", 230, 542, 1150, 520),
               ("M | Family", 230, 1072, 400, 550), ("M | ASM table", 640, 1072, 740, 550)]),
    "Executive": (1520, K4(["K | Sales", "K | Achievement", "K | Yearly target", "K | Run-rate"])
                  + K4(["K | Order book", "K | Backlog", "K | Partners", "K | Data"], 232)
                  + [("E | Monthly trend", 230, 352, 570, 320), ("E | Sales vs target", 810, 352, 570, 320),
                     ("E | Region table", 230, 682, 1150, 250),
                     ("E | Branch achievement", 230, 942, 570, 560), ("E | Branch growth", 810, 942, 570, 560)]),
    "Branches": (1780, K4(["K | Sales", "K | Achievement", "K | Yearly target", "K | Data"])
                 + [("B | Branch table", 230, 232, 1150, 560), ("B | Focus families", 230, 802, 1150, 480),
                    ("B | ASM table", 230, 1292, 1150, 470)]),
    "Products": (1820, K4(["K | Sales", "K | Wiser", "K | Zeta", "K | Data"])
                 + [("P | Family bars", 230, 232, 560, 420), ("P | Family table", 800, 232, 580, 420),
                    ("P | Activity", 230, 662, 560, 230), ("P | Focus by region", 800, 662, 580, 230),
                    ("P | Classification", 230, 902, 1150, 380),
                    ("P | Top SKUs", 230, 1292, 1150, 510)]),
    "Partners": (1700, K4(["K | Partners", "K | New partners", "K | Per partner", "K | Data"])
                 + [("Pa | Slabs", 230, 232, 370, 250), ("Pa | Tier", 620, 232, 370, 250), ("Pa | Partner type", 1010, 232, 370, 250),
                    ("Pa | New partners", 230, 492, 760, 250), ("Pa | Onboarded", 1000, 492, 380, 250),
                    ("Pa | Top partners", 230, 752, 1150, 460), ("Pa | At-risk partners", 230, 1222, 1150, 460)]),
    "Target vs Achievement": (2000, K4(["K | Achievement", "K | Yearly target", "K | Partner target", "K | Run-rate"])
                              + [("T | Branch", 230, 232, 1150, 560), ("T | Month heat", 230, 802, 1150, 500),
                                 ("T | Partner branches", 230, 1312, 400, 670), ("T | Partners", 640, 1312, 740, 670)]),
    "SSP": (1500, K4(["K | SSP sales", "K | SSP target", "K | Sales", "K | Data"])
            + [("S | SSP table", 230, 232, 1150, 300), ("S | With vs without", 230, 542, 560, 160), ("S | Trend", 800, 542, 580, 300),
               ("S | Partners", 230, 852, 1150, 630)]),
    "Pavilion": (1680, [("PV | Type", 230, 112, 560, 170), ("PV | Classification", 800, 112, 580, 170),
                        ("PV | Min slab", 230, 292, 560, 260), ("PV | Total target", 800, 292, 580, 260),
                        ("PV | App T&C", 230, 562, 560, 260), ("PV | Sales slabs", 230, 832, 1150, 420),
                        ("PV | Wiser slabs", 230, 1262, 1150, 400)]),
    "Order Book": (1520, K4(["K | Order book", "K | Sales to OB", "K | Backlog", "K | Data"])
                   + [("O | OB trend", 230, 232, 760, 320), ("O | Backlog ageing", 1000, 232, 380, 320),
                      ("O | OB by branch", 230, 562, 1150, 560), ("O | Backlog partners", 230, 1132, 1150, 370)]),
    "Explore": (1500, [("X | Explore table", 230, 112, 1150, 700), ("X | Explore bars", 230, 822, 1150, 660)]),
}
for r in REGIONS:
    p = f"{r} | "
    LAYOUT[f"Region - {r}"] = (1640, [(p + n, 230 + i * 230, 152, 220, 110) for i, n in
                                      enumerate(["Sales", "Achievement", "Yearly target", "Run-rate", "Order book"])]
                               + [(p + "Trend", 230, 272, 1150, 300), (p + "Branches", 230, 582, 1150, 300),
                                  (p + "Families", 230, 892, 560, 360), (p + "Focus", 800, 892, 580, 360),
                                  (p + "ASM", 230, 1262, 560, 360), (p + "Top partners", 800, 1262, 580, 360)])
