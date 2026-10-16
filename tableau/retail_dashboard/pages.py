"""Sheets and dashboards of the Retail dashboard (v4: one story in 12 chapters).

Every analysis appears once. Region-wise and branch-wise views come from the
Region / Branch filters on the left (they apply to every chapter) and from
clicking a row in the Regions / Branches scorecards.
"""
from core import C_DARK, C_MID, C_TARGET, PDS
from sheets import ACH_SCALE, GROWTH_SCALE, Sheet, kpi, nav, run, title_runs, is_value

REGIONS = ["East", "North", "South", "West1", "West2"]
# the story: (button label, dashboard name, one-line purpose)
CHAPTERS = [
    ("1 Overview", "Overview", "Where do we stand? Sales vs last year and target, three-year trend, regions."),
    ("2 Regions", "Regions", "Region-wise performance. Click a region to focus the page on it."),
    ("3 Branches", "Branches", "Branch-wise scorecard, focus families and ASM performance. Click a branch to see its ASMs."),
    ("4 Products", "Products", "Families, activities (FD / WD ...), new products, classification and top materials."),
    ("5 Partners", "Partners", "Partner base: achievement bands, slabs, tiers, new / revival, lapsed and at-risk partners."),
    ("6 All partners", "All Partners", "Numbers for every partner, and every partner month by month."),
    ("7 Targets", "Targets", "Branch targets by month and partner targets (QTOD / ATOD)."),
    ("8 SSP & reach", "SSP & Reach", "SSP performance, with vs without SSP, and new cities reached."),
    ("9 Pavilion", "Pavilion", "Pavilion partners vs target, slabs and qualification."),
    ("10 MTD", "MTD", "This month so far vs the same days last year and this month's target."),
    ("11 Orders", "Order Book", "Order book, conversion to sales, open backlog and every open order."),
    ("12 Explore", "Explore", "Any dimension against all measures."),
]
NAV = [(label, dash) for label, dash, _ in CHAPTERS]
DASHBOARDS = [dash for _, dash, _ in CHAPTERS]
NOTES = {dash: note for _, dash, note in CHAPTERS}
# click a row in the source sheet -> filter the rest of the dashboard
DRILL = {"Regions": ["R | Scorecard"], "Branches": ["B | Scorecard"], "Products": ["P | Family table"]}


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
    """Single-measure bar chart, sorted, labelled; value axis hidden."""
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
    """Lines (this year vs earlier years, or vs target) by month or day; value axis hidden."""
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


def grouped_bars(R, name, title, dim, measures, sub=None, nonnull="Has sales (agg)"):
    """Bars side by side per category (e.g. quarter: this year vs last year)."""
    s = Sheet(R, name, title_runs(title, sub, param_units=is_value(R, measures[0])), "chart")
    dref = s.use(dim, "dim")
    s.measures(measures)
    s.encodings.append(("color", s.mn))
    s.encodings.append(("text", s.mv))
    s.cols, s.rows = f"{dref} / {s.mn}", s.mv
    s.mark = "Bar"
    s.hide_axis.append((s.mv, "rows"))
    s.nonnull(nonnull)
    s.exclude_null(dref, R[dim])
    return s


def lines_by(R, name, title, dim, colour_dim, measure, sub=None, nonnull="Has sales (agg)"):
    """One line per member of colour_dim (e.g. monthly sales per region)."""
    s = Sheet(R, name, title_runs(title, sub, param_units=is_value(R, measure)), "chart")
    dref = s.use(dim, "dim")
    cref = s.use(F(R, colour_dim), "dim")
    val = s.use(measure, "usr")
    s.encodings.append(("color", cref))
    s.cols, s.rows = dref, val
    s.mark = "Line"
    s.hide_axis.append((val, "rows"))
    s.nonnull(nonnull)
    s.exclude_null(dref, R[dim])
    s.exclude_null(cref, F(R, colour_dim))
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


def matrix(R, name, title, rowdims, coldim, measure, sub=None, nonnull="Has sales (agg)", widths=None):
    """Text table: rows x columns (e.g. partner x month), one measure."""
    s = Sheet(R, name, title_runs(title, sub, param_units=is_value(R, measure)), "table")
    rrefs = [s.use(F(R, d), "dim") for d in rowdims]
    c = s.use(F(R, coldim), "dim")
    v = s.use(measure, "usr")
    s.rows, s.cols = " / ".join(rrefs), c
    s.rows_total = True
    s.encodings.append(("text", v))
    s.nonnull(nonnull)
    s.exclude_null(c, F(R, coldim))
    for i, w in enumerate(widths or []):
        s.widths.append((rrefs[i], w))
    s.cell_align = "right"
    s.fit = "fit-width"
    return s


def pivot(R, name, title, rowdims, coldim, measures, sub=None, nonnull="Has sales (agg)", widths=None):
    """Text table: row dimensions and measures down the side, a dimension (e.g. month) across."""
    s = Sheet(R, name, title_runs(title, sub, param_units=True), "table")
    rrefs = [s.use(F(R, d), "dim") for d in rowdims]
    c = s.use(F(R, coldim), "dim")
    s.measures(measures)
    s.rows = " / ".join(rrefs + [s.mn])
    s.cols = c
    s.cols_total = True
    s.encodings.append(("text", s.mv))
    s.nonnull(nonnull)
    s.exclude_null(c, F(R, coldim))
    for i, w in enumerate(widths or []):
        s.widths.append((rrefs[i], w))
    s.cell_align = "right"
    s.fit = "fit-width"
    return s


def story(R, name, caption, note=None):
    """The chapter headline: two or three short computed lines (one long line gets cut off)."""
    s = Sheet(R, name, "", "story")
    caps = sorted(c for c in R.by_caption if c.startswith(caption + " ") and c[len(caption) + 1:].isdigit())
    lines = []
    for cap in caps:
        ref = s.use(cap, "usrn")
        s.encodings.append(("text", ref))
        lines.append(run(f"<{ref}>", C_DARK, 10))
    if note:            # fixed text needs no calculation
        lines.append(run(note, C_DARK, 10))
    s.label = run("\u00c6\n", C_DARK, 4).join(lines)
    s.bg = "#eaf4ee"
    s.cell_align = "left"
    s.mark_color = C_DARK
    return s


def kpi_set(R, prefix):
    """The standard KPI cards (two values max per card)."""
    return [
        kpi(R, f"{prefix}Sales", "Sales", "Sales", ("", "Growth text")),
        kpi(R, f"{prefix}Achievement", "Monthly target achieved", "Monthly Ach %", ("Monthly target  ", "Monthly target")),
        kpi(R, f"{prefix}Yearly target", "Yearly target", "Yearly target", ("Achieved  ", "Yearly Ach %")),
        kpi(R, f"{prefix}Run-rate", "Monthly run-rate", "Current run-rate / month", ("Required  ", "Required run-rate / month")),
        kpi(R, f"{prefix}Order book", "Order book (OB)", "OB", ("", "OB growth text")),
        kpi(R, f"{prefix}Backlog", "Open backlog", "Backlog", ("Older than 50 days  ", "Backlog over 50 days")),
        kpi(R, f"{prefix}Partners", "Active partners", "Active partners", ("New / revival  ", "New / revival partners")),
        kpi(R, f"{prefix}Data", "Data as of", "Data as of", ("Period  ", "Period text")),
    ]


# ---------------------------------------------------------------- the sheets
def build_sheets(R):
    """Returns {name: Sheet}. scope: P (primary sales; region filter), PA (partner-level, also partner
    targets; region / branch / type filter on the partner), V (pavilion partners), Q (QTOD partner
    targets), S (SSP targets), PV (published pavilion tables), nav."""
    S = {}

    def add(s, scope="P", rls=None):
        s.scope, s.rls = scope, rls
        S[s.name] = s
        return s

    for label, dash in NAV:
        add(nav(R, dash, label), "nav")
    for label in ("‹ Back", "Next ›"):
        s = add(nav(R, label, label), "nav")
        s.nav_target = None             # set per dashboard (previous / next chapter)

    # ---- KPI cards
    for k in kpi_set(R, "K | "):
        add(k)
    add(kpi(R, "K | Wiser", "Wiser sales", "Wiser sales", ("Achieved  ", "Wiser Ach %")))
    add(kpi(R, "K | Zeta", "Miluz Zeta sales", "Zeta sales", ("Achieved  ", "Zeta Ach %")))
    add(kpi(R, "K | Quantity", "Quantity sold", "Quantity", ("Growth  ", "Quantity growth %"), units=False))
    add(kpi(R, "K | Per partner", "Sales per active partner", "Sales per partner", ("New / revival partner sales  ", "New partner sales")))
    add(kpi(R, "K | New partners", "New / revival partners", "New / revival partners", ("Their sales  ", "New partner sales"), units=False))
    add(kpi(R, "K | Lapsed", "Lapsed partners", "Lapsed partners", ("Active last year  ", "Active partners PY"), units=False))
    add(kpi(R, "K | Sales to OB", "Sales to OB", "Sales to OB %", ("Sales  ", "Sales")))
    add(kpi(R, "K | SSP sales", "Sales through SSPs", "SSP sales", ("Growth  ", "SSP growth %")))
    add(kpi(R, "K | SSP target", "SSP target", "SSP target", ("Achieved  ", "SSP Ach %")))
    add(kpi(R, "K | Cities", "Cities reached", "Active cities", ("New this year  ", "New cities"), units=False))
    add(kpi(R, "K | Partner FY target", "Partner targets (FY)", "Partner FY target", ("Achieved  ", "Partner FY Ach %")), "PA")
    add(kpi(R, "K | Pavilions", "Active pavilion partners", "Pavilions", ("Target  ", "Pavilion target"), units=False), "V")
    add(kpi(R, "K | Pavilion sales", "Pavilion partners' primary purchases", "Sales", ("Achieved  ", "Pavilion Ach %")), "V")

    # ---- story headlines
    notes = {"All Partners": "Type a name or code in Partner search to find one; click a column header to sort."}
    for dash, cap in [("Overview", "Overview"), ("Regions", "Regions"), ("Branches", "Branches"), ("Products", "Products"),
                      ("Partners", "Partners"), ("All Partners", "All partners"), ("Targets", "Targets"), ("SSP & Reach", "SSP"),
                      ("MTD", "MTD"), ("Order Book", "Orders")]:
        add(story(R, f"H | {dash}", f"Story | {cap}", notes.get(dash)), "PA" if dash in ("Partners", "All Partners", "Targets") else "P")
    add(story(R, "H | Pavilion", "Story | Pavilion",
              "ROs buy through distributors, so their sales show only in EazyDMS secondary data."), "V")

    # ---- 1 Overview
    add(trend(R, "O | Three years", "Monthly sales - this year and the two years before", ["Sales", "Sales PY", "Sales 2 years ago"],
              "Green = this year, grey = last year, light = two years ago", nonnull="Has sales (agg)"))
    add(grouped_bars(R, "O | Quarters", "Quarter-wise sales - this year vs last year", "Quarter", ["Sales", "Sales PY"]))
    add(hbar(R, "O | Region share", "Share of sales by region", "Region", "Share of sales %",
             "Colour: green = growing, red = declining vs last year", scale=("Growth %", GROWTH_SCALE)))
    add(hbar(R, "O | Partner type", "Sales by partner type (RD / KRO)", "RD/KRO", "Sales", scale=("Growth %", GROWTH_SCALE),
             exclude_null=True))

    # ---- 2 Regions
    add(table(R, "R | Scorecard", "Region scorecard", ["Region"],
              ["Sales", "Sales PY", "Growth %", "Share of sales %", "Monthly target", "Monthly Ach %", "Yearly target", "Yearly Ach %",
               "OB", "Backlog", "Active partners", "New / revival partners"], widths=[90]))
    add(lines_by(R, "R | Trend", "Monthly sales by region", "Month", "Region", "Sales"))
    add(heat(R, "R | Month heat", "Achievement by region and month", "Region", "Month", "Monthly Ach %", ACH_SCALE,
             "Red < 75%, amber 75-100%, green above 100%"))
    add(hbar(R, "R | Branches", "Branch sales in the selection", "Branch", "Sales",
             "Colour: green = growing, red = declining vs last year", scale=("Growth %", GROWTH_SCALE)))
    add(table(R, "R | Focus", "Focus families by region - Wiser and Miluz Zeta", ["Region"],
              ["Wiser sales", "Wiser target", "Wiser Ach %", "Zeta sales", "Zeta target", "Zeta Ach %"], widths=[90]))

    # ---- 3 Branches
    add(table(R, "B | Scorecard", "Branch scorecard", ["Region", "Branch"],
              ["Sales", "Sales PY", "Growth %", "Monthly target", "Monthly Ach %", "Gap to monthly target", "Yearly target", "Yearly Ach %",
               "OB", "Backlog", "Active partners"], widths=[70, 90]))
    add(table(R, "B | Focus families", "Focus families by branch - Wiser and Miluz Zeta", ["Region", "Branch"],
              ["Wiser sales", "Wiser target", "Wiser Ach %", "Zeta sales", "Zeta target", "Zeta Ach %"], widths=[70, 90]))
    add(table(R, "B | ASM table", "ASM (sales manager) performance", ["Region", "ASM"],
              ["Sales", "Sales PY", "Growth %", "Share of sales %", "Active partners", "New / revival partners", "Sales per partner"],
              widths=[70, 180], nonnull="Has sales (agg)", sort_by="Sales"))

    # ---- 4 Products
    add(hbar(R, "P | Family bars", "Sales by family", "Family", "Sales", "Top 15; green = growing, red = declining vs last year",
             scale=("Growth %", GROWTH_SCALE), top=(15, "Sales CY (row)")))
    add(table(R, "P | Family table", "Family performance - value and quantity", ["Family"],
              ["Sales", "Sales PY", "Growth %", "Share of sales %", "Quantity", "Quantity growth %"],
              sort_by="Sales", widths=[160], nonnull="Has sales (agg)"))
    add(hbar(R, "P | Activity", "Sales by activity (Helios code)", "Helios Code", "Sales", "HDCFD = FD, HDCWD = WD",
             scale=("Growth %", GROWTH_SCALE), exclude_null=True))
    add(hbar(R, "P | New products", "New products - Wiser, Miluz Zeta, Unica Pure Plus, UnicaX", "New product category", "Share of sales %",
             "Share of sales; colour = growth", scale=("Growth %", GROWTH_SCALE)))
    add(hbar(R, "P | Classification", "Sales by product classification", "Classification (Sheet1)", "Sales",
             scale=("Growth %", GROWTH_SCALE), exclude_null=True))
    add(table(R, "P | Top SKUs", "Top 25 materials", ["MaterialDescription"],
              ["Sales", "Sales PY", "Growth %", "Share of sales %", "Quantity"], sort_by="Sales",
              top=(25, "Sales CY (row)"), widths=[320], nonnull="Has sales (agg)", total=False))

    # ---- 5 Partners
    add(hbar(R, "Pa | Ach bands", "Partners by achievement of their target", "Partner ach band", "Partners in band",
             "Partner targets (gross) from the Partner Targets sheet", nonnull=None, exclude_null=True, sort=False), "PA", "P")
    add(hbar(R, "Pa | Slabs", "Partners by sales slab", "Partner slab", "Partners", "Partner sales in the selected period",
             nonnull="Partners", exclude_null=True, sort=False), "P", "P")
    add(hbar(R, "Pa | Tier", "Sales by tier", "Tier Category", "Sales", scale=("Growth %", GROWTH_SCALE), exclude_null=True), "P", "P")
    add(hbar(R, "Pa | New partners", "New / revival partners buying each month", "Month", "New / revival partners",
             "First sale this year (no sales in the previous two years)", nonnull="New / revival partners",
             exclude_null=True, sort=False, vertical=True), "P", "P")
    add(hbar(R, "Pa | Onboarded", "Partners by onboarded year", "Onboarded Year (Partner Master Data)", "Onboarded partners",
             "From the Partner Master", nonnull=None, exclude_null=True, sort=False, vertical=True), "P", "P")
    add(table(R, "Pa | Top partners", "Top 10 partners", ["Partner", "Branch"],
              ["Sales", "Sales PY", "Growth %", "Share of sales %"], sort_by="Sales", top=(10, "Sales CY (row)"),
              widths=[260, 70], total=False, nonnull="Has sales (agg)"), "P", "P")
    add(table(R, "Pa | At-risk partners", "Partners to watch - largest last-year partners now declining", ["Partner", "Branch", "ASM"],
              ["Sales PY", "Sales", "Growth", "Growth %", "Days since last sale"], sort_by=("Growth", "ASC"),
              top=(25, "Sales PY (row)"), widths=[280, 70, 170], total=False, nonnull="Has sales (agg)",
              scale=("Growth %", GROWTH_SCALE)), "P", "P")

    # ---- 6 All partners
    s = add(table(R, "AP | All partners", "All partners", ["Region", "Branch", "ASM", "Partner", "City (Geography mapping)", "RD/KRO", "Partner status"],
                  ["Sales", "Sales PY", "Growth %", "Partner FY target", "Partner FY Ach %", "OB", "Backlog", "Days since last sale"],
                  sub="Partner target = FY target from the Partner Targets sheet (gross; Net = gross x 0.83)",
                  widths=[52, 62, 120, 230, 85, 42, 85], sort_by="Sales", sort_dims=[3]), "PA", "P")
    s.bool_filter("Partner search match")
    s = add(matrix(R, "AP | By month", "Partner sales month by month", ["Branch", "Partner"], "Month", "Sales",
                   widths=[62, 260]), "PA", "P")
    s.bool_filter("Partner search match")

    # ---- 7 Targets
    add(table(R, "T | Monthly", "Monthly target vs sales and achievement", ["Month"],
              ["Monthly target", "Sales", "Monthly Ach %", "Gap to monthly target"],
              sub="Monthly targets from the Monthly Targets sheet (net; gross = net / 0.83)", widths=[80], nonnull="Has sales (agg)",
              scale=("Monthly Ach %", ACH_SCALE)))
    add(trend(R, "T | Monthly chart", "Monthly sales vs monthly target", ["Sales", "Monthly target"],
              "Green = sales, amber = monthly target", nonnull="Has sales (agg)"))
    add(pivot(R, "T | Branch x month", "Branch monthly target vs sales and achievement", ["Region", "Branch"], "Month",
              ["Monthly target", "Sales", "Monthly Ach %"], sub="For each branch: monthly target, sales and achievement by month",
              widths=[60, 80]))
    add(table(R, "T | Partner branches", "Partner targets by branch (QTOD / ATOD)", ["Target partner branch"],
              ["Partner target", "Partner achieved", "Partner Ach %", "Partner balance"],
              sub="View selected in 'Target view' (annual 100% / 108%, quarter, Wiser, Miluz Zeta)",
              nonnull="Has partner target", widths=[90], scale=("Partner Ach %", ACH_SCALE)), "Q", "Q")
    add(table(R, "T | Partners", "Partner target vs achievement (QTOD / ATOD)", ["Target partner branch", "Target partner"],
              ["Partner target", "Partner achieved", "Partner Ach %", "Partner balance"],
              sub="Partner targets are gross; Net = gross x 0.83",
              nonnull="Has partner target", widths=[80, 320], scale=("Partner Ach %", ACH_SCALE),
              sort_by="Partner target"), "Q", "Q")

    # ---- 8 SSP & reach
    add(table(R, "S | SSP table", "SSP target vs achievement", ["SSP", "Branch (SSA Targets)"],
              ["SSP target", "SSP sales", "SSP Ach %", "SSP sales PY", "SSP growth %"],
              sub="Sales of partners mapped to the SSP; SSP targets are used as they are",
              nonnull="SSP target", widths=[200, 90], scale=("SSP Ach %", ACH_SCALE), total=False), "S")
    add(table(R, "S | With vs without", "With vs without SSP", ["SSA group"],
              ["Sales", "Sales PY", "Growth %", "Share of sales %", "Active partners", "Sales per partner"], widths=[120]))
    add(trend(R, "S | Trend", "Monthly sales through SSPs - this year vs last year", ["SSP sales", "SSP sales PY"]))
    add(table(R, "S | Reach", "Geographic reach by branch", ["Region", "Branch"],
              ["Active cities", "New cities", "Active partners", "New / revival partners", "Sales"],
              sub="New city = first ever sale in that city this year", widths=[70, 90], nonnull="Has sales (agg)"))
    s = add(table(R, "S | Partners", "Top 25 SSP partners", ["SSA", "Partner"],
                  ["Sales", "Sales PY", "Growth %"], sort_by="Sales", top=(25, "SSA sales CY (row)"),
                  widths=[100, 200], total=False, nonnull="Has sales (agg)"), "P", "P")
    s.member_filter("SSA group", "With SSP")

    # ---- 9 Pavilion
    add(table(R, "V | Branches", "Pavilion partners by branch - primary purchases vs target",
              ["Pavilion region", "Pavilion branch", "KRO/RO"],
              ["Pavilions", "Pavilion target", "Sales", "Sales PY", "Growth %", "Pavilion Ach %"],
              sub="Active partners in the Pavilion Sales file, matched to primary sales by GSTIN; target as in the file",
              nonnull="Has pavilion", widths=[70, 80, 60]), "V")
    pf = lambda n: R.fields[f"{PDS}|{n}"]
    MS, PT, RA, SC, TT, WS = ("(Minimum!Slab!Achievement)", "(Pavilion!Type!Summary)", "(Retailer!App!TnC!Acceptance)",
                              "(Sales!Classification!Summary)", "(Total!Target!Achievement)", "(Wiser!Slabs!Unpivoted)")
    add(table(R, "PV | Type", "Pavilions by type", [pf("Pavilion Type"), pf(f"Sum of YTD Slab 2025 {PT}"), pf("Sum of YTD Slab 2026")],
              [(pf("Count of 2026 Classification"), "sum")], total=False, nonnull=None, widths=[110, 120, 120], units=False,
              sort_by=("Type row order", "ASC"), sort_dims=[0]), "PV")
    add(table(R, "PV | Classification", "Pavilions by classification",
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
    add(table(R, "PV | Sales slabs", "Sales slabs by ASM - 2025 vs 2026",
              [pf("Zone/Region"), pf("Sales Manager / Entity"), pf("Branch"), pf("2025 Sales Slab"), pf("2026 Sales Slab")],
              [(pf("2025-26 Growth %"), "sum")], total=False, nonnull=None, widths=[80, 200, 70, 120, 120], units=False,
              sort_by=("Slab row order", "ASC"), sort_dims=[0, 1]), "PV")
    s = add(table(R, "PV | Wiser slabs", "Wiser slab achievement by ASM",
                  [pf(f"Zone/Region {WS}"), pf(f"Sales Manager / Entity {WS}"), pf(f"Branch {WS}")],
                  [(pf("No. of KRO"), "sum"), (pf("No. of RO"), "sum"), (pf("No of KRO > 1 lac"), "sum"),
                   (pf("No. of RO > 25K"), "sum"), "Wiser qualifier %"], nonnull=None, widths=[90, 200, 70], units=False), "PV")
    s.member_filter(pf(f"Row Type {WS}"), "ASM")

    # ---- 10 MTD
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
    add(table(R, "M | ASM table", "MTD ASM performance - top performers first", ["Region", "ASM"],
              ["MTD sales", "MTD sales LY", "MTD growth %", "MTD OB", "MTD active partners"],
              widths=[70, 180], nonnull="Has MTD activity (agg)", sort_by="MTD sales"))
    add(hbar(R, "M | Family", "MTD sales by family", "Family", "MTD sales", top=(15, "Sales MTD (row)"),
             nonnull="Has MTD activity (agg)"))
    add(table(R, "M | Partners", "MTD top 20 partners", ["Partner", "Branch"], ["MTD sales", "MTD OB"],
              sort_by="MTD sales", top=(20, "Sales MTD (row)"), widths=[260, 80], total=False,
              nonnull="Has MTD activity (agg)"), "P", "P")

    # ---- 11 Orders
    add(trend(R, "O | OB trend", "Monthly order book - this year vs last year", ["OB", "OB PY"]))
    add(hbar(R, "O | Backlog ageing", "Open backlog by age", "Backlog ageing", "Backlog", nonnull="Backlog",
             exclude_null=True, sort=False, color=C_TARGET))
    add(table(R, "O | OB by branch", "Order book and backlog by branch", ["Region", "Branch"],
              ["OB", "OB PY", "OB growth %", "Sales", "Sales to OB %", "Backlog", "Backlog over 50 days"], widths=[70, 90]))
    add(table(R, "O | Backlog partners", "Top 20 partners by open backlog", ["Partner", "Branch"],
              ["Backlog", "Backlog over 50 days", "Sales"], sort_by="Backlog", top=(20, "Backlog (row)"),
              widths=[300, 90], total=False, nonnull="Backlog"), "P", "P")
    add(table(R, "O | Open orders", "Open orders (backlog) - every line",
              ["Order", "SalesOrderItem", "Order date", "Backlog ageing", "Partner", "Branch", "ASM", "MaterialDescription"],
              ["Backlog", "Backlog age (max days)"], nonnull="Backlog", sort_by="Backlog", sort_dims=[0],
              widths=[70, 40, 75, 95, 210, 55, 110, 230]), "P", "P")

    # ---- 12 Explore
    s = add(table(R, "X | Explore table", "Explore", ["Explore by"],
                  ["Sales", "Sales PY", "Growth %", "Share of sales %", "OB", "Backlog", "Active partners", "Quantity"],
                  sort_by="Sales", widths=[300]), "P", "P")
    s.title = title_runs("Explore by <[Parameters].[p Dim]>")
    add(hbar(R, "X | Explore bars", "Top 30 by sales", "Explore by", "Sales", scale=("Growth %", GROWTH_SCALE),
             top=(30, "Sales CY (row)")), "P", "P")

    # ---- filters: region selection, partner scope, pavilion scope, row-level security
    for name, s in S.items():
        if s.scope in ("P", "S"):
            s.bool_filter("Region selected", context=True)
        elif s.scope == "PA":
            s.bool_filter("Partner in scope", context=True)
        elif s.scope == "V":
            s.bool_filter("Pavilion in scope", context=True)
        elif s.scope == "Q":
            s.bool_filter("Target partner region selected", context=True)
        if s.rls == "P":
            s.bool_filter("Branch access (sales)", context=True)
        elif s.rls == "Q":
            s.bool_filter("Branch access (partner targets)", context=True)
    return S


# ---------------------------------------------------------------- dashboards
BASE_PARAMS = [("p Period", "compact"), ("p Start", "datetime"), ("p End", "datetime"), ("p Measure", "compact"),
               ("p Units", "compact"), ("p Region", "compact"), ("p Branch", "compact"), ("p Family", "compact"),
               ("p Activity", "compact"), ("p Channel", "compact")]
NO_PERIOD = [p for p in BASE_PARAMS if p[0] not in ("p Period", "p Start", "p End")]


def params_for(dash):
    if dash == "MTD":
        return NO_PERIOD
    if dash == "Pavilion":
        return [p for p in BASE_PARAMS if p[0] in ("p Period", "p Start", "p End", "p Measure", "p Units", "p Region", "p Branch")]
    if dash == "Targets":
        return BASE_PARAMS + [("p View", "compact")]
    if dash == "All Partners":
        return [("p Search", "type_in")] + BASE_PARAMS
    if dash == "Explore":
        return BASE_PARAMS + [("p Dim", "compact")]
    return BASE_PARAMS


SIDE_NOTE = ("Region / Branch apply to every chapter, so each page can be read for one region or one branch.\n"
             "Monthly targets come from the Monthly Targets sheet as they are (full months); yearly targets from the Branch Targets sheet.\n"
             "Gross / Net: monthly and yearly targets are net (gross = net / 0.83); partner targets are gross (net = gross x 0.83).\n"
             "Growth compares the same dates last year.\n"
             "Activity: HDCFD = FD, HDCWD = WD.")

X0, CW = 230, 1150          # content area
K4 = lambda names, y: [(n, X0 + i * 290, y, 280, 110) for i, n in enumerate(names)]
H = lambda dash, y=104: [(f"H | {dash}", X0, y, CW, 64)]
KPI1, KPI2 = 176, 294

LAYOUT = {
    "Overview": (1040, H("Overview")
                 + K4(["K | Sales", "K | Achievement", "K | Yearly target", "K | Run-rate"], KPI1)
                 + K4(["K | Order book", "K | Backlog", "K | Partners", "K | Data"], KPI2)
                 + [("O | Three years", X0, 412, CW, 300),
                    ("O | Quarters", X0, 722, 370, 300), ("O | Region share", 610, 722, 370, 300),
                    ("O | Partner type", 990, 722, 390, 300)]),
    "Regions": (1500, H("Regions")
                + [("R | Scorecard", X0, 176, CW, 230), ("R | Trend", X0, 416, CW, 300),
                   ("R | Month heat", X0, 726, CW, 220),
                   ("R | Branches", X0, 956, 560, 524), ("R | Focus", 800, 956, 580, 524)]),
    "Branches": (1716, H("Branches")
                 + [("B | Scorecard", X0, 176, CW, 560), ("B | Focus families", X0, 746, CW, 420),
                    ("B | ASM table", X0, 1176, CW, 520)]),
    "Products": (2030, H("Products")
                 + K4(["K | Sales", "K | Wiser", "K | Zeta", "K | Quantity"], KPI1)
                 + [("P | Family table", X0, 294, CW, 420),
                    ("P | Family bars", X0, 724, 560, 420), ("P | Classification", 800, 724, 580, 420),
                    ("P | Activity", X0, 1154, 560, 300), ("P | New products", 800, 1154, 580, 300),
                    ("P | Top SKUs", X0, 1464, CW, 546)]),
    "Partners": (1600, H("Partners")
                 + K4(["K | Partners", "K | New partners", "K | Lapsed", "K | Per partner"], KPI1)
                 + [("Pa | Ach bands", X0, 294, 370, 280), ("Pa | Slabs", 610, 294, 370, 280), ("Pa | Tier", 990, 294, 390, 280),
                    ("Pa | New partners", X0, 584, 760, 280), ("Pa | Onboarded", 1000, 584, 380, 280),
                    ("Pa | Top partners", X0, 874, CW, 300), ("Pa | At-risk partners", X0, 1184, CW, 400)]),
    "All Partners": (2300, H("All Partners")
                     + [("AP | All partners", X0, 176, CW, 1300), ("AP | By month", X0, 1486, CW, 800)]),
    "Targets": (2190, H("Targets")
                + K4(["K | Achievement", "K | Yearly target", "K | Partner FY target", "K | Run-rate"], KPI1)
                + [("T | Monthly", X0, 294, 560, 400), ("T | Monthly chart", 800, 294, 580, 400),
                   ("T | Branch x month", X0, 704, CW, 760),
                   ("T | Partner branches", X0, 1474, 400, 700), ("T | Partners", 640, 1474, 740, 700)]),
    "SSP & Reach": (1490, H("SSP & Reach")
                    + K4(["K | SSP sales", "K | SSP target", "K | Cities", "K | Data"], KPI1)
                    + [("S | SSP table", X0, 294, CW, 300),
                       ("S | With vs without", X0, 604, 560, 170), ("S | Trend", 800, 604, 580, 300),
                       ("S | Reach", X0, 784, 560, 690), ("S | Partners", 800, 914, 580, 560)]),
    "Pavilion": (2270, H("Pavilion")
                 + K4(["K | Pavilions", "K | Pavilion sales"], KPI1)
                 + [("V | Branches", X0, 294, CW, 400),
                    ("PV | Type", X0, 704, 560, 170), ("PV | Classification", 800, 704, 580, 170),
                    ("PV | Min slab", X0, 884, 560, 260), ("PV | Total target", 800, 884, 580, 260),
                    ("PV | App T&C", X0, 1154, 560, 260), ("PV | Sales slabs", X0, 1424, CW, 420),
                    ("PV | Wiser slabs", X0, 1854, CW, 400)]),
    "MTD": (1860, H("MTD")
            + K4(["M | Sales", "M | Achievement", "M | OB", "M | Yesterday"], KPI1)
            + [("M | Daily trend", X0, 294, CW, 300), ("M | Branch table", X0, 604, CW, 520),
               ("M | Family", X0, 1134, 400, 710), ("M | ASM table", 640, 1134, 740, 350),
               ("M | Partners", 640, 1494, 740, 350)]),
    "Order Book": (2110, H("Order Book")
                   + K4(["K | Order book", "K | Sales to OB", "K | Backlog", "K | Data"], KPI1)
                   + [("O | OB trend", X0, 294, 760, 320), ("O | Backlog ageing", 1000, 294, 380, 320),
                      ("O | OB by branch", X0, 624, CW, 520), ("O | Backlog partners", X0, 1154, CW, 330),
                      ("O | Open orders", X0, 1494, CW, 600)]),
    "Explore": (1500, [("X | Explore table", X0, 104, CW, 700), ("X | Explore bars", X0, 814, CW, 670)]),
}
