"""Sheets and dashboards of the Retail dashboard (v6).

Every analysis appears once. Region, Branch, ASM, Partner, Sold-to code, Family,
Activity, Partner type and Tier are dashboard filters shared by every sheet.
Dashboards are built from layout containers (rows of sheets), so they can be
rearranged in Tableau.
"""
from core import C_DARK, C_MID, C_TARGET, PDS
from model import MONTHS, MTD_VIEWS, TARGET_VIEWS
from sheets import (ACH_SCALE, GROWTH_SCALE, Sheet, column_scale, kpi, nav, run, title_runs, is_value)

# the story: (button label, dashboard name, one-line purpose)
CHAPTERS = [
    ("Overview", "Overview", "Where do we stand? Sales vs last year and the yearly target, three-year trend, regions."),
    ("Regions", "Regions", "Region-wise performance. Click a region to focus the page on it."),
    ("Branches", "Branches", "Branch scorecard against yearly targets, focus families and ASMs. Click a branch to see its ASMs."),
    ("Products", "Products", "Families, activities (FD / WD ...), new products, classification and top materials."),
    ("Partners", "Partners", "Partner base: achievement bands, slabs, tiers, new / revival partners and partners to watch."),
    ("All partners", "All Partners", "Numbers for every partner, and every partner month by month."),
    ("New Ptnr/Geo", "New Partner - New Geo", "Policy categories (Old, Handholding, New), new partners by category, region, tier and type, and new cities."),
    ("Monthly", "Monthly Targets", "Monthly targets vs sales and achievement."),
    ("Yearly", "Yearly Targets", "Yearly branch targets vs sales, balance and run-rate."),
    ("Partner tgt", "Partner Targets", "Partner target vs achievement: annual, each quarter, Wiser and Zeta (pick in 'Select a View')."),
    ("SSP", "SSP", "SSP performance, with vs without SSP, and new partners / new geo brought in by SSPs."),
    ("Pavilion", "Pavilion", "Pavilion partners, KRO primary purchases, slabs and qualification."),
    ("MTD", "MTD", "This month so far: target, sales, order book and the views of the original MTD dashboard."),
    ("Orders", "Order Book", "Order book, conversion to sales, open backlog and every open order."),
    ("Explore", "Explore", "Any dimension against all measures."),
]
NAV = [(label, dash) for label, dash, _ in CHAPTERS]
DASHBOARDS = [dash for _, dash, _ in CHAPTERS]
NOTES = {dash: note for _, dash, note in CHAPTERS}
# click a row in the source sheet -> filter the rest of the dashboard
DRILL = {"Regions": ["R | Scorecard"], "Branches": ["B | Scorecard"], "Products": ["P | Family table"]}

# shared dashboard filters: (field, filter group, context). Context filters also limit the LOD calculations.
FILTERS = [("Region", 1, True), ("Branch", 2, True), ("ASM", 3, True), ("Partner", 4, True),
           ("SoldToPartyID", 5, True), ("Family", 6, False), ("Helios Code", 7, False),
           ("RD/KRO", 8, True), ("Tier Category", 9, True), ("Policy category", 10, False), ("Partner growth band", 11, False)]
PAV_FILTERS = [("Zone/Region", 21, True), ("Branch (Pavilion Sales)", 22, True), ("KRO/RO", 23, True)]


# ---------------------------------------------------------------- generic builders
# expected number of rows of a dimension, so small tables get the height they need (no scrolling)
EST = {"Region": 5, "Branch": 27, "Policy category": 5, "Policy group": 3, "FY PY sales bucket": 4, "Month": 12, "Quarter": 4,
       "Tier Category": 3, "RD/KRO": 2, "SSA group": 2, "Partner status": 2, "Helios Code": 4, "Partner growth band": 7,
       "Top ASM MTD": 5, "SSA": 6, "SSP": 10, "Branch (SSA Targets)": 10, "Backlog ageing": 4, "Zone/Region": 6, "Branch (Pavilion Sales)": 27}
EST_PART = {"Zone/Region": 7, "Pavilion Type": 3, "Classification": 5}


def est_rows(dims):
    """Rows a table with these row dimensions shows (the finest one decides), or None if it can be long."""
    n = 0
    for d in dims:
        name = d if isinstance(d, str) else d.name
        k = EST.get(name) or next((v for part, v in EST_PART.items() if part in name), None)
        if k is None:
            return None
        n = max(n, k)
    return n


def F(R, x):
    return x if not isinstance(x, str) else (R[x] if x in R.by_caption else R.fields[x])


def colour_columns(R, s, measures):
    """Font colour per column: growth red / green, achievement red / amber / green, the rest plain."""
    s.encodings.append(("color", s.mv))
    for m in measures:
        f = F(R, m[0]) if isinstance(m, tuple) else R[m]
        kind = m[1] if isinstance(m, tuple) else "usr"
        ref = s.use(f, kind)
        s.scales.append((ref, column_scale(f.caption or f.name)))


def table(R, name, title, dims, measures, sub=None, total=True, sort_by=None, top=None, nonnull="Has activity (agg)",
          widths=None, units=None, sort_dims=None, col_width=None, fit="fit-width"):
    caps = [m if isinstance(m, str) else None for m in measures]
    if units is None:
        units = any(c and is_value(R, c) for c in caps)
    s = Sheet(R, name, title_runs(title, sub, param_units=units), "table")
    drefs = [s.use(F(R, d), "dim") for d in dims]
    s.measures(measures)
    s.rows = " / ".join(drefs) if drefs else s.mn
    s.rows_total = total and bool(drefs)
    s.cols = s.mn if drefs else ""
    s.encodings.append(("text", s.mv))
    colour_columns(R, s, measures)
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
    if col_width:
        s.widths.append((s.mn, col_width))
    s.cell_align = "right"
    s.fit = fit
    s.est_rows = (est_rows(dims) if dims else len(measures))
    if s.est_rows is not None and s.rows_total:
        s.est_rows += 1
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


def grouped_bars(R, name, title, dim, measures, sub=None, nonnull="Has sales (agg)", horizontal=False):
    s = Sheet(R, name, title_runs(title, sub, param_units=is_value(R, measures[0])), "chart")
    f = F(R, dim)
    dref = s.use(f, "dim")
    s.measures(measures)
    s.encodings.append(("color", s.mn))
    s.encodings.append(("text", s.mv))
    if horizontal:
        s.rows, s.cols = f"{dref} / {s.mn}", s.mv
        s.hide_axis.append((s.mv, "cols"))
    else:
        s.cols, s.rows = f"{dref} / {s.mn}", s.mv
        s.hide_axis.append((s.mv, "rows"))
    s.mark = "Bar"
    s.nonnull(nonnull)
    s.exclude_null(dref, f)
    return s


def lines_by(R, name, title, dim, colour_dim, measure, sub=None, nonnull="Has sales (agg)"):
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


def matrix(R, name, title, rowdims, coldim, measure, sub=None, nonnull="Has sales (agg)", widths=None, fit="fit-width"):
    s = Sheet(R, name, title_runs(title, sub, param_units=is_value(R, measure)), "table")
    rrefs = [s.use(F(R, d), "dim") for d in rowdims]
    c = s.use(F(R, coldim), "dim")
    v = s.use(measure, "usr")
    s.rows, s.cols = " / ".join(rrefs), c
    s.rows_total = True
    s.cols_total = True
    s.encodings.append(("text", v))
    if nonnull:
        s.nonnull(nonnull)
    s.exclude_null(c, F(R, coldim))
    for i, w in enumerate(widths or []):
        s.widths.append((rrefs[i], w))
    s.cell_align = "right"
    s.fit = fit
    s.est_rows = est_rows(rowdims)
    if s.est_rows is not None:
        s.est_rows += 1
    return s


def pivot(R, name, title, rowdims, coldim, measures, sub=None, nonnull="Has sales (agg)", widths=None):
    """Row dimensions and measures down the side, a dimension (e.g. month) across."""
    s = Sheet(R, name, title_runs(title, sub, param_units=True), "table")
    rrefs = [s.use(F(R, d), "dim") for d in rowdims]
    c = s.use(F(R, coldim), "dim")
    s.measures(measures)
    s.rows = " / ".join(rrefs + [s.mn])
    s.cols = c
    s.cols_total = True
    s.encodings.append(("text", s.mv))
    colour_columns(R, s, measures)
    s.nonnull(nonnull)
    s.exclude_null(c, F(R, coldim))
    for i, w in enumerate(widths or []):
        s.widths.append((rrefs[i], w))
    s.cell_align = "right"
    s.fit = "fit-width"
    return s


def story(R, name, caption, note=None):
    """The chapter headline: two or three short computed lines."""
    s = Sheet(R, name, "", "story")
    caps = sorted(c for c in R.by_caption if c.startswith(caption + " ") and c[len(caption) + 1:].isdigit())
    lines = []
    for cap in caps:
        ref = s.use(cap, "usrn")
        s.encodings.append(("text", ref))
        lines.append(run(f"<{ref}>", C_DARK, 10))
    if note:
        lines.append(run(note, C_DARK, 10))
    s.label = run("Æ\n", C_DARK, 4).join(lines)
    s.bg = "#eaf4ee"
    s.cell_align = "left"
    s.mark_color = C_DARK
    return s


# ---------------------------------------------------------------- the sheets
def build_sheets(R):
    """Returns {name: Sheet}. scope: P (retail data; shared filters), V (pavilion partners; pavilion filters),
    PV (published pavilion tables), nav."""
    S = {}

    def add(s, scope="P", rls=None):
        s.scope, s.rls = scope, rls
        S[s.name] = s
        return s

    for label, dash in NAV:
        add(nav(R, dash, label), "nav")
    for label in ("‹ Back", "Next ›"):
        s = add(nav(R, label, label), "nav")
        s.nav_target = None

    G = ["Sales", "Sales PY", "Growth %", "Growth (value)"]       # sales with growth

    # ---- KPI cards
    add(kpi(R, "K | Sales", "Sales", "Sales", ("", "Growth text")))
    add(kpi(R, "K | Yearly target", "Yearly target achieved", "Yearly Ach %", ("Yearly target  ", "Yearly target")))
    add(kpi(R, "K | Run-rate", "Monthly run-rate", "Current run-rate / month", ("Required  ", "Required run-rate (yearly)")))
    add(kpi(R, "K | Order book", "Order book (OB)", "OB", ("", "OB growth text")))
    add(kpi(R, "K | Backlog", "Open backlog", "Backlog", ("Ordered over 50 days ago  ", "Backlog over 50 days")))
    add(kpi(R, "K | Partners", "Active partners", "Active partners", ("New / revival  ", "New / revival partners")))
    add(kpi(R, "K | Data", "Data as of", "Data as of", ("Period  ", "Period text")))
    add(kpi(R, "K | Wiser", "Wiser sales", "Wiser sales", ("Growth  ", "Wiser growth %")))
    add(kpi(R, "K | Zeta", "Miluz Zeta sales", "Zeta sales", ("Growth  ", "Zeta growth %")))
    add(kpi(R, "K | Quantity", "Quantity sold", "Quantity", ("Growth  ", "Quantity growth %"), units=False))
    add(kpi(R, "K | Per partner", "Sales per active partner", "Sales per partner", ("Active last year  ", "Active partners PY")))
    add(kpi(R, "K | New partners", "New / revival partners", "New partners", ("Their sales  ", "New partner sales CY"), units=False))
    add(kpi(R, "K | Lapsed", "Lapsed partners", "Lapsed partners", ("Active last year  ", "Active partners PY"), units=False))
    add(kpi(R, "K | New share", "New partners' share of sales", "New partner sales %", ("Growth  ", "New partner growth %"), units=False))
    add(kpi(R, "K | NP partners", "New partners (policy)", "Policy new partners", ("Buying this period  ", "Policy new buying"), units=False))
    add(kpi(R, "K | NP sales", "New partners' sales", "Policy new sales", ("Share of sales  ", "Policy new share %")))
    add(kpi(R, "K | NP target", "New partner sales target achieved", "Policy new target Ach %", ("Target  ", "New partner sales target")))
    add(kpi(R, "K | New geo", "New geo cities", "New geo cities", ("Their sales  ", "New geo sales"), units=False))
    add(kpi(R, "K | Monthly target", "Monthly target achieved", "Monthly Ach %", ("Monthly target  ", "Monthly target")))
    add(kpi(R, "K | Monthly gap", "Gap to monthly target", "Gap to monthly target", ("Sales  ", "Sales")))
    add(kpi(R, "K | Yearly balance", "Balance to yearly target", "Balance to yearly target", ("Year gone  ", "Year gone %")))
    add(kpi(R, "K | Annual partner", "Annual target (QTOD)", "Annual target (100%)", ("Achieved  ", "Annual Ach %")))
    add(kpi(R, "K | Partner FY", "Partner target (FY)", "Partner FY target", ("Achieved  ", "Partner FY Ach %")))
    add(kpi(R, "K | Wiser target", "Wiser annual target (QTOD)", "Wiser annual target", ("Achieved  ", "Wiser annual Ach %")))
    add(kpi(R, "K | Zeta target", "Zeta annual target (QTOD)", "Zeta annual target", ("Achieved  ", "Zeta annual Ach %")))
    add(kpi(R, "K | SSP sales", "Sales through SSPs", "SSP sales", ("Growth  ", "SSP growth %")))
    add(kpi(R, "K | SSP target", "SSP target", "SSP target", ("Achieved  ", "SSP Ach %")))
    add(kpi(R, "K | Sales to OB", "Sales to OB", "Sales to OB %", ("Sales  ", "Sales")))
    add(kpi(R, "K | Total pavilions", "Total pavilions", "Total pavilions (file)", None, units=False), "PV")
    add(kpi(R, "K | Pavilion slab", "Sales slab 2025 > 2026", "Overall sales slab", ("Growth  ", "Overall slab growth %"), units=False), "PV")
    add(kpi(R, "K | KRO pavilions", "Active KRO / RO pavilions", "KRO pavilions", ("RO pavilions  ", "RO pavilions"), units=False), "V")
    s = add(kpi(R, "K | KRO purchases", "KRO primary purchases", "Sales", ("Growth  ", "Growth %")), "V")
    s.member_filter(R.fields["KRO/RO"], "KRO")

    # ---- story headlines
    notes = {"All Partners": "Use the Partner or Sold To Party ID filter on the left to find a partner; click a column header to sort."}
    for dash, cap in [("Overview", "Overview"), ("Regions", "Regions"), ("Branches", "Branches"), ("Products", "Products"),
                      ("Partners", "Partners"), ("All Partners", "All partners"), ("Monthly Targets", "Targets"),
                      ("SSP", "SSP"), ("MTD", "MTD"), ("Order Book", "Orders")]:
        add(story(R, f"H | {dash}", f"Story | {cap}", notes.get(dash)))
    add(story(R, "H | Pavilion", "Story | Pavilion", "Only KROs buy in primary; ROs buy through distributors."), "V")

    # ---- Overview
    add(trend(R, "O | Three years", "Monthly sales - this year and the two years before", ["Sales", "Sales PY", "Sales 2 years ago"],
              "Dark green = this year, grey = last year, light green = two years ago", nonnull="Has sales (agg)"))
    add(grouped_bars(R, "O | Quarters", "Quarter-wise sales - this year vs last year", "Quarter", ["Sales", "Sales PY"]))
    add(hbar(R, "O | Region share", "Share of sales by region", "Region", "Share of sales %",
             "Green = growing, red = declining vs last year", scale=("Growth %", GROWTH_SCALE)))
    add(hbar(R, "O | Partner type", "Sales by partner type (RD / KRO)", "RD/KRO", "Sales",
             "Green = growing, red = declining vs last year", scale=("Growth %", GROWTH_SCALE), exclude_null=True))

    # ---- Regions
    add(table(R, "R | Scorecard", "Region scorecard (yearly targets)", ["Region"],
              G + ["Share of sales %", "Yearly target", "Yearly Ach %", "Balance to yearly target",
                   "OB", "OB growth %", "Backlog", "Active partners", "New / revival partners"], widths=[80]))
    add(lines_by(R, "R | Trend", "Monthly sales by region", "Month", "Region", "Sales"))
    add(heat(R, "R | Growth heat", "Growth vs last year by region and month", "Region", "Month", "Growth %", GROWTH_SCALE,
             "Green = growing, red = declining"))
    add(hbar(R, "R | Branches", "Branch sales in the selection", "Branch", "Sales",
             "Green = growing, red = declining vs last year", scale=("Growth %", GROWTH_SCALE)))
    add(table(R, "R | Focus", "Focus families by region - Wiser and Miluz Zeta", ["Region"],
              ["Wiser sales", "Wiser sales PY", "Wiser growth %", "Zeta sales", "Zeta sales PY", "Zeta growth %"], widths=[80],
              col_width=72, fit=None))

    # ---- Branches
    add(table(R, "B | Scorecard", "Branch scorecard vs yearly targets", ["Region", "Branch"],
              G + ["Yearly target", "Yearly Ach %", "Balance to yearly target", "Required run-rate (yearly)",
                   "OB", "OB growth %", "Backlog", "Active partners"], widths=[60, 100], col_width=78, fit=None))
    add(table(R, "B | Focus families", "Focus families by branch - Wiser and Miluz Zeta", ["Region", "Branch"],
              ["Wiser sales", "Wiser sales PY", "Wiser growth %", "Zeta sales", "Zeta sales PY", "Zeta growth %"], widths=[60, 100]))
    add(table(R, "B | ASM table", "ASM (sales manager) performance", ["Region", "ASM"],
              G + ["Share of sales %", "Active partners", "New / revival partners", "Sales per partner"],
              widths=[60, 180], nonnull="Has sales (agg)", sort_by="Sales"))

    # ---- Products
    add(table(R, "P | Family table", "Family performance - value and quantity", ["Family"],
              G + ["Share of sales %", "Quantity", "Quantity PY", "Quantity growth %"],
              sort_by="Sales", widths=[170], nonnull="Has sales (agg)", col_width=90, fit=None))
    add(hbar(R, "P | Family bars", "Sales by family", "Family", "Sales", "Top 15; green = growing, red = declining vs last year",
             scale=("Growth %", GROWTH_SCALE), top=(15, "Sales CY (row)")))
    add(hbar(R, "P | Activity", "Sales by activity (Helios code)", "Helios Code", "Sales", "HDCFD = FD, HDCWD = WD; colour = growth",
             scale=("Growth %", GROWTH_SCALE), exclude_null=True))
    add(table(R, "P | New products", "New products - Wiser, Miluz Zeta, Unica Pure Plus, UnicaX", ["New product category"],
              G + ["Share of sales %"], widths=[120], nonnull="Has sales (agg)", col_width=72, fit=None))
    add(hbar(R, "P | Classification", "Sales by product classification", "Classification (Sheet1)", "Sales",
             "Colour = growth", scale=("Growth %", GROWTH_SCALE), exclude_null=True))
    add(table(R, "P | Top SKUs", "Top 25 materials", ["MaterialDescription"],
              G + ["Share of sales %", "Quantity"], sort_by="Sales",
              top=(25, "Sales CY (row)"), widths=[320], nonnull="Has sales (agg)", total=False))

    # ---- Partners
    add(hbar(R, "Pa | Ach bands", "Partners by achievement of their target", "Partner ach band", "Partners in band",
             "Partner targets are gross", nonnull=None, exclude_null=True, sort=False), "P", "P")
    add(hbar(R, "Pa | Slabs", "Partners by sales slab", "Partner slab", "Partners", "Partner sales in the selected period",
             nonnull="Partners", exclude_null=True, sort=False), "P", "P")
    add(hbar(R, "Pa | Tier", "Sales by tier", "Tier Category", "Sales", "Colour = growth", scale=("Growth %", GROWTH_SCALE),
             exclude_null=True), "P", "P")
    add(hbar(R, "Pa | New partners", "New / revival partners buying each month", "Month", "New partners",
             "New / revival = bought this year, nothing in the previous two years", nonnull="New partners",
             exclude_null=True, sort=False, vertical=True), "P", "P")
    add(hbar(R, "Pa | Onboarded", "Partners by onboarded year", "Onboarded Year (Partner Master Data)", "Onboarded partners",
             "From the Partner Master", nonnull=None, exclude_null=True, sort=False, vertical=True), "P", "P")
    add(table(R, "Pa | Top partners", "Top 10 partners", ["Partner", "Branch"], G + ["Share of sales %"],
              sort_by="Sales", top=(10, "Sales CY (row)"), widths=[260, 80], total=False, nonnull="Has sales (agg)"), "P", "P")
    s = add(table(R, "Pa | Declining", "Partners to watch - largest drops vs last year (declining partners only)",
                  ["Partner", "Branch", "ASM"], ["Sales PY", "Sales", "Growth (value)", "Growth %", "Days since last sale"],
                  sort_by=("Growth (value)", "ASC"), top=(25, "Decline (row)"), widths=[280, 80, 170], total=False,
                  nonnull="Has sales (agg)"), "P", "P")
    s.nonnull("Decline (agg)")

    # ---- All partners
    add(table(R, "AP | All partners", "All partners", ["Region", "Branch", "ASM", "Partner", "City (Geography mapping)", "RD/KRO",
                                                       "Tier Category", "Policy category"],
              ["FY PY sales", "Sales PY", "Sales", "Growth %", "Growth (value)", "Assigned target", "Policy Ach %",
               "Req. monthly run-rate", "OB", "Backlog", "Days since last sale"],
              sub="Assigned target = the partner's FY target (gross; Net = gross x 0.83); new partners (policy) get a flat 25 L",
              widths=[55, 75, 130, 250, 95, 45, 55, 130], col_width=72, sort_by="Sales", sort_dims=[3], fit=None), "P", "P")
    add(matrix(R, "AP | By month", "Partner sales month by month", ["Branch", "Partner"], "Month", "Sales",
               widths=[75, 260], fit=None), "P", "P")

    # ---- Policy tracker
    add(table(R, "PO | Summary", "Policy categories - partners and sales", ["Policy group", "Policy category"],
              ["Partners (count)", "FY PY sales", "Sales PY", "Sales", "Growth %", "Share of sales %"],
              sub="Old = bought before last year; Handholding = onboarded last year with over 5 L; New = new this year, Q4 last year or under 5 L",
              widths=[130, 150], nonnull="Has sales (agg)"))
    add(matrix(R, "PO | Buckets", "Handholding and new partners by last year's sales", ["FY PY sales bucket"], "Policy category",
               "Partners (count)", nonnull="Has sales (agg)", widths=[130]))
    S["PO | Buckets"].member_filter("Policy group", ["Handholding Partner", "New Partner"], context=False)

    # ---- New Partner - New Geo: new partners as the policy tracker groups them (this page only)
    NPSUB = "New partners = policy tracker's New Partner group: new this year, onboarded in Q4 last year, or under 5 L last year"
    s = add(table(R, "NG | Categories", "New partners by policy category",
                  ["Policy category"], ["Policy new partners", "Policy new buying", "Policy new sales PY", "Policy new sales",
                                        "Policy new growth %", "Policy new assigned target", "Policy new Ach %", "Policy new share %",
                                        "Policy new sales per partner"],
                  sub=NPSUB + "; assigned target flat 25 L (gross) each", widths=[170], nonnull="Policy new partners"))
    s.member_filter("Policy group", "New Partner", context=False)
    s.est_rows = 4                  # the three new-partner categories and the total
    add(table(R, "NG | Regions", "New partners by region - target vs actual", ["Region"],
              ["New partner count target", "Policy new buying", "Policy new count Ach %", "New partner sales target", "Policy new sales",
               "Policy new target Ach %", "Policy new sales PY", "Policy new growth %", "Policy new sales per partner"],
              sub="Region targets for new partners (count and gross value)", widths=[80], nonnull="Has sales (agg)"))
    s = add(matrix(R, "NG | Region x category", "New partners' sales by region and policy category", ["Region"], "Policy category",
                   "Policy new sales", nonnull="Policy new sales", widths=[80]))
    s.member_filter("Policy group", "New Partner", context=False)
    NT5 = ["Policy new partners", "Policy new buying", "Policy new sales PY", "Policy new sales", "Policy new growth %"]
    add(table(R, "NG | Tier", "New partners by tier", ["Tier Category"], NT5, widths=[90], nonnull="Policy new sales"))
    add(table(R, "NG | Type", "New partners by type", ["RD/KRO"], NT5, widths=[70], nonnull="Policy new sales"))
    add(matrix(R, "NG | Products", "New partners' sales by activity and region", ["Region"], "Helios Code", "Policy new sales",
               nonnull="Policy new sales", widths=[80]))
    add(table(R, "NG | Cities", "New geo cities - first sales this year (none in the previous two years)",
              ["Region", "Branch", "City (Geography mapping)"], ["New geo partners", "New geo sales"],
              widths=[70, 90, 160], nonnull="New geo sales", sort_by="New geo sales", sort_dims=[2]))

    # ---- Monthly targets
    add(table(R, "MT | Monthly", "Monthly target vs sales and achievement", ["Month"],
              ["Monthly target", "Sales", "Monthly Ach %", "Gap to monthly target", "Sales PY", "Growth %"],
              sub="Monthly targets are net; gross = net / 0.83", widths=[80], nonnull="Has sales (agg)"))
    add(trend(R, "MT | Chart", "Monthly sales vs monthly target", ["Sales", "Monthly target"],
              "Green = sales, amber = monthly target", nonnull="Has sales (agg)"))
    add(pivot(R, "MT | Branch x month", "Branch monthly target vs sales and achievement", ["Region", "Branch"], "Month",
              ["Monthly target", "Sales", "Monthly Ach %"], widths=[60, 100]))
    add(table(R, "MT | Focus", "Wiser and Miluz Zeta monthly targets by branch", ["Region", "Branch"],
              ["Wiser target", "Wiser sales", "Wiser Ach %", "Zeta target", "Zeta sales", "Zeta Ach %"], widths=[60, 100]))

    # ---- Yearly targets
    Y = ["Yearly target", "Sales", "Yearly Ach %", "Balance to yearly target", "Required run-rate (yearly)", "Sales PY", "Growth %"]
    add(table(R, "YT | Regions", "Yearly target by region", ["Region"], Y, widths=[80], nonnull="Has sales (agg)"))
    add(table(R, "YT | Branches", "Yearly target by branch", ["Region", "Branch"], Y, widths=[60, 100],
              nonnull="Has sales (agg)", col_width=96, fit=None))
    add(hbar(R, "YT | Ach bars", "Yearly target achieved by branch", "Branch", "Yearly Ach %",
             "Red < 75%, amber 75-100%, green above 100%", scale=("Yearly Ach %", ACH_SCALE), nonnull="Yearly Ach %"))

    # ---- Partner targets (QTOD / ATOD)
    PD = ["Sold-to Party (Partner Master Data)", "Patner (Partner Master Data)", "ASM", "City (Geography mapping)", "RD/KRO"]
    PW = [70, 185, 120, 90, 75]
    TV = {
        "Annual": ("Partner wise Annual Target vs Ach%",
                   ["Annual target (100%)", "Annual target (108%)", "Sales", "Annual Ach %", "Q1 sales", "Q2 sales", "Q3 sales",
                    "Q4 sales", "Sales PY", "Growth %"]),
        "Wiser": ("Partner wise Wiser Target vs Ach%",
                  ["Wiser annual target", "Wiser sales", "Wiser annual Ach %", "Wiser Q1 sales", "Wiser Q2 sales", "Wiser Q3 sales",
                   "Wiser Q4 sales", "Wiser sales PY", "Wiser growth %"]),
        "Zeta": ("Partner wise Zeta Target vs Ach%",
                 ["Zeta annual target", "Zeta sales", "Zeta annual Ach %", "Zeta Q1 sales", "Zeta Q2 sales", "Zeta Q3 sales",
                  "Zeta Q4 sales", "Zeta sales PY", "Zeta growth %"]),
    }
    for q in (1, 2, 3, 4):
        mons = MONTHS[(q - 1) * 3:q * 3]
        TV[f"Q{q}"] = (f"Partner wise Q{q} Target vs Ach% (QTOD)",
                       [f"Q{q} target", f"Q{q} sales", f"Q{q} Ach %", f"Q{q} balance"]
                       + [x for mo in mons for x in (f"{mo} sales", f"{mo} Ach %")])
    for key, _ in TARGET_VIEWS:
        title, ms = TV[key]
        s = add(table(R, f"PTV | {key}", title, PD, ms, sub="Partner targets are gross; Net = gross x 0.83",
                      nonnull="Has partner target", widths=PW, sort_by=ms[0], sort_dims=[1]), "P", "Q")
        s.bool_filter(f"Target view: {key}")

    # ---- SSP
    add(table(R, "S | SSP table", "SSP target vs achievement", ["SSP", "Branch (SSA Targets)"],
              ["SSP target", "SSP sales", "SSP Ach %", "SSP sales PY", "SSP growth %"],
              sub="Sales of partners mapped to the SSP",
              nonnull="SSP target", widths=[200, 90]))
    add(table(R, "S | With vs without", "With vs without SSP", ["SSA group"],
              G + ["Share of sales %", "Active partners", "Sales per partner"], widths=[100], col_width=62, fit=None))
    add(trend(R, "S | Trend", "Monthly sales through SSPs - this year vs last year", ["SSP sales", "SSP sales PY"]))
    s = add(table(R, "S | New by SSP", "New partners and new geo brought in by SSPs", ["SSA"],
                  ["Policy new partners", "Policy new buying", "Policy new sales", "New geo partners", "New geo cities", "New geo sales"],
                  sub="New partners as the policy tracker groups them; new geo = city with no sales in the previous two years",
                  widths=[200], nonnull="Policy new partners"))
    s.member_filter("SSA group", "With SSP", context=False)
    s = add(table(R, "S | New partners", "New partners onboarded through SSPs", ["SSA", "Partner", "Policy category", "Branch",
                                                                                "City (Geography mapping)", "Geo status"],
                  ["Policy new sales", "Sales PY", "Growth %"], widths=[150, 250, 150, 80, 120, 90],
                  sort_by="Policy new sales", sort_dims=[1], total=False, nonnull="Policy new sales"), "P", "P")
    s.member_filter("Policy group", "New Partner", context=False)
    s.member_filter("SSA group", "With SSP")

    # ---- Pavilion
    add(table(R, "V | Count", "Active pavilions by region and branch", ["Zone/Region", "Branch (Pavilion Sales)"],
              ["KRO pavilions", "RO pavilions", "Pavilion target"], sub="Active pavilion partners by region and branch",
              nonnull="Has pavilion", widths=[80, 90]), "V")
    s = add(table(R, "V | KRO", "KRO pavilions - primary purchases vs target", ["Zone/Region", "Branch (Pavilion Sales)"],
                  ["KRO pavilions", "Pavilion target", "Sales", "Sales PY", "Growth %", "Pavilion Ach %"],
                  sub="Only KROs buy in primary; matched to primary sales by GSTIN", nonnull="Has pavilion", widths=[80, 90]), "V")
    s.member_filter(R.fields["KRO/RO"], "KRO")
    s = add(table(R, "V | RO check", "Check: pavilions marked RO that look like KROs",
                  ["Name of RO/KRO", "Branch (Pavilion Sales)", "KRO/RO", "GSTIN no"],
                  ["Sales", "Sales PY"], sub="Marked RO but buying in primary (only KROs do) - please check the type",
                  nonnull="RO flag (agg)", widths=[240, 80, 50, 140], total=False), "V")
    s.mark_color = "#c0392b"
    pf = lambda n: R.fields[f"{PDS}|{n}"]
    MS, PT_, RA, SC, TT, WS = ("(Minimum!Slab!Achievement)", "(Pavilion!Type!Summary)", "(Retailer!App!TnC!Acceptance)",
                               "(Sales!Classification!Summary)", "(Total!Target!Achievement)", "(Wiser!Slabs!Unpivoted)")
    add(table(R, "PV | Type", "Pavilions by type", [pf("Pavilion Type"), pf(f"Sum of YTD Slab 2025 {PT_}"), pf("Sum of YTD Slab 2026")],
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

    # ---- MTD (the views of the original MTD dashboard)
    add(table(R, "M | KPIs", "KPIs - month to date", [],
              ["MTD target", "MTD sales", "MTD sales LY", "MTD growth %", "MTD OB", "MTD Ach %", "Sales last billing day", "OB last order day"],
              nonnull=None, total=False))
    add(table(R, "M | Focus", "Focus products - month target vs achievement", [],
              ["MTD Wiser target", "MTD Wiser sales", "MTD Wiser Ach %", "MTD Zeta target", "MTD Zeta sales", "MTD Zeta Ach %"],
              nonnull=None, total=False))
    add(table(R, "M | Leaderboard", "Sales leaderboard - top ASM in each region", ["Region", "Top ASM MTD"], ["MTD sales", "MTD growth %"],
              nonnull="Has MTD activity (agg)", widths=[70, 160], total=False))
    S["M | Leaderboard"].exclude_null(S["M | Leaderboard"].use(R["Top ASM MTD"], "dim"), R["Top ASM MTD"])
    add(table(R, "M | Regions", "Region wise month target, sales and OB", ["Region"],
              ["MTD target", "MTD sales", "MTD Ach %", "MTD growth %", "MTD OB"], nonnull="Has MTD activity (agg)", widths=[70]))
    add(trend(R, "M | Trend", "Day-wise sales and order book this month", ["MTD sales", "MTD OB"], "Green = sales",
              dim="Day", nonnull="Has MTD activity (agg)"))
    views = {
        "Branch": table(R, "MV | Branch", "Branch wise Sales, OB and Ach%", ["Region", "Branch"],
                        ["MTD target", "MTD sales", "MTD Ach %", "MTD sales LY", "MTD growth %", "MTD OB",
                         "MTD backlog (orders this month)", "Backlog", "Sales last billing day"], widths=[60, 100], nonnull="Has MTD activity (agg)"),
        "Activity": pivot(R, "MV | Activity", "Branch & Activity wise OB and Sales", ["Region", "Branch"], "Helios Code",
                          ["MTD sales", "MTD OB"], widths=[60, 100], nonnull="Has MTD activity (agg)"),
        "ASM": table(R, "MV | ASM", "ASM wise OB and Sales", ["Region", "ASM"],
                     ["MTD active partners", "MTD sales", "MTD sales LY", "MTD growth %", "MTD OB", "MTD backlog (orders this month)",
                      "Backlog", "Sales last billing day"], widths=[60, 180], nonnull="Has MTD activity (agg)", sort_by="MTD sales"),
        "Focus": table(R, "MV | Focus", "Focus Prod Month Target vs Ach% by branch", ["Region", "Branch"],
                       ["MTD Zeta target", "MTD Zeta sales", "MTD Zeta Ach %", "MTD Wiser target", "MTD Wiser sales", "MTD Wiser Ach %"],
                       widths=[60, 100], nonnull="Has MTD activity (agg)"),
        "Partner": table(R, "MV | Partner", "Partner wise OB and Sales", ["Partner", "RD/KRO", "ASM", "Region"],
                         ["MTD sales", "MTD sales LY", "MTD growth %", "MTD OB", "MTD backlog (orders this month)", "Backlog"],
                         widths=[260, 45, 160, 60], nonnull="Has MTD activity (agg)", sort_by="MTD sales", sort_dims=[0]),
        "Product": table(R, "MV | Product", "Product wise Sales, OB, Backlog", ["MaterialCode", "MaterialDescription", "Family"],
                         ["Qty sold MTD", "MTD sales", "Qty booked MTD", "MTD OB", "MTD backlog (orders this month)", "Backlog"],
                         widths=[90, 260, 120], nonnull="Has MTD activity (agg)", sort_by="MTD sales", sort_dims=[1]),
    }
    for key, s in views.items():
        add(s, "P", "P" if key == "Partner" else None)
        s.bool_filter(f"MTD view: {key}")

    # ---- Orders
    add(trend(R, "O | OB trend", "Monthly order book - this year vs last year", ["OB", "OB PY"]))
    add(hbar(R, "O | Backlog ageing", "Open backlog by age", "Backlog ageing", "Backlog", nonnull="Backlog",
             exclude_null=True, sort=False, color=C_TARGET))
    add(table(R, "O | OB by branch", "Order book and backlog by branch", ["Region", "Branch"],
              ["OB", "OB PY", "OB growth %", "Sales", "Sales to OB %", "Backlog", "Backlog over 50 days"], widths=[60, 100]))
    add(table(R, "O | Backlog partners", "Top 20 partners by open backlog", ["Partner", "Branch"],
              ["Backlog", "Backlog over 50 days", "Sales", "Growth %"], sort_by="Backlog", top=(20, "Backlog (row)"),
              widths=[300, 90], total=False, nonnull="Backlog"), "P", "P")
    add(table(R, "O | Open orders", "Open orders (backlog) - every line",
              ["Order", "SalesOrderItem", "Order date", "Backlog ageing", "Partner", "Branch", "ASM", "MaterialDescription"],
              ["Backlog", "Backlog age (max days)"], nonnull="Backlog", sort_by="Backlog", sort_dims=[0],
              widths=[70, 40, 75, 95, 210, 75, 110, 230]), "P", "P")

    # ---- Explore
    s = add(table(R, "X | Explore table", "Explore", ["Explore by"],
                  G + ["Share of sales %", "OB", "Backlog", "Active partners", "Quantity"], sort_by="Sales", widths=[300]), "P", "P")
    s.title = title_runs("Explore by <[Parameters].[p Dim]>")
    add(hbar(R, "X | Explore bars", "Top 30 by sales", "Explore by", "Sales", "Colour = growth", scale=("Growth %", GROWTH_SCALE),
             top=(30, "Sales CY (row)")), "P", "P")

    # ---- filters: shared dashboard filters, pavilion filters, row-level security
    for name, s in S.items():
        if s.scope == "P":
            s.filter_refs = {}
            for field, group, ctx in FILTERS:
                f = F(R, field)
                if s.top_field is not None and f.key == s.top_field.key:
                    continue
                if not take_over(s, f):
                    continue
                s.filter_refs[field] = s.shared_filter(field, group, ctx)
        elif s.scope == "V":
            s.filter_refs = {}
            s.bool_filter("Pavilion in scope", context=True)
            for field, group, ctx in PAV_FILTERS:
                if take_over(s, F(R, field)):
                    s.filter_refs[field] = s.shared_filter(field, group, ctx)
        if s.rls == "P":
            s.bool_filter("Branch access (sales)", context=True)
        elif s.rls == "Q":
            s.bool_filter("Branch access (partner targets)", context=True)
    return S


def take_over(s, f):
    """One filter per field and sheet: a shared filter replaces a 'hide blanks' filter on the same field,
    but a sheet's own selection (e.g. KRO only) stays and the shared filter is left out."""
    col = f"column='[{f.ds}].[none:{f.name}:nk]'"
    own = [x for x in s.filters if col in x]
    if not own:
        return True
    if all("function='except'" in x for x in own):
        for x in own:
            s.filters.remove(x)
        return True
    return False


# ---------------------------------------------------------------- dashboards
PERIOD = [("p Period", "compact"), ("p Start", "datetime"), ("p End", "datetime")]
VALUE = [("p Measure", "compact"), ("p Units", "compact")]


def params_for(dash):
    if dash == "MTD":
        return VALUE + [("p MTD view", "compact")]
    if dash == "Partner Targets":
        return PERIOD + VALUE + [("p Target view", "compact")]
    if dash == "Explore":
        return PERIOD + VALUE + [("p Dim", "compact")]
    return PERIOD + VALUE


def filters_for(dash):
    """Shared filters shown on the dashboard (field captions as in FILTERS / PAV_FILTERS)."""
    if dash == "Pavilion":
        return PAV_FILTERS
    return FILTERS


SIDE_NOTE = ("Filters apply to every sheet of every page. Growth compares the same dates last year "
             "(green = growth, red = decline).\n"
             "Monthly and yearly targets: "
             "Both net; gross = net / 0.83. Partner targets are gross; net = gross x 0.83.\n"
             "Activity: HDCFD = FD, HDCWD = WD.")

# rows of the content area: (height px, [cells]); cell = (sheet, width weight) | ("stack", weight, [(sheet, height)])
#                                                    | ("swap", weight, [sheets])  (one visible at a time)
# sheets swapped by a view parameter: the boolean field that shows each one
VISIBILITY = {**{f"PTV | {k}": f"Target view: {k}" for k, _ in TARGET_VIEWS}, **{f"MV | {k}": f"MTD view: {k}" for k, _ in MTD_VIEWS}}
H = lambda dash: (86, [(f"H | {dash}", 1)])
K = lambda *names: (128, [(n, 1) for n in names])

LAYOUT = {
    "Overview": [H("Overview"), K("K | Sales", "K | Yearly target", "K | Run-rate", "K | Data"),
                 K("K | Order book", "K | Backlog", "K | Partners", "K | New partners"),
                 (310, [("O | Three years", 1)]),
                 (300, [("O | Quarters", 1), ("O | Region share", 1), ("O | Partner type", 1)])],
    "Regions": [H("Regions"), (260, [("R | Scorecard", 1)]), (300, [("R | Trend", 1)]), (250, [("R | Growth heat", 1)]),
                (520, [("R | Branches", 1), ("R | Focus", 1)])],
    "Branches": [H("Branches"), (600, [("B | Scorecard", 1)]), (440, [("B | Focus families", 1)]), (520, [("B | ASM table", 1)])],
    "Products": [H("Products"), K("K | Sales", "K | Wiser", "K | Zeta", "K | Quantity"), (440, [("P | Family table", 1)]),
                 (420, [("P | Family bars", 1), ("P | Classification", 1)]),
                 (260, [("P | Activity", 1), ("P | New products", 1)]), (540, [("P | Top SKUs", 1)])],
    "Partners": [H("Partners"), K("K | Partners", "K | New partners", "K | Lapsed", "K | Per partner"),
                 (280, [("Pa | Ach bands", 1), ("Pa | Slabs", 1), ("Pa | Tier", 1)]),
                 (280, [("Pa | New partners", 2), ("Pa | Onboarded", 1)]),
                 (300, [("Pa | Top partners", 1)]), (440, [("Pa | Declining", 1)])],
    "All Partners": [H("All Partners"), (1200, [("AP | All partners", 1)]), (800, [("AP | By month", 1)])],
    "New Partner - New Geo": [K("K | NP partners", "K | NP sales", "K | NP target", "K | New geo"),
                              (300, [("PO | Summary", 3), ("PO | Buckets", 2)]), (230, [("NG | Categories", 1)]), (270, [("NG | Regions", 1)]),
                              (270, [("NG | Region x category", 1), ("NG | Tier", 1)]), (200, [("NG | Type", 1)]),
                              (260, [("NG | Products", 1)]), (520, [("NG | Cities", 1)])],
    "Monthly Targets": [H("Monthly Targets"), K("K | Monthly target", "K | Monthly gap", "K | Sales", "K | Data"),
                        (400, [("MT | Monthly", 1), ("MT | Chart", 1)]), (800, [("MT | Branch x month", 1)]),
                        (520, [("MT | Focus", 1)])],
    "Yearly Targets": [K("K | Yearly target", "K | Yearly balance", "K | Run-rate", "K | Sales"), (260, [("YT | Regions", 1)]),
                       (640, [("YT | Branches", 1)]), (620, [("YT | Ach bars", 1)])],
    "Partner Targets": [K("K | Annual partner", "K | Wiser target", "K | Zeta target", "K | Data"),
                        (1000, [("swap", 1, ["PTV | " + k for k, _ in TARGET_VIEWS])])],
    "SSP": [H("SSP"), K("K | SSP sales", "K | SSP target", "K | Sales", "K | Data"), (300, [("S | SSP table", 1)]),
            (300, [("S | With vs without", 1), ("S | Trend", 1)]), (300, [("S | New by SSP", 1)]), (520, [("S | New partners", 1)])],
    "Pavilion": [H("Pavilion"), (128, [("K | Total pavilions", 1), ("K | Pavilion slab", 2), ("K | KRO pavilions", 1), ("K | KRO purchases", 1)]),
                 (420, [("V | Count", 1), ("V | KRO", 1)]), (300, [("V | RO check", 1)]),
                 (180, [("PV | Type", 1), ("PV | Classification", 1)]), (260, [("PV | Min slab", 1), ("PV | Total target", 1)]),
                 (260, [("PV | App T&C", 1)]), (420, [("PV | Sales slabs", 1)]), (400, [("PV | Wiser slabs", 1)])],
    "MTD": [H("MTD"), (240, [("M | KPIs", 2), ("M | Focus", 2), ("M | Regions", 3)]),
            (300, [("M | Trend", 2), ("M | Leaderboard", 1)]),
            (760, [("swap", 1, ["MV | " + k for k, _ in MTD_VIEWS])])],
    "Order Book": [H("Order Book"), K("K | Order book", "K | Sales to OB", "K | Backlog", "K | Data"),
                   (320, [("O | OB trend", 2), ("O | Backlog ageing", 1)]), (520, [("O | OB by branch", 1)]),
                   (340, [("O | Backlog partners", 1)]), (600, [("O | Open orders", 1)])],
    "Explore": [(700, [("X | Explore table", 1)]), (680, [("X | Explore bars", 1)])],
}
