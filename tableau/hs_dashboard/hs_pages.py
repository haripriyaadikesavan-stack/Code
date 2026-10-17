"""Sheets and dashboards of the Home Solutions (HS) dashboard, in the style of the Retail dashboard.

Every page has the same filter panel: period, gross / net, units, business (All / Retail / Project),
with / without FD, the rebate rates (editable workbook only) and the shared dropdown filters.
"""
import pages as rp
from pages import F, est_rows, hbar, heat, lines_by, matrix, pivot, story, table, take_over, trend, grouped_bars
from hs_model import MTD_VIEWS
from sheets import ACH_SCALE, GROWTH_SCALE, kpi, nav

EDITABLE = True

CHAPTERS = [
    ("Overview", "Overview", "Where HS stands: sales vs last year, yearly target and commit, Retail vs Project, activities."),
    ("Regions", "Regions", "Region-wise performance. Click a region to focus the page on it."),
    ("Branches", "Branches", "Branch scorecard against yearly targets and commits. Click a branch to see its ASMs."),
    ("Activity", "Activity & Commit", "WD / FD / COS / SPE / EVC: sales and OB against the activity-wise commits."),
    ("Retail/Project", "Retail vs Project", "The two businesses side by side: sales, growth, targets and partners."),
    ("Products", "Products", "Families, product targets, new products (Wiser, Miluz Zeta, Miluz Lara) and top materials."),
    ("Partners", "Partners", "Partner base: achievement bands, slabs, tiers, new / revival partners and partners to watch."),
    ("All partners", "All Partners", "Numbers for every partner, and every partner month by month."),
    ("Monthly", "Monthly Targets", "Monthly targets vs sales and achievement, by month and branch."),
    ("MTD", "MTD", "This month so far: target, commit, sales, order book and the views of the original MTD dashboard."),
    ("Orders", "Order Book", "Order book, conversion to sales, open backlog and every open order."),
    ("Explore", "Explore", "Any dimension against all measures."),
]
NAV = [(label, dash) for label, dash, _ in CHAPTERS]
DASHBOARDS = [dash for _, dash, _ in CHAPTERS]
NOTES = {dash: note for _, dash, note in CHAPTERS}
DRILL = {"Regions": ["R | Scorecard"], "Branches": ["B | Scorecard"], "Activity & Commit": ["A | Activity"]}
FILTERS = [("Region", 1, True), ("Branch", 2, True), ("ASM", 3, True), ("Partner", 4, True), ("SoldToPartyID", 5, True),
           ("Activity", 6, False), ("Family", 7, False), ("RD/KRO", 8, True), ("Tier Category", 9, True),
           ("Partner growth band", 10, False)]
PHONE_FILTERS = ["Region", "Branch", "ASM", "Partner", "Activity", "Partner growth band"]
rp.EST.update({"Activity": 5, "GTM": 2, "Commit activity": 5})


def build_sheets(R):
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

    G = ["Sales", "Sales PY", "Growth %", "Growth (value)"]

    # ---- KPI cards
    add(kpi(R, "K | Sales", "Sales", "Sales", ("", "Growth text")))
    add(kpi(R, "K | Yearly target", "Yearly target achieved", "Yearly Ach %", ("Yearly target  ", "Yearly target")))
    add(kpi(R, "K | Commit", "Sales commit achieved", "Commit Ach %", ("Commit  ", "Sales commit")))
    add(kpi(R, "K | Run-rate", "Monthly run-rate", "Current run-rate / month", ("Required  ", "Required run-rate (yearly)")))
    add(kpi(R, "K | Order book", "Order book (OB)", "OB", ("", "OB growth text")))
    add(kpi(R, "K | OB commit", "OB commit achieved", "OB commit Ach %", ("OB commit  ", "OB commit")))
    add(kpi(R, "K | Backlog", "Open backlog", "Backlog", ("Ordered over 50 days ago  ", "Backlog over 50 days")))
    add(kpi(R, "K | Partners", "Active partners", "Active partners", ("New / revival  ", "New / revival partners")))
    add(kpi(R, "K | Retail share", "Retail share of sales", "Retail share %", ("FD share  ", "FD share %"), units=False))
    add(kpi(R, "K | Retail", "Retail sales", "Retail sales", ("Project  ", "Project sales")))
    add(kpi(R, "K | Data", "Data as of", "Data as of", ("Period  ", "Period text")))
    add(kpi(R, "K | Rebate", "Average rebate (net)", "Average rebate %", ("Data as of  ", "Data as of"), units=False))
    add(kpi(R, "K | Wiser", "Wiser sales", "Wiser sales", ("Growth  ", "Wiser growth %")))
    add(kpi(R, "K | Zeta", "Miluz Zeta sales", "Zeta sales", ("Growth  ", "Zeta growth %")))
    add(kpi(R, "K | Lara", "Miluz Lara sales", "Lara sales", ("Growth  ", "Lara growth %")))
    add(kpi(R, "K | Quantity", "Quantity sold", "Quantity", ("Growth  ", "Quantity growth %"), units=False))
    add(kpi(R, "K | Per partner", "Sales per active partner", "Sales per partner", ("Active last year  ", "Active partners PY")))
    add(kpi(R, "K | Lapsed", "Lapsed partners", "Lapsed partners", ("Active last year  ", "Active partners PY"), units=False))
    add(kpi(R, "K | Monthly target", "Monthly target achieved", "Monthly Ach %", ("Monthly target  ", "Monthly target")))
    add(kpi(R, "K | Monthly gap", "Gap to monthly target", "Gap to monthly target", ("Sales  ", "Sales")))
    add(kpi(R, "K | Sales to OB", "Sales to OB", "Sales to OB %", ("Sales  ", "Sales")))

    # ---- story headlines
    for dash, cap in [("Overview", "Overview"), ("Regions", "Regions"), ("Branches", "Branches"), ("Activity & Commit", "Activity"),
                      ("Products", "Products"), ("Partners", "Partners"), ("Monthly Targets", "Targets"), ("MTD", "MTD"),
                      ("Order Book", "Orders")]:
        add(story(R, f"H | {dash}", f"Story | {cap}"))

    # ---- Overview
    add(trend(R, "O | Three years", "Sales by month - this year, last year and two years ago",
              ["Sales", "Sales PY", "Sales 2 years ago"], "Same dates each year", nonnull="Has sales (agg)"))
    add(grouped_bars(R, "O | Quarters", "Sales by quarter vs last year", "Quarter", ["Sales", "Sales PY"]))
    add(hbar(R, "O | Region share", "Share of sales by region", "Region", "Share of sales %", "Colour = growth",
             scale=("Growth %", GROWTH_SCALE)))
    add(grouped_bars(R, "O | GTM", "Retail vs Project vs last year", "GTM", ["Sales", "Sales PY"]))
    add(hbar(R, "O | Activity", "Sales by activity", "Activity", "Sales", "Colour = growth", scale=("Growth %", GROWTH_SCALE)))

    # ---- Regions
    add(table(R, "R | Scorecard", "Region scorecard", ["Region"],
              G + ["Yearly target", "Yearly Ach %", "Sales commit", "Commit Ach %", "OB", "OB growth %", "Backlog", "Active partners"],
              widths=[80]))
    add(lines_by(R, "R | Trend", "Monthly sales by region", "Month", "Region", "Sales"))
    add(heat(R, "R | Growth heat", "Growth % by region and month (red = decline)", "Region", "Month", "Growth %", GROWTH_SCALE))
    add(matrix(R, "R | Activity", "Region x activity sales", ["Region"], "Activity", "Sales", widths=[80]))
    add(table(R, "R | Branches", "Branch sales and growth", ["Region", "Branch"], G + ["Share of sales %", "Active partners"],
              widths=[70, 100]))

    # ---- Branches
    add(table(R, "B | Scorecard", "Branch scorecard vs yearly targets and commits", ["Region", "Branch"],
              ["Yearly target", "Sales", "Yearly Ach %", "Balance to yearly target", "Required run-rate (yearly)", "Sales PY",
               "Growth %", "Sales commit", "Commit Ach %", "OB", "Backlog"], widths=[60, 100]))
    add(matrix(R, "B | Activity", "Branch x activity sales", ["Region", "Branch"], "Activity", "Sales", widths=[60, 100]))
    add(table(R, "B | ASM table", "ASM performance", ["Region", "ASM"], G + ["OB", "Backlog", "Active partners", "Sales per partner"],
              widths=[70, 180], sort_by="Sales"), "P", "P")

    # ---- Activity & Commit
    add(table(R, "A | Activity", "Activity performance vs commit", ["Activity"],
              ["Sales PY", "Sales", "Growth %", "Sales commit", "Commit Ach %", "OB", "OB commit", "OB commit Ach %", "Backlog",
               "Share of sales %"], sub="Commits are net; gross = net / 0.87", widths=[80]))
    add(table(R, "A | Monthly commit", "Month-wise commit vs sales and OB", ["Month"],
              ["Sales commit", "Sales", "Commit Ach %", "Gap to commit", "OB commit", "OB", "OB commit Ach %"], widths=[80],
              nonnull="Has commit (agg)"))
    add(trend(R, "A | Commit chart", "Monthly sales vs sales commit", ["Sales", "Sales commit"], "Green = sales",
              nonnull="Has commit (agg)"))
    add(table(R, "A | Branch commit", "Branch-wise commit vs sales and OB", ["Region", "Branch"],
              ["Sales commit", "Sales", "Commit Ach %", "OB commit", "OB", "OB commit Ach %"], widths=[60, 100],
              nonnull="Has commit (agg)"))
    add(pivot(R, "A | Branch x activity", "Branch x activity - commit vs sales", ["Region", "Branch"], "Activity",
              ["Sales commit", "Sales", "Commit Ach %"], widths=[60, 100], nonnull="Has commit (agg)"))

    # ---- Retail vs Project
    add(table(R, "GP | GTM", "Retail vs Project", ["GTM"], G + ["Share of sales %", "OB", "OB growth %", "Backlog", "Active partners",
                                                               "Sales per partner"], widths=[90]))
    add(matrix(R, "GP | Region", "Region x business sales", ["Region"], "GTM", "Sales", widths=[80]))
    add(matrix(R, "GP | Branch", "Branch x business sales", ["Region", "Branch"], "GTM", "Sales", widths=[60, 100]))
    add(matrix(R, "GP | Activity", "Activity x business sales", ["Activity"], "GTM", "Sales", widths=[80]))
    add(lines_by(R, "GP | Trend", "Monthly sales - Retail vs Project", "Month", "GTM", "Sales"))

    # ---- Products
    add(table(R, "P | Family table", "Family performance (value and quantity)", ["Family"],
              G + ["Family target", "Family Ach %", "Share of sales %", "Quantity", "Quantity growth %"],
              sub="Family targets by region, where set for the year", widths=[170], sort_by="Sales"))
    add(hbar(R, "P | Family bars", "Top 15 families by sales", "Family", "Sales", "Colour = growth", scale=("Growth %", GROWTH_SCALE),
             top=(15, "Sales CY (row)")))
    add(hbar(R, "P | Classification", "Sales by product classification", "Classification (Sheet1)", "Sales", "Colour = growth",
             scale=("Growth %", GROWTH_SCALE), exclude_null=True))
    add(table(R, "P | New products", "New products", ["New product category"], G + ["Share of sales %"], widths=[130]))
    add(table(R, "P | Top SKUs", "Top 25 materials", ["MaterialCode", "MaterialDescription", "Family"],
              G + ["Quantity", "Quantity growth %"], sort_by="Sales", top=(25, "Sales CY (row)"), widths=[90, 280, 120], total=False))

    # ---- Partners
    add(hbar(R, "Pa | Ach bands", "Partners by target achievement", "Partner ach band", "Partners in band",
             "Partner targets are gross", nonnull=None, exclude_null=True, sort=False), "P", "P")
    add(hbar(R, "Pa | Slabs", "Partners by sales slab", "Partner slab", "Active partners", nonnull=None, exclude_null=True,
             sort=False), "P", "P")
    add(hbar(R, "Pa | Tier", "Sales by tier", "Tier Category", "Sales", "Colour = growth", scale=("Growth %", GROWTH_SCALE),
             exclude_null=True), "P", "P")
    add(hbar(R, "Pa | New partners", "New / revival partners buying each month", "Month", "New / revival partners",
             "New / revival = bought this year, nothing in the previous two years", nonnull="New / revival partners",
             exclude_null=True, sort=False, vertical=True), "P", "P")
    add(table(R, "Pa | Top partners", "Top 10 partners", ["Partner", "Branch"], G + ["Share of sales %"],
              sort_by="Sales", top=(10, "Sales CY (row)"), widths=[260, 80], total=False, nonnull="Has sales (agg)"), "P", "P")
    s = add(table(R, "Pa | Declining", "Partners to watch - largest drops vs last year (declining partners only)",
                  ["Partner", "Branch", "ASM"], ["Sales PY", "Sales", "Growth (value)", "Growth %", "Days since last sale"],
                  sort_by=("Growth (value)", "ASC"), top=(25, "Decline (row)"), widths=[280, 80, 170], total=False,
                  nonnull="Has sales (agg)"), "P", "P")
    s.nonnull("Decline (agg)")

    # ---- All partners
    add(table(R, "AP | All partners", "All partners", ["Region", "Branch", "ASM", "Partner", "City (Geography mapping)", "RD/KRO",
                                                       "Tier Category", "Partner status"],
              G + ["Partner FY target", "Partner FY Ach %", "OB", "Backlog", "Days since last sale"],
              sub="Partner targets are gross; Net = gross x 0.83",
              widths=[55, 75, 130, 250, 95, 45, 55, 90], col_width=72, sort_by="Sales", sort_dims=[3], fit=None), "P", "P")
    add(matrix(R, "AP | By month", "Partner sales month by month", ["Branch", "Partner"], "Month", "Sales",
               widths=[75, 260], fit=None), "P", "P")

    # ---- Monthly targets
    add(table(R, "MT | Monthly", "Monthly target vs sales and achievement", ["Month"],
              ["Monthly target", "Sales", "Monthly Ach %", "Gap to monthly target", "Sales PY", "Growth %"],
              sub="Monthly targets are net; gross = net / 0.83", widths=[80], nonnull="Has sales (agg)"))
    add(trend(R, "MT | Chart", "Monthly sales vs monthly target", ["Sales", "Monthly target"], "Green = sales, amber = monthly target",
              nonnull="Has sales (agg)"))
    add(pivot(R, "MT | Branch x month", "Branch monthly target vs sales and achievement", ["Region", "Branch"], "Month",
              ["Monthly target", "Sales", "Monthly Ach %"], widths=[60, 100]))
    add(table(R, "MT | Focus", "Wiser and Miluz Zeta monthly targets by branch", ["Region", "Branch"],
              ["Wiser target", "Wiser sales", "Wiser Ach %", "Zeta target", "Zeta sales", "Zeta Ach %"], widths=[60, 100]))

    # ---- MTD
    add(table(R, "M | KPIs", "KPIs - month to date", [],
              ["MTD target", "MTD commit", "MTD sales", "MTD sales LY", "MTD growth %", "MTD Ach %", "MTD commit Ach %", "MTD OB",
               "MTD OB commit Ach %", "Sales last billing day", "OB last order day"], nonnull=None, total=False))
    add(table(R, "M | Focus", "Focus products - month target vs achievement", [],
              ["MTD Wiser target", "MTD Wiser sales", "MTD Wiser Ach %", "MTD Zeta target", "MTD Zeta sales", "MTD Zeta Ach %",
               "MTD Lara sales"], nonnull=None, total=False))
    add(table(R, "M | Leaderboard", "Sales leaderboard - top ASM in each region", ["Region", "Top ASM MTD"], ["MTD sales", "MTD growth %"],
              nonnull="Has MTD activity (agg)", widths=[70, 160], total=False))
    S["M | Leaderboard"].exclude_null(S["M | Leaderboard"].use(R["Top ASM MTD"], "dim"), R["Top ASM MTD"])
    add(table(R, "M | Regions", "Region wise month target, commit, sales and OB", ["Region"],
              ["MTD target", "MTD commit", "MTD sales", "MTD Ach %", "MTD commit Ach %", "MTD growth %", "MTD OB"],
              nonnull="Has MTD activity (agg)", widths=[70]))
    add(trend(R, "M | Trend", "Day-wise sales and order book this month", ["MTD sales", "MTD OB"], "Green = sales",
              dim="Day", nonnull="Has MTD activity (agg)"))
    views = {
        "Branch": table(R, "MV | Branch", "Branch wise Sales, OB and Ach%", ["Region", "Branch"],
                        ["MTD target", "MTD commit", "MTD sales", "MTD Ach %", "MTD commit Ach %", "MTD sales LY", "MTD growth %",
                         "MTD OB", "MTD backlog (orders this month)", "Backlog"], widths=[60, 100], nonnull="Has MTD activity (agg)"),
        "Activity": pivot(R, "MV | Activity", "Branch & Activity wise OB and Sales", ["Region", "Branch"], "Activity",
                          ["MTD sales", "MTD OB"], widths=[60, 100], nonnull="Has MTD activity (agg)"),
        "ASM": table(R, "MV | ASM", "ASM wise OB and Sales", ["Region", "ASM"],
                     ["MTD active partners", "MTD sales", "MTD sales LY", "MTD growth %", "MTD OB", "MTD backlog (orders this month)",
                      "Backlog"], widths=[60, 180], nonnull="Has MTD activity (agg)", sort_by="MTD sales"),
        "Partner": table(R, "MV | Partner", "Partner wise OB and Sales", ["Partner", "RD/KRO", "ASM", "Region"],
                         ["MTD sales", "MTD sales LY", "MTD growth %", "MTD OB", "MTD backlog (orders this month)", "Backlog"],
                         widths=[260, 60, 160, 60], nonnull="Has MTD activity (agg)", sort_by="MTD sales", sort_dims=[0]),
        "Product": table(R, "MV | Product", "Product wise Sales, OB, Backlog", ["MaterialCode", "MaterialDescription", "Family"],
                         ["Qty sold MTD", "MTD sales", "Qty booked MTD", "MTD OB", "MTD backlog (orders this month)", "Backlog"],
                         widths=[90, 260, 120], nonnull="Has MTD activity (agg)", sort_by="MTD sales", sort_dims=[1]),
    }
    for key, s in views.items():
        add(s, "P", "P" if key == "Partner" else None)
        s.bool_filter(f"MTD view: {key}")

    # ---- Orders
    add(trend(R, "O | OB trend", "Order book vs sales by month", ["OB", "Sales"], "Green = OB", nonnull="Has activity (agg)"))
    add(hbar(R, "O | Backlog ageing", "Open backlog by age", "Backlog ageing", "Backlog", "By order creation date",
             nonnull=None, exclude_null=True, sort=False))
    add(table(R, "O | OB by branch", "Order book, sales and backlog by branch", ["Region", "Branch"],
              ["OB", "OB PY", "OB growth %", "Sales", "Sales to OB %", "Backlog", "Backlog over 50 days"], widths=[60, 100]))
    add(table(R, "O | Backlog partners", "Top 20 partners by open backlog", ["Partner", "Branch"],
              ["Backlog", "Backlog over 50 days", "Sales", "Growth %"], sort_by="Backlog", top=(20, "Backlog (row)"),
              widths=[280, 90], total=False, nonnull=None), "P", "P")
    s = add(table(R, "O | Open orders", "Every open order line (backlog)", ["Branch", "Activity", "Order", "Order date", "Partner",
                                                                           "MaterialCode", "MaterialDescription"],
                  ["Backlog", "Backlog age (max days)"], sort_by="Backlog", sort_dims=[2], total=False, nonnull=None,
                  widths=[70, 50, 85, 85, 230, 90, 230]), "P", "P")
    s.nonnull("Backlog")

    # ---- Explore
    s = add(table(R, "X | Explore table", "Explore", ["Explore by"],
                  G + ["Share of sales %", "OB", "Backlog", "Active partners", "Quantity"], sort_by="Sales", widths=[300]), "P", "P")
    s.title = rp.title_runs("Explore by <[Parameters].[p Dim]>")
    add(hbar(R, "X | Explore bars", "Top 30 by sales", "Explore by", "Sales", "Colour = growth", scale=("Growth %", GROWTH_SCALE),
             top=(30, "Sales CY (row)")), "P", "P")

    # ---- filters: shared dashboard filters, FD / business scope, row-level security
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
            s.bool_filter("In scope (FD / GTM)", context=True)
            s.bool_filter("Branch access (sales)", context=True)
    return S


# ---------------------------------------------------------------- dashboards
PERIOD = [("p Period", "compact"), ("p Start", "datetime"), ("p End", "datetime")]
VALUE = [("p Measure", "compact"), ("p Units", "compact")]
SCOPE = [("p GTM", "compact"), ("p FD", "compact")]
REBATE = [("p Rebate retail CY", "type_in"), ("p Rebate project CY", "type_in"),
          ("p Rebate retail PY", "type_in"), ("p Rebate project PY", "type_in")]


def params_for(dash):
    rebate = REBATE if EDITABLE else []
    if dash == "MTD":
        return VALUE + SCOPE + rebate + [("p MTD view", "compact")]
    if dash == "Explore":
        return PERIOD + VALUE + SCOPE + rebate + [("p Dim", "compact")]
    return PERIOD + VALUE + SCOPE + rebate


def filters_for(dash):
    return FILTERS


def side_note():
    reb = ("Net = gross x (1 - rebate); type the rebate rates above (0.13 = 13%). " if EDITABLE else
           "Net = gross x (1 - rebate): Retail 13% / Project 5.1% this year, 12.2% / 4.5% last year. ")
    return ("Filters apply to every sheet of every page. Growth compares the same dates last year (green = growth, red = decline).\n"
            + reb + "Retail = pricing class EKR / ERW, or FD; the rest is Project.\n"
            "Monthly and yearly targets are net (gross = net / 0.83); commits are net (gross = net / 0.87); partner targets are gross.\n"
            "Without FD removes FD sales, OB, backlog, FD commits and FD yearly targets.")


SIDE_NOTE = side_note()
VISIBILITY = {f"MV | {k}": f"MTD view: {k}" for k, _ in MTD_VIEWS}
H = lambda dash: (86, [(f"H | {dash}", 1)])
K = lambda *names: (128, [(n, 1) for n in names])

LAYOUT = {
    "Overview": [H("Overview"), K("K | Sales", "K | Yearly target", "K | Commit", "K | Data"),
                 K("K | Order book", "K | Backlog", "K | Partners", "K | Retail share"),
                 (310, [("O | Three years", 1)]),
                 (300, [("O | Quarters", 1), ("O | GTM", 1), ("O | Region share", 1), ("O | Activity", 1)])],
    "Regions": [H("Regions"), (260, [("R | Scorecard", 1)]), (300, [("R | Trend", 1)]), (250, [("R | Growth heat", 1)]),
                (230, [("R | Activity", 1)]), (520, [("R | Branches", 1)])],
    "Branches": [H("Branches"), (600, [("B | Scorecard", 1)]), (600, [("B | Activity", 1)]), (520, [("B | ASM table", 1)])],
    "Activity & Commit": [H("Activity & Commit"), K("K | Commit", "K | OB commit", "K | Sales", "K | Order book"),
                          (230, [("A | Activity", 1)]), (360, [("A | Monthly commit", 1), ("A | Commit chart", 1)]),
                          (600, [("A | Branch commit", 1)]), (900, [("A | Branch x activity", 1)])],
    "Retail vs Project": [K("K | Retail", "K | Retail share", "K | Sales", "K | Rebate"), (180, [("GP | GTM", 1)]),
                          (300, [("GP | Trend", 2), ("GP | Activity", 1)]), (230, [("GP | Region", 1)]),
                          (600, [("GP | Branch", 1)])],
    "Products": [H("Products"), K("K | Sales", "K | Wiser", "K | Zeta", "K | Lara"), (600, [("P | Family table", 1)]),
                 (420, [("P | Family bars", 1), ("P | Classification", 1)]), (230, [("P | New products", 1)]),
                 (540, [("P | Top SKUs", 1)])],
    "Partners": [H("Partners"), K("K | Partners", "K | Per partner", "K | Lapsed", "K | Quantity"),
                 (280, [("Pa | Ach bands", 1), ("Pa | Slabs", 1), ("Pa | Tier", 1)]),
                 (280, [("Pa | New partners", 1)]), (300, [("Pa | Top partners", 1)]), (440, [("Pa | Declining", 1)])],
    "All Partners": [(1200, [("AP | All partners", 1)]), (800, [("AP | By month", 1)])],
    "Monthly Targets": [H("Monthly Targets"), K("K | Monthly target", "K | Monthly gap", "K | Sales", "K | Data"),
                        (400, [("MT | Monthly", 1), ("MT | Chart", 1)]), (800, [("MT | Branch x month", 1)]),
                        (520, [("MT | Focus", 1)])],
    "MTD": [H("MTD"), (300, [("M | KPIs", 2), ("M | Focus", 2), ("M | Regions", 3)]),
            (300, [("M | Trend", 2), ("M | Leaderboard", 1)]),
            (760, [("swap", 1, ["MV | " + k for k, _ in MTD_VIEWS])])],
    "Order Book": [H("Order Book"), K("K | Order book", "K | Sales to OB", "K | Backlog", "K | Data"),
                   (320, [("O | OB trend", 2), ("O | Backlog ageing", 1)]), (520, [("O | OB by branch", 1)]),
                   (340, [("O | Backlog partners", 1)]), (600, [("O | Open orders", 1)])],
    "Explore": [(700, [("X | Explore table", 1)]), (680, [("X | Explore bars", 1)])],
}
