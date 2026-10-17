"""Sheets and dashboards of the Wiser dashboard (sales and services), in the style of the Retail dashboard."""
import pages as rp
from pages import F, hbar, heat, lines_by, matrix, pivot, story, table, take_over, trend, grouped_bars
from sheets import GROWTH_SCALE, kpi, nav
from wiser_model import MTD_VIEWS

CHAPTERS = [
    ("Overview", "Overview", "Where Wiser stands: sales vs last year and the yearly target, Retail vs Project, Wiser 1.0 vs 2.0."),
    ("Branches", "Regions & Branches", "Region and branch performance against the yearly Wiser targets. Click a branch to filter the page."),
    ("Products", "Products", "Wiser 1.0 vs 2.0 and every Wiser material."),
    ("Partners", "Partners", "Partners buying Wiser: presence, slabs, new / revival, top partners and partners to watch."),
    ("All partners", "All Partners", "Numbers for every partner buying Wiser, and month by month."),
    ("Monthly", "Monthly Targets", "Monthly Wiser targets vs sales and achievement, by month and branch."),
    ("MTD", "MTD", "This month so far: target, sales, order book and the views of the original MTD dashboard."),
    ("Orders", "Order Book", "Order book, cancellations, open backlog and every open order."),
    ("Services", "Services", "Wiser service requests: installations, repairs, turnaround times, problems and engineers."),
    ("Explore", "Explore", "Any dimension against all measures."),
]
NAV = [(label, dash) for label, dash, _ in CHAPTERS]
DASHBOARDS = [dash for _, dash, _ in CHAPTERS]
NOTES = {dash: note for _, dash, note in CHAPTERS}
DRILL = {"Regions & Branches": ["B | Scorecard"]}
FILTERS = [("Region", 1, True), ("Branch", 2, True), ("SPOC", 3, True), ("Partner", 4, True), ("SoldToPartyID", 5, True),
           ("Wiser version", 6, False), ("RD/KRO", 7, True), ("Partner growth band", 8, False)]
SV_FILTERS = [("Service region", 21, True), ("Service branch", 22, True), ("Work order group", 23, False),
              ("Problem category", 24, False)]
PHONE_FILTERS = ["Region", "Branch", "SPOC", "Partner", "Wiser version", "Partner growth band",
                 "Service region", "Service branch", "Work order group", "Problem category"]
rp.EST.update({"Wiser version": 3, "GTM": 2, "Service region": 6, "Service branch": 27, "Work order group": 3,
               "Problem category": 8, "WO month": 12, "Work Order Status *": 4, "Top SPOC MTD": 5})


def build_sheets(R):
    S = {}

    def add(s, scope="P"):
        s.scope = scope
        S[s.name] = s
        return s

    for label, dash in NAV:
        add(nav(R, dash, label), "nav")
    for label in ("‹ Back", "Next ›"):
        s = add(nav(R, label, label), "nav")
        s.nav_target = None

    G = ["Sales", "Sales PY", "Growth %", "Growth (value)"]
    SVN = "Has service (agg)"

    # ---- KPI cards
    add(kpi(R, "K | Sales", "Wiser sales", "Sales", ("", "Growth text")))
    add(kpi(R, "K | Yearly target", "Yearly Wiser target achieved", "Yearly Ach %", ("Yearly target  ", "Yearly target")))
    add(kpi(R, "K | Monthly target", "Monthly target achieved", "Monthly Ach %", ("Monthly target  ", "Monthly target")))
    add(kpi(R, "K | Monthly gap", "Gap to monthly target", "Gap to monthly target", ("Sales  ", "Sales")))
    add(kpi(R, "K | Order book", "Order book (OB)", "OB", ("", "OB growth text")))
    add(kpi(R, "K | Backlog", "Open backlog", "Backlog", ("Ordered over 50 days ago  ", "Backlog over 50 days")))
    add(kpi(R, "K | Cancelled", "Cancelled orders", "Cancelled orders", ("OB  ", "OB")))
    add(kpi(R, "K | Partners", "Partners buying Wiser", "Active partners", ("New / revival  ", "New / revival partners")))
    add(kpi(R, "K | Per partner", "Sales per partner", "Sales per partner", ("Active last year  ", "Active partners PY")))
    add(kpi(R, "K | Lapsed", "Lapsed partners", "Lapsed partners", ("Active last year  ", "Active partners PY"), units=False))
    add(kpi(R, "K | Retail", "Retail Wiser sales", "Retail sales", ("Retail share  ", "Retail share %")))
    add(kpi(R, "K | Quantity", "Wiser units sold", "Quantity", ("Growth  ", "Quantity growth %"), units=False))
    add(kpi(R, "K | Sales to OB", "Sales to OB", "Sales to OB %", ("Sales  ", "Sales")))
    add(kpi(R, "K | Data", "Data as of", "Data as of", ("Period  ", "Period text")))
    add(kpi(R, "SK | Requests", "Service requests", "Service requests", ("Growth  ", "Service requests growth %"), units=False), "SV")
    add(kpi(R, "SK | Installations", "Installations", "Installations", ("Avg TAT (days)  ", "Avg installation TAT (days)"), units=False), "SV")
    add(kpi(R, "SK | Repairs", "Repairs", "Repairs", ("Avg TAT (days)  ", "Avg repair TAT (days)"), units=False), "SV")
    add(kpi(R, "SK | On time", "Done within 3 days", "Done within 3 days %", ("Completed  ", "Completed %"), units=False), "SV")

    for dash, cap in [("Overview", "Overview"), ("Regions & Branches", "Branches"), ("Partners", "Partners"),
                      ("Monthly Targets", "Targets"), ("MTD", "MTD"), ("Order Book", "Orders"), ("Services", "Services")]:
        add(story(R, f"H | {dash}", f"Story | {cap}"), "SV" if dash == "Services" else "P")

    # ---- Overview
    add(trend(R, "O | Three years", "Wiser sales by month - this year, last year and two years ago",
              ["Sales", "Sales PY", "Sales 2 years ago"], "Same dates each year", nonnull="Has sales (agg)"))
    add(grouped_bars(R, "O | Quarters", "Wiser sales by quarter vs last year", "Quarter", ["Sales", "Sales PY"]))
    add(grouped_bars(R, "O | GTM", "Retail vs Project vs last year", "GTM", ["Sales", "Sales PY"]))
    add(grouped_bars(R, "O | Version", "Wiser 1.0 vs 2.0 vs last year", "Wiser version", ["Sales", "Sales PY"]))
    add(hbar(R, "O | Region share", "Share of Wiser sales by region", "Region", "Share of sales %", "Colour = growth",
             scale=("Growth %", GROWTH_SCALE)))

    # ---- Regions & Branches
    add(table(R, "B | Regions", "Region scorecard", ["Region"],
              G + ["Yearly target", "Yearly Ach %", "OB", "Backlog", "Active partners"], widths=[80]))
    add(table(R, "B | Scorecard", "Branch scorecard vs yearly Wiser targets", ["Region", "Branch"],
              ["Yearly target", "Sales", "Yearly Ach %", "Balance to yearly target", "Required run-rate (yearly)", "Sales PY",
               "Growth %", "OB", "Backlog", "Active partners"], sub="Yearly Wiser targets by branch, Retail and Project, as set",
              widths=[60, 100]))
    add(table(R, "B | GTM targets", "Branch - Retail and Project targets", ["Region", "Branch"],
              ["Retail target", "Retail sales", "Retail Ach %", "Project target", "Project sales", "Project Ach %"], widths=[60, 100]))
    add(lines_by(R, "B | Trend", "Monthly Wiser sales by region", "Month", "Region", "Sales"))
    add(table(R, "B | SPOC", "Wiser SPOC performance", ["Region", "SPOC"], G + ["OB", "Backlog", "Active partners"],
              widths=[70, 180], sort_by="Sales"))

    # ---- Products
    add(table(R, "P | Versions", "Wiser 1.0 vs 2.0", ["Wiser version"], G + ["Share of sales %", "Quantity", "Quantity growth %",
                                                                          "OB", "Backlog"], widths=[110]))
    add(lines_by(R, "P | Version trend", "Monthly sales - Wiser 1.0 vs 2.0", "Month", "Wiser version", "Sales"))
    add(table(R, "P | Materials", "Every Wiser material", ["Wiser version", "MaterialCode", "MaterialDescription"],
              G + ["Quantity", "Quantity growth %", "OB", "Backlog"], sort_by="Sales", sort_dims=[2], widths=[90, 110, 300]))
    add(matrix(R, "P | Material x region", "Material sales by region", ["MaterialDescription"], "Region", "Sales", widths=[300],
               fit=None))

    # ---- Partners
    add(hbar(R, "Pa | Slabs", "Partners by Wiser sales slab", "Partner slab", "Active partners", nonnull=None, exclude_null=True,
             sort=False))
    add(hbar(R, "Pa | Type", "Wiser sales by partner type", "RD/KRO", "Sales", "Colour = growth", scale=("Growth %", GROWTH_SCALE),
             exclude_null=True))
    add(table(R, "Pa | Presence", "Partner presence by region", ["Region"], ["Active partners", "Active partners PY", "New / revival partners",
                                                                           "Lapsed partners", "Sales per partner"], widths=[80]))
    add(hbar(R, "Pa | New partners", "New / revival Wiser partners buying each month", "Month", "New / revival partners",
             "New / revival = bought Wiser this year, nothing in the previous two years", nonnull="New / revival partners",
             exclude_null=True, sort=False, vertical=True))
    add(table(R, "Pa | Top partners", "Top 10 Wiser partners", ["Partner", "Branch"], G + ["Share of sales %"],
              sort_by="Sales", top=(10, "Sales CY (row)"), widths=[260, 80], total=False, nonnull="Has sales (agg)"))
    s = add(table(R, "Pa | Declining", "Partners to watch - largest drops vs last year", ["Partner", "Branch", "SPOC"],
                  ["Sales PY", "Sales", "Growth (value)", "Growth %", "Days since last sale"],
                  sort_by=("Growth (value)", "ASC"), top=(25, "Decline (row)"), widths=[280, 80, 170], total=False,
                  nonnull="Has sales (agg)"))
    s.nonnull("Decline (agg)")

    # ---- All partners
    add(table(R, "AP | All partners", "All Wiser partners", ["Region", "Branch", "SPOC", "Partner", "City (Geography mapping)", "RD/KRO",
                                                             "Partner status"],
              G + ["Months active", "Quantity", "OB", "Backlog", "Days since last sale"],
              widths=[55, 75, 130, 250, 95, 50, 90], col_width=72, sort_by="Sales", sort_dims=[3], fit=None))
    add(matrix(R, "AP | By month", "Partner Wiser sales month by month", ["Branch", "Partner"], "Month", "Sales",
               widths=[75, 260], fit=None))

    # ---- Monthly targets
    add(table(R, "MT | Monthly", "Monthly Wiser target vs sales and achievement", ["Month"],
              ["Monthly target", "Sales", "Monthly Ach %", "Gap to monthly target", "Sales PY", "Growth %"],
              sub="Monthly targets are net; gross = net / 0.83", widths=[80], nonnull="Has sales (agg)"))
    add(trend(R, "MT | Chart", "Monthly Wiser sales vs monthly target", ["Sales", "Monthly target"], "Green = sales, amber = target",
              nonnull="Has sales (agg)"))
    add(pivot(R, "MT | Branch x month", "Branch monthly target vs sales and achievement", ["Region", "Branch"], "Month",
              ["Monthly target", "Sales", "Monthly Ach %"], widths=[60, 100]))

    # ---- MTD
    add(table(R, "M | KPIs", "KPIs - month to date", [],
              ["MTD target", "MTD sales", "MTD sales LY", "MTD growth %", "MTD Ach %", "MTD OB", "Sales last billing day",
               "OB last order day"], nonnull=None, total=False))
    add(table(R, "M | Leaderboard", "Leaderboard - top SPOC in each region", ["Region", "Top SPOC MTD"], ["MTD sales", "MTD growth %"],
              nonnull="Has MTD activity (agg)", widths=[70, 160], total=False))
    S["M | Leaderboard"].exclude_null(S["M | Leaderboard"].use(R["Top SPOC MTD"], "dim"), R["Top SPOC MTD"])
    add(table(R, "M | Regions", "Region wise month target, sales and OB", ["Region"],
              ["MTD target", "MTD sales", "MTD Ach %", "MTD growth %", "MTD OB"], nonnull="Has MTD activity (agg)", widths=[70]))
    add(trend(R, "M | Trend", "Day-wise Wiser sales and order book this month", ["MTD sales", "MTD OB"], "Green = sales",
              dim="Day", nonnull="Has MTD activity (agg)"))
    views = {
        "Branch": table(R, "MV | Branch", "Branch wise Sales, OB and Ach%", ["Region", "Branch"],
                        ["MTD target", "MTD sales", "MTD Ach %", "MTD sales LY", "MTD growth %", "MTD OB",
                         "MTD backlog (orders this month)", "Backlog"], widths=[60, 100], nonnull="Has MTD activity (agg)"),
        "SPOC": table(R, "MV | SPOC", "SPOC wise OB and Sales", ["Region", "SPOC"],
                      ["MTD active partners", "MTD sales", "MTD sales LY", "MTD growth %", "MTD OB", "MTD backlog (orders this month)",
                       "Backlog"], widths=[60, 180], nonnull="Has MTD activity (agg)", sort_by="MTD sales"),
        "Partner": table(R, "MV | Partner", "Partner wise OB and Sales", ["Partner", "RD/KRO", "SPOC", "Region"],
                         ["MTD sales", "MTD sales LY", "MTD growth %", "MTD OB", "MTD backlog (orders this month)", "Backlog"],
                         widths=[260, 60, 160, 60], nonnull="Has MTD activity (agg)", sort_by="MTD sales", sort_dims=[0]),
        "Material": table(R, "MV | Material", "Material wise Sales, OB, Backlog", ["MaterialCode", "MaterialDescription"],
                          ["Qty sold MTD", "MTD sales", "Qty booked MTD", "MTD OB", "MTD backlog (orders this month)", "Backlog"],
                          widths=[110, 300], nonnull="Has MTD activity (agg)", sort_by="MTD sales", sort_dims=[1]),
    }
    for key, s in views.items():
        add(s)
        s.bool_filter(f"MTD view: {key}")

    # ---- Orders
    add(trend(R, "O | OB trend", "Wiser order book vs sales by month", ["OB", "Sales"], "Green = OB", nonnull="Has activity (agg)"))
    add(hbar(R, "O | Backlog ageing", "Open backlog by age", "Backlog ageing", "Backlog", "By order creation date",
             nonnull=None, exclude_null=True, sort=False))
    add(table(R, "O | OB by branch", "Order book, cancellations, sales and backlog by branch", ["Region", "Branch"],
              ["OB", "OB PY", "OB growth %", "Cancelled orders", "Sales", "Sales to OB %", "Backlog", "Backlog over 50 days"],
              widths=[60, 100]))
    s = add(table(R, "O | Cancellations", "Cancelled order lines (negative OB) in the period",
                  ["Branch", "OB date", "Order", "Partner", "MaterialDescription"], ["Cancelled orders"],
                  sort_by="Cancelled orders", sort_dims=[2], total=False, nonnull=None, widths=[80, 85, 95, 250, 280]))
    s.nonnull("Cancelled orders")
    s = add(table(R, "O | Open orders", "Every open Wiser order line (backlog)", ["Branch", "Order", "Order date", "Partner",
                                                                                 "MaterialDescription"],
                  ["Backlog", "Backlog age (max days)"], sort_by="Backlog", sort_dims=[1], total=False, nonnull=None,
                  widths=[80, 95, 85, 250, 280]))
    s.nonnull("Backlog")

    # ---- Services (work orders; their own filters)
    add(lines_by(R, "SV | Trend", "Service requests month on month", "WO month", "Work order group", "Service requests",
                 nonnull=SVN), "SV")
    add(table(R, "SV | Branches", "Service requests and turnaround by branch", ["Service region", "Service branch"],
              ["Service requests", "Service requests PY", "Service requests growth %", "Installations", "Avg installation TAT (days)",
               "Repairs", "Avg repair TAT (days)", "Done within 3 days %", "Products serviced"], widths=[70, 100], nonnull=SVN), "SV")
    add(hbar(R, "SV | Problems", "Service requests by problem category", "Problem category", "Service requests", nonnull=SVN,
             exclude_null=True), "SV")
    add(table(R, "SV | Sub-categories", "Problem sub-categories", ["Problem category", "Sub-Category"],
              ["Service requests", "Avg repair TAT (days)", "Avg installation TAT (days)"], widths=[120, 220], nonnull=SVN,
              sort_by="Service requests"), "SV")
    add(table(R, "SV | Status", "Work order status", ["Work Order Status *"], ["Service requests", "Completed %"], widths=[170],
              nonnull=SVN), "SV")
    add(table(R, "SV | Cities", "Cities serviced", ["Service region", "City (Processed Data)"],
              ["Service requests", "Installations", "Repairs", "Avg installation TAT (days)", "Avg repair TAT (days)"],
              widths=[70, 140], nonnull=SVN, sort_by="Service requests", sort_dims=[1]), "SV")
    add(table(R, "SV | Engineers", "Wiser engineers by branch and city", ["Branch (Wiser - Engineers)", "City (Wiser - Engineers)"],
              [(R.fields["Engineers"], "sum")], widths=[90, 140], nonnull=None, units=False), "E")

    # ---- Explore
    s = add(table(R, "X | Explore table", "Explore", ["Explore by"],
                  G + ["Share of sales %", "OB", "Backlog", "Active partners", "Quantity"], sort_by="Sales", widths=[300]))
    s.title = rp.title_runs("Explore by <[Parameters].[p Dim]>")
    add(hbar(R, "X | Explore bars", "Top 30 by Wiser sales", "Explore by", "Sales", "Colour = growth", scale=("Growth %", GROWTH_SCALE),
             top=(30, "Sales CY (row)")))

    # ---- filters: Wiser scope, shared filters, row-level security
    for name, s in S.items():
        s.filter_refs = {}
        if s.scope == "P":
            for field, group, ctx in FILTERS:
                f = F(R, field)
                if s.top_field is not None and f.key == s.top_field.key:
                    continue
                if take_over(s, f):
                    s.filter_refs[field] = s.shared_filter(field, group, ctx)
            s.bool_filter("Is Wiser", context=True)
            s.bool_filter("Branch access (sales)", context=True)
        elif s.scope == "SV":
            for field, group, ctx in SV_FILTERS:
                if take_over(s, F(R, field)):
                    s.filter_refs[field] = s.shared_filter(field, group, ctx)
            s.bool_filter("Branch access (services)", context=True)
        elif s.scope == "E":
            s.bool_filter("Branch access (engineers)", context=True)
    return S


PERIOD = [("p Period", "compact"), ("p Start", "datetime"), ("p End", "datetime")]
VALUE = [("p Measure", "compact"), ("p Units", "compact"), ("p GTM", "compact")]


def params_for(dash):
    if dash == "MTD":
        return VALUE + [("p MTD view", "compact")]
    if dash == "Services":
        return PERIOD
    if dash == "Explore":
        return PERIOD + VALUE + [("p Dim", "compact")]
    return PERIOD + VALUE


def filters_for(dash):
    return SV_FILTERS if dash == "Services" else FILTERS


SIDE_NOTE = ("Filters apply to every sheet of every page. Wiser = family Wiser in the material master. Growth compares the same "
             "dates last year (green = growth, red = decline).\n"
             "Retail = pricing class EKR / ERW; the rest is Project. Yearly Wiser targets as set; monthly targets are net "
             "(gross = net / 0.83).\n"
             "Services: counts are work orders created in the period; TAT = completed - created, in days.")
VISIBILITY = {f"MV | {k}": f"MTD view: {k}" for k, _ in MTD_VIEWS}
H = lambda dash: (86, [(f"H | {dash}", 1)])
K = lambda *names: (128, [(n, 1) for n in names])

LAYOUT = {
    "Overview": [H("Overview"), K("K | Sales", "K | Yearly target", "K | Monthly target", "K | Data"),
                 K("K | Order book", "K | Backlog", "K | Partners", "K | Retail"),
                 (310, [("O | Three years", 1)]),
                 (300, [("O | Quarters", 1), ("O | GTM", 1), ("O | Version", 1), ("O | Region share", 1)])],
    "Regions & Branches": [H("Regions & Branches"), (230, [("B | Regions", 1)]), (600, [("B | Scorecard", 1)]),
                           (600, [("B | GTM targets", 1)]), (300, [("B | Trend", 1)]), (520, [("B | SPOC", 1)])],
    "Products": [K("K | Sales", "K | Quantity", "K | Order book", "K | Backlog"), (180, [("P | Versions", 1)]),
                 (300, [("P | Version trend", 1)]), (520, [("P | Materials", 1)]), (520, [("P | Material x region", 1)])],
    "Partners": [H("Partners"), K("K | Partners", "K | Per partner", "K | Lapsed", "K | Retail"),
                 (280, [("Pa | Slabs", 1), ("Pa | Type", 1)]), (230, [("Pa | Presence", 1)]),
                 (280, [("Pa | New partners", 1)]), (300, [("Pa | Top partners", 1)]), (440, [("Pa | Declining", 1)])],
    "All Partners": [(1200, [("AP | All partners", 1)]), (800, [("AP | By month", 1)])],
    "Monthly Targets": [H("Monthly Targets"), K("K | Monthly target", "K | Monthly gap", "K | Sales", "K | Data"),
                        (400, [("MT | Monthly", 1), ("MT | Chart", 1)]), (800, [("MT | Branch x month", 1)])],
    "MTD": [H("MTD"), (240, [("M | KPIs", 2), ("M | Regions", 3)]), (300, [("M | Trend", 2), ("M | Leaderboard", 1)]),
            (760, [("swap", 1, ["MV | " + k for k, _ in MTD_VIEWS])])],
    "Order Book": [H("Order Book"), K("K | Order book", "K | Cancelled", "K | Backlog", "K | Sales to OB"),
                   (320, [("O | OB trend", 2), ("O | Backlog ageing", 1)]), (520, [("O | OB by branch", 1)]),
                   (420, [("O | Cancellations", 1)]), (600, [("O | Open orders", 1)])],
    "Services": [H("Services"), K("SK | Requests", "SK | Installations", "SK | Repairs", "SK | On time"),
                 (300, [("SV | Trend", 2), ("SV | Problems", 1)]), (600, [("SV | Branches", 1)]),
                 (360, [("SV | Sub-categories", 2), ("SV | Status", 1)]), (520, [("SV | Cities", 2), ("SV | Engineers", 1)])],
    "Explore": [(700, [("X | Explore table", 1)]), (680, [("X | Explore bars", 1)])],
}
