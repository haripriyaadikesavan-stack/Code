"""Business model: parameters and calculated fields for the Retail dashboard.

Business rules (same as the existing Retail workbook):
  * Sales = rows with Type 'Sales' dated by InvoiceDate; OB = Type 'OB' by
    Calendar Day; Backlog = Type 'Backlog' (open orders, by OrderCreationDate).
  * Net Sales already carries the yearly rebate (from the notebook).
  * Monthly branch targets (Master Targets File > Monthly Targets) are NET;
    gross target = net / 0.83.
  * Partner targets (QTOD and ATOD file) are GROSS; net = gross x 0.83.
  * Annual branch targets (Branch Targets) are NET.
"""
from core import Registry

GROSS_FACTOR = 0.83


def build_model(branches, branch_zone, families, rls_rules):
    R = Registry()
    b = R.base
    # ---- primary table (Sales + OB + Backlog union, with masters joined)
    b("Type", "string", "dimension", "nominal")
    b("InvoiceDate", "datetime", "dimension", "ordinal", "Invoice Date")
    b("Calendar Day", "date", "dimension", "ordinal")
    b("OrderCreationDate", "date", "dimension", "ordinal")
    b("Gross Sales", "real", "measure", "quantitative")
    b("Net Sales", "real", "measure", "quantitative")
    b("BillingQuantity", "integer", "measure", "quantitative", "Billing Quantity")
    b("SoldToPartyID", "string", "dimension", "nominal", "Sold To Party ID")
    b("SoldToPartyName", "string", "dimension", "nominal", "Sold To Party Name")
    b("Patner (Partner Master Data)", "string", "dimension", "nominal", "Partner Name")
    b("RD/KRO", "string", "dimension", "nominal", "Partner Type (RD/KRO)")
    b("ASM", "string", "dimension", "nominal")
    b("Tier Category", "string", "dimension", "nominal", "Tier (Geography)")
    b("Onboarded Year (Partner Master Data)", "string", "dimension", "nominal", "Onboarded Year")
    b("SSA", "string", "dimension", "nominal")
    b("Zone", "string", "dimension", "nominal", "Region")
    b("Branch (Geography mapping)", "string", "dimension", "nominal", "Branch")
    b("City (Geography mapping)", "string", "dimension", "nominal", "City")
    b("Family Actual (Sheet1)", "string", "dimension", "nominal", "Family")
    b("Classification (Sheet1)", "string", "dimension", "nominal", "Product Classification")
    b("Helios Code", "string", "dimension", "nominal", "Activity (Helios)")
    b("MaterialCode", "string", "dimension", "nominal", "Material Code")
    b("MaterialDescription", "string", "dimension", "nominal", "Material Description")
    # ---- monthly targets (NET)
    b("Branch", "string", "dimension", "nominal", "Target Branch", table="T")
    b("Zone (Monthly Targets)", "string", "dimension", "nominal", "Target Region", table="T")
    b("Month and year", "date", "dimension", "ordinal", "Target Month", table="T")
    b("Target (Monthly Targets)", "integer", "measure", "quantitative", "Monthly Target (net)", table="T")
    b("Wiser", "integer", "measure", "quantitative", "Wiser Target (net)", table="T")
    b("Miluz Zeta", "integer", "measure", "quantitative", "Miluz Zeta Target (net)", table="T")
    # ---- annual branch targets (NET)
    b("Branch (Branch Targets1)", "string", "dimension", "nominal", "Annual Target Branch", table="B")
    b("Year (Branch Targets1)", "integer", "dimension", "ordinal", "Annual Target Year", table="B")
    b("Target (Branch Targets1)", "integer", "measure", "quantitative", "Annual Branch Target (net)", table="B")
    # ---- partner targets QTOD / ATOD (GROSS)
    b("CC", "string", "dimension", "nominal", "Partner Code (targets)", table="Q")
    b("Partner", "string", "dimension", "nominal", "Partner (targets)", table="Q")
    b("Branch (QTOD - ATOD)", "string", "dimension", "nominal", "Branch (targets)", table="Q")
    b("RD/ KRO", "string", "dimension", "nominal", "Partner Type (targets)", table="Q")
    b("Target Year", "integer", "dimension", "ordinal", "Target Year", table="Q")
    b("Target 100%", "integer", "measure", "quantitative", "Partner Annual Target (gross)", table="Q")
    b("Qtr 1 Target", "integer", "measure", "quantitative", table="Q")
    b("Qtr 2 Target", "string", "dimension", "nominal", table="Q")
    b("Qtr 3 Target", "string", "dimension", "nominal", table="Q")
    b("Qtr 4 Target", "string", "dimension", "nominal", table="Q")
    b("Zeta Target", "integer", "measure", "quantitative", "Partner Zeta Target (gross)", table="Q")
    b("Wiser Target", "integer", "measure", "quantitative", "Partner Wiser Target (gross)", table="Q")

    # ================================================================ params
    P = R.param
    P("p Period", "Period", "string", "YTD",
      ["YTD", "QTD", "MTD", "Last month", "Full year", "Custom range"])
    P("p Start", "Custom start", "date", "2026-01-01", rng=("2023-01-01", "2028-12-31", None))
    P("p End", "Custom end", "date", "2026-12-31", rng=("2023-01-01", "2028-12-31", None))
    P("p Measure", "Gross / Net", "string", "Gross", ["Gross", "Net"])
    P("p Units", "Units", "string", "Mn", ["INR", "Lakh", "Mn", "Cr"])
    P("p Region", "Region", "string", "All", ["All", "East", "North", "South", "West1", "West2"])
    P("p Branch", "Branch", "string", "All", ["All"] + branches)
    P("p Family", "Family", "string", "All", ["All"] + families)
    P("p Channel", "Partner type", "string", "All", ["All", "RD", "KRO", "PD"])
    P("p Basis", "Target basis", "string", "Pro-rated to date", ["Pro-rated to date", "Full months"])
    P("p Dim", "Explore by", "string", "Branch",
      ["Region", "Branch", "ASM", "Partner", "Family", "Product classification", "Activity (Helios)",
       "Tier", "Partner type", "SSA", "City", "Month", "Quarter"])

    C = R.calc
    fmt_v, fmt_p, fmt_n = "n#,##0.0;-#,##0.0", "p0.0%", "n#,##0;-#,##0"
    PU, PM = "[Parameters].[p Units]", "[Parameters].[p Measure]"

    # ================================================================ period (parameter-only, usable on any table)
    C("Units divisor", f"CASE {PU} WHEN 'INR' THEN 1 WHEN 'Lakh' THEN 100000 WHEN 'Mn' THEN 1000000 WHEN 'Cr' THEN 10000000 ELSE 1000000 END")
    C("Period start", "DATE(CASE [Parameters].[p Period] "
      "WHEN 'YTD' THEN DATETRUNC('year', TODAY()) "
      "WHEN 'QTD' THEN DATETRUNC('quarter', TODAY()) "
      "WHEN 'MTD' THEN DATETRUNC('month', TODAY()) "
      "WHEN 'Last month' THEN DATEADD('month', -1, DATETRUNC('month', TODAY())) "
      "WHEN 'Full year' THEN DATETRUNC('year', TODAY()) "
      "ELSE [Parameters].[p Start] END)", "date")
    C("Period end", "DATE(CASE [Parameters].[p Period] "
      "WHEN 'Last month' THEN DATEADD('day', -1, DATETRUNC('month', TODAY())) "
      "WHEN 'Full year' THEN DATEADD('day', -1, DATEADD('year', 1, DATETRUNC('year', TODAY()))) "
      "WHEN 'Custom range' THEN [Parameters].[p End] "
      "ELSE TODAY() END)", "date")
    C("PY start", "DATEADD('year', -1, {{Period start}})", "date")
    C("PY end", "DATEADD('year', -1, {{Period end}})", "date")
    C("Target factor", f"IF {PM} = 'Gross' THEN 1 / {GROSS_FACTOR} ELSE 1 END")   # net target -> chosen basis
    C("Partner target factor", f"IF {PM} = 'Net' THEN {GROSS_FACTOR} ELSE 1 END")  # gross partner target -> basis

    # ================================================================ primary table: scope & values
    C("Sales value (INR)", f"IF {PM} = 'Net' THEN [Net Sales] ELSE [Gross Sales] END")
    C("Value", "{{Sales value (INR)}} / {{Units divisor}}")
    C("Region (clean)", "IFNULL([Zone], 'Unmapped')", "string")
    C("Branch (clean)", "IFNULL([Branch (Geography mapping)], 'Unmapped')", "string")
    C("Family (clean)", "IFNULL([Family Actual (Sheet1)], 'Unmapped')", "string")
    C("Partner (sales)", "IFNULL([Patner (Partner Master Data)], [SoldToPartyName]) + ' (' + [SoldToPartyID] + ')'", "string")
    C("In filter scope",
      "([Parameters].[p Region] = 'All' OR [Zone] = [Parameters].[p Region]) "
      "AND ([Parameters].[p Branch] = 'All' OR [Branch (Geography mapping)] = [Parameters].[p Branch]) "
      "AND ([Parameters].[p Family] = 'All' OR [Family Actual (Sheet1)] = [Parameters].[p Family]) "
      "AND ([Parameters].[p Channel] = 'All' OR [RD/KRO] = [Parameters].[p Channel])", "boolean")
    C("Is sale in period", "[Type] = 'Sales' AND {{In filter scope}} AND DATE([InvoiceDate]) >= {{Period start}} AND DATE([InvoiceDate]) <= {{Period end}}", "boolean")
    C("Is sale in PY period", "[Type] = 'Sales' AND {{In filter scope}} AND DATE([InvoiceDate]) >= {{PY start}} AND DATE([InvoiceDate]) <= {{PY end}}", "boolean")
    C("Is OB in period", "[Type] = 'OB' AND {{In filter scope}} AND [Calendar Day] >= {{Period start}} AND [Calendar Day] <= {{Period end}}", "boolean")
    C("Is OB in PY period", "[Type] = 'OB' AND {{In filter scope}} AND [Calendar Day] >= {{PY start}} AND [Calendar Day] <= {{PY end}}", "boolean")
    C("Sales CY (row)", "IF {{Is sale in period}} THEN {{Value}} END", fmt=fmt_v)
    C("Sales PY (row)", "IF {{Is sale in PY period}} THEN {{Value}} END", fmt=fmt_v)
    C("OB CY (row)", "IF {{Is OB in period}} THEN {{Value}} END", fmt=fmt_v)
    C("OB PY (row)", "IF {{Is OB in PY period}} THEN {{Value}} END", fmt=fmt_v)
    C("Backlog (row)", "IF [Type] = 'Backlog' AND {{In filter scope}} THEN {{Value}} END", fmt=fmt_v)
    C("Qty CY (row)", "IF {{Is sale in period}} THEN [BillingQuantity] END", "integer", fmt=fmt_n)
    C("Active partner CY (row)", "IF {{Is sale in period}} AND [Gross Sales] <> 0 THEN [SoldToPartyID] END", "string")
    C("Active partner PY (row)", "IF {{Is sale in PY period}} AND [Gross Sales] <> 0 THEN [SoldToPartyID] END", "string")
    C("Has activity", "IF {{Is sale in period}} OR {{Is sale in PY period}} OR {{Is OB in period}} OR ([Type] = 'Backlog' AND {{In filter scope}}) THEN 1 END", "integer")
    C("Has sales either year", "IF {{Is sale in period}} OR {{Is sale in PY period}} THEN 1 END", "integer")
    # transaction month for trends (Sales by invoice date, OB by calendar day)
    C("Txn date", "CASE [Type] WHEN 'Sales' THEN DATE([InvoiceDate]) WHEN 'OB' THEN [Calendar Day] ELSE [OrderCreationDate] END", "date")
    C("Month", "IF MONTH({{Txn date}}) < 10 THEN '0' ELSE '' END + STR(MONTH({{Txn date}})) + ' ' + LEFT(DATENAME('month', {{Txn date}}), 3)", "string")
    C("Quarter", "'Q' + STR(DATEPART('quarter', {{Txn date}}))", "string")
    C("Day of month", "IF DAY({{Txn date}}) < 10 THEN '0' ELSE '' END + STR(DAY({{Txn date}}))", "string")
    # partner lifecycle (data starts two years back, so first sale this year = new or revived)
    C("Partner first sale", "{FIXED [SoldToPartyID] : MIN(IF [Type] = 'Sales' AND [Gross Sales] > 0 THEN DATE([InvoiceDate]) END)}", "date")
    C("Partner status", "IF {{Partner first sale}} >= DATETRUNC('year', {{Period start}}) THEN 'New / Revival' ELSE 'Existing' END", "string")
    C("Partner sales in period (INR)", "{FIXED [SoldToPartyID] : SUM(IF {{Is sale in period}} THEN {{Sales value (INR)}} END)}")
    C("Partner slab",
      "IF ISNULL({{Partner sales in period (INR)}}) OR {{Partner sales in period (INR)}} <= 0 THEN NULL "
      "ELSEIF {{Partner sales in period (INR)}} < 100000 THEN '1. Below 1 L' "
      "ELSEIF {{Partner sales in period (INR)}} < 500000 THEN '2. 1 - 5 L' "
      "ELSEIF {{Partner sales in period (INR)}} < 2500000 THEN '3. 5 - 25 L' "
      "ELSEIF {{Partner sales in period (INR)}} < 10000000 THEN '4. 25 L - 1 Cr' "
      "ELSE '5. Above 1 Cr' END", "string")
    C("Backlog age (days)", "IF [Type] = 'Backlog' THEN DATEDIFF('day', [OrderCreationDate], TODAY()) END", "integer", fmt=fmt_n)
    C("Backlog ageing",
      "IF [Type] <> 'Backlog' THEN NULL "
      "ELSEIF {{Backlog age (days)}} <= 30 THEN '1. 0-30 days' "
      "ELSEIF {{Backlog age (days)}} <= 60 THEN '2. 31-60 days' "
      "ELSEIF {{Backlog age (days)}} <= 90 THEN '3. 61-90 days' "
      "ELSE '4. Over 90 days' END", "string")
    C("Focus family", "IF [Family Actual (Sheet1)] = 'Wiser' THEN 'Wiser' ELSEIF [Family Actual (Sheet1)] = 'Miluz Zeta' THEN 'Miluz Zeta' ELSE 'Other families' END", "string")
    C("Wiser sales CY (row)", "IF {{Is sale in period}} AND [Family Actual (Sheet1)] = 'Wiser' THEN {{Value}} END", fmt=fmt_v)
    C("Zeta sales CY (row)", "IF {{Is sale in period}} AND [Family Actual (Sheet1)] = 'Miluz Zeta' THEN {{Value}} END", fmt=fmt_v)
    for q in (1, 2, 3, 4):
        C(f"Sales Q{q} (row)", f"IF [Type] = 'Sales' AND {{{{In filter scope}}}} AND YEAR([InvoiceDate]) = YEAR({{{{Period end}}}}) AND DATEPART('quarter', [InvoiceDate]) = {q} THEN {{{{Value}}}} END", fmt=fmt_v)
    C("Explore by",
      "CASE [Parameters].[p Dim] "
      "WHEN 'Region' THEN {{Region (clean)}} WHEN 'Branch' THEN {{Branch (clean)}} "
      "WHEN 'ASM' THEN IFNULL([ASM], 'Unmapped') WHEN 'Partner' THEN {{Partner (sales)}} "
      "WHEN 'Family' THEN {{Family (clean)}} WHEN 'Product classification' THEN IFNULL([Classification (Sheet1)], 'Unmapped') "
      "WHEN 'Activity (Helios)' THEN IFNULL([Helios Code], 'Unmapped') WHEN 'Tier' THEN IFNULL([Tier Category], 'Unmapped') "
      "WHEN 'Partner type' THEN IFNULL([RD/KRO], 'Unmapped') WHEN 'SSA' THEN IFNULL([SSA], 'Unmapped') "
      "WHEN 'City' THEN IFNULL([City (Geography mapping)], 'Unmapped') WHEN 'Month' THEN {{Month}} "
      "WHEN 'Quarter' THEN {{Quarter}} END", "string")
    C("SSA group", "IF ISNULL([SSA]) OR [SSA] = 'No SSA' THEN 'Without SSA' ELSE 'With SSA' END", "string")

    # ================================================================ monthly targets (table T)
    C("Target branch (clean)", "CASE [Branch] WHEN 'NCR(DEL+ROH)' THEN 'NCR' WHEN 'NCR(NOI+GHAZ)' THEN 'NOIDA-GHZ' ELSE [Branch] END", "string")
    C("Target in scope",
      "([Parameters].[p Region] = 'All' OR [Zone (Monthly Targets)] = [Parameters].[p Region]) "
      "AND ([Parameters].[p Branch] = 'All' OR {{Target branch (clean)}} = [Parameters].[p Branch]) "
      "AND [Parameters].[p Channel] = 'All'", "boolean")
    C("Target month start", "DATE(DATETRUNC('month', [Month and year]))", "date")
    C("Target month end", "DATEADD('day', -1, DATEADD('month', 1, {{Target month start}}))", "date")
    C("Target overlap days",
      "DATEDIFF('day', MAX({{Target month start}}, {{Period start}}), MIN({{Target month end}}, {{Period end}})) + 1", "integer")
    C("Target month share",
      "IF {{Target overlap days}} <= 0 THEN 0 "
      "ELSEIF [Parameters].[p Basis] = 'Full months' THEN 1 "
      "ELSE {{Target overlap days}} / (DATEDIFF('day', {{Target month start}}, {{Target month end}}) + 1) END")
    C("Target (net INR)",
      "CASE [Parameters].[p Family] WHEN 'All' THEN [Target (Monthly Targets)] "
      "WHEN 'Wiser' THEN [Wiser] WHEN 'Miluz Zeta' THEN [Miluz Zeta] END")
    C("Target (row)", "IF {{Target in scope}} THEN {{Target (net INR)}} * {{Target factor}} * {{Target month share}} / {{Units divisor}} END", fmt=fmt_v)
    C("Wiser target (row)", "IF {{Target in scope}} THEN [Wiser] * {{Target factor}} * {{Target month share}} / {{Units divisor}} END", fmt=fmt_v)
    C("Zeta target (row)", "IF {{Target in scope}} THEN [Miluz Zeta] * {{Target factor}} * {{Target month share}} / {{Units divisor}} END", fmt=fmt_v)
    C("Target month", "IF MONTH([Month and year]) < 10 THEN '0' ELSE '' END + STR(MONTH([Month and year])) + ' ' + LEFT(DATENAME('month', [Month and year]), 3)", "string")
    C("Target month (row)", "IF {{Target in scope}} AND YEAR([Month and year]) = YEAR({{Period end}}) THEN {{Target (net INR)}} * {{Target factor}} / {{Units divisor}} END", fmt=fmt_v)

    # ================================================================ annual branch targets (table B)
    zone_case = " ".join(f"WHEN '{br}' THEN '{z}'" for br, z in sorted(branch_zone.items()))
    C("Annual target region", f"CASE [Branch (Branch Targets1)] {zone_case} ELSE 'Other' END", "string")
    C("Annual target (row)",
      "IF [Year (Branch Targets1)] = YEAR({{Period end}}) "
      "AND ([Parameters].[p Region] = 'All' OR {{Annual target region}} = [Parameters].[p Region]) "
      "AND ([Parameters].[p Branch] = 'All' OR [Branch (Branch Targets1)] = [Parameters].[p Branch]) "
      "AND [Parameters].[p Family] = 'All' AND [Parameters].[p Channel] = 'All' "
      "THEN [Target (Branch Targets1)] * {{Target factor}} / {{Units divisor}} END", fmt=fmt_v)

    # ================================================================ partner targets (table Q)
    C("Target partner branch",
      "CASE UPPER(TRIM([Branch (QTOD - ATOD)])) WHEN 'TN -1' THEN 'TN-1' WHEN 'TN2' THEN 'TN-2' "
      "ELSE UPPER(TRIM([Branch (QTOD - ATOD)])) END", "string")
    qzone = " ".join(f"WHEN '{br}' THEN '{z}'" for br, z in sorted(branch_zone.items()))
    C("Target partner region", f"CASE {{{{Target partner branch}}}} {qzone} ELSE 'Other' END", "string")
    C("Target partner", "IFNULL([Partner], 'Unknown') + ' (' + [CC] + ')'", "string")
    C("Partner target in scope",
      "[Target Year] = YEAR({{Period end}}) "
      "AND ([Parameters].[p Region] = 'All' OR {{Target partner region}} = [Parameters].[p Region]) "
      "AND ([Parameters].[p Branch] = 'All' OR {{Target partner branch}} = [Parameters].[p Branch]) "
      "AND ([Parameters].[p Channel] = 'All' OR TRIM([RD/ KRO]) = [Parameters].[p Channel])", "boolean")
    C("Partner annual target (row)", "IF {{Partner target in scope}} THEN [Target 100%] * {{Partner target factor}} / {{Units divisor}} END", fmt=fmt_v)
    C("Partner Q1 target (row)", "IF {{Partner target in scope}} THEN [Qtr 1 Target] * {{Partner target factor}} / {{Units divisor}} END", fmt=fmt_v)
    for q in (2, 3, 4):
        C(f"Partner Q{q} target (row)", f"IF {{{{Partner target in scope}}}} THEN FLOAT([Qtr {q} Target]) * {{{{Partner target factor}}}} / {{{{Units divisor}}}} END", fmt=fmt_v)
    C("Partner Wiser target (row)", "IF {{Partner target in scope}} THEN [Wiser Target] * {{Partner target factor}} / {{Units divisor}} END", fmt=fmt_v)
    C("Partner Zeta target (row)", "IF {{Partner target in scope}} THEN [Zeta Target] * {{Partner target factor}} / {{Units divisor}} END", fmt=fmt_v)

    # ================================================================ row-level security (from the 'Branch RLS' user filter)
    def rls_expr(field):
        parts = []
        for group, members, include_null in rls_rules:
            cond = []
            if members:
                cond.append(f"{field} IN (" + ", ".join("'" + m + "'" for m in members) + ")")
            if include_null:
                cond.append(f"ISNULL({field})")
            if cond:
                parts.append(f"(ISMEMBEROF('{group}') AND ({' OR '.join(cond)}))")
        return " OR ".join(parts) if parts else "FALSE"
    C("Branch access (sales)", rls_expr("[Branch (Geography mapping)]"), "boolean")
    C("Branch access (partner targets)", rls_expr("{{Target partner branch}}"), "boolean")

    # ================================================================ aggregates
    A = lambda cap, f, dt="real", fmt=fmt_v: C(cap, f, dt, role="measure" if dt in ("real", "integer") else "dimension", fmt=fmt)
    A("Sales", "SUM({{Sales CY (row)}})")
    A("Sales PY", "SUM({{Sales PY (row)}})")
    A("Growth %", "IF SUM({{Sales PY (row)}}) <> 0 THEN (SUM({{Sales CY (row)}}) - SUM({{Sales PY (row)}})) / ABS(SUM({{Sales PY (row)}})) END", fmt=fmt_p)
    A("Growth", "SUM({{Sales CY (row)}}) - SUM({{Sales PY (row)}})")
    A("Target", "SUM({{Target (row)}})")
    A("Ach %", "IF SUM({{Target (row)}}) > 0 THEN SUM({{Sales CY (row)}}) / SUM({{Target (row)}}) END", fmt=fmt_p)
    A("Gap to target", "SUM({{Sales CY (row)}}) - SUM({{Target (row)}})")
    A("Annual target", "SUM({{Annual target (row)}})")
    A("Annual Ach %", "IF SUM({{Annual target (row)}}) > 0 THEN SUM({{Sales CY (row)}}) / SUM({{Annual target (row)}}) END", fmt=fmt_p)
    A("Balance to annual target", "SUM({{Annual target (row)}}) - SUM({{Sales CY (row)}})")
    A("Current run-rate / month", "SUM({{Sales CY (row)}}) / MAX(1, DATEDIFF('day', MIN({{Period start}}), MIN(MIN({{Period end}}), TODAY())) + 1) * 30.4")
    A("Required run-rate / month",
      "IF SUM({{Annual target (row)}}) - SUM({{Sales CY (row)}}) > 0 THEN "
      "(SUM({{Annual target (row)}}) - SUM({{Sales CY (row)}})) / MAX(0.5, DATEDIFF('day', TODAY(), DATEADD('day', -1, DATEADD('year', 1, DATETRUNC('year', TODAY()))))  / 30.4) "
      "ELSE 0 END")
    A("OB", "SUM({{OB CY (row)}})")
    A("OB PY", "SUM({{OB PY (row)}})")
    A("OB growth %", "IF SUM({{OB PY (row)}}) <> 0 THEN (SUM({{OB CY (row)}}) - SUM({{OB PY (row)}})) / ABS(SUM({{OB PY (row)}})) END", fmt=fmt_p)
    A("Sales to OB %", "IF SUM({{OB CY (row)}}) <> 0 THEN SUM({{Sales CY (row)}}) / SUM({{OB CY (row)}}) END", fmt=fmt_p)
    A("Backlog", "SUM({{Backlog (row)}})")
    A("Backlog over 60 days", "SUM(IF {{Backlog age (days)}} > 60 THEN {{Backlog (row)}} END)")
    A("Active partners", "COUNTD({{Active partner CY (row)}})", "integer", fmt_n)
    A("Active partners PY", "COUNTD({{Active partner PY (row)}})", "integer", fmt_n)
    A("Partner change", "COUNTD({{Active partner CY (row)}}) - COUNTD({{Active partner PY (row)}})", "integer", fmt_n)
    A("New / revival partners", "COUNTD(IF {{Partner status}} = 'New / Revival' THEN {{Active partner CY (row)}} END)", "integer", fmt_n)
    A("New partner sales", "SUM(IF {{Partner status}} = 'New / Revival' THEN {{Sales CY (row)}} END)")
    A("Sales per partner", "IF COUNTD({{Active partner CY (row)}}) > 0 THEN SUM({{Sales CY (row)}}) / COUNTD({{Active partner CY (row)}}) END")
    A("Quantity", "SUM({{Qty CY (row)}})", "integer", fmt_n)
    A("Wiser sales", "SUM({{Wiser sales CY (row)}})")
    A("Zeta sales", "SUM({{Zeta sales CY (row)}})")
    A("Wiser target", "SUM({{Wiser target (row)}})")
    A("Zeta target", "SUM({{Zeta target (row)}})")
    A("Wiser Ach %", "IF SUM({{Wiser target (row)}}) > 0 THEN SUM({{Wiser sales CY (row)}}) / SUM({{Wiser target (row)}}) END", fmt=fmt_p)
    A("Zeta Ach %", "IF SUM({{Zeta target (row)}}) > 0 THEN SUM({{Zeta sales CY (row)}}) / SUM({{Zeta target (row)}}) END", fmt=fmt_p)
    A("Monthly target", "SUM({{Target month (row)}})")
    A("Share of sales %", "SUM({{Sales CY (row)}}) / MIN({FIXED : SUM({{Sales CY (row)}})})", fmt=fmt_p)
    A("Partners (count)", "COUNTD({{Active partner CY (row)}})", "integer", fmt_n)
    C("Data as of", "STR(MAX(IF [Type] = 'Sales' THEN DATE([InvoiceDate]) END))", "string", role="measure", ftype="nominal")
    A("Has activity (agg)", "SUM({{Has activity}})", "integer", fmt_n)
    A("Has sales (agg)", "SUM({{Has sales either year}})", "integer", fmt_n)
    # partner targets vs achievement
    A("Partner annual target", "SUM({{Partner annual target (row)}})")
    A("Partner Ach %", "IF SUM({{Partner annual target (row)}}) > 0 THEN SUM({{Sales CY (row)}}) / SUM({{Partner annual target (row)}}) END", fmt=fmt_p)
    A("Partner balance", "SUM({{Partner annual target (row)}}) - SUM({{Sales CY (row)}})")
    for q in (1, 2, 3, 4):
        A(f"Q{q} target", f"SUM({{{{Partner Q{q} target (row)}}}})")
        A(f"Q{q} sales", f"SUM({{{{Sales Q{q} (row)}}}})")
        A(f"Q{q} Ach %", f"IF SUM({{{{Partner Q{q} target (row)}}}}) > 0 THEN SUM({{{{Sales Q{q} (row)}}}}) / SUM({{{{Partner Q{q} target (row)}}}}) END", fmt=fmt_p)
    A("Partner Wiser target", "SUM({{Partner Wiser target (row)}})")
    A("Partner Zeta target", "SUM({{Partner Zeta target (row)}})")
    A("Partner Wiser Ach %", "IF SUM({{Partner Wiser target (row)}}) > 0 THEN SUM({{Wiser sales CY (row)}}) / SUM({{Partner Wiser target (row)}}) END", fmt=fmt_p)
    A("Partner Zeta Ach %", "IF SUM({{Partner Zeta target (row)}}) > 0 THEN SUM({{Zeta sales CY (row)}}) / SUM({{Partner Zeta target (row)}}) END", fmt=fmt_p)
    A("Has partner target", "SUM(IF {{Partner target in scope}} THEN 1 END)", "integer", fmt_n)
    # bands (discrete, aggregate)
    band = lambda m: (f"IF ISNULL({m}) THEN 'No target' ELSEIF {m} >= 1 THEN '1. 100% and above' "
                      f"ELSEIF {m} >= 0.9 THEN '2. 90-100%' ELSEIF {m} >= 0.75 THEN '3. 75-90%' ELSE '4. Below 75%' END")
    C("Ach band", band("{{Ach %}}"), "string", role="measure", ftype="nominal")
    C("Partner Ach band", band("{{Partner Ach %}}"), "string", role="measure", ftype="nominal")
    C("Growth band", "IF ISNULL({{Growth %}}) THEN 'New' ELSEIF {{Growth %}} >= 0 THEN 'Growing' ELSE 'Declining' END", "string", role="measure", ftype="nominal")
    # KPI text helpers
    C("Growth text", "IF ISNULL({{Growth %}}) THEN 'no sales last year' ELSE (IF {{Growth %}} >= 0 THEN '+' ELSE '' END) + STR(ROUND({{Growth %}} * 100, 1)) + '% vs last year' END", "string", role="measure", ftype="nominal")
    C("Ach text", "IF ISNULL({{Ach %}}) THEN 'no target for this selection' ELSE STR(ROUND({{Ach %}} * 100, 1)) + '% of target' END", "string", role="measure", ftype="nominal")
    C("Period text", "STR(MIN({{Period start}})) + '  to  ' + STR(MIN({{Period end}}))", "string", role="measure", ftype="nominal")
    C("Nav count", "COUNT([Type])", "integer", role="measure", fmt=fmt_n)
    return R
