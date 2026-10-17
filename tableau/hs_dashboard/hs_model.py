"""Business model of the Home Solutions (HS) dashboard: parameters and calculated fields.

Rules (as in the Home Solutions workbook):
  * Sales = Type 'Sales' by InvoiceDate; OB = Type 'OB' by Calendar Day; Backlog = Type 'Backlog' (open orders).
  * Activity: HeliosCode grouped into WD (HDCWD, HSCWD), FD (HDCFD, PPCFD), COS, SPE, EVC (PREVC).
  * GTM: Retail when the pricing class is EKR or ERW, or the activity is FD; otherwise Project.
  * Net = Gross x (1 - rebate). The rebate depends on GTM and year: this year's rate for this year's
    rows, last year's rate for earlier rows. In the 'editable' workbook the four rates are parameters
    the user can type; in the 'static' workbook they are fixed (REBATES below).
  * Monthly and yearly branch targets are net; gross = net / 0.83. Commits are net; gross = net / 0.87.
    Partner targets are gross. Product (family) targets are used as they are.
  * 'Without FD' removes FD sales, OB, backlog and FD commits / yearly targets.
"""
from core import Registry

# rebate rates: (retail this year, project this year, retail last year, project last year)
REBATES = {"retail_cy": 0.13, "project_cy": 0.051, "retail_py": 0.122, "project_py": 0.045}
TARGET_NET = 0.83      # monthly / yearly targets: gross = net / 0.83
COMMIT_NET = 0.87      # activity commits: gross = net / 0.87
MTD_VIEWS = [("Branch", "Branch wise Sales, OB and Ach%"), ("Activity", "Branch & Activity wise OB and Sales"),
             ("ASM", "ASM wise OB and Sales"), ("Partner", "Partner wise OB and Sales"),
             ("Product", "Product wise Sales, OB, Backlog")]
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
FMT_V, FMT_P, FMT_N = "n#,##0.0;-#,##0.0", "p0.0%", "n#,##0;-#,##0"
ACTIVITIES = {"WD": ["HDCWD", "HSCWD"], "FD": ["HDCFD", "PPCFD"], "COS": ["HDCOS", "HSCOS"], "SPE": ["HDSPE", "HSSPE"],
              "EVC": ["PREVC"]}


def activity_expr(col):
    return ("CASE " + col + " " + " ".join(f"WHEN '{c}' THEN '{a}'" for a, cs in ACTIVITIES.items() for c in cs)
            + " ELSE 'Other' END")


def build_model(rls_rules, editable=True):
    R = Registry()
    b = R.base
    # ---- primary table (Sales + OB + Backlog union, with masters joined)
    b("Type", "string", "dimension", "nominal")
    b("InvoiceDate", "datetime", "dimension", "ordinal", "Invoice Date")
    b("Calendar Day", "date", "dimension", "ordinal")
    b("OrderCreationDate", "date", "dimension", "ordinal")
    b("Gross Sales", "real", "measure", "quantitative")
    b("Net Sales", "real", "measure", "quantitative", "Net Sales (source)")
    b("BillingQuantity", "integer", "measure", "quantitative", "Billing Quantity")
    b("HeliosCode", "string", "dimension", "nominal", "Activity code")
    b("PricingClassL3", "string", "dimension", "nominal", "Pricing class")
    b("SoldToPartyID", "string", "dimension", "nominal", "Sold To Party ID")
    b("SoldToPartyName", "string", "dimension", "nominal", "Sold To Party Name")
    b("Patner (Partner Master Data)", "string", "dimension", "nominal", "Partner Name")
    b("Sold-to Party (Partner Master Data)", "string", "dimension", "nominal", "Partner Code")
    b("RD/KRO", "string", "dimension", "nominal", "Partner Type")
    b("ASM", "string", "dimension", "nominal")
    b("Tier Category", "string", "dimension", "nominal", "Tier")
    b("SSA", "string", "dimension", "nominal", "SSP / SSA")
    b("Zone", "string", "dimension", "nominal", "Zone (geography)")
    b("Branch (Geography mapping)", "string", "dimension", "nominal", "Branch (geography)")
    b("City (Geography mapping)", "string", "dimension", "nominal", "City")
    b("Family Actual (Sheet1)", "string", "dimension", "nominal", "Family Actual")
    b("Classification (Sheet1)", "string", "dimension", "nominal", "Product Classification")
    b("MaterialCode", "string", "dimension", "nominal", "Material Code")
    b("MaterialDescription", "string", "dimension", "nominal", "Material")
    b("SalesOrderHeader", "integer", "dimension", "nominal", "Sales Order")
    b("SalesOrderItem", "string", "dimension", "nominal", "Order Item")
    b("PlantName", "string", "dimension", "nominal", "Plant")
    # ---- monthly targets (by branch and Retail / Project; NET)
    b("Branch", "string", "dimension", "nominal", "Monthly Target Branch", table="T")
    b("Month and year", "date", "dimension", "ordinal", "Target Month", table="T")
    b("Category (Monthly Targets)", "string", "dimension", "nominal", "Monthly Target GTM", table="T")
    b("Target (Monthly Targets)", "integer", "measure", "quantitative", "Monthly Target (net)", table="T")
    b("Wiser", "integer", "measure", "quantitative", "Wiser Monthly Target (net)", table="T")
    b("Miluz Zeta", "integer", "measure", "quantitative", "Miluz Zeta Monthly Target (net)", table="T")
    # ---- yearly branch targets (by branch, Retail / Project and activity; NET)
    b("Branch (Branch Targets1)", "string", "dimension", "nominal", "Yearly Target Branch", table="B")
    b("Category (Branch Targets1)", "string", "dimension", "nominal", "Yearly Target GTM", table="B")
    b("Helios Code (Branch Targets1)", "string", "dimension", "nominal", "Yearly Target Activity", table="B")
    b("Year (Branch Targets1)", "integer", "dimension", "ordinal", "Yearly Target Year", table="B")
    b("Target (Branch Targets1)", "integer", "measure", "quantitative", "Yearly Branch Target (net)", table="B")
    # ---- activity-wise month commits (by branch, Retail / Project, activity and month; NET)
    b("Branch (Activity wise Month Target)", "string", "dimension", "nominal", "Commit Branch", table="A")
    b("Month and year (Activity wise Month Target)", "date", "dimension", "ordinal", "Commit Month", table="A")
    b("Category (Activity wise Month Target)", "string", "dimension", "nominal", "Commit GTM", table="A")
    b("Helios Code (Activity wise Month Target)", "string", "dimension", "nominal", "Commit Activity", table="A")
    b("Commit", "integer", "measure", "quantitative", "Sales Commit (net)", table="A")
    b("OB Commit", "real", "measure", "quantitative", "OB Commit (net)", table="A")
    # ---- partner targets (pivoted by year; GROSS)
    b("Pivot Field Names", "integer", "dimension", "ordinal", "Partner Target Year", table="PT")
    b("Pivot Field Values", "real", "measure", "quantitative", "Partner Target (gross)", table="PT")
    # ---- product (family) targets by region and year
    b("Family", "string", "dimension", "nominal", "Product Target Family", table="PR")
    b("Region", "string", "dimension", "nominal", "Product Target Region", table="PR")
    b("Year (Product Targets)", "integer", "dimension", "ordinal", "Product Target Year", table="PR")
    b("Family Target", "real", "measure", "quantitative", "Family Target", table="PR")

    # ================================================================ params
    P = R.param
    P("p Period", "Period", "string", "YTD", ["YTD", "QTD", "MTD", "Last month", "Full year", "Custom range"])
    P("p Start", "Custom start", "date", "2026-01-01", rng=("2023-01-01", "2028-12-31", None))
    P("p End", "Custom end", "date", "2026-12-31", rng=("2023-01-01", "2028-12-31", None))
    P("p Measure", "Gross / Net", "string", "Gross", ["Gross", "Net"])
    P("p Units", "Units", "string", "Mn", ["INR", "Lakh", "Mn", "Cr"])
    P("p FD", "With / without FD", "string", "With FD", ["With FD", "Without FD"])
    P("p GTM", "Business (GTM)", "string", "All", ["All", "Retail", "Project"])
    P("p MTD view", "Select a view", "string", MTD_VIEWS[0][1], [v for _, v in MTD_VIEWS])
    P("p Dim", "Explore by", "string", "Branch",
      ["Region", "Branch", "ASM", "Partner", "Activity", "GTM", "Family", "Product classification", "Pricing class",
       "Tier", "Partner type", "SSP / SSA", "City", "Month", "Quarter", "New product category", "Partner status"])
    if editable:
        P("p Rebate retail CY", "Retail rebate - this year", "real", REBATES["retail_cy"])
        P("p Rebate project CY", "Project rebate - this year", "real", REBATES["project_cy"])
        P("p Rebate retail PY", "Retail rebate - last year", "real", REBATES["retail_py"])
        P("p Rebate project PY", "Project rebate - last year", "real", REBATES["project_py"])
        RB = {k: f"[Parameters].[p Rebate {k.split('_')[0]} {k.split('_')[1].upper()}]" for k in REBATES}
    else:
        RB = {k: str(v) for k, v in REBATES.items()}

    C = R.calc
    PU, PM = "[Parameters].[p Units]", "[Parameters].[p Measure]"

    # ================================================================ parameter-only helpers
    C("Units divisor", f"CASE {PU} WHEN 'INR' THEN 1 WHEN 'Lakh' THEN 100000 WHEN 'Mn' THEN 1000000 WHEN 'Cr' THEN 10000000 ELSE 1000000 END")
    C("Period start", "DATE(CASE [Parameters].[p Period] "
      "WHEN 'YTD' THEN DATETRUNC('year', TODAY()) WHEN 'QTD' THEN DATETRUNC('quarter', TODAY()) "
      "WHEN 'MTD' THEN DATETRUNC('month', TODAY()) WHEN 'Last month' THEN DATEADD('month', -1, DATETRUNC('month', TODAY())) "
      "WHEN 'Full year' THEN DATETRUNC('year', TODAY()) ELSE [Parameters].[p Start] END)", "date")
    C("Period end", "DATE(CASE [Parameters].[p Period] "
      "WHEN 'Last month' THEN DATEADD('day', -1, DATETRUNC('month', TODAY())) "
      "WHEN 'Full year' THEN DATEADD('day', -1, DATEADD('year', 1, DATETRUNC('year', TODAY()))) "
      "WHEN 'Custom range' THEN [Parameters].[p End] ELSE TODAY() END)", "date")
    C("PY start", "DATEADD('year', -1, {{Period start}})", "date")
    C("PY end", "DATEADD('year', -1, {{Period end}})", "date")
    C("PY2 start", "DATEADD('year', -2, {{Period start}})", "date")
    C("PY2 end", "DATEADD('year', -2, {{Period end}})", "date")
    C("MTD start", "DATE(DATETRUNC('month', TODAY()))", "date")
    C("MTD LY start", "DATEADD('year', -1, {{MTD start}})", "date")
    C("MTD LY end", "DATEADD('year', -1, TODAY())", "date")
    C("Current year", "YEAR({{Period end}})", "integer", fmt=FMT_N)
    C("Target net to gross", f"IF {PM} = 'Gross' THEN 1 / {TARGET_NET} ELSE 1 END")
    C("Commit net to gross", f"IF {PM} = 'Gross' THEN 1 / {COMMIT_NET} ELSE 1 END")
    C("Without FD", "[Parameters].[p FD] = 'Without FD'", "boolean")

    # ================================================================ primary table
    C("Activity", activity_expr("[HeliosCode]"), "string")
    C("GTM", "IF [PricingClassL3] = 'EKR' OR [PricingClassL3] = 'ERW' THEN 'Retail' "
      "ELSEIF {{Activity}} <> 'FD' THEN 'Project' ELSE 'Retail' END", "string")
    C("In scope (FD / GTM)", "(NOT {{Without FD}} OR {{Activity}} <> 'FD') AND "
      "([Parameters].[p GTM] = 'All' OR {{GTM}} = [Parameters].[p GTM])", "boolean")
    C("Txn date", "CASE [Type] WHEN 'Sales' THEN DATE([InvoiceDate]) WHEN 'OB' THEN [Calendar Day] ELSE [OrderCreationDate] END", "date")
    C("Rebate rate",
      f"IF [Type] = 'Backlog' OR YEAR({{{{Txn date}}}}) >= {{{{Current year}}}} THEN "
      f"(IF {{{{GTM}}}} = 'Retail' THEN {RB['retail_cy']} ELSE {RB['project_cy']} END) "
      f"ELSE (IF {{{{GTM}}}} = 'Retail' THEN {RB['retail_py']} ELSE {RB['project_py']} END) END", fmt=FMT_P)
    C("Sales value (INR)", f"IF {PM} = 'Net' THEN [Gross Sales] * (1 - {{{{Rebate rate}}}}) ELSE [Gross Sales] END")
    C("Value", "{{Sales value (INR)}} / {{Units divisor}}")
    C("Region", "IFNULL([Zone], 'Unmapped')", "string")
    C("Branch", "IFNULL([Branch (Geography mapping)], 'Unmapped')", "string")
    C("Family", "IFNULL([Family Actual (Sheet1)], 'Unmapped')", "string")
    C("Partner", "IFNULL([Patner (Partner Master Data)], [SoldToPartyName]) + ' (' + [SoldToPartyID] + ')'", "string")
    C("Is sale in period", "[Type] = 'Sales' AND DATE([InvoiceDate]) >= {{Period start}} AND DATE([InvoiceDate]) <= {{Period end}}", "boolean")
    C("Is sale in PY period", "[Type] = 'Sales' AND DATE([InvoiceDate]) >= {{PY start}} AND DATE([InvoiceDate]) <= {{PY end}}", "boolean")
    C("Is OB in period", "[Type] = 'OB' AND [Calendar Day] >= {{Period start}} AND [Calendar Day] <= {{Period end}}", "boolean")
    C("Is OB in PY period", "[Type] = 'OB' AND [Calendar Day] >= {{PY start}} AND [Calendar Day] <= {{PY end}}", "boolean")
    C("Is sale MTD", "[Type] = 'Sales' AND DATE([InvoiceDate]) >= {{MTD start}} AND DATE([InvoiceDate]) <= TODAY()", "boolean")
    C("Is sale MTD LY", "[Type] = 'Sales' AND DATE([InvoiceDate]) >= {{MTD LY start}} AND DATE([InvoiceDate]) <= {{MTD LY end}}", "boolean")
    C("Is OB MTD", "[Type] = 'OB' AND [Calendar Day] >= {{MTD start}} AND [Calendar Day] <= TODAY()", "boolean")
    C("Sales CY (row)", "IF {{Is sale in period}} THEN {{Value}} END", fmt=FMT_V)
    C("Sales PY (row)", "IF {{Is sale in PY period}} THEN {{Value}} END", fmt=FMT_V)
    C("Sales PY2 (row)", "IF [Type] = 'Sales' AND DATE([InvoiceDate]) >= {{PY2 start}} AND DATE([InvoiceDate]) <= {{PY2 end}} THEN {{Value}} END", fmt=FMT_V)
    C("OB CY (row)", "IF {{Is OB in period}} THEN {{Value}} END", fmt=FMT_V)
    C("OB PY (row)", "IF {{Is OB in PY period}} THEN {{Value}} END", fmt=FMT_V)
    C("Sales MTD (row)", "IF {{Is sale MTD}} THEN {{Value}} END", fmt=FMT_V)
    C("Sales MTD LY (row)", "IF {{Is sale MTD LY}} THEN {{Value}} END", fmt=FMT_V)
    C("OB MTD (row)", "IF {{Is OB MTD}} THEN {{Value}} END", fmt=FMT_V)
    C("Last sales date", "{FIXED : MAX(IF [Type] = 'Sales' THEN DATE([InvoiceDate]) END)}", "date")
    C("Last OB date", "{FIXED : MAX(IF [Type] = 'OB' THEN [Calendar Day] END)}", "date")
    C("Sales last billing day (row)", "IF [Type] = 'Sales' AND DATE([InvoiceDate]) = {{Last sales date}} THEN {{Value}} END", fmt=FMT_V)
    C("OB last order day (row)", "IF [Type] = 'OB' AND [Calendar Day] = {{Last OB date}} THEN {{Value}} END", fmt=FMT_V)
    C("Backlog (row)", "IF [Type] = 'Backlog' THEN {{Value}} END", fmt=FMT_V)
    C("Backlog MTD (row)", "IF [Type] = 'Backlog' AND [OrderCreationDate] >= {{MTD start}} THEN {{Value}} END", fmt=FMT_V)
    C("Qty CY (row)", "IF {{Is sale in period}} THEN [BillingQuantity] END", "integer", fmt=FMT_N)
    C("Qty PY (row)", "IF {{Is sale in PY period}} THEN [BillingQuantity] END", "integer", fmt=FMT_N)
    C("Qty MTD (row)", "IF {{Is sale MTD}} THEN [BillingQuantity] END", "integer", fmt=FMT_N)
    C("OB qty MTD (row)", "IF {{Is OB MTD}} THEN [BillingQuantity] END", "integer", fmt=FMT_N)
    C("Active partner CY (row)", "IF {{Is sale in period}} AND [Gross Sales] <> 0 THEN [SoldToPartyID] END", "string")
    C("Active partner PY (row)", "IF {{Is sale in PY period}} AND [Gross Sales] <> 0 THEN [SoldToPartyID] END", "string")
    C("Active partner MTD (row)", "IF {{Is sale MTD}} AND [Gross Sales] <> 0 THEN [SoldToPartyID] END", "string")
    C("Has activity", "IF {{Is sale in period}} OR {{Is sale in PY period}} OR {{Is OB in period}} OR [Type] = 'Backlog' THEN 1 END", "integer")
    C("Has sales either year", "IF {{Is sale in period}} OR {{Is sale in PY period}} THEN 1 END", "integer")
    C("Has MTD activity", "IF {{Is sale MTD}} OR {{Is sale MTD LY}} OR {{Is OB MTD}} THEN 1 END", "integer")
    C("Month", "IF MONTH({{Txn date}}) < 10 THEN '0' ELSE '' END + STR(MONTH({{Txn date}})) + ' ' + LEFT(DATENAME('month', {{Txn date}}), 3)", "string")
    C("Quarter", "'Q' + STR(DATEPART('quarter', {{Txn date}}))", "string")
    C("Day", "IF DAY({{Txn date}}) < 10 THEN '0' ELSE '' END + STR(DAY({{Txn date}}))", "string")
    C("Partner sales in period (INR)", "{FIXED [SoldToPartyID] : SUM(IF {{Is sale in period}} THEN {{Sales value (INR)}} END)}")
    C("Partner sales PY (INR)", "{FIXED [SoldToPartyID] : SUM(IF {{Is sale in PY period}} THEN {{Sales value (INR)}} END)}")
    C("Partner slab",
      "IF ISNULL({{Partner sales in period (INR)}}) OR {{Partner sales in period (INR)}} <= 0 THEN NULL "
      "ELSEIF {{Partner sales in period (INR)}} < 100000 THEN '1. Below 1 L' "
      "ELSEIF {{Partner sales in period (INR)}} < 500000 THEN '2. 1 - 5 L' "
      "ELSEIF {{Partner sales in period (INR)}} < 2500000 THEN '3. 5 - 25 L' "
      "ELSEIF {{Partner sales in period (INR)}} < 10000000 THEN '4. 25 L - 1 Cr' ELSE '5. Above 1 Cr' END", "string")
    C("Partner sales CY (growth lod)", "{FIXED [SoldToPartyID] : SUM({{Sales CY (row)}})}", fmt=FMT_V)
    C("Partner sales PY (growth lod)", "{FIXED [SoldToPartyID] : SUM({{Sales PY (row)}})}", fmt=FMT_V)
    C("Partner growth band",
      "IF ISNULL([SoldToPartyID]) THEN NULL "
      "ELSEIF ZN({{Partner sales PY (growth lod)}}) <= 0 AND ZN({{Partner sales CY (growth lod)}}) > 0 THEN '5. New this period (no sales last year)' "
      "ELSEIF ZN({{Partner sales PY (growth lod)}}) <= 0 THEN '7. No sales either year' "
      "ELSEIF ZN({{Partner sales CY (growth lod)}}) <= 0 THEN '6. No sales this period' "
      "ELSEIF {{Partner sales CY (growth lod)}} / {{Partner sales PY (growth lod)}} - 1 >= 0.25 THEN '1. Growth 25% and above' "
      "ELSEIF {{Partner sales CY (growth lod)}} >= {{Partner sales PY (growth lod)}} THEN '2. Growth 0-25%' "
      "ELSEIF {{Partner sales CY (growth lod)}} / {{Partner sales PY (growth lod)}} - 1 >= -0.25 THEN '3. Decline 0-25%' "
      "ELSE '4. Decline over 25%' END", "string")
    for k, lab in ((0, "CY"), (1, "PY"), (2, "PY2")):
        C(f"Partner gross {lab} (lod)",
          f"{{FIXED [SoldToPartyID] : SUM(IF [Type] = 'Sales' AND [HeliosCode] <> 'PREVC' AND YEAR([InvoiceDate]) = {{{{Current year}}}} - {k} THEN [Gross Sales] END)}}")
    C("Partner new CY", "ZN({{Partner gross CY (lod)}}) > 0 AND ZN({{Partner gross PY (lod)}}) = 0 AND ZN({{Partner gross PY2 (lod)}}) = 0", "boolean")
    C("Partner status", "IF {{Partner new CY}} THEN 'New / Revival' ELSE 'Existing' END", "string")
    C("New partner CY (row)", "IF {{Is sale in period}} AND {{Partner new CY}} AND [Gross Sales] <> 0 THEN [SoldToPartyID] END", "string")
    C("Lapsed partner (row)", "IF (ISNULL({{Partner sales in period (INR)}}) OR {{Partner sales in period (INR)}} <= 0) AND {{Partner sales PY (INR)}} > 0 THEN [SoldToPartyID] END", "string")
    C("Backlog age (days)", "IF [Type] = 'Backlog' THEN DATEDIFF('day', [OrderCreationDate], TODAY()) END", "integer", fmt=FMT_N)
    C("Backlog ageing",
      "IF [Type] <> 'Backlog' THEN NULL ELSEIF {{Backlog age (days)}} < 20 THEN '1. Under 20 days' "
      "ELSEIF {{Backlog age (days)}} <= 50 THEN '2. 20-50 days' ELSEIF {{Backlog age (days)}} <= 100 THEN '3. 51-100 days' "
      "ELSE '4. Over 100 days' END", "string")
    for fam, key in (("Wiser", "Wiser"), ("Miluz Zeta", "Zeta"), ("Miluz Lara", "Lara")):
        C(f"{key} sales CY (row)", f"IF {{{{Is sale in period}}}} AND [Family Actual (Sheet1)] = '{fam}' THEN {{{{Value}}}} END", fmt=FMT_V)
        C(f"{key} sales PY (row)", f"IF {{{{Is sale in PY period}}}} AND [Family Actual (Sheet1)] = '{fam}' THEN {{{{Value}}}} END", fmt=FMT_V)
        C(f"{key} sales MTD (row)", f"IF {{{{Is sale MTD}}}} AND [Family Actual (Sheet1)] = '{fam}' THEN {{{{Value}}}} END", fmt=FMT_V)
    for q in (1, 2, 3, 4):
        C(f"Sales Q{q} (row)", f"IF [Type] = 'Sales' AND YEAR([InvoiceDate]) = {{{{Current year}}}} AND DATEPART('quarter', [InvoiceDate]) = {q} THEN {{{{Value}}}} END", fmt=FMT_V)
    C("Retail sales CY (row)", "IF {{GTM}} = 'Retail' THEN {{Sales CY (row)}} END", fmt=FMT_V)
    C("Project sales CY (row)", "IF {{GTM}} = 'Project' THEN {{Sales CY (row)}} END", fmt=FMT_V)
    C("FD sales CY (row)", "IF {{Activity}} = 'FD' THEN {{Sales CY (row)}} END", fmt=FMT_V)
    C("SSA group", "IF ISNULL([SSA]) OR [SSA] = 'No SSA' THEN 'Without SSP' ELSE 'With SSP' END", "string")
    C("New product category",
      "IF CONTAINS([Family Actual (Sheet1)], 'Miluz Zeta') THEN 'Miluz Zeta' ELSEIF CONTAINS([Family Actual (Sheet1)], 'Wiser') THEN 'Wiser' "
      "ELSEIF CONTAINS([Family Actual (Sheet1)], 'Miluz Lara') THEN 'Miluz Lara' ELSEIF CONTAINS([Family Actual (Sheet1)], 'Unica Pure Plus') THEN 'Unica Pure Plus' "
      "ELSEIF CONTAINS([Family Actual (Sheet1)], 'UnicaX') THEN 'UnicaX' ELSE 'Other families' END", "string")
    C("Sale date (row)", "IF [Type] = 'Sales' AND [Gross Sales] > 0 THEN DATE([InvoiceDate]) END", "date")
    C("Order", "STR([SalesOrderHeader])", "string")
    C("Order date", "IF [Type] = 'Backlog' THEN STR([OrderCreationDate]) END", "string")
    C("Active city (row)", "IF {{Is sale in period}} AND [Gross Sales] <> 0 THEN [City (Geography mapping)] END", "string")
    C("Decline (row)", "ZN({{Sales PY (row)}}) - ZN({{Sales CY (row)}})", fmt=FMT_V)
    for cap, dim in (("Region", "[Zone]"), ("Branch", "[Branch (Geography mapping)]"), ("Family", "[Family Actual (Sheet1)]"),
                     ("Activity", "[HeliosCode]")):
        C(f"{cap} sales CY (lod)", f"{{FIXED {dim} : SUM({{{{Sales CY (row)}}}})}}")
        C(f"{cap} sales PY (lod)", f"{{FIXED {dim} : SUM({{{{Sales PY (row)}}}})}}")
        C(f"{cap} growth (lod)", f"IF {{{{{cap} sales PY (lod)}}}} > 0 THEN {{{{{cap} sales CY (lod)}}}} / {{{{{cap} sales PY (lod)}}}} - 1 END")
        C(f"Top {cap.lower()} (row)", f"IF {{{{{cap} sales CY (lod)}}}} = {{FIXED : MAX({{{{{cap} sales CY (lod)}}}})}} THEN {{{{{cap}}}}} END", "string")
    C("Fastest region (row)", "IF {{Region growth (lod)}} = {FIXED : MAX({{Region growth (lod)}})} THEN {{Region}} END", "string")
    C("Fastest branch (row)", "IF {{Branch growth (lod)}} = {FIXED : MAX({{Branch growth (lod)}})} THEN {{Branch}} END", "string")
    C("Slowest branch (row)", "IF {{Branch growth (lod)}} = {FIXED : MIN({{Branch growth (lod)}})} THEN {{Branch}} END", "string")
    C("ASM MTD sales (lod)", "{FIXED [Zone], [ASM] : SUM({{Sales MTD (row)}})}")
    C("Top ASM MTD", "IF {{ASM MTD sales (lod)}} = {FIXED [Zone] : MAX({{ASM MTD sales (lod)}})} AND {{ASM MTD sales (lod)}} > 0 THEN [ASM] END", "string")
    for key, label in MTD_VIEWS:
        C(f"MTD view: {key}", f"[Parameters].[p MTD view] = '{label}'", "boolean")
    C("Explore by",
      "CASE [Parameters].[p Dim] "
      "WHEN 'Region' THEN {{Region}} WHEN 'Branch' THEN {{Branch}} WHEN 'ASM' THEN IFNULL([ASM], 'Unmapped') "
      "WHEN 'Partner' THEN {{Partner}} WHEN 'Activity' THEN {{Activity}} WHEN 'GTM' THEN {{GTM}} WHEN 'Family' THEN {{Family}} "
      "WHEN 'Product classification' THEN IFNULL([Classification (Sheet1)], 'Unmapped') "
      "WHEN 'Pricing class' THEN IFNULL([PricingClassL3], 'Unmapped') WHEN 'Tier' THEN IFNULL([Tier Category], 'Unmapped') "
      "WHEN 'Partner type' THEN IFNULL([RD/KRO], 'Unmapped') WHEN 'SSP / SSA' THEN IFNULL([SSA], 'Unmapped') "
      "WHEN 'City' THEN IFNULL([City (Geography mapping)], 'Unmapped') WHEN 'Month' THEN {{Month}} "
      "WHEN 'Quarter' THEN {{Quarter}} WHEN 'New product category' THEN {{New product category}} "
      "WHEN 'Partner status' THEN {{Partner status}} END", "string")

    # ================================================================ targets
    GTM_OK = lambda col: f"([Parameters].[p GTM] = 'All' OR {col} = [Parameters].[p GTM])"
    # monthly targets (net) by branch and GTM
    C("Target month start", "DATE(DATETRUNC('month', [Month and year]))", "date")
    C("Target month end", "DATEADD('day', -1, DATEADD('month', 1, {{Target month start}}))", "date")
    C("Target month in period", "{{Target month end}} >= {{Period start}} AND {{Target month start}} <= {{Period end}}", "boolean")
    C("Target GTM in scope", GTM_OK("[Category (Monthly Targets)]"), "boolean")
    # the Wiser / Zeta targets are repeated on the Retail and the Project row of each branch and month: take one
    C("Focus target row", "[Category (Monthly Targets)] = 'Retail'", "boolean")
    C("Target (row)", "IF {{Target month in period}} AND {{Target GTM in scope}} THEN [Target (Monthly Targets)] * {{Target net to gross}} / {{Units divisor}} END", fmt=FMT_V)
    C("Wiser target (row)", "IF {{Target month in period}} AND {{Focus target row}} THEN [Wiser] * {{Target net to gross}} / {{Units divisor}} END", fmt=FMT_V)
    C("Zeta target (row)", "IF {{Target month in period}} AND {{Focus target row}} THEN [Miluz Zeta] * {{Target net to gross}} / {{Units divisor}} END", fmt=FMT_V)
    C("Is current month target", "{{Target month start}} = {{MTD start}}", "boolean")
    C("Target MTD (row)", "IF {{Is current month target}} AND {{Target GTM in scope}} THEN [Target (Monthly Targets)] * {{Target net to gross}} / {{Units divisor}} END", fmt=FMT_V)
    C("Wiser target MTD (row)", "IF {{Is current month target}} AND {{Focus target row}} THEN [Wiser] * {{Target net to gross}} / {{Units divisor}} END", fmt=FMT_V)
    C("Zeta target MTD (row)", "IF {{Is current month target}} AND {{Focus target row}} THEN [Miluz Zeta] * {{Target net to gross}} / {{Units divisor}} END", fmt=FMT_V)
    # yearly branch targets (net) by branch, GTM and activity
    C("Yearly target (row)",
      "IF [Year (Branch Targets1)] = {{Current year}} AND " + GTM_OK("[Category (Branch Targets1)]") + " "
      "AND NOT ({{Without FD}} AND [Helios Code (Branch Targets1)] = 'PPCFD') "
      "THEN [Target (Branch Targets1)] * {{Target net to gross}} / {{Units divisor}} END", fmt=FMT_V)
    # activity-wise month commits (net)
    C("Commit activity", activity_expr("[Helios Code (Activity wise Month Target)]"), "string")
    # FD sales are always Retail (GTM rule), so 'Project' FD commit lines can never be achieved; they are left out
    C("Project FD commit", "[Category (Activity wise Month Target)] = 'Project' AND {{Commit activity}} = 'FD'", "boolean")
    C("Commit month start", "DATE(DATETRUNC('month', [Month and year (Activity wise Month Target)]))", "date")
    C("Commit in scope",
      "DATEADD('day', -1, DATEADD('month', 1, {{Commit month start}})) >= {{Period start}} AND {{Commit month start}} <= {{Period end}} "
      "AND " + GTM_OK("[Category (Activity wise Month Target)]") + " AND NOT ({{Without FD}} AND {{Commit activity}} = 'FD') "
      "AND NOT {{Project FD commit}}", "boolean")
    C("Commit MTD in scope", "{{Commit month start}} = {{MTD start}} AND " + GTM_OK("[Category (Activity wise Month Target)]")
      + " AND NOT ({{Without FD}} AND {{Commit activity}} = 'FD') AND NOT {{Project FD commit}}", "boolean")
    C("Sales commit (row)", "IF {{Commit in scope}} THEN [Commit] * {{Commit net to gross}} / {{Units divisor}} END", fmt=FMT_V)
    C("OB commit (row)", "IF {{Commit in scope}} THEN [OB Commit] * {{Commit net to gross}} / {{Units divisor}} END", fmt=FMT_V)
    C("Sales commit MTD (row)", "IF {{Commit MTD in scope}} THEN [Commit] * {{Commit net to gross}} / {{Units divisor}} END", fmt=FMT_V)
    C("OB commit MTD (row)", "IF {{Commit MTD in scope}} THEN [OB Commit] * {{Commit net to gross}} / {{Units divisor}} END", fmt=FMT_V)
    # product (family) targets, as they are
    C("Family target (row)", "IF [Year (Product Targets)] = {{Current year}} THEN [Family Target] / {{Units divisor}} END", fmt=FMT_V)
    # partner targets (gross)
    C("Gross to net (partner)", f"IF {PM} = 'Net' THEN {TARGET_NET} ELSE 1 END")
    C("Partner FY target (row)", "IF [Pivot Field Names] = {{Current year}} THEN [Pivot Field Values] * {{Gross to net (partner)}} / {{Units divisor}} END", fmt=FMT_V)
    C("Partner FY ach (lod)", "{FIXED [SoldToPartyID] : ZN(MAX({{Partner sales in period (INR)}})) / {{Units divisor}} / MAX({{Partner FY target (row)}})}", fmt=FMT_P)
    C("Partner ach band",
      "IF ISNULL({{Partner FY ach (lod)}}) THEN '7. No target' ELSEIF {{Partner FY ach (lod)}} >= 1 THEN '1. 100% and above' "
      "ELSEIF {{Partner FY ach (lod)}} >= 0.75 THEN '2. 75-100%' ELSEIF {{Partner FY ach (lod)}} >= 0.5 THEN '3. 50-75%' "
      "ELSEIF {{Partner FY ach (lod)}} >= 0.25 THEN '4. 25-50%' ELSEIF {{Partner FY ach (lod)}} > 0 THEN '5. Below 25%' ELSE '6. No sales' END", "string")

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
        if not rls_rules:
            return "TRUE"       # the source workbook has no branch user filter
        return " OR ".join(parts) if parts else "FALSE"
    C("Branch access (sales)", rls_expr("[Branch (Geography mapping)]"), "boolean")

    # ================================================================ aggregates
    def A(cap, f, dt="real", fmt=FMT_V):
        return C(cap, f, dt, role="measure", fmt=fmt)

    def S(cap, f):
        return C(cap, f, "string", role="measure", ftype="nominal")

    def pct(cap, num, den):
        return A(cap, f"IF SUM({den}) <> 0 THEN SUM({num}) / SUM({den}) END", fmt=FMT_P)

    def growth(cap, cy, py):
        return A(cap, f"IF SUM({py}) <> 0 THEN (SUM({cy}) - SUM({py})) / ABS(SUM({py})) END", fmt=FMT_P)

    A("Sales", "SUM({{Sales CY (row)}})")
    A("Sales PY", "SUM({{Sales PY (row)}})")
    A("Sales 2 years ago", "SUM({{Sales PY2 (row)}})")
    growth("Growth %", "{{Sales CY (row)}}", "{{Sales PY (row)}}")
    A("Growth (value)", "SUM({{Sales CY (row)}}) - SUM({{Sales PY (row)}})")
    A("Decline (agg)", "IF SUM({{Sales CY (row)}}) < SUM({{Sales PY (row)}}) THEN 1 END", "integer", FMT_N)
    A("Monthly target", "SUM({{Target (row)}})")
    pct("Monthly Ach %", "{{Sales CY (row)}}", "{{Target (row)}}")
    A("Gap to monthly target", "SUM({{Sales CY (row)}}) - SUM({{Target (row)}})")
    A("Yearly target", "SUM({{Yearly target (row)}})")
    pct("Yearly Ach %", "{{Sales CY (row)}}", "{{Yearly target (row)}}")
    A("Balance to yearly target", "SUM({{Yearly target (row)}}) - SUM({{Sales CY (row)}})")
    A("Required run-rate (yearly)", "IF SUM({{Yearly target (row)}}) - SUM({{Sales CY (row)}}) > 0 THEN "
      "(SUM({{Yearly target (row)}}) - SUM({{Sales CY (row)}})) / (12 - MONTH(MIN({{Period end}})) + 1) ELSE 0 END")
    A("Current run-rate / month", "SUM({{Sales CY (row)}}) / MAX(1, DATEDIFF('day', MIN({{Period start}}), MIN(MIN({{Period end}}), TODAY())) + 1) * 30.4")
    A("Year gone %", "DATEDIFF('day', DATETRUNC('year', MIN({{Period end}})), MIN({{Period end}})) / 365", fmt=FMT_P)
    A("Sales commit", "SUM({{Sales commit (row)}})")
    pct("Commit Ach %", "{{Sales CY (row)}}", "{{Sales commit (row)}}")
    A("Gap to commit", "SUM({{Sales CY (row)}}) - SUM({{Sales commit (row)}})")
    A("OB commit", "SUM({{OB commit (row)}})")
    pct("OB commit Ach %", "{{OB CY (row)}}", "{{OB commit (row)}}")
    A("OB", "SUM({{OB CY (row)}})")
    A("OB PY", "SUM({{OB PY (row)}})")
    growth("OB growth %", "{{OB CY (row)}}", "{{OB PY (row)}}")
    pct("Sales to OB %", "{{Sales CY (row)}}", "{{OB CY (row)}}")
    A("Backlog", "SUM({{Backlog (row)}})")
    A("Backlog over 50 days", "SUM(IF {{Backlog age (days)}} > 50 THEN {{Backlog (row)}} END)")
    A("Backlog age (max days)", "MAX({{Backlog age (days)}})", "integer", FMT_N)
    A("Active partners", "COUNTD({{Active partner CY (row)}})", "integer", FMT_N)
    A("Active partners PY", "COUNTD({{Active partner PY (row)}})", "integer", FMT_N)
    A("New / revival partners", "COUNTD({{New partner CY (row)}})", "integer", FMT_N)
    A("New partner sales", "SUM(IF {{Partner new CY}} THEN {{Sales CY (row)}} END)")
    A("Lapsed partners", "COUNTD({{Lapsed partner (row)}})", "integer", FMT_N)
    A("Sales per partner", "IF COUNTD({{Active partner CY (row)}}) > 0 THEN SUM({{Sales CY (row)}}) / COUNTD({{Active partner CY (row)}}) END")
    A("Onboarded partners", "COUNTD([SoldToPartyID])", "integer", FMT_N)
    A("Active cities", "COUNTD({{Active city (row)}})", "integer", FMT_N)
    A("Days since last sale", "DATEDIFF('day', MAX({{Sale date (row)}}), TODAY())", "integer", FMT_N)
    A("Quantity", "SUM({{Qty CY (row)}})", "integer", FMT_N)
    A("Quantity PY", "SUM({{Qty PY (row)}})", "integer", FMT_N)
    growth("Quantity growth %", "{{Qty CY (row)}}", "{{Qty PY (row)}}")
    for key in ("Wiser", "Zeta", "Lara"):
        A(f"{key} sales", f"SUM({{{{{key} sales CY (row)}}}})")
        A(f"{key} sales PY", f"SUM({{{{{key} sales PY (row)}}}})")
        growth(f"{key} growth %", f"{{{{{key} sales CY (row)}}}}", f"{{{{{key} sales PY (row)}}}}")
        A(f"MTD {key} sales", f"SUM({{{{{key} sales MTD (row)}}}})")
    for key in ("Wiser", "Zeta"):
        A(f"{key} target", f"SUM({{{{{key} target (row)}}}})")
        pct(f"{key} Ach %", f"{{{{{key} sales CY (row)}}}}", f"{{{{{key} target (row)}}}}")
        A(f"MTD {key} target", f"SUM({{{{{key} target MTD (row)}}}})")
        pct(f"MTD {key} Ach %", f"{{{{{key} sales MTD (row)}}}}", f"{{{{{key} target MTD (row)}}}}")
    for q in (1, 2, 3, 4):
        A(f"Q{q} sales", f"SUM({{{{Sales Q{q} (row)}}}})")
    A("Retail sales", "SUM({{Retail sales CY (row)}})")
    A("Project sales", "SUM({{Project sales CY (row)}})")
    pct("Retail share %", "{{Retail sales CY (row)}}", "{{Sales CY (row)}}")
    pct("FD share %", "{{FD sales CY (row)}}", "{{Sales CY (row)}}")
    A("Share of sales %", "SUM({{Sales CY (row)}}) / MIN({FIXED : SUM({{Sales CY (row)}})})", fmt=FMT_P)
    A("Family target", "SUM({{Family target (row)}})")
    pct("Family Ach %", "{{Sales CY (row)}}", "{{Family target (row)}}")
    A("Partner FY target", "SUM({{Partner FY target (row)}})")
    pct("Partner FY Ach %", "{{Sales CY (row)}}", "{{Partner FY target (row)}}")
    A("Partners in band", "COUNTD([SoldToPartyID])", "integer", FMT_N)
    A("Partners at 100%+", "COUNTD(IF {{Partner ach band}} = '1. 100% and above' THEN [SoldToPartyID] END)", "integer", FMT_N)
    A("Partners with target", "COUNTD(IF NOT ISNULL({{Partner FY ach (lod)}}) THEN [SoldToPartyID] END)", "integer", FMT_N)
    A("Average rebate %", "IF SUM(IF {{Is sale in period}} THEN [Gross Sales] END) <> 0 THEN "
      "SUM(IF {{Is sale in period}} THEN [Gross Sales] * {{Rebate rate}} END) / SUM(IF {{Is sale in period}} THEN [Gross Sales] END) END", fmt=FMT_P)
    A("Has activity (agg)", "SUM({{Has activity}})", "integer", FMT_N)
    A("Has sales (agg)", "SUM({{Has sales either year}})", "integer", FMT_N)
    A("Has MTD activity (agg)", "SUM({{Has MTD activity}})", "integer", FMT_N)
    A("Has commit (agg)", "SUM(IF {{Commit in scope}} THEN 1 END)", "integer", FMT_N)
    A("Has partner target", "SUM(IF [Pivot Field Names] = {{Current year}} THEN 1 END)", "integer", FMT_N)
    S("Data as of", "STR(MAX(IF [Type] = 'Sales' THEN DATE([InvoiceDate]) END))")
    S("Period text", "STR(MIN({{Period start}})) + '  to  ' + STR(MIN({{Period end}}))")
    S("Growth text", "IF ISNULL({{Growth %}}) THEN 'no sales last year' ELSE (IF {{Growth %}} >= 0 THEN '+' ELSE '' END) + STR(ROUND({{Growth %}} * 100, 1)) + '% vs last year' END")
    S("OB growth text", "IF ISNULL({{OB growth %}}) THEN 'no OB last year' ELSE (IF {{OB growth %}} >= 0 THEN '+' ELSE '' END) + STR(ROUND({{OB growth %}} * 100, 1)) + '% vs last year' END")
    S("Nav text", "IF COUNT([Type]) > 0 THEN '' END")
    # MTD
    A("MTD sales", "SUM({{Sales MTD (row)}})")
    A("MTD sales LY", "SUM({{Sales MTD LY (row)}})")
    growth("MTD growth %", "{{Sales MTD (row)}}", "{{Sales MTD LY (row)}}")
    A("MTD target", "SUM({{Target MTD (row)}})")
    pct("MTD Ach %", "{{Sales MTD (row)}}", "{{Target MTD (row)}}")
    A("MTD commit", "SUM({{Sales commit MTD (row)}})")
    pct("MTD commit Ach %", "{{Sales MTD (row)}}", "{{Sales commit MTD (row)}}")
    A("MTD OB", "SUM({{OB MTD (row)}})")
    A("MTD OB commit", "SUM({{OB commit MTD (row)}})")
    pct("MTD OB commit Ach %", "{{OB MTD (row)}}", "{{OB commit MTD (row)}}")
    A("Sales last billing day", "SUM({{Sales last billing day (row)}})")
    A("OB last order day", "SUM({{OB last order day (row)}})")
    A("MTD active partners", "COUNTD({{Active partner MTD (row)}})", "integer", FMT_N)
    A("MTD backlog (orders this month)", "SUM({{Backlog MTD (row)}})")
    A("Qty sold MTD", "SUM({{Qty MTD (row)}})", "integer", FMT_N)
    A("Qty booked MTD", "SUM({{OB qty MTD (row)}})", "integer", FMT_N)

    # ---- story headlines
    N = lambda x: f"IFNULL(STR(ROUND({x}, 1)), '-')"
    Pc = lambda x: f"IFNULL(STR(ROUND({x} * 100, 1)) + '%', 'n/a')"
    G = lambda x: f"IFNULL((IF {x} >= 0 THEN '+' ELSE '' END) + STR(ROUND({x} * 100, 1)) + '%', 'n/a')"
    U = "' ' + [Parameters].[p Units]"
    I = lambda x: f"IFNULL(STR({x}), '0')"

    def story(cap, *lines):
        for k, f in enumerate(lines, 1):
            S(f"Story | {cap} {k}", f)

    scope = ("'Showing: ' + [Parameters].[p GTM] + ' business, ' + LOWER([Parameters].[p FD]) + ', ' + [Parameters].[p Measure] + "
             "(IF [Parameters].[p Measure] = 'Net' THEN ' (average rebate ' + " + Pc("{{Average rebate %}}") + " + ')' ELSE '' END) + '.'")
    story("Overview",
          f"'Sales: ' + {N('{{Sales}}')} + {U} + ', ' + {G('{{Growth %}}')} + ' on the same period last year; Retail ' + {Pc('{{Retail share %}}')} + ' of it, FD ' + {Pc('{{FD share %}}')} + '.'",
          f"'Yearly target: ' + {Pc('{{Yearly Ach %}}')} + ' achieved with ' + {Pc('{{Year gone %}}')} + ' of the year gone.  Commit: ' + {Pc('{{Commit Ach %}}')} + ' achieved.'",
          scope)
    story("Regions",
          f"'Largest region: ' + IFNULL(MAX({{{{Top region (row)}}}}), '-') + ' with ' + {Pc('MAX({{Region sales CY (lod)}}) / SUM({{Sales CY (row)}})')} + ' of sales.  Fastest growth: ' + IFNULL(MAX({{{{Fastest region (row)}}}}), '-') + ' (' + {G('MAX({{Region growth (lod)}})')} + ').'",
          f"{I('COUNTD(IF {{Region growth (lod)}} > 0 THEN [Zone] END)')} + ' of ' + {I('COUNTD(IF {{Region sales CY (lod)}} > 0 THEN [Zone] END)')} + ' regions are growing.  Click a region in the scorecard to filter the page.'")
    story("Branches",
          f"{I('COUNTD(IF {{Branch growth (lod)}} > 0 THEN [Branch (Geography mapping)] END)')} + ' of ' + {I('COUNTD(IF {{Branch sales CY (lod)}} > 0 THEN [Branch (Geography mapping)] END)')} + ' branches are growing.  Largest: ' + IFNULL(MAX({{{{Top branch (row)}}}}), '-') + '.'",
          f"'Fastest growth: ' + IFNULL(MAX({{{{Fastest branch (row)}}}}), '-') + ' (' + {G('MAX({{Branch growth (lod)}})')} + ');  weakest: ' + IFNULL(MAX({{{{Slowest branch (row)}}}}), '-') + ' (' + {G('MIN({{Branch growth (lod)}})')} + ').  Click a branch to see its ASMs.'")
    story("Activity",
          f"'Largest activity: ' + IFNULL(MAX({{{{Top activity (row)}}}}), '-') + '.  Sales commit ' + {N('{{Sales commit}}')} + {U} + ', achieved ' + {Pc('{{Commit Ach %}}')} + '; OB commit achieved ' + {Pc('{{OB commit Ach %}}')} + '.'")
    story("Products",
          f"'Top family: ' + IFNULL(MAX({{{{Top family (row)}}}}), '-') + ' with ' + {Pc('MAX({{Family sales CY (lod)}}) / SUM({{Sales CY (row)}})')} + ' of sales.'",
          f"'Wiser: ' + {N('{{Wiser sales}}')} + {U} + ' (' + {G('{{Wiser growth %}}')} + ').  Miluz Zeta: ' + {N('{{Zeta sales}}')} + {U} + ' (' + {G('{{Zeta growth %}}')} + ').  Miluz Lara: ' + {N('{{Lara sales}}')} + {U} + '.'")
    story("Partners",
          f"{I('{{Active partners}}')} + ' partners bought in the period (' + {I('{{Active partners PY}}')} + ' in the same period last year).  ' + {I('{{New / revival partners}}')} + ' are new or revived; ' + {I('{{Lapsed partners}}')} + ' bought last year but not this year.'",
          f"'Average sales per partner: ' + {N('{{Sales per partner}}')} + {U} + '.  ' + {I('{{Partners at 100%+}}')} + ' of ' + {I('{{Partners with target}}')} + ' partners with a target are at 100% or more.'")
    story("Targets",
          f"'Monthly targets: ' + {N('{{Monthly target}}')} + {U} + ', sales ' + {N('{{Sales}}')} + {U} + ', achievement ' + {Pc('{{Monthly Ach %}}')} + '.'",
          f"'Yearly target: ' + {Pc('{{Yearly Ach %}}')} + ' achieved, ' + {N('{{Balance to yearly target}}')} + {U} + ' to go.  Commit: ' + {Pc('{{Commit Ach %}}')} + ' achieved.'")
    story("MTD",
          f"'Month to date: ' + {N('{{MTD sales}}')} + {U} + ', ' + {G('{{MTD growth %}}')} + ' on the same days last year; ' + {Pc('{{MTD Ach %}}')} + ' of the month target and ' + {Pc('{{MTD commit Ach %}}')} + ' of the commit.'",
          f"'Last billing day (' + STR(MIN({{{{Last sales date}}}})) + '): ' + {N('{{Sales last billing day}}')} + {U} + ' sales.  Last order day (' + STR(MIN({{{{Last OB date}}}})) + '): ' + {N('{{OB last order day}}')} + {U} + ' orders.'")
    story("Orders",
          f"'Order book: ' + {N('{{OB}}')} + {U} + ' (' + {G('{{OB growth %}}')} + ' on last year); sales are ' + {Pc('{{Sales to OB %}}')} + ' of it.'",
          f"'Open backlog: ' + {N('{{Backlog}}')} + {U} + ', of which ' + {Pc('{{Backlog over 50 days}} / {{Backlog}}')} + ' was ordered more than 50 days ago.'")
    return R
