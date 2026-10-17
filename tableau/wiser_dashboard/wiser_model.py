"""Business model of the Wiser dashboard: Wiser sales (primary) and Wiser services (work orders).

Rules (as in the Wiser Performance workbook):
  * Wiser sales = primary rows whose family (material master) is Wiser. Sales by InvoiceDate, OB by Calendar Day,
    Backlog = open orders. Gross / Net use the source Gross Sales / Net Sales columns.
  * GTM: Retail when the pricing class is EKR or ERW, otherwise Project (as the Wiser workbook).
  * Wiser version (1.0 / 2.0) = the material master classification.
  * Yearly Wiser targets (Wiser Branch Targets: branch x Retail / Project x year) are used as they are.
  * Monthly Wiser targets (Monthly Targets sheet, net; gross = net / 0.83): the sheet lists every branch and
    month twice with the same Wiser target, so half the sum is used.
  * Services: one work order can have several product rows, so counts are distinct work orders and turnaround
    times are averaged per work order. Installation = Commissioning & Installation or Assessment / Inspection;
    Repair = Repair or Decommissioning. TAT = completed on - work order created, in days.
"""
from core import Registry

TARGET_NET = 0.83
EXPECTED_DAYS = 3      # service turnaround target (days), as the Wiser workbook's 'Expected Days'
MTD_VIEWS = [("Branch", "Branch wise Sales, OB and Ach%"), ("SPOC", "SPOC wise OB and Sales"),
             ("Partner", "Partner wise OB and Sales"), ("Material", "Material wise Sales, OB, Backlog")]
FMT_V, FMT_P, FMT_N, FMT_D = "n#,##0.0;-#,##0.0", "p0.0%", "n#,##0;-#,##0", "n#,##0.0;-#,##0.0"
WO = "[Work Order: Work Order Number *]"


def build_model(rls_rules):
    R = Registry()
    b = R.base
    # ---- primary table
    b("Type", "string", "dimension", "nominal")
    b("InvoiceDate", "datetime", "dimension", "ordinal", "Invoice Date")
    b("Calendar Day", "date", "dimension", "ordinal")
    b("OrderCreationDate", "date", "dimension", "ordinal")
    b("Gross Sales", "real", "measure", "quantitative")
    b("Net Sales", "real", "measure", "quantitative")
    b("BillingQuantity", "integer", "measure", "quantitative", "Billing Quantity")
    b("PricingClassL3", "string", "dimension", "nominal", "Pricing class")
    b("SoldToPartyID", "string", "dimension", "nominal", "Sold To Party ID")
    b("SoldToPartyName", "string", "dimension", "nominal", "Sold To Party Name")
    b("Patner (Partner Master Data)", "string", "dimension", "nominal", "Partner Name")
    b("RD/KRO", "string", "dimension", "nominal", "Partner Type")
    b("ASM", "string", "dimension", "nominal")
    b("SPOC (ASM!Branch Mapping1)", "string", "dimension", "nominal", "Wiser SPOC")
    b("Zone", "string", "dimension", "nominal", "Zone (geography)")
    b("Branch (Geography mapping)", "string", "dimension", "nominal", "Branch (geography)")
    b("City (Geography mapping)", "string", "dimension", "nominal", "City")
    b("Family Actual (Sheet1)", "string", "dimension", "nominal", "Family Actual")
    b("MaterialCode", "string", "dimension", "nominal", "Material Code")
    b("MaterialDescription", "string", "dimension", "nominal", "Material")
    b("SalesOrderHeader", "string", "dimension", "nominal", "Sales Order")
    b("SalesOrderItem", "string", "dimension", "nominal", "Order Item")
    # ---- material master (Wiser 1.0 / 2.0)
    b("Classification", "string", "dimension", "nominal", "Material classification", table="M")
    # ---- monthly targets (net); every branch and month appears twice
    b("Branch", "string", "dimension", "nominal", "Monthly Target Branch", table="T")
    b("Month and year", "date", "dimension", "ordinal", "Target Month", table="T")
    b("Wiser", "integer", "measure", "quantitative", "Wiser Monthly Target (net)", table="T")
    # ---- yearly Wiser targets by branch, Retail / Project and year
    b("Branch (Wiser Branch Targets)", "string", "dimension", "nominal", "Wiser Target Branch", table="W")
    b("Category (Wiser Branch Targets)", "string", "dimension", "nominal", "Wiser Target GTM", table="W")
    b("Year (Wiser Branch Targets)", "integer", "dimension", "ordinal", "Wiser Target Year", table="W")
    b("Target (Wiser Branch Targets)", "integer", "measure", "quantitative", "Wiser Branch Target", table="W")
    # ---- services (work orders)
    sv = lambda n, dt, role, ft, cap=None: b(n, dt, role, ft, cap, table="SV")
    sv("Work Order: Work Order Number *", "string", "dimension", "nominal", "Work order")
    sv("WO Created Date/Time", "datetime", "dimension", "ordinal", "WO created")
    sv("Completed On *", "datetime", "dimension", "ordinal", "WO completed")
    sv("Work Order Type *", "string", "dimension", "nominal", "Work order type")
    sv("Work Order Status *", "string", "dimension", "nominal", "Work order status")
    sv("Category (Processed Data)", "string", "dimension", "nominal", "Problem category (raw)")
    sv("Sub-Category", "string", "dimension", "nominal", "Problem sub-category")
    sv("City (Processed Data)", "string", "dimension", "nominal", "Service city")
    sv("Branch (Processed Data)", "string", "dimension", "nominal", "Service branch (file)")
    sv("Zone (Processed Data)", "string", "dimension", "nominal", "Service region (file)")
    sv("Pivot Field Values", "integer", "measure", "quantitative", "Products serviced (qty)")
    # ---- engineers
    b("Branch (Wiser - Engineers)", "string", "dimension", "nominal", "Engineer branch", table="E")
    b("City (Wiser - Engineers)", "string", "dimension", "nominal", "Engineer city", table="E")
    b("Engineers", "integer", "measure", "quantitative", "Engineers", table="E")

    # ================================================================ params
    P = R.param
    P("p Period", "Period", "string", "YTD", ["YTD", "QTD", "MTD", "Last month", "Full year", "Custom range"])
    P("p Start", "Custom start", "date", "2026-01-01", rng=("2022-01-01", "2028-12-31", None))
    P("p End", "Custom end", "date", "2026-12-31", rng=("2022-01-01", "2028-12-31", None))
    P("p Measure", "Gross / Net", "string", "Gross", ["Gross", "Net"])
    P("p Units", "Units", "string", "Lakh", ["INR", "Lakh", "Mn", "Cr"])
    P("p GTM", "Business (GTM)", "string", "All", ["All", "Retail", "Project"])
    P("p MTD view", "Select a view", "string", MTD_VIEWS[0][1], [v for _, v in MTD_VIEWS])
    P("p Dim", "Explore by", "string", "Branch",
      ["Region", "Branch", "Wiser SPOC", "ASM", "Partner", "Wiser version", "GTM", "Pricing class", "Partner type", "City",
       "Material", "Month", "Quarter", "Partner status"])

    C = R.calc
    PU, PM = "[Parameters].[p Units]", "[Parameters].[p Measure]"
    C("Units divisor", f"CASE {PU} WHEN 'INR' THEN 1 WHEN 'Lakh' THEN 100000 WHEN 'Mn' THEN 1000000 WHEN 'Cr' THEN 10000000 ELSE 100000 END")
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

    # ================================================================ primary
    C("GTM", "IF [PricingClassL3] = 'EKR' OR [PricingClassL3] = 'ERW' THEN 'Retail' ELSE 'Project' END", "string")
    C("Is Wiser", "[Family Actual (Sheet1)] = 'Wiser' AND ([Parameters].[p GTM] = 'All' OR {{GTM}} = [Parameters].[p GTM])", "boolean")
    C("Wiser version", "IFNULL([Classification], 'Unmapped')", "string")
    C("Sales value (INR)", f"IF {PM} = 'Net' THEN [Net Sales] ELSE [Gross Sales] END")
    C("Value", "{{Sales value (INR)}} / {{Units divisor}}")
    C("Region", "IFNULL([Zone], 'Unmapped')", "string")
    C("Branch", "IFNULL([Branch (Geography mapping)], 'Unmapped')", "string")
    C("SPOC", "IFNULL([SPOC (ASM!Branch Mapping1)], 'Unmapped')", "string")
    C("Partner", "IFNULL([Patner (Partner Master Data)], [SoldToPartyName]) + ' (' + [SoldToPartyID] + ')'", "string")
    C("Txn date", "CASE [Type] WHEN 'Sales' THEN DATE([InvoiceDate]) WHEN 'OB' THEN [Calendar Day] ELSE [OrderCreationDate] END", "date")
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
    C("Cancelled OB (row)", "IF {{Is OB in period}} AND [Gross Sales] < 0 THEN -{{Value}} END", fmt=FMT_V)
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
    C("Active month (row)", "IF {{Is sale in period}} AND [Gross Sales] > 0 THEN DATETRUNC('month', [InvoiceDate]) END", "datetime")
    C("Has activity", "IF {{Is sale in period}} OR {{Is sale in PY period}} OR {{Is OB in period}} OR [Type] = 'Backlog' THEN 1 END", "integer")
    C("Has sales either year", "IF {{Is sale in period}} OR {{Is sale in PY period}} THEN 1 END", "integer")
    C("Has MTD activity", "IF {{Is sale MTD}} OR {{Is sale MTD LY}} OR {{Is OB MTD}} THEN 1 END", "integer")
    C("Month", "IF MONTH({{Txn date}}) < 10 THEN '0' ELSE '' END + STR(MONTH({{Txn date}})) + ' ' + LEFT(DATENAME('month', {{Txn date}}), 3)", "string")
    C("Quarter", "'Q' + STR(DATEPART('quarter', {{Txn date}}))", "string")
    C("Day", "IF DAY({{Txn date}}) < 10 THEN '0' ELSE '' END + STR(DAY({{Txn date}}))", "string")
    C("Partner sales in period (INR)", "{FIXED [SoldToPartyID] : SUM(IF {{Is sale in period}} AND [Family Actual (Sheet1)] = 'Wiser' THEN {{Sales value (INR)}} END)}")
    C("Partner sales PY (INR)", "{FIXED [SoldToPartyID] : SUM(IF {{Is sale in PY period}} AND [Family Actual (Sheet1)] = 'Wiser' THEN {{Sales value (INR)}} END)}")
    C("Partner slab",
      "IF ISNULL({{Partner sales in period (INR)}}) OR {{Partner sales in period (INR)}} <= 0 THEN NULL "
      "ELSEIF {{Partner sales in period (INR)}} < 25000 THEN '1. Below 25 K' ELSEIF {{Partner sales in period (INR)}} < 100000 THEN '2. 25 K - 1 L' "
      "ELSEIF {{Partner sales in period (INR)}} < 500000 THEN '3. 1 - 5 L' ELSEIF {{Partner sales in period (INR)}} < 2500000 THEN '4. 5 - 25 L' "
      "ELSE '5. Above 25 L' END", "string")
    C("Partner growth band",
      "IF ISNULL([SoldToPartyID]) THEN NULL "
      "ELSEIF ZN({{Partner sales PY (INR)}}) <= 0 AND ZN({{Partner sales in period (INR)}}) > 0 THEN '5. New this period (no sales last year)' "
      "ELSEIF ZN({{Partner sales PY (INR)}}) <= 0 THEN '7. No sales either year' "
      "ELSEIF ZN({{Partner sales in period (INR)}}) <= 0 THEN '6. No sales this period' "
      "ELSEIF {{Partner sales in period (INR)}} / {{Partner sales PY (INR)}} - 1 >= 0.25 THEN '1. Growth 25% and above' "
      "ELSEIF {{Partner sales in period (INR)}} >= {{Partner sales PY (INR)}} THEN '2. Growth 0-25%' "
      "ELSEIF {{Partner sales in period (INR)}} / {{Partner sales PY (INR)}} - 1 >= -0.25 THEN '3. Decline 0-25%' "
      "ELSE '4. Decline over 25%' END", "string")
    for k, lab in ((0, "CY"), (1, "PY"), (2, "PY2")):
        C(f"Partner wiser {lab} (lod)",
          f"{{FIXED [SoldToPartyID] : SUM(IF [Type] = 'Sales' AND [Family Actual (Sheet1)] = 'Wiser' AND YEAR([InvoiceDate]) = {{{{Current year}}}} - {k} THEN [Gross Sales] END)}}")
    C("Partner new CY", "ZN({{Partner wiser CY (lod)}}) > 0 AND ZN({{Partner wiser PY (lod)}}) = 0 AND ZN({{Partner wiser PY2 (lod)}}) = 0", "boolean")
    C("Partner status", "IF {{Partner new CY}} THEN 'New / Revival' ELSE 'Existing' END", "string")
    C("New partner CY (row)", "IF {{Is sale in period}} AND {{Partner new CY}} AND [Gross Sales] <> 0 THEN [SoldToPartyID] END", "string")
    C("Lapsed partner (row)", "IF (ISNULL({{Partner sales in period (INR)}}) OR {{Partner sales in period (INR)}} <= 0) AND {{Partner sales PY (INR)}} > 0 THEN [SoldToPartyID] END", "string")
    C("Backlog age (days)", "IF [Type] = 'Backlog' THEN DATEDIFF('day', [OrderCreationDate], TODAY()) END", "integer", fmt=FMT_N)
    C("Backlog ageing",
      "IF [Type] <> 'Backlog' THEN NULL ELSEIF {{Backlog age (days)}} < 20 THEN '1. Under 20 days' "
      "ELSEIF {{Backlog age (days)}} <= 50 THEN '2. 20-50 days' ELSEIF {{Backlog age (days)}} <= 100 THEN '3. 51-100 days' "
      "ELSEIF {{Backlog age (days)}} <= 200 THEN '4. 101-200 days' ELSE '5. Over 200 days' END", "string")
    C("Retail sales CY (row)", "IF {{GTM}} = 'Retail' THEN {{Sales CY (row)}} END", fmt=FMT_V)
    C("Project sales CY (row)", "IF {{GTM}} = 'Project' THEN {{Sales CY (row)}} END", fmt=FMT_V)
    C("Sale date (row)", "IF [Type] = 'Sales' AND [Gross Sales] > 0 THEN DATE([InvoiceDate]) END", "date")
    C("Order", "[SalesOrderHeader]", "string")
    C("Order date", "IF [Type] = 'Backlog' THEN STR([OrderCreationDate]) END", "string")
    C("OB date", "IF [Type] = 'OB' THEN STR([Calendar Day]) END", "string")
    C("Decline (row)", "ZN({{Sales PY (row)}}) - ZN({{Sales CY (row)}})", fmt=FMT_V)
    for cap, dim in (("Region", "[Zone]"), ("Branch", "[Branch (Geography mapping)]")):
        C(f"{cap} sales CY (lod)", f"{{FIXED {dim} : SUM({{{{Sales CY (row)}}}})}}")
        C(f"{cap} sales PY (lod)", f"{{FIXED {dim} : SUM({{{{Sales PY (row)}}}})}}")
        C(f"{cap} growth (lod)", f"IF {{{{{cap} sales PY (lod)}}}} > 0 THEN {{{{{cap} sales CY (lod)}}}} / {{{{{cap} sales PY (lod)}}}} - 1 END")
        C(f"Top {cap.lower()} (row)", f"IF {{{{{cap} sales CY (lod)}}}} = {{FIXED : MAX({{{{{cap} sales CY (lod)}}}})}} THEN {{{{{cap}}}}} END", "string")
    C("Fastest branch (row)", "IF {{Branch growth (lod)}} = {FIXED : MAX({{Branch growth (lod)}})} THEN {{Branch}} END", "string")
    C("SPOC MTD sales (lod)", "{FIXED [Zone], [SPOC (ASM!Branch Mapping1)] : SUM({{Sales MTD (row)}})}")
    C("Top SPOC MTD", "IF {{SPOC MTD sales (lod)}} = {FIXED [Zone] : MAX({{SPOC MTD sales (lod)}})} AND {{SPOC MTD sales (lod)}} > 0 THEN [SPOC (ASM!Branch Mapping1)] END", "string")
    for key, label in MTD_VIEWS:
        C(f"MTD view: {key}", f"[Parameters].[p MTD view] = '{label}'", "boolean")
    C("Explore by",
      "CASE [Parameters].[p Dim] "
      "WHEN 'Region' THEN {{Region}} WHEN 'Branch' THEN {{Branch}} WHEN 'Wiser SPOC' THEN {{SPOC}} WHEN 'ASM' THEN IFNULL([ASM], 'Unmapped') "
      "WHEN 'Partner' THEN {{Partner}} WHEN 'GTM' THEN {{GTM}} "
      "WHEN 'Pricing class' THEN IFNULL([PricingClassL3], 'Unmapped') WHEN 'Partner type' THEN IFNULL([RD/KRO], 'Unmapped') "
      "WHEN 'City' THEN IFNULL([City (Geography mapping)], 'Unmapped') WHEN 'Material' THEN IFNULL([MaterialDescription], 'Unmapped') "
      "WHEN 'Month' THEN {{Month}} WHEN 'Quarter' THEN {{Quarter}} WHEN 'Partner status' THEN {{Partner status}} END", "string")

    # ================================================================ targets
    C("Target month start", "DATE(DATETRUNC('month', [Month and year]))", "date")
    C("Target month in period", "DATEADD('day', -1, DATEADD('month', 1, {{Target month start}})) >= {{Period start}} AND {{Target month start}} <= {{Period end}}", "boolean")
    C("Monthly target (row)", "IF {{Target month in period}} THEN [Wiser] / 2 * {{Target net to gross}} / {{Units divisor}} END", fmt=FMT_V)
    C("Monthly target MTD (row)", "IF {{Target month start}} = {{MTD start}} THEN [Wiser] / 2 * {{Target net to gross}} / {{Units divisor}} END", fmt=FMT_V)
    for cap, cond in (("Yearly target", "TRUE"), ("Retail target", "[Category (Wiser Branch Targets)] = 'Retail'"),
                      ("Project target", "[Category (Wiser Branch Targets)] = 'Project'")):
        C(f"{cap} (row)",
          f"IF [Year (Wiser Branch Targets)] = {{{{Current year}}}} AND {cond} AND ([Parameters].[p GTM] = 'All' OR "
          f"[Category (Wiser Branch Targets)] = [Parameters].[p GTM]) THEN [Target (Wiser Branch Targets)] / {{{{Units divisor}}}} END", fmt=FMT_V)

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
    C("Branch access (services)", rls_expr("[Branch (Processed Data)]"), "boolean")
    C("Branch access (engineers)", rls_expr("[Branch (Wiser - Engineers)]"), "boolean")

    # ================================================================ services
    C("Service region", "IFNULL([Zone (Processed Data)], 'Unmapped')", "string")
    C("Service branch", "IFNULL([Branch (Processed Data)], 'Unmapped')", "string")
    C("Work order group", "IF [Work Order Type *] = 'Commissioning & Installation' OR [Work Order Type *] = 'Assessment/ Inspection' THEN 'Installation' "
      "ELSEIF [Work Order Type *] = 'Repair' OR [Work Order Type *] = 'Decommissioning' THEN 'Repair' ELSE 'Other' END", "string")
    C("Problem category", "IFNULL([Category (Processed Data)], 'Unmapped')", "string")
    C("WO in period", "DATE([WO Created Date/Time]) >= {{Period start}} AND DATE([WO Created Date/Time]) <= {{Period end}}", "boolean")
    C("WO in PY period", "DATE([WO Created Date/Time]) >= {{PY start}} AND DATE([WO Created Date/Time]) <= {{PY end}}", "boolean")
    C("WO month", "IF MONTH([WO Created Date/Time]) < 10 THEN '0' ELSE '' END + STR(MONTH([WO Created Date/Time])) + ' ' + LEFT(DATENAME('month', [WO Created Date/Time]), 3)", "string")
    C("WO (row)", f"IF {{{{WO in period}}}} THEN {WO} END", "string")
    C("WO PY (row)", f"IF {{{{WO in PY period}}}} THEN {WO} END", "string")
    C("Installation WO (row)", f"IF {{{{WO in period}}}} AND {{{{Work order group}}}} = 'Installation' THEN {WO} END", "string")
    C("Repair WO (row)", f"IF {{{{WO in period}}}} AND {{{{Work order group}}}} = 'Repair' THEN {WO} END", "string")
    C("WO rows (lod)", f"{{FIXED {WO} : COUNT({WO})}}", "integer")
    C("TAT days (row)", "DATEDIFF('hour', [WO Created Date/Time], [Completed On *]) / 24", fmt=FMT_D)
    C("Installation TAT share (row)", "IF {{WO in period}} AND {{Work order group}} = 'Installation' AND NOT ISNULL([Completed On *]) THEN {{TAT days (row)}} / {{WO rows (lod)}} END", fmt=FMT_D)
    C("Repair TAT share (row)", "IF {{WO in period}} AND {{Work order group}} = 'Repair' AND NOT ISNULL([Completed On *]) THEN {{TAT days (row)}} / {{WO rows (lod)}} END", fmt=FMT_D)
    C("Completed installation (row)", f"IF {{{{WO in period}}}} AND {{{{Work order group}}}} = 'Installation' AND NOT ISNULL([Completed On *]) THEN {WO} END", "string")
    C("Completed repair (row)", f"IF {{{{WO in period}}}} AND {{{{Work order group}}}} = 'Repair' AND NOT ISNULL([Completed On *]) THEN {WO} END", "string")
    C("On-time WO (row)", f"IF {{{{WO in period}}}} AND NOT ISNULL([Completed On *]) AND {{{{TAT days (row)}}}} <= {EXPECTED_DAYS} THEN {WO} END", "string")
    C("Completed WO (row)", f"IF {{{{WO in period}}}} AND NOT ISNULL([Completed On *]) THEN {WO} END", "string")
    C("Products serviced (row)", "IF {{WO in period}} THEN [Pivot Field Values] END", "integer", fmt=FMT_N)

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
    A("Yearly target", "SUM({{Yearly target (row)}})")
    pct("Yearly Ach %", "{{Sales CY (row)}}", "{{Yearly target (row)}}")
    A("Balance to yearly target", "SUM({{Yearly target (row)}}) - SUM({{Sales CY (row)}})")
    A("Required run-rate (yearly)", "IF SUM({{Yearly target (row)}}) - SUM({{Sales CY (row)}}) > 0 THEN "
      "(SUM({{Yearly target (row)}}) - SUM({{Sales CY (row)}})) / (12 - MONTH(MIN({{Period end}})) + 1) ELSE 0 END")
    A("Retail target", "SUM({{Retail target (row)}})")
    A("Project target", "SUM({{Project target (row)}})")
    A("Retail sales", "SUM({{Retail sales CY (row)}})")
    A("Project sales", "SUM({{Project sales CY (row)}})")
    pct("Retail Ach %", "{{Retail sales CY (row)}}", "{{Retail target (row)}}")
    pct("Project Ach %", "{{Project sales CY (row)}}", "{{Project target (row)}}")
    pct("Retail share %", "{{Retail sales CY (row)}}", "{{Sales CY (row)}}")
    A("Monthly target", "SUM({{Monthly target (row)}})")
    pct("Monthly Ach %", "{{Sales CY (row)}}", "{{Monthly target (row)}}")
    A("Gap to monthly target", "SUM({{Sales CY (row)}}) - SUM({{Monthly target (row)}})")
    A("Year gone %", "DATEDIFF('day', DATETRUNC('year', MIN({{Period end}})), MIN({{Period end}})) / 365", fmt=FMT_P)
    A("OB", "SUM({{OB CY (row)}})")
    A("OB PY", "SUM({{OB PY (row)}})")
    growth("OB growth %", "{{OB CY (row)}}", "{{OB PY (row)}}")
    pct("Sales to OB %", "{{Sales CY (row)}}", "{{OB CY (row)}}")
    A("Cancelled orders", "SUM({{Cancelled OB (row)}})")
    A("Backlog", "SUM({{Backlog (row)}})")
    A("Backlog over 50 days", "SUM(IF {{Backlog age (days)}} > 50 THEN {{Backlog (row)}} END)")
    A("Backlog age (max days)", "MAX({{Backlog age (days)}})", "integer", FMT_N)
    A("Active partners", "COUNTD({{Active partner CY (row)}})", "integer", FMT_N)
    A("Active partners PY", "COUNTD({{Active partner PY (row)}})", "integer", FMT_N)
    A("New / revival partners", "COUNTD({{New partner CY (row)}})", "integer", FMT_N)
    A("Lapsed partners", "COUNTD({{Lapsed partner (row)}})", "integer", FMT_N)
    A("Sales per partner", "IF COUNTD({{Active partner CY (row)}}) > 0 THEN SUM({{Sales CY (row)}}) / COUNTD({{Active partner CY (row)}}) END")
    A("Months active", "COUNTD({{Active month (row)}})", "integer", FMT_N)
    A("Days since last sale", "DATEDIFF('day', MAX({{Sale date (row)}}), TODAY())", "integer", FMT_N)
    A("Quantity", "SUM({{Qty CY (row)}})", "integer", FMT_N)
    A("Quantity PY", "SUM({{Qty PY (row)}})", "integer", FMT_N)
    growth("Quantity growth %", "{{Qty CY (row)}}", "{{Qty PY (row)}}")
    A("Share of sales %", "SUM({{Sales CY (row)}}) / MIN({FIXED : SUM({{Sales CY (row)}})})", fmt=FMT_P)
    A("Has activity (agg)", "SUM({{Has activity}})", "integer", FMT_N)
    A("Has sales (agg)", "SUM({{Has sales either year}})", "integer", FMT_N)
    A("Has MTD activity (agg)", "SUM({{Has MTD activity}})", "integer", FMT_N)
    S("Data as of", "STR(MAX(IF [Type] = 'Sales' THEN DATE([InvoiceDate]) END))")
    S("Period text", "STR(MIN({{Period start}})) + '  to  ' + STR(MIN({{Period end}}))")
    S("Growth text", "IF ISNULL({{Growth %}}) THEN 'no sales last year' ELSE (IF {{Growth %}} >= 0 THEN '+' ELSE '' END) + STR(ROUND({{Growth %}} * 100, 1)) + '% vs last year' END")
    S("OB growth text", "IF ISNULL({{OB growth %}}) THEN 'no OB last year' ELSE (IF {{OB growth %}} >= 0 THEN '+' ELSE '' END) + STR(ROUND({{OB growth %}} * 100, 1)) + '% vs last year' END")
    S("Nav text", "IF COUNT([Type]) > 0 THEN '' END")
    A("MTD sales", "SUM({{Sales MTD (row)}})")
    A("MTD sales LY", "SUM({{Sales MTD LY (row)}})")
    growth("MTD growth %", "{{Sales MTD (row)}}", "{{Sales MTD LY (row)}}")
    A("MTD target", "SUM({{Monthly target MTD (row)}})")
    pct("MTD Ach %", "{{Sales MTD (row)}}", "{{Monthly target MTD (row)}}")
    A("MTD OB", "SUM({{OB MTD (row)}})")
    A("Sales last billing day", "SUM({{Sales last billing day (row)}})")
    A("OB last order day", "SUM({{OB last order day (row)}})")
    A("MTD active partners", "COUNTD({{Active partner MTD (row)}})", "integer", FMT_N)
    A("MTD backlog (orders this month)", "SUM({{Backlog MTD (row)}})")
    A("Qty sold MTD", "SUM({{Qty MTD (row)}})", "integer", FMT_N)
    A("Qty booked MTD", "SUM({{OB qty MTD (row)}})", "integer", FMT_N)
    # services
    A("Service requests", "COUNTD({{WO (row)}})", "integer", FMT_N)
    A("Service requests PY", "COUNTD({{WO PY (row)}})", "integer", FMT_N)
    A("Service requests growth %", "IF COUNTD({{WO PY (row)}}) > 0 THEN COUNTD({{WO (row)}}) / COUNTD({{WO PY (row)}}) - 1 END", fmt=FMT_P)
    A("Installations", "COUNTD({{Installation WO (row)}})", "integer", FMT_N)
    A("Repairs", "COUNTD({{Repair WO (row)}})", "integer", FMT_N)
    A("Avg installation TAT (days)", "IF COUNTD({{Completed installation (row)}}) > 0 THEN SUM({{Installation TAT share (row)}}) / COUNTD({{Completed installation (row)}}) END", fmt=FMT_D)
    A("Avg repair TAT (days)", "IF COUNTD({{Completed repair (row)}}) > 0 THEN SUM({{Repair TAT share (row)}}) / COUNTD({{Completed repair (row)}}) END", fmt=FMT_D)
    A(f"Done within {EXPECTED_DAYS} days %", "IF COUNTD({{Completed WO (row)}}) > 0 THEN COUNTD({{On-time WO (row)}}) / COUNTD({{Completed WO (row)}}) END", fmt=FMT_P)
    A("Completed %", "IF COUNTD({{WO (row)}}) > 0 THEN COUNTD({{Completed WO (row)}}) / COUNTD({{WO (row)}}) END", fmt=FMT_P)
    A("Products serviced", "SUM({{Products serviced (row)}})", "integer", FMT_N)
    A("Engineers (count)", "SUM([Engineers])", "integer", FMT_N)
    A("Has service (agg)", "COUNTD({{WO (row)}}) + COUNTD({{WO PY (row)}})", "integer", FMT_N)
    S("Service period text", "STR(MIN({{Period start}})) + '  to  ' + STR(MIN({{Period end}}))")

    # ---- story headlines
    N = lambda x: f"IFNULL(STR(ROUND({x}, 1)), '-')"
    Pc = lambda x: f"IFNULL(STR(ROUND({x} * 100, 1)) + '%', 'n/a')"
    G = lambda x: f"IFNULL((IF {x} >= 0 THEN '+' ELSE '' END) + STR(ROUND({x} * 100, 1)) + '%', 'n/a')"
    U = "' ' + [Parameters].[p Units]"
    I = lambda x: f"IFNULL(STR({x}), '0')"

    def story(cap, *lines):
        for k, f in enumerate(lines, 1):
            S(f"Story | {cap} {k}", f)

    story("Overview",
          f"'Wiser sales: ' + {N('{{Sales}}')} + {U} + ', ' + {G('{{Growth %}}')} + ' on the same period last year; Retail ' + {Pc('{{Retail share %}}')} + ' of it.'",
          f"'Yearly Wiser target: ' + {Pc('{{Yearly Ach %}}')} + ' achieved with ' + {Pc('{{Year gone %}}')} + ' of the year gone.  ' + {I('{{Active partners}}')} + ' partners bought Wiser (' + {I('{{New / revival partners}}')} + ' new or revived).'")
    story("Branches",
          f"'Largest branch: ' + IFNULL(MAX({{{{Top branch (row)}}}}), '-') + ' with ' + {Pc('MAX({{Branch sales CY (lod)}}) / SUM({{Sales CY (row)}})')} + ' of Wiser sales.  Fastest growth: ' + IFNULL(MAX({{{{Fastest branch (row)}}}}), '-') + ' (' + {G('MAX({{Branch growth (lod)}})')} + ').'",
          f"'Largest region: ' + IFNULL(MAX({{{{Top region (row)}}}}), '-') + '.  Click a branch to filter the page.'")
    story("Partners",
          f"{I('{{Active partners}}')} + ' partners bought Wiser in the period (' + {I('{{Active partners PY}}')} + ' last year); ' + {I('{{Lapsed partners}}')} + ' bought last year but not this year.  Average ' + {N('{{Sales per partner}}')} + {U} + ' per partner.'")
    story("Targets",
          f"'Monthly Wiser targets: ' + {N('{{Monthly target}}')} + {U} + ', sales ' + {N('{{Sales}}')} + {U} + ', achievement ' + {Pc('{{Monthly Ach %}}')} + '.  Yearly target: ' + {Pc('{{Yearly Ach %}}')} + ' achieved.'")
    story("MTD",
          f"'Month to date: ' + {N('{{MTD sales}}')} + {U} + ', ' + {G('{{MTD growth %}}')} + ' on the same days last year and ' + {Pc('{{MTD Ach %}}')} + ' of the month target.'",
          f"'Last billing day (' + STR(MIN({{{{Last sales date}}}})) + '): ' + {N('{{Sales last billing day}}')} + {U} + ' sales.  Last order day (' + STR(MIN({{{{Last OB date}}}})) + '): ' + {N('{{OB last order day}}')} + {U} + ' orders.'")
    story("Orders",
          f"'Order book: ' + {N('{{OB}}')} + {U} + ' (' + {G('{{OB growth %}}')} + ' on last year); cancelled order value ' + {N('{{Cancelled orders}}')} + {U} + '.'",
          f"'Open backlog: ' + {N('{{Backlog}}')} + {U} + ', of which ' + {Pc('{{Backlog over 50 days}} / {{Backlog}}')} + ' was ordered more than 50 days ago.'")
    story("Services",
          f"{I('{{Service requests}}')} + ' service requests in the period (' + {G('{{Service requests growth %}}')} + ' on last year): ' + {I('{{Installations}}')} + ' installations and ' + {I('{{Repairs}}')} + ' repairs.'",
          f"'Average turnaround: installation ' + {N('{{Avg installation TAT (days)}}')} + ' days, repair ' + {N('{{Avg repair TAT (days)}}')} + ' days; ' + {Pc('{{Done within 3 days %}}')} + ' done within 3 days.'")
    return R
