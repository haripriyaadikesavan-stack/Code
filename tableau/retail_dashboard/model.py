"""Business model: parameters and calculated fields for the Retail dashboard.

Business rules (same conversions as the existing Retail workbook):
  * Sales = rows with Type 'Sales' dated by InvoiceDate; OB = Type 'OB' by
    Calendar Day; Backlog = Type 'Backlog' (open orders, by OrderCreationDate).
  * Net Sales already carries the yearly rebate (from the notebook).
  * Monthly branch targets (Monthly Targets sheet) are NET; gross = net / 0.83.
  * Yearly branch targets (Branch Targets sheet) are NET; gross = net / 0.83.
  * Partner targets (QTOD and ATOD file) are GROSS; net = gross x 0.83.
  * SSP (SSA) targets are used as they are, as in the old workbook.
"""
from core import PDS, Registry

GROSS_FACTOR = 0.83
# the views of the MTD dashboard (as in the old workbook's 'Select a View')
MTD_VIEWS = [("Branch", "Branch wise Sales, OB and Ach%"), ("Activity", "Branch & Activity wise OB and Sales"),
             ("ASM", "ASM wise OB and Sales"), ("Focus", "Focus Prod Month Target vs Ach%"),
             ("Partner", "Partner wise OB and Sales"), ("Product", "Product wise Sales, OB, Backlog")]
# the partner target vs achievement views (as the old Target vs Ach dashboard's drop-down)
TARGET_VIEWS = [("Annual", "Annual Target vs Ach"), ("Q1", "QTOD - Q1"), ("Q2", "QTOD - Q2"), ("Q3", "QTOD - Q3"),
                ("Q4", "QTOD - Q4"), ("Wiser", "Wiser Target vs Ach"), ("Zeta", "Zeta Target vs Ach")]
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
FMT_V, FMT_P, FMT_N = "n#,##0.0;-#,##0.0", "p0.0%", "n#,##0;-#,##0"


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
    b("RD/KRO", "string", "dimension", "nominal", "Partner Type")
    b("ASM", "string", "dimension", "nominal")
    b("Tier Category", "string", "dimension", "nominal", "Tier")
    b("Onboarded Year (Partner Master Data)", "string", "dimension", "nominal", "Onboarded Year")
    b("SSA", "string", "dimension", "nominal", "SSP / SSA")
    b("Zone", "string", "dimension", "nominal", "Zone (geography)")
    b("Branch (Geography mapping)", "string", "dimension", "nominal", "Branch (geography)")
    b("City (Geography mapping)", "string", "dimension", "nominal", "City")
    b("Family Actual (Sheet1)", "string", "dimension", "nominal", "Family Actual")
    b("Classification (Sheet1)", "string", "dimension", "nominal", "Product Classification")
    b("Helios Code", "string", "dimension", "nominal", "Activity (Helios)")
    b("MaterialCode", "string", "dimension", "nominal", "Material Code")
    b("MaterialDescription", "string", "dimension", "nominal", "Material")
    b("Sold-to Party (Partner Master Data)", "string", "dimension", "nominal", "Partner Code")
    b("SalesOrderHeader", "integer", "dimension", "nominal", "Sales Order")
    b("SalesOrderItem", "string", "dimension", "nominal", "Order Item")
    b("PlantName", "string", "dimension", "nominal", "Plant")
    # ---- partner targets (Partner Targets sheet, pivoted by year; GROSS)
    b("Pivot Field Names", "integer", "dimension", "ordinal", "Partner Target Year", table="PT")
    b("Pivot Field Values", "real", "measure", "quantitative", "Partner Target (gross)", table="PT")
    # ---- pavilion partners (Pavilion Sales sheet)
    b("Name of RO/KRO", "string", "dimension", "nominal", "Pavilion Partner", table="V")
    b("Branch (Pavilion Sales)", "string", "dimension", "nominal", "Pavilion Branch", table="V")
    b("Zone/Region", "string", "dimension", "nominal", "Pavilion Region", table="V")
    b("KRO/RO", "string", "dimension", "nominal", "Pavilion Type (KRO / RO)", table="V")
    b("City (Pavilion Sales)", "string", "dimension", "nominal", "Pavilion City", table="V")
    b("Active/Removed", "string", "dimension", "nominal", "Pavilion Status", table="V")
    b("Sales Manager", "string", "dimension", "nominal", "Pavilion Sales Manager", table="V")
    b("GSTIN no", "string", "dimension", "nominal", "Pavilion GSTIN", table="V")
    b("2025 Target (Pavilion Sales)", "real", "measure", "quantitative", "Pavilion Target 2025", table="V")
    # ---- monthly targets (NET)
    b("Branch", "string", "dimension", "nominal", "Monthly Target Branch", table="T")
    b("Zone (Monthly Targets)", "string", "dimension", "nominal", "Monthly Target Region", table="T")
    b("Month and year", "date", "dimension", "ordinal", "Target Month", table="T")
    b("Target (Monthly Targets)", "integer", "measure", "quantitative", "Monthly Target (net)", table="T")
    b("Wiser", "integer", "measure", "quantitative", "Wiser Monthly Target (net)", table="T")
    b("Miluz Zeta", "integer", "measure", "quantitative", "Miluz Zeta Monthly Target (net)", table="T")
    # ---- yearly branch targets (NET)
    b("Branch (Branch Targets1)", "string", "dimension", "nominal", "Yearly Target Branch", table="B")
    b("Year (Branch Targets1)", "integer", "dimension", "ordinal", "Yearly Target Year", table="B")
    b("Target (Branch Targets1)", "integer", "measure", "quantitative", "Yearly Branch Target (net)", table="B")
    # ---- SSP / SSA targets (as is)
    b("SSP", "string", "dimension", "nominal", "SSP", table="S")
    b("Branch (SSA Targets)", "string", "dimension", "nominal", "SSP Branch", table="S")
    b("Year (SSA Targets)", "integer", "dimension", "ordinal", "SSP Target Year", table="S")
    b("Target (SSA Targets)", "integer", "measure", "quantitative", "SSP Target", table="S")
    # ---- partner targets QTOD / ATOD (GROSS)
    b("CC", "string", "dimension", "nominal", "Partner Code (targets)", table="Q")
    b("Partner", "string", "dimension", "nominal", "Partner (targets)", table="Q")
    b("Branch (QTOD - ATOD)", "string", "dimension", "nominal", "Branch (targets)", table="Q")
    b("RD/ KRO", "string", "dimension", "nominal", "Partner Type (targets)", table="Q")
    b("Target Year", "integer", "dimension", "ordinal", "Target Year", table="Q")
    b("Target 100%", "integer", "measure", "quantitative", "Partner Target 100% (gross)", table="Q")
    b("Target 108%", "integer", "measure", "quantitative", "Partner Target 108% (gross)", table="Q")
    b("Qtr 1 Target", "integer", "measure", "quantitative", table="Q")
    b("Qtr 2 Target", "string", "dimension", "nominal", table="Q")
    b("Qtr 3 Target", "string", "dimension", "nominal", table="Q")
    b("Qtr 4 Target", "string", "dimension", "nominal", table="Q")
    b("Zeta Target", "integer", "measure", "quantitative", "Partner Zeta Target (gross)", table="Q")
    b("Wiser Target", "integer", "measure", "quantitative", "Partner Wiser Target (gross)", table="Q")
    # ---- Pavilion summary tables (Pavilion_Unpivoted.xlsx)
    pv = lambda n, dt, role, ft, cap, t: b(n, dt, role, ft, cap, table=t, ds=PDS)
    pv("Zone/Region", "string", "dimension", "nominal", "Region", "PV1")
    pv("Row Type", "string", "dimension", "nominal", "Row type (slabs)", "PV1")
    pv("Sales Manager / Entity", "string", "dimension", "nominal", "ASM", "PV1")
    pv("Branch", "string", "dimension", "nominal", "Branch", "PV1")
    pv("2025 Sales Slab", "string", "dimension", "nominal", "2025 Sales Slab", "PV1")
    pv("2026 Sales Slab", "string", "dimension", "nominal", "2026 Sales Slab", "PV1")
    pv("2025-26 Growth %", "real", "measure", "quantitative", "Growth % (slabs)", "PV1")
    MS = "(Minimum!Slab!Achievement)"
    pv(f"Zone/Region {MS}", "string", "dimension", "nominal", "Region (min slab)", "PV2")
    pv(f"Row Type {MS}", "string", "dimension", "nominal", "Row type (min slab)", "PV2")
    pv(f"No. of KRO {MS}", "integer", "measure", "quantitative", "KRO (min slab)", "PV2")
    pv(f"No. of RO {MS}", "integer", "measure", "quantitative", "RO (min slab)", "PV2")
    pv("No of KRO Achieved", "integer", "measure", "quantitative", "KRO qualified (min slab)", "PV2")
    pv("No. of RO Achieved", "integer", "measure", "quantitative", "RO qualified (min slab)", "PV2")
    PT = "(Pavilion!Type!Summary)"
    pv("Pavilion Type", "string", "dimension", "nominal", "Pavilion Type", "PV3")
    pv(f"Row Type {PT}", "string", "dimension", "nominal", "Row type (type)", "PV3")
    pv("Count of 2026 Classification", "integer", "measure", "quantitative", "Pavilions (type)", "PV3")
    pv(f"Sum of YTD Slab 2025 {PT}", "string", "dimension", "nominal", "2025 YTD Slab", "PV3")
    pv("Sum of YTD Slab 2026", "string", "dimension", "nominal", "2026 YTD Slab", "PV3")
    RA = "(Retailer!App!TnC!Acceptance)"
    pv(f"Zone/Region {RA}", "string", "dimension", "nominal", "Region (app T&C)", "PV4")
    pv(f"Row Type {RA}", "string", "dimension", "nominal", "Row type (app T&C)", "PV4")
    pv("Total no. of RO Pavilions", "integer", "measure", "quantitative", "RO pavilions (app T&C)", "PV4")
    pv("Accepted", "integer", "measure", "quantitative", "Accepted (app T&C)", "PV4")
    pv("Pending", "integer", "measure", "quantitative", "Pending (app T&C)", "PV4")
    SC = "(Sales!Classification!Summary)"
    pv("Classification", "string", "dimension", "nominal", "Pavilion Classification", "PV5")
    pv(f"Row Type {SC}", "string", "dimension", "nominal", "Row type (class)", "PV5")
    pv(f"Count of 2026 Classification {SC}", "integer", "measure", "quantitative", "Pavilions (class)", "PV5")
    pv("Sum of YTD Slab 2025", "string", "dimension", "nominal", "2025 YTD Slab (class)", "PV5")
    pv(f"Sum of YTD Slab 2026 {SC}", "string", "dimension", "nominal", "2026 YTD Slab (class)", "PV5")
    pv(f"2025-26 Growth % {SC}", "real", "measure", "quantitative", "Growth % (class)", "PV5")
    TT = "(Total!Target!Achievement)"
    pv(f"Zone/Region {TT}", "string", "dimension", "nominal", "Region (total target)", "PV6")
    pv(f"Row Type {TT}", "string", "dimension", "nominal", "Row type (total target)", "PV6")
    pv(f"No. of KRO {TT}", "integer", "measure", "quantitative", "KRO (total target)", "PV6")
    pv(f"No. of RO {TT}", "integer", "measure", "quantitative", "RO (total target)", "PV6")
    pv(f"No of KRO Achieved {TT}", "integer", "measure", "quantitative", "KRO qualified (total target)", "PV6")
    pv(f"No. of RO Achieved {TT}", "integer", "measure", "quantitative", "RO qualified (total target)", "PV6")
    WS = "(Wiser!Slabs!Unpivoted)"
    pv(f"Zone/Region {WS}", "string", "dimension", "nominal", "Region (Wiser slab)", "PV7")
    pv(f"Row Type {WS}", "string", "dimension", "nominal", "Row type (Wiser slab)", "PV7")
    pv(f"Sales Manager / Entity {WS}", "string", "dimension", "nominal", "ASM (Wiser slab)", "PV7")
    pv(f"Branch {WS}", "string", "dimension", "nominal", "Branch (Wiser slab)", "PV7")
    pv("No. of KRO", "integer", "measure", "quantitative", "KRO (Wiser slab)", "PV7")
    pv("No. of RO", "integer", "measure", "quantitative", "RO (Wiser slab)", "PV7")
    pv("No of KRO > 1 lac", "integer", "measure", "quantitative", "KRO above 1 L (Wiser)", "PV7")
    pv("No. of RO > 25K", "integer", "measure", "quantitative", "RO above 25 K (Wiser)", "PV7")

    # ================================================================ params
    P = R.param
    P("p Period", "Period", "string", "YTD",
      ["YTD", "QTD", "MTD", "Last month", "Full year", "Custom range"])
    P("p Start", "Custom start", "date", "2026-01-01", rng=("2023-01-01", "2028-12-31", None))
    P("p End", "Custom end", "date", "2026-12-31", rng=("2023-01-01", "2028-12-31", None))
    P("p Measure", "Gross / Net", "string", "Gross", ["Gross", "Net"])
    P("p Units", "Units", "string", "Mn", ["INR", "Lakh", "Mn", "Cr"])
    P("p Target view", "Select a View - QTOD/ATOD", "string", TARGET_VIEWS[0][1], [v for _, v in TARGET_VIEWS])
    P("p MTD view", "Select a view", "string", MTD_VIEWS[0][1], [v for _, v in MTD_VIEWS])
    P("p Show", "Show", "string", "Value", ["Value", "Quantity"])
    P("p Dim2", "Explore across", "string", "None",
      ["None", "Family", "Activity (Helios)", "Product classification", "New product category", "Region", "Branch",
       "Partner type", "Tier", "Month", "Quarter", "Partner status"])
    P("p Dim", "Explore by", "string", "Branch",
      ["Region", "Branch", "ASM", "Partner", "Family", "Product classification", "Activity (Helios)",
       "Tier", "Partner type", "SSP / SSA", "City", "Month", "Quarter", "New product category", "Partner status",
       "Material code"])

    C = R.calc
    PU, PM, GV, NV = "[Parameters].[p Units]", "[Parameters].[p Measure]", "Gross", "Net"

    # ================================================================ parameter-only helpers
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
    C("MTD start", "DATE(DATETRUNC('month', TODAY()))", "date")
    C("MTD LY start", "DATEADD('year', -1, {{MTD start}})", "date")
    C("MTD LY end", "DATEADD('year', -1, TODAY())", "date")
    # conversions exactly as in the old workbook
    C("Net to gross", f"IF {PM} = '{GV}' THEN 1 / {GROSS_FACTOR} ELSE 1 END")      # monthly / yearly targets (net)
    C("Gross to net", f"IF {PM} = '{NV}' THEN {GROSS_FACTOR} ELSE 1 END")            # partner targets (gross)

    # ================================================================ primary table
    C("Sales value (INR)", f"IF {PM} = '{NV}' THEN [Net Sales] ELSE [Gross Sales] END")
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
    C("OB CY (row)", "IF {{Is OB in period}} THEN {{Value}} END", fmt=FMT_V)
    C("OB PY (row)", "IF {{Is OB in PY period}} THEN {{Value}} END", fmt=FMT_V)
    C("Sales MTD (row)", "IF {{Is sale MTD}} THEN {{Value}} END", fmt=FMT_V)
    C("Sales MTD LY (row)", "IF {{Is sale MTD LY}} THEN {{Value}} END", fmt=FMT_V)
    C("OB MTD (row)", "IF {{Is OB MTD}} THEN {{Value}} END", fmt=FMT_V)
    # the latest day in the data (not 'yesterday', which is empty on Mondays and before a refresh)
    C("Last sales date", "{FIXED : MAX(IF [Type] = 'Sales' THEN DATE([InvoiceDate]) END)}", "date")
    C("Last OB date", "{FIXED : MAX(IF [Type] = 'OB' THEN [Calendar Day] END)}", "date")
    C("Sales last billing day (row)", "IF [Type] = 'Sales' AND DATE([InvoiceDate]) = {{Last sales date}} THEN {{Value}} END", fmt=FMT_V)
    C("OB last order day (row)", "IF [Type] = 'OB' AND [Calendar Day] = {{Last OB date}} THEN {{Value}} END", fmt=FMT_V)
    C("Backlog (row)", "IF [Type] = 'Backlog' THEN {{Value}} END", fmt=FMT_V)
    C("Qty CY (row)", "IF {{Is sale in period}} THEN [BillingQuantity] END", "integer", fmt=FMT_N)
    C("Qty PY (row)", "IF {{Is sale in PY period}} THEN [BillingQuantity] END", "integer", fmt=FMT_N)
    C("Active partner CY (row)", "IF {{Is sale in period}} AND [Gross Sales] <> 0 THEN [SoldToPartyID] END", "string")
    C("Active partner PY (row)", "IF {{Is sale in PY period}} AND [Gross Sales] <> 0 THEN [SoldToPartyID] END", "string")
    C("Active partner MTD (row)", "IF {{Is sale MTD}} AND [Gross Sales] <> 0 THEN [SoldToPartyID] END", "string")
    C("Has activity", "IF {{Is sale in period}} OR {{Is sale in PY period}} OR {{Is OB in period}} OR [Type] = 'Backlog' THEN 1 END", "integer")
    C("Has sales either year", "IF {{Is sale in period}} OR {{Is sale in PY period}} THEN 1 END", "integer")
    C("Has MTD activity", "IF {{Is sale MTD}} OR {{Is sale MTD LY}} OR {{Is OB MTD}} THEN 1 END", "integer")
    C("Txn date", "CASE [Type] WHEN 'Sales' THEN DATE([InvoiceDate]) WHEN 'OB' THEN [Calendar Day] ELSE [OrderCreationDate] END", "date")
    C("Month", "IF MONTH({{Txn date}}) < 10 THEN '0' ELSE '' END + STR(MONTH({{Txn date}})) + ' ' + LEFT(DATENAME('month', {{Txn date}}), 3)", "string")
    C("Quarter", "'Q' + STR(DATEPART('quarter', {{Txn date}}))", "string")
    C("Day", "IF DAY({{Txn date}}) < 10 THEN '0' ELSE '' END + STR(DAY({{Txn date}}))", "string")
    C("Partner first sale", "{FIXED [SoldToPartyID] : MIN(IF [Type] = 'Sales' AND [Gross Sales] > 0 THEN DATE([InvoiceDate]) END)}", "date")

    C("Partner sales in period (INR)", "{FIXED [SoldToPartyID] : SUM(IF {{Is sale in period}} THEN {{Sales value (INR)}} END)}")
    C("Partner slab",
      "IF ISNULL({{Partner sales in period (INR)}}) OR {{Partner sales in period (INR)}} <= 0 THEN NULL "
      "ELSEIF {{Partner sales in period (INR)}} < 100000 THEN '1. Below 1 L' "
      "ELSEIF {{Partner sales in period (INR)}} < 500000 THEN '2. 1 - 5 L' "
      "ELSEIF {{Partner sales in period (INR)}} < 2500000 THEN '3. 5 - 25 L' "
      "ELSEIF {{Partner sales in period (INR)}} < 10000000 THEN '4. 25 L - 1 Cr' "
      "ELSE '5. Above 1 Cr' END", "string")
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
    C("Backlog age (days)", "IF [Type] = 'Backlog' THEN DATEDIFF('day', [OrderCreationDate], TODAY()) END", "integer", fmt=FMT_N)
    C("Backlog ageing",
      "IF [Type] <> 'Backlog' THEN NULL "
      "ELSEIF {{Backlog age (days)}} < 20 THEN '1. Under 20 days' "
      "ELSEIF {{Backlog age (days)}} <= 50 THEN '2. 20-50 days' "
      "ELSEIF {{Backlog age (days)}} <= 100 THEN '3. 51-100 days' "
      "ELSE '4. Over 100 days' END", "string")
    C("Wiser sales CY (row)", "IF {{Is sale in period}} AND [Family Actual (Sheet1)] = 'Wiser' THEN {{Value}} END", fmt=FMT_V)
    C("Zeta sales CY (row)", "IF {{Is sale in period}} AND [Family Actual (Sheet1)] = 'Miluz Zeta' THEN {{Value}} END", fmt=FMT_V)
    C("Wiser sales MTD (row)", "IF {{Is sale MTD}} AND [Family Actual (Sheet1)] = 'Wiser' THEN {{Value}} END", fmt=FMT_V)
    C("Zeta sales MTD (row)", "IF {{Is sale MTD}} AND [Family Actual (Sheet1)] = 'Miluz Zeta' THEN {{Value}} END", fmt=FMT_V)
    for q in (1, 2, 3, 4):
        C(f"Sales Q{q} (row)", f"IF [Type] = 'Sales' AND YEAR([InvoiceDate]) = YEAR({{{{Period end}}}}) AND DATEPART('quarter', [InvoiceDate]) = {q} THEN {{{{Value}}}} END", fmt=FMT_V)
    C("SSA group", "IF ISNULL([SSA]) OR [SSA] = 'No SSA' THEN 'Without SSP' ELSE 'With SSP' END", "string")
    C("SSA sales CY (row)", "IF {{Is sale in period}} AND {{SSA group}} = 'With SSP' THEN {{Value}} END", fmt=FMT_V)
    C("SSA sales PY (row)", "IF {{Is sale in PY period}} AND {{SSA group}} = 'With SSP' THEN {{Value}} END", fmt=FMT_V)
    # ---- v4: three-year view, product groups, reach, partner health, orders
    C("PY2 start", "DATEADD('year', -2, {{Period start}})", "date")
    C("PY2 end", "DATEADD('year', -2, {{Period end}})", "date")
    C("Sales PY2 (row)", "IF [Type] = 'Sales' AND DATE([InvoiceDate]) >= {{PY2 start}} AND DATE([InvoiceDate]) <= {{PY2 end}} THEN {{Value}} END", fmt=FMT_V)
    C("New product category",
      "IF CONTAINS([Family Actual (Sheet1)], 'Miluz Zeta') THEN 'Miluz Zeta' ELSEIF CONTAINS([Family Actual (Sheet1)], 'Wiser') THEN 'Wiser' "
      "ELSEIF CONTAINS([Family Actual (Sheet1)], 'Unica Pure Plus') THEN 'Unica Pure Plus' "
      "ELSEIF CONTAINS([Family Actual (Sheet1)], 'UnicaX') THEN 'UnicaX' ELSE 'Other families' END", "string")
    C("City first sale", "{FIXED [City (Geography mapping)] : MIN(IF [Type] = 'Sales' AND [Gross Sales] > 0 THEN DATE([InvoiceDate]) END)}", "date")
    C("New city (row)", "IF {{Is sale in period}} AND [Gross Sales] <> 0 AND {{City first sale}} >= DATE(DATETRUNC('year', {{Period start}})) THEN [City (Geography mapping)] END", "string")
    C("Active city (row)", "IF {{Is sale in period}} AND [Gross Sales] <> 0 THEN [City (Geography mapping)] END", "string")
    C("Partner sales PY (INR)", "{FIXED [SoldToPartyID] : SUM(IF {{Is sale in PY period}} THEN {{Sales value (INR)}} END)}")
    C("Lapsed partner (row)", "IF (ISNULL({{Partner sales in period (INR)}}) OR {{Partner sales in period (INR)}} <= 0) AND {{Partner sales PY (INR)}} > 0 THEN [SoldToPartyID] END", "string")
    C("Sale date (row)", "IF [Type] = 'Sales' AND [Gross Sales] > 0 THEN DATE([InvoiceDate]) END", "date")
    C("Order", "STR([SalesOrderHeader])", "string")
    C("Order date", "IF [Type] = 'Backlog' THEN STR([OrderCreationDate]) END", "string")
    C("Region sales CY (lod)", "{FIXED [Zone] : SUM({{Sales CY (row)}})}")
    C("Region sales PY (lod)", "{FIXED [Zone] : SUM({{Sales PY (row)}})}")
    C("Region growth (lod)", "IF {{Region sales PY (lod)}} > 0 THEN {{Region sales CY (lod)}} / {{Region sales PY (lod)}} - 1 END")
    C("Top region (row)", "IF {{Region sales CY (lod)}} = {FIXED : MAX({{Region sales CY (lod)}})} THEN {{Region}} END", "string")
    C("Fastest region (row)", "IF {{Region growth (lod)}} = {FIXED : MAX({{Region growth (lod)}})} THEN {{Region}} END", "string")
    C("Branch sales CY (lod)", "{FIXED [Branch (Geography mapping)] : SUM({{Sales CY (row)}})}")
    C("Branch sales PY (lod)", "{FIXED [Branch (Geography mapping)] : SUM({{Sales PY (row)}})}")
    C("Branch growth (lod)", "IF {{Branch sales PY (lod)}} > 0 THEN {{Branch sales CY (lod)}} / {{Branch sales PY (lod)}} - 1 END")
    C("Top branch (row)", "IF {{Branch sales CY (lod)}} = {FIXED : MAX({{Branch sales CY (lod)}})} THEN {{Branch}} END", "string")
    C("Fastest branch (row)", "IF {{Branch growth (lod)}} = {FIXED : MAX({{Branch growth (lod)}})} THEN {{Branch}} END", "string")
    C("Slowest branch (row)", "IF {{Branch growth (lod)}} = {FIXED : MIN({{Branch growth (lod)}})} THEN {{Branch}} END", "string")
    C("Family sales CY (lod)", "{FIXED [Family Actual (Sheet1)] : SUM({{Sales CY (row)}})}")
    C("Family sales PY (lod)", "{FIXED [Family Actual (Sheet1)] : SUM({{Sales PY (row)}})}")
    C("Top family (row)", "IF {{Family sales CY (lod)}} = {FIXED : MAX({{Family sales CY (lod)}})} THEN {{Family}} END", "string")
    # ---- policy categories and new geo (same rules as the Channel Analytics notebook, gross sales, PREVC excluded)
    C("Current year", "YEAR({{Period end}})", "integer", fmt=FMT_N)
    for k, lab in ((0, "CY"), (1, "PY"), (2, "PY2"), (3, "PY3")):
        C(f"Partner gross {lab} (lod)",
          f"{{FIXED [SoldToPartyID] : SUM(IF [Type] = 'Sales' AND [Helios Code] <> 'PREVC' AND YEAR([InvoiceDate]) = {{{{Current year}}}} - {k} THEN [Gross Sales] END)}}")
    C("Partner first sale PY (lod)",
      "{FIXED [SoldToPartyID] : MIN(IF [Type] = 'Sales' AND [Helios Code] <> 'PREVC' AND [Gross Sales] > 0 AND YEAR([InvoiceDate]) = {{Current year}} - 1 THEN DATE([InvoiceDate]) END)}", "date")
    C("Partner new CY", "ZN({{Partner gross CY (lod)}}) > 0 AND ZN({{Partner gross PY (lod)}}) = 0 AND ZN({{Partner gross PY2 (lod)}}) = 0 AND [SoldToPartyID] <> '500073231'", "boolean")
    C("Partner new PY", "ZN({{Partner gross PY (lod)}}) > 0 AND ZN({{Partner gross PY2 (lod)}}) = 0 AND ZN({{Partner gross PY3 (lod)}}) = 0 AND [SoldToPartyID] <> '500073231'", "boolean")
    C("Policy category",
      "IF {{Partner new CY}} THEN 'New Partner ' + STR({{Current year}}) "
      "ELSEIF {{Partner new PY}} AND MONTH({{Partner first sale PY (lod)}}) >= 10 THEN \"Q4'\" + STR({{Current year}} - 1) + ' Onboarded' "
      "ELSEIF {{Partner new PY}} AND ZN({{Partner gross PY (lod)}}) >= 500000 THEN '>5L ' + STR({{Current year}} - 1) + ' Onboarded' "
      "ELSEIF {{Partner new PY}} THEN '<5L ' + STR({{Current year}} - 1) + ' Onboarded' ELSE 'Old Partner' END", "string")
    C("Policy group",
      "IF {{Partner new CY}} OR ({{Partner new PY}} AND (MONTH({{Partner first sale PY (lod)}}) >= 10 OR ZN({{Partner gross PY (lod)}}) < 500000)) THEN 'New Partner' "
      "ELSEIF {{Partner new PY}} THEN 'Handholding Partner' ELSE 'Old Partner' END", "string")
    C("FY PY sales bucket",
      "IF ZN({{Partner gross PY (lod)}}) > 1000000 THEN '1. Above 10 L' ELSEIF ZN({{Partner gross PY (lod)}}) > 500000 THEN '2. 5-10 L' "
      "ELSEIF ZN({{Partner gross PY (lod)}}) > 0 THEN '3. Below 5 L' ELSE '4. Nil last year' END", "string")
    C("Sales FY PY (row)", "IF [Type] = 'Sales' AND YEAR([InvoiceDate]) = {{Current year}} - 1 THEN {{Value}} END", fmt=FMT_V)
    C("New partner sale CY (row)", "IF {{Is sale in period}} AND {{Partner new CY}} THEN {{Value}} END", fmt=FMT_V)
    C("New partner sale PY (row)", "IF {{Is sale in PY period}} AND {{Partner new PY}} THEN {{Value}} END", fmt=FMT_V)
    C("New partner CY (row)", "IF {{Is sale in period}} AND {{Partner new CY}} AND [Gross Sales] <> 0 THEN [SoldToPartyID] END", "string")
    C("New partner PY (row)", "IF {{Is sale in PY period}} AND {{Partner new PY}} AND [Gross Sales] <> 0 THEN [SoldToPartyID] END", "string")
    for k, lab in ((0, "CY"), (1, "PY"), (2, "PY2")):
        C(f"City gross {lab} (lod)",
          f"{{FIXED [City (Geography mapping)] : SUM(IF [Type] = 'Sales' AND [Helios Code] <> 'PREVC' AND YEAR([InvoiceDate]) = {{{{Current year}}}} - {k} THEN [Gross Sales] END)}}")
    C("Geo status",
      "IF ISNULL([City (Geography mapping)]) THEN 'Unmapped' "
      "ELSEIF ZN({{City gross CY (lod)}}) > 0 AND ZN({{City gross PY (lod)}}) = 0 AND ZN({{City gross PY2 (lod)}}) = 0 THEN 'New geo' ELSE 'Existing geo' END", "string")
    C("New geo city (row)", "IF {{Is sale in period}} AND [Gross Sales] <> 0 AND {{Geo status}} = 'New geo' THEN [City (Geography mapping)] END", "string")
    C("New geo partner (row)", "IF {{Is sale in period}} AND [Gross Sales] <> 0 AND {{Geo status}} = 'New geo' THEN [SoldToPartyID] END", "string")
    C("New geo sale (row)", "IF {{Is sale in period}} AND {{Geo status}} = 'New geo' THEN {{Value}} END", fmt=FMT_V)
    # ---- MTD views
    C("Backlog MTD (row)", "IF [Type] = 'Backlog' AND [OrderCreationDate] >= {{MTD start}} THEN {{Value}} END", fmt=FMT_V)
    C("Qty MTD (row)", "IF {{Is sale MTD}} THEN [BillingQuantity] END", "integer", fmt=FMT_N)
    C("OB qty MTD (row)", "IF {{Is OB MTD}} THEN [BillingQuantity] END", "integer", fmt=FMT_N)
    C("ASM MTD sales (lod)", "{FIXED [Zone], [ASM] : SUM({{Sales MTD (row)}})}")
    C("Top ASM MTD", "IF {{ASM MTD sales (lod)}} = {FIXED [Zone] : MAX({{ASM MTD sales (lod)}})} AND {{ASM MTD sales (lod)}} > 0 THEN [ASM] END", "string")
    for key, label in MTD_VIEWS:
        C(f"MTD view: {key}", f"[Parameters].[p MTD view] = '{label}'", "boolean")
    # ---- growth for focus families, decline
    C("Wiser sales PY (row)", "IF {{Is sale in PY period}} AND [Family Actual (Sheet1)] = 'Wiser' THEN {{Value}} END", fmt=FMT_V)
    C("Zeta sales PY (row)", "IF {{Is sale in PY period}} AND [Family Actual (Sheet1)] = 'Miluz Zeta' THEN {{Value}} END", fmt=FMT_V)
    C("Decline (row)", "ZN({{Sales PY (row)}}) - ZN({{Sales CY (row)}})", fmt=FMT_V)
    C("Partner status", "IF {{Partner new CY}} THEN 'New / Revival' ELSE 'Existing' END", "string")
    def explore(param):
        return ("CASE [Parameters].[" + param + "] "
                "WHEN 'Region' THEN {{Region}} WHEN 'Branch' THEN {{Branch}} "
                "WHEN 'ASM' THEN IFNULL([ASM], 'Unmapped') WHEN 'Partner' THEN {{Partner}} "
                "WHEN 'Family' THEN {{Family}} WHEN 'Product classification' THEN IFNULL([Classification (Sheet1)], 'Unmapped') "
                "WHEN 'Activity (Helios)' THEN IFNULL([Helios Code], 'Unmapped') WHEN 'Tier' THEN IFNULL([Tier Category], 'Unmapped') "
                "WHEN 'Partner type' THEN IFNULL([RD/KRO], 'Unmapped') WHEN 'SSP / SSA' THEN IFNULL([SSA], 'Unmapped') "
                "WHEN 'City' THEN IFNULL([City (Geography mapping)], 'Unmapped') WHEN 'Month' THEN {{Month}} "
                "WHEN 'Quarter' THEN {{Quarter}} WHEN 'New product category' THEN {{New product category}} "
                "WHEN 'Partner status' THEN {{Partner status}} "
                "WHEN 'Material code' THEN IFNULL([MaterialCode], 'Unmapped') + '  ' + IFNULL([MaterialDescription], '') "
                "ELSE 'All' END")
    C("Explore by", explore("p Dim"), "string")
    C("Explore across", explore("p Dim2"), "string")
    # value or quantity, as picked in 'Show' (Products and Explore pages)
    C("Shown CY (row)", "IF [Parameters].[p Show] = 'Quantity' THEN FLOAT({{Qty CY (row)}}) ELSE {{Sales CY (row)}} END", fmt=FMT_V)
    C("Shown PY (row)", "IF [Parameters].[p Show] = 'Quantity' THEN FLOAT({{Qty PY (row)}}) ELSE {{Sales PY (row)}} END", fmt=FMT_V)
    # TvA view: achievement matching the selected partner-target view

    # ================================================================ monthly targets (table T, net)
    C("Target month start", "DATE(DATETRUNC('month', [Month and year]))", "date")
    C("Target month end", "DATEADD('day', -1, DATEADD('month', 1, {{Target month start}}))", "date")
    # every month of the Monthly Targets sheet that falls in the period counts with its full target (no pro-rating)
    C("Target month in period", "{{Target month end}} >= {{Period start}} AND {{Target month start}} <= {{Period end}}", "boolean")
    C("Target (net INR)",
      "[Target (Monthly Targets)]")
    C("Target (row)", "IF {{Target month in period}} THEN {{Target (net INR)}} * {{Net to gross}} / {{Units divisor}} END", fmt=FMT_V)
    C("Wiser target (row)", "IF {{Target month in period}} THEN [Wiser] * {{Net to gross}} / {{Units divisor}} END", fmt=FMT_V)
    C("Zeta target (row)", "IF {{Target month in period}} THEN [Miluz Zeta] * {{Net to gross}} / {{Units divisor}} END", fmt=FMT_V)
    C("Is current month target", "{{Target month start}} = {{MTD start}}", "boolean")
    C("Target MTD (row)", "IF {{Is current month target}} THEN {{Target (net INR)}} * {{Net to gross}} / {{Units divisor}} END", fmt=FMT_V)
    C("Wiser target MTD (row)", "IF {{Is current month target}} THEN [Wiser] * {{Net to gross}} / {{Units divisor}} END", fmt=FMT_V)
    C("Zeta target MTD (row)", "IF {{Is current month target}} THEN [Miluz Zeta] * {{Net to gross}} / {{Units divisor}} END", fmt=FMT_V)

    # ================================================================ yearly branch targets (table B, net)
    C("Yearly target (row)",
      "IF [Year (Branch Targets1)] = YEAR({{Period end}}) "
      "THEN [Target (Branch Targets1)] * {{Net to gross}} / {{Units divisor}} END", fmt=FMT_V)

    # ================================================================ SSP targets (table S, as is)
    C("SSP target (row)",
      "IF [Year (SSA Targets)] = YEAR({{Period end}}) "
      "THEN [Target (SSA Targets)] / {{Units divisor}} END", fmt=FMT_V)

    # ================================================================ partner targets (table Q, gross)
    # branch names are used exactly as they are in the files
    C("Target partner branch", "[Branch (QTOD - ATOD)]", "string")
    C("Target partner", "IFNULL([Partner], 'Unknown') + ' (' + [CC] + ')'", "string")
    C("Partner target in scope", "[Target Year] = YEAR({{Period end}})", "boolean")
    # annual QTOD / ATOD targets (gross; Net = gross x 0.83)
    for col, cap in (("Target 100%", "Annual target (100%)"), ("Target 108%", "Annual target (108%)"),
                     ("Zeta Target", "Zeta annual target"), ("Wiser Target", "Wiser annual target")):
        C(f"{cap} (row)", f"IF {{{{Partner target in scope}}}} THEN [{col}] * {{{{Gross to net}}}} / {{{{Units divisor}}}} END", fmt=FMT_V)


    # ================================================================ partner targets (Partner Targets sheet, gross; as the old Partner Ach%)
    C("Partner FY target (row)",
      "IF [Pivot Field Names] = YEAR({{Period end}}) "
      "THEN [Pivot Field Values] * {{Gross to net}} / {{Units divisor}} END", fmt=FMT_V)
    # partner-level achievement (sales and target meet through the Sold-to Party relationship)
    C("Partner FY ach (lod)", "{FIXED [SoldToPartyID] : ZN(MAX({{Partner sales in period (INR)}})) / {{Units divisor}} / MAX({{Partner FY target (row)}})}", fmt=FMT_P)
    C("Partner ach band",
      "IF ISNULL({{Partner FY ach (lod)}}) THEN '7. No target' "
      "ELSEIF {{Partner FY ach (lod)}} >= 1 THEN '1. 100% and above' "
      "ELSEIF {{Partner FY ach (lod)}} >= 0.75 THEN '2. 75-100%' "
      "ELSEIF {{Partner FY ach (lod)}} >= 0.5 THEN '3. 50-75%' "
      "ELSEIF {{Partner FY ach (lod)}} >= 0.25 THEN '4. 25-50%' "
      "ELSEIF {{Partner FY ach (lod)}} > 0 THEN '5. Below 25%' ELSE '6. No sales' END", "string")

    # assigned target: flat 25 L (gross) for new partners, else the Partner Targets sheet
    C("Partner FY target (lod)", "{FIXED [SoldToPartyID] : MAX({{Partner FY target (row)}})}", fmt=FMT_V)
    C("Policy target (row)",
      "IF {{Policy group}} = 'New Partner' THEN 2500000 * {{Gross to net}} / {{Units divisor}} ELSE {{Partner FY target (lod)}} END", fmt=FMT_V)

    # ================================================================ pavilion partners (Pavilion Sales sheet)
    C("Pavilion in scope", "[Active/Removed] = 'Active'", "boolean")
    C("Pavilion KRO (row)", "IF {{Pavilion in scope}} AND [KRO/RO] = 'KRO' THEN [GSTIN no] END", "string")
    C("Pavilion RO (row)", "IF {{Pavilion in scope}} AND [KRO/RO] = 'RO' THEN [GSTIN no] END", "string")
    C("Pavilion target (row)", "IF {{Pavilion in scope}} THEN [2025 Target (Pavilion Sales)] / {{Units divisor}} END", fmt=FMT_V)
    C("Pavilion partner (row)", "IF {{Pavilion in scope}} THEN [GSTIN no] END", "string")

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
        if not rls_rules:
            return "TRUE"       # no branch user filter in the source workbook
        return " OR ".join(parts) if parts else "FALSE"
    C("Branch access (sales)", rls_expr("[Branch (Geography mapping)]"), "boolean")
    C("Branch access (partner targets)", rls_expr("{{Target partner branch}}"), "boolean")

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
    growth("Growth %", "{{Sales CY (row)}}", "{{Sales PY (row)}}")
    A("Monthly target", "SUM({{Target (row)}})")
    pct("Monthly Ach %", "{{Sales CY (row)}}", "{{Target (row)}}")
    A("Gap to monthly target", "SUM({{Sales CY (row)}}) - SUM({{Target (row)}})")
    A("Yearly target", "SUM({{Yearly target (row)}})")
    pct("Yearly Ach %", "{{Sales CY (row)}}", "{{Yearly target (row)}}")
    A("Balance to yearly target", "SUM({{Yearly target (row)}}) - SUM({{Sales CY (row)}})")
    A("Current run-rate / month", "SUM({{Sales CY (row)}}) / MAX(1, DATEDIFF('day', MIN({{Period start}}), MIN(MIN({{Period end}}), TODAY())) + 1) * 30.4")
    A("Required run-rate / month",
      "IF SUM({{Yearly target (row)}}) - SUM({{Sales CY (row)}}) > 0 THEN "
      "(SUM({{Yearly target (row)}}) - SUM({{Sales CY (row)}})) / MAX(0.5, DATEDIFF('day', TODAY(), DATEADD('day', -1, DATEADD('year', 1, DATETRUNC('year', TODAY())))) / 30.4) "
      "ELSE 0 END")
    A("OB", "SUM({{OB CY (row)}})")
    A("OB PY", "SUM({{OB PY (row)}})")
    growth("OB growth %", "{{OB CY (row)}}", "{{OB PY (row)}}")
    pct("Sales to OB %", "{{Sales CY (row)}}", "{{OB CY (row)}}")
    A("Backlog", "SUM({{Backlog (row)}})")
    A("Backlog over 50 days", "SUM(IF {{Backlog age (days)}} > 50 THEN {{Backlog (row)}} END)")
    A("Active partners", "COUNTD({{Active partner CY (row)}})", "integer", FMT_N)
    A("Active partners PY", "COUNTD({{Active partner PY (row)}})", "integer", FMT_N)
    A("New / revival partners", "COUNTD(IF {{Partner status}} = 'New / Revival' THEN {{Active partner CY (row)}} END)", "integer", FMT_N)
    A("New partner sales", "SUM(IF {{Partner status}} = 'New / Revival' THEN {{Sales CY (row)}} END)")
    A("Sales per partner", "IF COUNTD({{Active partner CY (row)}}) > 0 THEN SUM({{Sales CY (row)}}) / COUNTD({{Active partner CY (row)}}) END")
    A("Quantity", "SUM({{Qty CY (row)}})", "integer", FMT_N)
    A("Shown", "SUM({{Shown CY (row)}})", fmt=FMT_V)
    A("Shown PY", "SUM({{Shown PY (row)}})", fmt=FMT_V)
    growth("Shown growth %", "{{Shown CY (row)}}", "{{Shown PY (row)}}")
    A("Quantity PY", "SUM({{Qty PY (row)}})", "integer", FMT_N)
    growth("Quantity growth %", "{{Qty CY (row)}}", "{{Qty PY (row)}}")
    A("Wiser sales", "SUM({{Wiser sales CY (row)}})")
    A("Zeta sales", "SUM({{Zeta sales CY (row)}})")
    A("Wiser target", "SUM({{Wiser target (row)}})")
    A("Zeta target", "SUM({{Zeta target (row)}})")
    pct("Wiser Ach %", "{{Wiser sales CY (row)}}", "{{Wiser target (row)}}")
    pct("Zeta Ach %", "{{Zeta sales CY (row)}}", "{{Zeta target (row)}}")
    A("Share of sales %", "SUM({{Sales CY (row)}}) / MIN({FIXED : SUM({{Sales CY (row)}})})", fmt=FMT_P)
    A("Partners", "COUNTD({{Active partner CY (row)}})", "integer", FMT_N)
    A("Onboarded partners", "COUNTD([SoldToPartyID])", "integer", FMT_N)
    A("Has activity (agg)", "SUM({{Has activity}})", "integer", FMT_N)
    A("Has sales (agg)", "SUM({{Has sales either year}})", "integer", FMT_N)
    A("Has MTD activity (agg)", "SUM({{Has MTD activity}})", "integer", FMT_N)
    S("Data as of", "STR(MAX(IF [Type] = 'Sales' THEN DATE([InvoiceDate]) END))")
    S("Period text", "STR(MIN({{Period start}})) + '  to  ' + STR(MIN({{Period end}}))")
    S("Growth text", "IF ISNULL({{Growth %}}) THEN 'no sales last year' ELSE (IF {{Growth %}} >= 0 THEN '+' ELSE '' END) + STR(ROUND({{Growth %}} * 100, 1)) + '% vs last year' END")
    S("OB growth text", "IF ISNULL({{OB growth %}}) THEN 'no OB last year' ELSE (IF {{OB growth %}} >= 0 THEN '+' ELSE '' END) + STR(ROUND({{OB growth %}} * 100, 1)) + '% vs last year' END")
    A("Nav count", "COUNT([Type])", "integer", FMT_N)
    # text (not a number) behind the navigation buttons: a number that does not fit shows ####
    S("Nav text", "IF COUNT([Type]) > 0 THEN '' END")
    # MTD
    A("MTD sales", "SUM({{Sales MTD (row)}})")
    A("MTD sales LY", "SUM({{Sales MTD LY (row)}})")
    growth("MTD growth %", "{{Sales MTD (row)}}", "{{Sales MTD LY (row)}}")
    A("MTD target", "SUM({{Target MTD (row)}})")
    pct("MTD Ach %", "{{Sales MTD (row)}}", "{{Target MTD (row)}}")
    A("MTD OB", "SUM({{OB MTD (row)}})")
    A("Sales last billing day", "SUM({{Sales last billing day (row)}})")
    A("OB last order day", "SUM({{OB last order day (row)}})")
    A("MTD Wiser sales", "SUM({{Wiser sales MTD (row)}})")
    A("MTD Zeta sales", "SUM({{Zeta sales MTD (row)}})")
    A("MTD Wiser target", "SUM({{Wiser target MTD (row)}})")
    A("MTD Zeta target", "SUM({{Zeta target MTD (row)}})")
    pct("MTD Wiser Ach %", "{{Wiser sales MTD (row)}}", "{{Wiser target MTD (row)}}")
    pct("MTD Zeta Ach %", "{{Zeta sales MTD (row)}}", "{{Zeta target MTD (row)}}")
    A("MTD active partners", "COUNTD({{Active partner MTD (row)}})", "integer", FMT_N)
    S("MTD growth text", "IF ISNULL({{MTD growth %}}) THEN 'no sales last year' ELSE (IF {{MTD growth %}} >= 0 THEN '+' ELSE '' END) + STR(ROUND({{MTD growth %}} * 100, 1)) + '% vs same days last year' END")
    S("Month text", "DATENAME('month', MIN({{MTD start}})) + ' ' + STR(YEAR(MIN({{MTD start}}))) + ', day ' + STR(DAY(TODAY()))")
    # partner targets (TvA)
    A("Has partner target", "SUM(IF {{Partner target in scope}} THEN 1 END)", "integer", FMT_N)
    # SSP
    A("SSP target", "SUM({{SSP target (row)}})")
    A("SSP sales", "SUM({{SSA sales CY (row)}})")
    A("SSP sales PY", "SUM({{SSA sales PY (row)}})")
    pct("SSP Ach %", "{{SSA sales CY (row)}}", "{{SSP target (row)}}")
    growth("SSP growth %", "{{SSA sales CY (row)}}", "{{SSA sales PY (row)}}")

    # ---- v4 aggregates
    A("Sales 2 years ago", "SUM({{Sales PY2 (row)}})")
    A("Lapsed partners", "COUNTD({{Lapsed partner (row)}})", "integer", FMT_N)
    A("New cities", "COUNTD({{New city (row)}})", "integer", FMT_N)
    A("Active cities", "COUNTD({{Active city (row)}})", "integer", FMT_N)
    A("Days since last sale", "DATEDIFF('day', MAX({{Sale date (row)}}), TODAY())", "integer", FMT_N)
    A("Backlog age (max days)", "MAX({{Backlog age (days)}})", "integer", FMT_N)
    A("Partner FY target", "SUM({{Partner FY target (row)}})")
    pct("Partner FY Ach %", "{{Sales CY (row)}}", "{{Partner FY target (row)}}")
    A("Partner FY balance", "SUM({{Partner FY target (row)}}) - SUM({{Sales CY (row)}})")
    A("Partners in band", "COUNTD([SoldToPartyID])", "integer", FMT_N)
    A("Partners at 100%+", "COUNTD(IF {{Partner ach band}} = '1. 100% and above' THEN [SoldToPartyID] END)", "integer", FMT_N)
    A("Partners with target", "COUNTD(IF NOT ISNULL({{Partner FY ach (lod)}}) THEN [SoldToPartyID] END)", "integer", FMT_N)
    A("Pavilions", "COUNTD({{Pavilion partner (row)}})", "integer", FMT_N)
    A("Pavilion target", "SUM({{Pavilion target (row)}})")
    pct("Pavilion Ach %", "{{Sales CY (row)}}", "{{Pavilion target (row)}}")
    A("Has pavilion", "SUM(IF {{Pavilion in scope}} THEN 1 END)", "integer", FMT_N)
    pct("With-SSP share %", "{{SSA sales CY (row)}}", "{{Sales CY (row)}}")
    A("Year gone %", "DATEDIFF('day', DATETRUNC('year', MIN({{Period end}})), MIN({{Period end}})) / 365", fmt=FMT_P)

    # ---- v6 aggregates
    A("Growth (value)", "SUM({{Sales CY (row)}}) - SUM({{Sales PY (row)}})")
    A("Decline (agg)", "IF SUM({{Sales CY (row)}}) < SUM({{Sales PY (row)}}) THEN 1 END", "integer", FMT_N)
    A("Wiser sales PY", "SUM({{Wiser sales PY (row)}})")
    A("Zeta sales PY", "SUM({{Zeta sales PY (row)}})")
    growth("Wiser growth %", "{{Wiser sales CY (row)}}", "{{Wiser sales PY (row)}}")
    growth("Zeta growth %", "{{Zeta sales CY (row)}}", "{{Zeta sales PY (row)}}")
    # yearly targets
    A("Required run-rate (yearly)", "IF SUM({{Yearly target (row)}}) - SUM({{Sales CY (row)}}) > 0 THEN "
      "(SUM({{Yearly target (row)}}) - SUM({{Sales CY (row)}})) / (12 - MONTH(MIN({{Period end}})) + 1) ELSE 0 END")
    # partner (QTOD / ATOD) annual and quarter
    for cap in ("Annual target (100%)", "Annual target (108%)", "Zeta annual target", "Wiser annual target"):
        A(cap, f"SUM({{{{{cap} (row)}}}})")
    pct("Annual Ach %", "{{Sales CY (row)}}", "{{Annual target (100%) (row)}}")
    pct("Annual Ach % (108%)", "{{Sales CY (row)}}", "{{Annual target (108%) (row)}}")
    pct("Zeta annual Ach %", "{{Zeta sales CY (row)}}", "{{Zeta annual target (row)}}")
    pct("Wiser annual Ach %", "{{Wiser sales CY (row)}}", "{{Wiser annual target (row)}}")
    for q in (1, 2, 3, 4):
        A(f"Q{q} sales", f"SUM({{{{Sales Q{q} (row)}}}})")
    # ---- partner target vs achievement, one view per quarter (QTOD), annual, Wiser and Zeta
    for key, label in TARGET_VIEWS:
        C(f"Target view: {key}", f"[Parameters].[p Target view] = '{label}'", "boolean")
    for q in (1, 2, 3, 4):
        col = "[Qtr 1 Target]" if q == 1 else f"FLOAT([Qtr {q} Target])"
        C(f"Q{q} target (row)", f"IF {{{{Partner target in scope}}}} THEN {col} * {{{{Gross to net}}}} / {{{{Units divisor}}}} END", fmt=FMT_V)
        A(f"Q{q} target", f"SUM({{{{Q{q} target (row)}}}})")
        pct(f"Q{q} Ach %", f"{{{{Sales Q{q} (row)}}}}", f"{{{{Q{q} target (row)}}}}")
        A(f"Q{q} balance", f"SUM({{{{Q{q} target (row)}}}}) - SUM({{{{Sales Q{q} (row)}}}})")
        for k in (1, 2, 3):
            mo = (q - 1) * 3 + k
            mon = MONTHS[mo - 1]
            C(f"{mon} sales (row)", f"IF [Type] = 'Sales' AND YEAR([InvoiceDate]) = {{{{Current year}}}} AND MONTH([InvoiceDate]) = {mo} THEN {{{{Value}}}} END", fmt=FMT_V)
            A(f"{mon} sales", f"SUM({{{{{mon} sales (row)}}}})")
            A(f"{mon} Ach %", f"IF SUM({{{{Q{q} target (row)}}}}) <> 0 THEN SUM({{{{{mon} sales (row)}}}}) / (SUM({{{{Q{q} target (row)}}}}) / 3) END", fmt=FMT_P)
        for fam, cap in (("Wiser", "Wiser"), ("Miluz Zeta", "Zeta")):
            C(f"{cap} Q{q} sales (row)", f"IF [Type] = 'Sales' AND YEAR([InvoiceDate]) = {{{{Current year}}}} AND DATEPART('quarter', [InvoiceDate]) = {q} "
              f"AND [Family Actual (Sheet1)] = '{fam}' THEN {{{{Value}}}} END", fmt=FMT_V)
            A(f"{cap} Q{q} sales", f"SUM({{{{{cap} Q{q} sales (row)}}}})")

    # policy tracker (per partner)
    A("FY PY sales", "SUM({{Sales FY PY (row)}})")
    A("Assigned target", "MAX({{Policy target (row)}})")
    A("Policy Ach %", "IF MAX({{Policy target (row)}}) > 0 THEN SUM({{Sales CY (row)}}) / MAX({{Policy target (row)}}) END", fmt=FMT_P)
    A("Req. monthly run-rate", "IF MAX({{Policy target (row)}}) > SUM({{Sales CY (row)}}) THEN "
      "(MAX({{Policy target (row)}}) - ZN(SUM({{Sales CY (row)}}))) / (12 - MONTH(MIN({{Period end}})) + 1) ELSE 0 END")
    A("Partners (count)", "COUNTD([SoldToPartyID])", "integer", FMT_N)
    # new partners and new geo
    A("New partners", "COUNTD({{New partner CY (row)}})", "integer", FMT_N)
    A("New partners PY", "COUNTD({{New partner PY (row)}})", "integer", FMT_N)
    A("New partner sales CY", "SUM({{New partner sale CY (row)}})")
    A("New partner sales PY", "SUM({{New partner sale PY (row)}})")
    growth("New partner growth %", "{{New partner sale CY (row)}}", "{{New partner sale PY (row)}}")
    A("Avg sales per new partner", "IF COUNTD({{New partner CY (row)}}) > 0 THEN SUM({{New partner sale CY (row)}}) / COUNTD({{New partner CY (row)}}) END")
    A("Avg sales per new partner PY", "IF COUNTD({{New partner PY (row)}}) > 0 THEN SUM({{New partner sale PY (row)}}) / COUNTD({{New partner PY (row)}}) END")
    pct("New partner sales %", "{{New partner sale CY (row)}}", "{{Sales CY (row)}}")
    NT = {"East": (30, 50), "North": (50, 130), "South": (50, 130), "West1": (20, 60), "West2": (30, 60)}   # notebook config
    A("New partner count target", " + ".join(f"ZN(MAX(IF [Zone] = '{z}' THEN {c} END))" for z, (c, s) in NT.items()), "integer", FMT_N)
    A("New partner sales target", "(" + " + ".join(f"ZN(MAX(IF [Zone] = '{z}' THEN {s} END))" for z, (c, s) in NT.items())
      + ") * 1000000 * MIN({{Gross to net}}) / MIN({{Units divisor}})")
    A("New partner sales Ach %", "IF {{New partner sales target}} > 0 THEN {{New partner sales CY}} / {{New partner sales target}} END", fmt=FMT_P)
    # new partners as the policy tracker defines them (New Partner group: new this year, Q4 last year, under 5 L last year);
    # used on the New Partner - New Geo page only
    C("Policy new ID (row)", "IF {{Policy group}} = 'New Partner' THEN [SoldToPartyID] END", "string")
    C("Policy new active (row)", "IF {{Policy group}} = 'New Partner' THEN {{Active partner CY (row)}} END", "string")
    C("Policy new sale CY (row)", "IF {{Policy group}} = 'New Partner' THEN {{Sales CY (row)}} END", fmt=FMT_V)
    C("Policy new sale PY (row)", "IF {{Policy group}} = 'New Partner' THEN {{Sales PY (row)}} END", fmt=FMT_V)
    A("Policy new partners", "COUNTD({{Policy new ID (row)}})", "integer", FMT_N)
    A("Policy new buying", "COUNTD({{Policy new active (row)}})", "integer", FMT_N)
    A("Policy new sales", "SUM({{Policy new sale CY (row)}})")
    A("Policy new sales PY", "SUM({{Policy new sale PY (row)}})")
    growth("Policy new growth %", "{{Policy new sale CY (row)}}", "{{Policy new sale PY (row)}}")
    pct("Policy new share %", "{{Policy new sale CY (row)}}", "{{Sales CY (row)}}")
    A("Policy new assigned target", "COUNTD({{Policy new ID (row)}}) * 2500000 * MIN({{Gross to net}}) / MIN({{Units divisor}})")
    A("Policy new Ach %", "IF {{Policy new assigned target}} > 0 THEN {{Policy new sales}} / {{Policy new assigned target}} END", fmt=FMT_P)
    A("Policy new count Ach %", "IF {{New partner count target}} > 0 THEN {{Policy new buying}} / {{New partner count target}} END", fmt=FMT_P)
    A("Policy new target Ach %", "IF {{New partner sales target}} > 0 THEN {{Policy new sales}} / {{New partner sales target}} END", fmt=FMT_P)
    A("Policy new sales per partner", "IF {{Policy new buying}} > 0 THEN {{Policy new sales}} / {{Policy new buying}} END")
    A("New geo cities", "COUNTD({{New geo city (row)}})", "integer", FMT_N)
    A("New geo partners", "COUNTD({{New geo partner (row)}})", "integer", FMT_N)
    A("New geo sales", "SUM({{New geo sale (row)}})")
    # MTD views
    A("MTD backlog (orders this month)", "SUM({{Backlog MTD (row)}})")
    A("Qty sold MTD", "SUM({{Qty MTD (row)}})", "integer", FMT_N)
    A("Qty booked MTD", "SUM({{OB qty MTD (row)}})", "integer", FMT_N)
    # pavilion
    A("KRO pavilions", "COUNTD({{Pavilion KRO (row)}})", "integer", FMT_N)
    A("RO pavilions", "COUNTD({{Pavilion RO (row)}})", "integer", FMT_N)
    A("RO flag (agg)", "IF MIN([KRO/RO]) = 'RO' AND (SUM({{Sales CY (row)}}) > 0 OR SUM({{Sales PY (row)}}) > 0) THEN 1 END", "integer", FMT_N)

    # ---- story headlines: two or three short lines per chapter, computed from the current filters
    N = lambda x: f"IFNULL(STR(ROUND({x}, 1)), '-')"
    P = lambda x: f"IFNULL(STR(ROUND({x} * 100, 1)) + '%', 'n/a')"
    G = lambda x: f"IFNULL((IF {x} >= 0 THEN '+' ELSE '' END) + STR(ROUND({x} * 100, 1)) + '%', 'n/a')"
    U = "' ' + [Parameters].[p Units]"
    I = lambda x: f"IFNULL(STR({x}), '0')"

    def story(cap, *lines):
        for k, f in enumerate(lines, 1):
            S(f"Story | {cap} {k}", f)

    story("Overview",
          f"'Sales: ' + {N('{{Sales}}')} + {U} + ', ' + {G('{{Growth %}}')} + ' on the same period last year.'",
          f"'Yearly target: ' + {P('{{Yearly Ach %}}')} + ' achieved with ' + {P('{{Year gone %}}')} + ' of the year gone; ' + {N('{{Balance to yearly target}}')} + {U} + ' to go.'",
          f"IF {{{{Required run-rate / month}}}} > {{{{Current run-rate / month}}}} THEN 'Run-rate needed for the yearly target: ' + {N('{{Required run-rate / month}}')} + {U} + ' a month, against ' + {N('{{Current run-rate / month}}')} + {U} + ' so far.'"
          f" ELSE 'At the current run-rate of ' + {N('{{Current run-rate / month}}')} + {U} + ' a month the yearly target is on track.' END")
    story("Regions",
          f"'Largest region: ' + IFNULL(MAX({{{{Top region (row)}}}}), '-') + ' with ' + {P('MAX({{Region sales CY (lod)}}) / SUM({{Sales CY (row)}})')} + ' of sales.  Fastest growth: ' + IFNULL(MAX({{{{Fastest region (row)}}}}), '-') + ' (' + {G('MAX({{Region growth (lod)}})')} + ').'",
          f"{I('COUNTD(IF {{Region growth (lod)}} > 0 THEN [Zone] END)')} + ' of ' + {I('COUNTD(IF {{Region sales CY (lod)}} > 0 THEN [Zone] END)')} + ' regions are growing.  Click a region in the scorecard to filter the page.'")
    story("Branches",
          f"{I('COUNTD(IF {{Branch growth (lod)}} > 0 THEN [Branch (Geography mapping)] END)')} + ' of ' + {I('COUNTD(IF {{Branch sales CY (lod)}} > 0 THEN [Branch (Geography mapping)] END)')} + ' branches are growing.  Largest: ' + IFNULL(MAX({{{{Top branch (row)}}}}), '-') + ' (' + {P('MAX({{Branch sales CY (lod)}}) / SUM({{Sales CY (row)}})')} + ' of sales).'",
          f"'Fastest growth: ' + IFNULL(MAX({{{{Fastest branch (row)}}}}), '-') + ' (' + {G('MAX({{Branch growth (lod)}})')} + ');  weakest: ' + IFNULL(MAX({{{{Slowest branch (row)}}}}), '-') + ' (' + {G('MIN({{Branch growth (lod)}})')} + ').  Click a branch to see its ASMs.'")
    story("Products",
          f"'Top family: ' + IFNULL(MAX({{{{Top family (row)}}}}), '-') + ' with ' + {P('MAX({{Family sales CY (lod)}}) / SUM({{Sales CY (row)}})')} + ' of sales.'",
          f"'Wiser: ' + {N('{{Wiser sales}}')} + {U} + ' (' + {P('{{Wiser Ach %}}')} + ' of target).  Miluz Zeta: ' + {N('{{Zeta sales}}')} + {U} + ' (' + {P('{{Zeta Ach %}}')} + ' of target).'")
    story("Partners",
          f"{I('{{Active partners}}')} + ' partners bought in the period (' + {I('{{Active partners PY}}')} + ' in the same period last year).  ' + {I('{{New / revival partners}}')} + ' are new or revived; ' + {I('{{Lapsed partners}}')} + ' bought last year but not this year.'",
          f"'Average sales per partner: ' + {N('{{Sales per partner}}')} + {U} + '.  ' + {I('{{Partners at 100%+}}')} + ' of ' + {I('{{Partners with target}}')} + ' partners with a target are at 100% or more.'")
    story("All partners",
          "'Every partner with sales, orders or backlog in the period or last year: ' + IFNULL(STR(COUNTD(IF {{Has activity}} = 1 THEN [SoldToPartyID] END)), '0') + ' partners.'")
    story("Targets",
          f"'Monthly targets: ' + {N('{{Monthly target}}')} + {U} + ', sales ' + {N('{{Sales}}')} + {U} + ', achievement ' + {P('{{Monthly Ach %}}')} + ' (gap ' + {N('{{Gap to monthly target}}')} + {U} + ').'",
          f"'Yearly target: ' + {P('{{Yearly Ach %}}')} + ' achieved, ' + {N('{{Balance to yearly target}}')} + {U} + ' to go.  Partner targets: ' + {P('{{Partner FY Ach %}}')} + ' achieved overall.'")
    story("SSP",
          f"'Partners with an SSP bring ' + {P('{{With-SSP share %}}')} + ' of sales; SSP sales are ' + {G('{{SSP growth %}}')} + ' on last year and ' + {P('{{SSP Ach %}}')} + ' of SSP targets.'",
          f"{I('{{Active cities}}')} + ' cities bought in the period, ' + {I('{{New cities}}')} + ' of them for the first time.'")
    story("Pavilion",
          f"{I('{{Pavilions}}')} +  ' active pavilion partners, with a target of ' + {N('{{Pavilion target}}')} + {U} + '.  Their primary purchases (matched by GSTIN): ' + {N('{{Sales}}')} + {U} + ' (' + {P('{{Pavilion Ach %}}')} + ').'")
    story("MTD",
          f"'Month to date: ' + {N('{{MTD sales}}')} + {U} + ', ' + {G('{{MTD growth %}}')} + ' on the same days last year and ' + {P('{{MTD Ach %}}')} + ' of the target for this month.'",
          f"'Last billing day (' + STR(MIN({{Last sales date}})) + '): ' + {N('{{Sales last billing day}}')} + {U} + ' sales.  Last order day (' + STR(MIN({{Last OB date}})) + '): ' + {N('{{OB last order day}}')} + {U} + ' orders.'")
    story("Orders",
          f"'Order book: ' + {N('{{OB}}')} + {U} + ' (' + {G('{{OB growth %}}')} + ' on last year); sales are ' + {P('{{Sales to OB %}}')} + ' of it.'",
          f"'Open backlog: ' + {N('{{Backlog}}')} + {U} + ', of which ' + {P('{{Backlog over 50 days}} / {{Backlog}}')} +  ' was ordered more than 50 days ago.'")

    # ================================================================ Pavilion (summary tables, as published)
    PC = lambda cap, f, dt="real", fmt=FMT_N: C(cap, f, dt, role="measure", fmt=fmt, ds=PDS)
    PC("Min slab pavilions", f"SUM([No. of KRO {MS}]) + SUM([No. of RO {MS}])", "integer")
    PC("Min slab qualified", "SUM([No of KRO Achieved]) + SUM([No. of RO Achieved])", "integer")
    PC("Min slab qualifier %", f"IF SUM([No. of KRO {MS}]) + SUM([No. of RO {MS}]) > 0 THEN (SUM([No of KRO Achieved]) + SUM([No. of RO Achieved])) / (SUM([No. of KRO {MS}]) + SUM([No. of RO {MS}])) END", fmt=FMT_P)
    PC("Total target qualified", f"SUM([No of KRO Achieved {TT}]) + SUM([No. of RO Achieved {TT}])", "integer")
    PC("Total target qualifier %", f"IF SUM([No. of KRO {TT}]) + SUM([No. of RO {TT}]) > 0 THEN (SUM([No of KRO Achieved {TT}]) + SUM([No. of RO Achieved {TT}])) / (SUM([No. of KRO {TT}]) + SUM([No. of RO {TT}])) END", fmt=FMT_P)
    PC("Wiser qualified", "SUM([No of KRO > 1 lac]) + SUM([No. of RO > 25K])", "integer")
    PC("Wiser qualifier %", "IF SUM([No. of KRO]) + SUM([No. of RO]) > 0 THEN (SUM([No of KRO > 1 lac]) + SUM([No. of RO > 25K])) / (SUM([No. of KRO]) + SUM([No. of RO])) END", fmt=FMT_P)
    PC("Acceptance %", "IF SUM([Total no. of RO Pavilions]) > 0 THEN SUM([Accepted]) / SUM([Total no. of RO Pavilions]) END", fmt=FMT_P)
    PC("Pavilion growth %", "SUM([2025-26 Growth %])", fmt=FMT_P)
    PC("Pavilion class growth %", f"SUM([2025-26 Growth % {SC}])", fmt=FMT_P)
    PC("Total pavilions (file)", f"SUM(IF [Row Type {PT}] = 'Grand Total' THEN [Count of 2026 Classification] END)", "integer")
    C("Overall sales slab", "MIN(IF [Row Type] = 'Overall Total' THEN [2025 Sales Slab] END) + '  >  ' + MIN(IF [Row Type] = 'Overall Total' THEN [2026 Sales Slab] END)",
      "string", role="measure", ftype="nominal", ds=PDS)
    PC("Overall slab growth %", "SUM(IF [Row Type] = 'Overall Total' THEN [2025-26 Growth %] END)", fmt=FMT_P)
    PC("Slab row order", "MIN(IF [Row Type] = 'Overall Total' THEN 3 ELSEIF [Row Type] = 'Region Total' THEN 2 ELSE 1 END)", "integer")
    PC("Type row order", f"MIN(IF [Row Type {PT}] = 'Grand Total' THEN 2 ELSE 1 END)", "integer")
    PC("Class row order", f"MIN(IF [Row Type {SC}] = 'Grand Total' THEN 2 ELSE 1 END)", "integer")
    return R
