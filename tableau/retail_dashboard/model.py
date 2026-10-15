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
from core import MERGE, OLD_CUR, OLD_END, OLD_GN, OLD_START, PDS, Registry

GROSS_FACTOR = 0.83
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
    if MERGE:
        # the existing workbook's own controls (definitions are taken from that workbook when merging)
        P(OLD_START, "Start Date", "date", "2026-01-01")
        P(OLD_END, "End Date", "date", "2026-10-14")
        P(OLD_GN, "Gross / Net", "string", "Gross Sales", ["Gross Sales", "Net Sales"])
        P(OLD_CUR, "Currency Selector", "string", "MINR", ["INR", "MINR"])
    else:
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
    P("p Basis", "Monthly target basis", "string", "Full months", ["Full months", "Pro-rated to date"])
    P("p View", "Target view", "string", "Annual (100%)",
      ["Annual (100%)", "Annual (108%)", "Q1", "Q2", "Q3", "Q4", "Wiser", "Miluz Zeta"])
    P("p Dim", "Explore by", "string", "Branch",
      ["Region", "Branch", "ASM", "Partner", "Family", "Product classification", "Activity (Helios)",
       "Tier", "Partner type", "SSP / SSA", "City", "Month", "Quarter"])

    C = R.calc
    if MERGE:
        PU, PM, GV, NV = f"[Parameters].[{OLD_CUR}]", f"[Parameters].[{OLD_GN}]", "Gross Sales", "Net Sales"
    else:
        PU, PM, GV, NV = "[Parameters].[p Units]", "[Parameters].[p Measure]", "Gross", "Net"

    # ================================================================ parameter-only helpers
    if MERGE:
        C("Units divisor", f"IF {PU} = 'MINR' THEN 1000000 ELSE 1 END")
        C("Period start", f"DATE([Parameters].[{OLD_START}])", "date")
        C("Period end", f"DATE([Parameters].[{OLD_END}])", "date")
    else:
        C("Units divisor", f"CASE {PU} WHEN 'INR' THEN 1 WHEN 'Lakh' THEN 100000 WHEN 'Mn' THEN 1000000 WHEN 'Cr' THEN 10000000 ELSE 1000000 END")
    if not MERGE:
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
    C("Region selected", "[Parameters].[p Region] = 'All' OR [Zone] = [Parameters].[p Region]", "boolean")
    C("In filter scope",
      "([Parameters].[p Branch] = 'All' OR [Branch (Geography mapping)] = [Parameters].[p Branch]) "
      "AND ([Parameters].[p Family] = 'All' OR [Family Actual (Sheet1)] = [Parameters].[p Family]) "
      "AND ([Parameters].[p Channel] = 'All' OR [RD/KRO] = [Parameters].[p Channel])", "boolean")
    C("Is sale in period", "[Type] = 'Sales' AND {{In filter scope}} AND DATE([InvoiceDate]) >= {{Period start}} AND DATE([InvoiceDate]) <= {{Period end}}", "boolean")
    C("Is sale in PY period", "[Type] = 'Sales' AND {{In filter scope}} AND DATE([InvoiceDate]) >= {{PY start}} AND DATE([InvoiceDate]) <= {{PY end}}", "boolean")
    C("Is OB in period", "[Type] = 'OB' AND {{In filter scope}} AND [Calendar Day] >= {{Period start}} AND [Calendar Day] <= {{Period end}}", "boolean")
    C("Is OB in PY period", "[Type] = 'OB' AND {{In filter scope}} AND [Calendar Day] >= {{PY start}} AND [Calendar Day] <= {{PY end}}", "boolean")
    C("Is sale MTD", "[Type] = 'Sales' AND {{In filter scope}} AND DATE([InvoiceDate]) >= {{MTD start}} AND DATE([InvoiceDate]) <= TODAY()", "boolean")
    C("Is sale MTD LY", "[Type] = 'Sales' AND {{In filter scope}} AND DATE([InvoiceDate]) >= {{MTD LY start}} AND DATE([InvoiceDate]) <= {{MTD LY end}}", "boolean")
    C("Is OB MTD", "[Type] = 'OB' AND {{In filter scope}} AND [Calendar Day] >= {{MTD start}} AND [Calendar Day] <= TODAY()", "boolean")
    C("Sales CY (row)", "IF {{Is sale in period}} THEN {{Value}} END", fmt=FMT_V)
    C("Sales PY (row)", "IF {{Is sale in PY period}} THEN {{Value}} END", fmt=FMT_V)
    C("OB CY (row)", "IF {{Is OB in period}} THEN {{Value}} END", fmt=FMT_V)
    C("OB PY (row)", "IF {{Is OB in PY period}} THEN {{Value}} END", fmt=FMT_V)
    C("Sales MTD (row)", "IF {{Is sale MTD}} THEN {{Value}} END", fmt=FMT_V)
    C("Sales MTD LY (row)", "IF {{Is sale MTD LY}} THEN {{Value}} END", fmt=FMT_V)
    C("OB MTD (row)", "IF {{Is OB MTD}} THEN {{Value}} END", fmt=FMT_V)
    C("Sales yesterday (row)", "IF [Type] = 'Sales' AND {{In filter scope}} AND DATE([InvoiceDate]) = DATEADD('day', -1, TODAY()) THEN {{Value}} END", fmt=FMT_V)
    C("OB yesterday (row)", "IF [Type] = 'OB' AND {{In filter scope}} AND [Calendar Day] = DATEADD('day', -1, TODAY()) THEN {{Value}} END", fmt=FMT_V)
    C("Backlog (row)", "IF [Type] = 'Backlog' AND {{In filter scope}} THEN {{Value}} END", fmt=FMT_V)
    C("Qty CY (row)", "IF {{Is sale in period}} THEN [BillingQuantity] END", "integer", fmt=FMT_N)
    C("Active partner CY (row)", "IF {{Is sale in period}} AND [Gross Sales] <> 0 THEN [SoldToPartyID] END", "string")
    C("Active partner PY (row)", "IF {{Is sale in PY period}} AND [Gross Sales] <> 0 THEN [SoldToPartyID] END", "string")
    C("Active partner MTD (row)", "IF {{Is sale MTD}} AND [Gross Sales] <> 0 THEN [SoldToPartyID] END", "string")
    C("Has activity", "IF {{Is sale in period}} OR {{Is sale in PY period}} OR {{Is OB in period}} OR ([Type] = 'Backlog' AND {{In filter scope}}) THEN 1 END", "integer")
    C("Has sales either year", "IF {{Is sale in period}} OR {{Is sale in PY period}} THEN 1 END", "integer")
    C("Has MTD activity", "IF {{Is sale MTD}} OR {{Is sale MTD LY}} OR {{Is OB MTD}} THEN 1 END", "integer")
    C("Txn date", "CASE [Type] WHEN 'Sales' THEN DATE([InvoiceDate]) WHEN 'OB' THEN [Calendar Day] ELSE [OrderCreationDate] END", "date")
    C("Month", "IF MONTH({{Txn date}}) < 10 THEN '0' ELSE '' END + STR(MONTH({{Txn date}})) + ' ' + LEFT(DATENAME('month', {{Txn date}}), 3)", "string")
    C("Quarter", "'Q' + STR(DATEPART('quarter', {{Txn date}}))", "string")
    C("Day", "IF DAY({{Txn date}}) < 10 THEN '0' ELSE '' END + STR(DAY({{Txn date}}))", "string")
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
    C("Backlog age (days)", "IF [Type] = 'Backlog' THEN DATEDIFF('day', [OrderCreationDate], TODAY()) END", "integer", fmt=FMT_N)
    C("Backlog ageing",
      "IF [Type] <> 'Backlog' THEN NULL "
      "ELSEIF {{Backlog age (days)}} <= 30 THEN '1. 0-30 days' "
      "ELSEIF {{Backlog age (days)}} <= 60 THEN '2. 31-60 days' "
      "ELSEIF {{Backlog age (days)}} <= 90 THEN '3. 61-90 days' "
      "ELSE '4. Over 90 days' END", "string")
    C("Wiser sales CY (row)", "IF {{Is sale in period}} AND [Family Actual (Sheet1)] = 'Wiser' THEN {{Value}} END", fmt=FMT_V)
    C("Zeta sales CY (row)", "IF {{Is sale in period}} AND [Family Actual (Sheet1)] = 'Miluz Zeta' THEN {{Value}} END", fmt=FMT_V)
    C("Wiser sales MTD (row)", "IF {{Is sale MTD}} AND [Family Actual (Sheet1)] = 'Wiser' THEN {{Value}} END", fmt=FMT_V)
    C("Zeta sales MTD (row)", "IF {{Is sale MTD}} AND [Family Actual (Sheet1)] = 'Miluz Zeta' THEN {{Value}} END", fmt=FMT_V)
    for q in (1, 2, 3, 4):
        C(f"Sales Q{q} (row)", f"IF [Type] = 'Sales' AND {{{{In filter scope}}}} AND YEAR([InvoiceDate]) = YEAR({{{{Period end}}}}) AND DATEPART('quarter', [InvoiceDate]) = {q} THEN {{{{Value}}}} END", fmt=FMT_V)
    C("SSA group", "IF ISNULL([SSA]) OR [SSA] = 'No SSA' THEN 'Without SSP' ELSE 'With SSP' END", "string")
    C("SSA sales CY (row)", "IF {{Is sale in period}} AND {{SSA group}} = 'With SSP' THEN {{Value}} END", fmt=FMT_V)
    C("SSA sales PY (row)", "IF {{Is sale in PY period}} AND {{SSA group}} = 'With SSP' THEN {{Value}} END", fmt=FMT_V)
    C("Explore by",
      "CASE [Parameters].[p Dim] "
      "WHEN 'Region' THEN {{Region}} WHEN 'Branch' THEN {{Branch}} "
      "WHEN 'ASM' THEN IFNULL([ASM], 'Unmapped') WHEN 'Partner' THEN {{Partner}} "
      "WHEN 'Family' THEN {{Family}} WHEN 'Product classification' THEN IFNULL([Classification (Sheet1)], 'Unmapped') "
      "WHEN 'Activity (Helios)' THEN IFNULL([Helios Code], 'Unmapped') WHEN 'Tier' THEN IFNULL([Tier Category], 'Unmapped') "
      "WHEN 'Partner type' THEN IFNULL([RD/KRO], 'Unmapped') WHEN 'SSP / SSA' THEN IFNULL([SSA], 'Unmapped') "
      "WHEN 'City' THEN IFNULL([City (Geography mapping)], 'Unmapped') WHEN 'Month' THEN {{Month}} "
      "WHEN 'Quarter' THEN {{Quarter}} END", "string")
    # TvA view: achievement matching the selected partner-target view
    C("View sales (row)",
      "CASE [Parameters].[p View] WHEN 'Q1' THEN {{Sales Q1 (row)}} WHEN 'Q2' THEN {{Sales Q2 (row)}} "
      "WHEN 'Q3' THEN {{Sales Q3 (row)}} WHEN 'Q4' THEN {{Sales Q4 (row)}} "
      "WHEN 'Wiser' THEN {{Wiser sales CY (row)}} WHEN 'Miluz Zeta' THEN {{Zeta sales CY (row)}} "
      "ELSE {{Sales CY (row)}} END", fmt=FMT_V)

    # ================================================================ monthly targets (table T, net)
    C("Target branch", "CASE [Branch] WHEN 'NCR(DEL+ROH)' THEN 'NCR' WHEN 'NCR(NOI+GHAZ)' THEN 'NOIDA-GHZ' ELSE [Branch] END", "string")
    C("Target in scope",
      "([Parameters].[p Branch] = 'All' OR {{Target branch}} = [Parameters].[p Branch]) AND [Parameters].[p Channel] = 'All'", "boolean")
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
    C("Target (row)", "IF {{Target in scope}} THEN {{Target (net INR)}} * {{Net to gross}} * {{Target month share}} / {{Units divisor}} END", fmt=FMT_V)
    C("Wiser target (row)", "IF {{Target in scope}} THEN [Wiser] * {{Net to gross}} * {{Target month share}} / {{Units divisor}} END", fmt=FMT_V)
    C("Zeta target (row)", "IF {{Target in scope}} THEN [Miluz Zeta] * {{Net to gross}} * {{Target month share}} / {{Units divisor}} END", fmt=FMT_V)
    C("Is current month target", "{{Target month start}} = {{MTD start}}", "boolean")
    C("Target MTD (row)", "IF {{Target in scope}} AND {{Is current month target}} THEN {{Target (net INR)}} * {{Net to gross}} / {{Units divisor}} END", fmt=FMT_V)
    C("Wiser target MTD (row)", "IF {{Target in scope}} AND {{Is current month target}} THEN [Wiser] * {{Net to gross}} / {{Units divisor}} END", fmt=FMT_V)
    C("Zeta target MTD (row)", "IF {{Target in scope}} AND {{Is current month target}} THEN [Miluz Zeta] * {{Net to gross}} / {{Units divisor}} END", fmt=FMT_V)

    # ================================================================ yearly branch targets (table B, net)
    C("Yearly target (row)",
      "IF [Year (Branch Targets1)] = YEAR({{Period end}}) "
      "AND ([Parameters].[p Branch] = 'All' OR [Branch (Branch Targets1)] = [Parameters].[p Branch]) "
      "AND [Parameters].[p Family] = 'All' AND [Parameters].[p Channel] = 'All' "
      "THEN [Target (Branch Targets1)] * {{Net to gross}} / {{Units divisor}} END", fmt=FMT_V)

    # ================================================================ SSP targets (table S, as is)
    C("SSP target (row)",
      "IF [Year (SSA Targets)] = YEAR({{Period end}}) "
      "AND ([Parameters].[p Branch] = 'All' OR [Branch (SSA Targets)] = [Parameters].[p Branch]) "
      "THEN [Target (SSA Targets)] / {{Units divisor}} END", fmt=FMT_V)

    # ================================================================ partner targets (table Q, gross)
    C("Target partner branch",
      "CASE UPPER(TRIM([Branch (QTOD - ATOD)])) WHEN 'TN -1' THEN 'TN-1' WHEN 'TN2' THEN 'TN-2' "
      "ELSE UPPER(TRIM([Branch (QTOD - ATOD)])) END", "string")
    zone_case = " ".join(f"WHEN '{br}' THEN '{z}'" for br, z in sorted(branch_zone.items()))
    C("Target partner region", f"CASE {{{{Target partner branch}}}} {zone_case} ELSE 'Other' END", "string")
    C("Target partner region selected", "[Parameters].[p Region] = 'All' OR {{Target partner region}} = [Parameters].[p Region]", "boolean")
    C("Target partner", "IFNULL([Partner], 'Unknown') + ' (' + [CC] + ')'", "string")
    C("Partner target in scope",
      "[Target Year] = YEAR({{Period end}}) "
      "AND ([Parameters].[p Branch] = 'All' OR {{Target partner branch}} = [Parameters].[p Branch]) "
      "AND ([Parameters].[p Channel] = 'All' OR TRIM([RD/ KRO]) = [Parameters].[p Channel])", "boolean")
    C("View target (row)",
      "IF {{Partner target in scope}} THEN (CASE [Parameters].[p View] "
      "WHEN 'Annual (108%)' THEN [Target 108%] WHEN 'Q1' THEN [Qtr 1 Target] "
      "WHEN 'Q2' THEN FLOAT([Qtr 2 Target]) WHEN 'Q3' THEN FLOAT([Qtr 3 Target]) WHEN 'Q4' THEN FLOAT([Qtr 4 Target]) "
      "WHEN 'Wiser' THEN [Wiser Target] WHEN 'Miluz Zeta' THEN [Zeta Target] ELSE [Target 100%] END) "
      "* {{Gross to net}} / {{Units divisor}} END", fmt=FMT_V)

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
    A("Growth", "SUM({{Sales CY (row)}}) - SUM({{Sales PY (row)}})")
    A("Target", "SUM({{Target (row)}})")
    pct("Ach %", "{{Sales CY (row)}}", "{{Target (row)}}")
    A("Gap to target", "SUM({{Sales CY (row)}}) - SUM({{Target (row)}})")
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
    A("Backlog over 60 days", "SUM(IF {{Backlog age (days)}} > 60 THEN {{Backlog (row)}} END)")
    A("Active partners", "COUNTD({{Active partner CY (row)}})", "integer", FMT_N)
    A("Active partners PY", "COUNTD({{Active partner PY (row)}})", "integer", FMT_N)
    A("New / revival partners", "COUNTD(IF {{Partner status}} = 'New / Revival' THEN {{Active partner CY (row)}} END)", "integer", FMT_N)
    A("New partner sales", "SUM(IF {{Partner status}} = 'New / Revival' THEN {{Sales CY (row)}} END)")
    A("Sales per partner", "IF COUNTD({{Active partner CY (row)}}) > 0 THEN SUM({{Sales CY (row)}}) / COUNTD({{Active partner CY (row)}}) END")
    A("Quantity", "SUM({{Qty CY (row)}})", "integer", FMT_N)
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
    # MTD
    A("MTD sales", "SUM({{Sales MTD (row)}})")
    A("MTD sales LY", "SUM({{Sales MTD LY (row)}})")
    growth("MTD growth %", "{{Sales MTD (row)}}", "{{Sales MTD LY (row)}}")
    A("MTD target", "SUM({{Target MTD (row)}})")
    pct("MTD Ach %", "{{Sales MTD (row)}}", "{{Target MTD (row)}}")
    A("MTD OB", "SUM({{OB MTD (row)}})")
    A("Sales yesterday", "SUM({{Sales yesterday (row)}})")
    A("OB yesterday", "SUM({{OB yesterday (row)}})")
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
    A("Partner target", "SUM({{View target (row)}})")
    A("Partner achieved", "SUM({{View sales (row)}})")
    pct("Partner Ach %", "{{View sales (row)}}", "{{View target (row)}}")
    A("Partner balance", "SUM({{View target (row)}}) - SUM({{View sales (row)}})")
    A("Has partner target", "SUM(IF {{Partner target in scope}} THEN 1 END)", "integer", FMT_N)
    # SSP
    A("SSP target", "SUM({{SSP target (row)}})")
    A("SSP sales", "SUM({{SSA sales CY (row)}})")
    A("SSP sales PY", "SUM({{SSA sales PY (row)}})")
    pct("SSP Ach %", "{{SSA sales CY (row)}}", "{{SSP target (row)}}")
    growth("SSP growth %", "{{SSA sales CY (row)}}", "{{SSA sales PY (row)}}")

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
    PC("Slab row order", "MIN(IF [Row Type] = 'Overall Total' THEN 3 ELSEIF [Row Type] = 'Region Total' THEN 2 ELSE 1 END)", "integer")
    PC("Type row order", f"MIN(IF [Row Type {PT}] = 'Grand Total' THEN 2 ELSE 1 END)", "integer")
    PC("Class row order", f"MIN(IF [Row Type {SC}] = 'Grand Total' THEN 2 ELSE 1 END)", "integer")
    return R
