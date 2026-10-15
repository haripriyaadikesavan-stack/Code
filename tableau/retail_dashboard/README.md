# Retail Dashboard (Tableau) — generator

Builds `Retail Dashboard.twbx` in the eCommerce dashboard's style:
- the same green theme;
- a navigation bar on every page;
- one filter panel on every page;
- KPI cards on top.

It reuses the data source of `Overall Retail Dashboard.twbx` unchanged:
- the same SharePoint files and relationships;
- the same extract, so the refresh works as before.

All old calculations, sheets and dashboards are replaced.

```
python build.py "Overall Retail Dashboard.twbx" "Retail Dashboard.twbx"
python validate.py "Retail Dashboard.twb"      # static checks on the generated XML
python order_check.py "Retail Dashboard.twb" "Overall Retail Dashboard.twb" "eCommerce_Analysis.twb"   # element order / attributes vs real Tableau files
```

With `RETAIL_MERGE=1`, only the analysis pages are built (Executive, Branches, Products, Partners, Target vs Achievement, Order Book, Explore). They use the existing workbook's own Gross/Net, Currency Selector, Start Date and End Date parameters, and "Home" opens its Retail Overall Dashboard. `../retail_dashboard_qc/merge.py` adds them to that workbook.

Needs `lxml`. The `.twbx` / `.twb` / `.hyper` files contain business data and
are git-ignored.

## Pages

| Page | What it answers |
|---|---|
| Home | Headline KPIs, region table, monthly trend |
| MTD | Month to date vs the same days last year, this month's target, yesterday's sales / OB, branch and ASM MTD tables, top partners |
| Executive | Sales vs last year and vs target, region table, branch achievement and growth |
| Regions | One dashboard per region (East, North, South, West1, West2), always filtered to that region |
| Branches | Branch scorecard, Wiser / Miluz Zeta by branch, ASM table |
| Products | Families, Helios activity, product classification, focus families by region, top 25 materials |
| Partners | Slabs, tier, RD/KRO, new / revival partners, onboarded year, top and at-risk partners |
| Target vs Achievement | Branch target vs achievement (monthly and yearly targets), branch x month heat map, partner targets (annual 100% / 108%, quarter, Wiser, Miluz Zeta) |
| SSP | SSP target vs achievement, with vs without SSP, monthly trend, top SSP partners |
| Pavilion | Pavilion type and classification slabs, minimum slab / total target / Wiser slab qualification, retailer app T&C |
| Order Book | OB vs last year, sales-to-OB conversion, backlog by branch and age |
| Explore | Any dimension against all measures |

## Filters (left panel, every page)

- **Period:** YTD, QTD, MTD, Last month, Full year, or Custom range. Periods end today.
- **Gross / Net.**
- **Units:** INR, Lakh, Mn, Cr.
- **Region, Branch, Family, Partner type.**
- **Target basis:** "Pro-rated to date" or "Full months".

## Business rules

- **Sales, OB and Backlog dates:**
  - Sales = `Type = Sales`, dated by `InvoiceDate`.
  - OB = `Type = OB`, dated by `Calendar Day`.
  - Backlog = `Type = Backlog`, the open orders. They are aged from `OrderCreationDate`.
- **Last year:** the same dates one year earlier. Growth % compares against this.
- **Monthly branch targets** (Monthly Targets sheet) are **net**. Gross = net ÷ 0.83.
  "Full months" (default) counts each month's full target; "Pro-rated to date" counts the current month only up to today.
- **Yearly branch targets** (Branch Targets sheet) are **net**. Gross = net ÷ 0.83.
- **Partner targets** (QTOD and ATOD file) are **gross**. Net = gross × 0.83.
- **SSP targets** are used as they are (as in the old workbook).
- **New / revival partner:** the first sale is in the current year. The data starts two years back, so this also means no sales in the previous two years.
- **Region filter:** applied as a sheet filter, so the region dashboards always show their own region.
- **Pavilion:** shown as published in Pavilion_Unpivoted.xlsx (the slabs are text ranges); the filters do not apply.
- **Row-level security:** the old *Branch RLS* user filter is rebuilt as two calculations, `Branch access (sales)` and `Branch access (partner targets)`. They use the same Tableau groups. They are applied to the partner-level pages (Partners, Partner Targets, Explore, Backlog partners), where the old workbook applied it.
