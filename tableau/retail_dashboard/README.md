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

Needs `lxml`. The `.twbx` / `.twb` / `.hyper` files contain business data and
are git-ignored.

## Pages

| Page | What it answers |
|---|---|
| Home | Headline KPIs, region table, monthly trend |
| Executive | Are we on track? Sales vs last year and vs target, region table, branch achievement and growth |
| Branches | Branch scorecard (replaces the 20+ per-branch dashboards), Wiser / Miluz Zeta by branch, ASM table |
| Products | Families, Helios activity, product classification, top 25 materials, focus families by region |
| Partners | Partner slabs, tier, RD/KRO, new / revival partners, top and at-risk partners, SSA |
| Partner Targets | QTOD / ATOD partner targets vs achievement, annual and by quarter |
| Order Book | OB vs last year, sales-to-OB conversion, backlog by branch and age |
| Explore | Any dimension (Region, Branch, ASM, Partner, Family, …) against all measures |

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
- **Monthly branch targets** (Master Targets File → Monthly Targets):
  - They are **net**. Gross = net ÷ 0.83.
  - "Pro-rated to date" counts the current month's target only up to today.
- **Annual branch targets** (Branch Targets) are net.
- **Partner targets** (QTOD and ATOD file) are **gross**. Net = gross × 0.83.
- **New / revival partner:** the first sale is in the current year. The data starts two years back, so this also means no sales in the previous two years.
- **Row-level security:** the old *Branch RLS* user filter is rebuilt as two calculations, `Branch access (sales)` and `Branch access (partner targets)`. They use the same Tableau groups. They are applied to the partner-level pages (Partners, Partner Targets, Explore, Backlog partners), where the old workbook applied it.
