# Retail Dashboard (Tableau) — built from scratch

Builds `Retail Dashboard v11.twbx` from the data source of `Overall Retail Dashboard.twbx`. It reuses the same SharePoint files, relationships and extract, so the refresh works as before.

```
python build.py "Overall Retail Dashboard.twbx" "Retail Dashboard v11.twbx"
python validate.py "Retail Dashboard v11.twb"        # references, tables, aggregation, filters, actions
python order_check.py "Retail Dashboard v11.twb" "Overall Retail Dashboard.twb" "eCommerce_Analysis.twb"
python wire.py "Retail Dashboard v11.twb" wireframes  # optional PNG wireframes (needs matplotlib)
```

Needs `lxml`. The source can be the original workbook or a later Retail Dashboard whose extract has had tables removed: sheets that need a removed table are left out (and listed when building). The `.twbx`, `.twb` and `.hyper` files contain business data and are git-ignored.

## Layout

Every dashboard is built from layout containers, so rows, columns and the filter panel can be moved or resized in Tableau:
- header;
- navigation bar with Back / Next;
- filter panel on the left;
- one horizontal container per row of sheets.

Small tables (regions, categories, KPI lists, up to about 30 rows) get the height their rows need, so they do not scroll. Table columns are sized to fill their box; very wide tables keep 60 px columns and scroll sideways. Column headers are white on green, wrap, and are as tall as the longest header needs.

Where a 'Select a view' parameter swaps sheets (Partner Targets, MTD), each sheet's visibility is controlled by the parameter (Control visibility using value), so only the selected sheet and its title are shown.

## Filters (left panel, every page)

- **Parameters:** Period (YTD, QTD, MTD, last month, full year, custom), Gross / Net, Units.
- **Dropdown filters with search:** Region, Branch, ASM, Partner (name and code), Sold To Party ID, Family, Activity, Partner type, Tier, Policy category (e.g. only new partners), Partner growth band (growth or decline vs the same period last year, new, no sales).
  - They are shared by every sheet; the Pavilion page has its own Region / Branch / KRO-RO filters.
  - Branch names are used exactly as they are in the files.

## Pages

| Page | Content |
|---|---|
| Overview | KPIs (yearly target), three-year trend, quarters, region share, RD / KRO |
| Regions | Region scorecard with yearly targets and growth, monthly sales by region, growth heat map, branch sales, Wiser / Zeta growth |
| Branches | Branch scorecard vs yearly targets (balance, required run-rate), Wiser / Zeta, ASM table |
| Products | Families (value and quantity with growth), activity, new products (Wiser, Miluz Zeta, Unica Pure Plus, UnicaX), classification, top materials |
| Partners | Achievement bands, slabs, tiers, new / revival by month, onboarded year, top 10, declining partners |
| All partners | Every partner with tier, policy category, last FY sales, growth, assigned target (policy), Ach %, required monthly run-rate, OB, backlog; partner x month |
| New Partner - New Geo | Policy categories (Old, Handholding, New) and last-year sales buckets; new partners as the policy tracker groups them (New Partner group: new this year, Q4 last year onboarded, under 5 L last year onboarded), on this page only: by policy category with the flat 25 L assigned target, by region vs the count and value targets, region x category, tier, type, activity; new geo cities |
| Monthly Targets | Monthly Targets sheet: by month, chart, branch x month, Wiser / Zeta |
| Yearly Targets | Branch Targets sheet: by region and branch, achievement bars |
| Partner Targets | QTOD / ATOD partner-wise target vs achievement, as the old Target vs Ach dashboard: 'Select a View' picks Annual, QTOD Q1 / Q2 / Q3 / Q4 (quarter target, sales, Ach %, balance, each month's sales and Ach %), Wiser or Zeta |
| SSP | SSP target vs achievement (with totals), with vs without SSP, trend, new partners and new geo brought in by each SSP, and the new partners onboarded through SSPs |
| Pavilion | KPIs and published tables; KRO primary purchases vs target; RO pavilions buying in primary are listed for checking |
| MTD | The original MTD views: KPIs, focus products, region table, daily trend, leaderboard, and 'Select a view' (branch, activity, ASM, focus products, partner, product). 'Last billing day' is the latest invoice date in the data |
| Orders | Order book, backlog ageing, by branch and partner, every open order line |
| Explore | Any dimension against all measures |

## Rules

- **Tables:** growth columns are green for growth and red for decline. Achievement columns are red below 75%, amber from 75 to 100%, and green above 100%. Grand totals are bold.
- **Monthly targets:** from the Monthly Targets sheet as they are. They are used on Monthly Targets and MTD only.
- **Yearly targets:** branch-level, from the Branch Targets sheet. They are used in the scorecards.
- **Target basis:** monthly and yearly targets are net; gross = net ÷ 0.83. Partner targets are gross; net = gross × 0.83.
- **New / revival partner:** bought this year, nothing in the previous two years.
- **Policy categories (as the Channel Analytics notebook):**
  - New Partner (this year).
  - Q4 last year onboarded.
  - >5L last year onboarded (Handholding).
  - <5L last year onboarded.
  - Old Partner.

  New partners get a flat 25 L assigned target. Partner 500073231 is always Old.

  The extract holds sales from January 2024 only, so last year's onboarded partners cannot be checked against sales three years back.
- **New geo city:** sales this year, none in the previous two years.
- **Pavilion:** only KROs buy in primary, matched by GSTIN.
- **Backlog over 50 days:** open backlog (ordered, not yet invoiced) whose order was created more than 50 days ago.

## Phone layout

Every dashboard has a Phone layout: header, chapter buttons three to a row, the settings and the main filters (Region, Branch, ASM, Partner, Policy category, Growth band), then every sheet in page order, KPI cards two to a row. It is a separate device layout; the desktop (Default) layout is not changed by it.
