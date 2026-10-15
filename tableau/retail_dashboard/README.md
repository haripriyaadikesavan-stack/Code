# Retail Dashboard (Tableau) — v4, built from scratch

Builds `Retail Dashboard v4.twbx` from the data source of `Overall Retail Dashboard.twbx`. It reuses the same SharePoint files, relationships and extract, so the refresh works as before. Every calculation, sheet and dashboard is new.

```
python build.py "Overall Retail Dashboard.twbx" "Retail Dashboard v4.twbx"
python validate.py "Retail Dashboard v4.twb"        # references, tables, aggregation, layout, actions
python order_check.py "Retail Dashboard v4.twb" "Overall Retail Dashboard.twb" "eCommerce_Analysis.twb"
python wire.py "Retail Dashboard v4.twb" wireframes  # optional PNG wireframes (needs matplotlib)
```

Needs `lxml`. The `.twbx`, `.twb` and `.hyper` files contain business data and are git-ignored.

## The story: 12 chapters

Every page has:
- a navigation bar with all 12 chapters;
- Back / Next buttons;
- a headline sentence, computed from the current filters, that states the main finding.

Each analysis appears once.

| # | Chapter | Content |
|---|---|---|
| 1 | Overview | KPIs, three-year monthly trend, quarters vs last year, sales vs monthly target, region share, RD / KRO split |
| 2 | Regions | Region scorecard, monthly sales by region, region x month achievement, branch sales, Wiser / Miluz Zeta by region. Click a region to filter the page. |
| 3 | Branches | Branch scorecard, Wiser / Miluz Zeta by branch, ASM performance. Click a branch to see its ASMs. |
| 4 | Products | Family value and quantity, activity (FD / WD ...), new products, classification, top 25 materials |
| 5 | Partners | Achievement bands, slabs, tiers, new / revival by month, onboarded year, top 10, partners to watch |
| 6 | All partners | Every partner's numbers: sales, last year, growth, FY target, Ach %, OB, backlog, days since last sale. Also every partner month by month, with Partner search. |
| 7 | Targets | Branch x month achievement and QTOD / ATOD partner targets (annual 100% / 108%, quarters, Wiser, Zeta) |
| 8 | SSP & reach | SSP targets, with vs without SSP, SSP trend and top SSP partners, active and new cities by branch |
| 9 | Pavilion | Pavilion partners vs target by branch, with primary purchases matched by GSTIN, plus the published slab and qualification tables |
| 10 | MTD | Month to date vs the same days last year and this month's target; branches, ASMs, families, top partners |
| 11 | Orders | Order book vs last year, sales-to-OB conversion, backlog ageing, backlog by branch and partner, every open order line |
| 12 | Explore | Any dimension against all measures |

## Overall, region-wise and branch-wise

The Region, Branch, Family, Activity and Partner type filters on the left apply to every chapter. Each chapter can therefore be read for the whole business, for one region or for one branch, without separate copies of the dashboards.

## Business rules

- **Sales, OB and Backlog:**
  - Sales = `Type = Sales`, dated by `InvoiceDate`.
  - OB = `Type = OB`, dated by `Calendar Day`.
  - Backlog = `Type = Backlog`, aged from `OrderCreationDate`, in the buckets < 20, 20-50, 51-100 and over 100 days.
- **Growth:** compares the same dates last year. The three-year trend also shows the same dates two years ago.
- **Monthly and yearly branch targets** are **net**. Gross = net ÷ 0.83.
- **Partner targets** are **gross**. Net = gross × 0.83.
  - The FY target comes from the Partner Targets sheet, as the old Partner Ach% did.
  - The quarter, Wiser and Zeta views come from the QTOD / ATOD file.
- **SSP targets** are used as they are.
- **New / revival partner:** first sale is this year. **Lapsed partner:** bought in the same period last year but not this year.
- **New city:** the first ever sale in that city is this year.
- **Pavilion partners:** active rows of the Pavilion Sales sheet, matched to primary sales by GSTIN. The target is used as it is in the file.
  - ROs buy through distributors, so their sales appear only in the EazyDMS secondary data. That data ends in December 2024 in the current extract.
- **Activity:** HDCFD = FD, so the old FD dashboard is chapter 1-11 with Activity = HDCFD.
- **Row-level security:** the old *Branch RLS* user filter is rebuilt as `Branch access (sales)` and `Branch access (partner targets)` on the partner-level sheets.
