# Wiser Dashboard (Tableau) — sales and services

Built from the data source of `Wiser Performance Dashboard.twbx`. It keeps the same files, relationships and extract. The layout engine is the Retail dashboard's (`../retail_dashboard`).

```
python build_wiser.py "Wiser Performance Dashboard.twbx" "Wiser Dashboard.twbx"
python ../retail_dashboard/validate.py "Wiser Dashboard.twb"
```

The `.twbx`, `.twb` and `.hyper` files contain business data and are git-ignored.

## Rules

- **Wiser sales:** primary rows of family Wiser. Gross / Net come from the source columns.
  - The *Business (GTM)* switch shows All, Retail (pricing class EKR / ERW) or Project.
- **Wiser version:** 1.0 / 2.0, from the material master classification.
- **Yearly Wiser targets:** used as set (branch × Retail / Project × year).
- **Monthly Wiser targets:** net; gross = net / 0.83. The sheet lists every branch and month twice, so half the sum is used.
- **Cancellations:** negative OB lines in the period.
- **Services:** work orders, counted once each (a work order has several product rows).
  - Installation = Commissioning & Installation or Assessment / Inspection; Repair = Repair or Decommissioning.
  - TAT = completed − created, in days, averaged per work order. *Done within 3 days* uses the old workbook's expected days.
  - The Services page has its own filters (service region, branch, work-order group, problem category).

## Pages

Overview · Regions & Branches · Products · Partners · All Partners · Monthly Targets · MTD · Order Book · Services · Explore. Each has a phone layout.
