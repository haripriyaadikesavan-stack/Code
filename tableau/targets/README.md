# HS & Wiser targets workbook

`make_targets_template.py` writes the targets workbook the HS and Wiser dashboards are meant to read
(one file for both), with sample values or empty (`--empty`).

    python make_targets_template.py "HS_Wiser_Targets_2026.xlsx" --year 2026 [--empty]

| Sheet | One row per | Columns |
|---|---|---|
| Branches | target branch | Branch, Region, Retail data branch (partner map), Project data branch (NS Branches), Active |
| Branch Targets | year × month × branch × GTM × activity | Year, Month, Region, Branch, GTM, Activity, Target, Basis |
| Commits | same keys as Branch Targets | ..., Commit, Basis |
| Family Targets | year × month × branch × GTM × family | ..., Family, Target, Target quantity, Basis |
| Partner Targets | year × quarter × partner × family | Year, Quarter, Partner Code, Partner, Branch, Region, GTM, Partner type, Family, Target 100%, Target 108%, Basis |
| Wiser Service | year × month × branch | Installations target, Demos target, Expected TAT days |
| Lists | — | allowed values for the dropdowns |
| Check | — | totals per branch; unknown branches, duplicate keys, Project FD rows, missing basis (all should be 0) |

What it fixes compared with the current target sheets:

- Wiser / Zeta targets are entered once with their GTM, not repeated on a Retail and a Project row.
- Every amount carries its basis (Net / Gross) instead of each sheet having an unwritten convention.
- Yearly = sum of the months, so yearly and monthly targets cannot disagree.
- Quarter targets are numbers (Q2–Q4 are text in the QTOD sheet).
- One branch list maps the Retail and Project branch names (MUM / Mum, PUNE / Pune, KAR / Kar, RAJ / Raj).
- Long tables: Tableau reads them without pivots; adding a month, branch or family is adding rows.

The branch list and its Retail / Project names are taken from the 2026 HS data and need confirming.

## Organising the real target files (`organize_targets.py`)

Turns the current Retail target files (`Master Targets File.xlsx`, `QTOD and ATOD file.xlsx`) and the HS file
(`Master Targets File (Retail+Project).xlsx`) into `Channel_Targets_<year>.xlsx` (the layout above, Retail and
Project together), splitting every target by last year's sales, and writes `Sales_Contribution_<last year>.xlsx`.
It runs as section 10 of `Channel_Analytics_v9.ipynb`; `notebook_cell.py` generates that cell from this module.

Sheets: **Yearly Targets** (year × branch × GTM) and **Monthly Targets** (the monthly commit: month × branch × GTM)
at branch level, **Product Targets Monthly / Yearly** (only products that have a target), Partner Targets, SSP Targets,
Region Targets, Issues and Check. (This replaces the Branch / Family Targets and Commits sheets of the template above.)

Rules (column `Source` on every row says where a number comes from):

- **Branch targets as given.** Yearly ('Branch Targets') and monthly ('Monthly Targets', the monthly commit) are
  separate: the months need not add up to the yearly target, and months without a target yet stay empty.
- **Product targets only where one exists.** Monthly Miluz Zeta / Wiser from 'Monthly Targets'; 'Product Targets'
  (region level) shared out to the region's branches by each branch's sales of that product the year before.
- **Project:** yearly targets (HS Branch Targets) and monthly commits (HS Activity wise Month Target), by activity.
  The HS file's Retail rows are replaced by the Retail file, so both dashboards show the same Retail targets. Project
  FD commits and months that repeat the whole yearly target are left out (listed under Issues).
- **Partners:** QTOD quarters as given (108% = quarter × 1.08); Zeta / Wiser by the partner's quarters.
- Branch spellings are aligned with the sales data (PNQ → PUNE, TN2 → TN-2, NCR(DEL+ROH) → NCR ...); partner branches
  come from the partner master.

In Tableau relate, on the columns notebook cell 2.5 writes into the dashboard input files:

| Sheet | Fields |
|---|---|
| Yearly Targets | Target Branch = Branch, Target GTM = GTM, Target Year = Year |
| Monthly Targets | Target Branch = Branch, Target GTM = GTM, Target Month = Month |
| Product Targets Monthly | Target Branch, Target GTM, Target Month, Product Group = Product group |
| Product Targets Yearly | Target Branch, Target GTM, Target Year, Product Group = Product group |
| Partner Targets | Partner key = Partner Code, Target Year = Year, Target Quarter = Quarter |
| SSP Targets | Target Branch = Branch, Target Year = Year |
