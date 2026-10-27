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
