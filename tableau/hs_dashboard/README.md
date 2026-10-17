# Home Solutions (HS) Dashboard (Tableau) — two workbooks

Built from the data source of `Home Solutions Dashboard.twbx`. It keeps the same files, relationships and extract, so refresh works as before. The layout engine is the Retail dashboard's (`../retail_dashboard`).

```
python build_hs.py "Home Solutions Dashboard.twbx" "HS Dashboard (editable rebate).twbx" editable
python build_hs.py "Home Solutions Dashboard.twbx" "HS Dashboard (static rebate).twbx" static
python ../retail_dashboard/validate.py "HS Dashboard (editable rebate).twb"
```

The `.twbx`, `.twb` and `.hyper` files contain business data and are git-ignored.

## The two workbooks

- **Editable rebate:** every page's filter panel has four rebate rates that can be typed in (0.13 = 13%): Retail this year, Project this year, Retail last year, Project last year. Net figures follow.
- **Static rebate:** the same rates are fixed in the calculations. They are set in `hs_model.REBATES`: Retail 13% / Project 5.1% this year, Retail 12.2% / Project 4.5% last year.

## Rules

- **Net:** Gross × (1 − rebate). This year's rate applies to this year's rows (and open backlog); last year's rate applies to earlier rows.
- **GTM:** Retail = pricing class EKR / ERW, or activity FD. Everything else is Project. The pricing class only defines GTM; there is no pricing-class analysis.
  - The panel's *Business (GTM)* switch shows All, Retail or Project.
  - Targets and commits use their own Retail / Project category.
- **Activity:** WD (HDCWD, HSCWD), FD (HDCFD, PPCFD), COS, SPE, EVC (PREVC).
- **With / without FD:** *Without FD* removes FD sales, OB and backlog, FD commits and FD yearly targets.
- **Monthly and yearly targets:** branch targets are net; gross = net / 0.83.
- **Commits:** activity-wise month commits are net; gross = net / 0.87.
  - 'Project' FD commit lines are left out. FD sales are always Retail, so they can never be achieved.
- **Partner targets:** gross. Family targets are used as they are, where set for the year.
- **Wiser / Zeta monthly targets:** the target sheet repeats them on the Retail and the Project row of each branch and month, so only one row is used.
- **Month keys:** the build adds the month to the relationships of the monthly-target and commit tables. Without it, a month-by-month view repeats each month's target in every month.

## Pages

Overview · Regions · Branches · Activity & Commit · Retail vs Project · Products · Partners · All Partners · Monthly Targets · MTD · Order Book · Explore.

Each page has the shared dropdown filters (Region, Branch, ASM, Partner, Sold To Party ID, Activity, Family, Partner type, Tier, Partner growth band), small tables sized to their rows, and a phone layout.
