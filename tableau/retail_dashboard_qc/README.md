# Overall Retail Dashboard — QC fixes, theme and analysis pages

Improves the existing workbook in place. Every existing dashboard is kept; the data sources and refresh are unchanged.

## v3 (full build)

```
cd ../retail_dashboard
RETAIL_MERGE=1 python build.py "Overall Retail Dashboard.twbx" gen_merge.twbx                  # analysis pages on the workbook's own parameters
cd ../retail_dashboard_qc
python build_v3.py "Overall Retail Dashboard.twbx" ../retail_dashboard/gen_merge.twb "Overall Retail Dashboard v3.twbx"
python check.py "Overall Retail Dashboard v3.twb"         # references, parameters, actions, windows
python check_merge.py "Overall Retail Dashboard v3.twb"   # the added pages against the data source
python wire.py "Overall Retail Dashboard v3.twb" wireframes   # optional: PNG wireframe of every dashboard (needs matplotlib)
```

`build_v3.py` runs, in order:
1. `fix.py`: the QC fixes below.
2. `theme.py`: one theme on every dashboard and sheet (colours, borders and fonts only; nothing moves).
   - Background `#f2f5f3`.
   - Charts and tables on white cards with a thin border.
   - Section headers: green band with white text.
   - Page titles: dark green, bold 18.
   - Filter panel: `#e9f0ec`.
   - Table headers and panes: soft greens instead of beige, lilac and olive.
   - Brown text and borders: dark green.
3. `layout.py`:
   - Rows of charts that stopped short of the edge are stretched to close the gap.
   - Left-panel controls and export buttons are stacked with even spacing.
4. `merge.py`: adds 7 analysis dashboards and a "More analysis" button on the Retail Overall Dashboard:
   - Executive Summary
   - Branch Scorecard
   - Product Insights
   - Partner Health
   - Target Tracker
   - Order Book & Backlog
   - Explore (Any Dimension)

   They use the workbook's own Gross / Net, Currency Selector, Start Date and End Date parameters, so one set of controls drives every dashboard.

## v2 (QC fixes only)

```
python fix.py "Overall Retail Dashboard.twbx" "Overall Retail Dashboard v2.twbx"   # writes fix_log.txt
python check.py "Overall Retail Dashboard v2.twb"                                  # references, parameters, actions, windows
python gn.py                                                                       # which Gross/Net parameter each sheet follows
```

Fixes:
- **Gross/Net:** the two Gross/Net parameters are merged into one.
- **Targets:** FY Target 108% and Curr. yr. Partner Target now convert with Gross/Net. Partner targets are gross; net = gross × 0.83.
- **Partner Ach%:** follows Gross/Net.
- **Q1 Ach% (Zeta):** now uses Q1 sales (it used Q2).
- **Monthly and quarterly Net:** uses the Net Sales column.
- **Clean-up:** unused sheets and calculated fields are removed.
- **Formatting:** sheet titles share one style, left-panel controls are aligned, and typos in names are corrected.
