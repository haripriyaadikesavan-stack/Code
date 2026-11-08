# Channel Analytics — dashboard reports in Excel

`Channel_Analytics_Dashboard_Reports.ipynb` (repo root) writes the analysis of the Retail and HS dashboards as six workbooks:

| File | Scope | Values |
|---|---|---|
| `Retail_Analysis_Gross_<Mon>_<year>.xlsx` / `…_Net_…` | Retail | Gross / Net |
| `HS_with_FD_Analysis_Gross_…` / `…_Net_…` | HS Retail + Project, FD included | Gross / Net |
| `HS_without_FD_Analysis_Gross_…` / `…_Net_…` | HS Retail + Project, FD left out | Gross / Net |

Settings (cell 1): `REBATES` (Retail / Project, by year; Net = Gross × (1 − rebate)), `AS_OF`, `SOURCE` (`sharepoint` or `local`), output folder.
Targets in the target files are net; the Gross files show them ÷ (1 − rebate). QTOD partner targets are gross and are netted for the Net files.
Without FD there is no Retail target (the Retail branch targets include FD).

Inputs: the two dashboard input files written by `Channel_Analytics_v9.ipynb`, Material Master, Data Dictionary,
Master Targets File, QTOD and ATOD file, Master Targets File (Retail+Project).

## Editing
Edit `dashboard_reports.py` (cells separated by `# %%`), then rebuild:

    python channel_reports/build_notebook.py

The builder fills in the SharePoint client and report style from `Channel_Analytics_v9.ipynb` and the target rules from
`tableau/targets/organize_targets.py`, so all three notebooks use the same code.
