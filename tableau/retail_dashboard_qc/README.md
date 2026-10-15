# Overall Retail Dashboard — QC fixes

Applies the QC fixes to the existing workbook and keeps its layout, data sources and refresh. The full rebuild is in `../retail_dashboard`.


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
