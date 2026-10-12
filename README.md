# Channel Analytics

`Channel_Analytics_v8.ipynb` is the current notebook (`Channel_Analytics_v7_3.ipynb` is the previous version, kept for reference).

## Running it
1. Put `SP_TENANT_ID`, `SP_CLIENT_ID`, `SP_CLIENT_SECRET` in a `.env` file next to the notebook (never commit it).
2. Run cell **0.1 Setup** once per machine.
3. Run All. Cell **0.2** asks: `1` = use cached data (fast, nothing on SharePoint changes except the reports), `0` = check SharePoint for new data (full refresh, updates master data, saves a new cache). Set `CHANNEL_USE_CACHE=0/1` to skip the prompt.
4. The **Run summary** cell at the end lists each report as ok / skipped / FAILED.

Switch reports on/off in `REPORTS` (cell 0.3). All shared settings — paths, years, Retail filter, rebates, SSP map — live in that cell.

## Layout
| Section | Contents |
|---|---|
| 0 | Setup, run mode, configuration, SharePoint client, report style, shared helpers |
| 1 | Data: SAP broadcast → raw → pre-processed → `df_Sales_final` / `df_OB_final` / `df_backlog_final` (+ local cache) |
| 2 | Master data: Partner & Material Master updates, one load of all master files |
| 3 | Retail reports |
| 4 | HS (Retail + Project) reports |
| 5 | Partner Analysis, New GEO, Insights Summary |
| 6 | SSP performance |
| 7 | Advanced analytics & ad-hoc tools (off by default) |
| 8 | Archive — superseded code as raw cells (never executed) |

## Report style
Every workbook uses the house style in cell 0.5 (`THEME`): Calibri 10, dark-green headers, banded rows, frozen header, filters, bold totals, and the same number formats everywhere (₹ `#,##0`, ₹ Mn `#,##0.0`, `0.0%`, `dd-mmm-yyyy`). Data files read by Power BI / other reports keep their header on row 1.
