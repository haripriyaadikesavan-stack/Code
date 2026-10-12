# Channel Analytics

`Channel_Analytics_v8.ipynb` is the current notebook (`Channel_Analytics_v7_3.ipynb` is the previous version, kept for reference).

## Running it
1. Put `SP_TENANT_ID`, `SP_CLIENT_ID`, `SP_CLIENT_SECRET` in a `.env` file (never commit it). It is found automatically next to the notebook or in a parent folder, your home / Documents / OneDrive folder; otherwise set `ENV_FILE_PATH` in cell 0.3 (or the `CHANNEL_ENV_FILE` environment variable). If none is found, the notebook asks for the three values (hidden input, not saved).
2. Run cell **0.1 Setup** once per machine.
3. Run All. Cell **0.2** asks: `1` = use cached data with **no refresh** — Sales and OB from the SharePoint `parquet_cache` folders (Pre-processed Sales Data / Pre-processed OB Data), Backlog from its source folder (no cache); no SAP broadcast, no conversion, no uploads, no master-data updates; only reports are written. `0` = check SharePoint for new data (full refresh, updates master data and the parquet caches). Set `CHANNEL_USE_CACHE=0/1` to skip the prompt.
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
| 9 | Master-data review (last code cell): unmapped cities, partners missing from Partner Master, unmapped Sales Offices, products missing from Material Master / without Family Actual, and duplicate keys in Geography mapping / Material Master / Partner Master / NS Branches / SSA Mapping (with Excel row numbers) → `Master_Data_Review_<date>.xlsx` |

## Report style
Every workbook uses the house style in cell 0.5 (`THEME`): Calibri 10, dark-green headers, banded rows, frozen header, filters, bold totals, and the same number formats everywhere (₹ `#,##0`, ₹ Mn `#,##0.0`, `0.0%`, `dd-mmm-yyyy`). Data files read by Power BI / other reports keep their header on row 1.

## Duplicates in master data
A key listed twice in a mapping sheet (a city in Geography mapping, a material in Material Master, a Sold-to Party in Partner Master) would copy every matching sales row in a join — in the notebook and in Tableau. The notebook keeps the first row of each key when it loads master data and stops a report if a join ever adds rows; the review file lists every duplicate so it can be removed from the source file (which is what protects Tableau).

## Progress log
Every message shows the clock time and time since the run started (`10:42:07 +03:15 │ …`). Each report logs `[n/total] ▶ name (last run 38s)` when it starts and `✅ done in …` / `❌ failed …` when it ends, plus every sheet being written (rows × columns) and every file saved/uploaded. The same log is written to `logs/Channel_Analytics_<date>_<time>.log` next to the notebook; `logs/report_timings.json` keeps each report's last duration.
