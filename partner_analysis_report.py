"""
Partner Analysis Report Generator
==================================
Rebuilt from the original PTP_test3 script.

WHAT CHANGED vs. the original:
  1. All xlsxwriter formats are now defined ONCE in build_formats() and reused
     everywhere (the original defined two near-identical format blocks).
  2. write_summary_table() and write_complex_table() were duplicated logic
     (period aggregation + table writing). Replaced with a single
     `aggregate_periods()` + `write_kpi_table()` pair used by BOTH the
     Summary tab and the New_Partner_Summary tab.
  3. New: a "New Partner Dashboard" block at the TOP of New_Partner_Summary,
     matching your 3 screenshots:
        - KPI card: New Partners App./Sales split by RD vs KRO (Image 1)
        - BU Target context box (Image 1)
        - Region-wise New Partner Target vs Actual, 2026 YTD vs 2025 YTD (Image 3)
        - NEW: Helios-code x Region split, NEW PARTNERS ONLY (Image 2, filtered)
  4. generate_policy_handholding_tracker(): the row-wise .apply() categorization
     was replaced with a vectorised np.select (same logic, faster on large data).
  5. Sheet tabs colour-coded, freeze panes added for readability.
  6. Manual "always Old Partner" override for specific Sold-to Party IDs
     (see MANUAL_OLD_PARTNER_IDS) -- applied at the two places this script
     independently derives "new partner" status, so a partner on this list
     is never counted as New/Revival/New Partner/Handholding anywhere in
     the workbook (dashboard, aggregate tables, detail sheets,
     Lifetime_Sales/MoM_2026 Segment column, AND the Partner_Policy_Tracker
     sheet's own Policy Category / Old-New-Handholding columns).
  7. NEW_PARTNER_TARGETS_BY_ZONE: hardcoded New Partner Count/Sales targets
     by region (Image 3), since the Master Targets File's Branch Targets
     sheet doesn't carry this yet. Used directly in the Region-wise New
     Partner Target vs Actual table.
  8. Sheet1 (slimmed Region-actual-only view, no targets) and Geo Expansion
     (policy-tracker-driven "New Partner" population -- broader than
     Stat_2026 -- with its own Region/RD-KRO breakdowns and an FY2025
     sales-bucket split) sheets, matching the separate
     "New_GEO_{year}_YTD_{month}.xlsx"-style report. SSP sheet intentionally
     NOT included -- no Branch/Distributor(SSA)-level target+sales source
     is loaded anywhere in this script; add it once that source is known.
  9. Geo Expansion now also includes the Handholding-partner region
     breakdown table (Region / Partner Count / YTD'25 / YTD'26 / Grw%),
     which was present in your reference file but missing from this
     script, and the New Geo Expansion Analysis table has been moved to
     match the reference file's actual table order (Snapshot -> Policy
     rollup -> BU Target box -> Sales-bucket split -> Handholding region
     table -> New Partners region table -> New Geo Expansion Analysis ->
     RD/KRO breakdown -> Product breakup). The reference's 'Chls cont.'
     and '<5 lacs' columns on the Handholding table are still left out --
     no traceable definition for either yet.

REQUIRES (unchanged from your original, not modified here):
  - df_Sales_final must already exist in scope before this script runs
    (loaded from wherever your extraction step lives upstream).
  - msal, requests, pandas, numpy, xlsxwriter installed.

NOTE ON SECRETS: credentials are read from environment variables
(SP_TENANT_ID / SP_CLIENT_ID / SP_CLIENT_SECRET), never hardcoded.
"""

import io
import os
from datetime import datetime, timedelta
from urllib.parse import quote

import msal
import numpy as np
import pandas as pd
import requests

# =====================================================================
# 1. CONFIG & DATE ANCHORS
# =====================================================================
TARGET_YEAR = 2026
PREV_YEAR = 2025
LOCAL_OUTPUT_PATH = os.path.join(OUTPUT_DIR, "Partner_Analysis_2026_PTP_test3.xlsx")
GEO_OUTPUT_PATH = os.path.join(OUTPUT_DIR, "New_GEO_2026_YTD_Aug.xlsx")

today = datetime.now()
yesterday = today - timedelta(days=1)
last_year_yesterday = yesterday.replace(year=PREV_YEAR)
last_y_ytd_date = yesterday.replace(year=yesterday.year - 1)

MONTHS_ORDER = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]

MANUAL_OLD_PARTNER_IDS = {"500073231"}

TENANT_ID = os.environ["SP_TENANT_ID"]
CLIENT_ID = os.environ["SP_CLIENT_ID"]
CLIENT_SECRET = os.environ["SP_CLIENT_SECRET"]


# =====================================================================
# 2. AUTH & SHAREPOINT FETCH
# =====================================================================
def get_graph_headers():
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET,
    )
    token = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    return {"Authorization": f"Bearer {token['access_token']}"}


def get_drive_id(headers, site_path="schneiderelectric.sharepoint.com:/sites/HDIndiaTeamGroup"):
    site_res = requests.get(f"https://graph.microsoft.com/v1.0/sites/{site_path}", headers=headers).json()
    drives = requests.get(f"https://graph.microsoft.com/v1.0/sites/{site_res['id']}/drives", headers=headers).json()["value"]
    return next(d for d in drives if d["name"] == "Documents")["id"]


def fetch_sp_file(drive_id, headers, file_path):
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{quote(file_path)}"
    res = requests.get(url, headers=headers).json()
    return io.BytesIO(requests.get(res["@microsoft.graph.downloadUrl"]).content)


print("Authenticating with Graph API...")
headers_auth = get_graph_headers()
drive_id = get_drive_id(headers_auth)

# =====================================================================
# 3. MASTER DATA
# =====================================================================
print("Fetching Master Data...")
BASE_MASTER_PATH = "Test - haripriya/Data Downloads - Input Files/Master data"

Partner_Raw = pd.read_excel(fetch_sp_file(drive_id, headers_auth, f"{BASE_MASTER_PATH}/Data Dictionary.xlsx"), sheet_name="Partner Master Data")
Geo_Map = pd.read_excel(fetch_sp_file(drive_id, headers_auth, f"{BASE_MASTER_PATH}/Data Dictionary.xlsx"), sheet_name="Geography mapping")
df_targets = pd.read_excel(fetch_sp_file(drive_id, headers_auth, f"{BASE_MASTER_PATH}/Master Targets File.xlsx"), sheet_name="Partner Targets")
df_region_target = pd.read_excel(fetch_sp_file(drive_id, headers_auth, f"{BASE_MASTER_PATH}/Master Targets File.xlsx"), sheet_name="Branch Targets")
df_material_master = pd.read_excel(fetch_sp_file(drive_id, headers_auth, f"{BASE_MASTER_PATH}/Material Master Data.xlsx"), sheet_name="Sheet1")

Partner_Raw.columns = Partner_Raw.columns.str.strip()
Geo_Map.columns = Geo_Map.columns.str.strip()
df_region_target.columns = df_region_target.columns.str.strip()

Partner_Master = Partner_Raw.merge(Geo_Map.drop_duplicates("City"), on="City", how="left")
Partner_Master["Sold-to Party"] = Partner_Master["Sold-to Party"].astype(str).str.strip()

tier_col = next((c for c in Partner_Master.columns if "tier" in c.lower()), "Tier Category")
tier_map = {"Metro": "Metro & T1", "T1": "Metro & T1", "T2&T3": "T2 & T3"}
Partner_Master["Tier_Bucket"] = Partner_Master[tier_col].map(tier_map).fillna("Other")
kro_col = next((c for c in Partner_Master.columns if "partner type" in c.lower() or "kro" in c.lower()), "Partner Type")

NEW_PARTNER_TARGETS_BY_ZONE = pd.DataFrame({
    "Zone": ["East", "North", "South", "West1", "West2"],
    "Partner Count Target": [30, 50, 50, 20, 30],
    "Gr.Sales Target": [50, 130, 130, 60, 60],
}).set_index("Zone")

# =====================================================================
# 4. PROCESSING & SEGMENTATION
# =====================================================================
df_Sales_final["InvoiceDate"] = pd.to_datetime(df_Sales_final["InvoiceDate"])

mask = (
    df_Sales_final["PricingClassL2"].isin(["ED1", "ED2", "RE3"])
    & df_Sales_final["PricingClassL3"].isin(["ERW", "EKR", "R36", "ZD1"])
    & df_Sales_final["HeliosCode"].isin(
        ["HDCFD", "HDCWD", "HDCOS", "HDSPE", "HDCLD", "HDCLN", "HDCLS",
         "HDFDA", "HDFDB", "HDFDS", "HDLSD", "HDPNL", "HSCOS", "HSCWD", "HSSPE", "PPCFD"]
    )
)

df_retail = df_Sales_final[mask].copy()
df_retail["Year"] = df_retail["InvoiceDate"].dt.year
df_retail["Month"] = pd.Categorical(df_retail["InvoiceDate"].dt.month_name(), categories=MONTHS_ORDER, ordered=True)
df_retail["SoldToPartyID"] = df_retail["SoldToPartyID"].astype(str).str.strip()

sales_matrix = df_retail.groupby(["SoldToPartyID", "Year"])["Gross Sales"].sum().unstack(fill_value=0)
stat_26 = np.where((sales_matrix.get(2026, 0) > 0) & (sales_matrix.get(2025, 0) == 0) & (sales_matrix.get(2024, 0) == 0), "New/Revival", "Old")
stat_25 = np.where((sales_matrix.get(2025, 0) > 0) & (sales_matrix.get(2024, 0) == 0) & (sales_matrix.get(2023, 0) == 0), "New/Revival", "Old")
partner_status = pd.DataFrame({"Stat_2026": stat_26, "Stat_2025": stat_25}, index=sales_matrix.index)

override_idx = partner_status.index.intersection(MANUAL_OLD_PARTNER_IDS)
partner_status.loc[override_idx, ["Stat_2026", "Stat_2025"]] = "Old"

df_final = df_retail.merge(Partner_Master, left_on="SoldToPartyID", right_on="Sold-to Party", how="left")
df_final = df_final.merge(partner_status, left_on="SoldToPartyID", right_index=True, how="left")

geo_history = df_final.groupby(["City", "Year"])["Gross Sales"].sum().unstack(fill_value=0)
is_new_geo = (geo_history.get(2026, 0) > 0) & (geo_history.get(2025, 0) == 0) & (geo_history.get(2024, 0) == 0)
new_geo_list = geo_history[is_new_geo].index.tolist()
is_new_geo_25 = (geo_history.get(2025, 0) > 0) & (geo_history.get(2024, 0) == 0) & (geo_history.get(2023, 0) == 0)
new_geo_list_25 = geo_history[is_new_geo_25].index.tolist()

df_active = df_final[(df_final["Gross Sales"] != 0) & (df_final["Gross Sales"].notna())]


# =====================================================================
# 5. FORMATS (single source of truth)
# =====================================================================
def build_formats(workbook):
    return {
        "title": workbook.add_format({"bold": True, "font_size": 12}),
        "subtitle": workbook.add_format({"bold": True, "font_size": 11}),
        "hdr": workbook.add_format({"bold": True, "bg_color": "#50D08C", "border": 1, "align": "center", "valign": "vcenter"}),
        "hdr_wrap": workbook.add_format({"bold": True, "bg_color": "#50D08C", "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True}),
        "hdr_alt": workbook.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1, "align": "center"}),
        "hdr_pol": workbook.add_format({"bold": True, "bg_color": "#CFE2F3", "border": 1, "align": "center", "valign": "vcenter"}),
        "dec": workbook.add_format({"num_format": "#,##0.0", "border": 1, "align": "center"}),
        "int": workbook.add_format({"num_format": "#,##0", "border": 1, "align": "center"}),
        "pct": workbook.add_format({"num_format": "0.0%", "border": 1, "align": "center"}),
        "tot_dec": workbook.add_format({"bold": True, "bg_color": "#DAEFDF", "border": 1, "num_format": "#,##0.0", "align": "center"}),
        "tot_int": workbook.add_format({"bold": True, "bg_color": "#DAEFDF", "border": 1, "num_format": "#,##0", "align": "center"}),
        "tot_pct": workbook.add_format({"bold": True, "bg_color": "#DAEFDF", "border": 1, "num_format": "0.0%", "align": "center"}),
        "num_det": workbook.add_format({"num_format": "#,##0", "border": 1, "align": "right"}),
        "tot_det": workbook.add_format({"bold": True, "bg_color": "#DAEFDF", "num_format": "#,##0", "border": 1, "align": "right"}),
        "txt_det": workbook.add_format({"border": 1, "align": "left"}),
        "kpi_label": workbook.add_format({"bold": True, "bg_color": "#EAF6EF", "border": 1, "align": "left", "font_size": 11}),
        "kpi_value": workbook.add_format({"bold": True, "border": 1, "align": "center", "font_size": 11, "num_format": "#,##0.0"}),
        "kpi_value_int": workbook.add_format({"bold": True, "border": 1, "align": "center", "font_size": 11, "num_format": "#,##0"}),
        "kpi_pct_hero": workbook.add_format({"bold": True, "bg_color": "#50D08C", "border": 1, "align": "center", "font_size": 13, "num_format": "0.0%"}),
        "policy_num": workbook.add_format({"num_format": "#,##0", "border": 1, "align": "right"}),
        "policy_runrate": workbook.add_format({"num_format": "#,##0", "border": 1, "align": "right", "bg_color": "#FFF2CC", "bold": True, "font_color": "#9C0006"}),
    }


# =====================================================================
# 6. GENERIC PERIOD AGGREGATION (used by Summary tab + New Partner tab)
# =====================================================================
def aggregate_periods(group_col, is_segment=False, new_only=False, with_avg=False):
    s25_full = df_active[df_active["Year"] == 2025]
    s25_ytd = df_active[(df_active["Year"] == 2025) & (df_active["InvoiceDate"] <= last_year_yesterday)]
    s26_ytd = df_active[(df_active["Year"] == 2026) & (df_active["InvoiceDate"] <= yesterday)]

    if is_segment:
        group_25, group_26 = "Stat_2025", "Stat_2026"
    else:
        group_25, group_26 = group_col, group_col
        if new_only:
            s25_full = s25_full[s25_full["Stat_2025"] == "New/Revival"]
            s25_ytd = s25_ytd[s25_ytd["Stat_2025"] == "New/Revival"]
            s26_ytd = s26_ytd[s26_ytd["Stat_2026"] == "New/Revival"]

    def agg(df, gcol):
        return df.groupby(gcol).agg(P=("SoldToPartyID", "nunique"), S=("Gross Sales", lambda x: x.sum() / 1e6))

    a25f, a25y, a26y = agg(s25_full, group_25), agg(s25_ytd, group_25), agg(s26_ytd, group_26)
    a25f.columns = ["P25F", "S25F"]
    a25y.columns = ["P25Y", "S25Y"]
    a26y.columns = ["P26Y", "S26Y"]

    res = pd.concat([a25f, a25y, a26y], axis=1).fillna(0)
    res.columns = ["P25F", "S25F", "P25Y", "S25Y", "P26Y", "S26Y"]
    return res


def write_kpi_table(ws, wb, f, start_row, res, title, label, with_avg=True):
    ws.write(start_row, 0, title, f["title"] if with_avg else f["subtitle"])
    hdr_row = start_row + (2 if with_avg else 1)
    sub_row = hdr_row + 1

    if with_avg:
        ws.merge_range(hdr_row, 1, hdr_row, 3, "FY 2025 (Full Year)", f["hdr"])
        ws.merge_range(hdr_row, 4, hdr_row, 6, f"2025 YTD ({last_year_yesterday.strftime('%d-%b')})", f["hdr"])
        ws.merge_range(hdr_row, 7, hdr_row, 9, f"2026 YTD ({yesterday.strftime('%d-%b')})", f["hdr"])
        ws.merge_range(hdr_row, 10, hdr_row, 11, "Performance", f["hdr"])
        sub_headers = [label, "Partners", "Sales Mnr", "Avg Sales",
                       "Partners", "Sales Mnr", "Avg Sales",
                       "Partners", "Sales Mnr", "Avg Sales",
                       "Contr %", "Growth %"]
    else:
        ws.merge_range(hdr_row, 1, hdr_row, 2, "FY 2025", f["hdr"])
        ws.merge_range(hdr_row, 3, hdr_row, 4, f"2025 YTD ({last_year_yesterday.strftime('%d-%b')})", f["hdr"])
        ws.merge_range(hdr_row, 5, hdr_row, 7, f"2026 YTD ({yesterday.strftime('%d-%b')})", f["hdr"])
        ws.merge_range(hdr_row, 8, hdr_row, 9, "Performance", f["hdr"])
        sub_headers = [label, "Partners", "Sales", "Partners", "Sales",
                       "Partners", "Sales", "Avg Sales", "Contr %", "Growth %"]

    for c, h in enumerate(sub_headers):
        ws.write(sub_row, c, h, f["hdr"])

    total_s26 = res["S26Y"].sum() if with_avg else res["S26"].sum() if "S26" in res.columns else res["S26Y"].sum()
    r = sub_row + 1
    for idx, row in res.iterrows():
        s25y = row["S25Y"]
        s26y = row["S26Y"]
        growth = (s26y - s25y) / s25y if s25y > 0 else 0
        ws.write(r, 0, idx, f["txt_det"])
        if with_avg:
            ws.write(r, 1, row["P25F"], f["int"]); ws.write(r, 2, row["S25F"], f["dec"])
            ws.write(r, 3, row["S25F"] / row["P25F"] if row["P25F"] > 0 else 0, f["dec"])
            ws.write(r, 4, row["P25Y"], f["int"]); ws.write(r, 5, row["S25Y"], f["dec"])
            ws.write(r, 6, row["S25Y"] / row["P25Y"] if row["P25Y"] > 0 else 0, f["dec"])
            ws.write(r, 7, row["P26Y"], f["int"]); ws.write(r, 8, row["S26Y"], f["dec"])
            ws.write(r, 9, s26y / row["P26Y"] if row["P26Y"] > 0 else 0, f["dec"])
            ws.write(r, 10, s26y / total_s26 if total_s26 > 0 else 0, f["pct"])
            ws.write(r, 11, growth, f["pct"])
        else:
            ws.write(r, 1, row["P25F"], f["int"]); ws.write(r, 2, row["S25F"], f["dec"])
            ws.write(r, 3, row["P25Y"], f["int"]); ws.write(r, 4, row["S25Y"], f["dec"])
            ws.write(r, 5, row["P26Y"], f["int"]); ws.write(r, 6, row["S26Y"], f["dec"])
            ws.write(r, 7, s26y / row["P26Y"] if row["P26Y"] > 0 else 0, f["dec"])
            ws.write(r, 8, s26y / total_s26 if total_s26 > 0 else 0, f["pct"])
            ws.write(r, 9, growth, f["pct"])
        r += 1

    sums = res.sum()
    ws.write(r, 0, "Grand Total", f["tot_int"])
    if with_avg:
        ws.write(r, 1, sums["P25F"], f["tot_int"]); ws.write(r, 2, sums["S25F"], f["tot_dec"])
        ws.write(r, 3, sums["S25F"] / sums["P25F"] if sums["P25F"] > 0 else 0, f["tot_dec"])
        ws.write(r, 4, sums["P25Y"], f["tot_int"]); ws.write(r, 5, sums["S25Y"], f["tot_dec"])
        ws.write(r, 6, sums["S25Y"] / sums["P25Y"] if sums["P25Y"] > 0 else 0, f["tot_dec"])
        ws.write(r, 7, sums["P26Y"], f["tot_int"]); ws.write(r, 8, sums["S26Y"], f["tot_dec"])
        ws.write(r, 9, sums["S26Y"] / sums["P26Y"] if sums["P26Y"] > 0 else 0, f["tot_dec"])
        ws.write(r, 10, 1.0, f["tot_pct"])
        ws.write(r, 11, (sums["S26Y"] - sums["S25Y"]) / sums["S25Y"] if sums["S25Y"] > 0 else 0, f["tot_pct"])
    else:
        ws.write(r, 1, sums["P25F"], f["tot_int"]); ws.write(r, 2, sums["S25F"], f["tot_dec"])
        ws.write(r, 3, sums["P25Y"], f["tot_int"]); ws.write(r, 4, sums["S25Y"], f["tot_dec"])
        ws.write(r, 5, sums["P26Y"], f["tot_int"]); ws.write(r, 6, sums["S26Y"], f["tot_dec"])
        ws.write(r, 7, sums["S26Y"] / sums["P26Y"] if sums["P26Y"] > 0 else 0, f["tot_dec"])
        ws.write(r, 8, 1.0, f["tot_pct"])
        ws.write(r, 9, (sums["S26Y"] - sums["S25Y"]) / sums["S25Y"] if sums["S25Y"] > 0 else 0, f["tot_pct"])

    return r + 3


# =====================================================================
# 7. NEW PARTNER DASHBOARD  (top of New_Partner_Summary sheet)
# =====================================================================
def write_new_partner_dashboard(ws, wb, f, start_row=0):
    df_new_26_ytd = df_active[
        (df_active["Year"] == 2026) & (df_active["InvoiceDate"] <= yesterday) & (df_active["Stat_2026"] == "New/Revival")
    ]
    total_new_partners = df_new_26_ytd["SoldToPartyID"].nunique()
    total_new_sales = df_new_26_ytd["Gross Sales"].sum() / 1e6

    kro_split = df_new_26_ytd.groupby(kro_col).agg(
        P=("SoldToPartyID", "nunique"), S=("Gross Sales", lambda x: x.sum() / 1e6)
    )

    r = start_row
    ws.merge_range(r, 0, r, 3, f"New Partner Snapshot — YTD {yesterday.strftime('%d-%b-%Y')}", f["title"])
    r += 1
    ws.write(r, 0, "", f["hdr"])
    ws.write(r, 1, "New Partner's App.", f["hdr"])
    ws.write(r, 2, "Sales (Mn)", f["hdr"])
    r += 1
    ws.write(r, 0, "New Partners YTD", f["kpi_label"])
    ws.write(r, 1, total_new_partners, f["kpi_value_int"])
    ws.write(r, 2, total_new_sales, f["kpi_value"])
    r += 1
    for label in ["RD", "KRO"]:
        row = kro_split.loc[label] if label in kro_split.index else pd.Series({"P": 0, "S": 0})
        ws.write(r, 0, label, f["txt_det"])
        ws.write(r, 1, int(row["P"]), f["int"])
        ws.write(r, 2, row["S"], f["dec"])
        r += 1
    r += 1

    bu_target_mn = df_region_target.loc[df_region_target["Year"] == TARGET_YEAR, "Target"].sum() / 1e6
    total_ytd_sales_mn = df_active[
        (df_active["Year"] == 2026) & (df_active["InvoiceDate"] <= yesterday)
    ]["Gross Sales"].sum() / 1e6
    new_pct = total_new_sales / total_ytd_sales_mn if total_ytd_sales_mn else 0

    ws.write(r, 0, "Gross Total BU Target", f["kpi_label"]); ws.write(r, 1, bu_target_mn, f["kpi_value"]); r += 1
    ws.write(r, 0, f"Total YTD {yesterday.strftime('%b')} Sales", f["kpi_label"]); ws.write(r, 1, total_ytd_sales_mn, f["kpi_value"]); r += 1
    ws.write(r, 0, "New Partner's 2026 Sales", f["kpi_label"]); ws.write(r, 1, total_new_sales, f["kpi_value"])
    ws.merge_range(r, 2, r, 3, "New Partner's Sales %", f["hdr"]); r += 1
    ws.merge_range(r, 2, r, 3, new_pct, f["kpi_pct_hero"])
    r += 3

    ws.write(r, 0, "New Partner Performance by Region — Target vs Actual", f["title"])
    r += 1
    hdr_r = r
    ws.write(hdr_r, 0, "Region", f["hdr"])
    ws.write(hdr_r, 1, "Partner Count Target", f["hdr"])
    ws.write(hdr_r, 2, "Gr.Sales Target", f["hdr"])
    ws.merge_range(hdr_r, 3, hdr_r, 5, f"2026 YTD ({yesterday.strftime('%d-%b')})", f["hdr"])
    ws.merge_range(hdr_r, 6, hdr_r, 8, f"2025 YTD ({last_year_yesterday.strftime('%d-%b')})", f["hdr"])
    ws.merge_range(hdr_r, 9, hdr_r, 10, "Performance", f["hdr"])
    r += 1
    sub_r = r
    sub_headers = ["", "", "", "Partners", "Sales Mnr", "Avg Sales",
                   "Partners", "Sales Mnr", "Avg Sales", "% of Ach.", "Gwt %"]
    for c, h in enumerate(sub_headers):
        if h:
            ws.write(sub_r, c, h, f["hdr"])
    r += 1

    region_new = aggregate_periods("Zone", new_only=True)
    partner_target_by_zone = NEW_PARTNER_TARGETS_BY_ZONE["Partner Count Target"]
    sales_target_by_zone = NEW_PARTNER_TARGETS_BY_ZONE["Gr.Sales Target"]

    all_zones = sorted(set(partner_target_by_zone.index) | set(region_new.index))

    for zone in all_zones:
        row = region_new.loc[zone] if zone in region_new.index else pd.Series({"P25Y": 0, "S25Y": 0, "P26Y": 0, "S26Y": 0})
        p_target = int(partner_target_by_zone.get(zone, 0))
        s_target = sales_target_by_zone.get(zone, 0)
        s26 = row["S26Y"]
        s25 = row["S25Y"]
        pct_ach = s26 / s_target if s_target else 0
        gwt = (s26 - s25) / s25 if s25 else 0
        ws.write(r, 0, zone, f["txt_det"])
        ws.write(r, 1, p_target, f["int"])
        ws.write(r, 2, s_target, f["dec"])
        ws.write(r, 3, int(row["P26Y"]), f["int"])
        ws.write(r, 4, s26, f["dec"])
        ws.write(r, 5, s26 / row["P26Y"] if row["P26Y"] else 0, f["dec"])
        ws.write(r, 6, int(row["P25Y"]), f["int"])
        ws.write(r, 7, s25, f["dec"])
        ws.write(r, 8, s25 / row["P25Y"] if row["P25Y"] else 0, f["dec"])
        ws.write(r, 9, pct_ach, f["pct"])
        ws.write(r, 10, gwt, f["pct"])
        r += 1

    tot_p_target = int(partner_target_by_zone.sum())
    tot_s_target = sales_target_by_zone.sum()
    tot_s26 = region_new["S26Y"].sum()
    tot_s25 = region_new["S25Y"].sum()
    ws.write(r, 0, "Grand Total", f["tot_int"])
    ws.write(r, 1, tot_p_target, f["tot_int"])
    ws.write(r, 2, tot_s_target, f["tot_dec"])
    ws.write(r, 3, int(region_new["P26Y"].sum()), f["tot_int"])
    ws.write(r, 4, tot_s26, f["tot_dec"])
    ws.write(r, 5, tot_s26 / region_new["P26Y"].sum() if region_new["P26Y"].sum() else 0, f["tot_dec"])
    ws.write(r, 6, int(region_new["P25Y"].sum()), f["tot_int"])
    ws.write(r, 7, tot_s25, f["tot_dec"])
    ws.write(r, 8, tot_s25 / region_new["P25Y"].sum() if region_new["P25Y"].sum() else 0, f["tot_dec"])
    ws.write(r, 9, tot_s26 / tot_s_target if tot_s_target else 0, f["tot_pct"])
    ws.write(r, 10, (tot_s26 - tot_s25) / tot_s25 if tot_s25 else 0, f["tot_pct"])
    r += 4

    r = write_helios_region_pivot(ws, wb, f, r, new_partners_only=True,
                                   title="Product-wise Break-up — NEW PARTNERS ONLY")
    r += 2

    r = write_new_geo_table(ws, wb, f, r)

    return r + 2


def write_new_geo_table(ws, wb, f, start_row):
    ws.write(start_row, 0, "New Geo Expansion Analysis (Active 2026, Nil in 24/25)", f["subtitle"])
    hdr_r = start_row + 1
    headers_geo = ["City", "Partners (2026)", f"Sales Mnr ({yesterday.strftime('%d-%b')})"]
    for c, h in enumerate(headers_geo):
        ws.write(hdr_r, c, h, f["hdr"])

    geo_data = df_final[
        (df_final["City"].isin(new_geo_list)) & (df_final["Year"] == TARGET_YEAR) & (df_final["InvoiceDate"] <= yesterday)
    ]
    geo_sum = geo_data.groupby("City").agg(P=("SoldToPartyID", "nunique"), S=("Gross Sales", lambda x: x.sum() / 1e6))

    r = hdr_r + 1
    for idx, row in geo_sum.iterrows():
        ws.write(r, 0, idx, f["txt_det"])
        ws.write(r, 1, int(row["P"]), f["int"])
        ws.write(r, 2, row["S"], f["dec"])
        r += 1

    ws.write(r, 0, "Grand Total", f["tot_int"])
    ws.write(r, 1, int(geo_sum["P"].sum()), f["tot_int"])
    ws.write(r, 2, geo_sum["S"].sum(), f["tot_dec"])
    return r + 1


def write_helios_region_pivot(ws, wb, f, start_row, new_partners_only=False, title=None, new_partner_ids=None):
    df_src = df_retail.merge(
        df_material_master[["Material", "Family Actual"]],
        left_on="MaterialCode", right_on="Material", how="left",
    )
    zone_map = df_final[["SoldToPartyID", "Zone", "Stat_2026"]].drop_duplicates("SoldToPartyID")
    df_src = df_src.merge(zone_map, on="SoldToPartyID", how="left")

    df_src["Helios_Display"] = np.where(
        (df_src["HeliosCode"] == "HSCOS") & (df_src["Family Actual"].astype(str).str.upper() == "WISER"),
        "Wiser",
        np.where(df_src["HeliosCode"] == "HSCOS", "HSCOS Others", df_src["HeliosCode"]),
    )

    df_yr = df_src[df_src["Year"] == TARGET_YEAR]
    if new_partners_only:
        if new_partner_ids is not None:
            df_yr = df_yr[df_yr["SoldToPartyID"].astype(str).isin(new_partner_ids)]
        else:
            df_yr = df_yr[df_yr["Stat_2026"] == "New/Revival"]

    if df_yr.empty:
        ws.write(start_row, 0, (title or "Product-wise Break-up") + " — no data", f["title"])
        return start_row + 2

    pivot = df_yr.pivot_table(index="Zone", columns="Helios_Display", values="Gross Sales", aggfunc="sum", fill_value=0) / 1e6
    pivot["Grand Total"] = pivot.sum(axis=1)
    pivot.loc["Grand Total"] = pivot.sum()

    ws.write(start_row, 0, title or "Product-wise Break-up", f["title"])
    hdr_r = start_row + 1
    ws.write(hdr_r, 0, "Region", f["hdr"])
    for c, col in enumerate(pivot.columns, start=1):
        ws.write(hdr_r, c, col, f["hdr"])

    r = hdr_r + 1
    for idx, row in pivot.iterrows():
        row_fmt = f["tot_int"] if idx == "Grand Total" else f["txt_det"]
        val_fmt = f["tot_dec"] if idx == "Grand Total" else f["dec"]
        ws.write(r, 0, idx, row_fmt)
        for c, val in enumerate(row, start=1):
            ws.write(r, c, val, val_fmt)
        r += 1

    return r


# =====================================================================
# 8. DETAIL SHEET HELPER (pivot by partner, unchanged logic)
# =====================================================================
def write_detail(writer, wb, f, df_filtered, sheet_name, target_year):
    df_yr = df_filtered[df_filtered["Year"] == target_year]
    if df_yr.empty:
        print(f"No data for {sheet_name} in {target_year}, sheet skipped.")
        return

    pivot = df_yr.pivot_table(
        index=["SoldToPartyID", "Patner", "Zone", "Branch", "City", "Tier_Bucket", "RD/KRO", "ASM"],
        columns="Month", values="Gross Sales", aggfunc="sum", fill_value=0,
    )
    pivot["Grand Total"] = pivot.sum(axis=1)
    pivot = pivot.sort_values("Grand Total", ascending=False).reset_index()
    pivot.to_excel(writer, sheet_name=sheet_name, index=False)

    ws = writer.sheets[sheet_name]
    for c, col in enumerate(pivot.columns):
        ws.write(0, c, col, f["hdr"])
        ws.set_column(c, c, 18, f["num_det"] if c > 7 else f["txt_det"])
    ws.freeze_panes(1, 2)


# =====================================================================
# 9. POLICY / HANDHOLDING TRACKER (vectorised categorisation)
# =====================================================================
def generate_policy_handholding_tracker(writer, wb, f):
    df_retail["SoldToPartyID"] = df_retail["SoldToPartyID"].astype(str).str.strip()
    df_targets["Sold-to Party"] = df_targets["Sold-to Party"].astype(str).str.strip()

    target_col = [c for c in df_targets.columns if "2026" in str(c)][0]
    custom_targets = df_targets.set_index("Sold-to Party")[target_col].to_dict()
    rem_months = 12 - (yesterday.month - 1)

    sm = df_retail.groupby(["SoldToPartyID", "Year"])["Gross Sales"].sum().unstack(fill_value=0)
    sm.index = sm.index.astype(str)
    s25_val = sm[2025] if 2025 in sm.columns else pd.Series(0, index=sm.index)
    s24_val = sm[2024] if 2024 in sm.columns else pd.Series(0, index=sm.index)
    s23_val = sm[2023] if 2023 in sm.columns else pd.Series(0, index=sm.index)

    first_sale_2025 = (
        df_retail.loc[df_retail["Year"] == 2025]
        .groupby("SoldToPartyID")["InvoiceDate"]
        .min()
    )
    first_sale_2025.index = first_sale_2025.index.astype(str)

    tracker = df_final.groupby("SoldToPartyID").agg(
        Patner=("Patner", "first"), Zone=("Zone", "first"), Branch=("Branch", "first"),
        City=("City", "first"), ASM=("ASM", "first"), Stat_2026=("Stat_2026", "first"),
    )
    tracker.index = tracker.index.astype(str)

    s25 = s25_val.reindex(tracker.index).fillna(0)
    s24 = s24_val.reindex(tracker.index).fillna(0)
    s23 = s23_val.reindex(tracker.index).fillna(0)
    first_sale = first_sale_2025.reindex(tracker.index)

    stat_26_str = tracker["Stat_2026"].astype(str).str.lower().str.strip()
    is_new_2026 = stat_26_str.str.contains("new") | stat_26_str.str.contains("revival")
    is_new_25 = (s25 > 0) & (s24 == 0) & (s23 == 0)
    is_q4_2025 = is_new_25 & first_sale.dt.month.isin([10, 11, 12])
    is_5l_plus = s25 >= 500000

    tracker["Policy Category"] = np.select(
        [
            is_new_2026,
            is_new_25 & is_q4_2025,
            is_new_25 & ~is_q4_2025 & is_5l_plus,
            is_new_25 & ~is_q4_2025 & ~is_5l_plus,
        ],
        [
            "New Partner 2026",
            "Q4'2025 Onboarded",
            ">5L 2025 Onboarded",
            "<5L 2025 Onboarded",
        ],
        default="Old Partner",
    )

    is_new_partner_bucket = is_new_2026 | is_q4_2025 | (is_new_25 & ~is_q4_2025 & ~is_5l_plus)
    is_handholding_bucket = is_new_25 & ~is_q4_2025 & is_5l_plus

    tracker["Old/New/Handholding"] = np.select(
        [is_new_partner_bucket, is_handholding_bucket],
        ["New Partner", "Handholding partner"],
        default="Old partner",
    )

    override_mask = tracker.index.isin(MANUAL_OLD_PARTNER_IDS)
    tracker.loc[override_mask, "Policy Category"] = "Old Partner"
    tracker.loc[override_mask, "Old/New/Handholding"] = "Old partner"

    perf_26 = (
        df_retail[(df_retail["Year"] == 2026) & (df_retail["InvoiceDate"] <= yesterday)]
        .groupby("SoldToPartyID")
        .agg(YTD_2026=("Gross Sales", "sum"))
    )
    perf_25_ytd = (
        df_retail[(df_retail["Year"] == 2025) & (df_retail["InvoiceDate"] <= last_year_yesterday)]
        .groupby("SoldToPartyID")
        .agg(YTD_2025=("Gross Sales", "sum"))
    )
    tracker = tracker.join(s25.rename("Sales_2025_FY")).join(perf_25_ytd).join(perf_26)
    tracker[["Sales_2025_FY", "YTD_2025", "YTD_2026"]] = tracker[["Sales_2025_FY", "YTD_2025", "YTD_2026"]].fillna(0)

    is_flat_target = tracker["Policy Category"].isin(
        ["New Partner 2026", "Q4'2025 Onboarded", "<5L 2025 Onboarded"]
    )
    custom_target_series = tracker.index.to_series().astype(str).map(custom_targets).fillna(0)
    tracker["2026 Assigned Target"] = np.where(is_flat_target, 2_500_000, custom_target_series)
    tracker["Remaining Months"] = rem_months
    tracker["Req. Monthly Run Rate"] = np.where(
        tracker["2026 Assigned Target"] > 0,
        (tracker["2026 Assigned Target"] - tracker["YTD_2026"]).clip(lower=0) / rem_months,
        0,
    )

    tracker["Achievement %"] = np.where(
        tracker["2026 Assigned Target"] > 0,
        tracker["YTD_2026"] / tracker["2026 Assigned Target"],
        0,
    )

    tracker = tracker.sort_values(by="YTD_2026", ascending=False).reset_index()

    final_cols = [
        "SoldToPartyID", "Patner", "Policy Category", "Old/New/Handholding",
        "Zone", "Branch", "ASM",
        "Sales_2025_FY", "YTD_2025", "YTD_2026", "2026 Assigned Target", "Achievement %",
        "Remaining Months", "Req. Monthly Run Rate",
    ]
    display_df = tracker[final_cols].rename(columns={
        "SoldToPartyID": "Partner ID", "Patner": "Partner Name",
        "Sales_2025_FY": "2025 FY Actual", "YTD_2025": "2025 YTD Actual", "YTD_2026": "2026 YTD Actual",
    })

    display_df.to_excel(writer, sheet_name="Partner_Policy_Tracker", index=False)
    ws = writer.sheets["Partner_Policy_Tracker"]

    for c, val in enumerate(display_df.columns):
        ws.write(0, c, val, f["hdr_pol"])

    n_rows = len(display_df)
    last_row = n_rows + 1

    txt_border = f["txt_det"]

    ws.set_column("A:A", 14, f["policy_num"])
    ws.set_column("B:B", 30, txt_border)
    ws.set_column("C:C", 22, txt_border)
    ws.set_column("D:D", 20, txt_border)
    ws.set_column("E:E", 16, txt_border)
    ws.set_column("F:F", 16, txt_border)
    ws.set_column("G:G", 16, txt_border)
    ws.set_column("H:K", 18, f["policy_num"])
    ws.set_column("L:L", 15, f["pct"])
    ws.set_column("M:M", 16, f["int"])
    ws.set_column("N:N", 22, f["policy_runrate"])
    ws.freeze_panes(1, 2)
    ws.set_tab_color("#F4B183")

    policy_range = f"C2:C{last_row}"
    onh_range = f"D2:D{last_row}"

    policy_colors = {
        "New Partner 2026": "#C6EFCE",
        "Q4'2025 Onboarded": "#D4EEBD",
        ">5L 2025 Onboarded": "#BDD7EE",
        "<5L 2025 Onboarded": "#EEDABD",
        "Old Partner": "#D9D9D9",
    }
    for val, color in policy_colors.items():
        ws.conditional_format(policy_range, {
            "type": "text", "criteria": "containing", "value": val,
            "format": wb.add_format({"bg_color": color, "border": 1}),
        })

    onh_colors = {
        "New Partner": "#C6EFCE",
        "Handholding partner": "#BDD7EE",
        "Old partner": "#D9D9D9",
    }
    for val, color in onh_colors.items():
        ws.conditional_format(onh_range, {
            "type": "text", "criteria": "containing", "value": val,
            "format": wb.add_format({"bg_color": color, "border": 1}),
        })

    ws.conditional_format(f"N2:N{last_row}", {
        "type": "3_color_scale",
        "min_color": "#FFFFFF",
        "mid_color": "#FFEB9C",
        "max_color": "#F8696B",
    })

    ws.conditional_format(f"L2:L{last_row}", {
        "type": "3_color_scale",
        "min_color": "#F8696B",
        "mid_color": "#FFEB9C",
        "max_color": "#63BE7B",
    })

    ws.conditional_format(f"J2:J{last_row}", {
        "type": "data_bar",
        "bar_color": "#63C384",
    })

    return tracker


# =====================================================================
# 10. ZONE TARGET TRACKER (all partners, with formulas)
# =====================================================================
def generate_zone_target_tracker(writer, wb, f):
    df_temp = df_retail.copy()
    if "Zone" not in df_temp.columns:
        df_temp = df_temp.merge(df_final[["SoldToPartyID", "Zone"]].drop_duplicates(), on="SoldToPartyID", how="left")

    perf_26 = df_temp[(df_temp["Year"] == 2026) & (df_temp["InvoiceDate"] <= yesterday)].groupby("Zone").agg(YTD_26=("Gross Sales", "sum")).reset_index()
    perf_25 = df_temp[(df_temp["Year"] == 2025) & (df_temp["InvoiceDate"] <= last_y_ytd_date)].groupby("Zone").agg(YTD_25=("Gross Sales", "sum")).reset_index()
    target_26 = df_region_target[df_region_target["Year"] == 2026].groupby("Zone").agg(Target=("Target", "sum")).reset_index()

    summary = pd.merge(target_26, perf_26, on="Zone", how="outer")
    summary = pd.merge(summary, perf_25, on="Zone", how="outer").fillna(0)
    summary["Growth %"] = (summary["YTD_26"] - summary["YTD_25"]) / summary["YTD_25"].replace(0, np.nan)
    summary["Achievement %"] = summary["YTD_26"] / summary["Target"].replace(0, 1)
    summary["Gap"] = summary["Target"] - summary["YTD_26"]
    summary = summary.sort_values("Target", ascending=False)

    cols = ["Zone", "Target", "YTD_26", "Gap", "Achievement %", "YTD_25", "Growth %"]
    display_df = summary[cols].rename(columns={
        "YTD_26": f"2026 YTD ({yesterday.strftime('%d-%b')})",
        "YTD_25": f"2025 YTD ({last_y_ytd_date.strftime('%d-%b')})",
    })

    sheet_name = "Zone_Target_Tracker"
    display_df.to_excel(writer, sheet_name=sheet_name, index=False)
    ws = writer.sheets[sheet_name]
    for c, val in enumerate(display_df.columns):
        ws.write(0, c, val, f["hdr_alt"])
    ws.set_column("A:A", 25)
    ws.set_column("B:D", 18, f["policy_num"])
    ws.set_column("E:E", 15, f["pct"])
    ws.set_column("F:F", 18, f["policy_num"])
    ws.set_column("G:G", 15, f["pct"])

    last_row = len(display_df) + 1
    tot_fmt = wb.add_format({"bold": True, "bg_color": "#F2F2F2", "border": 1, "num_format": "#,##0"})
    tot_pct_fmt = wb.add_format({"bold": True, "num_format": "0.0%", "bg_color": "#F2F2F2", "border": 1})
    ws.write(last_row, 0, "Grand Total", tot_fmt)
    ws.write_formula(last_row, 1, f"=SUM(B2:B{last_row})", tot_fmt)
    ws.write_formula(last_row, 2, f"=SUM(C2:C{last_row})", tot_fmt)
    ws.write_formula(last_row, 3, f"=B{last_row + 1}-C{last_row + 1}", tot_fmt)
    ws.write_formula(last_row, 4, f"=C{last_row + 1}/B{last_row + 1}", tot_pct_fmt)
    ws.write_formula(last_row, 5, f"=SUM(F2:F{last_row})", tot_fmt)
    ws.write_formula(last_row, 6, f"=(C{last_row + 1}-F{last_row + 1})/F{last_row + 1}", tot_pct_fmt)
    ws.freeze_panes(1, 1)
    ws.set_tab_color("#9DC3E6")
    return summary


# =====================================================================
# 10b. Sheet1 + Geo Expansion (for the separate New_GEO report)
# =====================================================================
SALES_BUCKET_LABELS = [">10 lacs", ">5 lacs", "less than 5 lacs"]


def sales_bucket(sales_2025_fy):
    return np.select(
        [sales_2025_fy >= 1_000_000, sales_2025_fy >= 500_000],
        [SALES_BUCKET_LABELS[0], SALES_BUCKET_LABELS[1]],
        default=SALES_BUCKET_LABELS[2],
    )


def write_sheet1(workbook, fmt, region_new):
    ws = workbook.add_worksheet("Sheet1")
    ws.write(0, 0, "New Partner Performance by Region — Target vs Actual", fmt["title"])
    ws.merge_range(1, 1, 1, 3, f"2026 YTD ({yesterday.strftime('%d-%b')})", fmt["hdr"])
    ws.merge_range(1, 4, 1, 6, f"2025 YTD ({last_year_yesterday.strftime('%d-%b')})", fmt["hdr"])
    ws.write(1, 7, "", fmt["hdr"])
    headers = ["Region", "Partners", "Sales Mnr", "Avg Sales", "Partners", "Sales Mnr", "Avg Sales", "Gwt %"]
    for c, h in enumerate(headers):
        ws.write(2, c, h, fmt["hdr"])

    r = 3
    for zone, row in region_new.iterrows():
        gwt = (row["S26Y"] - row["S25Y"]) / row["S25Y"] if row["S25Y"] else 0
        ws.write(r, 0, zone, fmt["txt_det"])
        ws.write(r, 1, int(row["P26Y"]), fmt["int"])
        ws.write(r, 2, row["S26Y"], fmt["dec"])
        ws.write(r, 3, row["S26Y"] / row["P26Y"] if row["P26Y"] else 0, fmt["dec"])
        ws.write(r, 4, int(row["P25Y"]), fmt["int"])
        ws.write(r, 5, row["S25Y"], fmt["dec"])
        ws.write(r, 6, row["S25Y"] / row["P25Y"] if row["P25Y"] else 0, fmt["dec"])
        ws.write(r, 7, gwt, fmt["pct"])
        r += 1

    sums = region_new.sum()
    gwt_tot = (sums["S26Y"] - sums["S25Y"]) / sums["S25Y"] if sums["S25Y"] else 0
    ws.write(r, 0, "Grand Total", fmt["tot_int"])
    ws.write(r, 1, int(sums["P26Y"]), fmt["tot_int"])
    ws.write(r, 2, sums["S26Y"], fmt["tot_dec"])
    ws.write(r, 3, sums["S26Y"] / sums["P26Y"] if sums["P26Y"] else 0, fmt["tot_dec"])
    ws.write(r, 4, int(sums["P25Y"]), fmt["tot_int"])
    ws.write(r, 5, sums["S25Y"], fmt["tot_dec"])
    ws.write(r, 6, sums["S25Y"] / sums["P25Y"] if sums["P25Y"] else 0, fmt["tot_dec"])
    ws.write(r, 7, gwt_tot, fmt["tot_pct"])
    ws.set_column("A:A", 14)
    ws.set_column("B:H", 15)


def write_geo_expansion_sheet(workbook, fmt, tracker, kro_col_name):
    """
    'Geo Expansion' sheet: policy-tracker-driven "New Partner" population
    (broader than Stat_2026), with Region/RD-KRO breakdowns and an FY2025
    sales-bucket split.

    Table order matches the reference file: Snapshot -> Policy rollup ->
    BU Target box -> Sales-bucket split -> Handholding region table ->
    New Partners region table -> New Geo Expansion Analysis -> RD/KRO
    breakdown -> Product breakup.

    SKIPPED per confirmed scope: the SSP sheet (no Branch/Distributor(SSA)
    target+sales source loaded anywhere), the two stray manual-note values
    from the reference file, and the 'Chls cont.'/'<5 lacs' columns on the
    Handholding region table (no traceable definition -- ask before adding).
    """
    ws = workbook.add_worksheet("Geo Expansion")
    ws.set_tab_color("#ED7D31")

    NEW_PARTNER_SUBCATS = ["<5L 2025 Onboarded", "New Partner 2026", "Q4'2025 Onboarded"]
    HANDHOLDING_CAT = ">5L 2025 Onboarded"

    is_new_partner_tracker = tracker["Old/New/Handholding"] == "New Partner"
    is_handholding_tracker = tracker["Old/New/Handholding"] == "Handholding partner"

    total_new_sales_ytd26 = tracker.loc[is_new_partner_tracker, "YTD_2026"].sum() / 1e6
    bu_target_mn = df_region_target.loc[df_region_target["Year"] == TARGET_YEAR, "Target"].sum() / 1e6
    total_ytd_sales_mn = df_final[
        (df_final["Year"] == 2026) & (df_final["InvoiceDate"] <= yesterday)
    ]["Gross Sales"].sum() / 1e6

    # --- A/B) Snapshot ------------------------------------------------
    r = 0
    ws.merge_range(r, 0, r, 3, f"New Partner Snapshot — YTD {yesterday.strftime('%b-%Y')}", fmt["title"]); r += 2
    ws.write(r, 0, "Description", fmt["hdr"])
    ws.write(r, 1, "No.of Partner", fmt["hdr"])
    ws.write(r, 2, f"YTD {last_year_yesterday.strftime('%b').upper()}'25 (Mn)", fmt["hdr"])
    ws.write(r, 3, f"YTD {yesterday.strftime('%b').upper()}'26 (Mn)", fmt["hdr"])
    r += 1

    def write_cat_row(label, mask, bold=False):
        nonlocal r
        n = int(mask.sum())
        s25 = tracker.loc[mask, "Sales_2025_FY"].sum() / 1e6
        s26 = tracker.loc[mask, "YTD_2026"].sum() / 1e6
        f_int = fmt["tot_int"] if bold else fmt["int"]
        f_txt = fmt["tot_int"] if bold else fmt["txt_det"]
        f_num = fmt["tot_dec"] if bold else fmt["dec"]
        ws.write(r, 0, label, f_txt)
        ws.write(r, 1, n, f_int)
        ws.write(r, 2, s25, f_num)
        ws.write(r, 3, s26, f_num)
        r += 1

    write_cat_row("Handholding partner", is_handholding_tracker, bold=True)
    write_cat_row(HANDHOLDING_CAT, tracker["Policy Category"] == HANDHOLDING_CAT)
    r += 1
    write_cat_row("New Partner", is_new_partner_tracker, bold=True)
    for cat in NEW_PARTNER_SUBCATS:
        write_cat_row(cat, tracker["Policy Category"] == cat)
    r += 1
    write_cat_row("Grand Total", is_new_partner_tracker | is_handholding_tracker, bold=True)
    r += 2

    ws.write(r, 0, "Gross Total BU Target", fmt["kpi_label"]); ws.write(r, 1, bu_target_mn, fmt["kpi_value"]); r += 1
    ws.write(r, 0, f"Total YTD {yesterday.strftime('%b')} Sales", fmt["kpi_label"]); ws.write(r, 1, total_ytd_sales_mn, fmt["kpi_value"]); r += 1
    ws.write(r, 0, "New Partner's 2026 Sales", fmt["kpi_label"]); ws.write(r, 1, total_new_sales_ytd26, fmt["kpi_value"])
    ws.write(r, 2, "New Partner's Sales %", fmt["hdr"]); r += 1
    ws.write(r, 2, total_new_sales_ytd26 / total_ytd_sales_mn if total_ytd_sales_mn else 0, fmt["kpi_pct_hero"])
    r += 3

    # --- C) Partners by FY2025 Sales Bucket ---------------------------
    ws.write(r, 0, "Partners by FY2025 Sales Bucket (Handholding + New Partner)", fmt["subtitle"]); r += 1
    bucketed = tracker.loc[is_new_partner_tracker | is_handholding_tracker].copy()
    bucketed["Bucket"] = sales_bucket(bucketed["Sales_2025_FY"])
    bucket_counts = bucketed["Bucket"].value_counts().reindex(SALES_BUCKET_LABELS, fill_value=0)
    ws.write(r, 0, "Bucket", fmt["hdr"]); ws.write(r, 1, "Partner Count", fmt["hdr"]); r += 1
    for label, cnt in bucket_counts.items():
        ws.write(r, 0, label, fmt["txt_det"]); ws.write(r, 1, int(cnt), fmt["int"]); r += 1
    r += 2

    # --- D) Handholding partner region breakdown -----------------------
    # Region / Partner Count Target / YTD'25 / YTD'26 / Grw% for the
    # Handholding bucket specifically. Header says "Partner Count Target"
    # to match the reference file's own label exactly, even though the
    # values are a partner COUNT, not a target (there's no per-zone
    # Handholding target defined anywhere in this script, same as the
    # reference -- its own numbers are counts under that header too).
    # 'Chls cont.' and '<5 lacs' columns from the reference are still left
    # out -- no traceable definition for either (see chat).
    ws.write(r, 0, "Handholding Partner — Region Breakdown", fmt["title"]); r += 1
    hh_tracker = tracker.loc[is_handholding_tracker]
    headers_hh = ["Region", "Partner Count Target", f"YTD {last_year_yesterday.strftime('%b').upper()}'25 (Mn)",
                  f"YTD {yesterday.strftime('%b').upper()}'26 (Mn)", "Grw %"]
    for c, h in enumerate(headers_hh):
        ws.write(r, c, h, fmt["hdr"])
    r += 1
    hh_zones = sorted(hh_tracker["Zone"].dropna().unique())
    for zone in hh_zones:
        sub = hh_tracker[hh_tracker["Zone"] == zone]
        s25 = sub["Sales_2025_FY"].sum() / 1e6
        s26 = sub["YTD_2026"].sum() / 1e6
        gwt = (s26 - s25) / s25 if s25 else 0
        ws.write(r, 0, zone, fmt["txt_det"])
        ws.write(r, 1, len(sub), fmt["int"])
        ws.write(r, 2, s25, fmt["dec"])
        ws.write(r, 3, s26, fmt["dec"])
        ws.write(r, 4, gwt, fmt["pct"])
        r += 1
    hh_s25_tot = hh_tracker["Sales_2025_FY"].sum() / 1e6
    hh_s26_tot = hh_tracker["YTD_2026"].sum() / 1e6
    hh_gwt_tot = (hh_s26_tot - hh_s25_tot) / hh_s25_tot if hh_s25_tot else 0
    ws.write(r, 0, "Grand Total", fmt["tot_int"])
    ws.write(r, 1, len(hh_tracker), fmt["tot_int"])
    ws.write(r, 2, hh_s25_tot, fmt["tot_dec"])
    ws.write(r, 3, hh_s26_tot, fmt["tot_dec"])
    ws.write(r, 4, hh_gwt_tot, fmt["tot_pct"])
    r += 3

    # --- E) New Partner Performance by Region — Target vs Actual (extended) --
    ws.write(r, 0, "New Partner Performance by Region — Target vs Actual", fmt["title"]); r += 1
    np_tracker = tracker.loc[is_new_partner_tracker]
    # region_new gives the confirmed "New Partners 2025/2026" headcount
    # columns (Stat_2025/Stat_2026 New/Revival partner counts by zone, the
    # same numbers already used on the Snapshot dashboard).
    region_new = aggregate_periods("Zone", new_only=True)

    # Sales (Mn) columns for each sub-category, then a partner-COUNT column
    # for each of the same 3 sub-categories -- same count-per-category logic
    # already used in the RD/KRO breakdown table below, just grouped by
    # region here instead.
    sales_headers = [f"{cat} (Mn)" for cat in NEW_PARTNER_SUBCATS]
    count_headers = [f"{cat} (Count)" for cat in NEW_PARTNER_SUBCATS]
    headers_np = (["Region", "Partner Count Target", "Gr.Sales Target"] + sales_headers
                  + ["Total (Mn)"] + count_headers + ["New Partners 2025", "New Partners 2026"])
    for c, h in enumerate(headers_np):
        ws.write(r, c, h, fmt["hdr_wrap"])
    ws.set_row(r, 30)
    r += 1
    all_zones = sorted(set(NEW_PARTNER_TARGETS_BY_ZONE.index) | set(np_tracker["Zone"].dropna().unique()))
    for zone in all_zones:
        sub = np_tracker[np_tracker["Zone"] == zone]
        ws.write(r, 0, zone, fmt["txt_det"])
        ws.write(r, 1, int(NEW_PARTNER_TARGETS_BY_ZONE["Partner Count Target"].get(zone, 0)), fmt["int"])
        ws.write(r, 2, NEW_PARTNER_TARGETS_BY_ZONE["Gr.Sales Target"].get(zone, 0), fmt["dec"])
        row_total = 0
        for c, cat in enumerate(NEW_PARTNER_SUBCATS, start=3):
            val = sub.loc[sub["Policy Category"] == cat, "YTD_2026"].sum() / 1e6
            ws.write(r, c, val, fmt["dec"])
            row_total += val
        ws.write(r, 6, row_total, fmt["dec"])
        for c, cat in enumerate(NEW_PARTNER_SUBCATS, start=7):
            cnt = int((sub["Policy Category"] == cat).sum())
            ws.write(r, c, cnt, fmt["int"])
        zrow = region_new.loc[zone] if zone in region_new.index else pd.Series({"P25Y": 0, "P26Y": 0})
        ws.write(r, 10, int(zrow["P25Y"]), fmt["int"])
        ws.write(r, 11, int(zrow["P26Y"]), fmt["int"])
        r += 1
    tot_target_p = int(NEW_PARTNER_TARGETS_BY_ZONE["Partner Count Target"].sum())
    tot_target_s = NEW_PARTNER_TARGETS_BY_ZONE["Gr.Sales Target"].sum()
    ws.write(r, 0, "Grand Total", fmt["tot_int"])
    ws.write(r, 1, tot_target_p, fmt["tot_int"])
    ws.write(r, 2, tot_target_s, fmt["tot_dec"])
    grand_total = 0
    for c, cat in enumerate(NEW_PARTNER_SUBCATS, start=3):
        val = np_tracker.loc[np_tracker["Policy Category"] == cat, "YTD_2026"].sum() / 1e6
        ws.write(r, c, val, fmt["tot_dec"])
        grand_total += val
    ws.write(r, 6, grand_total, fmt["tot_dec"])
    for c, cat in enumerate(NEW_PARTNER_SUBCATS, start=7):
        cnt = int((np_tracker["Policy Category"] == cat).sum())
        ws.write(r, c, cnt, fmt["tot_int"])
    ws.write(r, 10, int(region_new["P25Y"].sum()), fmt["tot_int"])
    ws.write(r, 11, int(region_new["P26Y"].sum()), fmt["tot_int"])
    r += 3

    # --- F) New Geo Expansion Analysis ---------------------------------
    # MOVED: this now runs right after the New Partners region table, to
    # match the reference file's table order (it previously ran last).
    r = write_new_geo_table(ws, workbook, fmt, r)
    r += 2

    # --- G) RD/KRO breakdown of the same 3 sub-categories --------------
    # FIXED: sales totals were in raw rupees, and there was no per-category
    # sales breakdown (only counts) -- now matches the region table's
    # shape: a count column AND a sales-Mn column per sub-category.
    ws.write(r, 0, "New Partners by RD/KRO", fmt["subtitle"]); r += 1
    kro_tracker = tracker.loc[is_new_partner_tracker].copy()
    count_headers_kro = [f"{cat} (Count)" for cat in NEW_PARTNER_SUBCATS]
    sales_headers_kro = [f"{cat} (Mn)" for cat in NEW_PARTNER_SUBCATS]
    headers_kro = (["RD/KRO"] + count_headers_kro + sales_headers_kro + [
        f"YTD {last_year_yesterday.strftime('%b').upper()} '25 (Mn)",
        f"YTD {yesterday.strftime('%b').upper()} '26 (Mn)",
    ])
    for c, h in enumerate(headers_kro):
        ws.write(r, c, h, fmt["hdr_wrap"])
    ws.set_row(r, 30)
    r += 1
    kro_map = df_final[["SoldToPartyID", kro_col_name]].drop_duplicates("SoldToPartyID")
    kro_map["SoldToPartyID"] = kro_map["SoldToPartyID"].astype(str)
    kro_tracker = kro_tracker.merge(kro_map, on="SoldToPartyID", how="left")
    for kro_val, sub in kro_tracker.groupby(kro_col_name):
        ws.write(r, 0, kro_val, fmt["txt_det"])
        for c, cat in enumerate(NEW_PARTNER_SUBCATS, start=1):
            n = int((sub["Policy Category"] == cat).sum())
            ws.write(r, c, n, fmt["int"])
        for c, cat in enumerate(NEW_PARTNER_SUBCATS, start=4):
            val = sub.loc[sub["Policy Category"] == cat, "YTD_2026"].sum() / 1e6
            ws.write(r, c, val, fmt["dec"])
        ws.write(r, 7, sub["Sales_2025_FY"].sum() / 1e6, fmt["dec"])
        ws.write(r, 8, sub["YTD_2026"].sum() / 1e6, fmt["dec"])
        r += 1
    ws.write(r, 0, "Grand Total", fmt["tot_int"])
    for c, cat in enumerate(NEW_PARTNER_SUBCATS, start=1):
        ws.write(r, c, int((kro_tracker["Policy Category"] == cat).sum()), fmt["tot_int"])
    for c, cat in enumerate(NEW_PARTNER_SUBCATS, start=4):
        val = kro_tracker.loc[kro_tracker["Policy Category"] == cat, "YTD_2026"].sum() / 1e6
        ws.write(r, c, val, fmt["tot_dec"])
    ws.write(r, 7, kro_tracker["Sales_2025_FY"].sum() / 1e6, fmt["tot_dec"])
    ws.write(r, 8, kro_tracker["YTD_2026"].sum() / 1e6, fmt["tot_dec"])
    r += 3

    # --- H) Product-wise Break-up, New Partner (tracker) population ---
    new_partner_ids = set(tracker.loc[is_new_partner_tracker, "SoldToPartyID"].astype(str))
    r = write_helios_region_pivot(ws, workbook, fmt, r, new_partners_only=True,
                                   title="Product-wise Break-up — NEW PARTNERS (Policy Tracker)",
                                   new_partner_ids=new_partner_ids)

    ws.set_column("A:A", 26)
    ws.set_column("B:L", 16)
    return r


# =====================================================================
# 11. WRITE WORKBOOKS
# =====================================================================
with pd.ExcelWriter(LOCAL_OUTPUT_PATH, engine="xlsxwriter") as writer:
    workbook = writer.book
    fmt = build_formats(workbook)

    ws_sum = workbook.add_worksheet("Summary")
    ws_sum.set_tab_color("#70AD47")
    next_r = write_kpi_table(ws_sum, workbook, fmt, 0, aggregate_periods("Zone"), "Regional Partner Performance", "Region", with_avg=False)
    next_r = write_kpi_table(ws_sum, workbook, fmt, next_r, aggregate_periods("Tier_Bucket"), "Tier Category Analysis", "Tier Category", with_avg=False)
    next_r = write_kpi_table(ws_sum, workbook, fmt, next_r, aggregate_periods(kro_col), "Partner Type Split", "Type", with_avg=False)
    next_r = write_kpi_table(ws_sum, workbook, fmt, next_r, aggregate_periods(None, is_segment=True), "Partner Segment Analysis", "Segment", with_avg=False)

    next_r = write_new_geo_table(ws_sum, workbook, fmt, next_r)
    ws_sum.set_column("A:A", 35); ws_sum.set_column("B:J", 16)

    ws_new = workbook.add_worksheet("New_Partner_Summary")
    ws_new.set_tab_color("#50D08C")
    nr = write_new_partner_dashboard(ws_new, workbook, fmt, start_row=0)
    nr = write_kpi_table(ws_new, workbook, fmt, nr, aggregate_periods(None, is_segment=True), "Partner Segment Analysis (All Partners)", "Segment")
    nr = write_kpi_table(ws_new, workbook, fmt, nr, aggregate_periods("Zone", new_only=True), "NEW PARTNERS: Regional Performance", "Region")
    nr = write_kpi_table(ws_new, workbook, fmt, nr, aggregate_periods("Tier_Bucket", new_only=True), "NEW PARTNERS: Tier Analysis", "Tier Category")
    nr = write_kpi_table(ws_new, workbook, fmt, nr, aggregate_periods(kro_col, new_only=True), "NEW PARTNERS: Partner Type Split", "Type")
    ws_new.set_column("A:A", 24)
    ws_new.set_column("B:L", 12)

    df_mom = df_final[df_final["Year"] == TARGET_YEAR].pivot_table(
        index=["SoldToPartyID", "Patner", "Stat_2026", "ASM", "Branch", "Zone", "City", "Tier_Bucket"],
        columns="Month", values="Gross Sales", aggfunc="sum", fill_value=0,
    )
    df_mom["Partner Total"] = df_mom.sum(axis=1).round(0)
    df_mom = df_mom.reset_index().rename(columns={"Stat_2026": "Segment"})
    df_mom.to_excel(writer, sheet_name="MoM_2026", index=False)
    ws_mom = writer.sheets["MoM_2026"]
    for c, val in enumerate(df_mom.columns):
        ws_mom.write(0, c, val, fmt["hdr"])
        if val == "Partner Total":
            ws_mom.set_column(c, c, 18, fmt["tot_det"])
        elif val in MONTHS_ORDER:
            ws_mom.set_column(c, c, 18, fmt["num_det"])
        else:
            ws_mom.set_column(c, c, 22, fmt["txt_det"])
    ws_mom.freeze_panes(1, 2)

    life_pivot = df_final.groupby(["SoldToPartyID", "Year"])["Gross Sales"].sum().unstack(fill_value=0)
    life_pivot["Lifetime Total"] = life_pivot.sum(axis=1).round(0)
    crit_master = ["Sold-to Party", "Patner", "ASM", "Branch", "Zone", "City", "Tier_Bucket", kro_col]
    life_df = Partner_Master[crit_master].merge(partner_status[["Stat_2026"]], left_on="Sold-to Party", right_index=True, how="left")
    life_df = life_df.merge(life_pivot.reset_index(), left_on="Sold-to Party", right_on="SoldToPartyID", how="left")
    life_df.drop(columns=["SoldToPartyID"], inplace=True, errors="ignore")
    life_df.rename(columns={"Stat_2026": "Segment"}, inplace=True)
    life_df.to_excel(writer, sheet_name="Lifetime_Sales", index=False)
    ws_life = writer.sheets["Lifetime_Sales"]
    for c, val in enumerate(life_df.columns):
        ws_life.write(0, c, val, fmt["hdr"])
        if val == "Lifetime Total":
            ws_life.set_column(c, c, 18, fmt["tot_det"])
        elif isinstance(val, int) or (isinstance(val, str) and val.isdigit()):
            ws_life.set_column(c, c, 18, fmt["num_det"])
        else:
            ws_life.set_column(c, c, 22, fmt["txt_det"])
    ws_life.freeze_panes(1, 2)

    df_new_26 = df_final[df_final["Stat_2026"] == "New/Revival"]
    write_detail(writer, workbook, fmt, df_new_26, "New_Partner_Detail_2026", 2026)

    mask_25_ytd_new = (df_final["Stat_2025"] == "New/Revival") & (df_final["Year"] == 2025) & (df_final["InvoiceDate"] <= last_year_yesterday)
    write_detail(writer, workbook, fmt, df_final[mask_25_ytd_new].copy(), "New_Partner_Detail_2025_YTD", 2025)

    write_detail(writer, workbook, fmt, df_final[df_final["City"].isin(new_geo_list)], "Geo_Expansion_Detail_2026", 2026)
    write_detail(writer, workbook, fmt, df_final[df_final["City"].isin(new_geo_list_25)], "Geo_Expansion_Detail_2025", 2025)

    print("Generating Policy Tracker...")
    tracker = generate_policy_handholding_tracker(writer, workbook, fmt)

    print("Generating Zone Target Tracker...")
    generate_zone_target_tracker(writer, workbook, fmt)

    print("Generating Helios vs Region Analysis (all partners)...")
    ws_helios = workbook.add_worksheet("Helios_Region_Analysis")
    ws_helios.set_tab_color("#BF8F00")
    write_helios_region_pivot(ws_helios, workbook, fmt, 0, new_partners_only=False, title="Product-wise Break-up — All Partners")
    ws_helios.set_column("A:A", 25)
    ws_helios.set_column("B:H", 14)

print(f"Report saved to: {LOCAL_OUTPUT_PATH}")

# =====================================================================
# 11g. Separate, narrower "New_GEO_{year}_YTD_{month}.xlsx" report
# =====================================================================
with pd.ExcelWriter(GEO_OUTPUT_PATH, engine="xlsxwriter") as geo_writer:
    geo_workbook = geo_writer.book
    geo_fmt = build_formats(geo_workbook)

    ws_geo_new = geo_workbook.add_worksheet("New_Partner_Summary")
    ws_geo_new.set_tab_color("#50D08C")
    gr = write_new_partner_dashboard(ws_geo_new, geo_workbook, geo_fmt, start_row=0)
    gr = write_kpi_table(ws_geo_new, geo_workbook, geo_fmt, gr, aggregate_periods(None, is_segment=True), "Partner Segment Analysis (All Partners)", "Segment")
    gr = write_kpi_table(ws_geo_new, geo_workbook, geo_fmt, gr, aggregate_periods("Zone", new_only=True), "NEW PARTNERS: Regional Performance", "Region")
    gr = write_kpi_table(ws_geo_new, geo_workbook, geo_fmt, gr, aggregate_periods("Tier_Bucket", new_only=True), "NEW PARTNERS: Tier Analysis", "Tier Category")
    gr = write_kpi_table(ws_geo_new, geo_workbook, geo_fmt, gr, aggregate_periods(kro_col, new_only=True), "NEW PARTNERS: Partner Type Split", "Type")
    ws_geo_new.set_column("A:A", 24)
    ws_geo_new.set_column("B:L", 12)

    write_geo_expansion_sheet(geo_workbook, geo_fmt, tracker, kro_col)

    # SSP sheet intentionally skipped -- no Branch/Distributor(SSA)
    # target+sales source is loaded in this script yet. Its tab would slot
    # in right here (3rd position) once you point me to that source.

    region_new_for_sheet1 = aggregate_periods("Zone", new_only=True)
    write_sheet1(geo_workbook, geo_fmt, region_new_for_sheet1)

print(f"Geo report saved to: {GEO_OUTPUT_PATH}")
