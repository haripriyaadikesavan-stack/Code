"""Build 'Retail Dashboard.twbx' from the existing Retail workbook's data source.

usage: python build.py <Overall Retail Dashboard.twbx> <output.twbx>
"""
import os
import re
import shutil
import sys
import tempfile
import uuid
import zipfile

from lxml import etree

from core import (C_BG, C_DARK, C_MID, C_MUTED, C_PY, C_TARGET, C_TEXT, DS, X, attr, inst, iref, new_uuid, txt)
from model import build_model
from pages import LAYOUT, NOTES, PAGES, SIDEBAR_PARAMS, build_sheets
from sheets import MN, run

W = 1400
EXTRA_PARAMS = {"Explore": [("p Dim", "compact")]}
SRC_TWB = "Overall Retail Dashboard.twb"


def read_source(twbx):
    tmp = tempfile.mkdtemp()
    with zipfile.ZipFile(twbx) as z:
        z.extractall(tmp)
    return tmp, etree.parse(os.path.join(tmp, SRC_TWB))


def rls_rules(ds):
    """Read the 'Branch RLS' user filter: [(group, [branches], include_null)]."""
    g = [x for x in ds.findall("group") if x.get("name") == "[User Filter 2]"][0]
    union = g[0][1]
    rules = []
    for f in union:
        expr = f.get("expression") or ""
        m = re.match(r"ISMEMBEROF\('(.+)'\)$", expr)
        if not m:
            continue
        members = [x.get("member") for x in f.iter("groupfilter") if x.get("function") == "member"]
        names = [v.strip('"') for v in members if v and v != "%null%"]
        rules.append((m.group(1), names, "%null%" in members))
    return rules


def clean_datasource(ds, R):
    """Keep the connection / extract / relationships; replace all calculations."""
    keep_base = {f.name for f in R.fields.values() if not f.is_calc}
    for c in list(ds):
        tag = c.tag
        if tag == "column":
            if c.find("calculation") is not None or c.get("name") == "[:Measure Names]":
                ds.remove(c)
            elif c.get("name", "").strip("[]") in keep_base:
                f = R.fields[c.get("name").strip("[]")]
                if f.caption:
                    c.set("caption", f.caption)
                if c.get("hidden") == "true":
                    del c.attrib["hidden"]
        elif tag in ("column-instance", "group", "style", "default-sorts", "field-sort-info",
                     "datasource-dependencies", "drill-paths"):
            ds.remove(c)
    # base fields that had no <column> element yet -> add one so captions show
    present = {c.get("name") for c in ds.findall("column")}
    anchor = ds.find("extract")
    idx = list(ds).index(anchor)
    adds = [f for f in R.fields.values() if not f.is_calc and f.ref not in present and f.caption]
    calcs = [f for f in R.fields.values() if f.is_calc]
    for f in adds + calcs:
        ds.insert(idx, X(f.column_xml()))
        idx += 1
    # parameters used by the data source calcs
    params = set()
    for f in calcs:
        params |= f.params
    deps = X("<datasource-dependencies datasource='Parameters'>"
             + "".join(R.params[p].column_xml(with_members=False) for p in sorted(params)) + "</datasource-dependencies>")
    # colours for discrete fields (Measure Names and bands): after <layout>
    lay = ds.find("layout")
    ds.insert(list(ds).index(lay) + 1, X(style_xml(R)))
    og = [c for c in ds if "object-graph" in str(c.tag)][0]
    ds.insert(list(ds).index(og), deps)


def style_xml(R):
    def mn(caption):
        f = R[caption]
        return f'"[{DS}].[usr:{f.name}:qk]"'
    green = ["Sales", "OB", "Wiser sales", "Zeta sales", "New / revival partners", "Partners (count)"]
    grey = ["Sales PY", "OB PY"]
    amber = ["Target", "Wiser target", "Zeta target"]
    maps = "".join(f"<map to='{C_MID}'><bucket>{txt(mn(c))}</bucket></map>" for c in green)
    maps += "".join(f"<map to='{C_PY}'><bucket>{txt(mn(c))}</bucket></map>" for c in grey)
    maps += "".join(f"<map to='{C_TARGET}'><bucket>{txt(mn(c))}</bucket></map>" for c in amber)
    band_cols = {"1. 100% and above": "#1f7a4d", "2. 90-100%": "#7fbf7b", "3. 75-90%": "#e8a33d",
                 "4. Below 75%": "#c0392b", "No target": "#9aa5a0"}
    bands = lambda cap: (f"<encoding attr='color' field='[usr:{R[cap].name}:nk]' type='palette'>"
                         + "".join(f"<map to='{c}'><bucket>&quot;{k}&quot;</bucket></map>" for k, c in band_cols.items())
                         + "</encoding>")
    growth = (f"<encoding attr='color' field='[usr:{R['Growth band'].name}:nk]' type='palette'>"
              f"<map to='#1f7a4d'><bucket>&quot;Growing&quot;</bucket></map><map to='#c0392b'><bucket>&quot;Declining&quot;</bucket></map>"
              f"<map to='#9aa5a0'><bucket>&quot;New&quot;</bucket></map></encoding>")
    return ("<style><style-rule element='mark'>"
            f"<encoding attr='color' field='[:Measure Names]' type='palette'>{maps}</encoding>"
            + bands("Ach band") + bands("Partner Ach band") + growth
            + "</style-rule></style>")


def params_datasource(R):
    return X("<datasource hasconnection='false' inline='true' name='Parameters' version='18.1'><aliases enabled='yes' />"
             + "".join(p.column_xml() for p in R.params.values()) + "</datasource>")


# ---------------------------------------------------------------- dashboards
def u(px, total):
    return str(int(round(px * 100000 / total)))


class Zones:
    def __init__(self, H):
        self.H, self.n, self.items = H, 1, []

    def zone(self, x, y, w, h, inner, bg=None, margin=0, padding=0, **a):
        self.n += 1
        extra = "".join(f" {k.replace('_', '-')}={attr(v)}" for k, v in a.items())
        style = ("<zone-style><format attr='border-color' value='#000000' /><format attr='border-style' value='none' />"
                 f"<format attr='border-width' value='0' /><format attr='margin' value='{margin}' />"
                 + (f"<format attr='padding' value='{padding}' />" if padding else "")
                 + (f"<format attr='background-color' value='{bg}' />" if bg else "") + "</zone-style>")
        self.items.append(f"<zone h='{u(h, self.H)}' id='{self.n}'{extra} w='{u(w, W)}' x='{u(x, W)}' y='{u(y, self.H)}'>{inner}{style}</zone>")


def dashboard(R, page, S):
    H, items = LAYOUT[page]
    Z = Zones(H)
    head = (run("Retail Performance  ·  ", "#cfe9d7", 13) + run(page if page != "Home" else "Overview", "#ffffff", 18, True)
            + run("      " + NOTES[page], "#cfe9d7", 10))
    Z.zone(0, 0, W, 56, f"<formatted-text>{head}</formatted-text>", bg=C_DARK, padding=10, type_v2="text")
    # navigation row
    bw, gap, x0 = 160, 8, 14
    for i, p in enumerate(PAGES):
        Z.zone(x0 + i * (bw + gap), 64, bw, 32, "", bg=C_MID, name=f"Nav | {p}", show_title="false")
    # sidebar
    Z.zone(14, 112, 200, 26, "<formatted-text>" + run("Filters & settings", C_DARK, 10, True) + "</formatted-text>", type_v2="text")
    y = 140
    for pname, mode in SIDEBAR_PARAMS + EXTRA_PARAMS.get(page, []):
        Z.zone(14, y, 200, 52, "", bg=C_BG, margin=2, mode=mode, param=f"[Parameters].[{pname}]", type_v2="paramctrl")
        y += 54
    note = ("Period: YTD / QTD / MTD run to today; 'Custom range' uses the two dates.\n"
            "Growth compares the same dates last year.\n"
            "Targets: monthly branch targets (net); Gross shows net / 0.83. "
            "'Pro-rated to date' counts the current month's target only up to today.\n"
            "Partner pages follow the Branch RLS groups.")
    Z.zone(14, y + 6, 200, 230, "<formatted-text>" + run(note, C_MUTED, 8) + "</formatted-text>", type_v2="text")
    for name, x, yy, w, h in items:
        Z.zone(x, yy, w, h, "", bg="#ffffff", margin=4, padding=4, name=name, show_title="true")
    used = {n for n, *_ in items} | {f"Nav | {p}" for p in PAGES}
    params = sorted({p for p, _ in SIDEBAR_PARAMS + EXTRA_PARAMS.get(page, [])})
    pdeps = ("<datasources><datasource name='Parameters' /></datasources><datasource-dependencies datasource='Parameters'>"
             + "".join(R.params[p].column_xml() for p in params) + "</datasource-dependencies>")
    style = (f"<style><style-rule element='table'><format attr='background-color' value='{C_BG}' /></style-rule>"
             "<style-rule element='parameter-ctrl'><format attr='font-size' value='9' /><format attr='color' value='#333333' />"
             "<format attr='border-color' value='#d9d9d9' /><format attr='border-width' value='1' /><format attr='border-style' value='solid' />"
             "<format attr='background-color' value='#ffffff' /></style-rule>"
             f"<style-rule element='parameter-ctrl-title'><format attr='font-size' value='8' /><format attr='font-weight' value='bold' /><format attr='color' value='{C_DARK}' /></style-rule></style>")
    root = (f"<zone h='100000' id='1' type-v2='layout-basic' w='100000' x='0' y='0'>{''.join(Z.items)}"
            f"<zone-style><format attr='background-color' value='{C_BG}' /></zone-style></zone>")
    xml = (f"<dashboard name={attr(page)}>{style}<size maxheight='{H}' maxwidth='{W}' minheight='{H}' minwidth='{W}' />"
           f"{pdeps}<zones>{root}</zones><simple-id uuid='{new_uuid()}' /></dashboard>")
    return X(xml), used


def actions_xml(dash_used):
    out = []
    for page, used in dash_used.items():
        for p in PAGES:
            if p == page:
                continue
            out.append(f"<action caption={attr('Go to ' + p)} name='[Action_{uuid.uuid4().hex.upper()}]'>"
                       f"<activation auto-clear='true' type='on-select' />"
                       f"<source dashboard={attr(page)} type='sheet' worksheet={attr('Nav | ' + p)} />"
                       f"<command command='tsc:tsl-filter'><param name='special-fields' value='all' /><param name='target' value={attr(p)} /></command></action>")
    return X("<actions>" + "".join(out) + "</actions>")


def windows_xml(S, dash_used):
    out = []
    for i, (page, used) in enumerate(dash_used.items()):
        vps = "".join(f"<viewpoint name={attr(n)}><zoom type='entire-view' /></viewpoint>" for n in sorted(used))
        mx = " maximized='true'" if i == 0 else ""
        out.append(f"<window class='dashboard'{mx} name={attr(page)}><viewpoints>{vps}</viewpoints><active id='-1' />"
                   f"<simple-id uuid='{new_uuid()}' /></window>")
    for name in S:
        out.append(f"<window class='worksheet' hidden='true' name={attr(name)}><cards><edge name='left'><strip size='160'>"
                   "<card type='pages' /><card type='filters' /><card type='marks' /></strip></edge><edge name='top'>"
                   "<strip size='2147483647'><card type='columns' /></strip><strip size='2147483647'><card type='rows' /></strip>"
                   f"<strip size='31'><card type='title' /></strip></edge></cards><simple-id uuid='{new_uuid()}' /></window>")
    return X("<windows source-height='37'>" + "".join(out) + "</windows>")


def main(src, out):
    tmp, tree = read_source(src)
    root = tree.getroot()
    ds_list = root.find("datasources")
    retail = [d for d in ds_list if d.get("name") == DS][0]
    rules = rls_rules(retail)
    branch_zone = {"BBN": "East", "CHG": "East", "GAU": "East", "JKD": "East", "NPL": "East", "WB": "East",
                   "CHD": "North", "HR": "North", "J&K": "North", "NCR": "North", "NOIDA-GHZ": "North", "UP-1": "North", "UP-2": "North",
                   "AP": "South", "KAR": "South", "KER": "South", "TEL": "South", "TN-1": "South", "TN-2": "South",
                   "INDORE": "West1", "MUM": "West1", "NAGPUR": "West1", "NASHIK": "West1", "PCMC": "West1", "PUNE": "West1",
                   "GUJ": "West2", "RAJ": "West2"}
    families = ["Unica Pure", "Miluz Zeta", "Livia", "Easy9 MCB 6kA SP", "Easy9 RCCB", "Easy9 MCB 10kA MP", "Easy9 MCB 6kA MP",
                "AvatarOn", "Easy9 DB TPN", "Acti9", "Easy9 Isolator", "Easy9 MCB 10kA SP", "WD- Others", "Easy9 DB SPN",
                "Zencelo", "Opale", "Wiser", "Easy9 DB Phase Selector", "Unica Pure Plus", "Easy9 DB VTPN", "UnicaX", "Electra",
                "Easy9 DB Selector Switch", "PratiKa", "FD- Others", "Easy9 Surge Arrestor", "Easy9 Enclosure",
                "Easy9 Metal P and S", "SPE- Others", "ClipsalX", "COS- Others", "EVC", "KNX"]
    R = build_model(sorted(branch_zone), branch_zone, families, rules)

    clean_datasource(retail, R)
    for d in list(ds_list):
        ds_list.remove(d)
    ds_list.append(params_datasource(R))
    ds_list.append(retail)

    S = build_sheets(R)
    for tag in ("repository-location", "actions", "worksheets", "dashboards", "windows", "thumbnails"):
        e = root.find(tag)
        if e is not None:
            root.remove(e)
    ws = etree.SubElement(root, "worksheets")
    for s in S.values():
        ws.append(X(s.xml()))
    dbs = etree.SubElement(root, "dashboards")
    dash_used = {}
    for p in PAGES:
        d, used = dashboard(R, p, S)
        dbs.append(d)
        dash_used[p] = used
    root.append(actions_xml(dash_used))
    root.append(windows_xml(S, dash_used))
    # Tableau validates the order of the top-level elements
    order = ["document-format-change-manifest", "repository-location", "preferences", "style-theme", "style",
             "local-data", "datasources", "datasource-relationships", "mapsources", "shared-views", "actions",
             "worksheets", "dashboards", "windows", "datagraph", "thumbnails", "external"]
    rank = lambda e: next((i for i, t in enumerate(order) if str(e.tag) == t or str(e.tag).endswith("..." + t)), 99)
    kids = sorted(list(root), key=rank)
    if any(rank(k) == 99 for k in kids):
        raise ValueError(f"unknown top-level element(s): {[k.tag for k in kids if rank(k) == 99]}")
    for k in kids:
        root.remove(k)
        root.append(k)

    # ---- package: workbook + the primary extract only (the Pavilion extract is not used)
    twb = etree.tostring(tree, xml_declaration=True, encoding="utf-8", pretty_print=True)
    twb = twb.replace(b"<?xml version='1.0' encoding='utf-8'?>", b"<?xml version='1.0' encoding='utf-8' ?>", 1)
    name = os.path.splitext(os.path.basename(out))[0]
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr(f"{name}.twb", twb)
        hyper = "Data/Extracts/federated_14wc2mc1wc09dr1b4kqf30.hyper"
        z.write(os.path.join(tmp, hyper), hyper)
    open(out.replace(".twbx", ".twb"), "wb").write(twb)
    shutil.rmtree(tmp)
    print(f"built {out}: {len(S)} sheets, {len(PAGES)} dashboards, {sum(1 for f in R.fields.values() if f.is_calc)} calculations")
    return R, S


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
