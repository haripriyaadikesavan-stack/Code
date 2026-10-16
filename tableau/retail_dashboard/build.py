"""Build 'Retail Dashboard.twbx' from the existing Retail workbook's data sources.

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

from core import C_BG, C_DARK, C_MID, C_MUTED, C_PY, C_TARGET, DS, DS_CAPTION, PDS, X, attr, inst_xml, new_uuid, txt

SIDE_BG = "#e9f0ec"
from model import build_model
from pages import DASHBOARDS, DRILL, LAYOUT, NAV, NOTES, SIDE_NOTE, build_sheets, filters_for, params_for
from sheets import run

W = 1400
SRC_TWB = "Overall Retail Dashboard.twb"
TOP_ORDER = ["document-format-change-manifest", "repository-location", "preferences", "style-theme", "style",
             "local-data", "datasources", "datasource-relationships", "mapsources", "shared-views", "actions",
             "worksheets", "dashboards", "windows", "datagraph", "thumbnails", "external"]
BRANCH_ZONE = {"BBN": "East", "CHG": "East", "GAU": "East", "JKD": "East", "NPL": "East", "WB": "East",
               "CHD": "North", "HR": "North", "J&K": "North", "NCR": "North", "NOIDA-GHZ": "North", "UP-1": "North", "UP-2": "North",
               "AP": "South", "KAR": "South", "KER": "South", "TEL": "South", "TN-1": "South", "TN-2": "South",
               "INDORE": "West1", "MUM": "West1", "NAGPUR": "West1", "NASHIK": "West1", "PCMC": "West1", "PUNE": "West1",
               "GUJ": "West2", "RAJ": "West2"}
FAMILIES = ["Unica Pure", "Miluz Zeta", "Livia", "Easy9 MCB 6kA SP", "Easy9 RCCB", "Easy9 MCB 10kA MP", "Easy9 MCB 6kA MP",
            "AvatarOn", "Easy9 DB TPN", "Acti9", "Easy9 Isolator", "Easy9 MCB 10kA SP", "WD- Others", "Easy9 DB SPN",
            "Zencelo", "Opale", "Wiser", "Easy9 DB Phase Selector", "Unica Pure Plus", "Easy9 DB VTPN", "UnicaX", "Electra",
            "Easy9 DB Selector Switch", "PratiKa", "FD- Others", "Easy9 Surge Arrestor", "Easy9 Enclosure",
            "Easy9 Metal P and S", "SPE- Others", "ClipsalX", "COS- Others", "EVC", "KNX"]
# short column headers for the Pavilion measures (Measure Names aliases; trailing spaces keep
# them unique, the same way the old workbook did it)
PAV_ALIASES = {
    "No. of KRO (Minimum!Slab!Achievement)": "No. of KRO", "No. of RO (Minimum!Slab!Achievement)": "No. of RO",
    "No of KRO Achieved": "KRO qualified", "No. of RO Achieved": "RO qualified",
    "No. of KRO (Total!Target!Achievement)": "No. of KRO ", "No. of RO (Total!Target!Achievement)": "No. of RO ",
    "No of KRO Achieved (Total!Target!Achievement)": "KRO qualified ", "No. of RO Achieved (Total!Target!Achievement)": "RO qualified ",
    "No. of KRO": "No. of KRO  ", "No. of RO": "No. of RO  ", "No of KRO > 1 lac": "KRO above 1 L", "No. of RO > 25K": "RO above 25 K",
    "Count of 2026 Classification": "Pavilions", "Count of 2026 Classification (Sales!Classification!Summary)": "Pavilions ",
    "2025-26 Growth %": "Growth %", "2025-26 Growth % (Sales!Classification!Summary)": "Growth % ",
    "Total no. of RO Pavilions": "RO pavilions", "Accepted": "Accepted", "Pending": "Pending",
}


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
        m = re.match(r"ISMEMBEROF\('(.+)'\)$", f.get("expression") or "")
        if not m:
            continue
        members = [x.get("member") for x in f.iter("groupfilter") if x.get("function") == "member"]
        names = [v.strip('"') for v in members if v and v != "%null%"]
        rules.append((m.group(1), names, "%null%" in members))
    return rules


def display_name(c):
    return c.get("caption") or c.get("name", "").strip("[]")


def clean_datasource(ds, R, dsname, aliases=None):
    """Keep connection / extract / relationships; replace all calculations, groups and sets."""
    mine = {f.name: f for f in R.fields.values() if f.ds == dsname}
    for c in list(ds):
        tag = c.tag
        if tag == "column":
            if c.find("calculation") is not None or c.get("name") == "[:Measure Names]":
                ds.remove(c)
            elif c.get("name", "").strip("[]") in mine:
                f = mine[c.get("name").strip("[]")]
                if f.caption:
                    c.set("caption", f.caption)
                if f.fmt:
                    c.set("default-format", f.fmt)
                if c.get("hidden") == "true":
                    del c.attrib["hidden"]
        elif tag in ("column-instance", "group", "style", "default-sorts", "field-sort-info",
                     "datasource-dependencies", "drill-paths"):
            ds.remove(c)
    # rename other fields whose display name clashes with ours
    taken = {f.caption or f.name for f in mine.values()}
    for c in ds.findall("column"):
        n = c.get("name", "").strip("[]")
        if n not in mine and display_name(c) in taken:
            c.set("caption", display_name(c) + " (source)")
    present = {c.get("name") for c in ds.findall("column")}
    idx = list(ds).index(ds.find("extract"))
    adds = [f for f in mine.values() if not f.is_calc and f.ref not in present and (f.caption or f.fmt)]
    calcs = [f for f in mine.values() if f.is_calc]
    if aliases:
        al = "".join(f"<alias key={attr(chr(34) + f'[{dsname}].[sum:{k}:qk]' + chr(34))} value={attr(v)} />" for k, v in aliases.items())
        ds.insert(idx, X(f"<column datatype='string' name='[:Measure Names]' role='dimension' type='nominal'><aliases>{al}</aliases></column>"))
        idx += 1
    for f in adds + calcs:
        ds.insert(idx, X(f.column_xml()))
        idx += 1
    params = set()
    for f in calcs:
        params |= f.params
    og = [c for c in ds if "object-graph" in str(c.tag)][0]
    if params:
        ds.insert(list(ds).index(og), X("<datasource-dependencies datasource='Parameters'>"
                                        + "".join(R.params[p].column_xml(with_members=False) for p in sorted(params))
                                        + "</datasource-dependencies>"))
    return ds


def style_xml(R):
    """Measure Names colours: this year green, last year grey, two years ago light green, targets amber."""
    mn = lambda c: f'"[{DS}].[usr:{R[c].name}:qk]"'
    green = ["Sales", "OB", "MTD sales", "SSP sales"]
    grey = ["Sales PY", "OB PY", "MTD sales LY", "SSP sales PY"]
    amber = ["Monthly target"]
    light = ["Sales 2 years ago"]
    maps = "".join(f"<map to='{C_MID}'><bucket>{txt(mn(c))}</bucket></map>" for c in green)
    maps += "".join(f"<map to='{C_PY}'><bucket>{txt(mn(c))}</bucket></map>" for c in grey)
    maps += "".join(f"<map to='{C_TARGET}'><bucket>{txt(mn(c))}</bucket></map>" for c in amber)
    maps += "".join(f"<map to='#9fd3b5'><bucket>{txt(mn(c))}</bucket></map>" for c in light)
    return f"<style><style-rule element='mark'><encoding attr='color' field='[:Measure Names]' type='palette'>{maps}</encoding></style-rule></style>"


def params_datasource(R):
    return X("<datasource hasconnection='false' inline='true' name='Parameters' version='18.1'><aliases enabled='yes' />"
             + "".join(p.column_xml() for p in R.params.values()) + "</datasource>")


# ---------------------------------------------------------------- dashboards
def u(px, total):
    return str(int(round(px * 100000 / total)))


SIDE_W, HEAD_H, NAV_H, CTRL_H = 224, 56, 38, 58


class Layout:
    """Builds the zone tree of one dashboard from layout containers (Tableau 'layout-flow'),
    so rows, columns and the filter panel can be rearranged in Tableau."""

    def __init__(self, H):
        self.H, self.n = H, 1

    def _id(self):
        self.n += 1
        return self.n

    def box(self, x, y, w, h, inner="", children=(), style=None, fixed=None, **a):
        extra = "".join(f" {k.replace('_', '-')}={attr(v)}" for k, v in a.items())
        if fixed is not None:
            extra += f" fixed-size='{int(fixed)}' is-fixed='true'"
        zs = ""
        if style is not None:
            bg, margin, padding = style
            zs = ("<zone-style><format attr='border-color' value='#000000' /><format attr='border-style' value='none' />"
                  f"<format attr='border-width' value='0' /><format attr='margin' value='{margin}' />"
                  + (f"<format attr='padding' value='{padding}' />" if padding else "")
                  + (f"<format attr='background-color' value='{bg}' />" if bg else "") + "</zone-style>")
        return (f"<zone h='{u(h, self.H)}' id='{self._id()}'{extra} w='{u(w, W)}' x='{u(x, W)}' y='{u(y, self.H)}'>"
                f"{inner}{''.join(children)}{zs}</zone>")

    def flow(self, x, y, w, h, direction, children, fixed=None, style=None):
        return self.box(x, y, w, h, children=children, fixed=fixed, style=style, param=direction, type_v2="layout-flow")


def content_height(rows):
    return sum(h for h, _ in rows)


def dashboard(R, S, dash):
    rows = LAYOUT[dash]
    params = params_for(dash)
    filters = filters_for(dash)
    on_page = [name for _, cells in rows for c in cells
               for name in ([c[0]] if c[0] not in ("stack", "swap") else ([n for n, _ in c[2]] if c[0] == "stack" else c[2]))]
    side_min = 30 + (len(params) + len(filters)) * CTRL_H + 220
    body_h = max(content_height(rows), side_min)
    H = HEAD_H + NAV_H + body_h + 8
    L = Layout(H)
    used = set(on_page)
    i = DASHBOARDS.index(dash)

    # ---- header
    head = (run("Retail Performance  ·  ", "#cfe9d7", 12) + run(f"{i + 1} / {len(DASHBOARDS)}   ", "#cfe9d7", 10)
            + run(dash, "#ffffff", 18, True) + run("      " + NOTES[dash], "#cfe9d7", 10))
    header = L.box(0, 0, W, HEAD_H, f"<formatted-text>{head}</formatted-text>", style=(C_DARK, 0, 10), fixed=HEAD_H, type_v2="text")

    # ---- navigation: chapters, then back / next
    nav_items = []
    bn = 92
    extra = [n for n, ok in (("Nav | ‹ Back", i > 0), ("Nav | Next ›", i < len(DASHBOARDS) - 1)) if ok]
    n = len(NAV)
    bw = (W - len(extra) * bn) / n
    x = 0
    for label, target in NAV:
        nav_items.append(L.box(x, HEAD_H, bw, NAV_H, style=(C_DARK if target == dash else C_MID, 3, 0),
                               name=f"Nav | {label}", show_title="false"))
        used.add(f"Nav | {label}")
        x += bw
    for name in extra:
        nav_items.append(L.box(x, HEAD_H, bn, NAV_H, style=(C_TARGET, 3, 0), fixed=bn, name=name, show_title="false"))
        used.add(name)
        x += bn
    nav = L.flow(0, HEAD_H, W, NAV_H, "horz", nav_items, fixed=NAV_H, style=(C_BG, 0, 0))

    # ---- filter panel
    y0 = HEAD_H + NAV_H
    side = [L.box(0, y0, SIDE_W, 30, "<formatted-text>" + run("Filters & settings", C_DARK, 10, True) + "</formatted-text>",
                  style=(None, 4, 4), fixed=30, type_v2="text")]
    y = y0 + 30
    for pname, mode in params:
        side.append(L.box(0, y, SIDE_W, CTRL_H, style=("#ffffff", 3, 0), fixed=CTRL_H, mode=mode,
                          param=f"[Parameters].[{pname}]", type_v2="paramctrl"))
        y += CTRL_H
    filter_deps = []
    for field, _, _ in filters:
        src = next((s for s in on_page if field in getattr(S.get(s), "filter_refs", {})), None)
        if src is None:
            continue
        ref = S[src].filter_refs[field]
        side.append(L.box(0, y, SIDE_W, CTRL_H, style=("#ffffff", 3, 0), fixed=CTRL_H, mode="checkdropdown", name=src,
                          param=ref, show_apply="true", type_v2="filter", values="relevant"))
        filter_deps.append(field)
        y += CTRL_H
    note = SIDE_NOTE if dash != "Pavilion" else (
        "Slab and qualification tables come from Pavilion_Unpivoted.xlsx as published; the filters do not apply to them.\n" + SIDE_NOTE)
    side.append(L.box(0, y, SIDE_W, y0 + body_h - y, "<formatted-text>" + run(note, C_MUTED, 8) + "</formatted-text>",
                      style=(None, 6, 4), type_v2="text"))
    sidebar = L.flow(0, y0, SIDE_W, body_h, "vert", side, fixed=SIDE_W, style=(SIDE_BG, 0, 0))

    # ---- content: one horizontal container per row
    cx, cw = SIDE_W, W - SIDE_W
    content = []
    y = y0
    for h, cells in rows:
        total = sum(c[1] if c[0] in ("stack", "swap") else c[1] for c in cells)
        x = cx
        kids = []
        for c in cells:
            wgt = c[1]
            w = cw * wgt / total
            if c[0] == "stack":
                yy, sub = y, []
                for name, hh in c[2]:
                    sub.append(sheet_zone(L, S, name, x, yy, w, hh))
                    yy += hh
                kids.append(L.flow(x, y, w, h, "vert", sub))
            elif c[0] == "swap":       # one sheet shows at a time; the others are empty and collapse
                k = len(c[2])
                sub = [sheet_zone(L, S, name, x, y + j * h / k, w, h / k, title=False) for j, name in enumerate(c[2])]
                kids.append(L.flow(x, y, w, h, "vert", sub))
            else:
                kids.append(sheet_zone(L, S, c[0], x, y, w, h))
            x += w
        content.append(L.flow(cx, y, cw, h, "horz", kids, fixed=h))
        y += h
    if y < y0 + body_h:
        content.append(L.box(cx, y, cw, y0 + body_h - y, type_v2="empty"))
    body = L.flow(cx, y0, cw, body_h, "vert", content)
    main = L.flow(0, y0, W, body_h, "horz", [sidebar, body])
    page = L.flow(0, 0, W, H, "vert", [header, nav, main])

    # ---- dependencies: parameters on the panel, fields of the filters
    pnames = sorted({p for p, _ in params})
    deps = "<datasources><datasource name='Parameters' />" + (f"<datasource caption={attr(DS_CAPTION)} name='{DS}' />" if filter_deps else "") + "</datasources>"
    deps += ("<datasource-dependencies datasource='Parameters'>" + "".join(R.params[p].column_xml() for p in pnames)
             + "</datasource-dependencies>")
    if filter_deps:
        fields = [R[f] if f in R.by_caption else R.fields[f] for f in filter_deps]
        closure, _ = R.closure({f.key for f in fields})
        cols = sorted((R.fields[k] for k in closure), key=lambda f: f.name)
        deps += (f"<datasource-dependencies datasource='{DS}'>" + "".join(f.column_xml() for f in cols)
                 + "".join(inst_xml(f, "dim") for f in fields) + "</datasource-dependencies>")
    style = (f"<style><style-rule element='table'><format attr='background-color' value='{C_BG}' /></style-rule>"
             "<style-rule element='parameter-ctrl'><format attr='font-size' value='9' /><format attr='color' value='#333333' />"
             "<format attr='border-color' value='#d9d9d9' /><format attr='border-width' value='1' /><format attr='border-style' value='solid' />"
             "<format attr='background-color' value='#ffffff' /></style-rule>"
             f"<style-rule element='parameter-ctrl-title'><format attr='font-size' value='8' /><format attr='font-weight' value='bold' /><format attr='color' value='{C_DARK}' /></style-rule></style>")
    root = (f"<zone h='100000' id='1' type-v2='layout-basic' w='100000' x='0' y='0'>{page}"
            f"<zone-style><format attr='background-color' value='{C_BG}' /></zone-style></zone>")
    xml = (f"<dashboard name={attr(dash)}>{style}<size maxheight='{H}' maxwidth='{W}' minheight='{H}' minwidth='{W}' />"
           f"{deps}<zones>{root}</zones><simple-id uuid='{new_uuid()}' /></dashboard>")
    return X(xml), used


def sheet_zone(L, S, name, x, y, w, h, title=True):
    if name.startswith("H | "):
        return L.box(x, y, w, h, style=("#eaf4ee", 4, 6), name=name, show_title="false")
    return L.box(x, y, w, h, style=("#ffffff", 4, 4), name=name, show_title="true" if title else "false")


def _action(caption, dash, sheet, target):
    return (f"<action caption={attr(caption)} name='[Action_{uuid.uuid4().hex.upper()}]'>"
            f"<activation auto-clear='true' type='on-select' />"
            f"<source dashboard={attr(dash)} type='sheet' worksheet={attr(sheet)} />"
            f"<command command='tsc:tsl-filter'><param name='special-fields' value='all' /><param name='target' value={attr(target)} /></command></action>")


def actions_xml(S, dash_used):
    out = []
    for dash, used in dash_used.items():
        i = DASHBOARDS.index(dash)
        for name in sorted(used):
            s = S[name]
            if not name.startswith("Nav | "):
                continue
            if name == "Nav | ‹ Back":
                target = DASHBOARDS[i - 1]
            elif name == "Nav | Next ›":
                target = DASHBOARDS[i + 1]
            else:
                target = s.nav_target
            if target == dash:
                continue
            out.append(_action("Go to " + target, dash, name, target))
        # click a row: filter the rest of this dashboard (as the old dashboards did)
        for name in DRILL.get(dash, []):
            out.append(_action(f"Filter {dash} by {name.split(' | ')[1].lower()}", dash, name, dash))
    return X("<actions>" + "".join(out) + "</actions>")


def windows_xml(S, dash_used):
    out = []
    for i, (dash, used) in enumerate(dash_used.items()):
        vps = "".join(f"<viewpoint name={attr(n)}>" + (f"<zoom type='{S[n].fit}' />" if S[n].fit else "") + "</viewpoint>"
                      for n in sorted(used))
        mx = " maximized='true'" if i == 0 else ""
        out.append(f"<window class='dashboard'{mx} name={attr(dash)}><viewpoints>{vps}</viewpoints><active id='-1' />"
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
    pav = [d for d in ds_list if d.get("name") == PDS][0]
    R = build_model(sorted(BRANCH_ZONE), BRANCH_ZONE, FAMILIES, rls_rules(retail))
    for k in ("2025-26 Growth %", "2025-26 Growth % (Sales!Classification!Summary)"):
        R.fields[f"{PDS}|{k}"].fmt = "p0.0%"

    clean_datasource(retail, R, DS)
    retail.insert(list(retail).index(retail.find("layout")) + 1, X(style_xml(R)))
    clean_datasource(pav, R, PDS, aliases=PAV_ALIASES)
    for d in list(ds_list):
        ds_list.remove(d)
    ds_list.append(params_datasource(R))
    ds_list.append(retail)
    ds_list.append(pav)

    S = build_sheets(R)
    for tag in ("repository-location", "actions", "worksheets", "dashboards", "windows", "thumbnails"):
        e = root.find(tag)
        if e is not None:
            root.remove(e)
    ws = etree.SubElement(root, "worksheets")
    dbs = etree.SubElement(root, "dashboards")
    dash_used = {}
    for d in DASHBOARDS:
        el, used = dashboard(R, S, d)
        dbs.append(el)
        dash_used[d] = used
    missing = sorted({n for used in dash_used.values() for n in used if n not in S})
    if missing:
        raise ValueError(f"dashboards reference missing sheets: {missing}")
    on_dash = set().union(*dash_used.values())
    S = {n: s for n, s in S.items() if n in on_dash}        # drop sheets not placed anywhere
    for s in S.values():
        ws.append(X(s.xml()))
    root.append(actions_xml(S, dash_used))
    root.append(windows_xml(S, dash_used))
    rank = lambda e: next((i for i, t in enumerate(TOP_ORDER) if str(e.tag) == t or str(e.tag).endswith("..." + t)), 99)
    kids = sorted(list(root), key=rank)
    if any(rank(k) == 99 for k in kids):
        raise ValueError(f"unknown top-level element(s): {[k.tag for k in kids if rank(k) == 99]}")
    for k in kids:
        root.remove(k)
        root.append(k)

    twb = etree.tostring(tree, xml_declaration=True, encoding="utf-8", pretty_print=True)
    twb = twb.replace(b"<?xml version='1.0' encoding='utf-8'?>", b"<?xml version='1.0' encoding='utf-8' ?>", 1)
    name = os.path.splitext(os.path.basename(out))[0]
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr(f"{name}.twb", twb)
        for hyper in ("Data/Extracts/federated_14wc2mc1wc09dr1b4kqf30.hyper", "Data/Extracts/federated_00walea07guk6w1fe7zhg0.hyper"):
            z.write(os.path.join(tmp, hyper), hyper)
    open(out.replace(".twbx", ".twb"), "wb").write(twb)
    shutil.rmtree(tmp)
    print(f"built {out}: {len(S)} sheets, {len(DASHBOARDS)} dashboards, {sum(1 for f in R.fields.values() if f.is_calc)} calculations")
    return R, S


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
