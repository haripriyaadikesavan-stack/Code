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

from core import NP_ALIASES, C_BG, C_DARK, C_MID, C_MUTED, C_PY, C_TARGET, DS, DS_CAPTION, PDS, X, attr, inst_xml, new_uuid, txt

SIDE_BG = "#e9f0ec"
from model import build_model
from pages import DASHBOARDS, DRILL, LAYOUT, NAV, NOTES, SIDE_NOTE, VISIBILITY, build_sheets, filters_for, params_for
from sheets import run

W = 1500
TITLE = "Retail Performance"
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
    name = SRC_TWB if os.path.exists(os.path.join(tmp, SRC_TWB)) else [f for f in os.listdir(tmp) if f.endswith(".twb")][0]
    return tmp, etree.parse(os.path.join(tmp, name))


def missing_fields(R, ds_by_name):
    """Base fields of the model whose table is not in the data source any more (e.g. a table removed from the
    extract), and every calculation that depends on one."""
    present = {}
    for name, ds in ds_by_name.items():
        present[name] = {m.findtext("local-name") for m in ds.find("connection").iter("metadata-record")
                         if m.get("class") == "column"}
    bad = {k for k, f in R.fields.items() if not f.is_calc and f.ref not in present.get(f.ds, set())}
    changed = True
    while changed:
        changed = False
        for k, f in R.fields.items():
            if f.is_calc and k not in bad and f.deps & bad:
                bad.add(k)
                changed = True
    return bad


def drop_missing(R, S, bad):
    """Remove the sheets that use a missing field, the layout cells that showed them, and the fields themselves."""
    gone = set()
    for n, s in list(S.items()):
        keys, _ = R.closure({f.key for f, _ in s.uses} | {f.key for f in s.extra_fields})
        if keys & bad:
            gone.add(n)
            del S[n]
    for dash, rows in LAYOUT.items():
        new = []
        for h, cells in rows:
            keep = [c for c in cells if all(n in S for n in cell_names(c))]
            if keep:
                new.append((h, keep))
        rows[:] = new
    for k in bad:
        f = R.fields.pop(k)
        if f.is_calc:
            R.by_caption.pop(f.caption, None)
    return sorted(gone)


def rls_rules(ds):
    """Read the 'Branch RLS' user filter: [(group, [branches], include_null)]."""
    gs = [x for x in ds.findall("group") if x.get("name") == "[User Filter 2]"]
    if not gs:
        return []           # no user filter in this workbook: everyone sees every branch
    g = gs[0]
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
        key = lambda k: k if ":" in k else f"sum:{k}:qk"       # a column-instance name, or a field summed
        al = "".join(f"<alias key={attr(chr(34) + f'[{dsname}].[{key(k)}]' + chr(34))} value={attr(v)} />" for k, v in aliases.items())
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

    def __init__(self, H, width=None, start=1):
        self.H, self.W, self.n = H, width or W, start

    def _id(self):
        self.n += 1
        return self.n

    def box(self, x, y, w, h, inner="", children=(), style=None, fixed=None, zid=None, **a):
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
        self.last = zid if zid is not None else self._id()
        return (f"<zone h='{u(h, self.H)}' id='{self.last}'{extra} w='{u(w, self.W)}' x='{u(x, self.W)}' y='{u(y, self.H)}'>"
                f"{inner}{''.join(children)}{zs}</zone>")

    def flow(self, x, y, w, h, direction, children, fixed=None, style=None):
        return self.box(x, y, w, h, children=children, fixed=fixed, style=style, param=direction, type_v2="layout-flow")


def title_text(s):
    t = re.sub(r"<[^>]+>", "", s.title or "")
    return re.sub(r"&lt;\[Parameters\]\.\[p Units\]&gt;", "Mn", t).replace("&amp;", "&")


def need_height(s, w, fit=True):
    """Height a table needs to show all its rows (estimated), or None for charts and long tables."""
    if s.kind != "table" or s.est_rows is None:
        return None
    if fit:
        fit_columns(s, w)
    title, *subs = title_text(s).split("\n")
    lines = lambda text, px: max(1, -(-int(len(text) * px) // max(40, int(w) - 28)))
    head = 0
    if s.cols:
        hs = s.header_style()
        head = int(re.search(r"value='(\d+)'", hs).group(1)) if hs else 26
    return 26 + 22 * lines(title, 7.4) + sum(14 * lines(x, 5.2) for x in subs) + head + 22 * s.est_rows + 6


def cell_names(c):
    return [c[0]] if c[0] not in ("stack", "swap") else ([n for n, _ in c[2]] if c[0] == "stack" else list(c[2]))


def fit_rows(S, rows, cw):
    """Row heights: rows of small tables get exactly the height their tables need; other rows keep their height."""
    out = []
    for h, cells in rows:
        total = sum(c[1] for c in cells)
        needs, other = [], False
        for c in cells:
            n = need_height(S[c[0]], cw * c[1] / total) if c[0] not in ("stack", "swap") else None
            if n is None:
                other = True
            else:
                needs.append(n)
        if needs and max(needs) <= 900:
            h = max(needs + ([h] if other else []))
        out.append((h, cells))
    return out


def configure(pages, title):
    """Use the layout engine for another workbook: its pages module (CHAPTERS, LAYOUT, filters ...) and header title."""
    global DASHBOARDS, DRILL, LAYOUT, NAV, NOTES, SIDE_NOTE, VISIBILITY, filters_for, params_for, TITLE, PHONE_FILTERS
    DASHBOARDS, DRILL, LAYOUT, NAV, NOTES = pages.DASHBOARDS, pages.DRILL, pages.LAYOUT, pages.NAV, pages.NOTES
    SIDE_NOTE, VISIBILITY, filters_for, params_for = pages.SIDE_NOTE, pages.VISIBILITY, pages.filters_for, pages.params_for
    PHONE_FILTERS = getattr(pages, "PHONE_FILTERS", PHONE_FILTERS)
    TITLE = title


PHONE_W = 375
PHONE_FILTERS = ["Region", "Branch", "ASM", "Partner", "Policy category", "Partner growth band"]


def dashboard(R, S, dash):
    params = params_for(dash)
    filters = filters_for(dash)
    cx, cw = SIDE_W, W - SIDE_W
    rows = fit_rows(S, LAYOUT[dash], cw)
    on_page = [name for _, cells in rows for c in cells for name in cell_names(c)]
    side_min = 30 + (len(params) + len(filters)) * CTRL_H + 220
    body_h = max(sum(h for h, _ in rows), side_min)
    H = HEAD_H + NAV_H + body_h + 8
    L = Layout(H)
    used = set(on_page)
    i = DASHBOARDS.index(dash)
    ids = {}            # zone ids by item, shared with the phone layout

    # ---- header
    head = (run(f"{TITLE}  ·  ", "#cfe9d7", 12) + run(f"{i + 1} / {len(DASHBOARDS)}   ", "#cfe9d7", 10)
            + run(dash, "#ffffff", 18, True) + run("      " + NOTES[dash], "#cfe9d7", 10))
    head_xml = f"<formatted-text>{head}</formatted-text>"
    header = L.box(0, 0, W, HEAD_H, head_xml, style=(C_DARK, 0, 10), fixed=HEAD_H, type_v2="text")
    ids["header"] = L.last

    # ---- navigation: chapters, then back / next
    nav_items = []
    bn = 92
    extra = [n for n, ok in (("Nav | ‹ Back", i > 0), ("Nav | Next ›", i < len(DASHBOARDS) - 1)) if ok]
    n = len(NAV)
    bw = (W - len(extra) * bn) / n
    x = 0
    nav_names = []
    for label, target in NAV:
        name = f"Nav | {label}"
        nav_items.append(L.box(x, HEAD_H, bw, NAV_H, style=(C_DARK if target == dash else C_MID, 3, 0), name=name, show_title="false"))
        ids[name] = L.last
        nav_names.append((name, C_DARK if target == dash else C_MID))
        used.add(name)
        x += bw
    for name in extra:
        nav_items.append(L.box(x, HEAD_H, bn, NAV_H, style=(C_TARGET, 3, 0), fixed=bn, name=name, show_title="false"))
        ids[name] = L.last
        nav_names.append((name, C_TARGET))
        used.add(name)
        x += bn
    nav = L.flow(0, HEAD_H, W, NAV_H, "horz", nav_items, fixed=NAV_H, style=(C_BG, 0, 0))

    # ---- filter panel
    y0 = HEAD_H + NAV_H
    side = [L.box(0, y0, SIDE_W, 30, "<formatted-text>" + run("Filters & settings", C_DARK, 10, True) + "</formatted-text>",
                  style=(None, 4, 4), fixed=30, type_v2="text")]
    y = y0 + 30
    controls = []       # (key, zone attributes) for the phone layout
    for pname, mode in params:
        a = dict(mode=mode, param=f"[Parameters].[{pname}]", type_v2="paramctrl")
        side.append(L.box(0, y, SIDE_W, CTRL_H, style=("#ffffff", 3, 0), fixed=CTRL_H, **a))
        controls.append((pname, L.last, a))
        y += CTRL_H
    filter_deps = []
    for field, _, _ in filters:
        src = next((s for s in on_page if field in getattr(S.get(s), "filter_refs", {})), None)
        if src is None:
            continue
        ref = S[src].filter_refs[field]
        a = dict(mode="checkdropdown", name=src, param=ref, show_apply="true", type_v2="filter", values="relevant")
        side.append(L.box(0, y, SIDE_W, CTRL_H, style=("#ffffff", 3, 0), fixed=CTRL_H, **a))
        controls.append((field, L.last, a))
        filter_deps.append(field)
        y += CTRL_H
    note = SIDE_NOTE if dash != "Pavilion" else (
        "Slab and qualification tables are as published; the filters do not apply to them.\n" + SIDE_NOTE)
    side.append(L.box(0, y, SIDE_W, y0 + body_h - y, "<formatted-text>" + run(note, C_MUTED, 8) + "</formatted-text>",
                      style=(None, 6, 4), type_v2="text"))
    sidebar = L.flow(0, y0, SIDE_W, body_h, "vert", side, fixed=SIDE_W, style=(SIDE_BG, 0, 0))

    # ---- content: one horizontal container per row
    content = []
    swaps = []
    y = y0
    for h, cells in rows:
        total = sum(c[1] for c in cells)
        x = cx
        kids = []
        for c in cells:
            w = cw * c[1] / total
            if c[0] == "stack":
                yy, sub = y, []
                for name, hh in c[2]:
                    sub.append(sheet_zone(L, S, name, x, yy, w, hh))
                    ids[name] = L.last
                    yy += hh
                kids.append(L.flow(x, y, w, h, "vert", sub))
            elif c[0] == "swap":
                # one sheet shows at a time: each zone's visibility is controlled by a boolean field of the
                # view parameter (Control visibility using value), so the hidden sheets and their titles collapse
                sub = []
                for k, name in enumerate(c[2]):
                    sub.append(sheet_zone(L, S, name, x, y, w, h, title=True, hidden=k > 0))
                    ids[name] = L.last
                    swaps.append((L.last, VISIBILITY[name]))
                kids.append(L.flow(x, y, w, h, "vert", sub))
            else:
                kids.append(sheet_zone(L, S, c[0], x, y, w, h))
                ids[c[0]] = L.last
            x += w
        content.append(L.flow(cx, y, cw, h, "horz", kids, fixed=h))
        y += h
    if y < y0 + body_h:
        content.append(L.box(cx, y, cw, y0 + body_h - y, type_v2="empty"))
    body = L.flow(cx, y0, cw, body_h, "vert", content)
    main = L.flow(0, y0, W, body_h, "horz", [sidebar, body])
    page = L.flow(0, 0, W, H, "vert", [header, nav, main])

    phone = phone_layout(S, dash, rows, ids, head_xml, nav_names, controls, L.n)

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
    sid = new_uuid()
    xml = (f"<dashboard name={attr(dash)}>{style}<size maxheight='{H}' maxwidth='{W}' minheight='{H}' minwidth='{W}' />"
           f"{deps}<zones>{root}</zones>{phone}<simple-id uuid='{sid}' /></dashboard>")
    return X(xml), used, [(sid, zid, field) for zid, field in swaps]


def phone_layout(S, dash, rows, ids, head_xml, nav_names, controls, last_id):
    """Phone layout (its own layout of the same dashboard; the desktop layout is not changed): one column with the
    header, the chapter buttons three to a row, the settings and main filters, then every sheet in page order
    (KPI cards two to a row). Items keep their desktop zone ids, as Tableau does."""
    PW = PHONE_W
    blocks = []         # (height, builder(L, y) -> xml)

    def one(name, h, hidden=False):
        a = {"hidden_by_user": "true"} if hidden else {}
        if name.startswith("H | "):
            return lambda L, y: L.box(0, y, PW, h, style=("#eaf4ee", 4, 6), fixed=h, zid=ids[name], name=name, show_title="false")
        return lambda L, y: L.box(0, y, PW, h, style=("#ffffff", 4, 4), fixed=h, zid=ids[name], name=name, show_title="true", **a)

    blocks.append((64, lambda L, y: L.box(0, y, PW, 64, head_xml, style=(C_DARK, 0, 8), fixed=64, zid=ids["header"], type_v2="text")))
    for k in range(0, len(nav_names), 3):
        trio = nav_names[k:k + 3]

        def navrow(L, y, trio=trio):
            kids = [L.box(j * PW / 3, y, PW / 3, 32, style=(colour, 2, 0), zid=ids[name], name=name, show_title="false")
                    for j, (name, colour) in enumerate(trio)]
            return L.flow(0, y, PW, 32, "horz", kids, fixed=32, style=(C_BG, 0, 0))
        blocks.append((32, navrow))
    for key, zid, a in controls:
        if a["type_v2"] == "filter" and dash != "Pavilion" and key not in PHONE_FILTERS:
            continue
        blocks.append((54, lambda L, y, zid=zid, a=a: L.box(0, y, PW, 54, style=("#ffffff", 3, 0), fixed=54, zid=zid, **a)))
    for h, cells in rows:
        names = [n for c in cells for n in cell_names(c)]
        if all(n.startswith("K | ") for n in names):
            for k in range(0, len(names), 2):
                pair = names[k:k + 2]

                def kpis(L, y, pair=pair):
                    kids = [L.box(j * PW / 2, y, PW / 2, 112, style=("#ffffff", 4, 4), zid=ids[n], name=n, show_title="true")
                            for j, n in enumerate(pair)]
                    return L.flow(0, y, PW, 112, "horz", kids, fixed=112)
                blocks.append((112, kpis))
            continue
        for c in cells:
            if c[0] == "swap":
                hh = min(h, 640)

                def swap(L, y, hh=hh, names=c[2]):
                    kids = [one(n, hh, hidden=j > 0)(L, y) for j, n in enumerate(names)]
                    return L.flow(0, y, PW, hh, "vert", kids, fixed=hh)
                blocks.append((hh, swap))
                continue
            for name in cell_names(c):
                s = S[name]
                if name.startswith("H | "):
                    hh = 150
                else:
                    need = need_height(s, PW, fit=False) if s.kind == "table" and s.est_rows is not None else None
                    hh = need if need is not None else min(h, 560 if s.kind == "table" else 320)
                blocks.append((hh, one(name, hh)))
    PH = sum(h for h, _ in blocks) + 8
    L = Layout(PH, width=PW, start=last_id)
    y, kids = 0, []
    for h, build in blocks:
        kids.append(build(L, y))
        y += h
    col = L.flow(0, 0, PW, PH, "vert", kids)
    root = (f"<zone h='100000' id='{L._id()}' type-v2='layout-basic' w='100000' x='0' y='0'>{col}"
            f"<zone-style><format attr='background-color' value='{C_BG}' /></zone-style></zone>")
    return (f"<devicelayouts><devicelayout name='Phone'><size maxheight='{PH}' minheight='{PH}' sizing-mode='vscroll' />"
            f"<zones>{root}</zones></devicelayout></devicelayouts>")


def fit_columns(s, w):
    """Tables with measures across: size the measure columns so the table fills its box (no squeezed #### columns,
    no empty space on the right). Wide tables keep a 60 px minimum and scroll sideways."""
    if s.kind != "table" or not s.cols or not s.cols.endswith("[:Measure Names]"):
        return
    dims = sum(wd for ref, wd in s.widths if not ref.endswith("[:Measure Names]"))
    n = getattr(s, "n_measures", 0)
    if not n:
        return
    col = max(60, int((w - 8 - 8 - 18 - dims) / n))
    s.widths = [(ref, wd) for ref, wd in s.widths if not ref.endswith("[:Measure Names]")] + [(s.mn, col)]
    s.fit = None


def sheet_zone(L, S, name, x, y, w, h, title=True, hidden=False):
    fit_columns(S[name], w)
    if name.startswith("H | "):
        return L.box(x, y, w, h, style=("#eaf4ee", 4, 6), name=name, show_title="false")
    extra = {"hidden_by_user": "true"} if hidden else {}
    return L.box(x, y, w, h, style=("#ffffff", 4, 4), name=name, show_title="true" if title else "false", **extra)


def datagraph_xml(R, zones):
    """Dynamic zone visibility: each (dashboard uuid, zone id, boolean field) shows the zone while the field is true."""
    g = lambda: str(uuid.uuid4())
    sub = g()
    nodes, edges, pairs = [], [], []
    for sid, zid, field in zones:
        fnode, znode, out, inp = g(), g(), g(), g()
        nodes.append(f"<single-value-field-node fieldname={attr(f'[{DS}].[{R[field].name}]')} fieldname-input-guid='{g()}' "
                     f"node-guid='{fnode}' value-output-guid='{out}' />")
        nodes.append(f"<dashboard-zone-visibility-node dashboard-identifier='{sid}' node-guid='{znode}' "
                     f"visibility-input-guid='{inp}' zone-id='{zid}' />")
        pairs += [f"<pair execution-subgraph-guid='{sub}' node-guid='{n}' />" for n in (fnode, znode)]
        edges.append(f"<edge from='{out}' to='{inp}' />")
    return X(f"<datagraph><graph><properties><default-execution-subgraph-guid value='{sub}' /></properties>"
             f"<node-execution-subgraphs>{''.join(pairs)}</node-execution-subgraphs><nodes>{''.join(nodes)}</nodes>"
             f"<edges>{''.join(edges)}</edges><pin-values /></graph></datagraph>")


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

    S = build_sheets(R)
    bad = missing_fields(R, {DS: retail, PDS: pav})
    if bad:
        gone = drop_missing(R, S, bad)
        print(f"tables missing from the data source: {len(bad)} fields; sheets left out: {', '.join(gone)}")
        if any(n.startswith("V | ") for n in gone):
            NOTES["Pavilion"] = "Pavilion slabs, qualification and app acceptance (as published)."

    clean_datasource(retail, R, DS, aliases={f"usr:{R[k].name}:qk": v for k, v in NP_ALIASES.items()})
    retail.insert(list(retail).index(retail.find("layout")) + 1, X(style_xml(R)))
    clean_datasource(pav, R, PDS, aliases=PAV_ALIASES)
    for d in list(ds_list):
        ds_list.remove(d)
    ds_list.append(params_datasource(R))
    ds_list.append(retail)
    ds_list.append(pav)
    for tag in ("repository-location", "actions", "worksheets", "dashboards", "windows", "datagraph", "thumbnails"):
        # also the copies Tableau saves under a version-tagged name (e.g. '_.fcp.DatagraphCoreV1.true...datagraph')
        for e in [e for e in root if str(e.tag) == tag or str(e.tag).endswith("..." + tag)]:
            root.remove(e)
    ws = etree.SubElement(root, "worksheets")
    dbs = etree.SubElement(root, "dashboards")
    dash_used = {}
    vis = []
    for d in DASHBOARDS:
        el, used, z = dashboard(R, S, d)
        vis += z
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
    if vis:
        root.append(datagraph_xml(R, vis))
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
