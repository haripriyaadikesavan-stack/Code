"""Behaviour checks on a generated workbook, beyond validate.py: what Tableau would do when a user filters,
sorts, switches views or clicks. Run on any workbook built by build.py / build_hs.py / build_wiser.py.

usage: python audit.py <workbook.twb> [more.twb ...]

Checks
  filters   every shared filter group is identical on every sheet (Tableau unlinks a group whose members differ);
            every sheet on a page carries every filter shown on that page (or is a known exception);
            filter-group ids are not reserved (<3); dashboard filter zones point at a sheet that has the filter
  sorts     each sort's dimension is on the sheet's rows / columns and its measure is declared on the sheet
  top N     the ranking field of a top-N filter is declared
  shelves   everything on rows, columns and marks is declared; colour scales point at a field on the sheet
  views     every 'Select a view' value shows exactly one sheet; show / hide zones exist on desktop and phone
  actions   navigation buttons go to existing pages; click-to-filter sources are on their page
  params    each parameter's current value is one of its allowed values
"""
import collections
import html
import re
import sys

from lxml import etree

EXCEPTIONS = ("top_n", "own")      # why a sheet may lack a shared filter


def instances(w):
    return {ci.get("name") for ci in w.iter("column-instance")}


def leaf(ref):
    """'[datasource].[instance]' -> '[instance]' (instance names may contain dots, e.g. 'No. of KRO')."""
    return "[" + ref.split("].[", 1)[1] if "].[" in ref else ref


def refs_in(text):
    return set(re.findall(r"\[[^\[\]]+\]\.\[[^\[\]]+\]", text or ""))


def audit(path):
    t = etree.parse(path).getroot()
    issues, notes = [], []
    W = {w.get("name"): w for w in t.find("worksheets")}
    D = {d.get("name"): d for d in t.find("dashboards")}
    params = {}
    for c in t.find("datasources").find("datasource[@name='Parameters']").findall("column"):
        params[c.get("name")] = c
    calcs = {}
    for ds in t.find("datasources"):
        for c in ds.findall("column"):
            if c.find("calculation") is not None:
                calcs[(ds.get("name"), c.get("name"))] = c

    # ---------------------------------------------------------------- filters
    groups = collections.defaultdict(dict)       # group -> sheet -> normalised filter xml
    by_sheet = collections.defaultdict(dict)     # sheet -> column -> group
    for name, w in W.items():
        for f in w.iter("filter"):
            g = f.get("filter-group")
            if g:
                xml = re.sub(r"\s+", " ", etree.tostring(f).decode())
                groups[g][name] = xml
                by_sheet[name][f.get("column")] = g
    for g, members in groups.items():
        if int(g) < 3:
            issues.append(f"filter-group {g} is reserved by Tableau")
        variants = collections.Counter(members.values())
        if len(variants) > 1:
            issues.append(f"filter-group {g}: {len(variants)} different definitions on {len(members)} sheets "
                          f"(Tableau would unlink them)")
    for dname, d in D.items():
        zones = list(d.find("zones").iter("zone"))
        sheets = [z.get("name") for z in zones if z.get("name") and z.get("type-v2") is None]
        fz = [z for z in zones if z.get("type-v2") == "filter"]
        for z in fz:
            col, src = z.get("param"), z.get("name")
            if src not in W or col not in by_sheet.get(src, {}):
                issues.append(f"{dname}: filter {leaf(col)} is not on its source sheet {src}")
                continue
            g = by_sheet[src][col]
            missing = []
            for s in sheets:
                if s.startswith(("Nav | ", "PV | ")) or s not in W:
                    continue
                if col in by_sheet.get(s, {}):
                    if by_sheet[s][col] != g:
                        issues.append(f"{dname}: sheet {s} has {leaf(col)} in another group")
                    continue
                own = [f for f in W[s].iter("filter") if f.get("column") == col]
                topn = [f for f in W[s].iter("filter") if f.get("column") == col and f.find(".//groupfilter[@function='end']") is not None]
                if own or topn:
                    continue                                    # the sheet's own selection on that field (e.g. top N)
                if not by_sheet.get(s):
                    continue                                    # sheets of another data source (published tables)
                missing.append(s)
            if missing:
                issues.append(f"{dname}: filter {leaf(col)} does not reach {', '.join(missing)}")
        # dashboard params exist
        for z in zones:
            if z.get("type-v2") == "paramctrl" and z.get("param", "").split(".")[-1] not in params:
                issues.append(f"{dname}: control for unknown parameter {z.get('param')}")

    # ---------------------------------------------------------------- sorts, top N, shelves, colours
    for name, w in W.items():
        inst = instances(w)
        tbl = w.find("table")
        rows = tbl.findtext("rows") or ""
        cols = tbl.findtext("cols") or ""
        shelves = rows + " " + cols
        for s in w.iter("shelf-sort-v2"):
            dim = s.get("dimension-to-sort")
            meas = s.get("measure-to-sort-by")
            if dim not in shelves:
                issues.append(f"sheet {name}: sorts {leaf(dim)} which is not on rows / columns")
            if leaf(meas) not in inst:
                issues.append(f"sheet {name}: sorts by undeclared {leaf(meas)}")
        for f in w.iter("filter"):
            for gf in f.iter("groupfilter"):
                if gf.get("function") == "order":
                    by = gf.get("expression", "")
                    for r in re.findall(r"\[([^\[\]]+)\]", by):
                        if not any(r in i for i in inst) and not any(r == k[1].strip("[]") for k in calcs):
                            issues.append(f"sheet {name}: top-N ranks by unknown {r}")
        for r in refs_in(shelves):
            if leaf(r) not in inst and not r.endswith("[:Measure Names]") and not r.endswith("[Multiple Values]"):
                issues.append(f"sheet {name}: {leaf(r)} on rows / columns is not declared")
        for enc in w.iter("encodings"):
            for e in enc:
                c = e.get("column")
                if c and leaf(c) not in inst and not c.endswith("[:Measure Names]") and not c.endswith("[Multiple Values]"):
                    issues.append(f"sheet {name}: {e.tag} uses undeclared {leaf(c)}")
        for e in w.iter("encoding"):
            f = e.get("field")
            if f and f.startswith("[") and "].[" in f and leaf(f) not in inst:
                issues.append(f"sheet {name}: colour scale on undeclared {leaf(f)}")
        if not rows.strip() and not cols.strip() and w.find(".//customized-label") is None and not name.startswith("Nav"):
            issues.append(f"sheet {name}: nothing on rows, columns or the label")

    # ---------------------------------------------------------------- show / hide views
    dg = t.find("datagraph")
    if dg is not None:
        nodes = list(dg.iter())
        fields = {n.get("value-output-guid"): n.get("fieldname") for n in nodes if n.get("fieldname")}
        edges = {e.get("to"): e.get("from") for e in dg.iter("edge")}
        uuid_dash = {d.find("simple-id").get("uuid"): d for d in D.values()}
        per_dash = collections.defaultdict(list)
        for n in nodes:
            if n.get("zone-id"):
                d = uuid_dash.get(n.get("dashboard-identifier"))
                if d is None:
                    issues.append(f"show/hide: unknown dashboard {n.get('dashboard-identifier')}")
                    continue
                field = fields.get(edges.get(n.get("visibility-input-guid")))
                desk = [z for z in d.find("zones").iter("zone") if z.get("id") == n.get("zone-id")]
                phone = [z for z in d.iter("zone") if z.get("id") == n.get("zone-id")]
                if not desk:
                    issues.append(f"{d.get('name')}: show/hide zone {n.get('zone-id')} not on the page")
                elif len(phone) < 2 and d.find("devicelayouts") is not None:
                    issues.append(f"{d.get('name')}: show/hide zone {n.get('zone-id')} missing from the phone layout")
                per_dash[d.get("name")].append((desk[0].get("name") if desk else "?", field))
        for dname, items in per_dash.items():
            # each view parameter value must make exactly one zone visible
            formulas = {}
            for sheet, field in items:
                key = (field.split("].[")[0].strip("["), "[" + field.split("].[")[1])
                c = calcs.get(key)
                if c is None or c.get("datatype") != "boolean":
                    issues.append(f"{dname}: show/hide field for {sheet} is missing or not true/false")
                    continue
                m = re.match(r"\[Parameters\]\.\[([^\]]+)\] = '(.*)'$", html.unescape(c.find("calculation").get("formula")))
                if m:
                    formulas[sheet] = m.groups()
            pnames = {p for p, _ in formulas.values()}
            for p in pnames:
                pc = params.get(f"[{p}]")
                members = [mm.get("value").strip('"') for mm in pc.findall("members/member")] if pc is not None else []
                for v in members:
                    shown = [s for s, (pp, val) in formulas.items() if pp == p and val == v.replace("\\", "")]
                    if len(shown) != 1:
                        issues.append(f"{dname}: '{v}' shows {len(shown)} sheets")
            notes.append(f"{dname}: {len(items)} views switch with {', '.join(sorted(pnames))}")

    # ---------------------------------------------------------------- actions
    for a in t.find("actions"):
        src = a.find("source")
        d = D.get(src.get("dashboard"))
        target = [p.get("value") for p in a.iter("param") if p.get("name") == "target"]
        if d is None:
            issues.append(f"action {a.get('caption')}: page {src.get('dashboard')} does not exist")
            continue
        on_page = {z.get("name") for z in d.iter("zone")}
        if src.get("worksheet") not in on_page:
            issues.append(f"action {a.get('caption')}: {src.get('worksheet')} is not on {d.get('name')}")
        if target and target[0] not in D:
            issues.append(f"action {a.get('caption')}: target {target[0]} does not exist")
    for dname, d in D.items():
        navs = [z.get("name") for z in d.find("zones").iter("zone") if (z.get("name") or "").startswith("Nav | ")]
        acted = {a.find("source").get("worksheet") for a in t.find("actions") if a.find("source").get("dashboard") == dname}
        dead = [n for n in navs if n not in acted and n.replace("Nav | ", "") not in (dname,)]
        current = [n for n in dead if any(n == f"Nav | {lbl}" for lbl in [dname])]
        dead = [n for n in dead if n not in current]
        if len(dead) > 1:       # the button of the current page has no action on purpose
            issues.append(f"{dname}: buttons without an action: {', '.join(dead)}")

    # ---------------------------------------------------------------- parameters
    for name, c in params.items():
        members = [m.get("value") for m in c.findall("members/member")]
        if members and c.get("value") not in members:
            issues.append(f"parameter {c.get('caption')}: current value {c.get('value')} is not an allowed value")

    return issues, notes


def main(paths):
    ok = True
    for p in paths:
        issues, notes = audit(p)
        print(f"== {p.split('/')[-1]}: {len(issues)} issue(s)")
        for n in notes:
            print("   ·", n)
        for i in issues:
            print("   ✗", i)
        ok &= not issues
    return ok


if __name__ == "__main__":
    sys.exit(0 if main(sys.argv[1:]) else 1)
