"""Add the analysis pages built by ../gen (RETAIL_MERGE=1) to the existing workbook.

The pages use the workbook's own Gross/Net, Currency, Start Date and End
Date parameters, so one set of controls drives every dashboard.
Only the calculations those pages need are added; nothing existing changes
except: two new parameters' worth of controls, and a "More analysis" button
on the Retail Overall Dashboard.
"""
import copy
import uuid

from lxml import etree

DS = "federated.14wc2mc1wc09dr1b4kqf30rnj8dg"
OLD_PARAMS = {"[Gross/Net (copy)_1650287831742312449]", "[Parameter 10]", "[Parameter 2]", "[Parameter 1]"}
GN = "[Gross/Net (copy)_1650287831742312449]"
RENAME = {"Executive": "Executive Summary", "Branches": "Branch Scorecard", "Products": "Product Insights",
          "Partners": "Partner Health", "Target vs Achievement": "Target Tracker",
          "Order Book": "Order Book & Backlog", "Explore": "Explore (Any Dimension)"}
# captions that already exist in the workbook's data source
CAPTION_RENAME = {"Region": "Sales Region", "Family": "Product Family"}
HOME = "Retail Overall Dashboard"
LOG = []


def uid():
    return "{" + str(uuid.uuid4()).upper() + "}"


def ds_of(root, name):
    return [d for d in root.find("datasources") if d.get("name") == name][0]


def strip_domain(col):
    c = copy.deepcopy(col)
    for tag in ("members", "range"):
        for e in c.findall(tag):
            c.remove(e)
    return c


def merge(v, g):
    vds, gds = ds_of(v, DS), ds_of(g, DS)
    vpar, gpar = ds_of(v, "Parameters"), ds_of(g, "Parameters")
    vpar_cols = {c.get("name"): c for c in vpar.findall("column")}
    vds_cols = {c.get("name"): c for c in vds.findall("column")}
    gds_cols = {c.get("name"): c for c in gds.findall("column")}

    sheets = list(g.find("worksheets"))
    dashes = list(g.find("dashboards"))

    # ---- caption clashes: rename everywhere in the generated content
    clash = {c.get("name"): CAPTION_RENAME[c.get("caption")] for c in gds_cols.values()
             if c.find("calculation") is not None and c.get("caption") in CAPTION_RENAME}
    for el in [gds] + sheets:
        for c in el.iter("column"):
            if c.get("name") in clash:
                c.set("caption", clash[c.get("name")])

    # ---- calculations / base columns the pages need
    needed = set()
    for w in sheets:
        for dd in w.iter("datasource-dependencies"):
            if dd.get("datasource") == DS:
                needed |= {c.get("name") for c in dd.findall("column")}
    calcs = [gds_cols[n] for n in gds_cols if n in needed and gds_cols[n].find("calculation") is not None]
    bases = [gds_cols[n] for n in gds_cols if n in needed and gds_cols[n].find("calculation") is None and n not in vds_cols]
    vdisp = {(c.get("caption") or c.get("name").strip("[]")) for c in vds.findall("column")}
    for c in calcs + bases:
        if c.get("name") in vds_cols:
            raise ValueError(f"field name clash {c.get('name')}")
        if (c.get("caption") or c.get("name").strip("[]")) in vdisp:
            raise ValueError(f"caption clash {c.get('caption')}")
    last_col = vds.findall("column")[-1]
    idx = list(vds).index(last_col) + 1
    for c in bases + calcs:
        vds.insert(idx, copy.deepcopy(c))
        idx += 1
    LOG.append(f"added {len(calcs)} calculated fields and {len(bases)} field declarations")

    # base-column copies in the sheets' dependencies: use the workbook's own definitions
    for w in sheets:
        for dd in w.iter("datasource-dependencies"):
            if dd.get("datasource") != DS:
                continue
            for c in dd.findall("column"):
                n = c.get("name")
                if n in vds_cols and c.find("calculation") is None and vds_cols[n].find("calculation") is None:
                    dd.replace(c, copy.deepcopy(vds_cols[n]))

    # ---- measure colours (this year green, last year grey, target amber)
    vstyle = vds.find("style")
    vmark = [r for r in vstyle.findall("style-rule") if r.get("element") == "mark"]
    genc = gds.find("style/style-rule/encoding")
    if genc is not None:
        if any(e.get("field") == "[:Measure Names]" for r in vmark for e in r.findall("encoding")):
            enc = [e for r in vmark for e in r.findall("encoding") if e.get("field") == "[:Measure Names]"][0]
            for m in genc.findall("map"):
                enc.append(copy.deepcopy(m))
        else:
            vmark[0].append(copy.deepcopy(genc))

    # ---- parameters: new ones go into the Parameters data source
    added = []
    for c in gpar.findall("column"):
        if c.get("name") not in vpar_cols and c.get("name") not in OLD_PARAMS:
            vpar.append(copy.deepcopy(c))
            added.append(c.get("caption"))
    LOG.append(f"added parameters: {', '.join(added)}")
    vpar_cols = {c.get("name"): c for c in vpar.findall("column")}

    def fix_param_deps(el):
        for dd in el.iter("datasource-dependencies"):
            if dd.get("datasource") != "Parameters":
                continue
            for c in dd.findall("column"):
                if c.get("name") in OLD_PARAMS:
                    dd.replace(c, strip_domain(vpar_cols[c.get("name")]))

    for el in sheets + dashes:
        fix_param_deps(el)
    # the data source's own Parameters dependency block
    vdd = [dd for dd in vds.findall("datasource-dependencies") if dd.get("datasource") == "Parameters"]
    gdd = [dd for dd in gds.findall("datasource-dependencies") if dd.get("datasource") == "Parameters"]
    if gdd:
        if not vdd:
            new = etree.Element("datasource-dependencies", datasource="Parameters")
            vds.insert(list(vds).index(vds.find("object-graph") if vds.find("object-graph") is not None else vds[-1]), new)
            vdd = [new]
        have = {c.get("name") for c in vdd[0].findall("column")}
        for c in gdd[0].findall("column"):
            if c.get("name") not in have:
                vdd[0].append(strip_domain(vpar_cols[c.get("name")]))

    # ---- rename dashboards
    for d in dashes:
        old = d.get("name")
        new = RENAME[old]
        d.set("name", new)
        for r in d.iter("run"):
            if r.text == old:
                r.text = new
    for a in g.find("actions"):
        src = a.find("source")
        src.set("dashboard", RENAME.get(src.get("dashboard"), src.get("dashboard")))
        for p in a.iter("param"):
            if p.get("name") == "target":
                p.set("value", RENAME.get(p.get("value"), p.get("value")))
        t = a.get("caption").replace("Go to ", "")
        a.set("caption", "Go to " + RENAME.get(t, t))
    for w in g.find("windows"):
        w.set("name", RENAME.get(w.get("name"), w.get("name")))
        if "maximized" in w.attrib:
            del w.attrib["maximized"]

    # ---- append
    for w in sheets:
        v.find("worksheets").append(w)
    for d in dashes:
        v.find("dashboards").append(d)
    for a in list(g.find("actions")):
        v.find("actions").append(a)
    vwin = v.find("windows")
    for w in list(g.find("windows")):
        vwin.append(w)
    LOG.append(f"added {len(dashes)} dashboards ({', '.join(d.get('name') for d in dashes)}) and {len(sheets)} sheets")

    add_entry_button(v)
    # one caption for the merged Gross/Net parameter in every copy
    for c in v.iter("column"):
        if c.get("name") == GN and c.get("param-domain-type"):
            c.set("caption", "Gross / Net")


def add_entry_button(v):
    """'More analysis' button on the Retail Overall Dashboard, below its filter panel."""
    ws = v.find("worksheets")
    src = [w for w in ws if w.get("name") == "Nav | Executive"][0]
    w = copy.deepcopy(src)
    name = "Nav | More analysis"
    w.set("name", name)
    for r in w.iter("run"):
        if r.text == "Executive":
            r.text = "More analysis  ›"
    w.find("simple-id").set("uuid", uid())
    ws.append(w)
    d = [d for d in v.find("dashboards") if d.get("name") == HOME][0]
    size = d.find("size")
    W, H = int(size.get("maxwidth")), int(size.get("maxheight"))
    top = d.find("zones").findall("zone")
    side = [z for z in top[1:] if int(z.get("x")) * W / 100000 < 200 and (int(z.get("x")) + int(z.get("w"))) * W / 100000 <= 240]
    bottom = max(int(z.get("y")) + int(z.get("h")) for z in side)
    x = min(int(z.get("x")) for z in side if z.get("type-v2") in ("paramctrl", "filter"))
    wd = max(int(z.get("w")) for z in side if z.get("type-v2") in ("paramctrl", "filter"))
    ids = [int(z.get("id")) for z in d.iter("zone") if z.get("id")]
    z = etree.SubElement(d.find("zones"), "zone")
    for k, val in (("h", str(round(34 * 100000 / H))), ("id", str(max(ids) + 1)), ("name", name), ("show-title", "false"),
                   ("w", str(wd)), ("x", str(x)), ("y", str(bottom + round(10 * 100000 / H)))):
        z.set(k, val)
    act = etree.SubElement(v.find("actions"), "action", caption="Go to Executive Summary",
                           name=f"[Action_{uuid.uuid4().hex.upper()}]")
    etree.SubElement(act, "activation", **{"auto-clear": "true", "type": "on-select"})
    etree.SubElement(act, "source", dashboard=HOME, type="sheet", worksheet=name)
    cmd = etree.SubElement(act, "command", command="tsc:tsl-filter")
    etree.SubElement(cmd, "param", name="special-fields", value="all")
    etree.SubElement(cmd, "param", name="target", value="Executive Summary")
    win = [w for w in v.find("windows") if w.get("name") == HOME][0]
    vps = win.find("viewpoints")
    ref = [vp for w in v.find("windows") for vp in w.iter("viewpoint") if vp.get("name") == "Nav | Executive"][0]
    vp = copy.deepcopy(ref)
    vp.set("name", name)
    vps.append(vp)
    hw = copy.deepcopy([x for x in v.find("windows") if x.get("name") == "Nav | Executive"][0])
    hw.set("name", name)
    hw.find("simple-id").set("uuid", uid())
    v.find("windows").append(hw)
    LOG.append("added a 'More analysis' button to the Retail Overall Dashboard filter panel (opens Executive Summary)")
