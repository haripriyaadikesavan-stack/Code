"""Apply QC fixes to the user's own 'Overall Retail Dashboard' workbook.

usage: python fix.py <Overall Retail Dashboard.twbx> <output.twbx>
Writes a log of every change to fix_log.txt.
"""
import copy
import os
import re
import shutil
import sys
import tempfile
import zipfile
from collections import Counter

from lxml import etree

DS = "federated.14wc2mc1wc09dr1b4kqf30rnj8dg"
P6 = "[Parameter 6]"
GN = "[Gross/Net (copy)_1650287831742312449]"
GNP = f"[Parameters].{GN}"
LOG = []


def log(section, msg):
    LOG.append((section, msg))


# ---------------------------------------------------------------- text-level fixes
RENAMES = {   # typos in sheet / dashboard names (also fixes the titles shown on dashboards)
    "Branch Dashboard -GAU": "Branch Dashboard - GAU",
    "Branch Dashboard -JKD": "Branch Dashboard - JKD",
    "Branch Dashboard -TEL": "Branch Dashboard - TEL",
    "Dashboard 35": "New Partners Month on Month",
    "Min Slab Achivement": "Min Slab Achievement",
    "Total Slab Achivement": "Total Slab Achievement",
    "Zone wise Achievment": "Zone wise Achievement",
    "Tier Catergory wise Partners": "Tier Category wise Partners",
    "Partner wise Target vs Ach (WIser)": "Partner wise Target vs Ach (Wiser)",
}


def text_fixes(s):
    # 1. one Gross/Net parameter for the whole workbook
    n6 = s.count(P6)
    s = s.replace("[Parameters]." + P6, GNP)
    s = re.sub(r"(&quot;|\"|')Gross Sales/OB(&quot;|\"|')", r"\1Gross Sales\2", s)
    s = re.sub(r"(&quot;|\"|')Net Sales/OB(&quot;|\"|')", r"\1Net Sales\2", s)
    log("Gross / Net", f"Merged the two Gross/Net parameters: {n6} references to 'Gross/Net' ([Parameter 6], values "
        "'Gross Sales/OB' / 'Net Sales/OB') now use 'Net/Gross' ('Gross Sales' / 'Net Sales'), so one switch drives every sheet.")
    # 2. renames
    for old, new in RENAMES.items():
        n = len(re.findall(re.escape(old) + r"(?![\w])", s))
        if n:
            s = re.sub(re.escape(old) + r"(?![\w])", new, s)
            log("Names", f"Renamed '{old}' -> '{new}' ({n} references)")
    return s


# ---------------------------------------------------------------- formula fixes (applied to every copy of a calc)
def set_formula(root, name, fn, why):
    n = 0
    for c in root.iter("column"):
        if c.get("name") == name:
            cl = c.find("calculation")
            if cl is not None and cl.get("formula"):
                new = fn(cl.get("formula"))
                if new != cl.get("formula"):
                    cl.set("formula", new)
                    n += 1
    if n:
        log("Calculations", why)
    return n


def calc_names(ds):
    return {(c.get("caption") or c.get("name")): c.get("name") for c in ds.findall("column") if c.find("calculation") is not None}


def formula_fixes(root, ds):
    names = calc_names(ds)
    net = f"{GNP} = 'Net Sales'"
    # FY Target 108% ignored Gross/Net (partner targets are gross)
    set_formula(root, names["FY Target 108%"],
                lambda f: re.sub(r"THEN\s*\[Target 108%\]\s*END",
                                 f"THEN IF {net} THEN [Target 108%] * 0.83 ELSE [Target 108%] END END", f),
                "'FY Target 108%' now converts with Gross/Net like 'FY Target 100%' (partner targets are gross; net = gross x 0.83).")
    # Curr. yr. Partner Target ignored Gross/Net
    set_formula(root, names["Curr. yr. Partner Target"],
                lambda f: re.sub(r"THEN\s*\[Pivot Field Values\]\s*END",
                                 f"THEN IF {net} THEN [Pivot Field Values] * 0.83 ELSE [Pivot Field Values] END END", f),
                "'Curr. yr. Partner Target' now converts with Gross/Net (partner targets are gross; net = gross x 0.83).")
    # Partner Ach% used gross sales whatever the switch -> follow Gross/Net like its target
    set_formula(root, names["Partner Ach%"],
                lambda f: f.replace(names["Curr. yr. Primary Sales (gross)"], names["FY Sales YTD"]),
                "'Partner Ach%' now divides Gross or Net sales (per the switch) by the matching partner target; it always used gross sales before.")
    # Q1 Ach% (Zeta) used Q2 sales
    set_formula(root, names["Q1 Ach% (Zeta)"],
                lambda f: f.replace(names["Qtr 2 Sales (Zeta)"], names["Qtr 1 Sales (Zeta)"]),
                "'Q1 Ach% (Zeta)' divided Q2 Zeta sales by the target (copy error); it now uses Q1 Zeta sales.")
    # monthly / quarterly sales: Net was Gross x 0.83; use the Net Sales column (rebate by year) like FY Sales YTD
    pat = re.compile(r"IF\s+" + re.escape(GNP) + r"\s*=\s*'Net Sales'\s+THEN\s+\[Gross Sales\]\s*\*\s*0\.83\s+ELSE\s+\[Gross Sales\]\s+END")
    fixed = []
    for c in ds.findall("column"):
        cl = c.find("calculation")
        if cl is not None and cl.get("formula") and pat.search(cl.get("formula")):
            fixed.append(c.get("caption") or c.get("name"))
            set_formula(root, c.get("name"), lambda f: pat.sub(f"IF {net} THEN [Net Sales] ELSE [Gross Sales] END", f), "")
    if fixed:
        LOG.pop() if LOG and LOG[-1][1] == "" else None
        LOG[:] = [x for x in LOG if x[1]]
        log("Calculations", f"{len(fixed)} monthly / quarterly sales fields ({', '.join(sorted(fixed)[:6])}, ...) computed Net as Gross x 0.83; "
            "they now use the Net Sales column (yearly rebate), the same as FY Sales YTD, so monthly and YTD net figures agree for every year.")


# ---------------------------------------------------------------- parameters
def merge_parameters(root):
    params = root.find("datasources").find("datasource[@name='Parameters']")
    gn = [c for c in params.findall("column") if c.get("name") == GN][0]
    gn.set("caption", "Gross / Net")
    p6 = [c for c in params.findall("column") if c.get("name") == P6]
    for c in p6:
        params.remove(c)
    # every Parameters dependency block: replace the old parameter's definition by the merged one, drop duplicates
    for dd in root.iter("datasource-dependencies"):
        if dd.get("datasource") != "Parameters":
            continue
        for c in list(dd.findall("column")):
            if c.get("name") == P6:
                has_gn = any(x.get("name") == GN for x in dd.findall("column"))
                if has_gn:
                    dd.remove(c)
                else:
                    new = copy.deepcopy(gn)
                    if c.find("members") is None and new.find("members") is not None:
                        new.remove(new.find("members"))
                    dd.replace(c, new)
        seen = set()
        for c in list(dd.findall("column")):
            if c.get("name") in seen:
                dd.remove(c)
            seen.add(c.get("name"))


# ---------------------------------------------------------------- optimisation
def remove_unused(root, ds):
    sheets = {w.get("name"): w for w in root.find("worksheets")}
    on_dash = {z.get("name") for d in root.find("dashboards") for z in d.iter("zone") if z.get("name")}
    text = etree.tostring(root).decode()
    removed = []
    for name, w in sheets.items():
        if name in on_dash:
            continue
        q = name.replace("&", "&amp;")
        refs = len(re.findall(r'(?:name|worksheet|value|sheet)="' + re.escape(q) + '"', text))
        # its own worksheet + window (+ thumbnail) only
        others = refs - 1 - (1 if root.find(f"windows/window[@name='{name}']") is not None else 0) \
            - (1 if root.find(f"thumbnails/thumbnail[@name='{name}']") is not None else 0)
        if others <= 0:
            root.find("worksheets").remove(w)
            for tag in ("windows/window", "thumbnails/thumbnail"):
                for e in root.findall(f"{tag}[@name='{name}']"):
                    e.getparent().remove(e)
            removed.append(name)
    if removed:
        log("Optimisation", f"Removed {len(removed)} sheets not used on any dashboard: {', '.join(removed)}")
    # calculations referenced nowhere else
    text = etree.tostring(root).decode()
    dead = []
    for c in list(ds.findall("column")):
        if c.find("calculation") is None or c.find("calculation").get("class") != "tableau":
            continue
        key = c.get("name").strip("[]")
        n = len(re.findall(re.escape(key), text))
        if n <= 1:
            ds.remove(c)
            dead.append(c.get("caption") or c.get("name"))
            text = text.replace(key, "")
    if dead:
        log("Optimisation", f"Removed {len(dead)} calculated fields used nowhere: {', '.join(dead)}")


# ---------------------------------------------------------------- formatting
TITLE_COLOR = "#008360"


def format_titles(root):
    n = 0
    for w in root.find("worksheets"):
        t = w.find("layout-options/title/formatted-text")
        if t is None:
            continue
        for run in t.findall("run"):
            size = run.get("fontsize")
            if size in (None, "12") and (run.text or "").strip():
                before = (run.get("fontname"), run.get("fontsize"), run.get("fontcolor"), run.get("bold"))
                run.set("fontname", "Tableau Bold")
                run.set("fontsize", "12")
                run.set("fontcolor", TITLE_COLOR)
                run.set("bold", "true")
                if before != ("Tableau Bold", "12", TITLE_COLOR, "true"):
                    n += 1
    log("Formatting", f"Sheet titles now use one style (Tableau Bold 12, {TITLE_COLOR}); {n} title runs changed "
        "(they used 7 different colours: black, #4b966f, #008360, #006f00, #009500, #3a7455 and default).")


def align_sidebars(root):
    """Floating filter / parameter controls in the left panel: same left edge and width."""
    changed = []
    for d in root.find("dashboards"):
        size = d.find("size")
        W = int(size.get("maxwidth") or size.get("minwidth"))
        floating = [z for z in d.find("zones").findall("zone")[1:] if z.get("type-v2") in ("paramctrl", "filter")]
        side = [z for z in floating if int(z.get("x")) * W / 100000 < 175]
        if len(side) < 2:
            continue
        xs = Counter(int(z.get("x")) for z in side)
        x = min(int(z.get("x")) for z in side)
        w = max(int(z.get("w")) for z in side)
        if len(xs) > 1 or len({int(z.get("w")) for z in side}) > 1:
            for z in side:
                z.set("x", str(x))
                z.set("w", str(w))
            changed.append(d.get("name"))
    if changed:
        log("Layout", f"Aligned the filter / parameter controls in the left panel (same left edge and width) on {len(changed)} dashboards: "
            + ", ".join(changed))


# ---------------------------------------------------------------- main
def main(src, out):
    tmp = tempfile.mkdtemp()
    with zipfile.ZipFile(src) as z:
        z.extractall(tmp)
    twb = os.path.join(tmp, "Overall Retail Dashboard.twb")
    s = open(twb, encoding="utf-8").read()
    s = text_fixes(s)
    root = etree.fromstring(s.encode("utf-8"))
    ds = [d for d in root.find("datasources") if d.get("name") == DS][0]
    merge_parameters(root)
    formula_fixes(root, ds)
    remove_unused(root, ds)
    format_titles(root)
    align_sidebars(root)
    data = etree.tostring(root, xml_declaration=True, encoding="utf-8")
    data = data.replace(b"<?xml version='1.0' encoding='utf-8'?>", b"<?xml version='1.0' encoding='utf-8' ?>", 1)
    name = os.path.splitext(os.path.basename(out))[0]
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr(f"{name}.twb", data)
        for dp, _, files in os.walk(tmp):
            for f in files:
                p = os.path.join(dp, f)
                rel = os.path.relpath(p, tmp)
                if not rel.endswith(".twb"):
                    z.write(p, rel)
    open(out.replace(".twbx", ".twb"), "wb").write(data)
    shutil.rmtree(tmp)
    with open("fix_log.txt", "w") as f:
        for sec, msg in LOG:
            f.write(f"[{sec}] {msg}\n")
    print("\n".join(f"[{a}] {b[:200]}" for a, b in LOG))


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
