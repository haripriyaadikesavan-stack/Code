"""Static checks on the generated .twb (no Tableau available here)."""
import html
import re
import sys

from lxml import etree

AGG = re.compile(r"\b(SUM|COUNTD|COUNT|AVG|MIN|MAX|ATTR)\s*\(", re.I)
ORDER = ["document-format-change-manifest", "repository-location", "preferences", "style-theme", "style",
         "local-data", "datasources", "datasource-relationships", "mapsources", "shared-views", "actions",
         "worksheets", "dashboards", "windows", "datagraph", "thumbnails", "external"]
# metadata parent-name (physical table / union) -> logical table
LOGICAL = {"[Year+]": "P", "[Partner Master Data]": "P", "[Geography mapping]": "P", "[Sheet1]": "P", "[SSA Mapping]": "P",
           "[Monthly Targets]": "T", "[QTOD - ATOD]": "Q", "[Branch Targets1]": "B", "[SSA Targets]": "S",
           "[Partner Targets]": "PT", "[Pivot]": "PT", "[Pavilion Sales]": "V"}


def strip_lod(f):
    out, depth = [], 0
    for ch in f:
        if ch == "{":
            depth += 1
            out.append("LOD")
        elif ch == "}":
            depth -= 1
        elif depth == 0:
            out.append(ch)
    return "".join(out)


def strip_strings(f):
    return re.sub(r"'[^']*'|\"[^\"]*\"", "''", f)


def refs_of(f):
    f2 = re.sub(r"\[Parameters\]\.\[[^\]]+\]", "", f)
    return set(re.findall(r"\[[^\[\]]+\]", strip_strings(f2)))


def params_of(f):
    return {"[" + m + "]" for m in re.findall(r"\[Parameters\]\.\[([^\]]+)\]", f)}


def main(path):
    root = etree.parse(path).getroot()
    errs = []
    ranks = [next((i for i, t in enumerate(ORDER) if str(c.tag) == t or str(c.tag).endswith("..." + t)), 99) for c in root]
    if ranks != sorted(ranks) or 99 in ranks:
        errs.append(f"top-level element order: {[c.tag for c in root]}")
    dsl = root.find("datasources")
    params = {c.get("name") for c in dsl.find("datasource[@name='Parameters']").findall("column")}
    info = {}
    for ds in dsl:
        if ds.get("name") == "Parameters":
            continue
        cols = {c.get("name"): c for c in ds.findall("column")}
        calcs = {n: c for n, c in cols.items() if c.find("calculation") is not None}
        ext = {html.unescape(k) for k in re.findall(r'<map key="(\[[^"]+\])"', etree.tostring(ds.find("extract")).decode())}
        conn = {html.unescape(k) for k in re.findall(r'<map key="(\[[^"]+\])"', etree.tostring(ds.find("connection").find("cols")).decode())}
        meta = {}
        for m in ds.find("connection").find("metadata-records"):
            if m.findtext("local-name"):
                meta[m.findtext("local-name")] = m.findtext("parent-name")
        info[ds.get("name")] = (cols, calcs, ext, conn, meta)

    def is_agg(calcs, n, seen=()):
        f = calcs[n].find("calculation").get("formula")
        if AGG.search(strip_lod(strip_strings(f))):
            return True
        return any(is_agg(calcs, r, seen + (n,)) for r in refs_of(f) if r in calcs and r not in seen)

    for dsname, (cols, calcs, ext, conn, meta) in info.items():
        def tables(n, seen=None):
            seen = seen or set()
            if n in seen:
                return set()
            seen.add(n)
            if n in calcs:
                t = set()
                f = calcs[n].find("calculation").get("formula")
                m = re.match(r"\s*\{FIXED ([^:]*):", f)
                if m:      # a FIXED LOD lives at the level of its dimensions
                    f = m.group(1)
                for r in refs_of(f):
                    t |= tables(r, seen)
                return t
            tb = meta.get(n)
            return {LOGICAL.get(tb, tb)} if tb else set()
        names = {}
        for n, c in cols.items():
            dn = c.get("caption") or n.strip("[]")
            if n != "[:Measure Names]" and not n.startswith("[__tableau"):
                if dn in names:
                    errs.append(f"{dsname[:12]}: duplicate field name '{dn}' ({names[dn]} / {n})")
                names[dn] = n
        for n, c in calcs.items():
            f = c.find("calculation").get("formula")
            cap = c.get("caption")
            for r in refs_of(f):
                if r not in cols and r not in conn:
                    errs.append(f"{cap}: unknown field {r}")
                elif r not in calcs and r not in ext:
                    errs.append(f"{cap}: field {r} is not in the extract")
            for p in params_of(f):
                if p not in params:
                    errs.append(f"{cap}: unknown parameter {p}")
            # a FIXED LOD may combine related tables (as the old Partner Achievement Grouping did)
            if len(tables(n)) > 1 and not is_agg(calcs, n) and not f.lstrip().startswith("{FIXED"):
                errs.append(f"{cap}: row-level calc mixes tables {tables(n)}")
            outside = strip_lod(strip_strings(f))
            if "{" in f and AGG.search(outside) and re.search(r"(?<![A-Z(])LOD", outside) and not re.search(r"(MIN|MAX|SUM|AVG|ATTR)\(\s*LOD", outside):
                errs.append(f"{cap}: aggregate mixed with a bare LOD")
            if f.count("(") != f.count(")"):
                errs.append(f"{cap}: unbalanced parentheses")
            s = strip_strings(f)
            if len(re.findall(r"\bIF\b", s)) + len(re.findall(r"\bCASE\b", s)) != len(re.findall(r"\bEND\b", s)):
                errs.append(f"{cap}: IF/CASE/END mismatch")

    sheets = {w.get("name"): w for w in root.find("worksheets")}
    for name, w in sheets.items():
        s = etree.tostring(w).decode()
        for dsname, (cols, calcs, *_ ) in info.items():
            declared = set(re.findall(r"<column-instance [^>]*name=\"(\[[^\"]+\])\"", s))
            used = set(re.findall(r"\[" + re.escape(dsname) + r"\]\.(\[[^\]]+\])", s))
            for u in used - {"[:Measure Names]", "[Multiple Values]"}:
                if u not in declared:
                    errs.append(f"sheet {name}: instance {u} not declared")
        dep_cols = set(re.findall(r"<column [^>]*name=\"(\[[^\"]+\])\"", s))
        for ci in re.findall(r"<column-instance column=\"(\[[^\"]+\])\"", s):
            if ci not in dep_cols:
                errs.append(f"sheet {name}: column {ci} missing from dependencies")
        for p in re.findall(r"\[Parameters\]\.(\[[^\]]+\])", s):
            if p not in params:
                errs.append(f"sheet {name}: unknown parameter {p}")
        for dd in w.iter("datasource-dependencies"):
            dsname = dd.get("datasource")
            if dsname == "Parameters":
                continue
            cols, calcs = info[dsname][0], info[dsname][1]
            dcols = {c.get("name") for c in dd.findall("column")}
            for c in dd.findall("column"):
                calc = c.find("calculation")
                if calc is not None:
                    for r in refs_of(calc.get("formula")):
                        if r not in dcols:
                            errs.append(f"sheet {name}: dependency {c.get('caption')} needs {r}")
            for ci in dd.findall("column-instance"):
                n, der = ci.get("column"), ci.get("derivation")
                if n in calcs:
                    agg = is_agg(calcs, n)
                    if der == "User" and not agg:
                        errs.append(f"sheet {name}: usr instance on row-level calc {calcs[n].get('caption')}")
                    if der in ("None", "Sum") and agg:
                        errs.append(f"sheet {name}: {der} instance on aggregate calc {calcs[n].get('caption')}")
        # numbers that do not fit show as ####; text fields (":nk]") are cut off instead
        texts = sum(1 for e in w.iter("text") if e.get("column") and not e.get("column").endswith(":nk]"))
        if texts > 2 and w.find(".//mark").get("class") == "Text" and "Multiple Values" not in s:
            errs.append(f"sheet {name}: {texts} fields on one text mark (Tableau shows ####)")

    dnames = {d.get("name") for d in root.find("dashboards")}
    for d in root.find("dashboards"):
        for z in d.iter("zone"):
            if z.get("name") and z.get("name") not in sheets:
                errs.append(f"dashboard {d.get('name')}: unknown sheet {z.get('name')}")
            if z.get("param") and "[" + z.get("param").split("].[")[1] not in params:
                errs.append(f"dashboard {d.get('name')}: unknown param {z.get('param')}")
            if int(z.get("y")) + int(z.get("h")) > 100000 or int(z.get("x")) + int(z.get("w")) > 100000:
                errs.append(f"dashboard {d.get('name')}: zone {z.get('name') or z.get('id')} overflows")
    for a in root.find("actions"):
        src = a.find("source")
        if src.get("dashboard") not in dnames or src.get("worksheet") not in sheets:
            errs.append(f"action {a.get('caption')}: bad source")
        tgt = [p.get("value") for p in a.iter("param") if p.get("name") == "target"][0]
        if tgt not in dnames:
            errs.append(f"action {a.get('caption')}: bad target {tgt}")
    win = {w.get("name") for w in root.find("windows")}
    for n in list(sheets) + list(dnames):
        if n not in win:
            errs.append(f"no window for {n}")
    for dsname, (cols, *_ ) in info.items():
        dss = etree.tostring([d for d in dsl if d.get("name") == dsname][0]).decode()
        for ref in set(re.findall(r"\[(?:none|sum|usr|attr):([^:\]]+):", dss)):
            if "[" + html.unescape(ref) + "]" not in cols:
                errs.append(f"{dsname[:12]} references missing field [{ref}]")

    print(f"{sum(len(v[1]) for v in info.values())} calcs, {len(sheets)} sheets, {len(dnames)} dashboards")
    for e in errs[:80]:
        print("ERR ", e)
    print(f"{len(errs)} error(s)")
    return not errs


if __name__ == "__main__":
    sys.exit(0 if main(sys.argv[1]) else 1)
