"""Static checks on the generated .twb (no Tableau available here)."""
import re
import sys

from lxml import etree

AGG = re.compile(r"\b(SUM|COUNTD|COUNT|AVG|MIN|MAX|ATTR)\s*\(", re.I)


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


def main(path):
    root = etree.parse(path).getroot()
    errs, warns = [], []
    dsl = root.find("datasources")
    params = {c.get("name") for c in dsl.find("datasource[@name='Parameters']").findall("column")}
    ds = [d for d in dsl if d.get("name") != "Parameters"][0]
    dsname = ds.get("name")
    cols = {c.get("name"): c for c in ds.findall("column")}
    calcs = {n: c for n, c in cols.items() if c.find("calculation") is not None}
    extract_keys = set(re.findall(r'<map key="(\[[^"]+\])"', etree.tostring(ds.find("extract")).decode()))
    conn_keys = set(re.findall(r'<map key="(\[[^"]+\])"', etree.tostring(ds.find("connection").find("cols")).decode()))
    table_of = {}
    for m in ds.find("connection").find("metadata-records"):
        table_of["[" + (m.findtext("local-name") or "").strip("[]") + "]"] = m.findtext("parent-name")

    # 1. formulas reference existing fields / params; row-level calcs stay on one logical table
    def fields_in(f):
        f2 = re.sub(r"\[Parameters\]\.\[([^\]]+)\]", lambda m: (params_used.add("[" + m.group(1) + "]") or ""), f)
        return set(re.findall(r"\[[^\[\]]+\]", strip_strings(f2)))

    def base_tables(name, seen=None):
        seen = seen or set()
        if name in seen:
            return set()
        seen.add(name)
        if name in calcs:
            t = set()
            for r in fields_in(calcs[name].find("calculation").get("formula")):
                t |= base_tables(r, seen)
            return t
        tb = table_of.get(name)
        return {LOGICAL.get(tb, tb)} if tb else set()

    for n, c in calcs.items():
        params_used = set()
        f = c.find("calculation").get("formula")
        refs = fields_in(f)
        for r in refs:
            if r not in cols and r not in conn_keys:
                errs.append(f"{c.get('caption')}: unknown field {r}")
            elif r not in calcs and r not in extract_keys:
                errs.append(f"{c.get('caption')}: field {r} is not in the extract")
        for p in params_used:
            if p not in params:
                errs.append(f"{c.get('caption')}: unknown parameter {p}")
        is_agg = bool(AGG.search(strip_lod(strip_strings(f)))) or any(
            AGG.search(strip_lod(strip_strings(calcs[r].find("calculation").get("formula")))) for r in refs if r in calcs and r != n)
        tabs = base_tables(n)
        if len(tabs) > 1 and not is_agg:
            errs.append(f"{c.get('caption')}: row-level calc mixes tables {tabs}")
        # mixing an LOD (row-level) with an aggregate at top level
        if "{" in f and AGG.search(strip_lod(strip_strings(f))):
            body = strip_strings(f)
            outside = strip_lod(body)
            if re.search(r"(?<![A-Z(])LOD", outside) and not re.search(r"(MIN|MAX|SUM|AVG|ATTR)\(\s*LOD", outside):
                errs.append(f"{c.get('caption')}: aggregate mixed with a bare LOD")
        if f.count("(") != f.count(")"):
            errs.append(f"{c.get('caption')}: unbalanced parentheses")
        if len(re.findall(r"\bIF\b", strip_strings(f))) + len(re.findall(r"\bCASE\b", strip_strings(f))) != len(re.findall(r"\bEND\b", strip_strings(f))):
            errs.append(f"{c.get('caption')}: IF/CASE/END mismatch")

    # 2. worksheets: every referenced instance is declared
    sheets = {w.get("name"): w for w in root.find("worksheets")}
    for name, w in sheets.items():
        s = etree.tostring(w).decode()
        declared = set(re.findall(r"<column-instance [^>]*name=\"(\[[^\"]+\])\"", s))
        dep_cols = set(re.findall(r"<column [^>]*name=\"(\[[^\"]+\])\"", s))
        used = set(re.findall(r"\[" + re.escape(dsname) + r"\]\.(\[[^\]]+\])", s))
        for u in used:
            if u in ("[:Measure Names]", "[Multiple Values]"):
                continue
            if u not in declared:
                errs.append(f"sheet {name}: instance {u} not declared")
        for ci in re.findall(r"<column-instance column=\"(\[[^\"]+\])\"", s):
            if ci not in dep_cols:
                errs.append(f"sheet {name}: column {ci} missing from dependencies")
        for p in re.findall(r"\[Parameters\]\.(\[[^\]]+\])", s):
            if p not in params:
                errs.append(f"sheet {name}: unknown parameter {p}")
        # dependency calcs reference only declared columns
        for c in w.iter("column"):
            calc = c.find("calculation")
            if calc is not None and c.get("param-domain-type") is None:
                params_used = set()
                for r in fields_in(calc.get("formula")):
                    if r not in dep_cols:
                        errs.append(f"sheet {name}: dependency {c.get('caption')} needs {r}")

    # 2b. derivation matches the calc kind
    def is_aggregate(n, seen=()):
        f = calcs[n].find("calculation").get("formula")
        if AGG.search(strip_lod(strip_strings(f))):
            return True
        return any(is_aggregate(r, seen + (n,)) for r in re.findall(r"\[[^\[\]]+\]", strip_strings(f)) if r in calcs and r not in seen)
    for name, w in sheets.items():
        s = etree.tostring(w).decode()
        for der, col in re.findall(r'derivation="([^"]+)" name="\[(?:none|usr|sum|attr):([^:]+):', s):
            n = "[" + col + "]"
            if n not in calcs:
                continue
            agg = is_aggregate(n)
            if der == "User" and not agg:
                errs.append(f"sheet {name}: usr instance on row-level calc {calcs[n].get('caption')}")
            if der in ("None", "Sum") and agg:
                errs.append(f"sheet {name}: {der} instance on aggregate calc {calcs[n].get('caption')}")

    # 3. dashboards and actions
    dnames = {d.get("name") for d in root.find("dashboards")}
    for d in root.find("dashboards"):
        for z in d.iter("zone"):
            if z.get("name") and z.get("name") not in sheets:
                errs.append(f"dashboard {d.get('name')}: unknown sheet {z.get('name')}")
            if z.get("param") and "[" + z.get("param").split("].[")[1] not in params:
                errs.append(f"dashboard {d.get('name')}: unknown param {z.get('param')}")
            for k in ("x", "y", "w", "h"):
                v = int(z.get(k))
                if v < 0 or v > 100000:
                    errs.append(f"dashboard {d.get('name')}: zone {z.get('id')} {k}={v}")
            if int(z.get("y")) + int(z.get("h")) > 100000 or int(z.get("x")) + int(z.get("w")) > 100000:
                errs.append(f"dashboard {d.get('name')}: zone {z.get('name') or z.get('id')} overflows")
    for a in root.find("actions"):
        src = a.find("source")
        if src.get("dashboard") not in dnames or src.get("worksheet") not in sheets:
            errs.append(f"action {a.get('caption')}: bad source")
        tgt = [p.get("value") for p in a.iter("param") if p.get("name") == "target"][0]
        if tgt not in dnames:
            errs.append(f"action {a.get('caption')}: bad target")
    win = {w.get("name") for w in root.find("windows")}
    for n in list(sheets) + list(dnames):
        if n not in win:
            errs.append(f"no window for {n}")

    # 4. leftovers from the old workbook
    dss = etree.tostring(ds).decode()
    defined = set(cols)
    for ref in set(re.findall(r"\[(?:none|sum|usr|attr):([^:\]]+):", dss)):
        if "[" + ref + "]" not in defined:
            warns.append(f"data source still references [{ref}]")
    for ref in set(re.findall(r"\[(Calculation_\d+|[^\]]*\(copy\)_\d+)\]", dss)):
        if "[" + ref + "]" not in defined:
            errs.append(f"data source references removed calc [{ref}]")

    print(f"{len(calcs)} calcs, {len(sheets)} sheets, {len(dnames)} dashboards")
    for w in warns[:30]:
        print("WARN", w)
    for e in errs[:80]:
        print("ERR ", e)
    print(f"{len(errs)} error(s), {len(warns)} warning(s)")
    return not errs


# metadata parent-name (physical table / union) -> logical table
LOGICAL = {"[Year+]": "P", "[Partner Master Data]": "P", "[Geography mapping]": "P", "[Sheet1]": "P", "[SSA Mapping]": "P",
           "[Monthly Targets]": "T", "[QTOD - ATOD]": "Q", "[Branch Targets1]": "B"}

if __name__ == "__main__":
    sys.exit(0 if main(sys.argv[1]) else 1)
