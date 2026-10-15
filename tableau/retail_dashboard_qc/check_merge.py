"""Consistency checks for the pages merged in by merge.py."""
import re, sys
from lxml import etree
from load import refs
DS = "federated.14wc2mc1wc09dr1b4kqf30rnj8dg"
r = etree.parse(sys.argv[1]).getroot()
ds = [d for d in r.find("datasources") if d.get("name") == DS][0]
import html
cols = {c.get("name"): c for c in ds.findall("column")}
conn = {html.unescape(k) for k in re.findall(r'<map key="(\[[^"]+\])"', etree.tostring(ds.find("connection")).decode())}
params = {c.get("name"): c for c in r.find("datasources").find("datasource[@name='Parameters']").findall("column")}
new = [d for d in r.find("dashboards") if d.find(".//zone[@name]") is not None and any((z.get("name") or "").startswith("Nav | ") for z in d.iter("zone")) and d.get("name") != "Retail Overall Dashboard"]
names = {z.get("name") for d in new for z in d.iter("zone") if z.get("name")}
errs = []
caps = {}
for c in ds.findall("column"):
    k = c.get("caption") or c.get("name").strip("[]")
    if k in caps and not c.get("name").startswith("[__"): errs.append(f"duplicate caption {k}")
    caps[k] = 1
for w in r.find("worksheets"):
    if w.get("name") not in names: continue
    s = etree.tostring(w).decode()
    for dd in w.iter("datasource-dependencies"):
        dcols = {c.get("name"): c for c in dd.findall("column")}
        if dd.get("datasource") == "Parameters":
            for n, c in dcols.items():
                if n not in params: errs.append(f"{w.get('name')}: parameter {n} missing")
                elif c.get("value") != params[n].get("value") and n != "[Parameter 1]": errs.append(f"{w.get('name')}: parameter {n} value differs")
            continue
        for n, c in dcols.items():
            if n not in cols:
                if n not in conn: errs.append(f"{w.get('name')}: field {n} not in data source")
                continue
            a, b = c.find("calculation"), cols[n].find("calculation")
            if (a is None) != (b is None) or (a is not None and a.get("formula") != b.get("formula")):
                errs.append(f"{w.get('name')}: {n} differs from data source")
            if a is not None and a.get("formula"):
                fs, ps = refs(a.get("formula"))
                for f in fs:
                    if f not in dcols: errs.append(f"{w.get('name')}: {c.get('caption')} needs {f} in dependencies")
                for p in ps:
                    if p not in params: errs.append(f"{w.get('name')}: unknown parameter {p}")
        for ci in dd.findall("column-instance"):
            if ci.get("column") not in dcols: errs.append(f"{w.get('name')}: instance of undeclared {ci.get('column')}")
    for p in set(re.findall(r"\[Parameters\]\.(\[[^\]]+\])", s)):
        if p not in params: errs.append(f"{w.get('name')}: unknown parameter {p}")
for d in new:
    for z in d.iter("zone"):
        p = z.get("param")
        if p and p.startswith("[Parameters]") and "[" + p.split("].[")[1] not in params: errs.append(f"{d.get('name')}: control for missing {p}")
print(f"{len(new)} merged dashboards, {len(names)} zones checked; {len(errs)} error(s)")
print("\n".join(sorted(set(errs))[:40]))
