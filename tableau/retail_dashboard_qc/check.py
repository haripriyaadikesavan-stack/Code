import re, sys, html
from lxml import etree
from load import refs
path = sys.argv[1]
r = etree.parse(path).getroot()
errs = []
s = open(path, encoding="utf-8").read()
if "[Parameter 6]" in s: errs.append("Parameter 6 still referenced")
params = {c.get("name") for c in r.find("datasources").find("datasource[@name='Parameters']").findall("column")}
for p in set(re.findall(r"\[Parameters\]\.(\[[^\]]+\])", s)):
    if p not in params: errs.append(f"unknown parameter {p}")
for ds in r.find("datasources"):
    if ds.get("name") == "Parameters": continue
    cols = {c.get("name") for c in ds.findall("column")}
    conn = {html.unescape(k) for k in re.findall(r'<map key="(\[[^"]+\])"', etree.tostring(ds.find("connection")).decode())}
    groups = {g.get("name") for g in ds.findall("group")}
    for c in ds.findall("column"):
        cl = c.find("calculation")
        if cl is None or not cl.get("formula"): continue
        fs, _ = refs(cl.get("formula"))
        for f in fs:
            if f not in cols and f not in conn and f not in groups:
                errs.append(f"{ds.get('caption','')[:10]} calc {c.get('caption')}: unknown field {f}")
sheets = {w.get("name") for w in r.find("worksheets")}
dash = {d.get("name") for d in r.find("dashboards")}
for d in r.find("dashboards"):
    for z in d.iter("zone"):
        if z.get("name") and z.get("name") not in sheets and z.get("type-v2") not in ("text",):
            errs.append(f"{d.get('name')}: zone refers to missing sheet {z.get('name')}")
for a in r.find("actions"):
    src = a.find("source")
    if src is not None:
        if src.get("dashboard") and src.get("dashboard") not in dash: errs.append(f"action {a.get('caption')}: source dashboard {src.get('dashboard')}")
        if src.get("worksheet") and src.get("worksheet") not in sheets: errs.append(f"action {a.get('caption')}: source sheet {src.get('worksheet')}")
    for p in a.iter("param"):
        if p.get("name") == "target" and p.get("value") not in dash | sheets: errs.append(f"action {a.get('caption')}: target {p.get('value')}")
        if p.get("name") == "exclude" :
            for x in (p.get("value") or "").split(","):
                if x and x not in sheets: errs.append(f"action {a.get('caption')}: exclude {x}")
wins = {w.get("name") for w in r.find("windows")}
for n in sheets | dash:
    if n not in wins: errs.append(f"no window for {n}")
for w in r.find("windows"):
    if w.get("name") not in sheets | dash: errs.append(f"window for missing {w.get('name')}")
    for vp in w.iter("viewpoint"):
        if vp.get("name") and vp.get("name") not in sheets: errs.append(f"window {w.get('name')}: viewpoint {vp.get('name')} missing")
# button navigation targets
for m in re.findall(r'(?:sheet|target)=&quot;([^&]+)&quot;', s):
    pass
print(len(errs), "errors"); print("\n".join(errs[:40]))
