from load import *
import re
from collections import defaultdict
tree, r, ds, calcs = load()
sheets = {w.get("name"): w for w in r.find("worksheets")}

def sheet_fields(w):
    s = etree.tostring(w).decode()
    inst = set(re.findall(r'column-instance column="(\[[^"]+\])"', s))
    # measures actually placed: references in rows/cols/encodings/filters
    return inst

TARGET_WORDS = ("target", "tgt", "estimate", "ach", "est.")
def is_target(n):
    c = cap(calcs, n).lower()
    return any(t in c for t in TARGET_WORDS)

report = defaultdict(list)
for d in r.find("dashboards"):
    dname = d.get("name")
    ps = [z.get("param") for z in d.iter("zone") if z.get("type-v2") == "paramctrl"]
    ctrl = {"P6" if P6 in p else "COPY" for p in ps if P6 in p or PC in p}
    for z in d.iter("zone"):
        n = z.get("name")
        if not n or n not in sheets: continue
        fields = sheet_fields(sheets[n])
        for f in fields:
            if f not in calcs: continue
            _, params = closure(calcs, [f])
            dep = {"P6" if p == P6 else "COPY" for p in params if p in (P6, PC)}
            if dep and ctrl and not (dep & ctrl):
                report[dname].append((n, cap(calcs, f), f"follows {dep} but dashboard shows {ctrl}"))
            elif not dep and ctrl and is_target(f) and "%" not in cap(calcs, f):
                report[dname].append((n, cap(calcs, f), "target value ignores Gross/Net"))
            elif dep and not ctrl:
                report[dname].append((n, cap(calcs, f), f"follows {dep} but dashboard has no Gross/Net control"))
for d, items in report.items():
    seen = set()
    print("==", d)
    for it in items:
        key = (it[1], it[2])
        if key in seen: continue
        seen.add(key); print("   ", it)
