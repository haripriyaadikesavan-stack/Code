import re
from lxml import etree
SRC = "Overall Retail Dashboard.twb"   # extracted from the .twbx
DS = "federated.14wc2mc1wc09dr1b4kqf30rnj8dg"
P6, PC = "[Parameter 6]", "[Gross/Net (copy)_1650287831742312449]"

def load(path=SRC):
    tree = etree.parse(path)
    r = tree.getroot()
    ds = [d for d in r.find("datasources") if d.get("name") == DS][0]
    calcs = {}
    for c in ds.findall("column"):
        cl = c.find("calculation")
        if cl is not None and cl.get("class") == "tableau":
            calcs[c.get("name")] = c
    return tree, r, ds, calcs

def refs(f):
    f = re.sub(r"//[^\n]*", "", f or "")
    f = re.sub(r"'[^']*'|\"[^\"]*\"", "''", f)
    ps = set(re.findall(r"\[Parameters\]\.(\[[^\]]+\])", f))
    f2 = re.sub(r"\[Parameters\]\.\[[^\]]+\]", "", f)
    fs = set(re.findall(r"\[[^\[\]]+\]", f2))
    return fs, ps

def closure(calcs, start):
    seen, params, stack = set(), set(), list(start)
    while stack:
        n = stack.pop()
        if n in seen: continue
        seen.add(n)
        if n in calcs:
            fs, ps = refs(calcs[n].find("calculation").get("formula"))
            params |= ps
            stack += list(fs)
    return seen, params

def cap(calcs, n):
    c = calcs.get(n)
    return (c.get("caption") if c is not None else None) or n
