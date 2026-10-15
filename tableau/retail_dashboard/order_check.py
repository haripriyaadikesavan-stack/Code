"""Check child-element order against the order observed in real Tableau files."""
import sys
from collections import defaultdict
from lxml import etree

def tag(e):
    t = str(e.tag)
    return t.split("...")[-1] if "..." in t else t

def path_key(e):
    p = e.getparent()
    return (tag(p) if p is not None else "", tag(e))

def learn(files):
    before = defaultdict(set)   # parent -> {(a,b)}: a seen before b
    known = defaultdict(set)
    for f in files:
        for e in etree.parse(f).getroot().iter():
            if not isinstance(e.tag, str):
                continue
            kids = [tag(c) for c in e if isinstance(c.tag, str)]
            known[tag(e)] |= set(kids)
            for i, a in enumerate(kids):
                for b in kids[i + 1:]:
                    if a != b:
                        before[tag(e)].add((a, b))
    return before, known

def learn_attrs(files):
    seen = defaultdict(set)
    for f in files:
        for e in etree.parse(f).getroot().iter():
            if isinstance(e.tag, str):
                seen[tag(e)] |= set(e.attrib)
    return seen

def check(path, before, known):
    issues = set()
    for e in etree.parse(path).getroot().iter():
        if not isinstance(e.tag, str):
            continue
        kids = [tag(c) for c in e if isinstance(c.tag, str)]
        for i, a in enumerate(kids):
            for b in kids[i + 1:]:
                if a != b and (b, a) in before[tag(e)] and (a, b) not in before[tag(e)]:
                    issues.add(f"<{tag(e)}>: <{a}> placed before <{b}>")
        if tag(e) in known:
            for k in set(kids) - known[tag(e)]:
                issues.add(f"<{tag(e)}>: child <{k}> never seen in reference files")
    return issues

if __name__ == "__main__":
    b, k = learn(sys.argv[2:])
    iss = check(sys.argv[1], b, k)
    seen = learn_attrs(sys.argv[2:])
    for e in etree.parse(sys.argv[1]).getroot().iter():
        if isinstance(e.tag, str) and tag(e) in seen:
            for a in set(e.attrib) - seen[tag(e)]:
                iss.add(f"<{tag(e)}> attribute '{a}' never seen in reference files")
        elif isinstance(e.tag, str):
            iss.add(f"element <{tag(e)}> never seen in reference files")
    for i in sorted(iss):
        print(i)
    print(len(iss), "issue(s)")
