"""Layout tidy-up for the existing dashboards (no sheet is added or removed).

1. fill_gaps: a row of charts (horizontal container) whose items stop short
   of the container's edge leaves an empty strip; the charts are stretched
   proportionally to fill it.
2. stack_sidebar: the floating controls / buttons in the left filter panel
   are stacked top-down in their current order with an even 6 px gap, so the
   export buttons no longer float half-way down the page.
"""

STATS = {}
GAP_MIN = 12          # px: smaller gaps are rounding, not holes
SIDEBAR_MAX_X = 240   # px: floating items left of this belong to the filter panel
SIDE_GAP = 6


def bump(k, n=1):
    STATS[k] = STATS.get(k, 0) + n


def geom(z):
    return [int(z.get(k)) for k in ("x", "y", "w", "h")]


def visible_children(c):
    return [z for z in c.findall("zone")]


def overlapping(kids, axis):
    p, s = ("x", "w") if axis == "h" else ("y", "h")
    kids = sorted(kids, key=lambda z: int(z.get(p)))
    for a, b in zip(kids, kids[1:]):
        if int(b.get(p)) < int(a.get(p)) + int(a.get(s)) - 300:
            return True
    return False


def rescale(c, axis, first, k, total_px):
    """Map every descendant coordinate along the axis: v' = c0 + (v - first) * k."""
    p, s = ("x", "w") if axis == "h" else ("y", "h")
    c0 = int(c.get(p))
    for z in c.iter("zone"):
        if z is c:
            continue
        v, n = int(z.get(p)), int(z.get(s))
        nv = c0 + round((v - first) * k)
        z.set(p, str(nv))
        z.set(s, str(max(1, round(n * k))))
        if z.get("fixed-size") and z.getparent().get("param") == ("horz" if axis == "h" else "vert"):
            z.set("fixed-size", str(round(int(z.get("fixed-size")) * k)))


def fill_gaps(d, name, W, H):
    root = d.find("zones").findall("zone")[0]
    for c in list(root.iter("zone")):
        if c.get("type-v2") != "layout-flow" or c.get("param") not in ("horz", "vert"):
            continue
        axis = "h" if c.get("param") == "horz" else "v"
        kids = visible_children(c)
        if axis != "h" or not kids or overlapping(kids, axis):
            continue
        # rows of charts only: every item is a sheet, or a container holding sheets
        def has_sheet(z):
            return any(x.get("name") and x.get("type-v2") is None for x in z.iter("zone"))
        if not all(has_sheet(z) for z in kids):
            continue
        names = [z.get("name") for z in kids if z.get("name")]
        if len(names) != len(set(names)):
            continue
        p, s, T = ("x", "w", W) if axis == "h" else ("y", "h", H)
        start = min(int(z.get(p)) for z in kids)
        end = max(int(z.get(p)) + int(z.get(s)) for z in kids)
        c0, c1 = int(c.get(p)), int(c.get(p)) + int(c.get(s))
        gap_px = ((c1 - end) + (start - c0)) * T / 100000
        if gap_px < GAP_MIN or end <= start:
            continue
        k = (c1 - c0) / (end - start)
        rescale(c, axis, start, k, T)
        bump("rows / columns stretched to close a gap")
        STATS.setdefault("gaps", []).append(f"{name}: {', '.join(z.get('name') or z.get('type-v2') for z in kids)[:90]} (+{int(gap_px)} px)")


def stack_sidebar(d, W, H):
    floating = d.find("zones").findall("zone")[1:]
    side = [z for z in floating if int(z.get("x")) * W / 100000 < SIDEBAR_MAX_X - 40
            and (int(z.get("x")) + int(z.get("w"))) * W / 100000 <= SIDEBAR_MAX_X
            and z.get("type-v2") not in ("color", "size", "shape")]
    if len(side) < 3:
        return False
    side.sort(key=lambda z: int(z.get("y")))
    gap = round(SIDE_GAP * 100000 / H)
    y = int(side[0].get("y"))
    moved = False
    # groups that sit side by side (export buttons) move together
    i = 0
    while i < len(side):
        row = [side[i]]
        while i + len(row) < len(side) and abs(int(side[i + len(row)].get("y")) - int(side[i].get("y"))) < 300:
            row.append(side[i + len(row)])
        dy = y - int(row[0].get("y"))
        if dy:
            moved = True
            for z in row:
                for zz in z.iter("zone"):
                    zz.set("y", str(int(zz.get("y")) + dy))
        y = max(int(z.get("y")) + int(z.get("h")) for z in row) + gap
        i += len(row)
    return moved


def apply(root):
    STATS.clear()
    stacked = []
    for d in root.find("dashboards"):
        size = d.find("size")
        W = int(size.get("maxwidth") or size.get("minwidth"))
        H = int(size.get("maxheight") or size.get("minheight"))
        fill_gaps(d, d.get("name"), W, H)
        if stack_sidebar(d, W, H):
            stacked.append(d.get("name"))
    STATS["filter panels re-stacked"] = len(stacked)
    return dict(STATS)
