"""One visual theme for every dashboard and sheet of the Retail workbook.

Changes colours, borders, fonts and backgrounds only; no zone is moved or
resized and no sheet content changes.

    dashboard background   #f2f5f3 (light grey-green)
    charts / tables        white cards with a thin #d5e0d9 border
    section headers        #1f7a4d band, white bold 11
    page titles            #0b4d33 bold 18 on white
    filter panel           #e9f0ec
    controls               white box, thin border, dark green title
    table headers / panes  soft greens instead of beige / lilac / olive
"""
from lxml import etree

DARK = "#0b4d33"
MID = "#1f7a4d"
BG = "#f2f5f3"
CARD = "#ffffff"
LINE = "#d5e0d9"
SIDE = "#e9f0ec"
TEXT = "#333333"
HEAD = "#e9f0ec"
PANE_A, PANE_B, PANE_C = "#eef5f0", "#e3eee7", "#f7faf8"

BROWN = {"#a55f1d", "#550000"}
OLD_BORDERS = BROWN | {"#00aa00", "#009500", "#3a7455", "#d3d2d1", "#e0f0fa", "#c0c0c0"}
SECTION_BG = {"#f5ead7", "#e9f3f2", "#f3ebf3"}
TITLE_BG = {"#f5f5f5", "#fefaf1"}
PASTEL = {"#f3faf9", "#f7faf0", "#f3ebf3", "#eff3e3", "#fefaf1", "#faf5f0", "#f5f5f5"}
TITLE_GREENS = {"#3a7455", "#4b966f", "#000000", "#005500", "#009500", "#00aa00", "#006f00", "#008360"}
# worksheet-level background colours (table headers and panes)
SHEET_BG = {"#faf4e7": PANE_A, "#f5eddc": PANE_B, "#e8edda": PANE_A, "#f3ebf3": PANE_B,
            "#eff3e3": PANE_A, "#faf5f0": PANE_A, "#fefaf1": PANE_C}

STATS = {}


def bump(k, n=1):
    STATS[k] = STATS.get(k, 0) + n


def fmt_map(style):
    return {f.get("attr"): f for f in style.findall("format")} if style is not None else {}


def set_fmt(style, attr, value):
    f = fmt_map(style).get(attr)
    if f is None:
        f = etree.SubElement(style, "format")
        f.set("attr", attr)
    f.set("value", value)


def zone_style(z):
    s = z.find("zone-style")
    if s is None:
        s = etree.SubElement(z, "zone-style")
    return s


def px(z, key, total):
    return int(z.get(key)) * total / 100000


def is_section_header(z, H):
    s = fmt_map(z.find("zone-style"))
    bg = s["background-color"].get("value") if "background-color" in s else None
    return z.get("type-v2") == "text" and bg in SECTION_BG and px(z, "h", H) <= 60


def is_title_text(z):
    runs = list(z.iter("run"))
    return z.get("type-v2") == "text" and any(r.get("fontsize") in ("15", "16", "18") for r in runs)


def style_runs_title(runs):
    for r in runs:
        r.set("fontcolor", DARK)
        r.set("fontname", "Tableau Bold")
        r.set("bold", "true")
        if r.get("fontsize") in (None, "15", "16", "18"):
            r.set("fontsize", "18")
    if runs and runs[0].text:
        runs[0].text = runs[0].text.lstrip()


def theme_dashboard(d):
    size = d.find("size")
    W = int(size.get("maxwidth") or size.get("minwidth"))
    H = int(size.get("maxheight") or size.get("minheight"))
    top = d.find("zones").findall("zone")
    root = top[0]

    # dashboard background
    s = zone_style(root)
    set_fmt(s, "background-color", BG)
    if fmt_map(s).get("border-style") is not None and fmt_map(s)["border-style"].get("value") == "solid":
        set_fmt(s, "border-style", "none")
        set_fmt(s, "border-width", "0")

    tiled = set(root.iter("zone"))
    for z in d.iter("zone"):
        s = z.find("zone-style")
        f = fmt_map(s)
        t = z.get("type-v2")
        bg = f["background-color"].get("value") if "background-color" in f else None
        border = f["border-color"].get("value") if "border-color" in f else None
        solid = "border-style" in f and f["border-style"].get("value") == "solid"

        # borders: the brown frame goes, other coloured frames become thin grey-green lines
        if solid and border in OLD_BORDERS:
            if t == "layout-basic" and border == "#a55f1d":
                set_fmt(s, "border-style", "none")
                set_fmt(s, "border-width", "0")
                set_fmt(s, "border-color", "#000000")
                bump("brown frames removed")
            else:
                set_fmt(s, "border-color", LINE)
                set_fmt(s, "border-width", "1")
                bump("borders softened")

        if is_section_header(z, H):
            set_fmt(s, "background-color", MID)
            for r in z.iter("run"):
                r.set("fontcolor", "#ffffff")
                r.set("fontname", "Tableau Bold")
                r.set("bold", "true")
                r.set("fontsize", "11")
            bump("section headers")
        elif t == "text" and is_title_text(z):
            if bg is not None:
                set_fmt(s, "background-color", CARD)
            style_runs_title(list(z.iter("run")))
            bump("page titles")
        elif t in ("title", "empty") and bg in TITLE_BG:
            set_fmt(s, "background-color", CARD)
            bump("title bands")
        elif t in ("layout-flow", "layout-basic") and bg in PASTEL:
            if px(z, "w", W) <= 240:
                set_fmt(s, "background-color", SIDE)
                bump("filter panels")
            else:
                s.remove(f["background-color"])
                bump("container backgrounds cleared")
        elif t == "paramctrl" and bg is not None:
            set_fmt(s, "background-color", CARD)
        elif z.get("name") and t is None and bg in TITLE_BG | PASTEL:
            set_fmt(s, "background-color", CARD)

        # charts and tables in the tiled layout become white cards
        if z.get("name") and t is None and z in tiled and z is not root:
            s = zone_style(z)
            set_fmt(s, "background-color", CARD)
            set_fmt(s, "border-style", "solid")
            set_fmt(s, "border-width", "1")
            set_fmt(s, "border-color", LINE)
            if "margin" not in fmt_map(s):
                set_fmt(s, "margin", "4")
            bump("sheet cards")

        # greeting text: brighter green name
        for r in z.iter("run"):
            if r.get("fontcolor") == "#0a936c":
                r.set("fontcolor", MID)

    # dashboard title (layout-options)
    lo = d.find("layout-options/title/formatted-text")
    if lo is not None:
        style_runs_title(lo.findall("run"))
        bump("dashboard titles")

    # controls: one look for parameter controls; keep per-field rules, recolour them
    st = d.find("style")
    if st is None:
        st = etree.Element("style")
        d.insert(list(d).index(d.find("size")), st)
    rules = {sr.get("element"): sr for sr in st.findall("style-rule")}
    for el in ("table", "parameter-ctrl", "parameter-ctrl-title"):
        if el not in rules:
            rules[el] = etree.SubElement(st, "style-rule")
            rules[el].set("element", el)
    order = ["table", "parameter-ctrl", "parameter-ctrl-title"]
    for el in order:          # Tableau writes them in this order
        st.remove(rules[el])
        st.append(rules[el])
    for el, attrs in (("table", [("background-color", BG)]),
                      ("parameter-ctrl", [("font-size", "9"), ("color", TEXT), ("border-color", "#d9d9d9"),
                                          ("border-width", "1"), ("border-style", "solid"), ("background-color", CARD)]),
                      ("parameter-ctrl-title", [("font-size", "9"), ("font-weight", "bold"), ("color", DARK)])):
        sr = rules[el]
        have = {f.get("attr") for f in sr.findall("format") if f.get("field") is None}
        for a, v in attrs:
            if a not in have:
                f = etree.SubElement(sr, "format")
                f.set("attr", a)
                f.set("value", v)
        for f in sr.findall("format"):
            if f.get("field") is None and f.get("attr") == "color":
                f.set("value", DARK if el == "parameter-ctrl-title" else TEXT)
            elif f.get("attr") == "color" and f.get("value") in BROWN | TITLE_GREENS:
                f.set("value", DARK)
            elif f.get("attr") == "background-color" and f.get("field") is not None:
                f.set("value", CARD)


def theme_worksheet(w):
    st = w.find("table/style")
    n = 0
    if st is not None:
        for sr in st.findall("style-rule"):
            el = sr.get("element")
            for f in sr.findall("format"):
                a, v = f.get("attr"), f.get("value")
                if v in BROWN and a in ("color", "tick-color"):
                    f.set("value", DARK if a == "color" else "#9fb3a7")
                    n += 1
                elif a in ("background-color", "band-color") and v in SHEET_BG:
                    f.set("value", HEAD if el == "header" and a == "background-color" else SHEET_BG[v])
                    n += 1
                elif el in ("quick-filter-title",) and a == "color":
                    f.set("value", DARK)
                    n += 1
                elif el == "cell" and a == "border-color" and v == "#000000":
                    f.set("value", "#c9d6ce")
                    n += 1
                elif el == "worksheet" and a == "color" and v == "#000000":
                    f.set("value", TEXT)
                    n += 1
    # grey (#787878) title sub-runs stay; title colour is set by fix.format_titles
    if n:
        bump("sheet formats recoloured", n)


ZS_ORDER = ["border-color", "border-style", "border-width", "margin", "padding", "background-color"]


def sort_zone_styles(d):
    """Keep the format order Tableau writes in zone-style."""
    for s in d.iter("zone-style"):
        fs = s.findall("format")
        for f in fs:
            s.remove(f)
        for f in sorted(fs, key=lambda f: ZS_ORDER.index(f.get("attr")) if f.get("attr") in ZS_ORDER else 99):
            s.append(f)


def apply(root):
    STATS.clear()
    for d in root.find("dashboards"):
        theme_dashboard(d)
        sort_zone_styles(d)
    for w in root.find("worksheets"):
        theme_worksheet(w)
    return dict(STATS)
