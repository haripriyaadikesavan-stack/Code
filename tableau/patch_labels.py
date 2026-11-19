"""Patch a built Retail / HS / Wiser workbook in place: month labels Jan..Dec, band labels without
'1. ' numbering, and a fixed manual sort for those fields on every worksheet.
usage: python patch_labels.py <in.twbx> <out.twbx>"""
import re, sys, zipfile
from lxml import etree

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
ORDERS = {
    "Month": MONTHS, "WO month": MONTHS, "Explore by": MONTHS, "Explore across": MONTHS,
    "Partner slab": ["Below 25 K", "25 K - 1 L", "Below 1 L", "1 - 5 L", "5 - 25 L", "Above 25 L", "25 L - 1 Cr", "Above 1 Cr"],
    "Partner growth band": ["Growth 25% and above", "Growth 0-25%", "Decline 0-25%", "Decline over 25%",
                            "New this period (no sales last year)", "No sales this period", "No sales either year"],
    "Partner ach band": ["100% and above", "75-100%", "50-75%", "25-50%", "Below 25%", "No sales", "No target"],
    "Backlog ageing": ["Under 20 days", "20-50 days", "51-100 days", "101-200 days", "Over 100 days", "Over 200 days"],
    "FY PY sales bucket": ["Above 10 L", "5-10 L", "Below 5 L", "Nil last year"],
}
BANDS = {"Partner slab", "Partner growth band", "Partner ach band", "Backlog ageing", "FY PY sales bucket"}
MONTH_CALCS = {"Month", "WO month"}
OLD_MONTH = re.compile(r"^IF MONTH\((\[[^\]]+\])\) < 10 THEN '0' ELSE '' END \+ STR\(MONTH\(\1\)\) \+ ' ' \+ LEFT\(DATENAME\('month', \1\), 3\)$")


def patch(twb_bytes):
    root = etree.fromstring(twb_bytes)
    stats = {"formulas": 0, "sorts": 0}
    names = {}                                   # calc name -> caption, for the fields we sort
    for col in root.iter("column"):
        cap, calc = col.get("caption"), col.find("calculation")
        if calc is None or not (col.get("name") or "").startswith("[Calculation_"):
            continue
        f = calc.get("formula") or ""
        new = f
        if cap in MONTH_CALCS:
            m = OLD_MONTH.match(f)
            if m:
                new = f"LEFT(DATENAME('month', {m.group(1)}), 3)"
        elif cap in BANDS:
            new = re.sub(r"'\d\. ", "'", f)
        elif cap == "Partners at 100%+":
            new = f.replace("'1. 100% and above'", "'100% and above'")
        if new != f:
            calc.set("formula", new)
            stats["formulas"] += 1
        if cap in ORDERS:
            names[col.get("name")] = cap
    for view in root.iter("view"):
        if view.getparent().tag != "table":
            continue
        have = {e.get("column") for e in view.findall("manual-sort")}
        by_measure = {e.get("dimension-to-sort") for e in view.iter("shelf-sort-v2")}
        add = []
        for deps in view.findall("datasource-dependencies"):
            ds = deps.get("datasource")
            for ci in deps.findall("column-instance"):
                cap = names.get(ci.get("column"))
                ref = f"[{ds}].{ci.get('name')}"
                if cap and ci.get("derivation") == "None" and ref not in have and ref not in by_measure:
                    ms = etree.Element("manual-sort", column=ref, direction="ASC")
                    d = etree.SubElement(ms, "dictionary")
                    for v in ORDERS[cap]:
                        etree.SubElement(d, "bucket").text = f'"{v}"'
                    add.append(ms)
        if add:
            kids = list(view)
            after = [i for i, k in enumerate(kids) if k.tag in ("datasource-dependencies", "filter", "manual-sort")]
            pos = (after[-1] + 1) if after else 1
            for ms in reversed(add):
                view.insert(pos, ms)
            stats["sorts"] += len(add)
    out = etree.tostring(root, xml_declaration=True, encoding="utf-8")
    return out.replace(b"<?xml version='1.0' encoding='utf-8'?>", b"<?xml version='1.0' encoding='utf-8' ?>", 1), stats


def main(src, dst):
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.endswith(".twb"):
                data, stats = patch(data)
                print(f"{item.filename}: {stats['formulas']} formulas changed, {stats['sorts']} fixed sorts added")
            zout.writestr(item, data)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
