"""Build the Home Solutions (HS) dashboard from the data source of 'Home Solutions Dashboard.twbx'.

usage: python build_hs.py <Home Solutions Dashboard.twbx> <output.twbx> [editable|static]

editable: the four rebate rates are parameters on every page (type a new rate, the net figures follow).
static:   the rebate rates are fixed in the calculations (hs_model.REBATES).
"""
import os
import shutil
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "retail_dashboard"))

from lxml import etree  # noqa: E402

import build as engine  # noqa: E402  (the Retail layout engine)
import hs_model  # noqa: E402
import hs_pages  # noqa: E402
from core import C_MID, C_PY, C_TARGET, DS, X, txt  # noqa: E402

SRC_TWB = "Home Solutions Dashboard.twb"
HYPER = "Data/Extracts/federated_14wc2mc1wc09dr1b4kqf30.hyper"
# month keys added to the relationships of the month-wise target tables, so targets and commits
# can be shown month by month (without them each month's target would repeat in every month)
TXN_MONTH = "DATE(DATETRUNC('month', IF [Type] = 'OB' THEN [Calendar Day] ELSE DATE([InvoiceDate]) END))"
MONTH_KEYS = {"[Category (Monthly Targets)]": "[Month and year]",
              "[Category (Activity wise Month Target)]": "[Month and year (Activity wise Month Target)]"}


def style_xml(R):
    mn = lambda c: f'"[{DS}].[usr:{R[c].name}:qk]"'
    maps = "".join(f"<map to='{C_MID}'><bucket>{txt(mn(c))}</bucket></map>" for c in ["Sales", "OB", "MTD sales"])
    maps += "".join(f"<map to='{C_PY}'><bucket>{txt(mn(c))}</bucket></map>" for c in ["Sales PY", "OB PY", "MTD sales LY"])
    maps += "".join(f"<map to='{C_TARGET}'><bucket>{txt(mn(c))}</bucket></map>" for c in ["Monthly target", "Sales commit"])
    maps += f"<map to='#9fd3b5'><bucket>{txt(mn('Sales 2 years ago'))}</bucket></map>"
    return f"<style><style-rule element='mark'><encoding attr='color' field='[:Measure Names]' type='palette'>{maps}</encoding></style-rule></style>"


def add_month_keys(ds):
    og = [c for c in ds if "object-graph" in str(c.tag)][0]
    done = 0
    for rel in og.iter():
        if not str(rel.tag).endswith("relationship"):
            continue
        e = rel.find("expression")
        s = etree.tostring(e).decode()
        for cat, month in MONTH_KEYS.items():
            if cat in s and month not in s:
                eq = etree.SubElement(e, "expression", op="=")
                etree.SubElement(eq, "expression", op=TXN_MONTH)
                etree.SubElement(eq, "expression", op=month)
                done += 1
    if done != len(MONTH_KEYS):
        raise ValueError(f"month keys added to {done} relationships, expected {len(MONTH_KEYS)}")


def main(src, out, variant="editable"):
    editable = variant == "editable"
    hs_pages.EDITABLE = editable
    hs_pages.SIDE_NOTE = hs_pages.side_note()
    title = "Home Solutions Performance" + ("" if editable else " (fixed rebate)")
    engine.configure(hs_pages, title)

    tmp = tempfile.mkdtemp()
    with zipfile.ZipFile(src) as z:
        z.extractall(tmp)
    tree = etree.parse(os.path.join(tmp, SRC_TWB))
    root = tree.getroot()
    ds_list = root.find("datasources")
    ds = [d for d in ds_list if d.get("name") == DS][0]
    R = hs_model.build_model(engine.rls_rules(ds), editable=editable)
    engine.avoid_name_clashes(R, {DS: ds})
    engine.clean_datasource(ds, R, DS)
    ds.insert(list(ds).index(ds.find("layout")) + 1, X(style_xml(R)))
    add_month_keys(ds)
    for d in list(ds_list):
        ds_list.remove(d)
    ds_list.append(engine.params_datasource(R))
    ds_list.append(ds)

    S = hs_pages.build_sheets(R)
    for tag in ("repository-location", "actions", "worksheets", "dashboards", "windows", "datagraph", "thumbnails"):
        # also the copies Tableau saves under a version-tagged name (e.g. '_.fcp.DatagraphCoreV1.true...datagraph')
        for e in [e for e in root if str(e.tag) == tag or str(e.tag).endswith("..." + tag)]:
            root.remove(e)
    ws = etree.SubElement(root, "worksheets")
    dbs = etree.SubElement(root, "dashboards")
    dash_used, vis = {}, []
    for d in hs_pages.DASHBOARDS:
        el, used, z = engine.dashboard(R, S, d)
        dbs.append(el)
        dash_used[d] = used
        vis += z
    missing = sorted({n for used in dash_used.values() for n in used if n not in S})
    if missing:
        raise ValueError(f"dashboards reference missing sheets: {missing}")
    on_dash = set().union(*dash_used.values())
    S = {n: s for n, s in S.items() if n in on_dash}
    for s in S.values():
        ws.append(X(s.xml()))
    root.append(engine.actions_xml(S, dash_used))
    root.append(engine.windows_xml(S, dash_used))
    if vis:
        root.append(engine.datagraph_xml(R, vis))
    rank = lambda e: next((i for i, t in enumerate(engine.TOP_ORDER) if str(e.tag) == t or str(e.tag).endswith("..." + t)), 99)
    kids = sorted(list(root), key=rank)
    if any(rank(k) == 99 for k in kids):
        raise ValueError(f"unknown top-level element(s): {[k.tag for k in kids if rank(k) == 99]}")
    for k in kids:
        root.remove(k)
        root.append(k)

    twb = etree.tostring(tree, xml_declaration=True, encoding="utf-8", pretty_print=True)
    twb = twb.replace(b"<?xml version='1.0' encoding='utf-8'?>", b"<?xml version='1.0' encoding='utf-8' ?>", 1)
    name = os.path.splitext(os.path.basename(out))[0]
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr(f"{name}.twb", twb)
        z.write(os.path.join(tmp, HYPER), HYPER)
    open(out.replace(".twbx", ".twb"), "wb").write(twb)
    shutil.rmtree(tmp)
    print(f"built {out} ({variant}): {len(S)} sheets, {len(hs_pages.DASHBOARDS)} dashboards, "
          f"{sum(1 for f in R.fields.values() if f.is_calc)} calculations")
    return R, S


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else "editable")
