"""Build 'Overall Retail Dashboard v3.twbx' from the user's original workbook.

usage: python build_v3.py <Overall Retail Dashboard.twbx> <gen_merge.twb> <output.twbx>

  1. fix.py      QC fixes (Gross/Net, targets, names, clean-up, titles)
  2. theme.py    one visual theme on every dashboard and sheet
  3. layout.py   close gaps in chart rows, stack the filter panels
  4. merge.py    add the analysis pages built by ../gen with RETAIL_MERGE=1
"""
import json
import os
import sys
import tempfile
import zipfile

from lxml import etree

import fix
import layout
import merge
import theme


def main(src, gen_twb, out):
    tmp = tempfile.mkdtemp()
    v2 = os.path.join(tmp, "v2.twbx")
    fix.main(src, v2)
    root = etree.parse(v2.replace(".twbx", ".twb")).getroot()
    report = {"theme": theme.apply(root), "layout": layout.apply(root)}
    g = etree.parse(gen_twb).getroot()
    merge.merge(root, g)
    report["merge"] = merge.LOG
    data = etree.tostring(root, xml_declaration=True, encoding="utf-8")
    data = data.replace(b"<?xml version='1.0' encoding='utf-8'?>", b"<?xml version='1.0' encoding='utf-8' ?>", 1)
    name = os.path.splitext(os.path.basename(out))[0]
    with zipfile.ZipFile(v2) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        zout.writestr(f"{name}.twb", data)
        for item in zin.infolist():
            if not item.filename.endswith(".twb"):
                zout.writestr(item, zin.read(item.filename))
    open(out.replace(".twbx", ".twb"), "wb").write(data)
    with open("build_log.json", "w") as f:
        json.dump(report, f, indent=1)
    print(json.dumps({k: (v if k != "layout" else {kk: vv for kk, vv in v.items() if kk != "gaps"}) for k, v in report.items()}, indent=1))


if __name__ == "__main__":
    main(*sys.argv[1:4])
