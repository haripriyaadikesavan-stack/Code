"""Build Channel_Analytics_Dashboard_Reports.ipynb from dashboard_reports.py.

The script is split into cells at '# %%'. Three markers are filled from code that already
exists, so the notebook runs exactly the tested code:
  #@@SHAREPOINT@@     the SharePoint client of Channel_Analytics_v9.ipynb (cell 0.4)
  #@@STYLE@@          the report style of Channel_Analytics_v9.ipynb (cell 0.5)
  #@@TARGET_TOOLS@@   tableau/targets/organize_targets.py, wrapped as _target_split_tools()

usage: python build_notebook.py [out.ipynb]
"""
import json
import pathlib
import sys

HERE = pathlib.Path(__file__).parent
ROOT = HERE.parent
sys.path.insert(0, str(ROOT / "tableau" / "targets"))
import notebook_cell  # noqa: E402


def v9_cell(title):
    nb = json.loads((ROOT / "Channel_Analytics_v9.ipynb").read_text())
    for c in nb["cells"]:
        s = "".join(c["source"])
        if c["cell_type"] == "code" and title in s[:400]:
            return s
    raise KeyError(title)


def sharepoint_class():
    s = v9_cell("0.4 SHAREPOINT HELPERS")
    return s[s.index("from urllib.parse"):s.index("\n\nsp = SharePoint()")].rstrip() + "\n"


def style():
    s = v9_cell("0.5 REPORT STYLE")
    return s[s.index("from contextlib"):].rstrip() + "\n"


def target_tools():
    s = notebook_cell.keys_source()
    return s[s.index("def _target_split_tools"):s.index("\n\n\nTS = _target_split_tools()")].rstrip() + "\n"


def build(out):
    src = (HERE / "dashboard_reports.py").read_text()
    src = (src.replace("#@@SHAREPOINT@@\n", sharepoint_class())
              .replace("#@@STYLE@@\n", style())
              .replace("#@@TARGET_TOOLS@@\n", target_tools()))
    cells = []
    for part in src.split("# %%")[1:]:
        md = part.startswith(" [markdown]")
        body = part.split("\n", 1)[1].strip("\n")
        if md:
            body = "\n".join(ln[2:] if ln.startswith("# ") else ln.lstrip("#") for ln in body.split("\n"))
            cells.append({"cell_type": "markdown", "metadata": {}, "source": body.splitlines(keepends=True)})
        else:
            cells.append({"cell_type": "code", "execution_count": None, "metadata": {}, "outputs": [],
                          "source": body.splitlines(keepends=True)})
    nb = {"cells": cells, "metadata": {"kernelspec": {"display_name": "Python 3", "language": "python", "name": "python3"},
                                       "language_info": {"name": "python"}}, "nbformat": 4, "nbformat_minor": 5}
    pathlib.Path(out).write_text(json.dumps(nb, indent=1, ensure_ascii=False) + "\n")
    print(f"{out}: {len(cells)} cells")


if __name__ == "__main__":
    build(sys.argv[1] if len(sys.argv) > 1 else str(ROOT / "Channel_Analytics_Dashboard_Reports.ipynb"))
