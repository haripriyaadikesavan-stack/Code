"""Worksheet XML builders in the style of the eCommerce workbook."""
from core import (DS, DS_CAPTION, C_BAND, C_DARK, C_HEAD_BG, C_MID, C_MUTED, C_TEXT,
                  attr, inst, inst_xml, iref, new_uuid, txt)

MN = f"[{DS}].[:Measure Names]"
MV = f"[{DS}].[Multiple Values]"


def run(text, color=C_TEXT, size=9, bold=False):
    b = " bold='true'" if bold else ""
    font = "Tableau Bold" if bold else "Tableau Book"
    return f"<run{b} fontcolor={attr(color)} fontname={attr(font)} fontsize={attr(size)}>{txt(text)}</run>"


def title_runs(title, subtitle=None, param_units=True):
    r = run(title, C_DARK, 11, True)
    if param_units:
        r += run(" (", C_DARK, 11, True) + run("<[Parameters].[p Units]>", C_DARK, 11, True) + run(")", C_DARK, 11, True)
    if subtitle:
        r += run("\n" + subtitle, C_MUTED, 8)
    return r


STYLE_BASE = """
<style-rule element='table'><format attr='background-color' value='{bg}' /><format attr='show-null-value-warning' value='false' /></style-rule>
<style-rule element='worksheet'><format attr='font-family' value='Tableau Book' /><format attr='font-size' value='9' /><format attr='color' value='#333333' />
<format attr='display-field-labels' scope='cols' value='false' /><format attr='display-field-labels' scope='rows' value='{rowlabels}' /></style-rule>
<style-rule element='title'><format attr='background-color' value='#ffffff' /></style-rule>
"""
NO_LINES = """
<style-rule element='gridline'><format attr='stroke-size' scope='rows' value='0' /><format attr='line-visibility' scope='rows' value='off' />
<format attr='stroke-size' scope='cols' value='0' /><format attr='line-visibility' scope='cols' value='off' /></style-rule>
<style-rule element='zeroline'><format attr='stroke-size' value='0' /><format attr='line-visibility' value='off' /></style-rule>
<style-rule element='table-div'><format attr='stroke-size' scope='rows' value='0' /><format attr='line-visibility' scope='rows' value='off' />
<format attr='stroke-size' scope='cols' value='0' /><format attr='line-visibility' scope='cols' value='off' /></style-rule>
"""
TABLE_STYLE = f"""
<style-rule element='cell'><format attr='font-weight' data-class='total' value='bold' /><format attr='font-weight' data-class='subtotal' value='bold' />
<format attr='color' data-class='total' value='{C_DARK}' /><format attr='text-align' value='right' /><format attr='vertical-align' value='center' /></style-rule>
<style-rule element='header'><format attr='band-color' scope='rows' value='{C_BAND}' /><format attr='band-size' scope='rows' value='1' />
<format attr='background-color' scope='cols' value='{C_HEAD_BG}' /><format attr='font-weight' scope='cols' value='bold' /><format attr='color' scope='cols' value='{C_DARK}' />
<format attr='background-color' data-class='total' value='#dfeae4' /><format attr='background-color' data-class='subtotal' value='#eef3f0' />
<format attr='font-weight' data-class='total' value='bold' /><format attr='font-weight' data-class='subtotal' value='bold' /><format attr='color' data-class='total' value='{C_DARK}' /></style-rule>
<style-rule element='field-labels'><format attr='font-weight' value='bold' /><format attr='color' value='{C_DARK}' /><format attr='background-color' value='{C_HEAD_BG}' /></style-rule>
<style-rule element='pane'><format attr='band-color' scope='rows' value='{C_BAND}' /><format attr='band-size' scope='rows' value='1' />
<format attr='background-color' data-class='total' value='#dfeae4' /><format attr='background-color' data-class='subtotal' value='#eef3f0' /></style-rule>
<style-rule element='gridline'><format attr='line-visibility' scope='rows' value='off' /><format attr='line-visibility' scope='cols' value='off' /></style-rule>
<style-rule element='table-div'><format attr='stroke-color' scope='rows' value='#e3e7e5' /><format attr='stroke-size' scope='rows' value='1' /><format attr='line-visibility' scope='rows' value='on' />
<format attr='stroke-size' scope='cols' value='0' /><format attr='line-visibility' scope='cols' value='off' /></style-rule>
"""


class Sheet:
    """One worksheet. `uses` collects (Field, kind) column instances."""

    def __init__(self, R, name, title, kind="chart"):
        self.R, self.name, self.title, self.kind = R, name, title, kind
        self.uses = []          # (Field, kind)
        self.extra_fields = []  # fields needed only in dependencies
        self.filters = []       # xml strings
        self.slices = []        # instance refs
        self.sorts = []         # xml strings
        self.manual_sort = ""
        self.rows = ""
        self.cols = ""
        self.rows_total = False
        self.mark = "Text"
        self.encodings = []     # (type, ref)
        self.label = None       # customized label runs xml
        self.mark_color = None
        self.show_labels = True
        self.pane_style = ""
        self.cell_align = None
        self.extra_style = ""
        self.bg = "#ffffff"
        self.widths = []        # (instance ref, width)
        self.measure_names = None
        self.nav_target = None

    # -- helpers ---------------------------------------------------------
    def use(self, caption_or_field, kind):
        f = self.R[caption_or_field] if isinstance(caption_or_field, str) else caption_or_field
        if (f, kind) not in self.uses:
            self.uses.append((f, kind))
        return iref(f, kind)

    def base(self, name, kind="dim"):
        return self.use(self.R.fields[name], kind)

    def measures(self, captions):
        """Put several aggregate calcs on Measure Names / Measure Values."""
        refs = [self.use(c, "usr") for c in captions]
        self.measure_names = refs
        members = "".join(f"<groupfilter function='member' level='[:Measure Names]' member={attr(chr(34) + r + chr(34))} />" for r in refs)
        self.filters.insert(0, f"<filter class='categorical' column='{MN}'><groupfilter function='union' user:op='manual'>{members}</groupfilter></filter>")
        self.manual_sort = (f"<manual-sort column='{MN}' direction='ASC'><dictionary>"
                            + "".join(f"<bucket>{txt(chr(34) + r + chr(34))}</bucket>" for r in refs)
                            + "</dictionary></manual-sort>")
        self.slices.append(MN)
        return refs

    def bool_filter(self, caption, context=False):
        f = self.R[caption]
        ref = self.use(f, "dim")
        lvl = "[" + inst(f, "dim")[0] + "]"
        ctx = " context='true'" if context else ""
        self.filters.append(f"<filter class='categorical' column='{ref}'{ctx}><groupfilter function='member' level='{lvl}' member='true' "
                            f"user:ui-domain='database' user:ui-enumeration='inclusive' user:ui-marker='enumerate' /></filter>")
        self.slices.insert(0, ref)

    def nonnull(self, caption):
        ref = self.use(caption, "usr")
        self.filters.append(f"<filter class='quantitative' column='{ref}' included-values='non-null' />")
        self.slices.append(ref)

    def exclude_null(self, ref, field, kind="dim"):
        lvl = "[" + inst(field, kind)[0] + "]"
        self.filters.append(f"<filter class='categorical' column='{ref}'><groupfilter function='except' user:ui-domain='database' user:ui-enumeration='exclusive' user:ui-marker='enumerate'>"
                            f"<groupfilter function='level-members' level='{lvl}' /><groupfilter function='member' level='{lvl}' member='%null%' /></groupfilter></filter>")
        self.slices.append(ref)

    def top_n(self, ref, field, n, by_row_caption):
        lvl = "[" + inst(field, "dim")[0] + "]"
        expr = f"SUM({self.R[by_row_caption].ref})"
        self.extra_fields.append(self.R[by_row_caption])
        self.filters.append(
            f"<filter class='categorical' column='{ref}'><groupfilter count='{n}' end='top' function='end' units='records' user:ui-marker='end' user:ui-top-by-field='true'>"
            f"<groupfilter direction='DESC' expression={attr(expr)} function='order' user:ui-marker='order'>"
            f"<groupfilter function='except' user:ui-domain='database' user:ui-enumeration='exclusive' user:ui-marker='enumerate'>"
            f"<groupfilter function='level-members' level='{lvl}' /><groupfilter function='member' level='{lvl}' member='%null%' /></groupfilter>"
            f"</groupfilter></groupfilter></filter>")
        self.slices.append(ref)

    def sort_by(self, dim_ref, measure_ref, shelf="rows", direction="DESC"):
        self.sorts.append(f"<shelf-sort-v2 dimension-to-sort='{dim_ref}' direction='{direction}' measure-to-sort-by='{measure_ref}' shelf='{shelf}' />")

    # -- XML -------------------------------------------------------------
    def xml(self):
        R = self.R
        names = {f.name for f, _ in self.uses} | {f.name for f in self.extra_fields}
        closure, params = R.closure(names)
        cols = sorted((R.fields[n] for n in closure), key=lambda f: f.name)
        pdeps = ""
        title = self.title
        if "<[Parameters]." in title:
            import re
            params |= set(re.findall(r"<\[Parameters\]\.\[([^\]]+)\]>", title))
        if self.label and "<[Parameters]." in self.label:
            import re
            params |= set(re.findall(r"<\[Parameters\]\.\[([^\]]+)\]>", self.label))
        if params:
            pdeps = ("<datasource-dependencies datasource='Parameters'>"
                     + "".join(R.params[p].column_xml(with_members=False) for p in sorted(params))
                     + "</datasource-dependencies>")
        dsdeps = (f"<datasource-dependencies datasource='{DS}'>"
                  + "".join(f.column_xml() for f in cols)
                  + "".join(inst_xml(f, k) for f, k in self.uses)
                  + "</datasource-dependencies>")
        dss = f"<datasource caption={attr(DS_CAPTION)} name='{DS}' />" + ("<datasource name='Parameters' />" if params else "")
        sorts = ("<shelf-sorts>" + "".join(self.sorts) + "</shelf-sorts>") if self.sorts else ""
        slices = ("<slices>" + "".join(f"<column>{s}</column>" for s in self.slices) + "</slices>") if self.slices else ""
        view = (f"<view><datasources>{dss}</datasources>{pdeps}{dsdeps}"
                + "".join(self.filters) + self.manual_sort + sorts + slices + "<aggregation value='true' /></view>")

        bg = self.bg
        if self.kind == "table":
            style = STYLE_BASE.format(bg=bg, rowlabels="true") + TABLE_STYLE
        else:
            style = STYLE_BASE.format(bg=bg, rowlabels="false") + NO_LINES
            style += f"<style-rule element='pane'><format attr='background-color' value='{bg}' /></style-rule>"
        style += self.extra_style
        for ref, w in self.widths:
            rule = "cell" if ref == MN else "header"
            fmt = f"<format attr='width' field='{ref}' value='{w}' />"
            style = style.replace(f"<style-rule element='{rule}'>", f"<style-rule element='{rule}'>{fmt}", 1)
        style = f"<style>{style}</style>"

        enc = "".join(f"<{t} column='{r}' />" for t, r in self.encodings)
        enc = f"<encodings>{enc}</encodings>" if enc else ""
        label = f"<customized-label><formatted-text>{self.label}</formatted-text></customized-label>" if self.label else ""
        mstyle = ""
        if self.show_labels:
            mstyle += "<format attr='mark-labels-show' value='true' /><format attr='mark-labels-cull' value='true' />"
        if self.mark_color:
            mstyle += f"<format attr='mark-color' value='{self.mark_color}' />"
        cell = ""
        if self.cell_align:
            cell = f"<style-rule element='cell'><format attr='text-align' value='{self.cell_align}' /><format attr='vertical-align' value='center' /></style-rule>"
        pane = (f"<pane selection-relaxation-option='selection-relaxation-allow'><view><breakdown value='auto' /></view>"
                f"<mark class='{self.mark}' />{enc}{label}<style>{cell}<style-rule element='mark'>{mstyle}</style-rule>{self.pane_style}</style></pane>")
        rows = f"<rows total='true'>{self.rows}</rows>" if (self.rows and self.rows_total) else (f"<rows>{self.rows}</rows>" if self.rows else "<rows />")
        cols = f"<cols>{self.cols}</cols>" if self.cols else "<cols />"
        return (f"<worksheet name={attr(self.name)}><layout-options><title><formatted-text>{self.title}</formatted-text></title></layout-options>"
                f"<table>{view}{style}<panes>{pane}</panes>{rows}{cols}</table>"
                f"<simple-id uuid='{new_uuid()}' /></worksheet>")


# ====================================================================== builders
def kpi(R, name, title, big, lines, units=True):
    """KPI card: one big number + small context lines.
    big: caption of aggregate calc; lines: list of (prefix text, caption or None, suffix)."""
    s = Sheet(R, name, title_runs(title, param_units=units), "kpi")
    refs = {}
    allcaps = [big] + [c for _, c, _ in lines if c]
    for c in allcaps:
        kindc = "usrn" if R[c].datatype == "string" else "usr"
        refs[c] = s.use(c, kindc)
        s.encodings.append(("text", refs[c]))
    lab = run(f"<{refs[big]}>", C_TEXT, 18, True)
    for pre, c, suf in lines:
        lab += run("Æ\n", C_TEXT, 6)
        if pre:
            lab += run(pre, C_MUTED, 9)
        if c:
            lab += run(f"<{refs[c]}>", C_MID, 9, True)
        if suf:
            lab += run(suf, C_MUTED, 9)
    s.label = lab
    s.cell_align = "left"
    s.mark_color = C_MID
    return s


def nav(R, page):
    s = Sheet(R, f"Nav | {page}", "", "nav")
    s.title = ""
    f = R["Nav count"]
    ref = s.use(f, "usr")
    s.encodings.append(("text", ref))
    bg = C_MID
    s.bg = bg
    s.label = run(page, "#ffffff", 9, True)
    s.cell_align = "center"
    s.mark_color = bg
    s.nav_target = page
    return s
