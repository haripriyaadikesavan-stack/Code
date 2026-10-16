"""Worksheet XML builders in the style of the eCommerce workbook."""
import re

from core import (DS, DS_CAPTIONS, C_BAND, C_DARK, C_HEAD_BG, C_MID, C_MUTED, C_TEXT,
                  attr, inst, inst_xml, iref, new_uuid, txt)

MN = f"[{DS}].[:Measure Names]"
MV = f"[{DS}].[Multiple Values]"

UNITS_TOKEN = "<[Parameters].[p Units]>"


def mn(ds):
    return f"[{ds}].[:Measure Names]"


def mv(ds):
    return f"[{ds}].[Multiple Values]"

# stepped colour scales (worksheet level, as in the Retail workbook)
ACH_SCALE = (0.5, 1.5, 4, ["#c0392b", "#e8a33d", "#7fbf7b", "#1f7a4d"])      # <75% | 75-100% | 100-125% | >125%
GROWTH_SCALE = (-0.5, 0.5, 2, ["#c0392b", "#1f7a4d"])                         # declining | growing
NEUTRAL_SCALE = (0, 1, 2, ["#333333", "#333333"])                             # plain text colour


def column_scale(caption):
    """Font colour of a table column: growth red / green, achievement red / amber / green, others plain."""
    c = caption.lower()
    if "growth" in c:
        return GROWTH_SCALE
    if "ach %" in c or "qualifier %" in c:
        return ACH_SCALE
    return NEUTRAL_SCALE


def run(text, color=C_TEXT, size=9, bold=False):
    b = " bold='true'" if bold else ""
    font = "Tableau Bold" if bold else "Tableau Book"
    return f"<run{b} fontcolor={attr(color)} fontname={attr(font)} fontsize={attr(size)}>{txt(text)}</run>"


def title_runs(title, subtitle=None, param_units=True):
    r = run(title, C_DARK, 11, True)
    if param_units:
        r += run(" (", C_DARK, 11, True) + run(UNITS_TOKEN, C_DARK, 11, True) + run(")", C_DARK, 11, True)
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
<style-rule element='cell'><format attr='font-weight' data-class='total' scope='rows' value='bold' /><format attr='font-weight' data-class='subtotal' scope='rows' value='bold' />
<format attr='color' data-class='total' value='{C_DARK}' /><format attr='text-align' value='right' /><format attr='vertical-align' value='center' /></style-rule>
<style-rule element='header'><format attr='band-color' scope='rows' value='{C_BAND}' /><format attr='band-size' scope='rows' value='1' />
<format attr='background-color' scope='cols' value='{C_HEAD_BG}' /><format attr='font-weight' scope='cols' value='bold' /><format attr='color' scope='cols' value='{C_DARK}' />
<format attr='background-color' data-class='total' value='#dfeae4' /><format attr='background-color' data-class='subtotal' value='#eef3f0' />
<format attr='font-weight' data-class='total' scope='rows' value='bold' /><format attr='font-weight' data-class='subtotal' scope='rows' value='bold' /><format attr='color' data-class='total' value='{C_DARK}' /></style-rule>
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
        self.uses, self.extra_fields = [], []
        self.filters, self.slices, self.sorts = [], [], []
        self.manual_sort = ""
        self.rows = self.cols = ""
        self.rows_total = False
        self.cols_total = False
        self.mark = "Text"
        self.encodings = []
        self.label = None
        self.mark_color = None
        self.show_labels = True
        self.cell_align = None
        self.bg = "#ffffff"
        self.widths = []
        self.hide_axis = []         # (field ref, scope)
        self.scales = []            # (field ref, scale)
        self.fit = "entire-view"    # dashboard zoom: entire-view | fit-width
        self.nav_target = None

    # -- helpers ---------------------------------------------------------
    def use(self, caption_or_field, kind):
        f = self.R[caption_or_field] if isinstance(caption_or_field, str) else caption_or_field
        if (f, kind) not in self.uses:
            self.uses.append((f, kind))
        return iref(f, kind)

    def measures(self, items):
        """items: captions of aggregate calcs, or (field, kind) pairs."""
        refs = []
        for it in items:
            refs.append(self.use(*it) if isinstance(it, tuple) else self.use(it, "usr"))
        first = items[0]
        f0 = first[0] if isinstance(first, tuple) else self.R[first]
        MN, self.mn, self.mv = mn(f0.ds), mn(f0.ds), mv(f0.ds)
        members = "".join(f"<groupfilter function='member' level='[:Measure Names]' member={attr(chr(34) + r + chr(34))} />" for r in refs)
        self.filters.insert(0, f"<filter class='categorical' column='{MN}'><groupfilter function='union' user:op='manual'>{members}</groupfilter></filter>")
        self.manual_sort = (f"<manual-sort column='{MN}' direction='ASC'><dictionary>"
                            + "".join(f"<bucket>{txt(chr(34) + r + chr(34))}</bucket>" for r in refs)
                            + "</dictionary></manual-sort>")
        self.slices.append(MN)
        return refs

    def bool_filter(self, field, context=False):
        f = self.R[field] if isinstance(field, str) else field
        ref = self.use(f, "dim")
        lvl = "[" + inst(f, "dim")[0] + "]"
        ctx = " context='true'" if context else ""
        self.filters.append(f"<filter class='categorical' column='{ref}'{ctx}><groupfilter function='member' level='{lvl}' member='true' "
                            f"user:ui-domain='database' user:ui-enumeration='inclusive' user:ui-marker='enumerate' /></filter>")
        self.slices.insert(0, ref)

    def member_filter(self, field, values, context=True):
        f = self.R[field] if isinstance(field, str) else field
        ref = self.use(f, "dim")
        lvl = "[" + inst(f, "dim")[0] + "]"
        ctx = " context='true'" if context else ""
        values = [values] if isinstance(values, str) else values
        mem = lambda v: (f"<groupfilter function='member' level='{lvl}' member={attr(chr(34) + v + chr(34))} "
                         f"user:ui-domain='database' user:ui-enumeration='inclusive' user:ui-marker='enumerate' />")
        if len(values) == 1:
            inner = mem(values[0])
        else:
            inner = ("<groupfilter function='union' user:ui-domain='database' user:ui-enumeration='inclusive' user:ui-marker='enumerate'>"
                     + "".join(f"<groupfilter function='member' level='{lvl}' member={attr(chr(34) + v + chr(34))} />" for v in values)
                     + "</groupfilter>")
        self.filters.append(f"<filter class='categorical' column='{ref}'{ctx}>{inner}</filter>")
        self.slices.insert(0, ref)

    def shared_filter(self, field, group, context=True):
        """All values selected; the same filter (filter-group) on every sheet makes it one dashboard filter."""
        f = self.R[field] if isinstance(field, str) and field in self.R.by_caption else (self.R.fields[field] if isinstance(field, str) else field)
        ref = self.use(f, "dim")
        lvl = "[" + inst(f, "dim")[0] + "]"
        ctx = " context='true'" if context else ""
        self.filters.append(f"<filter class='categorical' column='{ref}'{ctx} filter-group='{group}'><groupfilter function='level-members' level='{lvl}' "
                            f"user:ui-enumeration='all' user:ui-marker='enumerate' /></filter>")
        self.slices.insert(0, ref)
        return ref

    def nonnull(self, caption):
        ref = self.use(caption, "usr")
        self.filters.append(f"<filter class='quantitative' column='{ref}' included-values='non-null' />")
        self.slices.append(ref)

    def exclude_null(self, ref, field):
        lvl = "[" + inst(field, "dim")[0] + "]"
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
        keys = {f.key for f, _ in self.uses} | {f.key for f in self.extra_fields}
        closure, params = R.closure(keys)
        for t in (self.title or "", self.label or ""):
            params |= set(re.findall(r"<\[Parameters\]\.\[([^\]]+)\]>", t))
        by_ds = {}
        for k in closure:
            f = R.fields[k]
            by_ds.setdefault(f.ds, []).append(f)
        dss = "".join(f"<datasource caption={attr(DS_CAPTIONS[d])} name='{d}' />" for d in sorted(by_ds))
        if params:
            dss += "<datasource name='Parameters' />"
        deps = ""
        if params:
            deps += ("<datasource-dependencies datasource='Parameters'>"
                     + "".join(R.params[p].column_xml(with_members=False) for p in sorted(params))
                     + "</datasource-dependencies>")
        for d in sorted(by_ds):
            cols = sorted(by_ds[d], key=lambda f: f.name)
            deps += (f"<datasource-dependencies datasource='{d}'>"
                     + "".join(f.column_xml() for f in cols)
                     + "".join(inst_xml(f, k) for f, k in self.uses if f.ds == d)
                     + "</datasource-dependencies>")
        sorts = ("<shelf-sorts>" + "".join(self.sorts) + "</shelf-sorts>") if self.sorts else ""
        slices = ("<slices>" + "".join(f"<column>{s}</column>" for s in self.slices) + "</slices>") if self.slices else ""
        view = (f"<view><datasources>{dss}</datasources>{deps}"
                + "".join(self.filters) + self.manual_sort + sorts + slices + "<aggregation value='true' /></view>")

        bg = self.bg
        style = ""
        for ref, scope in self.hide_axis:
            style += f"<style-rule element='axis'><format attr='display' class='0' field='{ref}' scope='{scope}' value='false' /></style-rule>"
        if self.kind == "table":
            style += STYLE_BASE.format(bg=bg, rowlabels="true") + TABLE_STYLE
        else:
            style += STYLE_BASE.format(bg=bg, rowlabels="false") + NO_LINES
            style += f"<style-rule element='pane'><format attr='background-color' value='{bg}' /></style-rule>"
        if self.scales:
            enc = ""
            for ref, (lo, hi, steps, colors) in self.scales:
                enc += (f"<encoding attr='color' field='{ref}' include-totals='true' max='{hi}' min='{lo}' num-steps='{steps}' type='custom-interpolated'>"
                        "<color-palette custom='true' name='' type='ordered-diverging'>"
                        + "".join(f"<color>{c}</color>" for c in colors) + "</color-palette></encoding>")
            style += f"<style-rule element='mark'>{enc}</style-rule>"
        for ref, w in self.widths:
            rule = "cell" if ref.endswith("[:Measure Names]") else "header"
            fmt = f"<format attr='width' field='{ref}' value='{w}' />"
            style = style.replace(f"<style-rule element='{rule}'>", f"<style-rule element='{rule}'>{fmt}", 1)
        style = f"<style>{style}</style>"

        order = ["color", "size", "shape", "lod", "text", "tooltip"]
        encs = sorted(self.encodings, key=lambda e: order.index(e[0]))
        # colour by Measure Values with one legend per measure (as the old workbook): per-column font colours
        enc = "".join(f"<{t} column='{r}'" + (" separate-domains='true'" if t == "color" and r.endswith("[Multiple Values]") else "") + " />"
                      for t, r in encs)
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
            if self.kind == "story":
                cell = cell.replace("</style-rule>", "<format attr='wrap' value='on' /></style-rule>")
        pane = (f"<pane selection-relaxation-option='selection-relaxation-allow'><view><breakdown value='auto' /></view>"
                f"<mark class='{self.mark}' />{enc}{label}<style>{cell}<style-rule element='mark'>{mstyle}</style-rule></style></pane>")
        rows = f"<rows total='true'>{self.rows}</rows>" if (self.rows and self.rows_total) else (f"<rows>{self.rows}</rows>" if self.rows else "<rows />")
        cols = (f"<cols total='true'>{self.cols}</cols>" if (self.cols and self.cols_total)
                else (f"<cols>{self.cols}</cols>" if self.cols else "<cols />"))
        return (f"<worksheet name={attr(self.name)}><layout-options><title><formatted-text>{self.title}</formatted-text></title></layout-options>"
                f"<table>{view}{style}<panes>{pane}</panes>{rows}{cols}</table>"
                f"<simple-id uuid='{new_uuid()}' /></worksheet>")


# ====================================================================== builders
def is_value(R, caption):
    return R[caption].fmt == "n#,##0.0;-#,##0.0"


def _grouped(i):
    """Tableau expression: integer expression i with thousands separators."""
    pad = lambda e: f"RIGHT('00' + STR({e}), 3)"
    return (f"IF {i} >= 1000000000 THEN STR(INT({i} / 1000000000)) + ',' + {pad(f'INT({i} / 1000000) % 1000')} + ',' + {pad(f'INT({i} / 1000) % 1000')} + ',' + {pad(f'{i} % 1000')} "
            f"ELSEIF {i} >= 1000000 THEN STR(INT({i} / 1000000)) + ',' + {pad(f'INT({i} / 1000) % 1000')} + ',' + {pad(f'{i} % 1000')} "
            f"ELSEIF {i} >= 1000 THEN STR(INT({i} / 1000)) + ',' + {pad(f'{i} % 1000')} ELSE STR({i}) END")


def as_text(R, cap):
    """A text version of a numeric measure for KPI cards. Tableau shows #### when a number is wider
    than the (narrow, fixed) cell of a sheet without headers; text is never replaced by ####."""
    f = R[cap]
    if f.datatype == "string":
        return cap
    tcap = f"{cap} (text)"
    if tcap in R.by_caption:
        return tcap
    x = "{{" + cap + "}}"
    if f.fmt == "p0.0%":
        expr = f"IFNULL(STR(ROUND({x} * 100, 1)) + '%', '-')"
    elif f.fmt == "n#,##0.0;-#,##0.0":
        i = f"INT(ROUND(ABS({x}) * 10, 0) / 10)"
        expr = (f"IFNULL((IF {x} < 0 THEN '-' ELSE '' END) + ({_grouped(i)}) + '.' + STR(INT(ROUND(ABS({x}) * 10, 0)) % 10), '-')")
    else:
        i = f"INT(ROUND(ABS({x}), 0))"
        expr = f"IFNULL((IF {x} < 0 THEN '-' ELSE '' END) + ({_grouped(i)}), '-')"
    R.calc(tcap, expr, "string", role="measure", ftype="nominal", ds=f.ds)
    return tcap


def kpi(R, name, title, big, line=None, units=None):
    """KPI card: one big value and at most one context value, both as text."""
    if units is None:
        units = is_value(R, big)
    s = Sheet(R, name, title_runs(title, param_units=units), "kpi")
    big = as_text(R, big)
    line = (line[0], as_text(R, line[1])) if line else None
    kind = lambda c: "usrn"
    rb = s.use(big, kind(big))
    s.encodings.append(("text", rb))
    lab = run(f"<{rb}>", C_TEXT, 16, True)
    if line:
        pre, cap = line
        rl = s.use(cap, kind(cap))
        s.encodings.append(("text", rl))
        lab += run("Æ\n", C_TEXT, 6) + run(pre, C_MUTED, 9) + run(f"<{rl}>", C_MID, 9, True)
    s.label = lab
    s.cell_align = "left"
    s.mark_color = C_MID
    return s


def nav(R, page, label=None):
    s = Sheet(R, f"Nav | {label or page}", "", "nav")
    ref = s.use("Nav count", "usr")
    s.encodings.append(("text", ref))
    s.bg = C_MID
    s.label = run(label or page, "#ffffff", 9, True)
    s.cell_align = "center"
    s.mark_color = C_MID
    s.nav_target = page
    return s
