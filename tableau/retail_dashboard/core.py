"""Core helpers for generating a Tableau workbook (.twb XML).

Field registry, calculated fields (with dependency closure), parameters,
and XML helpers. Everything is plain lxml so the output can be validated
before packaging.
"""
import re
import uuid
from xml.sax.saxutils import escape, quoteattr

from lxml import etree

DS = "federated.14wc2mc1wc09dr1b4kqf30rnj8dg"
DS_CAPTION = "Year+ (Retail_Primary_Sales_pre_processed)"
PDS = "federated.00walea07guk6w1fe7zhg073gkbk"
PDS_CAPTION = "Sales_Slabs_Unpivoted+ (Pavilion_Unpivoted)"
DS_CAPTIONS = {DS: DS_CAPTION, PDS: PDS_CAPTION}

# ---------------------------------------------------------------- palette
C_DARK = "#0b4d33"      # header band / titles
C_MID = "#1f7a4d"       # nav buttons / current year
C_ACCENT = "#3dcd58"    # Schneider green highlight
C_BG = "#f2f5f3"        # dashboard background
C_PY = "#b9c6c0"        # prior year / reference
C_TARGET = "#e8a33d"    # target
C_RED = "#c0392b"
C_AMBER = "#e8a33d"
C_TEXT = "#333333"
C_MUTED = "#6b6b6b"
C_HEAD_BG = "#e9f0ec"
C_BAND = "#f7f9f8"


def new_uuid():
    return "{" + str(uuid.uuid4()).upper() + "}"


def X(s):
    """Parse an XML fragment string into an element."""
    w = etree.fromstring('<w xmlns:user="http://www.tableausoftware.com/xml/user">' + s + "</w>")
    return w[0]


def attr(v):
    return quoteattr(str(v))


def txt(v):
    return escape(str(v))


# ---------------------------------------------------------------- fields
class Field:
    """A column of the Retail data source (base field or calculation)."""

    def __init__(self, name, datatype, role, ftype, caption=None, formula=None,
                 fmt=None, table="P", ds=DS):
        self.name = name                # local name without brackets
        self.datatype = datatype        # string, real, integer, date, datetime, boolean
        self.role = role                # dimension / measure
        self.ftype = ftype              # nominal / ordinal / quantitative
        self.caption = caption
        self.formula = formula          # resolved Tableau formula (None for base fields)
        self.fmt = fmt
        self.table = table              # P (primary), T (monthly targets), Q (QTOD), B (branch targets), S (SSA targets), X (none)
        self.ds = ds
        self.deps = set()               # names of fields referenced
        self.params = set()             # parameter names referenced

    @property
    def ref(self):
        return f"[{self.name}]"

    @property
    def is_calc(self):
        return self.formula is not None

    def column_xml(self, ns=True):
        a = []
        if self.caption:
            a.append(f"caption={attr(self.caption)}")
        a.append(f"datatype={attr(self.datatype)}")
        if self.fmt:
            a.append(f"default-format={attr(self.fmt)}")
        a.append(f"name={attr(self.ref)}")
        a.append(f"role={attr(self.role)}")
        a.append(f"type={attr(self.ftype)}")
        head = "<column " + " ".join(a)
        if self.is_calc:
            return head + f"><calculation class='tableau' formula={attr(self.formula)} /></column>"
        return head + " />"


class Param:
    def __init__(self, name, caption, datatype, default, members=None, rng=None):
        self.name, self.caption, self.datatype = name, caption, datatype
        self.default, self.members, self.rng = default, members, rng

    @property
    def ref(self):
        return f"[Parameters].[{self.name}]"

    def lit(self, v):
        if self.datatype == "string":
            return '"' + v + '"'
        if self.datatype == "date":
            return f"#{v}#"
        return str(v)

    def column_xml(self, with_members=True):
        dom = "list" if self.members else ("range" if self.rng else "any")
        ftype = "nominal" if self.datatype == "string" else "quantitative"
        v = self.lit(self.default)
        s = (f"<column caption={attr(self.caption)} datatype={attr(self.datatype)} name={attr('[' + self.name + ']')} "
             f"param-domain-type={attr(dom)} role='measure' type={attr(ftype)} value={attr(v)}>"
             f"<calculation class='tableau' formula={attr(v)} />")
        if with_members and self.members:
            s += "<members>" + "".join(f"<member value={attr(self.lit(m))} />" for m in self.members) + "</members>"
        if with_members and self.rng:
            lo, hi, gran = self.rng
            g = f" granularity={attr(gran)}" if gran else ""
            s += f"<range{g} max={attr(self.lit(hi))} min={attr(self.lit(lo))} />"
        return s + "</column>"


class Registry:
    """Holds base fields, parameters and calculations; resolves {Caption}
    tokens in formulas to calculation names and tracks dependencies."""

    def __init__(self):
        self.fields = {}        # name -> Field
        self.by_caption = {}    # caption -> Field (calcs only)
        self.params = {}        # name -> Param
        self._n = 0

    # --- registration -----------------------------------------------------
    def base(self, name, datatype, role, ftype, caption=None, table="P", ds=DS):
        f = Field(name, datatype, role, ftype, caption=caption, table=table, ds=ds)
        f.key = name if ds == DS else f"{ds}|{name}"
        self.fields[f.key] = f
        return f

    def param(self, name, caption, datatype, default, members=None, rng=None):
        p = Param(name, caption, datatype, default, members, rng)
        self.params[name] = p
        return p

    def calc(self, caption, formula, datatype="real", role=None, ftype=None, fmt=None, table=None, ds=DS):
        if caption in self.by_caption:
            raise ValueError(f"duplicate calc caption: {caption}")
        self._n += 1
        name = f"Calculation_83{self._n:017d}"
        if role is None:
            role = "measure" if datatype in ("real", "integer") else "dimension"
        if ftype is None:
            ftype = "quantitative" if role == "measure" else "nominal"
        f = Field(name, datatype, role, ftype, caption=caption, fmt=fmt, ds=ds)
        f.key = name
        f.formula, f.deps, f.params = self._resolve(formula, ds)
        # table = the logical table of the base fields it depends on
        tables = {self.fields[d].table for d in f.deps} - {"X"}
        if table:
            f.table = table
        elif len(tables) == 1:
            f.table = tables.pop()
        elif not tables:
            f.table = "X"
        else:
            f.table = "MIX"
        self.fields[name] = f
        self.by_caption[caption] = f
        return f

    def _resolve(self, formula, ds=DS):
        deps, params = set(), set()

        def tok(m):
            cap = m.group(1)
            f = self.by_caption.get(cap)
            if f is None:
                raise KeyError(f"unknown calc token {{{cap}}}")
            deps.add(f.key)
            return f.ref
        out = re.sub(r"\{\{([^{}]+)\}\}", tok, formula)
        for m in re.finditer(r"\[Parameters\]\.\[([^\]]+)\]", out):
            if m.group(1) not in self.params:
                raise KeyError(f"unknown parameter {m.group(1)}")
            params.add(m.group(1))
        # base field refs: [Name] not preceded by '].' (parameter) and not a calc we added
        stripped = re.sub(r"\[Parameters\]\.\[[^\]]+\]", "", out)
        for m in re.finditer(r"\[([^\[\]]+)\]", stripped):
            n = m.group(1)
            k = n if ds == DS else f"{ds}|{n}"
            if k in self.fields:
                deps.add(k)
            elif n in self.fields and self.fields[n].is_calc:
                deps.add(n)
            else:
                raise KeyError(f"unknown field [{n}] in formula: {formula[:120]}")
        return out, deps, params

    def __getitem__(self, caption):
        return self.by_caption[caption]

    # --- dependency closure ---------------------------------------------
    def closure(self, names):
        seen, stack, params = set(), list(names), set()
        while stack:
            n = stack.pop()
            if n in seen:
                continue
            seen.add(n)
            f = self.fields[n]
            params |= f.params
            stack.extend(f.deps)
        return seen, params


# ---------------------------------------------------------------- instances
def inst(f, kind):
    """Column-instance for a field.
    kind: 'dim' (discrete dimension), 'sum', 'usr' (aggregate calc measure),
          'usrn' (aggregate calc, discrete string), 'cnt'"""
    if kind == "dim":
        return f"none:{f.name}:nk", "None", "nominal"
    if kind == "sum":
        return f"sum:{f.name}:qk", "Sum", "quantitative"
    if kind == "usr":
        return f"usr:{f.name}:qk", "User", "quantitative"
    if kind == "usrn":
        return f"usr:{f.name}:nk", "User", "nominal"
    if kind == "attr":
        return f"attr:{f.name}:nk", "Attribute", "nominal"
    raise ValueError(kind)


def inst_xml(f, kind):
    n, d, t = inst(f, kind)
    return f"<column-instance column={attr(f.ref)} derivation={attr(d)} name={attr('[' + n + ']')} pivot='key' type={attr(t)} />"


def iref(f, kind):
    return f"[{f.ds}].[{inst(f, kind)[0]}]"
