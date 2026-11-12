"""
xlchart.py - write .xlsx workbooks with native, editable Excel charts
using ONLY the Python standard library (zipfile + hand-written XML).
No openpyxl, no xlsxwriter, no pip install.

Chart types
-----------
Classic (work in every Excel version, LibreOffice, Google Sheets):
    column, bar, line, area      subtype: clustered | stacked | percent_stacked
                                 (line/area: standard | stacked | percent_stacked)
    pie, doughnut
    scatter                      subtype: marker_only | straight_with_markers |
                                          straight | smooth_with_markers | smooth
    bubble
    radar                        subtype: standard | with_markers | filled
    stock                        (high-low-close, open-high-low-close)
    combo                        chart.combine(other_chart), secondary axis via
                                 series option 'y2_axis': True
    waterfall (classic)          wb.add_waterfall_chart(...) - built from a line
                                 chart with up/down bars, works everywhere

Modern "chartEx" types (Excel 2016 / Microsoft 365 only):
    waterfall, histogram, pareto, box_whisker, funnel, treemap, sunburst

Quick start
-----------
    from xlchart import Workbook
    wb = Workbook()
    ws = wb.add_worksheet("Data")
    ws.write_rows("A1", [["Month", "Sales"], ["Jan", 10], ["Feb", 14], ["Mar", 12]])
    ch = wb.add_chart({"type": "column"})
    ch.add_series({"name": "=Data!$B$1",
                   "categories": "=Data!$A$2:$A$4",
                   "values": "=Data!$B$2:$B$4",
                   "color": "#1F4E79"})
    ch.set_title({"name": "Sales", "font": {"size": 14, "bold": True}})
    ch.set_x_axis({"line": {"color": "000000", "width": 1},
                   "font": {"color": "C00000"}})
    ch.set_y_axis({"min": 0, "major_gridlines": {"color": "EEEEEE"}})
    ws.insert_chart("D2", ch)
    wb.save("demo.xlsx")

See README.md for the full option reference.
"""
from __future__ import annotations

import datetime as _dt
import math
import re
import uuid
import zipfile
from xml.sax.saxutils import escape as _xesc

__all__ = ["Workbook", "Worksheet", "Chart", "ChartEx", "PALETTE"]

# ----------------------------------------------------------------------------
# constants
# ----------------------------------------------------------------------------
PALETTE = ["4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47",
           "264478", "9E480E", "636363", "997300", "255E91", "43682B"]
TEXT_COLOR = "595959"
GRID_COLOR = "D9D9D9"
DEFAULT_FONT = {"name": "Calibri", "size": 9, "color": TEXT_COLOR}

NAMED_COLORS = {
    "black": "000000", "white": "FFFFFF", "red": "FF0000", "dark_red": "C00000",
    "green": "00B050", "dark_green": "375623", "blue": "0070C0",
    "navy": "1F3864", "gray": "808080", "grey": "808080",
    "light_gray": "D9D9D9", "light_grey": "D9D9D9", "dark_gray": "595959",
    "dark_grey": "595959", "orange": "ED7D31", "yellow": "FFC000",
    "purple": "7030A0", "teal": "008080",
}

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
NS_CX = "http://schemas.microsoft.com/office/drawing/2014/chartex"
NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

X_ID, Y_ID, X2_ID, Y2_ID = 50010001, 50020002, 50030003, 50040004

CLASSIC_TYPES = {"column", "bar", "line", "area", "pie", "doughnut",
                 "scatter", "bubble", "radar", "stock"}
CHARTEX_TYPES = {"waterfall", "histogram", "pareto", "box_whisker",
                 "funnel", "treemap", "sunburst"}

STYLE_IDS = {None: 0, "bold": 1, "header": 2, "number": 3, "percent": 4,
             "integer": 5, "date": 6}

DASHES = {"solid": "solid", "dot": "sysDot", "round_dot": "sysDot",
          "square_dot": "sysDash", "dash": "dash", "dash_dot": "dashDot",
          "long_dash": "lgDash", "long_dash_dot": "lgDashDot",
          "long_dash_dot_dot": "lgDashDotDot"}

MARKERS = {"automatic": "auto", "auto": "auto", "none": "none",
           "circle": "circle", "square": "square", "diamond": "diamond",
           "triangle": "triangle", "x": "x", "star": "star", "plus": "plus",
           "dot": "dot", "dash": "dash", "short_dash": "dash",
           "long_dash": "dash"}

LABEL_POS = {"center": "ctr", "centre": "ctr", "right": "r", "left": "l",
             "above": "t", "top": "t", "below": "b", "bottom": "b",
             "inside_base": "inBase", "inside_end": "inEnd",
             "outside_end": "outEnd", "best_fit": "bestFit"}

TRENDLINES = {"linear": "linear", "exponential": "exp", "log": "log",
              "polynomial": "poly", "power": "power",
              "moving_average": "movingAvg"}

DISPLAY_UNITS = {"hundreds": "hundreds", "thousands": "thousands",
                 "ten_thousands": "tenThousands",
                 "hundred_thousands": "hundredThousands",
                 "millions": "millions", "ten_millions": "tenMillions",
                 "hundred_millions": "hundredMillions",
                 "billions": "billions", "trillions": "trillions"}


# ----------------------------------------------------------------------------
# small helpers
# ----------------------------------------------------------------------------
def _e(s):
    return _xesc(str(s), {'"': "&quot;"})


def _color(c):
    if c is None:
        return None
    s = str(c).strip()
    if s.lower() in NAMED_COLORS:
        return NAMED_COLORS[s.lower()]
    s = s.lstrip("#").upper()
    if not re.fullmatch(r"[0-9A-F]{6}", s):
        raise ValueError(f"Bad colour {c!r}: use 'RRGGBB', '#RRGGBB' or a name "
                         f"from {sorted(NAMED_COLORS)}")
    return s


def _is_num(v):
    return (isinstance(v, (int, float)) and not isinstance(v, bool)
            and not (isinstance(v, float) and (math.isnan(v) or math.isinf(v))))


def _num(v):
    if isinstance(v, bool):
        return str(int(v))
    if isinstance(v, int):
        return str(v)
    f = float(v)
    if f.is_integer() and abs(f) < 1e15:
        return str(int(f))
    return repr(f)


def _excel_date(v):
    """date/datetime -> Excel serial number."""
    if isinstance(v, _dt.datetime):
        delta = v - _dt.datetime(1899, 12, 30)
        return delta.days + delta.seconds / 86400.0
    if isinstance(v, _dt.date):
        return (v - _dt.date(1899, 12, 30)).days
    return v


def _uid():
    return "{" + str(uuid.uuid4()).upper() + "}"


def col_letter(n):
    """1-based column index -> letters (1 -> A)."""
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def col_index(letters):
    n = 0
    for ch in letters.upper():
        n = n * 26 + ord(ch) - 64
    return n


_CELL_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+)$")
_RANGE_RE = re.compile(
    r"^(?:'((?:[^']|'')+)'|([^'!]+))!(\$?[A-Za-z]{1,3}\$?\d+)(?::(\$?[A-Za-z]{1,3}\$?\d+))?$")


def parse_cell(ref):
    """'B3' -> (row, col), both 1-based."""
    m = _CELL_RE.match(ref.strip())
    if not m:
        raise ValueError(f"Bad cell reference {ref!r}")
    return int(m.group(2)), col_index(m.group(1))


def quote_sheet(name):
    if (re.fullmatch(r"[A-Za-z_][A-Za-z0-9_.]*", name)
            and not re.fullmatch(r"[A-Za-z]{1,3}\d+", name)
            and name.upper() not in ("TRUE", "FALSE")):
        return name
    return "'" + name.replace("'", "''") + "'"


def abs_ref(sheet, r1, c1, r2=None, c2=None):
    a = f"${col_letter(c1)}${r1}"
    if r2 is not None and (r2, c2) != (r1, c1):
        a += f":${col_letter(c2)}${r2}"
    return f"{quote_sheet(sheet)}!{a}"


def _is_range_list(ref):
    return (isinstance(ref, (list, tuple)) and len(ref) == 5 and isinstance(ref[0], str)
            and all(isinstance(x, int) and not isinstance(x, bool) for x in ref[1:]))


def _merge(*dicts):
    out = {}
    for d in dicts:
        if d:
            out.update(d)
    return out


# ----------------------------------------------------------------------------
# DrawingML fragments (shared by classic charts and chartEx)
# ----------------------------------------------------------------------------
def _solid(color, transparency=None):
    alpha = ""
    if transparency:
        alpha = f'<a:alpha val="{int(round((100 - transparency) * 1000))}"/>'
    return f'<a:solidFill><a:srgbClr val="{_color(color)}">{alpha}</a:srgbClr></a:solidFill>'


def _fill_xml(fill):
    """fill: 'RRGGBB' | {'color','transparency'} | {'none':True} |
             {'gradient': {'colors': [...], 'positions': [...], 'angle': 90}}"""
    if fill is None or fill is False:
        return ""
    if isinstance(fill, str):
        fill = {"color": fill}
    if fill.get("none"):
        return "<a:noFill/>"
    if "gradient" in fill:
        g = fill["gradient"]
        cols = g["colors"]
        pos = g.get("positions") or [round(i * 100 / max(1, len(cols) - 1)) for i in range(len(cols))]
        stops = "".join(f'<a:gs pos="{int(p * 1000)}"><a:srgbClr val="{_color(c)}"/></a:gs>'
                        for c, p in zip(cols, pos))
        ang = int(g.get("angle", 90) * 60000)
        return f'<a:gradFill rotWithShape="1"><a:gsLst>{stops}</a:gsLst><a:lin ang="{ang}" scaled="0"/></a:gradFill>'
    if fill.get("color") is None:
        return ""
    return _solid(fill["color"], fill.get("transparency"))


def _line_xml(line):
    """line: 'RRGGBB' | {'color','width'(pt),'dash','transparency'} | {'none':True}"""
    if line is None or line is False:
        return ""
    if isinstance(line, str):
        line = {"color": line}
    if line.get("none"):
        return "<a:ln><a:noFill/></a:ln>"
    w = f' w="{int(round(float(line["width"]) * 12700))}"' if line.get("width") is not None else ""
    inner = _solid(line["color"], line.get("transparency")) if line.get("color") else ""
    if line.get("dash"):
        inner += f'<a:prstDash val="{DASHES.get(line["dash"], line["dash"])}"/>'
    if not w and not inner:
        return ""
    return f'<a:ln{w} cap="rnd">{inner}<a:round/></a:ln>' if line.get("round") else f"<a:ln{w}>{inner}</a:ln>"


def _sppr(fill=None, line=None, ns="c"):
    body = _fill_xml(fill) + _line_xml(line)
    return f"<{ns}:spPr>{body}</{ns}:spPr>" if body else ""


def _rpr(font, tag="a:defRPr", lang=False):
    font = font or {}
    attrs = ' lang="en-US"' if lang else ""
    if font.get("size") is not None:
        attrs += f' sz="{int(round(float(font["size"]) * 100))}"'
    if "bold" in font:
        attrs += f' b="{1 if font["bold"] else 0}"'
    if "italic" in font:
        attrs += f' i="{1 if font["italic"] else 0}"'
    if font.get("underline"):
        attrs += ' u="sng"'
    kids = ""
    if font.get("color"):
        kids += _solid(font["color"])
    if font.get("name"):
        n = _e(font["name"])
        kids += f'<a:latin typeface="{n}"/><a:ea typeface="{n}"/><a:cs typeface="{n}"/>'
    return f"<{tag}{attrs}>{kids}</{tag}>" if kids else f"<{tag}{attrs}/>"


def _bodypr(rotation=None):
    if rotation is None:
        return "<a:bodyPr/>"
    return f'<a:bodyPr rot="{int(-float(rotation) * 60000)}" vert="horz"/>'


def _txpr(font=None, rotation=None, ns="c"):
    if not font and rotation is None:
        return ""
    return (f"<{ns}:txPr>{_bodypr(rotation)}<a:lstStyle/><a:p><a:pPr>{_rpr(font)}</a:pPr>"
            f'<a:endParaRPr lang="en-US"/></a:p></{ns}:txPr>')


def _rich(text, font=None, rotation=None, ns="c"):
    paras = ""
    for line in str(text).split("\n"):
        paras += (f"<a:p><a:pPr>{_rpr(font)}</a:pPr><a:r>{_rpr(font, 'a:rPr', lang=True)}"
                  f"<a:t>{_e(line)}</a:t></a:r></a:p>")
    return f"<{ns}:rich>{_bodypr(rotation)}<a:lstStyle/>{paras}</{ns}:rich>"


# ----------------------------------------------------------------------------
# Worksheet
# ----------------------------------------------------------------------------
class Worksheet:
    def __init__(self, wb, name):
        self.wb = wb
        self.name = name
        self.cells = {}          # (row, col) 1-based -> {'v','f','s'}
        self.col_widths = {}     # col (1-based) -> width in characters
        self.hidden_cols = set()
        self.charts = []         # (chart, row0, col0, opts)
        self.freeze = None

    # --- cell writing -------------------------------------------------------
    @staticmethod
    def _rc(cell):
        if isinstance(cell, str):
            return parse_cell(cell)
        return cell[0] + 1, cell[1] + 1   # (row, col) 0-based tuple

    def write(self, cell, value, style=None):
        """Write a number, string, bool, date or '=FORMULA'."""
        r, c = self._rc(cell)
        if isinstance(value, str) and value.startswith("=") and len(value) > 1:
            self.cells[(r, c)] = {"f": value[1:], "v": None, "s": style}
        else:
            if isinstance(value, (_dt.date, _dt.datetime)) and style is None:
                style = "date"
            self.cells[(r, c)] = {"v": value, "s": style}

    def write_formula(self, cell, formula, value=None, style=None):
        """value = cached result (shown by viewers that don't recalculate)."""
        r, c = self._rc(cell)
        self.cells[(r, c)] = {"f": formula.lstrip("="), "v": value, "s": style}

    def write_row(self, cell, values, style=None):
        r, c = self._rc(cell)
        for i, v in enumerate(values):
            self.write((r - 1, c - 1 + i), v, style)

    def write_column(self, cell, values, style=None):
        r, c = self._rc(cell)
        for i, v in enumerate(values):
            self.write((r - 1 + i, c - 1), v, style)

    def write_rows(self, cell, rows, header=True, style=None):
        """Write a 2-D list. First row gets the 'header' style if header=True."""
        r, c = self._rc(cell)
        for i, row in enumerate(rows):
            st = "header" if (header and i == 0) else style
            self.write_row((r - 1 + i, c - 1), row, st)

    def set_column(self, cols, width=None, hidden=False):
        """cols: 'B' or 'B:D'. width in Excel character units."""
        a, _, b = cols.partition(":")
        for c in range(col_index(a), col_index(b or a) + 1):
            if width is not None:
                self.col_widths[c] = width
            if hidden:
                self.hidden_cols.add(c)

    def freeze_panes(self, cell):
        self.freeze = self._rc(cell)

    def insert_chart(self, cell, chart, width=480, height=288, x_offset=0, y_offset=0):
        """Place a chart with its top-left corner at `cell`. Size in pixels."""
        r, c = self._rc(cell)
        self.charts.append((chart, r - 1, c - 1,
                            {"width": width, "height": height,
                             "x_offset": x_offset, "y_offset": y_offset}))

    def value(self, r, c):
        cell = self.cells.get((r, c))
        if not cell:
            return None
        return _excel_date(cell.get("v"))

    # --- XML -------------------------------------------------------------------
    def _col_px(self, c):  # c 1-based
        if c in self.hidden_cols:
            return 0
        w = self.col_widths.get(c)
        return 64 if w is None else int(w * 7 + 5)

    def _anchor_xml(self, r0, c0, o):
        def walk(start, off, size_fn):
            i, rem = start, off
            while rem >= size_fn(i + 1) and size_fn(i + 1) > 0:
                rem -= size_fn(i + 1)
                i += 1
            return i, rem
        row_px = lambda _r: 20
        fc, fco = walk(c0, o["x_offset"], self._col_px)
        fr, fro = walk(r0, o["y_offset"], row_px)
        tc, tco = walk(fc, fco + o["width"], self._col_px)
        tr, tro = walk(fr, fro + o["height"], row_px)
        e = 9525
        return (f"<xdr:from><xdr:col>{fc}</xdr:col><xdr:colOff>{fco * e}</xdr:colOff>"
                f"<xdr:row>{fr}</xdr:row><xdr:rowOff>{fro * e}</xdr:rowOff></xdr:from>"
                f"<xdr:to><xdr:col>{tc}</xdr:col><xdr:colOff>{tco * e}</xdr:colOff>"
                f"<xdr:row>{tr}</xdr:row><xdr:rowOff>{tro * e}</xdr:rowOff></xdr:to>")

    def _xml(self, selected, has_drawing):
        rows = {}
        maxr = maxc = 1
        for (r, c), cell in self.cells.items():
            rows.setdefault(r, []).append((c, cell))
            maxr, maxc = max(maxr, r), max(maxc, c)
        out = []
        for r in sorted(rows):
            cs = []
            for c, cell in sorted(rows[r], key=lambda x: x[0]):
                ref = f"{col_letter(c)}{r}"
                sid = STYLE_IDS.get(cell.get("s"), 0)
                sa = f' s="{sid}"' if sid else ""
                v, f = cell.get("v"), cell.get("f")
                v = _excel_date(v)
                if f is not None:
                    fx = f"<f>{_e(f)}</f>"
                    if v is None:
                        cs.append(f'<c r="{ref}"{sa}>{fx}</c>')
                    elif isinstance(v, str) and v.startswith("#"):
                        cs.append(f'<c r="{ref}"{sa} t="e">{fx}<v>{_e(v)}</v></c>')
                    elif isinstance(v, bool):
                        cs.append(f'<c r="{ref}"{sa} t="b">{fx}<v>{int(v)}</v></c>')
                    elif _is_num(v):
                        cs.append(f'<c r="{ref}"{sa}>{fx}<v>{_num(v)}</v></c>')
                    else:
                        cs.append(f'<c r="{ref}"{sa} t="str">{fx}<v>{_e(v)}</v></c>')
                else:
                    if v is None:
                        if sid:
                            cs.append(f'<c r="{ref}"{sa}/>')
                    elif isinstance(v, bool):
                        cs.append(f'<c r="{ref}"{sa} t="b"><v>{int(v)}</v></c>')
                    elif _is_num(v):
                        cs.append(f'<c r="{ref}"{sa}><v>{_num(v)}</v></c>')
                    else:
                        cs.append(f'<c r="{ref}"{sa} t="inlineStr"><is>'
                                  f'<t xml:space="preserve">{_e(v)}</t></is></c>')
            out.append(f'<row r="{r}">{"".join(cs)}</row>')

        cols = ""
        allc = sorted(set(self.col_widths) | self.hidden_cols)
        if allc:
            parts = []
            for c in allc:
                w = self.col_widths.get(c, 8.43)
                hid = ' hidden="1"' if c in self.hidden_cols else ""
                parts.append(f'<col min="{c}" max="{c}" width="{w}" customWidth="1"{hid}/>')
            cols = f"<cols>{''.join(parts)}</cols>"

        pane = ""
        if self.freeze:
            fr, fc = self.freeze
            xs = f' xSplit="{fc - 1}"' if fc > 1 else ""
            ys = f' ySplit="{fr - 1}"' if fr > 1 else ""
            pane = (f'<pane{xs}{ys} topLeftCell="{col_letter(fc)}{fr}" '
                    f'activePane="bottomRight" state="frozen"/>')
        sel = ' tabSelected="1"' if selected else ""
        drawing = '<drawing r:id="rId1"/>' if has_drawing else ""
        return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                f'<worksheet xmlns="{NS_MAIN}" xmlns:r="{NS_R}">'
                f'<dimension ref="A1:{col_letter(maxc)}{maxr}"/>'
                f'<sheetViews><sheetView workbookViewId="0"{sel}>{pane}</sheetView></sheetViews>'
                f'<sheetFormatPr defaultRowHeight="15"/>{cols}'
                f'<sheetData>{"".join(out)}</sheetData>'
                f'<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
                f'{drawing}</worksheet>')


# ----------------------------------------------------------------------------
# Classic chart
# ----------------------------------------------------------------------------
class Chart:
    """A classic DrawingML chart. Create with Workbook.add_chart()."""

    def __init__(self, wb, options):
        self.wb = wb
        self.type = options["type"]
        st = options.get("subtype")
        defaults = {"column": "clustered", "bar": "clustered", "line": "standard",
                    "area": "standard", "scatter": "marker_only",
                    "radar": "standard"}
        self.subtype = st or defaults.get(self.type)
        self.opts = dict(options)
        self.series = []
        self.combined = []
        self.title = None
        self.axes = {"x": {}, "y": {}, "x2": {}, "y2": {}}
        self.legend = {}
        self.plotarea = {}
        self.chartarea = {}
        self.font = {}
        self.data_table = None
        self.show_blanks = "gap"
        self.show_hidden = False

    # --- public API -----------------------------------------------------------
    def add_series(self, options=None, **kw):
        """See README for keys: name, categories, values, color, fill, border,
        line, marker, data_labels, points, trendline, y_error_bars, ..."""
        s = _merge(options, kw)
        if "values" not in s:
            raise ValueError("series needs 'values'")
        self.series.append(s)
        return self

    def set_title(self, options=None, **kw):
        o = _merge(options, kw)
        self.title = o
        return self

    def set_x_axis(self, options=None, **kw):
        self.axes["x"].update(_merge(options, kw)); return self

    def set_y_axis(self, options=None, **kw):
        self.axes["y"].update(_merge(options, kw)); return self

    def set_x2_axis(self, options=None, **kw):
        self.axes["x2"].update(_merge(options, kw)); return self

    def set_y2_axis(self, options=None, **kw):
        self.axes["y2"].update(_merge(options, kw)); return self

    def set_legend(self, options=None, **kw):
        self.legend.update(_merge(options, kw)); return self

    def set_plotarea(self, options=None, **kw):
        self.plotarea.update(_merge(options, kw)); return self

    def set_chartarea(self, options=None, **kw):
        self.chartarea.update(_merge(options, kw)); return self

    def set_font(self, options=None, **kw):
        """Default font for every text element in the chart."""
        self.font.update(_merge(options, kw)); return self

    def set_table(self, options=None, **kw):
        """Show a data table under the chart."""
        self.data_table = _merge({"horizontal": True, "vertical": True,
                                  "outline": True, "show_keys": True}, options, kw)
        return self

    def set_options(self, **kw):
        """Plot-level options: gap, overlap, hole_size, first_slice_angle,
        high_low_lines, drop_lines, up_down_bars, series_lines, bubble_scale,
        vary_colors."""
        self.opts.update(kw); return self

    def show_blanks_as(self, how):
        self.show_blanks = {"gap": "gap", "zero": "zero", "span": "span"}[how]; return self

    def show_hidden_data(self, flag=True):
        self.show_hidden = flag; return self

    def combine(self, other):
        """Overlay another chart (e.g. a line on a column chart)."""
        self.combined.append(other); return self

    # --- helpers ---------------------------------------------------------------
    @property
    def _kind(self):
        if self.type in ("pie", "doughnut"):
            return "pie"
        if self.type in ("scatter", "bubble"):
            return "xy"
        return "cat"

    def _font(self, *extra):
        return _merge(DEFAULT_FONT, self.font, *extra)

    def _data(self, ref, tag, kind="auto"):
        if ref is None:
            return ""
        if isinstance(ref, (list, tuple)) and not _is_range_list(ref):
            vals = list(ref)
            if kind == "auto":
                kind = "num" if vals and all(_is_num(_excel_date(v)) for v in vals) else "str"
            pts = "".join(f'<c:pt idx="{i}"><c:v>{_num(_excel_date(v)) if kind == "num" else _e(v)}</c:v></c:pt>'
                          for i, v in enumerate(vals) if v is not None)
            if kind == "num":
                return (f"<c:{tag}><c:numLit><c:formatCode>General</c:formatCode>"
                        f'<c:ptCount val="{len(vals)}"/>{pts}</c:numLit></c:{tag}>')
            return f'<c:{tag}><c:strLit><c:ptCount val="{len(vals)}"/>{pts}</c:strLit></c:{tag}>'
        f, vals = self.wb._resolve_values(ref)
        if kind == "auto":
            nonnull = [v for v in vals if v is not None]
            kind = "num" if nonnull and all(_is_num(v) for v in nonnull) else "str"
        if kind == "num":
            pts = "".join(f'<c:pt idx="{i}"><c:v>{_num(v)}</c:v></c:pt>'
                          for i, v in enumerate(vals) if _is_num(v))
            return (f"<c:{tag}><c:numRef><c:f>{_e(f)}</c:f><c:numCache><c:formatCode>General</c:formatCode>"
                    f'<c:ptCount val="{len(vals)}"/>{pts}</c:numCache></c:numRef></c:{tag}>')
        pts = "".join(f'<c:pt idx="{i}"><c:v>{_e(v)}</c:v></c:pt>'
                      for i, v in enumerate(vals) if v is not None)
        return (f"<c:{tag}><c:strRef><c:f>{_e(f)}</c:f><c:strCache>"
                f'<c:ptCount val="{len(vals)}"/>{pts}</c:strCache></c:strRef></c:{tag}>')

    def _n_points(self, s):
        v = s["values"]
        if isinstance(v, (list, tuple)) and not _is_range_list(v):
            return len(v)
        return len(self.wb._resolve_values(v)[1])

    def _name_xml(self, s):
        name = s.get("name")
        if name is None:
            return ""
        if (isinstance(name, str) and name.startswith("=")) or _is_range_list(name):
            f, vals = self.wb._resolve_values(name)
            v = vals[0] if vals else ""
            return (f"<c:tx><c:strRef><c:f>{_e(f)}</c:f><c:strCache><c:ptCount val=\"1\"/>"
                    f'<c:pt idx="0"><c:v>{_e("" if v is None else v)}</c:v></c:pt></c:strCache></c:strRef></c:tx>')
        return f"<c:tx><c:v>{_e(name)}</c:v></c:tx>"

    def _marker(self, m, color):
        if m is None:
            return ""
        if m is False:
            m = {"type": "none"}
        if isinstance(m, str):
            m = {"type": m}
        sym = MARKERS.get(m.get("type", "circle"), m.get("type", "circle"))
        if sym == "none":
            return '<c:marker><c:symbol val="none"/></c:marker>'
        size = f'<c:size val="{int(m.get("size", 5))}"/>'
        fill = m.get("fill", {"color": color})
        border = m.get("border", {"color": color, "width": 0.75})
        return f'<c:marker><c:symbol val="{sym}"/>{size}{_sppr(fill, border)}</c:marker>'

    def _dlbls(self, dl, plot_type, stacked=False):
        if not dl:
            return ""
        if dl is True:
            dl = {"value": True}
        x = ""
        if dl.get("num_format"):
            x += f'<c:numFmt formatCode="{_e(dl["num_format"])}" sourceLinked="0"/>'
        x += _sppr(dl.get("fill"), dl.get("border"))
        x += _txpr(self._font(dl.get("font")), dl.get("rotation"))
        pos = dl.get("position")
        if pos and plot_type not in ("area", "doughnut", "radar") and not (stacked and pos == "outside_end"):
            x += f'<c:dLblPos val="{LABEL_POS.get(pos, pos)}"/>'
        is_pie = plot_type in ("pie", "doughnut")
        x += (f'<c:showLegendKey val="{int(bool(dl.get("legend_key")))}"/>'
              f'<c:showVal val="{int(bool(dl.get("value")))}"/>'
              f'<c:showCatName val="{int(bool(dl.get("category")))}"/>'
              f'<c:showSerName val="{int(bool(dl.get("series_name")))}"/>'
              f'<c:showPercent val="{int(bool(dl.get("percentage")) and is_pie)}"/>'
              f'<c:showBubbleSize val="{int(bool(dl.get("bubble_size")) and plot_type == "bubble")}"/>')
        if dl.get("separator"):
            sep = {"comma": ", ", "semicolon": "; ", "period": ". ", "newline": "\n", "space": " "}
            x += f"<c:separator>{_e(sep.get(dl['separator'], dl['separator']))}</c:separator>"
        if is_pie and dl.get("leader_lines"):
            x += '<c:showLeaderLines val="1"/>'
        return f"<c:dLbls>{x}</c:dLbls>"

    def _trendline(self, t):
        if not t:
            return ""
        if isinstance(t, str):
            t = {"type": t}
        x = f"<c:name>{_e(t['name'])}</c:name>" if t.get("name") else ""
        x += _sppr(None, t.get("line", {"color": "595959", "width": 1.25, "dash": "dash"}))
        x += f'<c:trendlineType val="{TRENDLINES.get(t.get("type", "linear"), t.get("type"))}"/>'
        if t.get("type") == "polynomial":
            x += f'<c:order val="{int(t.get("order", 2))}"/>'
        if t.get("type") == "moving_average":
            x += f'<c:period val="{int(t.get("period", 2))}"/>'
        if t.get("forward"):
            x += f'<c:forward val="{_num(t["forward"])}"/>'
        if t.get("backward"):
            x += f'<c:backward val="{_num(t["backward"])}"/>'
        if t.get("intercept") is not None:
            x += f'<c:intercept val="{_num(t["intercept"])}"/>'
        x += f'<c:dispRSqr val="{int(bool(t.get("display_r_squared")))}"/>'
        x += f'<c:dispEq val="{int(bool(t.get("display_equation")))}"/>'
        if t.get("display_r_squared") or t.get("display_equation"):
            x += '<c:trendlineLbl><c:numFmt formatCode="General" sourceLinked="0"/></c:trendlineLbl>'
        return f"<c:trendline>{x}</c:trendline>"

    def _errbars(self, e, direction=None):
        if not e:
            return ""
        types = {"fixed": "fixedVal", "percentage": "percentage",
                 "standard_deviation": "stdDev", "standard_error": "stdErr",
                 "custom": "cust"}
        x = f'<c:errDir val="{direction}"/>' if direction else ""
        x += f'<c:errBarType val="{e.get("direction", "both")}"/>'
        x += f'<c:errValType val="{types[e.get("type", "fixed")]}"/>'
        x += f'<c:noEndCap val="{0 if e.get("end_cap", True) else 1}"/>'
        if e.get("type") == "custom":
            x += self._data(e.get("plus_values"), "plus", "num")
            x += self._data(e.get("minus_values"), "minus", "num")
        elif e.get("type") != "standard_error":
            x += f'<c:val val="{_num(e.get("value", 1))}"/>'
        x += _sppr(None, e.get("line", {"color": "404040", "width": 0.75}))
        return f"<c:errBars>{x}</c:errBars>"

    def _dpts(self, s, plot_type, n=0, default_fill=None, default_border=None):
        pts = s.get("points")
        entries = {}
        if isinstance(pts, dict):
            entries = {int(k): v for k, v in pts.items()}
        elif isinstance(pts, (list, tuple)):
            entries = {i: p for i, p in enumerate(pts) if p}
        entries = {k: ({"fill": v} if isinstance(v, str) else dict(v)) for k, v in entries.items()}
        if default_fill:
            for i in range(n):
                p = entries.setdefault(i, {})
                if "fill" not in p and "color" not in p:
                    p["fill"] = default_fill(i)
                if default_border and "border" not in p:
                    p["border"] = default_border
        out = ""
        line_like = plot_type in ("line", "scatter", "radar", "stock")
        for i in sorted(entries):
            p = entries[i]
            x = f'<c:idx val="{i}"/>'
            if plot_type in ("column", "bar", "bubble"):
                x += '<c:invertIfNegative val="0"/>'
            if line_like and "marker" in p:
                x += self._marker(p["marker"], _color(p.get("color")) or "000000")
            if plot_type in ("pie", "doughnut", "bubble"):
                x += '<c:bubble3D val="0"/>'
            if plot_type in ("pie", "doughnut") and p.get("explosion"):
                x += f'<c:explosion val="{int(p["explosion"])}"/>'
            if line_like:
                ln = p.get("line") or ({"color": p["color"], "width": 2.25} if p.get("color") else None)
                x += _sppr(None, ln)
            else:
                x += _sppr(p.get("fill", p.get("color")), p.get("border"))
            out += f"<c:dPt>{x}</c:dPt>"
        return out

    # --- series XML per plot type ----------------------------------------------
    def _ser_xml(self, s, idx, plot):
        t, st = plot.type, plot.subtype
        color = _color(s.get("color")) or PALETTE[idx % len(PALETTE)]
        head = f'<c:idx val="{idx}"/><c:order val="{idx}"/>{self._name_xml(s)}'
        stacked = st in ("stacked", "percent_stacked")
        dl = self._dlbls(s.get("data_labels"), t, stacked)
        tl = self._trendline(s.get("trendline"))
        cat = self._data(s.get("categories"), "cat")
        val = self._data(s["values"], "val", "num")

        if t in ("column", "bar"):
            fill = s.get("fill", {"color": color})
            sp = _sppr(fill, s.get("border"))
            inv = f'<c:invertIfNegative val="{int(bool(s.get("invert_if_negative")))}"/>'
            eb = self._errbars(s.get("y_error_bars"))
            return (f"<c:ser>{head}{sp}{inv}{self._dpts(s, t)}{dl}{tl}{eb}{cat}{val}</c:ser>")

        if t == "area":
            fill = s.get("fill", {"color": color})
            sp = _sppr(fill, s.get("border"))
            return f"<c:ser>{head}{sp}{self._dpts(s, t)}{dl}{tl}{cat}{val}</c:ser>"

        if t in ("line", "stock"):
            if t == "stock":
                line = s.get("line", {"none": True})
                m = s.get("marker", None)
                if m is None:
                    m = False
            else:
                line = s.get("line", {"color": color, "width": 2.25, "round": True})
                m = s.get("marker", {"type": "circle", "size": 5} if st == "with_markers" or plot.opts.get("markers") else False)
            sp = _sppr(None, line)
            eb = self._errbars(s.get("y_error_bars"))
            smooth = f'<c:smooth val="{int(bool(s.get("smooth")))}"/>'
            return (f"<c:ser>{head}{sp}{self._marker(m, color)}{self._dpts(s, t)}{dl}{tl}{eb}"
                    f"{cat}{val}{smooth}</c:ser>")

        if t in ("pie", "doughnut"):
            n = self._n_points(s)
            fills = s.get("colors") or PALETTE
            dp = self._dpts(s, t, n, lambda i: {"color": _color(fills[i % len(fills)])},
                            s.get("border", {"color": "FFFFFF", "width": 1}))
            expl = f'<c:explosion val="{int(s["explosion"])}"/>' if s.get("explosion") else ""
            return f"<c:ser>{head}{expl}{dp}{dl}{cat}{val}</c:ser>"

        if t == "scatter":
            line_on = st not in ("marker_only",)
            marker_on = st in ("marker_only", "straight_with_markers", "smooth_with_markers")
            line = s.get("line", {"color": color, "width": 2.25, "round": True} if line_on else {"none": True})
            m = s.get("marker", {"type": "circle", "size": 6} if marker_on else False)
            smooth = s.get("smooth", st in ("smooth", "smooth_with_markers"))
            eb = self._errbars(s.get("x_error_bars"), "x") + self._errbars(s.get("y_error_bars"), "y")
            xv = self._data(s.get("categories"), "xVal")
            yv = self._data(s["values"], "yVal", "num")
            return (f"<c:ser>{head}{_sppr(None, line)}{self._marker(m, color)}{self._dpts(s, t)}"
                    f'{dl}{tl}{eb}{xv}{yv}<c:smooth val="{int(bool(smooth))}"/></c:ser>')

        if t == "bubble":
            fill = s.get("fill", {"color": color, "transparency": 25})
            sp = _sppr(fill, s.get("border", {"color": color, "width": 0.75}))
            xv = self._data(s.get("categories"), "xVal")
            yv = self._data(s["values"], "yVal", "num")
            bs = self._data(s.get("sizes"), "bubbleSize", "num")
            return (f'<c:ser>{head}{sp}<c:invertIfNegative val="0"/>{self._dpts(s, t)}{dl}{tl}'
                    f'{xv}{yv}{bs}<c:bubble3D val="0"/></c:ser>')

        if t == "radar":
            if st == "filled":
                sp = _sppr(s.get("fill", {"color": color, "transparency": 40}),
                           s.get("border", {"color": color, "width": 1.5}))
                m = ""
            else:
                sp = _sppr(None, s.get("line", {"color": color, "width": 2.25}))
                m = self._marker(s.get("marker", {"type": "circle", "size": 5} if st == "with_markers" else False), color)
            return f"<c:ser>{head}{sp}{m}{self._dpts(s, t)}{dl}{cat}{val}</c:ser>"

        raise ValueError(f"unknown chart type {t}")

    # --- plot (chart-type element) ---------------------------------------------
    def _plot_xml(self, plot, sers, secondary, idx0):
        t, st, o = plot.type, plot.subtype, plot.opts
        axids = (f'<c:axId val="{X2_ID if secondary else X_ID}"/>'
                 f'<c:axId val="{Y2_ID if secondary else Y_ID}"/>')
        sx = "".join(self._ser_xml(s, idx0 + i, plot) for i, s in enumerate(sers))
        grouping = {"clustered": "clustered", "standard": "standard", "stacked": "stacked",
                    "percent_stacked": "percentStacked", "with_markers": "standard"}
        vary = f'<c:varyColors val="{int(bool(o.get("vary_colors", t in ("pie", "doughnut"))))}"/>'

        def lines(tag, v):
            if not v:
                return ""
            return f"<c:{tag}>{_sppr(None, v if isinstance(v, dict) else {'color': '404040', 'width': 0.75})}</c:{tag}>"

        def updown(v):
            if not v:
                return ""
            v = v if isinstance(v, dict) else {}
            up = v.get("up", {})
            dn = v.get("down", {})
            return (f'<c:upDownBars><c:gapWidth val="{int(v.get("gap", 150))}"/>'
                    f'<c:upBars>{_sppr(up.get("fill", {"color": "FFFFFF"}), up.get("border", {"color": "404040", "width": 0.75}))}</c:upBars>'
                    f'<c:downBars>{_sppr(dn.get("fill", {"color": "404040"}), dn.get("border", {"color": "404040", "width": 0.75}))}</c:downBars>'
                    f"</c:upDownBars>")

        if t in ("column", "bar"):
            g = grouping.get(st, "clustered")
            gap = int(o.get("gap", 150 if g == "clustered" else 100))
            ov = o.get("overlap", 100 if g != "clustered" else None)
            ovx = f'<c:overlap val="{int(ov)}"/>' if ov is not None else ""
            sl = lines("serLines", o.get("series_lines")) if g != "clustered" else ""
            return (f'<c:barChart><c:barDir val="{"bar" if t == "bar" else "col"}"/>'
                    f'<c:grouping val="{g}"/>{vary}{sx}<c:gapWidth val="{gap}"/>{ovx}{sl}{axids}</c:barChart>')
        if t == "line":
            g = grouping.get(st, "standard")
            return (f'<c:lineChart><c:grouping val="{g}"/>{vary}{sx}'
                    f'{lines("dropLines", o.get("drop_lines"))}{lines("hiLowLines", o.get("high_low_lines"))}'
                    f'{updown(o.get("up_down_bars"))}<c:marker val="1"/>{axids}</c:lineChart>')
        if t == "area":
            g = grouping.get(st, "standard")
            return (f'<c:areaChart><c:grouping val="{g}"/>{vary}{sx}'
                    f'{lines("dropLines", o.get("drop_lines"))}{axids}</c:areaChart>')
        if t == "pie":
            return (f"<c:pieChart>{vary}{sx}"
                    f'<c:firstSliceAng val="{int(o.get("first_slice_angle", 0))}"/></c:pieChart>')
        if t == "doughnut":
            return (f"<c:doughnutChart>{vary}{sx}"
                    f'<c:firstSliceAng val="{int(o.get("first_slice_angle", 0))}"/>'
                    f'<c:holeSize val="{int(o.get("hole_size", 50))}"/></c:doughnutChart>')
        if t == "scatter":
            style = "smoothMarker" if st in ("smooth", "smooth_with_markers") else "lineMarker"
            return f'<c:scatterChart><c:scatterStyle val="{style}"/>{vary}{sx}{axids}</c:scatterChart>'
        if t == "bubble":
            return (f"<c:bubbleChart>{vary}{sx}"
                    f'<c:bubbleScale val="{int(o.get("bubble_scale", 100))}"/>'
                    f'<c:showNegBubbles val="0"/>{axids}</c:bubbleChart>')
        if t == "radar":
            style = {"filled": "filled", "with_markers": "marker"}.get(st, "marker")
            return f'<c:radarChart><c:radarStyle val="{style}"/>{vary}{sx}{axids}</c:radarChart>'
        if t == "stock":
            hl = o.get("high_low_lines", True)
            return (f"<c:stockChart>{sx}{lines('hiLowLines', hl)}"
                    f'{updown(o.get("up_down_bars", len(sers) == 4))}{axids}</c:stockChart>')
        raise ValueError(t)

    # --- axes -----------------------------------------------------------------
    def _gridlines(self, tag, g):
        if not g:
            return ""
        line = {"color": GRID_COLOR, "width": 0.75}
        if isinstance(g, dict):
            if not g.get("visible", True):
                return ""
            line = g.get("line") or _merge(line, {k: v for k, v in g.items() if k != "visible"})
        return f"<c:{tag}>{_sppr(None, line)}</c:{tag}>"

    def _axis_defaults(self, role, tag):
        d = {"visible": True, "major_tick_mark": "none", "minor_tick_mark": "none",
             "label_position": "nextTo"}
        if role == "x":
            d["line"] = {"color": GRID_COLOR, "width": 0.75}
            d["major_gridlines"] = tag == "valAx" or self.type == "radar"
        elif role == "y":
            d["line"] = {"none": True} if self.type != "radar" else {"color": GRID_COLOR, "width": 0.75}
            d["major_gridlines"] = True
        elif role == "x2":
            d["visible"] = False
        elif role == "y2":
            d["line"] = {"none": True}
            d["major_gridlines"] = False
        return d

    def _axis_xml(self, role, tag, ax_id, cross_id, pos, vertical, cross_between=None):
        o = _merge(self._axis_defaults(role, tag), self.axes[role])
        x = f'<c:axId val="{ax_id}"/>'
        sc = ""
        if o.get("log_base"):
            sc += f'<c:logBase val="{_num(o["log_base"])}"/>'
        sc += f'<c:orientation val="{"maxMin" if o.get("reverse") else "minMax"}"/>'
        if o.get("max") is not None:
            sc += f'<c:max val="{_num(_excel_date(o["max"]))}"/>'
        if o.get("min") is not None:
            sc += f'<c:min val="{_num(_excel_date(o["min"]))}"/>'
        x += f"<c:scaling>{sc}</c:scaling>"
        x += f'<c:delete val="{0 if o.get("visible", True) else 1}"/>'
        x += f'<c:axPos val="{pos}"/>'
        x += self._gridlines("majorGridlines", o.get("major_gridlines"))
        x += self._gridlines("minorGridlines", o.get("minor_gridlines"))
        title = o.get("title", o.get("name"))
        if title:
            if isinstance(title, str):
                title = {"name": title}
            tfont = self._font({"size": 10, "bold": False}, title.get("font"))
            rot = title.get("rotation", -90 if vertical else None)
            x += (f"<c:title><c:tx>{_rich(title['name'], tfont, rot)}</c:tx>"
                  f'<c:overlay val="0"/></c:title>')
        if o.get("num_format"):
            x += f'<c:numFmt formatCode="{_e(o["num_format"])}" sourceLinked="0"/>'
        elif tag == "valAx":
            x += '<c:numFmt formatCode="General" sourceLinked="1"/>'
        elif tag == "dateAx":
            x += '<c:numFmt formatCode="yyyy-mm" sourceLinked="0"/>'
        x += f'<c:majorTickMark val="{o["major_tick_mark"]}"/>'
        x += f'<c:minorTickMark val="{o["minor_tick_mark"]}"/>'
        lp = {"next_to": "nextTo", "nextTo": "nextTo", "low": "low", "high": "high", "none": "none"}
        x += f'<c:tickLblPos val="{lp.get(o["label_position"], o["label_position"])}"/>'
        x += _sppr(o.get("fill"), o.get("line"))
        x += _txpr(self._font(o.get("font")), (o.get("font") or {}).get("rotation"))
        x += f'<c:crossAx val="{cross_id}"/>'
        cr = o.get("crosses", "max" if role == "y2" else "auto")
        if _is_num(cr):
            x += f'<c:crossesAt val="{_num(cr)}"/>'
        else:
            x += f'<c:crosses val="{ {"auto": "autoZero", "min": "min", "max": "max"}.get(cr, cr)}"/>'
        if tag == "catAx":
            x += '<c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/>'
            if o.get("interval_unit"):
                x += f'<c:tickLblSkip val="{int(o["interval_unit"])}"/>'
            if o.get("interval_tick"):
                x += f'<c:tickMarkSkip val="{int(o["interval_tick"])}"/>'
            x += '<c:noMultiLvlLbl val="0"/>'
        elif tag == "valAx":
            x += f'<c:crossBetween val="{cross_between or "between"}"/>'
            if o.get("major_unit"):
                x += f'<c:majorUnit val="{_num(o["major_unit"])}"/>'
            if o.get("minor_unit"):
                x += f'<c:minorUnit val="{_num(o["minor_unit"])}"/>'
            if o.get("display_units"):
                du = DISPLAY_UNITS.get(o["display_units"], o["display_units"])
                lbl = "<c:dispUnitsLbl/>" if o.get("display_units_visible", True) else ""
                x += f'<c:dispUnits><c:builtInUnit val="{du}"/>{lbl}</c:dispUnits>'
        elif tag == "dateAx":
            x += '<c:auto val="1"/><c:lblOffset val="100"/>'
            x += f'<c:baseTimeUnit val="{o.get("base_unit", "days")}"/>'
            if o.get("major_unit"):
                x += f'<c:majorUnit val="{_num(o["major_unit"])}"/>'
                x += f'<c:majorTimeUnit val="{o.get("major_unit_type", "months")}"/>'
            if o.get("minor_unit"):
                x += f'<c:minorUnit val="{_num(o["minor_unit"])}"/>'
                x += f'<c:minorTimeUnit val="{o.get("minor_unit_type", "days")}"/>'
        return f"<c:{tag}>{x}</c:{tag}>"

    def _cross_between(self, plot_type):
        pa = self.axes["x"].get("position_axis")
        if pa:
            return {"on_tick": "midCat", "between": "between"}[pa]
        if plot_type in ("scatter", "bubble"):
            return "midCat"
        return "between"

    def _axes_xml(self, has_secondary, sec_type):
        if self._kind == "pie":
            return ""
        horiz = self.type == "bar"
        out = ""
        if self._kind == "xy":
            out += self._axis_xml("x", "valAx", X_ID, Y_ID, "b", False, "midCat")
            out += self._axis_xml("y", "valAx", Y_ID, X_ID, "l", True, "midCat")
        else:
            xtag = "dateAx" if self.axes["x"].get("date_axis") else "catAx"
            out += self._axis_xml("x", xtag, X_ID, Y_ID, "l" if horiz else "b", horiz)
            out += self._axis_xml("y", "valAx", Y_ID, X_ID, "b" if horiz else "l", not horiz,
                                  self._cross_between(self.type))
        if has_secondary:
            if sec_type in ("scatter", "bubble"):
                out += self._axis_xml("x2", "valAx", X2_ID, Y2_ID, "b", False, "midCat")
            else:
                out += self._axis_xml("x2", "catAx", X2_ID, Y2_ID, "l" if horiz else "b", horiz)
            out += self._axis_xml("y2", "valAx", Y2_ID, X2_ID, "t" if horiz else "r", not horiz,
                                  self._cross_between(sec_type))
        return out

    def _legend_xml(self):
        lg = self.legend
        pos = lg.get("position", "bottom")
        if pos in ("none", None) or lg.get("none"):
            return ""
        p = {"bottom": "b", "top": "t", "left": "l", "right": "r", "top_right": "tr",
             "overlay_right": "r", "overlay_left": "l"}.get(pos, pos)
        entries = "".join(f'<c:legendEntry><c:idx val="{int(i)}"/><c:delete val="1"/></c:legendEntry>'
                          for i in lg.get("delete_series", []))
        ov = int(bool(lg.get("overlay")) or pos.startswith("overlay"))
        return (f'<c:legend><c:legendPos val="{p}"/>{entries}<c:overlay val="{ov}"/>'
                f'{_sppr(lg.get("fill"), lg.get("border"))}{_txpr(self._font(lg.get("font")))}</c:legend>')

    def _title_xml(self):
        t = self.title
        if not t or t.get("none"):
            return "", '<c:autoTitleDeleted val="1"/>'
        name = t.get("name", "")
        font = self._font({"size": 14, "bold": False}, t.get("font"))
        if isinstance(name, str) and name.startswith("="):
            f, vals = self.wb._resolve_values(name)
            v = vals[0] if vals else ""
            tx = (f"<c:tx><c:strRef><c:f>{_e(f)}</c:f><c:strCache><c:ptCount val=\"1\"/>"
                  f'<c:pt idx="0"><c:v>{_e(v)}</c:v></c:pt></c:strCache></c:strRef></c:tx>')
            extra = _txpr(font)
        else:
            tx = f"<c:tx>{_rich(name, font)}</c:tx>"
            extra = ""
        return (f'<c:title>{tx}<c:overlay val="{int(bool(t.get("overlay")))}"/>'
                f'{_sppr(t.get("fill"), t.get("border"))}{extra}</c:title>',
                '<c:autoTitleDeleted val="0"/>')

    def _xml(self):
        groups = []
        for ch in [self] + self.combined:
            prim = [s for s in ch.series if not s.get("y2_axis")]
            sec = [s for s in ch.series if s.get("y2_axis")]
            if prim:
                groups.append((ch, prim, False))
            if sec:
                groups.append((ch, sec, True))
        if not groups:
            raise ValueError("chart has no series")
        plots, idx = "", 0
        for ch, sers, secondary in groups:
            plots += self._plot_xml(ch, sers, secondary, idx)
            idx += len(sers)
        has_sec = any(g[2] for g in groups)
        sec_type = next((g[0].type for g in groups if g[2]), None)
        title, atd = self._title_xml()
        dt = ""
        if self.data_table:
            d = self.data_table
            dt = (f'<c:dTable><c:showHorzBorder val="{int(d["horizontal"])}"/>'
                  f'<c:showVertBorder val="{int(d["vertical"])}"/>'
                  f'<c:showOutline val="{int(d["outline"])}"/>'
                  f'<c:showKeys val="{int(d["show_keys"])}"/>'
                  f'{_txpr(self._font(d.get("font")))}</c:dTable>')
        pa = _sppr(self.plotarea.get("fill"), self.plotarea.get("border"))
        ca = _sppr(self.chartarea.get("fill", {"color": "FFFFFF"}),
                   self.chartarea.get("border", {"color": GRID_COLOR, "width": 0.75}))
        rounded = int(bool(self.chartarea.get("rounded_corners")))
        return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                f'<c:chartSpace xmlns:c="{NS_C}" xmlns:a="{NS_A}" xmlns:r="{NS_R}">'
                f'<c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="{rounded}"/>'
                f"<c:chart>{title}{atd}<c:plotArea><c:layout/>{plots}"
                f"{self._axes_xml(has_sec, sec_type)}{dt}{pa}</c:plotArea>"
                f"{self._legend_xml()}"
                f'<c:plotVisOnly val="{0 if self.show_hidden else 1}"/>'
                f'<c:dispBlanksAs val="{self.show_blanks}"/></c:chart>'
                f"{ca}{_txpr(self._font())}</c:chartSpace>")


# ----------------------------------------------------------------------------
# ChartEx (Excel 2016+ chart types)
# ----------------------------------------------------------------------------
class ChartEx:
    """waterfall, histogram, pareto, box_whisker, funnel, treemap, sunburst."""

    LAYOUT = {"waterfall": "waterfall", "histogram": "clusteredColumn",
              "pareto": "clusteredColumn", "box_whisker": "boxWhisker",
              "funnel": "funnel", "treemap": "treemap", "sunburst": "sunburst"}

    def __init__(self, wb, options):
        self.wb = wb
        self.type = options["type"]
        self.opts = dict(options)
        self.series = []
        self.title = None
        self.axes = {"x": {}, "y": {}, "y2": {}}
        self.legend = {}
        self.plotarea = {}
        self.chartarea = {}
        self.font = {}

    add_series = Chart.add_series
    set_title = Chart.set_title
    set_x_axis = Chart.set_x_axis
    set_y_axis = Chart.set_y_axis
    set_y2_axis = Chart.set_y2_axis
    set_legend = Chart.set_legend
    set_plotarea = Chart.set_plotarea
    set_chartarea = Chart.set_chartarea
    set_font = Chart.set_font
    set_options = Chart.set_options
    _font = Chart._font

    @property
    def requires(self):
        if self.type == "funnel":
            return "cx2", "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex"
        return "cx1", "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex"

    def _data_xml(self, s, did):
        parts = ""
        if s.get("categories") is not None:
            f, sheet, r1, c1, r2, c2 = self.wb._resolve(s["categories"])
            name = self.wb._xlchart_name(f)
            ws = self.wb.get_worksheet(sheet)
            n = r2 - r1 + 1
            lvls = ""
            for c in range(c2, c1 - 1, -1):          # leaf level first
                pts = ""
                for i in range(n):
                    v = ws.value(r1 + i, c)
                    if v is not None and v != "":
                        pts += f'<cx:pt idx="{i}">{_e(_num(v) if _is_num(v) else v)}</cx:pt>'
                lvls += f'<cx:lvl ptCount="{n}">{pts}</cx:lvl>'
            d = ' dir="row"' if c2 > c1 else ""
            parts += f'<cx:strDim type="cat"><cx:f{d}>{name}</cx:f>{lvls}</cx:strDim>'
        f, vals = self.wb._resolve_values(s["values"])
        name = self.wb._xlchart_name(f)
        dim = "size" if self.type in ("treemap", "sunburst") else "val"
        pts = "".join(f'<cx:pt idx="{i}">{_num(v)}</cx:pt>' for i, v in enumerate(vals) if _is_num(v))
        parts += (f'<cx:numDim type="{dim}"><cx:f>{name}</cx:f>'
                  f'<cx:lvl ptCount="{len(vals)}" formatCode="General">{pts}</cx:lvl></cx:numDim>')
        return f'<cx:data id="{did}">{parts}</cx:data>', vals

    def _dlbls(self, dl):
        defaults = {"waterfall": {"value": True, "position": "outside_end"},
                    "funnel": {"value": True, "position": "center"},
                    "treemap": {"category": True, "position": "inside_end"},
                    "sunburst": {"category": True, "position": "center"}}
        if dl is None:
            dl = defaults.get(self.type)
        if not dl:
            return ""
        if dl is True:
            dl = {"value": True}
        pos = dl.get("position")
        pa = f' pos="{LABEL_POS.get(pos, pos)}"' if pos else ""
        x = ""
        if dl.get("num_format"):
            x += f'<cx:numFmt formatCode="{_e(dl["num_format"])}" sourceLinked="0"/>'
        x += _sppr(dl.get("fill"), dl.get("border"), ns="cx")
        x += _txpr(self._font(dl.get("font")), None, ns="cx")
        x += (f'<cx:visibility seriesName="{int(bool(dl.get("series_name")))}" '
              f'categoryName="{int(bool(dl.get("category")))}" value="{int(bool(dl.get("value")))}"/>')
        return f"<cx:dataLabels{pa}>{x}</cx:dataLabels>"

    def _series_xml(self, s, i, vals):
        t = self.type
        color = _color(s.get("color")) or PALETTE[i % len(PALETTE)]
        x = ""
        name = s.get("name")
        if name is not None:
            if isinstance(name, str) and name.startswith("="):
                f, nv = self.wb._resolve_values(name)
                x += (f"<cx:tx><cx:txData><cx:f>{self.wb._xlchart_name(f)}</cx:f>"
                      f"<cx:v>{_e(nv[0] if nv else '')}</cx:v></cx:txData></cx:tx>")
            else:
                x += f"<cx:tx><cx:txData><cx:v>{_e(name)}</cx:v></cx:txData></cx:tx>"
        # series fill (treemap/sunburst keep automatic per-branch colours)
        if t not in ("treemap", "sunburst") or s.get("color") or s.get("fill"):
            x += _sppr(s.get("fill", {"color": color}), s.get("border"), ns="cx")
        # per-point colours
        pts = {}
        if t == "waterfall":
            totals = set(s.get("totals", s.get("subtotals", [])))
            inc = s.get("increase_color", "70AD47")
            dec = s.get("decrease_color", "E15759")
            tot = s.get("total_color", "4472C4")
            for k, v in enumerate(vals):
                if not _is_num(v):
                    continue
                pts[k] = {"fill": tot if k in totals else (inc if v >= 0 else dec)}
        up = s.get("points")
        if isinstance(up, dict):
            up = {int(k): v for k, v in up.items()}
        elif isinstance(up, (list, tuple)):
            up = {k: v for k, v in enumerate(up) if v}
        for k, p in (up or {}).items():
            pts[k] = {"fill": p} if isinstance(p, str) else p
        for k in sorted(pts):
            p = pts[k]
            x += f'<cx:dataPt idx="{k}">{_sppr(p.get("fill", p.get("color")), p.get("border"), ns="cx")}</cx:dataPt>'
        x += self._dlbls(s.get("data_labels"))
        x += f'<cx:dataId val="{i}"/>'
        lp = ""
        if t == "waterfall":
            lp += f'<cx:visibility connectorLines="{int(bool(s.get("connector_lines", True)))}"/>'
            totals = sorted(set(s.get("totals", s.get("subtotals", []))))
            if totals:
                lp += "<cx:subtotals>" + "".join(f'<cx:idx val="{k}"/>' for k in totals) + "</cx:subtotals>"
        elif t == "treemap":
            lp += f'<cx:parentLabelLayout val="{s.get("parent_labels", "overlapping")}"/>'
        elif t == "box_whisker":
            lp += (f'<cx:visibility meanLine="{int(bool(s.get("mean_line", False)))}" '
                   f'meanMarker="{int(bool(s.get("mean_marker", True)))}" '
                   f'nonoutliers="{int(bool(s.get("inner_points", False)))}" '
                   f'outliers="{int(bool(s.get("outliers", True)))}"/>')
            lp += f'<cx:statistics quartileMethod="{s.get("quartile_method", "exclusive")}"/>'
        elif t in ("histogram", "pareto"):
            if t == "pareto" and s.get("categories") is not None:
                lp += "<cx:aggregation/>"
            else:
                b = s.get("bins", {})
                attrs = f' intervalClosed="{b.get("closed", "r")}"'
                if b.get("underflow") is not None:
                    attrs += f' underflow="{_num(b["underflow"])}"'
                if b.get("overflow") is not None:
                    attrs += f' overflow="{_num(b["overflow"])}"'
                inner = ""
                if b.get("width") is not None:
                    inner = f'<cx:binSize val="{_num(b["width"])}"/>'
                elif b.get("count") is not None:
                    inner = f'<cx:binCount val="{int(b["count"])}"/>'
                lp += f"<cx:binning{attrs}>{inner}</cx:binning>"
        if lp:
            x += f"<cx:layoutPr>{lp}</cx:layoutPr>"
        if t == "pareto":
            x += '<cx:axisId val="1"/>'
        hidden = ""
        return f'<cx:series layoutId="{self.LAYOUT[t]}" uniqueId="{_uid()}"{hidden}>{x}</cx:series>'

    def _axis(self, aid, role, kind, default_grid=False, units=""):
        o = self.axes.get(role, {})
        hidden = ' hidden="1"' if o.get("visible", True) is False else ""
        if kind == "cat":
            gap = o.get("gap", self.opts.get("gap"))
            if gap is None:
                gap = {"histogram": 0, "pareto": 0, "funnel": 6}.get(self.type, 50)
            sc = f'<cx:catScaling gapWidth="{float(gap) / 100:g}"/>'
        else:
            attrs = ""
            for k, a in (("max", "max"), ("min", "min"), ("major_unit", "majorUnit"),
                         ("minor_unit", "minorUnit")):
                if o.get(k) is not None:
                    attrs += f' {a}="{_num(o[k])}"'
            sc = f"<cx:valScaling{attrs}/>"
        x = sc
        title = o.get("title", o.get("name"))
        if title:
            if isinstance(title, str):
                title = {"name": title}
            rot = title.get("rotation", -90 if role.startswith("y") else None)
            x += (f"<cx:title><cx:tx>{_rich(title['name'], self._font({'size': 10}, title.get('font')), rot, ns='cx')}"
                  f"</cx:tx></cx:title>")
        x += units
        g = o.get("major_gridlines", default_grid)
        if g:
            line = {"color": GRID_COLOR, "width": 0.75}
            if isinstance(g, dict):
                line = g.get("line") or _merge(line, {k: v for k, v in g.items() if k != "visible"})
            x += f"<cx:majorGridlines>{_sppr(None, line, ns='cx')}</cx:majorGridlines>"
        if o.get("major_tick_mark") and o["major_tick_mark"] != "none":
            x += f'<cx:majorTickMarks type="{o["major_tick_mark"]}"/>'
        x += "<cx:tickLabels/>"
        if o.get("num_format"):
            x += f'<cx:numFmt formatCode="{_e(o["num_format"])}" sourceLinked="0"/>'
        x += _sppr(o.get("fill"), o.get("line", {"color": GRID_COLOR, "width": 0.75} if kind == "cat" else {"none": True}), ns="cx")
        x += _txpr(self._font(o.get("font")), (o.get("font") or {}).get("rotation"), ns="cx")
        return f'<cx:axis id="{aid}"{hidden}>{x}</cx:axis>'

    def _xml(self):
        if not self.series:
            raise ValueError("chart has no series")
        t = self.type
        data, sers = "", ""
        for i, s in enumerate(self.series):
            dx, vals = self._data_xml(s, i)
            data += dx
            sers += self._series_xml(s, i, vals)
        if t == "pareto":
            pl = self.series[0].get("pareto_line", {"color": "ED7D31", "width": 2.25})
            sers += (f'<cx:series layoutId="paretoLine" ownerIdx="0" uniqueId="{_uid()}">'
                     f'{_sppr(None, pl, ns="cx")}<cx:axisId val="2"/></cx:series>')
        axes = ""
        if t in ("waterfall", "histogram", "box_whisker"):
            axes = self._axis(0, "x", "cat") + self._axis(1, "y", "val", True)
        elif t == "pareto":
            axes = (self._axis(0, "x", "cat") + self._axis(1, "y", "val", True)
                    + self._axis(2, "y2", "val", False, '<cx:units unit="percentage"/>'))
        elif t == "funnel":
            axes = self._axis(0, "x", "cat")
        title = ""
        if self.title and not self.title.get("none"):
            font = self._font({"size": 14}, self.title.get("font"))
            title = (f'<cx:title pos="t" align="ctr" overlay="0"><cx:tx>'
                     f'{_rich(self.title.get("name", ""), font, ns="cx")}</cx:tx></cx:title>')
        lg = self.legend
        default_pos = "bottom" if (t == "box_whisker" and len(self.series) > 1) else "none"
        pos = lg.get("position", default_pos)
        legend = ""
        if pos not in ("none", None):
            p = {"bottom": "b", "top": "t", "left": "l", "right": "r"}.get(pos, pos)
            legend = (f'<cx:legend pos="{p}" align="ctr" overlay="0">'
                      f'{_sppr(lg.get("fill"), lg.get("border"), ns="cx")}'
                      f'{_txpr(self._font(lg.get("font")), ns="cx")}</cx:legend>')
        pa = _sppr(self.plotarea.get("fill"), self.plotarea.get("border"), ns="cx")
        ca = _sppr(self.chartarea.get("fill", {"color": "FFFFFF"}),
                   self.chartarea.get("border", {"color": GRID_COLOR, "width": 0.75}), ns="cx")
        return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                f'<cx:chartSpace xmlns:a="{NS_A}" xmlns:r="{NS_R}" xmlns:cx="{NS_CX}">'
                f"<cx:chartData>{data}</cx:chartData>"
                f"<cx:chart>{title}<cx:plotArea><cx:plotAreaRegion>{sers}</cx:plotAreaRegion>"
                f"{axes}{pa}</cx:plotArea>{legend}</cx:chart>"
                f"{ca}{_txpr(self._font(), ns='cx')}</cx:chartSpace>")


# ----------------------------------------------------------------------------
# Workbook
# ----------------------------------------------------------------------------
_THEME = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    f'<a:theme xmlns:a="{NS_A}" name="Office Theme"><a:themeElements>'
    '<a:clrScheme name="Office">'
    '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>'
    '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'
    '<a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>'
    + "".join(f'<a:accent{i + 1}><a:srgbClr val="{c}"/></a:accent{i + 1}>' for i, c in enumerate(PALETTE[:6]))
    + '<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>'
    '</a:clrScheme>'
    '<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/>'
    '<a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/>'
    '<a:cs typeface=""/></a:minorFont></a:fontScheme>'
    '<a:fmtScheme name="Office"><a:fillStyleLst>'
    + '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' * 3
    + '</a:fillStyleLst><a:lnStyleLst>'
    + "".join(f'<a:ln w="{w}"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>' for w in (6350, 12700, 19050))
    + '</a:lnStyleLst><a:effectStyleLst>'
    + '<a:effectStyle><a:effectLst/></a:effectStyle>' * 3
    + '</a:effectStyleLst><a:bgFillStyleLst>'
    + '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' * 3
    + '</a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>'
)

_STYLES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    f'<styleSheet xmlns="{NS_MAIN}">'
    '<numFmts count="2"><numFmt numFmtId="164" formatCode="0.0%"/>'
    '<numFmt numFmtId="165" formatCode="yyyy\\-mm\\-dd"/></numFmts>'
    '<fonts count="2">'
    '<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>'
    '<font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>'
    '</fonts>'
    '<fills count="3"><fill><patternFill patternType="none"/></fill>'
    '<fill><patternFill patternType="gray125"/></fill>'
    '<fill><patternFill patternType="solid"><fgColor rgb="FFD9E1F2"/><bgColor indexed="64"/></patternFill></fill>'
    '</fills>'
    '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>'
    '<border><left/><right/><top/><bottom style="thin"><color rgb="FF8EA9DB"/></bottom><diagonal/></border></borders>'
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    '<cellXfs count="7">'
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
    '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>'
    '<xf numFmtId="4" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
    '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
    '<xf numFmtId="3" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
    '<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
    '</cellXfs>'
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
    '</styleSheet>'
)


class Workbook:
    def __init__(self):
        self.sheets = []
        self._xlnames = {}

    def add_worksheet(self, name=None):
        name = name or f"Sheet{len(self.sheets) + 1}"
        if any(s.name.lower() == name.lower() for s in self.sheets):
            raise ValueError(f"duplicate sheet name {name!r}")
        if len(name) > 31 or re.search(r"[\[\]:*?/\\]", name):
            raise ValueError(f"invalid sheet name {name!r}")
        ws = Worksheet(self, name)
        self.sheets.append(ws)
        return ws

    def get_worksheet(self, name):
        for s in self.sheets:
            if s.name == name:
                return s
        raise KeyError(f"no sheet named {name!r}")

    def add_chart(self, options=None, **kw):
        o = _merge(options, kw)
        t = o.get("type")
        if t in CLASSIC_TYPES:
            return Chart(self, o)
        if t in CHARTEX_TYPES:
            return ChartEx(self, o)
        raise ValueError(f"unknown chart type {t!r}; choose from "
                         f"{sorted(CLASSIC_TYPES | CHARTEX_TYPES)}")

    # --- references -------------------------------------------------------------
    def _resolve(self, ref):
        """-> (formula, sheet, r1, c1, r2, c2) with 1-based indices."""
        if isinstance(ref, (list, tuple)):
            sheet, r1, c1, r2, c2 = ref                 # 0-based, xlsxwriter style
            r1, c1, r2, c2 = r1 + 1, c1 + 1, r2 + 1, c2 + 1
        else:
            m = _RANGE_RE.match(str(ref).lstrip("=").strip())
            if not m:
                raise ValueError(f"Bad range {ref!r}; use e.g. \"=Data!$A$2:$A$10\"")
            sheet = m.group(1).replace("''", "'") if m.group(1) else m.group(2)
            r1, c1 = parse_cell(m.group(3))
            r2, c2 = parse_cell(m.group(4)) if m.group(4) else (r1, c1)
        r1, r2 = min(r1, r2), max(r1, r2)
        c1, c2 = min(c1, c2), max(c1, c2)
        return abs_ref(sheet, r1, c1, r2, c2), sheet, r1, c1, r2, c2

    def _resolve_values(self, ref):
        f, sheet, r1, c1, r2, c2 = self._resolve(ref)
        try:
            ws = self.get_worksheet(sheet)
        except KeyError:
            return f, []
        vals = [ws.value(r, c) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)]
        return f, vals

    def _xlchart_name(self, formula):
        if formula not in self._xlnames:
            self._xlnames[formula] = f"_xlchart.v1.{len(self._xlnames)}"
        return self._xlnames[formula]

    # --- classic waterfall --------------------------------------------------------
    def add_waterfall_chart(self, labels, values, totals=(), helper_cell=None,
                            increase_color="70AD47", decrease_color="E15759",
                            total_color="4472C4", title=None, gap=60,
                            hide_helpers=False, connector_lines=False):
        """Waterfall that works in ANY Excel version, LibreOffice and Google Sheets.

        Built as a stacked column chart with invisible "base" segments, so bars
        that cross zero are drawn correctly too.

        labels / values : single-column ranges on one sheet, e.g. "=Data!$A$2:$A$9".
        totals          : 0-based positions that are totals/subtotals (bar drawn
                          from zero to the running total). The value cell of a
                          total row is ignored, except for the first row, where
                          it is the opening level.
        helper_cell     : top-left cell of the 11 helper columns (formulas, so
                          the chart follows the data). Default: two columns to
                          the right of the values column, header in the row above.
        hide_helpers    : hide the helper columns (the chart still plots them).
        Returns a Chart; style it further and place it with ws.insert_chart().
        """
        lf, *_ = self._resolve(labels)
        vf, vsheet, vr1, vc1, vr2, _ = self._resolve(values)
        ws = self.get_worksheet(vsheet)
        n = vr2 - vr1 + 1
        totals = set(totals)
        if helper_cell:
            hr, hc = parse_cell(helper_cell)
            hr += 1                                    # header goes in helper_cell
        else:
            hr, hc = vr1, vc1 + 2
        names = ["WF cum", "WF lo", "WF hi", "WF up", "WF base+", "WF base-",
                 "WF inc+", "WF inc-", "WF dec+", "WF dec-", "WF total"]
        C = {k: hc + i for i, k in enumerate(["cum", "lo", "hi", "up", "bp", "bn",
                                              "ip", "in", "dp", "dn", "tot"])}
        L = col_letter
        if hr > 1:
            ws.write_row((hr - 2, hc - 1), names, "header")
        cum = prev = 0.0
        for i in range(n):
            r, vr = hr + i, vr1 + i
            v = ws.value(vr, vc1)
            v = v if _is_num(v) else 0
            cell = lambda k: f"{L(C[k])}{r}"
            vcell, pcum = f"{L(vc1)}{vr}", f"{L(C['cum'])}{r - 1}"
            is_tot = i in totals
            prev = cum
            if i == 0:
                cum, f_cum = v, vcell
            elif is_tot:
                f_cum = pcum
            else:
                cum, f_cum = cum + v, f"{pcum}+{vcell}"
            p_ref = "0" if i == 0 else pcum
            if is_tot:
                lo = hi = 0.0
                up = 0
                f_lo = f_hi = "0"
                f_up = "0"
            else:
                lo, hi = min(prev, cum), max(prev, cum)
                up = int(cum >= prev)
                f_lo = f"MIN({p_ref},{cell('cum')})"
                f_hi = f"MAX({p_ref},{cell('cum')})"
                f_up = f"IF({cell('cum')}>={p_ref},1,0)"
            pos_part = max(hi, 0) - max(lo, 0)
            neg_part = min(lo, 0) - min(hi, 0)
            vals = {
                "cum": (f_cum, cum), "lo": (f_lo, lo), "hi": (f_hi, hi), "up": (f_up, up),
                "bp": (f"MAX({cell('lo')},0)", max(lo, 0)),
                "bn": (f"MIN({cell('hi')},0)", min(hi, 0)),
                "ip": (f"IF({cell('up')}=1,MAX({cell('hi')},0)-MAX({cell('lo')},0),0)", pos_part if up else 0),
                "in": (f"IF({cell('up')}=1,MIN({cell('lo')},0)-MIN({cell('hi')},0),0)", neg_part if up else 0),
                "dp": (f"IF({cell('up')}=0,MAX({cell('hi')},0)-MAX({cell('lo')},0),0)", 0 if up else pos_part),
                "dn": (f"IF({cell('up')}=0,MIN({cell('lo')},0)-MIN({cell('hi')},0),0)", 0 if up else neg_part),
                "tot": (cell("cum") if is_tot else "0", cum if is_tot else 0),
            }
            for k, (f, cv) in vals.items():
                ws.write_formula((r - 1, C[k] - 1), f, cv, "number")
        if hide_helpers:
            ws.set_column(f"{L(C['cum'])}:{L(C['tot'])}", hidden=True)
        rng = lambda k: abs_ref(vsheet, hr, C[k], hr + n - 1, C[k])
        ch = self.add_chart({"type": "column", "subtype": "stacked", "gap": gap,
                             "series_lines": ({"color": "808080", "width": 0.75, "dash": "dash"}
                                              if connector_lines else None)})
        spec = [("Base+", "bp", None), ("Base-", "bn", None),
                ("Increase", "ip", increase_color), ("Increase ", "in", increase_color),
                ("Decrease", "dp", decrease_color), ("Decrease ", "dn", decrease_color),
                ("Total", "tot", total_color)]
        for nm, k, colr in spec:
            ch.add_series({"name": nm.strip(), "categories": lf, "values": rng(k),
                           "fill": {"color": colr} if colr else {"none": True}})
        ch.set_legend({"position": "none", "delete_series": [0, 1, 3, 5]})
        ch.set_x_axis({"label_position": "low"})
        if hide_helpers:
            ch.show_hidden_data(True)
        if title:
            ch.set_title({"name": title})
        return ch


    # --- save ------------------------------------------------------------------
    def save(self, path):
        if not self.sheets:
            self.add_worksheet()
        self._xlnames = {}
        files = {}
        ct = ['<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
              '<Default Extension="xml" ContentType="application/xml"/>',
              '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
              '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
              '<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>']
        chart_n = chartex_n = drawing_n = 0
        for si, ws in enumerate(self.sheets, 1):
            has_drawing = bool(ws.charts)
            files[f"xl/worksheets/sheet{si}.xml"] = ws._xml(si == 1, has_drawing)
            ct.append(f'<Override PartName="/xl/worksheets/sheet{si}.xml" '
                      f'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
            if not has_drawing:
                continue
            drawing_n += 1
            files[f"xl/worksheets/_rels/sheet{si}.xml.rels"] = (
                f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                f'<Relationship Id="rId1" Type="{REL}/drawing" Target="../drawings/drawing{drawing_n}.xml"/></Relationships>')
            anchors, rels = [], []
            for k, (chart, r0, c0, o) in enumerate(ws.charts, 1):
                rid = f"rId{k}"
                pos = ws._anchor_xml(r0, c0, o)
                cx, cy = o["width"] * 9525, o["height"] * 9525
                if isinstance(chart, ChartEx):
                    chartex_n += 1
                    files[f"xl/charts/chartEx{chartex_n}.xml"] = chart._xml()
                    ct.append(f'<Override PartName="/xl/charts/chartEx{chartex_n}.xml" ContentType="application/vnd.ms-office.chartex+xml"/>')
                    rels.append(f'<Relationship Id="{rid}" Type="http://schemas.microsoft.com/office/2014/relationships/chartEx" Target="../charts/chartEx{chartex_n}.xml"/>')
                    pfx, pns = chart.requires
                    anchors.append(
                        f'<xdr:twoCellAnchor editAs="oneCell">{pos}'
                        f'<mc:AlternateContent xmlns:mc="{NS_MC}">'
                        f'<mc:Choice xmlns:{pfx}="{pns}" Requires="{pfx}">'
                        f'<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="{k + 1}" name="Chart {k}"/>'
                        f'<xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>'
                        f'<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>'
                        f'<a:graphic><a:graphicData uri="{NS_CX}"><cx:chart xmlns:cx="{NS_CX}" r:id="{rid}"/>'
                        f'</a:graphicData></a:graphic></xdr:graphicFrame></mc:Choice>'
                        f'<mc:Fallback><xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="{k + 1}" name="Chart {k}"/>'
                        f'<xdr:cNvSpPr><a:spLocks noTextEdit="1"/></xdr:cNvSpPr></xdr:nvSpPr>'
                        f'<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
                        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:prstClr val="white"/></a:solidFill>'
                        f'<a:ln w="1"><a:solidFill><a:prstClr val="green"/></a:solidFill></a:ln></xdr:spPr>'
                        f'<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip"/><a:lstStyle/>'
                        f'<a:p><a:r><a:rPr lang="en-US" sz="1100"/><a:t>This {chart.type.replace("_", " ")} chart needs '
                        f'Excel 2016 or later (Microsoft 365).</a:t></a:r></a:p></xdr:txBody></xdr:sp></mc:Fallback>'
                        f'</mc:AlternateContent><xdr:clientData/></xdr:twoCellAnchor>')
                else:
                    chart_n += 1
                    files[f"xl/charts/chart{chart_n}.xml"] = chart._xml()
                    ct.append(f'<Override PartName="/xl/charts/chart{chart_n}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>')
                    rels.append(f'<Relationship Id="{rid}" Type="{REL}/chart" Target="../charts/chart{chart_n}.xml"/>')
                    anchors.append(
                        f'<xdr:twoCellAnchor editAs="oneCell">{pos}'
                        f'<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="{k + 1}" name="Chart {k}"/>'
                        f'<xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>'
                        f'<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>'
                        f'<a:graphic><a:graphicData uri="{NS_C}"><c:chart xmlns:c="{NS_C}" r:id="{rid}"/>'
                        f'</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>')
            files[f"xl/drawings/drawing{drawing_n}.xml"] = (
                f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                f'<xdr:wsDr xmlns:xdr="{NS_XDR}" xmlns:a="{NS_A}" xmlns:r="{NS_R}">{"".join(anchors)}</xdr:wsDr>')
            files[f"xl/drawings/_rels/drawing{drawing_n}.xml.rels"] = (
                f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                f'{"".join(rels)}</Relationships>')
            ct.append(f'<Override PartName="/xl/drawings/drawing{drawing_n}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>')

        sheets = "".join(f'<sheet name="{_e(s.name)}" sheetId="{i}" r:id="rId{i}"/>'
                         for i, s in enumerate(self.sheets, 1))
        names = ""
        if self._xlnames:
            names = "<definedNames>" + "".join(
                f'<definedName name="{n}" hidden="1">{_e(f)}</definedName>'
                for f, n in self._xlnames.items()) + "</definedNames>"
        files["xl/workbook.xml"] = (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<workbook xmlns="{NS_MAIN}" xmlns:r="{NS_R}"><bookViews><workbookView/></bookViews>'
            f'<sheets>{sheets}</sheets>{names}<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>')
        n = len(self.sheets)
        wrels = "".join(f'<Relationship Id="rId{i}" Type="{REL}/worksheet" Target="worksheets/sheet{i}.xml"/>'
                        for i in range(1, n + 1))
        wrels += f'<Relationship Id="rId{n + 1}" Type="{REL}/styles" Target="styles.xml"/>'
        wrels += f'<Relationship Id="rId{n + 2}" Type="{REL}/theme" Target="theme/theme1.xml"/>'
        files["xl/_rels/workbook.xml.rels"] = (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{wrels}</Relationships>')
        files["xl/styles.xml"] = _STYLES
        files["xl/theme/theme1.xml"] = _THEME
        files["_rels/.rels"] = (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        files["[Content_Types].xml"] = (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            f'{"".join(ct)}</Types>')

        order = ["[Content_Types].xml", "_rels/.rels", "xl/workbook.xml", "xl/_rels/workbook.xml.rels"]
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            for name in order + sorted(k for k in files if k not in order):
                z.writestr(name, files[name])
        return path
