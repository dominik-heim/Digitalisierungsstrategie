"""Zeichen-Bausteine für PowerPoint (python-pptx): Folien, Tabellen, Diagramme, Matrizen, Pfeile."""
from __future__ import annotations

import math
from pathlib import Path
from typing import Any, Callable, Iterable

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_TICK_LABEL_POSITION
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

GREY = "6B7280"
LIGHT = "F3F4F6"
MID = "D1D5DB"
DARK = "1F2937"
WHITE = "FFFFFF"
RED = "C0392B"
GREEN = "2E8B57"
AMBER = "E0A21F"
BLUE = "2F6DB5"
PALETTE = ["1F3A5F", "E07A1F", "2E8B57", "2F6DB5", "8E44AD", "C0392B", "16A085", "7F8C8D"]


def rgb(hexstr: str) -> RGBColor:
    return RGBColor.from_string(hexstr.replace("#", "").upper())


def fit_size(text: str, w_in: float, h_in: float, max_pt: int = 14, min_pt: int = 7) -> int:
    """Grobe Schätzung der Schriftgröße, damit Text in eine Box passt."""
    text = text or ""
    for pt in range(max_pt, min_pt - 1, -1):
        cw = pt * 0.55 / 72  # mittlere Zeichenbreite in Zoll
        lh = pt * 1.25 / 72
        per_line = max(1, int(w_in / cw))
        lines = sum(max(1, math.ceil(len(par) / per_line)) for par in text.split("\n"))
        if lines * lh <= h_in:
            return pt
    return min_pt


class Deck:
    def __init__(self, theme: dict[str, Any] | None = None, footer: str = "") -> None:
        from .themes import resolve
        theme = resolve(theme)
        tpl = theme.get("pptx_template")
        self.prs = Presentation(tpl) if tpl and Path(tpl).exists() else Presentation()
        if not (tpl and Path(tpl).exists()):
            self.prs.slide_width, self.prs.slide_height = Inches(13.333), Inches(7.5)
        self.W = self.prs.slide_width / 914400
        self.H = self.prs.slide_height / 914400
        self.primary = theme.get("primary", "1F3A5F")
        self.accent = theme.get("accent", "E07A1F")
        self.font = theme.get("font", "Calibri")
        self.secondary = theme.get("secondary", BLUE)
        self.light = theme.get("light", LIGHT)
        self.palette = [self.primary, self.accent, self.secondary, theme.get("tertiary", "16A085"),
                        theme.get("soft", "8E44AD")] + PALETTE[5:]
        self.logo = theme.get("logo") or ""
        self.footer = footer
        self.blank = min(self.prs.slide_layouts, key=lambda l: len(l.placeholders))
        self.n = 0

    # ------------------------------------------------------------ Grundfolien
    def slide(self, title: str, subtitle: str | None = None, source: str | None = None):
        s = self.prs.slides.add_slide(self.blank)
        self.n += 1
        tsz = fit_size(title, self.W - 1.0, 0.5, 24, 15)
        self.text(s, title, 0.5, 0.3, self.W - 1.0, 0.6, size=tsz, bold=True, color=self.primary)
        if subtitle:
            self.text(s, subtitle, 0.5, 0.88, self.W - 1.0, 0.4, size=13, color=GREY)
        self.rect(s, 0.5, 1.3, 1.2, 0.05, fill=self.accent)
        self.text(s, self.footer, 0.5, self.H - 0.4, self.W - 2, 0.3, size=9, color=GREY)
        self.text(s, str(self.n), self.W - 1.0, self.H - 0.4, 0.5, 0.3, size=9, color=GREY, align="r")
        if source:
            self.text(s, f"Quelle: {source}", 0.5, self.H - 0.65, self.W - 1, 0.25, size=8, color=GREY)
        if self.logo and Path(self.logo).exists():
            s.shapes.add_picture(self.logo, Inches(self.W - 1.9), Inches(0.3), height=Inches(0.5))
        return s

    def title_slide(self, title: str, subtitle: str, meta: str):
        s = self.prs.slides.add_slide(self.blank)
        self.n += 1
        self.rect(s, 0, 0, self.W, self.H, fill=self.primary)
        self.rect(s, 0.8, 3.05, 1.6, 0.08, fill=self.accent)
        self.text(s, title, 0.8, 1.6, self.W - 1.6, 1.4, size=40, bold=True, color=WHITE)
        self.text(s, subtitle, 0.8, 3.3, self.W - 1.6, 0.8, size=20, color=WHITE)
        self.text(s, meta, 0.8, self.H - 1.2, self.W - 1.6, 0.5, size=12, color="C9D3E0")
        return s

    def section(self, title: str, subtitle: str = ""):
        s = self.prs.slides.add_slide(self.blank)
        self.n += 1
        self.rect(s, 0, 0, 0.35, self.H, fill=self.accent)
        self.text(s, title, 1.0, 2.6, self.W - 2, 1.0, size=34, bold=True, color=self.primary)
        if subtitle:
            self.text(s, subtitle, 1.0, 3.6, self.W - 2, 1.2, size=16, color=GREY)
        return s

    # ------------------------------------------------------------ Formen
    def rect(self, s, x, y, w, h, fill=None, line=None, shape=MSO_SHAPE.RECTANGLE, text="", size=11,
             color=DARK, bold=False, align="c", anchor="m", line_w=0.75, autosize=False):
        shp = s.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
        if fill:
            shp.fill.solid()
            shp.fill.fore_color.rgb = rgb(fill)
        else:
            shp.fill.background()
        if line:
            shp.line.color.rgb = rgb(line)
            shp.line.width = Pt(line_w)
        else:
            shp.line.fill.background()
        shp.shadow.inherit = False
        tf = shp.text_frame
        tf.word_wrap = True
        for side in ("margin_left", "margin_right"):
            setattr(tf, side, Inches(0.06))
        tf.margin_top = tf.margin_bottom = Inches(0.03)
        if text:
            sz = fit_size(text, w - 0.12, h - 0.06, size, 6) if autosize else size
            self._fill_text(tf, text, sz, color, bold, align)
        tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}[anchor]
        return shp

    def text(self, s, text, x, y, w, h, size=12, color=DARK, bold=False, align="l", anchor="t",
             autosize=False, italic=False):
        tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Inches(0.02)
        tf.margin_top = tf.margin_bottom = Inches(0.02)
        sz = fit_size(text, w, h, size, 7) if autosize else size
        self._fill_text(tf, text or "", sz, color, bold, align, italic)
        tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}[anchor]
        return tb

    def _fill_text(self, tf, text, size, color, bold, align, italic=False):
        al = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[align]
        for i, line in enumerate(str(text).split("\n")):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = al
            r = p.add_run()
            r.text = line
            f = r.font
            f.size, f.bold, f.italic, f.name = Pt(size), bold, italic, self.font
            f.color.rgb = rgb(color)

    def bullets(self, s, items: Iterable[str], x, y, w, h, size=14, color=DARK, bullet="•"):
        items = [i for i in items if i]
        sz = fit_size("\n".join(items), w - 0.3, h, size, 8)
        tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        for i, it in enumerate(items):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.space_after = Pt(sz * 0.5)
            r = p.add_run()
            r.text = f"{bullet} {it}" if bullet else it
            r.font.size, r.font.name = Pt(sz), self.font
            r.font.color.rgb = rgb(color)
        return tb

    def arrow(self, s, x1, y1, x2, y2, color=GREY, width=1.25, dash=False, head=True):
        c = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
        c.line.color.rgb = rgb(color)
        c.line.width = Pt(width)
        ln = c.line._get_or_add_ln()
        if dash:
            pd = ln.makeelement(qn("a:prstDash"), {"val": "dash"})
            ln.append(pd)
        if head:
            ln.append(ln.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"}))
        return c

    def elbow(self, s, x1, y1, x2, y2, color=GREY, width=1.25):
        """Rechtwinklige Verbindung (horizontal - vertikal - horizontal) aus Einzellinien."""
        if abs(y1 - y2) < 0.01:
            return self.arrow(s, x1, y1, x2, y2, color, width)
        xm = x1 + (x2 - x1) / 2 if x2 > x1 + 0.2 else x1 + 0.12
        self.arrow(s, x1, y1, xm, y1, color, width, head=False)
        self.arrow(s, xm, y1, xm, y2, color, width, head=False)
        return self.arrow(s, xm, y2, x2, y2, color, width)

    def kpi(self, s, x, y, w, h, value: str, label: str, color=None):
        self.rect(s, x, y, w, h, fill=self.light)
        self.rect(s, x, y, 0.07, h, fill=color or self.accent)
        self.text(s, value, x + 0.2, y + 0.08, w - 0.3, h * 0.55, size=24, bold=True, color=self.primary)
        self.text(s, label, x + 0.2, y + h * 0.58, w - 0.3, h * 0.4, size=10, color=GREY, autosize=True)

    # ------------------------------------------------------------ Tabelle
    def table(self, s, rows: list[list[Any]], x, y, w, h=None, col_w: list[float] | None = None,
              size=10, header=True, fill_fn: Callable[[int, int, Any], str | None] | None = None,
              row_h=0.3, bold_first_col=False, color_fn: Callable[[int, int, Any], str | None] | None = None):
        nr, nc = len(rows), len(rows[0])
        h = h or row_h * nr
        gt = s.shapes.add_table(nr, nc, Inches(x), Inches(y), Inches(w), Inches(h))
        tbl = gt.table
        tblPr = tbl._tbl.tblPr
        style = tblPr.find(qn("a:tableStyleId"))
        if style is not None:
            style.text = "{5940675A-B579-460E-94D1-54222C63F5DA}"  # 'No Style, Table Grid'
        if col_w:
            tot = sum(col_w)
            for i, cw in enumerate(col_w):
                tbl.columns[i].width = Emu(int(Inches(w) * cw / tot))
        for i in range(nr):
            tbl.rows[i].height = Inches(h / nr)
        for i, row in enumerate(rows):
            for j, val in enumerate(row):
                cell = tbl.cell(i, j)
                txt = "" if val is None else (val if isinstance(val, str) else _fmt(val))
                cell.text = ""
                tf = cell.text_frame
                tf.word_wrap = True
                cell.margin_left = cell.margin_right = Inches(0.05)
                cell.margin_top = cell.margin_bottom = Inches(0.02)
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                is_head = header and i == 0
                fcol = WHITE if is_head else (color_fn(i, j, val) if color_fn else None) or DARK
                self._fill_text(tf, txt, size, fcol, is_head or (bold_first_col and j == 0), "l")
                fill = self.primary if is_head else (fill_fn(i, j, val) if fill_fn else None) or (
                    self.light if i % 2 == 0 else WHITE)
                cell.fill.solid()
                cell.fill.fore_color.rgb = rgb(fill)
        return tbl

    # ------------------------------------------------------------ Diagramme (nativ, editierbar)
    def bar_chart(self, s, categories, series: dict[str, list[float]], x, y, w, h, horizontal=False,
                  stacked=False, number_format='0;-0;;', legend=True, colors=None, title=None, labels=True):
        cd = CategoryChartData()
        cd.categories = list(categories)
        for name, vals in series.items():
            cd.add_series(name, [float(v or 0) for v in vals])
        if horizontal:
            ct = XL_CHART_TYPE.BAR_STACKED if stacked else XL_CHART_TYPE.BAR_CLUSTERED
        else:
            ct = XL_CHART_TYPE.COLUMN_STACKED if stacked else XL_CHART_TYPE.COLUMN_CLUSTERED
        ch = s.shapes.add_chart(ct, Inches(x), Inches(y), Inches(w), Inches(h), cd).chart
        self._style_chart(ch, legend and len(series) > 1, title)
        colors = colors or self.palette
        for i, ser in enumerate(ch.series):
            ser.format.fill.solid()
            ser.format.fill.fore_color.rgb = rgb(colors[i % len(colors)])
        if labels:
            pl = ch.plots[0]
            pl.has_data_labels = True
            pl.data_labels.number_format = number_format
            pl.data_labels.number_format_is_linked = False
            pl.data_labels.font.size = Pt(9)
        if horizontal:
            ch.category_axis.reverse_order = True
        ch.plots[0].gap_width = 60
        return ch

    def line_chart(self, s, categories, series: dict[str, list[float]], x, y, w, h, number_format='#,##0',
                   title=None):
        cd = CategoryChartData()
        cd.categories = list(categories)
        for name, vals in series.items():
            cd.add_series(name, [float(v or 0) for v in vals])
        ch = s.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, Inches(x), Inches(y), Inches(w), Inches(h), cd).chart
        self._style_chart(ch, True, title)
        ch.value_axis.tick_labels.number_format = number_format
        ch.value_axis.tick_labels.number_format_is_linked = False
        ch.category_axis.tick_label_position = XL_TICK_LABEL_POSITION.LOW
        for i, ser in enumerate(ch.series):
            ser.format.line.color.rgb = rgb(self.palette[i % len(self.palette)])
            ser.format.line.width = Pt(2.25)
            ser.smooth = False
        return ch

    def _style_chart(self, ch, legend: bool, title: str | None):
        ch.font.size = Pt(10)
        ch.font.name = self.font
        ch.has_legend = legend
        if legend:
            ch.legend.position = XL_LEGEND_POSITION.BOTTOM
            ch.legend.include_in_layout = False
        if title:
            ch.has_title = True
            ch.chart_title.text_frame.text = title
            ch.chart_title.text_frame.paragraphs[0].runs[0].font.size = Pt(12)
        else:
            ch.has_title = False
        try:
            ch.value_axis.has_major_gridlines = True
            ch.value_axis.major_gridlines.format.line.color.rgb = rgb("E5E7EB")
            ch.value_axis.format.line.fill.background()
        except (AttributeError, ValueError):
            pass

    # ------------------------------------------------------------ 2×2-Matrix
    def matrix(self, s, points: list[dict], x, y, w, h, x_label: str, y_label: str,
               quadrants: tuple[str, str, str, str], q_colors: tuple[str, str, str, str],
               lo=1.0, hi=5.0, mid=3.0, point_color: Callable[[dict], str] | None = None):
        """Zeichnet eine 2×2-Matrix. quadrants = (oben-links, oben-rechts, unten-links, unten-rechts).
        points: {'x','y','label'}"""
        fx = (mid - lo) / (hi - lo)
        mx, my = x + w * fx, y + h * (1 - fx)
        boxes = [(x, y, mx - x, my - y), (mx, y, x + w - mx, my - y),
                 (x, my, mx - x, y + h - my), (mx, my, x + w - mx, y + h - my)]
        for (bx, by, bw, bh), lab, col in zip(boxes, quadrants, q_colors):
            self.rect(s, bx, by, bw, bh, fill=col, line=WHITE, line_w=2)
            self.text(s, lab, bx + 0.08, by + 0.05, bw - 0.16, 0.3, size=11, bold=True, color=GREY)
        self.text(s, x_label + " →", x, y + h + 0.05, w, 0.3, size=11, bold=True, color=GREY, align="c")
        tb = self.text(s, y_label + " →", x - 0.55 - h / 2 + 0.15, y + h / 2 - 0.15, h, 0.3, size=11, bold=True,
                       color=GREY, align="c")
        tb.rotation = -90
        occupied: dict[tuple[int, int], int] = {}
        for p in points:
            px = x + (p["x"] - lo) / (hi - lo) * w
            py = y + h - (p["y"] - lo) / (hi - lo) * h
            key = (round(px, 1), round(py, 1))
            k = occupied.get(key, 0)
            occupied[key] = k + 1
            if k:  # identische Koordinaten: versetzt untereinander darstellen
                py += 0.3 * k
            col = point_color(p) if point_color else self.primary
            r = 0.13
            self.rect(s, px - r, py - r, 2 * r, 2 * r, fill=col, shape=MSO_SHAPE.OVAL, line=WHITE)
            lx = px + r + 0.03
            align = "l"
            if lx + 1.9 > x + w:
                lx, align = px - r - 1.93, "r"
            self.text(s, p["label"], lx, py - 0.14, 1.9, 0.28, size=9, color=DARK, align=align)


def _fmt(v: Any) -> str:
    if isinstance(v, float):
        if abs(v) >= 1000:
            return f"{v:,.0f}".replace(",", ".")
        return f"{v:.1f}".replace(".", ",") if v != int(v) else str(int(v))
    if isinstance(v, (list, tuple)):
        return ", ".join(map(str, v))
    return str(v)


def chunks(seq: list, n: int) -> list[list]:
    return [seq[i:i + n] for i in range(0, len(seq), n)] or [[]]


def eur(v: float | None) -> str:
    if v is None:
        return "–"
    return f"{v / 1000:,.0f} T€".replace(",", ".")


def pct(v: float | None) -> str:
    return "–" if v is None else f"{v:.0%}"
