"""Ergebnis-Arbeitsmappe (Excel) mit Formeln, bedingter Formatierung und Diagrammen.

Rechenintensive Blätter (Pain-Point-Score, Wertstrom, Nutzwertanalyse, Business Case) enthalten
echte Excel-Formeln – Gewichte und Annahmen können im Management-Termin live verändert werden.
"""
from __future__ import annotations

from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference, ScatterChart, Series
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

from .analysis import management_summary
from .themes import resolve

THIN = Side(style="thin", color="D1D5DB")


class XL:
    def __init__(self, theme: dict | None):
        t = resolve(theme)
        self.primary, self.accent, self.light = t["primary"], t["accent"], t.get("light", "F3F4F6")
        self.wb = Workbook()
        self.wb.remove(self.wb.active)

    def sheet(self, title: str, heading: str, sub: str = ""):
        ws = self.wb.create_sheet(title[:31])
        ws["A1"] = heading
        ws["A1"].font = Font(size=15, bold=True, color=self.primary)
        if sub:
            ws["A2"] = sub
            ws["A2"].font = Font(italic=True, color="6B7280")
        ws.sheet_view.showGridLines = False
        return ws

    def table(self, ws, row: int, headers: list[str], rows: list[list[Any]], widths: list[int] | None = None,
              col: int = 1) -> tuple[int, int]:
        for j, h in enumerate(headers):
            c = ws.cell(row, col + j, h)
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor=self.primary)
            c.alignment = Alignment(wrap_text=True, vertical="center")
            c.border = Border(bottom=THIN)
        for i, r in enumerate(rows, start=1):
            for j, v in enumerate(r):
                if isinstance(v, list):
                    v = ", ".join(map(str, v))
                c = ws.cell(row + i, col + j, v)
                c.alignment = Alignment(wrap_text=True, vertical="top")
                c.border = Border(bottom=THIN)
                if i % 2 == 0:
                    c.fill = PatternFill("solid", fgColor=self.light)
        for j, w in enumerate(widths or []):
            ws.column_dimensions[L(col + j)].width = w
        ws.freeze_panes = ws.cell(row + 1, 1)
        return row + 1, row + len(rows)


def build_workbook(r: dict[str, Any], cfg: dict[str, Any], path: Path, summary: list[str] | None = None) -> Path:
    x = XL(cfg.get("theme"))
    _summary(x, r, cfg, summary or management_summary(r, cfg))
    if "stakeholder" in r:
        _stakeholder(x, r["stakeholder"])
    if "prozess" in r:
        _prozess(x, r["prozess"])
    if "vsm" in r:
        _vsm(x, r["vsm"])
    if "painpoints" in r:
        _painpoints(x, r["painpoints"])
    if "ursachen" in r and r["ursachen"]["whys"]:
        ws = x.sheet("5Whys", "Ursachenketten (5 Whys)")
        x.table(ws, 4, ["Pain Point", "Problem", "Warum 1", "Warum 2", "Warum 3", "Warum 4", "Warum 5", "Grundursache",
                        "Digital lösbar"],
                [[c["pain_point_id"], c["problem"], c.get("w1"), c.get("w2"), c.get("w3"), c.get("w4"), c.get("w5"),
                  c["grundursache"], c.get("digital_loesbar")] for c in r["ursachen"]["whys"]],
                [10, 30, 25, 25, 25, 25, 25, 30, 12])
    if "userstories" in r:
        _userstories(x, r["userstories"])
    if "handlungsfelder" in r:
        ws = x.sheet("Handlungsfelder", "Handlungsfelder & Potenzial")
        x.table(ws, 4, ["Handlungsfeld", "Stories", "Story Points", "Use Cases", "Pain Points", "Potenzial"],
                [[h["name"], h["stories"], h["story_points"], h["use_cases"], h["pain_points"], h["potenzial"]]
                 for h in r["handlungsfelder"]["items"]], [30, 10, 12, 20, 30, 12])
        _bar(ws, "Potenzial je Handlungsfeld", 6, 5, 4 + len(r["handlungsfelder"]["items"]), "H4", horizontal=True)
    if "usecases" in r:
        _usecases(x, r)
    if "fitgap" in r:
        _fitgap(x, r["fitgap"])
    if "nwa" in r:
        _nwa(x, r["nwa"])
    if "impact_effort" in r and r["impact_effort"]["items"]:
        _impact(x, r["impact_effort"])
    if "businesscase" in r:
        _businesscase(x, r["businesscase"])
    if "roadmap" in r:
        _roadmap(x, r["roadmap"])
    if "raci" in r:
        _raci(x, r["raci"])
    if r.get("issues"):
        ws = x.sheet("Datenqualität", "Hinweise der Plausibilitätsprüfung")
        x.table(ws, 4, ["Hinweis"], [[i] for i in r["issues"]], [120])
    path.parent.mkdir(parents=True, exist_ok=True)
    x.wb.save(path)
    return path


def _summary(x: XL, r: dict, cfg: dict, bullets: list[str]) -> None:
    ws = x.sheet("Übersicht", f"{cfg.get('projekt')} – {cfg.get('kunde')}", "Management Summary & Kennzahlen")
    row = 4
    for b in bullets:
        ws.cell(row, 1, "• " + b).alignment = Alignment(wrap_text=True)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
        ws.row_dimensions[row].height = 32
        row += 1
    row += 1
    rows = []
    for k, v in r.get("kennzahlen", {}).items():
        rows.append([k, v])
    x.table(ws, row, ["Kennzahl", "Wert"], rows, [45, 18, 18, 18, 18, 18])
    for i, (k, v) in enumerate(rows, start=row + 1):
        c = ws.cell(i, 2)
        c.number_format = '#,##0 "€"' if "EUR" in k else ("0%" if isinstance(v, float) and v <= 1.5 else "0")


def _stakeholder(x: XL, s: dict) -> None:
    ws = x.sheet("Stakeholder", "Stakeholder-Matrix", "Quadrant per Formel: Einfluss/Interesse ≥ 3 = hoch")
    rows = []
    for i, it in enumerate(s["items"], start=5):
        rows.append([it["name"], it["rolle"], it.get("bereich"), it["einfluss"], it["interesse"], it.get("haltung"),
                     f'=IF(AND(D{i}>=3,E{i}>=3),"Eng einbinden",IF(D{i}>=3,"Zufriedenstellen",'
                     f'IF(E{i}>=3,"Informieren","Beobachten")))', it.get("massnahme")])
    a, b = x.table(ws, 4, ["Name", "Rolle", "Bereich", "Einfluss", "Interesse", "Haltung", "Quadrant", "Maßnahme"],
                   rows, [22, 22, 16, 10, 10, 13, 18, 40])
    ch = ScatterChart()
    ch.title, ch.style = "Einfluss/Interesse", 13
    ch.x_axis.title, ch.y_axis.title = "Interesse", "Einfluss"
    ch.x_axis.scaling.min = ch.y_axis.scaling.min = 0
    ch.x_axis.scaling.max = ch.y_axis.scaling.max = 5.5
    ser = Series(Reference(ws, min_col=4, min_row=a, max_row=b), Reference(ws, min_col=5, min_row=a, max_row=b),
                 title="Stakeholder")
    ser.marker.symbol, ser.graphicalProperties.line.noFill = "circle", True
    ch.series.append(ser)
    ch.height, ch.width = 9, 12
    ws.add_chart(ch, "J4")


def _prozess(x: XL, p: dict) -> None:
    ws = x.sheet("Prozesse", "Ist-Prozesse: Kennzahlen und Medienbrüche")
    rows = [[n, v["anzahl_schritte"], v["uebergaben"], len(v["medienbrueche"]), v["digitalgrad"]]
            for n, v in p["prozesse"].items()]
    a, b = x.table(ws, 4, ["Prozess", "Schritte", "Übergaben", "Medienbrüche", "Digitalgrad"], rows, [32, 12, 12, 14, 12])
    for i in range(a, b + 1):
        ws.cell(i, 5).number_format = "0%"
    row = b + 3
    rows = [[n, m["von"], m["nach"], m["art"]] for n, v in p["prozesse"].items() for m in v["medienbrueche"]]
    if rows:
        ws.cell(row - 1, 1, "Medienbrüche").font = Font(bold=True, color=x.primary)
        x.table(ws, row, ["Prozess", "Von", "Nach", "Art"], rows)


def _vsm(x: XL, v: dict) -> None:
    for name, p in v["prozesse"].items():
        ws = x.sheet(f"VSM {name}"[:31], f"Wertstrom: {name}", "Kennzahlen als Formeln – Zeiten anpassbar")
        rows = []
        for i, s in enumerate(p["steps"], start=5):
            rows.append([s["nr"], s["schritt"], s.get("bz") or 0, s.get("rz") or 0, s.get("wz") or 0,
                         f"=C{i}+D{i}+E{i}", (s.get("qr") if s.get("qr") is not None else 100) / 100,
                         s.get("wertschoepfend") or "Ja", s.get("medium"), s.get("system")])
        a, b = x.table(ws, 4, ["Nr", "Schritt", "Bearbeitung [min]", "Rüsten [min]", "Warten [min]", "Summe [min]",
                               "FPY", "Wertschöpfend", "Medium", "System"], rows, [6, 30, 14, 12, 12, 12, 8, 13, 18, 18])
        for i in range(a, b + 1):
            ws.cell(i, 7).number_format = "0%"
        k = b + 2
        kpis = [("Durchlaufzeit [min]", f"=SUM(F{a}:F{b})", "0"),
                ("Wertschöpfende Zeit [min]", f'=SUMIF(H{a}:H{b},"Ja",C{a}:C{b})', "0"),
                ("Flussgrad", f"=B{k + 1}/B{k}", "0.0%"),
                ("Rolled Throughput Yield", f"=PRODUCT(G{a}:G{b})", "0.0%"),
                ("Wartezeitanteil", f"=SUM(E{a}:E{b})/B{k}", "0.0%")]
        for j, (lab, f, fmt) in enumerate(kpis):
            ws.cell(k + j, 1, lab).font = Font(bold=True)
            c = ws.cell(k + j, 2, f)
            c.number_format, c.font = fmt, Font(bold=True, color=x.accent)
        ch = BarChart()
        ch.type, ch.grouping, ch.overlap = "col", "stacked", 100
        ch.title = "Zeitanteile je Schritt [min]"
        ch.add_data(Reference(ws, min_col=3, max_col=5, min_row=4, max_row=b), titles_from_data=True)
        ch.set_categories(Reference(ws, min_col=2, min_row=a, max_row=b))
        ch.height, ch.width = 9, 16
        ws.add_chart(ch, "L4")


def _painpoints(x: XL, p: dict) -> None:
    ws = x.sheet("PainPoints", "Pain Points – priorisiert", "Score = Häufigkeit × Auswirkung (Formel)")
    rows = []
    for i, it in enumerate(p["items"], start=5):
        rows.append([it["id"], it.get("prozess"), it.get("schritt"), it["beschreibung"], it.get("kategorie"),
                     it.get("haeufigkeit") or 3, it.get("auswirkung") or 3, f"=F{i}*G{i}", it["stimmen"],
                     f"=RANK(H{i},$H$5:$H${4 + len(p['items'])})"])
    a, b = x.table(ws, 4, ["ID", "Prozess", "Schritt", "Beschreibung", "Kategorie", "Häufigkeit", "Auswirkung", "Score",
                           "Stimmen", "Rang"], rows, [8, 18, 20, 50, 15, 10, 10, 8, 8, 7])
    ws.conditional_formatting.add(f"H{a}:H{b}", ColorScaleRule(start_type="num", start_value=1, start_color="FCF3CF",
                                                               end_type="num", end_value=25, end_color="E74C3C"))
    ws2 = x.sheet("Pareto", "Pareto der Pain-Point-Kategorien")
    rows = []
    n = len(p["pareto"])
    for i, it in enumerate(p["pareto"], start=5):
        rows.append([it["kategorie"], f'=SUMIF(PainPoints!$E$5:$E${b},A{i},PainPoints!$H$5:$H${b})',
                     f'=COUNTIF(PainPoints!$E$5:$E${b},A{i})', f"=SUM($B$5:B{i})/SUM($B$5:$B${4 + n})"])
    a2, b2 = x.table(ws2, 4, ["Kategorie", "Score", "Anzahl", "Kumuliert"], rows, [22, 10, 10, 12])
    for i in range(a2, b2 + 1):
        ws2.cell(i, 4).number_format = "0%"
    bar = BarChart()
    bar.title = "Pareto"
    bar.add_data(Reference(ws2, min_col=2, min_row=4, max_row=b2), titles_from_data=True)
    bar.set_categories(Reference(ws2, min_col=1, min_row=a2, max_row=b2))
    line = LineChart()
    line.add_data(Reference(ws2, min_col=4, min_row=4, max_row=b2), titles_from_data=True)
    line.y_axis.axId = 200
    line.y_axis.scaling.max = 1
    line.y_axis.number_format = "0%"
    line.y_axis.crosses = "max"
    bar += line
    bar.height, bar.width = 9, 16
    ws2.add_chart(bar, "F4")


def _userstories(x: XL, u: dict) -> None:
    ws = x.sheet("UserStories", "User Stories mit INVEST-Check")
    rows = []
    for s in u["items"]:
        ch = s["invest"]["checks"]
        rows.append([s["id"], s["invest"]["satz"], s.get("akzeptanz"), s.get("aktivitaet"), s.get("release"),
                     s.get("story_points")] + ["✓" if ch[k] else "✗" for k in "INVEST"] +
                    [s["invest"]["score"], "; ".join(s["invest"]["hinweise"])])
    a, b = x.table(ws, 4, ["ID", "User Story", "Akzeptanzkriterien", "Aktivität", "Release", "SP", "I", "N", "V", "E",
                           "S", "T", "INVEST", "Hinweise"], rows, [8, 55, 50, 18, 10, 6, 4, 4, 4, 4, 4, 4, 9, 50])
    for i in range(a, b + 1):
        ws.cell(i, 13).number_format = "0%"
    ws.conditional_formatting.add(f"G{a}:L{b}", CellIsRule(operator="equal", formula=['"✗"'],
                                                           font=Font(color="C0392B", bold=True)))
    ws.conditional_formatting.add(f"G{a}:L{b}", CellIsRule(operator="equal", formula=['"✓"'],
                                                           font=Font(color="2E8B57", bold=True)))


def _usecases(x: XL, r: dict) -> None:
    ws = x.sheet("UseCases", "Use-Case-Steckbriefe (Übersicht)")
    ie = {i["id"]: i for i in r.get("impact_effort", {}).get("items", [])}
    rows = [[u["id"], u["titel"], u.get("handlungsfeld"), u.get("problem"), u.get("loesung"), u.get("nutzen"),
             u.get("datenbasis"), u.get("systeme"), u.get("aufwand"), u.get("owner"), u.get("pain_point_ids"),
             ie.get(u["id"], {}).get("quadrant")] for u in r["usecases"]["items"]]
    x.table(ws, 4, ["ID", "Titel", "Handlungsfeld", "Problem", "Lösung", "Nutzen", "Datenbasis", "Systeme", "Aufwand",
                    "Owner", "Pain Points", "Quadrant"], rows, [7, 28, 20, 35, 35, 35, 25, 18, 25, 15, 14, 15])


def _fitgap(x: XL, f: dict) -> None:
    ws = x.sheet("FitGap", "Fit-Gap-Analyse", "Auswertung per ZÄHLENWENNS – Liste anpassbar")
    rows = [[i["id"], i["anforderung"], i.get("prozess"), i.get("prioritaet"), i.get("system"), i.get("abdeckung"),
             i.get("loesung"), i.get("aufwand_pt")] for i in f["items"]]
    a, b = x.table(ws, 4, ["ID", "Anforderung", "Prozess", "Priorität", "System", "Abdeckung", "Lösungsansatz", "PT"],
                   rows, [8, 45, 18, 10, 14, 14, 35, 8])
    col = {"Standard": "D5F5E3", "Konfiguration": "E8F6EF", "Erweiterung": "FCF3CF", "Gap": "FADBD8"}
    for k, c in col.items():
        ws.conditional_formatting.add(f"F{a}:F{b}", CellIsRule(operator="equal", formula=[f'"{k}"'],
                                                               fill=PatternFill("solid", fgColor=c)))
    start = b + 3
    ws.cell(start - 1, 1, "Auswertung").font = Font(bold=True, color=x.primary)
    heads = ["Priorität"] + f["kategorien"] + ["Summe"]
    rows = []
    for k, p in enumerate(["Muss", "Soll", "Kann"]):
        rr = start + 1 + k
        rows.append([p] + [f'=COUNTIFS($D${a}:$D${b},$A{rr},$F${a}:$F${b},"{c}")' for c in f["kategorien"]]
                    + [f"=SUM(B{rr}:E{rr})"])
    rows.append(["Aufwand [PT]"] + [f'=SUMIF($F${a}:$F${b},"{c}",$H${a}:$H${b})' for c in f["kategorien"]]
                + [f"=SUM(B{start + 4}:E{start + 4})"])
    rows.append(["Fit-Quote", f"=(SUM(B{start + 1}:C{start + 3}))/SUM(F{start + 1}:F{start + 3})"])
    x.table(ws, start, heads, rows)
    ws.cell(start + 5, 2).number_format = "0%"
    ch = BarChart()
    ch.grouping, ch.overlap, ch.title = "stacked", 100, "Abdeckung nach Priorität"
    ch.add_data(Reference(ws, min_col=2, max_col=5, min_row=start, max_row=start + 3), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=start + 1, max_row=start + 3))
    ws.add_chart(ch, f"J{start}")


def _nwa(x: XL, n: dict) -> None:
    ws = x.sheet("Nutzwertanalyse", "Nutzwertanalyse (Scoring-Modell)",
                 "Gewichte in Zeile 5 ändern → Nutzwert und Rang rechnen live neu")
    crit = n["kriterien"]
    k = len(crit)
    ws.cell(4, 1, "Kriterium").font = Font(bold=True)
    ws.cell(5, 1, "Gewicht").font = Font(bold=True)
    ws.cell(6, 1, "Kategorie").font = Font(bold=True)
    for j, c in enumerate(crit, start=3):
        h = ws.cell(4, j, c["kriterium"])
        h.font, h.fill = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor=x.primary)
        h.alignment = Alignment(wrap_text=True)
        w = ws.cell(5, j, round(c["w"], 4))
        w.number_format, w.fill = "0%", PatternFill("solid", fgColor="FFF4E0")
        ws.cell(6, j, c["kategorie"])
        ws.column_dimensions[L(j)].width = 15
    last = L(2 + k)
    ws.cell(5, 3 + k, f"=SUM(C5:{last}5)").number_format = "0%"
    ws.cell(4, 3 + k, "Σ / Nutzwert").font = Font(bold=True)
    ws.cell(4, 4 + k, "in %").font = Font(bold=True)
    ws.cell(4, 5 + k, "Rang").font = Font(bold=True)
    ws.cell(7, 1, "Use Case").font = Font(bold=True)
    first = 8
    lastrow = first + len(n["items"]) - 1
    for i, it in enumerate(n["items"], start=first):
        ws.cell(i, 1, it["id"])
        ws.cell(i, 2, it["titel"])
        for j, c in enumerate(crit, start=3):
            ws.cell(i, j, it["werte"].get(c["kriterium"]))
        ws.cell(i, 3 + k, f"=SUMPRODUCT($C$5:${last}$5,C{i}:{last}{i})/$" + L(3 + k) + "$5").number_format = "0.00"
        ws.cell(i, 4 + k, f"={L(3 + k)}{i}/5").number_format = "0%"
        ws.cell(i, 5 + k, f"=RANK({L(3 + k)}{i},${L(3 + k)}${first}:${L(3 + k)}${lastrow})")
    ws.column_dimensions["A"].width, ws.column_dimensions["B"].width = 9, 34
    ws.conditional_formatting.add(f"C{first}:{last}{lastrow}", ColorScaleRule(
        start_type="num", start_value=1, start_color="F5B7B1", mid_type="num", mid_value=3, mid_color="FCF3CF",
        end_type="num", end_value=5, end_color="ABEBC6"))
    ws.conditional_formatting.add(f"{L(5 + k)}{first}:{L(5 + k)}{lastrow}",
                                  CellIsRule(operator="lessThanOrEqual", formula=["3"],
                                             fill=PatternFill("solid", fgColor="D5F5E3"), font=Font(bold=True)))
    ch = BarChart()
    ch.type, ch.title = "bar", "Nutzwert in %"
    ch.add_data(Reference(ws, min_col=4 + k, min_row=7, max_row=lastrow), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=2, min_row=first, max_row=lastrow))
    ch.y_axis.number_format = "0%"
    ch.x_axis.scaling.orientation = "maxMin"
    ch.height, ch.width = 10, 16
    ws.add_chart(ch, f"{L(7 + k)}4")
    r0 = lastrow + 3
    ws.cell(r0 - 1, 1, "Sensitivität (Rangspanne bei ±20 % je Gewicht)").font = Font(bold=True, color=x.primary)
    x.table(ws, r0, ["Use Case", "Titel", "Rang", "bester Rang", "schlechtester Rang"],
            [[i["id"], i["titel"], i["rang"], i["rang_min"], i["rang_max"]] for i in n["items"]])


def _impact(x: XL, ie: dict) -> None:
    ws = x.sheet("ImpactEffort", "Impact/Effort-Matrix", "Quadrant per Formel (Schwelle 5,5)")
    rows = []
    for i, it in enumerate(ie["items"], start=5):
        rows.append([it["id"], it["titel"], it["impact"], it["effort"],
                     f'=IF(C{i}>=5.5,IF(D{i}<5.5,"Quick Win","Großes Vorhaben"),IF(D{i}<5.5,"Lückenfüller","Nicht machen"))',
                     it["quelle"]])
    a, b = x.table(ws, 4, ["ID", "Use Case", "Impact", "Effort", "Quadrant", "Quelle"], rows, [8, 34, 9, 9, 18, 16])
    for k, c in {"Quick Win": "D5F5E3", "Großes Vorhaben": "D6EAF8", "Lückenfüller": "FCF3CF",
                 "Nicht machen": "FADBD8"}.items():
        ws.conditional_formatting.add(f"E{a}:E{b}", CellIsRule(operator="equal", formula=[f'"{k}"'],
                                                               fill=PatternFill("solid", fgColor=c)))
    ch = ScatterChart()
    ch.title, ch.style = "Impact/Effort", 13
    ch.x_axis.title, ch.y_axis.title = "Effort", "Impact"
    for ax in (ch.x_axis, ch.y_axis):
        ax.scaling.min, ax.scaling.max = 0, 10
    s = Series(Reference(ws, min_col=3, min_row=a, max_row=b), Reference(ws, min_col=4, min_row=a, max_row=b),
               title="Use Cases")
    s.marker.symbol, s.graphicalProperties.line.noFill = "circle", True
    ch.series.append(s)
    ws.add_chart(ch, "H4")


def _businesscase(x: XL, bc: dict) -> None:
    ws = x.sheet("BusinessCase", "Business Case & ROI",
                 "Kalkulationszins in B4 ändern → Kapitalwert, IRR und Amortisation rechnen neu")
    ws["A4"], ws["B4"] = "Kalkulationszins", bc["zinssatz"]
    ws["B4"].number_format, ws["B4"].fill = "0.0%", PatternFill("solid", fgColor="FFF4E0")
    n = bc["jahre"]
    heads = ["Use Case", "Zeile"] + [f"Jahr {t}" for t in range(n + 1)]
    row = 6
    summary = []
    for uc, u in bc["use_cases"].items():
        costs = [0.0] * (n + 1)
        bens = [0.0] * (n + 1)
        for p in u["positionen"]:
            amt = abs(p.get("betrag") or 0)
            if p["typ"] in ("Investition", "Laufende Kosten"):
                costs[p.get("jahr") or 0] += amt
            else:
                bens[p.get("jahr") or 0] += amt
        r1, r2, r3, r4 = row + 1, row + 2, row + 3, row + 4
        rows = [[uc, "Kosten"] + costs, [uc, "Nutzen"] + bens,
                [uc, "Netto-Cashflow"] + [f"={L(3 + t)}{r2}-{L(3 + t)}{r1}" for t in range(n + 1)],
                [uc, "Kumuliert"] + [f"=SUM($C{r3}:{L(3 + t)}{r3})" for t in range(n + 1)]]
        x.table(ws, row, heads, rows, [14, 16] + [13] * (n + 1))
        for rr in (r1, r2, r3, r4):
            for t in range(n + 1):
                ws.cell(rr, 3 + t).number_format = '#,##0 "€";[Red]-#,##0 "€"'
        summary.append((uc, r1, r2, r3, r4))
        row += 7
    ws.cell(row, 1, "Kennzahlen").font = Font(bold=True, color=x.primary)
    last = L(3 + n)
    rows = []
    for uc, r1, r2, r3, r4 in summary:
        rows.append([uc, f"=C{r3}+NPV($B$4,D{r3}:{last}{r3})", f"=(SUM(C{r2}:{last}{r2})-SUM(C{r1}:{last}{r1}))/SUM(C{r1}:{last}{r1})",
                     f'=IFERROR(IRR(C{r3}:{last}{r3}),"–")',
                     f'=IFERROR(MATCH(TRUE,INDEX(C{r4}:{last}{r4}>=0,0),0)-1,"> Horizont")'])
    a, b = x.table(ws, row + 1, ["Use Case", "Kapitalwert", "ROI", "IRR", "Break-even (Jahr)"], rows)
    for i in range(a, b + 1):
        ws.cell(i, 2).number_format = '#,##0 "€";[Red]-#,##0 "€"'
        ws.cell(i, 3).number_format = ws.cell(i, 4).number_format = "0%"


def _roadmap(x: XL, rm: dict) -> None:
    ws = x.sheet("Roadmap", "Roadmap (Gantt nach Quartalen)", "* = automatisch vorgeschlagen")
    qs = rm["quartale"]
    heads = ["ID", "Use Case", "Horizont", "Start", "Dauer [Mon.]", "MVP", "Pilot", "KPIs"] + qs
    rows = []
    for it in rm["items"]:
        rows.append([it["id"], it["titel"] + (" *" if it["vorschlag"] else ""), it["horizont"], it["start"], it["dauer"],
                     it.get("mvp"), it.get("pilot"), it.get("kpi")] +
                    ["■" if it["start_idx"] <= k < it["ende_idx"] else "" for k in range(len(qs))])
    a, b = x.table(ws, 4, heads, rows, [7, 32, 9, 9, 8, 30, 20, 25] + [8] * len(qs))
    rng = f"I{a}:{L(8 + len(qs))}{b}"
    for hz, c in {"Now": "00AFE9", "Next": "004F9F", "Later": "9CA3AF"}.items():
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'AND(I{a}="■",$C{a}="{hz}")'],
                                                       fill=PatternFill("solid", fgColor=c),
                                                       font=Font(color=c)))


def _raci(x: XL, r: dict) -> None:
    ws = x.sheet("RACI", "RACI-Matrix", "Prüfspalte: genau ein A und mindestens ein R je Zeile")
    roles = r["rollen"]
    rows = []
    last = L(1 + len(roles))
    for i, it in enumerate(r["items"], start=5):
        rows.append([it["aufgabe"]] + [it["werte"].get(k, "") for k in roles] +
                    [f'=IF(AND(COUNTIF(B{i}:{last}{i},"*A*")=1,COUNTIF(B{i}:{last}{i},"*R*")>=1),"✓","prüfen")'])
    a, b = x.table(ws, 4, ["Aufgabe"] + roles + ["Prüfung"], rows, [36] + [14] * len(roles) + [10])
    rng = f"B{a}:{last}{b}"
    for k, c in {"A": "F5B7B1", "R": "D6EAF8", "C": "FCF3CF", "I": "E5E7EB"}.items():
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT(B{a},1)="{k}"'],
                                                       fill=PatternFill("solid", fgColor=c)))
    pc = L(2 + len(roles))
    ws.conditional_formatting.add(f"{pc}{a}:{pc}{b}", CellIsRule(operator="equal", formula=['"prüfen"'],
                                                                 font=Font(color="C0392B", bold=True)))


def _bar(ws, title: str, val_col: int, cat_col: int, max_row: int, anchor: str, horizontal=False) -> None:
    ch = BarChart()
    ch.title = title
    if horizontal:
        ch.type = "bar"
        ch.x_axis.scaling.orientation = "maxMin"
    ch.add_data(Reference(ws, min_col=val_col, min_row=4, max_row=max_row), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=5, max_row=max_row))
    ws.add_chart(ch, anchor)
