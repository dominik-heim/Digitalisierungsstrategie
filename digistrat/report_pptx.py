"""Erzeugt das Strategie-Deck (PowerPoint) aus den Analyse-Ergebnissen."""
from __future__ import annotations

import datetime as dt
from collections import defaultdict
from pathlib import Path
from typing import Any

from pptx.enum.shapes import MSO_SHAPE

from .analysis import RELEASES, management_summary
from .model import PHASES, methods_by_phase
from .pptx_builder import (AMBER, BLUE, DARK, GREEN, GREY, LIGHT, MID, PALETTE, RED, WHITE, Deck, chunks,
                           eur, pct)

HZ_COL = {"Now": "00AFE9", "Next": "004F9F", "Later": "9CA3AF"}
Q_COL = {"Quick Win": GREEN, "Großes Vorhaben": BLUE, "Lückenfüller": AMBER, "Nicht machen": RED}
HALTUNG_COL = {"Promotor": GREEN, "Unterstützer": "7DBB91", "Neutral": GREY, "Skeptiker": AMBER, "Blockierer": RED}


def build_deck(results: dict[str, Any], cfg: dict[str, Any], path: Path,
               summary: list[str] | None = None) -> Path:
    d = Deck(cfg.get("theme"), footer=f"{cfg.get('kunde')} · {cfg.get('projekt')}")
    d.title_slide(cfg.get("projekt", "Digitalisierungsstrategie"), cfg.get("kunde", ""),
                  f"{cfg.get('berater') or ''}  ·  Stand {dt.date.today():%d.%m.%Y}".strip(" ·"))
    _summary(d, results, summary or management_summary(results, cfg))
    _vorgehen(d)

    d.section(PHASES["ist"], "Stakeholder · Prozesse · Wertstrom · Pain Points · Ursachen · User Stories")
    if "stakeholder" in results:
        _stakeholder(d, results["stakeholder"])
    if "sipoc" in results:
        _sipoc(d, results["sipoc"])
    if "prozess" in results:
        _prozess(d, results["prozess"])
    if "vsm" in results:
        _vsm(d, results["vsm"])
    if "painpoints" in results:
        _painpoints(d, results["painpoints"])
    if "ursachen" in results:
        _ursachen(d, results["ursachen"])
    if "userstories" in results:
        _userstories(d, results["userstories"])

    d.section(PHASES["synthese"], "Story Map · Handlungsfelder · Use-Case-Steckbriefe · Canvas · Fit-Gap")
    if "storymap" in results:
        _storymap(d, results["storymap"])
    if "handlungsfelder" in results and results["handlungsfelder"]["items"]:
        _handlungsfelder(d, results["handlungsfelder"])
    if "usecases" in results:
        _usecases(d, results)
    if "vpc" in results:
        _vpc(d, results["vpc"], results)
    if "bmc" in results:
        _bmc(d, results["bmc"], results)
    if "fitgap" in results:
        _fitgap(d, results["fitgap"])

    d.section(PHASES["bewertung"], "Nutzwertanalyse · Impact/Effort · Business Case · Roadmap · RACI")
    if "nwa" in results:
        _nwa(d, results["nwa"])
    if "impact_effort" in results and results["impact_effort"]["items"]:
        _impact_effort(d, results["impact_effort"])
    if "businesscase" in results:
        _businesscase(d, results["businesscase"], results)
    if "roadmap" in results:
        _roadmap(d, results["roadmap"])
    if "raci" in results:
        _raci(d, results["raci"])
    _next_steps(d, results)
    if results.get("issues"):
        _issues(d, results["issues"])
    path.parent.mkdir(parents=True, exist_ok=True)
    d.prs.save(path)
    return path


# ---------------------------------------------------------------- Rahmen
def _summary(d: Deck, r: dict, bullets: list[str]) -> None:
    s = d.slide("Management Summary", "Kernaussagen der Digitalisierungsstrategie")
    kz = list(r.get("kennzahlen", {}).items())[:6]
    w = (d.W - 1.0 - 0.2 * 5) / 6
    for i, (k, v) in enumerate(kz):
        if isinstance(v, float):
            v = eur(v) if "EUR" in k else pct(v)
        d.kpi(s, 0.5 + i * (w + 0.2), 1.55, w, 1.05, str(v), k.replace(" [EUR]", ""), d.palette[i % len(d.palette)])
    d.bullets(s, bullets, 0.5, 2.9, d.W - 1.0, d.H - 3.5, size=16)


def _vorgehen(d: Deck) -> None:
    s = d.slide("Vorgehen und Methodenkoffer", "Von der Ist-Analyse über die Synthese zur priorisierten Roadmap")
    bp = methods_by_phase()
    w = (d.W - 1.0 - 0.3) / 3
    for i, (ph, title) in enumerate(PHASES.items()):
        x = 0.5 + i * (w + 0.15)
        d.rect(s, x, 1.6, w, 0.8, fill=d.palette[i], shape=MSO_SHAPE.CHEVRON if i else MSO_SHAPE.PENTAGON,
               text=title, size=14, bold=True, color=WHITE)
        items = [m.title for m in bp[ph]]
        d.rect(s, x, 2.55, w, d.H - 3.3, fill=d.light)
        d.bullets(s, items, x + 0.15, 2.7, w - 0.3, d.H - 3.6, size=13)


def _next_steps(d: Deck, r: dict) -> None:
    s = d.slide("Nächste Schritte", "Vom Strategiepapier in die Umsetzung")
    steps = ["Roadmap und Priorisierung im Lenkungskreis verabschieden; Budget für NOW-Horizont freigeben"]
    if "roadmap" in r:
        now = r["roadmap"]["horizonte"]["Now"]
        for it in now[:4]:
            steps.append(f"{it['id']} {it['titel']}: MVP schneiden, Pilot "
                         f"{it.get('pilot') or 'festlegen'}, Start {it['start']}")
    if "raci" in r and r["raci"]["fehler"]:
        steps.append(f"RACI bereinigen: {len(r['raci']['fehler'])} Vorhaben ohne eindeutige Verantwortung")
    if "fitgap" in r and r["fitgap"]["kritische_gaps"]:
        steps.append(f"{len(r['fitgap']['kritische_gaps'])} kritische Muss-Gaps mit Anbieter/Prozess klären")
    steps.append("Governance etablieren: Digitalisierungs-Board, KPI-Tracking je Use Case, Quartals-Review der Roadmap")
    d.bullets(s, steps, 0.5, 1.6, d.W - 1, d.H - 2.3, size=18, bullet="→")


def _issues(d: Deck, issues: list[str]) -> None:
    for i, part in enumerate(chunks(issues, 22)):
        s = d.slide("Anhang: Datenqualität der Workshop-Erfassung" + (f" ({i + 1})" if i else ""),
                    "Automatische Plausibilitätsprüfung – bitte vor finaler Version bereinigen")
        d.bullets(s, part, 0.5, 1.55, d.W - 1, d.H - 2.2, size=11)


# ---------------------------------------------------------------- Phase 1
def _stakeholder(d: Deck, r: dict) -> None:
    s = d.slide("Stakeholder-Mapping", "Einfluss/Interesse-Matrix mit Einbindungsstrategie",
                "Stakeholder-Analyse (vertraulich)")
    pts = [{"x": i["interesse"], "y": i["einfluss"], "label": i["name"], "h": i.get("haltung")} for i in r["items"]]
    d.matrix(s, pts, 1.2, 1.6, 6.0, 5.0, "Interesse", "Einfluss",
             ("Zufriedenstellen", "Eng einbinden", "Beobachten", "Informieren"),
             ("FDF0E4", "E3EDF7", "F3F4F6", "EAF5EE"), lo=0.5, hi=5.5, mid=2.75, point_color=lambda p: HALTUNG_COL.get(p["h"], d.primary))
    x0 = 7.7
    d.text(s, "Haltung:", x0, 1.55, 1.0, 0.3, size=10, bold=True, color=GREY)
    for i, (h, c) in enumerate(HALTUNG_COL.items()):
        d.rect(s, x0 + 0.85 + i * 0.95, 1.62, 0.16, 0.16, fill=c, shape=MSO_SHAPE.OVAL)
        d.text(s, h, x0 + 1.04 + i * 0.95, 1.55, 0.8, 0.3, size=9, color=GREY)
    strat = {"Eng einbinden": "Kernteam/Lenkungskreis, aktiv in Workshops",
             "Zufriedenstellen": "Regelmäßige Entscheidungsvorlagen, kurze Updates",
             "Informieren": "Workshop-Teilnahme, Newsletter, Demo-Sessions",
             "Beobachten": "Bei Bedarf informieren"}
    rows = [["Quadrant", "Personen", "Strategie"]]
    for q, names in r["quadranten"].items():
        rows.append([q, ", ".join(names) or "–", strat[q]])
    d.table(s, rows, x0, 2.0, d.W - x0 - 0.5, col_w=[1.3, 2.0, 2.2], size=9, row_h=0.62)
    if r["risiken"]:
        txt = "Risiko: " + "; ".join(f"{x['name']} ({x['haltung']}, Einfluss {x['einfluss']})" for x in r["risiken"])
        d.rect(s, x0, 5.4, d.W - x0 - 0.5, 0.9, fill="FDECEA", text=txt, size=10, color=RED, align="l",
               autosize=True)


def _sipoc(d: Deck, r: dict) -> None:
    for name, p in r["prozesse"].items():
        s = d.slide(f"SIPOC: {name}", "Prozessrahmen – Lieferanten, Inputs, Hauptschritte, Outputs, Kunden")
        heads = ["Supplier", "Input", "Process", "Output", "Customer"]
        keys = ["supplier", "input", "process", "output", "customer"]
        w = (d.W - 1.0 - 0.4) / 5
        for i, (h, k) in enumerate(zip(heads, keys)):
            x = 0.5 + i * (w + 0.1)
            d.rect(s, x, 1.6, w, 0.6, fill=d.primary if k != "process" else d.accent, text=h, size=16,
                   bold=True, color=WHITE)
            d.rect(s, x, 2.3, w, d.H - 3.0, fill=d.light)
            vals = p[k]
            if k == "process":
                for j, v in enumerate(vals[:9]):
                    d.rect(s, x + 0.15, 2.45 + j * 0.5, w - 0.3, 0.42, fill=WHITE, line=d.accent, text=f"{j + 1}. {v}",
                           size=11, autosize=True)
            else:
                d.bullets(s, vals, x + 0.1, 2.4, w - 0.2, d.H - 3.2, size=12)


def _levels(steps: list[dict], succ: dict[int, list[int]]) -> dict[int, int]:
    nrs = [s["nr"] for s in steps]
    preds = defaultdict(list)
    for a, bs in succ.items():
        for b in bs:
            preds[b].append(a)
    level: dict[int, int] = {}
    order = {n: i for i, n in enumerate(nrs)}
    for n in nrs:  # Vorwärtskanten bestimmen die Spalte, Rücksprünge werden ignoriert
        fwd = [p for p in preds[n] if order.get(p, 1e9) < order[n]]
        level[n] = max((level[p] + 1 for p in fwd if p in level), default=0)
    return level


def _prozess(d: Deck, r: dict) -> None:
    for name, p in r["prozesse"].items():
        s = d.slide(f"Ist-Prozess: {name}", "Swimlane-Darstellung · rote Pfeile = Medienbruch",
                    "Prozessaufnahme; BPMN-2.0-Datei im Ergebnisordner")
        steps, lanes, succ = p["steps"], p["lanes"], p["succ"]
        lv = _levels(steps, succ)
        ncol = max(lv.values()) + 1 if lv else 1
        x0, y0, lane_w = 0.5, 1.55, 1.5
        W, H = d.W - 1.0, d.H - 2.35
        lh = H / max(1, len(lanes))
        cw = (W - lane_w) / ncol
        for i, ln in enumerate(lanes):
            y = y0 + i * lh
            d.rect(s, x0, y, lane_w, lh, fill=d.primary, line=WHITE, text=ln, size=11, bold=True, color=WHITE,
                   autosize=True)
            d.rect(s, x0 + lane_w, y, W - lane_w, lh, fill=d.light if i % 2 == 0 else WHITE, line=MID)
        pos = {}
        slot = defaultdict(int)
        cell_n = defaultdict(int)
        for st in steps:
            cell_n[(lanes.index(st.get("lane") or "–"), lv[st["nr"]])] += 1
        max_stack = max(cell_n.values(), default=1)
        bw = min(cw * 0.8, 1.6)
        bh = min(lh * 0.7 / max_stack, 0.75)
        for st in steps:
            li = lanes.index(st.get("lane") or "–")
            col = lv[st["nr"]]
            k = slot[(li, col)]
            slot[(li, col)] += 1
            n_cell = cell_n[(li, col)]
            cx = x0 + lane_w + col * cw + cw / 2
            cy = y0 + li * lh + lh / 2 + (k - (n_cell - 1) / 2) * (lh / n_cell)
            pos[st["nr"]] = (cx, cy)
            typ = st.get("typ") or "Aufgabe"
            label = f"{st['nr']}. {st['schritt']}"
            if typ in ("Start", "Ende"):
                rr = min(bh, 0.5)
                d.rect(s, cx - rr / 2, cy - rr / 2, rr, rr, fill=GREEN if typ == "Start" else RED,
                       shape=MSO_SHAPE.OVAL, line=WHITE)
                d.text(s, st["schritt"], cx - cw / 2, cy + rr / 2, cw, 0.3, size=8, align="c", color=GREY)
            elif typ == "Entscheidung":
                d.rect(s, cx - bw / 2, cy - bh / 2, bw, bh, fill="FFF4E0", line=AMBER, shape=MSO_SHAPE.DIAMOND,
                       text=label, size=9, autosize=True)
            else:
                nd = st.get("medium") in ("Papier", "Excel", "E-Mail", "Mündlich/Telefon", "Manuell/ohne Medium")
                d.rect(s, cx - bw / 2, cy - bh / 2, bw, bh, fill=WHITE, line=AMBER if nd else d.primary,
                       shape=MSO_SHAPE.ROUNDED_RECTANGLE, text=label, size=10, autosize=True, line_w=1.5)
                tag = (st.get("system") or st.get("medium") or "")[:22]
                if tag:
                    d.text(s, tag, cx - bw / 2, cy + bh / 2, bw, 0.22, size=7, color=GREY, align="c", italic=True)
        breaks = {(b["von_nr"], b["nach_nr"]) for b in p["medienbrueche"]}
        for a, bs in succ.items():
            for b in bs:
                if a not in pos or b not in pos:
                    continue
                (x1, y1), (x2, y2) = pos[a], pos[b]
                col = RED if (a, b) in breaks else GREY
                if x2 > x1:
                    d.elbow(s, x1 + bw / 2, y1, x2 - bw / 2, y2, color=col, width=2 if col == RED else 1.25)
                else:  # Rücksprung
                    d.arrow(s, x1, y1 - bh / 2, x2, y2 - bh / 2, color=col, dash=True)
        _process_kpis(d, name, p)


def _process_kpis(d: Deck, name: str, p: dict) -> None:
    s = d.slide(f"Ist-Prozess {name}: Kennzahlen und Medienbrüche")
    d.kpi(s, 0.5, 1.55, 2.8, 1.0, str(p["anzahl_schritte"]), "Prozessschritte")
    d.kpi(s, 3.5, 1.55, 2.8, 1.0, str(p["uebergaben"]), "Übergaben zwischen Rollen")
    d.kpi(s, 6.5, 1.55, 2.8, 1.0, str(len(p["medienbrueche"])), "Medienbrüche", RED)
    d.kpi(s, 9.5, 1.55, 3.3, 1.0, pct(p["digitalgrad"]), "Digitalgrad (integriert)", GREEN)
    rows = [["Von", "Nach", "Art des Bruchs"]] + [[b["von"], b["nach"], b["art"]] for b in p["medienbrueche"][:12]]
    if len(rows) > 1:
        d.table(s, rows, 0.5, 2.85, 7.6, col_w=[3, 3, 2.4], size=10, row_h=0.32)
    med = p["medien"]
    d.bar_chart(s, list(med), {"Schritte": list(med.values())}, 8.5, 2.8, 4.3, 3.9, horizontal=True,
                legend=False, title="Medien je Prozessschritt", colors=[d.primary])


def _vsm(d: Deck, r: dict) -> None:
    for name, v in r["prozesse"].items():
        s = d.slide(f"Wertstrom: {name}", "Bearbeitungs-, Rüst- und Wartezeiten, Qualität und Informationsfluss",
                    "Value Stream Mapping")
        steps = v["steps"][:9]
        n = len(steps)
        gap = 0.45
        bw = (d.W - 1.0 - gap * (n - 1)) / max(n, 1)
        y = 2.3
        for i, st in enumerate(steps):
            x = 0.5 + i * (bw + gap)
            d.rect(s, x, y, bw, 0.6, fill=d.primary, text=st["schritt"], size=10, bold=True, color=WHITE,
                   autosize=True)
            info = (f"BZ {st.get('bz') or 0:g} min\nRZ {st.get('rz') or 0:g} min\n"
                    f"FPY {st.get('qr') if st.get('qr') is not None else '–'} %\n{st.get('system') or ''}")
            d.rect(s, x, y + 0.6, bw, 1.05, fill=WHITE, line=d.primary, text=info, size=9, align="l", anchor="t",
                   autosize=True)
            if st.get("medium") and st["medium"] != "Digital integriert":
                d.text(s, f"⚠ {st['medium']}", x, y - 0.3, bw, 0.25, size=8, color=AMBER, align="c")
            if i > 0 and (st.get("wz") or 0) > 0:
                d.rect(s, x - gap + 0.06, y + 0.15, gap - 0.12, 0.35, fill=AMBER, shape=MSO_SHAPE.ISOSCELES_TRIANGLE)
        # Zeitleiter
        ly = 4.35
        for i, st in enumerate(steps):
            x = 0.5 + i * (bw + gap)
            wz = st.get("wz") or 0
            d.arrow(s, x - gap if i else x, ly, x + bw / 2, ly, color=RED, head=False, width=2)
            d.text(s, f"{_t(wz)}", x - gap if i else x, ly - 0.3, gap + bw / 2, 0.25, size=8, color=RED, align="c")
            d.arrow(s, x + bw / 2, ly + 0.35, x + bw, ly + 0.35, color=GREEN, head=False, width=2)
            d.arrow(s, x + bw / 2, ly, x + bw / 2, ly + 0.35, color=GREY, head=False)
            d.text(s, _t((st.get("bz") or 0) + (st.get("rz") or 0)), x + bw / 2, ly + 0.4, bw / 2, 0.25, size=8,
                   color=GREEN, align="c")
        kp = [(_t(v["durchlaufzeit"]), "Durchlaufzeit"), (_t(v["wertschoepfung"]), "Wertschöpfende Zeit"),
              (pct(v["flussgrad"]), "Flussgrad"), (pct(v["rty"]), "Rolled Throughput Yield"),
              (str(len(v["medienbrueche"])), "Medienbrüche"), (v["engpass"] or "–", "Größte Wartezeit vor")]
        w = (d.W - 1.0 - 0.15 * 5) / 6
        for i, (val, lab) in enumerate(kp):
            d.kpi(s, 0.5 + i * (w + 0.15), 5.35, w, 0.95, val, lab, [d.accent, GREEN, BLUE, BLUE, RED, AMBER][i])
        d.text(s, "Legende: ▲ Bestand/Wartezeit (rot), grüne Linie = Bearbeitungs- + Rüstzeit, ⚠ = nicht integriertes Medium",
               0.5, 1.5, d.W - 1, 0.3, size=9, color=GREY)
        s2 = d.slide(f"Wertstrom {name}: Zeitanteile je Prozessschritt")
        d.bar_chart(s2, [x["schritt"] for x in v["steps"]],
                    {"Bearbeitungszeit": [x.get("bz") or 0 for x in v["steps"]],
                     "Rüstzeit": [x.get("rz") or 0 for x in v["steps"]],
                     "Wartezeit": [x.get("wz") or 0 for x in v["steps"]]},
                    0.5, 1.5, d.W - 1, d.H - 2.2, stacked=True, colors=[GREEN, BLUE, RED], number_format='0')


def _t(minutes: float) -> str:
    if minutes >= 60 * 24:
        return f"{minutes / 60 / 24:.1f} d".replace(".", ",")
    if minutes >= 60:
        return f"{minutes / 60:.1f} h".replace(".", ",")
    return f"{minutes:.0f} min"


def _painpoints(d: Deck, r: dict) -> None:
    s = d.slide("Pain Points: Pareto nach Kategorie", "Score = Häufigkeit × Auswirkung, kumulierter Anteil")
    cats = [p["kategorie"] for p in r["pareto"]]
    d.bar_chart(s, cats, {"Score": [p["score"] for p in r["pareto"]]}, 0.5, 1.5, 8.0, d.H - 2.2,
                legend=False, colors=[d.accent])
    rows = [["Kategorie", "Anzahl", "Kumuliert"]] + [[p["kategorie"], p["anzahl"], pct(p["kumuliert"])]
                                                     for p in r["pareto"]]
    d.table(s, rows, 8.8, 1.6, 4.0, col_w=[2.2, 0.8, 1.0], size=10, row_h=0.32,
            fill_fn=lambda i, j, v: "FDF0E4" if i and r["pareto"][i - 1]["kumuliert"] <= 0.8 else None)
    d.text(s, "Hervorgehoben: Kategorien bis 80 % der Problemlast", 8.8, 1.6 + 0.32 * len(rows) + 0.1, 4, 0.3,
           size=9, color=GREY)

    for k, part in enumerate(chunks(r["items"][:20], 10)):
        s = d.slide("Top Pain Points" + (f" ({k + 1})" if k else ""), "Priorisiert nach Score und Dot-Voting")
        rows = [["#", "ID", "Pain Point", "Prozess / Schritt", "Kategorie", "H", "A", "Score", "Votes"]]
        for p in part:
            rows.append([p["rang"], p["id"], p["beschreibung"], f"{p.get('prozess') or ''} / {p.get('schritt') or ''}",
                         p.get("kategorie") or "", p.get("haeufigkeit") or "", p.get("auswirkung") or "", p["score"],
                         p["stimmen"]])
        d.table(s, rows, 0.5, 1.55, d.W - 1, col_w=[0.4, 0.7, 4.6, 2.6, 1.4, 0.4, 0.4, 0.6, 0.6], size=10,
                row_h=0.46, fill_fn=lambda i, j, v: _heat(v, 25) if j == 7 and i else None)

    procs = list(r["heatmap"])
    cats = r["kategorien"]
    if len(procs) > 1 or len(cats) > 1:
        s = d.slide("Pain-Point-Heatmap: Prozess × Kategorie", "Summe Score – wo brennt es am meisten?")
        mx = max((v for row in r["heatmap"].values() for v in row.values()), default=1)
        rows = [["Prozess"] + cats] + [[p] + [r["heatmap"][p].get(c, "") for c in cats] for p in procs]
        d.table(s, rows, 0.5, 1.6, d.W - 1, col_w=[2.2] + [1] * len(cats), size=10, row_h=0.5,
                fill_fn=lambda i, j, v: _heat(v, mx) if i and j and v != "" else None, bold_first_col=True)


def _heat(v: Any, mx: float) -> str | None:
    try:
        f = float(v) / mx
    except (TypeError, ValueError, ZeroDivisionError):
        return None
    if f > 0.66:
        return "F5B7B1"
    if f > 0.33:
        return "FAD7A0"
    return "FCF3CF"


def _ursachen(d: Deck, r: dict) -> None:
    for k, part in enumerate(chunks(r["whys"], 4)):
        if not part:
            break
        s = d.slide("Ursachenanalyse: 5 Whys" + (f" ({k + 1})" if k else ""),
                    "Von Symptom zur Grundursache – Digitalisierung setzt an der Ursache an")
        rh = (d.H - 2.2) / 4
        for i, c in enumerate(part):
            y = 1.55 + i * rh
            chain = [c["problem"]] + c["kette"]
            n = len(chain) + 1
            w = (d.W - 1.0) / n
            for j, t in enumerate(chain):
                d.rect(s, 0.5 + j * w, y, w - 0.25, rh - 0.15, fill=d.primary if j == 0 else d.light,
                       text=(f"{c['pain_point_id']}: " if j == 0 else f"Warum {j}:\n") + t,
                       size=11, color=WHITE if j == 0 else DARK, autosize=True)
                d.arrow(s, 0.5 + j * w + w - 0.23, y + (rh - 0.15) / 2, 0.5 + (j + 1) * w - 0.02,
                        y + (rh - 0.15) / 2, color=d.accent, width=2)
            col = {"Ja": GREEN, "Teilweise": AMBER, "Nein": GREY}.get(c.get("digital_loesbar"), d.accent)
            d.rect(s, 0.5 + (n - 1) * w, y, w - 0.02, rh - 0.15, fill=col,
                   text=f"Grundursache: {c['grundursache'] or '–'}\nDigital lösbar: {c.get('digital_loesbar') or '?'}",
                   size=9, color=WHITE, bold=True, autosize=True)
    for prob, cats in r["ishikawa"].items():
        _fishbone(d, prob, cats)


def _fishbone(d: Deck, problem: str, cats: dict[str, list[str]]) -> None:
    s = d.slide("Ishikawa-Diagramm", problem, "Ursache-Wirkungs-Analyse (✓ = bestätigt)")
    y_sp, top_y, bot_y = 4.2, 1.6, d.H - 0.85
    x0, x1 = 0.6, d.W - 2.6
    d.arrow(s, x0, y_sp, x1, y_sp, color=d.primary, width=3)
    d.rect(s, x1 + 0.05, y_sp - 0.6, 1.95, 1.2, fill=RED, text=problem, size=12, bold=True, color=WHITE,
           shape=MSO_SHAPE.ROUNDED_RECTANGLE, autosize=True)
    items = list(cats.items())
    top, bot = items[0::2], items[1::2]
    ncol = max(len(top), len(bot), 1)
    seg = (x1 - x0) / ncol
    for row, grp in ((-1, top), (1, bot)):
        for i, (cat, causes) in enumerate(grp):
            bx = x0 + seg * (i + 1) - 0.1  # Ansatzpunkt an der Hauptgräte
            lx = bx - min(seg - 0.15, 2.3)  # linker Rand des Segments
            lab_y = top_y if row < 0 else bot_y - 0.38
            d.rect(s, lx, lab_y, 1.6, 0.38, fill=d.primary, text=cat, size=12, bold=True, color=WHITE)
            end_x = lx + 1.65
            d.arrow(s, end_x, lab_y + (0.19), bx, y_sp, color=GREY, width=1.75)
            txt = "\n".join(f"• {c}" for c in causes[:6])
            ty = top_y + 0.45 if row < 0 else y_sp + 0.15
            d.text(s, txt, lx, ty, 1.55, 2.0, size=10, autosize=True)


def _userstories(d: Deck, r: dict) -> None:
    for k, part in enumerate(chunks(r["items"], 7)):
        s = d.slide("User Stories mit INVEST-Prüfung" + (f" ({k + 1})" if k else ""),
                    f"Ø INVEST-Erfüllung {pct(r['invest_quote'])} · {r['story_points']} Story Points gesamt")
        rows = [["ID", "User Story", "Akzeptanzkriterien (Given/When/Then)", "I", "N", "V", "E", "S", "T", "SP"]]
        for st in part:
            ch = st["invest"]["checks"]
            rows.append([st["id"], st["invest"]["satz"], (st.get("akzeptanz") or "–")[:260]] +
                        ["✓" if ch[x] else "✗" for x in "INVEST"] + [st.get("story_points") or "–"])
        d.table(s, rows, 0.5, 1.5, d.W - 1, col_w=[0.6, 4.6, 4.6, 0.3, 0.3, 0.3, 0.3, 0.3, 0.3, 0.4], size=9,
                row_h=0.68, color_fn=lambda i, j, v: GREEN if v == "✓" else RED if v == "✗" else None)
    bad = r["nachbessern"]
    if bad:
        s = d.slide("User Stories: Nachbesserungsbedarf", "Automatische INVEST-Hinweise")
        rows = [["ID", "Hinweise"]] + [[b["id"], "; ".join(b["invest"]["hinweise"])] for b in bad[:14]]
        d.table(s, rows, 0.5, 1.55, d.W - 1, col_w=[0.8, 10], size=10, row_h=0.36)


# ---------------------------------------------------------------- Phase 2
def _storymap(d: Deck, r: dict) -> None:
    acts, rels = r["aktivitaeten"], r["releases"]
    s = d.slide("User Story Map", "Backbone = Nutzeraktivitäten · Zeilen = Releases (MVP zuerst)")
    lw = 1.3
    cw = (d.W - 1.0 - lw) / max(1, len(acts))
    rh = (d.H - 2.6) / max(1, len(rels))
    for j, a in enumerate(acts):
        d.rect(s, 0.5 + lw + j * cw + 0.03, 1.55, cw - 0.06, 0.55, fill=d.primary, text=a, size=11, bold=True,
               color=WHITE, autosize=True)
    for i, rl in enumerate(rels):
        y = 2.2 + i * rh
        d.rect(s, 0.5, y, lw - 0.05, rh - 0.06, fill=d.accent if rl == "MVP" else MID,
               text=f"{rl}\n{r['story_points'].get(rl, 0)} SP", size=11, bold=True,
               color=WHITE if rl == "MVP" else DARK)
        d.arrow(s, 0.5, y + rh - 0.03, d.W - 0.5, y + rh - 0.03, color=MID, head=False, dash=True)
        for j, a in enumerate(acts):
            cards = r["grid"][a][rl]
            if not cards:
                continue
            ch = min(0.55, (rh - 0.1) / len(cards))
            for k, c in enumerate(cards):
                d.rect(s, 0.5 + lw + j * cw + 0.06, y + 0.03 + k * ch, cw - 0.12, ch - 0.04, fill="FFF6D6",
                       line="E6C860", text=f"{c['id']} {c.get('ziel') or ''}", size=9, autosize=True, align="l")


def _handlungsfelder(d: Deck, r: dict) -> None:
    s = d.slide("Handlungsfelder und Potenzial", "Potenzial = Summe der adressierten Pain-Point-Scores")
    items = r["items"]
    d.bar_chart(s, [h["name"] for h in items], {"Potenzial": [h["potenzial"] for h in items]}, 0.5, 1.5, 5.6,
                d.H - 2.2, horizontal=True, legend=False, colors=[d.accent])
    rows = [["Handlungsfeld", "Stories", "SP", "Use Cases", "Pain Points"]]
    for h in items:
        rows.append([h["name"], h["stories"], h["story_points"], ", ".join(h["use_cases"]) or "–",
                     ", ".join(h["pain_points"]) or "–"])
    d.table(s, rows, 6.4, 1.6, d.W - 6.9, col_w=[2.0, 0.7, 0.6, 1.4, 1.8], size=9, row_h=0.45)
    if r["nicht_adressiert"]:
        txt = "Nicht adressierte Pain Points: " + ", ".join(
            f"{p['id']} (Score {p['score']})" for p in r["nicht_adressiert"][:12])
        d.rect(s, 6.4, d.H - 1.4, d.W - 6.9, 0.65, fill="FDECEA", text=txt, size=10, color=RED, align="l",
               autosize=True)
    for h in items:
        if h.get("zielbild") or h.get("beschreibung"):
            break
    else:
        return
    s = d.slide("Handlungsfelder: Zielbilder")
    rows = [["Handlungsfeld", "Beschreibung", "Zielbild", "Verantwortlich"]] + [
        [h["name"], h.get("beschreibung") or "", h.get("zielbild") or "", h.get("verantwortlich") or ""]
        for h in items]
    d.table(s, rows, 0.5, 1.55, d.W - 1, col_w=[2, 4, 4, 1.6], size=10, row_h=0.6)


def _usecases(d: Deck, results: dict) -> None:
    items = results["usecases"]["items"]
    nwa = {i["id"]: i for i in results.get("nwa", {}).get("items", [])}
    ie = {i["id"]: i for i in results.get("impact_effort", {}).get("items", [])}
    bc = results.get("businesscase", {}).get("use_cases", {})
    s = d.slide("Use-Case-Portfolio", f"{len(items)} Use Cases aus dem Potenzial-Workshop")
    rows = [["ID", "Use Case", "Handlungsfeld", "Owner", "Pain Points", "Quadrant", "NWA"]]
    for u in items:
        rows.append([u["id"], u["titel"], u.get("handlungsfeld") or "", u.get("owner") or "",
                     ", ".join(u.get("pain_point_ids") or []), ie.get(u["id"], {}).get("quadrant", "–"),
                     pct(nwa[u["id"]]["prozent"]) if u["id"] in nwa else "–"])
    d.table(s, rows[:16], 0.5, 1.55, d.W - 1, col_w=[0.6, 3.6, 2.2, 1.6, 1.6, 1.4, 0.7], size=10, row_h=0.36,
            fill_fn=lambda i, j, v: {"Quick Win": "D5F5E3", "Großes Vorhaben": "D6EAF8", "Lückenfüller": "FCF3CF",
                                     "Nicht machen": "FADBD8"}.get(v) if j == 5 else None)
    for u in items:
        s = d.slide(f"Use-Case-Steckbrief {u['id']}: {u['titel']}",
                    f"Handlungsfeld: {u.get('handlungsfeld') or '–'}  ·  Owner: {u.get('owner') or '–'}")
        fields = [("Problem / Ausgangslage", u.get("problem")), ("Lösungsidee", u.get("loesung")),
                  ("Nutzen", u.get("nutzen")), ("Datenbasis", u.get("datenbasis")),
                  ("Systeme / Technologie", u.get("systeme")), ("Aufwand / Voraussetzungen", u.get("aufwand"))]
        bw, bh = (d.W - 1.0 - 3.3 - 0.3) / 2, (d.H - 2.35) / 3
        for i, (lab, val) in enumerate(fields):
            x = 0.5 + (i % 2) * (bw + 0.15)
            y = 1.55 + (i // 2) * (bh + 0.07)
            d.rect(s, x, y, bw, 0.32, fill=d.primary, text=lab, size=11, bold=True, color=WHITE, align="l")
            d.rect(s, x, y + 0.32, bw, bh - 0.32, fill=d.light, text=val or "–", size=11, align="l", anchor="t",
                   autosize=True)
        x = d.W - 0.5 - 3.3
        d.rect(s, x, 1.55, 3.3, 0.32, fill=d.accent, text="Bewertung", size=11, bold=True, color=WHITE, align="l")
        e = ie.get(u["id"], {})
        n = nwa.get(u["id"])
        b = bc.get(u["id"])
        lines = [f"Quadrant: {e.get('quadrant', '–')}",
                 f"Impact / Effort: {e.get('impact', '–')} / {e.get('effort', '–')}",
                 f"Nutzwert: {pct(n['prozent']) if n else '–'} (Rang {n['rang'] if n else '–'})"]
        if b:
            lines += [f"Investition: {eur(b['invest'])}", f"Kapitalwert: {eur(b['npv'])}",
                      f"Amortisation: {b['payback']:.1f} J." if b["payback"] is not None else "Amortisation: –"]
        lines += ["", f"Pain Points: {', '.join(u.get('pain_point_ids') or []) or '–'}",
                  f"User Stories: {', '.join(u.get('story_ids') or []) or '–'}"]
        d.rect(s, x, 1.87, 3.3, d.H - 2.6, fill=d.light, text="\n".join(lines), size=11, align="l", anchor="t",
               autosize=True)


def _uc_title(results: dict, uc: str) -> str:
    for u in results.get("usecases", {}).get("items", []):
        if u["id"] == uc:
            return f"{uc}: {u['titel']}"
    return uc


def _vpc(d: Deck, r: dict, results: dict) -> None:
    for uc, c in r["canvases"].items():
        s = d.slide(f"Value Proposition Canvas – {_uc_title(results, uc)}",
                    f"Kundensegment: {c.get('segment') or '–'} · Fit {pct(c['fit'])} der Pains/Gains adressiert")
        b = c["bereiche"]
        # Wertangebot (Quadrat, links)
        x, y, sz = 0.9, 1.55, 4.55
        d.rect(s, x, y, sz, sz, fill=d.light, line=d.primary, line_w=2)
        d.rect(s, x, y, sz, sz / 2, fill=None, line=MID, text="", size=10)
        d.rect(s, x + sz / 2, y + sz / 2, sz / 2, sz / 2, fill=None, line=MID)
        d.text(s, "Gain Creators\n" + _lst(b.get("Gain Creator")), x + 0.1, y + 0.1, sz - 0.2, sz / 2 - 0.2, size=10,
               autosize=True)
        d.text(s, "Produkte & Services\n" + _lst(b.get("Produkt/Service")), x + 0.1, y + sz / 2 + 0.1, sz / 2 - 0.2,
               sz / 2 - 0.2, size=10, autosize=True)
        d.text(s, "Pain Relievers\n" + _lst(b.get("Pain Reliever")), x + sz / 2 + 0.1, y + sz / 2 + 0.1,
               sz / 2 - 0.2, sz / 2 - 0.2, size=10, autosize=True)
        # Kundenprofil (Kreis, rechts)
        cx = 7.6
        d.rect(s, cx, y, sz, sz, fill=d.light, line=d.accent, shape=MSO_SHAPE.OVAL, line_w=2)
        d.text(s, "Gains\n" + _lst(b.get("Gain")), cx + 0.8, y + 0.45, sz - 1.6, 1.4, size=10, autosize=True)
        d.text(s, "Pains\n" + _lst(b.get("Pain")), cx + 0.6, y + 1.95, sz / 2 - 0.6, 1.9, size=10, autosize=True)
        d.text(s, "Customer Jobs\n" + _lst(b.get("Customer Job")), cx + sz / 2 + 0.05, y + 1.95, sz / 2 - 0.6, 1.9,
               size=10, autosize=True)
        d.arrow(s, x + sz + 0.1, y + sz / 2, cx - 0.1, y + sz / 2, color=d.accent, width=3)
        if c["offene_pains"] or c["offene_gains"]:
            txt = "Ohne Lösungsbezug: " + "; ".join(c["offene_pains"] + c["offene_gains"])
            d.rect(s, 0.7, d.H - 1.05, d.W - 1.4, 0.45, fill="FDECEA", text=txt, size=9, color=RED, align="l",
                   autosize=True)


def _lst(v: list[str] | None) -> str:
    return "\n".join(f"• {x}" for x in (v or [])) or "–"


def _bmc(d: Deck, r: dict, results: dict) -> None:
    layout = [("Schlüsselpartner", 0, 0, 1, 2), ("Schlüsselaktivitäten", 1, 0, 1, 1), ("Schlüsselressourcen", 1, 1, 1, 1),
              ("Wertangebot", 2, 0, 1, 2), ("Kundenbeziehungen", 3, 0, 1, 1), ("Kanäle", 3, 1, 1, 1),
              ("Kundensegmente", 4, 0, 1, 2), ("Kostenstruktur", 0, 2, 2.5, 1), ("Einnahmequellen", 2.5, 2, 2.5, 1)]
    for uc, c in r["canvases"].items():
        s = d.slide(f"Business Model Canvas – {_uc_title(results, uc)}")
        cw, rh = (d.W - 1.0) / 5, (d.H - 2.2) / 3
        for name, cx, cy, w, h in layout:
            x, y = 0.5 + cx * cw, 1.5 + cy * rh
            d.rect(s, x, y, w * cw, h * rh, fill=d.light if name != "Wertangebot" else "FDF0E4", line=WHITE, line_w=3)
            d.text(s, name, x + 0.08, y + 0.05, w * cw - 0.16, 0.3, size=11, bold=True, color=d.primary)
            d.text(s, _lst(c.get(name)), x + 0.08, y + 0.38, w * cw - 0.16, h * rh - 0.45, size=10, autosize=True)


def _fitgap(d: Deck, r: dict) -> None:
    s = d.slide("Fit-Gap-Analyse", f"Fit-Quote (Standard + Konfiguration): {pct(r['fit_quote'])} · "
                f"Aufwand gesamt {r['aufwand_gesamt']:.0f} PT")
    cats = r["kategorien"]
    prios = [p for p in r["matrix"] if sum(r["matrix"][p].values())]
    d.bar_chart(s, prios, {c: [r["matrix"][p][c] for p in prios] for c in cats}, 0.5, 1.5, 6.0, 4.0,
                stacked=True, colors=[GREEN, "7DBB91", AMBER, RED], title="Anforderungen nach Priorität und Abdeckung")
    for i, c in enumerate(cats):
        d.kpi(s, 0.5 + i * 1.5, 5.7, 1.4, 0.95, str(r["anzahl"][c]), f"{c} ({r['aufwand'][c]:.0f} PT)",
              [GREEN, "7DBB91", AMBER, RED][i])
    gaps = r["kritische_gaps"] or [x for x in r["items"] if x.get("abdeckung") in ("Gap", "Erweiterung")]
    title = "Kritische Gaps (Muss)" if r["kritische_gaps"] else "Gaps und Erweiterungen"
    d.text(s, title, 6.9, 1.5, 6, 0.3, size=13, bold=True, color=RED)
    rows = [["ID", "Anforderung", "System", "Lösungsansatz", "PT"]] + [
        [g["id"], g["anforderung"], g.get("system") or "", g.get("loesung") or "", g.get("aufwand_pt") or ""]
        for g in gaps[:10]]
    d.table(s, rows, 6.9, 1.85, d.W - 7.4, col_w=[0.6, 2.4, 1.0, 1.8, 0.5], size=9, row_h=0.42)


# ---------------------------------------------------------------- Phase 3
def _nwa(d: Deck, r: dict) -> None:
    s = d.slide("Nutzwertanalyse: Kriterien und Gewichtung", "Transparente, gewichtete Bewertung (Skala 1-5)")
    rows = [["Kriterium", "Kategorie", "Gewicht", "Skala"]] + [
        [c["kriterium"], c["kategorie"], pct(c["w"]), c.get("skala") or ""] for c in r["kriterien"]]
    d.table(s, rows, 0.5, 1.6, d.W - 1, col_w=[3, 1.5, 1, 6], size=11, row_h=0.42)

    s = d.slide("Nutzwertanalyse: Ranking der Use Cases",
                "Gewichteter Nutzwert in % des Maximums · Rangspanne bei ±20 % Gewichtsänderung je Kriterium")
    items = r["items"]
    cats = r["kategorien"]
    d.bar_chart(s, [f"{i['id']} {i['titel']}"[:40] for i in items],
                {k: [i["teil"][k] / 5 for i in items] for k in cats}, 0.5, 1.5, 7.8, d.H - 2.2, horizontal=True,
                stacked=True, number_format='0%', labels=False)
    rows = [["Rang", "Use Case", "Nutzwert", "Rangspanne"]] + [
        [i["rang"], i["id"], pct(i["prozent"]), f"{i['rang_min']}–{i['rang_max']}" if i["rang_min"] != i["rang_max"]
         else "stabil"] for i in items[:14]]
    d.table(s, rows, 8.6, 1.6, d.W - 9.1, col_w=[0.6, 1, 1, 1.2], size=10, row_h=0.34)


def _impact_effort(d: Deck, r: dict) -> None:
    s = d.slide("Impact/Effort-Matrix", "Quick Wins zuerst – große Vorhaben planen – „Nicht machen“ bewusst streichen")
    pts = [{"x": i["effort"], "y": i["impact"], "label": i["id"], "q": i["quadrant"]}
           for i in r["items"]]
    d.matrix(s, pts, 1.2, 1.6, 6.6, 5.0, "Aufwand (Effort)", "Nutzen (Impact)",
             ("Quick Wins", "Große Vorhaben", "Lückenfüller", "Nicht machen"),
             ("EAF5EE", "E3EDF7", "FEF9E7", "FDECEA"), lo=0.5, hi=10.5, mid=5.5,
             point_color=lambda p: Q_COL[p["q"]])
    x = 8.3
    y = 1.6
    for q, its in r["quadranten"].items():
        d.rect(s, x, y, d.W - x - 0.5, 0.32, fill=Q_COL[q], text=f"{q} ({len(its)})", size=11, bold=True,
               color=WHITE, align="l")
        txt = "\n".join(f"{i['id']} {i['titel']}" for i in its) or "–"
        h = max(0.35, 0.22 * len(its) + 0.1)
        d.text(s, txt, x + 0.1, y + 0.36, d.W - x - 0.7, h, size=10, autosize=True)
        y += 0.36 + h + 0.12


def _businesscase(d: Deck, r: dict, results: dict) -> None:
    ucs = r["use_cases"]
    s = d.slide("Business Case und ROI je Use Case",
                f"Kalkulationszins {pct(r['zinssatz'])} · Betrachtung Jahr 0–{r['jahre']}")
    rows = [["Use Case", "Investition", "Kosten ges.", "Nutzen ges.", "Kapitalwert", "ROI", "IRR", "Amortisation"]]
    for uc, u in ucs.items():
        rows.append([_uc_title(results, uc)[:42], eur(u["invest"]), eur(u["kosten"]), eur(u["nutzen"]),
                     eur(u["npv"]), pct(u["roi"]), pct(u["irr"]),
                     f"{u['payback']:.1f} J." if u["payback"] is not None else "> Horizont"])
    p = r["portfolio"]
    rows.append(["Portfolio gesamt", eur(p["invest"]), "", "", eur(p["npv"]), pct(p["roi"]), pct(p["irr"]),
                 f"{p['payback']:.1f} J." if p["payback"] is not None else "> Horizont"])
    d.table(s, rows, 0.5, 1.55, d.W - 1, col_w=[3.6, 1.2, 1.2, 1.2, 1.2, 0.8, 0.8, 1.2], size=10, row_h=0.38,
            fill_fn=lambda i, j, v: "FDF0E4" if i == len(rows) - 1 else None,
            color_fn=lambda i, j, v: RED if j == 4 and isinstance(v, str) and v.startswith("-") else None)
    s = d.slide("Business Case: kumulierter Cashflow", "Break-even, wenn die Kurve die Nulllinie schneidet")
    yrs = [f"Jahr {t}" for t in range(r["jahre"] + 1)]
    series = {uc: u["kumuliert"] for uc, u in list(ucs.items())[:7]}
    cum, c = [], 0.0
    for f in p["flows"]:
        c += f
        cum.append(c)
    series["Portfolio"] = cum
    d.line_chart(s, yrs, series, 0.5, 1.5, d.W - 1, d.H - 2.2, number_format='#,##0 "€"')


def _roadmap(d: Deck, r: dict) -> None:
    s = d.slide("Roadmap: Now · Next · Later", "Priorisierte Use Cases je Horizont (* = automatischer Vorschlag)")
    w = (d.W - 1.0 - 0.3) / 3
    labels = {"Now": "NOW – starten (0-6 Monate)", "Next": "NEXT – vorbereiten (6-12 Monate)",
              "Later": "LATER – beobachten (> 12 Monate)"}
    for i, hz in enumerate(("Now", "Next", "Later")):
        x = 0.5 + i * (w + 0.15)
        d.rect(s, x, 1.55, w, 0.5, fill=HZ_COL[hz], text=labels[hz], size=13, bold=True, color=WHITE)
        its = r["horizonte"][hz]
        ch = min(0.95, (d.H - 2.9) / max(1, len(its)))
        for k, it in enumerate(its):
            txt = f"{it['id']} {it['titel']}{' *' if it['vorschlag'] else ''}"
            if it.get("mvp"):
                txt += f"\nMVP: {it['mvp']}"
            elif it.get("quadrant"):
                txt += f"\n{it['quadrant']}" + (f" · Rang {it['rang']}" if it.get("rang") else "")
            d.rect(s, x + 0.08, 2.15 + k * ch, w - 0.16, ch - 0.08, fill=d.light if it["empfohlen"] else "FDECEA",
                   line=HZ_COL[hz], text=txt, size=10, align="l", autosize=True)

    s = d.slide("Umsetzungs-Roadmap (Gantt)", f"Quartalsplanung ab {r['start']} – Abhängigkeiten berücksichtigt")
    qs = r["quartale"]
    lw = 3.4
    qw = (d.W - 1.0 - lw) / len(qs)
    for k, q in enumerate(qs):
        d.rect(s, 0.5 + lw + k * qw, 1.5, qw - 0.02, 0.4, fill=d.primary if k % 4 == 0 else "3B5A80", text=q,
               size=10, bold=True, color=WHITE)
    items = [i for i in r["items"] if i["empfohlen"]]
    rh = min(0.42, (d.H - 2.6) / max(1, len(items)))
    for n, it in enumerate(items):
        y = 2.0 + n * rh
        if n % 2 == 0:
            d.rect(s, 0.5, y, d.W - 1.0, rh, fill=d.light)
        d.text(s, f"{it['id']} {it['titel']}"[:48], 0.55, y + 0.03, lw - 0.1, rh - 0.05, size=10, anchor="m")
        x = 0.5 + lw + it["start_idx"] * qw
        bw = (it["ende_idx"] - it["start_idx"]) * qw
        d.rect(s, x + 0.02, y + 0.06, bw - 0.04, rh - 0.12, fill=HZ_COL[it["horizont"]],
               shape=MSO_SHAPE.ROUNDED_RECTANGLE, text=f"{it['dauer']} Mon.", size=8, color=WHITE)
        d.rect(s, x + bw - 0.1, y + rh / 2 - 0.09, 0.18, 0.18, fill="E07A1F", shape=MSO_SHAPE.DIAMOND, line=WHITE)
    d.text(s, "◆ = Go-live MVP/Pilot · Farbe = Horizont (Cyan Now, Blau Next, Grau Later)", 0.5, d.H - 0.7,
           d.W - 1, 0.3, size=9, color=GREY)

    now = [i for i in r["items"] if i["horizont"] == "Now"]
    if now:
        s = d.slide("MVP-Definition und Pilotierung (Horizont NOW)")
        rows = [["Use Case", "MVP-Umfang", "Pilotbereich", "Erfolgskriterien/KPIs", "Start"]] + [
            [f"{i['id']} {i['titel']}", i.get("mvp") or "im Kick-off festlegen", i.get("pilot") or "–",
             i.get("kpi") or "–", i["start"]] for i in now]
        d.table(s, rows, 0.5, 1.55, d.W - 1, col_w=[2.4, 3.4, 2.2, 3.2, 0.8], size=10, row_h=0.7)


def _raci(d: Deck, r: dict) -> None:
    cols = {"R": "D6EAF8", "A": "F5B7B1", "C": "FCF3CF", "I": "E5E7EB"}
    for k, part in enumerate(chunks(r["items"], 12)):
        s = d.slide("RACI-Matrix" + (f" ({k + 1})" if k else ""),
                    "R = Responsible · A = Accountable (genau 1) · C = Consulted · I = Informed")
        rows = [["Aufgabe / Vorhaben"] + r["rollen"] + ["Prüfung"]]
        for it in part:
            rows.append([it["aufgabe"]] + [it["werte"].get(x, "") for x in r["rollen"]] +
                        ["✓" if not it["probleme"] else "; ".join(it["probleme"])])
        d.table(s, rows, 0.5, 1.55, d.W - 1, col_w=[3.2] + [1] * len(r["rollen"]) + [1.6], size=10, row_h=0.4,
                fill_fn=lambda i, j, v: cols.get(str(v)[:1]) if 0 < j <= len(r["rollen"]) and v else None,
                color_fn=lambda i, j, v: RED if j == len(r["rollen"]) + 1 and v != "✓" else
                (GREEN if v == "✓" else None))
