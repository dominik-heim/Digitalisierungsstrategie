"""Workshop-Planung: Workshop-Serie, Agenden, Teilnehmer, Moderationsleitfaden, Erfassungsvorlagen.

Ablauf:
  1. `plan(project)` legt 00_Workshopplanung/workshops.yaml an (editierbar: Termine, Dauer, Blöcke, Ort).
  2. `render(project)` erzeugt daraus Workshopplan.xlsx, Workshop-Agenden.pptx, Moderationsleitfäden (.md)
     und je Workshop eine Erfassungsvorlage, die nur die benötigten Methodenblätter enthält.
"""
from __future__ import annotations

import datetime as dt
from pathlib import Path
from typing import Any

import yaml
from openpyxl.styles import Alignment, Font

from .model import METHODS, PHASES
from .project import Project, write_input_templates

# Workshop-Katalog: (id, titel, phase, ziel, bloecke[(titel, minuten, methode|None, hinweis)], teilnehmer-hinweis)
CATALOG: list[dict[str, Any]] = [
    {"id": "WS0", "titel": "Kick-off & Stakeholder-Analyse", "phase": "ist",
     "ziel": "Projektauftrag, Zielbild und Erwartungen klären; Stakeholder und Einbindung festlegen",
     "teilnehmer": "Sponsor, Projektleitung, Kernteam (Stakeholder-Matrix vertraulich vorbereiten)",
     "bloecke": [("Begrüßung, Ziele, Projektauftrag", 20, None, "Ziele, Scope, Erfolgskriterien"),
                 ("Erwartungen & Befürchtungen", 25, None, "Kartenabfrage: Was muss am Ende erreicht sein?"),
                 ("Stakeholder-Mapping", 45, "stakeholder", ""),
                 ("Pause", 15, None, ""),
                 ("Kernprozesse & Scope festlegen", 30, None, "Welche 2-4 Prozesse werden aufgenommen?"),
                 ("Workshop-Serie & Termine abstimmen", 20, None, ""),
                 ("Zusammenfassung & nächste Schritte", 15, None, "")]},
    {"id": "WS1", "titel": "Ist-Aufnahme Prozesse & Wertstrom", "phase": "ist",
     "ziel": "Kernprozesse abgrenzen und im Detail aufnehmen; Zeiten, Medien und Systeme sichtbar machen",
     "teilnehmer": "Prozessbeteiligte je Kernprozess (operativ), Prozessverantwortliche, IT",
     "bloecke": [("Begrüßung & Ziele", 15, None, ""),
                 ("Prozessrahmen SIPOC", 45, "sipoc", ""),
                 ("Ist-Prozess Swimlane/BPMN", 120, "prozess", ""),
                 ("Mittagspause", 45, None, ""),
                 ("Wertstrom: Zeiten, Bestände, Qualität", 90, "vsm", "Vorab Gemba-Walk empfohlen"),
                 ("Zusammenfassung & nächste Schritte", 15, None, "")]},
    {"id": "WS2", "titel": "Pain Points & Ursachenanalyse", "phase": "ist",
     "ziel": "Schmerzpunkte am Prozessbild sammeln, priorisieren und auf Grundursachen zurückführen",
     "teilnehmer": "Prozessbeteiligte, Führungskräfte der Bereiche, Qualität",
     "bloecke": [("Begrüßung & Rückblick Prozessbilder", 15, None, ""),
                 ("Pain Points sammeln & voten", 90, "painpoints", ""),
                 ("Pause", 15, None, ""),
                 ("5 Whys für Top-Pain-Points", 60, "whys", ""),
                 ("Ishikawa für Kernproblem", 45, "ishikawa", ""),
                 ("Zusammenfassung & nächste Schritte", 15, None, "")]},
    {"id": "WS3", "titel": "Anforderungen: User Stories & Story Map", "phase": "ist",
     "ziel": "Anforderungen aus Nutzersicht formulieren, Story Map aufbauen, Handlungsfelder ableiten",
     "teilnehmer": "Key User je Rolle, Product Owner/Projektleitung, IT",
     "bloecke": [("Begrüßung & Rollen klären", 15, None, ""),
                 ("User Stories schreiben (INVEST, Given/When/Then)", 90, "userstories", ""),
                 ("Pause", 15, None, ""),
                 ("User Story Mapping & Handlungsfelder", 90, "handlungsfelder", ""),
                 ("Zusammenfassung & nächste Schritte", 15, None, "")]},
    {"id": "WS4", "titel": "Potenzial-Workshop: Use Cases", "phase": "synthese",
     "ziel": "Lösungsideen je Handlungsfeld als Use-Case-Steckbriefe ausarbeiten und Wertbeitrag prüfen",
     "teilnehmer": "Bereichsleitungen, Key User, IT, ggf. Vertrieb/Service für neue Services",
     "bloecke": [("Begrüßung & Handlungsfelder", 15, None, ""),
                 ("Inspiration: Technologie-/Best-Practice-Impulse", 30, None, "Beispiele aus der Branche"),
                 ("Use-Case-Steckbriefe in Kleingruppen", 90, "usecases", ""),
                 ("Mittagspause", 45, None, ""),
                 ("Gallery Walk", 30, None, "Rückfragen, Dubletten zusammenführen"),
                 ("Value Proposition Canvas (neue Services)", 60, "vpc", ""),
                 ("Business Model Canvas (neue Services)", 45, "bmc", "nur bei Geschäftsmodell-Ideen"),
                 ("Zusammenfassung & nächste Schritte", 15, None, "")]},
    {"id": "WS5", "titel": "Fit-Gap-Analyse Standardsoftware", "phase": "synthese", "optional": True,
     "ziel": "Anforderungen gegen Standard (z.B. SAP, MES, APS) abgleichen, Gaps und Aufwand bewerten",
     "teilnehmer": "Key User, IT, Anbieter/Implementierungspartner",
     "bloecke": [("Begrüßung & Vorgehen", 15, None, ""),
                 ("Fit-Gap je Prozess (Show & Tell)", 180, "fitgap", ""),
                 ("Mittagspause", 45, None, ""),
                 ("Gaps bewerten: Prozess anpassen vs. System erweitern", 45, None, ""),
                 ("Zusammenfassung & nächste Schritte", 15, None, "")]},
    {"id": "WS6", "titel": "Bewertung & Priorisierung", "phase": "bewertung",
     "ziel": "Use Cases transparent bewerten (Nutzwert, Impact/Effort, Business Case) und priorisieren",
     "teilnehmer": "Geschäftsführung/Entscheider, Bereichsleitungen, Controlling, IT",
     "bloecke": [("Begrüßung & Use-Case-Übersicht", 15, None, ""),
                 ("Kriterien & Gewichtung bestätigen", 30, "kriterien", ""),
                 ("Scoring je Use Case", 90, "nwa", ""),
                 ("Pause", 15, None, ""),
                 ("Business Case der Top-Use-Cases", 60, "businesscase", ""),
                 ("Impact/Effort-Matrix diskutieren", 30, None, "Automatisch aus Scoring erzeugt"),
                 ("Zusammenfassung & nächste Schritte", 15, None, "")]},
    {"id": "WS7", "titel": "Roadmap & Verantwortlichkeiten", "phase": "bewertung",
     "ziel": "Now/Next/Later-Roadmap mit MVP, Pilot und KPIs festlegen; Verantwortlichkeiten klären",
     "teilnehmer": "Geschäftsführung, Projektleitung, Use-Case-Owner, IT, Betriebsrat (Information)",
     "bloecke": [("Begrüßung & Priorisierungsergebnis", 15, None, ""),
                 ("Roadmap Now/Next/Later, MVP & Pilot", 120, "roadmap", ""),
                 ("Pause", 15, None, ""),
                 ("RACI-Matrix", 45, "raci", ""),
                 ("Governance, Budget, Kommunikation", 30, None, ""),
                 ("Zusammenfassung & Entscheidungsvorlage", 15, None, "")]},
]


def plan(project: Project, start: str | None = None, interval_days: int = 7, include_optional: bool = True,
         start_time: str = "09:00") -> Path:
    """Erzeugt workshops.yaml mit Terminvorschlag (überschreibt nicht, wenn vorhanden)."""
    f = project.workshop_dir / "workshops.yaml"
    if f.exists():
        return f
    d0 = dt.date.fromisoformat(start) if start else _next_monday()
    out = []
    k = 0
    for w in CATALOG:
        if w.get("optional") and not include_optional:
            continue
        out.append({"id": w["id"], "titel": w["titel"], "phase": w["phase"], "ziel": w["ziel"],
                    "datum": (d0 + dt.timedelta(days=interval_days * k)).isoformat(), "beginn": start_time,
                    "ort": "vor Ort beim Kunden", "teilnehmer_hinweis": w["teilnehmer"],
                    "bloecke": [{"titel": b[0], "minuten": b[1], "methode": b[2], "hinweis": b[3]}
                                for b in w["bloecke"]]})
        k += 1
    f.parent.mkdir(parents=True, exist_ok=True)
    f.write_text(yaml.safe_dump({"workshops": out}, allow_unicode=True, sort_keys=False), encoding="utf-8")
    return f


def _next_monday() -> dt.date:
    t = dt.date.today()
    return t + dt.timedelta(days=(7 - t.weekday()) % 7 or 7)


def load(project: Project) -> list[dict[str, Any]]:
    f = project.workshop_dir / "workshops.yaml"
    if not f.exists():
        plan(project)
    ws = yaml.safe_load(f.read_text(encoding="utf-8"))["workshops"]
    for w in ws:
        h, m = map(int, str(w.get("beginn", "09:00")).split(":"))
        t = dt.datetime.combine(dt.date.today(), dt.time(h, m))
        for b in w["bloecke"]:
            b["von"] = t.strftime("%H:%M")
            t += dt.timedelta(minutes=int(b["minuten"]))
            b["bis"] = t.strftime("%H:%M")
        w["ende"] = t.strftime("%H:%M")
        w["dauer_h"] = sum(int(b["minuten"]) for b in w["bloecke"]) / 60
    return ws


def participants(project: Project) -> dict[str, list[dict]]:
    """Schlägt Teilnehmer je Workshop aus der Stakeholder-Matrix vor."""
    from .analysis import stakeholder
    try:
        rows = project.load().get("stakeholder")
    except FileNotFoundError:
        rows = []
    if not rows:
        return {}
    st = stakeholder(rows)["items"]
    out: dict[str, list[dict]] = {}
    for w in CATALOG:
        if w["id"] in ("WS0", "WS6", "WS7"):
            sel = [s for s in st if s["quadrant"] in ("Eng einbinden", "Zufriedenstellen")]
        else:
            sel = [s for s in st if s["quadrant"] in ("Eng einbinden", "Informieren")]
        out[w["id"]] = sel
    return out


def render(project: Project, questions: dict[str, list[str]] | None = None) -> dict[str, Path]:
    ws = load(project)
    tn = participants(project)
    files = {"xlsx": _excel(project, ws, tn), "pptx": _pptx(project, ws, tn, questions or {})}
    for w in ws:
        keys = [b["methode"] for b in w["bloecke"] if b.get("methode")]
        if keys:
            files[f"vorlage:{w['id']}"] = write_input_templates(
                project.workshop_dir / "Erfassungsvorlagen" / f"{w['id']}_{_safe(w['titel'])}.xlsx", keys)
        files[f"leitfaden:{w['id']}"] = _guide(project, w, tn.get(w["id"], []), (questions or {}).get(w["id"], []))
    return files


def _safe(s: str) -> str:
    import re
    return re.sub(r"_+", "_", "".join(c if c.isalnum() else "_" for c in s))[:40].strip("_")


def _excel(project: Project, ws: list[dict], tn: dict) -> Path:
    from .report_excel import XL
    x = XL(project.config.get("theme"))
    sh = x.sheet("Workshop-Serie", f"Workshop-Serie – {project.config.get('kunde')}",
                 "Termine/Blöcke in workshops.yaml anpassen und neu erzeugen")
    rows = [[w["id"], w["titel"], PHASES[w["phase"]], w["datum"], f"{w['beginn']}–{w['ende']}",
             f"{w['dauer_h']:.1f}", w["ziel"], ", ".join(METHODS[b["methode"]].title for b in w["bloecke"]
                                                        if b.get("methode")),
             ", ".join(p["name"] for p in tn.get(w["id"], [])) or w["teilnehmer_hinweis"], w.get("ort")]
            for w in ws]
    x.table(sh, 4, ["ID", "Workshop", "Phase", "Datum", "Zeit", "Dauer [h]", "Ziel", "Methoden", "Teilnehmer", "Ort"],
            rows, [6, 32, 28, 11, 13, 9, 45, 40, 40, 18])
    for w in ws:
        s = x.sheet(w["id"], f"{w['id']} {w['titel']}", f"{w['datum']} · {w['beginn']}–{w['ende']} · {w.get('ort')}")
        s["A3"] = "Ziel: " + w["ziel"]
        rows = []
        for b in w["bloecke"]:
            m = METHODS.get(b.get("methode") or "")
            rows.append([f"{b['von']}–{b['bis']}", b["titel"], b["minuten"], m.title if m else "",
                         (m.workshop if m else "") + (" " + b["hinweis"] if b.get("hinweis") else ""),
                         m.material if m else "", m.output if m else ""])
        x.table(s, 5, ["Zeit", "Block", "Min.", "Methode", "Ablauf/Hinweis", "Material", "Ergebnis"], rows,
                [13, 34, 6, 30, 60, 28, 34])
        r = 7 + len(rows)
        s.cell(r, 1, "Teilnehmer").font = Font(bold=True, color=x.primary)
        ppl = tn.get(w["id"], [])
        if ppl:
            x.table(s, r + 1, ["Name", "Rolle", "Bereich", "Quadrant", "Haltung", "Hinweis"],
                    [[p["name"], p["rolle"], p.get("bereich"), p["quadrant"], p.get("haltung"),
                      "Vorab Einzelgespräch" if p["risiko"] else ""] for p in ppl])
        else:
            s.cell(r + 1, 1, w["teilnehmer_hinweis"]).alignment = Alignment(wrap_text=True)
    f = project.workshop_dir / "Workshopplan.xlsx"
    x.wb.save(f)
    return f


def _pptx(project: Project, ws: list[dict], tn: dict, questions: dict) -> Path:
    from pptx.enum.shapes import MSO_SHAPE

    from .pptx_builder import GREY, WHITE, Deck
    cfg = project.config
    d = Deck(cfg.get("theme"), footer=f"{cfg.get('kunde')} · Workshop-Serie")
    d.title_slide("Workshop-Serie Digitalisierungsstrategie", cfg.get("kunde", ""),
                  f"{len(ws)} Workshops · {ws[0]['datum']} bis {ws[-1]['datum']}" if ws else "")
    s = d.slide("Überblick Workshop-Serie", "Von der Ist-Analyse zur priorisierten Roadmap")
    rows = [["ID", "Workshop", "Datum", "Zeit", "Ziel"]] + [
        [w["id"], w["titel"], w["datum"], f"{w['beginn']}–{w['ende']}", w["ziel"]] for w in ws]
    d.table(s, rows, 0.5, 1.55, d.W - 1, col_w=[0.6, 3.2, 1.2, 1.3, 6], size=11, row_h=0.6)
    for w in ws:
        s = d.slide(f"{w['id']} · {w['titel']}", f"{w['datum']} · {w['beginn']}–{w['ende']} · {w.get('ort')}")
        d.rect(s, 0.5, 1.55, d.W - 1, 0.5, fill=d.light, text="Ziel: " + w["ziel"], size=13, align="l")
        total = sum(int(b["minuten"]) for b in w["bloecke"])
        x = 0.5
        y = 2.3
        width = d.W - 1.0
        for b in w["bloecke"]:
            bw = width * int(b["minuten"]) / total
            col = GREY if "pause" in b["titel"].lower() else (d.accent if b.get("methode") else d.primary)
            d.rect(s, x, y, bw - 0.04, 0.9, fill=col, text=f"{b['von']}\n{b['titel']}", size=10, color=WHITE,
                   autosize=True)
            x += bw
        rows = [["Zeit", "Block", "Methode / Ablauf"]]
        for b in w["bloecke"]:
            m = METHODS.get(b.get("methode") or "")
            rows.append([f"{b['von']}–{b['bis']}", b["titel"],
                         (m.workshop if m else b.get("hinweis") or "")[:180]])
        d.table(s, rows, 0.5, 3.4, d.W - 4.4, col_w=[1.2, 2.8, 5.0], size=9, row_h=0.42)
        ppl = tn.get(w["id"], [])
        txt = "\n".join(f"• {p['name']} ({p['rolle']})" for p in ppl) or w["teilnehmer_hinweis"]
        d.rect(s, d.W - 3.7, 3.4, 3.2, 0.35, fill=d.primary, text="Teilnehmer", size=11, bold=True, color=WHITE)
        d.rect(s, d.W - 3.7, 3.75, 3.2, d.H - 4.5, fill=d.light, text=txt, size=10, align="l", anchor="t",
               autosize=True)
        if questions.get(w["id"]):
            s2 = d.slide(f"{w['id']} · Leitfragen für die Moderation", "KI-gestützt auf den Kundenkontext zugeschnitten")
            d.bullets(s2, questions[w["id"]], 0.5, 1.55, d.W - 1, d.H - 2.2, size=16)
    # Methodenkarten
    d.section("Methodenkarten", "Ziel, Ablauf, Material und Ergebnis je Methode")
    for m in METHODS.values():
        s = d.slide(f"Methodenkarte: {m.title}", PHASES[m.phase])
        boxes = [("Ziel", m.purpose), ("Ablauf im Workshop", m.workshop),
                 ("Material", m.material), ("Ergebnis", m.output),
                 ("Dauer", f"ca. {m.duration_min} Minuten"), ("Tipp", m.tips or "–")]
        bw, bh = (d.W - 1.15) / 2, (d.H - 2.3) / 3
        for i, (lab, val) in enumerate(boxes):
            bx, by = 0.5 + (i % 2) * (bw + 0.15), 1.55 + (i // 2) * (bh + 0.05)
            d.rect(s, bx, by, bw, 0.34, fill=d.primary, text=lab, size=12, bold=True, color=WHITE, align="l")
            d.rect(s, bx, by + 0.34, bw, bh - 0.36, fill=d.light, text=val, size=13, align="l", anchor="t",
                   autosize=True)
        d.text(s, f"Erfassung: Blatt '{m.sheet}' – Spalten: " + ", ".join(m.headers[:10]), 0.5, d.H - 0.68,
               d.W - 1, 0.3, size=8, color=GREY)
    f = project.workshop_dir / "Workshop-Agenden.pptx"
    d.prs.save(f)
    return f


def _guide(project: Project, w: dict, ppl: list[dict], questions: list[str]) -> Path:
    lines = [f"# Moderationsleitfaden {w['id']}: {w['titel']}", "",
             f"**Kunde:** {project.config.get('kunde')}  ", f"**Termin:** {w['datum']}, {w['beginn']}–{w['ende']}  ",
             f"**Ort:** {w.get('ort')}  ", f"**Ziel:** {w['ziel']}", "", "## Teilnehmer", ""]
    lines += [f"- {p['name']} – {p['rolle']}" + (" ⚠ vorab Einzelgespräch (kritische Haltung)" if p["risiko"] else "")
              for p in ppl] or [f"- {w['teilnehmer_hinweis']}"]
    lines += ["", "## Ablauf", ""]
    for b in w["bloecke"]:
        m = METHODS.get(b.get("methode") or "")
        lines.append(f"### {b['von']}–{b['bis']} · {b['titel']} ({b['minuten']} min)")
        if m:
            lines += ["", f"*Methode:* {m.title} – {m.purpose}", "", f"*Ablauf:* {m.workshop}", "",
                      f"*Material:* {m.material}  ", f"*Ergebnis:* {m.output}  ",
                      f"*Erfassung:* Blatt `{m.sheet}` in der Erfassungsvorlage"]
            if m.tips:
                lines.append(f"*Tipp:* {m.tips}")
        if b.get("hinweis"):
            lines += ["", f"Hinweis: {b['hinweis']}"]
        lines.append("")
    if questions:
        lines += ["## Leitfragen (KI-generiert, bitte prüfen)", ""] + [f"- {q}" for q in questions] + [""]
    lines += ["## Nach dem Workshop", "",
              "1. Ergebnisse (Fotos, Miro-Export, Notizen) in die Erfassungsvorlage übertragen – "
              "oder mit `digistrat ki-import` automatisch extrahieren lassen.",
              "2. Datei nach `01_Workshop-Input/` legen.", "3. `digistrat auswerten` ausführen → PowerPoint & Excel."]
    f = project.workshop_dir / "Moderationsleitfaeden" / f"{w['id']}_{_safe(w['titel'])}.md"
    f.parent.mkdir(parents=True, exist_ok=True)
    f.write_text("\n".join(lines), encoding="utf-8")
    return f
