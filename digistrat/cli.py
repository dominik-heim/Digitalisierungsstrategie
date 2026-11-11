"""Kommandozeile: `digistrat <befehl> …` (oder `python -m digistrat …`)."""
from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path

from .model import METHODS, PHASES, methods_by_phase


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(prog="digistrat", description="KI-gestützter Methodenkoffer Digitalisierungsstrategie")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("init", help="Neues Kundenprojekt anlegen")
    p.add_argument("ordner")
    p.add_argument("--kunde", default="Muster GmbH")
    p.add_argument("--branche", default="Produzierendes Unternehmen")
    p.add_argument("--berater", default="")
    p.add_argument("--roadmap-start", default=None, help="z.B. 2027-Q1")

    p = sub.add_parser("demo", help="Beispielprojekt 'Muster Maschinenbau' anlegen und auswerten")
    p.add_argument("ordner")

    sub.add_parser("methoden", help="Methodenkoffer anzeigen")

    p = sub.add_parser("vorlage", help="Excel-Erfassungsvorlage (neu) erzeugen")
    p.add_argument("ordner")
    p.add_argument("--methoden", nargs="*", help=f"Auswahl aus: {', '.join(METHODS)}")
    p.add_argument("--datei", default=None)

    p = sub.add_parser("workshops", help="Workshop-Serie planen und Agenden/Leitfäden/Vorlagen erzeugen")
    p.add_argument("ordner")
    p.add_argument("--start", help="Datum des ersten Workshops (JJJJ-MM-TT)")
    p.add_argument("--intervall", type=int, default=7, help="Tage zwischen Workshops")
    p.add_argument("--ohne-optional", action="store_true", help="optionalen Fit-Gap-Workshop weglassen")
    p.add_argument("--ki", action="store_true", help="kundenspezifische Leitfragen per KI erzeugen")
    p.add_argument("--kontext", default="", help="Zusatzkontext für die KI (Ziele, Besonderheiten)")
    p.add_argument("--ja", action="store_true", help="Datenübermittlung an Anthropic ohne Rückfrage bestätigen")

    p = sub.add_parser("ki-import", help="Workshop-Rohdaten (Fotos, Notizen, PDFs) per KI in ein Methoden-Tool übernehmen")
    p.add_argument("ordner")
    p.add_argument("--methode", required=True, choices=[k for k, m in METHODS.items() if not m.wide])
    p.add_argument("--datei", nargs="*", default=[], help="Fotos (jpg/png), PDF, DOCX, PPTX, XLSX, TXT/MD")
    p.add_argument("--text", default=None, help="Notizen als Text")
    p.add_argument("--kontext", default="")
    p.add_argument("--ja", action="store_true", help="Datenübermittlung an Anthropic ohne Rückfrage bestätigen")

    p = sub.add_parser("ki-ableiten", help="Folge-Artefakte per KI vorschlagen (z.B. User Stories aus Pain Points)")
    p.add_argument("ordner")
    p.add_argument("--ziel", required=True,
                   choices=["handlungsfelder", "userstories", "usecases", "vpc", "fitgap", "roadmap", "whys"])
    p.add_argument("--hinweis", default="")
    p.add_argument("--ja", action="store_true", help="Datenübermittlung an Anthropic ohne Rückfrage bestätigen")

    p = sub.add_parser("pruefen", help="Workshop-Input auf Vollständigkeit/Plausibilität prüfen")
    p.add_argument("ordner")

    p = sub.add_parser("auswerten", help="Alles auswerten → PowerPoint, Excel, BPMN")
    p.add_argument("ordner")
    p.add_argument("--ki", action="store_true", help="Management Summary per KI formulieren")
    p.add_argument("--ja", action="store_true", help="Datenübermittlung an Anthropic ohne Rückfrage bestätigen")

    p = sub.add_parser("ki", help="KI für ein Projekt erlauben oder sperren")
    p.add_argument("ordner")
    p.add_argument("status", choices=["an", "aus", "status"])

    p = sub.add_parser("app", help="Web-Oberfläche starten (Streamlit)")

    a = ap.parse_args(argv)
    return globals()[f"cmd_{a.cmd.replace('-', '_')}"](a) or 0


def _confirm(lines: list[str], yes: bool) -> bool:
    print("Diese Daten werden an Anthropic (Claude) gesendet:")
    for x in lines:
        print(f"  - {x}")
    if any(x.startswith("⚠") for x in lines):
        print("  (⚠ = enthält personenbezogene Angaben)")
    if yes:
        return True
    try:
        return input("Übermittlung bestätigen? [j/N] ").strip().lower() in ("j", "ja", "y", "yes")
    except EOFError:
        return False


def _ki_allowed(p) -> bool:
    if not p.config.get("ki", {}).get("aktiv", True):
        print("🔒 KI ist für dieses Projekt deaktiviert – es werden keine Daten gesendet. "
              "Freigabe: digistrat ki <ordner> an")
        return False
    return True


def cmd_ki(a) -> int:
    from .project import Project
    p = Project.open(a.ordner)
    if a.status != "status":
        p.config.setdefault("ki", {})["aktiv"] = a.status == "an"
        p.save_config()
    on = p.config.get("ki", {}).get("aktiv", True)
    print("KI für dieses Projekt: " + ("erlaubt" if on else "🔒 deaktiviert"))
    return 0


def cmd_init(a) -> None:
    from .project import Project
    kw = {"kunde": a.kunde, "branche": a.branche, "berater": a.berater}
    if a.roadmap_start:
        kw["roadmap_start"] = a.roadmap_start
    p = Project.create(a.ordner, **kw)
    print(f"Projekt angelegt: {p.root}\n  Erfassungsvorlage: {p.input_dir / 'Workshop-Erfassung.xlsx'}\n"
          f"  Nächster Schritt: digistrat workshops {a.ordner}")


def cmd_demo(a) -> None:
    from .demo import create_demo
    from .pipeline import run
    from .workshop import plan, render
    p = create_demo(Path(a.ordner))
    plan(p, start=None)
    for k, f in render(p).items():
        if k in ("xlsx", "pptx"):
            print(f"→ Workshop-{k}: {f}")
    run(p)


def cmd_methoden(a) -> None:
    for ph, ms in methods_by_phase().items():
        print(f"\n{PHASES[ph]}")
        for m in ms:
            print(f"  {m.key:<16} {m.title:<48} Blatt '{m.sheet}'")


def cmd_vorlage(a) -> None:
    from .project import Project, write_input_templates
    p = Project.open(a.ordner)
    f = Path(a.datei) if a.datei else p.input_dir / "Workshop-Erfassung_leer.xlsx"
    print(write_input_templates(f, a.methoden or None))


def cmd_workshops(a) -> None:
    from .project import Project
    from .workshop import load, plan, render
    p = Project.open(a.ordner)
    f = plan(p, start=a.start, interval_days=a.intervall, include_optional=not a.ohne_optional)
    print(f"Planung: {f} (Termine/Blöcke dort anpassen und Befehl erneut ausführen)")
    questions = None
    if a.ki and _ki_allowed(p):
        from .ai import describe_payload, workshop_questions
        if _confirm(describe_payload("questions", p.config, context=a.kontext), a.ja):
            questions = workshop_questions(p.config, load(p), a.kontext)
        else:
            print("Abgebrochen – Unterlagen werden ohne KI-Leitfragen erzeugt.")
    for k, v in render(p, questions).items():
        print(f"→ {k}: {v}")


def cmd_ki_import(a) -> int:
    from .ai import describe_payload, extract_rows, save_rows
    from .project import Project
    p = Project.open(a.ordner)
    if not _ki_allowed(p):
        return 1
    if not _confirm(describe_payload("import", p.config, files=a.datei, text=a.text, context=a.kontext), a.ja):
        print("Abgebrochen – nichts gesendet.")
        return 1
    rows, hint = extract_rows(a.methode, a.text, [Path(x) for x in a.datei], p.config, a.kontext)
    f = save_rows(p, a.methode, rows)
    print(f"{len(rows)} Zeilen → {f}\nBitte prüfen/ergänzen, dann 'digistrat auswerten'.")
    if hint:
        print(f"Hinweise der KI: {hint}")
    return 0


def cmd_ki_ableiten(a) -> int:
    from .ai import describe_payload, derive_rows, save_rows
    from .analysis import analyze
    from .project import Project
    p = Project.open(a.ordner)
    if not _ki_allowed(p):
        return 1
    res = analyze(p.load(), p.config)
    if not _confirm(describe_payload("derive", p.config, results=res), a.ja):
        print("Abgebrochen – nichts gesendet.")
        return 1
    rows, hint = derive_rows(a.ziel, res, p.config, a.hinweis)
    f = save_rows(p, a.ziel, rows)
    print(f"{len(rows)} Vorschläge → {f}\nBitte prüfen: nicht übernommene Zeilen löschen.")
    if hint:
        print(f"Hinweise der KI: {hint}")
    return 0


def cmd_pruefen(a) -> int:
    from .project import Project
    d = Project.open(a.ordner).load()
    for k, m in METHODS.items():
        n = len(d.get(k)) or len(d.wide.get(k, ([], []))[1])
        print(f"  {'✓' if n else '·'} {m.title:<48} {n:>3} Einträge")
    for i in d.issues:
        print(f"  ⚠ {i}")
    return 1 if d.issues else 0


def cmd_auswerten(a) -> None:
    from .analysis import analyze
    from .pipeline import run
    from .project import Project
    p = Project.open(a.ordner)
    use_ai = a.ki and _ki_allowed(p)
    if use_ai:
        from .ai import describe_payload
        use_ai = _confirm(describe_payload("summary", p.config, results=analyze(p.load(), p.config)), a.ja)
        if not use_ai:
            print("Ohne KI fortgesetzt – regelbasierte Management Summary.")
    run(p, use_ai=use_ai)


def cmd_app(a) -> None:
    app = Path(__file__).resolve().parent.parent / "app" / "streamlit_app.py"
    subprocess.call([sys.executable, "-m", "streamlit", "run", str(app), "--server.address", "localhost",
                     "--browser.gatherUsageStats", "false"])


if __name__ == "__main__":
    sys.exit(main())
