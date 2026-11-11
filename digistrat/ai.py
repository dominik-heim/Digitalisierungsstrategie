"""KI-Unterstützung mit Claude (Anthropic API).

Funktionen
  * extract_rows   – Workshop-Rohdaten (Notizen, Transkripte, Fotos von Flipcharts/Stickies, PDFs, Miro-Exporte)
                     strukturiert in das Schema eines Methoden-Tools überführen.
  * derive_rows    – Folge-Artefakte vorschlagen: User Stories aus Pain Points/Ursachen, Handlungsfelder,
                     Use-Case-Steckbriefe, VPC, Fit-Gap-Anforderungen, Roadmap-MVPs …
  * ai_management_summary – Kernaussagen für das Management aus allen Analyseergebnissen.
  * workshop_questions    – kundenspezifische Leitfragen je Workshop.

Alle KI-Ergebnisse werden als separate Dateien (KI_<Blatt>.xlsx) in 01_Workshop-Input abgelegt und sind
damit vor der Auswertung prüf- und änderbar ("Human in the Loop").

Authentifizierung: ANTHROPIC_API_KEY (oder `ant auth login`). Modell/Effort in projekt.yaml unter `ki`.
"""
from __future__ import annotations

import base64
import json
import mimetypes
from pathlib import Path
from typing import Any

from .model import METHODS, Field, Method

DEFAULT_MODEL = "claude-opus-5-5"

SYSTEM = (
    "Du bist ein erfahrener Berater für Digitalisierungsstrategien in produzierenden Unternehmen "
    "(Lean Management, Prozessmanagement, ERP/MES, Industrie 4.0). Du arbeitest präzise, auf Deutsch, "
    "und erfindest keine Fakten: Was nicht in den Workshop-Unterlagen steht, bleibt leer bzw. null. "
    "Formuliere knapp und so, wie es ein Fachbereich im Workshop sagen würde."
)


class AIError(RuntimeError):
    pass


# ---------------------------------------------------------------- Basis
def is_enabled(cfg: dict | None) -> bool:
    """KI ist je Projekt abschaltbar (projekt.yaml → ki.aktiv)."""
    return bool((cfg or {}).get("ki", {}).get("aktiv", True))


def ensure_enabled(cfg: dict | None) -> None:
    if not is_enabled(cfg):
        raise AIError("KI ist für dieses Projekt deaktiviert (projekt.yaml → ki.aktiv: false). "
                      "Es werden keine Daten an Anthropic gesendet.")


def _client():
    try:
        import anthropic
    except ImportError as e:  # pragma: no cover
        raise AIError("Paket 'anthropic' fehlt: pip install anthropic") from e
    return anthropic.Anthropic()


def _json_type(f: Field) -> dict[str, Any]:
    if f.type == "int":
        return {"anyOf": [{"type": "integer"}, {"type": "null"}], "description": f.header}
    if f.type == "float":
        return {"anyOf": [{"type": "number"}, {"type": "null"}], "description": f.header}
    if f.type == "choice" and f.choices:
        return {"type": "string", "enum": f.choices + [""], "description": f.header}
    desc = f.header + (" – mehrere IDs kommagetrennt" if f.type == "ids" else "") + (f" ({f.help})" if f.help else "")
    return {"type": "string", "description": desc}


def rows_schema(m: Method, extra_fields: list[Field] | None = None) -> dict[str, Any]:
    fields = list(m.fields) + list(extra_fields or [])
    item = {"type": "object", "properties": {f.key: _json_type(f) for f in fields},
            "required": [f.key for f in fields], "additionalProperties": False}
    return {"type": "object", "properties": {"rows": {"type": "array", "items": item},
                                              "hinweise": {"type": "string",
                                                           "description": "Unklarheiten/Annahmen für den Berater"}},
            "required": ["rows", "hinweise"], "additionalProperties": False}


def ask_json(content: list[dict] | str, schema: dict, cfg: dict | None = None, system: str = SYSTEM) -> dict:
    """Ein Claude-Aufruf mit strukturierter Ausgabe (JSON-Schema)."""
    ensure_enabled(cfg)
    ki = (cfg or {}).get("ki", {})
    client = _client()
    with client.beta.messages.stream(
        model=ki.get("model", DEFAULT_MODEL),
        max_tokens=64000,
        system=system,
        messages=[{"role": "user", "content": content}],
        output_config={"effort": ki.get("effort", "high"), "format": {"type": "json_schema", "schema": schema}},
        # Bei einer Sicherheits-Ablehnung beantwortet die API die Anfrage automatisch mit einem Fallback-Modell.
        betas=["server-side-fallback-2026-07-01"],
        fallbacks="default",
    ) as stream:
        resp = stream.get_final_message()
    if resp.stop_reason == "refusal":
        raise AIError("Anfrage wurde vom Modell abgelehnt.")
    if resp.stop_reason == "max_tokens":
        raise AIError("Antwort abgeschnitten (max_tokens) – Eingabe aufteilen.")
    text = next((b.text for b in resp.content if b.type == "text"), None)
    if not text:
        raise AIError("Keine Textantwort erhalten.")
    return json.loads(text)


# ---------------------------------------------------------------- Dateien → Content-Blöcke
def file_blocks(paths: list[Path]) -> list[dict]:
    blocks: list[dict] = []
    for p in paths:
        p = Path(p)
        mime = mimetypes.guess_type(p.name)[0] or ""
        if mime in ("image/png", "image/jpeg", "image/gif", "image/webp"):
            blocks.append({"type": "image", "source": {"type": "base64", "media_type": mime,
                                                       "data": base64.standard_b64encode(p.read_bytes()).decode()}})
            blocks.append({"type": "text", "text": f"(Foto: {p.name})"})
        elif mime == "application/pdf":
            blocks.append({"type": "document", "source": {"type": "base64", "media_type": "application/pdf",
                                                          "data": base64.standard_b64encode(p.read_bytes()).decode()},
                           "title": p.name})
        elif p.suffix.lower() == ".docx":
            from docx import Document  # python-docx
            txt = "\n".join(par.text for par in Document(p).paragraphs)
            blocks.append({"type": "text", "text": f"<datei name='{p.name}'>\n{txt}\n</datei>"})
        elif p.suffix.lower() in (".xlsx", ".xlsm"):
            from openpyxl import load_workbook
            wb = load_workbook(p, data_only=True, read_only=True)
            parts = []
            for ws in wb.worksheets:
                parts.append(f"## Blatt {ws.title}")
                for r in ws.iter_rows(values_only=True):
                    if any(v is not None for v in r):
                        parts.append(" | ".join("" if v is None else str(v) for v in r))
            blocks.append({"type": "text", "text": f"<datei name='{p.name}'>\n" + "\n".join(parts) + "\n</datei>"})
        elif p.suffix.lower() == ".pptx":
            from pptx import Presentation
            parts = []
            for i, s in enumerate(Presentation(p).slides, 1):
                texts = [sh.text_frame.text for sh in s.shapes if sh.has_text_frame and sh.text_frame.text.strip()]
                parts.append(f"## Folie {i}\n" + "\n".join(texts))
            blocks.append({"type": "text", "text": f"<datei name='{p.name}'>\n" + "\n".join(parts) + "\n</datei>"})
        else:
            blocks.append({"type": "text",
                           "text": f"<datei name='{p.name}'>\n{p.read_text(encoding='utf-8', errors='replace')}\n</datei>"})
    return blocks


# ---------------------------------------------------------------- Extraktion
def extract_rows(method_key: str, text: str | None = None, files: list[Path] | None = None,
                 cfg: dict | None = None, context: str = "") -> tuple[list[dict], str]:
    ensure_enabled(cfg)
    m = METHODS[method_key]
    if m.wide:
        raise AIError(f"{m.title}: Matrix-Methoden bitte direkt in Excel erfassen.")
    content = file_blocks(files or [])
    if text:
        content.append({"type": "text", "text": f"<notizen>\n{text}\n</notizen>"})
    if not content:
        raise AIError("Keine Eingabe (Text oder Dateien) angegeben.")
    content.append({"type": "text", "text": (
        f"Aufgabe: Überführe die obigen Workshop-Ergebnisse in das Methoden-Tool „{m.title}“.\n"
        f"Zweck der Methode: {m.purpose}\n"
        f"Spalten: {', '.join(f.header for f in m.fields)}.\n"
        f"{_method_rules(method_key)}\n"
        f"Kontext: {context or '–'}\n"
        "Übernimm jede inhaltlich eigenständige Aussage als eigene Zeile, fasse Dubletten zusammen. "
        "Lies Fotos von Stickies/Flipcharts vollständig aus. Nicht genannte Bewertungen = null.")})
    out = ask_json(content, rows_schema(m), cfg)
    return out["rows"], out.get("hinweise", "")


def _method_rules(key: str) -> str:
    return {
        "painpoints": "IDs fortlaufend PP-01, PP-02 … vergeben. Kategorie passend wählen.",
        "userstories": ("Format strikt: Rolle / Ziel / Nutzen. Akzeptanzkriterien je Zeile im Format "
                        "'Given … When … Then …'. INVEST beachten: ein Anliegen je Story, ≤ 8 Story Points, "
                        "keine Lösung vorwegnehmen. IDs US-01 …"),
        "usecases": "IDs UC-01 …; Problem, Lösung, Nutzen, Datenbasis, Aufwand und Owner möglichst konkret.",
        "prozess": ("Schritte chronologisch nummerieren; Typ Start/Aufgabe/Entscheidung/Ende; bei Entscheidungen "
                    "Nachfolger-Nummern angeben; Medium (Papier, Excel, E-Mail …) und System je Schritt."),
        "vsm": "Zeiten in Minuten umrechnen (1 Tag = 1440 min, 1 Schicht = 480 min).",
        "fitgap": "IDs FG-01 …; Abdeckung Standard/Konfiguration/Erweiterung/Gap.",
        "whys": "Je Pain Point eine Zeile mit der Kette Warum 1-5 und der Grundursache.",
    }.get(key, "")


# ---------------------------------------------------------------- Ableitung
DERIVE_PROMPTS = {
    "handlungsfelder": ("Clustere die Pain Points, Ursachen und User Stories zu 3-6 Handlungsfeldern der "
                        "Digitalisierung (prägnante Namen, Zielbild messbar formuliert)."),
    "userstories": ("Leite aus Pain Points und Grundursachen User Stories ab (je relevante Rolle). Strikt INVEST, "
                    "Akzeptanzkriterien Given/When/Then, Pain-Point-IDs referenzieren, Aktivität (Story-Map-"
                    "Backbone), Handlungsfeld und Release (MVP zuerst) vorschlagen. Vorhandene Stories nicht doppeln."),
    "usecases": ("Leite aus Handlungsfeldern, User Stories und Pain Points konkrete Digitalisierungs-Use-Cases ab "
                 "(typische Lösungen im produzierenden Umfeld: BDE/MDE, MES, APS, CAQ, DMS/PDM am Shopfloor, "
                 "IoT/OEE, KI-Assistenzen, Workflow-Automatisierung, Self-Service-Portale). Ein Steckbrief je "
                 "Use Case; Impact/Effort nur grob schätzen (1-10), wenn begründbar, sonst null."),
    "vpc": ("Erstelle für die Use Cases mit Service-/Kundencharakter Value Proposition Canvas-Einträge. Jeder "
            "Pain Reliever/Gain Creator nennt in 'Bezug' den adressierten Pain/Gain wörtlich."),
    "fitgap": "Leite aus den User Stories prüfbare Anforderungen für eine Fit-Gap-Analyse ab (Abdeckung leer lassen, "
              "wenn unbekannt → 'Gap' nur wenn klar).",
    "roadmap": ("Schlage für die priorisierten Use Cases MVP-Umfang, Pilotbereich und messbare KPIs vor; "
                "Horizont Now/Next/Later gemäß Priorisierung, Abhängigkeiten angeben."),
    "whys": "Führe für die Top-Pain-Points eine plausible 5-Why-Analyse als Hypothese durch (im Workshop validieren).",
}


def derive_rows(target: str, results: dict, cfg: dict | None = None, hint: str = "") -> tuple[list[dict], str]:
    if target not in DERIVE_PROMPTS:
        raise AIError(f"Ableitung für '{target}' nicht verfügbar. Möglich: {', '.join(DERIVE_PROMPTS)}")
    ensure_enabled(cfg)
    m = METHODS[target]
    payload = json.dumps(_slim(derive_context(results)), ensure_ascii=False, default=str)
    content = [{"type": "text", "text": f"<projekt>\nKunde: {(cfg or {}).get('kunde')}, Branche: "
                                        f"{(cfg or {}).get('branche')}\n</projekt>\n<analyse>\n{payload}\n</analyse>"},
               {"type": "text", "text": f"Aufgabe: {DERIVE_PROMPTS[target]}\nZiel-Methode: {m.title} "
                                        f"(Spalten: {', '.join(f.header for f in m.fields)}).\n"
                                        f"{_method_rules(target)}\nZusatzhinweis des Beraters: {hint or '–'}"}]
    out = ask_json(content, rows_schema(m), cfg)
    return out["rows"], out.get("hinweise", "")


DERIVE_KEYS = ("painpoints", "ursachen", "userstories", "handlungsfelder", "usecases", "nwa", "impact_effort",
               "roadmap", "prozess", "vsm")


def derive_context(results: dict) -> dict:
    return {k: results[k] for k in DERIVE_KEYS if k in results}


def summary_context(results: dict) -> dict:
    return {k: v for k, v in results.items() if k != "issues"}


# ---------------------------------------------------------------- Transparenz: was wird gesendet?
SECTION_LABELS = {
    "stakeholder": "Stakeholder-Matrix (Namen, Rollen, Einfluss/Interesse, Haltung)",
    "sipoc": "SIPOC", "prozess": "Ist-Prozesse (Schritte, Rollen, Systeme, Medien)",
    "vsm": "Wertstrom (Zeiten, Qualität, Systeme)", "painpoints": "Pain Points",
    "ursachen": "Ursachenanalysen (5 Whys, Ishikawa)", "userstories": "User Stories",
    "storymap": "Story Map", "handlungsfelder": "Handlungsfelder", "usecases": "Use-Case-Steckbriefe",
    "vpc": "Value Proposition Canvas", "bmc": "Business Model Canvas", "fitgap": "Fit-Gap-Anforderungen",
    "nwa": "Nutzwertanalyse", "impact_effort": "Impact/Effort-Bewertung", "raci": "RACI (Rollen/Aufgaben)",
    "businesscase": "Business Case (Kosten, Einsparungen)", "roadmap": "Roadmap", "kennzahlen": "Kennzahlen",
}

PERSONAL = {"stakeholder", "raci"}


def _count(v: Any) -> str:
    if isinstance(v, dict):
        for k in ("items", "prozesse", "canvases", "use_cases", "whys"):
            if k in v and hasattr(v[k], "__len__"):
                return f" ({len(v[k])})"
    return ""


def describe_sections(ctx: dict) -> list[str]:
    out = []
    for k, v in ctx.items():
        lab = SECTION_LABELS.get(k, k) + _count(v)
        out.append(("⚠ " if k in PERSONAL else "") + lab)
    return out


def describe_payload(kind: str, cfg: dict | None, results: dict | None = None, files: list | None = None,
                     text: str | None = None, context: str = "") -> list[str]:
    """Für den Hinweis vor jedem KI-Aufruf: Liste der Daten, die an Anthropic gesendet werden."""
    cfg = cfg or {}
    out: list[str] = []
    if kind == "import":
        for f in files or []:
            out.append(f"Datei: {getattr(f, 'name', f)}")
        if text:
            out.append(f"Notizen/Transkript ({len(text)} Zeichen)")
        if context:
            out.append(f"Kontexttext ({len(context)} Zeichen)")
    elif kind == "derive":
        out.append(f"Kunde und Branche: {cfg.get('kunde')}, {cfg.get('branche')}")
        out += describe_sections(derive_context(results or {}))
    elif kind == "summary":
        out += describe_sections(summary_context(results or {}))
    elif kind == "questions":
        out.append(f"Kunde und Branche: {cfg.get('kunde')}, {cfg.get('branche')}")
        out.append("Titel, Ziele und Blöcke der Workshops")
        if context:
            out.append(f"Kontexttext ({len(context)} Zeichen)")
    model = cfg.get("ki", {}).get("model", DEFAULT_MODEL)
    out.append(f"Empfänger: Anthropic-API (Modell {model})")
    return out


def _slim(obj: Any) -> Any:
    """Entfernt sehr große, redundante Teile (z.B. Rohschritte in Gantt-Daten) aus dem Kontext."""
    if isinstance(obj, dict):
        return {k: _slim(v) for k, v in obj.items() if k not in ("succ", "grid", "quartale")}
    if isinstance(obj, list):
        return [_slim(x) for x in obj]
    return obj


# ---------------------------------------------------------------- Texte
def ai_management_summary(results: dict, cfg: dict) -> list[str]:
    schema = {"type": "object", "properties": {"kernaussagen": {"type": "array", "items": {"type": "string"}}},
              "required": ["kernaussagen"], "additionalProperties": False}
    ensure_enabled(cfg)
    payload = json.dumps(_slim(summary_context(results)), ensure_ascii=False, default=str)
    out = ask_json([{"type": "text", "text": f"<analyse>\n{payload}\n</analyse>"},
                    {"type": "text", "text": (
                        f"Formuliere 6-8 Kernaussagen für die Management Summary der Digitalisierungsstrategie von "
                        f"{cfg.get('kunde')} ({cfg.get('branche')}). Pyramidenprinzip: erst Gesamtaussage, dann "
                        "Belege mit Zahlen aus der Analyse (Pain Points, Medienbrüche, Flussgrad, Quick Wins, "
                        "Nutzwert-Ranking, Business Case, Roadmap). Jede Aussage max. 30 Wörter, keine erfundenen Zahlen.")}],
                   schema, cfg)
    return out["kernaussagen"]


def workshop_questions(cfg: dict, workshops: list[dict], context: str = "") -> dict[str, list[str]]:
    ensure_enabled(cfg)
    schema = {"type": "object", "properties": {"workshops": {"type": "array", "items": {
        "type": "object", "properties": {"id": {"type": "string"}, "fragen": {"type": "array", "items": {"type": "string"}}},
        "required": ["id", "fragen"], "additionalProperties": False}}},
        "required": ["workshops"], "additionalProperties": False}
    ws = [{"id": w["id"], "titel": w["titel"], "ziel": w["ziel"],
           "bloecke": [b["titel"] for b in w["bloecke"]]} for w in workshops]
    out = ask_json([{"type": "text", "text": (
        f"Kunde: {cfg.get('kunde')}, Branche: {cfg.get('branche')}. Kontext: {context or '–'}\n"
        f"Workshops: {json.dumps(ws, ensure_ascii=False)}\n"
        "Formuliere je Workshop 6-10 offene, branchenspezifische Leitfragen für die Moderation "
        "(z.B. für Pain-Point-Sammlung, Ursachenklärung, Bewertung).")}], schema, cfg)
    return {w["id"]: w["fragen"] for w in out["workshops"]}


def save_rows(project, method_key: str, rows: list[dict]) -> Path:
    from .project import append_rows
    m = METHODS[method_key]
    return append_rows(project.input_dir / f"KI_{m.sheet}.xlsx", method_key, rows)
