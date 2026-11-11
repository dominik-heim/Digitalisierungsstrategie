"""Methoden-Registry: zentrale Definition aller Methoden-Tools.

Jede Methode beschreibt
  * die Phase im Vorgehensmodell,
  * das Erfassungsschema (Spalten der Excel-Eingabevorlage),
  * Methodenkarte (Ziel, Ablauf im Workshop, Dauer, Material, Ergebnis).

Aus dieser einen Quelle werden Excel-Vorlagen, Validierung, KI-Import-Schemata,
Workshop-Agenden und Methodenkarten erzeugt.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

PHASES = {
    "ist": "1 · Ist-Analyse & Problemverständnis",
    "synthese": "2 · Synthese",
    "bewertung": "3 · Bewertung, Priorisierung & Roadmap",
}


@dataclass
class Field:
    key: str
    header: str
    type: str = "str"  # str | text | int | float | choice | ids
    required: bool = False
    choices: list[str] | None = None
    min: float | None = None
    max: float | None = None
    help: str = ""


@dataclass
class Method:
    key: str
    sheet: str
    phase: str
    title: str
    purpose: str
    fields: list[Field] = field(default_factory=list)
    wide: bool = False  # Matrix-Format mit dynamischen Spalten (NWA, RACI)
    workshop: str = ""  # Ablauf im Workshop
    duration_min: int = 60
    material: str = ""
    output: str = ""
    tips: str = ""

    @property
    def headers(self) -> list[str]:
        return [f.header for f in self.fields]

    def field_by_header(self) -> dict[str, Field]:
        return {f.header.lower(): f for f in self.fields}


S15 = dict(type="int", min=1, max=5)
S110 = dict(type="int", min=1, max=10)

METHODS: dict[str, Method] = {}


def _m(m: Method) -> Method:
    METHODS[m.key] = m
    return m


# ---------------------------------------------------------------- Phase 1
_m(Method(
    "stakeholder", "Stakeholder", "ist", "Stakeholder-Mapping (Einfluss/Interesse)",
    "Klärt, wer entscheidet, wer betroffen ist und wer blockieren kann. Gehört vor jeden Workshop.",
    [
        Field("name", "Name", required=True),
        Field("rolle", "Rolle/Funktion", required=True),
        Field("bereich", "Bereich"),
        Field("einfluss", "Einfluss (1-5)", required=True, **S15),
        Field("interesse", "Interesse (1-5)", required=True, **S15),
        Field("haltung", "Haltung", type="choice",
              choices=["Promotor", "Unterstützer", "Neutral", "Skeptiker", "Blockierer"]),
        Field("erwartungen", "Erwartungen/Befürchtungen", type="text"),
        Field("massnahme", "Maßnahme/Einbindung", type="text"),
    ],
    workshop="Vorab im Projektteam/mit Sponsor: Personen sammeln, Einfluss und Interesse je 1-5 "
             "einschätzen, Haltung markieren, Einbindungsstrategie je Quadrant ableiten.",
    duration_min=45, material="Matrix-Poster oder Miro-Board, Namenskarten",
    output="Stakeholder-Matrix, Einbindungsstrategie, Teilnehmerliste der Workshops",
    tips="Vertraulich behandeln – Haltungs-Einschätzung nicht im Kundenworkshop zeigen.",
))

_m(Method(
    "sipoc", "SIPOC", "ist", "Prozessrahmen (SIPOC)",
    "Liefert den groben Rahmen eines Prozesses: Lieferanten, Inputs, Prozessschritte, Outputs, Kunden.",
    [
        Field("prozess", "Prozess", required=True),
        Field("supplier", "Supplier"),
        Field("input", "Input"),
        Field("process", "Process (Hauptschritt)"),
        Field("output", "Output"),
        Field("customer", "Customer"),
    ],
    workshop="Mit 5 Spalten an der Wand starten: zuerst Process (5-7 Hauptschritte), dann Output & "
             "Customer, danach Input & Supplier. Start- und Endpunkt explizit festlegen.",
    duration_min=45, material="5 Spalten-Poster, Stickies", output="Abgegrenzter Prozessrahmen je Kernprozess",
))

_m(Method(
    "prozess", "Prozessaufnahme", "ist", "Ist-Prozessaufnahme (BPMN/Swimlane)",
    "Detailliertes Ist-Modell mit Rollen (Swimlanes), Systemen und Medien. Basis für BPMN-Export und Medienbruch-Analyse.",
    [
        Field("prozess", "Prozess", required=True),
        Field("nr", "Nr", type="int", required=True),
        Field("schritt", "Prozessschritt", required=True),
        Field("lane", "Rolle/Lane", required=True),
        Field("system", "System/Tool"),
        Field("medium", "Medium", type="choice",
              choices=["Digital integriert", "Excel", "E-Mail", "Papier", "Mündlich/Telefon", "Manuell/ohne Medium"]),
        Field("typ", "Typ", type="choice", choices=["Start", "Aufgabe", "Entscheidung", "Ende"]),
        Field("nachfolger", "Nachfolger (Nr, kommagetrennt)", type="ids",
              help="Leer = nächster Schritt. Bei Entscheidungen mehrere Nr angeben, z.B. '5,7'."),
        Field("bemerkung", "Bemerkung", type="text"),
    ],
    workshop="Brown-Paper oder Miro: Swimlanes je Rolle anlegen, Schritte chronologisch kleben, "
             "pro Schritt System und Medium (Papier/Excel/E-Mail…) notieren. Entscheidungen als Raute.",
    duration_min=120, material="Brown-Paper, Stickies in 3 Farben, Rautenkarten",
    output="Swimlane-Diagramm, BPMN-2.0-Datei (Import in Signavio/Camunda/bpmn.io), Medienbruch-Liste",
))

_m(Method(
    "vsm", "Wertstrom", "ist", "Value Stream Mapping",
    "Macht Durchlaufzeiten, Wartezeiten, Qualitätsverluste und Medienbrüche sichtbar.",
    [
        Field("prozess", "Prozess", required=True),
        Field("nr", "Nr", type="int", required=True),
        Field("schritt", "Prozessschritt", required=True),
        Field("bz", "Bearbeitungszeit [min]", type="float", required=True),
        Field("wz", "Wartezeit davor [min]", type="float"),
        Field("rz", "Rüstzeit [min]", type="float"),
        Field("qr", "Qualitätsrate/FPY [%]", type="float", min=0, max=100),
        Field("system", "System/Informationsfluss"),
        Field("medium", "Medium", type="choice",
              choices=["Digital integriert", "Excel", "E-Mail", "Papier", "Mündlich/Telefon", "Manuell/ohne Medium"]),
        Field("wertschoepfend", "Wertschöpfend", type="choice", choices=["Ja", "Nein", "Notwendig"]),
        Field("bestand", "Bestand/Queue [Stk/Vorgänge]", type="float"),
    ],
    workshop="Gemba/Prozessbegehung, anschließend Wertstrom am Board: je Prozesskasten Bearbeitungszeit, "
             "Wartezeit, FPY und Informationsfluss erfassen. Zeiten mit Stoppuhr/Systemdaten plausibilisieren.",
    duration_min=180, material="VSM-Symbole, Datenkästen, Stoppuhr, ggf. System-Auswertungen",
    output="Durchlaufzeit, Flussgrad, Rolled Throughput Yield, Zeitleiste, Medienbrüche",
))

_m(Method(
    "painpoints", "PainPoints", "ist", "Pain-Point-Analyse",
    "Teilnehmende sammeln Schmerzpunkte direkt am Prozessbild (rote Stickies) und bewerten Häufigkeit und Auswirkung.",
    [
        Field("id", "ID", required=True, help="z.B. PP-01"),
        Field("prozess", "Prozess"),
        Field("schritt", "Prozessschritt"),
        Field("beschreibung", "Beschreibung", type="text", required=True),
        Field("kategorie", "Kategorie", type="choice",
              choices=["Medienbruch", "Transparenz", "Qualität", "Durchlaufzeit", "Kosten",
                       "IT-System", "Datenqualität", "Organisation", "Kompetenz"]),
        Field("haeufigkeit", "Häufigkeit (1-5)", **S15),
        Field("auswirkung", "Auswirkung (1-5)", **S15),
        Field("stimmen", "Stimmen (Dot-Voting)", type="int", min=0),
        Field("quelle", "Genannt von"),
    ],
    workshop="Prozessbild aufhängen, 10 Min. stilles Brainstorming mit roten Stickies direkt am Schritt, "
             "Clustering, dann Dot-Voting (3 Punkte p.P.) und gemeinsame Einschätzung von Häufigkeit/Auswirkung.",
    duration_min=90, material="Rote Stickies, Klebepunkte, Prozessposter",
    output="Priorisierte Pain-Point-Liste, Pareto, Heatmap Prozess × Kategorie",
))

_m(Method(
    "whys", "5Whys", "ist", "Ursachenanalyse – 5 Whys",
    "Fragt je priorisiertem Pain Point fünfmal ‚Warum?‘ – verhindert, dass nur Symptome digitalisiert werden.",
    [
        Field("pain_point_id", "Pain-Point-ID", required=True),
        Field("problem", "Problem/Symptom", type="text", required=True),
        Field("w1", "Warum 1", type="text"),
        Field("w2", "Warum 2", type="text"),
        Field("w3", "Warum 3", type="text"),
        Field("w4", "Warum 4", type="text"),
        Field("w5", "Warum 5", type="text"),
        Field("grundursache", "Grundursache", type="text"),
        Field("digital_loesbar", "Digital lösbar?", type="choice", choices=["Ja", "Teilweise", "Nein"]),
    ],
    workshop="Top-Pain-Points (aus Voting) in Kleingruppen: Kette bis zur beeinflussbaren Grundursache, "
             "anschließend prüfen: ist die Ursache digital, organisatorisch oder prozessual lösbar?",
    duration_min=60, material="5-Why-Vorlage je Pain Point", output="Ursachenketten, Grundursachen",
))

_m(Method(
    "ishikawa", "Ishikawa", "ist", "Ursachenanalyse – Ishikawa",
    "Strukturiert mögliche Ursachen eines Problems entlang der Kategorien (6M + Daten/IT).",
    [
        Field("problem", "Problem (Fischkopf)", required=True),
        Field("kategorie", "Kategorie", type="choice", required=True,
              choices=["Mensch", "Maschine", "Methode", "Material", "Messung", "Mitwelt", "Management", "Daten/IT"]),
        Field("ursache", "Ursache", type="text", required=True),
        Field("bestaetigt", "Bestätigt?", type="choice", choices=["Ja", "Vermutet", "Nein"]),
    ],
    workshop="Fischgräte auf Poster, Problem im Kopf. Je Kategorie Ursachen sammeln, danach mit Daten/"
             "Erfahrung bestätigen oder verwerfen.",
    duration_min=60, material="Fischgräten-Poster", output="Ishikawa-Diagramm je Kernproblem",
))

_m(Method(
    "userstories", "UserStories", "ist", "User Stories (INVEST, Given/When/Then)",
    "Anforderungen aus Nutzersicht: „Als [Rolle] möchte ich [Ziel], damit [Nutzen]“ mit Akzeptanzkriterien.",
    [
        Field("id", "ID", required=True, help="z.B. US-01"),
        Field("rolle", "Als (Rolle)", required=True),
        Field("ziel", "möchte ich (Ziel)", type="text", required=True),
        Field("nutzen", "damit (Nutzen)", type="text", required=True),
        Field("akzeptanz", "Akzeptanzkriterien (Given/When/Then)", type="text",
              help="Je Kriterium eine Zeile: 'Given … When … Then …' bzw. 'Gegeben … Wenn … Dann …'"),
        Field("aktivitaet", "Aktivität (Story-Map-Backbone)"),
        Field("handlungsfeld", "Handlungsfeld"),
        Field("release", "Release", type="choice", choices=["MVP", "Release 2", "Release 3", "Backlog"]),
        Field("prioritaet", "Priorität (MoSCoW)", type="choice", choices=["Must", "Should", "Could", "Won't"]),
        Field("story_points", "Story Points", type="int", min=0),
        Field("pain_point_ids", "Pain-Point-IDs", type="ids"),
    ],
    workshop="Je Rolle aus Pain Points und Grundursachen Stories formulieren, INVEST-Check im Plenum, "
             "Akzeptanzkriterien im Given/When/Then-Format mit Fachbereich schärfen.",
    duration_min=90, material="Story-Karten (Vorlage), INVEST-Checkliste",
    output="Story-Backlog mit INVEST-Prüfung",
))

# ---------------------------------------------------------------- Phase 2
_m(Method(
    "handlungsfelder", "Handlungsfelder", "synthese", "User Story Mapping & Handlungsfelder",
    "Ordnet Stories entlang der Nutzer-Aktivitäten (Backbone) und Releases; verdichtet sie zu Handlungsfeldern mit Potenzial.",
    [
        Field("name", "Handlungsfeld", required=True),
        Field("beschreibung", "Beschreibung", type="text"),
        Field("zielbild", "Zielbild", type="text"),
        Field("verantwortlich", "Verantwortlich"),
    ],
    workshop="Story-Karten unter Aktivitäten anordnen (links→rechts = Nutzerreise, oben→unten = Priorität), "
             "Release-Linien ziehen, Cluster als Handlungsfelder benennen.",
    duration_min=120, material="Story-Map-Wand, Release-Klebeband",
    output="Story Map, Handlungsfeld-Landkarte mit Potenzial",
))

_m(Method(
    "usecases", "UseCases", "synthese", "Use-Case-Steckbriefe",
    "Halten pro Idee Problem, Nutzen, Datenbasis, Aufwand und Owner fest – Kern des Potenzial-Workshops.",
    [
        Field("id", "ID", required=True, help="z.B. UC-01"),
        Field("titel", "Titel", required=True),
        Field("handlungsfeld", "Handlungsfeld"),
        Field("problem", "Problem/Ausgangslage", type="text"),
        Field("loesung", "Lösungsidee", type="text"),
        Field("nutzen", "Nutzen (qualitativ/quantitativ)", type="text"),
        Field("datenbasis", "Datenbasis", type="text"),
        Field("systeme", "Systeme/Technologie"),
        Field("aufwand", "Aufwand/Voraussetzungen", type="text"),
        Field("owner", "Owner"),
        Field("pain_point_ids", "Pain-Point-IDs", type="ids"),
        Field("story_ids", "User-Story-IDs", type="ids"),
        Field("impact", "Impact (1-10)", help="Optional – sonst aus Nutzwertanalyse abgeleitet", **S110),
        Field("effort", "Effort (1-10)", help="Optional – sonst aus Nutzwertanalyse abgeleitet", **S110),
    ],
    workshop="Je Use Case ein Steckbrief-Poster in Kleingruppen (20 Min.), danach Gallery Walk mit Rückfragen.",
    duration_min=120, material="Steckbrief-Vorlage A3", output="Use-Case-Steckbriefe",
))

_m(Method(
    "vpc", "ValueProposition", "synthese", "Value Proposition Canvas",
    "Prüft den Wertbeitrag (Fit zwischen Kunden-Jobs/Pains/Gains und dem Angebot), v.a. bei neuen digitalen Services.",
    [
        Field("use_case_id", "Use-Case-ID", required=True),
        Field("segment", "Kundensegment"),
        Field("bereich", "Bereich", type="choice", required=True,
              choices=["Customer Job", "Pain", "Gain", "Produkt/Service", "Pain Reliever", "Gain Creator"]),
        Field("eintrag", "Eintrag", type="text", required=True),
        Field("bezug", "Bezug (adressierter Pain/Gain)"),
    ],
    workshop="Rechts (Kundenprofil) starten: Jobs, Pains, Gains. Dann links Produkte, Pain Reliever, "
             "Gain Creator – je Reliever den adressierten Pain referenzieren.",
    duration_min=90, material="VPC-Poster", output="Value Proposition Canvas mit Fit-Check",
))

_m(Method(
    "bmc", "BusinessModel", "synthese", "Business Model Canvas",
    "Beschreibt das Geschäftsmodell eines digitalen Services in 9 Bausteinen.",
    [
        Field("use_case_id", "Use-Case-ID", required=True),
        Field("baustein", "Baustein", type="choice", required=True,
              choices=["Schlüsselpartner", "Schlüsselaktivitäten", "Schlüsselressourcen", "Wertangebot",
                       "Kundenbeziehungen", "Kanäle", "Kundensegmente", "Kostenstruktur", "Einnahmequellen"]),
        Field("eintrag", "Eintrag", type="text", required=True),
    ],
    workshop="Reihenfolge: Kundensegmente → Wertangebot → Kanäle → Beziehungen → Einnahmen → "
             "Ressourcen → Aktivitäten → Partner → Kosten.",
    duration_min=90, material="BMC-Poster", output="Business Model Canvas",
))

_m(Method(
    "fitgap", "FitGap", "synthese", "Fit-Gap-Analyse",
    "Gleicht Anforderungen mit dem Standard ab – Pflicht bei SAP- und Standardsoftware-Einführungen.",
    [
        Field("id", "ID", required=True),
        Field("anforderung", "Anforderung", type="text", required=True),
        Field("prozess", "Prozess"),
        Field("prioritaet", "Priorität", type="choice", choices=["Muss", "Soll", "Kann"]),
        Field("system", "System/Modul"),
        Field("abdeckung", "Abdeckung", type="choice", required=True,
              choices=["Standard", "Konfiguration", "Erweiterung", "Gap"]),
        Field("loesung", "Lösungsansatz", type="text"),
        Field("aufwand_pt", "Aufwand [PT]", type="float", min=0),
    ],
    workshop="Anforderungen je Prozess im Standard-System demonstrieren (Show & Tell), Abdeckung gemeinsam "
             "einstufen. Gaps: Prozess anpassen vor System anpassen.",
    duration_min=180, material="Demo-System, Anforderungsliste", output="Fit-Gap-Liste, Abdeckungsgrad, Aufwand",
))

# ---------------------------------------------------------------- Phase 3
_m(Method(
    "kriterien", "NWA_Kriterien", "bewertung", "Nutzwertanalyse – Kriterien",
    "Gewichtete Kriterien machen Entscheidungen nachvollziehbar, auch gegenüber dem Management.",
    [
        Field("kriterium", "Kriterium", required=True),
        Field("kategorie", "Kategorie", type="choice", required=True,
              choices=["Nutzen", "Machbarkeit", "Risiko", "Strategie"]),
        Field("gewicht", "Gewicht [%]", type="float", required=True, min=0, max=100),
        Field("skala", "Skala/Beschreibung (1-5)", type="text"),
    ],
    workshop="Kriterien vorab mit Sponsor abstimmen; Gewichte per Paarvergleich oder 100-Punkte-Verteilung.",
    duration_min=45, material="Kriterienliste", output="Gewichtetes Kriterienset",
    tips="Machbarkeit/Risiko so formulieren, dass 5 = gut (geringer Aufwand / geringes Risiko).",
))

_m(Method(
    "nwa", "NWA_Bewertung", "bewertung", "Nutzwertanalyse / Scoring",
    "Bewertet jeden Use Case je Kriterium (1-5). Matrix: Zeilen = Use Cases, Spalten = Kriterien.",
    [Field("use_case_id", "Use-Case-ID", required=True)],
    wide=True,
    workshop="Bewertung je Kriterium im Plenum oder per Online-Voting (Median), Ausreißer diskutieren.",
    duration_min=90, material="Bewertungsmatrix, Voting-Tool", output="Gewichteter Score & Ranking",
))

_m(Method(
    "roadmap", "Roadmap", "bewertung", "Roadmap (Now/Next/Later, MVP, Pilot)",
    "Überführt priorisierte Use Cases in Horizonte mit MVP-Definition, Pilotbereich und KPIs.",
    [
        Field("use_case_id", "Use-Case-ID", required=True),
        Field("horizont", "Horizont", type="choice", choices=["Now", "Next", "Later"],
              help="Leer = automatischer Vorschlag aus Priorisierung"),
        Field("start", "Start (JJJJ-Qn)", help="z.B. 2027-Q1; leer = automatisch"),
        Field("dauer", "Dauer [Monate]", type="int", min=1),
        Field("mvp", "MVP-Umfang", type="text"),
        Field("pilot", "Pilotbereich", type="text"),
        Field("kpi", "Erfolgskriterien/KPIs", type="text"),
        Field("abhaengigkeiten", "Abhängigkeiten (Use-Case-IDs)", type="ids"),
    ],
    workshop="Vorschlag aus Priorisierung zeigen, Abhängigkeiten und Kapazitäten prüfen, je Now-Vorhaben "
             "MVP-Schnitt, Pilotbereich und messbare Erfolgskriterien festlegen.",
    duration_min=120, material="Roadmap-Board", output="Now/Next/Later-Roadmap, Gantt, MVP/Pilot-Steckbriefe",
))

_m(Method(
    "raci", "RACI", "bewertung", "RACI-Matrix",
    "Klärt Verantwortlichkeiten: Responsible, Accountable (genau eine), Consulted, Informed.",
    [Field("aufgabe", "Aufgabe/Vorhaben", required=True)],
    wide=True,
    workshop="Zeilen = Vorhaben/Aufgaben, Spalten = Rollen. Pro Zeile genau ein A, mindestens ein R.",
    duration_min=45, material="RACI-Matrix", output="Geprüfte RACI-Matrix",
))

_m(Method(
    "businesscase", "BusinessCase", "bewertung", "Business Case & ROI",
    "Macht aus Ideen budgetierbare Vorhaben: Investitionen, laufende Kosten, Einsparungen über die Jahre.",
    [
        Field("use_case_id", "Use-Case-ID", required=True),
        Field("position", "Position", required=True),
        Field("typ", "Typ", type="choice", required=True,
              choices=["Investition", "Laufende Kosten", "Einsparung", "Mehrumsatz"]),
        Field("jahr", "Jahr (0 = Invest-Jahr)", type="int", required=True, min=0, max=10),
        Field("betrag", "Betrag [EUR]", type="float", required=True, min=0,
              help="Immer positiv eintragen – das Vorzeichen ergibt sich aus dem Typ."),
        Field("annahme", "Annahme/Herleitung", type="text"),
    ],
    workshop="Mengengerüst mit Controlling: Zeiteinsparung × Stundensatz, vermiedene Fehlerkosten, "
             "Lizenz/Projekt-/Betriebskosten. Annahmen dokumentieren.",
    duration_min=90, material="Business-Case-Vorlage", output="NPV, ROI, Amortisationszeit je Use Case",
))

DEFAULT_CRITERIA = [
    ("Wirtschaftlicher Nutzen", "Nutzen", 25, "1 = kaum messbar … 5 = >250 T€/Jahr"),
    ("Beitrag zu Pain Points", "Nutzen", 20, "1 = einzelne Nennung … 5 = Top-Pain-Points gelöst"),
    ("Strategischer Fit", "Strategie", 15, "1 = kein Bezug … 5 = Kern der Unternehmensstrategie"),
    ("Umsetzbarkeit/Aufwand", "Machbarkeit", 20, "1 = sehr hoher Aufwand … 5 = geringer Aufwand"),
    ("Datenverfügbarkeit", "Machbarkeit", 10, "1 = Daten fehlen … 5 = Daten vorhanden & sauber"),
    ("Umsetzungsrisiko", "Risiko", 10, "1 = hohes Risiko … 5 = geringes Risiko"),
]

DEFAULT_RACI_ROLES = ["Geschäftsführung", "Projektleitung", "IT", "Fachbereich", "Produktion", "Qualität"]


def methods_by_phase() -> dict[str, list[Method]]:
    out: dict[str, list[Method]] = {p: [] for p in PHASES}
    for m in METHODS.values():
        out[m.phase].append(m)
    return out


def sheet_to_method(sheet: str) -> Method | None:
    s = sheet.strip().lower()
    for m in METHODS.values():
        if m.sheet.lower() == s or m.key == s:
            return m
    return None


def coerce(f: Field, value: Any) -> Any:
    """Konvertiert einen Zellwert entsprechend Feldtyp (tolerant)."""
    if value is None:
        return None
    if isinstance(value, float) and value != value:  # NaN
        return None
    if isinstance(value, str):
        value = value.strip()
        if value == "":
            return None
    if f.type == "int":
        try:
            return int(round(float(str(value).replace(",", "."))))
        except ValueError:
            return None
    if f.type == "float":
        try:
            return float(str(value).replace(",", ".").replace("%", ""))
        except ValueError:
            return None
    if f.type == "ids":
        return [x.strip() for x in str(value).replace(";", ",").split(",") if x.strip()]
    if f.type == "choice" and f.choices:
        sv = str(value).strip()
        for c in f.choices:
            if c.lower() == sv.lower():
                return c
        return sv
    return str(value)
