# Methodenkoffer

Automatisch aus `digistrat/model.py` erzeugt (`python -m digistrat methoden`). Jede Methode hat ein eigenes Erfassungsblatt in der Excel-Vorlage.

## 1 · Ist-Analyse & Problemverständnis

### Stakeholder-Mapping (Einfluss/Interesse)

**Ziel:** Klärt, wer entscheidet, wer betroffen ist und wer blockieren kann. Gehört vor jeden Workshop.  
**Ablauf im Workshop:** Vorab im Projektteam/mit Sponsor: Personen sammeln, Einfluss und Interesse je 1-5 einschätzen, Haltung markieren, Einbindungsstrategie je Quadrant ableiten.  
**Dauer:** ca. 45 min · **Material:** Matrix-Poster oder Miro-Board, Namenskarten  
**Ergebnis:** Stakeholder-Matrix, Einbindungsstrategie, Teilnehmerliste der Workshops  
**Tipp:** Vertraulich behandeln – Haltungs-Einschätzung nicht im Kundenworkshop zeigen.  

Blatt `Stakeholder`:

| Spalte | Typ | Pflicht | Werte/Hinweis |
|---|---|---|---|
| Name | str | ja |  |
| Rolle/Funktion | str | ja |  |
| Bereich | str |  |  |
| Einfluss (1-5) | int | ja | 1–5 |
| Interesse (1-5) | int | ja | 1–5 |
| Haltung | choice |  | Promotor, Unterstützer, Neutral, Skeptiker, Blockierer |
| Erwartungen/Befürchtungen | text |  |  |
| Maßnahme/Einbindung | text |  |  |

### Prozessrahmen (SIPOC)

**Ziel:** Liefert den groben Rahmen eines Prozesses: Lieferanten, Inputs, Prozessschritte, Outputs, Kunden.  
**Ablauf im Workshop:** Mit 5 Spalten an der Wand starten: zuerst Process (5-7 Hauptschritte), dann Output & Customer, danach Input & Supplier. Start- und Endpunkt explizit festlegen.  
**Dauer:** ca. 45 min · **Material:** 5 Spalten-Poster, Stickies  
**Ergebnis:** Abgegrenzter Prozessrahmen je Kernprozess  

Blatt `SIPOC`:

| Spalte | Typ | Pflicht | Werte/Hinweis |
|---|---|---|---|
| Prozess | str | ja |  |
| Supplier | str |  |  |
| Input | str |  |  |
| Process (Hauptschritt) | str |  |  |
| Output | str |  |  |
| Customer | str |  |  |

### Ist-Prozessaufnahme (BPMN/Swimlane)

**Ziel:** Detailliertes Ist-Modell mit Rollen (Swimlanes), Systemen und Medien. Basis für BPMN-Export und Medienbruch-Analyse.  
**Ablauf im Workshop:** Brown-Paper oder Miro: Swimlanes je Rolle anlegen, Schritte chronologisch kleben, pro Schritt System und Medium (Papier/Excel/E-Mail…) notieren. Entscheidungen als Raute.  
**Dauer:** ca. 120 min · **Material:** Brown-Paper, Stickies in 3 Farben, Rautenkarten  
**Ergebnis:** Swimlane-Diagramm, BPMN-2.0-Datei (Import in Signavio/Camunda/bpmn.io), Medienbruch-Liste  

Blatt `Prozessaufnahme`:

| Spalte | Typ | Pflicht | Werte/Hinweis |
|---|---|---|---|
| Prozess | str | ja |  |
| Nr | int | ja |  |
| Prozessschritt | str | ja |  |
| Rolle/Lane | str | ja |  |
| System/Tool | str |  |  |
| Medium | choice |  | Digital integriert, Excel, E-Mail, Papier, Mündlich/Telefon, Manuell/ohne Medium |
| Typ | choice |  | Start, Aufgabe, Entscheidung, Ende |
| Nachfolger (Nr, kommagetrennt) | ids |  | Leer = nächster Schritt. Bei Entscheidungen mehrere Nr angeben, z.B. '5,7'. |
| Bemerkung | text |  |  |

### Value Stream Mapping

**Ziel:** Macht Durchlaufzeiten, Wartezeiten, Qualitätsverluste und Medienbrüche sichtbar.  
**Ablauf im Workshop:** Gemba/Prozessbegehung, anschließend Wertstrom am Board: je Prozesskasten Bearbeitungszeit, Wartezeit, FPY und Informationsfluss erfassen. Zeiten mit Stoppuhr/Systemdaten plausibilisieren.  
**Dauer:** ca. 180 min · **Material:** VSM-Symbole, Datenkästen, Stoppuhr, ggf. System-Auswertungen  
**Ergebnis:** Durchlaufzeit, Flussgrad, Rolled Throughput Yield, Zeitleiste, Medienbrüche  

Blatt `Wertstrom`:

| Spalte | Typ | Pflicht | Werte/Hinweis |
|---|---|---|---|
| Prozess | str | ja |  |
| Nr | int | ja |  |
| Prozessschritt | str | ja |  |
| Bearbeitungszeit [min] | float | ja |  |
| Wartezeit davor [min] | float |  |  |
| Rüstzeit [min] | float |  |  |
| Qualitätsrate/FPY [%] | float |  | 0–100 |
| System/Informationsfluss | str |  |  |
| Medium | choice |  | Digital integriert, Excel, E-Mail, Papier, Mündlich/Telefon, Manuell/ohne Medium |
| Wertschöpfend | choice |  | Ja, Nein, Notwendig |
| Bestand/Queue [Stk/Vorgänge] | float |  |  |

### Pain-Point-Analyse

**Ziel:** Teilnehmende sammeln Schmerzpunkte direkt am Prozessbild (rote Stickies) und bewerten Häufigkeit und Auswirkung.  
**Ablauf im Workshop:** Prozessbild aufhängen, 10 Min. stilles Brainstorming mit roten Stickies direkt am Schritt, Clustering, dann Dot-Voting (3 Punkte p.P.) und gemeinsame Einschätzung von Häufigkeit/Auswirkung.  
**Dauer:** ca. 90 min · **Material:** Rote Stickies, Klebepunkte, Prozessposter  
**Ergebnis:** Priorisierte Pain-Point-Liste, Pareto, Heatmap Prozess × Kategorie  

Blatt `PainPoints`:

| Spalte | Typ | Pflicht | Werte/Hinweis |
|---|---|---|---|
| ID | str | ja | z.B. PP-01 |
| Prozess | str |  |  |
| Prozessschritt | str |  |  |
| Beschreibung | text | ja |  |
| Kategorie | choice |  | Medienbruch, Transparenz, Qualität, Durchlaufzeit, Kosten, IT-System, Datenqualität, Organisation, Kompetenz |
| Häufigkeit (1-5) | int |  | 1–5 |
| Auswirkung (1-5) | int |  | 1–5 |
| Stimmen (Dot-Voting) | int |  |  |
| Genannt von | str |  |  |

### Ursachenanalyse – 5 Whys

**Ziel:** Fragt je priorisiertem Pain Point fünfmal ‚Warum?‘ – verhindert, dass nur Symptome digitalisiert werden.  
**Ablauf im Workshop:** Top-Pain-Points (aus Voting) in Kleingruppen: Kette bis zur beeinflussbaren Grundursache, anschließend prüfen: ist die Ursache digital, organisatorisch oder prozessual lösbar?  
**Dauer:** ca. 60 min · **Material:** 5-Why-Vorlage je Pain Point  
**Ergebnis:** Ursachenketten, Grundursachen  

Blatt `5Whys`:

| Spalte | Typ | Pflicht | Werte/Hinweis |
|---|---|---|---|
| Pain-Point-ID | str | ja |  |
| Problem/Symptom | text | ja |  |
| Warum 1 | text |  |  |
| Warum 2 | text |  |  |
| Warum 3 | text |  |  |
| Warum 4 | text |  |  |
| Warum 5 | text |  |  |
| Grundursache | text |  |  |
| Digital lösbar? | choice |  | Ja, Teilweise, Nein |

### Ursachenanalyse – Ishikawa

**Ziel:** Strukturiert mögliche Ursachen eines Problems entlang der Kategorien (6M + Daten/IT).  
**Ablauf im Workshop:** Fischgräte auf Poster, Problem im Kopf. Je Kategorie Ursachen sammeln, danach mit Daten/Erfahrung bestätigen oder verwerfen.  
**Dauer:** ca. 60 min · **Material:** Fischgräten-Poster  
**Ergebnis:** Ishikawa-Diagramm je Kernproblem  

Blatt `Ishikawa`:

| Spalte | Typ | Pflicht | Werte/Hinweis |
|---|---|---|---|
| Problem (Fischkopf) | str | ja |  |
| Kategorie | choice | ja | Mensch, Maschine, Methode, Material, Messung, Mitwelt, Management, Daten/IT |
| Ursache | text | ja |  |
| Bestätigt? | choice |  | Ja, Vermutet, Nein |

### User Stories (INVEST, Given/When/Then)

**Ziel:** Anforderungen aus Nutzersicht: „Als [Rolle] möchte ich [Ziel], damit [Nutzen]“ mit Akzeptanzkriterien.  
**Ablauf im Workshop:** Je Rolle aus Pain Points und Grundursachen Stories formulieren, INVEST-Check im Plenum, Akzeptanzkriterien im Given/When/Then-Format mit Fachbereich schärfen.  
**Dauer:** ca. 90 min · **Material:** Story-Karten (Vorlage), INVEST-Checkliste  
**Ergebnis:** Story-Backlog mit INVEST-Prüfung  

Blatt `UserStories`:

| Spalte | Typ | Pflicht | Werte/Hinweis |
|---|---|---|---|
| ID | str | ja | z.B. US-01 |
| Als (Rolle) | str | ja |  |
| möchte ich (Ziel) | text | ja |  |
| damit (Nutzen) | text | ja |  |
| Akzeptanzkriterien (Given/When/Then) | text |  | Je Kriterium eine Zeile: 'Given … When … Then …' bzw. 'Gegeben … Wenn … Dann …' |
| Aktivität (Story-Map-Backbone) | str |  |  |
| Handlungsfeld | str |  |  |
| Release | choice |  | MVP, Release 2, Release 3, Backlog |
| Priorität (MoSCoW) | choice |  | Must, Should, Could, Won't |
| Story Points | int |  |  |
| Pain-Point-IDs | ids |  |  |

## 2 · Synthese

### User Story Mapping & Handlungsfelder

**Ziel:** Ordnet Stories entlang der Nutzer-Aktivitäten (Backbone) und Releases; verdichtet sie zu Handlungsfeldern mit Potenzial.  
**Ablauf im Workshop:** Story-Karten unter Aktivitäten anordnen (links→rechts = Nutzerreise, oben→unten = Priorität), Release-Linien ziehen, Cluster als Handlungsfelder benennen.  
**Dauer:** ca. 120 min · **Material:** Story-Map-Wand, Release-Klebeband  
**Ergebnis:** Story Map, Handlungsfeld-Landkarte mit Potenzial  

Blatt `Handlungsfelder`:

| Spalte | Typ | Pflicht | Werte/Hinweis |
|---|---|---|---|
| Handlungsfeld | str | ja |  |
| Beschreibung | text |  |  |
| Zielbild | text |  |  |
| Verantwortlich | str |  |  |

### Use-Case-Steckbriefe

**Ziel:** Halten pro Idee Problem, Nutzen, Datenbasis, Aufwand und Owner fest – Kern des Potenzial-Workshops.  
**Ablauf im Workshop:** Je Use Case ein Steckbrief-Poster in Kleingruppen (20 Min.), danach Gallery Walk mit Rückfragen.  
**Dauer:** ca. 120 min · **Material:** Steckbrief-Vorlage A3  
**Ergebnis:** Use-Case-Steckbriefe  

Blatt `UseCases`:

| Spalte | Typ | Pflicht | Werte/Hinweis |
|---|---|---|---|
| ID | str | ja | z.B. UC-01 |
| Titel | str | ja |  |
| Handlungsfeld | str |  |  |
| Problem/Ausgangslage | text |  |  |
| Lösungsidee | text |  |  |
| Nutzen (qualitativ/quantitativ) | text |  |  |
| Datenbasis | text |  |  |
| Systeme/Technologie | str |  |  |
| Aufwand/Voraussetzungen | text |  |  |
| Owner | str |  |  |
| Pain-Point-IDs | ids |  |  |
| User-Story-IDs | ids |  |  |
| Impact (1-10) | int |  | 1–10; Optional – sonst aus Nutzwertanalyse abgeleitet |
| Effort (1-10) | int |  | 1–10; Optional – sonst aus Nutzwertanalyse abgeleitet |

### Value Proposition Canvas

**Ziel:** Prüft den Wertbeitrag (Fit zwischen Kunden-Jobs/Pains/Gains und dem Angebot), v.a. bei neuen digitalen Services.  
**Ablauf im Workshop:** Rechts (Kundenprofil) starten: Jobs, Pains, Gains. Dann links Produkte, Pain Reliever, Gain Creator – je Reliever den adressierten Pain referenzieren.  
**Dauer:** ca. 90 min · **Material:** VPC-Poster  
**Ergebnis:** Value Proposition Canvas mit Fit-Check  

Blatt `ValueProposition`:

| Spalte | Typ | Pflicht | Werte/Hinweis |
|---|---|---|---|
| Use-Case-ID | str | ja |  |
| Kundensegment | str |  |  |
| Bereich | choice | ja | Customer Job, Pain, Gain, Produkt/Service, Pain Reliever, Gain Creator |
| Eintrag | text | ja |  |
| Bezug (adressierter Pain/Gain) | str |  |  |

### Business Model Canvas

**Ziel:** Beschreibt das Geschäftsmodell eines digitalen Services in 9 Bausteinen.  
**Ablauf im Workshop:** Reihenfolge: Kundensegmente → Wertangebot → Kanäle → Beziehungen → Einnahmen → Ressourcen → Aktivitäten → Partner → Kosten.  
**Dauer:** ca. 90 min · **Material:** BMC-Poster  
**Ergebnis:** Business Model Canvas  

Blatt `BusinessModel`:

| Spalte | Typ | Pflicht | Werte/Hinweis |
|---|---|---|---|
| Use-Case-ID | str | ja |  |
| Baustein | choice | ja | Schlüsselpartner, Schlüsselaktivitäten, Schlüsselressourcen, Wertangebot, Kundenbeziehungen, Kanäle, Kundensegmente, Kostenstruktur, Einnahmequellen |
| Eintrag | text | ja |  |

### Fit-Gap-Analyse

**Ziel:** Gleicht Anforderungen mit dem Standard ab – Pflicht bei SAP- und Standardsoftware-Einführungen.  
**Ablauf im Workshop:** Anforderungen je Prozess im Standard-System demonstrieren (Show & Tell), Abdeckung gemeinsam einstufen. Gaps: Prozess anpassen vor System anpassen.  
**Dauer:** ca. 180 min · **Material:** Demo-System, Anforderungsliste  
**Ergebnis:** Fit-Gap-Liste, Abdeckungsgrad, Aufwand  

Blatt `FitGap`:

| Spalte | Typ | Pflicht | Werte/Hinweis |
|---|---|---|---|
| ID | str | ja |  |
| Anforderung | text | ja |  |
| Prozess | str |  |  |
| Priorität | choice |  | Muss, Soll, Kann |
| System/Modul | str |  |  |
| Abdeckung | choice | ja | Standard, Konfiguration, Erweiterung, Gap |
| Lösungsansatz | text |  |  |
| Aufwand [PT] | float |  |  |

## 3 · Bewertung, Priorisierung & Roadmap

### Nutzwertanalyse – Kriterien

**Ziel:** Gewichtete Kriterien machen Entscheidungen nachvollziehbar, auch gegenüber dem Management.  
**Ablauf im Workshop:** Kriterien vorab mit Sponsor abstimmen; Gewichte per Paarvergleich oder 100-Punkte-Verteilung.  
**Dauer:** ca. 45 min · **Material:** Kriterienliste  
**Ergebnis:** Gewichtetes Kriterienset  
**Tipp:** Machbarkeit/Risiko so formulieren, dass 5 = gut (geringer Aufwand / geringes Risiko).  

Blatt `NWA_Kriterien`:

| Spalte | Typ | Pflicht | Werte/Hinweis |
|---|---|---|---|
| Kriterium | str | ja |  |
| Kategorie | choice | ja | Nutzen, Machbarkeit, Risiko, Strategie |
| Gewicht [%] | float | ja | 0–100 |
| Skala/Beschreibung (1-5) | text |  |  |

### Nutzwertanalyse / Scoring

**Ziel:** Bewertet jeden Use Case je Kriterium (1-5). Matrix: Zeilen = Use Cases, Spalten = Kriterien.  
**Ablauf im Workshop:** Bewertung je Kriterium im Plenum oder per Online-Voting (Median), Ausreißer diskutieren.  
**Dauer:** ca. 90 min · **Material:** Bewertungsmatrix, Voting-Tool  
**Ergebnis:** Gewichteter Score & Ranking  

Erfassung im Matrix-Format, Blatt `NWA_Bewertung`: erste Spalte *Use-Case-ID*, weitere Spalten = Kriterien (Werte 1-5).

### Roadmap (Now/Next/Later, MVP, Pilot)

**Ziel:** Überführt priorisierte Use Cases in Horizonte mit MVP-Definition, Pilotbereich und KPIs.  
**Ablauf im Workshop:** Vorschlag aus Priorisierung zeigen, Abhängigkeiten und Kapazitäten prüfen, je Now-Vorhaben MVP-Schnitt, Pilotbereich und messbare Erfolgskriterien festlegen.  
**Dauer:** ca. 120 min · **Material:** Roadmap-Board  
**Ergebnis:** Now/Next/Later-Roadmap, Gantt, MVP/Pilot-Steckbriefe  

Blatt `Roadmap`:

| Spalte | Typ | Pflicht | Werte/Hinweis |
|---|---|---|---|
| Use-Case-ID | str | ja |  |
| Horizont | choice |  | Now, Next, Later; Leer = automatischer Vorschlag aus Priorisierung |
| Start (JJJJ-Qn) | str |  | z.B. 2027-Q1; leer = automatisch |
| Dauer [Monate] | int |  |  |
| MVP-Umfang | text |  |  |
| Pilotbereich | text |  |  |
| Erfolgskriterien/KPIs | text |  |  |
| Abhängigkeiten (Use-Case-IDs) | ids |  |  |

### RACI-Matrix

**Ziel:** Klärt Verantwortlichkeiten: Responsible, Accountable (genau eine), Consulted, Informed.  
**Ablauf im Workshop:** Zeilen = Vorhaben/Aufgaben, Spalten = Rollen. Pro Zeile genau ein A, mindestens ein R.  
**Dauer:** ca. 45 min · **Material:** RACI-Matrix  
**Ergebnis:** Geprüfte RACI-Matrix  

Erfassung im Matrix-Format, Blatt `RACI`: erste Spalte *Aufgabe/Vorhaben*, weitere Spalten = Rollen (R/A/C/I).

### Business Case & ROI

**Ziel:** Macht aus Ideen budgetierbare Vorhaben: Investitionen, laufende Kosten, Einsparungen über die Jahre.  
**Ablauf im Workshop:** Mengengerüst mit Controlling: Zeiteinsparung × Stundensatz, vermiedene Fehlerkosten, Lizenz/Projekt-/Betriebskosten. Annahmen dokumentieren.  
**Dauer:** ca. 90 min · **Material:** Business-Case-Vorlage  
**Ergebnis:** NPV, ROI, Amortisationszeit je Use Case  

Blatt `BusinessCase`:

| Spalte | Typ | Pflicht | Werte/Hinweis |
|---|---|---|---|
| Use-Case-ID | str | ja |  |
| Position | str | ja |  |
| Typ | choice | ja | Investition, Laufende Kosten, Einsparung, Mehrumsatz |
| Jahr (0 = Invest-Jahr) | int | ja | 0–10 |
| Betrag [EUR] | float | ja | Immer positiv eintragen – das Vorzeichen ergibt sich aus dem Typ. |
| Annahme/Herleitung | text |  |  |
