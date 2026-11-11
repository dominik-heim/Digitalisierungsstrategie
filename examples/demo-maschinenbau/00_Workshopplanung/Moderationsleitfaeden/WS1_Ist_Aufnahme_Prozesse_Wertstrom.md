# Moderationsleitfaden WS1: Ist-Aufnahme Prozesse & Wertstrom

**Kunde:** Muster Maschinenbau GmbH  
**Termin:** 2026-11-23, 09:00–14:30  
**Ort:** vor Ort beim Kunden  
**Ziel:** Kernprozesse abgrenzen und im Detail aufnehmen; Zeiten, Medien und Systeme sichtbar machen

## Teilnehmer

- Dr. K. Berger – CEO
- S. Albrecht – CFO ⚠ vorab Einzelgespräch (kritische Haltung)
- M. Yilmaz – Leiter Produktion
- J. Hoffmann – IT-Leiter
- A. Schulz – Leiterin Vertrieb
- P. Wagner – Meister Montage
- L. Fischer – Qualitätsleitung
- R. Becker – Disponent

## Ablauf

### 09:00–09:15 · Begrüßung & Ziele (15 min)

### 09:15–10:00 · Prozessrahmen SIPOC (45 min)

*Methode:* Prozessrahmen (SIPOC) – Liefert den groben Rahmen eines Prozesses: Lieferanten, Inputs, Prozessschritte, Outputs, Kunden.

*Ablauf:* Mit 5 Spalten an der Wand starten: zuerst Process (5-7 Hauptschritte), dann Output & Customer, danach Input & Supplier. Start- und Endpunkt explizit festlegen.

*Material:* 5 Spalten-Poster, Stickies  
*Ergebnis:* Abgegrenzter Prozessrahmen je Kernprozess  
*Erfassung:* Blatt `SIPOC` in der Erfassungsvorlage

### 10:00–12:00 · Ist-Prozess Swimlane/BPMN (120 min)

*Methode:* Ist-Prozessaufnahme (BPMN/Swimlane) – Detailliertes Ist-Modell mit Rollen (Swimlanes), Systemen und Medien. Basis für BPMN-Export und Medienbruch-Analyse.

*Ablauf:* Brown-Paper oder Miro: Swimlanes je Rolle anlegen, Schritte chronologisch kleben, pro Schritt System und Medium (Papier/Excel/E-Mail…) notieren. Entscheidungen als Raute.

*Material:* Brown-Paper, Stickies in 3 Farben, Rautenkarten  
*Ergebnis:* Swimlane-Diagramm, BPMN-2.0-Datei (Import in Signavio/Camunda/bpmn.io), Medienbruch-Liste  
*Erfassung:* Blatt `Prozessaufnahme` in der Erfassungsvorlage

### 12:00–12:45 · Mittagspause (45 min)

### 12:45–14:15 · Wertstrom: Zeiten, Bestände, Qualität (90 min)

*Methode:* Value Stream Mapping – Macht Durchlaufzeiten, Wartezeiten, Qualitätsverluste und Medienbrüche sichtbar.

*Ablauf:* Gemba/Prozessbegehung, anschließend Wertstrom am Board: je Prozesskasten Bearbeitungszeit, Wartezeit, FPY und Informationsfluss erfassen. Zeiten mit Stoppuhr/Systemdaten plausibilisieren.

*Material:* VSM-Symbole, Datenkästen, Stoppuhr, ggf. System-Auswertungen  
*Ergebnis:* Durchlaufzeit, Flussgrad, Rolled Throughput Yield, Zeitleiste, Medienbrüche  
*Erfassung:* Blatt `Wertstrom` in der Erfassungsvorlage

Hinweis: Vorab Gemba-Walk empfohlen

### 14:15–14:30 · Zusammenfassung & nächste Schritte (15 min)

## Nach dem Workshop

1. Ergebnisse (Fotos, Miro-Export, Notizen) in die Erfassungsvorlage übertragen – oder mit `digistrat ki-import` automatisch extrahieren lassen.
2. Datei nach `01_Workshop-Input/` legen.
3. `digistrat auswerten` ausführen → PowerPoint & Excel.