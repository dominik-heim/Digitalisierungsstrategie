"""Beispielprojekt 'Muster Maschinenbau GmbH' – vollständig ausgefüllte Workshop-Erfassung zum Ausprobieren."""
from __future__ import annotations

from pathlib import Path

from .model import DEFAULT_CRITERIA, DEFAULT_RACI_ROLES
from .project import Project, append_rows, write_input_templates


def _d(keys: str, *rows):
    ks = keys.split("|")
    return [dict(zip(ks, r)) for r in rows]


STAKEHOLDER = _d("name|rolle|bereich|einfluss|interesse|haltung|erwartungen|massnahme",
    ("Dr. K. Berger", "CEO", "Geschäftsführung", 5, 4, "Promotor", "Wettbewerbsfähigkeit, kürzere Lieferzeiten", "Sponsor, Lenkungskreis"),
    ("S. Albrecht", "CFO", "Finanzen", 5, 3, "Skeptiker", "Belastbarer ROI, keine Budgetüberschreitung", "Business Case früh abstimmen"),
    ("M. Yilmaz", "Leiter Produktion", "Produktion", 4, 5, "Unterstützer", "Transparenz Auftragsstatus, weniger Suchzeiten", "Kernteam"),
    ("J. Hoffmann", "IT-Leiter", "IT", 4, 4, "Neutral", "Wartbarkeit, IT-Sicherheit, Kapazität", "Kernteam, Architektur-Board"),
    ("A. Schulz", "Leiterin Vertrieb", "Vertrieb", 3, 4, "Promotor", "Schnellere Angebote, Liefertermintreue", "Workshops"),
    ("T. Klein", "Betriebsrat", "Betriebsrat", 4, 2, "Skeptiker", "Leistungskontrolle, Arbeitsplätze", "Frühzeitig informieren, BV vorbereiten"),
    ("P. Wagner", "Meister Montage", "Produktion", 2, 5, "Unterstützer", "Weniger Papier, Zeichnungen digital", "Pilotnutzer"),
    ("L. Fischer", "Qualitätsleitung", "Qualität", 3, 3, "Neutral", "Rückverfolgbarkeit", "Workshops"),
    ("R. Becker", "Disponent", "Arbeitsvorbereitung", 2, 4, "Unterstützer", "Weniger Excel-Pflege", "Pilotnutzer"),
    ("E. Braun", "Einkauf", "Einkauf", 2, 2, "Neutral", "", "Informieren"),
)

SIPOC = _d("prozess|supplier|input|process|output|customer",
    ("Auftragsabwicklung", "Kunde", "Anfrage, Spezifikation", "Anfrage prüfen", "Angebot", "Kunde"),
    ("Auftragsabwicklung", "Konstruktion", "Stückliste, Zeichnungen", "Angebot kalkulieren", "Auftragsbestätigung", "Vertrieb"),
    ("Auftragsabwicklung", "Einkauf", "Lieferzeiten Zukaufteile", "Liefertermin klären", "Fertigungsauftrag", "Produktion"),
    ("Auftragsabwicklung", "", "", "Auftrag anlegen (ERP)", "", "Arbeitsvorbereitung"),
    ("Auftragsabwicklung", "", "", "Auftrag freigeben", "", ""),
)

PROZESS = _d("prozess|nr|schritt|lane|system|medium|typ|nachfolger|bemerkung",
    ("Auftragsabwicklung", 1, "Anfrage eingegangen", "Vertrieb", "Outlook", "E-Mail", "Start", "", ""),
    ("Auftragsabwicklung", 2, "Anfrage erfassen", "Vertrieb", "Excel", "Excel", "Aufgabe", "", "Anfrageliste in Excel"),
    ("Auftragsabwicklung", 3, "Technische Klärung", "Konstruktion", "CAD", "E-Mail", "Aufgabe", "", "Rückfragen per Mail"),
    ("Auftragsabwicklung", 4, "Machbar?", "Konstruktion", "", "Mündlich/Telefon", "Entscheidung", "5,11", ""),
    ("Auftragsabwicklung", 5, "Kalkulation erstellen", "Vertrieb", "Excel", "Excel", "Aufgabe", "", "Kalkulationsmappe"),
    ("Auftragsabwicklung", 6, "Liefertermin anfragen", "Arbeitsvorbereitung", "ERP", "Mündlich/Telefon", "Aufgabe", "", "Telefonische Abfrage"),
    ("Auftragsabwicklung", 7, "Angebot versenden", "Vertrieb", "Word", "E-Mail", "Aufgabe", "", ""),
    ("Auftragsabwicklung", 8, "Auftrag im ERP anlegen", "Vertrieb", "ERP", "Digital integriert", "Aufgabe", "", "Abtippen aus Excel"),
    ("Auftragsabwicklung", 9, "Fertigungspapiere drucken", "Arbeitsvorbereitung", "ERP", "Papier", "Aufgabe", "", "Laufkarte"),
    ("Auftragsabwicklung", 10, "Auftrag freigegeben", "Arbeitsvorbereitung", "ERP", "Digital integriert", "Ende", "", ""),
    ("Auftragsabwicklung", 11, "Absage an Kunden", "Vertrieb", "Outlook", "E-Mail", "Ende", "", ""),
)

VSM = _d("prozess|nr|schritt|bz|wz|rz|qr|system|medium|wertschoepfend|bestand",
    ("Fertigungsauftrag", 1, "Arbeitsvorbereitung", 45, 960, 0, 92, "ERP + Excel", "Excel", "Notwendig", 14),
    ("Fertigungsauftrag", 2, "Materialbereitstellung", 30, 1440, 0, 88, "Papier-Laufkarte", "Papier", "Nein", 22),
    ("Fertigungsauftrag", 3, "Zerspanung", 180, 720, 60, 95, "MES-Insel", "Digital integriert", "Ja", 9),
    ("Fertigungsauftrag", 4, "Montage", 480, 2880, 30, 90, "Papierzeichnung", "Papier", "Ja", 6),
    ("Fertigungsauftrag", 5, "Prüfung/QS", 60, 480, 0, 97, "Excel-Prüfprotokoll", "Excel", "Notwendig", 4),
    ("Fertigungsauftrag", 6, "Rückmeldung & Versand", 40, 600, 0, 99, "ERP", "Digital integriert", "Notwendig", 3),
)

PAINPOINTS = _d("id|prozess|schritt|beschreibung|kategorie|haeufigkeit|auswirkung|stimmen|quelle",
    ("PP-01", "Auftragsabwicklung", "Anfrage erfassen", "Anfragen werden doppelt in Excel und später ins ERP getippt", "Medienbruch", 5, 3, 7, "Vertrieb"),
    ("PP-02", "Auftragsabwicklung", "Liefertermin anfragen", "Liefertermine können nicht verbindlich zugesagt werden – keine Kapazitätssicht", "Transparenz", 4, 5, 9, "Vertrieb"),
    ("PP-03", "Auftragsabwicklung", "Kalkulation erstellen", "Kalkulation dauert 2-3 Tage, Wissen steckt in Köpfen", "Durchlaufzeit", 4, 4, 6, "Vertrieb"),
    ("PP-04", "Fertigungsauftrag", "Materialbereitstellung", "Suchzeiten nach Material und Laufkarten in der Halle", "Transparenz", 5, 4, 8, "Produktion"),
    ("PP-05", "Fertigungsauftrag", "Montage", "Montage arbeitet teils mit veralteten Papierzeichnungen", "Qualität", 3, 5, 8, "Montage"),
    ("PP-06", "Fertigungsauftrag", "Prüfung/QS", "Prüfprotokolle in Excel, keine Rückverfolgbarkeit je Seriennummer", "Datenqualität", 3, 4, 4, "Qualität"),
    ("PP-07", "Fertigungsauftrag", "Rückmeldung & Versand", "Rückmeldungen erfolgen gesammelt am Schichtende, Status im ERP veraltet", "Datenqualität", 5, 3, 5, "AV"),
    ("PP-08", "Auftragsabwicklung", "Technische Klärung", "Rückfragen per E-Mail gehen verloren, keine Historie", "Medienbruch", 3, 3, 3, "Konstruktion"),
    ("PP-09", "Fertigungsauftrag", "Zerspanung", "Maschinenstillstände werden nicht systematisch erfasst", "Transparenz", 4, 3, 4, "Instandhaltung"),
    ("PP-10", "Fertigungsauftrag", "Arbeitsvorbereitung", "Feinplanung in Excel, ständige manuelle Umplanung", "IT-System", 5, 4, 7, "AV"),
    ("PP-11", "Auftragsabwicklung", "Angebot versenden", "Angebote ohne einheitliche Vorlage, Fehler in Preisen", "Qualität", 2, 3, 1, "Vertrieb"),
    ("PP-12", "Fertigungsauftrag", "Montage", "Fehlende Kompetenzen im Umgang mit digitalen Tools", "Kompetenz", 3, 2, 2, "Betriebsrat"),
)

WHYS = _d("pain_point_id|problem|w1|w2|w3|w4|w5|grundursache|digital_loesbar",
    ("PP-02", "Liefertermine nicht verbindlich", "Kapazitätsauslastung unbekannt", "Feinplanung nur in Excel der AV",
     "ERP-Planung nicht gepflegt", "Rückmeldungen kommen zu spät", "Keine BDE an den Arbeitsplätzen",
     "Fehlende Echtzeit-Rückmeldung aus der Fertigung", "Ja"),
    ("PP-04", "Suchzeiten nach Material", "Lagerort nicht dokumentiert", "Bereitstellung über Zuruf",
     "Kein Lagerverwaltungssystem für Halle", "", "", "Keine Lagerplatzverwaltung in der Fertigung", "Ja"),
    ("PP-05", "Veraltete Zeichnungen in Montage", "Zeichnungen werden einmalig gedruckt", "Änderungen erreichen Montage nicht",
     "Kein Änderungsmanagement bis Shopfloor", "", "", "Kein digitaler Zugriff auf freigegebene Dokumente am Arbeitsplatz", "Ja"),
    ("PP-03", "Kalkulation dauert lange", "Erfahrungswerte fehlen", "Vorkalkulation nicht mit Nachkalkulation verknüpft",
     "", "", "", "Keine Rückkopplung Ist-Kosten in Kalkulation", "Teilweise"),
)

ISHIKAWA = _d("problem|kategorie|ursache|bestaetigt",
    ("Liefertermintreue 68 %", "Methode", "Terminierung ohne Kapazitätsabgleich", "Ja"),
    ("Liefertermintreue 68 %", "Daten/IT", "Rückmeldungen verspätet", "Ja"),
    ("Liefertermintreue 68 %", "Daten/IT", "Feinplanung in Excel-Insel", "Ja"),
    ("Liefertermintreue 68 %", "Material", "Zukaufteile verspätet", "Vermutet"),
    ("Liefertermintreue 68 %", "Mensch", "Wissen bei einzelnen Disponenten", "Ja"),
    ("Liefertermintreue 68 %", "Maschine", "Ungeplante Stillstände", "Vermutet"),
    ("Liefertermintreue 68 %", "Management", "Eilaufträge ohne Priorisierungsregel", "Ja"),
    ("Liefertermintreue 68 %", "Messung", "Keine KPI zur Termintreue je Bereich", "Ja"),
)

USERSTORIES = _d("id|rolle|ziel|nutzen|akzeptanz|aktivitaet|handlungsfeld|release|prioritaet|story_points|pain_point_ids",
    ("US-01", "Vertriebsmitarbeiter", "bei Anfrage sofort einen belastbaren Liefertermin sehen",
     "ich dem Kunden verbindlich zusagen kann",
     "Given ein Artikel mit Arbeitsplan When ich eine Menge eingebe Then zeigt das System den frühesten Termin mit Kapazitätsprüfung",
     "Angebot erstellen", "Transparente Auftragsplanung", "MVP", "Must", 8, "PP-02"),
    ("US-02", "Disponent", "Fertigungsaufträge per Drag&Drop umplanen", "Engpässe ohne Excel-Pflege auflösen",
     "Given ein Engpass When ich einen Auftrag verschiebe Then werden Folgetermine automatisch neu berechnet",
     "Fertigung planen", "Transparente Auftragsplanung", "MVP", "Must", 13, "PP-10"),
    ("US-03", "Werker", "Arbeitsgänge am Terminal an- und abmelden", "der Auftragsstatus in Echtzeit stimmt",
     "Given ein gestarteter Arbeitsgang When ich abmelde Then ist die Rückmeldung innerhalb 1 Minute im ERP",
     "Fertigen", "Shopfloor-Digitalisierung", "MVP", "Must", 5, "PP-07, PP-02"),
    ("US-04", "Monteur", "die aktuell freigegebene Zeichnung am Tablet sehen", "ich nicht nach veralteten Unterlagen montiere",
     "Given eine geänderte Zeichnung When ich den Auftrag öffne Then sehe ich nur den freigegebenen Stand",
     "Fertigen", "Shopfloor-Digitalisierung", "MVP", "Must", 5, "PP-05"),
    ("US-05", "Logistiker", "den Lagerplatz jedes Materials scannen und finden", "Suchzeiten entfallen",
     "Given ein Material When ich den Auftrag scanne Then wird der Lagerplatz angezeigt",
     "Material bereitstellen", "Shopfloor-Digitalisierung", "Release 2", "Should", 8, "PP-04"),
    ("US-06", "Qualitätsprüfer", "Prüfergebnisse digital je Seriennummer erfassen und auswerten", "Reklamationen nachvollziehbar sind",
     "Given eine Seriennummer When ich prüfe Then wird das Protokoll revisionssicher gespeichert",
     "Prüfen", "Qualität & Rückverfolgbarkeit", "Release 2", "Should", 8, "PP-06"),
    ("US-07", "Kalkulator", "Kalkulationen aus Vorlagen und Nachkalkulation ableiten", "Angebote in Stunden statt Tagen entstehen",
     "Given ein ähnlicher Altauftrag When ich ihn als Vorlage wähle Then werden Ist-Zeiten übernommen",
     "Angebot erstellen", "Digitaler Vertrieb", "Release 2", "Should", 13, "PP-03"),
    ("US-08", "Vertriebsmitarbeiter", "Anfragen direkt im CRM erfassen", "keine Doppelerfassung nötig ist",
     "Anfrage wird ins ERP übernommen", "Anfrage erfassen", "Digitaler Vertrieb", "MVP", "Should", 5, "PP-01"),
    ("US-09", "Instandhalter", "Stillstände automatisch aus der Maschine erhalten", "ich Ursachen auswerten kann",
     "Given ein Maschinenstillstand > 5 min When die Maschine stoppt Then wird ein Ereignis mit Zeitstempel angelegt",
     "Fertigen", "Shopfloor-Digitalisierung", "Release 3", "Could", 8, "PP-09"),
    ("US-10", "Konstrukteur", "Rückfragen zum Auftrag in einem Kanal dokumentieren", "die Historie für alle sichtbar ist",
     "Given eine Rückfrage When ich antworte Then ist sie am Vorgang gespeichert",
     "Anfrage erfassen", "Digitaler Vertrieb", "Release 3", "Could", 3, "PP-08"),
)

HANDLUNGSFELDER = _d("name|beschreibung|zielbild|verantwortlich",
    ("Transparente Auftragsplanung", "Kapazitätsbasierte Termin- und Feinplanung", "Verbindliche Termine in Echtzeit, Termintreue > 90 %", "M. Yilmaz"),
    ("Shopfloor-Digitalisierung", "BDE/MDE, digitale Dokumente und Material am Arbeitsplatz", "Papierlose Fertigung mit Echtzeitstatus", "M. Yilmaz"),
    ("Digitaler Vertrieb", "Durchgängiger Prozess Anfrage → Angebot → Auftrag", "Angebot in < 24 h, keine Doppelerfassung", "A. Schulz"),
    ("Qualität & Rückverfolgbarkeit", "Digitale Prüfprozesse je Seriennummer", "Lückenlose Rückverfolgbarkeit", "L. Fischer"),
)

USECASES = _d("id|titel|handlungsfeld|problem|loesung|nutzen|datenbasis|systeme|aufwand|owner|pain_point_ids|story_ids|impact|effort",
    ("UC-01", "Betriebsdatenerfassung (BDE) mit Terminals", "Shopfloor-Digitalisierung",
     "Rückmeldungen erst am Schichtende, Status veraltet", "Terminals je Arbeitsplatz, Rückmeldung in Echtzeit ins ERP",
     "Aktueller Auftragsstatus, Basis für Feinplanung; ca. 1.200 h/Jahr weniger Statusabfragen",
     "Arbeitspläne, Rückmeldungen", "ERP, MES-Modul", "Hardware, Schnittstelle ERP, Schulung, BV", "M. Yilmaz",
     "PP-07, PP-02", "US-03", "", ""),
    ("UC-02", "APS-Feinplanung mit Kapazitätsabgleich", "Transparente Auftragsplanung",
     "Feinplanung in Excel, Termine unzuverlässig", "APS-Tool mit Engpassplanung, gekoppelt an ERP/BDE",
     "Termintreue +20 %-Punkte, weniger Eilaufträge", "Arbeitspläne, Kapazitäten, BDE-Daten", "APS, ERP",
     "Stammdatenbereinigung, Integration", "R. Becker", "PP-02, PP-10", "US-01, US-02", "", ""),
    ("UC-03", "Digitale Zeichnung & Arbeitsanweisung am Tablet", "Shopfloor-Digitalisierung",
     "Veraltete Papierzeichnungen in der Montage", "PDM-Viewer auf Tablets, Freigabestatus sichtbar",
     "Weniger Nacharbeit, kein Druck", "PDM-Dokumente", "PDM, Tablets", "WLAN Halle, Tablets", "P. Wagner",
     "PP-05", "US-04", "", ""),
    ("UC-04", "Lagerplatzverwaltung Halle mit Scan", "Shopfloor-Digitalisierung",
     "Suchzeiten nach Material", "Barcode-Scan, Lagerplätze im ERP-LVS", "-50 % Suchzeiten",
     "Materialstamm, Lagerplätze", "ERP-LVS, Scanner", "Lagerplatzaufnahme", "E. Braun", "PP-04", "US-05", "", ""),
    ("UC-05", "Digitales Prüfprotokoll je Seriennummer", "Qualität & Rückverfolgbarkeit",
     "Excel-Prüfprotokolle", "CAQ-Prüfplan mit Tablet-Erfassung", "Rückverfolgbarkeit, Auswertung von Fehlern",
     "Prüfpläne", "CAQ", "CAQ-Lizenz, Prüfplanerstellung", "L. Fischer", "PP-06", "US-06", "", ""),
    ("UC-06", "Kalkulationsassistent mit Ähnlichkeitssuche (KI)", "Digitaler Vertrieb",
     "Kalkulation dauert Tage", "Ähnliche Altaufträge per KI finden, Ist-Zeiten übernehmen",
     "Angebot in < 24 h", "Auftragshistorie, Nachkalkulation", "ERP, KI-Service",
     "Datenqualität Historie unsicher", "A. Schulz", "PP-03", "US-07", "", ""),
    ("UC-07", "CRM-Anfragemanagement mit ERP-Übergabe", "Digitaler Vertrieb",
     "Doppelerfassung Excel/ERP", "CRM-Modul mit Schnittstelle", "Weniger Fehler, schnellere Angebote",
     "Kundenstamm", "CRM, ERP", "Lizenz, Schnittstelle", "A. Schulz", "PP-01, PP-08", "US-08, US-10", "", ""),
    ("UC-08", "Maschinendatenerfassung (MDE) & OEE-Dashboard", "Shopfloor-Digitalisierung",
     "Stillstände unbekannt", "IoT-Gateway an Maschinen, OEE-Dashboard", "OEE-Transparenz, gezielte Instandhaltung",
     "Maschinensignale", "IoT-Plattform", "Heterogener Maschinenpark, Retrofit", "J. Hoffmann", "PP-09", "US-09", "", ""),
)

VPC = _d("use_case_id|segment|bereich|eintrag|bezug",
    ("UC-06", "Vertriebsinnendienst", "Customer Job", "Angebote schnell und korrekt kalkulieren", ""),
    ("UC-06", "Vertriebsinnendienst", "Pain", "Kalkulation dauert 2-3 Tage", ""),
    ("UC-06", "Vertriebsinnendienst", "Pain", "Wissen nur bei erfahrenen Kollegen", ""),
    ("UC-06", "Vertriebsinnendienst", "Gain", "Höhere Angebots-Trefferquote", ""),
    ("UC-06", "Vertriebsinnendienst", "Produkt/Service", "KI-Kalkulationsassistent im ERP", ""),
    ("UC-06", "Vertriebsinnendienst", "Pain Reliever", "Vorschlag aus ähnlichen Altaufträgen in Sekunden", "Kalkulation dauert"),
    ("UC-06", "Vertriebsinnendienst", "Pain Reliever", "Erfahrungswissen in Daten gesichert", "Wissen nur bei"),
    ("UC-06", "Vertriebsinnendienst", "Gain Creator", "Preis-Benchmark aus Nachkalkulation", ""),
)

BMC = _d("use_case_id|baustein|eintrag",
    ("UC-08", "Kundensegmente", "Bestandskunden mit installierten Sondermaschinen"),
    ("UC-08", "Wertangebot", "Verfügbarkeits-Monitoring & Predictive Service als Abo"),
    ("UC-08", "Kanäle", "Kundenportal, Servicetechniker"),
    ("UC-08", "Kundenbeziehungen", "Servicevertrag, persönlicher Ansprechpartner"),
    ("UC-08", "Einnahmequellen", "Monatliche Lizenz je Maschine; Service-Einsätze"),
    ("UC-08", "Schlüsselressourcen", "IoT-Plattform, Maschinendaten, Service-Know-how"),
    ("UC-08", "Schlüsselaktivitäten", "Datenanalyse, Plattformbetrieb, Retrofit"),
    ("UC-08", "Schlüsselpartner", "IoT-Plattformanbieter, Cloud-Provider"),
    ("UC-08", "Kostenstruktur", "Plattformlizenz, Entwicklung, Support"),
)

FITGAP = _d("id|anforderung|prozess|prioritaet|system|abdeckung|loesung|aufwand_pt",
    ("FG-01", "Rückmeldung Arbeitsgang mit Mengen und Zeiten", "Fertigung", "Muss", "ERP-MES", "Standard", "", 2),
    ("FG-02", "Ausschussgründe nach eigenem Katalog", "Fertigung", "Muss", "ERP-MES", "Konfiguration", "Katalog pflegen", 3),
    ("FG-03", "Kapazitätsterminierung mit Engpassplanung", "Planung", "Muss", "APS", "Standard", "", 5),
    ("FG-04", "Plantafel Drag&Drop", "Planung", "Soll", "APS", "Standard", "", 1),
    ("FG-05", "Terminzusage im Angebot (ATP/CTP)", "Vertrieb", "Muss", "ERP", "Gap", "Erweiterung CTP-Service oder APS-API", 25),
    ("FG-06", "Anzeige PDM-Zeichnung zum Arbeitsgang", "Fertigung", "Muss", "MES", "Erweiterung", "Viewer-Integration", 8),
    ("FG-07", "Personalzeiterfassung am selben Terminal", "Fertigung", "Soll", "MES", "Konfiguration", "", 4),
    ("FG-08", "Prüfmerkmale je Seriennummer", "Qualität", "Muss", "CAQ", "Standard", "", 3),
    ("FG-09", "Eigenes Layout Laufkarte", "Fertigung", "Kann", "ERP", "Konfiguration", "Formular", 2),
    ("FG-10", "Rüstoptimierte Reihenfolge", "Planung", "Kann", "APS", "Gap", "Nicht umsetzen", 15),
)

NWA = {  # Use Case -> Punkte je Standardkriterium (Reihenfolge wie DEFAULT_CRITERIA)
    "UC-01": [4, 5, 4, 4, 4, 4], "UC-02": [5, 5, 5, 2, 3, 2], "UC-03": [3, 4, 3, 5, 4, 5],
    "UC-04": [3, 4, 2, 4, 3, 4], "UC-05": [3, 3, 3, 3, 4, 4], "UC-06": [4, 3, 4, 2, 2, 2],
    "UC-07": [2, 3, 3, 3, 4, 3], "UC-08": [3, 2, 4, 1, 2, 2],
}

ROADMAP = _d("use_case_id|horizont|start|dauer|mvp|pilot|kpi|abhaengigkeiten",
    ("UC-01", "Now", "", 4, "BDE an 10 Arbeitsplätzen Zerspanung/Montage", "Halle 2, Linie Montage A",
     "Rückmeldequote > 95 % in Echtzeit", ""),
    ("UC-03", "Now", "", 3, "Tablets für 2 Montageteams", "Montage A", "Nacharbeit wegen Zeichnungen −50 %", ""),
    ("UC-02", "", "", 9, "Engpassplanung für Zerspanung", "Zerspanung", "Termintreue > 85 %", "UC-01"),
)

RACI = [
    {"_": "Lenkungskreis & Budget", "Geschäftsführung": "A", "Projektleitung": "R", "IT": "C", "Fachbereich": "C", "Produktion": "I", "Qualität": "I"},
    {"_": "UC-01 BDE einführen", "Geschäftsführung": "I", "Projektleitung": "A", "IT": "R", "Fachbereich": "C", "Produktion": "R", "Qualität": "I"},
    {"_": "UC-02 APS-Feinplanung", "Geschäftsführung": "I", "Projektleitung": "A", "IT": "R", "Fachbereich": "R", "Produktion": "C", "Qualität": ""},
    {"_": "UC-03 Digitale Zeichnung", "Geschäftsführung": "", "Projektleitung": "C", "IT": "R", "Fachbereich": "", "Produktion": "A", "Qualität": "C"},
    {"_": "Betriebsvereinbarung BDE", "Geschäftsführung": "A", "Projektleitung": "R", "IT": "C", "Fachbereich": "", "Produktion": "C", "Qualität": ""},
    {"_": "Stammdatenbereinigung", "Geschäftsführung": "", "Projektleitung": "A", "IT": "C", "Fachbereich": "R", "Produktion": "R", "Qualität": "A"},
]

BUSINESSCASE = _d("use_case_id|position|typ|jahr|betrag|annahme",
    ("UC-01", "Terminals & Lizenzen", "Investition", 0, 85000, "20 Terminals à 2.500 € + Lizenz"),
    ("UC-01", "Einführung/Schnittstelle", "Investition", 0, 60000, "40 PT extern"),
    ("UC-01", "Wartung", "Laufende Kosten", 1, 12000, ""), ("UC-01", "Wartung", "Laufende Kosten", 2, 12000, ""),
    ("UC-01", "Wartung", "Laufende Kosten", 3, 12000, ""), ("UC-01", "Wartung", "Laufende Kosten", 4, 12000, ""),
    ("UC-01", "Wartung", "Laufende Kosten", 5, 12000, ""),
    ("UC-01", "Weniger Statusabfragen/Suchzeiten", "Einsparung", 1, 70000, "1.200 h × 58 €/h"),
    ("UC-01", "Weniger Statusabfragen/Suchzeiten", "Einsparung", 2, 85000, ""),
    ("UC-01", "Weniger Statusabfragen/Suchzeiten", "Einsparung", 3, 85000, ""),
    ("UC-01", "Weniger Statusabfragen/Suchzeiten", "Einsparung", 4, 85000, ""),
    ("UC-01", "Weniger Statusabfragen/Suchzeiten", "Einsparung", 5, 85000, ""),
    ("UC-02", "APS-Lizenz & Einführung", "Investition", 0, 180000, ""),
    ("UC-02", "Betrieb", "Laufende Kosten", 1, 30000, ""), ("UC-02", "Betrieb", "Laufende Kosten", 2, 30000, ""),
    ("UC-02", "Betrieb", "Laufende Kosten", 3, 30000, ""), ("UC-02", "Betrieb", "Laufende Kosten", 4, 30000, ""),
    ("UC-02", "Betrieb", "Laufende Kosten", 5, 30000, ""),
    ("UC-02", "Weniger Eilaufträge/Sonderfahrten", "Einsparung", 1, 40000, ""),
    ("UC-02", "Weniger Eilaufträge/Sonderfahrten", "Einsparung", 2, 90000, ""),
    ("UC-02", "Mehrumsatz durch Termintreue", "Mehrumsatz", 2, 50000, "Deckungsbeitrag"),
    ("UC-02", "Weniger Eilaufträge/Sonderfahrten", "Einsparung", 3, 110000, ""),
    ("UC-02", "Mehrumsatz durch Termintreue", "Mehrumsatz", 3, 80000, ""),
    ("UC-02", "Weniger Eilaufträge/Sonderfahrten", "Einsparung", 4, 110000, ""),
    ("UC-02", "Mehrumsatz durch Termintreue", "Mehrumsatz", 4, 80000, ""),
    ("UC-02", "Weniger Eilaufträge/Sonderfahrten", "Einsparung", 5, 110000, ""),
    ("UC-03", "Tablets & WLAN", "Investition", 0, 45000, ""),
    ("UC-03", "Support", "Laufende Kosten", 1, 6000, ""), ("UC-03", "Support", "Laufende Kosten", 2, 6000, ""),
    ("UC-03", "Support", "Laufende Kosten", 3, 6000, ""),
    ("UC-03", "Weniger Nacharbeit & Druck", "Einsparung", 1, 30000, ""),
    ("UC-03", "Weniger Nacharbeit & Druck", "Einsparung", 2, 35000, ""),
    ("UC-03", "Weniger Nacharbeit & Druck", "Einsparung", 3, 35000, ""),
)


def create_demo(root: Path) -> Project:
    p = Project.create(root, kunde="Muster Maschinenbau GmbH", branche="Sondermaschinenbau, 600 MA",
                       berater="Digital Solutions", roadmap_start="2027-Q1")
    f = p.input_dir / "Workshop-Erfassung.xlsx"
    write_input_templates(f, use_case_ids=list(NWA))
    for key, rows in [("stakeholder", STAKEHOLDER), ("sipoc", SIPOC), ("prozess", PROZESS), ("vsm", VSM),
                      ("painpoints", PAINPOINTS), ("whys", WHYS), ("ishikawa", ISHIKAWA),
                      ("userstories", USERSTORIES), ("handlungsfelder", HANDLUNGSFELDER), ("usecases", USECASES),
                      ("vpc", VPC), ("bmc", BMC), ("fitgap", FITGAP), ("roadmap", ROADMAP),
                      ("businesscase", BUSINESSCASE)]:
        append_rows(f, key, rows)
    append_rows(f, "kriterien", [{"kriterium": c[0], "kategorie": c[1], "gewicht": c[2], "skala": c[3]}
                                 for c in DEFAULT_CRITERIA])
    from openpyxl import load_workbook
    wb = load_workbook(f)
    ws = wb["NWA_Bewertung"]
    for i, (uc, vals) in enumerate(NWA.items(), start=2):
        ws.cell(i, 1, uc)
        for j, v in enumerate(vals, start=2):
            ws.cell(i, j, v)
    ws = wb["RACI"]
    for i, row in enumerate(RACI, start=2):
        ws.cell(i, 1, row["_"])
        for j, role in enumerate(DEFAULT_RACI_ROLES, start=2):
            ws.cell(i, j, row.get(role, ""))
    wb.save(f)
    return p
