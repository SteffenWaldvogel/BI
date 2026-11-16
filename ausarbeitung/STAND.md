# Projektstand BI-Portfolio, Thema 1.2 (Stand 16.11.2026)

Übergabedokument für einen neuen Chat. Maßgeblich für den aktuellen Stand; die Kennzahlen in `README.md`
(200 Kunden) sind veraltet.

## 1. Aufgabe und Rahmen

- Portfolioleistung Business Intelligence, Prof. Kirchberg, DHBW Ravensburg, Thema 1.2:
  drei Quellsysteme (ERP, Online-Shop, CRM) eines Dart-Zubehör-Händlers entwerfen und mit Excel Power Query
  in einen gemeinsamen, integrierten Datenbestand für verkäuferbezogene Auswertungen überführen.
  Alle im Skript genannten Integrationsprobleme müssen vorkommen und gelöst werden; der Ablauf ist mit
  Beispieldaten zu verdeutlichen; die Erzeugung der Demodaten ist zu zeigen.
- Abgabe: **27.11.2026, 24 Uhr**, PDF per Mail von der DHBW-Adresse an den Dozenten.
- Umfang: 8 bis 12 Inhaltsseiten als Richtlinie, bei vielen Abbildungen mehr.
  Elemente: Einleitung, Themenbearbeitung, Zusammenfassung. Keine Grundlagen, keine Reflexion, kein Ausblick.
- Bewertung wie ein Programmentwurf: Weg, begründete Entscheidungen, Alternativen, Eigenleistung.
  "Aufgabe erfüllt" entspricht 20 von 30 Punkten. Literaturarbeit zählt nicht.
- Präsentation am letzten Termin: maximal zwei Folien, fünf Minuten, Folien auf dem T-Laufwerk.
- Fallbeispiel: fiktive **DartPoint GmbH** (DACH), sechs Fragen F1 bis F6 (Umsatz je Verkäufer nach Kundentyp,
  Land/Region, Saison, Produktgruppe, Kanalanteil betreuter Kunden, Liefer- und Zahlungsqualität).

## 2. Vorgaben des Autors für den Text (verbindlich)

- Gliederung nur in drei Teilen: Einleitung, Themenbearbeitung, Zusammenfassung.
- Wissenschaftlicher Stil: sachlich, unpersönlich, nicht erzählend.
- Umfang je Abschnitt: nur was das Thema verlangt (Problem, Lösung in Power Query, Begründung, verworfene
  Alternative). Ein ausführliches Zielschema-Kapitel war zu viel und wurde auf einen Absatz gekürzt.
- Die frühere eigene Projektarbeit und das ServiceTool werden **nirgends** in der Ausarbeitung erwähnt.
  Sie waren nur Hilfestellung bei der Erarbeitung.
- Formulierung: "muss manuell überprüft und entschieden werden", nicht "ein Mensch prüft".
- Keine Quellenverweise, kein Literaturverzeichnis, kein KI-Verzeichnis.
- Alles in einer `main.tex`. Schwerpunkt sind die Abbildungen mit Beschreibung und Entscheidungen.
- Vorher/Nachher-Abbildungen als kleine Ausschnitte zu allen Fehlerbildern.
- Maximal rund 500 Kundendatensätze.
- Eine automatische Zusammenführung umgezogener Kunden (Brückenregel) wird nicht eingebaut, aber im Text
  bewertet: möglich, jedoch mit Risiko von Datenverlust; stattdessen Prüfliste und manuelle Entscheidung.

## 3. Ordner

| Pfad | Inhalt |
|---|---|
| `generate_sources.py` | Generator der Beispieldaten, Seed 20261127, Parameter `FEHLERQUOTE`. Kopie in `ausarbeitung/latex/code/`. |
| `quellen/erp`, `quellen/shop`, `quellen/crm` | erzeugte Quelldateien (ERP .txt Windows-1252, Shop .csv UTF-8, CRM .xlsx) |
| `sollwerte/` | bekannte Wahrheit: `kunden_mapping.csv` (Zuordnung) und `fehler_log.csv` (318 eingebrachte Fehler) |
| `integration/DartPoint_Integration.xlsx` | die Mappe: 60 Abfragen, 18 Blätter, Datenmodell, Dashboard |
| `integration/Messung6.xlsx`, `Messung7.xlsx` | Messkopien mit umgestellten Parametern (Beleg für die Messreihe) |
| `ausarbeitung/latex/main.tex` | die Ausarbeitung (pdfLaTeX, für Overleaf) |
| `ausarbeitung/latex/bilder/` | 19 Abbildungen |
| `ausarbeitung/latex/code/` | `Section1.m` (M-Code der Mappe, Quelle der Listings), `BELEG_*.m`, Generator, Listing-Stil |
| `ausarbeitung/overleaf_upload.zip` | `latex/` gezippt; nach jeder Änderung an `latex/` neu bauen |
| `ausarbeitung/fehlerkatalog.md`, `messprotokoll.md` | interne Arbeitsunterlagen (nennen die frühere Projektarbeit; nicht Teil der Abgabe) |
| `ausarbeitung/OFFEN.md` | kurze Liste offener Punkte (in Abschnitt 7 übernommen) |

## 4. Datenstand (v3)

- 285 eindeutige Kunden, daraus 499 Kundensätze: ERP 134 (4 Dubletten, 2 Testkunden), Shop 201 (1 Test),
  CRM 164 (7 Dubletten, 1 Test). 478 ERP-Aufträge, 549 Shop-Bestellungen.
- Präsenz: ERP+Shop+CRM 64, ERP+CRM 28, Shop+CRM 43, nur CRM 21, nur Shop 93, nur ERP 36.
- 318 protokollierte Fehler: Formate, Kodierungen, Feldmissbrauch, Dubletten, Sammel-E-Mail (5 Paare),
  gleiche Adresse bei verschiedenen Kunden (8 Vereinsheime), Fehler in den Bewegungsdaten
  (4 Positionen ohne Auftrag, 3 Doppelimporte, 4 Lieferungen vor Auftragsdatum, 5 Gutschriften).

## 5. Mappe

- Schichten: Rohimport (`KUNDE`, `customers`, `Accounts`, ...), Hilfsfunktionen (`fnClean`, `fnNormName`,
  `fnNameOk`, `fnLand`, `fnTyp`, `fnJN`, `fnStrasse`, `fnMatchkey`, `fnTel`, `fnEmails`), Mapping-Tabellen
  (`mapLand`, `mapKundentyp`, `mapRegion`, `mapMwSt`), Parameter (`pPfad`, `pSchwelle`, `pKursCHF`,
  `pAbsicherung`, `pBelegtRegel`), Harmonisierung (`KUNDE_h`, `SHOP_h`, `CRM_h`), Kaskade (`KUNDE_k`, `CRM_k`,
  `SHOP_k`, `MAP_SCHLUESSEL`), Ziel (`DIM_KUNDE`, `DIM_KUNDE_ADRESSE`, `DIM_VERKAEUFER`, `DIM_PRODUKT`,
  `FAKT_ERP`, `FAKT_SHOP`, `FAKT_VERKAUF`), Prüfung (`CHECK_Sollwerte`, `CHECK_Detail`, `CHECK_Fakten`,
  `CHECK_Pruefliste`, `CHECK_Fehlerlog`), Belege (`BELEG_Heterogenitaet`, `BELEG_Fehlerbilder`).
- Kaskade je Quelle: 1 ERP-Schlüssel mit Bestätigung durch Name oder Adresse, 2 E-Mail (bei Mehrfachverwendung
  mit Namensbestätigung), 3 Adressschlüssel mit Namensbestätigung, 4 Fuzzy mit fachlicher Nachregel und
  Belegt-Regel, sonst neuer Schlüssel.
- Messreihe (`CHECK_Sollwerte`): Messung 6 (jedes Merkmal allein) 17 zerrissen / 9 verschmolzen;
  Messung 7 (abgesichert) 7 / 1; Messung 8 (mit Belegt-Regel, aktueller Stand) **7 / 0**, 0 Quell-IDs ohne Mapping.
- Prüfliste: 161 Befunde an 139 von 499 Kundensätzen. `CHECK_Fakten`: alle Soll-Ist-Vergleiche erfüllt
  (FAKT_VERKAUF 2529 Zeilen; 131 Shop-Bestellungen auch im ERP, 56 storniert oder erstattet).
- Blätter: sechs Pivot-Blätter P1 bis P6, `Dashboard` (sechs PivotCharts), Prüf- und Zielblätter,
  `BELEG_Heterogenitaet`, `BELEG_Fehlerbilder`.
- **`pPfad` zeigt derzeit auf den anderen Rechner** (`C:\Users\SteffenWaldvogel\Documents\DHBW\BI-Seminararbeit\`).
  Vor einer Aktualisierung auf diesem Rechner auf den hiesigen Projektordner umstellen (mit abschließendem `\`).
- Laufzeiten: `CHECK_Sollwerte` rund 8 Minuten, "Alle aktualisieren" rund 20 Minuten. Excel zeigt währenddessen
  "Keine Rückmeldung". Die Mappe nicht über Standby offen lassen (zweimal aufgehängt, ohne Datenverlust).
- Die M-Kommentare in der Mappe nennen noch die frühere Projektarbeit ("PA1"). In `code/Section1.m` sind sie
  bereinigt; in die PDF gelangen sie nicht.
- Im Datenmodell liegt noch ein Duplikat `FAKT_VERKAUF 1`.
- Sechs der sieben Pivot-Tabellen teilen sich eine Datenquelle, eine hängt am Datenmodell (wichtig für Datenschnitte).

## 6. Ausarbeitung

Gliederung in `main.tex`:

1. Einleitung
2. Themenbearbeitung
   - 2.1 Quellsysteme und Beispieldaten
   - 2.2 Integrierter Datenbestand
   - 2.3 Import und Harmonisierung
   - 2.4 Kundenintegration und Erfolgsmessung
   - 2.5 Prüfliste: automatisch, hybrid, manuell
   - 2.6 Fakten und Auswertung
3. Zusammenfassung

Anhang A (M-Code-Ausschnitte), B (Generator-Ausschnitte), C (Vorher/Nachher-Belege), ehrenwörtliche Erklärung.

- **Der Text ist vollständig geschrieben** (alle Abschnitte, Einleitung und Zusammenfassung). Einzige verbliebene
  `\todo`-Stelle: Dozent auf dem Deckblatt.
- Abbildungen: `vn-kunde-6`, `vn-kunde-34` (ein Kunde in drei Quellen), `vn-1-format` bis `vn-7-bewegung`
  (Vorher/Nachher je Fehlergruppe), `pq-import-erp`, `pq-mapping`, `pq-strasse`, `pq-kaskade`, `pq-check`,
  `pq-pruefliste`, `pq-beleg`, `check-fakten`, `dashboard`. Dazu drei TikZ-Schaubilder im Text.
- **Seit dem Einbau der neuen Abbildungen, der neuen Gliederung und der Texte wurde `main.tex` nicht kompiliert.**
  Erster Schritt im neuen Chat: ZIP in Overleaf hochladen, Fehler beheben, Seitenzahl feststellen.
  Wahrscheinliche Stellen: Querverweise auf zusammengelegte Abschnitte (mehrere `\label` an einer Überschrift),
  Sonderzeichen in Bildunterschriften.
- Im Text offen angesprochene Grenzen: Sommerzeit bei der UTC-Umrechnung unberücksichtigt; Straßennamen mit Zahl
  würden falsch getrennt; Ersetzungsliste nur für deutsche Umlaute; Stammattribute nach fester Reihenfolge
  ERP, CRM, Shop (erster vorhandener Wert).

## 7. Offene Punkte

1. In Overleaf kompilieren, Fehler beheben, Seitenzahl prüfen; danach entscheiden, welche Abbildungen in den
   Anhang wandern (2.3 trägt acht Abbildungen und eine Tabelle).
2. Dozent auf dem Deckblatt eintragen (Titel und Vorname prüfen).
3. Analyse-Blatt in der Mappe (vom Autor gewünscht, nach dem Text): Pivot-Tabelle mit Hierarchien
   (Land > Region > Verkäufer; Saisonjahr > Quartal > Monat) für Drill-down und Roll-up, dazu Datenschnitte
   (Produktgruppe, Kundentyp, Kanal). Das Dashboard bleibt unverändert, weil seine Titel feste Aussagen nennen.
   Danach ein Screenshot und zwei bis drei Sätze in 2.6.
4. Schwache Vorher/Nachher-Beispiele ersetzen (`DE -> DE` bei den Ländervarianten, `ja -> ja` bei J/N),
   dafür `BELEG_Fehlerbilder` anpassen und `vn-2-kodierung` neu aufnehmen.
5. Optional: zweite Erfolgsmessung (`CHECK_Fehlerlog`) mit Zahlen in 2.5; je Frage ein Satz Ergebnis in 2.6.
6. `FAKT_VERKAUF 1` aus dem Datenmodell entfernen; M-Kommentare in der Mappe bereinigen.
7. PDF erzeugen und abgeben; zwei Folien für die Präsentation.

## 8. Arbeitsweise und Technik

- Zwei Rechner, per Git synchron (`origin`: GitHub-Repository `SteffenWaldvogel/BI`). Stand 16.11.: beide Seiten
  gleich, nichts ungepusht. Commits werden lokal erstellt, gepusht wird vom Autor.
- Nach Änderungen an `ausarbeitung/latex/`: ZIP neu bauen
  (`cd ausarbeitung/latex && zip -r ../overleaf_upload.zip . -x '*.aux' '*.log' '*.out' '*.toc' '*.pdf'`).
- Listings im Anhang lesen feste Zeilenbereiche aus `code/Section1.m` und `code/generate_sources.py`.
  Änderungen an diesen Dateien dürfen die Zeilenzahl nicht verschieben, sonst die Bereiche in `main.tex` anpassen.
- Abfragen in die Mappe bringen, ohne sie zu schließen: M-Text als UTF-16-Datei ablegen und im VBA-Direktfenster
  `Workbooks("DartPoint_Integration.xlsx").Queries.Add "Name", CreateObject("Scripting.FileSystemObject").OpenTextFile(Pfad, 1, False, -1).ReadAll`
  ausführen; Laden in ein Blatt über `ListObjects.Add` mit der Mashup-Verbindung. Vorhandene Abfrage ändern:
  `.Queries("Name").Formula = ...`, danach `QueryTable.Refresh False`.
- Screenshots der Vorher/Nachher-Tabellen: Spalten Nr, Gruppe, ID, Regel ausblenden, nach Gruppe filtern,
  Zoom 150 %, Menüband ausblenden, danach alles zurückstellen.
- Im Ordner des ServiceTools liegen Dateien mit Zugangsdaten (`key.txt`, `mail.txt`). Nicht öffnen.
