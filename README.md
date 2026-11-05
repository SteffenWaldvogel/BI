# BI-Portfolio – Thema 1.2: Power-Query-Integration DartPoint GmbH

Abgabe: 27.11.2026, 24 Uhr, PDF an kirchberg@dhbw-ravensburg.de (von der DHBW-Adresse).
Entwurf der Quellsysteme und Checkliste: Claude-Doc "Quellsysteme Dart-Händler".

## Ordner
- `generate_sources.py` – erzeugt alle Quelldateien deterministisch (Seed 20261127). Gehört in den Anhang.
- `quellen/erp/`  – ERP-Export: .txt, Windows-1252, Semikolon, TT.MM.JJJJ, Dezimalkomma
- `quellen/shop/` – Shop-Export: .csv, UTF-8, Komma, ISO-Datum UTC, Dezimalpunkt, Brutto
- `quellen/crm/`  – CRM-Export: CRM_Export.xlsx (Accounts, Betreuer, Aktivitaet), US-Datum, GUIDs
- `sollwerte/`    – Ground Truth (kunden_mapping.csv, parameter.json). NUR zur Erfolgsmessung, nicht Teil der Quellen.
- `integration/`  – hier entsteht die Excel mit den Power-Query-Abfragen (leer)
- `ausarbeitung/` – Text, Screenshots, Folien (leer)

## Neu generieren
    python generate_sources.py      (benötigt: pip install openpyxl)

## Fehlerdichte
`FEHLERQUOTE` im Skript skaliert alle manuell entstandenen Fehler (1.0 = Standard). Zusätzlich zu den
Integrationskonflikten enthalten die Quellen: ungültige PLZ (ERP), E-Mail ohne @ (Shop), Mojibake Ã¼ (CRM),
Anrede-Varianten Hr./Herrn/HERR (CRM), leere Land/PLZ-Felder (CRM), Datum nur MM/YYYY (CRM),
Telefonnummern in 5 Schreibweisen (alle), 3 Dubletten im ERP, 5 im CRM.

## Kennzahlen des Datenbestands
200 eindeutige Kunden: ERP 90 (+3 Dubletten = 93) · Shop 140 · CRM 110 (+5 Dubletten = 115 Accounts)
ERP-Aufträge 352 (davon 90 Shop-Importe mit ShopRef) · Shop-Orders 408
Überlappende Kunden 95: 40 % identische Adresse, 40 % Schreibweise, 20 % Umzug (CRM neuer)
Manuell entstandene Störungen: nachgestellte/doppelte/geschützte Leerzeichen in Firma, Straße, Adresse; Nur-Whitespace-Bemerkungen; E-Mail-Transportmüll (mailto:, "Name <mail>", GROSS, ;/. am Ende, (at))
Vertauschte Straße/Hausnummer: ~3 % im ERP-Einzelfeld ("12 Goethestr."), 2 CRM-Accounts mit vertauschten Feldern
ERP KUNDE.txt: 90 Datensätze auf ~124 physischen Zeilen – Freitextfeld Bemerkung enthält eingebettete Zeilenumbrüche (CRLF und LF)
