# Fehlerkatalog und Behandlungsregeln (Datenstand v3, 11.11.2026)

Arbeitsdokument zur BI-Portfolioleistung, Thema 1.2. Es hält fest, welche Fehlerbilder in den Beispieldaten stecken,
woher sie stammen, wie sie behandelt werden und warum. Es ist die Grundlage für Kap. 2 (Entwurf), Kap. 4 (Umsetzung)
und die Diskussion automatisch / hybrid / manuell.

## 1. Herkunft der Fehlerbilder

| Quelle | Was übernommen wurde |
|---|---|
| Skript Kap. 7–8 | Heterogenitätsarten, Harmonisierung, Schlüsselbehandlung, Record Linkage, Mängelklassen (Tab. 8.1), Historisierung – die bisherige 49-Zeilen-Checkliste |
| Projektarbeit 1 (Mandantenkonsolidierung Beck/Hokubema, rund 43.000 Kundeninstanzen, 109 Felder, OpenRefine) | Feldmissbrauch laut Kriterienkatalog (richtiger Wert im falschen Feld, Text statt Wert), Länderpräfix in der PLZ, PLZ im Ortsfeld, Länder-Varianten, Telefon mit Text, mehrere E-Mails je Feld, Sperrvermerke im Freitext, Feldlängen, erlaubte Mehrfachkunden an einer Adresse; dazu die Arbeitsregeln (siehe 2.) |
| ServiceTool (`matchingService.js`, Importlogik) | Sammel-E-Mail bei verschiedenen Kunden, Tippfehler im Namen, Fremdschlüssel per Name nicht auflösbar, Pflichtfeldregel; das Punktemodell als Alternative zur Kaskade |

## 2. Arbeitsregeln aus der PA1, die hier gelten

1. **Schwelle statt Dogma.** In der PA1 galt je Fehlerbild: unter 100 Fällen händisch, darüber eine Methode definieren
   (bei 43.000 Datensätzen). Übertragen auf 499 Kundensätze: eine eigene Regel lohnt ab etwa 5 Fällen. Die Klasse eines
   Fehlers hängt damit an drei Fragen: Ist er per Regel sicher erkennbar? Ist die Korrektur eindeutig? Was kostet eine
   falsche Korrektur gegenüber keiner – und wie viele Fälle sind es?
2. **Harte und weiche Regeln.** Hart = eindeutiger Formverstoß, wird korrigiert oder abgewiesen. Weich = auffällig,
   aber möglicherweise zulässig, wird markiert. Entspricht Tab. 8.1 im Skript (Klasse 1 automatisch, Klasse 2
   automatisch erkennbar und manuell zu korrigieren, Klasse 3 nur manuell).
3. **Reihenfolge.** Globale Vorbereitung (Whitespace, Zeilenumbrüche, Zeichensatz) vor allen Feldregeln; Land vor
   Telefon (die Vorwahl braucht ein gültiges Land); Gesperrte und Testkunden vor der Dublettensuche aussortieren.
4. **Prüfliste erst am Ende.** Wie die Tag-Spalte der PA1: markiert wird nur, was nach Standardisierung, Validierung
   und Zusammenführung noch übrig ist (`CHECK_Pruefliste`).
5. **Nichts raten.** Kein Datensatz wird wegen schlechter Qualität gelöscht. Verworfen werden nur Dubletten,
   Testkunden, Stornos und Waisen – jeweils gezählt. Unbrauchbar ist ein Kundensatz nur, wenn er weder einen Namen
   noch einen Schlüsselkandidaten (KdNr-Verweis, E-Mail, Adresse) hat.

## 3. Datenstand v3

285 eindeutige Kunden, 499 Kundensätze (ERP 134, Shop 201, CRM 164), 478 ERP-Aufträge, 549 Shop-Bestellungen.
Zwei Sollwert-Dateien, beide nicht Teil der Quellen:

- `sollwerte/kunden_mapping.csv` – welcher Kunde steckt unter welcher ID in welchem System.
- `sollwerte/fehler_log.csv` – welcher eingebaute Fehler steckt in welchem Quelldatensatz (318 Einträge). Damit wird
  neben der Kundenzuordnung auch die Erkennungsquote der Prüfregeln messbar.

## 4. Fehlerbilder, Fallzahlen, Behandlung

Fälle = Einträge im Fehlerlog (ERP / Shop / CRM). Klasse: A = automatisch, H = hybrid (Regel liefert Vorschlag oder
braucht eine zweite Bedingung, Rest auf die Prüfliste), M = manuell (nur erkennen und melden).

### 4.1 Zeichen, Format, Kodierung

| Fehlerbild | Fälle | Behandlung | Klasse | Alternative / Grenze |
|---|---|---|---|---|
| E-Mail-Transportmüll (`mailto:`, `Name <…>`, Großschrift, `;` am Ende, `(at)`) | 0 / 20 / 22 | `fnNormEmail` | A | – |
| Groß-/Kleinschreibung im Namen | 0 / 12 / 0 | Normalisierung auf Kleinschrift für den Abgleich; Anzeige unverändert | A | Proper-Case für die Anzeige wäre möglich, zerstört aber `McDonald`, `von der Heide` |
| Mojibake (`Ã¼` statt `ü`) | 0 / 0 / 12 | feste Ersetzungsliste in `fnClean` (7 Sequenzen) | A | Rückkodierung über Bytes geht in M nicht direkt; die Liste deckt nur deutsche Umlaute |
| PLZ als Zahl (Excel), führende Null fehlt | 0 / 0 / 16 | in Text wandeln, in DE auf 5 Stellen auffüllen | A | – |
| PLZ mit Länderpräfix (`D-88212`) | 7 / 0 / 3 | Präfix abtrennen; das Präfix füllt ein fehlendes Land | A | – |
| Land-Variante (`BRD`, `Germany`, `AUT`, `Deutschl.`) | 9 / 0 / 11 | Mapping-Tabelle `mapLand` | A (Tabelle manuell gepflegt) | Unbekannter Wert bleibt null und kommt auf die Prüfliste; Fuzzy auf Ländernamen wäre riskant (`Österreich`/`Australien`) |
| Land fehlt | 0 / 0 / 6 | aus PLZ-Präfix, sonst aus der PLZ-Länge (5 = DE) | A | AT und CH sind beide vierstellig: dort bleibt das Land leer → Prüfliste |
| Datum nur `MM/YYYY` | 0 / 0 / 12 | Standardtag 1., `Flag_Datum_ungenau` | A mit Flag | – |
| Zahlungsziel als Text (`30 Tage`) | 5 / 0 / 0 | Ziffern extrahieren | A | – |
| J/N-Feld mit Fremdwert (`X`, `?`, `ja`) | 5 / 0 / 0 | Mapping; Unbekanntes → null | A | – |
| Platzhalter statt leer (`-`, `k.A.`, `xxx`) | 4 / 0 / 2 | Platzhalterliste → null | A (Liste manuell) | – |

### 4.2 Feldmissbrauch (PA1)

| Fehlerbild | Fälle | Behandlung | Klasse | Alternative / Grenze |
|---|---|---|---|---|
| Hausnummer vor Straße / Felder vertauscht | 3 / 0 / 4 | Rücktausch-Regel in ERP **und** CRM | A | `Straße des 17. Juni 5` würde falsch behandelt; kommt nicht vor |
| PLZ im Ortsfeld | 5 / 0 / 0 | führende Ziffern aus dem Ort ziehen, wenn die PLZ leer ist | A | – |
| Postfach im Straßenfeld | 3 / 0 / 0 | erkennen; kein Adress-Matchkey | H | Ein Postfach ist keine Lieferadresse und kein Linkage-Merkmal |
| Hausnummer-Bereich (`2-5`, `5 + 8`, `12/1`) | 3 / 0 / 0 | Split nimmt alles ab der ersten Ziffer nach dem Straßennamen; für den Matchkey zählt nur die erste Nummer | H | – |
| Telefon mit Text oder Durchwahl | 13 / 5 / 9 | Ziffern bis zum ersten Buchstaben → E.164 (braucht Land); Resttext in `Telefon_Rest` | H | Durchwahl anhängen oder abschneiden ist eine Fachentscheidung (PA1: Durchwahl gehört zum Ansprechpartner) |
| Zwei E-Mails in einem Feld | 0 / 4 / 0 | am Trennzeichen teilen; erste Adresse führt, beide sind Match-Kandidaten | A | – |
| E-Mail mit Leerzeichen | 0 / 4 / 0 | Leerzeichen entfernen, dann prüfen | A | – |
| Bemerkung im E-Mail-Feld (`keine Mails!`) | 0 / 4 / 0 | kein `@` → E-Mail null, Text bleibt als Hinweis | A | – |
| E-Mail ohne `@` | 0 / 2 / 0 | `Email_ok = false` → Prüfliste | M | nicht ableitbar |
| Sperrvermerk im Freitext oder Firmenfeld (`INSOLVENT`, `NICHT VERWENDEN`) | 1 / 0 / 3 | Schlüsselwortliste setzt Kontaktstatus `gesperrt` und entfernt den Vermerk aus dem Namen | A (Liste manuell) | – |
| Ansprechpartner im Firmenfeld (`Dart-Shop / Hr. Koch`) | 2 / 0 / 0 | nur melden | M | Einfacher Fix für genau dieses Muster möglich (ab ` / Hr.` abschneiden); jede andere Schreibweise entgeht ihm. Unter der Schwelle |
| Privatkunde nur im Firmenfeld (`Müller Hans`) | 0 / 0 / 6 | Name wird für den Abgleich verwendet; Feldkorrektur manuell | H | Heuristik „zwei Wörter, keine Rechtsform" rät die Reihenfolge |
| Betreuer nicht auflösbar (`S. Huber`, `STEFAN HUBER`) | 0 / 0 / 7 | Normalisierung, dann Abgleich über Nachname; Rest Prüfliste | H | bei 7 Vertretern reicht das; bei 700 bräuchte es einen eigenen Abgleich |
| Feldlänge: Name nach 30 Zeichen abgeschnitten | 7 / 0 / 0 | Namensprüfung per Präfixvergleich statt Gleichheit | H | – |

### 4.3 Identität und Zuordnung

| Fehlerbild | Fälle | Behandlung | Klasse | Alternative / Grenze |
|---|---|---|---|---|
| Dubletten im ERP / im CRM | 4 / 0 / 7 | Gruppierung über Adresse + normalisierten Namen; ältester Datensatz führt | A | – |
| Fremdschlüssel `ERP_KdNr` mit Tippfehler | 0 / 0 / 3 | Schlüssel zählt nur, wenn der Name passt (`Flag_Schluessel_Name`) | H | – |
| Tippfehler im Namen (Buchstabendreher) | 6 / 0 / 0 | Fuzzy-Kandidat, bestätigt durch PLZ oder Ort | H | – |
| **Sammel-E-Mail** (zwei Kunden, ein Postfach) | 5 Paare | E-Mail-Stufe greift nur, wenn die Adresse im Ziel eindeutig ist oder der Name passt | H | Bisherige Kaskade vertraut der E-Mail blind → falsche Verschmelzung |
| **Gleiche Adresse, verschiedene Kunden** (Vereinsheim in der Gaststätte) | 8 Paare | Adress-Stufe braucht Namensbestätigung | H | PA1: Filialen und Sammeladressen wurden bewusst getrennt gelassen |
| Umzieher ohne Schlüssel und ohne E-Mail | aus dem Mapping | bleibt getrennt → Prüfliste | M | mit Stammdaten allein nicht von einem gleichnamigen Betrieb zu unterscheiden |
| Testkunden | 2 / 1 / 1 | Namensliste, filtern und zählen | A | – |
| PLZ passt nicht zum Land / PLZ fehlt / PLZ ungültig | 1 / 0 / 6 | erkennen (`Flag_PLZ_Land`, `Flag_PLZ_ok`) | M | Korrektur bräuchte ein Referenzverzeichnis; welche Seite stimmt, ist offen |

### 4.4 Bewegungsdaten

| Fehlerbild | Fälle | Behandlung | Klasse | Alternative / Grenze |
|---|---|---|---|---|
| Gutschrift (negative Menge, Status `G`) | 5 Aufträge | bleibt in den Fakten und mindert den Umsatz | A (Fachregel) | Alternative: herausfiltern und getrennt ausweisen |
| Position ohne Auftrag (Waise) | 4 | per Inner Join verworfen, in `CHECK_Fakten` gezählt | A | – |
| Shop-Auftrag doppelt ins ERP importiert | 3 | je `ShopRef` nur der erste ERP-Auftrag | A | – |
| Versand vor Auftragsdatum | 4 | Lieferdauer wird null statt negativ; Prüfliste | M | Ursache (falsches Auftrags- oder Versanddatum) nicht entscheidbar |

## 5. Änderungen an der Match-Kaskade

| Stufe | bisher | neu | Grund |
|---|---|---|---|
| 0 Vorfilter | – | Testkunden entfernen; Sperrvermerk auswerten | PA1-Reihenfolge |
| 1 Schlüssel | Name muss gleich sein | Name gleich **oder** einer ist Präfix des anderen (ab 12 Zeichen) | abgeschnittene ERP-Namen |
| 2 E-Mail | jede Übereinstimmung | nur bei eindeutiger E-Mail im Ziel, sonst zusätzlich Namensvergleich | Sammel-E-Mail |
| 3 Adresse | Matchkey gleich | Matchkey gleich **und** Name gleich/Präfix; sonst nur Kandidat für Stufe 4 | Vereinsheim, Filialen |
| 4 Fuzzy | Kandidat + PLZ-/Ort-Regel | unverändert | – |

Erwartung: Messung 6 läuft auf Datenstand v3 noch mit der alten Kaskade und zeigt falsche Verschmelzungen durch
Sammel-E-Mail und Vereinsheim. Messung 7 mit der neuen Kaskade.

## 6. Verworfene Alternative: Punktemodell

Das ServiceTool bewertet jeden Kandidaten additiv (gleicher Name 90, ähnlicher Name 70, PLZ + Ort 40, Straße 20,
E-Mail 60) und zeigt ab 60 Punkten einen Hinweis, ab 80 eine Warnung; ein Mensch bestätigt. Das ist ein
probabilistisches Record Linkage im Kleinen und per Konstruktion hybrid. Für die Integration hier nicht gewählt, weil
(1) Power Query keinen Ähnlichkeitswert aus dem Fuzzy-Join liefert, (2) die Gewichte ohne Trainingsdaten geschätzt
wären und (3) die Kaskade je Stufe einzeln messbar ist. Als Werkzeug für die Abarbeitung der Prüfliste wäre es die
passende Ergänzung.
