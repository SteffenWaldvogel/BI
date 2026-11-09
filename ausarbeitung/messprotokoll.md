# Messprotokoll Kundenintegration (Kap. 4.5)

Erfolgsmessung über `CHECK_Sollwerte` gegen `sollwerte/kunden_mapping.csv` (200 eindeutige Kunden).
Zerrissen = ein echter Kunde hat mehrere KundeKeys. Verschmolzen = ein KundeKey enthält mehrere echte Kunden.
Alle Messungen ohne manuelle Nacharbeit; Fuzzy-Schwellwert `pSchwelle` = 0,8, sofern nicht anders angegeben.

| Nr. | Datum | Stand | zerrissen | verschmolzen | ohne Mapping | Fuzzy-Treffer CRM / Shop |
| --- | --- | --- | --- | --- | --- | --- |
| 1 | 05.11.2026 | Kaskade Schlüssel(+Namensprüfung) → E-Mail → Adresse → Fuzzy(Name+Land, PLZ-Regel außer Vereine) | 5 | 12 | 0 | 19 / 0 |
| 2 | 05.11.2026 | wie 1, Quelldaten korrigiert: Namen je Ort eindeutig | 7 | 8 | 0 | 13 / 0 |
| 3a | 05.11.2026 | wie 2, Fuzzy-Treffer nur bei gleicher PLZ oder (Verein und identischer Name_norm) | 7 | 3 | 0 | 5 / 0 |
| 3b | 05.11.2026 | Fuzzy-Treffer: Verein → nur identischer Name_norm; sonst gleiche PLZ oder (identischer Name_norm und gleicher Ort) | 4 | 0 | 0 | 5 / 0 |
| 4 | 05.11.2026 | wie 3b, Schwellwert 0,9 | 4 | 0 | 0 | 5 / 0 |
| 5 | 05.11.2026 | wie 3b, Schwellwert 0,7 | 4 | 0 | 0 | 5 / 0 |

## Zwischenergebnisse und Erkenntnisse

**Vorstufe (ohne Messung):** Fuzzy-Join ohne Einschränkung lieferte in CRM_k 122 Zeilen statt 115 – mehrere Kandidaten über dem
Schwellwert vervielfachen Zeilen. Korrektur: je Zeile nur der beste Treffer. Danach 31 Fuzzy-Treffer; mit PLZ-Regel
(Treffer nur bei gleicher PLZ, Ausnahme Vereine) 19. Die Differenz waren gleichnamige Kneipen/Händler in verschiedenen Städten.

**Messung 1 → Ursachen:**
- Verschmolzen: Namenskollisionen in den generierten Quelldaten (zwei verschiedene Vereine „Black Arrows Halle (Saale)",
  zwei Kneipen „Zum Bulls Eye" in derselben Stadt). Unrealistisch dicht → Generator korrigiert (Name je Ort eindeutig;
  gleicher Name in verschiedenen Städten bleibt erlaubt, das fängt die PLZ-Regel ab).
- Zerrissen: Umzieher (Gastro/Händler) ohne gültigen ERP-Schlüssel. Die PLZ-Regel, die Verschmelzungen verhindert,
  blockiert per Definition Umzüge. Zielkonflikt Precision/Recall – bewusste Entscheidung, in der Arbeit zu begründen.

**Messung 2 → Ursachen der verbleibenden 8 Verschmelzungen:**
Alle sind Vereine mit gleichem Namensstamm in verschiedenen Städten (steeldartersaugsburg ↔ steeldartersstuttgart,
dartfreundefreiburg ↔ dartfreundeleipzig, reddragonsdresden ↔ reddragonsberlin). Bei Schwellwert 0,8 reichen
12 von 20 gleichen Zeichen. Die Vereinsausnahme von der PLZ-Regel („Ort steckt im Namen") ist falsch, weil Fuzzy alle
Zeichen gleich gewichtet und der Ortsanteil klein ist.
Geplante Korrektur: Fuzzy auf Name und Ort als getrennte Spalten – beide müssen ähnlich sein. Umzieher-Vereine
(alter Ort im Namen, neuer Ort im CRM) bleiben dann bewusst „zerrissen" → manuelle Prüfliste (Tab. 8.1, Klasse 2).

**Messung 3a → Ursache der verbleibenden 3 Verschmelzungen:** drei Vereinspaare in derselben Stadt und PLZ
(dartdevilsleipzig ↔ dartfreundeleipzig, dartsportfreiburg ↔ dartdevilsfreiburg, dartdevilsfriedrichshafen ↔ dartsportfriedrichshafen).
Die PLZ-Regel greift hier nicht, weil die PLZ gleich ist. Ursache in Messung 2 außerdem: `k{0}` nahm den *ersten* Fuzzy-Kandidaten,
nicht den ähnlichsten – Power Query liefert keinen Ähnlichkeitswert, deshalb muss die Auswahl über eigene Regeln laufen.
Entscheidung statt „Fuzzy auf Name + Ort als getrennte Spalten" (ursprünglicher Plan): bei Vereinen steckt der Ort im Namen,
also zählt nur der identische normierte Name. Für Händler/Gastro gleiche PLZ **oder** identischer Name bei gleichem Ort
(fängt die ERP-PLZ-Tippfehler 07017/70173 und 30110/3011 sowie die CRM-PLZ aus dem falschen Land ab).

**Messung 3b → die 4 verbleibenden Zerreißungen** sind alle vom Typ „Umzug ohne Schlüssel/E-Mail" (freizeit-center AT,
zumbullseye, gasthausweber, dartwelt): neue PLZ, neuer Ort, kein ERP_KdNr im CRM, keine E-Mail im ERP. Mit Stammdaten allein
nicht sicher von einem gleichnamigen Betrieb in einer anderen Stadt zu unterscheiden („dart-shop" existiert in mehreren Städten).
Bewusst „zerrissen" gelassen → Prüfliste für den Fachbereich (Tab. 8.1 Klasse 2). Alternative wäre identischer Name genügt
(Recall ↑, Precision ↓) – verworfen, weil falsche Verschmelzungen in der Umsatzanalyse schwerer wiegen als getrennte Kunden.

**Messungen 4–5 → Schwellwert ist wirkungslos geworden:** 0,7, 0,8 und 0,9 liefern identische Ergebnisse. Der Fuzzy-Join
dient nach 3b nur noch als Kandidatenlieferant; die Entscheidung trifft die Nachregel (identischer Name / gleiche PLZ / gleicher Ort).
Das ist das Gegenteil von Messung 1, wo der Schwellwert allein 19 Treffer erzeugte. Folgerung für die Arbeit: ein Ähnlichkeitsmaß
ohne fachliche Nachregel ist für Firmennamen mit Ortsbestandteil ungeeignet; der Schwellwert ist der falsche Stellhebel.
Endstand für die Faktenseite: Variante 3b mit pSchwelle = 0,8.

**Weitere Befunde unterwegs (jeweils korrigiert):**
- fnNormName: Präfix „sv " vor „tsv " ersetzt → Rest „t" blieb stehen (tdartsportravensburg). Reihenfolge längere Präfixe zuerst.
- Derselbe Fehler im Generator (TTreffpunkt 180 Weingarten in den CRM-Daten).
- Anlagedatum in KUNDE_h als Text durchgerutscht → ERP-Dubletten: falscher Datensatz als führend gewählt, Nulldatum nicht erkannt.
- Erstkontakt MM/YYYY wird von Power Query stillschweigend zum 1. des Monats → Flag_Datum_ungenau vor der Typumwandlung.
- SHOP_k: alle 75 Überlappungen über Schlüssel (18) und E-Mail (57) gefunden, Adresse/Fuzzy kamen nie zum Zug →
  Wert eines verlässlichen gemeinsamen Attributs; Gegenbeispiel CRM_k (ERP ohne E-Mail).

## Stand der Matchmethoden (Messung 3b)

CRM_k: 1 ERP-Schlüssel 38 · 3 Adresse exakt 18 · 4 Fuzzy 5 · neu 49 (15 Leads + 30 nur Shop+CRM + 4 Umzieher)
SHOP_k: 1 ERP-Schlüssel 24 · 2 E-Mail CRM 51 · neu 65 (= die 65 reinen Privatkunden)
(Zahlen aus CHECK_Sollwerte, Zeilen mit Sollwert.)

**Gegenprüfung Schlüsseltreffer (09.11.2026), Nachrechnung direkt auf den Quelldateien:** 115 CRM-Accounts, davon 39 mit ERP_KdNr.
38 davon zeigen auf eine existierende KdNr mit identischem Name_norm → 38 ist korrekt. Der 39. (ERP_KdNr 10368 am Account „Dartwelt OHG")
zeigt auf „Treffpunkt 180 Luzern e.V." – der eingebaute Tippfehler im Fremdschlüssel, den die Namensprüfung (Checkliste 30) richtig abweist
(Flag_Schluessel_Name = true). Keine Referenz zeigt auf eine verworfene ERP-Dublette. Die 43 aus Messung 2 stammen aus einem älteren
Generatorlauf: jede Generatoränderung (hier die Präfix-Korrektur) verschiebt den Zufallsstrom und damit die 60-%-Auswahl, welche Accounts
eine ERP_KdNr tragen. Die Zahl ist also kein Fehler der Integration, sondern eine Eigenschaft des jeweiligen Datenstands.

**Offener Befund im Generator (nicht korrigiert, um die Messreihe nicht zu invalidieren):** `_nkey` ersetzt „sv " vor „tsv " – derselbe
Präfixfehler wie ursprünglich in fnNormName. Folge: „TSV Steeldarters Augsburg" (10206) und „DC Steeldarters Augsburg" (10381) haben
dieselbe Name_norm in derselben Stadt, obwohl der Generator das verhindern soll. In Messung 3b harmlos (10206 hat ERP-Schlüssel im CRM,
10381 ist reiner ERP-Altkunde), aber ein realistischer Grenzfall für Kap. 8: zwei Vereine mit gleichem Namensstamm in einer Stadt
sind über Namen allein nicht trennbar → Prüfliste.
