section Section1;

shared pPfad = "C:\Users\steff\DHBW\Semester 5 (Okt-Dez 2026)\04 Business Intelligence (Kirchberg)\BI\";

shared KUNDE = let
    Quelle = Csv.Document(File.Contents(pPfad & "quellen\erp\KUNDE.txt"),[Delimiter=";", Columns=17, Encoding=1252, QuoteStyle=QuoteStyle.Csv]),
    #"Höher gestufte Header" = Table.PromoteHeaders(Quelle, [PromoteAllScalars=true]),
    #"Ersetzter Wert" = Table.ReplaceValue(#"Höher gestufte Header","#(cr)#(lf)","|",Replacer.ReplaceText,{"Bemerkung"}),
    #"Gefilterte Zeilen1" = Table.SelectRows(#"Ersetzter Wert", each true),
    #"Geänderter Typ mit Gebietsschema" = Table.TransformColumnTypes(#"Gefilterte Zeilen1", {{"Anlagedatum", type date}}, "de-DE"),
    #"Geänderter Typ mit Gebietsschema1" = Table.TransformColumnTypes(#"Geänderter Typ mit Gebietsschema", {{"GesperrtBis", type date}}, "de-DE"),
    Zeilenvorschub_LF_ersetzt = Table.ReplaceValue(#"Geänderter Typ mit Gebietsschema1", "#(lf)", " | ", Replacer.ReplaceText, {"Bemerkung"})
in
    Zeilenvorschub_LF_ersetzt;

shared customers = let
    Quelle = Csv.Document(File.Contents(pPfad & "quellen\shop\customers.csv"),[Delimiter=",", Columns=13, Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    #"Höher gestufte Header" = Table.PromoteHeaders(Quelle, [PromoteAllScalars=true]),
    #"Geänderter Typ mit Gebietsschema" = Table.TransformColumnTypes(#"Höher gestufte Header", {{"created_at", type datetimezone}}, "de-DE")
in
    #"Geänderter Typ mit Gebietsschema";

shared Accounts = let
    Quelle = Excel.Workbook(File.Contents(pPfad & "quellen\crm\CRM_Export.xlsx"), null, true),
    Accounts_Sheet = Quelle{[Item="Accounts",Kind="Sheet"]}[Data],
    #"Höher gestufte Header" = Table.PromoteHeaders(Accounts_Sheet, [PromoteAllScalars=true]),
    Flag_PLZ_Zahl = Table.AddColumn(#"Höher gestufte Header", "Flag_PLZ_Zahl", each Value.Is([PLZ], type number), type logical),
    Flag_Datum_ungenau = Table.AddColumn(Flag_PLZ_Zahl, "Flag_Datum_ungenau", each Text.Length(Text.From([Erstkontakt])) = 7, type logical),
    #"Geänderter Typ mit Gebietsschema" = Table.TransformColumnTypes(Flag_Datum_ungenau, {{"Erstkontakt", type date}}, "en-US"),
    #"Geänderter Typ mit Gebietsschema1" = Table.TransformColumnTypes(#"Geänderter Typ mit Gebietsschema", {{"LetzteAenderung", type date}}, "en-US"),
    #"Geänderter Typ" = Table.TransformColumnTypes(#"Geänderter Typ mit Gebietsschema1",{{"PLZ", type text}, {"ERP_KdNr", type text}})
in
    #"Geänderter Typ";

shared fnClean = // Globale Vorbereitung (PA1): Zeichensatzfehler, Steuerzeichen, geschützte/doppelte Leerzeichen, Platzhalter -> null
(s as nullable text) as nullable text =>
let
    moji = {{"Ã¼","ü"}, {"Ã¶","ö"}, {"Ã¤","ä"}, {"ÃŸ","ß"}, {"Ãœ","Ü"}, {"Ã–","Ö"}, {"Ã„","Ä"}},
    m0 = if s = null then null else List.Accumulate(moji, s, (acc, p) => Text.Replace(acc, p{0}, p{1})),
    t0 = if m0 = null then null else Text.Clean(m0),
    t1 = if t0 = null then null else Text.Replace(t0, Character.FromNumber(160), " "),
    t2 = if t1 = null then null else Text.Remove(t1, {Character.FromNumber(173), Character.FromNumber(8203)}),
    t3 = if t2 = null then null else Text.Combine(List.Select(Text.Split(t2, " "), each _ <> ""), " "),
    t4 = if t3 = null or t3 = "" or List.Contains({"-", "?", "k.a.", "xxx", "n/a", "unbekannt"}, Text.Lower(t3)) then null else t3
in
    t4;

shared fnNormEmail = (s as nullable text) as nullable text =>
let
    t0 = fnClean(s),
    t1 = if t0 = null then null else Text.Lower(t0),
    t2 = if t1 = null then null else Text.Replace(t1, "mailto:", ""),
    t3 = if t2 = null then null else (if Text.Contains(t2, "<") then Text.BetweenDelimiters(t2, "<", ">") else t2),
    t4 = if t3 = null then null else Text.Replace(Text.Replace(t3, "(at)", "@"), "[at]", "@"),
    t5 = if t4 = null then null else Text.Trim(t4, {" ", ";", ",", "."}),
    t6 = if t5 = null then null else (if Text.Contains(t5, " ") then Text.BeforeDelimiter(t5, " ") else t5)
in
    t6;

shared fnNormStrasse = (s as nullable text) as nullable text =>
let
    t0 = if s = null then null else Text.Lower(Text.Trim(s)),
    t1 = if t0 = null then null else
         Text.Replace(Text.Replace(Text.Replace(Text.Replace(t0, "ß","ss"), "ä","ae"), "ö","oe"), "ü","ue"),
    t2 = if t1 = null then null else Text.Replace(Text.Replace(t1, "-", ""), " ", ""),
    t3 = if t2 = null then null else Text.Replace(t2, "str.", "strasse")
in
    t3;

shared fnNormName = (s as nullable text) as nullable text =>
let
    t0 = if s = null then null else Text.Lower(Text.Trim(s)),
    t1 = if t0 = null then null else
         Text.Replace(Text.Replace(Text.Replace(Text.Replace(t0, "ß","ss"), "ä","ae"), "ö","oe"), "ü","ue"),
    // Reihenfolge wichtig: längere Präfixe (tsv) vor kürzeren (sv), sonst bleibt ein Rest stehen
    stop = {"(insolvent)", "(gesperrt)", "(nicht verwenden)", " e.v.", " e. v.", " gmbh", " e.k.", " ohg", "tsv ", "sv ", "dc ", "sc ", "dr. "},
    t2 = if t1 = null then null else List.Accumulate(stop, t1, (acc, w) => Text.Replace(acc, w, " ")),
    t3 = if t2 = null then null else Text.Replace(Text.Trim(t2), " ", "")
in
    t3;

shared mapLand = #table(type table [Quelle = text, Wert = text, LandISO = text], {
    {"ERP", "D",  "DE"}, {"ERP", "A", "AT"}, {"ERP", "CH", "CH"},
    {"CRM", "Deutschland", "DE"}, {"CRM", "Österreich", "AT"}, {"CRM", "Schweiz", "CH"},
    {"Variante", "DE", "DE"}, {"Variante", "Deutsch", "DE"}, {"Variante", "BRD", "DE"}, {"Variante", "Germany", "DE"}, {"Variante", "Deutschl.", "DE"},
    {"Variante", "AT", "AT"}, {"Variante", "AUT", "AT"}, {"Variante", "Oesterreich", "AT"}, {"Variante", "Austria", "AT"},
    {"Variante", "SUI", "CH"}, {"Variante", "Switzerland", "CH"}
});

shared mapKundentyp = #table(type table [Quelle = text, Wert = text, Kundentyp = text], {
    {"ERP", "1", "Verein"}, {"ERP", "2", "Händler"}, {"ERP", "3", "Gastro"}, {"ERP", "4", "Privat"},
    {"CRM", "Verein", "Verein"}, {"CRM", "Fachhändler", "Händler"}, {"CRM", "Großhändler", "Händler"},
    {"CRM", "Gastronomie", "Gastro"}, {"CRM", "Privat", "Privat"}
});

shared mapRegion = #table(type table [LandISO = text, Praefix = text, Region = text], {
    {"DE","0","Ost"}, {"DE","1","Ost"}, {"DE","2","Nord"}, {"DE","3","Nord"}, {"DE","4","West"},
    {"DE","5","West"}, {"DE","6","Mitte"}, {"DE","7","Süd"}, {"DE","8","Süd"}, {"DE","9","Süd"},
    {"AT","","AT"}, {"CH","","CH"}
});

shared pSchwelle = 0.8 + 0.0;

shared pAbsicherung = true and true;

shared pBelegtRegel = true and true;

shared fnNameOk = // Namensbestätigung: gleich, oder gemeinsamer Anfang bis auf höchstens 4 Zeichen des kürzeren Namens (ab 12 Zeichen).
// Fängt im ERP nach 30 Zeichen abgeschnittene Namen ab, ohne kurze Namen wie "dartshop" zu verwechseln.
(a as nullable text, b as nullable text) as logical =>
let
    n = if a = null or b = null then 0 else List.Min({Text.Length(a), Text.Length(b)}),
    p = if n = 0 then 0 else List.Count(List.FirstN(List.Zip({Text.ToList(a), Text.ToList(b)}), each _{0} = _{1})),
    ok = if a = null or b = null then false else a = b or (n >= 12 and p >= n - 4)
in
    ok;

shared fnLand = (w as nullable text) as nullable text =>
let
    r = if w = null then null else List.First(Table.SelectRows(mapLand, each Text.Lower([Wert]) = Text.Lower(w))[LandISO], null)
in
    r;

shared fnTyp = (w as nullable text) as nullable text =>
let
    r = if w = null then null else List.First(Table.SelectRows(mapKundentyp, each [Wert] = w)[Kundentyp], null)
in
    r;

shared fnJN = (v as nullable text) as nullable logical =>
let
    t = if v = null then null else Text.Lower(v),
    r = if t = null then null else if List.Contains({"j", "ja", "x", "true", "1", "y"}, t) then true else if List.Contains({"n", "nein", "false", "0"}, t) then false else null
in
    r;

shared fnStrasse = // Straße und Hausnummer trennen: Postfach, Hausnummer vor der Straße, Hausnummer-Bereiche (2-5, 5 + 8, 12/1)
(s as nullable text) as record =>
let
    teile = if s = null then {} else Text.Split(s, " "),
    ziffer = (x as text) => Text.Select(Text.Start(x, 1), {"0".."9"}) <> "",
    postfach = s <> null and Text.StartsWith(Text.Lower(s), "postfach"),
    vorn = s <> null and ziffer(s) and List.Count(teile) > 1 and not ziffer(teile{1}),
    pos = if s = null then -1 else Text.PositionOfAny(s, {"0".."9"}),
    r = if s = null then [Strasse = null, Hausnr = null, Postfach = false, Vertauscht = false]
        else if postfach then [Strasse = s, Hausnr = null, Postfach = true, Vertauscht = false]
        else if vorn then [Strasse = Text.Combine(List.Skip(teile, 1), " "), Hausnr = teile{0}, Postfach = false, Vertauscht = true]
        else if pos > 0 then [Strasse = Text.Trim(Text.Start(s, pos)), Hausnr = Text.Trim(Text.Middle(s, pos)), Postfach = false, Vertauscht = false]
        else [Strasse = s, Hausnr = null, Postfach = false, Vertauscht = false]
in
    r;

shared fnMatchkey = // Adressschlüssel wie in der PA1: Land + PLZ + normalisierte Straße + erste Hausnummer. Fehlt ein Teil, gibt es keinen Schlüssel.
(land as nullable text, plz as nullable text, strasse as nullable text, hausnr as nullable text) as nullable text =>
let
    h = if hausnr = null then null else Text.Lower(Text.SplitAny(Text.Trim(hausnr), "-+/ "){0}),
    r = if land = null or plz = null or strasse = null or h = null then null else land & plz & "|" & fnNormStrasse(strasse) & "|" & h
in
    r;

shared fnTel = // Telefon nach E.164 (PA1): nur der Teil vor Durchwahl/Text; braucht ein gültiges Land für die Vorwahl
(t as nullable text, land as nullable text) as nullable text =>
let
    h0 = if t = null then null else Text.BeforeDelimiter(t & " -", " -"),
    pos = if h0 = null then -1 else Text.PositionOfAny(h0, {"a".."z"} & {"A".."Z"}),
    h = if h0 = null then null else if pos < 0 then h0 else Text.Start(h0, pos),
    d = if h = null then "" else Text.Select(h, {"0".."9"}),
    lk = if land = "DE" then "49" else if land = "AT" then "43" else if land = "CH" then "41" else null,
    r = if Text.Length(d) < 6 then null
        else if Text.StartsWith(Text.Trim(h), "+") then "+" & d
        else if Text.StartsWith(d, "00") then "+" & Text.Middle(d, 2)
        else if Text.StartsWith(d, "0") and lk <> null then "+" & lk & Text.Middle(d, 1)
        else null
in
    r;

shared fnHatText = (t as nullable text) as logical => t <> null and Text.PositionOfAny(t, {"a".."z"} & {"A".."Z"}) >= 0;

shared fnEmails = // E-Mail-Feld zerlegen: mehrere Adressen, Leerzeichen um das @, Transportmüll; liefert nur syntaktisch gültige Adressen
(s as nullable text) as list =>
let
    t = if s = null then null else Text.Replace(Text.Replace(s, " @ ", "@"), " / ", ";"),
    teile = if t = null then {} else List.Transform(Text.SplitAny(t, ";,"), each fnNormEmail(_)),
    ok = (e as nullable text) => e <> null and List.Count(Text.Split(e, "@")) = 2 and Text.Contains(Text.AfterDelimiter(e, "@"), ".") and not Text.Contains(e, " ") and Text.Length(Text.BeforeDelimiter(e, "@")) > 0,
    r = List.Select(teile, each ok(_))
in
    r;

shared KUNDE_h = // KUNDE_h: ERP-Kunden harmonisieren. Reihenfolge wie in der PA1: globale Vorbereitung -> PLZ/Ort -> Land -> Straße/Hausnr -> Telefon -> Kodierungen -> Befunde.
let
    Txt = {"KdNr", "Firma", "Name1", "Vorname", "Strasse", "PLZ", "Ort", "Land", "Kundengruppe", "VertreterNr", "Zahlungsziel", "Newsletter", "Gesperrt", "Bemerkung", "Telefon"},
    Q = Table.TransformColumns(KUNDE, List.Transform(Txt, (c) => {c, (v) => fnClean(if v = null then null else Text.From(v)), type text})),
    // PLZ steht im Ortsfeld ("88212 Ravensburg") oder trägt ein Länderpräfix ("D-88212")
    P1 = Table.AddColumn(Q, "PLZ_aus_Ort", each if [PLZ] = null and [Ort] <> null and Text.Select(Text.Start([Ort], 1), {"0".."9"}) <> "" then Text.BeforeDelimiter([Ort], " ") else null, type text),
    P2 = Table.AddColumn(P1, "Ort_ok", each if [PLZ_aus_Ort] <> null then Text.AfterDelimiter([Ort], " ") else [Ort], type text),
    P3 = Table.AddColumn(P2, "PLZ_Praefix", each if [PLZ] <> null and Text.Contains([PLZ], "-") then Text.Upper(Text.BeforeDelimiter([PLZ], "-")) else null, type text),
    P4 = Table.AddColumn(P3, "PLZ_roh", each if [PLZ_aus_Ort] <> null then [PLZ_aus_Ort] else if [PLZ_Praefix] <> null then Text.AfterDelimiter([PLZ], "-") else [PLZ], type text),
    L1 = Table.AddColumn(P4, "LandISO", each fnLand([Land]) ?? fnLand([PLZ_Praefix]), type text),
    P5 = Table.AddColumn(L1, "PLZ_norm", each if [PLZ_roh] = null then null else if [LandISO] = "DE" and Text.Length([PLZ_roh]) = 4 then "0" & [PLZ_roh] else [PLZ_roh], type text),
    P6 = Table.AddColumn(P5, "Flag_PLZ_ok", each [PLZ_norm] <> null and Text.Select([PLZ_norm], {"0".."9"}) = [PLZ_norm] and Text.Length([PLZ_norm]) = (if [LandISO] = "DE" then 5 else 4), type logical),
    // Straße / Hausnummer
    S1 = Table.AddColumn(P6, "S", each fnStrasse([Strasse])),
    S2 = Table.ExpandRecordColumn(S1, "S", {"Strasse", "Hausnr", "Postfach", "Vertauscht"}, {"Strasse_ok", "Hausnr", "Flag_Postfach", "Flag_Vertauscht"}),
    S3 = Table.AddColumn(S2, "Matchkey", each if [Flag_Postfach] or not [Flag_PLZ_ok] then null else fnMatchkey([LandISO], [PLZ_norm], [Strasse_ok], [Hausnr]), type text),
    // Name, Titel, Kodierungen
    N1 = Table.AddColumn(S3, "Titel", each if [Name1] <> null and Text.StartsWith([Name1], "Dr.") then "Dr." else null, type text),
    N2 = Table.AddColumn(N1, "Nachname", each if [Titel] <> null then Text.Trim(Text.AfterDelimiter([Name1], "Dr.")) else [Name1], type text),
    N3 = Table.AddColumn(N2, "Kundentyp", each fnTyp([Kundengruppe]), type text),
    N4 = Table.AddColumn(N3, "Name_norm", each fnNormName(if [Firma] <> null then [Firma] else [Nachname]), type text),
    N5 = Table.AddColumn(N4, "Flag_Test", each [Name_norm] <> null and Text.StartsWith([Name_norm], "test"), type logical),
    K1 = Table.AddColumn(N5, "Newsletter_ok", each fnJN([Newsletter]), type logical),
    K2 = Table.AddColumn(K1, "Gesperrt_ok", each fnJN([Gesperrt]) ?? false, type logical),
    K3 = Table.AddColumn(K2, "Flag_Sperrvermerk", each [Bemerkung] <> null and List.AnyTrue(List.Transform({"insolvent", "gesperrt", "nicht verwenden"}, (w) => Text.Contains(Text.Lower([Bemerkung]), w))), type logical),
    K4 = Table.AddColumn(K3, "Kontaktstatus", each if [Gesperrt_ok] or [Flag_Sperrvermerk] then "gesperrt" else null, type text),
    K5 = Table.AddColumn(K4, "Zahlungsziel_ok", each if [Zahlungsziel] = null or Text.StartsWith([Zahlungsziel], "-") or Text.Select([Zahlungsziel], {"0".."9"}) = "" then null else Number.FromText(Text.Select([Zahlungsziel], {"0".."9"})), Int64.Type),
    K6 = Table.AddColumn(K5, "Telefon_E164", each fnTel([Telefon], [LandISO]), type text),
    // Befunde: was nach der automatischen Bereinigung übrig bleibt (Prüfliste, Tab. 8.1 Klasse 2 und 3)
    B = Table.AddColumn(K6, "Befunde", each Text.Combine(List.RemoveNulls({
            if [Flag_Test] then "Testkunde (gefiltert)" else null,
            if [LandISO] = null then "Land fehlt oder unbekannt" else null,
            if not [Flag_PLZ_ok] then "PLZ fehlt oder ungültig" else null,
            if [Flag_Postfach] then "Postfach statt Straße (kein Adressschlüssel)" else null,
            if fnHatText([Telefon]) then "Telefon enthält Text" else null,
            if [Telefon] <> null and [Telefon_E164] = null then "Telefon nicht normalisierbar" else null,
            if [Firma] <> null and (Text.Contains([Firma], " / ") or Text.Contains([Firma], "Hr.") or Text.Contains([Firma], "Fr.")) then "Firma enthält Zusatz (Ansprechpartner?)" else null,
            if [Firma] <> null and Text.Length([Firma]) >= 30 then "Name evtl. abgeschnitten (Feldlänge 30)" else null,
            if [Flag_Sperrvermerk] and not [Gesperrt_ok] then "Sperrvermerk nur im Freitext" else null}), "; "), type text),
    Sel = Table.SelectColumns(B, {"KdNr", "Firma", "Name1", "Nachname", "Vorname", "PLZ_norm", "Ort_ok", "VertreterNr", "Zahlungsziel_ok", "Anlagedatum", "Newsletter_ok", "Gesperrt_ok", "GesperrtBis", "Bemerkung", "LandISO", "Strasse_ok", "Hausnr", "Titel", "Kundentyp", "Matchkey", "Name_norm", "Kontaktstatus", "Telefon_E164", "Flag_Test", "Flag_PLZ_ok", "Flag_Postfach", "Flag_Vertauscht", "Flag_Sperrvermerk", "Befunde"}),
    Ren = Table.RenameColumns(Sel, {{"Ort_ok", "Ort"}, {"Zahlungsziel_ok", "Zahlungsziel"}, {"Newsletter_ok", "Newsletter"}, {"Gesperrt_ok", "Gesperrt"}, {"Strasse_ok", "Strasse"}}),
    Typ = Table.TransformColumnTypes(Ren, {{"Strasse", type text}, {"Hausnr", type text}, {"Flag_Postfach", type logical}, {"Flag_Vertauscht", type logical}}),
    Nulldatum = Table.ReplaceValue(Typ, #date(9999, 12, 31), null, Replacer.ReplaceValue, {"GesperrtBis"}),
    Quelle = Table.AddColumn(Nulldatum, "Quelle", each "ERP", type text)
in
    Quelle;

shared SHOP_h = // SHOP_h: Shop-Kunden harmonisieren. Adresse steht in einem Feld ("Straße Nr, PLZ Ort"), E-Mail-Feld kann mehrere Adressen oder Text enthalten.
let
    Txt = {"customer_id", "email", "first_name", "last_name", "gender", "address", "country", "is_club", "newsletter", "erp_customer_no", "company", "phone"},
    Q = Table.TransformColumns(customers, List.Transform(Txt, (c) => {c, (v) => fnClean(if v = null then null else Text.From(v)), type text})),
    A1 = Table.AddColumn(Q, "Adr1", each if [address] = null then null else Text.BeforeDelimiter([address], ", "), type text),
    A2 = Table.AddColumn(A1, "Adr2", each if [address] = null or not Text.Contains([address], ", ") then null else Text.AfterDelimiter([address], ", "), type text),
    A3 = Table.AddColumn(A2, "S", each fnStrasse([Adr1])),
    A4 = Table.ExpandRecordColumn(A3, "S", {"Strasse", "Hausnr", "Postfach"}, {"Strasse", "Hausnr", "Flag_Postfach"}),
    A5 = Table.AddColumn(A4, "PLZ_norm", each if [Adr2] = null then null else Text.BeforeDelimiter([Adr2], " "), type text),
    A6 = Table.AddColumn(A5, "Ort", each if [Adr2] = null then null else Text.AfterDelimiter([Adr2], " "), type text),
    A7 = Table.AddColumn(A6, "LandISO", each fnLand([country]) ?? [country], type text),
    A8 = Table.AddColumn(A7, "Matchkey", each fnMatchkey([LandISO], [PLZ_norm], [Strasse], [Hausnr]), type text),
    K1 = Table.AddColumn(A8, "Kundentyp", each if fnJN([is_club]) = true then "Verein" else if [company] <> null then null else "Privat", type text),
    K2 = Table.AddColumn(K1, "Anrede", each if [gender] = "1" then "Herr" else if [gender] = "2" then "Frau" else null, type text),
    M1 = Table.AddColumn(K2, "Mails", each fnEmails([email])),
    M2 = Table.AddColumn(M1, "Email_norm", each List.First([Mails], null), type text),
    M3 = Table.AddColumn(M2, "Email_2", each if List.Count([Mails]) > 1 then [Mails]{1} else null, type text),
    N1 = Table.AddColumn(M3, "Name_norm", each fnNormName(if [company] <> null then [company] else [last_name]), type text),
    N2 = Table.AddColumn(N1, "Flag_Test", each ([Name_norm] <> null and Text.StartsWith([Name_norm], "test")) or [Email_norm] = "test@test.de", type logical),
    // UTC -> mitteleuropäische Zeit, sonst rutscht eine Bestellung kurz vor Mitternacht auf den Vortag
    D1 = Table.AddColumn(N2, "Anlagedatum", each Date.From(DateTimeZone.RemoveZone(DateTimeZone.SwitchZone([created_at], 1))), type date),
    T1 = Table.AddColumn(D1, "Telefon_E164", each fnTel([phone], [LandISO]), type text),
    B = Table.AddColumn(T1, "Befunde", each Text.Combine(List.RemoveNulls({
            if [Flag_Test] then "Testkunde (gefiltert)" else null,
            if [email] <> null and [Email_norm] = null then "E-Mail ungültig oder Text im E-Mail-Feld" else null,
            if [Email_2] <> null then "zwei E-Mails in einem Feld (erste führt)" else null,
            if [Matchkey] = null then "Adresse unvollständig (kein Adressschlüssel)" else null,
            if fnHatText([phone]) then "Telefon enthält Text" else null,
            if [phone] <> null and [Telefon_E164] = null then "Telefon nicht normalisierbar" else null}), "; "), type text),
    Ren = Table.RenameColumns(B, {{"customer_id", "ShopID"}, {"erp_customer_no", "ERP_KdNr"}, {"company", "Firma"}, {"first_name", "Vorname"}, {"last_name", "Nachname"}}),
    L1 = Table.AddColumn(Ren, "is_club_ok", each fnJN([is_club]) ?? false, type logical),
    L2 = Table.AddColumn(L1, "Newsletter", each fnJN([newsletter]), type logical),
    Sel = Table.SelectColumns(L2, {"ShopID", "Vorname", "Nachname", "Strasse", "Hausnr", "PLZ_norm", "Ort", "LandISO", "is_club_ok", "Newsletter", "Anlagedatum", "ERP_KdNr", "Firma", "Kundentyp", "Email_norm", "Email_2", "Anrede", "Matchkey", "Name_norm", "Telefon_E164", "Flag_Test", "Befunde"}),
    Ren2 = Table.RenameColumns(Sel, {{"is_club_ok", "is_club"}}),
    Typ = Table.TransformColumnTypes(Ren2, {{"Strasse", type text}, {"Hausnr", type text}}),
    Quelle = Table.AddColumn(Typ, "Quelle", each "Shop", type text)
in
    Quelle;

shared CRM_h = // CRM_h: CRM-Accounts harmonisieren. Vom Vertrieb gepflegt: Freitext, Excel-Artefakte, vertauschte Felder, Sperrvermerke im Firmenfeld.
let
    Txt = {"AccountId", "Firma", "Anrede", "Titel", "Vorname", "Nachname", "Strasse", "Hausnr", "PLZ", "Ort", "Land", "Kundentyp", "Segment", "Betreuer", "ERP_KdNr", "Newsletter", "Status", "E-Mail", "Telefon"},
    Q = Table.TransformColumns(Accounts, List.Transform(Txt, (c) => {c, (v) => fnClean(if v = null then null else Text.From(v)), type text})),
    // Sperrvermerk im Firmenfeld auswerten und aus dem Namen entfernen
    F1 = Table.AddColumn(Q, "Flag_Sperrvermerk", each [Firma] <> null and List.AnyTrue(List.Transform({"(insolvent)", "(gesperrt)", "(nicht verwenden)"}, (w) => Text.Contains(Text.Lower([Firma]), w))), type logical),
    F2 = Table.AddColumn(F1, "Firma_ok", each if [Flag_Sperrvermerk] then Text.Trim(Text.BeforeDelimiter([Firma], "(", {0, RelativePosition.FromEnd})) else [Firma], type text),
    // PLZ: Länderpräfix abtrennen; Zahl aus Excel ist durch den Texttyp schon Text, die führende Null fehlt aber
    P1 = Table.AddColumn(F2, "PLZ_Praefix", each if [PLZ] <> null and Text.Contains([PLZ], "-") then Text.Upper(Text.BeforeDelimiter([PLZ], "-")) else null, type text),
    P2 = Table.AddColumn(P1, "PLZ_roh", each if [PLZ_Praefix] <> null then Text.AfterDelimiter([PLZ], "-") else [PLZ], type text),
    L0 = Table.AddColumn(P2, "LandISO_roh", each fnLand([Land]) ?? fnLand([PLZ_Praefix]), type text),
    // fehlendes Land: fünfstellige PLZ ist in DACH nur in Deutschland möglich; vierstellig bleibt offen (AT oder CH)
    L1 = Table.AddColumn(L0, "LandISO", each if [LandISO_roh] <> null then [LandISO_roh] else if [PLZ_roh] <> null and Text.Length([PLZ_roh]) = 5 then "DE" else null, type text),
    P3 = Table.AddColumn(L1, "PLZ_norm", each if [PLZ_roh] = null then null else if [LandISO] = "DE" and Text.Length([PLZ_roh]) = 4 and [Flag_PLZ_Zahl] then "0" & [PLZ_roh] else [PLZ_roh], type text),
    P4 = Table.AddColumn(P3, "Flag_PLZ_Land", each [PLZ_norm] <> null and (([LandISO] = "DE" and Text.Length([PLZ_norm]) <> 5) or ([LandISO] <> "DE" and [LandISO] <> null and Text.Length([PLZ_norm]) <> 4)), type logical),
    // Straße/Hausnr vertauscht (Fehlbelegung bei manueller Erfassung)
    S0 = Table.AddColumn(P4, "Flag_Vertauscht", each [Strasse] <> null and [Hausnr] <> null and Text.Select(Text.Start([Strasse], 1), {"0".."9"}) <> "" and Text.Select(Text.Start([Hausnr], 1), {"0".."9"}) = "", type logical),
    S1 = Table.AddColumn(S0, "Strasse_ok", each if [Flag_Vertauscht] then [Hausnr] else [Strasse], type text),
    S2 = Table.AddColumn(S1, "Hausnr_ok", each if [Flag_Vertauscht] then [Strasse] else [Hausnr], type text),
    S3 = Table.AddColumn(S2, "Matchkey", each if [Flag_PLZ_Land] then null else fnMatchkey([LandISO], [PLZ_norm], [Strasse_ok], [Hausnr_ok]), type text),
    // Kodierungen
    K1 = Table.AddColumn(S3, "Kundentyp_ok", each fnTyp([Kundentyp]), type text),
    K2 = Table.AddColumn(K1, "Anrede_ok", each let a = Text.Lower(Text.Start([Anrede] ?? "", 2)) in if a = "he" or a = "hr" then "Herr" else if a = "fr" then "Frau" else null, type text),
    K3 = Table.AddColumn(K2, "Newsletter_ok", each fnJN([Newsletter]) ?? false, type logical),
    K4 = Table.AddColumn(K3, "Kontaktstatus", each if [Flag_Sperrvermerk] then "gesperrt" else [Status], type text),
    // Privatkunde nur im Firmenfeld ("Müller Hans"): für den Abgleich zählt das erste Wort als Nachname, die Feldkorrektur bleibt manuell
    N0 = Table.AddColumn(K4, "Flag_Privat_in_Firma", each [Kundentyp_ok] = "Privat" and [Nachname] = null and [Firma_ok] <> null, type logical),
    N1 = Table.AddColumn(N0, "Name_norm", each fnNormName(if [Flag_Privat_in_Firma] then Text.BeforeDelimiter([Firma_ok] & " ", " ") else if [Firma_ok] <> null then [Firma_ok] else [Nachname]), type text),
    N2 = Table.AddColumn(N1, "Flag_Test", each [Name_norm] <> null and Text.StartsWith([Name_norm], "test"), type logical),
    M1 = Table.AddColumn(N2, "Email_norm", each List.First(fnEmails([#"E-Mail"]), null), type text),
    T1 = Table.AddColumn(M1, "Telefon_E164", each fnTel([Telefon], [LandISO]), type text),
    // Betreuer ist ein Fremdschlüssel per Name: Abgleich über den normalisierten Nachnamen gegen das Blatt Betreuer
    BT = Table.Buffer(Table.AddColumn(Betreuer, "NN", each fnNormName(List.Last(Text.Split([Name], " "))), type text)),
    R1 = Table.AddColumn(T1, "Betreuer_ok", each if [Betreuer] = null then null else List.First(Table.SelectRows(BT, (b) => b[NN] = fnNormName(List.Last(Text.Split([Betreuer], " "))))[Name], null), type text),
    B = Table.AddColumn(R1, "Befunde", each Text.Combine(List.RemoveNulls({
            if [Flag_Test] then "Testkunde (gefiltert)" else null,
            if [LandISO] = null then "Land fehlt oder unbekannt" else null,
            if [PLZ_norm] = null then "PLZ fehlt oder ungültig" else null,
            if [Flag_PLZ_Land] then "PLZ passt nicht zum Land" else null,
            if [Flag_Datum_ungenau] then "Erstkontakt nur monatsgenau" else null,
            if [Betreuer] <> null and [Betreuer_ok] = null then "Betreuer nicht auflösbar" else null,
            if [Flag_Privat_in_Firma] then "Privatname im Firmenfeld" else null,
            if [#"E-Mail"] <> null and [Email_norm] = null then "E-Mail ungültig oder Text im E-Mail-Feld" else null,
            if fnHatText([Telefon]) then "Telefon enthält Text" else null,
            if [Telefon] <> null and [Telefon_E164] = null then "Telefon nicht normalisierbar" else null,
            if [Flag_Sperrvermerk] then "Sperrvermerk im Firmenfeld" else null}), "; "), type text),
    Sel = Table.SelectColumns(B, {"AccountId", "Firma_ok", "Anrede_ok", "Titel", "Vorname", "Nachname", "Strasse_ok", "Hausnr_ok", "PLZ", "Ort", "Segment", "Betreuer_ok", "ERP_KdNr", "Erstkontakt", "Newsletter_ok", "Kontaktstatus", "LetzteAenderung", "Telefon_E164", "LandISO", "PLZ_norm", "Kundentyp_ok", "Email_norm", "Name_norm", "Matchkey", "Flag_PLZ_Land", "Flag_Datum_ungenau", "Flag_Vertauscht", "Flag_Sperrvermerk", "Flag_Privat_in_Firma", "Flag_Test", "Befunde"}),
    Ren = Table.RenameColumns(Sel, {{"Firma_ok", "Firma"}, {"Anrede_ok", "Anrede"}, {"Strasse_ok", "Strasse"}, {"Hausnr_ok", "Hausnr"}, {"Segment", "ABC_Klasse"}, {"Betreuer_ok", "Betreuer"}, {"Newsletter_ok", "Newsletter"}, {"Kundentyp_ok", "Kundentyp"}}),
    Quelle = Table.AddColumn(Ren, "Quelle", each "CRM", type text)
in
    Quelle;

shared KUNDE_k = // KUNDE_k: ERP als Basis der Kundenintegration. Testkunden raus, Dubletten (gleiche Adresse + gleicher Name) auf den ältesten Datensatz, KundeKey als Surrogat.
let
    Quelle = Table.SelectRows(KUNDE_h, each not [Flag_Test]),
    #"Sortierte Zeilen" = Table.Buffer(Table.Sort(Quelle, {{"Matchkey", Order.Ascending}, {"Name_norm", Order.Ascending}, {"Anlagedatum", Order.Ascending}})),
    MitSchluessel = Table.Distinct(Table.SelectRows(#"Sortierte Zeilen", each [Matchkey] <> null), {"Name_norm", "Matchkey"}),
    OhneSchluessel = Table.SelectRows(#"Sortierte Zeilen", each [Matchkey] = null),
    Alle = Table.Combine({MitSchluessel, OhneSchluessel}),
    #"Hinzugefügte benutzerdefinierte Spalte" = Table.AddColumn(Alle, "Matchmethode", each "0 Basis ERP", type text),
    KundeKey_vergeben = Table.Buffer(Table.AddIndexColumn(#"Hinzugefügte benutzerdefinierte Spalte", "KundeKey", 1, 1, Int64.Type))
in
    KundeKey_vergeben;

shared ERP_Dubletten = // ERP_Dubletten (Zeile 47): Kunden, die im ERP mehrfach angelegt sind – gleicher Matchkey + normalisierter Name.
// Führend ist der älteste Datensatz (Anlagedatum); die anderen KdNr werden später auf denselben KundeKey gemappt.
let
    Q = Table.SelectRows(KUNDE_h, each [Matchkey] <> null and not [Flag_Test]),
    G = Table.Group(Q, {"Matchkey", "Name_norm"}, {{"Anzahl", each Table.RowCount(_), Int64.Type}, {"Zeilen", each _, type table}}),
    F = Table.SelectRows(G, each [Anzahl] > 1),
    E = Table.ExpandTableColumn(F, "Zeilen", {"KdNr", "Firma", "Anlagedatum"}),
    S = Table.Sort(E, {{"Matchkey", Order.Ascending}, {"Anlagedatum", Order.Ascending}}),
    // Führende KdNr je Gruppe = erste nach Sortierung
    L = Table.Group(S, {"Matchkey"}, {{"KdNr_fuehrend", each List.First([KdNr]), type text}}),
    J = Table.NestedJoin(S, {"Matchkey"}, L, {"Matchkey"}, "L", JoinKind.LeftOuter),
    X = Table.ExpandTableColumn(J, "L", {"KdNr_fuehrend"}),
    R = Table.AddColumn(X, "Rolle", each if [KdNr] = [KdNr_fuehrend] then "führend" else "verworfen", type text)
in
    R;

shared CRM_k = // CRM_k: jedem CRM-Account einen KundeKey zuordnen. Kaskade: 1 ERP-Schlüssel (mit Namens- oder Adressbestätigung) -> 3 Adresse (mit Namensbestätigung) -> 4 Fuzzy Name -> neu.
// Stufe 2 (E-Mail) entfällt, weil das ERP keine E-Mail führt. pAbsicherung = false schaltet auf die ungesicherte Kaskade zurück (Messung 6).
let
    S = Table.SelectRows(CRM_h, each not [Flag_Test]),
    ERP = Table.Buffer(Table.SelectColumns(KUNDE_k, {"KdNr", "KundeKey", "Name_norm", "Matchkey", "LandISO", "PLZ_norm", "Ort", "Kundentyp"})),
    // ---- Stufe 1: ERP-Schlüssel. Der Schlüssel zählt nur, wenn Name oder Adresse ihn bestätigen (Tippfehler im Fremdschlüssel).
    fKey = (kd as nullable text, name as nullable text, mk as nullable text) as record =>
        let t = if kd = null then null else Table.SelectRows(ERP, (r) => r[KdNr] = kd),
            e = if t = null or Table.IsEmpty(t) then null else t{0},
            ok = if e = null then false else if pAbsicherung then fnNameOk(e[Name_norm], name) or (mk <> null and e[Matchkey] = mk) else e[Name_norm] = name
        in [Key = if ok then e[KundeKey] else null, Flag = e <> null and not ok],
    A1 = Table.AddColumn(S, "K1", each fKey([ERP_KdNr], [Name_norm], [Matchkey])),
    A2 = Table.ExpandRecordColumn(A1, "K1", {"Key", "Flag"}, {"Key_Schluessel", "Flag_Schluessel_Name"}),
    // ---- Stufe 3: Adresse. Gleiche Adresse allein genügt nicht (Vereinsheim in der Gaststätte, Filialen) – der Name muss passen.
    fAdr = (mk as nullable text, name as nullable text) =>
        let t = if mk = null then null else Table.SelectRows(ERP, (r) => r[Matchkey] = mk),
            k = if t = null then null else if pAbsicherung then Table.SelectRows(t, (r) => fnNameOk(r[Name_norm], name)) else t
        in if k = null or Table.IsEmpty(k) then null else k{0}[KundeKey],
    A3 = Table.AddColumn(A2, "Key_Adr", each fAdr([Matchkey], [Name_norm]), Int64.Type),
    Vor = Table.AddColumn(A3, "Key_13", each [Key_Schluessel] ?? [Key_Adr], Int64.Type),
    // ---- Stufe 4: Fuzzy auf normalisierten Namen + Land, nur für noch offene Zeilen. Fuzzy liefert Kandidaten, die Nachregel entscheidet.
    Offen = Table.SelectRows(Vor, each [Key_13] = null),
    Zu = Table.SelectRows(Vor, each [Key_13] <> null),
    // Belegt-Regel: ein ERP-Kunde, dem schon ein Account per Schlüssel oder Adresse zugeordnet ist, nimmt keinen weiteren Account
    // allein über Name + PLZ/Ort auf (umgezogener Betrieb trifft am neuen Ort auf einen gleichnamigen).
    Belegt = List.Buffer(List.RemoveNulls(Zu[Key_13])),
    ERPn = Table.Distinct(Table.SelectColumns(ERP, {"Name_norm", "LandISO", "PLZ_norm", "Ort", "Kundentyp", "KundeKey", "Matchkey"}), {"Name_norm", "LandISO", "PLZ_norm"}),
    FJ = Table.FuzzyNestedJoin(Offen, {"Name_norm", "LandISO"}, ERPn, {"Name_norm", "LandISO"}, "FZ", JoinKind.LeftOuter, [Threshold = pSchwelle, IgnoreCase = true, IgnoreSpace = true]),
    FE = Table.RemoveColumns(Table.AddColumn(FJ, "Key_Fuzzy", each let k = Table.SelectRows([FZ], (r) =>
            (pAbsicherung and [Matchkey] <> null and r[Matchkey] = [Matchkey])
            or (if [Kundentyp] = "Verein" then r[Name_norm] = [Name_norm]
                else not (pBelegtRegel and List.Contains(Belegt, r[KundeKey])) and (([PLZ_norm] <> null and r[PLZ_norm] = [PLZ_norm]) or (r[Name_norm] = [Name_norm] and Text.StartsWith(Text.Lower([Ort] ?? ""), Text.Lower(r[Ort] ?? "#")))))) in if Table.IsEmpty(k) then null else k{0}[KundeKey]), {"FZ"}),
    OffenK = Table.RemoveColumns(Table.AddColumn(FE, "KundeKey", each [Key_Fuzzy], Int64.Type), {"Key_Fuzzy"}),
    ZuK = Table.AddColumn(Zu, "KundeKey", each [Key_13], Int64.Type),
    Beide = Table.Combine({ZuK, OffenK}),
    M = Table.AddColumn(Beide, "Matchmethode", each
          if [Key_Schluessel] <> null then "1 ERP-Schlüssel"
          else if [Key_Adr] <> null then "3 Adresse + Name"
          else if [KundeKey] <> null then "4 Fuzzy Name"
          else "neu", type text),
    // ---- Neue Keys ab 1001: Leads, Umzieher ohne Schlüssel, Rest. CRM-Dubletten teilen sich einen Key.
    Neu = Table.SelectRows(M, each [KundeKey] = null),
    Alt = Table.SelectRows(M, each [KundeKey] <> null),
    NeuG = Table.Group(Neu, {"Matchkey", "Name_norm"}, {{"Z", each _, type table}}),
    NeuI = Table.AddIndexColumn(NeuG, "NeuKey", 1001, 1, Int64.Type),
    NeuE = Table.ExpandTableColumn(NeuI, "Z", List.RemoveItems(Table.ColumnNames(Neu), {"Matchkey", "Name_norm", "KundeKey"})),
    NeuK = Table.RenameColumns(NeuE, {{"NeuKey", "KundeKey"}}),
    Fertig = Table.Buffer(Table.RemoveColumns(Table.Combine({Alt, NeuK}), {"Key_Adr", "Key_13"}))
in
    Fertig;

shared SHOP_k = // SHOP_k: jedem Shop-Kunden einen KundeKey zuordnen. Kaskade: 1 ERP-Schlüssel -> 2 E-Mail (CRM) -> 3 Adresse + Name (ERP, dann CRM) -> 4 Fuzzy (nur gewerblich) -> neu ab 2001.
let
    S = Table.SelectRows(SHOP_h, each not [Flag_Test]),
    ERP = Table.Buffer(Table.SelectColumns(KUNDE_k, {"KdNr", "KundeKey", "Name_norm", "Matchkey", "LandISO", "PLZ_norm", "Ort", "Kundentyp"})),
    CRM = Table.Buffer(Table.SelectColumns(CRM_k, {"KundeKey", "Name_norm", "Matchkey", "LandISO", "PLZ_norm", "Ort", "Kundentyp", "Email_norm"})),
    // ---- Stufe 1: ERP-Schlüssel mit Namens- oder Adressbestätigung
    fKey = (kd as nullable text, name as nullable text, mk as nullable text) as record =>
        let t = if kd = null then null else Table.SelectRows(ERP, (r) => r[KdNr] = kd),
            e = if t = null or Table.IsEmpty(t) then null else t{0},
            ok = if e = null then false else if pAbsicherung then fnNameOk(e[Name_norm], name) or (mk <> null and e[Matchkey] = mk) else e[Name_norm] = name
        in [Key = if ok then e[KundeKey] else null, Flag = e <> null and not ok],
    A1 = Table.AddColumn(S, "K1", each fKey([ERP_KdNr], [Name_norm], [Matchkey])),
    A2 = Table.ExpandRecordColumn(A1, "K1", {"Key", "Flag"}, {"Key_Schluessel", "Flag_Schluessel_Name"}),
    // ---- Stufe 2: E-Mail gegen CRM. Bei mehreren Kunden mit derselben Adresse (Sammelpostfach) entscheidet der Name;
    //      ohne Namensbestätigung zählt die E-Mail nur, wenn sie im CRM genau einem Kunden gehört.
    fMail = (mail as nullable text, name as nullable text) =>
        let t = if mail = null then null else Table.SelectRows(CRM, (r) => r[Email_norm] = mail),
            n = if t = null then null else Table.SelectRows(t, (r) => fnNameOk(r[Name_norm], name))
        in if t = null or Table.IsEmpty(t) then null
           else if not pAbsicherung then t{0}[KundeKey]
           else if not Table.IsEmpty(n) then n{0}[KundeKey]
           else if List.Count(List.Distinct(t[KundeKey])) = 1 then t{0}[KundeKey]
           else null,
    A3 = Table.AddColumn(A2, "Key_Mail", each fMail([Email_norm], [Name_norm]) ?? fMail([Email_2], [Name_norm]), Int64.Type),
    // ---- Stufe 3: Adresse + Name, erst ERP, dann CRM
    fAdr = (quelle as table, mk as nullable text, name as nullable text) =>
        let t = if mk = null then null else Table.SelectRows(quelle, (r) => r[Matchkey] = mk),
            k = if t = null then null else if pAbsicherung then Table.SelectRows(t, (r) => fnNameOk(r[Name_norm], name)) else t
        in if k = null or Table.IsEmpty(k) then null else k{0}[KundeKey],
    A4 = Table.AddColumn(A3, "Key_AdrERP", each fAdr(ERP, [Matchkey], [Name_norm]), Int64.Type),
    A5 = Table.AddColumn(A4, "Key_AdrCRM", each fAdr(CRM, [Matchkey], [Name_norm]), Int64.Type),
    Vor = Table.AddColumn(A5, "Key_123", each [Key_Schluessel] ?? [Key_Mail] ?? [Key_AdrERP] ?? [Key_AdrCRM], Int64.Type),
    // ---- Stufe 4: Fuzzy nur für gewerbliche Kunden; Privatnamen (Müller) sind zu häufig.
    Offen = Table.SelectRows(Vor, each [Key_123] = null and [Kundentyp] <> "Privat"),
    Rest  = Table.SelectRows(Vor, each [Key_123] <> null or [Kundentyp] = "Privat"),
    Sp = {"Name_norm", "LandISO", "PLZ_norm", "Ort", "Kundentyp", "KundeKey", "Matchkey"},
    Namen = Table.Distinct(Table.Combine({Table.SelectColumns(ERP, Sp), Table.SelectColumns(CRM, Sp)}), {"Name_norm", "LandISO", "PLZ_norm"}),
    FJ = Table.FuzzyNestedJoin(Offen, {"Name_norm", "LandISO"}, Namen, {"Name_norm", "LandISO"}, "FZ", JoinKind.LeftOuter, [Threshold = pSchwelle, IgnoreCase = true, IgnoreSpace = true]),
    FE = Table.RemoveColumns(Table.AddColumn(FJ, "Key_Fuzzy", each let k = Table.SelectRows([FZ], (r) =>
            (pAbsicherung and [Matchkey] <> null and r[Matchkey] = [Matchkey])
            or (if [Kundentyp] = "Verein" then r[Name_norm] = [Name_norm]
                else (([PLZ_norm] <> null and r[PLZ_norm] = [PLZ_norm]) or (r[Name_norm] = [Name_norm] and Text.StartsWith(Text.Lower([Ort] ?? ""), Text.Lower(r[Ort] ?? "#")))))) in if Table.IsEmpty(k) then null else k{0}[KundeKey]), {"FZ"}),
    OffenK = Table.RemoveColumns(Table.AddColumn(FE, "KundeKey", each [Key_Fuzzy], Int64.Type), {"Key_Fuzzy"}),
    RestK = Table.AddColumn(Rest, "KundeKey", each [Key_123], Int64.Type),
    Beide = Table.Combine({RestK, OffenK}),
    M = Table.AddColumn(Beide, "Matchmethode", each
          if [Key_Schluessel] <> null then "1 ERP-Schlüssel"
          else if [Key_Mail] <> null then "2 E-Mail CRM"
          else if [Key_AdrERP] <> null then "3 Adresse ERP"
          else if [Key_AdrCRM] <> null then "3 Adresse CRM"
          else if [KundeKey] <> null then "4 Fuzzy Name"
          else "neu", type text),
    // ---- Neue Keys ab 2001 (Privatkunden nur im Shop)
    Neu = Table.SelectRows(M, each [KundeKey] = null),
    Alt = Table.SelectRows(M, each [KundeKey] <> null),
    NeuI = Table.AddIndexColumn(Table.RemoveColumns(Neu, {"KundeKey"}), "KundeKey", 2001, 1, Int64.Type),
    Fertig = Table.Buffer(Table.RemoveColumns(Table.Combine({Alt, NeuI}), {"Key_Mail", "Key_AdrERP", "Key_AdrCRM", "Key_123"}))
in
    Fertig;

shared MAP_SCHLUESSEL = // MAP_SCHLUESSEL (Checkliste 22): Mapping Quellschlüssel <-> DW-Schlüssel für alle drei Quellen,
// inkl. der im ERP verworfenen Dubletten-KdNr, die auf den KundeKey des führenden Datensatzes zeigen.
let
    E = Table.SelectColumns(Table.AddColumn(KUNDE_k, "QuellID", each [KdNr], type text), {"Quelle", "QuellID", "KundeKey", "Matchmethode"}),
    D0 = Table.SelectRows(ERP_Dubletten, each [Rolle] = "verworfen"),
    D1 = Table.NestedJoin(D0, {"KdNr_fuehrend"}, KUNDE_k, {"KdNr"}, "K", JoinKind.LeftOuter),
    D2 = Table.ExpandTableColumn(D1, "K", {"KundeKey"}),
    D = Table.SelectColumns(Table.AddColumn(Table.AddColumn(Table.AddColumn(D2, "Quelle", each "ERP", type text), "QuellID", each [KdNr], type text), "Matchmethode", each "0 ERP-Dublette", type text), {"Quelle", "QuellID", "KundeKey", "Matchmethode"}),
    C = Table.SelectColumns(Table.AddColumn(CRM_k, "QuellID", each [AccountId], type text), {"Quelle", "QuellID", "KundeKey", "Matchmethode"}),
    S = Table.SelectColumns(Table.AddColumn(SHOP_k, "QuellID", each [ShopID], type text), {"Quelle", "QuellID", "KundeKey", "Matchmethode"}),
    Alle = Table.Combine({E, D, C, S}),
    Sort = Table.Sort(Alle, {{"KundeKey", Order.Ascending}, {"Quelle", Order.Ascending}})
in
    Sort;

shared CHECK_Sollwerte = // CHECK_Sollwerte: Erfolgsmessung der Kundenintegration gegen die bekannte Wahrheit aus der Datengenerierung.
// Zerrissen = ein echter Kunde hat mehrere KundeKeys. Falsch verschmolzen = ein KundeKey enthält mehrere echte Kunden.
let
    Roh = Csv.Document(File.Contents(pPfad & "sollwerte\kunden_mapping.csv"), [Delimiter = ";", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Soll = Table.PromoteHeaders(Roh, [PromoteAllScalars = true]),
    Lang1 = Table.RenameColumns(Table.SelectRows(Table.SelectColumns(Table.AddColumn(Soll, "Quelle", each "ERP"), {"Quelle", "ERP_KdNr", "KundeKey", "Abweichung"}), each [ERP_KdNr] <> ""), {{"ERP_KdNr", "QuellID"}, {"KundeKey", "SollKey"}}),
    Lang2 = Table.RenameColumns(Table.SelectRows(Table.SelectColumns(Table.AddColumn(Soll, "Quelle", each "CRM"), {"Quelle", "CRM_ID", "KundeKey", "Abweichung"}), each [CRM_ID] <> ""), {{"CRM_ID", "QuellID"}, {"KundeKey", "SollKey"}}),
    Lang3 = Table.RenameColumns(Table.SelectRows(Table.SelectColumns(Table.AddColumn(Soll, "Quelle", each "Shop"), {"Quelle", "Shop_ID", "KundeKey", "Abweichung"}), each [Shop_ID] <> ""), {{"Shop_ID", "QuellID"}, {"KundeKey", "SollKey"}}),
    Lang = Table.Combine({Lang1, Lang2, Lang3}),
    J = Table.NestedJoin(Lang, {"Quelle", "QuellID"}, MAP_SCHLUESSEL, {"Quelle", "QuellID"}, "M", JoinKind.LeftOuter),
    JE = Table.ExpandTableColumn(J, "M", {"KundeKey", "Matchmethode"}, {"IstKey", "Matchmethode"}),
    // Fehlerart 1: je SollKey mehr als ein IstKey
    Z = Table.Group(JE, {"SollKey"}, {{"IstKeys", each List.Count(List.Distinct([IstKey])), Int64.Type}}),
    Zerrissen = Table.RowCount(Table.SelectRows(Z, each [IstKeys] > 1)),
    // Fehlerart 2: je IstKey mehr als ein SollKey
    V = Table.Group(Table.SelectRows(JE, each [IstKey] <> null), {"IstKey"}, {{"SollKeys", each List.Count(List.Distinct([SollKey])), Int64.Type}}),
    Verschmolzen = Table.RowCount(Table.SelectRows(V, each [SollKeys] > 1)),
    Fehlend = Table.RowCount(Table.SelectRows(JE, each [IstKey] = null)),
    Detail = Table.Group(JE, {"Quelle", "Matchmethode", "Abweichung"}, {{"Anzahl", each Table.RowCount(_), Int64.Type}}),
    Ergebnis = Table.InsertRows(Detail, 0, {
        [Quelle = "GESAMT", Matchmethode = "zerrissene Kunden", Abweichung = "", Anzahl = Zerrissen],
        [Quelle = "GESAMT", Matchmethode = "falsch verschmolzene Keys", Abweichung = "", Anzahl = Verschmolzen],
        [Quelle = "GESAMT", Matchmethode = "Quell-IDs ohne Mapping", Abweichung = "", Anzahl = Fehlend]})
in
    Ergebnis;

shared CHECK_Detail = // CHECK_Detail: die konkreten Problemfälle aus der Erfolgsmessung – zerrissene Kunden und falsch verschmolzene Keys mit Namen.
let
    Roh = Csv.Document(File.Contents(pPfad & "sollwerte\kunden_mapping.csv"), [Delimiter = ";", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Soll = Table.PromoteHeaders(Roh, [PromoteAllScalars = true]),
    Lang1 = Table.RenameColumns(Table.SelectRows(Table.SelectColumns(Table.AddColumn(Soll, "Quelle", each "ERP"), {"Quelle", "ERP_KdNr", "KundeKey", "Abweichung", "Kundentyp"}), each [ERP_KdNr] <> ""), {{"ERP_KdNr", "QuellID"}, {"KundeKey", "SollKey"}}),
    Lang2 = Table.RenameColumns(Table.SelectRows(Table.SelectColumns(Table.AddColumn(Soll, "Quelle", each "CRM"), {"Quelle", "CRM_ID", "KundeKey", "Abweichung", "Kundentyp"}), each [CRM_ID] <> ""), {{"CRM_ID", "QuellID"}, {"KundeKey", "SollKey"}}),
    Lang3 = Table.RenameColumns(Table.SelectRows(Table.SelectColumns(Table.AddColumn(Soll, "Quelle", each "Shop"), {"Quelle", "Shop_ID", "KundeKey", "Abweichung", "Kundentyp"}), each [Shop_ID] <> ""), {{"Shop_ID", "QuellID"}, {"KundeKey", "SollKey"}}),
    Lang = Table.Combine({Lang1, Lang2, Lang3}),
    J = Table.NestedJoin(Lang, {"Quelle", "QuellID"}, MAP_SCHLUESSEL, {"Quelle", "QuellID"}, "M", JoinKind.LeftOuter),
    JE = Table.ExpandTableColumn(J, "M", {"KundeKey", "Matchmethode"}, {"IstKey", "Matchmethode"}),
    // Namen dazu, damit man die Fälle lesen kann
    N = Table.Combine({
        Table.SelectColumns(Table.AddColumn(KUNDE_h, "QuellID", each [KdNr], type text), {"Quelle", "QuellID", "Name_norm", "Matchkey"}),
        Table.SelectColumns(Table.AddColumn(CRM_h, "QuellID", each [AccountId], type text), {"Quelle", "QuellID", "Name_norm", "Matchkey"}),
        Table.SelectColumns(Table.AddColumn(SHOP_h, "QuellID", each [ShopID], type text), {"Quelle", "QuellID", "Name_norm", "Matchkey"})}),
    JN = Table.ExpandTableColumn(Table.NestedJoin(JE, {"Quelle", "QuellID"}, N, {"Quelle", "QuellID"}, "N", JoinKind.LeftOuter), "N", {"Name_norm", "Matchkey"}),
    // Gruppenzähler
    Z = Table.Group(JN, {"SollKey"}, {{"IstKeys_je_Soll", each List.Count(List.Distinct([IstKey])), Int64.Type}}),
    V = Table.Group(Table.SelectRows(JN, each [IstKey] <> null), {"IstKey"}, {{"SollKeys_je_Ist", each List.Count(List.Distinct([SollKey])), Int64.Type}}),
    A = Table.ExpandTableColumn(Table.NestedJoin(JN, {"SollKey"}, Z, {"SollKey"}, "Z", JoinKind.LeftOuter), "Z", {"IstKeys_je_Soll"}),
    B = Table.ExpandTableColumn(Table.NestedJoin(A, {"IstKey"}, V, {"IstKey"}, "V", JoinKind.LeftOuter), "V", {"SollKeys_je_Ist"}),
    P = Table.AddColumn(B, "Problem", each if [IstKeys_je_Soll] > 1 and [SollKeys_je_Ist] > 1 then "beides" else if [IstKeys_je_Soll] > 1 then "zerrissen" else if [SollKeys_je_Ist] > 1 then "verschmolzen" else null, type text),
    F = Table.SelectRows(P, each [Problem] <> null),
    S = Table.Sort(F, {{"Problem", Order.Ascending}, {"IstKey", Order.Ascending}, {"SollKey", Order.Ascending}})
in
    S;

shared STACK_KUNDE = let
    Quelle = let Spalten = {"KundeKey","Quelle","Firma","Titel","Vorname","Nachname","Anrede","Kundentyp","LandISO","PLZ_norm","Ort","Strasse","Hausnr","Matchkey","Newsletter","Anlagedatum","Aenderungsdatum","Kontaktstatus","VertreterNr","Betreuer","Email_norm","KdNr","ShopID","AccountId","Telefon_E164"}, E = Table.AddColumn(Table.AddColumn(KUNDE_k, "Prio", each 1), "Aenderungsdatum", each try [Anlagedatum] otherwise null, type date), C = Table.AddColumn(Table.AddColumn(Table.RenameColumns(CRM_k, {{"Erstkontakt","Anlagedatum"}}, MissingField.Ignore), "Prio", each 2), "Aenderungsdatum", each try [LetzteAenderung] otherwise null, type date), S = Table.AddColumn(Table.AddColumn(SHOP_k, "Prio", each 3), "Aenderungsdatum", each try [Anlagedatum] otherwise null, type date), Alle = Table.Combine({Table.SelectColumns(E, Spalten & {"Prio"}, MissingField.UseNull), Table.SelectColumns(C, Spalten & {"Prio"}, MissingField.UseNull), Table.SelectColumns(S, Spalten & {"Prio"}, MissingField.UseNull)}) in Alle
in
    Quelle;

shared DIM_KUNDE = let
    Quelle = let S = Table.Sort(STACK_KUNDE, {{"KundeKey", Order.Ascending}, {"Prio", Order.Ascending}}), erst = (l as list) => List.First(List.RemoveNulls(l)), G = Table.Group(S, {"KundeKey"}, {{"Firma", each erst([Firma]), type text}, {"Titel", each erst([Titel]), type text}, {"Vorname", each erst([Vorname]), type text}, {"Nachname", each erst([Nachname]), type text}, {"Anrede", each erst([Anrede]), type text}, {"Kundentyp", each erst([Kundentyp]) ?? "Privat", type text}, {"LandISO", each erst([LandISO]), type text}, {"Newsletter", each List.AnyTrue(List.RemoveNulls([Newsletter])), type logical}, {"Erstkontakt", each List.Min([Anlagedatum]), type date}, {"Kontaktstatus", each erst([Kontaktstatus]) ?? "Kunde", type text}, {"VertreterNr", each erst([VertreterNr]), type text}, {"Betreuer", each erst([Betreuer]), type text}, {"ERP_KdNr", each erst([KdNr]), type text}, {"ShopID", each erst([ShopID]), type text}, {"AccountId", each erst([AccountId]), type text}, {"Telefon", each erst([Telefon_E164]), type text}, {"Quellen", each Text.Combine(List.Distinct([Quelle]), "+"), type text}, {"AnzahlQuellen", each List.Count(List.Distinct([Quelle])), Int64.Type}}) in G
in
    Quelle;

shared DIM_KUNDE_ADRESSE = let
    Quelle = let A = Table.SelectRows(STACK_KUNDE, each [Matchkey] <> null), G = Table.Group(A, {"KundeKey","Matchkey"}, {{"GueltigVonQuelle", each List.Min([Aenderungsdatum]), type date}, {"Strasse", each List.First(List.RemoveNulls([Strasse])), type text}, {"Hausnr", each List.First(List.RemoveNulls([Hausnr])), type text}, {"PLZ", each List.First(List.RemoveNulls([PLZ_norm])), type text}, {"Ort", each List.First(List.RemoveNulls([Ort])), type text}, {"LandISO", each List.First(List.RemoveNulls([LandISO])), type text}}), S = Table.Sort(G, {{"KundeKey", Order.Ascending}, {"GueltigVonQuelle", Order.Ascending}}), V = Table.Group(S, {"KundeKey"}, {{"T", each Table.AddIndexColumn(_, "Version", 1, 1, Int64.Type), type table}}), VE = Table.ExpandTableColumn(V, "T", {"Matchkey","GueltigVonQuelle","Strasse","Hausnr","PLZ","Ort","LandISO","Version"}), Von = Table.AddColumn(VE, "GueltigVon", each if [Version] = 1 then #date(1900,1,1) else [GueltigVonQuelle], type date), Next = Table.TransformColumns(Table.SelectColumns(Von, {"KundeKey","Version","GueltigVon"}), {{"Version", each _ - 1, Int64.Type}}), J = Table.NestedJoin(Von, {"KundeKey","Version"}, Next, {"KundeKey","Version"}, "N", JoinKind.LeftOuter), JE = Table.ExpandTableColumn(J, "N", {"GueltigVon"}, {"NaechsteVon"}), Bis = Table.AddColumn(JE, "GueltigBis", each if [NaechsteVon] = null then null else Date.AddDays([NaechsteVon], -1), type date), Cur = Table.AddColumn(Bis, "CurrentFlag", each [GueltigBis] = null, type logical), Pr = Table.AddColumn(Cur, "Praefix", each if [LandISO] = "DE" then Text.Start([PLZ], 1) else "", type text), RJ = Table.NestedJoin(Pr, {"LandISO","Praefix"}, mapRegion, {"LandISO","Praefix"}, "R", JoinKind.LeftOuter), RE = Table.ExpandTableColumn(RJ, "R", {"Region"}), Fertig = Table.RemoveColumns(RE, {"NaechsteVon","Praefix"}) in Fertig
in
    Quelle;

shared AUFTRAG = let
    Quelle = let Quelle = Csv.Document(File.Contents(pPfad & "quellen\erp\AUFTRAG.txt"), [Delimiter=";", Encoding=1252, QuoteStyle=QuoteStyle.Csv]), Header = Table.PromoteHeaders(Quelle, [PromoteAllScalars=true]), Typen = Table.TransformColumnTypes(Header, {{"AuftragNr", type text}, {"KdNr", Int64.Type}, {"Datum", type date}, {"VertreterNr", type text}, {"Status", type text}, {"Waehrung", type text}, {"Lieferart", type text}, {"Zahlungsziel", Int64.Type}, {"ShopRef", type text}}, "de-DE"), LeerZuNull = Table.ReplaceValue(Typen, "", null, Replacer.ReplaceValue, {"VertreterNr", "ShopRef"}) in LeerZuNull
in
    Quelle;

shared AUFTRAGSPOS = let
    Quelle = let Quelle = Csv.Document(File.Contents(pPfad & "quellen\erp\AUFTRAGSPOS.txt"), [Delimiter=";", Encoding=1252, QuoteStyle=QuoteStyle.Csv]), Header = Table.PromoteHeaders(Quelle, [PromoteAllScalars=true]), Typen = Table.TransformColumnTypes(Header, {{"AuftragNr", type text}, {"Pos", Int64.Type}, {"ArtNr", type text}, {"Menge", Int64.Type}, {"Einheit", type text}, {"EPreisNetto", type number}}, "de-DE") in Typen
in
    Quelle;

shared ZAHLUNG = let
    Quelle = let Quelle = Csv.Document(File.Contents(pPfad & "quellen\erp\ZAHLUNG.txt"), [Delimiter=";", Encoding=1252, QuoteStyle=QuoteStyle.Csv]), Header = Table.PromoteHeaders(Quelle, [PromoteAllScalars=true]), Typen = Table.TransformColumnTypes(Header, {{"ZahlungNr", type text}, {"AuftragNr", type text}, {"Datum", type date}, {"Betrag", type number}, {"Zahlart", type text}}, "de-DE") in Typen
in
    Quelle;

shared LIEFERUNG = let
    Quelle = let Quelle = Csv.Document(File.Contents(pPfad & "quellen\erp\LIEFERUNG.txt"), [Delimiter=";", Encoding=1252, QuoteStyle=QuoteStyle.Csv]), Header = Table.PromoteHeaders(Quelle, [PromoteAllScalars=true]), Typen = Table.TransformColumnTypes(Header, {{"LieferNr", type text}, {"AuftragNr", type text}, {"Versanddatum", type date}, {"Lieferart", type text}}, "de-DE") in Typen
in
    Quelle;

shared VERTRETER = let
    Quelle = let Quelle = Csv.Document(File.Contents(pPfad & "quellen\erp\VERTRETER.txt"), [Delimiter=";", Encoding=1252, QuoteStyle=QuoteStyle.Csv]), Header = Table.PromoteHeaders(Quelle, [PromoteAllScalars=true]), Typen = Table.TransformColumnTypes(Header, {{"VertreterNr", type text}, {"Name", type text}, {"Gebiet", type text}, {"Austritt", type text}}) in Typen
in
    Quelle;

shared ARTIKEL = let
    Quelle = let Quelle = Csv.Document(File.Contents(pPfad & "quellen\erp\ARTIKEL.txt"), [Delimiter=";", Encoding=1252, QuoteStyle=QuoteStyle.Csv]), Header = Table.PromoteHeaders(Quelle, [PromoteAllScalars=true]), Typen = Table.TransformColumnTypes(Header, {{"ArtNr", type text}, {"Bezeichnung", type text}, {"EKPreis", type number}}, "de-DE") in Typen
in
    Quelle;

shared Betreuer = let
    Quelle = let Quelle = Excel.Workbook(File.Contents(pPfad & "quellen\crm\CRM_Export.xlsx"), null, true), Blatt = Quelle{[Item="Betreuer", Kind="Sheet"]}[Data], Header = Table.PromoteHeaders(Blatt, [PromoteAllScalars=true]), Typen = Table.TransformColumnTypes(Header, {{"Name", type text}, {"E-Mail", type text}, {"Rolle", type text}, {"Region", type text}}) in Typen
in
    Quelle;

shared orders = let
    Quelle = let Quelle = Csv.Document(File.Contents(pPfad & "quellen\shop\orders.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]), Header = Table.PromoteHeaders(Quelle, [PromoteAllScalars=true]), Typen = Table.TransformColumnTypes(Header, {{"order_id", type text}, {"customer_id", type text}, {"order_date", type datetimezone}, {"status", type text}, {"total_gross", type number}, {"shipping_method_id", Int64.Type}, {"coupon_code", type text}}, "en-US"), LeerZuNull = Table.ReplaceValue(Typen, "", null, Replacer.ReplaceValue, {"coupon_code"}) in LeerZuNull
in
    Quelle;

shared order_items = let
    Quelle = let Quelle = Csv.Document(File.Contents(pPfad & "quellen\shop\order_items.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]), Header = Table.PromoteHeaders(Quelle, [PromoteAllScalars=true]), Typen = Table.TransformColumnTypes(Header, {{"order_id", type text}, {"line_no", Int64.Type}, {"sku", type text}, {"qty", Int64.Type}, {"unit_price_gross", type number}}, "en-US") in Typen
in
    Quelle;

shared products = let
    Quelle = let Quelle = Csv.Document(File.Contents(pPfad & "quellen\shop\products.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]), Header = Table.PromoteHeaders(Quelle, [PromoteAllScalars=true]), Typen = Table.TransformColumnTypes(Header, {{"sku", type text}, {"name", type text}, {"category", type text}}) in Typen
in
    Quelle;

shared shipping_methods = let
    Quelle = let Quelle = Csv.Document(File.Contents(pPfad & "quellen\shop\shipping_methods.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]), Header = Table.PromoteHeaders(Quelle, [PromoteAllScalars=true]), Typen = Table.TransformColumnTypes(Header, {{"id", Int64.Type}, {"name", type text}}) in Typen
in
    Quelle;

shared mapMwSt = let
    Quelle = #table(type table [LandISO = text, MwStFaktor = number], {{"DE", 1.19}, {"AT", 1.20}, {"CH", 1.081}})
in
    Quelle;

shared pKursCHF = let
    Quelle = 0.96 + 0.0 + 0.0
in
    Quelle;

shared fnSaison = let
    Quelle = (d as date) as record => let m = Date.Month(d), j = Date.Year(d), Saison = if m >= 9 and m <= 10 then "1 Saisonstart (Sep-Okt)" else if m >= 11 or m = 1 then "2 WM-/Hauptsaison (Nov-Jan)" else if m >= 2 and m <= 4 then "3 Rückrunde (Feb-Apr)" else "4 Sommerpause (Mai-Aug)", Saisonjahr = if m >= 9 then Text.From(j) & "/" & Text.From(j + 1) else Text.From(j - 1) & "/" & Text.From(j) in [Saison = Saison, Saisonjahr = Saisonjahr]
in
    Quelle;

shared DIM_PRODUKT = let
    Quelle = let Q = ARTIKEL, S = Table.AddColumn(Q, "SKU_abgeleitet", each "1" & Text.End([ArtNr], 4), type text), G = Table.AddColumn(S, "GruppeCode", each Text.Start([ArtNr], 2), type text), P = Table.AddColumn(G, "Produktgruppe", each if [GruppeCode] = "FL" then "Flights" else if [GruppeCode] = "SH" then "Shafts" else if [GruppeCode] = "BA" then "Barrels" else if [GruppeCode] = "BO" then "Boards" else if [GruppeCode] = "ZU" then "Zubehör" else "unbekannt", type text), J = Table.NestedJoin(P, {"SKU_abgeleitet"}, products, {"sku"}, "Shop", JoinKind.LeftOuter), E = Table.ExpandTableColumn(J, "Shop", {"sku", "name", "category"}, {"Shop_sku", "Shop_name", "Shop_category"}), F1 = Table.AddColumn(E, "Flag_Gruppe_abweichend", each [Shop_category] <> null and [Shop_category] <> [Produktgruppe], type logical), F2 = Table.AddColumn(F1, "Flag_ohne_Shop", each [Shop_sku] = null, type logical), I = Table.AddIndexColumn(F2, "ProduktKey", 1, 1, Int64.Type), Sel = Table.SelectColumns(I, {"ProduktKey", "ArtNr", "Shop_sku", "Bezeichnung", "Shop_name", "Produktgruppe", "EKPreis", "Flag_Gruppe_abweichend", "Flag_ohne_Shop"}) in Sel
in
    Quelle;

shared DIM_VERKAEUFER = let
    Quelle = let V = VERTRETER, VS = Table.SplitColumn(V, "Name", Splitter.SplitTextByDelimiter(", ", QuoteStyle.None), {"Nachname", "Vorname"}), VA = Table.ReplaceValue(VS, "31.12.9999", null, Replacer.ReplaceValue, {"Austritt"}), VT = Table.TransformColumnTypes(VA, {{"Austritt", type date}}, "de-DE"), VK = Table.AddColumn(VT, "Aktiv", each [Austritt] = null, type logical), VN = Table.AddColumn(VK, "Name_norm", each fnNormName([Nachname] & [Vorname]), type text), B = Betreuer, BS = Table.SplitColumn(B, "Name", Splitter.SplitTextByEachDelimiter({" "}, QuoteStyle.None, false), {"Vorname_B", "Nachname_B"}), BN = Table.AddColumn(BS, "Name_norm", each fnNormName([Nachname_B] & [Vorname_B]), type text), J = Table.NestedJoin(VN, {"Name_norm"}, BN, {"Name_norm"}, "B", JoinKind.LeftOuter), E = Table.ExpandTableColumn(J, "B", {"E-Mail", "Rolle", "Region"}), K = Table.AddColumn(E, "VerkaeuferKey", each [VertreterNr], type text), O = #table(type table [VerkaeuferKey = text, Nachname = text, Vorname = text, Rolle = text, Region = text, Aktiv = logical], {{"ONLINE", "Online-Shop", null, "Kanal", "Online", true}}), Alle = Table.Combine({Table.SelectColumns(K, {"VerkaeuferKey", "Nachname", "Vorname", "Rolle", "Gebiet", "Region", "E-Mail", "Austritt", "Aktiv"}), O}) in Alle
in
    Quelle;

shared FAKT_ERP = let
    Quelle = let P = AUFTRAGSPOS, AOK = Table.Combine({Table.SelectRows(AUFTRAG, each [ShopRef] = null), Table.Distinct(Table.Buffer(Table.Sort(Table.SelectRows(AUFTRAG, each [ShopRef] <> null), {{"AuftragNr", Order.Ascending}})), {"ShopRef"})}), J1 = Table.NestedJoin(P, {"AuftragNr"}, AOK, {"AuftragNr"}, "A", JoinKind.Inner), E1 = Table.ExpandTableColumn(J1, "A", {"KdNr", "Datum", "VertreterNr", "Status", "Waehrung", "Lieferart", "Zahlungsziel", "ShopRef"}), L = Table.Group(LIEFERUNG, {"AuftragNr"}, {{"Versanddatum_letzt", each List.Max([Versanddatum]), type date}, {"AnzLieferungen", each Table.RowCount(_), Int64.Type}}), Z = Table.Group(ZAHLUNG, {"AuftragNr"}, {{"Zahldatum_letzt", each List.Max([Datum]), type date}, {"Zahlbetrag", each List.Sum([Betrag]), type number}, {"AnzZahlungen", each Table.RowCount(_), Int64.Type}}), J2 = Table.NestedJoin(E1, {"AuftragNr"}, L, {"AuftragNr"}, "L", JoinKind.LeftOuter), E2 = Table.ExpandTableColumn(J2, "L", {"Versanddatum_letzt"}), J3 = Table.NestedJoin(E2, {"AuftragNr"}, Z, {"AuftragNr"}, "Z", JoinKind.LeftOuter), E3 = Table.ExpandTableColumn(J3, "Z", {"Zahldatum_letzt", "Zahlbetrag"}), KH = Table.AddColumn(Table.SelectColumns(KUNDE_h, {"KdNr", "VertreterNr"}), "KdNr_txt", each Text.From([KdNr]), type text), Q = Table.AddColumn(E3, "QuellID", each Text.From([KdNr]), type text), J4 = Table.NestedJoin(Q, {"QuellID"}, KH, {"KdNr_txt"}, "K", JoinKind.LeftOuter), E4 = Table.ExpandTableColumn(J4, "K", {"VertreterNr"}, {"Stamm_VertreterNr"}), M = Table.AddColumn(E4, "MengeStk", each if [Einheit] = "Set" then [Menge] * 10 else [Menge], Int64.Type), Ku = Table.AddColumn(M, "Kurs", each if [Waehrung] = "CHF" then pKursCHF else 1, type number), N = Table.AddColumn(Ku, "NettoEUR", each Number.Round([MengeStk] * [EPreisNetto] * [Kurs], 2), type number), Ka = Table.AddColumn(N, "Kanal", each if [ShopRef] <> null then "Online" else "Direkt", type text), Vk = Table.AddColumn(Ka, "VerkaeuferKey", each if [VertreterNr] <> null then [VertreterNr] else if [ShopRef] <> null then "ONLINE" else [Stamm_VertreterNr], type text), Ld = Table.AddColumn(Vk, "Lieferdauer_Tage", each if [Versanddatum_letzt] = null or [Versanddatum_letzt] < [Datum] then null else Duration.Days([Versanddatum_letzt] - [Datum]), Int64.Type), Zv = Table.AddColumn(Ld, "Zahlungsverzug_Tage", each if [Zahldatum_letzt] = null then null else Duration.Days([Zahldatum_letzt] - Date.AddDays([Datum], [Zahlungsziel])), Int64.Type), Qu = Table.AddColumn(Zv, "Quelle", each "ERP", type text), R = Table.RenameColumns(Qu, {{"AuftragNr", "AuftragID"}, {"Datum", "Auftragsdatum"}, {"ShopRef", "Fremdbeleg"}}), Sel = Table.SelectColumns(R, {"Quelle", "QuellID", "AuftragID", "Pos", "Auftragsdatum", "VerkaeuferKey", "Kanal", "Status", "Lieferart", "ArtNr", "MengeStk", "NettoEUR", "Lieferdauer_Tage", "Zahlungsverzug_Tage", "Fremdbeleg"}) in Sel
in
    Quelle;

shared FAKT_SHOP = let
    Quelle = let I = order_items, J1 = Table.NestedJoin(I, {"order_id"}, orders, {"order_id"}, "O", JoinKind.LeftOuter), E1 = Table.ExpandTableColumn(J1, "O", {"customer_id", "order_date", "status", "shipping_method_id"}), F = Table.SelectRows(E1, each [status] = "shipped" or [status] = "paid"), A = Table.NestedJoin(F, {"order_id"}, AUFTRAG, {"ShopRef"}, "ERP", JoinKind.LeftAnti), A2 = Table.RemoveColumns(A, {"ERP"}), J2 = Table.NestedJoin(A2, {"customer_id"}, SHOP_h, {"ShopID"}, "K", JoinKind.LeftOuter), E2 = Table.ExpandTableColumn(J2, "K", {"LandISO"}), J3 = Table.NestedJoin(E2, {"LandISO"}, mapMwSt, {"LandISO"}, "M", JoinKind.LeftOuter), E3 = Table.ExpandTableColumn(J3, "M", {"MwStFaktor"}), J4 = Table.NestedJoin(E3, {"shipping_method_id"}, shipping_methods, {"id"}, "V", JoinKind.LeftOuter), E4 = Table.ExpandTableColumn(J4, "V", {"name"}, {"Versandart"}), La = Table.AddColumn(E4, "Lieferart", each if [Versandart] = "Abholung" then "Abholung" else "Paket", type text), N = Table.AddColumn(La, "NettoEUR", each Number.Round([qty] * [unit_price_gross] / [MwStFaktor], 2), type number), D = Table.AddColumn(N, "Auftragsdatum", each Date.From(DateTimeZone.RemoveZone(DateTimeZone.SwitchZone([order_date], 1))), type date), St = Table.AddColumn(D, "Status", each if [status] = "shipped" then "L" else "O", type text), J5 = Table.NestedJoin(St, {"sku"}, DIM_PRODUKT, {"Shop_sku"}, "P", JoinKind.LeftOuter), E5 = Table.ExpandTableColumn(J5, "P", {"ArtNr"}), C1 = Table.AddColumn(E5, "Quelle", each "Shop", type text), C2 = Table.AddColumn(C1, "VerkaeuferKey", each "ONLINE", type text), C3 = Table.AddColumn(C2, "Kanal", each "Online", type text), C4 = Table.AddColumn(C3, "Lieferdauer_Tage", each null, Int64.Type), C5 = Table.AddColumn(C4, "Zahlungsverzug_Tage", each null, Int64.Type), C6 = Table.AddColumn(C5, "Fremdbeleg", each null, type text), R = Table.RenameColumns(C6, {{"order_id", "AuftragID"}, {"line_no", "Pos"}, {"qty", "MengeStk"}, {"customer_id", "QuellID"}}), Sel = Table.SelectColumns(R, {"Quelle", "QuellID", "AuftragID", "Pos", "Auftragsdatum", "VerkaeuferKey", "Kanal", "Status", "Lieferart", "ArtNr", "MengeStk", "NettoEUR", "Lieferdauer_Tage", "Zahlungsverzug_Tage", "Fremdbeleg"}) in Sel
in
    Quelle;

shared FAKT_VERKAUF = let
    C = Table.Combine({FAKT_ERP, FAKT_SHOP}),
    MAPt = Table.Buffer(Table.TransformColumns(MAP_SCHLUESSEL, {{"KundeKey", each Text.From(_), type text}})),
    J1 = Table.NestedJoin(C, {"Quelle", "QuellID"}, MAPt, {"Quelle", "QuellID"}, "M", JoinKind.LeftOuter),
    E1 = Table.ExpandTableColumn(J1, "M", {"KundeKey"}),
    DKt = Table.Buffer(Table.TransformColumns(DIM_KUNDE, {{"KundeKey", each Text.From(_), type text}})),
    J2 = Table.NestedJoin(E1, {"KundeKey"}, DKt, {"KundeKey"}, "K", JoinKind.LeftOuter),
    // Stamm_VertreterNr = betreuender Vertreter laut Kundenstamm (nur ERP-Kunden), unabhaengig vom Verkaeufer des Auftrags -> F5 Kanalanteil
    E2 = Table.ExpandTableColumn(J2, "K", {"Kundentyp", "VertreterNr"}, {"Kundentyp", "Stamm_VertreterNr"}),
    DAt = Table.Buffer(Table.TransformColumns(DIM_KUNDE_ADRESSE, {{"KundeKey", each Text.From(_), type text}})),
    J3 = Table.NestedJoin(E2, {"KundeKey"}, DAt, {"KundeKey"}, "A", JoinKind.LeftOuter),
    Ad = Table.AddColumn(J3, "Adresse", each let d = [Auftragsdatum], t = Table.SelectRows([A], (r) => r[GueltigVon] <= d and (r[GueltigBis] = null or d <= r[GueltigBis])) in if Table.IsEmpty(t) then null else t{0}),
    E3 = Table.ExpandRecordColumn(Ad, "Adresse", {"LandISO", "Region", "PLZ"}, {"LandISO", "Region", "PLZ"}),
    R1 = Table.RemoveColumns(E3, {"A"}),
    Y = Table.AddColumn(R1, "Jahr", each Date.Year([Auftragsdatum]), Int64.Type),
    Qn = Table.AddColumn(Y, "Quartal", each "Q" & Text.From(Date.QuarterOfYear([Auftragsdatum])), type text),
    JQ = Table.AddColumn(Qn, "JahrQuartal", each Text.From([Jahr]) & "-" & [Quartal], type text),
    Mo = Table.AddColumn(JQ, "Monat", each Date.Month([Auftragsdatum]), Int64.Type),
    Sa = Table.AddColumn(Mo, "S", each fnSaison([Auftragsdatum])),
    ES = Table.ExpandRecordColumn(Sa, "S", {"Saison", "Saisonjahr"}),
    PRt = Table.Buffer(DIM_PRODUKT),
    J4 = Table.NestedJoin(ES, {"ArtNr"}, PRt, {"ArtNr"}, "P", JoinKind.LeftOuter),
    E4 = Table.ExpandTableColumn(J4, "P", {"ProduktKey", "Produktgruppe"}),
    Sel = Table.SelectColumns(E4, {"KundeKey", "Kundentyp", "Stamm_VertreterNr", "LandISO", "Region", "PLZ", "VerkaeuferKey", "Kanal", "Quelle", "AuftragID", "Pos", "Auftragsdatum", "Jahr", "Quartal", "JahrQuartal", "Monat", "Saison", "Saisonjahr", "ProduktKey", "Produktgruppe", "Status", "Lieferart", "MengeStk", "NettoEUR", "Lieferdauer_Tage", "Zahlungsverzug_Tage"}),
    Typ = Table.TransformColumnTypes(Sel, {{"Saison", type text}, {"Saisonjahr", type text}, {"LandISO", type text}, {"Region", type text}, {"PLZ", type text}, {"Stamm_VertreterNr", type text}})
in
    Typ;

shared CHECK_Fakten = let
    Quelle = let F = Table.Buffer(FAKT_VERKAUF), FE = Table.Buffer(FAKT_ERP), FS = Table.Buffer(FAKT_SHOP), VK = Table.Buffer(DIM_VERKAEUFER), dupAuftr = Table.RowCount(Table.SelectRows(Table.Group(Table.Distinct(Table.SelectColumns(F, {"AuftragID", "Quelle"})), {"AuftragID"}, {{"n", Table.RowCount}}), each [n] > 1)), Storno = Table.RowCount(Table.SelectRows(orders, each [status] = "cancelled" or [status] = "refunded")), Importiert = List.Count(List.Distinct(List.RemoveNulls(AUFTRAG[ShopRef]))), Doppelt = Table.RowCount(Table.SelectRows(AUFTRAG, each [ShopRef] <> null)) - Importiert, Waisen = Table.RowCount(Table.NestedJoin(AUFTRAGSPOS, {"AuftragNr"}, AUFTRAG, {"AuftragNr"}, "A", JoinKind.LeftAnti)), VorAuftrag = Table.RowCount(Table.SelectRows(Table.NestedJoin(LIEFERUNG, {"AuftragNr"}, AUFTRAG, {"AuftragNr"}, "A", JoinKind.Inner), each [Versanddatum] < [A]{0}[Datum])), Pruef = #table({"Pruefung", "Soll", "Ist"}, {{"Positionen ohne Auftrag (verworfen)", 4, Waisen}, {"Shop-Auftraege doppelt im ERP (verworfen)", 3, Doppelt}, {"Versand vor Auftrag (Lieferdauer null)", 4, VorAuftrag}, {"Zeilen FAKT_ERP", 1463, Table.RowCount(FE)}, {"Auftraege ERP", 475, Table.RowCount(Table.Distinct(Table.SelectColumns(FE, {"AuftragID"})))}, {"Bestellungen Shop gesamt", 549, Table.RowCount(orders)}, {"Shop storniert/erstattet (gefiltert)", null, Storno}, {"Shop im ERP importiert (Anti-Join)", 131, Importiert}, {"Bestellungen in FAKT_SHOP", 549 - Storno - Importiert, Table.RowCount(Table.Distinct(Table.SelectColumns(FS, {"AuftragID"})))}, {"AuftragID in beiden Quellen", 0, dupAuftr}, {"KundeKey null", 0, Table.RowCount(Table.SelectRows(F, each [KundeKey] = null))}, {"ProduktKey null", 0, Table.RowCount(Table.SelectRows(F, each [ProduktKey] = null))}, {"Region null", 0, Table.RowCount(Table.SelectRows(F, each [Region] = null))}, {"Verkaeufer unbekannt", 0, Table.RowCount(Table.SelectRows(F, each not List.Contains(VK[VerkaeuferKey], [VerkaeuferKey])))}, {"Zeilen FAKT_VERKAUF", Table.RowCount(FE) + Table.RowCount(FS), Table.RowCount(F)}, {"Summe NettoEUR ERP", null, Number.Round(List.Sum(FE[NettoEUR]), 2)}, {"Summe NettoEUR Shop", null, Number.Round(List.Sum(FS[NettoEUR]), 2)}}), OK = Table.AddColumn(Pruef, "OK", each if [Soll] = null then null else [Soll] = [Ist], type logical) in OK
in
    Quelle;

shared CHECK_Pruefliste = // CHECK_Pruefliste: Restbefunde nach der Bereinigung (entspricht der Tag-Spalte der PA1). Eine Zeile je Datensatz und Befund.
let
    E = Table.SelectColumns(Table.AddColumn(KUNDE_h, "ID", each [KdNr], type text), {"Quelle", "ID", "Name_norm", "Befunde"}),
    C = Table.SelectColumns(Table.AddColumn(CRM_h, "ID", each [AccountId], type text), {"Quelle", "ID", "Name_norm", "Befunde"}),
    S = Table.SelectColumns(Table.AddColumn(SHOP_h, "ID", each [ShopID], type text), {"Quelle", "ID", "Name_norm", "Befunde"}),
    // Befunde aus der Zuordnung: Fremdschlüssel ohne Bestätigung
    CK = Table.SelectColumns(Table.AddColumn(Table.AddColumn(Table.SelectRows(CRM_k, each [Flag_Schluessel_Name] = true), "ID", each [AccountId], type text), "Befunde2", each "Fremdschlüssel ERP_KdNr zeigt auf anderen Kunden", type text), {"Quelle", "ID", "Name_norm", "Befunde2"}),
    SK = Table.SelectColumns(Table.AddColumn(Table.AddColumn(Table.SelectRows(SHOP_k, each [Flag_Schluessel_Name] = true), "ID", each [ShopID], type text), "Befunde2", each "Fremdschlüssel ERP_KdNr zeigt auf anderen Kunden", type text), {"Quelle", "ID", "Name_norm", "Befunde2"}),
    // Befunde über mehrere Kunden: dieselbe E-Mail oder dieselbe Adresse bei verschiedenen KundeKeys
    ST = Table.Buffer(Table.SelectColumns(Table.Combine({
            Table.AddColumn(KUNDE_k, "ID", each [KdNr], type text),
            Table.AddColumn(CRM_k, "ID", each [AccountId], type text),
            Table.AddColumn(SHOP_k, "ID", each [ShopID], type text)}), {"Quelle", "ID", "Name_norm", "KundeKey", "Email_norm", "Matchkey"}, MissingField.UseNull)),
    GM = Table.SelectRows(Table.Group(Table.SelectRows(ST, each [Email_norm] <> null), {"Email_norm"}, {{"n", each List.Count(List.Distinct([KundeKey])), Int64.Type}}), each [n] > 1),
    GA = Table.SelectRows(Table.Group(Table.SelectRows(ST, each [Matchkey] <> null), {"Matchkey"}, {{"n", each List.Count(List.Distinct([KundeKey])), Int64.Type}}), each [n] > 1),
    XM = Table.SelectColumns(Table.AddColumn(Table.SelectRows(ST, each [Email_norm] <> null and List.Contains(GM[Email_norm], [Email_norm])), "Befunde2", each "E-Mail bei mehreren Kunden (Sammelpostfach?)", type text), {"Quelle", "ID", "Name_norm", "Befunde2"}),
    XA = Table.SelectColumns(Table.AddColumn(Table.SelectRows(ST, each [Matchkey] <> null and List.Contains(GA[Matchkey], [Matchkey])), "Befunde2", each "Adresse bei mehreren Kunden (Filiale, Vereinsheim?)", type text), {"Quelle", "ID", "Name_norm", "Befunde2"}),
    Lang1 = Table.ExpandListColumn(Table.AddColumn(Table.SelectRows(Table.Combine({E, C, S}), each [Befunde] <> null and [Befunde] <> ""), "Befund", each Text.Split([Befunde], "; ")), "Befund"),
    Lang2 = Table.RenameColumns(Table.Combine({CK, SK, XM, XA}), {{"Befunde2", "Befund"}}),
    Alle = Table.Combine({Table.SelectColumns(Lang1, {"Quelle", "ID", "Name_norm", "Befund"}), Lang2}),
    Typ = Table.TransformColumnTypes(Alle, {{"Befund", type text}}),
    Sort = Table.Sort(Typ, {{"Befund", Order.Ascending}, {"Quelle", Order.Ascending}, {"ID", Order.Ascending}})
in
    Sort;

shared CHECK_Fehlerlog = // CHECK_Fehlerlog: zweite Erfolgsmessung. Je eingebautem Fehlertyp: wie viele Fälle gibt es, wie viele davon stehen auf der Prüfliste.
// Automatisch behobene Fehler sollen NICHT auf der Prüfliste stehen, manuelle sollen vollständig dort stehen.
let
    Roh = Csv.Document(File.Contents(pPfad & "sollwerte\fehler_log.csv"), [Delimiter = ";", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Log = Table.TransformColumnTypes(Table.PromoteHeaders(Roh, [PromoteAllScalars = true]), {{"Quelle", type text}, {"Tabelle", type text}, {"ID", type text}, {"Fehlertyp", type text}}),
    Kunden = Table.SelectRows(Log, each [Tabelle] = "KUNDE" or [Tabelle] = "customers" or [Tabelle] = "Accounts"),
    Typ1 = Table.AddColumn(Kunden, "Typ", each if Text.StartsWith([Fehlertyp], "Dublette im ERP") then "Dublette im ERP" else [Fehlertyp], type text),
    PL = Table.Buffer(Table.Group(CHECK_Pruefliste, {"Quelle", "ID"}, {{"Befunde", each Text.Combine(List.Distinct([Befund]), "; "), type text}})),
    J = Table.NestedJoin(Typ1, {"Quelle", "ID"}, PL, {"Quelle", "ID"}, "P", JoinKind.LeftOuter),
    JE = Table.ExpandTableColumn(J, "P", {"Befunde"}),
    G = Table.Group(JE, {"Typ"}, {{"Faelle", each Table.RowCount(_), Int64.Type}, {"Auf_Pruefliste", each List.Count(List.RemoveNulls([Befunde])), Int64.Type}, {"Haeufigster_Befund", each List.First(List.RemoveNulls([Befunde]), null), type text}}),
    Sort = Table.Sort(G, {{"Faelle", Order.Descending}})
in
    Sort;
