section Section1;

shared KUNDE = let
    Quelle = Csv.Document(File.Contents("C:\Users\SteffenWaldvogel\Documents\DHBW\BI-Seminararbeit\quellen\erp\KUNDE.txt"),[Delimiter=";", Columns=16, Encoding=1252, QuoteStyle=QuoteStyle.Csv]),
    #"Höher gestufte Header" = Table.PromoteHeaders(Quelle, [PromoteAllScalars=true]),
    #"Ersetzter Wert" = Table.ReplaceValue(#"Höher gestufte Header","#(cr)#(lf)","|",Replacer.ReplaceText,{"Bemerkung"}),
    #"Gefilterte Zeilen1" = Table.SelectRows(#"Ersetzter Wert", each true),
    #"Geänderter Typ mit Gebietsschema" = Table.TransformColumnTypes(#"Gefilterte Zeilen1", {{"Anlagedatum", type date}}, "de-DE"),
    #"Geänderter Typ mit Gebietsschema1" = Table.TransformColumnTypes(#"Geänderter Typ mit Gebietsschema", {{"GesperrtBis", type date}}, "de-DE"),
    Zeilenvorschub_LF_ersetzt = Table.ReplaceValue(#"Geänderter Typ mit Gebietsschema1", "#(lf)", " | ", Replacer.ReplaceText, {"Bemerkung"})
in
    Zeilenvorschub_LF_ersetzt;

shared customers = let
    Quelle = Csv.Document(File.Contents("C:\Users\SteffenWaldvogel\Documents\DHBW\BI-Seminararbeit\quellen\shop\customers.csv"),[Delimiter=",", Columns=12, Encoding=65001, QuoteStyle=QuoteStyle.None]),
    #"Höher gestufte Header" = Table.PromoteHeaders(Quelle, [PromoteAllScalars=true]),
    #"Geänderter Typ mit Gebietsschema" = Table.TransformColumnTypes(#"Höher gestufte Header", {{"created_at", type datetimezone}}, "de-DE")
in
    #"Geänderter Typ mit Gebietsschema";

shared Accounts = let
    Quelle = Excel.Workbook(File.Contents("C:\Users\SteffenWaldvogel\Documents\DHBW\BI-Seminararbeit\quellen\crm\CRM_Export.xlsx"), null, true),
    Accounts_Sheet = Quelle{[Item="Accounts",Kind="Sheet"]}[Data],
    #"Höher gestufte Header" = Table.PromoteHeaders(Accounts_Sheet, [PromoteAllScalars=true]),
    Flag_Datum_ungenau = Table.AddColumn(#"Höher gestufte Header", "Flag_Datum_ungenau", each Text.Length(Text.From([Erstkontakt])) = 7, type logical),
    #"Geänderter Typ mit Gebietsschema" = Table.TransformColumnTypes(Flag_Datum_ungenau, {{"Erstkontakt", type date}}, "en-US"),
    #"Geänderter Typ mit Gebietsschema1" = Table.TransformColumnTypes(#"Geänderter Typ mit Gebietsschema", {{"LetzteAenderung", type date}}, "en-US"),
    #"Geänderter Typ" = Table.TransformColumnTypes(#"Geänderter Typ mit Gebietsschema1",{{"PLZ", type text}, {"ERP_KdNr", type text}})
in
    #"Geänderter Typ";

shared fnClean = (s as nullable text) as nullable text =>
let
    t0 = if s = null then null else Text.Clean(s),
    t1 = if t0 = null then null else Text.Replace(t0, Character.FromNumber(160), " "),
    t2 = if t1 = null then null else Text.Remove(t1, {Character.FromNumber(173), Character.FromNumber(8203)}),
    t3 = if t2 = null then null else Text.Combine(List.Select(Text.Split(t2, " "), each _ <> ""), " "),
    t4 = if t3 = "" then null else t3
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
    stop = {" e.v.", " e. v.", " gmbh", " e.k.", " ohg", "tsv ", "sv ", "dc ", "sc ", "dr. "},
    t2 = if t1 = null then null else List.Accumulate(stop, t1, (acc, w) => Text.Replace(acc, w, " ")),
    t3 = if t2 = null then null else Text.Replace(Text.Trim(t2), " ", "")
in
    t3;

shared mapLand = #table(type table [Quelle = text, Wert = text, LandISO = text], {
    {"ERP", "D",  "DE"}, {"ERP", "A", "AT"}, {"ERP", "CH", "CH"},
    {"CRM", "Deutschland", "DE"}, {"CRM", "Österreich", "AT"}, {"CRM", "Schweiz", "CH"}
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

shared KUNDE_h = let
    Quelle = KUNDE,
    fnClean = Table.TransformColumns(Table.TransformColumnTypes(Quelle, {{"Anlagedatum", type text}, {"GesperrtBis", type text}}, "de-DE"),{{"KdNr", Text.Trim, type text}, {"Firma", Text.Trim, type text}, {"Name1", Text.Trim, type text}, {"Vorname", Text.Trim, type text}, {"Strasse", Text.Trim, type text}, {"PLZ", Text.Trim, type text}, {"Ort", Text.Trim, type text}, {"Land", Text.Trim, type text}, {"Kundengruppe", Text.Trim, type text}, {"VertreterNr", Text.Trim, type text}, {"Zahlungsziel", Text.Trim, type text}, {"Anlagedatum", Text.Trim, type text}, {"Newsletter", Text.Trim, type text}, {"Gesperrt", Text.Trim, type text}, {"GesperrtBis", Text.Trim, type text}, {"Bemerkung", Text.Trim, type text}}),
    #"Geänderter Typ" = Table.TransformColumnTypes(fnClean,{{"KdNr", type text}}),
    #"Zusammengeführte Abfragen" = Table.NestedJoin(#"Geänderter Typ", {"Land"}, mapLand, {"Wert"}, "mapLand", JoinKind.LeftOuter),
    #"Erweiterte mapLand" = Table.ExpandTableColumn(#"Zusammengeführte Abfragen", "mapLand", {"LandISO"}, {"mapLand.LandISO"}),
    #"Entfernte Spalten" = Table.RemoveColumns(#"Erweiterte mapLand",{"Land"}),
    #"Umbenannte Spalten" = Table.RenameColumns(#"Entfernte Spalten",{{"mapLand.LandISO", "LandISO"}}),
    #"Hinzugefügte benutzerdefinierte Spalte" = Table.AddColumn(#"Umbenannte Spalten", "PLZ_norm", each if [LandISO] = "DE" then Text.PadStart([PLZ], 5, "0") else [PLZ]),
    #"Neu angeordnete Spalten" = Table.ReorderColumns(#"Hinzugefügte benutzerdefinierte Spalte",{"KdNr", "Firma", "Name1", "Vorname", "Strasse", "PLZ", "PLZ_norm", "Ort", "Kundengruppe", "VertreterNr", "Zahlungsziel", "Anlagedatum", "Newsletter", "Gesperrt", "GesperrtBis", "Bemerkung", "LandISO"}),
    #"Spalte nach Trennzeichen teilen" = Table.SplitColumn(#"Neu angeordnete Spalten", "Strasse", Splitter.SplitTextByEachDelimiter({" "}, QuoteStyle.Csv, true), {"Strasse.1", "Strasse.2"}),
    #"Geänderter Typ1" = Table.TransformColumnTypes(#"Spalte nach Trennzeichen teilen",{{"Strasse.1", type text}, {"Strasse.2", type text}}),
    #"Umbenannte Spalten1" = Table.RenameColumns(#"Geänderter Typ1",{{"Strasse.2", "Hausnummer"}, {"Strasse.1", "Strasse"}}),
    #"Umbenannte Spalten2" = Table.RenameColumns(#"Umbenannte Spalten1",{{"Hausnummer", "Hausnr"}}),
    #"Hinzugefügte benutzerdefinierte Spalte1" = Table.AddColumn(#"Umbenannte Spalten2", "Strasse_ok", each if Text.Select(Text.Start([Strasse], 1), {"0".."9"}) <> "" and Text.Select(Text.Start([Hausnr], 1), {"0".."9"}) = "" then [Hausnr] else [Strasse]),
    #"Hinzugefügte benutzerdefinierte Spalte2" = Table.AddColumn(#"Hinzugefügte benutzerdefinierte Spalte1", "Hausnr_ok", each if Text.Select(Text.Start([Strasse], 1), {"0".."9"}) <> "" and Text.Select(Text.Start([Hausnr], 1), {"0".."9"}) = "" then [Strasse] else [Hausnr]),
    #"Gefilterte Zeilen" = Table.SelectRows(#"Hinzugefügte benutzerdefinierte Spalte2", each true),
    #"Hinzugefügte bedingte Spalte" = Table.AddColumn(#"Gefilterte Zeilen", "Titel", each if Text.StartsWith([Name1], "Dr.") then "Dr." else null),
    #"Ersetzter Wert" = Table.ReplaceValue(#"Hinzugefügte bedingte Spalte","Dr. ","",Replacer.ReplaceText,{"Name1"}),
    #"Geänderter Typ2" = Table.TransformColumnTypes(#"Ersetzter Wert",{{"Kundengruppe", type text}}),
    #"Zusammengeführte Abfragen1" = Table.NestedJoin(#"Geänderter Typ2", {"Kundengruppe"}, mapKundentyp, {"Wert"}, "mapKundentyp", JoinKind.LeftOuter),
    #"Erweiterte mapKundentyp" = Table.ExpandTableColumn(#"Zusammengeführte Abfragen1", "mapKundentyp", {"Kundentyp"}, {"mapKundentyp.Kundentyp"}),
    #"Entfernte Spalten1" = Table.RemoveColumns(#"Erweiterte mapKundentyp",{"Kundengruppe"}),
    #"Umbenannte Spalten3" = Table.RenameColumns(#"Entfernte Spalten1",{{"mapKundentyp.Kundentyp", "Kundentyp"}}),
    #"Ersetzter Wert1" = Table.ReplaceValue(#"Umbenannte Spalten3","J","true",Replacer.ReplaceText,{"Newsletter"}),
    #"Ersetzter Wert2" = Table.ReplaceValue(#"Ersetzter Wert1","N","false",Replacer.ReplaceText,{"Newsletter"}),
    #"Ersetzter Wert3" = Table.ReplaceValue(#"Ersetzter Wert2","J","true",Replacer.ReplaceText,{"Gesperrt"}),
    #"Ersetzter Wert4" = Table.ReplaceValue(#"Ersetzter Wert3","N","false",Replacer.ReplaceText,{"Gesperrt"}),
    #"Ersetzter Wert5" = Table.ReplaceValue(#"Ersetzter Wert4","-1",null,Replacer.ReplaceValue,{"Zahlungsziel"}),
    #"Ersetzter Wert6" = Table.ReplaceValue(#"Ersetzter Wert5","31.12.9999",null,Replacer.ReplaceValue,{"GesperrtBis"}),
    #"Entfernte Spalten2" = Table.RemoveColumns(#"Ersetzter Wert6",{"Strasse", "Hausnr"}),
    #"Umbenannte Spalten4" = Table.RenameColumns(#"Entfernte Spalten2",{{"Hausnr_ok", "Hausnr"}, {"Strasse_ok", "Strasse"}}),
    #"Hinzugefügte benutzerdefinierte Spalte3" = Table.AddColumn(#"Umbenannte Spalten4", "Matchkey", each [LandISO] & [PLZ_norm] & "|" & fnNormStrasse([Strasse]) & "|" & Text.Lower([Hausnr])),
    #"Hinzugefügte benutzerdefinierte Spalte4" = Table.AddColumn(#"Hinzugefügte benutzerdefinierte Spalte3", "Name_norm", each fnNormName(if [Firma] <> null then [Firma] else [Nachname])),
    #"Geänderter Typ3" = Table.TransformColumnTypes(#"Hinzugefügte benutzerdefinierte Spalte4",{{"Name_norm", type text}}),
    #"Hinzugefügte benutzerdefinierte Spalte5" = Table.AddColumn(#"Geänderter Typ3", "Quelle", each "ERP"),
    #"Andere entfernte Spalten" = Table.SelectColumns(#"Hinzugefügte benutzerdefinierte Spalte5",{"KdNr", "Firma", "Name1", "Vorname", "PLZ_norm", "Ort", "VertreterNr", "Zahlungsziel", "Anlagedatum", "Newsletter", "Gesperrt", "GesperrtBis", "Bemerkung", "LandISO", "Strasse", "Hausnr", "Titel", "Kundentyp", "Matchkey", "Name_norm", "Quelle"}),
    #"Geänderter Typ4" = Table.TransformColumnTypes(#"Andere entfernte Spalten",{{"Newsletter", type logical}, {"Gesperrt", type logical}}),
    Benutzerdefiniert1 = Table.ReplaceValue(Table.TransformColumnTypes(#"Geänderter Typ4", {{"Anlagedatum", type date}, {"GesperrtBis", type date}}, "de-DE"), #date(9999,12,31), null, Replacer.ReplaceValue, {"GesperrtBis"})
in
    Benutzerdefiniert1;

shared SHOP_h = let
    Quelle = customers,
    fnClean = Table.TransformColumns(Quelle,{{"email", fnClean, type text}, {"customer_id", fnClean, type text}, {"first_name", fnClean, type text}, {"last_name", fnClean, type text}, {"address", fnClean, type text}, {"country", fnClean, type text}, {"erp_customer_no", fnClean, type text}, {"company", fnClean, type text}}),
    #"Spalte nach Trennzeichen teilen" = Table.SplitColumn(fnClean, "address", Splitter.SplitTextByEachDelimiter({", "}, QuoteStyle.Csv, false), {"address.1", "address.2"}),
    #"Geänderter Typ" = Table.TransformColumnTypes(#"Spalte nach Trennzeichen teilen",{{"gender", Int64.Type}, {"address.1", type text}, {"address.2", type text}, {"is_club", type logical}, {"newsletter", type logical}}),
    #"Spalte nach Trennzeichen teilen1" = Table.SplitColumn(#"Geänderter Typ", "address.1", Splitter.SplitTextByEachDelimiter({" "}, QuoteStyle.Csv, true), {"address.1.1", "address.1.2"}),
    #"Geänderter Typ1" = Table.TransformColumnTypes(#"Spalte nach Trennzeichen teilen1",{{"address.1.1", type text}, {"address.1.2", type text}}),
    #"Spalte nach Trennzeichen teilen2" = Table.SplitColumn(#"Geänderter Typ1", "address.2", Splitter.SplitTextByEachDelimiter({" "}, QuoteStyle.Csv, false), {"address.2.1", "address.2.2"}),
    #"Umbenannte Spalten" = Table.RenameColumns(#"Spalte nach Trennzeichen teilen2",{{"address.1.1", "Strasse"}, {"address.1.2", "Hausnr"}, {"address.2.1", "PLZ"}, {"address.2.2", "Ort"}}),
    #"Hinzugefügte bedingte Spalte" = Table.AddColumn(#"Umbenannte Spalten", "Benutzerdefiniert", each if [is_club] = true then "Verein" else if [company] <> null then null else "Privat"),
    #"Geänderter Typ2" = Table.TransformColumnTypes(#"Hinzugefügte bedingte Spalte",{{"Benutzerdefiniert", type text}}),
    #"Emails normalisiern" = Table.AddColumn(#"Geänderter Typ2", "Email_norm", each fnNormEmail([email])),
    #"Anrede aus gender ableiten" = Table.AddColumn(#"Emails normalisiern", "Benutzerdefiniert.1", each if [gender] = 1 then "Herr" else if [gender] = 2 then "Frau" else null),
    #"Umbenannte Spalten1" = Table.RenameColumns(#"Anrede aus gender ableiten",{{"Benutzerdefiniert.1", "Anrede"}}),
    #"Umbenannte Spalten2" = Table.RenameColumns(#"Umbenannte Spalten1",{{"PLZ", "PLZ_norm"}}),
    #"Hinzugefügte benutzerdefinierte Spalte" = Table.AddColumn(#"Umbenannte Spalten2", "Matchkey", each [country] & [PLZ_norm] & "|" & fnNormStrasse([Strasse]) & "|" & Text.Lower([Hausnr])),
    #"Hinzugefügte benutzerdefinierte Spalte1" = Table.AddColumn(#"Hinzugefügte benutzerdefinierte Spalte", "Name_norm", each fnNormName(if [company] <> null then [company] else [last_name])),
    #"Umbenannte Spalten3" = Table.RenameColumns(#"Hinzugefügte benutzerdefinierte Spalte1",{{"customer_id", "ShopID"}, {"erp_customer_no", "ERP_KdNr"}, {"company", "Firma"}, {"first_name", "Vorname"}, {"last_name", "Nachname"}, {"country", "LandISO"}, {"newsletter", "Newsletter"}, {"created_at", "Anlagedatum"}}),
    #"Hinzugefügte benutzerdefinierte Spalte2" = Table.AddColumn(#"Umbenannte Spalten3", "Quelle", each "Shop"),
    #"Andere entfernte Spalten" = Table.SelectColumns(#"Hinzugefügte benutzerdefinierte Spalte2",{"ShopID", "Vorname", "Nachname", "Strasse", "Hausnr", "PLZ_norm", "Ort", "LandISO", "is_club", "Newsletter", "Anlagedatum", "ERP_KdNr", "Firma", "Benutzerdefiniert", "Email_norm", "Anrede", "Matchkey", "Name_norm", "Quelle"}),
    Kundentyp_benannt_Typen_gesetzt = Table.TransformColumnTypes(Table.RenameColumns(#"Andere entfernte Spalten", {{"Benutzerdefiniert", "Kundentyp"}}), {{"Kundentyp", type text}, {"Anrede", type text}, {"Email_norm", type text}, {"Matchkey", type text}, {"Name_norm", type text}, {"Quelle", type text}, {"Anlagedatum", type date}})
in
    Kundentyp_benannt_Typen_gesetzt;

shared CRM_h = let
    Quelle = Accounts,
    #"Abgeschnittener Text" = Table.TransformColumns(Quelle,{{"Land", fnClean, type text}, {"Firma", fnClean, type text}, {"Anrede", fnClean, type text}, {"Titel", fnClean, type text}, {"Vorname", fnClean, type text}, {"Nachname", fnClean, type text}, {"Strasse", fnClean, type text}, {"Hausnr", fnClean, type text}, {"PLZ", fnClean, type text}, {"Ort", fnClean, type text}, {"Kundentyp", fnClean, type text}, {"ERP_KdNr", fnClean, type text}, {"Status", fnClean, type text}, {"E-Mail", fnClean, type text}}),
    #"Zusammengeführte Abfragen" = Table.NestedJoin(#"Abgeschnittener Text", {"Land"}, mapLand, {"Wert"}, "mapLand", JoinKind.LeftOuter),
    #"Erweiterte mapLand" = Table.ExpandTableColumn(#"Zusammengeführte Abfragen", "mapLand", {"LandISO"}, {"mapLand.LandISO"}),
    #"Entfernte Spalten" = Table.RemoveColumns(#"Erweiterte mapLand",{"Land"}),
    #"Umbenannte Spalten" = Table.RenameColumns(#"Entfernte Spalten",{{"mapLand.LandISO", "LandISO"}}),
    #"Hinzugefügte benutzerdefinierte Spalte" = Table.AddColumn(#"Umbenannte Spalten", "PLZ_norm", each if [LandISO] = "DE" then Text.PadStart([PLZ], 5, "0") else [PLZ]),
    #"Geänderter Typ" = Table.TransformColumnTypes(#"Hinzugefügte benutzerdefinierte Spalte",{{"PLZ_norm", type text}}),
    #"Umbenannte Spalten1" = Table.RenameColumns(#"Geänderter Typ",{{"Kundentyp", "Kundentyp_CRM"}}),
    #"Zusammengeführte Abfragen1" = Table.NestedJoin(#"Umbenannte Spalten1", {"Kundentyp_CRM"}, mapKundentyp, {"Wert"}, "mapKundentyp", JoinKind.LeftOuter),
    #"Erweiterte mapKundentyp" = Table.ExpandTableColumn(#"Zusammengeführte Abfragen1", "mapKundentyp", {"Kundentyp"}, {"mapKundentyp.Kundentyp"}),
    #"Umbenannte Spalten2" = Table.RenameColumns(#"Erweiterte mapKundentyp",{{"mapKundentyp.Kundentyp", "Kundentyp"}}),
    #"Entfernte Spalten1" = Table.RemoveColumns(#"Umbenannte Spalten2",{"Kundentyp_CRM"}),
    #"Ersetzter Wert" = Table.ReplaceValue(#"Entfernte Spalten1","X","true",Replacer.ReplaceText,{"Newsletter"}),
    #"Ersetzter Wert1" = Table.ReplaceValue(#"Ersetzter Wert","","false",Replacer.ReplaceValue,{"Newsletter"}),
    #"Ersetzter Wert2" = Table.ReplaceValue(#"Ersetzter Wert1",null,"false",Replacer.ReplaceValue,{"Newsletter"}),
    #"Geänderter Typ1" = Table.TransformColumnTypes(#"Ersetzter Wert2",{{"Newsletter", type logical}}),
    #"Umbenannte Spalten3" = Table.RenameColumns(#"Geänderter Typ1",{{"Status", "Kundenstatus"}, {"Segment", "ABC_Klasse"}, {"E-Mail", "Email"}}),
    #"Hinzugefügte benutzerdefinierte Spalte1" = Table.AddColumn(#"Umbenannte Spalten3", "Email_norm", each fnNormEmail([Email])),
    #"Hinzugefügte benutzerdefinierte Spalte2" = Table.AddColumn(#"Hinzugefügte benutzerdefinierte Spalte1", "Name_norm", each fnNormName(if [Firma] <> null then [Firma] else [Nachname])),
    #"Hinzugefügte benutzerdefinierte Spalte3" = Table.AddColumn(#"Hinzugefügte benutzerdefinierte Spalte2", "Quelle", each "CRM"),
    #"Hinzugefügte benutzerdefinierte Spalte4" = Table.AddColumn(#"Hinzugefügte benutzerdefinierte Spalte3", "Matchkey", each [LandISO] & [PLZ_norm] & "|" & fnNormStrasse([Strasse]) & "|" & Text.Lower([Hausnr])),
    #"Geänderter Typ2" = Table.TransformColumnTypes(#"Hinzugefügte benutzerdefinierte Spalte4",{{"Matchkey", type text}}),
    #"Andere entfernte Spalten" = Table.SelectColumns(#"Geänderter Typ2",{"AccountId", "Firma", "Anrede", "Titel", "Vorname", "Nachname", "Strasse", "Hausnr", "PLZ", "Ort", "ABC_Klasse", "Betreuer", "ERP_KdNr", "Erstkontakt", "Newsletter", "Kundenstatus", "LetzteAenderung", "Telefon", "LandISO", "PLZ_norm", "Kundentyp", "Email_norm", "Name_norm", "Quelle", "Matchkey"}),
    #"Umbenannte Spalten4" = Table.RenameColumns(#"Andere entfernte Spalten",{{"Kundenstatus", "Kontaktstatus"}}),
    Typen_Text_gesetzt = Table.TransformColumnTypes(#"Umbenannte Spalten4", {{"Email_norm", type text}, {"Matchkey", type text}, {"Name_norm", type text}, {"Quelle", type text}, {"PLZ_norm", type text}}),
    #"Hinzugefügte benutzerdefinierte Spalte5" = Table.AddColumn(Typen_Text_gesetzt, "Anrede_norm", each let a = Text.Lower(Text.Start([Anrede] ?? "", 2)) in
if a = "he" or a = "hr" then "Herr" else if a = "fr" then "Frau" else null),
    #"Entfernte Spalten2" = Table.RemoveColumns(#"Hinzugefügte benutzerdefinierte Spalte5",{"Anrede"}),
    #"Hinzugefügte benutzerdefinierte Spalte6" = Table.AddColumn(#"Entfernte Spalten2", "LandISO_norm", each if [LandISO] <> null then [LandISO]
else if Text.Length([PLZ_norm]) = 5 then "DE"
else null),
    #"Entfernte Spalten3" = Table.RemoveColumns(#"Hinzugefügte benutzerdefinierte Spalte6",{"LandISO"}),
    #"Umbenannte Spalten5" = Table.RenameColumns(#"Entfernte Spalten3",{{"LandISO_norm", "LandISO"}}),
    Anrede_Typen_PLZ_Plausibilitaet = Table.AddColumn(Table.TransformColumnTypes(Table.RenameColumns(#"Umbenannte Spalten5", {{"Anrede_norm", "Anrede"}}), {{"Anrede", type text}, {"LandISO", type text}}), "Flag_PLZ_Land", each ([LandISO] = "DE" and Text.Length([PLZ_norm]) <> 5) or ([LandISO] <> "DE" and [LandISO] <> null and Text.Length([PLZ_norm]) <> 4), type logical)
in
    Anrede_Typen_PLZ_Plausibilitaet;

shared KUNDE_k = let
    Quelle = KUNDE_h,
    #"Sortierte Zeilen" = Table.Sort(Quelle,{{"Matchkey", Order.Ascending}, {"Name_norm", Order.Ascending}, {"Anlagedatum", Order.Ascending}}),
    #"Entfernte Duplikate" = Table.Distinct(#"Sortierte Zeilen", {"Name_norm", "Matchkey"}),
    #"Hinzugefügte benutzerdefinierte Spalte" = Table.AddColumn(#"Entfernte Duplikate", "Matchmethode", each "0 Basis ERP"),
    KundeKey_vergeben = Table.TransformColumnTypes(Table.AddIndexColumn(#"Hinzugefügte benutzerdefinierte Spalte", "KundeKey", 1, 1, Int64.Type), {{"Matchmethode", type text}})
in
    KundeKey_vergeben;

shared ERP_Dubletten = // ERP_Dubletten (Zeile 47): Kunden, die im ERP mehrfach angelegt sind – gleicher Matchkey + normalisierter Name.
// Führend ist der älteste Datensatz (Anlagedatum); die anderen KdNr werden später auf denselben KundeKey gemappt.
let
    Q = KUNDE_h,
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

shared CRM_k = // CRM_k: jedem CRM-Account einen KundeKey zuordnen. Kaskade: 1 ERP-Schlüssel (mit Namensprüfung) -> 3 Adresse exakt -> 4 Fuzzy Name -> neu.
// Stufe 2 (E-Mail) entfällt, weil das ERP keine E-Mail führt (Checkliste 17).
let
    S = CRM_h,
    ERP = KUNDE_k,
    // ---- Stufe 1: ERP-Schlüssel (Checkliste 21)
    J1 = Table.NestedJoin(S, {"ERP_KdNr"}, ERP, {"KdNr"}, "K", JoinKind.LeftOuter),
    E1 = Table.ExpandTableColumn(J1, "K", {"KundeKey", "Name_norm"}, {"Key_S", "Name_ERP"}),
    // Plausibilität (30): Schlüssel nur akzeptieren, wenn der normalisierte Name auf beiden Seiten gleich ist
    P  = Table.AddColumn(E1, "Key_Schluessel", each if [Key_S] <> null and [Name_ERP] = [Name_norm] then [Key_S] else null, Int64.Type),
    F  = Table.AddColumn(P, "Flag_Schluessel_Name", each [Key_S] <> null and [Name_ERP] <> [Name_norm], type logical),
    // ---- Stufe 3: Adresse exakt über Matchkey (27)
    ERPadr = Table.Distinct(Table.SelectRows(Table.SelectColumns(ERP, {"Matchkey", "KundeKey"}), each [Matchkey] <> null), {"Matchkey"}),
    J3 = Table.NestedJoin(F, {"Matchkey"}, ERPadr, {"Matchkey"}, "A", JoinKind.LeftOuter),
    E3 = Table.ExpandTableColumn(J3, "A", {"KundeKey"}, {"Key_Adr"}),
    Vor = Table.AddColumn(E3, "Key_13", each if [Key_Schluessel] <> null then [Key_Schluessel] else [Key_Adr], Int64.Type),
    // ---- Stufe 4: Fuzzy auf normalisierten Namen + Land, nur für noch offene Zeilen
    Offen = Table.SelectRows(Vor, each [Key_13] = null),
    Zu = Table.SelectRows(Vor, each [Key_13] <> null),
    ERPn = Table.Distinct(Table.SelectColumns(ERP, {"Name_norm", "LandISO", "PLZ_norm", "Ort", "Kundentyp", "KundeKey"}), {"Name_norm", "LandISO", "PLZ_norm"}),
    FJ = Table.FuzzyNestedJoin(Offen, {"Name_norm", "LandISO"}, ERPn, {"Name_norm", "LandISO"}, "FZ", JoinKind.LeftOuter, [Threshold = pSchwelle, IgnoreCase = true, IgnoreSpace = true]),
    FE = Table.RemoveColumns(Table.AddColumn(FJ, "Key_Fuzzy", each let k = Table.SelectRows([FZ], (r) => if [Kundentyp] = "Verein" then r[Name_norm] = [Name_norm] else (r[PLZ_norm] = [PLZ_norm] or (r[Name_norm] = [Name_norm] and Text.StartsWith(Text.Lower([Ort]), Text.Lower(r[Ort]))))) in if Table.IsEmpty(k) then null else k{0}[KundeKey]), {"FZ"}),
    OffenK = Table.RemoveColumns(Table.AddColumn(FE, "KundeKey", each [Key_Fuzzy], Int64.Type), {"Key_Fuzzy"}),
    ZuK = Table.AddColumn(Zu, "KundeKey", each [Key_13], Int64.Type),
    Beide = Table.Combine({ZuK, OffenK}),
    M = Table.AddColumn(Beide, "Matchmethode", each
          if [Key_Schluessel] <> null then "1 ERP-Schlüssel"
          else if [Key_Adr] <> null then "3 Adresse exakt"
          else if [KundeKey] <> null then "4 Fuzzy Name"
          else "neu", type text),
    // ---- Neue Keys ab 1001: Leads, Umzieher ohne Schlüssel, Rest. CRM-Dubletten teilen sich einen Key (29)
    Neu = Table.SelectRows(M, each [KundeKey] = null),
    Alt = Table.SelectRows(M, each [KundeKey] <> null),
    NeuG = Table.Group(Neu, {"Matchkey", "Name_norm"}, {{"Z", each _, type table}}),
    NeuI = Table.AddIndexColumn(NeuG, "NeuKey", 1001, 1, Int64.Type),
    NeuE = Table.ExpandTableColumn(NeuI, "Z", List.RemoveItems(Table.ColumnNames(Neu), {"Matchkey", "Name_norm", "KundeKey"})),
    NeuK = Table.RenameColumns(NeuE, {{"NeuKey", "KundeKey"}}),
    Fertig = Table.RemoveColumns(Table.Combine({Alt, NeuK}), {"Key_S", "Name_ERP", "Key_Adr", "Key_13"})
in
    Fertig;

shared SHOP_k = // SHOP_k: jedem Shop-Kunden einen KundeKey zuordnen. Kaskade: 1 ERP-Schlüssel -> 2 E-Mail (CRM) -> 3 Adresse (ERP, dann CRM) -> 4 Fuzzy (nur gewerblich) -> neu ab 2001.
let
    S = SHOP_h,
    ERP = KUNDE_k,
    CRM = CRM_k,
    // ---- Stufe 1: ERP-Schlüssel mit Namensprüfung (21, 30)
    J1 = Table.NestedJoin(S, {"ERP_KdNr"}, ERP, {"KdNr"}, "K", JoinKind.LeftOuter),
    E1 = Table.ExpandTableColumn(J1, "K", {"KundeKey", "Name_norm"}, {"Key_S", "Name_ERP"}),
    P  = Table.AddColumn(E1, "Key_Schluessel", each if [Key_S] <> null and [Name_ERP] = [Name_norm] then [Key_S] else null, Int64.Type),
    F  = Table.AddColumn(P, "Flag_Schluessel_Name", each [Key_S] <> null and [Name_ERP] <> [Name_norm], type logical),
    // ---- Stufe 2: E-Mail gegen CRM (das ERP hat keine E-Mail)
    CRMm = Table.Distinct(Table.SelectRows(Table.SelectColumns(CRM, {"Email_norm", "KundeKey"}), each [Email_norm] <> null and [Email_norm] <> ""), {"Email_norm"}),
    J2 = Table.NestedJoin(F, {"Email_norm"}, CRMm, {"Email_norm"}, "M", JoinKind.LeftOuter),
    E2 = Table.ExpandTableColumn(J2, "M", {"KundeKey"}, {"Key_Mail"}),
    // ---- Stufe 3: Adresse exakt, erst ERP, dann CRM (27)
    ERPadr = Table.Distinct(Table.SelectRows(Table.SelectColumns(ERP, {"Matchkey", "KundeKey"}), each [Matchkey] <> null), {"Matchkey"}),
    CRMadr = Table.Distinct(Table.SelectRows(Table.SelectColumns(CRM, {"Matchkey", "KundeKey"}), each [Matchkey] <> null), {"Matchkey"}),
    J3a = Table.NestedJoin(E2, {"Matchkey"}, ERPadr, {"Matchkey"}, "A", JoinKind.LeftOuter),
    E3a = Table.ExpandTableColumn(J3a, "A", {"KundeKey"}, {"Key_AdrERP"}),
    J3b = Table.NestedJoin(E3a, {"Matchkey"}, CRMadr, {"Matchkey"}, "B", JoinKind.LeftOuter),
    E3b = Table.ExpandTableColumn(J3b, "B", {"KundeKey"}, {"Key_AdrCRM"}),
    Vor = Table.AddColumn(E3b, "Key_123", each if [Key_Schluessel] <> null then [Key_Schluessel] else if [Key_Mail] <> null then [Key_Mail] else if [Key_AdrERP] <> null then [Key_AdrERP] else [Key_AdrCRM], Int64.Type),
    // ---- Stufe 4: Fuzzy nur für gewerbliche Kunden; Privatnamen (Müller) sind zu häufig. Treffer nur bei gleicher PLZ oder bei Vereinen.
    Offen = Table.SelectRows(Vor, each [Key_123] = null and [Kundentyp] <> "Privat"),
    Rest  = Table.SelectRows(Vor, each [Key_123] <> null or [Kundentyp] = "Privat"),
    Namen = Table.Distinct(Table.Combine({Table.SelectColumns(ERP, {"Name_norm", "LandISO", "PLZ_norm", "Ort", "Kundentyp", "KundeKey"}), Table.SelectColumns(CRM, {"Name_norm", "LandISO", "PLZ_norm", "Ort", "Kundentyp", "KundeKey"})}), {"Name_norm", "LandISO", "PLZ_norm"}),
    FJ = Table.FuzzyNestedJoin(Offen, {"Name_norm", "LandISO"}, Namen, {"Name_norm", "LandISO"}, "FZ", JoinKind.LeftOuter, [Threshold = pSchwelle, IgnoreCase = true, IgnoreSpace = true]),
    FE = Table.RemoveColumns(Table.AddColumn(FJ, "Key_Fuzzy", each let k = Table.SelectRows([FZ], (r) => if [Kundentyp] = "Verein" then r[Name_norm] = [Name_norm] else (r[PLZ_norm] = [PLZ_norm] or (r[Name_norm] = [Name_norm] and Text.StartsWith(Text.Lower([Ort]), Text.Lower(r[Ort]))))) in if Table.IsEmpty(k) then null else k{0}[KundeKey]), {"FZ"}),
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
    Fertig = Table.RemoveColumns(Table.Combine({Alt, NeuI}), {"Key_S", "Name_ERP", "Key_Mail", "Key_AdrERP", "Key_AdrCRM", "Key_123"})
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
    Roh = Csv.Document(File.Contents("C:\Users\SteffenWaldvogel\Documents\DHBW\BI-Seminararbeit\sollwerte\kunden_mapping.csv"), [Delimiter = ";", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
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
    Roh = Csv.Document(File.Contents("C:\Users\SteffenWaldvogel\Documents\DHBW\BI-Seminararbeit\sollwerte\kunden_mapping.csv"), [Delimiter = ";", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
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
    Quelle = let Spalten = {"KundeKey","Quelle","Firma","Titel","Vorname","Nachname","Anrede","Kundentyp","LandISO","PLZ_norm","Ort","Strasse","Hausnr","Matchkey","Newsletter","Anlagedatum","Aenderungsdatum","Kontaktstatus","VertreterNr","Betreuer","Email_norm","KdNr","ShopID","AccountId"}, E = Table.AddColumn(Table.AddColumn(KUNDE_k, "Prio", each 1), "Aenderungsdatum", each try [Anlagedatum] otherwise null, type date), C = Table.AddColumn(Table.AddColumn(Table.RenameColumns(CRM_k, {{"Erstkontakt","Anlagedatum"}}, MissingField.Ignore), "Prio", each 2), "Aenderungsdatum", each try [LetzteAenderung] otherwise null, type date), S = Table.AddColumn(Table.AddColumn(SHOP_k, "Prio", each 3), "Aenderungsdatum", each try [Anlagedatum] otherwise null, type date), Alle = Table.Combine({Table.SelectColumns(E, Spalten & {"Prio"}, MissingField.UseNull), Table.SelectColumns(C, Spalten & {"Prio"}, MissingField.UseNull), Table.SelectColumns(S, Spalten & {"Prio"}, MissingField.UseNull)}) in Alle
in
    Quelle;

shared DIM_KUNDE = let
    Quelle = let S = Table.Sort(STACK_KUNDE, {{"KundeKey", Order.Ascending}, {"Prio", Order.Ascending}}), erst = (l as list) => List.First(List.RemoveNulls(l)), G = Table.Group(S, {"KundeKey"}, {{"Firma", each erst([Firma]), type text}, {"Titel", each erst([Titel]), type text}, {"Vorname", each erst([Vorname]), type text}, {"Nachname", each erst([Nachname]), type text}, {"Anrede", each erst([Anrede]), type text}, {"Kundentyp", each erst([Kundentyp]) ?? "Privat", type text}, {"LandISO", each erst([LandISO]), type text}, {"Newsletter", each List.AnyTrue(List.RemoveNulls([Newsletter])), type logical}, {"Erstkontakt", each List.Min([Anlagedatum]), type date}, {"Kontaktstatus", each erst([Kontaktstatus]) ?? "Kunde", type text}, {"VertreterNr", each erst([VertreterNr]), type text}, {"Betreuer", each erst([Betreuer]), type text}, {"ERP_KdNr", each erst([KdNr]), type text}, {"ShopID", each erst([ShopID]), type text}, {"AccountId", each erst([AccountId]), type text}, {"Quellen", each Text.Combine(List.Distinct([Quelle]), "+"), type text}, {"AnzahlQuellen", each List.Count(List.Distinct([Quelle])), Int64.Type}}) in G
in
    Quelle;

shared DIM_KUNDE_ADRESSE = let
    Quelle = let A = Table.SelectRows(STACK_KUNDE, each [Matchkey] <> null), G = Table.Group(A, {"KundeKey","Matchkey"}, {{"GueltigVonQuelle", each List.Min([Aenderungsdatum]), type date}, {"Strasse", each List.First(List.RemoveNulls([Strasse])), type text}, {"Hausnr", each List.First(List.RemoveNulls([Hausnr])), type text}, {"PLZ", each List.First(List.RemoveNulls([PLZ_norm])), type text}, {"Ort", each List.First(List.RemoveNulls([Ort])), type text}, {"LandISO", each List.First(List.RemoveNulls([LandISO])), type text}}), S = Table.Sort(G, {{"KundeKey", Order.Ascending}, {"GueltigVonQuelle", Order.Ascending}}), V = Table.Group(S, {"KundeKey"}, {{"T", each Table.AddIndexColumn(_, "Version", 1, 1, Int64.Type), type table}}), VE = Table.ExpandTableColumn(V, "T", {"Matchkey","GueltigVonQuelle","Strasse","Hausnr","PLZ","Ort","LandISO","Version"}), Von = Table.AddColumn(VE, "GueltigVon", each if [Version] = 1 then #date(1900,1,1) else [GueltigVonQuelle], type date), Next = Table.TransformColumns(Table.SelectColumns(Von, {"KundeKey","Version","GueltigVon"}), {{"Version", each _ - 1, Int64.Type}}), J = Table.NestedJoin(Von, {"KundeKey","Version"}, Next, {"KundeKey","Version"}, "N", JoinKind.LeftOuter), JE = Table.ExpandTableColumn(J, "N", {"GueltigVon"}, {"NaechsteVon"}), Bis = Table.AddColumn(JE, "GueltigBis", each if [NaechsteVon] = null then null else Date.AddDays([NaechsteVon], -1), type date), Cur = Table.AddColumn(Bis, "CurrentFlag", each [GueltigBis] = null, type logical), Pr = Table.AddColumn(Cur, "Praefix", each if [LandISO] = "DE" then Text.Start([PLZ], 1) else "", type text), RJ = Table.NestedJoin(Pr, {"LandISO","Praefix"}, mapRegion, {"LandISO","Praefix"}, "R", JoinKind.LeftOuter), RE = Table.ExpandTableColumn(RJ, "R", {"Region"}), Fertig = Table.RemoveColumns(RE, {"NaechsteVon","Praefix"}) in Fertig
in
    Quelle;

shared AUFTRAG = let
    Quelle = let Quelle = Csv.Document(File.Contents("C:\Users\SteffenWaldvogel\Documents\DHBW\BI-Seminararbeit\quellen\erp\AUFTRAG.txt"), [Delimiter=";", Encoding=1252, QuoteStyle=QuoteStyle.Csv]), Header = Table.PromoteHeaders(Quelle, [PromoteAllScalars=true]), Typen = Table.TransformColumnTypes(Header, {{"AuftragNr", type text}, {"KdNr", Int64.Type}, {"Datum", type date}, {"VertreterNr", type text}, {"Status", type text}, {"Waehrung", type text}, {"Lieferart", type text}, {"Zahlungsziel", Int64.Type}, {"ShopRef", type text}}, "de-DE"), LeerZuNull = Table.ReplaceValue(Typen, "", null, Replacer.ReplaceValue, {"VertreterNr", "ShopRef"}) in LeerZuNull
in
    Quelle;

shared AUFTRAGSPOS = let
    Quelle = let Quelle = Csv.Document(File.Contents("C:\Users\SteffenWaldvogel\Documents\DHBW\BI-Seminararbeit\quellen\erp\AUFTRAGSPOS.txt"), [Delimiter=";", Encoding=1252, QuoteStyle=QuoteStyle.Csv]), Header = Table.PromoteHeaders(Quelle, [PromoteAllScalars=true]), Typen = Table.TransformColumnTypes(Header, {{"AuftragNr", type text}, {"Pos", Int64.Type}, {"ArtNr", type text}, {"Menge", Int64.Type}, {"Einheit", type text}, {"EPreisNetto", type number}}, "de-DE") in Typen
in
    Quelle;

shared ZAHLUNG = let
    Quelle = let Quelle = Csv.Document(File.Contents("C:\Users\SteffenWaldvogel\Documents\DHBW\BI-Seminararbeit\quellen\erp\ZAHLUNG.txt"), [Delimiter=";", Encoding=1252, QuoteStyle=QuoteStyle.Csv]), Header = Table.PromoteHeaders(Quelle, [PromoteAllScalars=true]), Typen = Table.TransformColumnTypes(Header, {{"ZahlungNr", type text}, {"AuftragNr", type text}, {"Datum", type date}, {"Betrag", type number}, {"Zahlart", type text}}, "de-DE") in Typen
in
    Quelle;

shared LIEFERUNG = let
    Quelle = let Quelle = Csv.Document(File.Contents("C:\Users\SteffenWaldvogel\Documents\DHBW\BI-Seminararbeit\quellen\erp\LIEFERUNG.txt"), [Delimiter=";", Encoding=1252, QuoteStyle=QuoteStyle.Csv]), Header = Table.PromoteHeaders(Quelle, [PromoteAllScalars=true]), Typen = Table.TransformColumnTypes(Header, {{"LieferNr", type text}, {"AuftragNr", type text}, {"Versanddatum", type date}, {"Lieferart", type text}}, "de-DE") in Typen
in
    Quelle;

shared VERTRETER = let
    Quelle = let Quelle = Csv.Document(File.Contents("C:\Users\SteffenWaldvogel\Documents\DHBW\BI-Seminararbeit\quellen\erp\VERTRETER.txt"), [Delimiter=";", Encoding=1252, QuoteStyle=QuoteStyle.Csv]), Header = Table.PromoteHeaders(Quelle, [PromoteAllScalars=true]), Typen = Table.TransformColumnTypes(Header, {{"VertreterNr", type text}, {"Name", type text}, {"Gebiet", type text}, {"Austritt", type text}}) in Typen
in
    Quelle;

shared ARTIKEL = let
    Quelle = let Quelle = Csv.Document(File.Contents("C:\Users\SteffenWaldvogel\Documents\DHBW\BI-Seminararbeit\quellen\erp\ARTIKEL.txt"), [Delimiter=";", Encoding=1252, QuoteStyle=QuoteStyle.Csv]), Header = Table.PromoteHeaders(Quelle, [PromoteAllScalars=true]), Typen = Table.TransformColumnTypes(Header, {{"ArtNr", type text}, {"Bezeichnung", type text}, {"EKPreis", type number}}, "de-DE") in Typen
in
    Quelle;

shared Betreuer = let
    Quelle = let Quelle = Excel.Workbook(File.Contents("C:\Users\SteffenWaldvogel\Documents\DHBW\BI-Seminararbeit\quellen\crm\CRM_Export.xlsx"), null, true), Blatt = Quelle{[Item="Betreuer", Kind="Sheet"]}[Data], Header = Table.PromoteHeaders(Blatt, [PromoteAllScalars=true]), Typen = Table.TransformColumnTypes(Header, {{"Name", type text}, {"E-Mail", type text}, {"Rolle", type text}, {"Region", type text}}) in Typen
in
    Quelle;

shared orders = let
    Quelle = let Quelle = Csv.Document(File.Contents("C:\Users\SteffenWaldvogel\Documents\DHBW\BI-Seminararbeit\quellen\shop\orders.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]), Header = Table.PromoteHeaders(Quelle, [PromoteAllScalars=true]), Typen = Table.TransformColumnTypes(Header, {{"order_id", type text}, {"customer_id", type text}, {"order_date", type datetimezone}, {"status", type text}, {"total_gross", type number}, {"shipping_method_id", Int64.Type}, {"coupon_code", type text}}, "en-US"), LeerZuNull = Table.ReplaceValue(Typen, "", null, Replacer.ReplaceValue, {"coupon_code"}) in LeerZuNull
in
    Quelle;

shared order_items = let
    Quelle = let Quelle = Csv.Document(File.Contents("C:\Users\SteffenWaldvogel\Documents\DHBW\BI-Seminararbeit\quellen\shop\order_items.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]), Header = Table.PromoteHeaders(Quelle, [PromoteAllScalars=true]), Typen = Table.TransformColumnTypes(Header, {{"order_id", type text}, {"line_no", Int64.Type}, {"sku", type text}, {"qty", Int64.Type}, {"unit_price_gross", type number}}, "en-US") in Typen
in
    Quelle;

shared products = let
    Quelle = let Quelle = Csv.Document(File.Contents("C:\Users\SteffenWaldvogel\Documents\DHBW\BI-Seminararbeit\quellen\shop\products.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]), Header = Table.PromoteHeaders(Quelle, [PromoteAllScalars=true]), Typen = Table.TransformColumnTypes(Header, {{"sku", type text}, {"name", type text}, {"category", type text}}) in Typen
in
    Quelle;

shared shipping_methods = let
    Quelle = let Quelle = Csv.Document(File.Contents("C:\Users\SteffenWaldvogel\Documents\DHBW\BI-Seminararbeit\quellen\shop\shipping_methods.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]), Header = Table.PromoteHeaders(Quelle, [PromoteAllScalars=true]), Typen = Table.TransformColumnTypes(Header, {{"id", Int64.Type}, {"name", type text}}) in Typen
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
    Quelle = let P = AUFTRAGSPOS, J1 = Table.NestedJoin(P, {"AuftragNr"}, AUFTRAG, {"AuftragNr"}, "A", JoinKind.LeftOuter), E1 = Table.ExpandTableColumn(J1, "A", {"KdNr", "Datum", "VertreterNr", "Status", "Waehrung", "Lieferart", "Zahlungsziel", "ShopRef"}), L = Table.Group(LIEFERUNG, {"AuftragNr"}, {{"Versanddatum_letzt", each List.Max([Versanddatum]), type date}, {"AnzLieferungen", each Table.RowCount(_), Int64.Type}}), Z = Table.Group(ZAHLUNG, {"AuftragNr"}, {{"Zahldatum_letzt", each List.Max([Datum]), type date}, {"Zahlbetrag", each List.Sum([Betrag]), type number}, {"AnzZahlungen", each Table.RowCount(_), Int64.Type}}), J2 = Table.NestedJoin(E1, {"AuftragNr"}, L, {"AuftragNr"}, "L", JoinKind.LeftOuter), E2 = Table.ExpandTableColumn(J2, "L", {"Versanddatum_letzt"}), J3 = Table.NestedJoin(E2, {"AuftragNr"}, Z, {"AuftragNr"}, "Z", JoinKind.LeftOuter), E3 = Table.ExpandTableColumn(J3, "Z", {"Zahldatum_letzt", "Zahlbetrag"}), KH = Table.AddColumn(Table.SelectColumns(KUNDE_h, {"KdNr", "VertreterNr"}), "KdNr_txt", each Text.From([KdNr]), type text), Q = Table.AddColumn(E3, "QuellID", each Text.From([KdNr]), type text), J4 = Table.NestedJoin(Q, {"QuellID"}, KH, {"KdNr_txt"}, "K", JoinKind.LeftOuter), E4 = Table.ExpandTableColumn(J4, "K", {"VertreterNr"}, {"Stamm_VertreterNr"}), M = Table.AddColumn(E4, "MengeStk", each if [Einheit] = "Set" then [Menge] * 10 else [Menge], Int64.Type), Ku = Table.AddColumn(M, "Kurs", each if [Waehrung] = "CHF" then pKursCHF else 1, type number), N = Table.AddColumn(Ku, "NettoEUR", each Number.Round([MengeStk] * [EPreisNetto] * [Kurs], 2), type number), Ka = Table.AddColumn(N, "Kanal", each if [ShopRef] <> null then "Online" else "Direkt", type text), Vk = Table.AddColumn(Ka, "VerkaeuferKey", each if [VertreterNr] <> null then [VertreterNr] else if [ShopRef] <> null then "ONLINE" else [Stamm_VertreterNr], type text), Ld = Table.AddColumn(Vk, "Lieferdauer_Tage", each if [Versanddatum_letzt] = null then null else Duration.Days([Versanddatum_letzt] - [Datum]), Int64.Type), Zv = Table.AddColumn(Ld, "Zahlungsverzug_Tage", each if [Zahldatum_letzt] = null then null else Duration.Days([Zahldatum_letzt] - Date.AddDays([Datum], [Zahlungsziel])), Int64.Type), Qu = Table.AddColumn(Zv, "Quelle", each "ERP", type text), R = Table.RenameColumns(Qu, {{"AuftragNr", "AuftragID"}, {"Datum", "Auftragsdatum"}, {"ShopRef", "Fremdbeleg"}}), Sel = Table.SelectColumns(R, {"Quelle", "QuellID", "AuftragID", "Pos", "Auftragsdatum", "VerkaeuferKey", "Kanal", "Status", "Lieferart", "ArtNr", "MengeStk", "NettoEUR", "Lieferdauer_Tage", "Zahlungsverzug_Tage", "Fremdbeleg"}) in Sel
in
    Quelle;

shared FAKT_SHOP = let
    Quelle = let I = order_items, J1 = Table.NestedJoin(I, {"order_id"}, orders, {"order_id"}, "O", JoinKind.LeftOuter), E1 = Table.ExpandTableColumn(J1, "O", {"customer_id", "order_date", "status", "shipping_method_id"}), F = Table.SelectRows(E1, each [status] = "shipped" or [status] = "paid"), A = Table.NestedJoin(F, {"order_id"}, AUFTRAG, {"ShopRef"}, "ERP", JoinKind.LeftAnti), A2 = Table.RemoveColumns(A, {"ERP"}), J2 = Table.NestedJoin(A2, {"customer_id"}, SHOP_h, {"ShopID"}, "K", JoinKind.LeftOuter), E2 = Table.ExpandTableColumn(J2, "K", {"LandISO"}), J3 = Table.NestedJoin(E2, {"LandISO"}, mapMwSt, {"LandISO"}, "M", JoinKind.LeftOuter), E3 = Table.ExpandTableColumn(J3, "M", {"MwStFaktor"}), J4 = Table.NestedJoin(E3, {"shipping_method_id"}, shipping_methods, {"id"}, "V", JoinKind.LeftOuter), E4 = Table.ExpandTableColumn(J4, "V", {"name"}, {"Versandart"}), La = Table.AddColumn(E4, "Lieferart", each if [Versandart] = "Abholung" then "Abholung" else "Paket", type text), N = Table.AddColumn(La, "NettoEUR", each Number.Round([qty] * [unit_price_gross] / [MwStFaktor], 2), type number), D = Table.AddColumn(N, "Auftragsdatum", each Date.From([order_date]), type date), St = Table.AddColumn(D, "Status", each if [status] = "shipped" then "L" else "O", type text), J5 = Table.NestedJoin(St, {"sku"}, DIM_PRODUKT, {"Shop_sku"}, "P", JoinKind.LeftOuter), E5 = Table.ExpandTableColumn(J5, "P", {"ArtNr"}), C1 = Table.AddColumn(E5, "Quelle", each "Shop", type text), C2 = Table.AddColumn(C1, "VerkaeuferKey", each "ONLINE", type text), C3 = Table.AddColumn(C2, "Kanal", each "Online", type text), C4 = Table.AddColumn(C3, "Lieferdauer_Tage", each null, Int64.Type), C5 = Table.AddColumn(C4, "Zahlungsverzug_Tage", each null, Int64.Type), C6 = Table.AddColumn(C5, "Fremdbeleg", each null, type text), R = Table.RenameColumns(C6, {{"order_id", "AuftragID"}, {"line_no", "Pos"}, {"qty", "MengeStk"}, {"customer_id", "QuellID"}}), Sel = Table.SelectColumns(R, {"Quelle", "QuellID", "AuftragID", "Pos", "Auftragsdatum", "VerkaeuferKey", "Kanal", "Status", "Lieferart", "ArtNr", "MengeStk", "NettoEUR", "Lieferdauer_Tage", "Zahlungsverzug_Tage", "Fremdbeleg"}) in Sel
in
    Quelle;

shared FAKT_VERKAUF = let
    Quelle = let C = Table.Combine({FAKT_ERP, FAKT_SHOP}), MAPt = Table.Buffer(Table.TransformColumns(MAP_SCHLUESSEL, {{"KundeKey", each Text.From(_), type text}})), J1 = Table.NestedJoin(C, {"Quelle", "QuellID"}, MAPt, {"Quelle", "QuellID"}, "M", JoinKind.LeftOuter), E1 = Table.ExpandTableColumn(J1, "M", {"KundeKey"}), DKt = Table.Buffer(Table.TransformColumns(DIM_KUNDE, {{"KundeKey", each Text.From(_), type text}})), J2 = Table.NestedJoin(E1, {"KundeKey"}, DKt, {"KundeKey"}, "K", JoinKind.LeftOuter), E2 = Table.ExpandTableColumn(J2, "K", {"Kundentyp"}), DAt = Table.Buffer(Table.TransformColumns(DIM_KUNDE_ADRESSE, {{"KundeKey", each Text.From(_), type text}})), J3 = Table.NestedJoin(E2, {"KundeKey"}, DAt, {"KundeKey"}, "A", JoinKind.LeftOuter), Ad = Table.AddColumn(J3, "Adresse", each let d = [Auftragsdatum], t = Table.SelectRows([A], (r) => r[GueltigVon] <= d and (r[GueltigBis] = null or d <= r[GueltigBis])) in if Table.IsEmpty(t) then null else t{0}), E3 = Table.ExpandRecordColumn(Ad, "Adresse", {"LandISO", "Region", "PLZ"}, {"LandISO", "Region", "PLZ"}), R1 = Table.RemoveColumns(E3, {"A"}), Y = Table.AddColumn(R1, "Jahr", each Date.Year([Auftragsdatum]), Int64.Type), Qn = Table.AddColumn(Y, "Quartal", each "Q" & Text.From(Date.QuarterOfYear([Auftragsdatum])), type text), JQ = Table.AddColumn(Qn, "JahrQuartal", each Text.From([Jahr]) & "-" & [Quartal], type text), Mo = Table.AddColumn(JQ, "Monat", each Date.Month([Auftragsdatum]), Int64.Type), Sa = Table.AddColumn(Mo, "S", each fnSaison([Auftragsdatum])), ES = Table.ExpandRecordColumn(Sa, "S", {"Saison", "Saisonjahr"}), PRt = Table.Buffer(DIM_PRODUKT), J4 = Table.NestedJoin(ES, {"ArtNr"}, PRt, {"ArtNr"}, "P", JoinKind.LeftOuter), E4 = Table.ExpandTableColumn(J4, "P", {"ProduktKey", "Produktgruppe"}), Sel = Table.SelectColumns(E4, {"KundeKey", "Kundentyp", "LandISO", "Region", "PLZ", "VerkaeuferKey", "Kanal", "Quelle", "AuftragID", "Pos", "Auftragsdatum", "Jahr", "Quartal", "JahrQuartal", "Monat", "Saison", "Saisonjahr", "ProduktKey", "Produktgruppe", "Status", "Lieferart", "MengeStk", "NettoEUR", "Lieferdauer_Tage", "Zahlungsverzug_Tage"}), Typ = Table.TransformColumnTypes(Sel, {{"Saison", type text}, {"Saisonjahr", type text}, {"LandISO", type text}, {"Region", type text}, {"PLZ", type text}}) in Typ
in
    Quelle;

shared CHECK_Fakten = let
    Quelle = let F = Table.Buffer(FAKT_VERKAUF), FE = Table.Buffer(FAKT_ERP), FS = Table.Buffer(FAKT_SHOP), VK = Table.Buffer(DIM_VERKAEUFER), dupAuftr = Table.RowCount(Table.SelectRows(Table.Group(Table.Distinct(Table.SelectColumns(F, {"AuftragID", "Quelle"})), {"AuftragID"}, {{"n", Table.RowCount}}), each [n] > 1)), Storno = Table.RowCount(Table.SelectRows(orders, each [status] = "cancelled" or [status] = "refunded")), Importiert = Table.RowCount(Table.SelectRows(AUFTRAG, each [ShopRef] <> null)), Pruef = #table({"Pruefung", "Soll", "Ist"}, {{"Zeilen FAKT_ERP = AUFTRAGSPOS", Table.RowCount(AUFTRAGSPOS), Table.RowCount(FE)}, {"Auftraege ERP", 354, Table.RowCount(Table.Distinct(Table.SelectColumns(FE, {"AuftragID"})))}, {"Bestellungen Shop gesamt", 384, Table.RowCount(orders)}, {"Shop storniert/erstattet (gefiltert)", null, Storno}, {"Shop im ERP importiert (Anti-Join)", 78, Importiert}, {"Bestellungen in FAKT_SHOP", 384 - Storno - Importiert, Table.RowCount(Table.Distinct(Table.SelectColumns(FS, {"AuftragID"})))}, {"AuftragID in beiden Quellen", 0, dupAuftr}, {"KundeKey null", 0, Table.RowCount(Table.SelectRows(F, each [KundeKey] = null))}, {"ProduktKey null", 0, Table.RowCount(Table.SelectRows(F, each [ProduktKey] = null))}, {"Region null", 0, Table.RowCount(Table.SelectRows(F, each [Region] = null))}, {"Verkaeufer unbekannt", 0, Table.RowCount(Table.SelectRows(F, each not List.Contains(VK[VerkaeuferKey], [VerkaeuferKey])))}, {"Zeilen FAKT_VERKAUF", Table.RowCount(FE) + Table.RowCount(FS), Table.RowCount(F)}, {"Summe NettoEUR ERP", null, Number.Round(List.Sum(FE[NettoEUR]), 2)}, {"Summe NettoEUR Shop", null, Number.Round(List.Sum(FS[NettoEUR]), 2)}}), OK = Table.AddColumn(Pruef, "OK", each if [Soll] = null then null else [Soll] = [Ist], type logical) in OK
in
    Quelle;