// BELEG_Heterogenitaet: derselbe Kunde in den drei Quellen. Je Merkmal der Rohwert aus ERP, Shop und CRM und daneben das harmonisierte Ergebnis.
// Stimmen die drei harmonisierten Werte überein, steht rechts nur ein Wert; sonst stehen die Varianten nebeneinander.
let
    Kunden = {"6", "34"},
    t = (v) => if v = null then "(leer)" else if v is text then (if Text.Trim(v) = "" then "(leer)" else v) else if v is logical then (if v then "ja" else "nein") else if v is date then Date.ToText(v, "dd.MM.yyyy") else Text.From(v, "de-DE"),
    eins = (l as list) => let d = List.Distinct(List.Transform(List.RemoveNulls(l), t)) in if List.IsEmpty(d) then "(leer)" else Text.Combine(d, "  /  "),
    rec = (T as table, col as text, id as text) as nullable record => List.First(Table.ToRecords(Table.SelectRows(T, each Text.From(Record.Field(_, col)) = id)), null),
    ER = Table.Buffer(Table.PromoteHeaders(Csv.Document(File.Contents(pPfad & "quellen\erp\KUNDE.txt"), [Delimiter = ";", Columns = 17, Encoding = 1252, QuoteStyle = QuoteStyle.Csv]), [PromoteAllScalars = true])),
    SR = Table.Buffer(Table.PromoteHeaders(Csv.Document(File.Contents(pPfad & "quellen\shop\customers.csv"), [Delimiter = ",", Columns = 13, Encoding = 65001, QuoteStyle = QuoteStyle.Csv]), [PromoteAllScalars = true])),
    CR = Table.Buffer(Table.PromoteHeaders(Excel.Workbook(File.Contents(pPfad & "quellen\crm\CRM_Export.xlsx"), null, true){[Item = "Accounts", Kind = "Sheet"]}[Data], [PromoteAllScalars = true])),
    EH = Table.Buffer(KUNDE_h), SH = Table.Buffer(SHOP_h), CH = Table.Buffer(CRM_h),
    Map = Table.Buffer(Table.PromoteHeaders(Csv.Document(File.Contents(pPfad & "sollwerte\kunden_mapping.csv"), [Delimiter = ";", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]), [PromoteAllScalars = true])),
    block = (k as text) as list =>
        let
            m = rec(Map, "KundeKey", k),
            e = rec(ER, "KdNr", m[ERP_KdNr]), s = rec(SR, "customer_id", m[Shop_ID]), c = rec(CR, "AccountId", m[CRM_ID]),
            eh = rec(EH, "KdNr", m[ERP_KdNr]), sh = rec(SH, "ShopID", m[Shop_ID]), ch = rec(CH, "AccountId", m[CRM_ID]),
            adr = (h) => t(h[Strasse]) & " | " & t(h[Hausnr]),
            Z = {
                {"Schlüssel", "KdNr = " & t(e[KdNr]), "customer_id = " & t(s[customer_id]), "AccountId = " & Text.Start(t(c[AccountId]), 8) & "…", "KundeKey (Surrogatschlüssel)", "Kaskade und MAP_SCHLUESSEL"},
                {"Name", "Firma = " & t(e[Firma]), "company = " & t(s[company]), "Firma = " & t(c[Firma]), eins({eh[Name_norm], sh[Name_norm], ch[Name_norm]}), "fnNormName: klein, Umlaute, ohne Rechtsform und Vereinskürzel"},
                {"Straße, Hausnummer", "Strasse = " & t(e[Strasse]), "address = " & t(s[address]), "Strasse = " & t(c[Strasse]) & " | Hausnr = " & t(c[Hausnr]), eins({adr(eh), adr(sh), adr(ch)}), "fnStrasse: Straße und Hausnummer trennen"},
                {"PLZ, Ort", "PLZ = " & t(e[PLZ]) & " | Ort = " & t(e[Ort]), "(im Feld address)", "PLZ = " & t(c[PLZ]) & " | Ort = " & t(c[Ort]), eins({t(eh[PLZ_norm]) & " " & t(eh[Ort]), t(sh[PLZ_norm]) & " " & t(sh[Ort]), t(ch[PLZ_norm]) & " " & t(ch[Ort])}), "PLZ als Text, führende Null, Länderpräfix ab"},
                {"Land", "Land = " & t(e[Land]), "country = " & t(s[country]), "Land = " & t(c[Land]), eins({eh[LandISO], sh[LandISO], ch[LandISO]}), "mapLand: Mapping-Tabelle auf ISO-Code"},
                {"Adressschlüssel", "–", "–", "–", eins({eh[Matchkey], sh[Matchkey], ch[Matchkey]}), "fnMatchkey: Land + PLZ + Straße normalisiert + erste Hausnummer"},
                {"Kundentyp", "Kundengruppe = " & t(e[Kundengruppe]), "is_club = " & t(s[is_club]), "Kundentyp = " & t(c[Kundentyp]), eins({eh[Kundentyp], sh[Kundentyp], ch[Kundentyp]}), "mapKundentyp; Shop kennt nur Verein ja/nein"},
                {"Newsletter", "Newsletter = " & t(e[Newsletter]), "newsletter = " & t(s[newsletter]), "Newsletter = " & t(c[Newsletter]), eins({eh[Newsletter], sh[Newsletter], ch[Newsletter]}), "fnJN: J/N, true/false, X/leer auf ja/nein"},
                {"Anlagedatum", "Anlagedatum = " & t(e[Anlagedatum]), "created_at = " & t(s[created_at]), "Erstkontakt = " & t(c[Erstkontakt]), eins({eh[Anlagedatum], sh[Anlagedatum], ch[Erstkontakt]}), "Gebietsschema de-DE, ISO mit UTC, en-US"},
                {"Ansprechpartner", "Name1 = " & t(e[Name1]) & " | Vorname = " & t(e[Vorname]), "first_name = " & t(s[first_name]) & " | last_name = " & t(s[last_name]), "Anrede = " & t(c[Anrede]) & " | Vorname = " & t(c[Vorname]) & " | Nachname = " & t(c[Nachname]), eins({t(eh[Vorname]) & " " & t(eh[Nachname]), t(sh[Vorname]) & " " & t(sh[Nachname]), t(ch[Vorname]) & " " & t(ch[Nachname])}), "Titel abgetrennt, Platzhalter zu leer"},
                {"E-Mail", "(kein Feld)", "email = " & t(s[email]), "E-Mail = " & t(c[#"E-Mail"]), eins({sh[Email_norm], ch[Email_norm]}), "fnEmails: klein, Transportmüll entfernt"},
                {"Telefon", "Telefon = " & t(e[Telefon]), "phone = " & t(s[phone]), "Telefon = " & t(c[Telefon]), eins({eh[Telefon_E164], sh[Telefon_E164], ch[Telefon_E164]}), "fnTel: E.164, Zusatztext abgeschnitten"}
            }
        in
            List.Transform(Z, (z) => {"Kunde " & k & " (" & t(m[Kundentyp]) & ", " & t(m[Land]) & ")"} & z),
    Alle = List.Combine(List.Transform(Kunden, block)),
    Tab = #table(type table [Kunde = text, Merkmal = text, #"ERP (KUNDE.txt)" = text, #"Shop (customers.csv)" = text, #"CRM (Accounts)" = text, Harmonisiert = text, Regel = text], Alle)
in
    Tab
