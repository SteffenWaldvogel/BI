"""
Generierung der Beispieldaten für die drei Quellsysteme der DartPoint GmbH
(Portfolioleistung Business Intelligence, Thema 1.2 – Power-Query-Integration).

Erzeugt:
  erp/   KUNDE, VERTRETER, AUFTRAG, AUFTRAGSPOS, ZAHLUNG, LIEFERUNG   (.txt, Windows-1252, ';', TT.MM.JJJJ, Dezimalkomma, CRLF)
  shop/  customers, orders, order_items, shipping_methods, products   (.csv, UTF-8, ',', ISO-Datum UTC, Dezimalpunkt, LF)
  crm/   CRM_Export.xlsx mit Blättern Accounts, Betreuer, Aktivitaet  (US-Datum MM/DD/YYYY, GUIDs)
  sollwerte/ kunden_mapping.csv – Ground Truth: welcher Kunde steckt unter welcher ID in welchem System,
             inkl. Art der eingebauten Abweichung. Nur zur Erfolgsmessung der Integration, NICHT Teil der Quellen.

Alle Konflikte sind absichtlich eingebaut (siehe Entwurf "Quellsysteme Dart-Händler", Kap. 6 Checkliste).
Deterministisch über SEED.
"""
import csv, json, os, random, uuid
from datetime import date, datetime, timedelta
from openpyxl import Workbook

SEED = 20261127
random.seed(SEED)
FEHLERQUOTE = 1.0   # Skalierung aller "manuellen" Fehler (1.0 = Standard, 2.0 = doppelt so viele, 0 = keine)
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "quellen")
FEHLER_LOG = []   # zweite Ground Truth: welcher eingebaute Fehler steckt in welchem Quelldatensatz
def log(quelle, tabelle, id_, typ): FEHLER_LOG.append((quelle, tabelle, str(id_), typ))

# ---------------------------------------------------------------- Stammdaten-Bausteine
VERTRETER = [  # VertreterNr, Nachname, Vorname, Rolle, ERP-Gebiet (PLZ-Präfixe), CRM-Region, Austritt
    ("V01", "Keller",   "Thomas",  "Außendienst", "70-79", "Süd",  "30.06.2024"),   # ausgeschieden, hat noch Aufträge
    ("V02", "Huber",    "Stefan",  "Außendienst", "80-89", "Süd",  "31.12.9999"),
    ("V03", "Schneider","Anja",    "Außendienst", "01-19", "Ost",  "31.12.9999"),
    ("V04", "Brunner",  "Reto",    "Außendienst", "CH",    "CH",   "31.12.9999"),
    ("V05", "Pichler",  "Lukas",   "Außendienst", "AT",    "AT",   "31.12.9999"),
    ("I01", "Vogt",     "Sabine",  "Innendienst", "",      "Innendienst", "31.12.9999"),
    ("I02", "Maier",    "Jonas",   "Innendienst", "",      "Innendienst", "31.12.9999"),
]
def vertreter_fuer(land, plz):
    if land == "CH": return "V04"
    if land == "AT": return "V05"
    p = int(plz[:2])
    if p < 20: return "V03"
    if p < 80: return "V01" if random.random() < 0.35 else "V02"
    return "V02"

ORTE = {  # (PLZ, Ort, Land) – DE 5-stellig, AT/CH 4-stellig
    "DE": [("88212","Ravensburg"),("88250","Weingarten"),("88045","Friedrichshafen"),("89073","Ulm"),("70173","Stuttgart"),
           ("72072","Tübingen"),("79098","Freiburg"),("80331","München"),("86150","Augsburg"),("90402","Nürnberg"),
           ("01067","Dresden"),("04109","Leipzig"),("10115","Berlin"),("18055","Rostock"),("06108","Halle (Saale)")],
    "AT": [("6900","Bregenz"),("6020","Innsbruck"),("5020","Salzburg"),("4020","Linz"),("1010","Wien"),("8010","Graz")],
    "CH": [("8001","Zürich"),("9000","St. Gallen"),("3011","Bern"),("4051","Basel"),("6003","Luzern"),("8400","Winterthur")],
}
STRASSEN = ["Goethestraße","Schillerstraße","Bahnhofstraße","Hauptstraße","Mozartstraße","Kirchplatz","Marktstraße",
            "Ringstraße","Lindenweg","Am Sportplatz","Gartenstraße","Bergstraße","Seestraße","Industriestraße","Zeppelinstraße"]
VORNAMEN_M = ["Max","Stefan","Jonas","Lukas","Felix","Tobias","Michael","Andreas","Daniel","Florian","Markus","Patrick"]
VORNAMEN_W = ["Anna","Lena","Julia","Sarah","Laura","Katharina","Sabine","Nina","Lisa","Jana","Claudia","Petra"]
NACHNAMEN  = ["Müller","Schmidt","Schneider","Fischer","Weber","Meyer","Wagner","Becker","Schulz","Hoffmann","Koch","Bauer",
              "Richter","Klein","Wolf","Schröder","Neumann","Schwarz","Zimmermann","Braun","Krüger","Hofmann","Lange","Werner",
              "Krause","Lehmann","Köhler","Herrmann","König","Jäger"]
VEREINE  = ["Dartfreunde","Bulls Eye","Triple 20","Red Dragons","Dartsport","Treffpunkt 180","Flight Club","Steeldarters","Dart Devils","Black Arrows"]
HAENDLER = ["Dart-Shop","Sport {n}","Dartwelt","Spielwaren {n}","Freizeit-Center","Sporthaus {n}","Dart Store"]
GASTRO   = ["Zum Bulls Eye","Gasthaus {n}","Sportsbar 180","Pub Triple Twenty","Kneipe zum Anker","Bistro {n}","Dart-Lounge"]

# ---------------------------------------------------------------- Produkte
PRODUKTE = []  # ERP ArtNr, Shop SKU, Name, Gruppe, Netto-EK-Preis
_gr = {"FL":"Flights","SH":"Shafts","BA":"Barrels","BO":"Boards","ZU":"Zubehör"}
_names = {"FL":["Standard Flights","Slim Flights","Kite Flights","Pear Flights","Turnier-Flights"],
          "SH":["Nylon Shafts","Alu Shafts","Carbon Shafts","Spinning Shafts"],
          "BA":["Steel Barrel 22g","Steel Barrel 24g","Tungsten Barrel 90%","Soft Barrel 18g","Brass Barrel 20g"],
          "BO":["Sisal Board Pro","Sisal Board Club","E-Dartboard","Board-Surround"],
          "ZU":["Flight Protector","Dart Case","Spitzen Steel","Spitzen Soft","Abwurflinie","Scoreboard","Wurfmatte"]}
_preis = {"FL":(1.9,3.5),"SH":(2.5,6.0),"BA":(18,65),"BO":(29,180),"ZU":(1.5,25)}
_sku = 10400
for g, names in _names.items():
    for i, n in enumerate(names, 1):
        var = "STD" if i == 1 else "PRO" if i == 2 else "VAR"
        _sku += 7
        PRODUKTE.append((f"{g}-{var}-{str(_sku)[1:]}", str(_sku), n, _gr[g], round(random.uniform(*_preis[g]), 2)))
SKU_SETS = {p[0] for p in PRODUKTE if p[3] in ("Flights","Shafts")}  # Vereine kaufen diese im 10er-Set

# ---------------------------------------------------------------- Kunden erzeugen
def land_zufall():
    return random.choices(["DE","AT","CH"], [70,15,15])[0]

def neuer_kunde(key, typ, praesenz):
    land = land_zufall()
    plz, ort = random.choice(ORTE[land])
    strasse = random.choice(STRASSEN); hnr = str(random.randint(1, 120)) + random.choice(["","","","a","b"])
    geschlecht = random.choice(["m","w"])
    vorname = random.choice(VORNAMEN_M if geschlecht=="m" else VORNAMEN_W)
    nachname = random.choice(NACHNAMEN)
    titel = "Dr." if random.random() < 0.08 else ""
    if typ == "Verein":
        v = random.choice(VEREINE); firma = f"{random.choice(['SV','TSV','DC','SC',''])} {v} {ort}".strip() + random.choice([" e.V."," e.V."," e. V.",""])
    elif typ == "Händler":
        firma = random.choice(HAENDLER).format(n=nachname) + random.choice([" GmbH"," GmbH"," e.K."," OHG",""])
    elif typ == "Gastro":
        firma = random.choice(GASTRO).format(n=nachname)
    else:
        firma = ""
    erst = date(2019,1,1) + timedelta(days=random.randint(0, 5*365))
    return dict(key=key, typ=typ, praesenz=praesenz, land=land, plz=plz, ort=ort, strasse=strasse, hnr=hnr,
                geschlecht=geschlecht, vorname=vorname, nachname=nachname, titel=titel, firma=firma,
                erst=erst, email=f"{vorname}.{nachname}".lower().replace("ü","ue").replace("ö","oe").replace("ä","ae").replace("ß","ss") + f"{random.randint(1,99)}@{random.choice(['gmail.com','web.de','gmx.de','outlook.com','t-online.de'])}",
                newsletter=random.random() < 0.55, abweichung="identisch", umzug=None)

VERTEILUNG = [  # (Präsenz, Anzahl, Typverteilung)
    ("ERP+SHOP+CRM", 64, ["Verein"]*5 + ["Händler"]*3 + ["Gastro"]*2),
    ("ERP+CRM",      28, ["Verein"]*4 + ["Händler"]*4 + ["Gastro"]*2),
    ("SHOP+CRM",     43, ["Verein"]*6 + ["Privat"]*3 + ["Gastro"]*1),
    ("CRM",          21, ["Verein"]*5 + ["Händler"]*3 + ["Gastro"]*2),       # Leads
    ("SHOP",         93, ["Privat"]*10),
    ("ERP",          36, ["Verein"]*3 + ["Händler"]*4 + ["Gastro"]*3),       # Altkunden, gesperrt
]
KUNDEN = []
k = 0
_namen_je_ort = set()   # (normalisierter Name, Ort) muss eindeutig sein – zwei gleichnamige Vereine/Kneipen in derselben Stadt sind unrealistisch
def _nkey(c):
    base = (c["firma"] or c["nachname"]).lower()
    for w in ["tsv ","sv ","dc ","sc "," e.v."," e. v."," gmbh"," e.k."," ohg"]: base = base.replace(w, " ")
    return (base.replace(" ", ""), c["ort"])
for praesenz, n, typen in VERTEILUNG:
    for _ in range(n):
        k += 1
        typ = random.choice(typen)
        for _versuch in range(50):
            c = neuer_kunde(k, typ, praesenz)
            if _nkey(c) not in _namen_je_ort: break
        _namen_je_ort.add(_nkey(c))
        KUNDEN.append(c)

# Abweichungen bei überlappenden Kunden: 40 % identisch, 40 % Schreibweise, 20 % Umzug (CRM neuer)
ueberlappend = [c for c in KUNDEN if "+" in c["praesenz"]]
random.shuffle(ueberlappend)
n_ul = len(ueberlappend)
for i, c in enumerate(ueberlappend):
    if i < int(n_ul*0.4): c["abweichung"] = "identisch"
    elif i < int(n_ul*0.8): c["abweichung"] = "schreibweise"
    else:
        c["abweichung"] = "umzug"
        plz, ort = random.choice(ORTE[c["land"]])
        c["umzug"] = dict(plz=plz, ort=ort, strasse=random.choice(STRASSEN), hnr=str(random.randint(1,120)),
                          datum=date(2024,1,1)+timedelta(days=random.randint(0,300)))

# ---------------------------------------------------------------- Schreibweisen-Varianten
def strasse_erp(c):   # ERP: Abkürzung "str.", Straße+Hausnr in einem Feld; 3 % vertauscht ("12 Goethestr.")
    s = c["strasse"].replace("straße","str.").replace("Straße","Str.")
    return f"{c['hnr']} {s}" if random.random() < 0.03 else f"{s} {c['hnr']}"
def strasse_shop(c):  # Shop: alles in einem Feld, Bindestrich-Schreibweise bei Abweichung
    s = c["strasse"]
    if c["abweichung"] == "schreibweise":
        s = s.replace("straße","-Straße") if "straße" in s else s
        s = s.replace("ß","ss") if random.random() < 0.3 else s
    return f"{s} {c['hnr']}, {c['plz']} {c['ort']}"
def strasse_crm(c):   # CRM: getrennt, Variante "-Str."
    src = c["umzug"] if c["umzug"] else c
    s = src["strasse"]
    if c["abweichung"] == "schreibweise":
        s = s.replace("straße","-Str.") if random.random() < 0.6 else s.replace("straße","strasse")
    return s, src["hnr"]
def name_crm(c):
    n = c["nachname"]
    if c["abweichung"] == "schreibweise" and random.random() < 0.5:
        n = n.replace("ü","ue").replace("ö","oe").replace("ä","ae")
    return n
def firma_crm(c):
    f = c["firma"]
    if c["typ"] == "Verein":
        for p in ["TSV ","SV ","DC ","SC "]: f = f.replace(p,"")
        f = f.replace(" e.V.","").replace(" e. V.","")
    return f

# Gleiche Adresse, verschiedene Kunden (PA1: Sammeladressen/Filialen): Verein hat sein Vereinsheim in der Gaststätte
VEREINSHEIM = []
_gastro = [c for c in KUNDEN if c["typ"]=="Gastro" and not c["umzug"]]
_frei = [c for c in KUNDEN if c["typ"]=="Verein" and not c["umzug"]]
for g in _gastro:
    kand = [v for v in _frei if v["ort"]==g["ort"]]
    if kand and len(VEREINSHEIM) < 8:
        v = kand[0]; _frei.remove(v)
        v["strasse"], v["hnr"], v["plz"] = g["strasse"], g["hnr"], g["plz"]
        VEREINSHEIM.append((g, v))
# Sammel-E-Mail (ServiceTool: gleiche E-Mail bei verschiedenen Kunden): zwei Kunden teilen sich ein Postfach
SAMMELMAIL = []
_mp = [c for c in KUNDEN if "SHOP" in c["praesenz"] and "CRM" in c["praesenz"] and c["typ"]!="Privat" and not c["umzug"]]
random.shuffle(_mp)
for a, b in zip(_mp[0:5], _mp[5:10]):
    b["email"] = a["email"]; SAMMELMAIL.append((a, b))

# ---------------------------------------------------------------- IDs vergeben
erp_nr = 10200; shop_nr = 100; crm_ids = {}
for c in KUNDEN:
    c["erp"]  = None; c["shop"] = None; c["crm"] = None
    if "ERP"  in c["praesenz"]: erp_nr += random.randint(1,4); c["erp"] = erp_nr
    if "SHOP" in c["praesenz"]: shop_nr += random.randint(1,3); c["shop"] = f"C-{shop_nr:05d}"
    if "CRM"  in c["praesenz"]: c["crm"] = str(uuid.UUID(int=random.getrandbits(128)))

for g, v in VEREINSHEIM:
    for q, t, kk in (("ERP","KUNDE","erp"),("Shop","customers","shop"),("CRM","Accounts","crm")):
        if v[kk]: log(q, t, v[kk], "gleiche Adresse wie anderer Kunde (Vereinsheim)")
for a, b in SAMMELMAIL:
    for x in (a, b):
        log("Shop","customers",x["shop"],"Sammel-E-Mail (zwei Kunden, ein Postfach)"); log("CRM","Accounts",x["crm"],"Sammel-E-Mail (zwei Kunden, ein Postfach)")

def d_de(d):  return d.strftime("%d.%m.%Y")
def d_us(d):  return d.strftime("%m/%d/%Y")
def d_iso(dt): return dt.strftime("%Y-%m-%dT%H:%M:%SZ")
def dez(x):   return f"{x:.2f}".replace(".", ",")

# ---------------------------------------------------------------- Aufträge
MWST = {"DE":1.19,"AT":1.20,"CH":1.081}
CHF_KURS = 0.96  # 1 CHF = 0,96 EUR (für die Netto-Umrechnung im Ziel – nicht in den Quellen enthalten!)
def saison_gewicht(m):  # Sep–Nov Ligastart stark, Dez Privat stark, Jul/Aug Sommerloch
    return {1:0.8,2:0.7,3:0.9,4:0.9,5:0.8,6:0.7,7:0.5,8:0.5,9:1.4,10:1.6,11:1.5,12:1.2}[m]

erp_auftraege, erp_pos, erp_zahlung, erp_lieferung = [], [], [], []
shop_orders, shop_items = [], []
auftrag_nr = {2023:0, 2024:0}; so_nr = 48000
tage = [date(2023,1,1)+timedelta(days=i) for i in range(730)]
gewichte = [saison_gewicht(t.month) for t in tage]

def positionen(c, n_pos):
    out = []
    for p in random.sample(PRODUKTE, n_pos):
        menge = random.randint(1,3) if c["typ"]=="Privat" else random.randint(2,12)
        einheit = "Stk"
        if c["typ"]=="Verein" and p[0] in SKU_SETS and random.random()<0.5:
            einheit, menge = "Set", random.randint(1,4)       # 1 Set = 10 Stück
        aufschlag = 1.6 if c["typ"]=="Privat" else 1.25 if c["typ"]=="Händler" else 1.4
        preis = round(p[4]*aufschlag, 2)
        out.append((p, menge, einheit, preis))
    return out

for c in KUNDEN:
    if c["praesenz"]=="CRM": continue                      # Leads: keine Aufträge
    gesperrt = c["praesenz"]=="ERP"
    n = random.randint(1,3) if gesperrt else random.randint(2,8) if c["typ"]!="Privat" else random.randint(1,4)
    for _ in range(n):
        t = random.choices(tage, gewichte)[0]
        if gesperrt and t > date(2023,9,30): t = date(2023, random.randint(1,9), random.randint(1,28))
        if t < c["erst"]: t = c["erst"] + timedelta(days=random.randint(5,60))
        if t > date(2024,12,31): continue
        online = c["shop"] is not None and (c["erp"] is None or random.random() < 0.45)
        pos = positionen(c, random.randint(1,5))
        if online:
            so_nr += random.randint(1,5); oid = f"SO-{so_nr}"
            status = random.choices(["shipped","paid","cancelled","refunded"],[80,10,6,4])[0]
            brutto = round(sum(m*(10 if e=="Set" else 1)*pr for _,m,e,pr in pos)*MWST[c["land"]], 2)
            shop_orders.append([oid, c["shop"], d_iso(datetime.combine(t, datetime.min.time())+timedelta(hours=random.randint(6,22),minutes=random.randint(0,59))),
                                status, f"{brutto:.2f}", random.choice([1,1,2,3]), random.choice(["","","","","LIGA24","XMAS23","CLUB10"])])
            for i,(p,m,e,pr) in enumerate(pos,1):
                shop_items.append([oid, i, p[1], m*(10 if e=="Set" else 1), f"{pr*MWST[c['land']]:.2f}"])
            # Shop-Auftrag wird ins ERP importiert (ohne Vertreter) – sofern der Kunde im ERP existiert und nicht storniert
            if c["erp"] and status in ("shipped","paid"):
                auftrag_nr[t.year]+=1; anr = f"{t.year}-{auftrag_nr[t.year]:05d}"
                erp_auftraege.append([anr, c["erp"], d_de(t), "", "L" if status=="shipped" else "O", "EUR",
                                      random.choice(["Paket","Paket","Spedition"]), 14, oid])
                for i,(p,m,e,pr) in enumerate(pos,1):
                    erp_pos.append([anr, i, p[0], m*(10 if e=="Set" else 1), "Stk", dez(pr)])
                if status=="shipped":
                    erp_lieferung.append([f"L{anr}", anr, d_de(t+timedelta(days=random.randint(1,4))), "Paket"])
                    erp_zahlung.append([f"Z{anr}", anr, d_de(t), dez(brutto/MWST[c["land"]]), random.choice(["PP","KK"])])
        else:
            auftrag_nr[t.year]+=1; anr = f"{t.year}-{auftrag_nr[t.year]:05d}"
            vnr = vertreter_fuer(c["land"], c["plz"]) if random.random()<0.8 else random.choice(["I01","I02"])
            waehrung = "CHF" if c["land"]=="CH" else "EUR"
            kurs = 1/CHF_KURS if waehrung=="CHF" else 1
            status = random.choices(["R","L","O"],[75,15,10])[0]
            ziel = random.choice([14,30,30,30,60])
            lieferart = random.choice(["Paket","Paket","Spedition","Abholung"])
            erp_auftraege.append([anr, c["erp"], d_de(t), vnr, status, waehrung, lieferart, ziel, ""])
            netto = 0
            for i,(p,m,e,pr) in enumerate(pos,1):
                erp_pos.append([anr, i, p[0], m, e, dez(pr*kurs)]); netto += m*(10 if e=="Set" else 1)*pr*kurs
            if status in ("L","R"):
                erp_lieferung.append([f"L{anr}", anr, d_de(t+timedelta(days=random.randint(1,12))), lieferart])
            if status=="R":
                verzug = random.choices([0,0,0,3,8,15,30],[40,20,10,10,10,6,4])[0]
                erp_zahlung.append([f"Z{anr}", anr, d_de(t+timedelta(days=ziel+verzug)), dez(netto), random.choice(["UE","UE","LS","BAR"])])

# ---------------------------------------------------------------- Fehler in den Bewegungsdaten
_gs = random.sample([x for x in erp_auftraege if not x[8] and x[4]=="R"], 5)          # Gutschrift: negative Menge, Status G
for a in _gs:
    for p in erp_pos:
        if p[0]==a[0]: p[3] = -abs(p[3])
    a[4] = "G"; log("ERP","AUFTRAG",a[0],"Gutschrift (negative Menge)")
_gsn = {a[0] for a in _gs}
erp_zahlung = [z for z in erp_zahlung if z[1] not in _gsn]; erp_lieferung = [l for l in erp_lieferung if l[1] not in _gsn]
for l in random.sample(erp_lieferung, 4):                                              # Versanddatum vor Auftragsdatum
    a = next(x for x in erp_auftraege if x[0]==l[1]); d = datetime.strptime(a[2], "%d.%m.%Y").date()
    l[2] = d_de(d - timedelta(days=random.randint(3,30))); log("ERP","LIEFERUNG",l[0],"Versand vor Auftrag")
for i in range(4):                                                                     # Waise: Position ohne Auftrag
    erp_pos.append([f"2023-9{i:04d}", 1, random.choice(PRODUKTE)[0], random.randint(1,5), "Stk", dez(9.9)])
    log("ERP","AUFTRAGSPOS",f"2023-9{i:04d}","Position ohne Auftrag (Waise)")
for a in random.sample([x for x in erp_auftraege if x[8]], 3):                         # Shop-Auftrag doppelt ins ERP importiert
    auftrag_nr[2024]+=1; anr = f"2024-{auftrag_nr[2024]:05d}"; dup = list(a); dup[0] = anr; erp_auftraege.append(dup)
    for p in [p for p in erp_pos if p[0]==a[0]]:
        q = list(p); q[0] = anr; erp_pos.append(q)
    log("ERP","AUFTRAG",anr,"Shop-Auftrag doppelt importiert")

# Exporte sind nach Schlüssel sortiert, wie ein echter Systemexport
erp_auftraege.sort(key=lambda r: r[0]); erp_pos.sort(key=lambda r: (r[0], r[1]))
erp_zahlung.sort(key=lambda r: r[1]); erp_lieferung.sort(key=lambda r: r[1])
shop_orders.sort(key=lambda r: r[0]); shop_items.sort(key=lambda r: (r[0], r[1]))

BEMERKUNGEN = ["", "", "", "", "", "Liefert nur Di/Do", "Rechnung an Zentrale", "Skonto 2 % vereinbart",
    "Anlieferung nur vormittags\r\nAnsprechpartner: Hr. Vogel\r\nTel. 0751/12345",        # eingebetteter CRLF
    "Turnierausstattung 2024\r\nSonderpreise Flights",
    "Achtung: neue Adresse ab 01/2024\nalte Rechnungen noch an alte Adresse",             # eingebetteter LF
    "Vereinsvorsitz wechselt jährlich;\r\nKontakt über Kassenwart"]                        # Semikolon + CRLF

# ---------------------------------------------------------------- Weitere Fehlertypen nach Skript 8.1.1.1
def f(p): return random.random() < p * FEHLERQUOTE
VORWAHL = {"DE": "751", "AT": "5574", "CH": "71"}
def telefon(land, stil=None):
    """Telefonnummer in einer von fünf Schreibweisen; E.164 ist das Ziel."""
    lk = {"DE":"49","AT":"43","CH":"41"}[land]; vw = VORWAHL[land]; nr = str(random.randint(20000, 99999))
    stil = stil or random.choice(["e164","national","klammer","intl00","punkt"])
    return {"e164": f"+{lk}{vw}{nr}", "national": f"0{vw}/{nr}", "klammer": f"(0{vw}) {nr}",
            "intl00": f"00{lk} {vw} {nr}", "punkt": f"0{vw}.{nr}"}[stil]
def mojibake(t):   # UTF-8-Bytes als Windows-1252 gelesen – klassischer Exportfehler
    return t.replace("ü","Ã¼").replace("ö","Ã¶").replace("ä","Ã¤").replace("ß","ÃŸ").replace("Ü","Ãœ").replace("Ö","Ã–").replace("Ä","Ã„")
ANREDE_VAR = {"Herr": ["Herr","Herr","Hr.","Herrn","HERR","herr"], "Frau": ["Frau","Frau","Fr.","FRAU","frau"]}

# ---------------------------------------------------------------- Manuell entstandene Störungen (Whitespace, E-Mail-Transportmüll)
NBSP = "\u00a0"
def stoer_ws(s, p_trail=0.10, p_double=0.05, p_nbsp=0.0):
    """Nachgestellte Leerzeichen, Doppel-Leerzeichen, geschütztes Leerzeichen – unsichtbar, aber join-zerstörend."""
    if not s: return s
    r = random.random()
    if r < p_trail: return s + " " * random.randint(1, 2)
    if r < p_trail + p_double and " " in s: return s.replace(" ", "  ", 1)
    if r < p_trail + p_double + p_nbsp and " " in s: return s.replace(" ", NBSP, 1)
    return s
def stoer_mail(m, vorname, nachname, p=0.15):
    """Copy-Paste-Varianten aus Mailprogrammen und Webseiten."""
    if random.random() > p: return m
    return random.choice([
        "mailto:" + m, f"{vorname} {nachname} <{m}>", m.upper(), m + ";", m + ".", m.replace("@", "(at)"), " " + m + " ",
    ])

# ---------------------------------------------------------------- ERP schreiben (Windows-1252, ';', CRLF)
def w_erp(name, header, rows):
    with open(os.path.join(OUT,"erp",name+".txt"), "w", encoding="cp1252", newline="") as f:
        wr = csv.writer(f, delimiter=";", lineterminator="\r\n"); wr.writerow(header); wr.writerows(rows)
os.makedirs(os.path.join(OUT,"erp"), exist_ok=True)
# Fehlerbilder aus der Praxis (PA1 Kriterienkatalog, ServiceTool): richtiger Wert im falschen Feld, Text statt Wert
ERP_FELDLAENGE = 30   # Firmenfeld im ERP begrenzt: längere Namen werden beim Speichern abgeschnitten
LAND_VAR_ERP = {"DE":["DE","Deutschland","Deutsch","BRD"], "AT":["AT","AUT","Österreich"], "CH":["Schweiz","SUI"]}
LAND_VAR_CRM = {"DE":["DE","Germany","Deutschl.","D"], "AT":["Oesterreich","AT","Austria"], "CH":["CH","Switzerland"]}
SPERRVERMERK = ["*INSOLVENT*", "GESPERRT", "NICHT VERWENDEN", "*** gesperrt lt. GF ***"]
def tel_text(t): return t + random.choice([" -16 Hr. Vogel", " (Durchwahl 22)", " Fr. Maier vormittags", " oder mobil"])
def tippfehler(s):   # Buchstabendreher
    i = random.randint(1, max(1, len(s)-3)); return s[:i] + s[i+1] + s[i] + s[i+2:]
kunden_erp = []
for c in KUNDEN:
    if not c["erp"]: continue
    E = lambda typ: log("ERP", "KUNDE", c["erp"], typ)
    gesperrt = c["praesenz"]=="ERP"
    plz = str(int(c["plz"]))  # Zahl: führende Null geht verloren (Dresden 01067 -> 1067)
    name1 = (c["titel"]+" " if c["titel"] else "") + c["nachname"]
    ort = c["ort"]; firma = c["firma"]; strasse = strasse_erp(c); vorname = c["vorname"]
    land = {"DE":"D","AT":"A","CH":"CH"}[c["land"]]
    if strasse[:1].isdigit(): E("Hausnummer vor Straße")
    if f(0.02): plz = random.choice([plz[:-1], plz + "0", plz[:2] + "O" + plz[3:]]); E("PLZ ungültig")   # zu kurz, zu lang, Buchstabe O statt 0
    elif f(0.06): plz = land + "-" + c["plz"]; E("PLZ mit Länderpräfix")
    elif f(0.03): ort = c["plz"] + " " + ort; plz = ""; E("PLZ im Ortsfeld")
    if f(0.08): land = random.choice(LAND_VAR_ERP[c["land"]]); E("Land-Variante")
    if "+" in c["praesenz"] and len(firma) > 8 and f(0.04): firma = tippfehler(firma); E("Tippfehler im Namen")
    if f(0.03): firma = firma + " / " + random.choice(["Hr.","Fr."]) + " " + c["nachname"]; E("Ansprechpartner im Firmenfeld")
    if len(firma) > ERP_FELDLAENGE: firma = firma[:ERP_FELDLAENGE]; E("Feldlänge: Name abgeschnitten")
    if f(0.02): strasse = "Postfach " + str(random.randint(1000,9999)); E("Postfach im Straßenfeld")
    elif not strasse[:1].isdigit() and f(0.03): strasse = strasse + random.choice(["-"+str(random.randint(2,9)), " + "+str(random.randint(2,9)), "/1"]); E("Hausnummer-Bereich")
    ziel = random.choice([30,30,14,60,-1])
    if ziel > 0 and f(0.05): ziel = random.choice([f"{ziel} Tage", f"{ziel}T netto", f"{ziel} Tg."]); E("Zahlungsziel als Text")
    nl = "J" if c["newsletter"] else "N"
    if f(0.03): nl = random.choice(["X","?","j","ja"]); E("J/N-Feld mit Fremdwert")
    if f(0.03): vorname = random.choice(["-","k.A.","?","xxx"]); E("Platzhalter statt leer")
    tel = telefon(c["land"])
    if f(0.08): tel = tel_text(tel); E("Telefon mit Text")
    bem = random.choice(BEMERKUNGEN + ["   ", " "])
    if not gesperrt and f(0.04): bem = random.choice(SPERRVERMERK); E("Sperrvermerk nur im Freitext")
    kunden_erp.append([c["erp"], stoer_ws(firma), name1, vorname, strasse, plz, ort,
                       land, {"Verein":1,"Händler":2,"Gastro":3,"Privat":4}[c["typ"]],
                       vertreter_fuer(c["land"], c["plz"]), ziel, d_de(c["erst"]),
                       nl, "J" if gesperrt else "N",
                       d_de(date(2023,random.randint(6,12),28)) if gesperrt else "31.12.9999", bem, tel])
# 4 Dubletten im ERP: gleicher Kunde zweimal angelegt (anderer Mitarbeiter, andere Schreibweise), neue KdNr
for row in random.sample(kunden_erp, 4):
    erp_nr += 1; dup = list(row); dup[0] = erp_nr
    dup[1] = dup[1].replace("e.V.", "").replace("e. V.", "").strip(); dup[4] = dup[4].replace("str.", "straße").replace("Str.", "Straße")
    dup[11] = d_de(date(2024, random.randint(1,12), random.randint(1,28)))
    kunden_erp.append(dup); log("ERP", "KUNDE", erp_nr, "Dublette im ERP (führend: " + str(row[0]) + ")")
# Testkunden, die nie gelöscht wurden
for nm in ["Testkunde", "TEST bitte nicht löschen"]:
    erp_nr += 1
    kunden_erp.append([erp_nr, nm, "Test", "Test", "Teststr. 1", "12345", "Teststadt", "D", 4, "I01", 30, d_de(date(2022,1,1)), "N", "N", "31.12.9999", "", ""])
    log("ERP", "KUNDE", erp_nr, "Testkunde")
w_erp("KUNDE", ["KdNr","Firma","Name1","Vorname","Strasse","PLZ","Ort","Land","Kundengruppe","VertreterNr","Zahlungsziel","Anlagedatum","Newsletter","Gesperrt","GesperrtBis","Bemerkung","Telefon"], kunden_erp)
w_erp("VERTRETER", ["VertreterNr","Name","Gebiet","Austritt"], [[v[0], f"{v[1]}, {v[2]}", v[4], v[6]] for v in VERTRETER])
w_erp("AUFTRAG", ["AuftragNr","KdNr","Datum","VertreterNr","Status","Waehrung","Lieferart","Zahlungsziel","ShopRef"], erp_auftraege)
w_erp("AUFTRAGSPOS", ["AuftragNr","Pos","ArtNr","Menge","Einheit","EPreisNetto"], erp_pos)
w_erp("ZAHLUNG", ["ZahlungNr","AuftragNr","Datum","Betrag","Zahlart"], erp_zahlung)
w_erp("LIEFERUNG", ["LieferNr","AuftragNr","Versanddatum","Lieferart"], erp_lieferung)
w_erp("ARTIKEL", ["ArtNr","Bezeichnung","EKPreis"], [[p[0],p[2],dez(p[4])] for p in PRODUKTE])

# ---------------------------------------------------------------- Shop schreiben (UTF-8, ',', LF)
def w_shop(name, header, rows):
    with open(os.path.join(OUT,"shop",name+".csv"), "w", encoding="utf-8", newline="") as f:
        wr = csv.writer(f, delimiter=",", lineterminator="\n"); wr.writerow(header); wr.writerows(rows)
os.makedirs(os.path.join(OUT,"shop"), exist_ok=True)
cust = []
for c in KUNDEN:
    if not c["shop"]: continue
    gender = {"m":1,"w":2}[c["geschlecht"]] if random.random()<0.85 else 0
    erp_ref = str(c["erp"]) if c["erp"] and random.random()<0.4 else ""
    S = lambda typ: log("Shop", "customers", c["shop"], typ)
    mail = stoer_mail(c["email"], c["vorname"], c["nachname"])
    if mail != c["email"]: S("E-Mail-Transportmüll (mailto, Name <>, Großschrift, Satzzeichen, (at))")
    if f(0.02): mail = c["email"].replace("@", random.choice(["", "#", " at "])); S("E-Mail ohne @")
    elif f(0.04): mail = c["email"] + random.choice(["; ", ", ", " / "]) + "buero." + c["email"]; S("zwei E-Mails in einem Feld")
    elif f(0.03): mail = c["email"].replace("@", " @ "); S("E-Mail mit Leerzeichen")
    elif f(0.02): mail = random.choice(["keine Mails!", "nix gefunden", "tel. bestellen", "-"]); S("Bemerkung im E-Mail-Feld")
    vn, nn = c["vorname"], c["nachname"]
    if f(0.05): nn = random.choice([nn.upper(), nn.lower()]); S("Groß-/Kleinschreibung")
    tel = telefon(c["land"]) if f(0.6) else ""
    if tel and f(0.06): tel = tel_text(tel); S("Telefon mit Text")
    cust.append([c["shop"], mail, vn, nn, gender, stoer_ws(strasse_shop(c), 0.0, 0.08), c["land"],
                 "true" if c["typ"]=="Verein" else "false", "true" if c["newsletter"] else "false",
                 d_iso(datetime.combine(c["erst"],datetime.min.time())+timedelta(hours=random.randint(7,23))), erp_ref, c["firma"] if c["typ"]!="Privat" else "", tel])
cust.append(["C-99999", "test@test.de", "Test", "Test", 0, "Teststr. 1, 12345 Teststadt", "DE", "false", "false", "2022-01-01T10:00:00Z", "", "", ""])
log("Shop", "customers", "C-99999", "Testkunde")
w_shop("customers", ["customer_id","email","first_name","last_name","gender","address","country","is_club","newsletter","created_at","erp_customer_no","company","phone"], cust)
w_shop("orders", ["order_id","customer_id","order_date","status","total_gross","shipping_method_id","coupon_code"], shop_orders)
w_shop("order_items", ["order_id","line_no","sku","qty","unit_price_gross"], shop_items)
w_shop("shipping_methods", ["id","name"], [[1,"DHL Paket"],[2,"DPD"],[3,"Abholung"]])
w_shop("products", ["sku","name","category"], [[p[1],p[2],p[3]] for p in PRODUKTE])

# ---------------------------------------------------------------- CRM schreiben (xlsx, US-Datum, GUIDs)
os.makedirs(os.path.join(OUT,"crm"), exist_ok=True)
wb = Workbook(); ws = wb.active; ws.title = "Accounts"
ws.append(["AccountId","Firma","Anrede","Titel","Vorname","Nachname","Strasse","Hausnr","PLZ","Ort","Land","Kundentyp","Segment",
           "Betreuer","ERP_KdNr","Erstkontakt","Newsletter","Status","LetzteAenderung","E-Mail","Telefon"])
betreuer_name = {v[0]: f"{v[2]} {v[1]}" for v in VERTRETER}
accounts = []
for c in KUNDEN:
    if not c["crm"]: continue
    src = c["umzug"] if c["umzug"] else c
    anrede = {"m":"Herr","w":"Frau"}[c["geschlecht"]]
    if random.random() < 0.08: anrede = "Frau" if anrede=="Herr" else "Herr"   # Widerspruch zum Shop
    if random.random() < 0.05: anrede = ""
    C = lambda typ: log("CRM", "Accounts", c["crm"], typ)
    if anrede and f(0.10): anrede = random.choice(ANREDE_VAR[anrede])          # Hr./Herrn/HERR
    if f(0.03): anrede = random.choice(["k.A.", "-", "?"]); C("Platzhalter statt leer")
    strasse, hnr = strasse_crm(c)
    ort = src["ort"] + (f" ({random.choice(['Zentrum','West','Altstadt','Industriegebiet'])})" if random.random()<0.10 else "")
    typ = c["typ"]
    if typ=="Händler": typ = random.choice(["Fachhändler","Fachhändler","Großhändler"])
    if typ=="Gastro":  typ = "Gastronomie"
    status = "Lead" if c["praesenz"]=="CRM" else random.choices(["Kunde","Inaktiv"],[92,8])[0]
    erp_ref = ""
    if c["erp"] and random.random()<0.6:
        erp_ref = str(c["erp"] + (random.choice([1,-1,10]) if random.random()<0.05 else 0))  # 5 % Tippfehler
        if erp_ref != str(c["erp"]): C("Fremdschlüssel ERP_KdNr mit Tippfehler")
    vnr = vertreter_fuer(c["land"], c["plz"])
    if vnr=="V01": vnr = "V02"   # CRM wurde nach Kellers Austritt umgehängt
    aend = src["datum"] if c["umzug"] else date(2022,6,1)+timedelta(days=random.randint(0,800))
    strasse = stoer_ws(strasse, 0.05, 0.0, 0.05)
    nachname = name_crm(c); firma = firma_crm(c); vorname = c["vorname"]
    if c["typ"]=="Privat" and f(0.30): firma, vorname, nachname = f"{c['nachname']} {c['vorname']}", "", ""; C("Privatkunde nur im Firmenfeld")
    if any(ch in nachname + firma for ch in "äöüßÄÖÜ") and f(0.30):
        nachname, firma = mojibake(nachname), mojibake(firma); C("Mojibake (Zeichensatzfehler beim Export)")
    if status=="Kunde" and f(0.03): firma = (firma or nachname) + " (" + random.choice(["INSOLVENT","GESPERRT","nicht verwenden"]) + ")"; C("Sperrvermerk im Firmenfeld")
    land_txt = {"DE":"Deutschland","AT":"Österreich","CH":"Schweiz"}[c["land"]]
    plz_txt = src["plz"]
    if f(0.03): land_txt = ""; C("Land fehlt")
    elif f(0.10): land_txt = random.choice(LAND_VAR_CRM[c["land"]]); C("Land-Variante")
    if f(0.02): plz_txt = ""; C("PLZ fehlt")
    elif f(0.05): plz_txt = {"DE":"D","AT":"A","CH":"CH"}[c["land"]] + "-" + plz_txt; C("PLZ mit Länderpräfix")
    elif f(0.10): plz_txt = int(plz_txt); C("PLZ als Zahl (Excel-Artefakt, führende Null fehlt)" if src["plz"].startswith("0") else "PLZ als Zahl (Excel-Artefakt)")
    erst_txt = d_us(c["erst"])
    if f(0.05): erst_txt = c["erst"].strftime("%m/%Y"); C("Datum ungenau (nur MM/YYYY)")
    betreuer = betreuer_name[vnr]
    if f(0.05): betreuer = random.choice([betreuer.replace("u", "ue", 1) if "u" in betreuer else betreuer.upper(), betreuer.split()[0][0] + ". " + betreuer.split()[1], betreuer.upper()]); C("Betreuer nicht auflösbar (Fremdschlüssel per Name)")
    mail = stoer_mail(c["email"], c["vorname"], c["nachname"])
    if mail != c["email"]: C("E-Mail-Transportmüll (mailto, Name <>, Großschrift, Satzzeichen, (at))")
    tel = telefon(c["land"]) if f(0.8) else ""
    if tel and f(0.06): tel = tel_text(tel); C("Telefon mit Text")
    row = [c["crm"], stoer_ws(firma, 0.08, 0.0, 0.07), anrede, c["titel"], vorname, nachname, strasse, hnr, plz_txt, ort,
           land_txt, typ, random.choice("ABBCCC"),
           betreuer, erp_ref, erst_txt, "X" if c["newsletter"] else "", status, d_us(aend), mail,
           tel]
    accounts.append(row)
# 5 % Dubletten: gleicher Kunde, zweiter Vertreter, leicht andere Schreibweise, ohne ERP_KdNr
for row in random.sample(accounts, max(1, len(accounts)//20)):
    dup = list(row); dup[0] = str(uuid.UUID(int=random.getrandbits(128)))
    dup[1] = dup[1].replace("GmbH","").strip(); dup[6] = dup[6].replace("straße","str.").replace("Straße","Str.")
    dup[13] = random.choice([betreuer_name["I01"], betreuer_name["I02"]]); dup[14] = ""; dup[12] = "C"
    accounts.append(dup); log("CRM", "Accounts", dup[0], "Dublette im CRM")
# 2 Accounts mit vertauschten Feldern Strasse <-> Hausnr (Fehlbelegung bei manueller Erfassung)
for row in random.sample(accounts, 4):
    row[6], row[7] = row[7], row[6]; log("CRM", "Accounts", row[0], "Felder vertauscht (Straße/Hausnr)")
# 2 Accounts mit PLZ aus dem falschen Land (Plausibilitätsprüfung)
for row in random.sample(accounts, 3):
    row[8] = "88212" if row[10]!="Deutschland" else "8001"; log("CRM", "Accounts", row[0], "PLZ passt nicht zum Land")
_tid = str(uuid.UUID(int=random.getrandbits(128)))
accounts.append([_tid, "Testfirma", "", "", "Test", "Test", "Teststr.", "1", "12345", "Teststadt", "Deutschland", "Privat", "C", betreuer_name["I01"], "", "01/01/2022", "", "Kunde", "01/01/2022", "test@test.de", ""])
log("CRM", "Accounts", _tid, "Testkunde")
random.shuffle(accounts)
for r in accounts: ws.append(r)
ws2 = wb.create_sheet("Betreuer"); ws2.append(["Name","E-Mail","Rolle","Region"])
for v in VERTRETER:
    ws2.append([f"{v[2]} {v[1]}", f"{v[2][0].lower()}.{v[1].lower().replace('ü','ue')}@dartpoint.de", v[3], v[5]])
ws3 = wb.create_sheet("Aktivitaet"); ws3.append(["ActivityId","AccountId","Typ","Datum","Betreuer"])
for r in accounts:
    for _ in range(random.randint(0,4)):
        d = date(2023,1,1)+timedelta(days=random.randint(0,700))
        ws3.append([str(uuid.UUID(int=random.getrandbits(128))), r[0], random.choice(["Besuch","Anruf","Angebot","E-Mail"]), d_us(d), r[13]])
wb.save(os.path.join(OUT,"crm","CRM_Export.xlsx"))

# ---------------------------------------------------------------- Sollwerte (Ground Truth)
os.makedirs(os.path.join(os.path.dirname(OUT),"sollwerte"), exist_ok=True)
with open(os.path.join(os.path.dirname(OUT),"sollwerte","kunden_mapping.csv"), "w", encoding="utf-8", newline="") as f:
    wr = csv.writer(f, delimiter=";"); wr.writerow(["KundeKey","Praesenz","Kundentyp","Land","ERP_KdNr","Shop_ID","CRM_ID","Abweichung","Umzug_am"])
    for c in KUNDEN:
        wr.writerow([c["key"], c["praesenz"], c["typ"], c["land"], c["erp"] or "", c["shop"] or "", c["crm"] or "", c["abweichung"],
                     d_de(c["umzug"]["datum"]) if c["umzug"] else ""])
with open(os.path.join(os.path.dirname(OUT),"sollwerte","fehler_log.csv"), "w", encoding="utf-8", newline="") as fl:
    wr = csv.writer(fl, delimiter=";"); wr.writerow(["Quelle","Tabelle","ID","Fehlertyp"]); wr.writerows(sorted(FEHLER_LOG))
with open(os.path.join(os.path.dirname(OUT),"sollwerte","parameter.json"), "w", encoding="utf-8") as f:
    json.dump({"seed":SEED, "fehlerquote":FEHLERQUOTE, "kunden":len(KUNDEN), "verteilung":{p:n for p,n,_ in VERTEILUNG},
               "erp_auftraege":len(erp_auftraege), "shop_orders":len(shop_orders), "crm_accounts":len(accounts),
               "crm_dubletten":len(accounts)-sum(1 for c in KUNDEN if c["crm"]), "chf_kurs_eur":CHF_KURS, "mwst":MWST}, f, indent=2, ensure_ascii=False)

print(f"Kunden: {len(KUNDEN)} | ERP-Kunden {len(kunden_erp)} | Shop-Kunden {len(cust)} | CRM-Accounts {len(accounts)} (davon Dubletten {len(accounts)-sum(1 for c in KUNDEN if c['crm'])})")
from collections import Counter
print(f"Kundensätze gesamt: {len(kunden_erp)+len(cust)+len(accounts)} | eingebaute Fehler laut Log: {len(FEHLER_LOG)}")
for (typ, q), n in sorted(Counter((t, q) for q, _, _, t in FEHLER_LOG).items()): print(f"  {n:3d}  {q:4s}  {typ}")
print(f"ERP-Aufträge {len(erp_auftraege)} ({sum(1 for a in erp_auftraege if a[8])} Shop-Importe) | Shop-Orders {len(shop_orders)} | Positionen ERP {len(erp_pos)} / Shop {len(shop_items)}")
