"""Erstellt die Jurten-Recherche XLSX fuer Algarve, Portugal."""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = Workbook()

HEADER_FILL = PatternFill(start_color="2E5C8A", end_color="2E5C8A", fill_type="solid")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
ALT_FILL = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
TITLE_FONT = Font(bold=True, size=14, color="2E5C8A")
THIN = Side(border_style="thin", color="CCCCCC")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def style_sheet(ws, headers, rows):
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=col, value=h)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
    for r_idx, row in enumerate(rows, 2):
        for c_idx, val in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = BORDER
            if r_idx % 2 == 0:
                cell.fill = ALT_FILL
    ws.row_dimensions[1].height = 30
    for col_idx in range(1, len(headers) + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = 22
    ws.freeze_panes = "A2"


# ============ Sheet 1: Hersteller Neu ============
ws1 = wb.active
ws1.title = "Hersteller Neu"
headers1 = [
    "Hersteller", "Land", "Website", "Groesse m2", "Preis EUR",
    "Material", "Klima-Eignung Algarve", "Versand nach PT",
    "Besonderheiten", "Gesamtkosten geschaetzt EUR",
]

# Sortiert nach Preis aufsteigend (Basis-Modell ca. 6m)
rows1 = [
    ["Adorjan Jurta", "Ungarn", "https://adorjan-jurta.hu",
     "28 m2 (6m)", 1750,
     "Holzgitter + Filz + Canvas",
     "Gering ohne Anpassung - traditionelles Modell fuer Mitteleuropa, UV-Schutz schwach",
     "Selbstabholung/Spedition - ca. 800-1500 EUR",
     "Sehr guenstig, lange Lieferzeit (Wartezeit 1+ Jahr), Handarbeit Gabor Adorjan",
     "2.500-3.500"],

    ["Kyrgyz Yurte (eBay DE - Windeck NRW)", "Kirgistan/Deutschland", "https://www.ebay.de/itm/270505657535",
     "12,6 m2 (4m)", 3868,
     "Filz + Canvas (traditionell)",
     "Sehr gering - Filz fuer Steppenklima, kaum UV-Schutz, Plane nur 2-3 Jahre Lebensdauer in PT-Sonne",
     "Lieferung DE moeglich, von DE nach PT ca. 500-800 EUR Spedition",
     "Traditionell, klein, Pickup in NRW oder Versand. Filz bei Hitze problematisch",
     "4.400-4.700"],

    ["Kyrgyz Yurte 5m (eBay DE)", "Kirgistan/Deutschland", "https://www.ebay.de/itm/270555654183",
     "20 m2 (5m)", 4265,
     "Filz + Canvas (traditionell)",
     "Gering - traditionelle Bauweise nicht fuer 40 Grad",
     "Spedition ca. 500-800 EUR ab DE",
     "Vorrat verfuegbar, kein UV-Schutz",
     "4.800-5.100"],

    ["Mongolei Shop / jurte.info (Bonn)", "Deutschland", "https://mongoleishop.de/5-wall-Mongolian-Yurt",
     "25 m2 (5-wand, ca. 5,5m)", 5390,
     "Filz + Canvas",
     "Mittel - originale mongolische Bauweise, Aussenplane unter UV anfaellig",
     "Abholung Bonn oder Versand +590 EUR (DE), Spedition nach PT ca. 1.500-2.000 EUR",
     "Wildberry Handelsges. seit Jahren etabliert, Ersatzteile vorraetig",
     "7.000-7.500"],

    ["Yourtepoque (5m)", "Frankreich", "https://yourtepoque.com/5.html",
     "20 m2 (5m)", 5800,
     "Holz + Baumwolle/Polyester Plane",
     "Mittel - europaeisch angepasst, UV-bestaendige Plane, fuer 40 Grad Beschattung noetig",
     "Spedition FR->PT ca. 800-1.200 EUR",
     "Bestes Preis-Leistungs-Verhaeltnis FR. Verschiedene Groessen",
     "6.700-7.200"],

    ["Silk Road Yurts (6,18m)", "Mongolei", "https://silkroadyurts.com/product/traditional-mongolian-yurt/",
     "30 m2 (6,18m)", 5040,
     "Filz + Canvas + Holz",
     "Gering - traditionell mongolisch, Anpassung an PT-Klima noetig",
     "Versand UB->EU ca. 2.300 EUR (USD 2.500)",
     "Direkt aus Mongolei, USD 5.450 Werkpreis, dann +Versand+Zoll+IVA",
     "9.000-10.500 inkl. Zoll/IVA"],

    ["Yurt Specialists (6m)", "UK", "https://yurtspecialists.com/product/6m-19-6ft-diameter-yurt/",
     "28 m2 (6m)", 5750,
     "Holz + UV-bestaendiges Canvas, Wollfilz",
     "Gut - explizit UV-bestaendiges Heavy-Duty Canvas, fuer EU-Klima angepasst",
     "UK->PT Spedition: ca. 1.500-2.500 EUR (Post-Brexit Zoll & IVA noetig)",
     "GBP 4.913 ex VAT. UK-Brexit: nochmal Zoll bei Einfuhr nach EU",
     "8.500-10.000"],

    ["Yourtepoque (6m)", "Frankreich", "https://yourtepoque.com/6.html",
     "27 m2 (6m)", 7240,
     "Holz + Baumwoll-Plane",
     "Mittel - europaeisch, Plane braucht Schatten in PT-Sommer",
     "Spedition FR->PT ca. 800-1.200 EUR",
     "Solide Mittelpreis-Option, transparenter Preis",
     "8.000-8.500"],

    ["Mongolian Yurts Direct (6m)", "UK/Mongolei", "https://www.yurtsdirect.com/sizes-and-prices",
     "28 m2 (6m)", 7000,
     "Holz + Wollfilz + wasserdichte Aussenplane",
     "Mittel - mongolische Bauweise, EU-angepasst",
     "UK->PT 1.500-2.500 EUR + EU-Einfuhrzoll/IVA",
     "GBP 5.995 inkl. VAT, Brexit-Zollkosten zusaetzlich",
     "9.000-11.000"],

    ["Yourtepoque (6,5m)", "Frankreich", "https://yourtepoque.com/65.html",
     "33 m2 (6,5m)", 8680,
     "Holz + Plane (UV)",
     "Mittel-Gut", "Spedition FR->PT ca. 1.000-1.500 EUR",
     "Beliebte Familiengroesse",
     "9.700-10.200"],

    ["Yourtepoque (7,5m)", "Frankreich", "http://yourtepoque.com/75.html",
     "44 m2 (7,5m)", 10500,
     "Holz + Plane (UV)",
     "Mittel-Gut", "Spedition FR->PT ca. 1.200-1.500 EUR",
     "Geraeumig fuer Wohnen",
     "11.700-12.000"],

    ["Yourtepoque (8m)", "Frankreich", "https://yourtepoque.com/8m.html",
     "50 m2 (8m)", 11820,
     "Holz + Plane (UV)",
     "Mittel-Gut", "Spedition FR->PT ca. 1.500-2.000 EUR",
     "Familienjurte / Wohnen ganzjaehrig",
     "13.300-13.800"],

    ["Jurte24 (Wendt) 6m", "Deutschland", "https://www.jurte24.de/p/moderne-mongolische-jurte-6m-durchmesser-ca-29m-komplett-winterfest-und-wasserdicht-170-cm-wand",
     "28-29 m2 (6m)", 12050,
     "Holz + 100% Schafwollfilz + winterfeste Aussenplane",
     "Mittel - winterfest ausgelegt, fuer PT Hitze braucht es Schatten/Lueftung",
     "Spedition DE->PT ca. 1.500-2.000 EUR",
     "Inkl. 19% MwSt, hohe deutsche Qualitaet, Lieferung ab DE",
     "13.500-14.000"],

    ["Yourteco Tiny Yurt 19m2", "Frankreich", "https://www.yourteco.com/la-surface-dune-tiny-yourte-est-egale-a-20-m2-ou-30-m2/",
     "19 m2 (5m)", 13180,
     "Massivholz + isolierte Wandpaneele",
     "Sehr gut - moderne Holzbauweise statt Plane, hitze-/UV-stabiler",
     "Spedition FR->PT ca. 1.500-2.500 EUR (groessere Holzkonstruktion)",
     "Made in France, zeitgenoessisch, langlebig",
     "15.000-16.000 (an Budget-Grenze)"],

    ["La Yourte Francaise (basis 30m2)", "Frankreich", "https://www.layourtefrancaise.fr/tarifs-des-yourtes-francaises/",
     "30 m2", 8100,
     "Holz + Aussenplane, Glas-Doppeltueren PVC",
     "Mittel - ERP-zugelassene franz. Bauweise",
     "Spedition FR->PT inkl. Montage: 2.000-3.500 EUR",
     "270 EUR/m2 ab Werk netto, Montage extra, fuer Gewerbe ERP-zertifiziert",
     "12.000-14.500 (Basis ERP)"],

    ["Yourteco Tiny Yurt 28m2", "Frankreich", "https://www.yourteco.com",
     "28 m2 (5,5m)", 14500,
     "Massivholz + isolierte Paneele",
     "Sehr gut", "Spedition FR->PT 1.500-2.500 EUR",
     "Premium Holz-Jurte, Markenfabrikant",
     "16.000-17.000 (UEBER BUDGET)"],

    ["Jurtenmanufaktur Jena (basis)", "Deutschland", "https://www.jurtenmanufaktur.de/jurten/preise",
     "ab ca. 20-30 m2", 14500,
     "Holz + Schafwollfilz + Aussenplane",
     "Mittel-Gut", "Spedition DE->PT 1.500-2.000 EUR",
     "Wohnjurten ab 14.500 EUR, individuelle Anfrage",
     "16.000-17.500 (UEBER BUDGET)"],

    ["Jurtenwerkstatt OG", "Oesterreich", "https://www.jurtenwerkstatt.at/preise",
     "20 m2 (5m) - 50 m2", "850-1000 EUR/m2",
     "Holz + Filz + Canvas",
     "Mittel - europaeisch optimiert",
     "Spedition AT->PT 1.500-2.500 EUR",
     "Winterfest, handgefertigt, 20m2 ~17.000-20.000 EUR -> ueber Budget",
     "ueber Budget bei 20m2+"],

    ["Bet Yurts (Algarve)", "Portugal (Algarve)", "https://www.betyurts.com/",
     "5-9 m Durchmesser", "Anfrage (vergleichbar 10-20k)",
     "Holz + Aussenhuelle, glasierte Tueren",
     "Sehr gut - lokal in Algarve gebaut, klima-optimiert, Schreiner+Designerin",
     "Inland - Transport innerhalb PT inklusive (entfernungsabh.)",
     "BESTE LOKALE OPTION. Carpenter+Designer Duo, lokale Materialien. +800 EUR pro extra Tuer/Fenster",
     "10.000-15.000 (geschaetzt)"],

    ["Casa dos Sonhos Yurts (Alto Alentejo, Besteiros de Cima/Portalegre)",
     "Portugal (Alto Alentejo)",
     "https://www.casadossonhos.co.uk/copy-of-environmental-policy",
     "5m=20m2 / 6,1m=30m2 / 7,3m=42m2 / 9,1m=64m2", "Anfrage",
     "Portugiesisches Holz (Tueren, Khana-Gitter, Sparren, Fenster), portugiesische Korkboeden + Korkdaemmung, moderne Isolierung, Aussenhuelle mit 10-J. UV-Garantie",
     "EXZELLENT - 10 Jahre UV-Garantie auf Aussenhuelle, Plane rottet/schimmelt nicht, braucht keine Nachbehandlung; speziell fuer portugiesisches Wetter gefertigt",
     "Innerhalb Festland-PT inkl. Lieferung & Aufbau (im Preis enthalten)",
     "STANDARD inkl.: 1x1 m doppelverglastes Alu-Fenster (9,1m: 120x120cm), Holztuer mit Fenster+Laden, Daemmung, oeffenbares Kuppel-Oberlicht. Selbsttragend ohne Mittelsaeule, Wand >2 m, Mittelring min. 3,5 m. Solarwerkstatt. OPTIONEN: groessere Fenster, Sturm-Kit, Extra-Daemmlagen, franz. Doppeltueren 3/4-Glas+Lade, Kaminzug, Mueckennetze",
     "8.000-15.000 (geschaetzt nach Groesse) - Lieferung+Aufbau PT inklusive"],

    ["Yurt Workshop Spain (Cadiar)", "Spanien (Granada)", "http://yurtworkshop.es/",
     "5-9m", "Anfrage",
     "Spanisch-Kastanienholz + flammenhemmendes Canvas + Wollfilz",
     "Gut - 15 Jahre Erfahrung in suedlich-mediterranem Klima (Alpujarras)",
     "Spanien->PT relativ guenstig 600-1.200 EUR Spedition",
     "Rob Matthews, 3 Spezifikationen: standard/isoliert/Glaswand. Hat Trainings in Kirgistan",
     "8.000-13.000"],

    ["TipiWakan (Girona)", "Spanien (Katalonien)", "https://www.tipis.es/en/our-structures/yurts/",
     "5/6/6,5 muros (Waende)", "Anfrage",
     "Canvas + Holz, handgefertigt",
     "Gut - mediterran erfahren",
     "ES->PT 800-1.500 EUR Spedition",
     "22+ Jahre Erfahrung, Lager verfuegbar fuer Sofortlieferung, Sonderrabatte",
     "ca. 6.000-12.000"],

    ["Horizon Yourte", "Frankreich (Cevennes)", "https://www.horizon-yourte.fr/",
     "11-60 m2 (4-9m)", "Anfrage",
     "Holzgitter + Wollfilz + Plane",
     "Mittel-Gut - europaeisch",
     "FR->PT 1.000-1.800 EUR",
     "15+ Jahre Erfahrung, 150+ Jurten geliefert",
     "8.000-15.000"],

    ["Yurt Workshop UK", "UK", "https://www.yurtworkshop.com/",
     "5-7m", "Anfrage GBP/EUR",
     "Eichen-/Kastanien-Holz + Canvas + Wollfilz",
     "Mittel - Brexit-Zoll Aufschlag",
     "UK->PT 1.500-2.500 EUR + EU-Zoll/IVA bei Brexit-Import",
     "Etabliert; Brexit verteuert Lieferung in EU erheblich",
     "9.000-13.000"],

    ["FamWest", "Deutschland", "https://www.famwest.de/en/mongolian-yurts/",
     "5-9m Durchmesser", "Anfrage",
     "Filz + Canvas",
     "Mittel", "DE->PT 1.500-2.000 EUR",
     "5/6/7/8/9 m Optionen, inkl 19% MwSt",
     "ca. 6.000-15.000"],

    ["Yourteco Flex Yurt 10m2", "Frankreich", "https://www.yourteco.com/",
     "10 m2", 3900,
     "Holz, kompakt",
     "Sehr gut für kleine Strukturen",
     "FR->PT 800-1.200 EUR",
     "Sehr klein - eher Gartenstudio",
     "4.700-5.100"],

    ["Mongolian Yurts Direct (5m)", "UK/Mongolei", "https://www.yurtsdirect.com/sizes-and-prices",
     "20 m2 (5m)", 4650,
     "Filz + Canvas + Holz",
     "Mittel", "UK->PT 1.500-2.500 EUR + EU-Zoll/IVA (Brexit)",
     "GBP 3.995 inkl VAT, plus Brexit-Einfuhr",
     "6.500-8.000"],
]

# sortiere nach Preis EUR (numerisch wo moeglich)
def sort_key(r):
    p = r[4]
    if isinstance(p, (int, float)):
        return p
    return 9999999  # "Anfrage" ans Ende
rows1.sort(key=sort_key)

style_sheet(ws1, headers1, rows1)
# Spaltenbreiten anpassen
widths1 = [26, 18, 38, 22, 14, 28, 32, 28, 38, 22]
for i, w in enumerate(widths1, 1):
    ws1.column_dimensions[get_column_letter(i)].width = w

# ============ Sheet 2: Gebrauchtmarkt ============
ws2 = wb.create_sheet("Gebrauchtmarkt")
headers2 = [
    "Angebot/Plattform", "Plattform", "Standort", "Groesse m2",
    "Preis EUR", "Zustand", "Link", "Versand nach PT moeglich", "Anmerkungen",
]
rows2 = [
    ["Kyrgyz Jurte 2,5m", "eBay.de / PicClick", "Windeck NRW (DE)",
     "ca. 5 m2 (2,5m)", 1550, "Neuwertig (Einzelangebot)",
     "https://picclick.de/Jurta-Jurte-aus-Kirgisien-kirgistan-Kasachstan-%88%85-25m-261376413355.html",
     "Ja - Spedition ab Windeck, ca. 500-800 EUR",
     "Sehr kleine Demo-Jurte"],

    ["Kyrgyz Jurte 4m", "eBay.de", "Windeck NRW (DE)",
     "12,6 m2 (4m)", 3868, "Gebraucht/Restposten",
     "https://www.ebay.de/itm/270505657535",
     "Ja - Spedition 500-800 EUR",
     "EUR 3.250 netto, 3.868 brutto"],

    ["Kyrgyz Jurte 5m", "eBay.de", "Windeck NRW (DE)",
     "20 m2 (5m)", 4265, "Gebraucht/Vorrat",
     "https://www.ebay.de/itm/270555654183",
     "Ja - Spedition 500-800 EUR", "Filz+Canvas traditionell"],

    ["Kleinanzeigen - Jurte Suche", "kleinanzeigen.de", "Deutschland weit",
     "diverse", "1.000-15.000",
     "Neu bis stark gebraucht",
     "https://www.kleinanzeigen.de/s-jurte/k0",
     "Ja - Spedition 500-1.500 EUR",
     "Marktplatz, wechselnde Angebote, regelmaessig pruefen"],

    ["eBay.de Jurten Kategorie", "eBay.de", "Deutschland",
     "diverse", "200-10.000+", "Neu/gebraucht",
     "https://www.ebay.de/b/Jurte-Zelt/bn_7005694261",
     "Versand variiert",
     "Auch Pfadfinder-Jurten (gross, billig, aber Plane nicht UV-fest)"],

    ["eBay.co.uk Yurt", "eBay UK", "UK",
     "diverse", "1.500-12.000", "Gebraucht",
     "https://www.ebay.co.uk/sch/i.html?_nkw=yurt",
     "Ja - Brexit-Zoll bei Import nach PT (~7% Zoll + 23% IVA)",
     "Brexit verteuert UK-Importe stark"],

    ["LeBonCoin Yourte", "leboncoin.fr", "Frankreich (variable)",
     "10-50 m2", "1.500-20.000", "Gebraucht",
     "https://www.leboncoin.fr/ck/ventes_immobilieres/yourte",
     "Ja - Spedition FR->PT 800-1.500 EUR",
     "Beispiel: 6m yourte mit Wollisolierung gelistet"],

    ["Wallapop Yurtas", "wallapop.com", "Spanien",
     "diverse", "5.000-25.000", "Gebraucht",
     "https://es.wallapop.com/muebles-deco-y-jardin/yurtas",
     "Ja - ES->PT guenstig 500-1.000 EUR",
     "9 Eintraege typischerweise gelistet"],

    ["Milanuncios Yurta", "milanuncios.com", "Spanien",
     "diverse 40-80 m2", "22.500-42.000", "Neu/gebraucht",
     "https://www.milanuncios.com/anuncios/yurt-yurta.htm",
     "Ja", "Komplett ausgestattet 27.500 EUR, ohne Moebel 22.500 EUR (Beispiel)"],

    ["OLX Portugal yurts", "olx.pt", "Portugal",
     "diverse", "variabel", "Neu/gebraucht",
     "https://www.olx.pt/moveis-casa-e-jardim/q-yurts/",
     "Ja - innerhalb PT",
     "INLAND PT - keine Zoll, kein Auslandsversand"],

    ["Yurt Workshop Secondhand", "yurtworkshop.com (Spanien)", "Spanien Granada",
     "5-7m", "3.000-10.000", "Gebraucht (vom Hersteller)",
     "https://www.yurtworkshop.com/Buy_Your_Yurt/Second-Hand-Yurts-For-Sale.html",
     "Ja - ES->PT 600-1.200 EUR",
     "Vom Hersteller selbst - Qualitaet gepruef"],

    ["Casa Dos Sonhos Secondhand", "Facebook Casa Dos Sonhos", "Portugal (Alentejo)",
     "5m", "Anfrage",
     "Gebraucht direkt vom Eigentuemer",
     "https://www.facebook.com/yurtscasadossonhos/posts/3660919220624833/",
     "Ja - inland PT", "Hersteller direkt, ggf. Garantie"],

    ["Yurt Market (worldwide)", "yurtmarket.com", "International",
     "diverse", "variabel", "Neu/gebraucht",
     "https://www.yurtmarket.com/",
     "Versand variiert je Anbieter",
     "Welt-Marktplatz"],

    ["Yurt Forum Classifieds", "yurtforum.com", "International",
     "diverse", "variabel", "Gebraucht",
     "https://www.yurtforum.com/classifieds/",
     "Versand variiert", "Community-Marktplatz, oft USA-fokussiert"],

    ["Groovy Yurts Clearance", "groovyyurts.com", "Kanada/USA",
     "diverse", "USD 2.790-18.950", "Restposten/leicht gebraucht",
     "https://www.groovyyurts.com/catalogue-clearance-and-sale",
     "Nein - hauptsaechlich Nordamerika; Versand teuer",
     "Eher fuer US/CA - hohe Versandkosten nach EU/PT"],

    ["Facebook Marketplace EU", "facebook.com/marketplace", "Diverse EU",
     "diverse", "variabel", "Gebraucht",
     "https://www.facebook.com/marketplace/category/search/?query=yurt",
     "Regional pruefen", "Lokale Suche pro Land empfohlen"],
]
rows2.sort(key=lambda r: r[4] if isinstance(r[4], (int, float)) else 9999999)
style_sheet(ws2, headers2, rows2)
widths2 = [28, 20, 24, 18, 16, 22, 42, 28, 32]
for i, w in enumerate(widths2, 1):
    ws2.column_dimensions[get_column_letter(i)].width = w


# ============ Sheet 3: Alibaba Direktimport ============
ws3 = wb.create_sheet("Alibaba Direktimport")
headers3 = [
    "Hersteller", "Link", "Preis ab Werk USD", "Versand geschaetzt EUR",
    "Zoll + IVA EUR", "Gesamtkosten EUR", "Groesse", "Material", "Bewertung Supplier",
]

# Berechnungslogik:
# 1 USD = 0,92 EUR (Annahme)
# Zoll Zelte (HS 6306.22): typisch 12% EU
# IVA Portugal: 23% (auf CIF + Zoll)
# Versand Container LCL China->Lisbon: ca. 1500-2500 EUR fuer 1 Jurte

def calc_alibaba(usd_low, usd_high, ship_eur=2000):
    usd_avg = (usd_low + usd_high) / 2
    eur_werk = usd_avg * 0.92
    cif = eur_werk + ship_eur
    zoll = cif * 0.12
    iva = (cif + zoll) * 0.23
    total = cif + zoll + iva
    return round(total)

rows3 = [
    ["Alibaba Generic Mongolian Yurt (durchschnitt)",
     "https://www.alibaba.com/showroom/mongolian-yurt-for-sale.html",
     "USD 2.684 (avg)", 1800, calc_alibaba(2500, 2900, 1800) - 2500*0.92 - 1800,
     calc_alibaba(2500, 2900, 1800),
     "5-7m", "Filz + Canvas, traditionell",
     "Marktplatz - Stichprobe noetig, Lieferantenpruefung"],

    ["Alibaba Mongolian Ger MOQ 1 Premium",
     "https://www.alibaba.com/showroom/mongolian-yurt.html",
     "USD 6.150-9.368", 2200, "siehe Gesamt",
     calc_alibaba(6150, 9368, 2200),
     "6-8m", "Holz + Filz + verstaerkte Plane",
     "Premium Lieferanten mit Trade Assurance"],

    ["Liri Tents (Guangdong) Aluminiumrahmen",
     "https://liri-tents.en.made-in-china.com/product/xvnmqaGVVNkU/",
     "USD 200-400 per m2 (FOB)", 2000, "siehe Gesamt",
     calc_alibaba(200*28, 400*28, 2000),
     "5-12m, MOQ 20 m2",
     "Aluminium-Rahmen + PVC doppelt beschichtet, UV-bestaendig, anti-rust",
     "Etablierter Hersteller, gut fuer Glamping-Betrieb, sehr UV-fest"],

    ["CZ Meister Luxury Mongolian Yurt 4-8m",
     "https://czmeister.en.made-in-china.com/product/MycxdCvbkAhY/",
     "USD 2.000-5.000 (geschaetzt)", 2000, "siehe Gesamt",
     calc_alibaba(2000, 5000, 2000),
     "4/5/6/8m", "Aluminium + Bambus + Plane",
     "Made-in-China verifiziert"],

    ["Henan Sinoyurt Outdoor Products",
     "https://jinyuanli.en.made-in-china.com/product-list-1.html",
     "USD 1.500-4.500 (geschaetzt)", 1800, "siehe Gesamt",
     calc_alibaba(1500, 4500, 1800),
     "diverse 4-8m", "Stahl/Filz/Canvas",
     "Etabliert auf Made-in-China"],

    ["Generic 6m Yurt Tent (Glamping-Style)",
     "https://www.alibaba.com/showroom/6m-yurt-tent.html",
     "USD 4.100-4.300", 1800, "siehe Gesamt",
     calc_alibaba(4100, 4300, 1800),
     "6m", "Aluminium + PVC",
     "2.200+ Listings - Marktpreis"],

    ["MORE TENT Aluminum Glamping Yurt",
     "https://www.glamping-hotel.com/products/yurt-tent/",
     "USD 3.000-6.000 (geschaetzt)", 2000, "siehe Gesamt",
     calc_alibaba(3000, 6000, 2000),
     "5-10m", "Aluminium + PVC",
     "Etablierter Glamping-Spezialist"],

    ["Alibaba Wholesale ab USD 19 (Bell Tent)",
     "https://www.alibaba.com/showroom/yurt-tent-sale.html",
     "USD 19-1.765", 1500, "siehe Gesamt",
     calc_alibaba(500, 1500, 1500),
     "klein bis mittel", "Canvas Bell-Tent (kein echtes Ger)",
     "Vorsicht - meist Glockenzelte, kein traditioneller Ger"],
]

# In rows3 wir setzen die Zoll+IVA-Spalte als Differenz
def fix_row(r):
    werk_eur = None
    p = r[2]
    if isinstance(p, str) and "USD" in p:
        # Wert nicht numerisch, lass 'siehe Gesamt'
        pass
    return r

style_sheet(ws3, headers3, rows3)
widths3 = [34, 38, 24, 18, 18, 22, 18, 32, 30]
for i, w in enumerate(widths3, 1):
    ws3.column_dimensions[get_column_letter(i)].width = w


# ============ Sheet 4: Top 5 Empfehlungen ============
ws4 = wb.create_sheet("Top 5 Empfehlungen")
headers4 = ["Rang", "Anbieter", "Typ", "Preis EUR", "Warum empfohlen", "Link"]
rows4 = [
    [1, "Casa dos Sonhos Yurts (Alto Alentejo)", "Hersteller PT - lokale Spitzenoption",
     "ca. 8.000-15.000 nach Groesse (Lieferung+Aufbau PT inkl.)",
     "BESTE KLIMA-EIGNUNG IM MARKT: 10 Jahre UV-Garantie auf Aussenhuelle, Plane rottet/schimmelt nicht, keine Nachbehandlung noetig. Standorte: 5m/6,1m/7,3m/9,1m (20/30/42/64 m2). SELBSTTRAGEND ohne Mittelsaeule, Waende >2 m, Mittelring 3,5 m -> super geraeumig. STANDARD inkl.: doppelverglastes Alu-Fenster, Holztuer m. Fenster+Laden, Daemmung, oeffenbares Kuppel-Oberlicht. Portugiesisches Holz + portug. Korkboeden + Korkdaemmung. Solarwerkstatt in Alto Alentejo (Portalegre). Lieferung und Aufbau in Festland-PT IM PREIS. Kein Import-Zoll/IVA. Optionen: Sturm-Kit (Algarve-Winterstuerme!), Extra-Daemmung, franz. Doppeltueren, Kaminzug, Mueckennetze.",
     "https://www.casadossonhos.co.uk/copy-of-environmental-policy  |  https://www.facebook.com/yurtscasadossonhos/"],

    [2, "Bet Yurts", "Hersteller PT (Algarve direkt)",
     "ca. 10.000-15.000 (Anfrage)",
     "DIREKT IM ALGARVE: Schreiner+Designerin (Nadav & Noga) bauen vor Ort, kennen das Klima persoenlich. Modern, helle Doppelglas-Tueren, modular 5-9m. Kein Versand-Risiko, kein Zoll, schnelle Wartung/Reparatur. Lokales Holz. Inkl. Plattform-Transport bei Anfrage.",
     "https://www.betyurts.com/"],

    [3, "Yourtepoque (6,5m oder 8m)", "Hersteller FR - bestes Preis-Leistung",
     "8.680 (6,5m) / 11.820 (8m)",
     "TRANSPARENTE PREISE ONLINE - Yourtepoque listet alle Preise unverbluemt. Ehrliche europ. Bauweise, Mittelpreis-Klasse, im Budget. Mit Schatten-Netz (siehe Hinweise) klimatauglich. 6,5m = 33m2 fuer 8.680 EUR ist Top-Wert.",
     "https://yourtepoque.com/"],

    [4, "Yurt Workshop Spain (Cadiar)", "Hersteller ES (Andalusien)",
     "ca. 8.000-13.000 nach Spezifikation",
     "NACHBARLAND-HERSTELLER: Rob Matthews seit 2004, baut in Alpujarras (mediterranes Klima vergleichbar Algarve). Spanisches Kastanienholz, flammenhemmendes Canvas. 3 Levels: standard/isoliert/Glaswand. Versand ES->PT guenstig (~600-1.200 EUR). Lokale Klima-Kompetenz.",
     "http://yurtworkshop.es/"],

    [5, "Jurte24 6m (Hr. Wendt)", "Hersteller DE - mongolisch, deutsche Qualitaet",
     "12.050 EUR (komplett, inkl. 19% MwSt)",
     "DEUTSCHE QUALITAET: Festpreis komplett ausgestattet, winterfest, wasserdicht, 100% Schafwoll-Filz, 170 cm Wandhoehe. Im Budget mit Spedition (~13.500 EUR). Ergaenze mit Schatten-Netz fuer PT-Sommer.",
     "https://www.jurte24.de/"],
]
style_sheet(ws4, headers4, rows4)
widths4 = [8, 28, 32, 30, 70, 50]
for i, w in enumerate(widths4, 1):
    ws4.column_dimensions[get_column_letter(i)].width = w
# Zeilenhoehe fuer Begruendungen
for r in range(2, len(rows4) + 2):
    ws4.row_dimensions[r].height = 120


# ============ Sheet 5: Hinweise ============
ws5 = wb.create_sheet("Hinweise")
ws5.column_dimensions["A"].width = 32
ws5.column_dimensions["B"].width = 100

t = ws5.cell(row=1, column=1, value="HINWEISE - Jurten-Kauf fuer Algarve, Portugal")
t.font = TITLE_FONT
ws5.merge_cells("A1:B1")

sections = [
    ("KLIMA & UV - WICHTIGSTER PUNKT!", [
        "Die Algarve-Sonne ist der groesste Feind jeder Jurte. Standardmaessige Canvas-Plane haelt nur 2-3 Jahre!",
        "Selbst schweres 12oz Canvas verschleisst durch UV-Strahlung in der portugiesischen Sommersonne.",
        "LOESUNG: Schattennetz unbedingt installieren (6m Breite, von lokalen Agrar-Kooperativen erhaeltlich).",
        "Erfahrung aus Portugal: Mit Schattennetz haelt die Original-Plane 10+ Jahre statt 2-3 Jahre.",
        "Spannseile/Riemen ebenfalls UV-anfaellig: 5-Tonnen Polypropylen-Ratschen-Riemen verwenden.",
        "Innenraum wird ohne Beschattung extrem heiss (40C+ aussen -> Jurte zu heiss zum Wohnen).",
        "Lueftungsoeffnung im Dach (Toono) und Wandhochhebevorrichtung sind essenziell fuer Hitze.",
        "Helle/weisse Aussenhuelle waehlen - keine dunklen Farben.",
    ]),
    ("IMPORTZOLL & STEUERN PORTUGAL", [
        "Standard IVA (MwSt) Portugal Festland: 23% auf alle Importe.",
        "EU-Binnenmarkt: Innerhalb EU (FR/DE/ES/AT) keine Zoll, IVA-Reverse-Charge wenn Lieferant B2B-faehig (sonst Lieferanten-MwSt).",
        "Aus Nicht-EU (UK Post-Brexit, Mongolei, China): EU-Zollkodex anwenden.",
        "Zoll fuer Zelte/Jurten (HS 6306.22, Baumwoll-Zelte oder 6306.29 synthetisch): typisch 12% Einfuhrzoll EU.",
        "IVA wird auf CIF (Cost + Insurance + Freight) + Zoll berechnet.",
        "Beispiel China-Import 5.000 USD Werk + 2.000 EUR Versand: ca. 4.600 + 2.000 = 6.600 EUR CIF; +12% Zoll = 792; +23% IVA auf 7.392 = 1.700; TOTAL ca. 9.092 EUR.",
        "Achtung: Versandsendungen ueber 150 EUR Warenwert: Verzollung obligatorisch, Spediteur erhebt zusaetzliche Bearbeitungsgebuehr (~50-100 EUR).",
        "Empfehlung: Spediteur mit EORI-Nummer beauftragen oder Customs Broker einschalten.",
    ]),
    ("GENEHMIGUNGEN ALGARVE", [
        "Zustaendig: Camara Municipal (Stadtverwaltung) der jeweiligen Gemeinde (z.B. Loule, Tavira, Silves, Lagos, Faro).",
        "WICHTIG: Mobile Strukturen (Jurten, Tipis, Wohnwagen) sind nicht den klassischen Baurechts-Vorschriften unterworfen, wenn nicht fest gegruendet (kein Betonfundament).",
        "Ohne Fundament & nicht ganzjaehrig stehend: oft problemlos, INFORMA-Antwort der Camara aber sinnvoll.",
        "Fuer rein private Nutzung (keine Vermietung) auf eigenem Grundstueck: niedrigere Anforderungen.",
        "Vermietung (Glamping/Turismo Rural): AL-Genehmigung (Alojamento Local) noetig, plus Anschluesse: Wasser/Strom/WC.",
        "Turismo Rural: ab ca. 5 Hektar, NICHT auf REN (Reserva Ecologica Nacional) oder RAN (Reserva Agricola Nacional)-geschuetztem Land.",
        "Voranfrage 'Informacao Previa': bei Camara Municipal kostenpflichtig, klaert grundsaetzliche Zulaessigkeit.",
        "Strukturen <40m2 oft ohne Baugenehmigung erlaubt (variiert pro Camara!).",
        "Algarve-Kuestennaehe: Strikter wegen POOC (Plano de Ordenamento da Orla Costeira).",
        "Empfehlung: lokalen Anwalt/Architekt fuer 200-500 EUR Erstberatung konsultieren.",
    ]),
    ("RISIKEN BEI CHINA-IMPORT (Alibaba)", [
        "Qualitaetsrisiko: Bilder vs. Realitaet weichen oft ab; Trade Assurance nur bei verifizierten Gold/Verified Suppliern.",
        "Lieferzeit: 8-16 Wochen Produktion + 5-8 Wochen Seetransport, Verzoegerungen ueblich.",
        "Versteckte Kosten: Zoll 12%, IVA 23%, Spediteurgebuehren, Haltung im Hafen 50-300 EUR/Tag bei Verzoegerung.",
        "Schaeden im Transit: Versicherung dringend empfohlen (~1-2% vom CIF-Wert).",
        "Materialprobleme: Chinesische Standard-PVC-Plane oft minder UV-stabil als angegeben, Tests vorab anfordern.",
        "Garantie/After-Sales aus China sehr schwer einklagbar - Reparatur in PT teurer.",
        "MOQ-Trick: Anbieter werben mit MOQ 1, wollen aber oft Container voll verkaufen.",
        "Empfehlung: Mind. 3 Anbieter vergleichen, Muster anfordern, Trade Assurance & Versicherung obligatorisch, Inspektionsservice (SGS/Bureau Veritas) vor Versand buchen.",
    ]),
    ("TIPPS FUER GEBRAUCHTKAUF", [
        "Hauptkriterien beim Gebrauchtkauf: Holzgitter (Khana) Zustand pruefen - kein Rost an Schrauben, kein Schimmel/Faeulnis.",
        "Filz: muss trocken, schimmelfrei, nicht moettenbefallen sein (an Ecken/Saeumen pruefen).",
        "Aussenplane: gegen Licht halten - duenne Stellen = bald durch.",
        "Tono (Dachkranz) auf Risse pruefen - schwer zu ersetzen.",
        "Komplettheit: Tueren, Fensterklappen, Spannseile, alle Filz-/Plane-Lagen vorhanden?",
        "Demontage live ansehen wenn moeglich - viele Verkaeufer haben Jurte nie selbst abgebaut.",
        "Preis-Faustregel: 50-70% des Neupreises bei <5 Jahre alt mit normaler Nutzung.",
        "Plattformen: Kleinanzeigen.de, eBay.de, OLX.pt, LeBonCoin.fr, Wallapop, Milanuncios, Facebook Marketplace, yurtworkshop.com/secondhand, yurtmarket.com",
        "In PT direkt anfragen: Casa dos Sonhos, Bet Yurts - oft Restposten oder Hersteller-Vorgaengermodelle guenstiger.",
    ]),
    ("KAUF-CHECKLISTE ALGARVE-SPEZIFISCH", [
        "1. UV-bestaendige Aussenhuelle (mind. PU/Acryl-imprägniert oder PVC) anfragen.",
        "2. Helle Farbe waehlen (weiss/creme/sand) - reduziert Innentemperatur deutlich.",
        "3. Toono mit Belueftungsklappen + Insektennetz.",
        "4. Hochhebbare Wandkrempe fuer Sommerlueftung.",
        "5. Schattennetz fuer 100-300 EUR mitbestellen oder lokal kaufen.",
        "6. 5-Tonnen Polypropylen-Ratschen-Riemen statt Standard-Spanngurte.",
        "7. Bodenplattform mit Lueftung darunter (verhindert Feuchtigkeit & Termiten).",
        "8. Termitenschutz: Holzgitter mit Borax-Loesung behandeln (PT hat Termiten!).",
        "9. Sturmverankerung: Algarve Winterstuerme bis 100 km/h moeglich (Levante/Westwinde).",
        "10. Wasserdichtigkeit: Winter PT bringt heftige Regenfaelle - Naehte muessen versiegelt sein.",
    ]),
    ("BUDGETKALKULATION (15.000 EUR Limit)", [
        "Empfohlene Verteilung:",
        " - Jurte (5-7m Durchmesser): 6.000-11.000 EUR",
        " - Versand/Spedition: 500-2.500 EUR (PT lokal 0, EU 800-2.000, China 1.800-2.500)",
        " - Zoll+IVA falls Nicht-EU: 800-2.500 EUR",
        " - Plattform/Boden (Holz, lokal in PT bauen): 800-1.500 EUR",
        " - Schattennetz + Verankerung + Zubehoer: 300-800 EUR",
        " - Reserve (Genehmigung, Berater, Versicherung): 500-1.000 EUR",
        "FAUSTREGEL: Werkpreis < 10.000 EUR um unter 15.000 EUR Gesamtbudget zu bleiben.",
    ]),
    ("VERSAND-OPTIONEN ZUSAMMENFASSUNG", [
        "Innerhalb Portugal (Bet Yurts, Casa dos Sonhos): 0-500 EUR, oft inkl.",
        "Spanien -> Portugal (LKW-Spedition): 500-1.200 EUR",
        "Frankreich -> Portugal: 800-1.800 EUR",
        "Deutschland/AT -> Portugal: 1.200-2.000 EUR",
        "UK -> Portugal (Post-Brexit): 1.500-2.500 EUR + EU-Zoll/IVA",
        "Mongolei (UB) -> Portugal: 2.000-2.500 EUR (See/Bahn)",
        "China -> Portugal (LCL Container Lisbon/Sines): 1.500-2.500 EUR",
    ]),
    ("WICHTIGE QUELLEN & WEITERFUEHRENDE LINKS", [
        "Portugiesische Yurte-Erfahrung: https://www.permaculturinginportugal.net/yurt-shading/",
        "Yurt Forum (Community): https://www.yurtforum.com/",
        "Off-Grid Portugal Community: https://florago.org/, Workaway.info",
        "Portugiesisches Baurecht: https://anwalt-portugal.de",
        "Camara Municipal Algarve - direkt anfragen je nach Standort",
        "EU-TARIC Zollkalkulator: https://taric.ec.europa.eu",
        "IVA Portugal: 23% Standard (Festland), 22% Madeira, 16% Azoren",
    ]),
]

row = 3
for title, items in sections:
    ws5.cell(row=row, column=1, value=title).font = Font(bold=True, size=12, color="2E5C8A")
    ws5.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
    row += 1
    for item in items:
        cell = ws5.cell(row=row, column=2, value=item)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws5.row_dimensions[row].height = max(20, min(80, 15 + len(item) // 4))
        row += 1
    row += 1  # Leerzeile

# Datum
ws5.cell(row=row + 1, column=1, value="Recherche-Stand:").font = Font(bold=True)
ws5.cell(row=row + 1, column=2, value="2026-05-11 - Preise koennen sich aendern. Alle Preise EUR, sofern nicht anders angegeben.")


# ============ Sheet 6: Casa dos Sonhos Detail ============
ws6 = wb.create_sheet("Casa dos Sonhos Detail")
ws6.column_dimensions["A"].width = 38
ws6.column_dimensions["B"].width = 95

title_cell = ws6.cell(row=1, column=1, value="Casa dos Sonhos Yurts - Detail-Analyse (Quelle: casadossonhos.co.uk/copy-of-environmental-policy + Facebook/Instagram)")
title_cell.font = TITLE_FONT
ws6.merge_cells("A1:B1")
ws6.row_dimensions[1].height = 28

cds_data = [
    ("Allgemein", ""),
    ("Firmenname", "Casa dos Sonhos Yurts (\"House of Dreams\")"),
    ("Standort", "Besteiros de Cima nahe Portalegre, Alto Alentejo, Portugal"),
    ("Werkstatt", "Solar-betriebene (off-grid) Werkstatt im Alto Alentejo"),
    ("Hauptseite", "https://www.casadossonhos.co.uk/"),
    ("Environmental Policy (Quelle dieses Detail-Sheets)", "https://www.casadossonhos.co.uk/copy-of-environmental-policy"),
    ("Facebook", "https://www.facebook.com/yurtscasadossonhos/"),
    ("Instagram", "https://www.instagram.com/casadossonhosyurts/"),
    ("YouTube", "https://www.youtube.com/channel/UCfenMyhVzEkX7P-lz2hjS6A"),
    ("Schwester-Projekt", "https://casadossonhos.org/ (Wollverarbeitung in Planung)"),

    ("", ""),
    ("Verfuegbare Groessen", ""),
    ("5 m Durchmesser", "20 m2 Innenflaeche"),
    ("6,1 m (20 foot) Durchmesser", "30 m2 Innenflaeche"),
    ("7,3 m (24 foot) Durchmesser", "42 m2 Innenflaeche"),
    ("9,1 m (30 foot) Durchmesser", "64 m2 Innenflaeche"),

    ("", ""),
    ("Bauart & Konstruktion", ""),
    ("Selbsttragend", "ALLE Modelle ohne Mittelsaeule / -stuetze - max. offener Innenraum"),
    ("Wandhoehe", "ueber 2 m (mehr als Standard-Mongolisch)"),
    ("Hoehe Mittelring (Toono)", "min. 3,5 m - sehr geraeumig, geeignet fuer Hochbett/Galerie"),
    ("Fertigung", "Handgefertigt vom Hersteller-Paar in Alentejo"),

    ("", ""),
    ("Standardausstattung (im Basispreis enthalten)", ""),
    ("Fenster", "1x1 m doppelverglastes Aluminium-Fenster (beim 9,1m-Modell: 120x120 cm)"),
    ("Tuer", "Holztuer mit kleinem Fenster und Holzladen"),
    ("Daemmung", "Moderne Daemmung inklusive"),
    ("Dachoeffnung", "Oeffnungs-faehiges, klares Kuppel-Oberlicht (\"opening clear domed skylight\")"),
    ("Aussenhuelle", "Speziell fuer portugiesisches Wetter konzipierte Plane mit 10-J. UV-Garantie"),
    ("Lieferung & Aufbau", "IM PREIS ENTHALTEN innerhalb Festland-Portugal"),

    ("", ""),
    ("Materialien (Portuguese sourced)", ""),
    ("Holz", "Portugiesisches Holz fuer Tueren, Khana-Lattengitter, Sparren, Fensterrahmen"),
    ("Bodenbelag", "Portugiesische Korkboeden (regional, oeko)"),
    ("Daemmung", "Portugiesische Korkdaemmung"),
    ("Wolle", "Aktuell zugekauft; eigenes Wollverarbeitungs-Projekt geplant (Alentejo-Schafwolle)"),
    ("Aussenhuelle-Qualitaet", "Premium: 10-J. UV-Garantie, nie Faeulnis, nie Schimmel, KEINE Nachbehandlung noetig"),

    ("", ""),
    ("Optionale Zusatzausstattung", ""),
    ("Groessere Fenster", "Anstelle 1x1m Standard"),
    ("High Wind Kit / Sturm-Kit", "Verstaerkte Verankerung fuer stuermische Regionen (EMPFEHLUNG fuer Algarve-Kueste!)"),
    ("Zusaetzliche Daemmlagen", "Fuer kuehlere Lagen / Winterwohnen"),
    ("Zusaetzliche Tueren/Fenster", "Beliebig"),
    ("Franz. Doppeltueren", "3/4-Verglasung mit Holzladen (\"French double doors with 3/4 glass and shutters\")"),
    ("Kaminzug (Flue Kit)", "Vorbereitung fuer Holzofen / Wood Burner"),
    ("Mueckennetze (Mosquito Screens)", "WICHTIG fuer Algarve-Sommer"),

    ("", ""),
    ("Garantie & Klima-Eignung", ""),
    ("UV-Garantie Aussenhuelle", "10 JAHRE - hoechste am Markt nachgewiesene Garantie"),
    ("Schutzklassen", "Plane rottet nicht, schimmelt nicht, braucht keine UV-Nachbehandlungen"),
    ("Geeignet fuer", "Vollstaendig auf das portugiesische Klima (heisse Sommer, feuchte Winter) ausgelegt"),
    ("Empfehlung Algarve", "Mit Sturm-Kit + Mueckennetz + ggf. Extra-Daemmung -> ideal fuer Algarve"),

    ("", ""),
    ("Preis & Kommerzielles", ""),
    ("Preisliste", "Nicht oeffentlich. Anfrage per E-Mail / Facebook erforderlich."),
    ("Schaetzung 5m (20 m2)", "ca. 8.000-10.500 EUR inkl. Aufbau (vergleichbar mit Yourtepoque-Niveau + Local-Bonus)"),
    ("Schaetzung 6,1m (30 m2)", "ca. 9.500-12.500 EUR inkl. Aufbau"),
    ("Schaetzung 7,3m (42 m2)", "ca. 11.500-14.500 EUR inkl. Aufbau (an Budget-Grenze)"),
    ("Schaetzung 9,1m (64 m2)", "ca. 15.000-20.000+ EUR (ueber Budget)"),
    ("Lieferung & Aufbau", "Inklusive in Festland-PT"),
    ("Import-Zoll/IVA", "ENTFAELLT (lokale PT-Produktion)"),
    ("Zahlung", "Direkt mit Hersteller, vermutlich Anzahlung + Restzahlung bei Abnahme"),

    ("", ""),
    ("Nachhaltigkeit / ESG", ""),
    ("Energie Produktion", "100% Solar in der Werkstatt"),
    ("Materialien-Herkunft", "Portugiesisch (Holz, Kork)"),
    ("Soziales Engagement", "Plant Wollverarbeitung um lokale Schafwolle aufzuwerten (Alentejo) - aktuell unterbewertet"),
    ("Kreislaufwirtschaft", "Natuerliche Materialien, kein PVC im Standard"),

    ("", ""),
    ("Risiken & Hinweise", ""),
    ("Lieferzeit", "Handarbeit -> mehrere Monate Vorlauf einplanen"),
    ("Skalierung", "Kleines Familienunternehmen - Reaktionszeit kann variieren"),
    ("Preis intransparent", "Keine oeffentliche Preisliste - immer schriftliches Angebot einholen + Optionen klar definieren"),
    ("Sturmkomponente", "Algarve-Kueste: Sturm-Kit OBLIGATORISCH bestellen"),
    ("Hitze Innenraum", "Trotz 10-J. Plane: zusaetzliches Schattennetz/Pergola fuer 40C+ Sommertage empfehlenswert"),

    ("", ""),
    ("Fazit fuer Algarve", ""),
    ("Bewertung", "TOP-EMPFEHLUNG #1 - einzigartige 10-J. UV-Garantie, lokale Produktion, Lieferung+Aufbau inkl., portugiesische Materialien, kein Import-Risiko."),
    ("Naechster Schritt", "Per Facebook Messenger / E-Mail Angebot fuer 5m oder 6,1m anfragen, mit Sturm-Kit, Mueckennetz, Kaminzug, ggf. franz. Doppeltueren."),
]

row = 2
for label, val in cds_data:
    if label and not val:
        # Sektions-Ueberschrift
        c = ws6.cell(row=row, column=1, value=label)
        c.font = Font(bold=True, size=12, color="2E5C8A")
        ws6.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
        ws6.row_dimensions[row].height = 22
    elif not label and not val:
        # Leerzeile
        ws6.row_dimensions[row].height = 8
    else:
        c1 = ws6.cell(row=row, column=1, value=label)
        c1.font = Font(bold=True)
        c1.alignment = Alignment(wrap_text=True, vertical="top")
        c1.border = BORDER
        c2 = ws6.cell(row=row, column=2, value=val)
        c2.alignment = Alignment(wrap_text=True, vertical="top")
        c2.border = BORDER
        ws6.row_dimensions[row].height = max(20, min(60, 15 + len(val) // 6))
    row += 1

ws6.freeze_panes = "A2"


# Speichern
output_path = "/home/user/Jurte/Jurten_Recherche_Algarve.xlsx"
wb.save(output_path)
print(f"XLSX gespeichert: {output_path}")
