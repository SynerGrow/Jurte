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

    ["Bet Yurts (Algarve - Noga Shimshon)", "Portugal (Algarve, Aljezur)", "https://www.betyurts.com/",
     "6 m=28m2 / 7 m=38,5m2 / 9 m=63,5m2 (Quote-PDFs Mai 2026)",
     "Jurte netto: 6m=12.600 / 7m=14.400 / 9m=19.800 EUR. Plattform: 6m=6.500 / 7m=7.200 / 9m=11.500 EUR. Zus. Tueren/Fenster: 800 EUR/Stk. Transport zur Aljezur: 6m=700 / 7m=900 / 9m=1.500. Installation: 6m=1.900 / 7m=2.500 / 9m=2.900. Alle Preise zzgl. 23% IVA.",
     "Quote-PDF nennt nur 'Nordic Pine roof rafters + lattice walls', 'Outer canvas with stove outlet', '80mm Insulation material', 'Inner cotton linen', '1.20m central ring'. KEINE Hersteller-/Markennennung trotz Marketing 'marine-grade' auf Website.",
     "Mittel-Gut: Pine-Rahmen (NICHT Laerche/Kastanie!). 'Outer canvas' ohne Spezifikation. NUR 2 JAHRE GARANTIE auf Plane + 1 J. Verarbeitung = NICHT marine-grade-Niveau (Sauleda waere 5-10 J. Garantie).",
     "Inland Algarve, Transport+Aufbau gegen Aufpreis ausgewiesen. Vorgefertigte gedaemmte Plattform mit Erdschrauben + Vinyl-Top-Flooring (kann selbst gebaut werden, dann Plattform-Preis sparen).",
     "NIF 307460320 (Empresario em Nome Individual, registriert auf Noga Shimshon). Bank: Privatkonto Noga Shimshon. ZAHLUNG: 50% NICHT-RUECKZAHLBAR upfront, 50% vor Lieferung. AUGUST 2026 SLOT verfuegbar. Werkstatt-Besuch nach 3 Wochen Reise. Kontakt: noga@betyurts.com / +351 961 592 792",
     "6m all-in (2 Fenster+Plattform+Transport+Install): ~28.060 / 7m: ~31.940 / 9m: ~44.870 EUR. DOPPELT SO TEUER WIE CASA DOS SONHOS bei kuerzerer Garantie."],

    ["Casa dos Sonhos Yurts (Sarah & Vladimir, Alegrete)",
     "Portugal (Alto Alentejo, Besteiros de Cima 7300-320 Alegrete)",
     "https://www.casadossonhos.co.uk/  |  https://www.facebook.com/yurtscasadossonhos/",
     "5m=20m2 / 6,1m=30m2 / 7,3m=42m2 / 9,1m=64m2",
     "5m 7.900 / 6,1m 9.700 / 7,3m 15.750 / 9,1m 18.000 EUR (delivered+erected PT). Konkretes Quote Mai 2026: 6,1m + 2.Tuer 850 + Flue-Kit 100 = 10.650 EUR all-in.",
     "Portugiesisches Holz (Tueren, Khana-Gitter, Sparren, Fenster), portugiesische Korkboeden + Korkdaemmung, moderne Isolierung, Aussenhuelle mit 10-J. UV-Garantie",
     "EXZELLENT - 10 Jahre UV-Garantie auf Aussenhuelle SCHRIFTLICH, Plane rottet/schimmelt nicht, braucht keine Nachbehandlung; speziell fuer portugiesisches Wetter gefertigt",
     "Innerhalb Festland-PT INKL. Lieferung+Aufbau im Preis. Plattform NICHT enthalten (Kunde stellt 'a base').",
     "Inhaber: Sarah (UK-Englisch) + Vladimir, Tochter 3J, leben in Alegrete. NIPC 260282812 (Empresario em Nome Individual - persoenlicher NIF als Geschaefts-NIPC). Slot Oktober 2026 wurde 15.05.2026 an anderen Kunden vergeben. Naechster Slot: 2027. Zahlung: 1.000 EUR Anzahlung + Restbetrag 2 Monate vor Lieferung (3 Monate beim 9m). Refs: Brett bei Quinta Glamping (https://www.quintaglamping.com), Paula Young verkauft gebrauchte 7,3m (FB 'Paula Vegan Chef')",
     "5m: 7.900 / 6,1m: 9.700 / 7,3m: 15.750 / 9,1m: 18.000 EUR alles inkl. PT-Lieferung+Aufbau. ACHTUNG: 6,1m all-in mit Plattform DIY 4.000 + Schattennetz 240 = 14.890 EUR. STATUS: Oktober 2026 SLOT WEG (15.05.2026), nur 2027 verfuegbar."],

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

    ["Yurt Made in Portugal (Loures, ueber OLX)", "Portugal (Loures, Lisboa)",
     "https://www.olx.pt/d/anuncio/yurt-made-in-portugal-IDIyzPL.html",
     "5m=20m2 / 6m=28m2 / 7-8m auf Anfrage",
     "5m 6.900 / 6m 8.400 / 7-8m Anfrage",
     "Holz + Filz + Plane (modern), portugiesisch handgefertigt",
     "Sehr gut - lokaler PT-Hersteller seit 2020, kein Import-Zoll/IVA",
     "Innerhalb PT - Lieferung+Aufbau EXTRA (standortabhaengig)",
     "Neuer kleiner PT-Hersteller bei Lissabon, 'modern take on ancient living'. OLX-Praesenz seit 2020.",
     "8.400-10.500 inkl. Aufbau (6m)"],

    ["Atilla es a Fehernep (jurtak.hu)", "Ungarn", "https://jurtak.hu/  |  https://fehernep.hu/",
     "4 / 6 / 8 / 10 m",
     "8m Rahmen 6.600 + 4 Fenster + Kuppel = 9.410 netto / +Plane (Sauleda+Airtex) 2.200 / +Versand PT 3.500",
     "Rahmen Buche (Standard) oder Eiche/Laerche (Upgrade), Rest Sibirische Kiefer lackiert (Sebastian wuenscht Oel-Finish), Plane Sauleda PVC Dach + Airtex Acryl Wand, Standard-Daemmung Spiegelfolie+Vatelin (fuer Algarve NICHT empfohlen - Sebastian sourct Wollfilz+Tyvek selbst)",
     "Mittel: Erfahrung ES/BE/DE/PL aber NICHT explizit Atlantik-Kueste. Buche-Standard fuer humide Algarve ungeeignet -> Holz-Upgrade obligatorisch.",
     "Versand HU->Algarve 3.500 EUR (hoch). EU-USt-IdNr noetig fuer steuerfreie IC-Lieferung (sonst 27% MwSt HU). Lieferzeit 1-2 Monate. Foundation-Plaene inkl.",
     "Atilla Berki, Jurtenbauer seit 1996, erster kommerzieller HU-Hersteller. Sozial: Rahmen aus Therapieprogramm fuer Drogenabhaengige/Jugendstraftaeter. Ehrlich (sagt selbst Airtex nicht atmungsaktiv). Kontakt: jurtak@gmail.com / +36 30 657 0740",
     "8m mit Eichen-Upgrade+Premium-Daemmung+Versand: 17.700-20.500 netto (UEBER 15k Budget)"],

    ["Les Tentes d'Avalon (vivrelayourte.fr) - NICHT Yourtepoque verwechseln!",
     "Frankreich (Val-du-Faby, Aude, Sued-FR)",
     "https://vivrelayourte.fr/",
     "Modelle Marco Polo (1 Tuer) / Da Vinci (2 Tueren) in 7m / 8m",
     "7m Marco Polo 19.370 / 7m Da Vinci 20.840 / 8m Marco Polo 22.100 / 8m Da Vinci 23.590 EUR (jeweils ohne Plattform/Versand)",
     "Laerchen-Rahmen, Schafwoll-Filz, gedaemmt, Innenliner inklusive, oeffenbare Plexikuppel separat 1.900",
     "Mittel - franz. Hersteller, keine UV-Garantie ausgewiesen, KEINE decennale-Bauhaftung",
     "Versand Val-du-Faby->Algarve ca. 1.500 km, 1,40 EUR/km hin+rueck = ca. 4.200 EUR. Aide-au-montage 450-550 EUR (nur 2 Experten, 6-8 Personen Selbsthilfe noetig).",
     "PREMIUM-Segment: Plancher Isole (Holz+55mm Kork) 6.250 (7m) / 9.000 (8m). 2/3-Anzahlung bei Bestellung (sehr hoch, branchenunueblich). Lieferzeit 2-4 Monate. Garantie nur 1 J. versteckte Maengel.",
     "7m Gesamt mit Plattform+Versand+Montage: ~33.000-34.500 EUR / 8m: ~38.600-40.000 EUR (DEUTLICH UEBER BUDGET)"],

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

    # === NEUE ANBIETER aus Tiefen-Recherche Mai 2026 ===
    ["Portugal Yurts (Yonatan / Natural Habitat)",
     "Portugal (Centro)", "https://portugalyurts.com/",
     "Anfrage", "Anfrage",
     "Natural Building Approach: Holz + Lehm + Stroh + lokale Materialien",
     "Hoch fuer humide Algarve - natuerliche Baustoffe, lokal beschafft, ca. 20 J. Erfahrung Yonatan in EU",
     "Innerhalb PT - kein Import-Zoll",
     "PT-LOKAL-OPTION #3 nach Casa und Bēt. Gruender Yonatan hat 'Natural Habitat'-Hintergrund. Eigenes Tradesmen-Team von Kauf bis Komplettinstallation. Keine oeffentlichen Preise - Direktanfrage erforderlich.",
     "Anfrage (Schaetzung 12.000-18.000 EUR fuer 6-7m)"],

    ["YourTent.com (Praha CZ)",
     "Tschechien (Praha)", "https://www.yourtent.com/  |  https://cz.yourtent.com/cenik",
     "4-8 m, individuell konfigurierbar", "Online-Kalkulator",
     "Laerchenholz-Rahmen, 50mm Schafwoll-Filz, 180cm Standard-Wandhoehe, franz. Doppeltuer Laerche, Skylight 150/180cm zweischichtig oeffenbar, Ofen-Vorbereitung",
     "Sehr gut - 5 JAHRE PLANE-GARANTIE (60 Monate fabric warranty), Mediterranean Track Record bis Kanarische Inseln + Polarkreis",
     "CZ->PT 1.500-2.500 EUR Spedition",
     "150 Jurten in 12 Jahren gebaut. BESTE LAERCHE-WAHL fuer Algarve nach Casa. Vergleichbar mit Casa-Qualitaet aber laengere Reise. 5-J. Garantie schlaegt Bēt (2 J.).",
     "Anfrage (Schaetzung 12.000-16.000 EUR 6m all-in inkl. Versand)"],

    ["ZenYurts Belgium",
     "Belgien (Vlaanderen)", "https://zenyurts.be/en/",
     "5,80m / 6,80m=35m2 / 7,80m=47m2", "Online transparent",
     "Schafwoll-Filz 15mm + Baumwoll-Canvas UV-impraegniert, UV-Holz-Oel-Behandlung, UV-Plexikuppel",
     "Mittel - mongolisch-traditioneller Stil, Plane-Garantie unklar",
     "BE->PT 1.500-2.000 EUR Spedition",
     "Sehr transparente Online-Preise mit SKU-Codes (z.B. SW10011.1). 7,80m mit Setup+Plattform = ca. 15.800 EUR. Halb so teuer wie Bēt Yurts bei aehnlicher Spec.",
     "ca. 14.000-17.000 EUR 6,80m all-in inkl. Versand+Sturm-Kit"],

    ["Yurts4ever (Y4E, Transylvania)",
     "Rumaenien (Siebenbuergen)", "https://yurts4ever.ro/",
     "5-8m typisch", "Anfrage",
     "Handgefertigt, traditionelle mongolische Bauweise, ca. 1 Monat Produktion",
     "Mittel - kontinentaler Klimaerfahrung, mediterrane Tracks (Italien)",
     "RO->PT 1.500-2.000 EUR Spedition",
     "Verkauft in RO, HU, DK, IT. Adorján-Klasse Preisniveau erwartbar bei Handwerksqualitaet. Neuer Player.",
     "Anfrage (Schaetzung 8.000-12.000 EUR 6m all-in)"],

    ["Yurts Hellas (Athens)",
     "Griechenland", "https://www.yurts-hellas.gr/",
     "diverse 4-8m", "Anfrage",
     "Mediterranean climate-experienced",
     "Sehr gut - direkte mediterrane Klima-Erfahrung",
     "GR->PT 1.800-2.500 EUR Spedition (See/Sammelladung)",
     "Telefon +30 210 8317431. Konkrete Preise via Anfrage. Versand-Kosten sind der Nachteil vs PT-lokal.",
     "Anfrage"],

    ["Tiny Homes Greece",
     "Griechenland", "https://www.tinyhomes.gr/",
     "diverse", "Anfrage",
     "Mediterran, Erfahrung GR+UK",
     "Sehr gut - 17 Jahre Erfahrung Mittelmeerklima",
     "GR->PT teuer (1.800-2.500 EUR)",
     "Verkauft Jurten fuer GR und UK Markt seit 17 J. Versand-Kostennachteil vs PT-Lokal.",
     "Anfrage"],

    ["Streetgy (AT vertreibt CZ)",
     "Tschechien / vertrieb Oesterreich",
     "https://streetgy.at/jurte-kaufen-und-erfahrungen/",
     "Diverse", "Online-Preisliste pro m2",
     "Mongolische Jurten via CZ-Hersteller",
     "Mittel - marketing 'bis zu 30% guenstiger' (gegenueber wem?)",
     "CZ->PT 1.500-2.500 EUR",
     "Bewerben sich als guenstige Alternative. Preisliste online vorhanden. Garantie unklar. Anfrage lohnt sich evtl.",
     "Anfrage"],

    ["Mongolian Magic (Niederlande)",
     "Niederlande", "https://mongolianmagic.com/mongolian-magic-ger-yurt/?lang=en",
     "diverse 5-8m", "Anfrage",
     "Direktimport aus Mongolei mit EU-Anpassung",
     "Mittel - mongolische Bauweise, Anpassung an EU-Klima",
     "NL->PT 1.800-2.200 EUR",
     "Aehnlich Silk Road Yurts. Lokale Service-Vertretung in NL.",
     "Anfrage"],

    ["Yin Yang Yurt (Belgium)",
     "Belgien (Brasschaat / Rochefort)", "https://yinyangyurt.com/en/",
     "diverse mongolisch", "Anfrage",
     "Mongolische Jurten, Daemmung gegen Kaelte+Naesse",
     "Mittel - europ. Wettererfahrung",
     "BE->PT ca. 1.500-2.000 EUR",
     "EU-angepasste mongolische Jurten. Weniger transparent als ZenYurts.",
     "Anfrage"],

    ["Green World (Ukraine, Liefert PT)",
     "Ukraine, Vertrieb weltweit", "https://greenworld.house/pt/yurts/",
     "24 / 50 / 56 / 78 / 113 / 254 m2 (50m2 = 8m Durchmesser)",
     "Anfrage (Preisliste ab 01.01.2024)",
     "Vorgefertigter Holz-Frame, modular, durchnummeriert, mit Anleitung",
     "10 JAHRE STRUKTUR-GARANTIE + 12 Mon. versteckte Maengel",
     "Lieferung in 9 Laender weltweit, auch PT. Versand ab UA per Spedition.",
     "VERDACHT: moeglicher Verkaeufer hinter OLX-Listings IDJatvu + IDJbO7n (beide 8m). Produktion 2 Monate. Plane-Spec unklar (vermutlich Standard, nicht marine-grade). Ukraine-Hersteller -> Versand- und politisches Risiko abklaeren. 8m / 50m2 koennte sehr wettbewerbsfaehig sein wenn Plane Algarve-tauglich ist.",
     "Anfrage (Schaetzung 8m/50m2 ca. 8.000-14.000 EUR + Versand)"],

    ["Woodspot Solutions (PT)",
     "Portugal", "https://www.woodspotsolutions.com/pt/yurts-pt",
     "diverse", "Anfrage",
     "Inspiriert traditionellem Design, modernisiert fuer mehr Komfort",
     "Mittel - PT-lokal aber wenig Info im Netz",
     "Innerhalb PT - kein Import-Zoll",
     "Weiterer PT-Anbieter, weniger sichtbar als Casa/Bēt/Yonatan. Direktanfrage erforderlich.",
     "Anfrage"],

    ["Celtic Yurts (Galicia, ES)",
     "Spanien (Galicia, NW Iberia)", "https://www.celticyurts.com/en",
     "diverse", "Preisliste auf Website",
     "Hochwertige Materialien, Strapazierfaehigkeit",
     "Mittel - Galicien-Klima vergleichbar Nord-PT/Atlantikkueste",
     "ES->PT relativ guenstig 500-1.500 EUR (Galicien grenzt an PT)",
     "Optional Selbstbau-Workshop (2-3 Wochen Werkstatt + fertige Jurte zum Mitnehmen).",
     "Anfrage"],

    ["Dani's Community-Projekt 7m Custom (Castelo de Vide, Alentejo)",
     "Portugal - Castelo de Vide (Alto Alentejo, ca. 400 km von Aljezur)",
     "7m Innenflaeche ca. 38 m2, Mittelhoehe 360 cm, Dachneigung 28 Grad",
     "16.000 EUR inkl. Aufbau (Transport+Plattform extra). UPGRADES MOEGLICH: Sauleda-Plane +1.100 EUR Material, 8cm Wolle +400 EUR Material. NICHT moeglich: zusaetzliche Fenster (Waende schon fertig).",
     "Khaana: PINIE-Scherengitter, 200cm Wandhoehe. Dachring (Tono): DOUGLASIE. Dachkuppel 150cm doppeltes Plexiglas (stufenlos oeffenbar, windfest). Tueren: 2 Hartholz-Doppeltueren 155x200cm zweifach verglast (= 4,8 m² Glasflaeche durch Tueren, plus 1,77 m² Skylight). KEINE zusaetzlichen Wandfenster moeglich. Daemmung: Standard 4cm Schafwolle (8cm Upgrade moeglich wenn Sebastian Material besorgt - Schafwolle PT-Bezug schwierig, Kontakt 'Chris' via Dani-Vermittlung). Innenstoff: 100% Biobaumwolltuch. Aussenstoff: BW/Polyester 50/50 Canvas Standard ueber ESVO NL (1.400 EUR fuer 7m); ALTERNATIVE STOFFE wenn unter 420g/m² (Sauleda Solar Pro 290-340g/m² oder Tempotest Marine ✓) - Sebastian besorgt selbst, Dani naeht.",
     "VERHANDELBAR: Dani ist sehr flexibel - naeht gerne anderen Stoff bis 420g/m² Maschinen-Limit, macht gerne dickere Daemmung wenn Sebastian Material liefert. SAULEDA SOLAR PRO ist langfristig guenstiger als 5-Jahres-Canvas-Tausch.",
     "Transport Castelo de Vide -> Aljezur: 800-1.500 EUR. Aufbau IM 16k-Preis. Plattform bauen sie KEINE (Sebastian baut lokal: 2.880 EUR Kork-Variante, siehe Sheet 17).",
     "Telegram Direkt-Kontakt 'Dani' (Gruppe Jurten Portugal). RECHTSFORM: AKTUELL Privatverkauf an Freunde (Verein in Gruendung). Zahlung: Anzahlung als Commitment + flexible Teilzahlungen, ueberweisen ODER PAYPAL (= PayPal-Kaeuferschutz potentiell nutzbar). Kontakt Chris (PT-Schafwolle) via Dani vermittelbar.",
     "TOP-EMPFEHLUNG WENN UPGRADES MACHBAR: Dani 16k + Sauleda-Upgrade 1.100 + 8cm-Wolle 400 + Plattform Kork 2.880 + Transport 1.000 + Schattennetz 240 = ~21.620 EUR all-in. Spar gegenueber Canvas-5J-Tausch ueber 10J ca. 1.200 EUR. PLUS: 4 Familien-Community vor Ort fuer Sebastians Familie sehr willkommen. PAYPAL-OPTION fuer Anzahlung = wichtiges Sicherheitsnetz fuer Privatverkauf."],
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
    # === Paula Young - TOP-LEAD ALGARVE (Casa-Quote-Empfehlung Mai 2026) ===
    ["Paula Young (FB 'Paula Vegan Chef') - gebrauchte Casa-Jurte 7,3m",
     "Facebook (Direktverkauf)", "Algarve / Monchique-Umgebung",
     "42 m² (7,3 m / 24 ft Casa dos Sonhos)",
     "Anfrage (Schaetzung 8.000-11.000 EUR = 50-70% vom Neupreis 15.750)",
     "Gebraucht direkt vom Eigentuemer",
     "Facebook-Suche 'Paula Vegan Chef' - kein Direktlink. Sarah von Casa kann Einfuehrung machen.",
     "Inland Algarve, Transport ca. 500-1.000 EUR",
     "TOP-EMPFEHLUNG seit Casa-Slot weg: gebrauchte Casa-Qualitaet, physisch verifizierbar, sofortige Verfuegbarkeit, kein Vorauszahlungs-Risiko. Wichtige Fragen: Alter? Zustand Plane/Filz? Grund Verkauf? Algarve-Erfahrung mit Sommer/Winter."],

    # === PT-Gebrauchtmarkt-Tiefenscan Mai 2026 ===
    ["PT-Marktanalyse: Kategorie 1-5 Uebersicht", "Multi-Plattform", "Portugal komplett",
     "Marktreport", "Marktbeobachtung",
     "Kategorisiert nach 5 Kategorien (Sofortkauf / Anfrage lohnt / Vorsicht / Plattform-Uebersicht / Spezielle Quellen)",
     "Siehe einzelne Eintraege",
     "ANALYSE: PT-Used-Markt klein - 3-5 aktive >=6m Listings gleichzeitig. Konzentration Mafra/Lissabon, kaum Algarve-Lokal. Off-OLX-Leads (Sarah-Vermittlung) wertvoller als oeffentliche Marktplaetze. Empfehlung: OLX-Daily-Alert + FB-Gruppe Algarve 'Wanted'-Post + Sarah um weitere Vermittler bitten."],

    ["FB Gruppe 'For Sale or Swap in the Algarve'",
     "facebook.com/groups", "Algarve - lokale Expat-Community",
     "Variabel", "Variabel",
     "Variabel - meist neu/gebraucht aus Expat-Besitz",
     "https://www.facebook.com/groups/FORSALEORSWAPINTHEALGARVE/",
     "Lokal Algarve - kein Versand",
     "Eigenen 'Wanted: 6-8m yurt' Post aufgeben. Beste Chance auf lokale Algarve-Funde direkt. Gross genug fuer Sichtbarkeit, lokal genug fuer Transport-Vorteil."],

    ["CustoJusto.pt", "custojusto.pt", "Portugal weit",
     "diverse", "selten - meist <500 EUR",
     "Meist Spielzelte oder kleine Hobby-Jurten",
     "https://www.custojusto.pt/portugal?o=1&q=yurt",
     "Inland PT",
     "Sekundaerplattform, deutlich weniger Listings als OLX. 1-2 Listings typisch. Beobachten alle 2-4 Wochen."],


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

    ["OLX Portugal yurts (Kategorie-Uebersicht)", "olx.pt", "Portugal",
     "diverse", "variabel", "Neu/gebraucht",
     "https://www.olx.pt/moveis-casa-e-jardim/jardim-e-bricolage/q-yurts/",
     "Ja - innerhalb PT",
     "INLAND PT - keine Zoll, kein Auslandsversand"],

    ["OLX PT: Yurts mongois (Igreja Nova/Cheleiros, Mafra)", "olx.pt", "PT - Mafra (bei Lissabon)",
     "28 m2 (6 m Durchmesser)", 5000, "Gebraucht / Restposten",
     "https://www.olx.pt/d/anuncio/yurts-mongis-mongolian-yurts-IDHQ8sJ.html",
     "Ja - inland PT, Transport+Aufbau gegen Aufpreis",
     "BESTPREIS 6m+ auf OLX! Filzdecke (warm/kuehl), mobile Aufbau auf jeder ebenen Flaeche. Optional Pelletofen gegen Aufpreis. Direkt vom Verkaeufer."],

    ["OLX PT: Tenda Yurt Original Quirguistao (Palmela)", "olx.pt", "PT - Palmela (Setubal)",
     "28 m2 (ca. 6 m Durchmesser, 30 Personen)", "Anfrage",
     "Neu/handgefertigt (Originalimport Issyk-Kul Quirguistao)",
     "https://www.olx.pt/d/anuncio/tenda-yurt-original-do-quirguisto-yurt-tent-original-from-kyrgystan-IDHEwz9.html",
     "Ja - PT-inland, Versand EXTRA (Region Lissabon)",
     "Marke 'Yourta', Modell 'Serenity'. Komplette Originalausstattung: Schilf-Geflecht, Wolle/Filz, Shyrdak-Teppich, Baskur-Band, Korpe-Kissen, geschnitzte Holztuer. Hand-Applikation. Keine Baugenehmigung noetig."],

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
    # NEU sortiert nach OLX-Fund Yury 8m 15k (Mai 2026) - Sebastian's Wunschgroesse 8m
    [1, "Dani Castelo de Vide 7m + Sauleda-Upgrade + 8cm Schafwolle + Kork-Plattform",
     "Custom-Bau PT (Community-Projekt) - flexibel verhandelbar",
     "~21.620 EUR all-in (16k Dani + 1.100 Sauleda + 400 Wolle 8cm + 2.880 Plattform + 1.000 Transport + 240 Schatten)",
     "NEUE TOP-EMPFEHLUNG seit Dani 23.05.2026: Maximum-flexibler Custom-Bau bei einer 8-Jahre-erfahrenen Familie in PT. Spec aushandelbar: Sauleda Solar Pro (290-340g/m², passt unter 420g/m² Maschinen-Limit) statt Standard-Canvas + 8cm Schafwolle statt 4cm (Sebastian besorgt selbst). Probeschlafen-Option, Community-Kontakt fuer Sebastians Familie. Casa-Wartezeit 2027 entfaellt, Lieferung bereits Aug/Sep 2026. STRUKTUR: 200cm Wand, 360cm Mittelhoehe, 150cm Tono Douglasie, 28° Dachneigung, 2 Hartholz-Doppeltueren 155x200cm. Standort Castelo de Vide ca. 400km - Transport ueberschaubar.",
     "https://t.me/ (Telegram Direktkontakt 'Dani' in Gruppe 'Jurten Portugal')"],

    [2, "Mikael/Bento Pacific Yurts 8m (Gouveia) = OLX 'Yury'",
     "Pacific Yurts USA Modell 'European Modern' - NEU mit Maengeln",
     "~22.040 EUR all-in (15k Pacific + 700 AC4 + 3.300 Plattform Premium + 600 Zusatz-Daemmung + 1k Versand + 500 Termiten + 200 Schrauben + 240 Schatten + 500 Reserve)",
     "STARK aber TEURER ALS DANI: 50 m² Pacific Yurts Premium-Marke, Mikael=Bento (1 Person 2 Channels: Telegram+OLX). EHRLICHER Verkaeufer (Rost offen kommuniziert). PROBLEME: 1) Schrauben-Rost - Komplett-Ersatz Edelstahl A4 (~200 EUR + 30 Std Arbeit). 2) NUR 7mm Schafwoll-Filz (R~0,18) - 5x DUENNER als Dani's 4cm. Zusatz-Daemmung in Algarve obligatorisch (+600 EUR). 3) Pinie ohne Insektizid = Termiten-Behandlung selbst (+500 EUR + 2 Tage). PLATTFORM nicht inklusive (Sebastian baut). AC4-Boden 700 EUR optional fair. Konstruktion: traditionelle Khaana, hochpraezise gefertigt. 1 Doppeltuer + 2 Glas-Fenster integriert.",
     "https://www.olx.pt/d/anuncio/yurt-de-8-metros-de-dimetro-com-janela-panormica-IDJatvu.html | Telegram 'Mickaël'"],

    [3, "DIY Modify Szenario A (Adorjan-Rahmen + Algarve-Sauleda-Sattler)",
     "DIY-Hybrid - lokaler PT-Sattler + HU-Rahmen",
     "ca. 10.300-11.000 EUR all-in (6m)",
     "NEUER SPITZENREITER seit Casa-Slot weg: Adorjan 6m Rahmen 1.750 + Versand 1.200 + Sauleda Marine-Acryl bei lokalem Algarve-Sattler 1.800-2.500 + PET-Daemmung 500 (Tobias Tumfart AT, 8,16€/m²) + Solitex 500 + Plattform DIY 3.000 + Sturm-Kit 400 + Schattennetz 240 + Reserve 500. PT-LOKALER SATTLER (z.B. Dune Algarve Sailmakers in Vilamoura) ist der Game-Changer: Sauleda Solar Pro Acryl ist objektiv BESSER als unspezifiziertes Casa-Canvas. 4.000-5.000 EUR Ersparnis vs Casa Wartezeit 2027. Sofortige Verfuegbarkeit. Volle Materialkontrolle. Risiko: Selbstaufbau-Skill, Maßabstimmung Rahmen->Plane, keine Komplett-Garantie.",
     "https://adorjan-jurta.hu/  |  https://www.dunesailmakers.com/"],

    [4, "Paula Young - gebrauchte Casa-Jurte 7,3m",
     "Used-Markt (Algarve direkt)",
     "Schaetzung 8.000-11.000 EUR + Transport 1.000",
     "ALGARVE-LEAD von Casa-Sarah selbst empfohlen: Paula Young (FB 'Paula Vegan Chef') verkauft 7,3m Casa-Jurte (42m²). Gebrauchte Casa-Qualitaet mit Original-Garantie-Resten. PHYSISCH VERIFIZIERBAR vor Kauf, kein Vorauszahlungs-Risiko, sofort verfuegbar, GROESSER als Casa 6,1m. Wichtige Pruefpunkte: Alter Plane (10-J. Casa-Garantie laeuft mit), Zustand Filz/Holz, Grund Verkauf, Algarve-Praxiserfahrung. Kontakt ueber Facebook-Suche oder Einfuehrung via Sarah Casa.",
     "Facebook: 'Paula Vegan Chef' / Sarah Casa als Vermittler"],

    [5, "Yurt Made in Portugal (Loures, OLX-Hersteller)",
     "Hersteller PT (Lissabon-Region)",
     "5m 6.900 / 6m 8.400 / 7-8m Anfrage",
     "PT-LOKAL ALTERNATIVE zu Casa: junger Hersteller seit 2020 bei Lissabon. Kein Import-Zoll, kein Versandstress. Preise transparent auf OLX. Klima-Garantie & UV-Spezifikation noch zu erfragen! Anfrage dringend lohnt sich, bevor andere Optionen verfolgt werden. Kontakt via OLX.pt.",
     "https://www.olx.pt/d/anuncio/yurt-made-in-portugal-IDIyzPL.html"],

    [6, "YourTent.com CZ (Praha)",
     "Hersteller CZ - 5J-Plane-Garantie",
     "Schaetzung 12.000-16.000 EUR 6m inkl. Versand",
     "BESTE PLANE-GARANTIE NACH CASA: 5 JAHRE Plane-Garantie (60 Monate). 150 Jurten in 12 Jahren gebaut, MEDITERRANEAN TRACK RECORD bis Kanarische Inseln + Polarkreis. LAERCHENHOLZ-Rahmen (Algarve-optimal). 180cm Standard-Wandhoehe, 50mm Schafwoll-Filz, franz. Doppeltuer Laerche, Skylight 150/180cm, Ofen-Vorbereitung. Online-Kalkulator + Anfrage. Versand CZ->PT ca. 1.500-2.500 EUR.",
     "https://www.yourtent.com/"],

    [7, "Casa dos Sonhos Yurts (2027er Slot)",
     "Hersteller PT - 10-J. UV-Garantie SCHRIFTLICH",
     "6,1m 9.700 / 7,3m 15.750 EUR delivered+erected PT",
     "BLEIBT TOP-OPTION FUER LAENGEREN ZEITHORIZONT (2027+): Oktober 2026-Slot wurde 15.05.2026 an anderen Kunden vergeben. Naechster verfuegbarer Slot in 2027. Trotzdem das technisch beste Angebot fuer Algarve mit 10 J. UV-Garantie auf Plane, alles inkl. Lieferung+Aufbau in Festland-PT, portugiesische Materialien. Sarah & Vladimir (UK-Familien-Betrieb) etabliert seit Jahren in Alegrete. NIPC: Empresario em Nome Individual. Verifizieren via Brett bei Quinta Glamping + Werkstatt-Besuch.",
     "https://www.casadossonhos.co.uk/  |  https://www.facebook.com/yurtscasadossonhos/"],

    [8, "Bēt Yurts (August 2026 Slot)",
     "Hersteller PT (Algarve direkt)",
     "6m 28.060 / 7m 31.940 / 9m 44.870 EUR all-in",
     "ALGARVE-LOKAL ABER PREMIUM-PREIS: Konkretes Quote-PDF erhalten. Halb so teuer waere besser - DOPPELT SO TEUER wie Casa. 50% NICHT-RUECKZAHLBAR upfront, nur 2 JAHRE Plane-Garantie. Material laut Quote: 'Nordic Pine + Outer canvas + 80mm Insulation' OHNE Marken-/Herstellernennung (Marketing 'marine-grade' nicht im Quote bestaetigt). AUGUST 2026 Slot verfuegbar - Zeit-Vorteil. Nur empfehlenswert wenn andere Optionen ausfallen + August-Lieferung zwingend.",
     "https://www.betyurts.com/"],
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
    ("Preis & Kommerzielles (BESTAETIGT per Mail an Sebastian Mai 2026)", ""),
    ("Preisliste", "Direktangebot per Facebook-Messenger erhalten - oeffentlich nicht gelistet"),
    ("5m (20 m2)", "7.900 EUR delivered + erected mainland PT"),
    ("6,1m / 20-foot (30 m2)", "9.700 EUR delivered + erected mainland PT - BESTES EUR/m2-VERHAELTNIS"),
    ("7,3m / 24-foot (42 m2)", "15.750 EUR delivered + erected mainland PT"),
    ("9,1m / 30-foot (64 m2)", "18.000 EUR delivered + erected (inkl. 120x120cm Alu-Fenster fuer Proportionalitaet)"),
    ("EUR pro m2 Vergleich", "5m: 395 / 6,1m: 323 / 7,3m: 375 / 9,1m: 281 - 6,1m und 9,1m am effizientesten"),
    ("Lieferung & Aufbau", "IM PREIS ENTHALTEN (Festland-PT) - keine versteckten Versandkosten"),
    ("Import-Zoll/IVA", "ENTFAELLT (lokale PT-Produktion)"),
    ("Optionen Preise", "Noch nicht bestaetigt - aktiv nachfragen: Sturm-Kit, franz. Doppeltueren, Extra-Daemmung, Kaminzug, Mezzanine"),
    ("Slot-Verfuegbarkeit (Mai 2026)", "Aktuell Dezember 2026 oder 2027 angeboten -> hohe Auslastung"),
    ("Zahlungsbedingungen BESTAETIGT", "1) Anzahlung sichert den Termin (Teil der Gesamtsumme). 2) VOLLZAHLUNG 2 Monate vor Lieferung (5m/6,1m/7,3m) bzw. 3 Monate vor Lieferung (9,1m). 3) Lieferung selbst KEIN Zahlungstrigger - alles vorher beglichen."),
    ("RISIKO-Bewertung Zahlung", "ROTE FLAGGE: 100% im Voraus, Monate vor Lieferung. Branchenstandard ist 30-50% Anzahlung + Restzahlung bei Abnahme. Bei CdS keinerlei Hebel mehr nach Vollzahlung. Insolvenz-/Qualitaetsrisiko liegt vollstaendig beim Kaeufer."),
    ("Mitigation-Optionen", "1) Per Kreditkarte zahlen (Chargeback-Schutz) - mit Bank vorab klaeren ob Limit reicht. 2) Treuhand-Konto (Escrow) in PT, kostet ~0,5-1% Volumen. 3) Etappenzahlung verhandeln: 30% Anzahlung / 40% bei Baustart mit Foto-/Video-Nachweis / 30% bei Abnahme vor Ort. 4) Persoenlicher Besuch+Werkstatt-Besichtigung VOR Vollzahlung erzwingen. 5) Referenzen (PT/Algarve) kontaktieren und Punkt: Termintreue+Qualitaet abfragen."),
    ("Social-Media-Kanaele zur Verifikation",
     "Facebook https://www.facebook.com/yurtscasadossonhos/ - 3.617 Follower, regelmaessige Posts seit Jahren / YouTube https://www.youtube.com/channel/UCfenMyhVzEkX7P-lz2hjS6A - dokumentierte Aufbau-Videos / TikTok https://www.tiktok.com/@casadossonhosyurts"),

    ("", ""),
    ("Nachhaltigkeit / ESG", ""),
    ("Energie Produktion", "100% Solar in der Werkstatt"),
    ("Materialien-Herkunft", "Portugiesisch (Holz, Kork)"),
    ("Soziales Engagement", "Plant Wollverarbeitung um lokale Schafwolle aufzuwerten (Alentejo) - aktuell unterbewertet"),
    ("Kreislaufwirtschaft", "Natuerliche Materialien, kein PVC im Standard"),

    ("", ""),
    ("Risiken & Hinweise", ""),
    ("Lieferzeit", "Handarbeit -> aktuell Dezember 2026 oder 2027. Slot frueh sichern."),
    ("Skalierung", "Kleines Familienunternehmen - Reaktionszeit kann variieren"),
    ("HAUPTRISIKO Zahlung", "100% Vorauszahlung 2-3 Monate vor Lieferung - vollstaendiges Kaeufer-Risiko ohne Sicherheit. Mitigation siehe Sektion 'Preis & Kommerzielles'."),
    ("Sturmkomponente", "Algarve-Kueste: Sturm-Kit OBLIGATORISCH bestellen (Preis offen)"),
    ("Hitze Innenraum", "Trotz 10-J. Plane: zusaetzliches Schattennetz/Pergola fuer 40C+ Sommertage empfehlenswert"),
    ("Verifikations-Schritte VOR Auftrag", "1) Show-Yurt-Besuch Portalegre + Werkstatt-Tour. 2) 2 Algarve/Monchique-Referenzen telefonisch sprechen. 3) Rechtsform der Firma (Lda/Unipessoal/Einzelperson) checken. 4) Zahlungsmodalitaeten verhandeln (Etappen oder Kreditkarte)."),

    ("", ""),
    ("Fazit fuer Algarve", ""),
    ("Bewertung", "BLEIBT TOP-OPTION TECHNISCH - aber STATUS 15.05.2026: Oktober-2026-Slot wurde an anderen Kunden vergeben. Naechster Slot 2027."),
    ("Sweet Spot", "6,1m / 30 m2 zu 9.700 EUR delivered+erected. Konkretes Quote 15.05.2026: 9.700 Basis + 850 Extra-Tuer + 100 Flue-Kit = 10.650 EUR. + Plattform DIY 4.000 + Schattennetz 240 = ~14.890 EUR all-in."),
    ("Konkretes Angebot in der Konversation", "6,1m mit 2.Tuer und Flue-Kit = 10.650 EUR; 1.000 EUR Anzahlung sichert Slot, 9.650 EUR Restbetrag 1. August 2026 zu zahlen (= 2 Monate vor Lieferung). Lieferung war 2. Oktoberwoche 2026 vorgesehen - SLOT INZWISCHEN WEG."),
    ("Inhaber-Hintergrund", "Sarah (englisch) + Vladimir (Mann) + Tochter 3J, leben in Alegrete. Sarah operiert die Mail-Kommunikation. Auch ueber Casa dos Sonhos Cottage (Ferienvermietung) seit Jahren in PT etabliert - dokumentiert auf babyfriendlyboltholes, homeexchange, peoplelikeus."),
    ("NIPC-Status", "260282812 - Anfangsziffer '26' = persoenlicher NIF, der als NIPC verwendet wird = 'Empresario em Nome Individual' (Einzelunternehmer ohne Haftungsbeschraenkung). Sarah oder Vladimir persoenlich als Geschaeftsbetreiber. Standard fuer PT-Kleinhandwerker, kein Showstopper, aber kein Lda-Schutz fuer Kaeufer."),
    ("Verifikations-Optionen", "Brett bei Quinta Glamping (https://www.quintaglamping.com/) hat eine Casa-Jurte als Lake View Yurt - Uebernachtung dort moeglich zur Casa-Qualitaets-Pruefung. Paula Young (FB 'Paula Vegan Chef') verkauft eine gebrauchte 7,3m Casa-Jurte in der Algarve - kann besichtigt werden + Algarve-Praxiserfahrung in 1-2 Std Gespraech erhalten."),
    ("Naechster Schritt", "1) Casa: Antwort auf 'Slot-weg'-Mail abwarten - 2027er-Optionen + ggf. Vermittlung Paula Young. 2) Quinta Glamping Lake View Yurt fuer Naechtigung buchen (Casa-Qualitaets-Test). 3) Paula Young direkt anschreiben - moeglicherweise sofortige Algarve-Loesung. 4) Parallel DIY-Modify-Szenario A verfolgen (siehe Sheet DIY Bauplan). 5) Wenn alle obigen ausfallen: 2027-Slot bei Casa annehmen oder Bēt August."),
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


# ============ Sheet 7: OLX 6m+ Funde ============
ws7 = wb.create_sheet("OLX 6m+ Funde")
headers7 = [
    "Titel/Listing", "Standort PT", "Durchmesser",
    "Innenflaeche m2", "Preis EUR", "Zustand", "Material/Bauart",
    "Versand/Aufbau", "Direktlink", "Anmerkungen / Bewertung Algarve",
]
rows7 = [
    ["Yurts mongois / Mongolian Yurts",
     "Igreja Nova e Cheleiros (Mafra, naehe Lissabon)",
     "6 m", "28 m2", 5000, "Gebraucht/Restposten - direkt vom Verkaeufer",
     "Klassische mongolische Bauweise: Filzdecke (warm im Winter, kuehl im Sommer), Holzgitter, mobil auf jeder ebenen Flaeche aufbaubar",
     "PT-inland, Transport+Aufbau gegen Aufpreis. Optional Pelletofen gegen Aufpreis.",
     "https://www.olx.pt/d/anuncio/yurts-mongis-mongolian-yurts-IDHQ8sJ.html",
     "TOP-FUND: Gueneststes 6m-Angebot auf OLX. ABER: traditioneller Filz/Canvas ohne UV-Schutz - mit Schattennetz nachruesten (siehe Hinweise-Sheet) sonst Plane nur 2-3 Jahre Lebensdauer in Algarve. Mit Pelletofen+Aufbau realistisch 6.000-7.500 EUR Total"],

    ["TENDA YURT Original do Quirguistao (Marke Yourta - Modell 'Serenity')",
     "Palmela (Setubal, naehe Lissabon)",
     "ca. 6 m (geschaetzt aus 28 m2)", "28 m2 (Kapazitaet 30 Personen)", 13000,
     "Neu / handgefertigt, original aus Quirguistao (Issyk-Kul See)",
     "Original kirgisisch: Schilf-Geflecht (Junco), Wolle und Filz. Komplette Originaldeko: Shyrdak-Teppich, Baskur-Band, Korpe-Linen-Kissen, geschnitzte Holztuer, Filz-Tuerverkleidung, dekorative Kuppelbaender",
     "Versand EXTRA (Region Lissabon - genauen Preis erfragen)",
     "https://www.olx.pt/d/anuncio/tenda-yurt-original-do-quirguisto-yurt-tent-original-from-kyrgystan-IDHEwz9.html",
     "KATEGORIE ROT (NICHT EMPFOHLEN ALGARVE): HOHE AUTHENTIZITAET aber NICHT fuer Atlantikkueste. Traditioneller Filz/Wolle nicht fuer 40C+ ausgelegt; Plane unter UV-Belastung anfaellig; Salzluft schimmelt Wolle. Sehr dekorativ - eher fuer kuehlere Bergstandorte. 13.000 EUR fuer eine Deko-Jurte ist viel."],

    ["Yurt de 8 metros com janela panoramica (OLX neu Mai 2026) - Modell 'Yury' von Benedikt",
     "Gouveia (Sao Pedro / Sao Juliao) - Serra da Estrela, Zentral-PT",
     "8 m", "50 m2", 15000,
     "NEU (Neuware) - veroeffentlicht 06.05.2026, Verkaeufer Benedikt seit 07.2025 auf OLX",
     "STANDARD INKL.: Tuer + 2 transparente Fenster + WEISSER INNENSTOFF + FILZISOLIERUNG + GRUENDACH (?). Modell 'Yury' (vermutlich Green World UA oder vergleichbar). Hersteller-Spezifikation muss erfragt werden.",
     "Innerhalb PT (Gouveia ~400 km von Aljezur). Transport/Aufbau muss erfragt werden ob enthalten.",
     "https://www.olx.pt/d/anuncio/yurt-de-8-metros-de-dimetro-com-janela-panormica-IDJatvu.html",
     "SEHR STARKES ANGEBOT FUER 8M! 15.000 EUR ist GUENSTIGSTER 8m-PREIS AM MARKT (vs Yourtepoque 19k / Atilla 22k / DIY 24k / Casa 7,3m 20,7k). Verkaeufer Benedikt = deutscher Reseller in PT, deutsche Inserat-Sprache = einfache Kommunikation. ACHTUNG: Privatverkauf ohne Verbraucherschutz - vor Kauf besichtigen + Spec klaeren (Plane-Material/UV-Garantie, was ist 'Gruendach', Hersteller, Garantie, Transport+Aufbau)."],

    ["Yurt luxuosa NOVA 8 metros (OLX neu Mai 2026)",
     "Standort tbd (Sebastian: Details bitte teilen)",
     "8 m", "ca. 50 m2", "Preis tbd",
     "NEU - Luxus-Variante",
     "Konstruktion tbd - Sebastian Details aus Listing bitte teilen",
     "tbd",
     "https://www.olx.pt/d/anuncio/yurt-luxuosa-nova-8-metros-de-dimetro-IDJbO7n.html",
     "NEUFUND Mai 2026: 'Luxus' + 'NEU' 8m. Wahrscheinlich vom selben Verkaeufer wie IDJatvu (zwei verschiedene Stueck auf Lager). MOEGLICHER ANBIETER: Green World UA. Details fehlen - Sebastian muss Preis/Verkaeufer/Standort teilen."],

    ["Mikael (Mickaël) = Bento - Pacific Yurts 8m Panoramic (Gouveia)",
     "Portugal (Gouveia, Guarda) - 10 Mon. in Van gelagert nach Grundstueck-Brand",
     "8 m, 2.10 m Wandhoehe", "ca. 50 m2",
     "15.000 EUR (NORMALPREIS 18.500 = 3.500 EUR Rabatt fuer Rost) + 700 AC4 + 1.000 Vinyl-Overlay optional + 500-800 Lieferung. PLATTFORM NICHT INKL!",
     "NEU - August/Oktober 2025 gekauft Pacific Yurts. BIG FIRE zerstoerte Grundstueck. 10 Mon in Van: lagerung TOP (kein Schimmel/Geruch/Verzug). Aber: 30% Schrauben sichtbar rostig + ALLE Schrauben sind 'regular steel' (Fertigungsfehler statt verzinkt) = ALLE bedrohlich in Algarve-Salzluft.",
     "HERSTELLER: PACIFIC YURTS USA - 8m Panoramic Model. KONSTRUKTION: TRADITIONELLE KHAANA-SCHEREN-GITTER. Holz: Massiv Nordische Kiefer, klar lackiert + INITIAL Insektenschutz-Spray bereits durchgefuehrt. Mikael macht zweite Spray-Anwendung am Abholtag (inkl.). ISOLIERUNG: 7mm Schafwoll-Filz Dichte 1.1 kg/m² (Standard Pacific Yurts, fuer Algarve duenn - Mikael kann dickere Daemmung beschaffen). Canvas-Dach Standard + Vinyl-Overlay-Option (NICHT in 15k, +1.000 EUR fuer 8m Variante, hat 2 zusaetzliche Oeffnungen). SKYLIGHT: 1.40m Kiefer+Eisen. KONFIG: 1 Doppeltuer + 2 Glas-Fenster in Khaana. AC4-Boden OPTIONAL 700 EUR (feuchtigkeitsresistent). PLATTFORM Sebastian baut selbst.",
     "Selbst-Abholung Mietvan 750 EUR (Real-Referenz). Mikael-Lieferung Castelo-Branco-Mitnahme 1. Juni (vor Sebastians Besuch = nicht nutzbar). Mikael hilft beim Beladen mit Maschine, Helfer auch beim Abladen.",
     "Telegram 'Mickaël' = Bento (OLX-Verkaeufer). Tragischer Brand-Opfer-Hintergrund.",
     "AKTUALISIERT 24.05.2026 nach voller Antwort: NORMALPREIS waere 18.500, sebastian zahlt 15k = 3.500 EUR fuer Rost. ALLE Schrauben muessen ersetzt werden (Mikael ~800 EUR oder Sebastian ~300 EUR + 50 Std). MIT VINYL-OVERLAY: ALL-IN ~22.320 EUR. OHNE VINYL: ~21.320 EUR. TIMING-KONFLIKT: 1. Juni Lieferung vs 14./21. Juni Besuch = unmoeglich kombiniert. Sebastian schlaegt Selbst-Abholung Ende Juni oder spaetere Liefertour vor."],

    ["Yurt Made in Portugal (Hersteller-Insertion auf OLX)",
     "Loures (naehe Lissabon)",
     "6 m verfuegbar (auch 5m, 7m, 8m)",
     "ca. 28 m2 (6m)", 8400, "NEU - Hersteller seit 2020",
     "Modern interpretierte traditionelle Bauweise; portugiesischer Hersteller",
     "Lieferung+Aufbau NICHT inklusive - extra je nach Standort und Groesse",
     "https://www.olx.pt/d/anuncio/yurt-made-in-portugal-IDIyzPL.html",
     "Lokaler PT-Hersteller (kein Importzoll), seit 2020 aktiv. Direktkontakt ueber OLX. 5m=6.900 EUR / 6m=8.400 EUR / 7-8m auf Anfrage. Hinweis: Klimaschutz/UV-Garantie nicht spezifiziert -> dringend nachfragen, ob Plane fuer Algarve geeignet ist."],
]
# Nach Preis sortieren
rows7.sort(key=lambda r: r[4] if isinstance(r[4], (int, float)) else 9999999)
style_sheet(ws7, headers7, rows7)
widths7 = [34, 30, 14, 18, 14, 28, 40, 32, 50, 60]
for i, w in enumerate(widths7, 1):
    ws7.column_dimensions[get_column_letter(i)].width = w
# Zeilenhoehe fuer lange Texte
for r in range(2, len(rows7) + 2):
    ws7.row_dimensions[r].height = 110

# Hinweise-Footer
note_row = len(rows7) + 3
note = ws7.cell(row=note_row, column=1,
                value=("Ausgeschlossen (unter 6 m Mindestgroesse): "
                       "'Yurt mongol autentico Mafra' (5,5 m, 8.000 EUR) - "
                       "https://www.olx.pt/d/anuncio/yurt-mongol-autntico-IDIR7sZ.html"))
note.alignment = Alignment(wrap_text=True, vertical="top")
note.font = Font(italic=True, color="666666")
ws7.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=10)
ws7.row_dimensions[note_row].height = 32

note2 = ws7.cell(row=note_row + 2, column=1,
                 value=("WICHTIG: OLX-Angebote aendern sich taeglich. Vor Anfrage Live-Suche unter "
                        "https://www.olx.pt/moveis-casa-e-jardim/jardim-e-bricolage/q-yurts/ "
                        "pruefen. Recherche-Stand: 2026-05-21."))
note2.alignment = Alignment(wrap_text=True, vertical="top")
note2.font = Font(italic=True, color="666666")
ws7.merge_cells(start_row=note_row + 2, start_column=1, end_row=note_row + 2, end_column=10)
ws7.row_dimensions[note_row + 2].height = 32

# PT-Used-Markt-Tiefenscan Mai 2026
note3 = ws7.cell(row=note_row + 4, column=1,
                 value=("PT-USED-MARKT TIEFENSCAN MAI 2026: Markt ist klein - max 3-5 aktive >=6m Listings gleichzeitig. "
                        "Konzentration Mafra/Lissabon, kaum direkt in Algarve. "
                        "EMPFOHLENE PLATTFORMEN ZUM BEOBACHTEN: "
                        "1) OLX.pt https://www.olx.pt/moveis-casa-e-jardim/q-yurts/ (Hauptkanal); "
                        "2) Facebook Marketplace Portugal + 'For Sale or Swap in the Algarve' Gruppe https://www.facebook.com/groups/FORSALEORSWAPINTHEALGARVE/ (eigenen Wanted-Post aufgeben!); "
                        "3) CustoJusto.pt https://www.custojusto.pt/ (selten); "
                        "4) Casa dos Sonhos Facebook https://www.facebook.com/yurtscasadossonhos/ (sporadische Secondhand-Posts); "
                        "5) Hipcamp + Workaway Algarve-Hosts direkt fragen (Glamping-Sites verkaufen manchmal). "
                        "WICHTIGSTES OFF-MARKT-LEAD: Paula Young via Sarah-Vermittlung (gebrauchte 7,3m Casa-Jurte). "
                        "EXPERTEN-FAZIT: Casa-Jurten kommen kaum auf Used-Markt (Hersteller-Wartelisten + Eigentuemer halten Jahre). "
                        "Off-OLX-Leads (Sarah-Vermittlung, FB-Gruppe Wanted-Post) sind wertvoller als oeffentliche Marktplaetze."))
note3.alignment = Alignment(wrap_text=True, vertical="top")
note3.font = Font(italic=True, color="2E5C8A")
ws7.merge_cells(start_row=note_row + 4, start_column=1, end_row=note_row + 4, end_column=10)
ws7.row_dimensions[note_row + 4].height = 120


# ============================================================
# ============ DIY EIGENBAU 8m JURTE FUER ALGARVE ============
# ============================================================
# Annahmen Sebastian (Mai 2026):
# - 8 m Durchmesser (~50 m2 Wohnflaeche)
# - Tono (Kronenring) fertig vom Lieferanten
# - Aussenplane Sauleda Solar Pro / Tempotest als Massanfertigung
# - Skill mittel (DIY, Holz schon mal bearbeitet)
# - Standort: Westliche Algarve 8650 (Aljezur), atlantisches Mikroklima


# ---------- Sheet 8: DIY Bauplan Uebersicht ----------
ws8 = wb.create_sheet("DIY 8 Bauplan Uebersicht")
ws8.column_dimensions["A"].width = 28
ws8.column_dimensions["B"].width = 22
ws8.column_dimensions["C"].width = 72

t = ws8.cell(row=1, column=1, value="DIY 8m Jurte - Konstruktionsuebersicht (Algarve-optimiert)")
t.font = TITLE_FONT
ws8.merge_cells("A1:C1")
ws8.row_dimensions[1].height = 28

bauplan = [
    ("Gesamtgeometrie", "", ""),
    ("Durchmesser", "8,00 m", "Aussenmass Khaana ausgebreitet, plus 100 mm Plane-Ueberlappung pro Seite"),
    ("Wandhoehe (Khaana)", "2,00 m", "Casa-dos-Sonhos-Standard (vs. mongolisch traditionell 1,60-1,70 m) - mehr Stehhoehe + bessere Belueftung"),
    ("Hoehe Mittelring (Tono)", "4,20 m", "Ueber Plattform-Oberkante; Dachneigung am Eaves ca. 28-30 Grad"),
    ("Tono-Durchmesser", "170-180 cm", "KORRIGIERT Mai 2026 nach technischem Feedback von Atilla + Adorjan: 120 cm fuer 8m ist statisch UNDERDIMENSIONIERT. Adorjan praktiziert 170 cm bei 8m (mit 48 Rafters 5x10cm), Atilla 180 cm (mit groesserer Anzahl Uni). Bei 60 Uni und 120 cm Tono waeren nur 6 cm Lochabstand am Ring."),
    ("Anzahl Dachstangen (Uni)", "60-80 (je nach Tono-Wahl)", "Bei 180 cm Tono mit jeder Khaana-Spitze=Uni: 64-80 Uni. Bei 170 cm Tono mit massiven 5x10cm Rafters (Adorjan-Bauart): 48 Uni reichen. Beides bewaehrte Engineering-Loesungen."),
    ("Anzahl Khaana-Sektionen", "8", "Jede Sektion ca. 3,14 m breit ausgebreitet (= pi * 8m / 8)"),
    ("Wohnflaeche innen", "~50 m2", "Inkl. 1,2 m Tono-Saeule-frei -> volle Fussbodenflaeche nutzbar"),

    ("", "", ""),
    ("Konstruktionsprinzip", "", ""),
    ("Bauweise", "Traditionelle Khaana mit Tono+Uni", "Selbsttragend ohne Mittelsaeule - Wandgitter haelt unter Druck der Dachstangen-Schubspannung. Zwingend: stabiles Spannband (Kuriye) auf Wandkronenhoehe."),
    ("Statik-Prinzip", "Druckring + Zugring", "Tono = Druckring (Dachstangen druecken nach innen). Spannband um Khaana-Top = Zugring (haelt Wand gegen Schub). Funktioniert seit 3000 Jahren."),
    ("Sturmsicherung Algarve", "Sturm-Kit dringend", "Atlantik-Winterstuerme 80-100 km/h. Erdanker an 8 Punkten unter Plattform, 5-Tonnen Ratschen-Riemen ueber Aussenplane, Wand+Dach-Verstaerkung mit zusaetzlichen Querstaeben."),

    ("", "", ""),
    ("Aufbau Schichten (innen nach aussen)", "", ""),
    ("Schicht 1 - Khaana-Rahmen", "Edelkastanie 30x10 mm", "Lokal in PT verfuegbar, naturlich tannin-/insektenresistent durch Gerbsaeure - Termiten-Schutz."),
    ("Schicht 2 - Innenliner", "Baumwoll-Stoff 200 g/m2", "Aesthetisch + faengt Daemmwolle ein. Maemost in PT bei Sattlerei beziehbar."),
    ("Schicht 3 - Daemmung", "Schafwollfilz 16 mm (ISOLENA)", "100% Wolle Ionic Protect, schimmelresistent, oeko. 50 m2 fuer Wand + 50 m2 fuer Dach = 100 m2 Gesamtbedarf."),
    ("Schicht 4 - Atmungsoffene Membran", "Solitex Plus / Tyvek 'Soft'", "OPTIONAL aber dringend empfohlen: zwischen Daemmung und Aussenplane, verhindert Kondensat in der Daemmwolle. Sebastian hat das auch bei Atilla geplant."),
    ("Schicht 5 - Aussenplane", "Sauleda Solar Pro Acryl", "Marine-grade UV, salzluftbestaendig, atmungsaktiv, 290-340 g/m2. ca. 110 m2 Bedarf (Dach + Wand + Ueberlappung)."),

    ("", "", ""),
    ("Holzwahl - Algarve-optimiert", "", ""),
    ("Khaana (Lattengitter)", "Edelkastanie (Castanho)", "Lokal in PT (Norden+Centro). Tannin-haltig = natuerlicher Termiten-/Insektenschutz. Leicht zu biegen, gut sägbar. Klasse 2 Dauerhaftigkeit nach EN 350."),
    ("Uni (Dachstangen)", "Sibirische Laerche", "Hohe Festigkeit bei geringem Gewicht, harzhaltig = rotresistent. Aus AT/RU-Import via DE oder direkt von Siero Lam (ES/PT). Klasse 3-4 Dauerhaftigkeit."),
    ("Tono (Kronenring)", "Gekauft - Hartholz", "Lieferant macht Wahl (typisch Eiche oder Esche laminiert). Achtung: Kontaktpartner fuer ALLE Bohrungen, weil die Geometrie 1.200 mm + 60 Bohrungen nicht einfach zu replizieren ist."),
    ("Tuerrahmen + Tuer", "Robinie / Falsche Akazie", "Klasse 1-2 Dauerhaftigkeit (auch ohne Behandlung), extrem hart, salzwasserresistent. Aus AT/HU-Import. Alternative: portugiesische Edelkastanie."),
    ("Plattform-Konstruktion", "Druckimpr. Kiefer + Laerchen-Deck", "Tragbalken Kiefer KDI (Klasse 4 Dauerhaftigkeit fuer Bodennaehe), Decke Laerche fuer Optik+Klima. Alternativ Robinie als Premium."),
    ("Plattform-Schraubpfaehle", "Stahl verzinkt Krinner / Stop", "Erdschrauben 800-1.000 mm tief, 8 Stueck fuer 8m-Plattform. Loest die Frage 'Beton oder nicht' (du wolltest keinen Beton)."),

    ("", "", ""),
    ("Plane / Aussenhuelle", "", ""),
    ("Material", "Sauleda Solar Pro 300 g/m2 (Acryl)", "Spanischer Hersteller seit 1897, Marine-grade. Solrain-Variante zusaetzlich impraegniert."),
    ("Alternative Tempotest Parà (IT)", "Tempotest Marine 290 g/m2", "Italienisches Premium, 6-J. Farbtreue-Garantie, Teflon EXTREME Finish. Vergleichbar Sauleda."),
    ("Auftragsfertigung", "Lokaler Persenningmacher / Sattlerei in PT", "Du lieferst die Stoffrolle, der Fachbetrieb naeht zu. Suchbegriffe PT: 'oficina de toldos', 'velejaria' (Sailmaker), 'persenning'."),
    ("Aussenflaeche Gesamt", "~110 m2 (mit Ueberlappung)", "Dach: 50 m2 Kegelmantel + 10% Ueberlappung. Wand: 8 x pi x 2,2 m = 55 m2. Plus Verschnitt 15% Reserve."),
    ("Plane-Lebensdauer", "10-15 Jahre", "Sauleda/Tempotest mit Schattennetz darueber. OHNE Schattennetz in Algarve: 5-8 Jahre realistic."),

    ("", "", ""),
    ("Plattform / Foundation", "", ""),
    ("Typ", "Geschraubte Pfahl-Plattform", "Krinner-Schraubpfaehle KSF-M 800mm. Kein Beton. Reversibel. Belueftung darunter verhindert Feuchte+Termiten."),
    ("Tragwerk", "Doppel-T aus 50x200 mm Kiefer KDI", "Konzentrische Rahmen-Konstruktion + radiale Balken alle 45 Grad."),
    ("Bodenaufbau (innen nach aussen)", "Estrich-Vlies + Kork 30mm + Laerchen-Deck 27mm", "Trittschall, Kaelte-Pufferung, schoene Optik. Kork = portugiesisches Material par excellence."),
    ("Diameter", "8,20 m (10 cm Ueberstand)", "Verhindert Regenwasser-Hochzug ins Innere; passt Plane-Wand-Saum exakt."),
]

for r_idx, (a, b, c) in enumerate(bauplan, 2):
    if a and not b and not c:
        # Sektions-Header
        cell = ws8.cell(row=r_idx, column=1, value=a)
        cell.font = Font(bold=True, size=12, color="2E5C8A")
        ws8.merge_cells(start_row=r_idx, start_column=1, end_row=r_idx, end_column=3)
        ws8.row_dimensions[r_idx].height = 22
    elif not a and not b and not c:
        ws8.row_dimensions[r_idx].height = 8
    else:
        ws8.cell(row=r_idx, column=1, value=a).font = Font(bold=True)
        ws8.cell(row=r_idx, column=2, value=b)
        ws8.cell(row=r_idx, column=3, value=c).alignment = Alignment(wrap_text=True, vertical="top")
        for col in (1, 2, 3):
            ws8.cell(row=r_idx, column=col).border = BORDER
            ws8.cell(row=r_idx, column=col).alignment = Alignment(wrap_text=True, vertical="top")
        ws8.row_dimensions[r_idx].height = max(20, min(60, 15 + len(c) // 6))

ws8.freeze_panes = "A2"


# ---------- Sheet 9: DIY Materialliste ----------
ws9 = wb.create_sheet("DIY 9 Materialliste")
headers9 = ["Position", "Bauteil", "Material/Spezifikation", "Menge", "Einheit",
            "Einzelpreis EUR", "Gesamtpreis EUR", "Bezugsquelle Vorschlag", "Bemerkung"]

# Berechnungs-Basis fuer 8m Yurte
material = [
    # KHAANA (Lattengitter Wand)
    ["1.0", "Khaana Latten Edelkastanie", "30x10 mm, getrocknet, gehobelt, 2,1 m Laenge",
     280, "Stueck", 4.50, 1260,
     "Mil Martins & Irmao (PT-Norden) / Industria de Madeiras / Albano Leite",
     "8 Sektionen x ~35 Latten/Sektion. 10% Verschnitt eingerechnet. Wahlweise auch in 30x12 mm fuer mehr Stabilitaet."],
    ["1.1", "Khaana Nieten/Bolzen", "M5 x 30 mm Senkkopf Edelstahl A4 + Hutmuttern",
     350, "Stueck", 0.20, 70,
     "Wuerth PT / Bauhaus PT / Schraubenking.de",
     "Pro Khaana-Kreuzung 1 Bolzen. ~38 Kreuzungen pro Sektion * 8 = 304 + Reserve."],
    ["1.2", "Khaana Hanfband / Lederband", "Hanfgurt 25 mm breit ODER Leder-Streifen",
     50, "Meter", 2.50, 125,
     "tipiwakan.com / handweber-PT / Saddlery Lousa",
     "Optional fuer traditionelle Optik statt Bolzen. Sebastian: Bolzen empfohlen (haltbarer in Salzluft)."],

    # TONO (Kronenring) - GEKAUFT
    ["2.0", "Tono - Kronenring fertig (KORRIGIERT 2026)", "170-180 cm Durchmesser, 60-80 Bohrungen Uni, 4-lagig verleimt (Adorjan-Standard) - NICHT 120 cm wie urspruenglich angenommen",
     1, "Stueck", 1500, 1500,
     "Camping Yurts (UK) / FAMTENTS (CZ) / Adorjan Jurta (HU, 170 cm) / Atilla (HU, 180 cm)",
     "WICHTIG: Atilla + Adorjan haben beide bestaetigt: 120 cm Tono fuer 8m statisch zu klein. 170-180 cm noetig. Preisaufschlag entsprechend (700-1.500 EUR). Adorjan baut komplettes Frame inkl. Tono fuer 8.150 EUR."],
    ["2.1", "Tono Versand nach PT", "Spedition Übermass",
     1, "Pauschal", 250, 250,
     "Logistik-Anbieter Wahl Lieferant",
     "Bei Anfrage einkalkulieren."],

    # UNI (Dachstangen)
    ["3.0", "Uni Dachstangen (KORRIGIERT 2026)", "Bei 170 cm Tono (Adorjan-Bauart): 48 Stk mit 5x10 cm massiv / Bei 180 cm Tono (Atilla-Bauart): 64-80 Stk mit 3,5x3,5 cm, jede Khaana-Spitze",
     80, "Stueck", 18, 1440,
     "Holz Schiller AT / Siero Lam (PT) / Holz Possling DE / Adorjan oder Atilla im Komplett-Frame",
     "WICHTIG: Wenn Adorjan/Atilla das ganze Frame liefert, sind Uni inklusive. Hier nur relevant fuer reines DIY. 5x10 cm massive Variante ist teurer aber statisch besser fuer 8m."],
    ["3.1", "Uni-Befestigung am Tono", "Edelstahl-Niet M6 x 40 mm A4",
     65, "Stueck", 0.80, 52,
     "Wuerth PT / Bauhaus",
     "Eine Bohrung im Uni-Kopf + Tono-Bohrung verbinden."],
    ["3.2", "Uni-Auflage auf Khaana", "Lederband / Hanfschlaufe alle 60 cm an Wandkrone",
     50, "Meter", 3, 150,
     "Sattlerei PT / Lederhandwerk",
     "Pro Sektion 8 Uni werden aufgelegt. Bestens: durchgehende Schlaufen-Schnur (Kuriye-Funktion)."],

    # KURIYE (Spannband)
    ["4.0", "Kuriye - Spannband Wandkrone", "Polypropylen-Gurt 50 mm, 8 Tonnen Zugfest",
     30, "Meter", 4, 120,
     "Industriebedarf PT / Wuerth / Schraubenking.de",
     "2 komplette Umrundungen (8m Durchmesser * pi = 25,1 m je Lage). Mit Klemmschnalle."],
    ["4.1", "Spannschloesser Edelstahl", "M10 Spannschloss A4",
     4, "Stueck", 18, 72,
     "Wuerth / Bootsbedarf PT (Acastillaje)",
     "An 4 Stellen Spannung einstellbar."],

    # TUER
    ["5.0", "Tuerrahmen Robinie", "Pfosten 100x100 mm, Sturz 100x150 mm",
     12, "Meter", 25, 300,
     "Holz Henkel AT / Forstamt PT / Holz Schiller",
     "Trockene Robinie, gehobelt. Alternative PT: Edelkastanie Bohlenware."],
    ["5.1", "Tuerblatt Doppeltuer", "2x je 90x200 cm Robinie/Eiche mit Doppelverglasung",
     1, "Set", 800, 800,
     "Lokale Schreinerei PT (Tischlerei) / Bauhaus IKEA-Tueren als Basis",
     "Selber bauen oder bei lokalem Tischler in PT in Auftrag geben. UPVC waere guenstiger aber waermer."],
    ["5.2", "Tuerbeschlaege Edelstahl A4", "Bandscharniere x 6 + Schloss + Drueckergarnitur",
     1, "Set", 250, 250,
     "Wuerth PT / Brico Marche",
     "A4 statt A2 wegen Salzluft."],

    # PLATTFORM
    ["6.0", "Schraubpfaehle Krinner KSF-M", "M-Profil, 800 mm Laenge, verzinkt",
     8, "Stueck", 35, 280,
     "Krinner GmbH DE / Bauhaus PT / Acreditados.pt",
     "Statt Beton. 8 Punkte alle 45 Grad auf 8m-Kreis."],
    ["6.1", "Tragbalken Kiefer KDI Klasse 4", "50 x 200 mm, 4,0 m Laenge",
     20, "Stueck", 28, 560,
     "Leroy Merlin PT / AKI / lokales Saege-werk",
     "Doppel-T-Rahmen + Radial-Balken. KDI = Kesseldruckimpragniert."],
    ["6.2", "Laerchen-Deck T+G", "27 x 137 mm Nut+Feder, 4 m Laenge",
     50, "m2", 38, 1900,
     "Mil Martins / Maderterraneo (ES) / Holz Possling",
     "ca. 50 m2 zu verlegen plus 10% Verschnitt."],
    ["6.3", "Kork-Daemmung Plattform", "30 mm gepresst Naturkork",
     50, "m2", 18, 900,
     "Amorim Cork (PT national) / Sofalca / Maderterraneo",
     "Portugiesisches Material par excellence. Zwischen Tragbalken und Decke."],
    ["6.4", "Estrich-Vlies + Dampfsperre", "PE Folie 200 my + Geo-Vlies 300 g/m2",
     50, "m2", 4, 200,
     "Bauhaus PT / AKI",
     "Unter Plattform gegen aufsteigende Feuchte."],
    ["6.5", "Plattform-Schrauben Edelstahl", "Spax A4 5x80 mm + 4x50 mm",
     1, "Pauschal", 200, 200,
     "Wuerth PT / Schraubenking.de",
     "A4 wegen Salzluft. Pauschal Sortiment."],

    # DAEMMUNG WAND+DACH
    ["7.0", "Schafwollfilz Isolena 16 mm", "100% Schafwolle Ionic Protect, 1,2 m breit",
     105, "m2", 16, 1680,
     "ISOLENA AT direkt / baustoffplus.de / DAEMWOOL AT",
     "Wand 55 m2 + Dach 50 m2 = 105 m2. Bei Bestellung in DE: USt-IdNr nutzen fuer IC-Lieferung."],
    ["7.1", "Innenliner Baumwoll-Sergeant", "200 g/m2, naturweiss/decorativ, 1,5 m breit",
     65, "m2", 12, 780,
     "Toldum.es / Sattlerei PT / Stoffhandel-Industrie",
     "Wand 55 m2 + Dach 50 m2 - Verschnitt = 65 m2 Stoffbedarf bei Verzug. Wird unterhalb Filz montiert."],

    # ATMUNGSAKTIVE MEMBRAN
    ["8.0", "Membran Solitex Plus / Tyvek Soft", "Diffusionsoffen, wasserdicht, 1,5 m breit",
     115, "m2", 6, 690,
     "Pro Clima AT / Tyvek via Bauhaus PT",
     "Zwischen Filz und Aussenplane. WICHTIG fuer Algarve: verhindert Kondensat in Daemmwolle."],

    # AUSSENPLANE (gekauft als Service)
    ["9.0", "Aussenplane MASSANFERTIG", "Sauleda Solar Pro 290-340 g/m2 / oder Tempotest Marine, Beige/Sand",
     1, "Set", 2200, 2200,
     "PRIORITAET: Dune Algarve Sailmakers (Vilamoura) / Toldos Etapaveloz (Loule) / Textilux (Algarve) / Toldos Chique / Toldegarve / SOMBRIARTE / Sombra&Constroi (Portimao). BACKUP: Yurt Workshop ES (Cadiar) / Atilla HU.",
     "GAME-CHANGER ERKENNTNIS: Lokale Algarve-Sattler arbeiten taeglich mit Sauleda Marine-Acryl, ideal fuer Jurten-Cover-Auftrag. Schaetzung 6m: 1.800-2.500 EUR (Material+Verarbeitung). Komplett-Set Dach+Wand+Kuppel-Cover, inkl. Saeume, Tunnel fuer Spannband, Verstaerkungen Eaves. PERSOENLICHER BESUCH MOEGLICH = Maßabstimmung perfekt + lokale Reparatur-Quelle."],
    ["9.1", "Kuppel-Oberlicht (Polycarbonat)", "Klar, oeffenbar, 1,2 m Durchmesser",
     1, "Stueck", 750, 750,
     "Bauhaus PT / Atilla Jurte (HU) / Custom",
     "Wahlweise von Tono-Lieferant gleich mitbestellen."],

    # STURM-KIT
    ["10.0", "Erdanker Sturm-Kit", "Schraubanker M16 1.000 mm + Stahlseil 6 mm",
     8, "Stueck", 35, 280,
     "Krinner / Wuerth PT",
     "Algarve-Atlantik OBLIGATORISCH. 8 Anker um Plattform, Seil ueber Dach gespannt."],
    ["10.1", "Ratschen-Riemen 5 Tonnen", "Polypropylen 50 mm / 8 m Laenge",
     8, "Stueck", 22, 176,
     "Wuerth PT",
     "Pro Erdanker 1 Ratsche."],

    # OFENROHR (optional aber empfohlen)
    ["11.0", "Kaminzug-Durchfuehrung", "Hitze-Manschette Silikon 600 Grad C, Adapter 150 mm",
     1, "Set", 280, 280,
     "Ofenbedarf PT / Heimdesign DE / Bauhaus",
     "Falls Holzofen geplant. Silikon hitzefest, an Plane verschweissbar."],

    # WERKZEUG (siehe Sheet 12)
    ["12.0", "Werkzeug-Pauschale", "siehe Sheet 'DIY 12 Werkzeuge'",
     1, "Pauschal", 800, 800,
     "Mietkauf / lokale Vermietung",
     "Inkl. Tischkreissaege Miete, Bohrer-Aufsaetze, Hobel-Verleih, Sicherheitsausruestung."],

    # NEBENKOSTEN
    ["13.0", "Schattennetz Algarve", "6 m breit, 90% Sonnenschutz",
     60, "m2", 4, 240,
     "Agrar-Kooperative Algarve / Cooperativa Agricola",
     "WICHTIG fuer Algarve. Ueber Plane gespannt = verdoppelt Plane-Lebensdauer."],
    ["13.1", "Reserve / Unvorhergesehenes", "10% des Materialbudgets",
     1, "Pauschal", 1500, 1500,
     "-",
     "Schrauben-Nachkauf, Reparaturen, Lieferzuschlaege."],
]

# Style Sheet 9
for col, h in enumerate(headers9, 1):
    c = ws9.cell(row=1, column=col, value=h)
    c.fill = HEADER_FILL
    c.font = HEADER_FONT
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = BORDER
for r_idx, row in enumerate(material, 2):
    for c_idx, val in enumerate(row, 1):
        cell = ws9.cell(row=r_idx, column=c_idx, value=val)
        cell.alignment = Alignment(vertical="top", wrap_text=True)
        cell.border = BORDER
        if r_idx % 2 == 0:
            cell.fill = ALT_FILL

# Summenzeile
sum_row = len(material) + 2
total_eur = sum(r[6] for r in material)
ws9.cell(row=sum_row, column=6, value="GESAMTKOSTEN").font = Font(bold=True)
ws9.cell(row=sum_row, column=7, value=total_eur).font = Font(bold=True, size=12, color="2E5C8A")
for col in range(1, 10):
    ws9.cell(row=sum_row, column=col).border = BORDER

widths9 = [8, 28, 38, 9, 9, 14, 14, 38, 50]
for i, w in enumerate(widths9, 1):
    ws9.column_dimensions[get_column_letter(i)].width = w
ws9.row_dimensions[1].height = 30
ws9.freeze_panes = "A2"


# ---------- Sheet 10: DIY Beschaffungsquellen ----------
ws10 = wb.create_sheet("DIY 10 Quellen")
headers10 = ["Kategorie", "Anbieter", "Land", "Spezialisierung", "Website / Kontakt",
             "Liefert nach PT?", "Anmerkung Algarve-Tauglichkeit"]
quellen = [
    # HOLZ
    ["Holz", "Mil Martins & Irmao", "PT (Norden)", "Edelkastanie, Eiche, Pinie",
     "https://www.milmartinsirmao.pt/",
     "Innerhalb PT",
     "Etablierte Saegerei, Castanho Premium getrocknet. Anfrage per Mail."],
    ["Holz", "Industria de Madeiras", "PT", "Castanho-Handel, Industrie",
     "https://www.industriademadeiras.com/",
     "Innerhalb PT", "Castanho-Bohlen, Latten zuschneiden lassen moeglich."],
    ["Holz", "Albano Leite da Silva", "PT", "Allgemeine Saegerei",
     "https://www.albanoleitesilva.pt/", "Innerhalb PT",
     "Castanho + Pinie + Eukalyptus, KDI Behandlung verfuegbar."],
    ["Holz", "Maderterraneo", "ES", "Castanho Boden, Laerche, EU-Premium",
     "https://maderterraneo.com/", "Ja, Versand ES->PT",
     "Boden-Komponente Laerche/Castanho premium. Online-Bestellung."],
    ["Holz", "Siero Lam", "ES (Asturien)", "Castanho-Bohlen Massivholz",
     "https://www.sierolam.com/pt-pt/madeira-tabua/", "Ja, EU-Spedition",
     "Spanischer Marktfuehrer Edelkastanie. Versand PT moeglich."],
    ["Holz", "Holz Schiller AT", "AT", "Sibir. Laerche, Robinie",
     "https://www.holz-schiller.at/", "Ja, EU-Spedition 200-400 EUR",
     "Premium-Importeur, online bestellbar, USt-IdNr noetig fuer IC."],
    ["Holz", "Holz Henkel AT", "AT", "Robinie, Eiche, Esche",
     "https://www.holz-henkel.com/", "Ja, EU-Spedition",
     "Robinie fuer Tuer-Konstruktion, sehr salzwasserresistent."],
    ["Holz", "Holz Possling DE", "DE", "Sibir. Laerche, Uni-Material",
     "https://www.holzpossling.de/", "Ja, ueber Spedition",
     "Online-Shop mit DE-Preisen. Versand nach PT moeglich."],

    # KORK
    ["Daemmung-Kork", "Amorim Cork", "PT (national)", "Welt-Markt-fuehrer Kork",
     "https://www.amorim.com/", "Innerhalb PT, Hauptsitz",
     "PT-Heimspiel. Korkboden 30mm direkt vom Hersteller anfragen."],
    ["Daemmung-Kork", "Sofalca", "PT", "Expandierter Naturkork",
     "https://www.sofalca.pt/", "Innerhalb PT",
     "ICB-Korkplatten, fuer Plattform-Daemmung optimal."],

    # WOLLFILZ
    ["Daemmung-Wolle", "ISOLENA AT", "AT (Burgenland)", "Schafwoll-Daemmung Premium",
     "https://www.isolena.com/", "Ja, EU-Spedition + USt-IdNr",
     "100% Wolle Ionic Protect 16mm. Direkt-Bestellung."],
    ["Daemmung-Wolle", "DAEMWOOL AT", "AT", "Schafwoll-Jurtenfilz",
     "https://www.daemwool.at/", "Ja, EU-Spedition",
     "Spezialisiert auf Jurten-Anwendung. 8 oder 16mm verfuegbar."],
    ["Daemmung-Wolle", "Tobias Tumfart Schafwolldaemmung", "AT", "Jurtefilz 8/16 mm",
     "https://www.schafwolldaemmung.at/", "Ja, EU-Spedition",
     "Preis 8mm = 8,16 EUR/m2 (excl VAT). Massmade verfuegbar."],
    ["Daemmung-Wolle", "Wollwerkstatt AT", "AT", "Nadelfilz nach Mass",
     "https://www.wollwerkstatt.at/", "Ja", "Online-Konfigurator, Auf-Bestellung."],

    # PLANE / FACHBETRIEB
    ["Plane", "Sauleda S.A.", "ES (Barcelona)", "Solar Pro / Nautic Acryl 125+ Jahre",
     "https://sauleda.com/", "Ja, B2B-Anfrage",
     "Hersteller selbst, Distributoren-Liste in PT anfragen."],
    ["Plane", "Coartal (Sauleda PT-Distri)", "PT (zu pruefen)", "Sauleda-Stoffe Vertrieb",
     "https://www.coartal.com/files/Catlogo-Acrlico-Sauleda.pdf", "Ja",
     "Lokaler Vertrieb der Sauleda-Produkte, direkt anfragen."],
    ["Plane", "Toldum.es", "ES (Murcia)", "Sauleda + Markenstoffe online",
     "https://www.toldum.com/nuestros-productos/lonas-y-tejidos/sauleda/",
     "Ja", "Online-Vertrieb auch Stoff-Rollen, EU-Versand."],
    ["Plane", "Parà Tempotest", "IT", "Tempotest Marine Acryl 6J-Garantie",
     "https://www.para.it/en/", "Ja, Distributor-Liste",
     "Alternative zu Sauleda, oft etwas teurer aber Teflon-Finish."],
    ["Plane-Fachbetrieb", "Yurt Workshop Spain (Cadiar)", "ES (Granada)", "Jurten-Plane Massanfertigung",
     "http://yurtworkshop.es/", "Ja, ES->PT 600-1.200 EUR",
     "Rob Matthews kann Plane separat fertigen. Erfahrung mit Sebastian's Klima."],
    ["Plane-Fachbetrieb", "Atilla Jurtak (HU)", "HU", "Sauleda PVC + Airtex Acryl Set",
     "https://jurtak.hu/", "Ja, HU->PT 3.500 EUR Spedition",
     "Komplettes Cover 8m = 2.200 EUR netto, sehr guter Preis. Versand teurer."],
    # ALGARVE-LOKALE SATTLER (NEUE ERKENNTNIS - kritisch fuer DIY-Modify-Szenario)
    ["Plane-Algarve-LOKAL", "Dune Algarve Sailmakers", "PT (Vilamoura)",
     "Bootssegel + Marine-Persenningen + Schatten-Segel; Sauleda-Marine-Acryl",
     "https://www.dunesailmakers.com/en/",
     "Innerhalb PT (Algarve direkt)",
     "STATUS Mai 2026: ANFRAGE GESENDET (PT). Antwort offen. PRIORITAET #1 FUER JURTEN-COVER-COMMISSION. Beste fachliche Eignung (3D-Persenning-Erfahrung, Sauleda im Sortiment). Persoenlicher Besuch leicht moeglich von Aljezur."],
    ["Plane-Algarve-LOKAL", "Toldos Etapaveloz Lda", "PT (Loule, seit 2007)",
     "Toldos, Markisen, Sicht-/Schutz-beschattungen Spezialfertigung",
     "https://www.toldosetapaveloz.com/en", "Innerhalb PT (naehe Aljezur)",
     "STATUS Mai 2026: ANFRAGE GESENDET (PT). Antwort offen. Aelteste Loule-Werkstatt, geographisch am naechsten zu Aljezur. Toldos = Acryl-Markisen-Erfahrung = Sauleda-Verarbeitung. Frage: Erfahrung mit 3D-/Jurten-Form?"],
    ["Plane-Algarve-LOKAL", "Textilux", "PT (Algarve)",
     "Markisen + Stoff-Sonderanfertigung Veranden/Terrassen",
     "https://www.textilux.com/en/", "Innerhalb PT",
     "STATUS Mai 2026: ABGELEHNT - 'nao sao da nossa especialidade'. Klassischer Markisen-Hersteller, 3D-Persenning nicht im Programm. Datenpunkt: Toldos-Hersteller machen meist KEINE Jurten-Cover. Echte Kandidaten = Segelmacher (Dune) oder Atilla/Yurt Workshop als Cover-Only."],
    ["Plane-Algarve-LOKAL", "Toldos Chique", "PT (Algarve)",
     "Pergolas, Markisen, Outdoor-Beschattungs-Loesungen massgefertigt",
     "https://toldos-chique.com/awnings/", "Innerhalb PT",
     "BACKUP #3: auch automatisierte Markisen, gut etabliert."],
    ["Plane-Algarve-LOKAL", "Toldegarve", "PT (Algarve)",
     "Markisen vertikal/horizontal manuell/elektrisch",
     "https://www.toldegarve.pt/en/", "Innerhalb PT", "BACKUP #4"],
    ["Plane-Algarve-LOKAL", "algartoldos SOMBRIARTE Lda", "PT (Algarve)",
     "Markisen Herstellung", "https://www.toldosombriarte.com/",
     "Innerhalb PT", "BACKUP #5"],
    ["Plane-Algarve-LOKAL", "Sombra&Constroi", "PT (Portimao)",
     "Toldos, Estores, Pergolas",
     "Caminho Horta de Sao Francisco 8500-145 Portimao",
     "Innerhalb PT", "BACKUP #6 - in Portimao naehe Aljezur"],
    ["Plane-Fachbetrieb-EU", "Lokaler Persenningmacher PT (Backup-Suche)", "PT", "Auftragsnaeherei",
     "Suche: 'oficina de toldos Algarve', 'velejaria PT'", "Innerhalb PT",
     "Generelle Suchbegriffe falls obige Liste nicht weiterhilft. Skizzen mitbringen, Maße ueber bereits aufgebautes Skelett nehmen."],

    # TONO (Kronenring)
    ["Tono", "Camping Yurts", "UK", "Tono fertig gebohrt",
     "https://www.campingyurts.com/yurt-parts/", "Ja, UK->PT post-Brexit Zoll",
     "Pre-drilled, sealed, varnished. Direkt online bestellbar."],
    ["Tono", "Groovy Yurts (CA, exportiert global)", "CA", "Sibir.-Kiefer Tono Mongolian-style",
     "https://www.groovyyurts.com/accessories-and-add-ons/toono-dome", "Versand teuer",
     "Authentisch mongolisch, eher teuer fuer EU."],
    ["Tono", "FAMTENTS", "CZ", "Tschechische Laerche Tono",
     "https://www.famtents.com/yurts", "Ja, EU-Spedition",
     "Premium-Hersteller EU, Tono separat anfragen."],
    ["Tono", "Atilla Jurtak (HU)", "HU", "Tono separat moeglich (Anfrage)",
     "https://jurtak.hu/", "Ja, HU-Versand", "Custom in Hartholz. Anfragen ob er nur Tono+Uni verkauft."],
    ["Tono", "Adorjan Jurta (HU)", "HU", "Tradi. mongolisch Tono",
     "https://adorjan-jurta.hu/", "Ja, HU-Versand", "Hartholz, traditionell, lange Lieferzeit (bookings 1J voraus)."],

    # SCHRAUBPFAHL / FOUNDATION
    ["Schraubpfaehle", "Krinner GmbH", "DE", "Schraubpfaehle Marktfuehrer",
     "https://www.krinner.de/", "Ja, EU-Spedition",
     "KSF-M Profil 800-1000 mm, alle EU-baurelevanten Zulassungen."],
    ["Schraubpfaehle", "Stop Schraubfundamente", "DE", "Alternative zu Krinner",
     "https://www.stop-schraubfundamente.de/", "Ja, EU-Spedition", "Etwas guenstiger als Krinner."],
    ["Schraubpfaehle", "Bauhaus / Leroy Merlin PT", "PT", "Krinner-Lager+Vertrieb",
     "https://www.bauhaus.pt/  |  https://www.leroymerlin.pt/", "Innerhalb PT",
     "Lokales Lager, schneller als DE-Direkt."],

    # BESCHLAEGE
    ["Beschlaege", "Wuerth Portugal", "PT", "Industrie-Schrauben/Beschlaege",
     "https://www.wurth.pt/", "Innerhalb PT, oft Same-Day",
     "Edelstahl A4 fuer Algarve obligatorisch. Schraubenmusterbox empfehlenswert."],
    ["Beschlaege", "Bauhaus PT", "PT", "DIY-Sortiment",
     "https://www.bauhaus.pt/", "Innerhalb PT",
     "Gut fuer Standard-Sortiment. A4-Schrauben anfragen."],
    ["Beschlaege", "Brico Marche", "PT", "DIY landesweit",
     "https://www.bricomarche.pt/", "Innerhalb PT", "Bauhaus-Alternative."],
    ["Beschlaege", "Acastillaje (Bootsbedarf)", "PT", "Marine-Edelstahl Spannschloss",
     "Suche 'acastillaje algarve'", "Lokal", "Spezialisiert auf Salzluft. Spannschlossglieder, Edelstahl-Karabiner."],

    # WERKZEUG
    ["Werkzeug-Verleih", "Bauhaus PT Mietservice", "PT", "Tischkreissaege, Bohrer-Profi",
     "https://www.bauhaus.pt/leihservice", "Lokal", "Tages-/Wochenmiete. Tischkreissaege 25-40 EUR/Tag."],
    ["Werkzeug-Verleih", "Lokale Aluguer-Equipamentos Algarve", "PT", "Bau-Maschinen-Verleih",
     "Suche 'aluguer equipamentos Algarve'", "Lokal", "Oft Wochenpakete guenstiger."],

    # LITERATUR / PLANS
    ["Literatur", "Paul King - The Complete Yurt Handbook", "UK", "Standard-Werk Bauanleitung",
     "Amazon / Yurtinfo.org", "PDF/Print", "ca. 25 EUR. Pflichtlektuere fuer DIY. Kein 8m-Plan, aber Skalierungs-Anleitung."],
    ["Literatur", "Paul King - Build Your Own Yurt (kostenloses PDF)", "UK", "Frueheres Werk, online",
     "https://azinelibrary.org/approved/build-your-own-yurt-1.pdf",
     "PDF kostenlos", "Aelter aber kostenlos. 3m Yurte als Skalierungs-Basis."],
    ["Literatur", "SimplyDifferently Yurt Calculator", "Web", "Onlinerechner Bemassung",
     "https://simplydifferently.org/Yurt_Notes", "kostenlos",
     "GOLD-STANDARD: Exakte Bohrwinkel, Latten-Anzahlen, Uni-Laengen fuer beliebige Durchmesser."],
    ["Literatur", "Yurt Forum Community", "Web", "DIY-Erfahrungsaustausch",
     "https://www.yurtforum.com/", "kostenlos", "Stelle Fragen, Bilder, Detail-Probleme."],
    ["Literatur", "Doit Yurtself", "Web", "DIY-Schritte, Tipps",
     "https://doityurtself.com/", "kostenlos", "Schreibblog mit Foto-Doku eines Komplett-Baus."],
    ["Literatur", "Yurt Plans Library DryUrts", "Web", "Plaene zum Kauf",
     "https://www.dryurts.com/yurt-plans.html", "ca. 20-50 USD", "Mehrere Yurten-Plaene downloadbar."],

    # SCHATTENNETZ
    ["Schattennetz", "Cooperativa Agricola Algarve", "PT (lokal)", "Agrar-Schattennetz 90%",
     "Lokal vor Ort suchen z.B. Sao Bras de Alportel, Lagos", "Innerhalb PT",
     "Wichtigster Algarve-Tipp! 6m Breite Standard, 4-8 EUR/m2."],
]

for col, h in enumerate(headers10, 1):
    c = ws10.cell(row=1, column=col, value=h)
    c.fill = HEADER_FILL
    c.font = HEADER_FONT
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = BORDER
for r_idx, row in enumerate(quellen, 2):
    for c_idx, val in enumerate(row, 1):
        cell = ws10.cell(row=r_idx, column=c_idx, value=val)
        cell.alignment = Alignment(vertical="top", wrap_text=True)
        cell.border = BORDER
        if r_idx % 2 == 0:
            cell.fill = ALT_FILL
widths10 = [16, 30, 14, 32, 50, 28, 50]
for i, w in enumerate(widths10, 1):
    ws10.column_dimensions[get_column_letter(i)].width = w
ws10.row_dimensions[1].height = 30
ws10.freeze_panes = "A2"


# ---------- Sheet 11: DIY Arbeitsschritte ----------
ws11 = wb.create_sheet("DIY 11 Arbeitsschritte")
headers11 = ["Phase", "Woche", "Aufgabe", "Dauer", "Werkzeug", "Material", "Output"]
schritte = [
    ["Vorbereitung", "Woche 0", "Pläne und Buch besorgen, SimplyDifferently durchrechnen",
     "2-3 Tage", "PC + Drucker",
     "Paul King Buch, Bemassungs-Ausdrucke",
     "Detail-Plan auf Papier mit allen Massen"],
    ["Vorbereitung", "Woche 0", "Werkstattflaeche organisieren",
     "1 Tag", "Massband",
     "10x10 m flacher trockener Untergrund",
     "Khaana kann komplett ausgebreitet werden"],
    ["Vorbereitung", "Woche 0", "Material bestellen (lange Lieferzeit zuerst)",
     "1 Tag",
     "PC",
     "Tono fertig (4-6 Wochen Lieferzeit), Wollfilz, Schraubpfaehle, Sauleda Plane Massanfertigung",
     "Bestellungen raus, Lieferzeit-Plan"],
    ["Vorbereitung", "Woche 0", "Werkzeug mieten/kaufen",
     "1 Tag", "-",
     "Tischkreissaege, Stichsaege, Akkuschrauber 18V, Bohrmaschine, Hobel, Bohraufsatz, Schraubzwingen",
     "Werkstatt aufgeruestet"],
    ["Standort", "Woche 1", "Standort markieren, Schraubpfaehle setzen",
     "1 Tag", "Erdraketenwerkzeug oder hydraulische Eindrehhilfe",
     "8 Krinner-Pfaehle, Wasserwaage",
     "Plattform-Fundament ueber Boden, perfekt nivelliert"],
    ["Standort", "Woche 1", "Plattform-Tragwerk bauen",
     "2 Tage", "Tischkreissaege, Akkuschrauber",
     "Kiefer KDI 50x200 Tragbalken, Edelstahl-Schrauben A4",
     "Tragwerk fertig, Doppel-T-Rahmen + Radial-Balken"],
    ["Standort", "Woche 1", "Plattform-Daemmung + Decke",
     "1 Tag", "Cuttermesser, Akkuschrauber",
     "Geo-Vlies, PE-Folie, Kork 30 mm, Laerchen-Deck T+G 27 mm",
     "Plattform begehbar, isoliert"],

    ["Khaana-Bau", "Woche 2", "Edelkastanie-Latten zuschneiden und hobeln",
     "2 Tage", "Tischkreissaege, Hobel, Schleifer",
     "280 Latten 30x10x2100 mm",
     "Alle Latten fertig, gleichmaessig"],
    ["Khaana-Bau", "Woche 2-3", "Bohrloch-Lehre bauen + alle Latten bohren",
     "2 Tage", "Stationaer-Bohrer, selbstgebaute Bohrlehre",
     "M5-Bohrer 5,2 mm, Schraubzwingen",
     "Alle Latten gebohrt, max. 1 mm Toleranz"],
    ["Khaana-Bau", "Woche 3", "Khaana-Sektionen zusammensetzen (Bolzen+Mutter)",
     "3-4 Tage", "Akkuschrauber + 8mm-Steckschluessel",
     "M5x30 Edelstahl Bolzen + Hutmuttern",
     "8 Khaana-Sektionen klappbar, Vorabtest stehen lassen"],
    ["Khaana-Bau", "Woche 3", "Khaana-Test: alle Sektionen verbinden + auf Plattform stellen",
     "1 Tag", "2 Personen",
     "Hanf-Schnur fuer temporaere Verbindung",
     "Khaana steht im Kreis, Tono-Hoehe stimmt"],

    ["Tono+Uni", "Woche 4", "Tono-Lieferung pruefen + nachbohren falls noetig",
     "1 Tag", "Akkuschrauber, Schleifer",
     "Sicherheitslack",
     "Tono fertig fuer Uni-Aufnahme"],
    ["Tono+Uni", "Woche 4", "Uni-Stangen zuschneiden, anschraegen",
     "2 Tage", "Tischkreissaege, Anschlagwinkel, Hobel",
     "65 Laerche 30x30x4000 mm",
     "Alle 60 Uni-Stangen einsatzbereit"],
    ["Tono+Uni", "Woche 4", "Uni-Schlaufen am Khaana-Wandkronenrand",
     "1 Tag", "Sattler-Werkzeug, Schraubendreher",
     "Lederband oder Hanfschlaufen",
     "Sattel-Punkte fuer 60 Uni alle 60 cm"],
    ["Tono+Uni", "Woche 4", "Test-Aufbau Tono+Uni (4-8 Helfer noetig!)",
     "1 Tag", "Leitern, 4-8 Helfer",
     "-",
     "Erste komplette Skelett-Stehprobe"],

    ["Spannband+Tuer", "Woche 5", "Kuriye-Spannband installieren",
     "1 Tag", "Spannschlossglieder, Edelstahl-Karabiner",
     "PP-Gurt 50mm 30m, 4 Spannschloesser",
     "Wand-Spannband fixiert, Vorspannung gleichmaessig"],
    ["Spannband+Tuer", "Woche 5", "Tuerrahmen einsetzen+ausrichten",
     "1 Tag", "Wasserwaage, Akkuschrauber",
     "Robinie-Rahmen 100x100 mm, A4-Schrauben",
     "Tuer steht 90 Grad, schwingt frei"],
    ["Spannband+Tuer", "Woche 5", "Tuerblatt einhaengen+justieren",
     "1 Tag", "Bandscharnier-Schrauber",
     "Tuerblatt 2x 90x200 cm + Beschlaege",
     "Tuer schliesst sauber, kein Spalt"],

    ["Daemmung+Plane", "Woche 6", "Innenliner befestigen",
     "1 Tag", "Tacker, Naehnadel",
     "Baumwoll-Sergeant 65 m2",
     "Wand- und Dachsegmente innen mit Liner bezogen"],
    ["Daemmung+Plane", "Woche 6", "Wollfilz auflegen (Wand+Dach)",
     "1 Tag", "Cuttermesser",
     "Wollfilz 16 mm, 105 m2",
     "Innenseite komplett gedaemmt"],
    ["Daemmung+Plane", "Woche 6", "Solitex-Membran ueber Daemmung",
     "1 Tag", "Cuttermesser, Tacker",
     "Solitex 115 m2",
     "Dampfsperre installiert, ueberlappend +10 cm"],
    ["Daemmung+Plane", "Woche 6", "Aussenplane aufziehen (Sauleda)",
     "1-2 Tage", "Saugen+Leitern, 4 Helfer",
     "Komplett-Plane 110 m2 inkl. Tunnel+Saeume",
     "Yurte wetterfest!"],

    ["Sturm-Kit", "Woche 7", "Erdanker Sturm-Kit einschrauben",
     "1 Tag", "Erdraketenwerkzeug",
     "8 M16 Anker 1.000 mm, Stahlseil 6 mm",
     "Sturm-Verankerung aktiv"],
    ["Sturm-Kit", "Woche 7", "Ratschen-Riemen ueber Dach spannen",
     "0,5 Tag", "-",
     "8 PP-Ratschen 5 Tonnen",
     "Plane gegen Aufwehung gesichert"],
    ["Sturm-Kit", "Woche 7", "Kaminzug-Durchfuehrung einbauen falls geplant",
     "1 Tag", "Akkuschrauber, Silikon-Spachtel",
     "Kaminzug-Manschette 600 Grad",
     "Holzofen einsetzbar"],

    ["Finale", "Woche 8", "Innen-Ausbau: Bett, Kueche, Boden-Teppich",
     "Variabel", "-", "-", "Wohnfertig"],
    ["Finale", "Woche 8", "Schattennetz ueber Dach spannen",
     "0,5 Tag", "Leitern",
     "Agrar-Schattennetz 90% 60 m2",
     "Algarve-Hitze entschaerft, Plane-Lebensdauer +X Jahre"],
    ["Finale", "Woche 8", "Foto-Doku + Wartungs-Checkliste",
     "0,5 Tag", "Kamera", "-",
     "Versicherung+Erinnerung+Dokumentation"],
]

for col, h in enumerate(headers11, 1):
    c = ws11.cell(row=1, column=col, value=h)
    c.fill = HEADER_FILL
    c.font = HEADER_FONT
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = BORDER
for r_idx, row in enumerate(schritte, 2):
    for c_idx, val in enumerate(row, 1):
        cell = ws11.cell(row=r_idx, column=c_idx, value=val)
        cell.alignment = Alignment(vertical="top", wrap_text=True)
        cell.border = BORDER
        if r_idx % 2 == 0:
            cell.fill = ALT_FILL
widths11 = [16, 12, 50, 12, 32, 38, 38]
for i, w in enumerate(widths11, 1):
    ws11.column_dimensions[get_column_letter(i)].width = w
ws11.row_dimensions[1].height = 30
ws11.freeze_panes = "A2"


# ---------- Sheet 12: DIY Werkzeuge ----------
ws12 = wb.create_sheet("DIY 12 Werkzeuge")
headers12 = ["Werkzeug", "Pflicht/Optional", "Kauf/Miete", "Preis EUR", "Bezugsquelle", "Bemerkung"]
werkzeug = [
    ["Tischkreissaege", "PFLICHT", "Miete empfohlen", 200,
     "Bauhaus PT Mietservice", "Hochpraezise Latten-Schnitte. 25-40 EUR/Tag x 7 Tage."],
    ["Stichsaege", "PFLICHT", "Kauf", 80, "Bauhaus / Leroy Merlin", "Fuer Tono-Korrekturen, Tuer-Ausschnitte. Festool/Bosch."],
    ["Akkuschrauber 18V", "PFLICHT", "Kauf", 180, "Bauhaus", "2 Akkus empfohlen. Bosch GSR oder Makita."],
    ["Bohrmaschine stationaer", "PFLICHT", "Miete oder Kauf", 150,
     "Bauhaus Miete o. Globus Baumarkt Kauf", "Fuer 280 Latten-Bohrungen praezise. Saeulenbohrmaschine ideal."],
    ["Bohrlehre selbstgebaut", "PFLICHT", "Eigenbau", 20, "Sperrholz-Rest",
     "MUSS-HAVE: 5 mm Bohrloch, Anschlag fuer Wiederholgenauigkeit. Selbst in 1 Std gebaut."],
    ["Bohraufsaetze M5+M6", "PFLICHT", "Kauf", 25, "Wuerth PT", "Edelstahl-tauglich, scharf."],
    ["Akku-Hobel ELEKTRISCH", "PFLICHT", "Miete", 100, "Bauhaus PT", "Latten-Oberflaeche glatt machen. 2 Wochen Miete."],
    ["Sortimentskasten Schrauben A4", "PFLICHT", "Kauf", 80, "Wuerth", "M5, M6, Spax 4x50, 5x80 Edelstahl A4."],
    ["Schraubzwingen (Set 8 Stueck)", "PFLICHT", "Kauf", 60, "Bauhaus", "Beim Khaana-Zusammenbau unentbehrlich."],
    ["Wasserwaage 2m", "PFLICHT", "Kauf", 35, "Bauhaus", "Plattform-Aufbau, Tuer-Justierung."],
    ["Massband 5m + 8m", "PFLICHT", "Kauf", 25, "Bauhaus", "Doppelpack."],
    ["Erdraketen-Werkzeug Krinner", "PFLICHT", "Miete", 80,
     "Mit Schraubpfahl-Lieferung mitbestellen", "Spezial-Werkzeug fuer Krinner-Schrauben. 8 Pfaehle 1 Tag."],
    ["Sicherheitsausruestung (Brille+Gehoer+Maske)", "PFLICHT", "Kauf", 60,
     "Bauhaus", "Pflicht fuer Saege+Hobel+Plane-Naehen."],

    ["Leiter 3 m", "Empfohlen", "Miete oder Kauf", 90,
     "Bauhaus", "Tono-Aufbau (4,20 m hoch). Klappleiter ausreichend."],
    ["4-8 Helfer fuer Aufbau", "Empfohlen", "Bekanntenkreis",
     0, "-", "Tono+Uni-Aufbau allein nicht moeglich. Pizza & Bier statt Geld."],
    ["Sattler-Werkzeug fuer Schlaufen", "Optional", "Kauf", 40,
     "Sattlerei PT", "Falls Du Lederband statt Bolzen am Khaana willst."],
    ["Industrie-Naehmaschine", "OPTIONAL (nur wenn selber naehen)", "Miete", 200,
     "Sattlerei lokal", "Du wolltest Plane bei Fachbetrieb beauftragen -> nicht noetig."],
    ["Tacker fuer Innenliner", "Empfohlen", "Kauf", 30, "Bauhaus", "Hand-Tacker mit 10 mm Klammern."],
    ["Schraubendreher-Set", "Empfohlen", "Kauf", 25, "Bauhaus", "Wera/Wiha Premium - lebenslange Investition."],
    ["Beize/Imprägnierung Holz (Hartoel)", "Empfohlen", "Kauf", 80,
     "Bauhaus Auro/Osmo", "Auro/Osmo Hartoel statt Lack - atmungsaktiv. ca. 4 Dosen fuer 8m."],
]
for col, h in enumerate(headers12, 1):
    c = ws12.cell(row=1, column=col, value=h)
    c.fill = HEADER_FILL
    c.font = HEADER_FONT
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = BORDER
for r_idx, row in enumerate(werkzeug, 2):
    for c_idx, val in enumerate(row, 1):
        cell = ws12.cell(row=r_idx, column=c_idx, value=val)
        cell.alignment = Alignment(vertical="top", wrap_text=True)
        cell.border = BORDER
        if r_idx % 2 == 0:
            cell.fill = ALT_FILL

# Summe
total_tool = sum(r[3] for r in werkzeug)
sum_row = len(werkzeug) + 2
ws12.cell(row=sum_row, column=3, value="WERKZEUG GESAMT").font = Font(bold=True)
ws12.cell(row=sum_row, column=4, value=total_tool).font = Font(bold=True, size=12, color="2E5C8A")
for col in range(1, 7):
    ws12.cell(row=sum_row, column=col).border = BORDER

widths12 = [38, 22, 18, 12, 38, 50]
for i, w in enumerate(widths12, 1):
    ws12.column_dimensions[get_column_letter(i)].width = w
ws12.row_dimensions[1].height = 30
ws12.freeze_panes = "A2"


# ---------- Sheet 13: DIY Kostenrechnung ----------
ws13 = wb.create_sheet("DIY 13 Kostenrechnung")
ws13.column_dimensions["A"].width = 36
ws13.column_dimensions["B"].width = 16
ws13.column_dimensions["C"].width = 60

t = ws13.cell(row=1, column=1, value="DIY 8m Jurte - Vollkostenrechnung + Vergleich Casa dos Sonhos")
t.font = TITLE_FONT
ws13.merge_cells("A1:C1")

kosten = [
    ("Block A: Material (siehe Sheet 9)", "", ""),
    ("Khaana Edelkastanie + Bolzen", 1455, "PT-Saegerei. Hauptbauteil Wand."),
    ("Tono fertig + Versand", 1000, "Hartholz-Kronenring + Spedition aus DE/HU."),
    ("Uni Sibir. Laerche + Befestigung", 982, "60 Dachstangen incl. Niete + Lederschlaufen."),
    ("Kuriye Spannband + Spannschloesser", 192, "Wand-Zugring."),
    ("Tuer + Tuerbeschlaege A4", 1350, "Robinie/Eiche, Doppeltuer mit Doppelverglasung."),
    ("Plattform komplett (Pfaehle, Tragwerk, Daemmung, Decke, Schrauben)", 4040,
     "Schraubpfaehle Krinner + Kiefer KDI + Kork 30mm + Laerchen-Deck T+G."),
    ("Daemmung Wand+Dach (Schafwollfilz Isolena 16mm)", 1680,
     "105 m2 inkl. Aufschlag, Versand AT->PT."),
    ("Innenliner Baumwolle", 780, "65 m2 Sergeant 200 g/m2."),
    ("Solitex-Membran (Diffusionsoffen)", 690, "115 m2 fuer Algarve-Klima essentiell."),
    ("Aussenplane Sauleda Massanfertigung", 3500,
     "Komplett-Set Dach+Wand+Kuppel-Cover, Sauleda Solar Pro Acryl."),
    ("Polycarbonat-Kuppel oeffenbar", 750, "1,2 m, klar, oeffnungs-faehig."),
    ("Sturm-Kit (Erdanker + Ratschen)", 456, "ALGARVE OBLIGATORISCH."),
    ("Kaminzug-Durchfuehrung", 280, "Falls Holzofen geplant."),
    ("Schattennetz Algarve", 240, "Verlaengert Plane-Lebensdauer drastisch."),
    ("Reserve (10%)", 1500, "Schraubenkauf, Lieferzuschlaege, Reparaturen."),
    ("Werkzeug (Miete+Kauf, siehe Sheet 12)", 1378, "Tischkreissaege Miete + Akku-Tools Kauf."),
    ("", "", ""),
    ("ZWISCHENSUMME MATERIAL+WERKZEUG", 20273, "Komplett-Materialliste + Werkzeug-Pauschale."),
    ("", "", ""),

    ("Block B: Versand + Transport", "", ""),
    ("Holz-Versand DE/AT->PT (Laerche, Robinie)", 600, "Spedition Sammelladung 2-3 Wochen."),
    ("Filz-Versand AT->PT (Isolena)", 250, "Paketdienst, ~100 kg."),
    ("Schraubpfaehle/Schrauben Inland PT", 0, "Lokaler Einkauf."),
    ("Plane-Lieferung", 150, "Falls Auftrag bei ES-Sattler statt PT."),
    ("Tono-Versand", 250, "Sperrgut DE/HU->PT."),
    ("", "", ""),
    ("ZWISCHENSUMME VERSAND", 1250, ""),
    ("", "", ""),

    ("Block C: Eigene Arbeitszeit (nicht monetarisiert)", "", ""),
    ("Geschaetzter Zeitaufwand", "350-450 Std", "Bei mittlerem Skill, einzeln + 4-8 Helfer-Tage."),
    ("Zeit-Aequivalent bei 30 EUR/h (rechnerisch)", "10.500-13.500 EUR", "Bewertung: Hobby/Eigenleistung. Nicht in Kostenrechnung."),
    ("", "", ""),

    ("GESAMTKOSTEN DIY (Material+Werkzeug+Versand)", 21523,
     "Inklusive Werkzeug und Algarve-Sturm-Kit. OHNE eigene Arbeitszeit."),

    ("", "", ""),
    ("Vergleich mit Casa dos Sonhos (Hersteller)", "", ""),
    ("Casa dos Sonhos 7,3m (42 m2)", 15750, "Lieferung+Aufbau PT INKL. Aber: keine Plattform, kein Sturm-Kit, kein Kaminzug."),
    ("Casa Plattform separat", 4000, "Wenn man dasselbe Niveau will (Schraubpfaehle + Kork + Laerchen-Deck)."),
    ("Casa Sturm-Kit (Optionspreis offen)", 800, "Geschaetzt - bei Casa konkret abfragen."),
    ("Casa Kaminzug (Optionspreis offen)", 600, "Geschaetzt - bei Casa konkret abfragen."),
    ("Casa Schattennetz", 240, "Selbst lokal kaufen."),
    ("CASA TOTAL APPLES-TO-APPLES (7,3m)", 21390,
     "Ueberraschung: DIY 8m und Casa 7,3m kommen aufs Gleiche raus!"),

    ("", "", ""),
    ("Casa 9,1m (64 m2)", 18000, "Inkl. Aufbau, ohne Plattform/Sturm-Kit/Kamin."),
    ("Casa 9,1m + Plattform + Sturm-Kit + Kamin", 23640, "Komplett-Setup 9,1m vs DIY 8m."),

    ("", "", ""),

    ("FAZIT ENTSCHEIDUNG DIY VS CASA", "", ""),
    ("Material-Differenz DIY 8m vs Casa 7,3m all-in", "-130 EUR (Casa minimal teurer)",
     "DIY ist NICHT signifikant guenstiger wenn alles ehrlich gerechnet wird."),
    ("Was DIY trotzdem fuer dich bringen kann", "Lernen + 50m2 statt 42m2 + 8m statt 7,3m + volle Kontrolle Materialwahl + keine 100% Vorauszahlung",
     "Du baust 8m statt 7,3m = 8 m2 mehr Wohnflaeche fuer denselben Preis."),
    ("Was DIY KOSTET (jenseits Geld)", "350-450 Stunden Lebenszeit + Risiko Konstruktionsfehler + Werkzeug-Lernkurve",
     "Casa liefert Sicherheit (10-J. UV-Garantie schriftlich)."),
    ("Empfehlung", "Hybrid pruefen: Casa Plane+Daemmung kaufen, Plattform+Aufbau selbst",
     "Frage Casa, ob sie Komponenten ohne Aufbau verkaufen. Spart 30-50% Hersteller-Marge."),

    ("", "", ""),
    ("NEU MAI 2026: DIY-MODIFY SZENARIO A (CHEAP BASE + PREMIUM UPGRADES)", "", ""),
    ("Konzept", "Skelett guenstig + Plane Premium + lokal verbauen", "Du sparst beim Holzrahmen (Adorjan HU 1.750 EUR fuer 6m), zahlst Premium bei Plane (Sauleda Solar Pro lokal in Algarve gefertigt) und Daemmung (PET-Vlies / Wolle), baust Plattform selbst."),
    ("Adorjan 6m Rahmen (Khaana + Tono + Uni + Tuer)", 1750, "Ungarn, Direktanfrage. Lieferzeit oft 1J Wartezeit aber neuer/aktueller Preis erfragen."),
    ("Versand HU -> Algarve (Sammelladung)", 1200, "Spedition uebergross, 2-3 Wochen Lieferzeit."),
    ("Sauleda Solar Pro Plane BEI LOKALEM ALGARVE-SATTLER", 2200,
     "GAME-CHANGER: Dune Algarve Sailmakers in Vilamoura oder Toldos Etapaveloz in Loule. Schaetzung 1.800-2.500 EUR. PERSOENLICHE Auftragsfertigung + lokale Reparatur."),
    ("PET-Vlies-Daemmung 8 cm (60 m2 Wand+Dach)", 500,
     "Tobias Tumfart AT 8,16 EUR/m2 netto = 489 EUR fuer 60 m2. Modern, schimmelfest. Vergleichbar Bēt-Yurts-Daemmung."),
    ("Solitex-Diffusions-Membran", 500, "Pro Clima AT, 115 m2. Verhindert Kondensat."),
    ("Tuer-Upgrade falls noetig (Glas-Doppeltuer)", 400, "Adorjan-Standard-Tuer eventuell zu einfach - Upgrade pruefen."),
    ("Plattform DIY (Schraubpfaehle + Holz)", 3000,
     "Vereinfachte Plattform: Krinner-Pfaehle + Kiefer KDI + einfaches Laerchen-Deck. Kork-Daemmung weglassen wenn Budget eng."),
    ("Sturm-Verankerung Algarve", 400, "Erdanker + Ratschen lokal."),
    ("Schattennetz Algarve", 240, "Lokale Agrar-Kooperative."),
    ("Reserve / Verschnitt / Schrauben", 500, ""),
    ("ZWISCHENSUMME SZENARIO A", 10690, "OHNE Werkzeugkosten (Sebastian hat ggf. bereits Werkzeug oder mietet))."),
    ("Werkzeug-Block bei DIY", 1378, "Wenn neu zu beschaffen, siehe Sheet DIY 12."),
    ("TOTAL SZENARIO A (mit Werkzeug)", 12068,
     "NEU EMPFEHLUNG #1 SEIT CASA-SLOT WEG: ca. 11.000-12.500 EUR all-in fuer 6m Algarve-Premium-Setup."),

    ("", "", ""),
    ("VERGLEICH SZENARIO A vs CASA 2027er Slot", "", ""),
    ("DIY Modify Szenario A 6m", 10690, "Material+Versand. Ohne Werkzeug-Kauf (haben oder mieten)."),
    ("Casa 6,1m (2027 Slot)", 14890, "9.700 Casa + 4.000 Plattform DIY + 240 Schattennetz + 950 Optionen (Tuer/Kamin)."),
    ("ERSPARNIS SZENARIO A", -4200,
     "ca. 4.000-4.500 EUR weniger als Casa 2027, mit objektiv besseren Materialien (Sauleda Solar Pro vs unspezifiziertes Casa-Canvas). Du gibst Casa's 10-J. UV-Garantie auf, gewinnst dafuer Sauleda-Marken-Garantie (5-10 J. je nach Hersteller) + sofortige Verfuegbarkeit + lokale Reparatur-Quelle."),
]

row = 2
for entry in kosten:
    a, b, c = entry
    if a and (b == "" and c == ""):
        cell = ws13.cell(row=row, column=1, value=a)
        cell.font = Font(bold=True, size=12, color="2E5C8A")
        ws13.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
        ws13.row_dimensions[row].height = 22
    elif not a and b == "" and c == "":
        ws13.row_dimensions[row].height = 8
    else:
        c1 = ws13.cell(row=row, column=1, value=a)
        c1.alignment = Alignment(wrap_text=True, vertical="top")
        c1.border = BORDER
        c2 = ws13.cell(row=row, column=2, value=b)
        c2.alignment = Alignment(horizontal="right", vertical="top")
        c2.border = BORDER
        if isinstance(b, (int, float)):
            c2.number_format = '#,##0 "EUR"'
            if a.startswith("GESAMT") or a.startswith("ZWISCHENSUMME") or a.startswith("CASA TOTAL"):
                c1.font = Font(bold=True)
                c2.font = Font(bold=True, size=12, color="2E5C8A")
        c3 = ws13.cell(row=row, column=3, value=c)
        c3.alignment = Alignment(wrap_text=True, vertical="top")
        c3.border = BORDER
        ws13.row_dimensions[row].height = max(20, min(45, 15 + len(str(c)) // 8))
    row += 1
ws13.freeze_panes = "A2"


# ---------- Sheet 14: 8m DIY-Modify Vollkalkulation (Sebastian Mai 2026) ----------
ws14 = wb.create_sheet("DIY 14 8m Modify Kalk")
ws14.column_dimensions["A"].width = 38
ws14.column_dimensions["B"].width = 16
ws14.column_dimensions["C"].width = 62

t = ws14.cell(row=1, column=1, value="DIY 8m Yurt fuer Algarve - 'Modify Cheap Base' mit Algarve-Sattler (Sebastians Praeferenz Mai 2026)")
t.font = TITLE_FONT
ws14.merge_cells("A1:C1")
ws14.row_dimensions[1].height = 30

kalk_8m = [
    ("AUSGANGSLAGE", "", ""),
    ("Sebastians Vorgabe", "8 m Durchmesser (50 m2)", "Begrundung: deutlich groesser als Casa 7,3m (42m2). Wohnflaeche-Plus = 8 m2 (ca. ein Mini-Schlafalkoven)."),
    ("Verfuegbare 8m-Anbieter", "Yourtepoque, Atilla, Adorjan, jurte24, YourTent CZ, Yurts4ever, DIY", "Casa und Bēt Yurts haben KEIN 8m im Sortiment. Casa naechste: 7,3m oder 9,1m."),
    ("Budget-Realitaet", "8m all-in ca. 17.000-22.000 EUR", "Original Budget 15.000 EUR reicht NICHT fuer 8m. Bei 8m muss Budget auf 17-18k flexen ODER signifikante Kompromisse."),

    ("", "", ""),
    ("KONZEPT DIY-MODIFY 8M", "", ""),
    ("Idee", "Cheap Frame (HU) + Premium Plane (Sauleda lokal Algarve) + Selbst-Aufbau", "Wiederholt das 6m-Schema von Sheet 13, hochskaliert auf 8m. Hauptvorteile: lokaler Algarve-Sattler fuer Reparatur, volle Materialkontrolle, sofortige Verfuegbarkeit, kein Slot-Warten."),
    ("Anbieter Frame", "Adorjan Jurta HU / Atilla HU / Yurts4ever RO", "Anfragen parallel. Adorjan + Atilla aktuell mit Frame-only-Anfrage angeschrieben Mai 2026."),
    ("Anbieter Plane", "Dune Algarve Sailmakers (Vilamoura) / Toldos Etapaveloz (Loule) / Textilux", "3 Anfragen Mai 2026 abgesendet. Erwartung: 1-3 Wochen Antwortzeit."),

    ("", "", ""),
    ("MATERIAL-KOSTEN 8M", "", ""),
    ("Adorjan 8m Frame-only (Khaana + Tono + Uni + Tuer)", "5.000",
     "SCHAETZUNG (6m bei 1.750 - 8m Skalierung mit Material-Mehrbedarf + groesserem Tono). Konkreter Preis ueber laufende Anfrage."),
    ("Versand HU -> Algarve 8m (heavier)", "2.000",
     "Spedition Uebergross, 2-3 Wochen. 8m benoetigt groesseres Fahrzeug als 6m."),
    ("Sauleda Solar Pro Plane 8m (~150 m2 Stoff)", "4.500",
     "DIE GROESSTE POSITION: 8m hat ca. 60 m2 Dach (Kegelmantel) + 50 m2 Wand + Verschnitt + Saum-/Tunnel-Verstaerkung = ca. 150 m2 Stoffbedarf. Bei 25-30 EUR/m2 + 50-100% Verarbeitungs-Aufschlag = 4.000-5.000 EUR komplett massgefertigt. Komplettes Set: Dach + Wand + Kuppel-Cap."),
    ("PET-Vlies-Daemmung 8 cm (~110 m2 Wand+Dach)", "900",
     "Tobias Tumfart AT 8,16 EUR/m2 netto = 898 EUR fuer 110 m2. Modern, schimmelfest fuer Salzluft."),
    ("Solitex Diffusions-Membran", "700",
     "Pro Clima AT. ~120 m2 inkl. Ueberlappung. Verhindert Kondensat in Daemmwolle."),
    ("Innenliner Baumwolle (~115 m2 Stoff)", "1.380",
     "200 g/m2 Sergeant fuer Aesthetik + Daemmwolle-Halt. Beim selben Sattler/Sattlerei mitbeauftragen oder online beziehen."),
    ("Kuppel-Oberlicht oeffenbar (Polycarbonat)", "1.500",
     "KORRIGIERT Mai 2026: Atilla+Adorjan bestaetigen Tono 170-180 cm fuer 8m (nicht 120 cm). Kuppel muss entsprechend 1,5-1,7 m gross sein. Polycarbonat in dieser Groesse: 1.200-1.800 EUR. Bei Frame-Lieferant gleich mitbestellen."),
    ("Tuer-Upgrade (Glas-Doppeltuer / franz. Doppeltuer)", "600",
     "Adorjan-Standard-Tuer kann zu einfach sein - Upgrade pruefen. Oder lokal bei PT-Schreiner massgefertigt."),

    ("", "", ""),
    ("PLATTFORM 8M", "", ""),
    ("Plattform-Konstruktion (~50 m2)", "4.500",
     "10 Schraubpfaehle Krinner KSF-M 800mm = 350 EUR. Kiefer KDI Tragbalken 50x200 ca. 25 Stueck = 700 EUR. Laerchen-Deck T+G 27 mm fuer 50 m2 = 1.900 EUR. Kork-Daemmung 30mm 50 m2 = 900 EUR. Geo-Vlies + Schrauben + Sundries = 650 EUR."),
    ("Vereinfachte Plattform (Budget-Variante)", "3.000",
     "Falls Budget eng: Kork weglassen, Standard-Druck-impr. Decking statt Laerche. Spart 1.500 EUR."),

    ("", "", ""),
    ("STURM + KLIMA", "", ""),
    ("Sturm-Kit (Erdanker + Ratschen 5T)", "500",
     "ALGARVE-OBLIGATORISCH. 10 Anker statt 8 fuer 8m. Verstaerkter Spannband + Ratschen-Set."),
    ("Schattennetz Algarve (~75 m2)", "350",
     "Lokale Agrar-Kooperative. 8m braucht mehr Flaeche als 6m."),

    ("", "", ""),
    ("BUFFER", "", ""),
    ("Reserve / Verschnitt / Schrauben", "800",
     "Bei 8m mehr Kleinteile, hoeheres Risiko von Anpassungen."),

    ("", "", ""),
    ("GESAMTKOSTEN DIY-MODIFY 8M (Budget-Plattform, eigenes Werkzeug)", "23.780",
     "AKTUALISIERT Mai 2026 nach Adorjan-Real-Quote + Atilla-Engineering-Feedback: Adorjan Frame 8.150 (statt 5.000 Schaetzung) + Tono+Kuppel 1.500 (statt 800 wegen 170-180 cm Tono) = +3.850 vs Original. 8,8k UEBER 15k Budget."),
    ("GESAMTKOSTEN DIY-MODIFY 8M (Premium-Plattform + neues Werkzeug)", "26.658",
     "AKTUALISIERT Mai 2026: mit Premium Plattform inkl. Kork 4.500 + Werkzeug-Pauschale 1.378. Realistisches Komplett-Setup wenn alles neu beschafft."),
    ("Yourtepoque 8m bleibt PREISSIEGER 8m", "19.170",
     "11.820 + 2.000 Versand + 4.500 Plattform + 850 Sturm/Schatten = 19.170 EUR. Keine UV-Garantie auf Plane aber 4.000-7.000 EUR guenstiger als DIY-Modify 8m."),
    ("BLEIBT die Kompromiss-Option Casa 7,3m", "20.700",
     "15.750 inkl. Aufbau + 4.000 Plattform DIY + 850 Sturm/Schatten + 100 Kamin = 20.700. 10-J. UV-Garantie. 8m² kleiner als 8m DIY."),
    ("NEU ERKENNTNIS Mai 2026", "Adorjan/Atilla Engineering ist solide aber teurer als gedacht",
     "DIY-Modify 8m verliert seinen Preisvorteil gegenueber Yourtepoque (rein finanziell). Hauptargumente fuer DIY-Modify bleiben: 1) sofortige Verfuegbarkeit, 2) volle Materialkontrolle (Sauleda Marine-grade vs Yourtepoque-Standard-Canvas), 3) Lokaler PT-Reparatur-Service via Algarve-Sattler."),

    ("", "", ""),
    ("VERGLEICH 8M-OPTIONEN ALL-IN ALGARVE", "", ""),
    ("Yourtepoque 8m (FR) + Plattform + Sturm/Schatten", "19.170",
     "Einfachste 8m-Komplett-Loesung. ABER keine UV-Garantie auf Plane, Versand 2k FR-PT."),
    ("DIY-Modify 8m (Adorjan + Algarve-Sattler) Budget-Variante", "17.652",
     "EMPFEHLUNG #1 FUER 8M. Schaerfster Preis bei premium Sauleda-Plane lokal. 350-450 h Eigenarbeit."),
    ("DIY-Modify 8m Premium-Variante", "22.030",
     "Mit Premium-Plattform inkl. Kork-Daemmung."),
    ("Atilla 8m + Eichen-Upgrade + Plattform DIY", "21.960",
     "Komplette Jurte mit Premium-Holz von Atilla, dann Plattform selbst. Versand 3.500 EUR ist der Kostentreiber."),
    ("YourTent CZ 8m + Versand + Plattform", "21.350",
     "5-J. Plane-Garantie ist Plus. Versand CZ-PT ca. 2k. Schaetzung."),
    ("jurte24 8m winterfest + Versand + Plattform", "26.250",
     "Teuerste 8m-Option. Kiefer-Rahmen (kritisch fuer Algarve), nur Werbe-Aussage statt schriftl. UV-Garantie."),

    ("", "", ""),
    ("ALTERNATIVE: SIZE-DOWN ZU 7,3M", "", ""),
    ("Casa 7,3m + Plattform DIY + Sturm/Schatten + Kamin", "20.700",
     "Casa hat keine 8m, aber 7,3m mit 10-J. UV-Garantie und Lieferung+Aufbau inkl. STATUS: Oktober 2026 Slot weg, fruehestens 2027 verfuegbar."),

    ("", "", ""),
    ("FAZIT 8M", "", ""),
    ("8m unter 18k machbar?", "JA - aber nur DIY-Modify Budget-Variante",
     "17.652 EUR sind realistic fuer DIY-Modify 8m mit Sauleda-Plane und einfacherer Plattform. Sebastian muss 2-3k ueber Original-Budget gehen."),
    ("8m unter 20k machbar?", "JA - mehrere Wege",
     "DIY-Modify Premium: 22k, DIY-Modify Budget: 17,6k, Yourtepoque: 19k. Bewegt sich im 17-22k-Korridor je nach Detailwahl."),
    ("Empfohlene Strategie", "Adorjan + Atilla parallel anfragen (Mai 2026 gesendet), Sattler-Antworten abwarten, dann finalkalkulieren",
     "Mit konkreten Adorjan-Frame-Preis und Sattler-Plane-Preis kann auf 500 EUR genau gerechnet werden. Aktuell noch Schaetzungen."),
    ("Worst Case", "8m geht nicht in Budget - dann Casa 7,3m als 2027-Slot oder Casa 6,1m sofort",
     "Falls Adorjan oder Sattler unverhaeltnismaessig teuer, ist 7,3m bei Casa der beste Kompromiss (Casa-Slot 2027 oder Paula Young gebraucht)."),
]

row = 2
for entry in kalk_8m:
    a, b, c = entry
    if a and not b and not c:
        cell = ws14.cell(row=row, column=1, value=a)
        cell.font = Font(bold=True, size=12, color="2E5C8A")
        ws14.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
        ws14.row_dimensions[row].height = 22
    elif not a and not b and not c:
        ws14.row_dimensions[row].height = 8
    else:
        c1 = ws14.cell(row=row, column=1, value=a)
        c1.alignment = Alignment(wrap_text=True, vertical="top")
        c1.border = BORDER
        c2 = ws14.cell(row=row, column=2, value=b)
        c2.alignment = Alignment(horizontal="right", vertical="top")
        c2.border = BORDER
        if isinstance(b, str) and b.replace(".", "").replace(",", "").isdigit():
            try:
                c2.value = int(b.replace(".", ""))
                c2.number_format = '#,##0 "EUR"'
            except ValueError:
                pass
        if isinstance(b, (int, float)):
            c2.number_format = '#,##0 "EUR"'
        if isinstance(a, str) and (a.startswith("GESAMT") or a.startswith("VERGLEICH") or a.startswith("FAZIT") or a.startswith("BUDGET")):
            c1.font = Font(bold=True)
            c2.font = Font(bold=True, size=12, color="2E5C8A")
        c3 = ws14.cell(row=row, column=3, value=c)
        c3.alignment = Alignment(wrap_text=True, vertical="top")
        c3.border = BORDER
        ws14.row_dimensions[row].height = max(20, min(50, 15 + len(str(c)) // 8))
    row += 1
ws14.freeze_panes = "A2"


# ---------- Sheet 15: Kommunikations-Log (Status aller laufenden Anfragen) ----------
ws15 = wb.create_sheet("15 Kommunikations-Log")
headers15 = ["Datum", "Anbieter", "Kanal", "Richtung", "Inhalt", "Status", "Naechste Aktion"]
log = [
    # === Frame-Anbieter ===
    ["~10.05.2026", "Atilla es a Fehernep (HU)", "E-Mail",
     "Sebastian -> Atilla", "Originale Anfrage 8m Komplett-Jurte",
     "✓ Beantwortet", "Folgemail Frame-only-Anfrage senden (Entwurf liegt vor)"],
    ["~18.05.2026", "Atilla es a Fehernep (HU)", "E-Mail",
     "Sebastian -> Atilla", "Frame-only Anfrage 8m (Khaana+Tono 120cm+60 Uni+Tuer)",
     "✓ Beantwortet", "Atilla akzeptiert Frame-only ABER lehnt 120cm Tono ab"],
    ["~19.05.2026", "Atilla es a Fehernep (HU)", "E-Mail", "Atilla -> Sebastian",
     "ENGINEERING-KORREKTUR: 120cm Tono fuer 8m statisch zu klein (6cm Loch-Abstand). Atilla nutzt 180cm fuer 8m und 160cm fuer 6m. 60 Uni waeren auch zu wenig (schwache Wand bei 2m Hoehe). Fragt: jede Khaana-Spitze = ein Uni? Wuerde nur mit seinen Massen Verantwortung uebernehmen.",
     "Info erhalten - technisch wertvoll", "Antwort: Atilla's Standard-Masse akzeptieren (180cm Tono, jede Khaana-Spitze=Uni). Quote neu erfragen + Larch/Skylight/Lieferzeit/Versand/Eigene-Spedition Fragen."],
    ["20.05.2026", "Atilla es a Fehernep (HU)", "E-Mail", "Sebastian -> Atilla",
     "Akzeptiert Atilla's Engineering-Standards (180cm Tono, jede Khaana-Spitze=Uni, 2m Wand). Fragt nach: Frame-only-Quote mit Atilla's Massen, Laerche statt Buche/Kiefer, Skylight-Preis 180cm Tono, Lieferzeit, Versand frame-only Algarve, Eigene-Spedition-Option, Werkstatt-Besuch.",
     "✓ Gesendet", "Antwort abwarten (3-7 Tage). Bei Verzug nach 10 Tagen WhatsApp +36 30 657 0740."],
    ["~10.05.2026", "Adorjan Jurta (HU)", "E-Mail",
     "Sebastian -> Adorjan", "Frame-only Anfrage 8m, Larch optional, Oel-Finish, Versand PT",
     "✓ Teilantwort", "Folgemail: 4 Punkte offen (Larch / Tuer-Variante / Lieferzeit / Eigener Spediteur / Werkstatt-Besuch)"],
    ["?", "Adorjan Jurta (HU)", "E-Mail", "Adorjan -> Sebastian",
     "8m Rahmen 3.220.000 HUF = 8.150 EUR netto. Spezifikation Top: 48 Rafters 5x10cm, Tono 170cm 4-lagig verleimt, Wand 2m, Pinie AT Biopin-impraegniert, isolierte Tuer Standard. Versand offen.",
     "Info erhalten", "Folge-Fragen senden"],
    ["?", "Adorjan Jurta (HU)", "E-Mail", "Adorjan -> Sebastian",
     "Versand-Spediteur wird gesucht, kein Preis. 50% Anzahlung + 50% vor Versand. Installation extra (4 Pers + Fluege + Maschinen).",
     "Info erhalten", "Folge-Fragen (Larch/Tuer/Lieferzeit) + Eigener-Spediteur-Option erfragen"],
    ["~10.05.2026", "Yourtetoiles & Kontempobois (FR)", "E-Mail",
     "Sebastian -> Yourtetoiles", "Anfrage 8m Yurte",
     "✓ Teilantwort", "WhatsApp-Videocall Termin festlegen, beide Linien (Klassisch + Kontempo) erfragen"],
    ["?", "Yourtetoiles & Kontempobois (FR)", "E-Mail",
     "Yourtetoiles -> Sebastian",
     "Empfehlen Kontempo (Holz-Rundhaus mit Zink-Dach) statt klassischer Jurte fuer Atlantik-Exposition. Devis komplett dauert noch. Bieten Visio-Call an.",
     "Info erhalten", "Visio-Call vereinbaren (3 Terminvorschlaege schicken)"],

    # === PT-Hersteller ===
    ["~12.05.2026", "Casa dos Sonhos (Sarah & Vladimir, PT)", "Facebook Messenger",
     "Sebastian -> Sarah", "Anfrage 5/6,1/7,3/9,1m Preise + Konditionen",
     "✓ Beantwortet", "-"],
    ["~13.05.2026", "Casa dos Sonhos (Sarah & Vladimir, PT)", "Facebook Messenger",
     "Sarah -> Sebastian",
     "Preise bestaetigt: 5m 7.900 / 6,1m 9.700 / 7,3m 15.750 / 9,1m 18.000 EUR delivered+erected PT. Refs Brett (Quinta Glamping) + Paula Young (verkauft 7,3m gebraucht).",
     "Quote erhalten", "Nachfrage NIPC/IVA/Bankdaten/Material-Spec/Vermittlung Paula"],
    ["15.05.2026", "Casa dos Sonhos (Sarah & Vladimir, PT)", "PDF Quote",
     "Sarah -> Sebastian",
     "Quote 6,1m + 2.Tuer 850 + Flue-Kit 100 = 10.650 EUR. 1.000 EUR Anzahlung + 9.650 EUR am 1.8.2026. Lieferung 2. Oktoberwoche 2026.",
     "Quote erhalten", "Quote schwach dokumentiert (kein IVA-Split, kein Quote-Nr, keine Bank), Klaerung gefordert"],
    ["15.05.2026", "Casa dos Sonhos (Sarah, PT)", "Facebook Messenger",
     "Sarah -> Sebastian",
     "SLOT OKTOBER 2026 ist an anderen Kunden vergeben. Naechster Slot 2027.",
     "✗ Slot verloren", "Antwort raus: 2027er Slot + Preis-Bestaetigung + Paula-Vermittlung erfragen"],

    ["~10.05.2026", "Bēt Yurts (Noga & Nadav, PT)", "E-Mail",
     "Sebastian -> Noga", "Anfrage 6/7/9m Preise + Plattform + Aufbau Algarve",
     "✓ Beantwortet", "-"],
    ["~12.05.2026", "Bēt Yurts (Noga, PT)", "E-Mail + 3 PDF-Quotes",
     "Noga -> Sebastian",
     "Preise: 6m=12.600 / 7m=14.400 / 9m=19.800 Jurte netto. Plattform separat: 6m=6.500 / 7m=7.200 / 9m=11.500. Transport+Install: 700/1.900, 900/2.500, 1.500/2.900. 23% IVA on top. 50% upfront NICHT-RUECKZAHLBAR. AUGUST 2026 Slot. Material laut Quote: Nordic Pine + Outer Canvas + 80mm Insulation (ohne Marken-Angabe). 2-J. Garantie auf Plane.",
     "Quote erhalten", "Mail raus: Konkurrenz-Preis + Material-Aufschluesselung (Sauleda? PET?) anfordern"],

    # === Algarve-Sattler (DIY-Modify) ===
    ["~17.05.2026", "Dune Algarve Sailmakers (PT)", "Web/E-Mail",
     "Sebastian -> Dune", "Anfrage Sauleda-Plane 8m Massanfertigung",
     "⏳ Antwort offen", "5 Tage warten, dann WhatsApp/Tel.-Nachfrage. Telefon-Skript bereit."],
    ["~17.05.2026", "Toldos Etapaveloz (Loule, PT)", "Web/E-Mail",
     "Sebastian -> Etapaveloz", "Anfrage Sauleda-Plane 8m Massanfertigung",
     "⏳ Antwort offen", "5 Tage warten, dann Folgekontakt"],
    ["~17.05.2026", "Textilux (Algarve, PT)", "Web/E-Mail",
     "Sebastian -> Textilux", "Anfrage Sauleda-Plane 8m Massanfertigung",
     "✗ ABGELEHNT", "'nao sao da nossa especialidade' - klassischer Markisen-Hersteller, kein 3D-Persenning. Keine weitere Aktion."],

    # === Paula Young Lead ===
    ["TBD", "Paula Young (FB 'Paula Vegan Chef')", "Facebook (offen)",
     "Sebastian -> Paula", "Anfrage zur gebrauchten 7,3m Casa-Jurte",
     "⏳ Kontakt steht aus", "Ueber FB suchen ODER Sarah um Vermittlung bitten (in der naechsten Casa-Antwort mit aufnehmen)"],

    # === Neue OLX-Funde Mai 2026 ===
    ["21.05.2026", "OLX 8m Listing IDJatvu (Benedikt 'Yury')", "OLX Direktansicht",
     "Sebastian -> Sebastian (Recherche)",
     "Yurt 8m / 50m2 / 15.000 EUR / Gouveia / Verkaeufer Benedikt / Modell 'Yury' / NEU / inkl. Tuer+2 Fenster+weisser Innenstoff+Filzisolierung+'Gruendach'",
     "✓ Details bestaetigt",
     "BENEDIKT KONTAKTIEREN: 11 konkrete Fragen (Hersteller? Plane-Material/UV? Filz-Typ? was ist 'Gruendach'? Holz? Tono-Durchmesser? Transport Aljezur? Aufbau? Garantie? Verkaufsgrund? Besichtigungstermin?). KRITISCH: vor Anzahlung physische Besichtigung Gouveia."],
    ["21.05.2026", "OLX 8m Listing IDJbO7n ('luxuosa nova')", "OLX Direktansicht",
     "Sebastian -> Sebastian (Recherche)",
     "Yurt luxuosa NOVA 8m - vermutlich gleicher Verkaeufer (Benedikt?) - Sebastian muss Details teilen",
     "⏳ Details fehlen",
     "Wenn 2 Listings vom selben Verkaeufer = Benedikt ist Reseller mit Lager. Wichtig: 'luxuosa' Differenz zu 'Yury' verstehen (Material? Groesse? Preis?)"],
    ["TBD", "Green World (Ukraine)", "Web/E-Mail (zu pruefen)",
     "Sebastian -> Green World",
     "Anfrage 8m / 50m2 Yurte fuer Algarve, 10-J. Strukturgarantie, Versandkosten PT, Plane-Spezifikation",
     "⏳ Anfrage moeglich",
     "WENN OLX-Listings nicht von ihnen: trotzdem direkt anfragen ueber greenworld.house/pt/yurts/ - moeglicher 4. PT-Lieferant neben Casa/Bēt/Yonatan."],

    # === Telegram-Gruppe Jurten Portugal Mai 2026 ===
    ["~22.05.2026", "Telegram-Gruppe 'Jurten Portugal'", "Telegram-Post (DE)",
     "Sebastian -> Gruppe", "Eigene Wanted-Anfrage 6-8m Jurte fuer dauerhaftes Wohnen West-Algarve, 15k Budget",
     "✓ Gepostet", "Auf Antworten warten - bisher 1 konkrete Antwort von 'Dani'"],
    ["23.05.2026 07:00", "Dani (Telegram, Community-Projekt PT?)", "Telegram DM",
     "Dani -> Sebastian",
     "Verkauft 7m Custom-Build noch im Bau. Eigenes Design, seit 8J in 8m gewohnt. Hatte fuer daenische Familie geplant - jetzt verfuegbar. 200cm Wand, 150cm Skylight, Schafwoll-Filz, 2 Doppeltueren. Fertig August/September 2026.",
     "✓ Antwort erhalten", "Sebastian hat zurueckgeschrieben mit Standort+Specs+Preisfrage - warten auf Folge-Antwort"],
    ["23.05.2026 07:07", "Dani (Telegram)", "Telegram DM",
     "Sebastian -> Dani",
     "Standort? Algarve? Selbst gebaut? Materialspec (Fenster, Daemmungsdicke)? Preisvorstellung?",
     "✓ Gesendet", "Auf Antwort warten - dann Material-/Preis-Bewertung"],
    ["23.05.2026", "Dani (Castelo de Vide)", "Telegram DM", "Dani -> Sebastian",
     "FULL SPECS+PREIS: 16k inkl. Aufbau. Standort Castelo de Vide. Khaana Pinie 200cm Wand, Tono Douglasie 150cm, 2 Hartholz-Doppeltueren 155x200, Mittelhoehe 360cm, 28 Grad Dachneigung, 4cm Schafwolle, BW-Innenliner Bio, Canvas Aussen BW/Poly 50/50. Plattform/Transport extra. UV-Schutz Empfehlung drueber legen. PROBESCHLAFEN IM 6M-MODELL ANGEBOTEN!",
     "✓ Volle Spec erhalten",
     "Antwort raus: Probeschlafen-Termin (Juni) + Plane-Upgrade-Anfrage (Sauleda statt Canvas) + Dickere Daemmung erfragen"],
    ["23.05.2026 08:20", "Dani (Castelo de Vide)", "Telegram DM", "Sebastian -> Dani",
     "Unverbindliche Anfrage: alternative Aussenhuelle moeglich? Falls nein - Lebensdauer/Erneuerung/Kosten Canvas-Tausch nach 5J? 8cm Schafwolle moeglich+Aufpreis? Probeschlafen mit Frau+2 Kindern - Wann passt es bei euch?",
     "✓ Gesendet", "Antwort erhalten 12:47"],
    ["23.05.2026 12:47", "Dani (Castelo de Vide)", "Telegram DM", "Dani -> Sebastian",
     "EXTREM HILFREICHE ANTWORT: 1) Plane: naeht gerne anderen Stoff bis MASCHINEN-LIMIT 420g/m². Standard ueber ESVO NL = 1.400 EUR fuer 7m. UV-Schutz drueber laesst Stoff laenger halten. 2) Daemmung: 8cm prinzipiell ok, Sebastian muss selbst Material besorgen (aktuell kein AT-Bezug). 3) Sommer-Tipps: Fallschirm/Schattennetz drueber, Kuppel-Ueberzieher gegen direkte Einstrahlung. 4) Probeschlafen: 6m-Jurte besetzt bis Mitte Juni, danach evtl. Luecke. Besuch jederzeit ok. 4 Familien+Kinder vor Ort.",
     "✓ Vollantwort - sehr flexibel", "Sebastian antwortet: Sauleda/Tempotest erwaehnen (passt unter 420g/m²), 8cm Wolle selber besorgen (Isolena/Tumfart), Besuch in 2./3. Juniwoche vorschlagen"],
    ["23.05.2026", "Dani (Castelo de Vide)", "Telegram DM", "Sebastian -> Dani",
     "Lockere Antwort in Sebastian-Stil: Sauleda/Tempotest checken (290-340g/m²), 8cm Schafwolle selber besorgen, Kuppel-Ueberzieher-Tipp aufgenommen, Besuch 2./3. Juniwoche vorschlagen (Wohnwagen/Hotel falls noetig), bei Glueck spaeter Uebernachtung in 6m.",
     "⏳ Zu senden", "Sebastian's Antwort raus - Termin warten"],
    ["23.05.2026 12:53", "Dani (Castelo de Vide)", "Telegram DM", "Sebastian -> Dani",
     "Findet alles super, Frau kommt naechste Woche, meldet sich dann fuer Termin. Daemmung kann auch in PT bezogen werden (Recherche). Aug/Sep-Lieferzeit passt. Fragt: Bezahlung? Aufpreis zusaetzliches Fenster?",
     "✓ Gesendet", "Antwort erhalten 14:06"],
    ["23.05.2026 14:06", "Dani (Castelo de Vide)", "Telegram DM", "Dani -> Sebastian",
     "WICHTIGE EINSCHRAENKUNG: ZUSAETZLICHES FENSTER NICHT MOEGLICH - Waende sind schon fertig. Umbau bis September nicht machbar. Schafwolle-Bezug schwierig - neuer Anbieter 'Chris' hat in Schafwolldaemmungs-Post in Gruppe geantwortet. Dani bietet Chris-Kontakt-Vermittlung an.",
     "✓ Eingeschraenkt aber transparent", "Chris-Vermittlung annehmen!"],
    ["23.05.2026 14:09", "Dani (Castelo de Vide)", "Telegram DM", "Sebastian -> Dani",
     "Fragt nach Bezahlungsmoeglichkeiten und Firmenstruktur",
     "✓ Gesendet", "Antwort 14:11"],
    ["23.05.2026 14:11", "Dani (Castelo de Vide)", "Telegram DM", "Dani -> Sebastian",
     "ZAHLUNG: Anzahlung als Commitment, danach Teilbetraege verhandelbar. UEBERWEISUNG ODER PAYPAL (aufteilbar!) - PayPal-Kaeuferschutz potentiell nutzbar. AKTUELL: Privatverkauf an Freunde (kein Gewerbe, keine MwSt., keine Verbraucher-Rechte). Verein in Gruendung - dann ueber Verein.",
     "✓ Flexibel + PayPal-Option", "PayPal als Risiko-Absicherung wichtig"],
    ["23.05.2026 15:29", "Dani (Castelo de Vide)", "Telegram DM", "Sebastian -> Dani",
     "'Super 👍🏼' - kurzes Acknowledgement",
     "✓ Gesendet", "FOLGE: Chris-Vermittlung explizit annehmen + Anfang-Juni Besuchstermin avisieren"],
    ["24.05.2026 (Sebastians Folge)", "Dani (Castelo de Vide)", "Telegram DM", "Sebastian -> Dani",
     "Chris-Vermittlung annehmen + Anfang Juni Besuchstermin avisieren (mit Frau)",
     "⏳ Zu senden", "Wartet auf Frauen-Rueckkehr fuer Termin-Fixierung"],
    ["TBD", "Chris (PT-Schafwoll-Anbieter via Dani)", "Telegram (geplant)",
     "Sebastian -> Chris",
     "Schafwolle 8cm fuer 7m Jurte, Liefermenge ca. 50 m², Algarve-Klima beraten, Aufpreis vs 4cm",
     "⏳ Vermittlung steht aus", "Wenn Dani vermittelt: Chris direkt anschreiben"],

    # === Telegram-Gruppe Mai 2026: weitere Kontakte ===
    ["23.05.2026", "Mikael (Telegram-Gruppe)", "Telegram-Post",
     "Mikael -> Gruppe",
     "Bietet NEUE komplette 8m Jurte an, 2.10m Wandhoehe, sofort lieferbar, auch Massanfertigung anderer moeglich",
     "✓ Angebot erhalten",
     "Sebastian -> Mikael DM: Standort? Preis? Cover-Material UV? Fotos? Hersteller oder privat?"],
    ["23.05.2026", "Mikael (Telegram)", "Telegram DM",
     "Sebastian -> Mikael",
     "Antwort kurz: 8m mit 2.10m Wand passt zu Plaenen Algarve. Frage: Standort? Preis? Cover-Material (UV-resistent?)? Fotos? Hersteller oder privat?",
     "✓ Gesendet", "Antwort von Mikael erhalten"],
    ["23.05.2026", "Mikael (Mickaël) Gouveia/Guarda", "Telegram DM", "Mikael -> Sebastian",
     "VOLLSPEC + 3 FOTOS: 8m Yurt, 15.000 EUR + Versand. Standort GOUVEIA/GUARDA (gleicher Ort wie Benedikt 'Yury'!). Material: Canvas Standard mit EXTRA VINYL-OVERLAY-OPTION (Vinyl-Dach hat 2 zusaetzliche Oeffnungen, muesste 2 extra Fenster bauen). KONSTRUKTION: HOLZWAND-Style (vertikale Bretter, NICHT traditionelle Khaana - sieht nach Pacific-Yurts-Style aus). AC4-LAMINAT-FLOORING INKLUSIVE. Empfiehlt Sommer-Schattennetz.",
     "✓ Erste Daten erhalten",
     "VERDACHT: Mikael+Benedikt selber Verkaeufer oder verbunden (Gouveia+15k+8m). Sebastian fragt direkt nach. Folge-Fragen: Hersteller? Holz? Was in 15k? Versand-Schaetzung? Aufgebaut/neu? Mehr Fotos? Besuch Anfang Juni?"],
    ["24.05.2026 (Sebastians Folge)", "Mikael (Mickaël)", "Telegram DM",
     "Sebastian -> Mikael",
     "Detaillierte Folge-Fragen in 4 Bloecken: 1) HOLZ: Wand/Tono/Uni/Tueren-Holz+Behandlung. 2) ISOLIERUNG: Material/Dicke Wand/Dach/Innenliner-Spec. 3) WAS IN 15k: Anzahl Tueren+Fenster, Skylight-Groesse, Canvas vs Vinyl, AC4-Flaeche/Verlegung. 4) PRAKTISCH: Hersteller+Garantie, Versand-Schaetzung, neu/aufgebaut. Plus: mehr Innenfotos, Besuchstermin Gouveia Anfang Juni. VERDACHT BENEDIKT-VERBINDUNG direkt angesprochen.",
     "✓ Gesendet", "Antwort 13:15 erhalten"],
    ["24.05.2026 13:15", "Mikael (Mickaël) Gouveia", "Telegram DM", "Mikael -> Sebastian",
     "ANTWORT: Hersteller PACIFIC YURTS USA, nie aufgebaut, customized. Holz: solid wood vermutlich Nordische Kiefer mit klarem Lack OHNE Insektizid (selbst behandeln noetig!). Isolierung: Schafwoll-Filz 1.1 kg/m² 'Dicke 0.7mm' (unklar). Skylight 1.40m Kiefer+Eisen. AC4 Boden inkl., Sebastian installiert selbst. Konfiguration: 1 Tuer + 1 Fenster pro Seite. Kennt Benedikt ANGEBLICH NICHT, bittet um OLX-Link.",
     "✓ Vollantwort erhalten", "WICHTIG: Pacific Yurts = Premium-Hersteller. ABER: Kein Termiten-Schutz, Filz-Dicke unklar, nur 1 Tuer. Sebastian Folge-Fragen + OLX-Link senden."],
    ["24.05.2026 (Sebastians Folge)", "Mikael (Mickaël)", "Telegram DM",
     "Sebastian -> Mikael",
     "Folge-Klarstellungs-Fragen + Fotos analysiert: 1) OLX-Link Benedikt mitgeschickt zur Pruefung. 2) Filz-Dicke 0.7mm = was genau (Canvas/Filz/Schreibfehler 7mm)? 3) Termiten-Behandlung in Algarve obligatorisch, was empfiehlt er + Kosten? 4) Pacific Yurts Modell-Jahr + Serie? Original Sunforger 13oz Canvas oder anders? 5) Genaue Fenster-Anzahl + Groesse. 6) Versand-Schaetzung Aljezur. 7) Besuch Anfang Juni Gouveia (ggf. mit Benedikt zusammen). 8) PLATTFORM-INKLUSION bestaetigen (sichtbar in Fotos).",
     "⏳ Zu senden", "Antwort wird klaeren ob ernsthafte Option oder durchfallen. PLATTFORM-KLAERUNG ist kritisch (spart 2.880 EUR)"],
    ["24.05.2026", "Mikael (Mickaël) Gouveia", "Telegram DM", "Mikael -> Sebastian (4 FOTOS)",
     "FOTO-UPDATE: Foto 1 = Jurte im Aufbau mit Canvas-Dach + Doppeltuer + AUFGESTAENDERTE PLATTFORM auf Holzpfaehlen. Foto 2 = TRADITIONELLE KHAANA-SCHEREN-GITTER (nicht Vertikal-Holzwand!) mit 2 Glas-Fenstern + Doppeltuer. Foto 3 = Innenansicht mit weißem Vlies + Uni-Dachsparren radial. Foto 4 = Pacific Yurts offizielles 'Traditional Structure' Marketing-Bild.",
     "✓ Fotos erhalten", "WICHTIG: Construction IST traditionelle Khaana (wie Dani), nur maschinen-praezise von Pacific Yurts. Plattform sichtbar im Preis enthalten. Verarbeitung sehr ordentlich. RESET vorherige Annahme 'Vertikal-Holzwand' = FALSCH."],
    ["24.05.2026 13:42", "Mikael (Mickaël) = Bento", "Telegram DM", "Mikael -> Sebastian",
     "KLARSTELLUNGEN: 1) ER IST DER OLX-VERKAEUFER (Bento, nicht Benedikt - Translator-Fehler). EINE Person, nicht zwei! 2) Filz BESTAETIGT 7mm (traditional thickness). 3) Termiten-Behandlung obligatorisch. 4) Modell 'European Modern' (Pacific Yurts Variante). 5) NEUES PROBLEM: SCHRAUBEN ROSTIG an Khaana-Waenden - muessen ersetzt werden, 'hence the price'. 6) Besuch jederzeit willkommen.",
     "✓ Klargestellt + Mangelinfo",
     "WICHTIG: Mikael war komplett ehrlich ueber Rost-Mangel = Vertrauensbeweis. PREIS 15K IST MAENGELPREIS. Rost-Ausmass muss vor Ort gepruefen werden. ~30 Std Eigenarbeit Schrauben-Tausch + 200 EUR Material Edelstahl A4."],
    ["24.05.2026 13:59", "Mikael (Bento)", "Telegram DM", "Mikael -> Sebastian",
     "Begeistert ueber geplanten Besuch ('Are you coming to Gouveia? That's great, you can see the yurt and its components 👍')",
     "✓ Termin-Bereitschaft", "Konkreten Datums-Vorschlag + Vorab-Fragen senden"],
    ["24.05.2026 (Sebastians Folge)", "Mikael (Bento)", "Telegram DM",
     "Sebastian -> Mikael",
     "Besuchsbestaetigung + 2 Datum-Optionen (Sa 13. oder Sa 20. Juni mit Frau+Kindern) + Vorab-Fragen: Rost-Ausmass (Verfaerbung vs durchgerostet?), welche Schrauben betroffen (nur Khaana oder auch Tono/Uni?), Original-Material (verzinkt?), Pacific Yurts Modell-Name+Jahr, Plattform-Inklusion bestaetigen. Hinweis: bringen Lunch mit, schlafen in Pension.",
     "⏳ Zu senden", "Bei Termin-Fixierung: gleiches Wochenende auch Dani Castelo de Vide besuchen (~130 km zwischen den Orten)"],
    ["24.05.2026 spaeter", "Mikael (Bento)", "Telegram DM", "Mikael -> Sebastian",
     "KORREKTUR: AC4-Boden 52 m² NICHT in 15k - extra 700 EUR (Selbstkostenpreis). Plattform NICHT in 15k - Sebastian baut selbst aus Saegewerk-Holz lokal (guenstiger). KLARSTELLUNG DAEMMUNG: '7mm refers to the thickness of the insulation, density 1,100 kg/m²' (vermutlich PT-Komma: 1,1 kg/m² = 157 kg/m³ Dichte = normal aber 7mm KRITISCH DUENN).",
     "✓ Wichtige Klarstellungen", "KRITISCH: 7mm Daemmung ist halb so viel wie Industrie-Standard fuer ganzjaehrigen Wohnen. Sebastian muss zusaetzlich daemmen ODER aktiv heizen."],
    ["24.05.2026 (Sebastians Folge)", "Mikael (Bento)", "Telegram DM",
     "Sebastian -> Mikael",
     "Bestaetigungs-Frage Daemmungs-Wert (1,1 vs 1.100 - Komma-Dezimal-Klaerung) + Erfahrungs-Frage: andere Pacific-Yurts-Besitzer in PT haben zusaetzlich gedaemmt? OK fuer Sebastian eine 8mm Zusatz-Schicht draufzulegen?",
     "⏳ Zu senden", "Bei Antwort: bei Besuch Daemmungs-Realitaet anfassen"],
    ["24.05.2026 spaeter", "Mikael (Bento) Gouveia", "Telegram DM", "Mikael -> Sebastian",
     "TRAGISCHE HINTERGRUNDGESCHICHTE: Pacific Yurts gekauft AUGUST 2025 (10 Mon. her), seitdem in einem Van gelagert. GROSSER BRAND zerstoerte ihre Grundstuecks-Plaene - haben deshalb nichts installiert. Erklaert: Verkaufsgrund (echtes Schicksal), Rost-Schrauben (Van-Lagerung mit Temperatur+Feuchte-Wechsel), nie aufgebaut. Empfiehlt Holz-Behandlung vor Installation.",
     "✓ Vollkontext geklaert",
     "MENSCHLICH stark - tragische Verkaufsmotivation. STRATEGISCH neues Risiko: 10 Mon. Van-Lagerung koennte Filz/Plane/Holz beeinflusst haben (Schimmel/Verzug)."],
    ["24.05.2026 (Sebastians empathische Folge)", "Mikael (Bento)", "Telegram DM",
     "Sebastian -> Mikael",
     "Empathische Antwort zum Brand. Bestaetigung Verstaendnis (15k + 700 AC4 + DIY Plattform + DIY Daemmungs-Nachruest+Termiten-Behandlung). Termin-Vorschlag Sa 14. oder Sa 21. Juni mit Familie+Pension. Wichtige Vorab-Fragen: 1) Lagerungs-Zustand Canvas/Filz/Holz nach 10 Mon. Van (Feuchte/Schimmel/Verzug?). 2) Rost-Ausmass: 10/50/100% der Schrauben?",
     "⏳ Zu senden",
     "Vor dem Besuch: Lagerungs-Zustand und Rost-Quantitaet abklaeren, um Trip-Aufwand zu rechtfertigen"],
    ["24.05.2026 spaeter", "Mikael (Bento)", "Telegram DM (PT)", "Mikael -> Sebastian",
     "VERSAND-REFERENZ: Vorheriger Kunde aus Sebastians Naehe hat Jurte selbst mit gemietetem grossem Van abgeholt - Kosten 750 EUR. Real-Referenz Selbst-Abholung billiger als Spedition.",
     "✓ Konkrete Versand-Option",
     "Sebastian Antwort: Frage nach Groesse der anderen Jurte (7m/8m?) und ob Mikael beim Beladen hilft"],
    ["24.05.2026 (Sebastians Folge)", "Mikael (Bento)", "Telegram DM",
     "Sebastian -> Mikael",
     "Versand-Optionen: 750 EUR Selbst-Abholung wird in Sebastians Plan integriert. Klaerung-Fragen: Welche Groesse hat die andere Jurte gekostet? Hilfst du beim Beladen? Plus Idee: Besuch und Abholung kombinieren in 1-2 Wochenenden",
     "⏳ Zu senden", "Bei Bestaetigung: Van-Miete in Aljezur recherchieren (typisch Renault Master 13-17 m³ ~80-150 EUR/Tag)"],
    ["24.05.2026 spaeter", "Mikael (Bento)", "Telegram DM", "Mikael -> Sebastian",
     "NEUES ANGEBOT: Mikael hat sowieso Lieferung nach Castelo Branco geplant - bietet an, Sebastians Jurte direkt nach Aljezur mitzunehmen mit Preisanpassung. Gouveia -> Castelo Branco ~110 km, dann Zusatz Castelo Branco -> Aljezur ~400 km.",
     "✓ Liefer-Option erhalten",
     "Sebastian Fragen: 1) Wann ist Castelo-Branco-Lieferung geplant (vor oder nach Juni-Besuch)? 2) Konkreter Aufpreis fuer Lieferung Aljezur? 3) Hilft Mikael beim Abladen?"],
    ["24.05.2026 (Sebastians Folge)", "Mikael (Bento)", "Telegram DM",
     "Sebastian -> Mikael",
     "Ja zur Liefer-Option grundsaetzlich interessiert. Wann ist Castelo-Branco-Termin? Wie viel Aufpreis fuer Aljezur? Wenn nach Mitte Juni: ideal - Besuch + Lieferung kombinieren.",
     "⏳ Zu senden",
     "Wenn Mikaels Liefertermin VOR Mitte Juni: Konflikt mit Besuch. Falls NACH Mitte Juni: ideal."],
    ["24.05.2026 KONSOLIDIERT", "Mikael (Bento)", "Telegram DM",
     "Sebastian -> Mikael",
     "KONSOLIDIERTE NACHRICHT MIT 8 BLOECKEN: 1) Lagerungs-Zustand 10 Mon Van (Canvas/Filz/Holz). 2) Rost-Quantifizierung (%/Position/Material). 3) Pacific Yurts Modell+Sunforger-Canvas+Vinyl-Inkl. 4) DAEMMUNG-UPGRADE moeglich (14mm oder doppelt)? Falls nein: ok wenn Sebastian 8mm Zusatzschicht selbst macht? 5) HOLZ-BEHANDLUNG durch Mikael vor Lieferung moeglich+Preis? Falls nein: Produkt-Empfehlung. 6) LIEFERUNG Castelo Branco Datum+Aufpreis Aljezur+Hilfe Abladen. 7) Visit-Termin Sa 14 vs Sa 21 Juni. 8) Andere Jurte 7m/8m, Mikael Beladen-Hilfe?",
     "✓ Gesendet", "Vollantwort 15:02 erhalten"],
    ["24.05.2026 15:02", "Mikael (Bento) Gouveia", "Telegram DM", "Mikael -> Sebastian",
     "VOLLANTWORT auf 8-Block-Nachricht: 1) LAGERUNG: 'Nothing rusty, moldy, mildew anymore, smells brand new' - TOP Zustand. 2) ROST: 30% Schrauben oberflaechlich, Ursache: REGULAR STEEL statt verzinkt (Fertigungsfehler!) - implizit ALLE Schrauben gefaehrdet. 3) Pacific Yurts komplett bestaetigt. NORMALPREIS 8m Panoramic = 18.500 EUR (15k ist 3.500 EUR Rabatt fuer Schrauben). 4) Dickere Daemmung kann Mikael beschaffen (Sheep wool, 'treated properly'). 5) Holz-Behandlung: Initial bereits erfolgt + Spray am Abholtag (kein Aufpreis erwaehnt). 6) LIEFERUNG CASTELO BRANCO AM 1. JUNI! Vor Sebastians Besuch (14./21. Juni). 7) Visit-Termin offen. 8) Vinyl-Overlay 8m = EXTRA 1.000 EUR (nicht in 15k). AC4 = 700 EUR. Mikael hilft beim Beladen mit Maschine, Lieferung kommt mit Helfer.",
     "✓ Vollantwort", "TIMING-PROBLEM: 1. Juni Lieferung vor Juni-Besuch = unmoeglich kombinierbar. Sebastian muss Selbst-Abholung wahlen oder spaeteren Termin verhandeln."],
    ["24.05.2026 (Sebastians Folge)", "Mikael (Bento)", "Telegram DM",
     "Sebastian -> Mikael",
     "Bedauert dass 1. Juni zu frueh - Frau kommt erst naechste Woche zurueck. Zwei Optionen: A) Selbst-Abholung Mietvan ~750 EUR Ende Juni/Anfang Juli nach Besuch. B) Mikael macht spaetere Liefertour im Mitte/Ende Juni nach Algarve mit Aufpreis. Plus: ALLE Schrauben durch Mikael ersetzen lassen (nicht nur 30%) + Vinyl-Overlay-Preis 1k fuer welche Variante genau?",
     "⏳ Zu senden", "Bei Mikael's Antwort: Logistik-Entscheidung + Final-Kalkulation Vinyl-Overlay"],

    ["23.05.2026", "Dror (Telegram-Gruppe)", "Telegram-Post",
     "Dror -> Gruppe",
     "Stellt sich vor: viel Erfahrung im Jurten-BAU UND REPARATUR, baut selbst (noch) keine. Bietet sich als Kontakt fuer Jurten-Fragen an.",
     "✓ Wertvoller Kontakt", "Sebastian -> Dror DM: warm danken, Projekt kurz erklaeren, signalisieren dass spaeter Sanity-Check kommen wird. Frage: Standort PT?"],
    ["23.05.2026", "Dror (Telegram)", "Telegram DM",
     "Sebastian -> Dror",
     "Warmer Erstkontakt: Vorstellung, Projekt-Beschreibung, signalisieren dass spaeter Sanity-Check kommt. Frage Standort PT.",
     "⏳ Zu senden",
     "DROR ALS WERTVOLLE RESSOURCE: Sanity-Check auf Dani/Benedikt/Mikael Angebote, Material-Beurteilung, Plattform-Engineering, Reparatur-Wissen fuer Zukunft. Konkrete Anfragen erst bei Auslosern (vor Anzahlung, vor Plattform-Bau, etc.)."],
    ["23.05.2026", "Sebastian -> Selbst-Klaerung", "Analyse", "Intern",
     "Frage 4cm Schafwolle Daemmung fuer Algarve-Winter ausreichend? Antwort: KNAPP. Dani's 4cm (R~1.05) ist halb so gut wie Bēt 8cm PET (R~2.1) und 1/3 vom DE-Niedrig-Bau-Standard. Aljezur-Atlantikkueste hat mildere Temperaturen als Castelo de Vide, aber starker Wind+Feuchtigkeit verstaerken Waermeverlust. MIT aktivem Holzofen+Top-Plattform-Daemmung (60-80mm Kork) = wohnen ja, mit Komfort-Limits. OHNE Heizen = nachts 8-12°C. DRINGEND: 8cm Upgrade bei Dani anfragen.",
     "Analyse fertig", "Bei Dani 8cm Upgrade anfragen ODER Plattform-Daemmung als Kompensation auslegen (60-80mm Kork)"],

    # === Brett bei Quinta Glamping ===
    ["TBD", "Brett, Quinta Glamping (Algarve)", "Online-Buchung", "Sebastian (geplant)",
     "Uebernachtung im Lake View Yurt (Casa dos Sonhos)",
     "⏳ noch nicht gebucht", "Booking direkt: https://www.quintaglamping.com/rooms/lake-view-yurt.html - vor Anzahlung an Casa wertvoll"],
]

for col, h in enumerate(headers15, 1):
    c = ws15.cell(row=1, column=col, value=h)
    c.fill = HEADER_FILL
    c.font = HEADER_FONT
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = BORDER
for r_idx, row in enumerate(log, 2):
    for c_idx, val in enumerate(row, 1):
        cell = ws15.cell(row=r_idx, column=c_idx, value=val)
        cell.alignment = Alignment(vertical="top", wrap_text=True)
        cell.border = BORDER
        if r_idx % 2 == 0:
            cell.fill = ALT_FILL
        # Status-Spalte einfaerben
        if c_idx == 6 and isinstance(val, str):
            if "ABGELEHNT" in val or "Slot verloren" in val:
                cell.fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
            elif "Antwort offen" in val or "Kontakt steht aus" in val or "noch nicht" in val:
                cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
            elif "Beantwortet" in val or "Quote erhalten" in val or "Info erhalten" in val:
                cell.fill = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")
widths15 = [14, 36, 22, 22, 60, 22, 50]
for i, w in enumerate(widths15, 1):
    ws15.column_dimensions[get_column_letter(i)].width = w
ws15.row_dimensions[1].height = 30
ws15.freeze_panes = "A2"


# ---------- Sheet 16: DIY Bauplan-Quellen (Deep Research Mai 2026) ----------
ws16 = wb.create_sheet("16 DIY Bauplan-Quellen")
headers16 = ["Kategorie", "Quelle", "Typ", "Sprache", "Link", "Bewertung 6-9m"]
quellen_bauplan = [
    # === PFLICHT-3 ===
    ["1 Pflicht-Quelle", "Paul King: 'The Complete Yurt Handbook'",
     "Buch (~25 EUR)", "Englisch",
     "https://www.amazon.co.uk/Complete-Yurt-Handbook-Paul-King/dp/1899233083",
     "Standardwerk: 121 Seiten, 3 Jurten-Typen in mehreren Groessen, Materiallisten, Worked Examples. UNVERZICHTBAR."],
    ["1 Pflicht-Quelle", "Paul King: 'Build Your Own Yurt' (Vorgaengerwerk)",
     "Free PDF", "Englisch",
     "https://azinelibrary.org/approved/build-your-own-yurt-1.pdf",
     "Kostenlos. 3m-Beispiel aber Logik skalierbar. Idealerweise VOR Buchkauf lesen."],
    ["1 Pflicht-Quelle", "SimplyDifferently.org Yurt Calculator",
     "Web-Tool kostenlos", "Englisch",
     "https://simplydifferently.org/Yurt_Notes",
     "GOLDSTANDARD. Du gibst Durchmesser ein -> exakte Latten-Anzahl, Bohrwinkel, Uni-Laengen, Tono-Bemassung. Pflicht."],
    ["1 Pflicht-Quelle", "SimplyDifferently Construction PDF",
     "Free PDF", "Englisch",
     "https://simplydifferently.org/DL/yurt_construction_document.pdf",
     "Vertieftes PDF mit Konstruktions-Detail."],

    # === PLATTFORM ===
    ["2 Plattform", "Pacific Yurts: 20 ft / 6,1 m Plattform-Plan",
     "Free PDF", "Englisch",
     "https://www.yurts.com/wp-content/uploads/2023/12/20ft-Pacific-Yurts-Platform-Plan.pdf",
     "Direkt fuer 6,1m verwendbar. Standard-USA-Bauweise."],
    ["2 Plattform", "Pacific Yurts: 24 ft / 7,3 m Plattform-Plan",
     "Free PDF", "Englisch",
     "https://www.yurts.com/wp-content/uploads/2023/12/24ft-Pacific-Yurts-Platform-Plan.pdf",
     "Direkt fuer 7,3m. Fuer 8m um ~10% skalieren."],
    ["2 Plattform", "Pacific Yurts Setup Manual 20/24/30 ft",
     "Free PDF", "Englisch",
     "https://www.yurts.com/wp-content/uploads/2023/12/Set-Up-Manual_20-24-30-Yurts.pdf",
     "Komplettes Aufbau-Manual."],
    ["2 Plattform", "Pacific Yurts Download-Hub",
     "Web-Hub", "Englisch", "https://www.yurts.com/downloads/",
     "Alle Pacific-Yurts-PDFs zentral."],
    ["2 Plattform", "That Yurt Blog: Foundations, Beams, Blocking",
     "Web-Artikel", "Englisch",
     "https://www.thatyurt.com/yurt-journal/yurt-platform-building-1/",
     "Detaillierte Foto-Doku Plattform-Aufbau."],
    ["2 Plattform", "Krinner Ground Screws",
     "Hersteller-Website", "Englisch/Deutsch",
     "https://www.krinner.io/en/",
     "Schraubpfaehle-Marktfuehrer. KSF-M 800mm Standard fuer 8m-Yurte."],
    ["2 Plattform", "Krinner Foundation Construction Guide",
     "Hersteller-Tutorial", "Englisch",
     "https://www.krinner.io/en/foundation-construction/ground-screws/",
     "Schritt-fuer-Schritt Schraubpfahl-Einbau."],
    ["2 Plattform", "Shelter Designs Platform Guide",
     "Web-Artikel", "Englisch",
     "https://www.shelterdesigns.net/learn-about-yurts/how-to-plan-your-yurt/how-to-build-a-yurt-platform/",
     "Alternative Plattform-Bauweise."],

    # === KHAANA (Lattengitter) ===
    ["3 Khaana", "The Delmer Yurt Blog: Khana",
     "Web-Foto-Doku", "Englisch",
     "http://thedelmeryurt.blogspot.com/p/khana.html",
     "Detaillierte Foto-Schritte Khaana-Konstruktion. Sehr nuetzlich."],
    ["3 Khaana", "The Yurt Is Born: Assembling Lattice Wall",
     "Web-Tutorial", "Englisch",
     "https://theyurtisborn.wordpress.com/2013/01/15/assembling-the-lattice-wall-section/",
     "Schritt-fuer-Schritt mit Tipps zur Bohr-Schablone."],
    ["3 Khaana", "ThePlywood.com - How to Make a Yurt",
     "Web-Tutorial", "Englisch", "https://theplywood.com/yurt/",
     "Materialliste + Bauanleitung."],
    ["3 Khaana", "Outdoor Happens Homestead - Build a Yurt",
     "Web-Artikel", "Englisch",
     "https://www.outdoorhappens.com/how-to-build-a-yurt/",
     "Materialliste + Kosten + DIY-Kit-Anbieter."],
    ["3 Khaana", "Instructables: Yurt Without Steel (10 Steps)",
     "Web-Tutorial", "Englisch", "https://www.instructables.com/Yurt/",
     "Holz-only-Variante ohne Stahl-Schrauben."],

    # === TONO (Kronenring) ===
    ["4 Tono", "The Delmer Yurt Blog: Crown",
     "Web-Foto-Doku", "Englisch",
     "http://thedelmeryurt.blogspot.com/p/crown.html",
     "Tono-Konstruktion: 4-lagig verleimt, Bohrwinkel-Schablone, Dichtung."],
    ["4 Tono", "Solaripedia: Yurt Roof Ring",
     "Web-Artikel", "Englisch",
     "https://www.solaripedia.com/13/318/3704/yurt_roof_ring.html",
     "Architektonische Detail-Beschreibung."],
    ["4 Tono", "Felting and Fiber Studio: Yurt Tono",
     "Blog-Artikel", "Englisch",
     "https://feltingandfiberstudio.com/2013/08/29/yurt-tono-and-a-felting-party/amp/",
     "Tono-Bau Schritt-fuer-Schritt mit Filz-Detail."],
    ["4 Tono", "Yurt Forum: DIY Center Ring",
     "Forum-Diskussion", "Englisch",
     "https://www.yurtforum.com/forums/building-a-yurt-f3/diy-yurt-center-ring-any-ideas-529.html",
     "Community-Erfahrungen Tono selbst bauen."],

    # === AUSSENPLANE / COVER ===
    ["5 Plane", "The Delmer Yurt Blog: Canvas Exterior",
     "Web-Foto-Doku", "Englisch",
     "http://thedelmeryurt.blogspot.com/p/canvas-exterior.html",
     "Schnittmuster + Naehtechnik. WICHTIG fuer Maß-Konsistenz."],
    ["5 Plane", "Outdoor Sewing Solutions UK: Designing Your Yurt",
     "Web-Tutorial", "Englisch",
     "https://outdoorsewingsolutions.co.uk/designing-your-yurt/",
     "Pattern-Designer-Tutorial mit Bemassungs-Logik."],
    ["5 Plane", "Yurting: Sewing Canvas with Sailrite",
     "Blog-Artikel", "Englisch",
     "https://diyode.com/blog/2014/01/yurting-sewing-the-canvas-with-the-sailrite-and-yurt-dwelling-tips/",
     "Industrie-Naehmaschinen-Tipps + Yurten-Leben."],
    ["5 Plane", "Beowulf Industrial Sewing: Yurt Covers",
     "Web-Anbieter", "Englisch", "https://www.beowulfsewing.com/yurt-covers",
     "Professionelle Cover-Naeher in USA. Konzepte uebertragbar."],
    ["5 Plane", "Yurtcovers.com Broker Advice",
     "Web-Artikel", "Englisch", "http://yurtcovers.com/yurtbroker-advice/",
     "Cover-Profi-Tipps."],

    # === VIDEO-TUTORIALS ===
    ["6 Video", "Kents of Cornwall: Full Tutorial 5m",
     "YouTube", "Englisch",
     "https://www.youtube.com/watch?v=tBRt4yEqwYA",
     "Klassisch 5m, Logik fuer 6-8m uebertragbar."],
    ["6 Video", "How to Build an ORON Yurt - Full Step by Step",
     "YouTube", "Englisch",
     "https://www.youtube.com/watch?v=8p_-3X1C7jA",
     "Moderne Bauweise, Komplett-Anleitung."],
    ["6 Video", "Building Yurt Platform & Raising Playlist (30ft)",
     "YouTube Playlist", "Englisch",
     "https://www.youtube.com/playlist?list=PLCu9QeIvVTyZjdRf3cY-y2CHOzrMahXHY",
     "30-ft (9m) komplette Doku - direkt anwendbar fuer Sebastians 8m."],
    ["6 Video", "Inside DIY Yurt - Registered Home in Germany",
     "YouTube", "Englisch",
     "https://www.youtube.com/watch?v=wNWlieH2ha0",
     "Wohnnutzung-Beispiel mit Bau-Genehmigung. Inspirations-Video."],
    ["6 Video", "Strongest Yurt - Full Build",
     "YouTube", "Englisch",
     "https://www.youtube.com/watch?v=zIcRn50oh_E",
     "Robuste Bauweise fuer extremes Klima."],
    ["6 Video", "Shelter Designs Video Library",
     "Web-Hub", "Englisch",
     "https://www.shelterdesigns.net/learn-about-yurts/yurt-building-videos/",
     "Mehrere Videos: Flooring, Framing, Wiring, Plumbing."],
    ["6 Video", "Couple Building YURT Timelapse",
     "YouTube", "Englisch",
     "https://www.youtube.com/watch?v=-6f_GIqqZ8Y",
     "Komplett-Aufbau visualisiert in 15min Timelapse."],

    # === KOMPAKTE GUIDES ===
    ["7 Step-by-Step Guide", "HomeBiogas: DIY Yurt Complete Guide",
     "Web-Guide", "Englisch", "https://www.homebiogas.com/blog/diy-yurt/",
     "Kompakte Komplett-Anleitung."],
    ["7 Step-by-Step Guide", "Outdoor Happens: Build a Yurt",
     "Web-Guide", "Englisch",
     "https://www.outdoorhappens.com/how-to-build-a-yurt/",
     "Materialliste, Kosten, DIY-Kit-Anbieter."],
    ["7 Step-by-Step Guide", "Woodworkers Institute: Traditional Yurt",
     "Web-Guide", "Englisch",
     "https://woodworkersinstitute.com/how-to-make-a-yurt/",
     "Handwerker-Perspektive."],
    ["7 Step-by-Step Guide", "Yurtnotes.com Illustrated Guide",
     "Web-Guide", "Englisch",
     "https://yurtnotes.com/yurt-construction-a-helpful-illustrated-guide/",
     "Illustrierte Schritte."],
    ["7 Step-by-Step Guide", "Milkwood Permaculture: Yurt Resources",
     "Web-Guide", "Englisch",
     "https://www.milkwood.net/2012/07/04/building-a-yurt-from-scratch-resources/",
     "Permakultur-Perspektive auf Bau."],

    # === COMMUNITIES & FOREN ===
    ["8 Community", "Yurt Forum",
     "Forum", "Englisch", "https://www.yurtforum.com/",
     "Hauptforum fuer DIY-Fragen, Bauprobleme."],
    ["8 Community", "Felting and Fiber Studio - Yurt Series",
     "Blog-Serie", "Englisch",
     "https://feltingandfiberstudio.com/category/yurt/",
     "51-Tage-Bau-Dokumentation."],
    ["8 Community", "YurtInfo.org - Plans",
     "Web-Hub", "Englisch", "https://www.yurtinfo.org/yurt-plans",
     "Linkliste zu Plaenen."],
    ["8 Community", "YurtInfo.org - Resources",
     "Web-Hub", "Englisch", "https://www.yurtinfo.org/yurt-resources",
     "Allgemeine Ressourcen."],
    ["8 Community", "YurtInfo.org - Bookstore",
     "Web-Hub", "Englisch", "https://www.yurtinfo.org/yurt-bookstore",
     "Buch-Empfehlungen."],

    # === ENGINEERING ===
    ["9 Engineering Wind", "Structural Basics: Wind Loads on Pitched Roofs",
     "Web-Artikel", "Englisch",
     "https://structuralbasics.substack.com/p/wind-loads-on-pitched-roofs",
     "Wind-Last-Berechnung Schraegdach. Algarve-Atlantik-Wind kritisch."],
    ["9 Engineering Wind", "Omnicalculator: Wind Load Calculator",
     "Online-Rechner", "Englisch",
     "https://www.omnicalculator.com/physics/wind-load",
     "Quick-Calc fuer Sturm-Auslegung."],

    # === ENGINEERING-ZUSAMMENFASSUNG fuer Sebastian 8m ===
    ["10 Sebastian-Specs (8m)", "Tono Durchmesser",
     "Spec", "DE",
     "180 cm (Atilla-Standard, mit jeder Khaana-Spitze=Uni) oder 170 cm (Adorjan-Standard, mit 48 massiven 5x10cm Rafters)",
     "EXTRA WICHTIG: 120 cm waere zu klein! Mehrfach bestaetigt durch Atilla + Adorjan."],
    ["10 Sebastian-Specs (8m)", "Uni Anzahl",
     "Spec", "DE",
     "48 (5x10cm massiv) ODER 64-80 (3,5x3,5cm klassisch)",
     "Wahl haengt von Tono-Bauart ab."],
    ["10 Sebastian-Specs (8m)", "Khaana-Latten",
     "Spec", "DE", "~380 Stueck, 30x10 mm, 2,1 m Laenge",
     "8 Sektionen. Edelkastanie ideal fuer Algarve."],
    ["10 Sebastian-Specs (8m)", "Wandhoehe",
     "Spec", "DE", "2,00 m (Casa-Standard)",
     "Hoeher als traditionell mongolisch (1,70m)."],
    ["10 Sebastian-Specs (8m)", "Bohrwinkel Tono fuer Uni",
     "Spec", "DE", "30 Grad",
     "Mit Bohrlehre arbeiten - Wiederholgenauigkeit kritisch."],
    ["10 Sebastian-Specs (8m)", "Sturm-Verankerung Algarve",
     "Spec", "DE", "8-10 Erdanker + 5-Tonnen Ratschen ueber Dach",
     "Atlantik-Wind bis 100 km/h - obligatorisch."],
]

for col, h in enumerate(headers16, 1):
    c = ws16.cell(row=1, column=col, value=h)
    c.fill = HEADER_FILL
    c.font = HEADER_FONT
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = BORDER
for r_idx, row in enumerate(quellen_bauplan, 2):
    for c_idx, val in enumerate(row, 1):
        cell = ws16.cell(row=r_idx, column=c_idx, value=val)
        cell.alignment = Alignment(vertical="top", wrap_text=True)
        cell.border = BORDER
        if r_idx % 2 == 0:
            cell.fill = ALT_FILL
widths16 = [22, 36, 18, 12, 50, 50]
for i, w in enumerate(widths16, 1):
    ws16.column_dimensions[get_column_letter(i)].width = w
ws16.row_dimensions[1].height = 30
ws16.freeze_panes = "A2"


# ---------- Sheet 17: Guenstigste Plattform 7m fuer Dani-Setup ----------
ws17 = wb.create_sheet("17 Plattform 7m Kostenrechnung")
ws17.column_dimensions["A"].width = 50
ws17.column_dimensions["B"].width = 12
ws17.column_dimensions["C"].width = 12
ws17.column_dimensions["D"].width = 14
ws17.column_dimensions["E"].width = 55

t = ws17.cell(row=1, column=1, value="Guenstigste isolierte Plattform 7m fuer Algarve-Winter (Sebastians Wunsch Mai 2026)")
t.font = TITLE_FONT
ws17.merge_cells("A1:E1")
ws17.row_dimensions[1].height = 28

for col, h in enumerate(["Position", "Menge", "Einheit", "Gesamt EUR", "Bemerkung"], 1):
    c = ws17.cell(row=2, column=col, value=h)
    c.fill = HEADER_FILL
    c.font = HEADER_FONT
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = BORDER

platt = [
    ("VARIANTE A - Absolut guenstig EPS-Daemmung", "", "", "",
     "Cheapest fuer 41 m² Plattform; Schimmel-Risiko langfristig"),
    ("Plattform-Durchmesser 7,2 m (41 m² Flaeche)", "7,2 m", "Aussen", "",
     "Yurt 7m + 10cm Ueberstand"),
    ("Fundament Stein-/min. Beton 8 Punkte", 8, "Stueck", 145,
     "Lokal beschaffbar"),
    ("Tragbalken Kiefer KDI 50x200x4m", 25, "Stueck", 700,
     "Leroy Merlin / AKI / Bauhaus PT"),
    ("Schrauben+Beschlaege A4", 1, "Pauschal", 150, "Wuerth PT"),
    ("PE-Folie 200 µ", 45, "m²", 90, "Bauhaus 2 EUR/m²"),
    ("EPS 100mm WLG 035 (R~2,8)", 41, "m²", 205,
     "Cheapest aber nicht atmungsaktiv"),
    ("OSB-Verlegeplatte 22mm Deck", 41, "m²", 410, "Bedeckbar mit Teppich"),
    ("Boden-Schrauben Spax A4", 1, "Pauschal", 50, ""),
    ("Reserve 10%", 1, "Pauschal", 175, ""),
    ("VARIANTE A TOTAL", "", "", 1925, "Absolutes Minimum"),

    ("", "", "", "", ""),
    ("VARIANTE B - Eco-guenstig mit KORK EMPFEHLUNG", "", "", "",
     "Portugiesisches Material, atmungsaktiv"),
    ("Fundament Stein-/Beton 8 Punkte", 8, "Stueck", 145, ""),
    ("Tragbalken Kiefer KDI 50x200x4m", 25, "Stueck", 700, ""),
    ("Schrauben+Beschlaege A4", 1, "Pauschal", 150, ""),
    ("PE-Folie 200 µ", 45, "m²", 90, ""),
    ("KORK 80mm Sofalca/Amorim (R~2,0)", 41, "m²", 738,
     "Portugiesisches Material, diffusionsoffen"),
    ("Kiefer KDI T+G Boden 22mm", 41, "m²", 738, "Schoenere Optik"),
    ("Boden-Schrauben", 1, "Pauschal", 50, ""),
    ("Reserve 10%", 1, "Pauschal", 270, ""),
    ("VARIANTE B TOTAL", "", "", 2880,
     "EMPFEHLUNG - 955 mehr als A, dafuer oeko+atmungsaktiv"),

    ("", "", "", "", ""),
    ("VARIANTE C - Variante B + Krinner statt Beton", "", "", "",
     "Kein Beton, vollstaendig reversibel"),
    ("Krinner KSF-M 800mm Schraubpfaehle", 8, "Stueck", 280,
     "Statt Stein/Beton"),
    ("Erdraketen-Miete oder Eigenhand", 1, "Pauschal", 80,
     "Bauhaus PT, oder per Hand 4 Std Arbeit gratis"),
    ("Rest wie Variante B (Tragwerk+Kork+Deck+Schrauben+Reserve)", "", "", 2735, ""),
    ("VARIANTE C TOTAL", "", "", 3095,
     "Top-Setup: kein Beton + oeko + Kork"),

    ("", "", "", "", ""),
    ("R-WERT VERGLEICH", "", "", "", ""),
    ("EPS 100mm Plattform", "", "", "R~2,8", "Sehr gut thermisch nicht atmungsaktiv"),
    ("Kork 80mm Plattform", "", "", "R~2,0", "Gut + diffusionsoffen"),
    ("Dani Wand 4cm Schafwolle", "", "", "R~1,0",
     "Plattform mit Kork ist DOPPELT so gut wie Danis Wand - kompensiert kalte Fuesse im Algarve-Winter"),

    ("", "", "", "", ""),
    ("PT-LIEFERANTEN", "", "", "", ""),
    ("Kiefer KDI Tragwerk", "", "", "", "Leroy Merlin / AKI / Bauhaus PT"),
    ("Kork 80mm", "", "", "",
     "Sofalca (Coruche) https://www.sofalca.pt/ ODER Amorim (Santa Maria) https://www.amorim.com/"),
    ("OSB / Kiefer-Boden", "", "", "", "Bauhaus / Leroy Merlin / AKI"),
    ("Schrauben A4", "", "", "", "Wuerth Portugal"),
    ("Krinner-Pfaehle", "", "", "", "Krinner DE-Versand oder Bauhaus PT"),

    ("", "", "", "", ""),
    ("INTEGRATION DANI 7M + KORK-PLATTFORM", "", "", "", ""),
    ("Dani 7m mit Aufbau", "", "", 16000, ""),
    ("Plattform Variante B (Kork DIY)", "", "", 2880, ""),
    ("Transport Castelo de Vide -> Aljezur", "", "", 1000, "Schaetzung 400km"),
    ("Werkzeug-Block (falls neu)", "", "", 300, "0 wenn vorhanden"),
    ("Schattennetz Algarve", "", "", 240, ""),
    ("TOTAL ALL-IN", "", "", 20420,
     "Kork-Plattform spart 1.000-2.000 vs Casa-Style Premium-Plattform = Spielraum fuer 8cm Schafwolle-Upgrade bei Dani"),
]

row = 3
for entry in platt:
    a, b, c, d, e = entry
    if isinstance(a, str) and a.startswith("VARIANTE") and "TOTAL" not in a:
        cell = ws17.cell(row=row, column=1, value=a)
        cell.font = Font(bold=True, size=12, color="2E5C8A")
        ce = ws17.cell(row=row, column=5, value=e)
        ce.font = Font(italic=True, color="666666")
        ce.alignment = Alignment(wrap_text=True, vertical="top")
        ws17.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
        ws17.row_dimensions[row].height = 22
    elif isinstance(a, str) and a in ("R-WERT VERGLEICH", "PT-LIEFERANTEN", "INTEGRATION DANI 7M + KORK-PLATTFORM"):
        cell = ws17.cell(row=row, column=1, value=a)
        cell.font = Font(bold=True, size=12, color="2E5C8A")
        ws17.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
        ws17.row_dimensions[row].height = 22
    elif not a and not b and not c and not d and not e:
        ws17.row_dimensions[row].height = 8
    else:
        c1 = ws17.cell(row=row, column=1, value=a)
        c1.alignment = Alignment(wrap_text=True, vertical="top")
        c1.border = BORDER
        c2 = ws17.cell(row=row, column=2, value=b)
        c2.alignment = Alignment(horizontal="center", vertical="top")
        c2.border = BORDER
        c3 = ws17.cell(row=row, column=3, value=c)
        c3.alignment = Alignment(horizontal="center", vertical="top")
        c3.border = BORDER
        c4 = ws17.cell(row=row, column=4, value=d)
        c4.alignment = Alignment(horizontal="right", vertical="top")
        c4.border = BORDER
        if isinstance(d, (int, float)):
            c4.number_format = '#,##0 "EUR"'
        c5 = ws17.cell(row=row, column=5, value=e)
        c5.alignment = Alignment(wrap_text=True, vertical="top")
        c5.border = BORDER
        if isinstance(a, str) and ("TOTAL" in a):
            c1.font = Font(bold=True)
            c4.font = Font(bold=True, size=12, color="2E5C8A")
        ws17.row_dimensions[row].height = max(20, min(45, 15 + len(str(e)) // 8))
    row += 1
ws17.freeze_panes = "A2"


# ---------- Sheet 18: Mikael Pacific Yurts 3-Szenarien Kostenrechnung ----------
ws18 = wb.create_sheet("18 Mikael 3 Szenarien")
ws18.column_dimensions["A"].width = 48
ws18.column_dimensions["B"].width = 14
ws18.column_dimensions["C"].width = 14
ws18.column_dimensions["D"].width = 14
ws18.column_dimensions["E"].width = 45

t = ws18.cell(row=1, column=1, value="Mikael Pacific Yurts 8m - 3 Szenarien Vollkosten (Mai 2026)")
t.font = TITLE_FONT
ws18.merge_cells("A1:E1")
ws18.row_dimensions[1].height = 28

for col, h in enumerate(["Position", "A: Selbst max", "B: Mittel", "C: Komfort", "Bemerkung"], 1):
    c = ws18.cell(row=2, column=col, value=h)
    c.fill = HEADER_FILL
    c.font = HEADER_FONT
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = BORDER

szenarien = [
    ("Pacific Yurts 8m (Mikael Basis)", 15000, 15000, 15000, "Jurte selbst, alle Komponenten"),
    ("AC4-Boden 52 m² optional", 700, 700, 700, "Mikaels Selbstkosten-Preis fair"),
    ("Lieferung Castelo-Branco-Mitnahme (Schaetzung)", 500, 500, 800, "Mit Aufbauhilfe in Szenario C"),
    ("Zusatz-Daemmung 8mm Schafwoll-Filz Material", 600, 600, 800, "C: Mikael waehlt dickeres Filz"),
    ("Termiten-Behandlung Holz", 500, 400, 500, "A: Material+Selbstarbeit / B+C: durch Mikael+Oel-Auffrischung"),
    ("Schrauben-Ersatz Edelstahl A4", 200, 600, 600, "A: Material+Selbstarbeit 30 Std / B+C: durch Mikael"),
    ("Plattform DIY Kork (Sheet 17)", 2880, 2880, 3300, "C: Premium Variante"),
    ("Schattennetz Algarve", 240, 240, 240, "Lokale Agrar-Koop"),
    ("Reserve / Werkzeug", 500, 500, 500, ""),
    ("", "", "", "", ""),
    ("TOTAL ALL-IN", 21120, 21420, 22440, "Spannweite 1.320 EUR"),
    ("Eigenarbeit (Stunden)", "~80 h", "~50 h", "~20 h", ""),
    ("", "", "", "", ""),
    ("VERGLEICH MIT DANI 7m Premium (21.620 EUR)", "", "", "", ""),
    ("Mikael Szenario A vs Dani", "-500", "", "", "Mikael 500 EUR guenstiger als Dani"),
    ("Mikael Szenario B vs Dani", "", "-200", "", "Praktisch gleich"),
    ("Mikael Szenario C vs Dani", "", "", "+820", "Komfort kostet 820 EUR Aufpreis"),
    ("", "", "", "", ""),
    ("ABER: Dani Vorteile", "", "", "", ""),
    ("Daemmung", "7+8 mm", "7+8 mm", "bis 14 mm", "Dani 40mm = 2x dicker als Mikael max"),
    ("Probeschlafen", "nein", "nein", "nein", "Dani JA - Live-Verifikation"),
    ("Community", "nein", "nein", "nein", "Dani: 4 Familien vor Ort"),
    ("Rost-Risiko", "moderat", "behoben", "behoben", "Dani: keiner"),
    ("Termiten-Risiko", "moderat", "behoben", "behoben", "Dani: keiner (atmungsaktive Bauweise)"),
    ("Lagerungs-Risiko 10 Mon Van", "ja", "ja", "ja", "Dani: frisch gebaut Aug/Sep"),
    ("Wohnflaeche", "50 m²", "50 m²", "50 m²", "Dani: 38 m²"),
    ("Markenwert", "Pacific Yurts Premium", "Pacific Yurts Premium", "Pacific Yurts Premium", "Dani: Selbstbau 8J Erfahrung"),
]

row = 3
for entry in szenarien:
    a, b, c, d, e = entry
    if not a and not b and not c and not d:
        ws18.row_dimensions[row].height = 8
    elif isinstance(a, str) and (a.startswith("VERGLEICH") or a == "ABER: Dani Vorteile"):
        cell = ws18.cell(row=row, column=1, value=a)
        cell.font = Font(bold=True, size=12, color="2E5C8A")
        ws18.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
        ws18.row_dimensions[row].height = 22
    else:
        c1 = ws18.cell(row=row, column=1, value=a)
        c1.alignment = Alignment(wrap_text=True, vertical="top")
        c1.border = BORDER
        c2 = ws18.cell(row=row, column=2, value=b)
        c2.alignment = Alignment(horizontal="right", vertical="top")
        c2.border = BORDER
        if isinstance(b, (int, float)):
            c2.number_format = '#,##0 "EUR"'
        c3 = ws18.cell(row=row, column=3, value=c)
        c3.alignment = Alignment(horizontal="right", vertical="top")
        c3.border = BORDER
        if isinstance(c, (int, float)):
            c3.number_format = '#,##0 "EUR"'
        c4 = ws18.cell(row=row, column=4, value=d)
        c4.alignment = Alignment(horizontal="right", vertical="top")
        c4.border = BORDER
        if isinstance(d, (int, float)):
            c4.number_format = '#,##0 "EUR"'
        c5 = ws18.cell(row=row, column=5, value=e)
        c5.alignment = Alignment(wrap_text=True, vertical="top")
        c5.border = BORDER
        if isinstance(a, str) and "TOTAL" in a:
            c1.font = Font(bold=True)
            c2.font = Font(bold=True, size=12, color="2E5C8A")
            c3.font = Font(bold=True, size=12, color="2E5C8A")
            c4.font = Font(bold=True, size=12, color="2E5C8A")
        ws18.row_dimensions[row].height = max(20, min(45, 15 + len(str(e)) // 8))
    row += 1
ws18.freeze_panes = "A2"


# Speichern
output_path = "/home/user/Jurte/Jurten_Recherche_Algarve.xlsx"
wb.save(output_path)
print(f"XLSX gespeichert: {output_path}")
