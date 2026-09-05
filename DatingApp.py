#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DatingApp  –  BLOCK 1 von 2
=================================================================
Dieses Skript besteht aus zwei Blöcken:
  Block 1 (diese Datei): Datenmodelle, Parser, Analyse-Engine
  Block 2 (anhängen):    GUI (tkinter)

Verwendung:  Blöcke 1 + 2 in eine .py-Datei zusammenfügen und starten.
Anforderungen: Python >= 3.8, nur Standardbibliothek.
Optional für Drag&Drop: pip install tkinterdnd2

Log-Datei: ~/datierungshilfe.log
"""

import csv, json, logging, re, sys, traceback, webbrowser
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple
import tkinter as tk   # nur für TYPE_CHECKING, kein GUI hier
from tkinter import filedialog, messagebox, scrolledtext
from tkinter import ttk

# ── Drag & Drop (optional) ────────────────────────────────────
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD as _TkDnD
    _APP_BASE = _TkDnD.Tk
    DND_OK = True
except Exception:
    _APP_BASE = tk.Tk
    DND_OK = False

# ── Version & Logging ─────────────────────────────────────────
VERSION     = "3.2.0"
APP_TITEL   = "DatingApp"
LOG_PATH    = Path.home() / "datierungshilfe.log"
PAYPAL_URL  = "https://paypal.me/PeMeurer"

_fmt = logging.Formatter("%(asctime)s [%(levelname)-8s] %(message)s", "%H:%M:%S")
_fh  = logging.FileHandler(LOG_PATH, encoding="utf-8", mode="a")
_fh.setFormatter(_fmt); _fh.setLevel(logging.DEBUG)
_ch  = logging.StreamHandler(sys.stdout)
_ch.setFormatter(_fmt); _ch.setLevel(logging.WARNING)
logging.basicConfig(level=logging.DEBUG, handlers=[_fh, _ch])
log = logging.getLogger("DatingApp")

def _exc_hook(t, v, tb):
    log.critical("Unbehandelte Ausnahme", exc_info=(t,v,tb))
sys.excepthook = _exc_hook

# ── Texte ─────────────────────────────────────────────────────
README_TEXT = """\
DatingApp/Datierungshilfe  v{v}  –  Hilfe
=========================================================

Kurzanleitung
-------------
1. Fundliste laden        : Reiter „Fundliste" → Datei wählen oder hineinziehen
2. Stellenkatalog laden   : Datei → Stellenkatalog laden
3. Stellen verwalten      : Reiter „Stellen & Harris"
4. Analyse starten        : F5  oder  Menü Analyse → Analyse ausführen
5. Ergebnis               : Reiter „Auswertung"
6. Harris-Matrix (grafisch): Reiter „Harris (Grafik)"
7. Projekt speichern      : Strg+S

Überblick
---------
Die DatierungsHilfe ist ein Python-basiertes Werkzeug zur strukturierten
Erfassung, Verwaltung und Auswertung archäologischer Funddaten sowie
zugehöriger Stelleninformationen. Ziel ist die Unterstützung bei der
relativen und absoluten Datierung durch konsistente Datenhaltung und
automatisierte Auswertung.

Das Programm kombiniert:
- Fundlisten (Einzelfunde)
- Stellenkataloge (Kontexte/Befunde)
- Datierungslogik (zusammengeführte Interpretation)

Funktionalitäten
----------------
1. Laden und Verwalten von Daten
   - Import von Fundlisten
   - Import von Stellenkatalogen
   - Konsolidierung der Datenbestände

2. Darstellung
   - Tabellarische Anzeige von Funden und Stellen
   - Strukturierte Visualisierung relevanter Attribute
   - Markierung von Auffälligkeiten (z. B. Konflikte, Unsicherheiten)

3. Datierung
   - Berechnung effektiver Datierungen auf Basis der Fundlage
   - Kennzeichnung von Datierungsquellen
   - Zusammenführung mehrerer Datierungsindikatoren

4. Qualitätskontrolle
   - Erkennung von Inkonsistenzen
   - Warnmeldungen bei widersprüchlichen Daten
   - Nachvollziehbarkeit der Datierungsentscheidungen

Technische Grundlagen
--------------------
- Programmiersprache: Python 3
- GUI: tkinter (Treeview-basierte Tabellenansicht)
- Architektur:
  - Trennung von Datenmodell und Benutzeroberfläche
  - Ereignisgesteuerte Aktualisierung der Anzeige
  - Zentrale Aktualisierungsroutine (_aktualisiere_alles)

Datenstruktur (konzeptionell)
----------------------------
Fund:
    Repräsentiert ein einzelnes Artefakt mit Datierungsinformation.

Stelle:
    Repräsentiert einen Befund/Zusammenhang, der mehrere Funde enthält.
    Enthält:
    - Beschreibung
    - Zugeordnete Funde
    - Abgeleitete Datierung

Wichtige Hinweise
-----------------
- Jede Tabellenzeile benötigt eine eindeutige ID (iid).
  Diese wird intern als String behandelt.
- Doppelte IDs führen zu Laufzeitfehlern.
- Beim Einfügen in Treeviews muss folgende Signatur eingehalten werden:

      tree.insert(parent, index, iid=..., values=...)

  Falsche Parameterreihenfolgen führen zu Fehlern wie:
      "got multiple values for argument 'iid'"

- Vor dem Einfügen sollte geprüft werden, ob eine ID bereits existiert.

Typische Fehlerquellen
---------------------
- Mehrfache Verwendung derselben ID
- Falsche Parameterreihenfolge bei Treeview.insert()
- Inkonsistente oder unvollständige Eingabedaten
- Fehlende Synchronisation zwischen Datenmodell und GUI

Start und Nutzung
----------------
1. Programm starten
2. Fundliste laden
3. Stellenkatalog laden
4. Automatische Aktualisierung der Anzeige erfolgt
5. Ergebnisse in den Tabs überprüfen

Erweiterungsmöglichkeiten
------------------------
- Exportfunktionen (CSV, Excel)
- Erweiterte Datierungsalgorithmen
- Filter- und Suchfunktionen
- Visualisierung von Zeiträumen (Timeline)
- Datenbankanbindung

Zielgruppe
----------
Das Programm richtet sich an:
- Archäolog:innen
- Grabungsdokumentation
- Wissenschaftliche Auswertung

Feedback und Fehlermeldungen
----------------------------
Rückmeldungen sind willkommen. Bitte die Log-Datei beifügen, sie liegt
als datierungshilfe.log im Benutzerverzeichnis.

Spenden
-------
Wenn das Programm nützt: Über „Über → Spenden" öffnet sich
https://paypal.me/PeMeurer im Browser.

""".format(v=VERSION)

LIZENZ_TEXT = """\
Lizenz
======

© 2026 Peter Meurer

OpenSource Software (GNU GPL v3)
Keine Gewährleistung – Nutzung auf eigene Verantwortung.
Dieses Programm ist freie Software: Sie können es unter den Bedingungen der
GNU General Public License Version 3 (GPLv3) weitergeben und/oder verändern.

Dieses Programm wird in der Hoffnung bereitgestellt, dass es nützlich ist,
aber OHNE JEDE GEWÄHRLEISTUNG, sogar ohne die implizite Gewährleistung der
MARKTREIFE oder EIGNUNG FÜR EINEN BESTIMMTEN ZWECK.

Die Nutzung erfolgt auf eigene Verantwortung. Der Autor übernimmt keine
Haftung für Schäden, Datenverluste oder Fehlinterpretationen der Ergebnisse.

Unterstützung und Weiterentwicklung sind willkommen.

"""



# ═══════════════════════════════════════════════════════════════
# KONSTANTEN
# ═══════════════════════════════════════════════════════════════

DELIMITER   = ";"
ENCODING    = "windows-1252"
PRAEZISIONEN = ["Tag","Monat","Jahr","Jahrzehnt","Jahrhundert","Jahrtausend"]

MODERN_BEGRIFFE: Set[str] = {
    "plastik","kunststoff","beton","glas (modern)","neuzeit","rezent",
    "modern","zeitgenössisch","19. jh.","20. jh.","21. jh.",
    "industriell","porzellan (modern)","aluminium","gummi","asphalt",
}

# Befundansprachen → automatisch Arbeitsbereich setzen
ARBEITSBEREICH_ANSPRACHEN = {
    "arbeitsfläche (archäologie)","arbeitsfläche","profil","geosondage",
    "kein befund","suchschnitt","sondierschnitt","kontrollschnitt",
}
# Befundansprachen → automatisch Störung setzen
STOERUNG_ANSPRACHEN = {
    "störung (archäologie)","störung","baumwurf (archäologie)","baumwurf",
    "bioturbation","tiergrabung","wurzeln","rohrleitungsgraben (modern)",
    "leitungsgraben","fundamentgraben (modern)",
}

# ═══════════════════════════════════════════════════════════════
# EINGEBETTETE WNK-DATIERUNGEN  (339 Begriffe, Stand DokuPrüfer 2.02)
# ═══════════════════════════════════════════════════════════════
# 339 WNK-Datierungsbegriffe eingebettet
WNK_DATIERUNGEN_EINGEBETTET: dict = {
    'Datierung unbekannt': (-542000000, None),
    'Geologisch': (-542000000, 2100),
    'Paläozoikum': (-538800000, -251902000),
    'Kambrium': (-538800000, -486850000),
    'Terreneuvium': (-538800000, -521000000),
    'Fortunium': (-538800000, -529000000),
    'Kambrium-Stufe 2': (-529000000, -521000000),
    'Kambrium-Serie 2': (-521000000, -506500000),
    'Kambrium-Stufe 3': (-521000000, -514500000),
    'Kambrium-Stufe 4': (-514500000, -506500000),
    'Miaolingium': (-506500000, -497000000),
    'Wuliuum': (-506500000, -504500000),
    'Drumium': (-504500000, -500500000),
    'Guzhangium': (-500500000, -497000000),
    'Furongium': (-497000000, -486850000),
    'Paibium': (-497000000, -494200000),
    'Jiangshanium': (-494200000, -491000000),
    'Kambrium-Stufe 10': (-491000000, -486850000),
    'Ordovizium': (-486850000, -443100000),
    'Unterordovizium': (-486850000, -471300000),
    'Tremadocium': (-486850000, -477100000),
    'Floium': (-477100000, -471300000),
    'Mittelordovizium': (-471300000, -458200000),
    'Dapingium': (-471300000, -469400000),
    'Darriwilium': (-469400000, -458200000),
    'Oberordovizium': (-458200000, -443100000),
    'Sandbium': (-458200000, -452800000),
    'Katium': (-452800000, -445200000),
    'Hirnantium': (-445200000, -443100000),
    'Silur': (-443100000, -419620000),
    'Llandovery': (-443100000, -432900000),
    'Rhuddanium': (-443100000, -440500000),
    'Aeronium': (-440500000, -438600000),
    'Telychium': (-438600000, -432900000),
    'Wenlock': (-432900000, -426700000),
    'Sheinwoodium': (-432900000, -430600000),
    'Homerium': (-430600000, -426700000),
    'Ludlow': (-426700000, -422700000),
    'Gorstium': (-426700000, -425000000),
    'Ludfordium': (-425000000, -422700000),
    'Pridoli': (-422700000, -419620000),
    'Devon': (-419620000, -358860000),
    'Unterdevon': (-419620000, -393470000),
    'Lochkovium': (-419620000, -413020000),
    'Pragium': (-413020000, -410620000),
    'Emsium': (-410620000, -393470000),
    'Mitteldevon': (-393470000, -382310000),
    'Eifelium': (-393470000, -387950000),
    'Givetium': (-387950000, -382310000),
    'Oberdevon': (-382310000, -358860000),
    'Frasnium': (-382310000, -372150000),
    'Famennium': (-372150000, -358860000),
    'Karbon': (-358860000, -298900000),
    'Mississippium': (-358860000, -323400000),
    'Tournaisium': (-358860000, -346700000),
    'Viseum': (-346700000, -330300000),
    'Serpukhovium': (-330300000, -323400000),
    'Pennsylvanium': (-323400000, -298900000),
    'Bashkirium': (-323400000, -315200000),
    'Moskovium': (-315200000, -307000000),
    'Kasimovium': (-307000000, -303700000),
    'Gzhelium': (-303700000, -298900000),
    'Perm': (-298900000, -251902000),
    'Cisuralium': (-298900000, -274400000),
    'Asselium': (-298900000, -293520000),
    'Sakmarium': (-293520000, -290100000),
    'Artinskium': (-290100000, -283300000),
    'Kungurium': (-283300000, -274400000),
    'Guadalupium': (-274400000, -259510000),
    'Roadium': (-274400000, -266900000),
    'Wordium': (-266900000, -264280000),
    'Capitanium': (-264280000, -259510000),
    'Lopingium': (-259510000, -251902000),
    'Wuchiapingium': (-259510000, -254140000),
    'Changhsingium': (-254140000, -251902000),
    'Mesozoikum': (-251902000, -66000000),
    'Trias': (-251902000, -201400000),
    'Untertrias': (-251902000, -246700000),
    'Indusium': (-251902000, -249900000),
    'Olenekium': (-249900000, -246700000),
    'Mitteltrias': (-246700000, -237000000),
    'Anisium': (-246700000, -241464000),
    'Ladinium': (-241464000, -237000000),
    'Obertrias': (-237000000, -201400000),
    'Karnium': (-237000000, -227300000),
    'Norium': (-227300000, -205700000),
    'Rhaetium': (-205700000, -201400000),
    'Jura (Mesozoikum)': (-201400000, -143100000),
    'Unterjura': (-201400000, -174700000),
    'Hettangium': (-201400000, -199500000),
    'Sinemurium': (-199500000, -192900000),
    'Pliensbachium': (-192900000, -184200000),
    'Toarcium': (-184200000, -174700000),
    'Mitteljura': (-174700000, -161500000),
    'Aalenium': (-174700000, -170900000),
    'Bajocium': (-170900000, -168200000),
    'Bathonium': (-168200000, -165300000),
    'Callovium': (-165300000, -161500000),
    'Oberjura': (-161500000, -143100000),
    'Oxfordium': (-161500000, -154800000),
    'Kimmeridgium': (-154800000, -149200000),
    'Tithonium': (-149200000, -143100000),
    'Kreide (Geologie)': (-143100000, -66000000),
    'Unterkreide': (-143100000, -100500000),
    'Berriasium': (-143100000, -137050000),
    'Valanginium': (-137050000, -132600000),
    'Hauterium': (-132600000, -125770000),
    'Barremium': (-125770000, -121400000),
    'Aptium': (-121400000, -113200000),
    'Albium': (-113200000, -100500000),
    'Oberkreide': (-100500000, -66000000),
    'Cenomanium': (-100500000, -93900000),
    'Turonium': (-93900000, -89800000),
    'Coniacium': (-89800000, -85700000),
    'Santonium': (-85700000, -83600000),
    'Campanium': (-83600000, -72200000),
    'Maastrichtium': (-72200000, -66000000),
    'Känozoikum': (-66000000, None),
    'Paläogen': (-66000000, -23040000),
    'Paläozän': (-66000000, -56000000),
    'Danium': (-66000000, -61660000),
    'Tertiär': (-65000000, -1800000),
    'Seelandium': (-61660000, -59240000),
    'Thanetium': (-59240000, -56000000),
    'Eozän': (-56000000, -33900000),
    'Ypresium': (-56000000, -48070000),
    'Lutetium (Eozän)': (-48070000, -41030000),
    'Bartonium': (-41030000, -37710000),
    'Priabonium': (-37710000, -33900000),
    'Oligozän': (-33900000, -23040000),
    'Rupelium': (-33900000, -27300000),
    'Chattium': (-27300000, -23040000),
    'Neogen': (-23040000, -2580000),
    'Miozän': (-23040000, -5333000),
    'Untermiozän': (-23040000, -15980000),
    'Aquitanium': (-23040000, -20450000),
    'Burdigalium': (-20450000, -15980000),
    'Mittelmiozän': (-15980000, -11630000),
    'Langhium': (-15980000, -13820000),
    'Serravallium': (-13820000, -11630000),
    'Obermiozän': (-11630000, -5333000),
    'Tortonium': (-11630000, -7246000),
    'Messinium': (-7246000, -5333000),
    'Pliozän': (-5333000, -2580000),
    'Zancleum': (-5333000, -3600000),
    'Piacenzium': (-3600000, -2580000),
    'Quartär': (-2580000, None),
    'Pleistozän': (-2580000, -11700),
    'Gelasium': (-2580000, -1800000),
    'Calabrium': (-1800000, -774000),
    'Chibanium': (-774000, -129000),
    'Oberpleistozän': (-129000, -11700),
    'Grönlandium': (-117000, -8200),
    'Holozän': (-11700, None),
    'Nordgrippium': (-8200, -4200),
    'Meghalayum': (-4200, None),
    'Urgeschichte bis Neuzeit': (-780000, 2100),
    'Urgeschichte': (-780000, -15),
    'Steinzeit': (-780000, -2150),
    'Paläolithikum bis Mesolithikum': (-780000, -5300),
    'Paläolithikum': (-780000, -9600),
    'Altpaläolithikum': (-780000, -300000),
    'Mittelpaläolithikum': (-300000, -35000),
    'Spätpaläolithikum': (-35000, -9600),
    'Jungpaläolithikum': (-35000, -13000),
    'Aurignacien': (-35000, -30000),
    'Gravettien': (-30000, -15000),
    'Magdalénien': (-15000, -13000),
    'Endpaläolithikum bis Mesolithikum': (-13000, -5300),
    'Endpaläolithikum': (-13000, -9600),
    'Federmessergruppen': (-13000, -12000),
    'Ahrensburger Kultur': (-12000, -9600),
    'Mesolithikum bis Metallzeit': (-9600, -15),
    'Mesolithikum bis Neolithikum': (-9600, -2150),
    'Mesolithikum': (-9600, -5300),
    'Neolithikum bis Metallzeit': (-5300, -15),
    'Neolithikum': (-5300, -2150),
    'Alt- bis Jungneolithikum': (-5300, -3500),
    'Alt- bis Mittelneolithikum': (-5300, -4300),
    'Altneolithikum/Bandkeramik': (-5300, -4900),
    'Mittelneolithikum bis Metallzeit': (-4900, -15),
    'Mittel- bis Spätneolithikum': (-4900, -2800),
    'Mittel- bis Jungneolithikum': (-4900, -3500),
    'Mittelneolithikum': (-4900, -4300),
    'Großgartach': (-4900, -4600),
    'Planig-Friedberg': (-4800, -4750),
    'Rössen': (-4750, -4600),
    'Bischheim': (-4600, -4300),
    'Jung- bis Endneolithikum': (-4300, -2150),
    'Jung- bis Spätneolithikum': (-4300, -2800),
    'Jungneolithikum/Michelsberg': (-4300, -3500),
    'Jungneolithikum': (None, None),
    'Spätneolithikum bis Bronzezeit': (-3500, -800),
    'Spät- bis Endneolithikum': (-3500, -2150),
    'Spätneolithikum/Wartberg': (-3500, -2800),
    'Endneolithikum/Rheinische Becherkultur': (-2800, -2150),
    'Schnurkeramik': (-2800, -2400),
    'Glockenbecher': (-2400, -2150),
    'Metallzeit bis Neuzeit': (-2150, 2100),
    'Metallzeit': (-2150, -15),
    'Bronzezeit bis Hallstattzeit': (-2150, -475),
    'Bronzezeit': (-2150, -800),
    'Bronzezeit bis Hallstattzeit A': (-2150, -1000),
    'Ältere Bronzezeit': (-2150, -1300),
    'Frühe Bronzezeit': (-2150, -1500),
    'Mittlere Bronzezeit': (-1500, -1300),
    'Urnenfelderzeit bis Eisenzeit': (-1300, -15),
    'Urnenfelderzeit bis Frühlatène': (-1300, -250),
    'Jüngere Bronzezeit': (-1300, -800),
    'Bronzezeit D': (-1300, -1200),
    'Hallstatt A': (-1200, -1000),
    'Hallstatt B/C': (-1000, -600),
    'Hallstatt B': (-1000, -800),
    'Eisenzeit bis Neuzeit': (-800, 2100),
    'Eisenzeit bis Hochmittelalter': (-800, 1300),
    'Eisenzeit bis Germanisch': (-800, 450),
    'Hallstatt C bis Frühlatène': (-800, 250),
    'Eisenzeit': (-800, -15),
    'Ältere Eisenzeit': (-800, -475),
    'Ältere Hallstattzeit': (-800, -600),
    'Hallstatt D bis Frühlatène': (-600, 250),
    'Hallstatt D bis Spätlatène': (-600, -15),
    'Jüngere Hallstattzeit': (-600, -475),
    'Jüngere Eisenzeit': (-475, -15),
    'Früh- bis Mittellatènezeit': (-475, -120),
    'Frühlatènezeit': (-475, -250),
    'Mittel- bis Spätlatènezeit': (-250, -15),
    'Mittellatènezeit': (-250, -120),
    'Römisch, 2.-1. Jh. v. Chr.': (-200, -1),
    'Römisch, 2. Jh. v. Chr.': (-200, -101),
    'Latène D bis frühe Römische Kaiserzeit': (-120, 70),
    'Spätlatènezeit': (-120, -15),
    'Römisch, 1. Jh. v. Chr. - 4. Jh.': (-100, 400),
    'Römisch, 1. Jh. v. Chr. - 3. Jh.': (-100, 300),
    'Römisch, 1. Jh. v. Chr. - 2. Jh.': (-100, 200),
    'Römisch, 1. Jh. v. Chr. - 1. Jh.': (-100, 100),
    'Römisch, 1. Jh. v. Chr.': (-100, -1),
    'Römische Kaiserzeit bis Neuzeit': (-27, None),
    'Römische Kaiserzeit bis Mittelalter': (-27, 1500),
    'Römisch': (-27, 450),
    'Germanisch/einheimisch': (-27, 450),
    'Römische Kaiserzeit': (-27, 450),
    'Germanisch A': (-27, 14),
    'Römisch, 1.-5. Jh.': (1, 500),
    'Römisch, 1.-4. Jh.': (1, 400),
    'Römisch, 1.-3. Jh.': (1, 300),
    'Römisch, 1.-2. Jh.': (1, 200),
    'Römisch, 1. Jh.': (1, 100),
    'Germanisch B/C2': (14, 305),
    'Germanisch B': (14, 180),
    'Germanisch B1': (14, 70),
    'Germanisch B2': (70, 180),
    'Römisch, 2.-5. Jh.': (101, 500),
    'Römisch, 2.-4. Jh.': (101, 400),
    'Römisch, 2.-3. Jh.': (101, 300),
    'Römisch, 2. Jh.': (101, 200),
    'Germanisch C': (180, 370),
    'Germanisch C1': (180, 235),
    'Römisch, 3.-5. Jh.': (201, 500),
    'Römisch, 3.-4. Jh.': (201, 400),
    'Römisch, 3. Jh.': (201, 300),
    'Germanisch C2': (235, 305),
    'Späte römische Kaiserzeit bis Frühmittelalter': (284, 500),
    'Römisch, 4.-5. Jh.': (301, 500),
    'Römisch, 4. Jh.': (301, 400),
    'Germanisch C3/D': (305, 450),
    'Germanisch C3': (305, 370),
    'Germanisch D': (370, 450),
    'Frühmittelalter, 5.-8. Jh.': (401, 800),
    'Frühmittelalter, 5.-7. Jh.': (401, 700),
    'Frühmittelalter, 5.-6. Jh.': (401, 600),
    'Römisch, 5. Jh.': (401, 500),
    'Mittelalter bis Neuzeit': (450, None),
    'Mittelalter bis Mitte 19. Jh.': (450, 1850),
    'Mittelalter bis Frühneuzeit': (450, 1600),
    'Früh- bis Spätmittelalter': (450, 1500),
    'Mittelalter': (450, 1500),
    'Früh- bis Hochmittelalter': (450, 1300),
    'Merowingisch bis ottonisch': (450, 1024),
    'Frühmittelalter': (450, 919),
    'Merowingerzeit': (450, 750),
    'Frühe Merowingerzeit': (450, 500),
    'Frühmittelalter, 6.-9. Jh.': (501, 900),
    'Frühmittelalter, 6.-8. Jh.': (501, 800),
    'Merowingerzeit, 6.-7. Jh.': (501, 700),
    'Merowingerzeit, 6. Jh.': (501, 600),
    'Frühmittelalter, 7.-9. Jh.': (601, 900),
    'Frühmittelalter, 7.-8. Jh.': (601, 800),
    'Merowingerzeit, 7. Jh.': (601, 700),
    'Frühmittelalter, 8.-9. Jh.': (701, 900),
    'Ausgehende Merowingerzeit': (701, 750),
    'Karolingisch bis Hochmittelalter': (750, 1300),
    'Karolingisch bis ottonisch': (750, 1024),
    'Karolingerzeit': (750, 919),
    'Frühe Karolingerzeit': (750, 820),
    'Mittlere Karolingerzeit': (820, 880),
    'Ausgehende Karolingerzeit': (880, 919),
    'Hochmittelalter bis Neuzeit': (901, 2100),
    'Hoch- bis Spätmittelalter': (901, 1500),
    'Hochmittelalter': (901, 1300),
    'Hochmittelalter, 10.-12. Jh.': (901, 1200),
    'Hochmittelalter, 10.-11. Jh.': (901, 1100),
    'Hochmittelalter, 10. Jh.': (901, 1000),
    'Ottonisch': (919, 1024),
    'Hochmittelalter, 11.-13. Jh.': (1001, 1300),
    'Hochmittelalter, 11.-12. Jh.': (1001, 1200),
    'Hochmittelalter, 11. Jh.': (1001, 1100),
    'Salisch': (1024, 1125),
    'Hochmittelalter, 12.-13. Jh.': (1101, 1300),
    'Hochmittelalter, 12. Jh.': (1101, 1200),
    'Staufisch': (1137, 1250),
    '13.-14. Jahrhundert n. Chr.': (1201, 1400),
    'Hochmittelalter, 13. Jh.': (1201, 1300),
    'Spätmittelalter': (1301, 1500),
    'Spätmittelalter, 14. Jh.': (1301, 1400),
    'Spätmittelalter bis Neuzeit': (1401, None),
    'Spätmittelalter bis Frühneuzeit, 15.-16. Jh.': (1401, 1600),
    'Spätmittelalter, 15. Jh.': (1401, 1500),
    'Neuzeit': (1501, None),
    'Neuzeit, 16.-20. Jh.': (1501, 2000),
    'Neuzeit, 16.-19. Jh.': (1501, 1900),
    'Neuzeit, 16.-18. Jh.': (1501, 1800),
    'Neuzeit, 16.-17. Jh.': (1501, 1700),
    'Neuzeit, 16. Jh.': (1501, 1600),
    'Neuzeit, 17.-21. Jh.': (1601, 2100),
    'Neuzeit, 17.-20. Jh.': (1601, 2000),
    'Neuzeit, 17.-19. Jh.': (1601, 1900),
    'Neuzeit, 17.-18. Jh.': (1601, 1800),
    'Neuzeit, 17. Jh.': (1601, 1700),
    'Neuzeit, 18.-21. Jh.': (1701, 2100),
    'Neuzeit, 18.-20. Jh.': (1701, 2000),
    'Neuzeit, 18.-19. Jh.': (1701, 1900),
    'Neuzeit, 18. Jh.': (1701, 1800),
    'Neuzeit, 19.-21. Jh.': (1801, 2100),
    'Neuzeit, 19.-20. Jh.': (1801, 2000),
    'Neuzeit, 19. Jh.': (1801, 1900),
    'Neuzeit, 20.-21. Jh.': (1901, 2100),
    'Neuzeit, 20. Jh.': (1901, 2000),
    'Neuzeit, 21. Jh.': (2001, 2100),
}

# Unbekannte Datierungsbegriffe getrennt nach Quelle, damit ein Neuladen
# der Fundliste die Katalog-Begriffe nicht löscht (und umgekehrt) und der
# rote Warn-Button nach einer korrigierten CSV wirklich verschwindet
_UNBEKANNT_FUND: set = set()
_UNBEKANNT_KATALOG: set = set()

def _unbekannte_datierungen() -> set:
    return _UNBEKANNT_FUND | _UNBEKANNT_KATALOG

_WNK_LOWER: Dict[str, Tuple[Optional[int], Optional[int]]] = {
    k.lower(): v for k, v in WNK_DATIERUNGEN_EINGEBETTET.items()
}
# Erweiterte Suche: auch Teilübereinstimmungen mit WNK-Schlüsseln
def _dat_lookup_erweitert(begriff: str):
    """Sucht in WNK auch nach Teilübereinstimmungen für neue Begriffe."""
    b = begriff.lower().strip()
    if b in _WNK_LOWER: return _WNK_LOWER[b]
    # Ersten passenden WNK-Eintrag der den Begriff enthält
    for k, v in _WNK_LOWER.items():
        if b in k or k in b: return v
    return (None, None)
# Linienfarben Harris-Matrix (Dark / Light)
_LFARBEN: dict = {
    "dark": {
        "manual":  ("#e8e8e8", False),
        "profil":  ("#b0b0b0", False),
        "dat":     ("#69d46e", False),
        "okuk":    ("#ffe066", True ),
        "widersp": ("#ff6b6b", False),
        "gleich":  ("#9e9e9e", True ),
    },
    "light": {
        "manual":  ("#333333", False),
        "profil":  ("#757575", False),
        "dat":     ("#2e7d32", False),
        "okuk":    ("#e65100", True ),
        "widersp": ("#c62828", False),
        "gleich":  ("#616161", True ),
    },
}

# Sortierte Begriffsliste (WNK-Datierungen für Jahres-Lookup)
_WNK_BEGRIFFE: List[str] = sorted(WNK_DATIERUNGEN_EINGEBETTET.keys())

# Befundansprachen (Stand der WNK-Liste, vgl. Liste_WNK-Befundansprachen.csv)
_ANSPRACHEN_LISTE: List[str] = [
    'A-Horizont',
    'Abbaukammer',
    'Abdruck (Vertiefung)',
    'Abfallgrube',
    'Abortgrube',
    'Abraumhalde',
    'Abri',
    'Abschnittsbefestigung',
    'Absetzbecken',
    'Absturzstelle',
    'Abzugsloch (Ofen)',
    'Alenkastell',
    'Altacker (Landwirtschaft)',
    'Amphitheater',
    'Anlegestelle',
    'Aquädukt',
    'Aquäduktbrücke',
    'Aquädukttunnel',
    'Arbeitsfläche (Archäologie)',
    'Arbeitsgrube',
    'Arbeitslager',
    'Aschengrube (Ofen)',
    'Aschenkiste',
    'Atriumhaus',
    'Auensediment',
    'Aufschüttung (Bauwerk)',
    'Ausbruchgraben',
    'Ausbruchgrube',
    'Ausbruchsbefund',
    'Ausfachung',
    'Auxiliarlager',
    'Backhaus',
    'Backofen',
    'Bad (Bauwerk)',
    'Badehaus',
    'Balken (Bauteil)',
    'Baptisterium',
    'Baracke',
    'Barbakane',
    'Basilika (Halle)',
    'Basilika (Kirche)',
    'Bastion',
    'Batterie (Festung)',
    'Baugrube',
    'Bauhütte (Bauwerk)',
    'Baumsarg',
    'Baumstandspur',
    'Baumwurfgrube',
    'Becken (Bauwerk)',
    'Befestigungsanlage',
    'Befund (Archäologie)',
    'Befund nicht interpretierbar',
    'Beigabennische',
    'Benefiziarierstation',
    'Bergbaugebiet',
    'Bergfried',
    'Bergwerk',
    'Bergwerksschacht',
    'Bergwerksstollen',
    'Berme',
    'Bestattungsplatz',
    'Bildstock (Bauwerk)',
    'Bioturbation',
    'Bleihütte (Wirtschaft)',
    'Bohrloch',
    'Bombentrichter',
    'Brandgrab',
    'Brandgrubengrab',
    'Brandschüttungsgrab',
    'Brauerei',
    'Brennraum (Ofen)',
    'Brennraumboden',
    'Brunnen',
    'Brunnenfassung',
    'Brunnenschacht',
    'Brunnenstube',
    'Brücke (Bauwerk)',
    'Bunker (Bauwerk)',
    'Burg',
    'Burgus',
    'Bustum',
    'Canabae legionis',
    'Cella memoria',
    'Colonia (Siedlung)',
    'Contrescarpe',
    'Darre',
    'Deckungsgraben',
    'Deich',
    'Depot (Archäologie)',
    'Dom',
    'Donjon',
    'Dorf',
    'Dorfkern',
    'Dormitorium',
    'Dungschleier',
    'Einfriedung',
    'Einkammerofen',
    'Einsiedelei',
    'Einzelfund',
    'Einzelhof',
    'Eisenbahnstrecke',
    'Eisenhütte',
    'Eisgrube',
    'Eiskeller',
    'Entwässerungsgraben',
    'Erdkeller',
    'Erdwerk (Archäologie)',
    'Erzröstplatz',
    'Erzwäsche (Vorrichtung)',
    'Eskarpe',
    'Explosionskrater',
    'Fabrica',
    'Fabrik (Baukomplex)',
    'Fachwerkständerbau',
    'Faschine',
    'Feldbegrenzung (Landwirtschaft)',
    'Feldbrandofen',
    'Feldstellung',
    'Felsbild',
    'Felsheiligtum',
    'Festes Haus',
    'Festung',
    'Feuerstelle',
    'Feuerungskanal',
    'Feuerungsraum',
    'Feuerungstür',
    'Findling (Geologie)',
    'Fischteich',
    'Flachgrab',
    'Flachsröste (Grube)',
    'Flakstellung',
    'Fließgewässer',
    'Flottenlager',
    'Fluchtgang',
    'Flughafen',
    'Flugzeug',
    'Flurform',
    'Flusskies',
    'Forum (Offener Raum)',
    'Fossillagerstätte',
    'Freilandstation',
    'Friedhof',
    'Frigidarium',
    'Fritteofen',
    'Fund (Objekt)',
    'Fundament',
    'Fundamentgraben',
    'Fundkonzentration',
    'Fundplatzindikator',
    'Furt',
    'Fußboden',
    'Fußbodenheizung',
    'Fußbodenmosaik',
    'Fäkalienrinne',
    'Färberei (Wirtschaft)',
    'Färbergrube',
    'Galgenhügel',
    'Gang (Architektur)',
    'Garage',
    'Garten',
    'Gasthof',
    'Gaststätte',
    'Gebäude',
    'Gefängnisgebäude',
    'Gerberei',
    'Gerbergrube',
    'Geschützstand',
    'Getreidedarre',
    'Gewölbe (Bauteil)',
    'Glacis',
    'Glashütte',
    'Glasofen',
    'Glaswerkstatt',
    'Glockengussgrube',
    'Grab',
    'Grabbau (Bauwerk)',
    'Grabeinfriedung',
    'Graben (Erdbauwerk)',
    'Grabenanlage',
    'Grabgarten',
    'Grabgrube',
    'Grabhügel',
    'Grabkammer',
    'Grabungsfläche',
    'Grenzpunkt',
    'Grenzstein',
    'Grube (Erdbauwerk)',
    'Grubenhaus',
    'Grubenkomplex',
    'Gruft',
    'Grundriss',
    'Gräberfeld',
    'Gräfte',
    'Gästehaus',
    'Hafen',
    'Hafenofen',
    'Halbbastion',
    'Halde (Wirtschaft)',
    'Hallenhaus',
    'Hammerwerk',
    'Hauptburg',
    'Hauptgebäude',
    'Heiligtum (Raum)',
    'Herberge',
    'Herdstelle',
    'Hochofen',
    'Hockergrab',
    'Hof (Landwirtschaft)',
    'Hohlweg',
    'Holzwerkstatt',
    'Horreum',
    'Hypokaustum',
    'Höckerlinie',
    'Höhenburg',
    'Höhensiedlung',
    'Höhle',
    'Höhlenheiligtum',
    'Hügelgrab',
    'Immunitätsbezirk',
    'Industrieanlage',
    'Innenhof',
    'Insula (Siedlung)',
    'Irdenwareofen',
    'Jüdischer Friedhof',
    'Kaldarium',
    'Kalkbrennerei',
    'Kalklöschgrube',
    'Kalkofen',
    'Kalvarienberg',
    'Kammergrab',
    'Kanal (Wasserbau)',
    'Kanalisation',
    'Kapelle (Bauwerk)',
    'Kapitelsaal',
    'Karner',
    'Kasematte',
    'Kaserne',
    'Kavalier (Festung)',
    'Kein Befund',
    'Kenotaph',
    'Kirchengebäude',
    'Kirchhof',
    'Klausur (Kloster)',
    'Kleinkastell',
    'Kloake (Bauwerk)',
    'Kloakenrinne',
    'Kloster (Architektur)',
    'Knochenmühle',
    'Knochensiederei',
    'Knochenwerkstatt',
    'Knüppeldamm',
    'Kochgrube',
    'Kohlenmeiler',
    'Kohortenkastell',
    'Kolluvium',
    'Konzentrationslager',
    'Krankenhaus',
    'Kreisgraben',
    'Kreuzgang',
    'Kreuzwegstation',
    'Kriegsgefangenenlager',
    'Kultopfer',
    'Kurtine',
    'Körpergrab',
    'Küche (Raum)',
    'Kühlofen',
    'Künette',
    'Landschaft',
    'Landwehr (Bauwerk)',
    'Langhaus (Bauwerk)',
    'Langhügel (Grab)',
    'Latrine',
    'Laufhorizont',
    'Layerinterface',
    'Legionslager',
    'Legionsziegelei',
    'Leichenbrandlager',
    'Leichenschatten',
    'Leprosorium',
    'Lichtschacht (Architektur)',
    'Liegender Ofen',
    'Limes (Grenzbefestigung)',
    'Lochstein (Vermessungswesen)',
    'Lochtenne',
    'Luftschutzkeller',
    'Löschteich',
    'Lünette (Festung)',
    'Mansio',
    'Markierung (Archäologie)',
    'Marktplatz',
    'Marschlager',
    'Massengrab',
    'Materialentnahmegrube',
    'Materialraub',
    'Mauer',
    'Mauerfundament',
    'Megalithgrab',
    'Mehrfachbestattung (Grab)',
    'Meilenstein (Entfernungsanzeiger)',
    'Meilerplatz',
    'Mikwe',
    'Militärdepot',
    'Militärlager',
    'Militärposten (Baukomplex)',
    'Mittelständer (Ofen)',
    'Mittelzunge (Ofen)',
    'Motte (Architektur)',
    'Muffelofen',
    'Mundloch (Bergbau)',
    'Mühle (Baukomplex)',
    'Mühlengraben',
    'Mühlenteich',
    'Münzschatzfund',
    'Nachbestattung (Grab)',
    'Natur',
    'Nebenbestattung (Grab)',
    'Nebengebäude',
    'Niederungsburg',
    'Noch nicht ermittelt',
    'Numeruskastell',
    'Oberflächenfund',
    'Ofen (Vorrichtung)',
    'Ofenfeuerung',
    'Ofenkuppel',
    'Ofenrost',
    'Ofenwandung',
    'Ofenzunge',
    'Opferfund',
    'Opfergrube',
    'Opferschacht',
    'Oppidum (Siedlung)',
    'Palas',
    'Palisade',
    'Palisadengraben',
    'Palisadenpfosten',
    'Panzergraben',
    'Panzersperre',
    'Papiermühle',
    'Park',
    'Pavillon (Bauwerk)',
    'Pechhütte',
    'Pechofen',
    'Peristyl',
    'Pfahlgründung',
    'Pfahlrost',
    'Pfalz (Bauwerk)',
    'Pfeifenbäckerei',
    'Pfeiler',
    'Pferch',
    'Pflaster (Oberflächenelement)',
    'Pflugspur',
    'Pfosten',
    'Pfostenbau (Bauwerk)',
    'Pfostengrube',
    'Pfostengrubenreihe',
    'Pfostenreihe',
    'Pfostenspur',
    'Pfostenstickung',
    'Pinge',
    'Pingenfeld',
    'Plagge (Landwirtschaft)',
    'Planierung',
    'Platz (Städtebau)',
    'Pochwerk',
    'Podiumstempel',
    'Porta decumana',
    'Porta praetoria',
    'Porta principalis dextra',
    'Porta principalis sinistra',
    'Portikus',
    'Postamt (Bauwerk)',
    'Praefurnium',
    'Praetorium',
    'Principia (Bauwerk)',
    'Profil (Archäologie)',
    'Pulvermühle',
    'Punktfundament',
    'Quelle (Gewässer)',
    'Quellfassung',
    'Quellheiligtum',
    'Radkasten',
    'Raubgräberloch',
    'Raum (Gebäudeteil)',
    'Ravelin',
    'Refektorium',
    'Reihengräberfeld',
    'Reithalle',
    'Remise',
    'Rennofen',
    'Revisionsschacht',
    'Richtstätte',
    'Ringofen',
    'Ringwall',
    'Risalit',
    'Räucherofen',
    'Römisches Militärlager',
    'Rösche (Bergbau)',
    'Sakralbau',
    'Sammelbecken (Bauwerk)',
    'Sarg',
    'Sarkophag',
    'Schacht (Bauwesen)',
    'Schanze (Befestigungsanlage)',
    'Scherbenhügel',
    'Scherbennest',
    'Scheune',
    'Schicht (Archäologie)',
    'Schießstand',
    'Schiff (Wasserfahrzeug)',
    'Schifffahrtskanal',
    'Schiffswerft',
    'Schlachtfeld',
    'Schlachtplatz (Wirtschaft)',
    'Schlackenhalde',
    'Schlagplatz',
    'Schleiferei (Wirtschaft)',
    'Schleuse (Wasserbau)',
    'Schleusenwehr',
    'Schlitzgrube',
    'Schlitztenne',
    'Schloss (Bauwerk)',
    'Schmelzofen',
    'Schmelzplatz',
    'Schmiede',
    'Schurf',
    'Schwarzerderelikt',
    'Schwellbalkenkonstruktion',
    'Schwelle (Fachwerk)',
    'Schwellenbau',
    'Schwimmbad',
    'Schädelbestattung (Grab)',
    'Schürfgrube (Bauwesen)',
    'Schürmündung',
    'Schützengraben',
    'Schützenloch',
    'Siechenhaus',
    'Siedlung',
    'Skelett (Körperbestandteil)',
    'Sohlgraben',
    'Soldatenfriedhof',
    'Speicher (Bauwerk)',
    'Spinnerei',
    'Spitzgraben',
    'Sportplatz',
    'Stadt (Siedlung)',
    'Stadtbefestigung',
    'Stadtgraben',
    'Stadtkern',
    'Stadtmauer',
    'Stall',
    'Stauanlage',
    'Staudamm',
    'Stauteich',
    'Steg (Bauwerk)',
    'Stehender Ofen',
    'Steinbruch',
    'Steinkeller',
    'Steinkistengrab',
    'Steinkonzentration',
    'Steinkreuz',
    'Steinplattengrab',
    'Steinsetzung',
    'Steinzeugofen',
    'Sternschanze',
    'Stickung',
    'Stiftsgebäude',
    'Straße',
    'Straßendamm',
    'Straßengraben',
    'Straßenposten',
    'Straßensperre',
    'Streifenflur',
    'Streifenhaus',
    'Stufenrain',
    'Ständerwand (Ofen)',
    'Störung (Archäologie)',
    'Stückofen',
    'Suspensura',
    'Synagoge',
    'Sägemühle',
    'Säule',
    'Teich',
    'Tempel',
    'Tepidarium',
    'Terrassenacker',
    'Thermen',
    'Tierbau',
    'Tiergrab',
    'Tonaufbereitungsgrube',
    'Tongrube',
    'Toranlage',
    'Totenbrett',
    'Treppe',
    'Trockenhaus',
    'Trockenmauergrab',
    'Trockenrinne',
    'Trümmerstelle',
    'Turm (Bauwerk)',
    'Turnhalle',
    'Töpferei',
    'Töpferofen',
    'Türschwelle',
    'Uferbefestigung (Bauwerk)',
    'Umgangstempel',
    'Unterstand (Militär)',
    'Urnengrab',
    'Ustrine',
    'Valetudinarium',
    'Verfüllung',
    'Verhüttungsplatz',
    'Verschanzung (Bauwerk)',
    'Via decumana',
    'Via praetoria',
    'Via principalis',
    'Via quintana',
    'Via sagularis',
    'Vicus',
    'Viehtrift',
    'Villa',
    'Villa rustica',
    'Villa urbana',
    'Vorburg',
    'Vorratsgrube',
    'Vorstadt',
    'Wachtturm',
    'Wagengrab',
    'Wagenspur (Verkehr)',
    'Waldweide',
    'Walkmühle',
    'Wall',
    'Wallanlage',
    'Wand (Architektur)',
    'Wandgraben',
    'Waschteich',
    'Wasserburg',
    'Wassergraben',
    'Wasserkunst',
    'Wasserleitung',
    'Wassermühle',
    'Weberei (Betrieb)',
    'Weg (Verkehr)',
    'Wehr (Stauanlage)',
    'Wehrturm',
    'Weiher',
    'Weiler',
    'Werft (Betrieb)',
    'Werkplatz',
    'Werkstatt',
    'Windmühle',
    'Wohnhaus',
    'Wohnstallhaus',
    'Wohnturm',
    'Wurt',
    'Wölbacker',
    'Wüstung',
    'Zangentor',
    'Zaun',
    'Zentralgrab',
    'Ziegelei',
    'Ziegelkonzentration',
    'Ziegelofen',
    'Ziegelplattengrab',
    'Zisterne',
    'Zitadelle',
    'Zollhaus',
    'Zug (Ofen)',
    'Zwinger (Architektur)',
    'Ärmchen (Ofen)',
    'Ölmühle',
    'Übungslager (Militär)',
]

# Datierungsbegriffe (Stand der WNK-Liste, vgl. Liste_WNK-Datierungen.csv)
_DATIERUNGEN_LISTE: List[str] = [
    'Datierung unbekannt',
    'Neuzeit',
    'Römisch',
    'Mittelalter',
    'Eisenzeit',
    'Urgeschichte',
    '13.-14. Jahrhundert n. Chr.',
    'Aalenium',
    'Aeronium',
    'Ahrensburger Kultur',
    'Albium',
    'Alt- bis Jungneolithikum',
    'Alt- bis Mittelneolithikum',
    'Altneolithikum/Bandkeramik',
    'Altpaläolithikum',
    'Anisium',
    'Aptium',
    'Aquitanium',
    'Artinskium',
    'Asselium',
    'Aurignacien',
    'Ausgehende Karolingerzeit',
    'Ausgehende Merowingerzeit',
    'Bajocium',
    'Barremium',
    'Bartonium',
    'Bashkirium',
    'Bathonium',
    'Berriasium',
    'Bischheim',
    'Bronzezeit',
    'Bronzezeit D',
    'Bronzezeit bis Hallstattzeit',
    'Bronzezeit bis Hallstattzeit A',
    'Burdigalium',
    'Calabrium',
    'Callovium',
    'Campanium',
    'Capitanium',
    'Cenomanium',
    'Changhsingium',
    'Chattium',
    'Chibanium',
    'Cisuralium',
    'Coniacium',
    'Danium',
    'Dapingium',
    'Darriwilium',
    'Devon',
    'Drumium',
    'Eifelium',
    'Eisenzeit bis Germanisch',
    'Eisenzeit bis Hochmittelalter',
    'Eisenzeit bis Neuzeit',
    'Emsium',
    'Endneolithikum/Rheinische Becherkultur',
    'Endpaläolithikum',
    'Endpaläolithikum bis Mesolithikum',
    'Eozän',
    'Famennium',
    'Federmessergruppen',
    'Floium',
    'Fortunium',
    'Frasnium',
    'Früh- bis Hochmittelalter',
    'Früh- bis Mittellatènezeit',
    'Früh- bis Spätmittelalter',
    'Frühe Bronzezeit',
    'Frühe Karolingerzeit',
    'Frühe Merowingerzeit',
    'Frühlatènezeit',
    'Frühmittelalter',
    'Frühmittelalter, 5.-6. Jh.',
    'Frühmittelalter, 5.-7. Jh.',
    'Frühmittelalter, 5.-8. Jh.',
    'Frühmittelalter, 6.-8. Jh.',
    'Frühmittelalter, 6.-9. Jh.',
    'Frühmittelalter, 7.-8. Jh.',
    'Frühmittelalter, 7.-9. Jh.',
    'Frühmittelalter, 8.-9. Jh.',
    'Furongium',
    'Gelasium',
    'Geologisch',
    'Germanisch A',
    'Germanisch B',
    'Germanisch B/C2',
    'Germanisch B1',
    'Germanisch B2',
    'Germanisch C',
    'Germanisch C1',
    'Germanisch C2',
    'Germanisch C3',
    'Germanisch C3/D',
    'Germanisch D',
    'Germanisch/einheimisch',
    'Givetium',
    'Glockenbecher',
    'Gorstium',
    'Gravettien',
    'Großgartach',
    'Grönlandium',
    'Guadalupium',
    'Guzhangium',
    'Gzhelium',
    'Hallstatt A',
    'Hallstatt B',
    'Hallstatt B/C',
    'Hallstatt C bis Frühlatène',
    'Hallstatt D bis Frühlatène',
    'Hallstatt D bis Spätlatène',
    'Hauterium',
    'Hettangium',
    'Hirnantium',
    'Hoch- bis Spätmittelalter',
    'Hochmittelalter',
    'Hochmittelalter bis Neuzeit',
    'Hochmittelalter, 10. Jh.',
    'Hochmittelalter, 10.-11. Jh.',
    'Hochmittelalter, 10.-12. Jh.',
    'Hochmittelalter, 11. Jh.',
    'Hochmittelalter, 11.-12. Jh.',
    'Hochmittelalter, 11.-13. Jh.',
    'Hochmittelalter, 12. Jh.',
    'Hochmittelalter, 12.-13. Jh.',
    'Hochmittelalter, 13. Jh.',
    'Holozän',
    'Homerium',
    'Indusium',
    'Jiangshanium',
    'Jung- bis Endneolithikum',
    'Jung- bis Spätneolithikum',
    'Jungneolithikum',
    'Jungneolithikum/Michelsberg',
    'Jungpaläolithikum',
    'Jura (Mesozoikum)',
    'Jüngere Bronzezeit',
    'Jüngere Eisenzeit',
    'Jüngere Hallstattzeit',
    'Kambrium',
    'Kambrium-Serie 2',
    'Kambrium-Stufe 10',
    'Kambrium-Stufe 2',
    'Kambrium-Stufe 3',
    'Kambrium-Stufe 4',
    'Karbon',
    'Karnium',
    'Karolingerzeit',
    'Karolingisch bis Hochmittelalter',
    'Karolingisch bis ottonisch',
    'Kasimovium',
    'Katium',
    'Kimmeridgium',
    'Kreide (Geologie)',
    'Kungurium',
    'Känozoikum',
    'Ladinium',
    'Langhium',
    'Latène D bis frühe Römische Kaiserzeit',
    'Llandovery',
    'Lochkovium',
    'Lopingium',
    'Ludfordium',
    'Ludlow',
    'Lutetium (Eozän)',
    'Maastrichtium',
    'Magdalénien',
    'Meghalayum',
    'Merowingerzeit',
    'Merowingerzeit, 6. Jh.',
    'Merowingerzeit, 6.-7. Jh.',
    'Merowingerzeit, 7. Jh.',
    'Merowingisch bis ottonisch',
    'Mesolithikum',
    'Mesolithikum bis Metallzeit',
    'Mesolithikum bis Neolithikum',
    'Mesozoikum',
    'Messinium',
    'Metallzeit',
    'Metallzeit bis Neuzeit',
    'Miaolingium',
    'Miozän',
    'Mississippium',
    'Mittel- bis Jungneolithikum',
    'Mittel- bis Spätlatènezeit',
    'Mittel- bis Spätneolithikum',
    'Mittelalter bis Frühneuzeit',
    'Mittelalter bis Mitte 19. Jh.',
    'Mittelalter bis Neuzeit',
    'Mitteldevon',
    'Mitteljura',
    'Mittellatènezeit',
    'Mittelmiozän',
    'Mittelneolithikum',
    'Mittelneolithikum bis Metallzeit',
    'Mittelordovizium',
    'Mittelpaläolithikum',
    'Mitteltrias',
    'Mittlere Bronzezeit',
    'Mittlere Karolingerzeit',
    'Moskovium',
    'Neogen',
    'Neolithikum',
    'Neolithikum bis Metallzeit',
    'Neuzeit, 16. Jh.',
    'Neuzeit, 16.-17. Jh.',
    'Neuzeit, 16.-18. Jh.',
    'Neuzeit, 16.-19. Jh.',
    'Neuzeit, 16.-20. Jh.',
    'Neuzeit, 17. Jh.',
    'Neuzeit, 17.-18. Jh.',
    'Neuzeit, 17.-19. Jh.',
    'Neuzeit, 17.-20. Jh.',
    'Neuzeit, 17.-21. Jh.',
    'Neuzeit, 18. Jh.',
    'Neuzeit, 18.-19. Jh.',
    'Neuzeit, 18.-20. Jh.',
    'Neuzeit, 18.-21. Jh.',
    'Neuzeit, 19. Jh.',
    'Neuzeit, 19.-20. Jh.',
    'Neuzeit, 19.-21. Jh.',
    'Neuzeit, 20. Jh.',
    'Neuzeit, 20.-21. Jh.',
    'Neuzeit, 21. Jh.',
    'Nordgrippium',
    'Norium',
    'Oberdevon',
    'Oberjura',
    'Oberkreide',
    'Obermiozän',
    'Oberordovizium',
    'Oberpleistozän',
    'Obertrias',
    'Olenekium',
    'Oligozän',
    'Ordovizium',
    'Ottonisch',
    'Oxfordium',
    'Paibium',
    'Paläogen',
    'Paläolithikum',
    'Paläolithikum bis Mesolithikum',
    'Paläozoikum',
    'Paläozän',
    'Pennsylvanium',
    'Perm',
    'Piacenzium',
    'Planig-Friedberg',
    'Pleistozän',
    'Pliensbachium',
    'Pliozän',
    'Pragium',
    'Priabonium',
    'Pridoli',
    'Quartär',
    'Rhaetium',
    'Rhuddanium',
    'Roadium',
    'Rupelium',
    'Römisch, 1. Jh.',
    'Römisch, 1. Jh. v. Chr.',
    'Römisch, 1. Jh. v. Chr. - 1. Jh.',
    'Römisch, 1. Jh. v. Chr. - 2. Jh.',
    'Römisch, 1. Jh. v. Chr. - 3. Jh.',
    'Römisch, 1. Jh. v. Chr. - 4. Jh.',
    'Römisch, 1.-2. Jh.',
    'Römisch, 1.-3. Jh.',
    'Römisch, 1.-4. Jh.',
    'Römisch, 1.-5. Jh.',
    'Römisch, 2. Jh.',
    'Römisch, 2. Jh. v. Chr.',
    'Römisch, 2.-1. Jh. v. Chr.',
    'Römisch, 2.-3. Jh.',
    'Römisch, 2.-4. Jh.',
    'Römisch, 2.-5. Jh.',
    'Römisch, 3. Jh.',
    'Römisch, 3.-4. Jh.',
    'Römisch, 3.-5. Jh.',
    'Römisch, 4. Jh.',
    'Römisch, 4.-5. Jh.',
    'Römisch, 5. Jh.',
    'Römische Kaiserzeit',
    'Römische Kaiserzeit bis Mittelalter',
    'Römische Kaiserzeit bis Neuzeit',
    'Rössen',
    'Sakmarium',
    'Salisch',
    'Sandbium',
    'Santonium',
    'Schnurkeramik',
    'Seelandium',
    'Serpukhovium',
    'Serravallium',
    'Sheinwoodium',
    'Silur',
    'Sinemurium',
    'Spät- bis Endneolithikum',
    'Späte römische Kaiserzeit bis Frühmittelalter',
    'Spätlatènezeit',
    'Spätmittelalter',
    'Spätmittelalter bis Frühneuzeit, 15.-16. Jh.',
    'Spätmittelalter bis Neuzeit',
    'Spätmittelalter, 14. Jh.',
    'Spätmittelalter, 15. Jh.',
    'Spätneolithikum bis Bronzezeit',
    'Spätneolithikum/Wartberg',
    'Spätpaläolithikum',
    'Staufisch',
    'Steinzeit',
    'Telychium',
    'Terreneuvium',
    'Tertiär',
    'Thanetium',
    'Tithonium',
    'Toarcium',
    'Tortonium',
    'Tournaisium',
    'Tremadocium',
    'Trias',
    'Turonium',
    'Unterdevon',
    'Unterjura',
    'Unterkreide',
    'Untermiozän',
    'Unterordovizium',
    'Untertrias',
    'Urgeschichte bis Neuzeit',
    'Urnenfelderzeit bis Eisenzeit',
    'Urnenfelderzeit bis Frühlatène',
    'Valanginium',
    'Viseum',
    'Wenlock',
    'Wordium',
    'Wuchiapingium',
    'Wuliuum',
    'Ypresium',
    'Zancleum',
    'Ältere Bronzezeit',
    'Ältere Eisenzeit',
    'Ältere Hallstattzeit',
]


def datierung_zu_jahre(dat: str) -> Tuple[Optional[int], Optional[int]]:
    if not dat: return None, None
    key = dat.strip().lower()
    # Sentinel: "Datierung unbekannt" ist keine echte Datierung — nie Jahre
    # liefern (der WNK-Wert -542000000 würde die Strat-Propagierung vergiften)
    if key == "datierung unbekannt": return None, None
    if key in _WNK_LOWER: return _WNK_LOWER[key]
    # Zahlen-Fallback. 3-stellige Zahlen NUR wenn der ganze Text eine
    # Jahresangabe ist (sonst würde "Pos. 189" zum Jahr 189);
    # 4-stellige Zahlen sind im Datierungsfeld praktisch immer Jahre.
    txt = dat.strip()
    jahresangabe = re.fullmatch(
        r'(?:um\s+|ca\.?\s*)?(-?\d{3,4})'
        r'(?:\s*(?:-|–|/|bis)\s*(-?\d{3,4}))?'
        r'(?:\s*(?:n|v)\.?\s*Chr\.?)?\.?', txt, re.IGNORECASE)
    if jahresangabe:
        # Strukturiert lesen — findall hätte bei "1650-1700" das zweite
        # Jahr als -1700 gedeutet
        j = [int(jahresangabe.group(1))]
        if jahresangabe.group(2):
            j.append(int(jahresangabe.group(2)))
    else:
        j = [int(x) for x in re.findall(r'(?<![\d.])\d{4}(?!\d)', txt)]
    if j:
        if re.search(r'v\.?\s*Chr', txt, re.IGNORECASE):
            j = [-abs(x) for x in j]
        return min(j), max(j)
    return None, None

def _sort_key(nr: str) -> tuple:
    m = re.match(r'^(\d+)', str(nr))
    return (int(m.group(1)) if m else 0, str(nr))

def norm_stellennr(nr) -> str:
    """Stellennummern vereinheitlichen, damit Fundliste, Katalog und Editor
    dieselbe Stelle treffen: '004' → '4', '4.0' (Excel-Export) → '4',
    Leerzeichen entfernt. Alphanumerische Nummern ('100a') bleiben unverändert."""
    t = str(nr).strip()
    if re.fullmatch(r'\d+', t):
        return str(int(t))
    if re.fullmatch(r'\d+[.,]0+', t):
        return str(int(re.split(r'[.,]', t)[0]))
    return t

# ═══════════════════════════════════════════════════════════════
# PROFILTEXT-PARSER
# ═══════════════════════════════════════════════════════════════

_PFX = r'(?:St\.?\s*|Stelle\s*)?'
_NR  = r'\d+[a-zA-Z]?(?:-[a-zA-Z0-9]+)?'

# Matcht Einzel- UND Listen-Stellennummern:
#   "30"  |  "St. 30"  |  "St. 30, 47 und 31"  |  "St. 8/20 u. 30"
_S = (
    _PFX + r'('
    + _NR
    + r'(?:\s*[,/]\s*' + _PFX + _NR + r')*'   # Komma ODER Schrägstrich
    + r'(?:\s+(?:und|u\.|sowie)\s+' + _PFX + _NR + r')?'
    + r')'
)

# Himmelsrichtungen (deutsch, kurz und ausgeschrieben)
_HIM = (r'(?:N|S|O|W|NO|NW|SO|SW|'
        r'Nord|Süd|Ost|West|'
        r'Nordost|Nordwest|Südost|Südwest)')

# Muster: diese Stelle ist JÜNGER als die genannten
# (aktiv: schneidet, liegt über, eingetieft in …)
_JUENGER = [
    rf'schneidet\s+{_S}',
    rf'überlagert\s+(?!von\b){_S}',
    rf'überprägt\s+(?!von\b){_S}',
    rf'liegt\s+(?:direkt\s+)?auf\s+{_S}',
    rf'auf\s+{_S}\s+aufliegend',
    rf'über\s+{_S}(?=\s|,|;|$)',
    rf'oberhalb\s+(?:v(?:on|\.)?\s+)?{_S}',
    rf'oberh\.\s+(?:v\.?\s+)?{_S}',
    rf'eingetieft\s+in\s+{_S}',
    rf'einget\.\s+in\s+{_S}',
    rf'tieft?\s+in\s+{_S}\s*ein',
    rf'in\s+{_HIM}\s+schneidet\s+{_S}',
    rf'im\s+{_HIM}\s+schneidet\s+{_S}',
    rf'stört\s+{_S}',
    # Bewusst KEIN 'st\.'-Kurzmuster: "St." ist fast immer das Stellen-Präfix
    # ("liegt unter St. 5" würde sonst zusätzlich als JÜNGER fehlgedeutet).
    # Mehrdeutige Kurzformen werden nicht geparst — nur Eindeutiges (Nutzerentscheid).
]

# Muster: diese Stelle ist ÄLTER als die genannten
# (passiv: wird geschnitten, liegt unter, wird überlagert …)
_AELTER = [
    rf'(?:wird\s+)?(?:gest\.|gestört)\s+von\s+{_S}',  # "wird gest. von St. 8/20 u. 30"
    rf'geschnitten\s+von\s+{_S}',
    rf'geschn\.\s+(?:von\s+)?{_S}',
    rf'gestört\s+von\s+{_S}',
    rf'gest\.\s+(?:von\s+)?{_S}',
    rf'überprägt\s+von\s+{_S}',
    rf'überlagert\s+von\s+{_S}',
    rf'von\s+{_S}\s+(?:geschn|gest)\.',
    rf'von\s+{_S}\s+(?:geschnitten|gestört|überlagert)',
    rf'unter\s+{_S}',
    rf'unterhalb\s+(?:v(?:on|\.)?\s+)?{_S}',
    rf'unterh\.\s+(?:v(?:on|\.)?\s+)?{_S}',
    rf'in\s+unterhalb\s+{_S}',
    # mit Himmelsrichtung – Passiv-Konstruktionen → diese Stelle ist älter
    rf'in\s+{_HIM}\s+von\s+{_S}\s+(?:geschnitten|gestört|überlagert|geschn\.|gest\.)',
    rf'im\s+{_HIM}\s+von\s+{_S}\s+(?:geschnitten|gestört|überlagert|geschn\.|gest\.)',
    rf'in\s+{_HIM}\s+(?:von\s+){_S}\s+(?:geschnitten|gestört|überlagert)',
    rf'in\s+{_HIM}\s+von\s+{_S}(?=\s*,)',   # "in N von St. 10," → Passiv folgt
    rf'im\s+{_HIM}\s+von\s+{_S}(?=\s*[,.])',
]

# Muster: gleichzeitig / Grenze unklar / fließender Übergang
_GLEICH = [
    rf'grenzt\s+(?:sich\s+)?(?:nur\s+)?(?:schwach\s+)?an\s+{_S}',
    rf'angrenzend\s+an\s+{_S}',
    rf'neben\s+{_S}',
    rf'undeutliche\s+[Gg]renzen?\s+zu\s+{_S}',
    rf'[Gg]renzen?\s+zu\s+{_S}\s+(?:verwaschen|unklar)',
    rf'[Gg]renze\s+zu\s+{_S}\s+(?:verwaschen|unklar)',
    rf'fließender?\s+[Üü]bergang\s+zu\s+{_S}',
    rf'geht\s+(?:{_HIM}\s+)?fließend\s+in\s+{_S}\s+über',
    rf'[Gg]renze\s+(?:in\s+{_HIM}\s+)?zu\s+{_S}\s+unklar',
    rf'[Aa]bgr\.?\s+zu\s+{_S}.*?undeutl?',
    rf'[Aa]bgrenzung\s+zu\s+{_S}.*?undeutl?',
    rf'grenzt\s+sich\s+(?:nur\s+)?schwach\s+von\s+{_S}\s+ab',
    rf'[Üü]bergang\s+(?:zu|in)\s+{_S}\s+(?:fließend|unklar)',
]

def _suche_nummern(muster_liste, text: str) -> List[str]:
    gefunden = []
    for muster in muster_liste:
        for treffer in re.finditer(muster, text, re.IGNORECASE):
            # Listenausdrücke ("30, 47 und 31", "8/20 u. 30") in einzelne
            # Stellennummern zerlegen — unzerteilt würden sie nie eine
            # existierende Stelle treffen und still verworfen
            for nr in re.findall(_NR, treffer.group(1)):
                nr = norm_stellennr(nr)
                if nr and nr not in gefunden:
                    gefunden.append(nr)
    return gefunden

def parse_profil_beziehungen(profil: str) -> Tuple[List[str], List[str], List[str]]:
    """
    Extrahiert stratigraphische Beziehungen aus Profilbeschreibungstext.
    Gibt zurück: (juenger_als, aelter_als, gleich_alt_wie).
    """
    if not profil:
        return [], [], []
    j = _suche_nummern(_JUENGER, profil)
    a = _suche_nummern(_AELTER, profil)
    g = _suche_nummern(_GLEICH, profil)
    log.debug(f"Profil-Parser: ↑{j}  ↓{a}  ={g}  |  Text: {profil[:60]}")
    return j, a, g

# ═══════════════════════════════════════════════════════════════
# DATENMODELLE
# ═══════════════════════════════════════════════════════════════

@dataclass
class Fund:
    stellennr: str = ""
    posnr: str = ""
    unternr: str = ""
    aktivitaet: str = ""
    material: str = ""
    ansprache: str = ""
    kommentar: str = ""
    datierung: str = ""
    von: Optional[int] = None
    bis: Optional[int] = None
    ansprache_von: str = ""
    datum: str = ""

    def kurztext(self) -> str:
        t = []
        if self.material:  t.append(self.material)
        if self.ansprache: t.append(self.ansprache)
        if self.datierung: t.append(f"[{self.datierung}]")
        return ", ".join(t) or "(kein Eintrag)"

    def ist_modern(self) -> bool:
        p = f"{self.material} {self.ansprache} {self.datierung}".lower()
        return any(b in p for b in MODERN_BEGRIFFE)


@dataclass
class Ankerpunkt:
    """Ein Ankerpunkt pro Art (tpq|taq|exakt): Ereignis + Jahr.
    Bewusst auf die Editor-Felder reduziert — mehr kann die Eingabemaske
    nicht darstellen, und für die Rechnung zählt ohnehin nur der jüngste
    TPQ bzw. älteste TAQ (Nutzerentscheid Option a, 2026-06-11)."""
    art: str = "exakt"          # tpq | taq | exakt
    beschreibung: str = ""
    von: Optional[int] = None
    praezision: str = "Jahr"

    def als_text(self) -> str:
        p = [self.beschreibung] if self.beschreibung else []
        if self.von is not None:
            p.append(f"({self.von} ≈{self.praezision})")
        return " ".join(p)


@dataclass
class Stelle:
    stellennr: str = ""
    befansprac: str = ""
    datierung_katalog: str = ""
    kein_befund: bool = False
    rest_befund: bool = False
    moderne_stoerung: bool = False
    durchschossen: bool = False
    sichtbar: str = ""          # sehr gut|gut|mäßig|schlecht
    zeichnung: str = ""
    kommentar: str = ""
    planum: str = ""
    profil: str = ""
    bearbeiter: str = ""
    tiefe_ok: Optional[float] = None
    tiefe_uk: Optional[float] = None
    box_farbe:     str  = ""           # Hex-Farbe für Harris-Box
    box_form:      str  = "rechteck"   # rechteck | dreieck | ellipse
    box_geschuetzt: bool = False
    box_x: Optional[int] = None   # manuelle X-Position in Harris-Grafik
    box_y: Optional[int] = None   # manuelle Y-Position
    strat_von: Optional[int] = None    # strat. Untergrenze (ephemer)
    strat_bis: Optional[int] = None    # strat. Obergrenze  (ephemer)
    dat_protokoll: str = ""            # Datierungsprotokoll
    dat_juenger_als:   List[str] = field(default_factory=list)  # aus Datierung
    dat_aelter_als:    List[str] = field(default_factory=list)  # aus Datierung
    dat_suppressed:    bool = False  # True = Nutzer hat alle Linien gelöscht → keine Auto-Dat
    # Vom Nutzer einzeln gelöschte Auto-Beziehungen ("juenger:5", "aelter:8",
    # "gleich:12") – werden nach jeder Neuberechnung wieder aussortiert,
    # damit Gelöschtes gelöscht bleibt, ohne es zu "manuell" zu befördern
    unterdrueckte_bez: List[str] = field(default_factory=list)
    # Bidirektional abgeleitet (ephemer, nicht gespeichert, jederzeit neuberechnet)
    _bidir_juenger_als: List[str] = field(default_factory=list, repr=False)
    _bidir_aelter_als:  List[str] = field(default_factory=list, repr=False)
    _bidir_gleich:      List[str] = field(default_factory=list, repr=False)

    # Harris-Beziehungen  (manuell – höchste Priorität)
    juenger_als:   List[str] = field(default_factory=list)
    aelter_als:    List[str] = field(default_factory=list)
    gleich_alt_wie: List[str] = field(default_factory=list)
    gestoert_von:  List[str] = field(default_factory=list)
    # Aus Profiltext geparst (niedrigere Priorität als manuell)
    profil_juenger_als:  List[str] = field(default_factory=list)
    profil_aelter_als:   List[str] = field(default_factory=list)
    profil_gleich:       List[str] = field(default_factory=list)
    # Aus OK/UK abgeleitet (niedrigste Priorität)
    okuk_juenger_als: List[str] = field(default_factory=list)
    okuk_aelter_als:  List[str] = field(default_factory=list)

    # Absolute Ankerpunkte
    ankerpunkte: List[Ankerpunkt] = field(default_factory=list)

    # Manuelle Datierungsüberschreibung
    manuell_datierung: str = ""
    manuell_von: Optional[int] = None
    manuell_bis: Optional[int] = None

    # Laufzeit (nicht persistiert)
    auto_von: Optional[int] = None
    auto_bis: Optional[int] = None
    auto_datierung_str: str = ""
    inferred_von: Optional[int] = None
    inferred_bis: Optional[int] = None
    inferred_datierung_str: str = ""
    inferred_quelle: str = ""
    funde: List[Fund] = field(default_factory=list)

    # ── kombinierte Beziehungslisten (manuell hat Vorrang) ────
    def alle_juenger_als(self) -> List[str]:
        kombi = list(self.juenger_als)
        for n in (self.profil_juenger_als + self.okuk_juenger_als
                  + self.dat_juenger_als + self._bidir_juenger_als):
            if n not in kombi: kombi.append(n)
        return kombi

    def alle_aelter_als(self) -> List[str]:
        kombi = list(self.aelter_als)
        for n in (self.profil_aelter_als + self.okuk_aelter_als
                  + self.dat_aelter_als + self._bidir_aelter_als):
            if n not in kombi: kombi.append(n)
        return kombi

    def alle_gleich_alt_wie(self) -> List[str]:
        kombi = list(self.gleich_alt_wie)
        for n in (self.profil_gleich + self._bidir_gleich):
            if n not in kombi: kombi.append(n)
        return kombi

    # ── effektive Datierung ───────────────────────────────────
    def _raw_von(self) -> Optional[int]:
        if self.manuell_von is not None: return self.manuell_von
        tpq = [a.von for a in self.ankerpunkte
               if a.art=="tpq" and a.von is not None]
        return max(tpq) if tpq else self.auto_von

    def _raw_bis(self) -> Optional[int]:
        if self.manuell_bis is not None: return self.manuell_bis
        taq = [a.von for a in self.ankerpunkte
               if a.art=="taq" and a.von is not None]
        return min(taq) if taq else self.auto_bis

    def effektiv_von(self) -> Optional[int]:
        if self.manuell_von is not None: return self.manuell_von
        tpq = [a.von for a in self.ankerpunkte
               if a.art == "tpq" and a.von is not None]
        base = max(tpq) if tpq else self.auto_von
        if base is not None and self.strat_von is not None:
            return max(base, self.strat_von)
        if base is not None: return base
        return self.strat_von if self.strat_von is not None else self.inferred_von

    def effektiv_bis(self) -> Optional[int]:
        if self.manuell_bis is not None: return self.manuell_bis
        taq = [a.von for a in self.ankerpunkte
               if a.art == "taq" and a.von is not None]
        base = min(taq) if taq else self.auto_bis
        if base is not None and self.strat_bis is not None:
            return min(base, self.strat_bis)
        if base is not None: return base
        return self.strat_bis if self.strat_bis is not None else self.inferred_bis

    def datierung_quelle(self) -> str:
        if (self.manuell_datierung
                or self.manuell_von is not None
                or self.manuell_bis is not None):
            return "Manuell"
        if self.ankerpunkte:
            return "Ankerpunkt"
        if self.auto_datierung_str:
            return "Fund"
        if self.inferred_datierung_str:
            return "Stratigraphie"
        if self.datierung_katalog:
            return "Katalog"
        return "–"

    def effektiv_datierung_str(self) -> str:
        if self.manuell_datierung: return self.manuell_datierung
        # Nur Jahreszahlen ohne WNK-Begriff → trotzdem anzeigen
        if self.manuell_von is not None or self.manuell_bis is not None:
            v = str(self.manuell_von) if self.manuell_von is not None else "?"
            b = str(self.manuell_bis) if self.manuell_bis is not None else "?"
            return v if self.manuell_von == self.manuell_bis else f"{v}–{b}"
        if self.auto_datierung_str: return self.auto_datierung_str
        if self.inferred_datierung_str:
            # Nur den WNK-Begriff anzeigen – kein ≤/≥ Präfix, keine Quellenangabe
            t = re.sub(r"^[\u2264\u2265<>]+\s*", "", self.inferred_datierung_str)
            t = re.sub(r"\s*\(wg\..*?\)\s*$", "", t).strip()
            return t or self.inferred_datierung_str
        if self.datierung_katalog: return self.datierung_katalog
        return "Datierung unbekannt"

    def harris_info_kurztext(self) -> str:
        """Kurzer Text für Harris-Matrix-Box: Datierungsherkunft."""
        q = self.datierung_quelle()
        if q == "Fund" and self.funde:
            bester = min((f for f in self.funde if f.von is not None),
                         key=lambda f: f.von, default=None)
            if bester:
                pos = f"St.{bester.stellennr}/{bester.posnr}"
                mat = (bester.material or bester.ansprache)[:20]
                return f"Fund {pos}: {mat}"
        if q == "Stratigraphie":
            return self.inferred_datierung_str[:40]
        if q == "Manuell":
            return "Manuell datiert"
        if q == "Ankerpunkt":
            ap = self.ankerpunkte[0] if self.ankerpunkte else None
            return ap.als_text()[:40] if ap else "Ankerpunkt"
        if self.tiefe_ok is not None and q == "–":
            return f"OK: {self.tiefe_ok:.2f} m NHN"
        return ""


@dataclass
class Warnung:
    stufe: int       # 1=Widerspruch  2=Hinweis  3=Notiz
    stelle: str
    nachricht: str
    STUFEN = {1:"⛔ Widerspruch", 2:"⚠️  Hinweis", 3:"ℹ️  Notiz"}
    FARBEN = {1:"#ffcccc", 2:"#fff3cd", 3:"#d4edda"}
    def stufe_str(self) -> str: return self.STUFEN.get(self.stufe,"?")
    def farbe(self)     -> str: return self.FARBEN.get(self.stufe,"white")

# ═══════════════════════════════════════════════════════════════
# CSV-PARSER
# ═══════════════════════════════════════════════════════════════

def _lese_csv(pfad: Path) -> List[Dict]:
    # utf-8-sig zuerst und STRIKT: echte UTF-8-Dateien (auch reines ASCII) werden
    # erkannt; windows-1252 mit Umlauten ist kein gültiges UTF-8 und fällt durch.
    # errors="replace" nur beim letzten Fallback, sonst greift der Fallback nie.
    for enc, fehler in [("utf-8-sig", "strict"), (ENCODING, "strict"),
                        ("latin-1", "replace")]:
        try:
            with open(pfad, encoding=enc, errors=fehler, newline="") as f:
                reader = csv.DictReader(f, delimiter=DELIMITER)
                zeilen = [{k.strip().lstrip('\ufeff'): (v or "").strip()
                           for k,v in z.items()} for z in reader]
            log.info(f"CSV ({enc}): {pfad.name}  {len(zeilen)} Zeilen")
            return zeilen
        except Exception as e:
            log.debug(f"Encoding {enc} für {pfad.name}: {e}")
    log.error(f"Konnte {pfad} nicht lesen"); return []


# Spalten-Signaturen zur Auto-Erkennung des CSV-Typs (Drag&Drop ins selbe Feld).
# Die Mengen sind ÜBERSCHNEIDUNGSFREI → eine echte Fundliste/­Katalog-Datei
# trifft nur jeweils eine davon.
_FUNDLISTE_SIG = {"posnr", "unternr", "material"}
_KATALOG_SIG   = {"befansprac", "profil", "planum", "tiefeok", "tiefeuk",
                  "juengerals", "aelterals", "keinbefund"}

def _csv_spalten(pfad: Path) -> set:
    """Header-Spaltennamen einer CSV, normalisiert (klein, ohne Leerzeichen/BOM)."""
    for enc, fehler in [("utf-8-sig", "strict"), (ENCODING, "strict"),
                        ("latin-1", "replace")]:
        try:
            with open(pfad, encoding=enc, errors=fehler, newline="") as f:
                for row in csv.reader(f, delimiter=DELIMITER):
                    return {c.strip().lstrip("﻿").lower().replace(" ", "")
                            for c in row}
            return set()
        except Exception:
            continue
    return set()

def erkenne_csv_typ(pfad: Path) -> Optional[str]:
    """Erkennt anhand der Spaltennamen, ob eine CSV eine Fundliste oder ein
    Stellenkatalog ist. Rückgabe: 'fundliste' | 'stellenkatalog' | None."""
    cols = _csv_spalten(pfad)
    if not cols:
        return None
    sf = len(cols & _FUNDLISTE_SIG)
    sk = len(cols & _KATALOG_SIG)
    if sk > sf:  return "stellenkatalog"
    if sf > sk:  return "fundliste"
    return None   # nicht eindeutig (keine oder gleich viele Signaturspalten)

def parse_liste(text: str) -> List[str]:
    if not text: return []
    return [norm_stellennr(x) for x in str(text).replace(";",",").split(",")
            if x.strip()]

def _float_oder_none(s: str) -> Optional[float]:
    try: return float(s.replace(",",".")) if s else None
    except: return None

def parse_fundliste(pfad: Path) -> List[Fund]:
    funde = []
    # Neu aufbauen: korrigierte Begriffe sollen den Warn-Button auch
    # wirklich verschwinden lassen
    _UNBEKANNT_FUND.clear()
    for i, z in enumerate(_lese_csv(pfad)):
        try:
            dat = z.get("Datierung","")
            _dat_norm = dat.strip()
            if (_dat_norm and
                    _dat_norm.lower() not in ("datierung unbekannt", "") and
                    _dat_norm not in WNK_DATIERUNGEN_EINGEBETTET):
                _UNBEKANNT_FUND.add(_dat_norm)
            von, bis = datierung_zu_jahre(dat)
            f = Fund(
                aktivitaet=z.get("Aktivitaet",""),
                stellennr=norm_stellennr(z.get("Stellennr","")),
                posnr=z.get("Posnr",""), unternr=z.get("Unternr",""),
                material=z.get("Material",""), ansprache=z.get("Ansprache",""),
                kommentar=z.get("Kommentar",""), datierung=dat,
                von=von, bis=bis,
                ansprache_von=z.get("AnspracheVon",""), datum=z.get("Datum",""),
            )
            if f.stellennr: funde.append(f)
        except Exception as e: log.warning(f"Fundliste Z.{i+2}: {e}")
    log.info(f"Fundliste: {len(funde)} Funde  ({pfad.name})")
    return funde

def parse_stellenkatalog(pfad: Path) -> List[Stelle]:
    stellen = []
    _UNBEKANNT_KATALOG.clear()
    for i, z in enumerate(_lese_csv(pfad)):
        try:
            nr = norm_stellennr(z.get("StellenNr",""))
            if not nr: continue
            ansp = z.get("BefAnsprac","").strip()
            ansp_low = ansp.lower()

            # Flags automatisch setzen
            kein_befund = (z.get("KeinBefund","").upper() == "X" or
                           ansp_low in ARBEITSBEREICH_ANSPRACHEN)
            stoerung    = ansp_low in STOERUNG_ANSPRACHEN

            # Profilbeziehungen parsen
            profil_txt = z.get("Profil","")
            pj, pa, pg = parse_profil_beziehungen(profil_txt)
            # Planum ebenfalls nach Harris-Beziehungen durchsuchen
            planum_txt = z.get("Planum","")
            if planum_txt and planum_txt.strip().lower() != "entfällt":
                pj2, pa2, pg2 = parse_profil_beziehungen(planum_txt)
                for x in pj2:
                    if x not in pj: pj.append(x)
                for x in pa2:
                    if x not in pa: pa.append(x)
                for x in pg2:
                    if x not in pg: pg.append(x)
            # Selbstreferenzen entfernen (z.B. "eingetieft in 108" in Stelle 108)
            pj = [x for x in pj if x != nr]
            pa = [x for x in pa if x != nr]
            pg = [x for x in pg if x != nr]

            _kat_dat = z.get("Datierung","").strip()
            if (_kat_dat and
                    _kat_dat.lower() not in ("datierung unbekannt", "") and
                    _kat_dat not in WNK_DATIERUNGEN_EINGEBETTET):
                _UNBEKANNT_KATALOG.add(_kat_dat)
            s = Stelle(
                stellennr=nr, befansprac=ansp,
                datierung_katalog=_kat_dat,
                kein_befund=kein_befund, rest_befund=z.get("RestBefund","").upper()=="X",
                moderne_stoerung=stoerung,
                sichtbar=z.get("Sichtbar","").strip().lower(),
                zeichnung=z.get("Zeichnung",""), kommentar=z.get("Kommentar",""),
                planum=z.get("Planum",""), profil=profil_txt,
                bearbeiter=z.get("Bearbeiter",""),
                tiefe_ok=_float_oder_none(z.get("TiefeOK","")),
                tiefe_uk=_float_oder_none(z.get("TiefeUK","")),
                juenger_als=parse_liste(z.get("JuengerAls","")),
                aelter_als=parse_liste(z.get("AelterAls","")),
                profil_juenger_als=pj, profil_aelter_als=pa, profil_gleich=pg,
            )
            stellen.append(s)
        except Exception as e: log.warning(f"Stellenkatalog Z.{i+2}: {e}")
    log.info(f"Stellenkatalog: {len(stellen)} Stellen  ({pfad.name})")
    return stellen

# ═══════════════════════════════════════════════════════════════
# ANALYSE-ENGINE
# ═══════════════════════════════════════════════════════════════

def auto_datiere_stellen(stellen: Dict[str, Stelle], funde: List[Fund]):
    """Funde zu Stellen zuordnen, Datierungsspannen berechnen."""
    for s in stellen.values():
        s.funde=[]; s.auto_von=None; s.auto_bis=None; s.auto_datierung_str=""

    for f in funde:
        if not f.stellennr: continue
        if f.stellennr not in stellen:
            stellen[f.stellennr] = Stelle(stellennr=f.stellennr)
        stellen[f.stellennr].funde.append(f)

    # Stellenkatalog-Datierung → auto_von/auto_bis (Fallback)
    # "Datierung unbekannt" liefert einen Sentinel (-542000000) der keine echte
    # Datierung darstellt und die Strat-Propagierung korrumpieren würde → überspringen.
    # Arbeitsflächen (kein_befund) sind datierungsirrelevant → ebenfalls überspringen.
    for s in stellen.values():
        if s.kein_befund: continue
        if s.auto_von is None and s.datierung_katalog:
            if s.datierung_katalog != 'Datierung unbekannt':
                s.auto_von, s.auto_bis = datierung_zu_jahre(s.datierung_katalog)

    for nr, s in stellen.items():
        if not s.funde: continue
        if s.kein_befund: continue  # Arbeitsflächen: Funde nicht für Datierung auswerten
        vw = [f.von for f in s.funde if f.von is not None]
        bw = [f.bis for f in s.funde if f.bis is not None]
        s.auto_von = min(vw) if vw else None
        s.auto_bis = max(bw) if bw else None
        dats = list(dict.fromkeys(
            f.datierung for f in s.funde
            if f.datierung and f.datierung.lower() not in ("datierung unbekannt","")
        ))
        s.auto_datierung_str = ", ".join(dats)
        log.debug(f"Stelle {nr}: {len(s.funde)} Funde → '{s.auto_datierung_str}' [{s.auto_von}…{s.auto_bis}]")

    # Störungen: immer Neuzeit. NACH der Fundauswertung, damit altes Material
    # (Residualfunde) die Störung nicht doch datiert — die Funde bleiben zur
    # Dokumentation an der Stelle, die Residual-Warnung kommt separat.
    for s in stellen.values():
        if s.kein_befund: continue
        if s.moderne_stoerung and not s.manuell_datierung:
            nz_von, nz_bis = datierung_zu_jahre("Neuzeit")
            war_alt = s.auto_von is not None and s.auto_von < 1500
            if s.auto_von is None or s.auto_von < 1500:
                s.auto_von = nz_von or 1500
            if s.auto_bis is None or s.auto_bis < s.auto_von:
                s.auto_bis = nz_bis or 2100
            if not s.auto_datierung_str or war_alt:
                s.auto_datierung_str = "Neuzeit"


def propagiere_beziehungen(stellen: Dict[str, Stelle]):
    """
    Bidirektionale Beziehungen vollständig neu aufbauen (_bidir_* Felder).
    Manuell eingetragene juenger_als/aelter_als bleiben UNBERÜHRT.
    Löschen eines Eintrags wirkt sofort, weil _bidir_* komplett neugebaut wird.
    Manuelles Wiederhinzufügen überschreibt jederzeit.
    """
    # _bidir_* vollständig leeren und neu berechnen
    for s in stellen.values():
        s._bidir_juenger_als = []
        s._bidir_aelter_als  = []
        s._bidir_gleich      = []

    for nr, s in stellen.items():
        # Alle Quellen die "nr jünger als xnr" bedeuten
        for xnr in (s.juenger_als + s.profil_juenger_als
                    + s.okuk_juenger_als + s.dat_juenger_als):
            if (xnr in stellen and nr not in stellen[xnr]._bidir_aelter_als
                    and f"aelter:{nr}" not in stellen[xnr].unterdrueckte_bez):
                stellen[xnr]._bidir_aelter_als.append(nr)
        # Alle Quellen die "nr älter als ynr" bedeuten
        for ynr in (s.aelter_als + s.profil_aelter_als
                    + s.okuk_aelter_als + s.dat_aelter_als):
            if (ynr in stellen and nr not in stellen[ynr]._bidir_juenger_als
                    and f"juenger:{nr}" not in stellen[ynr].unterdrueckte_bez):
                stellen[ynr]._bidir_juenger_als.append(nr)
        # gleich_alt_wie: symmetrisch über das EPHEMERE _bidir_gleich spiegeln.
        # Früher wurde in die persistente Liste der Gegenstelle geschrieben —
        # dadurch überlebte die Beziehung jede Löschung im Editor und wurde
        # als scheinbar manueller Eintrag mit abgespeichert.
        for gnr in (s.gleich_alt_wie + s.profil_gleich):
            if gnr in stellen and gnr != nr:
                g = stellen[gnr]
                if (nr not in g.gleich_alt_wie and nr not in g._bidir_gleich
                        and f"gleich:{nr}" not in g.unterdrueckte_bez):
                    g._bidir_gleich.append(nr)
        # gestoert_von: Störung jünger als Gestörtes
        for xnr in s.gestoert_von:
            if (xnr in stellen and nr not in stellen[xnr]._bidir_aelter_als
                    and f"aelter:{nr}" not in stellen[xnr].unterdrueckte_bez):
                stellen[xnr]._bidir_aelter_als.append(nr)


def propagiere_datierungen(stellen: Dict[str, Stelle]):
    """
    Leitet Datierungshinweise über Harris-Beziehungen weiter (iterativ).
    Priorität: Manuell > Ankerpunkt > Fund > Stratigraphie > OK/UK
    Schreibt nur in inferred_*, nie in auto_* oder manuell_*.
    """
    for s in stellen.values():
        s.inferred_von=None; s.inferred_bis=None
        s.inferred_datierung_str=""; s.inferred_quelle=""

    changed=True; it=0
    while changed and it < 40:
        changed=False; it+=1
        for nr, s in stellen.items():
            # s ist JÜNGER als xnr  →  xnr ist ÄLTER
            # xnr.bis wird durch s.von nach oben begrenzt
            for xnr in s.alle_juenger_als():
                if xnr not in stellen: continue
                x = stellen[xnr]
                # Manuelle Beziehungen nicht für Datierungsinferenz verwenden
                # (Gegenstück zum Guard in der aelter_als-Schleife unten)
                if xnr in s.juenger_als or nr in x.aelter_als:
                    continue
                s_von = s.effektiv_von()
                if s_von is not None:
                    if x.inferred_bis is None or s_von < x.inferred_bis:
                        x.inferred_bis = s_von; changed=True
                # Nur solide Datierungen propagieren (verhindert ≤≤≤-Kaskaden)
                s_dat_solid = (s.manuell_datierung or s.auto_datierung_str
                               or s.datierung_katalog or "")
                if s_dat_solid not in ("Datierung unbekannt","") \
                        and not x.auto_datierung_str and not x.manuell_datierung:
                    ns = f"≤ {s_dat_solid}  (wg. Stelle {nr})"
                    if x.inferred_datierung_str != ns:
                        x.inferred_datierung_str=ns; x.inferred_quelle=nr; changed=True

            # s ist ÄLTER als ynr  →  ynr ist JÜNGER
            for ynr in s.alle_aelter_als():
                if ynr not in stellen: continue
                y = stellen[ynr]
                # Manuelle Beziehungen nicht für Datierungsinferenz verwenden
                if ynr in s.aelter_als or nr in stellen[ynr].juenger_als:
                    continue
                s_bis = s.effektiv_bis()
                if s_bis is not None:
                    # Nicht propagieren wenn es y's eigenes Datum widerspricht
                    y_raw_bis = y._raw_bis()
                    if y_raw_bis is None or s_bis <= y_raw_bis:
                        if y.inferred_von is None or s_bis > y.inferred_von:
                            y.inferred_von = s_bis; changed=True
                s_dat_solid = (s.manuell_datierung or s.auto_datierung_str
                               or s.datierung_katalog or "")
                if s_dat_solid not in ("Datierung unbekannt","") \
                        and not y.auto_datierung_str and not y.manuell_datierung:
                    ns = f"≥ {s_dat_solid}  (wg. Stelle {nr})"
                    if y.inferred_datierung_str != ns:
                        y.inferred_datierung_str=ns; y.inferred_quelle=nr; changed=True


def okuk_beziehungen_berechnen(stellen: Dict[str, Stelle],
                                auswahl: Optional[List[str]] = None) -> int:
    """Vorläufige Harris-Beziehungen aus TiefeOK ableiten."""
    kand = {nr: s for nr, s in stellen.items()
            if s.tiefe_ok is not None and not s.kein_befund and not s.moderne_stoerung
            and (auswahl is None or nr in auswahl)}
    if len(kand) < 2: return 0
    for nr in kand: stellen[nr].okuk_juenger_als=[]; stellen[nr].okuk_aelter_als=[]
    sortiert = sorted(kand, key=lambda n: stellen[n].tiefe_ok, reverse=True)
    neu=0; schwelle=0.05
    for i in range(len(sortiert)-1):
        nr_j = sortiert[i]; nr_a = sortiert[i+1]
        diff = stellen[nr_j].tiefe_ok - stellen[nr_a].tiefe_ok
        if diff > schwelle and nr_a not in stellen[nr_j].juenger_als:
            stellen[nr_j].okuk_juenger_als.append(nr_a)
            stellen[nr_a].okuk_aelter_als.append(nr_j)
            neu += 1
    log.info(f"OK/UK: {neu} neue Beziehungen aus {len(kand)} Stellen")
    return neu


def _gleich_komponenten(stellen: Dict[str, Stelle]) -> Dict[str, str]:
    """Union-Find über alle 'gleich alt wie'-Kanten → jede Stelle bekommt einen
    Komponenten-Repräsentanten. Stellen derselben Komponente sind (transitiv)
    gleich alt und liegen in der Harris-Matrix auf einer Ebene."""
    parent = {nr: nr for nr in stellen}
    def find(x):
        r = x
        while parent[r] != r: r = parent[r]
        while parent[x] != r:           # Pfadkompression
            parent[x], x = r, parent[x]
        return r
    for nr, s in stellen.items():
        for g in s.alle_gleich_alt_wie():   # union-find ist symmetrisch →
            if g in parent:                 # einseitige Einträge genügen
                ra, rb = find(nr), find(g)
                if ra != rb: parent[ra] = rb
    return {nr: find(nr) for nr in stellen}


def _bereinige_gleich_ordnung(stellen: Dict[str, Stelle]) -> list:
    """'Gleich alt wie' gewinnt (Nutzerwunsch): eine direkte manuelle
    jünger/älter-Beziehung ZWISCHEN zwei gleich-alten Stellen wird aus den
    Datenfeldern entfernt und als unterdrückt vermerkt — damit 'Stellen & Harris'
    konsistent ist und die Beziehung nicht neu entsteht. Wird beim Zeichnen einer
    Gleich-Linie und beim Laden aufgerufen (nicht im Dauer-Rechentakt).
    Rückgabe: Liste (aelter, juenger) entfernter Beziehungen."""
    comp = _gleich_komponenten(stellen)
    entfernt = []
    def _weg(stelle, attr, wert, richtung):
        lst = getattr(stelle, attr)
        if wert in lst:
            setattr(stelle, attr, [x for x in lst if x != wert])
        key = f"{richtung}:{wert}"
        if key not in stelle.unterdrueckte_bez:
            stelle.unterdrueckte_bez.append(key)
    for nr, s in stellen.items():
        cn = comp.get(nr)
        for g in list(s.juenger_als):                 # nr jünger als g
            if g in stellen and g != nr and comp.get(g) == cn:
                _weg(s, "juenger_als", g, "juenger")
                _weg(stellen[g], "aelter_als", nr, "aelter")
                entfernt.append((g, nr))
        for g in list(s.aelter_als):                  # nr älter als g
            if g in stellen and g != nr and comp.get(g) == cn:
                _weg(s, "aelter_als", g, "aelter")
                _weg(stellen[g], "juenger_als", nr, "juenger")
                entfernt.append((nr, g))
    return entfernt


# Beziehungstypen → zugehörige Datenfelder + Gegen-Richtung auf der anderen Seite
_BEZ_FELDER = {
    "juenger": ("juenger_als", "profil_juenger_als", "okuk_juenger_als", "dat_juenger_als"),
    "aelter":  ("aelter_als",  "profil_aelter_als",  "okuk_aelter_als",  "dat_aelter_als"),
    "gleich":  ("gleich_alt_wie", "profil_gleich"),
}
_BEZ_GEGEN = {"juenger": "aelter", "aelter": "juenger", "gleich": "gleich"}

def beziehung_exklusiv(stellen: Dict[str, Stelle], nr: str, other: str, typ: str):
    """'Letzte Eingabe gewinnt': sorgt dafür, dass zwischen nr und other NUR die
    Beziehung `typ` besteht (typ: 'juenger' = nr jünger als other | 'aelter' |
    'gleich'). Alle WIDERSPRECHENDEN Typen werden auf BEIDEN Seiten aus den
    Datenfeldern entfernt und unterdrückt; der gewünschte Typ wird ent-unterdrückt.
    Setzt die gewünschte Beziehung NICHT selbst (das macht der Aufrufer) — räumt
    nur die Widersprüche weg. So aktualisiert eine neue Eingabe (Text ODER Linie)
    eine ältere widersprüchliche Eingabe zwischen demselben Stellenpaar."""
    s = stellen.get(nr); o = stellen.get(other)
    if not s or not o or nr == other:
        return
    def _strip(stelle, wert, t):
        for attr in _BEZ_FELDER[t]:
            lst = getattr(stelle, attr)
            if wert in lst:
                setattr(stelle, attr, [x for x in lst if x != wert])
    def _supp(stelle, t, wert, an):
        key = f"{t}:{wert}"
        if an and key not in stelle.unterdrueckte_bez:
            stelle.unterdrueckte_bez.append(key)
        elif not an and key in stelle.unterdrueckte_bez:
            stelle.unterdrueckte_bez.remove(key)
    for t in ("juenger", "aelter", "gleich"):
        g = _BEZ_GEGEN[t]
        if t == typ:
            _supp(s, t, other, False)   # gewünschten Typ nicht blockieren
            _supp(o, g, nr, False)
        else:
            _strip(s, other, t)         # Widerspruch auf beiden Seiten entfernen
            _strip(o, nr, g)
            _supp(s, t, other, True)
            _supp(o, g, nr, True)


def _harris_topologie(stellen: Dict[str, Stelle]):
    """Kern der Harris-Höhenberechnung mit GLEICH-GEWINNT-Semantik.

    Gleich-alte Stellen werden zu EINER Komponente verschmolzen. Ordnungs-
    beziehungen (jünger/älter) INNERHALB einer Komponente werden ignoriert
    (eine gezogene Gleichsetzung 'überstimmt' die alte Über-/Unterlagerung).
    Zwischen den Komponenten wird der Längste-Pfad gerechnet; sollten dort doch
    noch Zyklen entstehen, werden sie deterministisch aufgebrochen. Dadurch
    kann die Matrix NIE mehr in eine Reihe kollabieren (außer es ist wirklich
    alles gleich alt = eine einzige Komponente).

    Rückgabe: (levels, ignoriert)
      levels    : Dict Stelle → Level (0 = älteste); Gleich-Gruppe teilt Level.
      ignoriert : Liste (aelter_nr, juenger_nr) Original-Beziehungen, die wegen
                  einer Gleichsetzung übergangen wurden (für Hinweise)."""
    if not stellen:
        return {}, []
    comp = _gleich_komponenten(stellen)
    reps = sorted(set(comp.values()), key=_sort_key)

    # Ordnungskanten auf Komponentenebene sammeln; Original-Paare merken
    kanten: Dict[tuple, list] = {}     # (alt_rep, jung_rep) → [(alt_nr,jung_nr)]
    intra: list = []                   # innerhalb einer Komponente → ignoriert
    for nr, s in stellen.items():
        cn = comp[nr]
        for x in s.alle_juenger_als():        # nr jünger als x → x älter
            if x in comp:
                if comp[x] == cn: intra.append((x, nr))
                else: kanten.setdefault((comp[x], cn), []).append((x, nr))
        for y in s.alle_aelter_als():         # nr älter als y
            if y in comp:
                if comp[y] == cn: intra.append((nr, y))
                else: kanten.setdefault((cn, comp[y]), []).append((nr, y))

    adj = {r: set() for r in reps}
    for (a, b) in kanten: adj[a].add(b)

    # Zyklen deterministisch aufbrechen (DFS, Rückwärtskanten verwerfen)
    WHITE, GRAY, BLACK = 0, 1, 2
    farbe = {r: WHITE for r in reps}
    keep  = {r: set() for r in reps}
    for start in reps:
        if farbe[start] != WHITE: continue
        farbe[start] = GRAY
        stack = [(start, iter(sorted(adj[start], key=_sort_key)))]
        while stack:
            u, it = stack[-1]
            weiter = False
            for v in it:
                if farbe[v] == WHITE:
                    keep[u].add(v); farbe[v] = GRAY
                    stack.append((v, iter(sorted(adj[v], key=_sort_key))))
                    weiter = True; break
                elif farbe[v] == GRAY:
                    pass                      # Rückwärtskante → Zyklus → verwerfen
                else:
                    keep[u].add(v)            # Vorwärts-/Querkante → unkritisch
            if not weiter:
                farbe[u] = BLACK; stack.pop()

    # Längster Pfad auf dem nun zyklenfreien Komponentengraphen
    clevel = {r: 0 for r in reps}
    changed = True; it = 0
    while changed and it <= len(reps):
        changed = False; it += 1
        for u in reps:
            for v in keep[u]:
                if clevel[v] < clevel[u] + 1:
                    clevel[v] = clevel[u] + 1; changed = True

    levels = {nr: clevel[comp[nr]] for nr in stellen}

    # Übergangene Beziehungen einsammeln (innerhalb Komponente + verworfene
    # Komponentenkanten)
    ignoriert = list(intra)
    for (a, b), paare in kanten.items():
        if b not in keep.get(a, ()):
            ignoriert.extend(paare)
    return levels, ignoriert


def berechne_levels(stellen: Dict[str, Stelle]) -> Dict[str, int]:
    """
    Harris-Höhen. Level 0 = ÄLTESTE Stelle, höhere Level = JÜNGERE Stellen.
    Gleich-alte Stellen teilen sich ein Level (siehe _harris_topologie).
    """
    levels, _ = _harris_topologie(stellen)
    # Datierungs-Vorsortierung fuer isolierte Stellen
    # (keine Harris-Beziehungen -> Einordnung nach juengstem Fund)
    verbunden = {nr for nr, s in stellen.items()
                 if s.alle_juenger_als() or s.alle_aelter_als()
                 or s.alle_gleich_alt_wie()}
    isoliert  = [(nr, s.effektiv_bis())
                 for nr, s in stellen.items()
                 if nr not in verbunden
                 and s.effektiv_bis() is not None]
    if isoliert:
        max_struct = max((levels[nr] for nr in verbunden), default=0)
        max_stufen = max(max_struct, 4)
        bis_werte  = sorted(set(b for _, b in isoliert))
        n = len(bis_werte)
        if n > 1:
            bis_zu_level = {b: int(i * max_stufen / (n - 1))
                            for i, b in enumerate(bis_werte)}
            for nr, bis in isoliert:
                levels[nr] = bis_zu_level[bis]
    return levels


def topologische_sortierung(stellen: Dict[str, Stelle]) -> List[Stelle]:
    """Kahn-Algorithmus, älteste zuerst. Kanten: ältere → jüngere."""
    levels  = berechne_levels(stellen)
    in_deg  = {nr: 0 for nr in stellen}
    nachf: Dict[str, Set[str]] = {nr: set() for nr in stellen}

    kanten: Set[Tuple[str,str]] = set()
    for nr, s in stellen.items():
        for xnr in s.alle_juenger_als():    # xnr (älter) → nr (jünger)
            if xnr in stellen and (xnr,nr) not in kanten:
                kanten.add((xnr,nr))
                nachf[xnr].add(nr); in_deg[nr] += 1
        for ynr in s.alle_aelter_als():     # nr (älter) → ynr (jünger)
            if ynr in stellen and (nr,ynr) not in kanten:
                kanten.add((nr,ynr))
                nachf[nr].add(ynr); in_deg[ynr] += 1

    def sk(nr):
        s = stellen[nr]
        v = s.effektiv_von()  if s.effektiv_von()  is not None else 99999
        b = s.effektiv_bis() if s.effektiv_bis() is not None else 99999
        return (v, b, nr)

    queue = sorted([nr for nr,d in in_deg.items() if d==0], key=sk)
    sortiert=[]
    while queue:
        queue.sort(key=sk); cur=queue.pop(0)
        sortiert.append(stellen[cur])
        for nx in nachf.get(cur, set()):
            in_deg[nx]-=1
            if in_deg[nx]==0: queue.append(nx)

    done={s.stellennr for s in sortiert}
    sortiert += sorted([s for nr,s in stellen.items() if nr not in done],
                       key=lambda s: sk(s.stellennr))
    return sortiert


def analysiere_warnungen(stellen: Dict[str, Stelle]) -> List[Warnung]:
    ws: List[Warnung] = []
    def w(stufe, stelle, msg): ws.append(Warnung(stufe,stelle,msg))

    # 'Gleich alt wie' hat Vorrang vor älteren Über-/Unterlagerungen
    # (Nutzerwunsch: neu gezogene Gleich-Linien überstimmen). Wo eine
    # jünger/älter-Beziehung deswegen übergangen wurde, NUR informieren
    # (Stufe 3, kein Widerspruch) — die Matrix bleibt korrekt.
    _, ignoriert = _harris_topologie(stellen)
    gemeldet = set()
    for alt, jung in ignoriert:
        schluessel = (alt, jung)
        if schluessel in gemeldet: continue
        gemeldet.add(schluessel)
        w(3, jung, f"Beziehung '{jung} jünger als {alt}' wird in der Harris-Matrix "
                   f"übergangen, weil eine Gleichsetzung Vorrang hat. Falls die "
                   f"Über-/Unterlagerung doch gelten soll, Gleich-Verbindung entfernen.")

    for nr, s in stellen.items():
        # Störung mit altem Material (Residualfunde) — VOR dem Überspringen
        # prüfen, sonst ist diese Warnung unerreichbar (Störungen werden
        # darunter generell übersprungen)
        if s.moderne_stoerung:
            fund_von = [f.von for f in s.funde if f.von is not None]
            if fund_von and min(fund_von) < 1800:
                w(2,nr,f"Störung {nr} enthält Material ab ~{min(fund_von)} – "
                       f"mögliche Residualfunde.")
        # Arbeitsbereiche und Störungen überspringen (datierungsirrelevant)
        if s.kein_befund or s.moderne_stoerung: continue

        ef_von = s.effektiv_von(); ef_bis = s.effektiv_bis()

        # Modernes Material
        for f in s.funde:
            if f.ist_modern():
                w(3,nr,f"Modernes Material in Fund: '{f.material or f.ansprache}' [{f.datierung}]"); break

        # Schlechte Sichtbarkeit
        if s.sichtbar in ("schlecht","mäßig"):
            w(3,nr,f"Befund war {s.sichtbar} sichtbar – Datierung mit Vorsicht.")

        # Keine Funde, keine Zeichnung
        if not s.funde and not s.zeichnung:
            w(3,nr,"Keine Funde in Fundliste und keine Zeichnung. Fotos/Vermessung prüfen.")

        # Datierung aus Katalog ≠ aus Fundliste
        if s.datierung_katalog and s.auto_datierung_str:
            kat_von, kat_bis = datierung_zu_jahre(s.datierung_katalog)
            if (kat_von is not None and ef_von is not None
                    and abs(kat_von - ef_von) > 200):
                w(2,nr,f"Datierung Stellenkatalog ('{s.datierung_katalog}') weicht von "
                        f"Fundliste ('{s.auto_datierung_str}') ab.")

        # WNK-Begriffe ohne Jahreszahlen: bekannt, aber datierungslos — ohne
        # Hinweis erführe der Nutzer nie, warum kein Datum berechnet wird
        _ohne_jahre = []
        for beg in ([s.datierung_katalog]
                    + [f.datierung for f in s.funde if f.datierung]):
            beg = (beg or "").strip()
            if (beg and beg.lower() != "datierung unbekannt"
                    and beg.lower() in _WNK_LOWER
                    and datierung_zu_jahre(beg) == (None, None)
                    and beg not in _ohne_jahre):
                _ohne_jahre.append(beg)
        if _ohne_jahre:
            w(3,nr,f"Datierung '{', '.join(_ohne_jahre)}' ist im WNK bekannt, "
                   f"enthält dort aber keine Jahresangaben – fließt nicht in "
                   f"die Datumsberechnung ein.")

        # Datierung aus Stratigraphie kommentieren
        if s.inferred_datierung_str and s.inferred_quelle:
            w(3,nr,f"Datierung erschlossen: {s.inferred_datierung_str}")

        # OK/UK-Beziehungen kommentieren
        if s.okuk_juenger_als:
            w(3,nr,f"Vorläufige Reihenfolge aus OK/UK: jünger als {', '.join(s.okuk_juenger_als)} "
                   f"(OK={s.tiefe_ok:.2f}m) – bitte manuell prüfen!")

        # Profil-Parser-Ergebnisse kommentieren
        if s.profil_juenger_als or s.profil_aelter_als or s.profil_gleich:
            w(3,nr,f"Beziehungen aus Profilbeschreibung automatisch erkannt: "
                   f"↑{s.profil_juenger_als} ↓{s.profil_aelter_als} ={s.profil_gleich}")

        # Stratigraphische Datierungswidersprüche
        for jnr in s.alle_juenger_als():
            if jnr not in stellen: continue
            j=stellen[jnr]; j_von=j.effektiv_von(); j_bis=j.effektiv_bis()
            # Manuell gesetzte Beziehung (Linie gezeichnet oder per Hand eingetragen)
            # hat immer Vorrang – kein Widerspruch-Warning dafür erzeugen.
            ist_manuell = (jnr in s.juenger_als or nr in stellen[jnr].aelter_als)
            if ef_bis is not None and j_von is not None:
                if ef_bis < j_von:
                    if not ist_manuell:
                        w(1,nr,f"Widerspruch: {nr} (endet ~{ef_bis}) soll jünger als "
                               f"{jnr} (beginnt ~{j_von}) sein – zeitlich aber früher datiert.")
                elif (ef_von is not None and j_bis is not None
                      and ef_von < j_bis and not ist_manuell):
                    # Manuell gezeichnete/eingetragene Beziehungen lösen auch
                    # keinen Überlappungs-Hinweis aus (Nutzer hat entschieden)
                    w(2,nr,f"Zeitliche Überlappung: {nr} und {jnr} überlappen obwohl "
                           f"{nr} stratigraphisch jünger sein soll.")

        # Durchschossen
        if s.durchschossen:
            w(3,nr,f"Stelle {nr} wurde untertunnelt – Beziehungen zu Nachbarstellen prüfen.")

    return sorted(ws, key=lambda x: x.stufe)

# ═══════════════════════════════════════════════════════════════
# STRATIGRAPHISCHE DATIERUNGSPROPAGIERUNG
# ═══════════════════════════════════════════════════════════════

def propagiere_stratigraphische_datierungen(
        stellen: Dict[str, "Stelle"]) -> None:
    """Iterative bidir. Propagierung der Datierungsgrenzen.
    Aufwaerts: juengere Stellen erhalten strat_von = max(strat_von, Quelle).
    Abwaerts:  aeltere  Stellen erhalten strat_bis = min(strat_bis, Quelle).
    Undatierte Zwischenstellen leiten korrekt weiter.
    """
    for s in stellen.values():
        s.strat_von = None; s.strat_bis = None; s.dat_protokoll = ""

    changed = True; it = 0
    while changed and it < 80:
        changed = False; it += 1
        for nr, s in stellen.items():
            rv = s._raw_von()
            if s.strat_von is not None:
                rv = max(rv, s.strat_von) if rv is not None else s.strat_von
            rb = s._raw_bis()
            if s.strat_bis is not None:
                rb = min(rb, s.strat_bis) if rb is not None else s.strat_bis
            if rv is None and rb is None:
                continue
            for anr in s.alle_aelter_als():   # anr ist juenger als s
                if anr not in stellen: continue
                a = stellen[anr]
                # Manuelle Beziehungen definieren nur Struktur, keine Datierung.
                # Propagierung durch sie hindurch würde bei Datierungsumkehrungen
                # (z.B. Neuzeit-Grube manuell als älter als undatierte Schicht)
                # falsche strat_von-Werte erzeugen.
                if anr in s.aelter_als or nr in stellen[anr].juenger_als:
                    continue
                # Nicht propagieren wenn strat_von das eigene Datum der Zielstelle
                # widersprechen würde.
                a_raw_bis = a._raw_bis()
                if rv is not None and a_raw_bis is not None and rv > a_raw_bis:
                    continue
                if rv is not None and (a.strat_von is None or rv > a.strat_von):
                    a.strat_von = rv; changed = True
            for jnr in s.alle_juenger_als():  # jnr ist aelter als s
                if jnr not in stellen: continue
                j = stellen[jnr]
                # Manuelle Beziehungen definieren nur Struktur, keine Datierung
                # (Gegenstück zum Guard in der aelter_als-Schleife oben)
                if jnr in s.juenger_als or nr in j.aelter_als:
                    continue
                # Obergrenze für die ältere Stelle: s beginnt bei rv → tighteste
                # Grenze; ist nur rb (Endjahr) bekannt, gilt immerhin j.bis ≤ rb
                grenze = rv if rv is not None else rb
                j_raw_von = j._raw_von()
                if grenze is not None and (j_raw_von is None or grenze >= j_raw_von):
                    if j.strat_bis is None or grenze < j.strat_bis:
                        j.strat_bis = grenze; changed = True
            for gnr in s.alle_gleich_alt_wie():
                if gnr not in stellen: continue
                g = stellen[gnr]
                g_raw_bis = g._raw_bis(); g_raw_von = g._raw_von()
                if rv is not None and (g.strat_von is None or rv > g.strat_von):
                    if g_raw_bis is None or rv <= g_raw_bis:
                        g.strat_von = rv; changed = True
                if rb is not None and (g.strat_bis is None or rb < g.strat_bis):
                    if g_raw_von is None or rb >= g_raw_von:
                        g.strat_bis = rb; changed = True

    for nr, s in stellen.items():
        p = []
        if s.auto_datierung_str:
            p.append("Fundliste/Katalog: " + s.auto_datierung_str
                     + (" (" + str(s.auto_von) + "-" + str(s.auto_bis) + ")"
                        if s.auto_von else ""))
        if s.datierung_katalog and s.datierung_katalog != s.auto_datierung_str:
            p.append("Katalog: " + s.datierung_katalog)
        if s.manuell_datierung:
            p.append("Manuell: " + s.manuell_datierung
                     + (" (" + str(s.manuell_von) + "-" + str(s.manuell_bis) + ")"
                        if s.manuell_von else ""))
        for a in s.ankerpunkte:
            p.append("Ankerpunkt (" + a.art.upper() + "): " + a.als_text())
        if s.strat_von is not None:
            p.append("Strat. Minimum: ab " + str(s.strat_von))
        if s.strat_bis is not None:
            p.append("Strat. Maximum: bis " + str(s.strat_bis))
        ev = s.effektiv_von(); eb = s.effektiv_bis()
        ev_s = str(ev) if ev is not None else "?"
        eb_s = str(eb) if eb is not None else "?"
        if ev is not None or eb is not None:
            p.append("Effektiv: " + ev_s + " - " + eb_s)
        s.dat_protokoll = "\n".join(p) if p else "(keine Datierungsdaten)"


# ═══════════════════════════════════════════════════════════════
# DATIERUNGS-BEZIEHUNGEN (vorläufig aus nicht-überlappenden Datierungen)
# ═══════════════════════════════════════════════════════════════

def datierungs_beziehungen_berechnen(stellen: Dict[str, "Stelle"]) -> int:
    """Erstellt vorläufige Harris-Beziehungen aus klar nicht-überlappenden
    Datierungen. Bedingung: die relevante Datumsgrenze muss bekannt sein –
    also A.bis für „A endet vor B" bzw. B.von. Keine Spannenbeschränkung,
    damit auch Urgeschichte vs. Neuzeit erkannt wird. Max. 5 pro Stelle.
    Zyklenprüfung: dat-Beziehungen die der bestehenden Topologie (manuell/
    profil/okuk) widersprechen werden übersprungen → kein „alle nebeneinander"
    nach manuellem Linienzeichnen + F5."""
    for s in stellen.values():
        s.dat_juenger_als = []
        s.dat_aelter_als  = []
    # Topologie OHNE dat ermitteln (dat gerade gelöscht) → Grundlage für Zyklenprüfung
    lvls = berechne_levels(stellen)
    # Kandidaten: mindestens ein Datumswert bekannt, kein AB/Störung
    kand = [(nr, s) for nr, s in stellen.items()
            if not s.kein_befund and not s.moderne_stoerung
            and not s.dat_suppressed
            and (s.effektiv_von() is not None or s.effektiv_bis() is not None)]
    neu = 0
    for i, (nr_a, s_a) in enumerate(kand):
        for nr_b, s_b in kand[i+1:]:
            if (len(s_a.dat_aelter_als) >= 5 or len(s_b.dat_juenger_als) >= 5
                    or len(s_b.dat_aelter_als) >= 5 or len(s_a.dat_juenger_als) >= 5):
                continue
            bez = (nr_b in s_a.alle_juenger_als() or
                   nr_b in s_a.alle_aelter_als() or
                   nr_b in s_a.alle_gleich_alt_wie())
            if bez: continue
            # Vom Nutzer gelöschte Auto-Beziehung (Editor) → nicht neu erzeugen
            if (f"juenger:{nr_a}" in s_b.unterdrueckte_bez
                    or f"aelter:{nr_b}" in s_a.unterdrueckte_bez
                    or f"juenger:{nr_b}" in s_a.unterdrueckte_bez
                    or f"aelter:{nr_a}" in s_b.unterdrueckte_bez):
                continue
            a_bis = s_a.effektiv_bis()
            b_von = s_b.effektiv_von()
            b_bis = s_b.effektiv_bis()
            a_von = s_a.effektiv_von()
            if a_bis is not None and b_von is not None and a_bis < b_von:
                # A endet klar vor B → B jünger als A (B muss höher im Baum sein)
                # Widerspricht bestehender Topologie (B bereits tiefer als A)? → Zyklus → skip
                if lvls.get(nr_b, 0) < lvls.get(nr_a, 0):
                    continue
                if nr_a not in s_b.dat_juenger_als:
                    s_b.dat_juenger_als.append(nr_a)
                if nr_b not in s_a.dat_aelter_als:
                    s_a.dat_aelter_als.append(nr_b)
                neu += 1
            elif b_bis is not None and a_von is not None and b_bis < a_von:
                # B endet klar vor A → A jünger als B (A muss höher im Baum sein)
                # Widerspricht bestehender Topologie (A bereits tiefer als B)? → Zyklus → skip
                if lvls.get(nr_a, 0) < lvls.get(nr_b, 0):
                    continue
                if nr_b not in s_a.dat_juenger_als:
                    s_a.dat_juenger_als.append(nr_b)
                if nr_a not in s_b.dat_aelter_als:
                    s_b.dat_aelter_als.append(nr_a)
                neu += 1
    log.debug(f"Datierungs-Beziehungen: {neu}")
    return neu


# ═══════════════════════════════════════════════════════════════
# PROJEKT SPEICHERN / LADEN
# ═══════════════════════════════════════════════════════════════

def stelle_zu_dict(s: Stelle) -> dict:
    return {
        "stellennr":s.stellennr, "befansprac":s.befansprac,
        "datierung_katalog":s.datierung_katalog,
        "kein_befund":s.kein_befund, "rest_befund":s.rest_befund,
        "moderne_stoerung":s.moderne_stoerung, "durchschossen":s.durchschossen,
        "sichtbar":s.sichtbar, "zeichnung":s.zeichnung, "kommentar":s.kommentar,
        "planum":s.planum, "profil":s.profil, "bearbeiter":s.bearbeiter,
        "tiefe_ok":s.tiefe_ok, "tiefe_uk":s.tiefe_uk,
        # list(...) zwingend: Undo-Snapshots dürfen die Listenobjekte NICHT mit
        # der lebenden Stelle teilen, sonst verändert jedes .append() den Snapshot mit
        "juenger_als":list(s.juenger_als), "aelter_als":list(s.aelter_als),
        "gleich_alt_wie":list(s.gleich_alt_wie), "gestoert_von":list(s.gestoert_von),
        "okuk_juenger_als":list(s.okuk_juenger_als), "okuk_aelter_als":list(s.okuk_aelter_als),
        "profil_juenger_als":list(s.profil_juenger_als), "profil_aelter_als":list(s.profil_aelter_als),
        "profil_gleich":list(s.profil_gleich),
        "unterdrueckte_bez":list(s.unterdrueckte_bez),
        "manuell_datierung":s.manuell_datierung,
        "manuell_von":s.manuell_von, "manuell_bis":s.manuell_bis,
        "dat_suppressed":s.dat_suppressed,
        "box_farbe":s.box_farbe, "box_form":s.box_form,
        "box_geschuetzt":s.box_geschuetzt,
        "box_x":s.box_x, "box_y":s.box_y,
        "ankerpunkte":[{"art":a.art,"beschreibung":a.beschreibung,"von":a.von,
                        "praezision":a.praezision}
                       for a in s.ankerpunkte],
    }

def stelle_aus_dict(d: dict) -> Stelle:
    felder = ["stellennr","befansprac","datierung_katalog","kein_befund","rest_befund",
              "moderne_stoerung","durchschossen","sichtbar","zeichnung","kommentar",
              "planum","profil","bearbeiter","tiefe_ok","tiefe_uk",
              "manuell_datierung","manuell_von","manuell_bis"]
    kwargs = {k: d.get(k, "") for k in felder}
    for k in ["kein_befund","rest_befund","moderne_stoerung","durchschossen"]:
        kwargs[k] = bool(d.get(k, False))
    for k in ["tiefe_ok","tiefe_uk","manuell_von","manuell_bis"]:
        kwargs[k] = d.get(k)
    s = Stelle(**kwargs)
    s.dat_suppressed = bool(d.get("dat_suppressed", False))
    s.box_farbe      = d.get("box_farbe", "")
    s.box_form       = d.get("box_form",  "rechteck")
    s.box_geschuetzt = bool(d.get("box_geschuetzt", False))
    s.box_x          = d.get("box_x")
    s.box_y          = d.get("box_y")
    s.stellennr = norm_stellennr(s.stellennr)
    for ls in ["juenger_als","aelter_als","gleich_alt_wie","gestoert_von",
               "okuk_juenger_als","okuk_aelter_als",
               "profil_juenger_als","profil_aelter_als","profil_gleich"]:
        # list(...) + Normalisierung: keine geteilten Objekte mit dem Snapshot,
        # alte Projektdateien mit '004'-Nummern bleiben kompatibel
        setattr(s, ls, [norm_stellennr(x) for x in d.get(ls, [])])
    # Unterdrückte Auto-Beziehungen ("richtung:nr") – nr-Teil normalisieren
    s.unterdrueckte_bez = []
    for eintrag in d.get("unterdrueckte_bez", []):
        if ":" in str(eintrag):
            richtung, _, unr = str(eintrag).partition(":")
            s.unterdrueckte_bez.append(f"{richtung}:{norm_stellennr(unr)}")
    # Ein Anker pro Art (Editor-Modell). Alte Projektdateien können mehrere
    # bzw. Zusatzfelder (bis/monat/tag) enthalten — es zählt der jeweils
    # strengste Anker: jüngster TPQ, ältester TAQ, erster exakter.
    gesehen: Dict[str, Ankerpunkt] = {}
    for a in d.get("ankerpunkte",[]):
        neu = Ankerpunkt(
            art=a.get("art","exakt"), beschreibung=a.get("beschreibung",""),
            von=a.get("von"), praezision=a.get("praezision","Jahr"))
        alt = gesehen.get(neu.art)
        if alt is None:
            gesehen[neu.art] = neu
        elif neu.von is not None and (
                alt.von is None
                or (neu.art == "tpq" and neu.von > alt.von)
                or (neu.art == "taq" and neu.von < alt.von)):
            gesehen[neu.art] = neu
    s.ankerpunkte = list(gesehen.values())
    return s

# ── Ende Block 1 ──────────────────────────────────────────────
# Block 2 (GUI) direkt darunter einfügen.

# ═══════════════════════════════════════════════════════════════
# BLOCK 2 von 2  –  GUI
# ═══════════════════════════════════════════════════════════════

# ── Undo-Stack ────────────────────────────────────────────────

class UndoStack:
    """Speichert bis zu MAX Snapshots des Stellen-Dicts für Ctrl+Z."""
    MAX = 5

    def __init__(self):
        self._stack: List[List[dict]] = []

    def push(self, stellen: Dict[str, "Stelle"]):
        self._stack.append([stelle_zu_dict(s) for s in stellen.values()])
        if len(self._stack) > self.MAX:
            self._stack.pop(0)

    def pop(self) -> Optional[List[dict]]:
        return self._stack.pop() if self._stack else None

    def can_undo(self) -> bool:
        return bool(self._stack)

    def size(self) -> int:
        return len(self._stack)


# ── GUI-Hilfsmittel ───────────────────────────────────────────

def _ac_filtern(combo, liste, event=None):
    """Gemeinsame Filter- und Autocomplete-Logik für beide Combobox-Klassen.
    Filtert die Dropdown-Liste ab 1 Zeichen und öffnet den Dropdown automatisch.
    Pfeiltasten / Tab / Enter übernehmen den markierten Eintrag."""
    if event and event.keysym in (
            "Return","Tab","Escape","Up","Down","Left","Right"):
        return
    text = combo.get().lower().strip()
    if len(text) >= 1:
        vorne  = [v for v in liste if v.lower().startswith(text)]
        hinten = [v for v in liste if text in v.lower() and v not in vorne]
        neue   = vorne + hinten
        combo['values'] = neue
        if neue:
            try: combo.tk.call('ttk::combobox::Post', combo)
            except Exception: pass
    else:
        combo['values'] = liste


class AutocompleteCombo(ttk.Combobox):
    """Combobox mit Live-Filterung der WNK-Datierungsbegriffe.
    Breite orientiert am längsten Eintrag (~46 Zeichen)."""
    def __init__(self, parent, textvariable, **kw):
        kw.setdefault('width', 52)
        super().__init__(parent, textvariable=textvariable, **kw)
        self['values'] = _DATIERUNGEN_LISTE
        self.bind('<KeyRelease>',
                  lambda e: _ac_filtern(self, _DATIERUNGEN_LISTE, e))


class AutocompleteAnsprache(ttk.Combobox):
    """Combobox mit Live-Filterung der WNK-Befundansprachen.
    Breite orientiert am längsten Eintrag (~40 Zeichen)."""
    def __init__(self, parent, textvariable, **kw):
        kw.setdefault('width', 44)
        super().__init__(parent, textvariable=textvariable, **kw)
        self['values'] = _ANSPRACHEN_LISTE
        self.bind('<KeyRelease>',
                  lambda e: _ac_filtern(self, _ANSPRACHEN_LISTE, e))

class GUILogHandler(logging.Handler):
    def __init__(self):
        super().__init__()
        self.setFormatter(_fmt)
        self.callback = None
    def emit(self, r):
        if self.callback:
            try: self.callback(self.format(r), r.levelname)
            except: pass


class ToolTip:
    def __init__(self, widget, text: str):
        self._w = widget; self._t = text; self._tip = None
        widget.bind("<Enter>", self._show)
        widget.bind("<Leave>", self._hide)
    def _show(self, _=None):
        x = self._w.winfo_rootx() + 22
        y = self._w.winfo_rooty() + self._w.winfo_height() + 4
        self._tip = tw = tk.Toplevel(self._w)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        tk.Label(tw, text=self._t, background="#fffacd", relief="solid",
                 borderwidth=1, font=("", 9), wraplength=420,
                 justify="left").pack(ipadx=4, ipady=2)
    def _hide(self, _=None):
        if self._tip:
            self._tip.destroy()
            self._tip = None


def hint(parent, text: str, row: int, col: int = 3):
    lbl = ttk.Label(parent, text=text, foreground="#bbbbbb", font=("", 8))
    lbl.grid(row=row, column=col, sticky="w", padx=4)
    return lbl


def popup(parent, titel: str, text: str):
    d = tk.Toplevel(parent)
    d.title(titel)
    d.geometry("740x520")
    d.minsize(400, 300)
    t = scrolledtext.ScrolledText(d, wrap=tk.WORD, font=("", 10))
    t.insert("1.0", text)
    t.config(state="disabled")
    t.pack(fill=tk.BOTH, expand=True, padx=10, pady=(10, 0))
    ttk.Button(d, text="Schließen  [Esc]", command=d.destroy).pack(pady=8)
    d.bind("<Escape>", lambda _: d.destroy())
    d.focus_set()
    try: parent.style_dialog(d)
    except Exception: pass


def _mache_tree(parent, cols, heads, breiten, height=0) -> ttk.Treeview:
    f = ttk.Frame(parent)
    f.pack(fill=tk.BOTH, expand=True)
    kw = dict(columns=cols, show="headings")
    if height:
        kw["height"] = height
    tree = ttk.Treeview(f, **kw)
    for c, h, b in zip(cols, heads, breiten):
        tree.heading(c, text=h)
        tree.column(c, width=b, minwidth=28, stretch=True)
    sy = ttk.Scrollbar(f, orient=tk.VERTICAL,   command=tree.yview)
    sx = ttk.Scrollbar(f, orient=tk.HORIZONTAL, command=tree.xview)
    tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
    sy.pack(side=tk.RIGHT,  fill=tk.Y)
    sx.pack(side=tk.BOTTOM, fill=tk.X)
    tree.pack(fill=tk.BOTH, expand=True)
    return tree

# ═══════════════════════════════════════════════════════════════
# TAB: FUNDLISTE
# ═══════════════════════════════════════════════════════════════

class FundlisteTab(ttk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self._alle: List[Fund] = []
        self._build()

    def _build(self):
        tb = ttk.Frame(self)
        self._tb_ref = tb
        tb.pack(fill=tk.X, padx=6, pady=4)
        ttk.Button(tb, text="Fundliste laden…",
                   command=self.app.cmd_lade_fundliste).pack(side=tk.LEFT, padx=2)
        ttk.Button(tb, text="Auto-Datierung",
                   command=self.app.cmd_auto_datiere).pack(side=tk.LEFT, padx=2)
        self.info = tk.StringVar(value="Keine Fundliste geladen.")
        ttk.Label(tb, textvariable=self.info).pack(side=tk.LEFT, padx=10)
        ttk.Label(tb, text="Suche:").pack(side=tk.RIGHT)
        self.sv = tk.StringVar()
        self.sv.trace_add("write", lambda *_: self._filter())
        ttk.Entry(tb, textvariable=self.sv, width=18).pack(side=tk.RIGHT, padx=4)

        drop = ttk.LabelFrame(self, text="Drag & Drop")
        drop.pack(fill=tk.X, padx=6, pady=(0, 4))
        msg = ("📂  Fundliste ODER Stellenkatalog (CSV) hier reinziehen — "
               "Typ wird automatisch erkannt"
               if DND_OK else
               "📂  Fundliste ODER Stellenkatalog (CSV) hier reinziehen — "
               "Typ wird automatisch erkannt\n"
               "    (Drag & Drop aktivieren: pip install tkinterdnd2 — oder Buttons nutzen)")
        self.drop_lbl = tk.Label(drop, text=msg, height=3,
                                  background="#1a2a3a", foreground="#7ab3e0",
                                  font=("", 10), cursor="hand2",
                                  relief="groove", borderwidth=2)
        self.drop_lbl.pack(fill=tk.X, padx=6, pady=4)
        self.drop_lbl.bind("<Button-1>", lambda _: self.app.cmd_lade_fundliste())
        if DND_OK:
            try:
                self.drop_lbl.drop_target_register(DND_FILES)
                self.drop_lbl.dnd_bind("<<Drop>>", self._on_drop)
            except Exception as e:
                log.warning(f"DnD: {e}")

        cols    = ("Stelle","Pos","U","Material","Ansprache","Kommentar",
                   "Datierung","Von","Bis","Bearb.")
        heads   = ("Stelle","Pos","U-Nr","Material","Ansprache","Kommentar",
                   "Datierung","Von","Bis","Bearb.")
        breiten = [55, 40, 40, 130, 130, 200, 170, 55, 55, 110]
        self.tree = _mache_tree(self, cols, heads, breiten)
        self.tree.tag_configure("modern", background="#2a1a00",
                                foreground="#ffcc88")
        self.tree.bind("<Double-1>", self._on_dblclick)

    def _on_dblclick(self, event=None):
        """Doppelklick auf Fund → zugehörige Stelle in Stellen-Tab öffnen."""
        sel = self.tree.selection()
        if not sel: return
        nr = str(self.tree.item(sel[0], "values")[0])
        try:
            self.app.notebook.select(self.app.tab_stellen)
            if self.app.tab_stellen.tree.exists(nr):
                self.app.tab_stellen.tree.selection_set(nr)
                self.app.tab_stellen.tree.see(nr)
                self.app.tab_stellen._on_sel()
        except Exception: pass

    @staticmethod
    def _dnd_pfade(data: str) -> List[str]:
        """Drop-Daten in einzelne Pfade zerlegen. NICHT tk.splitlist nutzen —
        das interpretiert Backslashes in Windows-Pfaden (\\U, \\D …) als
        Tcl-Escapes und zerstört den Pfad. tkinterdnd2 klammert Pfade mit
        Leerzeichen in {…}; mehrere Dateien sind leerzeichengetrennt."""
        data = (data or "").strip()
        if not data:
            return []
        if "{" in data:
            teile = re.findall(r"\{([^}]*)\}", data)
            if teile:
                return [t for t in teile if t.strip()]
        # ohne Klammern: einzelner Pfad (Windows-Pfade haben i. d. R. keine
        # Leerzeichen ohne Klammern) → als ein Pfad behandeln
        return [data]

    def _on_drop(self, event):
        # Es können mehrere Dateien auf einmal fallen gelassen werden
        for roh in self._dnd_pfade(event.data):
            pfad = Path(roh.strip().strip("{}"))
            typ = erkenne_csv_typ(pfad)
            if typ == "stellenkatalog":
                self.app._lade_stellenkatalog_pfad(pfad)
            elif typ == "fundliste":
                self.app._lade_fundliste_pfad(pfad)
            else:
                # Spalten nicht eindeutig → den Nutzer entscheiden lassen
                als_fund = messagebox.askyesnocancel(
                    "CSV-Typ nicht erkannt",
                    f"Die Spalten von '{pfad.name}' konnten nicht eindeutig "
                    f"einer Fundliste oder einem Stellenkatalog zugeordnet werden.\n\n"
                    f"Als Fundliste laden?\n"
                    f"(Ja = Fundliste · Nein = Stellenkatalog · Abbrechen = nichts)")
                if als_fund is True:
                    self.app._lade_fundliste_pfad(pfad)
                elif als_fund is False:
                    self.app._lade_stellenkatalog_pfad(pfad)

    def aktualisiere(self, funde: List[Fund]):
        self._alle = funde
        self.info.set(f"{len(funde)} Funde geladen.")
        self._zeige(funde)
        n = len(_unbekannte_datierungen())
        if hasattr(self, '_warn_btn'):
            try: self._warn_btn.destroy()
            except: pass
        if n:
            self._warn_btn = tk.Button(
                self._tb_ref, text=f"⚠ {n} unbekannte Datierungen",
                foreground="red", background="#fff0f0",
                font=("", 9, "bold"), relief="flat", cursor="hand2",
                command=self._zeige_unbekannte_datierungen)
            self._warn_btn.pack(side=tk.LEFT, padx=6)

    def _zeige(self, funde):
        self.tree.delete(*self.tree.get_children())
        for f in funde:
            tag = ("modern",) if f.ist_modern() else ()
            self.tree.insert("", tk.END, tags=tag, values=(
                f.stellennr, f.posnr, f.unternr,
                f.material, f.ansprache,
                f.kommentar[:60] + ("…" if len(f.kommentar) > 60 else ""),
                f.datierung,
                f.von  if f.von  is not None else "",
                f.bis  if f.bis  is not None else "",
                f.ansprache_von,
            ))

    def _filter(self):
        t = self.sv.get().lower()
        self._zeige([f for f in self._alle if not t or
                     any(t in x.lower() for x in
                         [f.stellennr, f.material, f.ansprache,
                          f.datierung, f.kommentar])])

    def _zeige_unbekannte_datierungen(self):
        d = tk.Toplevel(self)
        d.title("Unbekannte Datierungsbegriffe")
        d.geometry("460x320")
        d.grab_set()
        ttk.Label(d, text="Folgende Datierungsbegriffe sind nicht im WNK-Katalog:",
                  font=("", 9, "bold")).pack(pady=(10, 4), padx=10, anchor="w")
        txt = scrolledtext.ScrolledText(d, width=55, height=14, font=("Courier", 9))
        txt.pack(fill=tk.BOTH, expand=True, padx=10, pady=4)
        for b in sorted(_unbekannte_datierungen()):
            txt.insert(tk.END, b + "\n")
        txt.config(state="disabled")
        ttk.Button(d, text="Schließen", command=d.destroy).pack(pady=6)
        self.app.style_dialog(d)

# ═══════════════════════════════════════════════════════════════
# TAB: STELLEN & HARRIS
# ═══════════════════════════════════════════════════════════════

class StellenTab(ttk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self._cur: Optional[str] = None
        self._snapshots: Dict[str, dict] = {}
        self._build()

    def _build(self):
        tb = ttk.Frame(self)
        tb.pack(fill=tk.X, padx=6, pady=4)
        ttk.Label(tb, text="Stellenbereich:").pack(side=tk.LEFT)
        self.v_von = tk.StringVar(value="1")
        self.v_bis = tk.StringVar(value="9999")
        ttk.Entry(tb, textvariable=self.v_von, width=7).pack(side=tk.LEFT, padx=2)
        ttk.Label(tb, text="–").pack(side=tk.LEFT)
        ttk.Entry(tb, textvariable=self.v_bis, width=7).pack(side=tk.LEFT, padx=2)
        ttk.Button(tb, text="Anlegen",        command=self._bereich).pack(side=tk.LEFT, padx=4)
        ttk.Button(tb, text="+ Neue Stelle",  command=self._neue_stelle).pack(side=tk.LEFT, padx=2)
        ttk.Button(tb, text="Harris aus OK/UK…", command=self._okuk_dialog).pack(side=tk.LEFT, padx=6)

        paned = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True, padx=6, pady=2)
        lf = ttk.LabelFrame(paned, text="Stellen")
        paned.add(lf, weight=1)
        self._build_liste(lf)
        rf = ttk.LabelFrame(paned, text="Stelle bearbeiten")
        paned.add(rf, weight=3)
        self._build_editor(rf)
        # Teiler beim Start schmaler setzen → mehr Platz für "Stelle bearbeiten".
        # Einmalig beim ersten echten Layout (danach bestimmt der Nutzer/weight).
        self._paned = paned
        self._sash_init = False
        def _init_sash(ev):
            if self._sash_init or ev.width < 300:
                return
            self._sash_init = True
            try: paned.sashpos(0, max(300, min(430, int(ev.width * 0.30))))
            except Exception: pass
        paned.bind("<Configure>", _init_sash)

    # ── Stellenliste ──────────────────────────────────────────
    def _build_liste(self, parent):
        sf = ttk.Frame(parent)
        sf.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(sf, text="Suche:").pack(side=tk.LEFT)
        self.sv = tk.StringVar()
        self.sv.trace_add("write", lambda *_: self.aktualisiere())
        ttk.Entry(sf, textvariable=self.sv, width=14).pack(
            side=tk.LEFT, padx=2, fill=tk.X, expand=True)

        inner = ttk.Frame(parent)
        inner.pack(fill=tk.BOTH, expand=True)
        cols    = ("Nr", "Ansprache", "Datierung", "Q", "F", "Flags")
        heads   = ("Stelle", "Ansprache", "Datierung", "Quelle", "F", "Flags")
        breiten = [60, 130, 180, 72, 28, 80]
        self.tree = ttk.Treeview(inner, columns=cols, show="headings",
                                  selectmode="browse")
        for c, h, b in zip(cols, heads, breiten):
            self.tree.heading(c, text=h)
            # stretch=False: Spalten behalten ihre Breite → bei schmaler Tabelle
            # erscheint der horizontale Bildlauf, statt die Spalten zu quetschen
            self.tree.column(c, width=b, minwidth=24, stretch=False)
        sy = ttk.Scrollbar(inner, orient=tk.VERTICAL, command=self.tree.yview)
        sx = ttk.Scrollbar(inner, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        sy.pack(side=tk.RIGHT, fill=tk.Y)
        sx.pack(side=tk.BOTTOM, fill=tk.X)   # horizontaler Bildlauf unten
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.tree.bind("<<TreeviewSelect>>", self._on_sel)
        self.tree.bind("<Control-c>", self._copy_tree)
        self.tree.tag_configure("ab",  foreground="#888888")
        self.tree.tag_configure("st",  foreground="#cc4400")
        self.tree.tag_configure("inf", foreground="#0055bb")
        self.tree.tag_configure("wid", background="#ffcccc")
        ToolTip(self.tree,
                "Q = Datierungsquelle: Manuell | Ankerpunkt | Fund | "
                "Stratigraphie | Katalog | –\n"
                "F = Anzahl Funde\n"
                "Flags: AB=Arbeitsbereich  ST=Störung  "
                "DU=Durchschossen  OKUK=vorläufig aus OK/UK  "
                "PROF=aus Profiltext")

    # ── Editor ────────────────────────────────────────────────
    def _build_editor(self, parent):
        cv = tk.Canvas(parent, borderwidth=0, highlightthickness=0)
        sb = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=cv.yview)
        cv.configure(yscrollcommand=sb.set)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        cv.pack(fill=tk.BOTH, expand=True)
        self._ef = ttk.Frame(cv)
        _wid = cv.create_window((0, 0), window=self._ef, anchor="nw")
        self._ef.bind("<Configure>",
                      lambda e: cv.configure(scrollregion=cv.bbox("all")))
        cv.bind("<Configure>",
                lambda e: cv.itemconfig(_wid, width=e.width))
        # Nur scrollen wenn Maus über diesem Canvas-Bereich ist (nicht bind_all!)
        def _cv_scroll(e):
            try:
                x, y = e.x_root, e.y_root
                if (cv.winfo_rootx() <= x <= cv.winfo_rootx() + cv.winfo_width() and
                        cv.winfo_rooty() <= y <= cv.winfo_rooty() + cv.winfo_height()):
                    cv.yview_scroll(int(-1*(e.delta/120)), "units")
            except Exception:
                pass
        cv.bind_all("<MouseWheel>", _cv_scroll)

        f = self._ef
        r = 0

        def L(t, row, c=0, bold=False):
            kw = {"font": ("", 9, "bold")} if bold else {}
            ttk.Label(f, text=t, **kw).grid(
                row=row, column=c, sticky="w", padx=6, pady=2)

        def E(var, row, c=1, w=28, ro=False):
            e = ttk.Entry(f, textvariable=var, width=w,
                          state="readonly" if ro else "normal")
            e.grid(row=row, column=c, columnspan=2, sticky="ew", padx=4, pady=2)
            return e

        def SEP(row):
            ttk.Separator(f, orient="horizontal").grid(
                row=row, column=0, columnspan=5,
                sticky="ew", pady=5, padx=4)

        # Basis
        L("Stelle Nr.:", r)
        self.v_nr = tk.StringVar()
        E(self.v_nr, r, ro=True)
        r += 1

        L("Befundansprache:", r)
        self.v_ansp = tk.StringVar()
        _ansp_combo = AutocompleteAnsprache(f, textvariable=self.v_ansp)
        _ansp_combo.grid(row=r, column=1, columnspan=2, sticky="ew", padx=4, pady=2)
        hint(f, "z. B. Graben, Pfostengrube", r)
        r += 1

        L("Flags:", r)
        ff = ttk.Frame(f)
        ff.grid(row=r, column=1, columnspan=4, sticky="w", pady=2)
        self.v_ab = tk.BooleanVar()
        self.v_st = tk.BooleanVar()
        self.v_rb = tk.BooleanVar()
        for txt, var in [("Arbeitsbereich", self.v_ab),
                          ("Störung",        self.v_st),
                          ("Restbefund",      self.v_rb)]:
            ttk.Checkbutton(ff, text=txt, variable=var).pack(side=tk.LEFT)
        r += 1

        L("Sichtbar / OK / UK:", r)
        ri = ttk.Frame(f)
        ri.grid(row=r, column=1, columnspan=4, sticky="w")
        self.v_sich = tk.StringVar()
        self.v_ok   = tk.StringVar()
        self.v_uk   = tk.StringVar()
        for lbl, var in [("Sichtbar:", self.v_sich),
                          ("OK (m):",  self.v_ok),
                          ("UK (m):",  self.v_uk)]:
            ttk.Label(ri, text=lbl).pack(side=tk.LEFT, padx=(6, 0))
            ttk.Entry(ri, textvariable=var, width=9,
                      state="readonly").pack(side=tk.LEFT, padx=2)
        ttk.Label(ri, text="(aus Katalog)", foreground="#aaaaaa",
                  font=("", 8)).pack(side=tk.LEFT, padx=4)
        r += 1

        SEP(r); r += 1
        L("── Datierung ──", r, bold=True); r += 1

        L("Auto (aus Funden):", r)
        self.v_aut = tk.StringVar()
        E(self.v_aut, r, ro=True, w=38)
        r += 1

        L("Stratigraph. Hinweis:", r)
        self.v_inf = tk.StringVar()
        E(self.v_inf, r, ro=True, w=38)
        hint(f, "Autom. z.B.: \"<= Mittelalter (Stratigraphie, wg. Stelle 30)\"", r)
        r += 1

        L("Manuell (WNK-Begriff):", r)
        self.v_man = tk.StringVar()
        _combo = AutocompleteCombo(f, textvariable=self.v_man)
        _combo.grid(row=r, column=1, columnspan=2, sticky="ew", padx=4, pady=2)
        hint(f, "Tippen → WNK-Begriffe werden vorgeschlagen", r)
        r += 1

        L("Von (Jahr):", r)
        self.v_vonj = tk.StringVar()
        ttk.Entry(f, textvariable=self.v_vonj, width=10).grid(
            row=r, column=1, sticky="w", padx=4, pady=2)
        ttk.Label(f, text="Bis (Jahr):").grid(row=r, column=2, sticky="w")
        self.v_bisj = tk.StringVar()
        ttk.Entry(f, textvariable=self.v_bisj, width=10).grid(
            row=r, column=3, sticky="w", padx=2)
        hint(f, "z. B.  -27  bis  450", r, col=4)
        r += 1


        SEP(r); r += 1
        L("── Absolute Ankerpunkte ──", r, bold=True); r += 1

        # Spaltenüberschriften
        ttk.Label(f, text="Ereignis",
                  foreground="#999999", font=("", 8)).grid(
            row=r, column=1, sticky="w", padx=4)
        ttk.Label(f, text="Jahr (pos. Zahl)",
                  foreground="#999999", font=("", 8)).grid(
            row=r, column=2, sticky="w", padx=2)
        r += 1

        self._ap_vars: Dict[str, dict] = {}
        for art, lbl_t in [("tpq",   "TPQ – jünger als:"),
                            ("taq",   "TAQ – älter als:"),
                            ("exakt", "Exaktes Ereignis:")]:
            L(lbl_t, r)
            vb   = tk.StringVar()
            vj   = tk.StringVar()
            vchr = tk.BooleanVar()
            self._ap_vars[art] = {"b": vb, "j": vj, "vchr": vchr}

            ttk.Entry(f, textvariable=vb, width=24).grid(
                row=r, column=1, sticky="ew", padx=4, pady=2)
            ttk.Entry(f, textvariable=vj, width=7).grid(
                row=r, column=2, sticky="w", padx=2)
            ttk.Checkbutton(f, text="v. Chr.", variable=vchr).grid(
                row=r, column=3, sticky="w", padx=2)
            hint(f, 'z. B. "Vesuvausbrauch" | 79', r, col=4)
            r += 1
        SEP(r); r += 1
        L("── Stratigraphie / Harris ──", r, bold=True); r += 1

        self._h_vars: Dict[str, tk.StringVar] = {}
        for attr, lbl_t, h_txt in [
            ("juenger_als",    "Jünger als (liegt über):", "z. B.  169, 170"),
            ("aelter_als",     "Älter als (liegt unter):", "z. B.  30"),
            ("gleich_alt_wie", "Gleich alt wie:",           "z. B.  34"),
            ("gestoert_von",   "Gestört von:",              "z. B.  45"),
        ]:
            L(lbl_t, r)
            var = tk.StringVar()
            self._h_vars[attr] = var
            ttk.Entry(f, textvariable=var, width=30).grid(
                row=r, column=1, columnspan=2, sticky="ew", padx=4, pady=2)
            hint(f, h_txt, r, col=3)
            r += 1

        # Datierungsbasierte Vorschläge (read-only, werden nicht in Felder gemischt)
        SEP(r); r += 1
        L("Aus Profiltext erkannt:", r, bold=True); r += 1
        self.txt_profil = tk.Text(f, height=3, width=56, state="disabled",
                                   background="#f0f8ff", relief="flat",
                                   borderwidth=1)
        self.txt_profil.grid(row=r, column=0, columnspan=5,
                              sticky="ew", padx=6, pady=2)
        r += 1

        # "Bekannte Beziehungen" entfernt – bidir-Info erscheint jetzt
        # direkt in den Feldern "Jünger als" / "Älter als".

        SEP(r); r += 1
        L("Zeichnung(en):", r)
        self.v_zeich = tk.StringVar()
        E(self.v_zeich, r)
        hint(f, "z. B.  1-15  (aus Katalog)", r)
        r += 1

        L("Kommentar:", r)
        self.txt_kom = tk.Text(f, height=3, width=42)
        self.txt_kom.grid(row=r, column=1, columnspan=3,
                           sticky="ew", padx=4, pady=2)
        r += 1

        SEP(r); r += 1
        L("Datierungsprotokoll:", r, bold=True); r += 1
        self.txt_protokoll = scrolledtext.ScrolledText(
            f, height=5, width=56, state="disabled",
            font=("Courier", 8), wrap=tk.WORD,
            background="#1a1a2a", foreground="#a0c8ff",
            relief="flat", borderwidth=1)
        self.txt_protokoll.grid(row=r, column=0, columnspan=5,
                                sticky="ew", padx=6, pady=2)
        r += 1
        rr = ttk.Frame(f)
        rr.grid(row=r, column=0, columnspan=5, sticky="w", padx=6, pady=2)
        ttk.Label(rr, text="Stratigr. zurücksetzen:").pack(side=tk.LEFT)
        ttk.Button(rr, text="Diese Stelle",
                   command=self._reset_strat_stelle).pack(side=tk.LEFT, padx=3)
        ttk.Button(rr, text="Verbundene Gruppe",
                   command=self._reset_strat_gruppe).pack(side=tk.LEFT, padx=3)
        ttk.Button(rr, text="Alle",
                   command=self._reset_strat_alle).pack(side=tk.LEFT, padx=3)
        r += 1

        L("Funde:", r, bold=True); r += 1
        self.txt_funde = tk.Text(f, height=5, width=56, state="disabled",
                                  background="#f8f8f8", relief="flat",
                                  borderwidth=1)
        self.txt_funde.grid(row=r, column=0, columnspan=5,
                             sticky="ew", padx=6, pady=2)
        r += 1

        bf = ttk.Frame(f)
        bf.grid(row=r, column=0, columnspan=5, pady=8)
        ttk.Button(bf, text="↩ Zurücksetzen", command=self._reset_editor).pack(side=tk.LEFT, padx=5)
        ttk.Button(bf, text="🗑 Löschen",      command=self._delete).pack(side=tk.LEFT, padx=5)
        f.columnconfigure(1, weight=3)   # Hauptspalte Eingabefelder
        f.columnconfigure(2, weight=1)   # zweite Eingabespalte
        f.columnconfigure(3, weight=1)   # Hinweis-Spalte
        # Auto-Save: Traces auf alle editierbaren Felder
        self._debounce_id = None
        self._saving = False
        _autosave_vars = ([self.v_ansp, self.v_man, self.v_vonj,
                           self.v_bisj, self.v_zeich]
                          + list(self._h_vars.values())
                          + [av[k] for av in self._ap_vars.values()
                             for k in ("b","j")]
                          + [self.v_ab, self.v_st, self.v_rb])
        for _v in _autosave_vars:
            _v.trace_add("write", self._sched_save)

    # ── Liste aktualisieren ───────────────────────────────────
    def aktualisiere(self):
        # Binding lösen damit selection_set kein _on_sel auslöst.
        # self._cur ist autoritativ — immer diese Stelle selektieren.
        self.tree.unbind("<<TreeviewSelect>>")
        try:
            s_txt = self.sv.get().lower()
            self.tree.delete(*self.tree.get_children())
            warn_nrs = {w.stelle for w in self.app.warnungen if w.stufe == 1}
            for nr in sorted(self.app.stellen, key=_sort_key):
                s = self.app.stellen[nr]
                if s_txt and s_txt not in nr.lower() \
                         and s_txt not in s.befansprac.lower():
                    continue
                flags = []
                if s.kein_befund:      flags.append("AB")
                if s.moderne_stoerung: flags.append("ST")
                if s.durchschossen:    flags.append("DU")
                if s.okuk_juenger_als: flags.append("OKUK")
                if s.profil_juenger_als or s.profil_aelter_als:
                    flags.append("PROF")
                tag = ""
                if nr in warn_nrs:                                      tag = "wid"
                elif s.kein_befund:                                     tag = "ab"
                elif s.moderne_stoerung:                                tag = "st"
                elif s.inferred_datierung_str and not s.auto_datierung_str:
                    tag = "inf"
                self.tree.insert("", tk.END, iid=nr, tags=(tag,), values=(
                    nr,
                    s.befansprac[:25],
                    s.effektiv_datierung_str(),
                    s.datierung_quelle(),
                    len(s.funde),
                    " ".join(flags),
                ))
            # Selektion auf self._cur setzen (autoritativ)
            if self._cur and self.tree.exists(self._cur):
                self.tree.selection_set(self._cur)
                self.tree.see(self._cur)
        finally:
            self.tree.bind("<<TreeviewSelect>>", self._on_sel)

    # ── Editor befüllen ───────────────────────────────────────
    def _on_sel(self, _=None):
        # Selektion VOR dem Speichern lesen — _save kann aktualisiere() triggern
        # und dabei die Selektion auf self._cur zurücksetzen.
        sel = self.tree.selection()
        new_nr = sel[0] if sel else None
        if not new_nr or new_nr == self._cur:
            return
        if self._cur and self._cur in self.app.stellen:
            # Stille Übergangsspeicherung: Felder → Modell, KEIN Listenrebuild
            self._save(silent=True)
        self._cur = new_nr          # Erst jetzt setzen → aktualisiere() nutzt new_nr
        self._load_editor()

    def _load_editor(self):
        nr = self._cur
        if not nr or nr not in self.app.stellen:
            return
        # Traces während des Ladens unterdrücken (verhindert Phantomspeicherungen)
        self._saving = True
        try:
            self._load_editor_inner(nr)
        finally:
            self._saving = False

    def _load_editor_inner(self, nr):
        s = self.app.stellen[nr]
        # Snapshot für Reset (einmalig beim ersten Laden dieser Stelle)
        if nr not in self._snapshots:
            self._snapshots[nr] = {
                'j': list(s.juenger_als), 'a': list(s.aelter_als),
                'g': list(s.gleich_alt_wie), 'gv': list(s.gestoert_von),
                'md': s.manuell_datierung, 'mv': s.manuell_von, 'mb': s.manuell_bis,
                'ba': s.befansprac, 'kb': s.kein_befund, 'ms': s.moderne_stoerung,
                'rb': s.rest_befund,
                'ze': s.zeichnung, 'ko': s.kommentar,
            }

        self.v_nr.set(s.stellennr)
        self.v_ansp.set(s.befansprac)
        self.v_ab.set(s.kein_befund)
        self.v_st.set(s.moderne_stoerung)
        self.v_rb.set(s.rest_befund)
        self.v_sich.set(s.sichtbar)
        self.v_ok.set(f"{s.tiefe_ok:.3f}" if s.tiefe_ok is not None else "")
        self.v_uk.set(f"{s.tiefe_uk:.3f}" if s.tiefe_uk is not None else "")
        self.v_aut.set(s.auto_datierung_str or s.datierung_katalog or "")
        # Stratigraph. Hinweis: vollständiger Text mit ≤/≥ und Quelle
        self.v_inf.set(s.inferred_datierung_str)
        self.v_man.set(s.manuell_datierung)
        # Nur manuelle Jahre anzeigen – auto_von/bis sind im Feld "Auto" sichtbar
        self.v_vonj.set(str(s.manuell_von) if s.manuell_von is not None else "")
        self.v_bisj.set(str(s.manuell_bis) if s.manuell_bis is not None else "")
        self.v_zeich.set(s.zeichnung)

        # Editorfelder: manuell + dat_* + bidir kombiniert.
        # _bidir_* = Beziehungen, die aus anderen Stellen umgekehrt abgeleitet wurden
        # (z.B. 104 weiß durch _bidir_aelter_als, dass 102 in es eingetieft ist).
        # _auto_im_feld merkt sich, welche Einträge beim Laden NUR automatisch
        # waren — _save nutzt das, um sie nicht zu manuellen zu befördern.
        self._auto_im_feld = {a: set() for a in self._h_vars}

        def _combined_rel(manual_attr, *auto_attrs):
            m = list(getattr(s, manual_attr))
            seen = set(m)
            autos = []
            for aa in auto_attrs:
                for x in getattr(s, aa):
                    if x not in seen:
                        seen.add(x); autos.append(x)
            self._auto_im_feld[manual_attr] = set(autos)
            return ", ".join(m + autos)

        for attr, var in self._h_vars.items():
            if attr == "juenger_als":
                var.set(_combined_rel("juenger_als", "dat_juenger_als",
                                      "_bidir_juenger_als"))
            elif attr == "aelter_als":
                var.set(_combined_rel("aelter_als", "dat_aelter_als",
                                      "_bidir_aelter_als"))
            elif attr == "gleich_alt_wie":
                var.set(_combined_rel("gleich_alt_wie", "_bidir_gleich"))
            else:
                var.set(", ".join(getattr(s, attr)))

        for art, av in self._ap_vars.items():
            ap = next((a for a in s.ankerpunkte if a.art == art), None)
            av["b"].set(ap.beschreibung if ap else "")
            if ap and ap.von is not None:
                av["j"].set(str(abs(ap.von)))   # immer positiv anzeigen
                av["vchr"].set(ap.von < 0)       # Haken wenn v. Chr.
            else:
                av["j"].set("")
                av["vchr"].set(False)

        # Profiltext-Beziehungen
        self.txt_profil.config(state="normal")
        self.txt_profil.delete("1.0", tk.END)
        pl = []
        # Eigene Profil-Beziehungen
        if s.profil_juenger_als:
            pl.append(f"↑ Jünger als:    {', '.join(s.profil_juenger_als)}")
        if s.profil_aelter_als:
            pl.append(f"↓ Älter als:     {', '.join(s.profil_aelter_als)}")
        if s.profil_gleich:
            pl.append(f"= Gleich alt:    {', '.join(s.profil_gleich)}")
        # Referenzen aus dem Profiltext ANDERER Stellen auf diese Stelle
        # (z. B. "eingetieft in 104" bei 102 bedeutet: 104 ist älter als 102)
        ref_ae = sorted([onr for onr, os in self.app.stellen.items()
                         if onr != nr and nr in os.profil_juenger_als],
                        key=_sort_key)
        ref_ju = sorted([onr for onr, os in self.app.stellen.items()
                         if onr != nr and nr in os.profil_aelter_als],
                        key=_sort_key)
        ref_gl = sorted([onr for onr, os in self.app.stellen.items()
                         if onr != nr and nr in os.profil_gleich],
                        key=_sort_key)
        if ref_ae or ref_ju or ref_gl:
            pl.append("── Aus Profiltexten anderer Stellen: ──")
        if ref_ae:
            pl.append(f"↓ Diese Stelle ist älter als:  {', '.join(ref_ae)}")
        if ref_ju:
            pl.append(f"↑ Diese Stelle ist jünger als: {', '.join(ref_ju)}")
        if ref_gl:
            pl.append(f"= Gleich alt wie:              {', '.join(ref_gl)}")
        self.txt_profil.insert("1.0",
            "\n".join(pl) if pl else "(keine Beziehungen im Profiltext erkannt)")
        self.txt_profil.config(state="disabled")

        self.txt_kom.delete("1.0", tk.END)
        self.txt_kom.insert("1.0", s.kommentar)

        self.txt_protokoll.config(state="normal")
        self.txt_protokoll.delete("1.0", tk.END)
        self.txt_protokoll.insert("1.0", s.dat_protokoll
                                  or "(Analyse ausführen um Protokoll zu erstellen)")
        self.txt_protokoll.config(state="disabled")

        self.txt_funde.config(state="normal")
        self.txt_funde.delete("1.0", tk.END)
        if s.funde:
            for f in s.funde:
                self.txt_funde.insert(
                    tk.END, f"Pos {f.posnr}/{f.unternr}: {f.kurztext()}\n")
        else:
            self.txt_funde.insert(tk.END, "(keine Funde in Fundliste)")
        self.txt_funde.config(state="disabled")

    # ── Editor speichern ──────────────────────────────────────
    def _save(self, silent=False):
        nr = self._cur
        if not nr or nr not in self.app.stellen:
            return
        # Nur explizite Speicherung (kein Auto-Save) sichert Undo-Schritt
        if not silent:
            self.app._push_undo()
        s = self.app.stellen[nr]
        s.befansprac       = self.v_ansp.get().strip()
        s.kein_befund      = self.v_ab.get()
        s.moderne_stoerung = self.v_st.get()
        s.rest_befund      = self.v_rb.get()
        s.manuell_datierung = self.v_man.get().strip()
        s.zeichnung        = self.v_zeich.get().strip()
        s.kommentar        = self.txt_kom.get("1.0", tk.END).strip()

        def _j(v):
            try:    return int(v.get().strip()) if v.get().strip() else None
            except: return None
        s.manuell_von = _j(self.v_vonj)
        s.manuell_bis = _j(self.v_bisj)

        # Beziehungsfelder zurücklesen. Auto-Einträge (dat_*/_bidir_*) werden
        # NICHT zu manuellen befördert: was beim Laden als Auto angezeigt wurde
        # und noch im Feld steht, bleibt Auto. Was der Nutzer davon gelöscht
        # hat, wird unterdrückt (bleibt auch nach F5/Neuberechnung weg).
        # Gelöschte manuelle Einträge werden auch auf der Gegenstelle entfernt,
        # sonst würde die Spiegelung sie sofort wiederbeleben.
        _richtung = {"juenger_als": "juenger", "aelter_als": "aelter",
                     "gleich_alt_wie": "gleich"}
        _gegen    = {"juenger_als": "aelter_als", "aelter_als": "juenger_als",
                     "gleich_alt_wie": "gleich_alt_wie"}
        auto_im_feld = getattr(self, "_auto_im_feld", {})
        neu_eintraege = []   # (richtung, other) – gerade NEU getippte Beziehungen
        for attr, var in self._h_vars.items():
            feld = parse_liste(var.get())
            if attr == "gestoert_von":
                s.gestoert_von = feld
                continue
            autos = auto_im_feld.get(attr, set())
            alt_manuell = list(getattr(s, attr))
            neu_manuell = [x for x in feld if x not in autos and x != nr]
            geloescht_auto    = autos - set(feld)
            geloescht_manuell = [x for x in alt_manuell if x not in feld]
            for x in geloescht_auto | set(geloescht_manuell):
                # Manuellen Gegen-Eintrag beim Peer entfernen (falls vorhanden)
                if x in self.app.stellen:
                    gegen_liste = getattr(self.app.stellen[x], _gegen[attr])
                    if nr in gegen_liste:
                        gegen_liste.remove(nr)
                # Zusätzlich unterdrücken: hält auch dat_/Spiegel-Ableitungen fern
                key = f"{_richtung[attr]}:{x}"
                if key not in s.unterdrueckte_bez:
                    s.unterdrueckte_bez.append(key)
            for x in neu_manuell:
                # Wieder eingetragen → Unterdrückung aufheben
                key = f"{_richtung[attr]}:{x}"
                if key in s.unterdrueckte_bez:
                    s.unterdrueckte_bez.remove(key)
                # neu im Feld (war vorher nicht manuell) → "letzte Eingabe gewinnt"
                if x not in alt_manuell:
                    neu_eintraege.append((_richtung[attr], x))
            setattr(s, attr, neu_manuell)
        # "Letzte Eingabe gewinnt": eine neu getippte Beziehung räumt jede
        # widersprechende ältere Beziehung zwischen demselben Stellenpaar weg
        # (auf beiden Seiten). Beispiel: '5 gleich 6' bestand, Nutzer tippt bei
        # 6 'älter als 5' → die alte Gleichsetzung wird entfernt.
        for richtung, other in neu_eintraege:
            beziehung_exklusiv(self.app.stellen, nr, other, richtung)
        # Sicherheitsnetz: verbliebene Mehrfach-Einträge desselben Ziels
        # (sollten nach beziehung_exklusiv nicht auftreten) deterministisch lösen
        seen_ex = set()
        for _attr in ["juenger_als","aelter_als","gleich_alt_wie"]:
            _vals = getattr(s, _attr)
            _clean = [x for x in _vals if x not in seen_ex]
            if _clean != _vals:
                setattr(s, _attr, _clean)
            seen_ex.update(_clean)
        # Ankerpunkte VOR der Neuberechnung übernehmen — sonst rechnet die
        # Propagierung mit den alten TPQ/TAQ-Werten und die Auswertung zeigt
        # bis zum nächsten Speichern veraltete Datierungen
        s.ankerpunkte = []
        for art, av in self._ap_vars.items():
            b  = av["b"].get().strip()
            js = av["j"].get().strip()
            if b or js:
                try:
                    jahr = int(js) if js else None
                    if jahr is not None and av["vchr"].get():
                        jahr = -abs(jahr)
                except ValueError:
                    jahr = None
                s.ankerpunkte.append(
                    Ankerpunkt(art=art, beschreibung=b,
                               von=jahr, praezision="Jahr"))
        # Stille Speicherung: Propagierung ja, Stellenliste NICHT neu aufbauen
        # (würde _cur-Selektion zurücksetzen und _on_sel-Logik stören)
        self.app._nach_aenderung(harris=False,
                                  update_stellen_list=not silent)

        if not silent:
            # _nach_aenderung wurde bereits oben aufgerufen (silent=False-Pfad)
            self.aktualisiere()
            try: self.tree.selection_set(nr)
            except: pass
            self._load_editor()
            self.app.status(f"Stelle {nr} gespeichert.")
            # Snapshot nach explizitem Speichern aktualisieren
            self._snapshots[nr] = {
                'j': list(s.juenger_als), 'a': list(s.aelter_als),
                'g': list(s.gleich_alt_wie), 'gv': list(s.gestoert_von),
                'md': s.manuell_datierung, 'mv': s.manuell_von, 'mb': s.manuell_bis,
                'ba': s.befansprac, 'kb': s.kein_befund, 'ms': s.moderne_stoerung,
                'rb': s.rest_befund,
                'ze': s.zeichnung, 'ko': s.kommentar,
            }
        # Harris-Matrix IMMER auffrischen — auch bei stiller Speicherung
        # (Texteingabe/Tab-Wechsel), damit die Grafik Änderungen sofort zeigt
        try: self.app.tab_harris.aktualisiere()
        except: pass

    def _reset_editor(self):
        """Leert ALLE manuellen Felder sofort (kein Dialog)."""
        nr = self._cur
        if not nr or nr not in self.app.stellen:
            return
        self.app._push_undo()
        s = self.app.stellen[nr]
        # Manuelle Gegen-Einträge der Peers mit entfernen (gezeichnete Linien
        # stehen auf beiden Seiten — sonst kommt die Beziehung als Spiegel zurück)
        for attr, gegen in [("juenger_als", "aelter_als"),
                            ("aelter_als", "juenger_als"),
                            ("gleich_alt_wie", "gleich_alt_wie")]:
            for x in getattr(s, attr):
                peer = self.app.stellen.get(x)
                if peer:
                    gl = getattr(peer, gegen)
                    if nr in gl:
                        setattr(peer, gegen, [y for y in gl if y != nr])
        s.juenger_als       = []
        s.aelter_als        = []
        s.gleich_alt_wie    = []
        s.gestoert_von      = []
        s.unterdrueckte_bez = []   # zurück auf reinen Automatik-Stand
        s.manuell_datierung = ""
        s.manuell_von       = None
        s.manuell_bis       = None
        s.ankerpunkte       = []
        s.zeichnung         = ""
        s.kommentar         = ""
        # Snapshot entfernen, damit kein alter Stand zurückkommt
        self._snapshots.pop(nr, None)
        self.app._aktualisiere_alles()
        self._load_editor()
        self.app.status(f"Stelle {nr}: Alle manuellen Eingaben geleert (Analyse aktualisiert).")


    def _reset_strat_stelle(self):
        nr = self._cur
        if not nr or nr not in self.app.stellen: return
        s = self.app.stellen[nr]
        s.strat_von = None; s.strat_bis = None
        self.app._nach_aenderung()
        self._load_editor_inner(nr)
        self.app.status(f"Stelle {nr}: Stratigraphie-Korrekturen zurückgesetzt.")

    def _reset_strat_gruppe(self):
        nr = self._cur
        if not nr or nr not in self.app.stellen: return
        besucht = set(); queue = [nr]
        while queue:
            cur = queue.pop(0)
            if cur in besucht: continue
            besucht.add(cur)
            if cur not in self.app.stellen: continue
            s = self.app.stellen[cur]
            s.strat_von = None; s.strat_bis = None
            for xnr in (s.alle_juenger_als() + s.alle_aelter_als()
                        + s.alle_gleich_alt_wie()):
                if xnr not in besucht: queue.append(xnr)
        self.app._nach_aenderung()
        self._load_editor_inner(nr)
        self.app.status(f"Gruppe ({len(besucht)} Stellen) zurückgesetzt.")

    def _reset_strat_alle(self):
        for s in self.app.stellen.values():
            s.strat_von = None; s.strat_bis = None
        self.app._nach_aenderung()
        if self._cur: self._load_editor_inner(self._cur)
        self.app.status("Alle Stratigraphie-Korrekturen zurückgesetzt.")

    def _sched_save(self, *_):
        if self._saving: return
        if self._debounce_id:
            try: self.after_cancel(self._debounce_id)
            except: pass
        self._debounce_id = self.after(600, self._do_autosave)

    def _do_autosave(self):
        self._debounce_id = None
        if self._cur and self._cur in self.app.stellen:
            self._saving = True
            try: self._save(silent=True)
            finally: self._saving = False

    def _neue_stelle(self):
        d = tk.Toplevel(self)
        d.title("Neue Stelle")
        d.geometry("300x110")
        d.grab_set()
        ttk.Label(d, text="Stellennummer (z. B. 17, 100d, Schnitt-A):").pack(pady=8)
        v = tk.StringVar()
        e = ttk.Entry(d, textvariable=v, width=20)
        e.pack()
        e.focus_set()
        def ok():
            nr = norm_stellennr(v.get().strip())
            if not nr: return
            if nr in self.app.stellen:
                messagebox.showinfo("Vorhanden",
                    f"Stelle '{nr}' existiert bereits.", parent=d)
                d.destroy(); return
            self.app.stellen[nr] = Stelle(stellennr=nr)
            self.aktualisiere()
            try: self.tree.selection_set(nr)
            except: pass
            self._on_sel()
            d.destroy()
        ttk.Button(d, text="Anlegen", command=ok).pack(pady=4)
        d.bind("<Return>", lambda _: ok())
        self.app.style_dialog(d)

    def _bereich(self):
        try:
            von, bis = int(self.v_von.get()), int(self.v_bis.get())
        except:
            messagebox.showerror("Eingabe", "Bitte Ganzzahlen eingeben.")
            return
        neu = 0
        for i in range(von, bis + 1):
            if str(i) not in self.app.stellen:
                self.app.stellen[str(i)] = Stelle(stellennr=str(i))
                neu += 1
        self.aktualisiere()
        self.app.status(f"Bereich {von}–{bis}: {neu} neue Stellen.")

    def _delete(self):
        nr = self._cur
        if nr and messagebox.askyesno("Löschen", f"Stelle '{nr}' löschen?"):
            self.app._push_undo()
            del self.app.stellen[nr]
            # Verweise auf die gelöschte Nummer überall mit entfernen — sonst
            # bleiben Geister-Einträge, die sich an eine später neu angelegte
            # Stelle gleicher Nummer heften würden
            _rel = ("juenger_als","aelter_als","gleich_alt_wie","gestoert_von",
                    "profil_juenger_als","profil_aelter_als","profil_gleich",
                    "okuk_juenger_als","okuk_aelter_als",
                    "dat_juenger_als","dat_aelter_als")
            for other in self.app.stellen.values():
                for attr in _rel:
                    lst = getattr(other, attr)
                    if nr in lst:
                        setattr(other, attr, [x for x in lst if x != nr])
                other.unterdrueckte_bez = [
                    e for e in other.unterdrueckte_bez
                    if ":" not in e or e.split(":", 1)[1] != nr]
            self._cur = None
            self.app._aktualisiere_alles()
            self.aktualisiere()


    def _copy_tree(self, event=None):
        tree = getattr(self, 'tree', None)
        if tree is None:
            return
        sel = tree.selection()
        if not sel:
            sel = tree.get_children()
        cols = tree['columns']
        header = '\t'.join(str(tree.heading(c)['text']) for c in cols)
        zeilen = [header]
        for iid in sel:
            vals = tree.item(iid, 'values')
            zeilen.append('\t'.join(str(v) for v in vals))
        self.clipboard_clear()
        self.clipboard_append('\n'.join(zeilen))

    def _okuk_dialog(self):
        mit_ok = [(nr, s.tiefe_ok) for nr, s in self.app.stellen.items()
                  if s.tiefe_ok is not None
                  and not s.kein_befund and not s.moderne_stoerung]
        if not mit_ok:
            messagebox.showinfo("OK/UK",
                "Keine Befundstellen mit TiefeOK-Werten.\n"
                "Bitte zuerst Stellenkatalog laden.")
            return
        d = tk.Toplevel(self)
        d.title("Harris-Vorschläge aus OK/UK")
        d.geometry("540x430")
        d.grab_set()
        ttk.Label(d, text=f"{len(mit_ok)} Befundstellen mit TiefeOK.\n"
                          "Wähle Anwendungsbereich:",
                  wraplength=490).pack(pady=8, padx=10)
        mode = tk.StringVar(value="alle")
        ttk.Radiobutton(d, text="Alle Befundstellen mit OK-Wert",
                        variable=mode, value="alle").pack(anchor="w", padx=20)
        ttk.Radiobutton(d, text="Nur Nummernbereich:",
                        variable=mode, value="auswahl").pack(anchor="w", padx=20)
        af = ttk.Frame(d); af.pack(fill=tk.X, padx=40)
        v2 = tk.StringVar(value="1"); b2 = tk.StringVar(value="9999")
        ttk.Label(af, text="Von:").pack(side=tk.LEFT)
        ttk.Entry(af, textvariable=v2, width=7).pack(side=tk.LEFT, padx=2)
        ttk.Label(af, text="Bis:").pack(side=tk.LEFT)
        ttk.Entry(af, textvariable=b2, width=7).pack(side=tk.LEFT, padx=2)
        ttk.Label(d, text="Vorschau (höchste OK = jüngste):",
                  foreground="#555555").pack(pady=(8, 2), padx=10, anchor="w")
        pf = ttk.Frame(d); pf.pack(fill=tk.BOTH, expand=True, padx=10)
        pt = _mache_tree(pf, ("Nr","TiefeOK","TiefeUK"),
                         ("Stelle","OK (m)","UK (m)"), [80,90,90], height=7)
        for nr, ok in sorted(mit_ok, key=lambda x: x[1], reverse=True):
            s = self.app.stellen[nr]
            pt.insert("", tk.END, values=(
                nr, f"{ok:.3f}",
                f"{s.tiefe_uk:.3f}" if s.tiefe_uk is not None else "–"))
        ttk.Label(d, text="⚠️  Vorläufig! Bitte manuell prüfen.",
                  foreground="#cc6600").pack(pady=4)
        def anwenden():
            auswahl = None
            if mode.get() == "auswahl":
                try: auswahl = [str(i) for i in range(int(v2.get()), int(b2.get())+1)]
                except: pass
            n = okuk_beziehungen_berechnen(self.app.stellen, auswahl)
            propagiere_datierungen(self.app.stellen)
            self.aktualisiere()
            self.app.status(f"OK/UK: {n} neue vorläufige Beziehungen.")
            d.destroy()
        ttk.Button(d, text="Beziehungen ableiten", command=anwenden).pack(pady=4)
        ttk.Button(d, text="Abbrechen", command=d.destroy).pack()
        self.app.style_dialog(d)

# ═══════════════════════════════════════════════════════════════
# TAB: AUSWERTUNG
# ═══════════════════════════════════════════════════════════════

class AuswertungTab(ttk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self._sortiert: List[Stelle] = []
        self._warnungen: List[Warnung] = []
        self._build()

    def _build(self):
        tb = ttk.Frame(self)
        tb.pack(fill=tk.X, padx=6, pady=4)
        ttk.Button(tb, text="Exportieren…",
                   command=self.app.cmd_exportiere).pack(side=tk.LEFT, padx=2)
        ttk.Label(tb, text="Zeige:").pack(side=tk.LEFT, padx=(16, 2))
        self.v_ab = tk.BooleanVar(value=True)
        self.v_st = tk.BooleanVar(value=True)
        ttk.Checkbutton(tb, text="Arbeitsbereiche", variable=self.v_ab,
                        command=self._filter).pack(side=tk.LEFT)
        ttk.Checkbutton(tb, text="Störungen", variable=self.v_st,
                        command=self._filter).pack(side=tk.LEFT)

        paned = ttk.PanedWindow(self, orient=tk.VERTICAL)
        paned.pack(fill=tk.BOTH, expand=True, padx=6, pady=2)

        tbl = ttk.LabelFrame(paned, text="Chronologische Reihenfolge  (alt → jung)")
        paned.add(tbl, weight=3)
        cols = ("Stelle","Ansprache","Datierung","Quelle","Von","Bis",
                "Ältester Fund","Jüngster Fund","Beziehungen","Flags")
        self.tree = _mache_tree(tbl, cols, cols,
                                [60,150,210,90,52,52,220,220,200,75])
        self.tree.tag_configure("wid", background="#ffcccc")
        self.tree.tag_configure("hin", background="#fff3cd")
        self.tree.tag_configure("not", background="#0d3320", foreground="#a0e0b0")
        self.tree.bind("<Double-1>", self._on_dblclick)
        self.tree.tag_configure("ab",  foreground="#888888")

        wf = ttk.LabelFrame(paned, text="Warnungen & Hinweise")
        paned.add(wf, weight=1)
        self.warn_tree = _mache_tree(
            wf, ("Stufe","Stelle","Nachricht"),
                ("Stufe","Stelle","Nachricht"), [130, 60, 720])
        self.warn_tree.tag_configure("wid", background="#ffcccc")
        self.warn_tree.tag_configure("hin", background="#fff3cd")
        self.warn_tree.tag_configure("not", background="#0d3320" ,  foreground="#a0e0b0")

    def _copy_tree(self, event=None):
        tree = getattr(self, 'tree', None)
        if tree is None:
            return
        sel = tree.selection()
        if not sel:
            sel = tree.get_children()
        cols = tree['columns']
        header = '\t'.join(str(tree.heading(c)['text']) for c in cols)
        zeilen = [header]
        for iid in sel:
            vals = tree.item(iid, 'values')
            zeilen.append('\t'.join(str(v) for v in vals))
        self.clipboard_clear()
        self.clipboard_append('\n'.join(zeilen))

    def _on_dblclick(self, event=None):
        """Doppelklick auf Auswertungszeile → Stelle in Stellen-Tab öffnen."""
        sel = self.tree.selection()
        if not sel: return
        nr = str(self.tree.item(sel[0], "values")[0])
        try:
            self.app.notebook.select(self.app.tab_stellen)
            if self.app.tab_stellen.tree.exists(nr):
                self.app.tab_stellen.tree.selection_set(nr)
                self.app.tab_stellen.tree.see(nr)
                self.app.tab_stellen._on_sel()
        except Exception: pass

    def aktualisiere(self, sortiert, warnungen):
        self._sortiert = sortiert
        self._warnungen = warnungen
        self._filter()
        self._zeige_warn()

    def _filter(self):
        self.tree.delete(*self.tree.get_children())
        ws = {w.stelle: min(w2.stufe for w2 in self._warnungen
                            if w2.stelle == w.stelle)
              for w in self._warnungen}
        for s in self._sortiert:
            if not self.v_ab.get() and s.kein_befund:      continue
            if not self.v_st.get() and s.moderne_stoerung: continue
            af = jf = ""
            md = [f for f in s.funde
                  if f.von is not None or f.bis is not None]
            if md:
                a = min(md, key=lambda f: f.von if f.von is not None else 99999)
                j = max(md, key=lambda f: f.bis if f.bis is not None else -99999)
                af = a.kurztext()[:55]
                jf = j.kurztext()[:55] if j != a else ""
            bez = []
            if s.juenger_als:        bez.append("↑ " + ", ".join(s.juenger_als[:4]))
            if s.aelter_als:         bez.append("↓ " + ", ".join(s.aelter_als[:4]))
            if s.gleich_alt_wie:     bez.append("= " + ", ".join(s.gleich_alt_wie[:4]))
            if s.profil_juenger_als: bez.append("↑P " + ", ".join(s.profil_juenger_als[:3]))
            if s.okuk_juenger_als:   bez.append("⊥OK " + ", ".join(s.okuk_juenger_als[:3]))
            flags = " ".join(
                (["AB"] if s.kein_befund      else []) +
                (["ST"] if s.moderne_stoerung else []) +
                (["DU"] if s.durchschossen    else []) +
                (["OKUK"] if s.okuk_juenger_als else []) +
                (["PROF"] if s.profil_juenger_als or s.profil_aelter_als else [])
            )
            ms = ws.get(s.stellennr, 0)
            tag = {1: "wid", 2: "hin", 3: "not"}.get(ms, "")
            if s.kein_befund and not tag: tag = "ab"
            self.tree.insert("", tk.END, tags=(tag,) if tag else (), values=(
                s.stellennr, s.befansprac[:30],
                s.effektiv_datierung_str(),
                s.datierung_quelle(),
                s.effektiv_von()  if s.effektiv_von()  is not None else "",
                s.effektiv_bis() if s.effektiv_bis() is not None else "",
                af, jf, " | ".join(bez), flags,
            ))

    def _zeige_warn(self):
        self.warn_tree.delete(*self.warn_tree.get_children())
        tm = {1: "wid", 2: "hin", 3: "not"}
        for w in self._warnungen:
            self.warn_tree.insert("", tk.END, tags=(tm.get(w.stufe, ""),),
                                  values=(w.stufe_str(), w.stelle, w.nachricht))

    def exportiere(self, pfad: Path):
        with open(pfad, "w", encoding="utf-8-sig", newline="") as f:
            wr = csv.writer(f, delimiter=DELIMITER)
            wr.writerow(["Stelle","Ansprache","Datierung","Quelle","Von","Bis",
                         "Ältester Fund","Jüngster Fund","Beziehungen","Flags"])
            for iid in self.tree.get_children():
                wr.writerow(self.tree.item(iid, "values"))
            wr.writerow([])
            wr.writerow(["WARNUNGEN"])
            wr.writerow(["Stufe","Stelle","Nachricht"])
            for w in self._warnungen:
                wr.writerow([w.stufe_str(), w.stelle, w.nachricht])

# ═══════════════════════════════════════════════════════════════
# TAB: HARRIS-MATRIX (GRAFISCH)
# Älteste Stelle UNTEN, jüngste OBEN – klassische Darstellung
# Nicht eingeordnete Stellen erscheinen in Seitenleiste rechts
# ═══════════════════════════════════════════════════════════════

class HarrisMatrixTab(ttk.Frame):
    BW = 128; BH = 44; HG = 20; VG = 52

    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self._pos: Dict[str, Tuple[int, int]] = {}
        self._drag_nr = None
        self._drag_start_pos = None
        self._drag_offset = (0, 0)
        self._drag_timer = None
        self._line_colors: dict = {}
        self._warn1: set = set()
        self._popup_win  = None
        self._draw_mode  = None
        self._draw_src   = None
        self._draw_just_done = False
        self._build()

    def _build(self):
        tb = ttk.Frame(self)
        tb.pack(fill=tk.X, padx=6, pady=4)
        ttk.Button(tb, text="Matrix aktualisieren",
                   command=self._aktualisiere_neu_sortiert).pack(side=tk.LEFT, padx=2)
        ttk.Button(tb, text="PNG exportieren…",
                   command=self._png_export).pack(side=tk.LEFT, padx=2)
        ttk.Button(tb, text="Stilregeln…",
                   command=self._oeffne_stilregeln).pack(side=tk.LEFT, padx=2)
        ttk.Button(tb, text="⊞ Positionen zurücksetzen",
                   command=self._reset_positionen).pack(side=tk.LEFT, padx=2)
        ttk.Label(tb, text="Zoom:").pack(side=tk.LEFT, padx=(12, 2))
        self.v_zoom = tk.DoubleVar(value=1.0)
        ttk.Scale(tb, from_=0.3, to=2.5, variable=self.v_zoom,
                  orient=tk.HORIZONTAL, length=130,
                  command=lambda _: self.aktualisiere()).pack(side=tk.LEFT)
        ttk.Label(tb, text="  Klick auf Box → öffnet Editor"
                  ).pack(side=tk.LEFT, padx=8)

        # Hauptbereich: Canvas links, Seitenleiste rechts
        haupt = ttk.Frame(self)
        haupt.pack(fill=tk.BOTH, expand=True, padx=6, pady=4)

        cf = ttk.Frame(haupt)
        cf.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.cv = tk.Canvas(cf, background="white", cursor="crosshair")
        sy = ttk.Scrollbar(cf, orient=tk.VERTICAL,   command=self.cv.yview)
        sx = ttk.Scrollbar(cf, orient=tk.HORIZONTAL, command=self.cv.xview)
        self.cv.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        sy.pack(side=tk.RIGHT,  fill=tk.Y)
        sx.pack(side=tk.BOTTOM, fill=tk.X)
        self.cv.pack(fill=tk.BOTH, expand=True)
        # Linke Maustaste: Box verschieben
        self.cv.bind("<ButtonPress-1>",   self._on_box_press)
        self.cv.bind("<B1-Motion>",        self._on_box_drag)
        self.cv.bind("<ButtonRelease-1>",  self._on_box_release)
        # Mittlere Maustaste: Panning
        self.cv.bind("<ButtonPress-2>",   lambda e: self.cv.scan_mark(e.x, e.y))
        self.cv.bind("<B2-Motion>",        lambda e: self.cv.scan_dragto(e.x, e.y, gain=1))
        self.cv.bind("<MouseWheel>",
                     lambda e: self.cv.yview_scroll(int(-1*(e.delta/120)), "units"))
        self.cv.bind("<Button-4>", lambda e: self.cv.yview_scroll(-1, "units"))
        self.cv.bind("<Button-5>", lambda e: self.cv.yview_scroll(1,  "units"))
        self.cv.bind("<Escape>",   lambda e: self._zeichenmodus_beenden())

        # ── Seitenleiste (geteilt) ───────────────────────────
        sf = ttk.Frame(haupt)
        sf.pack(side=tk.RIGHT, fill=tk.Y, padx=(6, 0))

        # Ausgewählte Stelle
        selbf = ttk.LabelFrame(sf, text="Ausgewählte Stelle", width=215)
        selbf.pack(fill=tk.X, padx=0, pady=(0,4))
        selbf.pack_propagate(False)
        self._sel_nr = None
        self._sel_lbl = ttk.Label(selbf, text="(Klick auf Box zum Auswählen)",
                                   wraplength=200, foreground="#888888")
        self._sel_lbl.pack(anchor="w", padx=6, pady=(4,2))
        _fr = ttk.Frame(selbf); _fr.pack(fill=tk.X, padx=6, pady=2)
        ttk.Label(_fr, text="Farbe:").pack(side=tk.LEFT)
        self._sel_farbe_var = tk.StringVar(value="")
        self._sel_farbe_btn = tk.Button(
            _fr, width=4, relief="solid", borderwidth=1,
            background="#e8f4fd", cursor="hand2",
            command=self._sel_farbe_waehlen)
        self._sel_farbe_btn.pack(side=tk.LEFT, padx=4)
        ttk.Button(_fr, text="Auto", command=self._sel_farbe_auto
                   ).pack(side=tk.LEFT)
        _fr2 = ttk.Frame(selbf); _fr2.pack(fill=tk.X, padx=6, pady=2)
        ttk.Label(_fr2, text="Form:").pack(side=tk.LEFT)
        self._sel_form_var = tk.StringVar(value="rechteck")
        ttk.Combobox(_fr2, textvariable=self._sel_form_var,
                     values=["rechteck","dreieck","ellipse"],
                     state="readonly", width=10).pack(side=tk.LEFT, padx=4)
        self._sel_schutz_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(selbf, text="Schützen (vor Stilregeln)",
                        variable=self._sel_schutz_var).pack(anchor="w", padx=6)
        ttk.Button(selbf, text="✓ Übernehmen",
                   command=self._sel_uebernehmen).pack(pady=4)

        # ── Linienlegende (vertikal) ──────────────────────
        self._llf = ttk.LabelFrame(sf, text="Linienlegende", width=215)
        self._llf.pack(fill=tk.X, padx=0, pady=(0,4))
        self._leg_eintraege = [
            ("manual",  False, "Manuell"),
            ("profil",  False, "Profil/Katalog"),
            ("dat",     False, "Fundliste"),
            ("okuk",    True,  "OK/UK"),
            ("widersp", False, "Widerspruch"),
            ("gleich",  True,  "Gleich alt"),
        ]
        self._leg_canvases = []
        self._leg_labels   = []
        for _key, _dash, _lbl in self._leg_eintraege:
            _lr = ttk.Frame(self._llf)
            _lr.pack(fill=tk.X, padx=6, pady=2)
            _lc = tk.Canvas(_lr, width=36, height=16,
                            highlightthickness=0, bd=0)
            _lc.pack(side=tk.LEFT)
            _lc.create_line(3, 8, 33, 8, fill="#888888", width=3,
                            **( {"dash":(5,3)} if _dash else {} ))
            self._leg_canvases.append((_lc, _dash))
            _ll = tk.Label(_lr, text=_lbl, font=("", 9))
            _ll.pack(side=tk.LEFT, padx=4)
            self._leg_labels.append(_ll)
        self._aktualisiere_legende()

        # Nicht eingeordnet
        legf = ttk.LabelFrame(sf, text="Nicht eingeordnet", width=215)
        legf.pack(fill=tk.BOTH, expand=True, padx=0)
        legf.pack_propagate(False)
        self.leg_txt = scrolledtext.ScrolledText(
            legf, width=25, wrap=tk.WORD, font=("", 9),
            state="disabled", relief="flat", background="#fafafa")
        self.leg_txt.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)
        self.leg_txt.tag_configure("titel", font=("", 9, "bold"),
                                   foreground="#444444")

    # ── Zeichnen ──────────────────────────────────────────────
    def aktualisiere(self):
        self.cv.delete("all")
        self._leg_leeren()
        st = self.app.stellen
        if not st:
            self.cv.create_text(300, 150, text="Keine Stellen vorhanden",
                                font=("", 14), fill="gray")
            return

        z   = max(0.3, min(2.5, self.v_zoom.get()))
        bw  = int(self.BW * z); bh = int(self.BH * z)
        hg  = int(self.HG * z); vg = int(self.VG * z)
        fs  = max(7, int(8  * z))
        fss = max(6, int(7  * z))
        mg  = max(20, int(30 * z))

        ws = {w.stelle: min(w2.stufe for w2 in self.app.warnungen
                            if w2.stelle == w.stelle)
              for w in self.app.warnungen}

        def ist_neben(nr, s):
            if s.kein_befund or s.moderne_stoerung:
                return True
            if s.alle_juenger_als() or s.alle_aelter_als() or s.alle_gleich_alt_wie():
                return False  # Verbunden (auch über gleich) → immer in Hauptmatrix
            dat = s.effektiv_datierung_str()
            if dat == 'Datierung unbekannt':
                return True
            von = s.effektiv_von(); bis = s.effektiv_bis()
            if von is None or bis is None:
                return True  # Keine Jahreszahlen → Sidebar
            if (bis - von) > 1200:
                return True  # Zu breite Spanne → Sidebar (manuell nachbessern)
            return False


        haupt_st = {nr: s for nr, s in st.items() if not ist_neben(nr, s)}
        neben_st = {nr: s for nr, s in st.items() if ist_neben(nr, s)}

        # ── Levels berechnen ────────────────────────────────────
        levels = berechne_levels(haupt_st)

        # ── Outlier-Erkennung auf unkomprimierten Levels ────────
        # Stellen mit Nachbarn, die ≥8 Levels entfernt sind → Seitenleiste
        for _nr in list(haupt_st):
            _my = levels.get(_nr, 0)
            _nachb = (haupt_st[_nr].alle_juenger_als() +
                      haupt_st[_nr].alle_aelter_als())
            _lv_nachb = [levels[_n] for _n in _nachb if _n in levels]
            if _lv_nachb and min(abs(_my - _nl) for _nl in _lv_nachb) >= 8:
                neben_st[_nr] = haupt_st.pop(_nr)
                del levels[_nr]

        # ── Level-Lücken komprimieren (kein riesiger Leerraum) ──
        if levels:
            _alle = sorted(set(levels.values()))
            _map  = {old: new for new, old in enumerate(_alle)}
            levels = {nr: _map[lv] for nr, lv in levels.items()}

        self._fulle_legende(neben_st)
        if not haupt_st:
            self.cv.create_text(
                300, 150,
                text="Keine eingeordneten Befunde vorhanden.",
                font=("", 11), fill="#888888", justify="center")
            return

        max_lv = max(levels.values()) if levels else 0

        lg: Dict[int, List[str]] = {}
        for nr, lv in levels.items():
            lg.setdefault(lv, []).append(nr)
        for lv in lg:
            lg[lv].sort(key=_sort_key)

        # Barycenter-Heuristik
        for _ in range(3):
            for lv in sorted(lg):
                sc = {}
                for nr in lg[lv]:
                    s = haupt_st[nr]; nb = []
                    for xnr in s.alle_juenger_als():   # xnr älter → lv-1
                        if xnr in levels and levels[xnr] == lv - 1:
                            lst = lg.get(lv - 1, [])
                            nb.append(lst.index(xnr) if xnr in lst else 0)
                    for ynr in s.alle_aelter_als():    # ynr jünger → lv+1
                        if ynr in levels and levels[ynr] == lv + 1:
                            lst = lg.get(lv + 1, [])
                            nb.append(lst.index(ynr) if ynr in lst else 0)
                    sc[nr] = (sum(nb) / len(nb)) if nb else lg[lv].index(nr)
                lg[lv].sort(key=lambda n: sc.get(n, 0))

        max_w = max((len(nrs)*(bw+hg) - hg) for nrs in lg.values()) if lg else bw
        cx = mg + max_w // 2

        # Positionen: Level 0 (älteste) unten (großes y); Level max (jüngste) oben (kleines y)
        self._pos = {}
        for lv, nrs in lg.items():
            y = (max_lv - lv) * (bh + vg) + mg
            lw = len(nrs) * (bw + hg) - hg
            x0 = cx - lw // 2
            for i, nr in enumerate(nrs):
                self._pos[nr] = (x0 + i * (bw + hg), y)
        # Manuelle Positionen uebernehmen
        for nr in list(self._pos):
            s2 = haupt_st.get(nr)
            if s2 and s2.box_x is not None and s2.box_y is not None:
                self._pos[nr] = (s2.box_x, s2.box_y)
        # Warn-Set + Kontext fuer Drag/Linienfarben
        self._warn1 = {w.stelle for w in self.app.warnungen if w.stufe == 1}
        self._haupt_st = haupt_st

        # Verbindungslinien (zuerst, Boxen kommen drueber)
        self._line_colors = {}
        lw_normal = max(2, int(2.25 * z))  # 1.5× dicker
        for nr, s in haupt_st.items():
            if nr not in self._pos: continue
            x1, y1 = self._pos[nr]; cx1 = x1 + bw // 2
            for jnr in s.alle_juenger_als():
                if jnr not in self._pos: continue
                x2, y2 = self._pos[jnr]; cx2 = x2 + bw // 2
                fill, dash = self._linie_stil(nr, jnr, s)
                htag = f"lh_{nr}_{jnr}"
                kw = dict(fill=fill, width=lw_normal,
                          tags=(htag, "lhover", "connection"))
                if dash: kw["dash"] = (6, 4)
                self.cv.create_line(cx2, y2, cx1, y1+bh, **kw)
                self._line_colors[htag] = (fill, lw_normal, bool(dash))
            for gnr in s.alle_gleich_alt_wie():
                if gnr not in self._pos or gnr <= nr: continue
                x2, y2 = self._pos[gnr]
                _F = _LFARBEN["dark" if getattr(self.app,"_dark",True) else "light"]
                fill_g = _F["manual"][0] if gnr in s.gleich_alt_wie else _F["gleich"][0]
                htag = f"lg_{nr}_{gnr}"
                self.cv.create_line(x1 + bw, y1 + bh//2, x2, y2 + bh//2,
                                    fill=fill_g, dash=(4, 4),
                                    width=max(1, int(z)),
                                    tags=(htag, "lhover", "connection"))
                self._line_colors[htag] = (fill_g, max(1,int(z)), True)
        # Hover-Bindings für alle Linien
        self.cv.tag_bind("lhover", "<Enter>", self._on_line_enter)
        self.cv.tag_bind("lhover", "<Leave>", self._on_line_leave)

        # Boxen (Farbe + Form frei wählbar; Rand immer schwarz)
        for nr, (x, y) in self._pos.items():
            s = haupt_st[nr]; ms = ws.get(nr, 0)
            # Füllfarbe: Benutzer-Farbe hat Vorrang, dann Warn-Farbe
            if s.box_farbe:
                fill = s.box_farbe
            elif ms == 1:              fill = "#ffcccc"
            elif ms == 2:             fill = "#fff3cd"
            elif s.okuk_juenger_als:  fill = "#f0e8ff"
            else:                     fill = "#e8f4fd"
            out = "#000000"  # immer schwarzer Rand (4.1)
            lw  = max(1, int(2 * z))
            tag = f"box_{nr}"
            form = (s.box_form or 'rechteck').lower()
            if form == 'dreieck':
                pts = [x+bw//2, y, x, y+bh, x+bw, y+bh]
                self.cv.create_polygon(
                    pts, fill=fill, outline=out, width=lw, tags=(tag,"box"))
            elif form == 'ellipse':
                self.cv.create_oval(
                    x, y+bh//4, x+bw, y+3*bh//4,
                    fill=fill, outline=out, width=lw, tags=(tag,"box"))
            else:  # rechteck (Standard)
                self.cv.create_rectangle(
                    x, y, x+bw, y+bh,
                    fill=fill, outline=out, width=lw, tags=(tag,"box"))
            lbl = f"{nr}\n{s.befansprac[:16]}" if s.befansprac else nr
            self.cv.create_text(x + bw//2, y + bh//2,
                                text=lbl, font=("", fs, "bold"),
                                tags=(tag, "boxtext"), justify="center")
            sub = s.harris_info_kurztext()
            if sub:
                self.cv.create_text(x + bw//2, y + bh + 2,
                                    text=sub, font=("", fss),
                                    fill="#777777", tags=(tag,),
                                    anchor="n", width=bw + hg // 2)

        self.cv.configure(scrollregion=self.cv.bbox("all"))
        self.cv.xview_moveto(0)
        self.cv.yview_moveto(0)
        self.cv.tag_bind("box",     "<Button-1>",        self._on_click)
        self.cv.tag_bind("boxtext", "<Button-1>",        self._on_click)
        self.cv.tag_bind("box",     "<Double-Button-1>", self._on_dblclick)
        self.cv.tag_bind("boxtext", "<Double-Button-1>", self._on_dblclick)
        self.cv.tag_bind("box",     "<Button-3>",        self._ctx_box)
        self.cv.tag_bind("boxtext", "<Button-3>",        self._ctx_box)
        self.cv.tag_bind("lhover",  "<Button-3>",        self._ctx_linie)
        self.cv.bind("<Motion>",    self._on_motion)

    # ── Seitenleiste ──────────────────────────────────────────
    def _png_export(self):
        pfad = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG","*.png"), ("EPS/PostScript","*.eps"), ("Alle","*.*")])
        if not pfad: return
        bbox = self.cv.bbox("all")
        if not bbox:
            messagebox.showinfo("Leer", "Keine Inhalte zum Exportieren.")
            return
        x1,y1,x2,y2 = bbox
        marge = 20
        import tempfile, os
        eps_pfad = pfad if pfad.endswith(".eps") else pfad + ".eps.tmp"
        self.cv.postscript(
            file=eps_pfad, colormode="color",
            x=x1-marge, y=y1-marge,
            width=x2-x1+2*marge, height=y2-y1+2*marge)
        if pfad.endswith(".eps"):
            self.app.status(f"Exportiert (EPS): {pfad}")
            return
        try:
            from PIL import Image
            # with-Block: PIL hält die EPS-Datei sonst offen → unlink schlägt
            # unter Windows mit PermissionError fehl und der Nutzer bekommt
            # einen "Export-Fehler", obwohl die PNG längst geschrieben ist
            with Image.open(eps_pfad) as img:
                img.save(pfad)
            try: os.unlink(eps_pfad)
            except OSError: pass
            self.app.status(f"Exportiert (PNG): {pfad}")
        except ImportError:
            import shutil
            eps_final = pfad.replace(".png", ".eps")
            shutil.move(eps_pfad, eps_final)
            messagebox.showinfo("PIL fehlt",
                f"Pillow nicht installiert \u2192 als EPS gespeichert:\n{eps_final}\n\n"
                "Zum Aktivieren: pip install pillow")
        except Exception as e:
            try: os.unlink(eps_pfad)
            except: pass
            messagebox.showerror("Export-Fehler", str(e))

    def _oeffne_stilregeln(self):
        d = tk.Toplevel(self); d.title("Stilregeln – Harris-Matrix")
        d.geometry("520x420"); d.grab_set()
        ttk.Label(d, text="Neue Regel: Farbe + Form für alle Stellen einer Kategorie.",
                  wraplength=490).pack(padx=10, pady=6)
        nf = ttk.LabelFrame(d, text="Neue Regel"); nf.pack(fill=tk.X, padx=10)
        _r0 = ttk.Frame(nf); _r0.pack(fill=tk.X, padx=6, pady=4)
        ttk.Label(_r0, text="Befundgattung:").pack(side=tk.LEFT)
        v_bg = tk.StringVar(value="")
        _bg = AutocompleteAnsprache(_r0, textvariable=v_bg)
        _bg.pack(side=tk.LEFT, padx=4)
        ttk.Label(_r0, text="(leer = alle)").pack(side=tk.LEFT)
        _r1 = ttk.Frame(nf); _r1.pack(fill=tk.X, padx=6, pady=4)
        ttk.Label(_r1, text="Datierung:    ").pack(side=tk.LEFT)
        v_dat = tk.StringVar(value="")
        _dat = AutocompleteCombo(_r1, textvariable=v_dat)
        _dat.pack(side=tk.LEFT, padx=4)
        ttk.Label(_r1, text="(leer = alle)").pack(side=tk.LEFT)
        _r2 = ttk.Frame(nf); _r2.pack(fill=tk.X, padx=6, pady=4)
        ttk.Label(_r2, text="Farbe:").pack(side=tk.LEFT)
        v_fc = tk.StringVar(value="#e8f4fd")
        _fb = tk.Button(_r2, width=4, background="#e8f4fd",
                         relief="solid", borderwidth=1)
        _fb.pack(side=tk.LEFT, padx=4)
        def _pick_col():
            from tkinter import colorchooser
            r = colorchooser.askcolor(color=v_fc.get(), parent=d)
            if r and r[1]: v_fc.set(r[1]); _fb.config(background=r[1])
        _fb.config(command=_pick_col)
        ttk.Label(_r2, text="  Form:").pack(side=tk.LEFT)
        v_form = tk.StringVar(value="rechteck")
        ttk.Combobox(_r2, textvariable=v_form,
                     values=["rechteck","dreieck","ellipse"],
                     state="readonly", width=10).pack(side=tk.LEFT, padx=4)
        def _anwenden():
            regel = {"bg": v_bg.get().strip().lower(),
                     "dat": v_dat.get().strip().lower(),
                     "farbe": v_fc.get(), "form": v_form.get()}
            count = 0
            for nr, s in self.app.stellen.items():
                if s.box_geschuetzt: continue
                m_bg  = (not regel["bg"]  or
                         regel["bg"] in s.befansprac.lower())
                m_dat = (not regel["dat"] or
                         regel["dat"] in s.effektiv_datierung_str().lower())
                if m_bg and m_dat:
                    s.box_farbe = regel["farbe"]
                    s.box_form  = regel["form"]
                    count += 1
            self.aktualisiere()
            self.app.status(f"Stilregel auf {count} Stellen angewendet.")
        ttk.Button(nf, text="▶ Auf alle Passenden anwenden",
                   command=_anwenden).pack(pady=6)
        ttk.Separator(d).pack(fill=tk.X, padx=10, pady=4)
        ttk.Label(d, text="Hinweis: Individuell mit 'Schützen' versehene Stellen\n"
                  "werden von Regeln nicht überschrieben.", foreground="#777"
                  ).pack(padx=10)
        ttk.Button(d, text="Schließen", command=d.destroy).pack(pady=8)
        self.app.style_dialog(d)


    # ── Kontextmenü: Box ─────────────────────────────────────
    def _ctx_box(self, event):
        nr = None
        for tag in self.cv.gettags("current"):
            if tag.startswith("box_"):
                nr = tag[4:]; break
        if not nr or nr not in self.app.stellen: return
        m = tk.Menu(self, tearoff=0)
        m.add_command(
            label=f"Alle Linien von Stelle {nr} löschen",
            command=lambda: self._alle_linien_loeschen(nr))
        m.add_command(
            label="📋 Beziehungen anzeigen",
            command=lambda: self._beziehungen_popup(nr))
        m.add_separator()
        for lbl, mode in [("Jünger als …","juenger"),
                          ("Älter als …","aelter"),
                          ("Gleich alt wie …","gleich")]:
            m.add_command(
                label=f"Neue Linie: {lbl}",
                command=lambda m=mode, n=nr: self._zeichenmodus_starten(n, m))
        m.tk_popup(event.x_root, event.y_root)

    def _alle_linien_loeschen(self, nr):
        s = self.app.stellen.get(nr)
        if not s: return
        self.app._push_undo()
        _alle_rel = ("juenger_als","aelter_als","gleich_alt_wie","gestoert_von",
                     "profil_juenger_als","profil_aelter_als","profil_gleich",
                     "okuk_juenger_als","okuk_aelter_als",
                     "dat_juenger_als","dat_aelter_als")
        # 1. Eigene Listen leeren
        for attr in _alle_rel:
            setattr(s, attr, [])
        s.dat_suppressed = True
        s.box_x = None; s.box_y = None   # Position zurücksetzen → korrekte Neuberechnung
        # 2. Stelle auch aus ALLEN anderen Stellen entfernen —
        #    sonst erzeugt propagiere_beziehungen sofort _bidir_*-Einträge zurück
        for other in self.app.stellen.values():
            if other.stellennr == nr: continue
            for attr in _alle_rel:
                lst = getattr(other, attr)
                if nr in lst:
                    setattr(other, attr, [x for x in lst if x != nr])
        self.app._aktualisiere_alles()
        try: self.app.tab_stellen._load_editor()
        except: pass
        self.app.status(f"Stelle {nr}: alle Verbindungen gelöscht (Analyse aktualisiert).")

    def _beziehungen_popup(self, nr):
        s = self.app.stellen.get(nr)
        if not s: return
        d = tk.Toplevel(self)
        d.title(f"Beziehungen: Stelle {nr}")
        d.geometry("420x360")
        d.transient(self.winfo_toplevel())

        ttk.Label(d, text=f"Stelle {nr}  –  Beziehungsübersicht",
                  font=("", 10, "bold")).pack(pady=(10, 4), padx=10, anchor="w")

        def _make_section(parent, titel, liste, attr):
            if not liste: return
            lf = ttk.LabelFrame(parent, text=titel)
            lf.pack(fill=tk.X, padx=10, pady=3)
            for val in list(liste):
                fr = ttk.Frame(lf)
                fr.pack(fill=tk.X, padx=4, pady=1)
                ttk.Label(fr, text=val, width=14).pack(side=tk.LEFT)
                def _del(v=val, a=attr):
                    self.app._push_undo()
                    cur = list(getattr(s, a))
                    if v in cur:
                        cur.remove(v)
                        setattr(s, a, cur)
                    # Gegenstück beim Peer entfernen + unterdrücken — sonst
                    # belebt die Spiegelung/Neuberechnung den Eintrag wieder
                    _gegen = {"juenger_als": "aelter_als",
                              "aelter_als": "juenger_als",
                              "gleich_alt_wie": "gleich_alt_wie"}[a]
                    _richt = {"juenger_als": "juenger", "aelter_als": "aelter",
                              "gleich_alt_wie": "gleich"}[a]
                    peer = self.app.stellen.get(v)
                    if peer:
                        gl = getattr(peer, _gegen)
                        if nr in gl:
                            setattr(peer, _gegen, [x for x in gl if x != nr])
                    key = f"{_richt}:{v}"
                    if key not in s.unterdrueckte_bez:
                        s.unterdrueckte_bez.append(key)
                    self.app._aktualisiere_alles()
                    d.destroy()
                    self._beziehungen_popup(nr)
                ttk.Button(fr, text="✕ Löschen", width=10,
                           command=_del).pack(side=tk.RIGHT)

        sf = ttk.Frame(d)
        sf.pack(fill=tk.BOTH, expand=True, padx=4)
        j_all = list(s.juenger_als)
        a_all = list(s.aelter_als)
        g_all = list(s.gleich_alt_wie)
        _make_section(sf, "Jünger als (manuell)", j_all, "juenger_als")
        _make_section(sf, "Älter als (manuell)",  a_all, "aelter_als")
        _make_section(sf, "Gleich alt wie",        g_all, "gleich_alt_wie")
        if not j_all and not a_all and not g_all:
            ttk.Label(sf, text="(keine manuellen Beziehungen)",
                      foreground="#888888").pack(pady=10)
        ttk.Button(d, text="Schließen", command=d.destroy).pack(pady=8)
        self.app.style_dialog(d)

    # ── Kontextmenü: Linie ───────────────────────────────────
    def _ctx_linie(self, event):
        tags  = self.cv.gettags("current")
        # Daten-Tag finden: 'lh_<nr>_<jnr>' (jünger/älter) ODER 'lg_<nr>_<gnr>'
        # (gleich alt). WICHTIG: 'lhover' nicht verwechseln (beginnt auch mit lh)
        htag  = next((t for t in tags
                      if t.startswith("lh_") or t.startswith("lg_")), None)
        if not htag: return
        parts = htag.split("_", 2)
        if len(parts) < 3: return
        art, nr, jnr = parts[0], parts[1], parts[2]
        ist_gleich = (art == "lg")
        m = tk.Menu(self, tearoff=0)
        label = (f"Gleich-Verbindung {nr} = {jnr} löschen" if ist_gleich
                 else f"Linie {nr} – {jnr} löschen")
        m.add_command(label=label,
                      command=lambda: self._linie_loeschen(nr, jnr, ist_gleich))
        m.tk_popup(event.x_root, event.y_root)

    def _linie_loeschen(self, nr, jnr, ist_gleich=False):
        self.app._push_undo()
        def ex(lst, v): return [x for x in lst if x != v]
        def _unterdruecke(stelle, richtung, wert):
            # Gelöschte Linie auch für künftige Neuberechnungen sperren —
            # sonst erzeugt der Datierungsvergleich dieselbe Kante beim
            # nächsten F5 einfach wieder
            key = f"{richtung}:{wert}"
            if key not in stelle.unterdrueckte_bez:
                stelle.unterdrueckte_bez.append(key)
        if ist_gleich:
            # Nur die GLEICH-Verbindung lösen (beide Seiten), Über-/Unterlagerung
            # bleibt unberührt
            for a, b in ((nr, jnr), (jnr, nr)):
                if a in self.app.stellen:
                    sa = self.app.stellen[a]
                    sa.gleich_alt_wie = ex(sa.gleich_alt_wie, b)
                    sa.profil_gleich  = ex(sa.profil_gleich, b)
                    _unterdruecke(sa, "gleich", b)
                    sa.box_x = None; sa.box_y = None
        else:
            if nr in self.app.stellen:
                s = self.app.stellen[nr]
                for a in ("juenger_als","profil_juenger_als",
                          "okuk_juenger_als","dat_juenger_als"):
                    setattr(s, a, ex(getattr(s, a), jnr))
                _unterdruecke(s, "juenger", jnr)
                s.box_x = None; s.box_y = None  # Position zurücksetzen
            if jnr in self.app.stellen:
                t = self.app.stellen[jnr]
                for a in ("aelter_als","profil_aelter_als",
                          "okuk_aelter_als","dat_aelter_als"):
                    setattr(t, a, ex(getattr(t, a), nr))
                _unterdruecke(t, "aelter", nr)
                t.box_x = None; t.box_y = None  # Position zurücksetzen
        self.app._aktualisiere_alles()
        try: self.app.tab_stellen._load_editor()
        except: pass
        _txt = "Gleich-Verbindung" if ist_gleich else "Linie"
        self.app.status(f"{_txt} {nr} – {jnr} gelöscht (Analyse aktualisiert).")

    # ── Neue Linie zeichnen ──────────────────────────────────
    def _zeichenmodus_starten(self, nr, mode):
        self._draw_mode = mode
        self._draw_src  = nr
        self.cv.config(cursor="crosshair")
        lbl = {"juenger":"Jünger als","aelter":"Älter als",
               "gleich":"Gleich alt wie"}[mode]
        self.app.status(
            f"Stelle {nr} {lbl} …  Zielbox anklicken  [Esc = Abbrechen]")

    def _zeichenmodus_beenden(self):
        self._draw_mode = None
        self._draw_src  = None
        self.cv.delete("draw_preview")
        self.cv.config(cursor="")
        self.app.status("Zeichenmodus beendet.")

    def _zeichne_linie_abschliessen(self, ziel_nr):
        src, mode = self._draw_src, self._draw_mode
        self._zeichenmodus_beenden()
        if not src or src == ziel_nr: return
        s = self.app.stellen.get(src)
        if not s: return
        self.app._push_undo()
        zt = self.app.stellen.get(ziel_nr)

        # ── Manuelle Linie hat Vorrang: direkte Gegenbeziehung entfernen ──
        # Verhindert Zyklen in berechne_levels (würde ohne Ende iterieren).
        # Beispiel: "104 jünger als 109" → entferne automatisch "109 jünger als 104"
        # aus ALLEN Quellen (profil_*, okuk_*, dat_*, manuell).
        def _entferne(stelle, attr, wert):
            if stelle is None: return
            lst = getattr(stelle, attr, [])
            if wert in lst:
                setattr(stelle, attr, [x for x in lst if x != wert])

        def _unterdrueckung_aufheben(stelle, richtung, wert):
            # Neu gezeichnete Linie schlägt frühere Löschung dieser Beziehung
            if stelle is None: return
            key = f"{richtung}:{wert}"
            if key in stelle.unterdrueckte_bez:
                stelle.unterdrueckte_bez.remove(key)

        if mode == "juenger":
            # src ist jünger als ziel → Gegenbeziehung: ziel jünger als src
            for a in ("juenger_als","profil_juenger_als","okuk_juenger_als","dat_juenger_als"):
                _entferne(zt, a, src)          # ziel.juenger_als enthielt src → weg
            for a in ("aelter_als","profil_aelter_als","okuk_aelter_als","dat_aelter_als"):
                _entferne(s, a, ziel_nr)       # src.aelter_als enthielt ziel → weg
            # Auch eine bestehende GLEICH-Beziehung widerspricht der neuen
            # Ordnung → beidseitig entfernen (sonst Level-Aufschaukeln)
            for a in ("gleich_alt_wie","profil_gleich"):
                _entferne(s,  a, ziel_nr)
                _entferne(zt, a, src)
            if ziel_nr not in s.juenger_als:
                s.juenger_als.append(ziel_nr)
            _unterdrueckung_aufheben(s,  "juenger", ziel_nr)
            _unterdrueckung_aufheben(zt, "aelter",  src)

        elif mode == "aelter":
            # src ist älter als ziel → Gegenbeziehung: ziel älter als src
            for a in ("aelter_als","profil_aelter_als","okuk_aelter_als","dat_aelter_als"):
                _entferne(zt, a, src)
            for a in ("juenger_als","profil_juenger_als","okuk_juenger_als","dat_juenger_als"):
                _entferne(s, a, ziel_nr)
            for a in ("gleich_alt_wie","profil_gleich"):
                _entferne(s,  a, ziel_nr)
                _entferne(zt, a, src)
            if ziel_nr not in s.aelter_als:
                s.aelter_als.append(ziel_nr)
            _unterdrueckung_aufheben(s,  "aelter",  ziel_nr)
            _unterdrueckung_aufheben(zt, "juenger", src)

        elif mode == "gleich":
            # Gleich alt → keine Gegenbeziehung nötig, aber jünger/älter entfernen
            for a in ("juenger_als","aelter_als","profil_juenger_als","profil_aelter_als",
                      "okuk_juenger_als","okuk_aelter_als","dat_juenger_als","dat_aelter_als"):
                _entferne(s,  a, ziel_nr)
                _entferne(zt, a, src)
            # Beidseitig manuell eintragen (gezeichnet = manuelle Entscheidung;
            # Löschen über Editor/Menü räumt ebenfalls beide Seiten auf)
            if ziel_nr not in s.gleich_alt_wie:
                s.gleich_alt_wie.append(ziel_nr)
            if zt and src not in zt.gleich_alt_wie:
                zt.gleich_alt_wie.append(src)
            _unterdrueckung_aufheben(s,  "gleich", ziel_nr)
            _unterdrueckung_aufheben(zt, "gleich", src)
            # Gleich gewinnt: jede jünger/älter-Beziehung INNERHALB der jetzt
            # verschmolzenen Gleich-Gruppe entfernen (auch transitive, z.B.
            # 9 jünger 6, während 8 gleich 6 und 8 gleich 9 gezogen wird) →
            # Stellen & Harris bleibt konsistent, keine Reihen-Kollaps.
            _bereinige_gleich_ordnung(self.app.stellen)

        s.dat_suppressed = False
        if zt:
            zt.dat_suppressed = False
        # Gesamten verbundenen Teilgraph traversieren (BFS).
        # OKUK-Gegenstücke (z.B. 106.okuk_aelter_als=["109"]) erzeugen nach
        # propagiere_beziehungen bidir-Kanten die den manuellen Umkehr-Zyklus
        # wieder aufbauen, auch wenn nur src+ziel gelöscht wurden.
        # Lösung: OKUK für ALLE Stellen im Teilgraph löschen + Positionen reset.
        _besucht: set = set()
        _queue: list = [src, ziel_nr]
        while _queue:
            _nr = _queue.pop()
            if _nr in _besucht or _nr not in self.app.stellen:
                continue
            _besucht.add(_nr)
            _s2 = self.app.stellen[_nr]
            for _xnr in (_s2.alle_juenger_als() + _s2.alle_aelter_als()
                         + _s2.alle_gleich_alt_wie()):
                if _xnr not in _besucht:
                    _queue.append(_xnr)
        for _nr in _besucht:
            _s2 = self.app.stellen.get(_nr)
            if _s2:
                _s2.box_x = None; _s2.box_y = None
                _s2.okuk_juenger_als = []   # OKUK im gesamten Teilgraph löschen —
                _s2.okuk_aelter_als  = []   # Gegenstücke sonst → Zyklus via bidir
                _s2.dat_juenger_als  = []   # Datierungs-Beziehungen auch löschen;
                _s2.dat_aelter_als   = []   # manuelle Linie hat Vorrang (F5 = neu)
        self.app._aktualisiere_alles()
        try: self.app.tab_stellen._load_editor()
        except: pass
        self.app.status(f"Linie hinzugefügt: {src} – {ziel_nr}  "
                        f"(Analyse automatisch aktualisiert)")

    def _on_motion(self, event):
        if not self._draw_mode:
            return
        self.cv.delete("draw_preview")
        if self._draw_src not in self._pos:
            # Sidebar-Stelle: kein Ursprungs-Anker, aber Zeichenmodus aktiv.
            # Kleines Fadenkreuz am Cursor zeigt dem Nutzer: Zielbox anklicken.
            cx = self.cv.canvasx(event.x)
            cy = self.cv.canvasy(event.y)
            r = 10
            self.cv.create_line(cx-r, cy, cx+r, cy,
                                fill="#888888", width=1, dash=(3,3),
                                tags="draw_preview")
            self.cv.create_line(cx, cy-r, cx, cy+r,
                                fill="#888888", width=1, dash=(3,3),
                                tags="draw_preview")
            self.cv.tag_raise("draw_preview")
            return
        z = self.v_zoom.get()
        bw = int(self.BW*z); bh = int(self.BH*z)
        x0, y0 = self._pos[self._draw_src]
        cx = self.cv.canvasx(event.x)
        cy = self.cv.canvasy(event.y)
        self.cv.create_line(x0+bw//2, y0+bh//2, cx, cy,
                            fill="#888888", width=2, dash=(6,4),
                            tags="draw_preview")
        self.cv.tag_raise("draw_preview")

    def _aktualisiere_legende(self):
        """Legende mit aktuellen Theme-Farben neu zeichnen."""
        if not hasattr(self, "_leg_canvases"): return
        dark = getattr(self.app, "_dark", True)
        F = _LFARBEN["dark" if dark else "light"]
        bg = "#1e1e2e" if dark else "#f0f0f0"
        fg = "#cdd6f4" if dark else "#111111"
        for (_lc, _dash), (_key, _, _) in zip(
                self._leg_canvases, self._leg_eintraege):
            col = F[_key][0]
            _lc.delete("all")
            _lc.configure(background=bg)
            kw = {"dash":(5,3)} if _dash else {}
            _lc.create_line(3, 8, 33, 8, fill=col, width=3, **kw)
        for _ll in self._leg_labels:
            _ll.configure(background=bg, foreground=fg)

    def _leg_leeren(self):
        self.leg_txt.config(state="normal")
        self.leg_txt.delete("1.0", tk.END)
        self.leg_txt.config(state="disabled")

    def _sidebar_ctx(self, event, nr):
        """Kontextmenü für Stelle in der Seitenleiste (Rechtsklick)."""
        m = tk.Menu(self, tearoff=0)
        for lbl, mode in [("Jünger als … (Linie zeichnen)", "juenger"),
                           ("Älter als … (Linie zeichnen)",  "aelter"),
                           ("Gleich alt wie … (Linie zeichnen)", "gleich")]:
            m.add_command(label=f"Stelle {nr}: {lbl}",
                          command=lambda mo=mode, n=nr:
                              self._zeichenmodus_starten(n, mo))
        m.add_separator()
        m.add_command(label=f"Stelle {nr} in Stellen-Tab öffnen",
                      command=lambda n=nr: self._sidebar_stellen_tab(n))
        m.tk_popup(event.x_root, event.y_root)

    def _sidebar_stellen_tab(self, nr):
        try:
            self.app.notebook.select(self.app.tab_stellen)
            if self.app.tab_stellen.tree.exists(nr):
                self.app.tab_stellen.tree.selection_set(nr)
                self.app.tab_stellen.tree.see(nr)
                self.app.tab_stellen._on_sel()
        except Exception:
            pass

    def _fulle_legende(self, neben: Dict[str, Stelle]):
        dark    = getattr(self.app, "_dark", True)
        link_fg = "#89b4fa" if dark else "#1565c0"

        arbeit = sorted([nr for nr, s in neben.items() if s.kein_befund],
                        key=_sort_key)
        stoer  = sorted([nr for nr, s in neben.items()
                         if s.moderne_stoerung and not s.kein_befund],
                        key=_sort_key)
        biotb  = sorted([nr for nr in stoer
                         if any(k in neben[nr].befansprac.lower()
                                for k in ("biot","baumwurf"))],
                        key=_sort_key)
        stoer2 = [nr for nr in stoer if nr not in biotb]
        unbek  = sorted([nr for nr, s in neben.items()
                         if not s.kein_befund and not s.moderne_stoerung],
                        key=_sort_key)

        self.leg_txt.config(state="normal")
        self.leg_txt.delete("1.0", tk.END)
        leer = True
        for titel, nrs in [("Arbeitsbereiche",          arbeit),
                            ("Stoerungen",               stoer2),
                            ("Bioturbationen/Baumwurf",  biotb),
                            ("Undatiert / ohne Bezug",   unbek)]:
            if not nrs: continue
            self.leg_txt.insert(tk.END, f"{titel}:\n", "titel")
            for i, nr in enumerate(nrs):
                tag = f"sid_{id(neben)}_{nr}"   # eindeutiger Tag
                self.leg_txt.insert(tk.END, nr, tag)
                self.leg_txt.tag_configure(tag, foreground=link_fg,
                                           underline=True)
                self.leg_txt.tag_bind(
                    tag, "<Button-3>",
                    lambda e, n=nr: self._sidebar_ctx(e, n))
                self.leg_txt.tag_bind(
                    tag, "<Button-1>",
                    lambda e, n=nr: self._sidebar_stellen_tab(n))
                if i < len(nrs) - 1:
                    self.leg_txt.insert(tk.END, ", ")
            self.leg_txt.insert(tk.END, "\n\n")
            leer = False
        if leer:
            self.leg_txt.insert(tk.END, "(keine)")
        self.leg_txt.config(state="disabled")


    def _linie_stil(self, nr, jnr, s):
        """Linienfarbe theme-abhängig aus Quellfeldern bestimmen."""
        F = _LFARBEN["dark" if getattr(self.app,"_dark",True) else "light"]
        t = (self._haupt_st or {}).get(jnr) or self.app.stellen.get(jnr)
        if nr in self._warn1 and jnr in self._warn1:
            return F["widersp"]
        if jnr in s.juenger_als or (t and nr in t.aelter_als):
            return F["manual"]
        if jnr in s.profil_juenger_als or (t and nr in t.profil_aelter_als):
            return F["profil"]
        if jnr in s.dat_juenger_als or (t and nr in t.dat_aelter_als):
            return F["dat"]
        if jnr in s.okuk_juenger_als or (t and nr in t.okuk_aelter_als):
            return F["okuk"]
        return F["manual"]

    def _on_box_press(self, event):
        # Zeichenmodus-Abschluss wurde bereits von uns gesetzt: Drag-Start verhindern
        if getattr(self, "_draw_just_done", False):
            self._draw_just_done = False
            return
        if self._draw_mode:
            # draw_preview liegt via tag_raise immer oben → tag_bind("box") feuert nie.
            # Daher gesamte Draw-Logik hier im canvas-level Binding erledigen.
            cx = self.cv.canvasx(event.x)
            cy = self.cv.canvasy(event.y)
            items = self.cv.find_overlapping(cx - 3, cy - 3, cx + 3, cy + 3)
            for item in items:
                for tag in self.cv.gettags(item):
                    if tag.startswith("box_"):
                        self._draw_just_done = True
                        self._zeichne_linie_abschliessen(tag[4:])
                        return
            # Keine Box getroffen → Draw-Modus abbrechen
            self._zeichenmodus_beenden()
            return
        self._drag_nr = None
        self._drag_start_pos = None
        for tag in self.cv.gettags("current"):
            if tag.startswith("box_"):
                nr = tag[4:]
                if nr in self._pos:
                    self._drag_nr = nr
                    self._drag_start_pos = self._pos[nr]
                    cx = self.cv.canvasx(event.x)
                    cy = self.cv.canvasy(event.y)
                    ox, oy = self._pos[nr]
                    self._drag_offset = (cx - ox, cy - oy)
                return

    def _on_box_drag(self, event):
        if self._drag_nr is None: return
        cx = self.cv.canvasx(event.x)
        cy = self.cv.canvasy(event.y)
        nx = int(cx - self._drag_offset[0])
        ny = int(cy - self._drag_offset[1])
        old_x, old_y = self._pos[self._drag_nr]
        self.cv.move(f"box_{self._drag_nr}", nx-old_x, ny-old_y)
        self._pos[self._drag_nr] = (nx, ny)
        if self._drag_timer:
            try: self.after_cancel(self._drag_timer)
            except: pass
        self._drag_timer = self.after(20, self._lines_update)

    def _lines_update(self):
        self._drag_timer = None
        self.cv.delete("connection")
        if not hasattr(self, '_haupt_st'): return
        z = self.v_zoom.get()
        bw = int(self.BW*z); bh = int(self.BH*z)
        lw = max(2, int(2.25*z))
        self._line_colors = {}
        for nr, s in self._haupt_st.items():
            if nr not in self._pos: continue
            x1, y1 = self._pos[nr]; cx1 = x1+bw//2
            for jnr in s.alle_juenger_als():
                if jnr not in self._pos: continue
                x2, y2 = self._pos[jnr]; cx2 = x2+bw//2
                fill, dash = self._linie_stil(nr, jnr, s)
                htag = f"lh_{nr}_{jnr}"
                kw = dict(fill=fill, width=lw, tags=(htag,"lhover","connection"))
                if dash: kw["dash"] = (6,4)
                self.cv.create_line(cx2, y2, cx1, y1+bh, **kw)
                self._line_colors[htag] = (fill, lw, bool(dash))
            for gnr in s.alle_gleich_alt_wie():
                if gnr not in self._pos or gnr <= nr: continue
                x2, y2 = self._pos[gnr]
                fg = "#222222" if gnr in s.gleich_alt_wie else "#888888"
                htag = f"lg_{nr}_{gnr}"
                self.cv.create_line(x1+bw, y1+bh//2, x2, y2+bh//2,
                    fill=fg, dash=(4,4), width=max(2,int(1.5*z)),
                    tags=(htag,"lhover","connection"))
                self._line_colors[htag] = (fg, max(1,int(z)), True)
        self.cv.tag_raise("box")
        self.cv.tag_bind("lhover","<Enter>",self._on_line_enter)
        self.cv.tag_bind("lhover","<Leave>",self._on_line_leave)

    def _on_box_release(self, event):
        if self._drag_nr is None: return
        nr = self._drag_nr
        s = self.app.stellen.get(nr)
        new_pos = self._pos.get(nr)
        # Undo nur wenn Box wirklich verschoben wurde (kein Push bei bloßem Klick)
        if s and new_pos is not None and new_pos != self._drag_start_pos:
            self.app._push_undo()   # speichert alten box_x/box_y vor dem Update
        if s and new_pos is not None:
            s.box_x, s.box_y = new_pos
        self._drag_nr = None
        if self._drag_timer:
            try: self.after_cancel(self._drag_timer)
            except: pass
            self._drag_timer = None
        self.aktualisiere()

    def _aktualisiere_neu_sortiert(self):
        """Matrix aktualisieren + alle manuellen Positionen zurücksetzen."""
        for s in self.app.stellen.values():
            s.box_x = None; s.box_y = None
        self.aktualisiere()
        self.app.status("Harris-Matrix neu sortiert (alle Positionen frisch berechnet).")

    def _reset_positionen(self):
        for s in self.app.stellen.values():
            s.box_x = None; s.box_y = None
        self.aktualisiere()
        self.app.status("Box-Positionen auf berechnetes Layout zurückgesetzt.")

    def _on_line_enter(self, event):
        for tag in self.cv.gettags("current"):
            if tag in self._line_colors:
                _, w, _ = self._line_colors[tag]
                self.cv.itemconfig("current", fill="#4fc3f7", width=max(3,w+2))
                return

    def _on_line_leave(self, event):
        for tag in self.cv.gettags("current"):
            if tag in self._line_colors:
                fill, w, _ = self._line_colors[tag]
                self.cv.itemconfig("current", fill=fill, width=w)
                return

    def _update_sel_panel(self, nr):
        self._sel_nr = nr
        if nr not in self.app.stellen: return
        s = self.app.stellen[nr]
        lbl = f"Stelle {nr}"
        if s.befansprac: lbl += f"\n{s.befansprac[:28]}"
        self._sel_lbl.config(text=lbl)
        farbe = s.box_farbe or "#e8f4fd"
        self._sel_farbe_var.set(s.box_farbe or "")
        self._sel_farbe_btn.config(background=farbe)
        self._sel_form_var.set(s.box_form or "rechteck")
        self._sel_schutz_var.set(s.box_geschuetzt)

    def _sel_farbe_waehlen(self):
        from tkinter import colorchooser
        aktuell = self._sel_farbe_var.get() or "#e8f4fd"
        r = colorchooser.askcolor(color=aktuell, parent=self,
                                   title="Farbe für Harris-Box")
        if r and r[1]:
            self._sel_farbe_var.set(r[1])
            self._sel_farbe_btn.config(background=r[1])

    def _sel_farbe_auto(self):
        self._sel_farbe_var.set("")
        self._sel_farbe_btn.config(background="#e8f4fd")

    def _sel_uebernehmen(self):
        nr = self._sel_nr
        if not nr or nr not in self.app.stellen: return
        s = self.app.stellen[nr]
        s.box_farbe      = self._sel_farbe_var.get()
        s.box_form       = self._sel_form_var.get()
        s.box_geschuetzt = self._sel_schutz_var.get()
        self.aktualisiere()
        self.app.status(f"Stil für Stelle {nr} übernommen.")

    def _on_click(self, event):
        """Klick auf Box: Panel aktualisieren (Draw-Modus wird in _on_box_press behandelt)."""
        if self._draw_mode:
            return  # Wurde bereits in _on_box_press abgeschlossen
        for tag in self.cv.gettags("current"):
            if tag.startswith("box_"):
                nr = tag[4:]
                self._update_sel_panel(nr)
                if (self._popup_win is not None
                        and self._popup_win.winfo_exists()):
                    self._popup_aktualisieren(nr)
                return

    def _on_dblclick(self, event):
        """Doppelklick: Stil-Popup öffnen oder fokussieren."""
        for tag in self.cv.gettags("current"):
            if tag.startswith("box_"):
                nr = tag[4:]
                self._update_sel_panel(nr)
                if (self._popup_win is not None
                        and self._popup_win.winfo_exists()):
                    self._popup_aktualisieren(nr)
                    self._popup_win.lift()
                    self._popup_win.focus_force()
                else:
                    self._stil_popup(nr,
                                     event.x_root, event.y_root)
                return

    def _popup_aktualisieren(self, nr):
        """Aktualisiert das offene Popup auf eine andere Stelle."""
        if (self._popup_win is None
                or not self._popup_win.winfo_exists()):
            return
        # Popup-Titel + Inhalt neu befüllen via destroy+reopen
        rx = self._popup_win.winfo_x()
        ry = self._popup_win.winfo_y()
        self._popup_win.destroy()
        self._popup_win = None
        self._stil_popup(nr, rx, ry)

    def _stil_popup(self, nr, rx, ry):
        s = self.app.stellen.get(nr)
        if not s: return
        d = tk.Toplevel(self)
        self._popup_win = d
        d.title(f"Box-Stil: Stelle {nr}")
        d.geometry(f"250x210+{rx+8}+{ry+8}")
        d.resizable(False, False)
        d.transient(self.winfo_toplevel())
        d.protocol("WM_DELETE_WINDOW",
                   lambda: (setattr(self, "_popup_win", None), d.destroy()))
        lbl = f"Stelle {nr}"
        if s.befansprac: lbl += f"  ·  {s.befansprac[:28]}"
        ttk.Label(d, text=lbl, font=("",9,"bold")).pack(pady=(8,4), padx=10)
        _r0 = ttk.Frame(d); _r0.pack(fill=tk.X, padx=10, pady=3)
        ttk.Label(_r0, text="Farbe:", width=7).pack(side=tk.LEFT)
        v_fc = tk.StringVar(value=s.box_farbe or "")
        _fb  = tk.Button(_r0, width=4, relief="solid", borderwidth=1,
                         background=s.box_farbe or "#e8f4fd")
        def _pick(btn=_fb, var=v_fc):
            from tkinter import colorchooser
            r = colorchooser.askcolor(color=var.get() or "#e8f4fd", parent=d)
            if r and r[1]: var.set(r[1]); btn.config(background=r[1])
        _fb.config(command=_pick)
        _fb.pack(side=tk.LEFT, padx=4)
        ttk.Button(_r0, text="Auto",
                   command=lambda: (v_fc.set(""), _fb.config(background="#e8f4fd"))
                   ).pack(side=tk.LEFT)
        _r1 = ttk.Frame(d); _r1.pack(fill=tk.X, padx=10, pady=3)
        ttk.Label(_r1, text="Form:", width=7).pack(side=tk.LEFT)
        v_fm = tk.StringVar(value=s.box_form or "rechteck")
        ttk.Combobox(_r1, textvariable=v_fm,
                     values=["rechteck","dreieck","ellipse"],
                     state="readonly", width=12).pack(side=tk.LEFT, padx=4)
        v_sc = tk.BooleanVar(value=s.box_geschuetzt)
        ttk.Checkbutton(d, text="Schützen (vor Stilregeln)",
                        variable=v_sc).pack(anchor="w", padx=10, pady=2)
        def _ok():
            s.box_farbe = v_fc.get()
            s.box_form  = v_fm.get()
            s.box_geschuetzt = v_sc.get()
            self._popup_win = None
            self.aktualisiere()
            self.app.status(f"Stil für Stelle {nr} gespeichert.")
            d.destroy()
        bf = ttk.Frame(d); bf.pack(pady=8)
        ttk.Button(bf, text="✓ Übernehmen", command=_ok).pack(side=tk.LEFT, padx=4)
        ttk.Button(bf, text="Abbrechen",
                   command=lambda: (setattr(self,"_popup_win",None),
                                   d.destroy())).pack(side=tk.LEFT, padx=4)
        d.bind("<Return>", lambda _: _ok())
        d.bind("<Escape>", lambda _: d.destroy())
        self.app.style_dialog(d)

# ═══════════════════════════════════════════════════════════════
# TAB: LOG
# ═══════════════════════════════════════════════════════════════

class LogTab(ttk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self._build()

    def _build(self):
        tb = ttk.Frame(self)
        tb.pack(fill=tk.X, padx=6, pady=4)
        ttk.Button(tb, text="Log-Datei laden…", command=self._laden).pack(side=tk.LEFT, padx=2)
        ttk.Button(tb, text="Log speichern…",   command=self._speichern).pack(side=tk.LEFT, padx=2)
        ttk.Button(tb, text="Leeren",            command=self._leeren).pack(side=tk.LEFT, padx=2)
        ttk.Label(tb, text=f"Auto-gespeichert: {LOG_PATH}",
                  foreground="#888888").pack(side=tk.LEFT, padx=10)
        self.txt = scrolledtext.ScrolledText(
            self, font=("Courier", 9), state="disabled",
            wrap=tk.WORD, bg="#1e1e1e", fg="#d4d4d4")
        self.txt.pack(fill=tk.BOTH, expand=True, padx=6, pady=4)
        for lv, fg in [("CRITICAL","#ff4444"), ("ERROR","#ff6666"),
                        ("WARNING","#ffaa00"), ("INFO","#d4d4d4"),
                        ("DEBUG","#777777")]:
            self.txt.tag_config(lv, foreground=fg)

    def append_log(self, msg, level="INFO"):
        try:
            self.txt.config(state="normal")
            self.txt.insert(tk.END, msg + "\n", level)
            self.txt.see(tk.END)
            self.txt.config(state="disabled")
        except: pass

    def _laden(self):
        p = filedialog.askopenfilename(
            filetypes=[("Log","*.log *.txt"), ("Alle","*.*")])
        if not p: return
        try:
            c = Path(p).read_text(encoding="utf-8", errors="replace")
            self.txt.config(state="normal")
            self.txt.delete("1.0", tk.END)
            self.txt.insert(tk.END, f"=== {p} ===\n\n")
            self.txt.insert(tk.END, c)
            self.txt.config(state="disabled")
        except Exception as e:
            messagebox.showerror("Fehler", str(e))

    def _speichern(self):
        p = filedialog.asksaveasfilename(
            defaultextension=".log",
            filetypes=[("Log","*.log"), ("Text","*.txt")])
        if p:
            try: Path(p).write_text(self.txt.get("1.0", tk.END), encoding="utf-8")
            except Exception as e: messagebox.showerror("Fehler", str(e))

    def _leeren(self):
        self.txt.config(state="normal")
        self.txt.delete("1.0", tk.END)
        self.txt.config(state="disabled")

# ═══════════════════════════════════════════════════════════════
# HAUPTFENSTER
# ═══════════════════════════════════════════════════════════════

class App(_APP_BASE):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_TITEL}  v{VERSION}")
        self.geometry("1380x870")
        self.minsize(980, 660)

        self.stellen:  Dict[str, Stelle] = {}
        self.funde:    List[Fund]         = []
        self.warnungen: List[Warnung]     = []
        self.undo_stack = UndoStack()
        self.fundliste_pfad:      Optional[Path] = None
        self.stellenkatalog_pfad: Optional[Path] = None
        self.stil_regeln: list = []

        self._gui_log = GUILogHandler()
        logging.getLogger().addHandler(self._gui_log)

        self._build_menu()
        self._build_tabs()
        self._build_status()
        self._gui_log.callback = self.tab_log.append_log

        log.info(f"{APP_TITEL} gestartet  v{VERSION}  |  "
                 f"DnD={'ja' if DND_OK else 'nein'}")
        self._dark = True
        self._apply_theme(dark=True)
        self.status(f"Bereit.  WNK-Datierungen eingebettet: "
                    f"{len(WNK_DATIERUNGEN_EINGEBETTET)} Begriffe.")

    def _build_menu(self):
        mb = tk.Menu(self)

        mf = tk.Menu(mb, tearoff=0)
        mf.add_command(label="Projekt speichern…",
                       command=self.cmd_speichern,           accelerator="Ctrl+S")
        mf.add_command(label="Projekt laden…",
                       command=self.cmd_laden,               accelerator="Ctrl+O")
        mf.add_separator()
        mf.add_command(label="Beenden", command=self.quit)
        mb.add_cascade(label="Datei", menu=mf)

        # ── Bearbeiten ──
        self._mbearbeiten = tk.Menu(mb, tearoff=0)
        self._mbearbeiten.add_command(
            label="Rückgängig", command=self.cmd_undo,
            accelerator="Ctrl+Z", state="disabled")
        self._mbearbeiten.add_separator()
        self._mbearbeiten.add_command(label="Analyse ausführen",
                                       command=self.cmd_analyse, accelerator="F5")
        self._mbearbeiten.add_command(label="Auswertung exportieren…",
                                       command=self.cmd_exportiere)
        mb.add_cascade(label="Bearbeiten", menu=self._mbearbeiten)

        mu = tk.Menu(mb, tearoff=0)
        mu.add_command(label="Hilfe / README…",
                       command=lambda: popup(self, "Hilfe", README_TEXT))
        mu.add_separator()
        mu.add_command(label="Spenden",
                       command=lambda: webbrowser.open(PAYPAL_URL))
        mu.add_separator()
        mu.add_command(label="Lizenz…",
                       command=lambda: popup(self, "Lizenz", LIZENZ_TEXT))
        mu.add_separator()
        mb.add_cascade(label="Über", menu=mu)

        self.config(menu=mb)
        self.bind("<Control-f>", lambda _: self.cmd_lade_fundliste())
        self.bind("<Control-k>", lambda _: self.cmd_lade_stellenkatalog())
        self.bind("<Control-s>", lambda _: self.cmd_speichern())
        self.bind("<Control-o>", lambda _: self.cmd_laden())
        self.bind("<Control-z>", self.cmd_undo)
        self.bind("<F5>",        lambda _: self.cmd_analyse())

    def _build_tabs(self):
        # Datei-Panel oben
        self._datei_panel = ttk.LabelFrame(self, text="Geladene Dateien")
        self._datei_panel.pack(fill=tk.X, padx=6, pady=(4,0))
        self._fl_var  = tk.StringVar(value="–")
        self._sk_var  = tk.StringVar(value="–")
        hint8 = ("", 8)

        # Zeile 0: Stellenkatalog (oben)
        r0 = ttk.Frame(self._datei_panel)
        r0.pack(fill=tk.X, padx=4, pady=(3,1))
        ttk.Label(r0, text="Stellenkatalog:", width=15, anchor='w').pack(side=tk.LEFT)
        ttk.Button(r0, text="📂 Öffnen",
                   command=self.cmd_lade_stellenkatalog).pack(side=tk.LEFT, padx=2)
        ttk.Button(r0, text="✕ Löschen",
                   command=self.cmd_entferne_stellenkatalog).pack(side=tk.LEFT, padx=2)
        ttk.Label(r0, textvariable=self._sk_var,
                  foreground='#4a9eff').pack(side=tk.LEFT, padx=6)
        ttk.Label(r0, text="(Löschen entfernt Profiltext-Beziehungen, OK/UK und Katalog-Datierungen)",
                  foreground="#888888", font=hint8).pack(side=tk.LEFT)

        # Zeile 1: Fundliste (darunter)
        r1 = ttk.Frame(self._datei_panel)
        r1.pack(fill=tk.X, padx=4, pady=(1,3))
        ttk.Label(r1, text="Fundliste:", width=15, anchor='w').pack(side=tk.LEFT)
        ttk.Button(r1, text="📂 Öffnen",
                   command=self.cmd_lade_fundliste).pack(side=tk.LEFT, padx=2)
        ttk.Button(r1, text="✕ Löschen",
                   command=self.cmd_entferne_fundliste).pack(side=tk.LEFT, padx=2)
        ttk.Label(r1, textvariable=self._fl_var,
                  foreground='#4a9eff').pack(side=tk.LEFT, padx=6)
        ttk.Label(r1, text="(Löschen entfernt alle aus der Fundliste stammenden Datierungen)",
                  foreground="#888888", font=hint8).pack(side=tk.LEFT)
        # Notebook
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)
        self.tab_fundliste  = FundlisteTab(self.notebook, self)
        self.tab_stellen    = StellenTab(self.notebook, self)
        self.tab_auswertung = AuswertungTab(self.notebook, self)
        self.tab_harris     = HarrisMatrixTab(self.notebook, self)
        self.tab_log        = LogTab(self.notebook, self)
        self.notebook.add(self.tab_fundliste,  text="  Fundliste  ")
        self.notebook.add(self.tab_stellen,    text="  Stellen & Harris  ")
        self.notebook.add(self.tab_auswertung, text="  Auswertung  ")
        self.notebook.add(self.tab_harris,     text="  Harris (Grafik)  ")
        self.notebook.add(self.tab_log,        text="  Log  ")

    def _build_status(self):
        sb = ttk.Frame(self, relief="sunken", borderwidth=1)
        sb.pack(side=tk.BOTTOM, fill=tk.X, padx=2, pady=1)
        self._theme_btn = ttk.Button(
            sb, text="☀ Hell", width=8,
            command=self._toggle_theme)
        self._theme_btn.pack(side=tk.RIGHT, padx=4, pady=1)
        self._sv = tk.StringVar(value="Bereit.")
        ttk.Label(sb, textvariable=self._sv, anchor="w").pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=4)

    def _toggle_theme(self):
        dark = not getattr(self, "_dark", True)
        self._apply_theme(dark)
        self._theme_btn.config(
            text="☀ Hell" if dark else "🌙 Dark")

    def status(self, msg: str):
        self._sv.set(msg)
        self.update_idletasks()

    # ── Kernaktualisierung ────────────────────────────────────
    # ── Undo ──────────────────────────────────────────────────
    def _push_undo(self):
        """Snapshot vor einer manuellen Aktion speichern."""
        self.undo_stack.push(self.stellen)
        self._update_undo_menu()

    def cmd_undo(self, _=None):
        data = self.undo_stack.pop()
        if data is None:
            self.status("Nichts zum Rückgängigmachen.")
            return
        self.stellen = {d["stellennr"]: stelle_aus_dict(d) for d in data}
        self._aktualisiere_alles()
        self._update_undo_menu()
        verbleibend = self.undo_stack.size()
        self.status(f"Rückgängig gemacht.  "
                    f"{'Noch ' + str(verbleibend) + ' Schritt(e) verfügbar.' if verbleibend else 'Keine weiteren Schritte.'}")

    def _update_undo_menu(self):
        try:
            lbl = (f"Rückgängig  [{self.undo_stack.size()}]"
                   if self.undo_stack.can_undo() else "Rückgängig")
            self._mbearbeiten.entryconfig(0, label=lbl,
                                          state="normal" if self.undo_stack.can_undo()
                                          else "disabled")
        except Exception:
            pass

    def _nach_aenderung(self, harris=True, update_stellen_list=True):
        """Volle Kaskadierung nach jeder Beziehungs-/Datierungsänderung.
        harris=False:             Harris-Grafik nicht neu zeichnen.
        update_stellen_list=False: Stellenliste nicht neu aufbauen (für stille
                                   Übergangsspeicherung beim Tab-Wechsel).
        datierungs_beziehungen_berechnen läuft NICHT hier, nur in
        _aktualisiere_alles (Laden / F5). Sonst würden gelöschte Datierungs-
        Linien sofort neu erzeugt.
        """
        auto_datiere_stellen(self.stellen, self.funde)
        propagiere_beziehungen(self.stellen)
        propagiere_datierungen(self.stellen)
        propagiere_stratigraphische_datierungen(self.stellen)
        sortiert = topologische_sortierung(self.stellen)
        self.warnungen = analysiere_warnungen(self.stellen)
        if update_stellen_list:
            self.tab_stellen.aktualisiere()
        try:
            self.tab_auswertung.aktualisiere(sortiert, self.warnungen)
        except Exception as e:
            log.error(f"Auswertung-Aktualisierung: {e}")
        if harris:
            try:
                self.tab_harris.aktualisiere()
            except Exception as e:
                log.error(f"Harris-Aktualisierung: {e}\n{traceback.format_exc()}")

    def _aktualisiere_alles(self):
        auto_datiere_stellen(self.stellen, self.funde)
        propagiere_beziehungen(self.stellen)
        propagiere_datierungen(self.stellen)
        datierungs_beziehungen_berechnen(self.stellen)
        propagiere_stratigraphische_datierungen(self.stellen)
        sortiert = topologische_sortierung(self.stellen)
        self.warnungen = analysiere_warnungen(self.stellen)
        self.tab_stellen.aktualisiere()
        try:
            self.tab_auswertung.aktualisiere(sortiert, self.warnungen)
        except Exception as e:
            log.error(f"Auswertung-Aktualisierung: {e}")
        try:
            self.tab_harris.aktualisiere()
        except Exception as e:
            log.error(f"Harris-Aktualisierung: {e}")

    # ── Befehle ───────────────────────────────────────────────
    def cmd_lade_fundliste(self):
        p = filedialog.askopenfilename(
            title="Fundliste laden",
            filetypes=[("CSV","*.csv"), ("Alle","*.*")])
        if p:
            self._lade_fundliste_pfad(Path(p))

    def _lade_fundliste_pfad(self, pfad: Path):
        try:
            self.fundliste_pfad = pfad
            self.funde = parse_fundliste(pfad)
            self.tab_fundliste.aktualisiere(self.funde)
            self._aktualisiere_alles()
            # Snapshots zurücksetzen (neue Daten könnten Datierungen ändern)
            self.tab_stellen._snapshots.clear()
            self._fl_var.set(pfad.name)
            self.status(f"Fundliste: {len(self.funde)} Funde  ·  {pfad.name}")
        except Exception as e:
            log.error(f"Fundliste: {e}\n{traceback.format_exc()}")
            messagebox.showerror("Fehler",
                f"Fundliste konnte nicht geladen werden:\n{e}")

    def cmd_lade_stellenkatalog(self):
        p = filedialog.askopenfilename(
            title="Stellenkatalog laden",
            filetypes=[("CSV","*.csv"), ("Alle","*.*")])
        if not p: return
        self._lade_stellenkatalog_pfad(Path(p))

    def _lade_stellenkatalog_pfad(self, pfad: Path):
        p = str(pfad)
        try:
            self.stellenkatalog_pfad = Path(pfad)
            neu = parse_stellenkatalog(self.stellenkatalog_pfad)
            for s in neu:
                if s.stellennr not in self.stellen:
                    self.stellen[s.stellennr] = s
                else:
                    ex = self.stellen[s.stellennr]
                    ex.befansprac        = s.befansprac        or ex.befansprac
                    ex.datierung_katalog = s.datierung_katalog or ex.datierung_katalog
                    ex.kein_befund       = s.kein_befund       or ex.kein_befund
                    ex.moderne_stoerung  = s.moderne_stoerung  or ex.moderne_stoerung
                    ex.rest_befund       = s.rest_befund       or ex.rest_befund
                    ex.sichtbar          = s.sichtbar          or ex.sichtbar
                    ex.zeichnung         = s.zeichnung         or ex.zeichnung
                    ex.kommentar         = s.kommentar         or ex.kommentar
                    ex.planum            = s.planum            or ex.planum
                    ex.profil            = s.profil            or ex.profil
                    ex.bearbeiter        = s.bearbeiter        or ex.bearbeiter
                    if s.tiefe_ok is not None: ex.tiefe_ok = s.tiefe_ok
                    if s.tiefe_uk is not None: ex.tiefe_uk = s.tiefe_uk
                    # Handarbeit im GUI gewinnt: Katalog-Beziehungen nur
                    # übernehmen, wenn im Projekt noch nichts eingetragen ist
                    # (sonst löscht jeder Katalog-Reimport die manuelle Arbeit)
                    if s.juenger_als and not ex.juenger_als:
                        ex.juenger_als = s.juenger_als
                    if s.aelter_als and not ex.aelter_als:
                        ex.aelter_als  = s.aelter_als
                    ex.profil_juenger_als = s.profil_juenger_als
                    ex.profil_aelter_als  = s.profil_aelter_als
                    ex.profil_gleich      = s.profil_gleich
            # Profiltext-Beziehungen → manuelle Felder vorausfüllen (nur wenn
            # noch nicht manuell belegt). Profiltexte sind explizite
            # Feldbeobachtungen und damit verlässliche Vorschläge.
            # OKUK-Beziehungen bleiben in okuk_* (nur Vorschlag, unsicher).
            for nr, s in self.stellen.items():
                # Vom Nutzer gelöschte Beziehungen (unterdrueckte_bez) werden
                # auch beim Vorausfüllen respektiert — gelöscht bleibt gelöscht
                if not s.juenger_als and s.profil_juenger_als:
                    s.juenger_als = [x for x in s.profil_juenger_als
                                     if x in self.stellen and x != nr
                                     and f"juenger:{x}" not in s.unterdrueckte_bez]
                if not s.aelter_als and s.profil_aelter_als:
                    s.aelter_als  = [x for x in s.profil_aelter_als
                                     if x in self.stellen and x != nr
                                     and f"aelter:{x}" not in s.unterdrueckte_bez]
                if not s.gleich_alt_wie and s.profil_gleich:
                    s.gleich_alt_wie = [x for x in s.profil_gleich
                                        if x in self.stellen and x != nr
                                        and f"gleich:{x}" not in s.unterdrueckte_bez]
            # OK/UK-Beziehungen automatisch ableiten
            n_okuk = okuk_beziehungen_berechnen(self.stellen)
            self._aktualisiere_alles()
            # Snapshots zurücksetzen damit Reset-Button korrekt arbeitet
            self.tab_stellen._snapshots.clear()
            self._sk_var.set(Path(p).name)
            self.status(
                f"Stellenkatalog: {len(neu)} Stellen  ·  {Path(p).name}"
                + (f"  ·  {n_okuk} OK/UK-Beziehungen" if n_okuk else ""))
        except Exception as e:
            log.error(f"Stellenkatalog: {e}\n{traceback.format_exc()}")
            messagebox.showerror("Fehler",
                f"Stellenkatalog konnte nicht geladen werden:\n{e}")

    def cmd_auto_datiere(self):
        self._aktualisiere_alles()
        self.status("Auto-Datierung abgeschlossen.")

    def cmd_analyse(self):
        try:
            self.status("Analysiere…")
            self._aktualisiere_alles()
            self.warnungen = analysiere_warnungen(self.stellen)
            sortiert = topologische_sortierung(self.stellen)
            self.tab_auswertung.aktualisiere(sortiert, self.warnungen)
            self.tab_harris.aktualisiere()
            self.tab_stellen.aktualisiere()
            n1 = sum(1 for w in self.warnungen if w.stufe == 1)
            n2 = sum(1 for w in self.warnungen if w.stufe == 2)
            n3 = sum(1 for w in self.warnungen if w.stufe == 3)
            self.status(f"Analyse: {len(sortiert)} Stellen  ·  "
                        f"⛔{n1}  ⚠️{n2}  ℹ️{n3}")
            self.notebook.select(self.tab_auswertung)
        except Exception as e:
            log.error(f"Analyse: {e}\n{traceback.format_exc()}")
            messagebox.showerror("Fehler", f"Analyse fehlgeschlagen:\n{e}")

    def cmd_exportiere(self):
        p = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV","*.csv"), ("Text","*.txt")])
        if p:
            try:
                self.tab_auswertung.exportiere(Path(p))
                self.status(f"Exportiert: {p}")
            except Exception as e:
                log.error(f"Export: {e}")
                messagebox.showerror("Fehler", str(e))

    def cmd_entferne_fundliste(self):
        """Fundliste entladen und alle davon abgeleiteten Datierungen entfernen."""
        if not self.funde and not self.fundliste_pfad:
            return
        if not messagebox.askyesno("Fundliste entfernen",
                "Fundliste entfernen?\n\n"
                "Alle Datierungen aus der Fundliste werden entfernt.\n"
                "Manuell eingetragene Werte bleiben erhalten."):
            return
        self._push_undo()   # Zustand vor dem Entfernen sichern (Ctrl+Z möglich)
        self.funde = []
        self.fundliste_pfad = None
        self._fl_var.set("–")
        _UNBEKANNT_FUND.clear()
        # Stellen NUR aus Fundliste (kein Katalog, nichts manuell) entfernen.
        # Der Dialog verspricht "Manuell eingetragene Werte bleiben erhalten" —
        # also wirklich JEDE Art manueller Eingabe prüfen
        entfernen = [nr for nr, s in self.stellen.items()
                     if not s.datierung_katalog and not s.befansprac
                     and not s.juenger_als and not s.aelter_als
                     and not s.gleich_alt_wie and not s.gestoert_von
                     and not s.manuell_datierung
                     and s.manuell_von is None and s.manuell_bis is None
                     and not s.ankerpunkte
                     and not s.kommentar and not s.zeichnung
                     and s.funde]
        for nr in entfernen: del self.stellen[nr]
        # Aus der Fundliste stammende Datierungen entfernen
        for s in self.stellen.values():
            s.funde = []
            s.auto_von = None
            s.auto_bis = None
            s.auto_datierung_str = ""
        self.tab_fundliste.aktualisiere([])
        self._aktualisiere_alles()
        self.tab_stellen._snapshots.clear()
        self.status("Fundliste entfernt. Manuelle Einträge bleiben erhalten.")

    def cmd_entferne_stellenkatalog(self):
        """Stellenkatalog entladen und abgeleitete Beziehungen/Werte entfernen."""
        if not self.stellenkatalog_pfad:
            return
        if not messagebox.askyesno("Stellenkatalog entfernen",
                "Stellenkatalog entfernen?\n\n"
                "Aus Profil/Planum erkannte Beziehungen und OK/UK-Ableitungen\n"
                "werden entfernt. Manuell eingetragene Werte bleiben erhalten."):
            return
        self._push_undo()   # Zustand vor dem Entfernen sichern (Ctrl+Z möglich)
        self.stellenkatalog_pfad = None
        self._sk_var.set("–")
        _UNBEKANNT_KATALOG.clear()
        for s in self.stellen.values():
            # Nur automatisch abgeleitete Felder löschen
            s.profil_juenger_als = []
            s.profil_aelter_als  = []
            s.profil_gleich      = []
            s.okuk_juenger_als   = []
            s.okuk_aelter_als    = []
            s.tiefe_ok           = None
            s.tiefe_uk           = None
            s.sichtbar           = ""
            # datierung_katalog und befansprac NUR löschen wenn NICHT
            # manuell überschrieben (manuell_datierung hat Vorrang sowieso)
        self._aktualisiere_alles()
        self.tab_stellen._snapshots.clear()
        self.status("Stellenkatalog entfernt. Manuelle Einträge bleiben erhalten.")

    def cmd_speichern(self):
        p = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("Projekt","*.json")])
        if not p: return
        try:
            daten = {
                "version": VERSION,
                "fundliste_pfad":      str(self.fundliste_pfad      or ""),
                "stellenkatalog_pfad": str(self.stellenkatalog_pfad or ""),
                "stellen": [stelle_zu_dict(s) for s in self.stellen.values()],
            }
            Path(p).write_text(
                json.dumps(daten, ensure_ascii=False, indent=2),
                encoding="utf-8")
            self.status(f"Projekt gespeichert: {p}")
        except Exception as e:
            log.error(f"Speichern: {e}")
            messagebox.showerror("Fehler", str(e))

    def cmd_laden(self):
        p = filedialog.askopenfilename(
            filetypes=[("Projekt","*.json")])
        if not p: return
        try:
            daten = json.loads(Path(p).read_text(encoding="utf-8"))
            # Alten Zustand vollständig räumen — sonst hängen die Funde des
            # VORHERIGEN Projekts an den neuen Stellen (Phantom-Datierungen)
            self.funde = []
            self.fundliste_pfad = None
            self._fl_var.set("–")
            self.tab_fundliste.aktualisiere([])
            self.tab_stellen._snapshots.clear()
            self.tab_stellen._cur = None
            try: self.tab_harris._zeichenmodus_beenden()
            except Exception: pass
            # Schlüssel über das normalisierte stellennr-Feld (alte Dateien
            # können '004' enthalten — stelle_aus_dict normalisiert)
            geladen = [stelle_aus_dict(d) for d in daten.get("stellen", [])]
            self.stellen = {s.stellennr: s for s in geladen}
            # 'Gleich gewinnt': in gespeicherten Projekten enthaltene jünger/
            # älter-Beziehungen innerhalb einer Gleich-Gruppe bereinigen, damit
            # Stellen & Harris konsistent ist (Nutzerwunsch).
            _bereinige_gleich_ordnung(self.stellen)
            fl = daten.get("fundliste_pfad", "")
            if fl and Path(fl).exists():
                self.fundliste_pfad = Path(fl)
                self.funde = parse_fundliste(self.fundliste_pfad)
                self.tab_fundliste.aktualisiere(self.funde)
                self._fl_var.set(Path(fl).name)
            elif fl:
                log.warning(f"Fundliste nicht gefunden: {fl}")
                messagebox.showwarning("Fundliste fehlt",
                    f"Die im Projekt verlinkte Fundliste wurde nicht "
                    f"gefunden:\n{fl}\n\nDas Projekt wird ohne Funde geladen.")
            sk = daten.get("stellenkatalog_pfad", "")
            self.stellenkatalog_pfad = Path(sk) if sk else None
            self._sk_var.set(Path(sk).name if sk else "–")
            # Unbekannte Datierungsbegriffe aus dem geladenen Stand neu aufbauen
            _UNBEKANNT_FUND.clear(); _UNBEKANNT_KATALOG.clear()
            for s in self.stellen.values():
                kd = (s.datierung_katalog or "").strip()
                if (kd and kd.lower() != "datierung unbekannt"
                        and kd not in WNK_DATIERUNGEN_EINGEBETTET):
                    _UNBEKANNT_KATALOG.add(kd)
            self._aktualisiere_alles()
            self.status(f"Projekt geladen: {p}")
        except Exception as e:
            log.error(f"Laden: {e}\n{traceback.format_exc()}")
            messagebox.showerror("Fehler", str(e))

# ═══════════════════════════════════════════════════════════════
# DESIGN / DARK MODE
# ═══════════════════════════════════════════════════════════════

    def _apply_theme(self, dark: bool = True):
        self._dark = dark
        bg  = "#1e1e2e" if dark else "#f0f0f0"
        fg  = "#cdd6f4" if dark else "#111111"
        ebg = "#313244" if dark else "#ffffff"
        sbg = "#89b4fa" if dark else "#0078d4"
        nbg = "#181825" if dark else "#c8c8c8"
        # Farben merken, damit später geöffnete Dialoge (Toplevel) sie
        # übernehmen können — die wurden bei _theme_widget noch nicht erfasst
        self._theme_bg = bg; self._theme_fg = fg; self._theme_ebg = ebg
        s   = ttk.Style()
        s.theme_use("clam")
        for w in (".", "TFrame","TLabelframe","TLabelframe.Label",
                  "TLabel","TCheckbutton","TRadiobutton","TPanedwindow",
                  "TSeparator"):
            s.configure(w, background=bg, foreground=fg)
        s.configure("TButton",    background=ebg, foreground=fg, borderwidth=1)
        s.configure("TEntry",     fieldbackground=ebg, foreground=fg,
                                  insertcolor=fg)
        s.configure("TCombobox",  fieldbackground=ebg, foreground=fg,
                                  selectbackground=sbg)
        s.configure("TScrollbar", background=nbg, troughcolor=nbg,
                                  arrowcolor=fg, borderwidth=0)
        s.configure("TScale",     background=bg,  troughcolor=nbg)
        s.configure("TNotebook",  background=nbg, tabmargins=[2,5,2,0])
        s.configure("TNotebook.Tab", background=ebg, foreground=fg,
                                     padding=[10,4])
        s.configure("Treeview",   background=ebg, foreground=fg,
                                  fieldbackground=ebg, rowheight=22)
        hover_bg = "#2d2d4a" if dark else "#d0d0d0"   # dunkle Hover-Farbe im Dark-Mode
        s.configure("Treeview.Heading", background=nbg, foreground=fg,
                                        relief="flat", font=("", 9, "bold"))
        s.map("Treeview.Heading",
              background=[("active",  hover_bg), ("pressed", hover_bg)],
              foreground=[("active",  fg),        ("pressed", fg)])
        s.map("TButton",      background=[("active", sbg)],
                              foreground=[("active", "#1e1e2e" if dark else fg)])
        s.map("TCheckbutton", background=[("active", hover_bg)],
                              foreground=[("active", fg)])
        s.map("TRadiobutton", background=[("active", hover_bg)],
                              foreground=[("active", fg)])
        s.map("TNotebook.Tab", background=[("selected", bg)],
                               foreground=[("selected", fg)])
        s.map("Treeview",  background=[("selected", sbg)],
                           foreground=[("selected", "#1e1e2e")])
        s.map("TCombobox", fieldbackground=[("readonly", ebg), ("focus", ebg)])
        # Hauptfenster
        self.configure(background=bg)
        # Alle tk-Widgets rekursiv anpassen
        self._theme_widget(self, bg, fg, ebg)
        # Harris-Canvas + Legende auffrischen
        try:
            self.tab_harris.cv.configure(
                background="#1a1a2e" if dark else "white")
            self.tab_harris._aktualisiere_legende()
            self.tab_harris.aktualisiere()
        except: pass
        # Fundliste: Tree-Tags und Drop-Label anpassen
        try:
            fl = self.tab_fundliste
            fl.drop_lbl.configure(
                background="#1a2a3a" if dark else "#f0f4f8",
                foreground="#7ab3e0" if dark else "#4477aa")
            fl.tree.tag_configure(
                "modern",
                background="#2a1a00" if dark else "#fff0e0",
                foreground="#ffcc88" if dark else "#000000")
        except: pass
        # Auswertung + StellenTab: Tree-Tags für Dark/Light anpassen
        wid_bg  = "#3d0000" if dark else "#ffcccc"
        wid_fg  = "#ffaaaa" if dark else "#880000"
        hin_bg  = "#2a1f00" if dark else "#fff3cd"
        hin_fg  = "#ffe070" if dark else "#664400"
        not_bg  = "#0d3320" if dark else "#d4edda"
        not_fg  = "#a0e0b0" if dark else "#1a5c1a"
        try:
            at = self.tab_auswertung
            for tr in [at.tree, at.warn_tree]:
                tr.tag_configure("wid", background=wid_bg, foreground=wid_fg)
                tr.tag_configure("hin", background=hin_bg, foreground=hin_fg)
                tr.tag_configure("not", background=not_bg, foreground=not_fg)
                tr.tag_configure("ab",  foreground="#888888")
        except: pass
        try:
            st = self.tab_stellen
            st.tree.tag_configure("wid", background=wid_bg, foreground=wid_fg)
            st.tree.tag_configure("ab",  foreground="#888888")
            st.tree.tag_configure("st",  foreground="#ff6644" if dark else "#cc4400")
            st.tree.tag_configure("inf", foreground="#4499ff" if dark else "#0055bb")
        except: pass
        # Log-Terminal
        try: self.tab_log.txt.configure(
            background="#0d1117" if dark else "#1e1e1e",
            foreground="#d4d4d4")
        except: pass

    def _scrollbar_kwargs(self):
        """Farben für KLASSISCHE tk.Scrollbar (z.B. in ScrolledText) — die folgen
        nicht dem ttk-Style und müssen einzeln eingefärbt werden."""
        if getattr(self, "_dark", True):
            return dict(background="#45475a", troughcolor="#181825",
                        activebackground="#585b70", highlightbackground="#181825",
                        borderwidth=0, highlightthickness=0)
        return dict(background="#c4c4c4", troughcolor="#e6e6e6",
                    activebackground="#a8a8a8", highlightbackground="#e6e6e6",
                    borderwidth=0, highlightthickness=0)

    def style_dialog(self, win):
        """Ein zur Laufzeit geöffnetes Toplevel an das aktuelle Theme angleichen.
        ttk-Widgets folgen dem globalen Style automatisch; der Fenster-
        Hintergrund und tk-Widgets (Label/Frame/Text/Entry) müssen aber
        nachgezogen werden, sonst bleibt das Fenster im Dark-Mode weiß.
        Buttons mit eigener Hintergrundfarbe (z.B. Farbwähler) werden in
        Ruhe gelassen."""
        bg  = getattr(self, "_theme_bg",  None)
        fg  = getattr(self, "_theme_fg",  None)
        ebg = getattr(self, "_theme_ebg", None)
        if bg is None:
            return
        try: win.configure(background=bg)
        except Exception: pass
        def _walk(w):
            try:
                wc = w.winfo_class()
                if wc == "Canvas":
                    w.configure(background=bg)
                elif wc == "Text":
                    w.configure(background=ebg, foreground=fg,
                                insertbackground=fg, selectbackground="#89b4fa")
                elif wc == "Entry":
                    w.configure(background=ebg, foreground=fg, insertbackground=fg)
                elif wc in ("Label","Frame","Checkbutton","Radiobutton","Message"):
                    w.configure(background=bg, foreground=fg)
                elif wc == "Toplevel":
                    w.configure(background=bg)
                elif wc == "Button":
                    # Farbwähler-Buttons tragen ihre Vorschaufarbe selbst →
                    # nur Buttons ohne gesetzte Spezialfarbe anpassen
                    cur = str(w.cget("background"))
                    if cur in ("", "SystemButtonFace", bg):
                        w.configure(background=ebg, foreground=fg)
                elif wc == "Scrollbar":
                    w.configure(**self._scrollbar_kwargs())
            except Exception:
                pass
            for c in w.winfo_children():
                _walk(c)
        for c in win.winfo_children():
            _walk(c)

    def _theme_widget(self, w, bg, fg, ebg):
        """Rekursiv tk-(nicht-ttk-)Widgets einfärben."""
        try:
            wc = w.winfo_class()
            if wc == "Canvas":
                w.configure(background=bg)
            elif wc == "Text":
                w.configure(background=ebg, foreground=fg,
                            insertbackground=fg, selectbackground="#89b4fa")
            elif wc in ("Label","Button","Frame","Checkbutton",
                        "Radiobutton","Message"):
                w.configure(background=bg, foreground=fg)
            elif wc == "Entry":
                w.configure(background=ebg, foreground=fg,
                            insertbackground=fg)
            elif wc == "Scrollbar":
                # klassische tk.Scrollbar (z.B. in ScrolledText) — folgt nicht
                # dem ttk-Style, daher hier dunkel einfärben
                w.configure(**self._scrollbar_kwargs())
        except: pass
        for c in w.winfo_children():
            self._theme_widget(c, bg, fg, ebg)


# ═══════════════════════════════════════════════════════════════
# EINSTIEGSPUNKT
# ═══════════════════════════════════════════════════════════════

def main():
    if sys.version_info < (3, 8):
        print(f"Python 3.8+ erforderlich (aktuell {sys.version})")
        sys.exit(1)
    try:
        import tkinter  # noqa
    except ImportError:
        print("tkinter fehlt.  Ubuntu/Debian:  sudo apt install python3-tk")
        sys.exit(1)
    App().mainloop()

if __name__ == "__main__":
    main()

# ── Ende Block 2 ──────────────────────────────────────────────