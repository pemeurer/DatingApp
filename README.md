# DatingApp

Datierungshilfe für archäologische Grabungsdaten. Das Programm führt Fundlisten
und Stellenkataloge zusammen, leitet daraus die Datierung der einzelnen Stellen
ab und stellt die stratigrafischen Beziehungen als Harris-Matrix dar.

Version 3.2.0. Läuft offline mit Python 3 und Tkinter.

## Start

```bash
python DatingApp.py
```

Unter Windows genügt ein Doppelklick auf `start.bat`.

Optional, nicht erforderlich:

| Paket | Nutzen |
|---|---|
| `tkinterdnd2` | Dateien per Ziehen und Ablegen laden |
| `pillow` | zusätzliche Bildfunktionen |

## Kurzanleitung

1. Fundliste laden, im Reiter „Fundliste" über den Dateidialog oder per Ziehen
   und Ablegen.
2. Stellenkatalog laden über das Menü „Datei".
3. Stellen im Reiter „Stellen & Harris" ordnen und verknüpfen.
4. Analyse mit F5 starten.
5. Ergebnis im Reiter „Auswertung" ansehen, die Grafik im Reiter
   „Harris (Grafik)".
6. Projekt mit Strg+S sichern.

## Was das Programm leistet

- **Daten zusammenführen.** Fundlisten und Stellenkataloge werden eingelesen und
  zu einem gemeinsamen Bestand verbunden.
- **Datierung berechnen.** Aus der Fundlage entsteht je Stelle eine effektive
  Datierung. Das Programm hält fest, woher jede Datierung stammt, und führt
  mehrere Anhaltspunkte zusammen.
- **Auffälligkeiten zeigen.** Widersprüche und unsichere Zuordnungen werden in
  der Tabelle markiert statt stillschweigend geglättet.
- **Nachvollziehbar bleiben.** Zu jeder Datierungsentscheidung ist erkennbar,
  auf welchen Funden sie beruht.

## Begriffslisten

Für Befundansprachen und Datierungen sind Begriffslisten eingebaut, die
Eingabefelder filtern die Vorschläge beim Tippen.

## Log

Das Programm schreibt ein Protokoll nach `datierungshilfe.log` im
Benutzerverzeichnis. Es hilft, wenn etwas nicht wie erwartet läuft.

## Hinweis zu Grabungsdaten

Fundlisten, Stellenkataloge und Projektdateien enthalten Daten realer Grabungen
und gehören nicht in dieses Repository. Die `.gitignore` hält Tabellen und
Bilddateien draußen.

## Lizenz

GNU General Public License Version 3, siehe [LICENSE](LICENSE).

Keine Gewährleistung. Die Nutzung erfolgt auf eigene Verantwortung, der Autor
übernimmt keine Haftung für Schäden, Datenverluste oder Fehlinterpretationen der
Ergebnisse. Die berechneten Datierungen sind ein Vorschlag und ersetzen die
fachliche Beurteilung nicht.

© 2026 Peter Meurer
