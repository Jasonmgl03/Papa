# Wartungsplaner

Ein Wartungsplan-Tool für Maschinen und Werkzeuge, das **komplett lokal** läuft – ohne Installation, ohne Internet, ohne Server, ohne Konto.

## Starten

1. `Wartungsplaner.html` doppelklicken – die Datei öffnet sich im Browser (Chrome, Edge oder Firefox).
2. Zum Ausprobieren auf **„Beispieldaten laden“** klicken – oder direkt eigene Objekte anlegen.

Tipp: Die Datei als Lesezeichen speichern oder eine Verknüpfung auf den Desktop legen.

## Funktionen

| Bereich | Was du damit machen kannst |
|---|---|
| **Übersicht** | Anzahl überfälliger, bald fälliger und erledigter Aufgaben; Liste der nächsten 30 Tage mit „Erledigt“-Knopf |
| **Maschinen & Werkzeuge** | Objekte mit Inventar-Nr., Standort, Hersteller, Modell und Seriennummer erfassen; farbiger Status zeigt auf einen Blick, ob etwas überfällig ist |
| **Wartungsplan** | Alle Aufgaben nach Fälligkeit sortiert; Filter nach Objekt, Status und Zeitraum, dazu eine Suche |
| **Jahresplan** | Übersicht über 12 Monate: geplant (●), überfällig (rot), erledigt (✓) |
| **Historie** | Nachweis aller durchgeführten Wartungen (Datum, durchgeführt von, Bemerkung) |
| **Erledigt** | Wird eine Aufgabe als erledigt markiert, berechnet das Tool die nächste Fälligkeit automatisch aus dem Intervall |

### Drucken (oder als PDF speichern)
- **Wartungsplan** – Liste mit Statusfarben und einer Spalte „Erledigt“ zum Abhaken
- **Jahresplan** – Übersicht auf einer Seite im Querformat
- **Wartungskarte** je Maschine – Stammdaten, Aufgaben und ein Wartungsnachweis mit leeren Zeilen und Unterschriftenfeld, zum Aufhängen an der Maschine
- **Alle Wartungskarten** auf einmal (eine Seite pro Maschine)

Im Druckdialog kannst du auch „Als PDF speichern“ wählen.

### Excel
- **Excel (.xlsx)** – eine echte Excel-Datei mit den Blättern *Wartungsplan*, *Objekte*, *Jahresplan* und *Historie*, inklusive Filter, Datumsformat und Farbmarkierung. Sie wird direkt im Browser erzeugt, ohne externe Bibliothek.
- **CSV** – mit Semikolon getrennt, öffnet sich direkt im deutschen Excel.

## Daten & Sicherung

- Alle Daten werden **nur im Browser auf diesem PC** gespeichert (localStorage). Es wird nichts verschickt.
- Unter **„Export & Daten“ → „Backup speichern“** lassen sich alle Daten als `.json`-Datei sichern. Das solltest du regelmäßig tun – auch, um auf einen anderen PC umzuziehen („Backup laden“).
- Achtung: Wer im Browser den Verlauf bzw. die Website-Daten löscht, löscht damit auch die Daten des Tools. Deshalb regelmäßig ein Backup speichern.

## Einstellungen
Unter „Export & Daten“:
- **Firma / Abteilung** – erscheint auf jedem Ausdruck
- **Vorwarnzeit** – wie viele Tage vor der Fälligkeit eine Aufgabe als „Bald fällig“ angezeigt wird (Standard: 14 Tage)

---

# Instandhaltungs-Audit

`Instandhaltungs-Audit.html` ist die digitale Fassung der Excel-Vorlage **Vorlage_Audit_Instandhaltung.xlsm**. Sie läuft wie der Wartungsplaner komplett lokal (Doppelklick, kein Internet nötig) und ist für Handy, Tablet und PC ausgelegt.

| Bereich | Inhalt |
|---|---|
| **Übersicht** | Kennzahlen, Erfüllungsgrad im Verlauf, letztes Audit mit Auswertung je Kapitel, offene Maßnahmen, Maschinenpark nach Kritikalität A/B/C |
| **Audits** | Vier Checklisten aus der Vorlage: *Instandhaltung allgemein* (39 Punkte, DE/EN), *Maschine* (13), *Werkzeuge* (10), *Kompakt-Audit Anlage · Werkzeuge · Ersatzteile* mit den Normfragen aus ISO/TS 16949 (24). Bewertung N.I.O / bed. OK / OK / n. b. wie in der Vorlage |
| **Bericht** | Ergebnis, Diagramme, Maßnahmenplan, alle Prüfpunkte und Unterschriftsfelder – druckbar bzw. als PDF |
| **Maßnahmen** | Alle Befunde mit N.I.O oder bed. OK, mit Verantwortlichem, Termin und Status (offen / überfällig / erledigt) |
| **Maschinen** | Stammblatt mit den 28 Angaben aus „Checkliste_Maschine_Daten“; die drei ausgefüllten Maschinen aus der Vorlage sind bereits übernommen |
| **Wissen** | DIN 31051, DIN EN 13306, DIN EN 15341, IATF 16949 8.5.1.5, BetrSichV, TRBS 1112, DGUV V3, AwSV, Strategien, Kennzahlen mit Rechner (MTBF, MTTR, Verfügbarkeit, OEE), Quellen |
| **Daten** | Backup (.json) speichern/laden, Maßnahmen als CSV für Excel, Einstellungen (Firma, Gewichtung „bed. OK“, Ampelgrenzen) |

**Erfüllungsgrad** = (OK + Gewicht × bed. OK) ÷ (N.I.O + bed. OK + OK). Die Vorlage enthält dafür keine Formel; Gewicht (Standard 50 %) und Ampelgrenzen (grün ab 85 %, gelb ab 60 %) sind deshalb einstellbar.

## Excel-Fassung

`Instandhaltungs-Audit.xlsx` enthält dieselben Checklisten als Excel-Mappe ohne Makros:
*Übersicht* (Einstellungen, Ergebnis je Checkliste, Diagramme), *Audit_Allgemein*, *Audit_Maschine*, *Audit_Werkzeuge*, *Audit_Kompakt* (Bewertung per Auswahlliste, Ampelfarben, Auswertung je Kapitel mit Diagramm), *Maßnahmen* (sammelt alle N.I.O/bed. OK automatisch), *Maschinen* (Stammdaten-Register), *Stammblatt* (eine Maschine druckfertig), *Wissen* (Normen, Strategien, Kennzahlen-Rechner) und *Anleitung*. Eingaben nur in gelben Zellen.
