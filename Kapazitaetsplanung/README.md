# Kapazitätsplanung BB2 Instandhaltung – Umsetzung

| Datei | Inhalt |
|---|---|
| `Kapaplanung_BB2_Instandhaltung_Planung_Umsetzung_optimiert.xlsx` | Überarbeitete Arbeitsmappe |
| `Kapaplanung_BB2_Instandhaltung_Planug_Umsetzung_ORIGINAL.xlsm` | Ausgangsdatei (Stand 21.07.2025), unverändert |
| `build_kapaplanung.py` | Skript, das die optimierte Datei aus dem Original erzeugt (`python build_kapaplanung.py <original.xlsm> <ziel.xlsx>`) |

## 1. Analyse der Ausgangsdatei

**Zweck:** Ermittlung des Personalbedarfs (FTE) im Umsetzungsteam der Instandhaltung BB2.
Rechenweg: Std. je Arbeitstag × 5 = Std./Woche, × 250 = Std./Jahr, ÷ 1.750 = FTE netto;
zuzüglich Urlaub (38/220 = 17,3 %) und Krankheit (4,1 %) ergibt sich der Brutto-Bedarf.
**Ergebnis:** 28,1 Std./Tag · 7.025 Std./Jahr · **4,01 FTE netto** · **4,87 FTE brutto**.

### Stärken
- Kompakt auf einer Seite, druckfertig, durchgängig mit Formeln gerechnet.
- Breites Tätigkeitsspektrum (19 Positionen), Netto- und Bruttobedarf getrennt, Quelle der Krankenquote angegeben.

### Schwächen
- Spaltenkopf „Dauer/Min./Teil“ ist falsch – gerechnet wird mit **Std. je Arbeitstag**. Spalte C (Std./Woche) hat keine Überschrift.
- Parameter stehen fest in den Formeln (×5, ×250, /1750, /7, 38/220); 1.750 und 7 werden getrennt gepflegt, obwohl sie voneinander abhängen.
- Das Kreisdiagramm erfasst nur Zeilen 4–18. Es fehlen 1.250 Std./Jahr (Balg-Bereich, Wartung, Administration), deshalb sind die Prozentwerte zu hoch.
- Farbmarkierungen, Namensliste (Nr. 1, 2, 4, 5) und „*Schichtübergreifend“ sind nicht erklärt.
- „Summe Direkte Mitarbeiter“ steht sowohl für netto als auch für brutto. Dazu kommen Rechtschreibfehler.
- Es gibt keinen Abgleich mit dem verfügbaren Personal und keine Verantwortlichkeiten, Risiken oder Ist-Werte.

### Fehlende Inhalte
- Die präventive Instandhaltung ist mit **0 Std.** bewertet (Wartung Betriebsmittel / allgemeine Maschinen).
- Es fehlen: Störungsbehebung, Inspektion/gesetzliche Prüfungen, Arbeitssicherheit, Fremdfirmensteuerung, Audits/QM, Energiemanagement, Einarbeitung/Wissenssicherung, Weiterbildungstage.

### Verbesserungspotenzial
Zentrale Parameter, Kategorisierung (Tätigkeitsbereich + DIN 31051), Risikobewertung, Kapazitätsabgleich mit Ampel, Dashboard, Plan-Ist-Vergleich, Kontrollformeln.

### Neue Struktur
Übersicht → Dashboard → Kapazitätsbedarf (Eingabe) → Personalbedarf → Parameter → Analyse → Änderungsprotokoll → Original (Archiv)

## 2. Wesentliche Änderungen und Begründung

| Änderung | Warum es besser ist |
|---|---|
| Alle Annahmen auf Blatt **Parameter** mit Quelle; 7 Std./Tag = 1.750/250 wird abgeleitet | Eine Änderung wirkt überall. Werte können sich nicht mehr widersprechen. |
| Einheit korrigiert („Std. / Arbeitstag“), alle Spalten beschriftet, FTE mit 2 Nachkommastellen | Die Tabelle wird richtig gelesen. Kleine Posten verschwinden nicht mehr als „0,0“. |
| Einheitliche Begriffe (FTE, Personalbedarf netto/brutto), Rechtschreibung korrigiert | Für Instandhaltung, Produktion, QM und Management gleich verständlich. |
| Kategorien: 9 Tätigkeitsbereiche + Instandhaltungsart nach **DIN 31051** | Gemeinsame Fachsprache; Auswertung nach Bereichen; Präventivquote messbar. |
| **8 ergänzte Tätigkeiten** ohne Stundenwert (Status „offen – zu bewerten“) | Die Lücken werden sichtbar, das Ergebnis ändert sich aber erst nach der Bewertung. |
| **Risikobewertung** (Auswirkung × Wahrscheinlichkeit, Skala 1–5, Ampel) als lila markierter Erstvorschlag | Ermöglicht eine Priorisierung und begründet den Personalbedarf gegenüber dem Management. |
| Spalten Verantwortlich, Schichtübergreifend, Datenbasis, Ist-Std., Abweichung | Klare Zuständigkeit, Transparenz über die Schätzgüte und einen Plan-Ist-Regelkreis. |
| **Kapazitätsabgleich** Team ↔ Brutto-Bedarf, Deckungsgrad, Ampel | Die Kernaussage für das Management (Unterdeckung) ist sofort sichtbar. |
| **Methodenvergleich** Zuschlag (Original, 4,87) vs. Verfügbarkeit (5,10 FTE) | Zeigt die methodische Bandbreite. Das Original bleibt die Basis. |
| **Dashboard**: KPI-Kacheln, Status, Bereiche, DIN-Arten, Pareto Top 10, Top-Risiken, 3 Balkendiagramme | Automatische Auswertung ohne Handarbeit; korrekte Diagramme über alle Positionen. |
| Kontrollzeilen „Wert Original“ mit ✔-Abgleich | Belegt, dass keine Bewertung verändert wurde. |
| Dropdowns, Eingabeprüfung (0–24 Std.), Filter, fixierte Kopfzeilen, Navigationslinks | Weniger Eingabefehler, schnellere Bedienung. |
| Farbmarkierungen und Namensliste als Text erhalten (Spalte „Kennzeichnung im Original“) | Keine Information geht verloren; die Bedeutung wird geklärt statt geraten. |
| Original als unveränderte Archivkopie im Blatt „Original 21.07.2025“ | Nachvollziehbarkeit (QM/Revision). |
| `.xlsx` statt `.xlsm` | Die Datei enthielt keine Makros, daher entfällt die Sicherheitswarnung. |

## 3. Zentrale Erkenntnisse
- Projekte/KVP (≈36 %) und Facility Management (≈21 %) binden zusammen mehr als 55 % der Kapazität.
- Der Präventivanteil (Wartung + Inspektion) liegt bei **0 %**. Zu klären ist, ob die Schicht-Instandhaltung diese Aufgaben abdeckt.
- Unter der Annahme 4 Personen × 1,0 FTE fehlen **≈0,87 FTE** (Deckungsgrad ≈82 %).
- Acht ergänzte Tätigkeiten sind noch unbewertet. Der tatsächliche Bedarf liegt daher voraussichtlich höher.

## 4. Vorgeschlagene zusätzliche Kennzahlen
Deckungsgrad, Präventivquote, Planungsgrad, Wartungsplan-Erfüllung, MTBF/MTTR, technische Verfügbarkeit, Plan-Ist-Abweichung, Ersatzteilverfügbarkeit, Überstundenquote, Instandhaltungskostenquote. Details mit Formeln und Datenquellen stehen auf dem Blatt „Analyse“.

## 5. Offene Punkte (vom Team zu klären)
1. Bedeutung der Farbmarkierungen und der Namensliste im Original
2. Zuständigkeit für Wartung und Inspektion (Schicht-IH oder Umsetzung)
3. Bewertung der 8 ergänzten Tätigkeiten
4. Validierung der Risikovorschläge
5. Bezugsbasis der 220 AT und damit die Methode für die Verteilzeiten
6. Tatsächliche FTE je Person (derzeit Annahme 1,0)
