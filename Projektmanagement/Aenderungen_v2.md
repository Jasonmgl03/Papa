# Projektmanagement WI – Version 2

## Eingabe (vereinfacht)
- **Ein Blatt „Zeiterfassung“** ersetzt die sechs Blätter Eingabe_Benny/Jose/Linda/nn1/nn2/nn3.
- Die Personen-Buttons auf dem Startblatt bleiben; sie öffnen die Zeiterfassung mit der passenden Person.
- Felder: Person · Datum (vorbelegt mit heute, änderbar) · Projekt (nur aktive Projekte) · Stunden · Kosten · Notiz.
- Prüfungen beim Speichern (Pflichtfelder, Stunden 0,25–24, gültiges Datum/Projekt/Person) mit klarer Meldung.
- Nach dem Speichern bleiben Person und Datum stehen; die Pivots/Dashboards werden automatisch aktualisiert.
- Unten auf dem Blatt: die letzten 5 Einträge der gewählten Person zur Kontrolle.
- Neue Person: nur in der Personenliste (Blatt „Personen“) ergänzen – kein neues Blatt, kein neues Makro.

## Korrigiert
- Eingabe_nn1–nn3 waren defekt (fehlende Felder, nn1 speicherte als nn2) → entfällt durch das gemeinsame Blatt.
- Datum war immer „heute“ → jetzt frei wählbar.
- Leere Zeile in „Projekte“ (leerer Eintrag in der Projektauswahl) und leere Zeile in „Daten“ entfernt.
- Spalte „Datum Heute“ → „Projektstart“; Zeitfortschritt im Dashboard rechnet ab Projektstart statt ab 1. Januar.
- „Projekt erstellen“: neues Feld Projektstart, Prüfung Ende ≥ Start, doppelte Namen werden abgelehnt.
- Auslastung je Person: Stunden des laufenden Jahres im Verhältnis zur anteiligen Jahreskapazität (vorher Gesamtstunden ÷ volle Jahreskapazität → Ampel fast immer grün). Spalte „Frei“ zeigt die Reststunden.
- Projekte: Fortschritt > 100 % und negative Reststunden werden rot markiert.
- #DIV/0!-Fehler im Rechenbereich des Dashboards beseitigt.
- Pivots aktualisieren sich beim Öffnen der Datei, nach dem Speichern und beim Wechsel ins Dashboard.

## Unverändert
- Design (Hintergrund, Farben, Buttons, Navigation), Dashboards, Diagramme, alle bisherigen Buchungen.
- Buchungen, bei denen in der Notiz ein Datum steht (36 Stück), wurden bewusst nicht verändert.
- Benötigt weiterhin Microsoft 365 (FILTER, XLOOKUP, LET).
