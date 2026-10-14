# Maintenance BB2 – Staffing & Capacity Dashboard

| Datei | Inhalt |
|---|---|
| `Maintenance_BB2_Staffing_Dashboard.xlsx` | Englische Arbeitsmappe: **Management Dashboard** + **Capacity Planning** (Eingabe) |
| `Kapaplanung_BB2_Instandhaltung_Planug_Umsetzung_ORIGINAL.xlsm` | Ausgangsdatei (Stand 21.07.2025), unverändert |
| `build_kapaplanung.py` | Erzeugt die Arbeitsmappe (`python build_kapaplanung.py <ziel.xlsx>`) |

## Bedienung
Auf dem Blatt **Capacity Planning** nur die gelben Felder ausfüllen (vor allem *Hours per day*).
Woche, Monat, Jahr, Mitarbeiter (FTE), Anteile, KPI-Kacheln und das Kuchendiagramm rechnen automatisch.

## Ergebnis (Stand 21.07.2025)
- Total hours per year: **7,025 h**
- Required headcount: **4.9** (5 persons rounded up) · current headcount 4 · staffing gap **−0.9**
- Avg. workload per employee: **1,756 h/year** (146 h/month, 35.1 h/week) vs. capacity 1,750 h → utilisation 100 %
