# Personalbedarf Instandhaltung BB2 – Umsetzung

| Datei | Inhalt |
|---|---|
| `Personalbedarf_BB2_Instandhaltung_einfach.xlsx` | Einfache Personalbedarfsrechnung (1 Blatt) + Originalblatt zum Nachschlagen |
| `Kapaplanung_BB2_Instandhaltung_Planug_Umsetzung_ORIGINAL.xlsm` | Ausgangsdatei (Stand 21.07.2025), unverändert |
| `build_kapaplanung.py` | Erzeugt die einfache Datei aus dem Original |

## So funktioniert die Rechnung

1. Für jede Aufgabe eintragen: **Stunden pro Tag**
2. Stunden pro Jahr = Stunden pro Tag × 250 Arbeitstage
3. Mitarbeiter = Stunden pro Jahr ÷ 1.750 Stunden je Mitarbeiter
4. Dazu kommen Urlaub (38 von 220 Tagen = 17,3 %) und Krankheit (4,1 %)

## Ergebnis (Stand 21.07.2025)

| | Mitarbeiter |
|---|---|
| Für die Aufgaben | 4,0 |
| + Urlaub | 0,7 |
| + Krankheit | 0,2 |
| **Benötigt gesamt** | **4,9 (aufgerundet 5 Personen)** |
| Vorhanden | 4 |
| **Es fehlen** | **0,9** |

Gelbe Felder können geändert werden, alles andere rechnet automatisch.
