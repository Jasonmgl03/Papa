# Wartungsintervallberechnung große Wartung – horizontale Simultanbalgmaschine (HSIM)

**Vorversion:** Stand 01 / 03.02.2009 (`.xlsm`, 10 Blätter, Makro „NeueZeile“)
**Aktuelle Version:** Stand 03 (Entwurf) / 17.10.2026 – `Wartungsplan_grosse_Wartung_horizontale_Balgmaschine_Stand03.xlsx`
(Stand 02 ist in der Git-Historie enthalten.)

---

## Stand 03 – Vereinfachung und flexible Intervalle

### Analyse von Stand 02

| Problem | Folge |
|---|---|
| 11 Blätter, Wartungsplan mit 36 Spalten | Unübersichtlich; man muss viel scrollen und zwischen Blättern wechseln. |
| Termine in einem eigenen Blatt „Durchführungsnachweis“ | Doppelte Pflege (Nr., Maßnahme, Intervall stehen zweimal); Risiko von Abweichungen. |
| Intervalle nur in Wochen (täglich … 16 Wochen, halbjährlich, jährlich) | Keine Monatsintervalle, keine mehrjährigen Prüfungen, keine freie Eingabe; Termine nicht kalendergenau. |
| Für ein Intervall drei Spalten (berechnet / festgelegt / Wochen) | Unklar, wo man etwas einträgt. |
| Risikobewertung (selten geändert) zwischen den Tagesdaten | Lenkt im Alltag ab. |
| Dashboard ohne Liste „Was ist jetzt zu tun?“ | Fällige Maßnahmen müssen im Plan gesucht werden. |

### Lösung (umgesetzt)

- **7 statt 11 Blätter:** Start · Dashboard · Wartungsplan · Bewertung & Methodik · Anlagenübersicht · Einstellungen · Dokumentation.
- **Ein Wartungsplan für alles:** Eine Zeile pro Maßnahme enthält Maßnahme, Intervall, letzte Durchführung, nächsten Termin, Ampel und Nachweis. Die Tagesansicht hat 16 Spalten. Risikobewertung und Vorversion lassen sich über [+] aufklappen, technische Hilfsspalten sind ausgeblendet.
- **Intervalle:**
  - Auswahlliste mit wöchentlich, monatlich, alle 3 Monate, alle 4 Monate, halbjährlich und jährlich, dazu 8/16 Wochen (Ergebnis der RPZ-Methode) sowie 2 und 5 Jahre.
  - Freie Eingabe als „Zahl + Einheit“, z. B. „2 Jahre“, „18 Monate“, „10 Tage“, „alle 2 Jahre“.
  - Monate und Jahre werden kalendergenau gerechnet (EDATE).
  - Nicht erkannte Eingaben werden rot als „Intervall prüfen“ markiert.
  - Den Katalog kann man in den Einstellungen erweitern.
- **Ampel:**
  - Rot = überfällig.
  - Gelb = fällig innerhalb der Vorwarnzeit. Standard sind 14 Tage, höchstens aber ¼ des Intervalls, damit wöchentliche Maßnahmen nicht dauerhaft gelb sind.
  - Grün = in Ordnung.
  - Grau = nicht erfasst, Intervall fehlt oder ist zu prüfen.
- **Dashboard:**
  - Kacheln für Überfällig, Fällig, In Ordnung, Erfüllungsgrad, Maßnahmen gesamt, nicht erfasst, Intervall fehlt/prüfen und Erfassungsgrad.
  - Liste „Zu erledigen“, nach Termin sortiert.
  - Maßnahmen pro Intervall mit Diagramm, Status und Risiko je Baugruppe.
  - Die Risikoauswertungen aus Stand 02 bleiben erhalten.
- **Zusammengeführte Spalten:**
  - „Hinweis“ und „Hilfsmittel“ zu einer Spalte.
  - „Begründung“ und „Bemerkung“ zu einer Spalte.
  - Die Positionsnummer steht jetzt im Baugruppentitel.
  - Für das Intervall gibt es nur noch „Empfehlung (RPZ)“ (farbig nach Risikoklasse) und „Intervall“ als Eingabe. Abweichungen von der Empfehlung werden blau markiert.

**Erhalten:** alle Maßnahmen, Texte, Bilder, Bewertungen und Verknüpfungen, alle Originaltexte und alten Nummern sowie alle Kennzahlen aus Stand 02. RPZ und Empfehlung sind für alle 88 Maßnahmen identisch zur Vorversion, maschinell geprüft.

**Bewusst geändert:** Für die Klasse „kritisch“ wird jetzt „monatlich“ statt „4 Wochen“ empfohlen. Derzeit ist keine Maßnahme betroffen.

**Offene Entscheidung:** Die RPZ-Methode empfiehlt 8 bzw. 16 Wochen. Wenn stattdessen Standardintervalle gelten sollen (z. B. hoch → monatlich, mittel → alle 3 oder 4 Monate), stellt man das unter Einstellungen → Abschnitt 2 um. Es wirkt sofort auf alle Maßnahmen ohne manuelle Festlegung.

---

# Stand 02 – Analyse und Überarbeitung der Vorversion

---

## Teil A – Analyse der bestehenden Datei

### Zweck und Bewertungssystematik (Ist-Zustand)

Die Datei legt für 87 Wartungsmaßnahmen an 21 Baugruppen die Wartungsintervalle risikobasiert fest:

| Größe | Bedeutung | Skala |
|---|---|---|
| F, IH, Q, AS, US, S | Bedeutung eines Ausfalls: Funktionalität, Instandsetzungsaufwand, Qualität/Kunde, Arbeitssicherheit, Umweltschutz, Schadenshöhe | 1–5 |
| **B** | = √(F²+IH²+Q²+AS²+US²+S²) / 2 | 1,22–6,12 |
| A | Ausfallwahrscheinlichkeit (> 3 Jahre … alle 6–8 Wochen) | 1–5 |
| E | Entdeckungswahrscheinlichkeit (sofort … nach 1 Monat) | 1–5 |
| SF | Schichtfaktor (0,33 / 0,66 / 1) | – |
| **RPZ** | = B × A × E × SF | |
| Auswertung | RPZ ≥ 66 → 8 Wochen · ≥ 15 → 16 Wochen · ≥ 1 → jährlich | |

Ergebnis der Vorversion: 6 × 8 Wochen, 56 × 16 Wochen, 24 × jährlich, 1 × ohne Intervall. Maximale RPZ 90.

### 1. Stärken

- **Durchdachte, mehrdimensionale Methode:** Sechs Folgekriterien plus Ausfall- und Entdeckungswahrscheinlichkeit – deutlich fundierter als reine Herstellerintervalle; FMEA-nah und damit anschlussfähig an IATF 16949.
- **Transparente Skalen:** Für jedes Kriterium gibt es eine 5-stufige Beschreibung, teils mit messbaren Ankern (Zeiten bei IH, A, E).
- **Praxisnah:** Konkrete Tätigkeiten mit Schmierstoffen, Hubzahlen, Grenzwerten (Öltemperatur < 75 °C), Sicherheitshinweisen und Fotos der Schmierstellen.
- **Anlagenbezug:** Übersichtszeichnung mit Positionsnummern; Baugruppen danach gegliedert.
- **Automatische Intervallermittlung** per Formel; Wiederverwendung von Bewertungen für gleichartige Objekte.
- **Grundlegende Dokumentenlenkung** (Verantwortlicher, Stand, Druckdatum in der Fußzeile).

### 2. Schwächen

| # | Schwäche | Folge |
|---|---|---|
| 1 | **Formellücke:** RPZ > 125 und RPZ < 1 liefern keine Auswertung (leere Zelle). RPZ bis 153 ist möglich. | Hochrisiko-Maßnahmen bekämen kein Intervall. |
| 2 | **Widersprüchliche Methodik-Doku:** Kommentar „RPZ = B × A × E“ (ohne SF) vs. Blatt „Berechnung RPZ“ mit SF; Grenzen dokumentiert als „16 < RPZ < 65“, Formel nutzt ≥ 15 / ≥ 66. | Unklar, was gilt; Lücken (15, 65–66) in der Doku. |
| 3 | **Unvollständige Bewertungen:** 4.1.4 Hauptführung Y-Schlitten ohne A/E (RPZ 0); Abschnitt 17 Druckfilter ganz ohne Bewertung. | Zwei Maßnahmen ohne Intervall, unbemerkt. |
| 4 | **Irreführende Summenzeile:** Σ B, Σ A, Σ E sind fachlich bedeutungslos; „RPZ gesamt“ ist ein Mittelwert nur über Zeilen 8–133 (Abschnitte 19–21 fehlen, RPZ 0 zählt mit). | Falsche Managementkennzahl (28,9). |
| 5 | **Inkonsistente Nummerierung:** Überschriften als 5.1.1 / 11.1.1, Maßnahmen als 9.01–9.03, Abschnitt 6 ohne Baugruppe, Verschlussbalken-Schmierung unter „Hydraulikzylinder“ einsortiert. | Schwer referenzierbar (z. B. in SAP-PM oder Audits). |
| 6 | **Uneinheitliche Formulierungen und Tippfehler** („Beduetung“, „Verschlußbalcken“, „Werkzuegträger“, „abschmiern“, „Riffen“, „vorn Arbeitsbegin“, „hoche“, „Generall-Überhollung“ …), Mischung aus „Prüfe“, „kontrolliere“, „überprüfe“. | Unprofessioneller Eindruck, Missverständnisse. |
| 7 | **Defektes Makro „NeueZeile“:** kopiert die Kopfzeile (Zeile 6) über die aktive Zeile und überschreibt Inhalte. `.xlsm` wird von IT-Richtlinien oft blockiert. | Datenverlust-Risiko, Sicherheitswarnungen. |
| 8 | **Tote Verknüpfungen:** 13 Namensbereiche verweisen auf eine externe Datei „2006 01 03 Jahresplanung REM.xls“. | Verknüpfungswarnungen beim Öffnen. |
| 9 | **8 nahezu gleich aufgebaute Skalen-Blätter** mit je 5 Zeilen Inhalt. | Viel Blätterei, keine Vergleichbarkeit. |
| 10 | **Spalte „Objekt“ nur mit Bildern**, Bilder frei über Zellen gelegt; Warn-/Info-Symbole nur als Grafik. | Nicht filter-/sortierbar; Bilder verrutschen. |
| 11 | **Kopierte Bewertungen (29×) ohne Kennzeichnung**, teils fachlich fraglich (Pneumatikventile mit US = 5 wie Hydraulik; Lüfter = Filtermatten). | Ungeprüfte Übernahmen; Änderungen wirken unbemerkt weiter. |
| 12 | **Kein Durchführungsnachweis, keine Fälligkeiten, kein Status.** | Kein Nachweis der vorbeugenden Instandhaltung (IATF 8.5.1.5). |
| 13 | Bewertungsstufen heißen „sehr gut … sehr schlecht“, obwohl 5 = höchstes Risiko. | Verwechslungsgefahr. |

### 3. Fehlende Inhalte

- Stammdaten (Equipmentnummer, Standort/Kostenstelle, Hersteller, Baujahr) – Felder waren leer.
- Freigabe (Instandhaltung, Produktion, QM), Überprüfungsrhythmus der Bewertung, Änderungshistorie.
- Sollzustand/Prüfkriterium – nur bei 2 von 87 Maßnahmen vorhanden.
- Zuständigkeit (Bediener / IH Mechanik / IH Elektrik / Fremdfirma) und Plan-Zeiten.
- Möglichkeit, ein abweichendes Intervall (Herstellervorgabe, Gesetz) begründet festzulegen.
- Kennzeichnung sicherheits-, qualitäts- und umweltrelevanter Maßnahmen.
- Kennzahlen/Dashboard (Risikoverteilung, Termintreue, Aufwand).
- Gesetzliche/normative Prüfungen: DGUV V3 (Elektrik), BetrSichV (Schutzeinrichtungen, Druckübersetzer), Hydraulikschlauch-Austausch (DGUV Regel 113-020), Ölwechsel/Ölanalyse, Filterelementwechsel.
- Positionen 4, 7, 9, 10, 12, 18, 21, 23, 27, 28 der Übersichtszeichnung kommen im Plan nicht vor.

### 4. Verbesserungspotenzial

- Methodik **unverändert lassen**, aber Lücken schließen und Grenzwerte als Parameter zentralisieren.
- Ampellogik (Risikoklasse) und automatische Plausibilitätsprüfung (unvollständige Bewertungen sichtbar).
- Eingabeprüfung (nur 1–5) und Skalen-Hilfetexte direkt in der Zelle.
- Schichtfaktor einmal zentral statt 87× einzeln.
- Fälligkeitssteuerung und Kennzahlen (PM-Compliance, Überfälligkeiten, n.i.O.-Quote, Aufwand).
- Rückverfolgbarkeit: alte Nummer und Originaltext in jeder Zeile.
- Makrofreie Datei mit Excel-Tabelle statt Makro.

### 5. Empfohlene neue Struktur (umgesetzt)

| Blatt | Zweck | Zielgruppe |
|---|---|---|
| Deckblatt | Stammdaten, Schichtmodell, Dokumentenlenkung, Freigabe, Historie, Anleitung, Legende | alle |
| Dashboard | Kennzahlen, Risikoverteilung, Intervalle, Kriterienprofil, Risikomatrix A×E, Top-10, Baugruppen + Diagramme | Management, QM, IH-Leitung |
| Wartungsplan | Kernblatt: Maßnahmen, Bewertung, RPZ, Ampel, Intervall, Planungsdaten | Instandhaltung, Produktion |
| Durchführungsnachweis | letzte Durchführung, Fälligkeit, Status, Befund | Instandhaltung, QM/Audit |
| Anlagenübersicht | Zeichnung + Positionsliste mit Abdeckungsprüfung | alle |
| Bewertungskatalog | alle Skalen auf einem Blatt + Bewertungsregeln | Bewertungsteam |
| RPZ-Methodik | Formeln, Wertebereiche, Klassen, Live-Rechenbeispiel | Bewertungsteam, QM |
| Parameter | Grenzwerte, Intervalle, SF, Auswahllisten | IH-Leitung |
| Offene Punkte | Befunde der Überarbeitung mit Maßnahmenverfolgung | Bewertungsteam |
| Weiterentwicklung | zusätzliche Kriterien und KPIs | Management, QM |
| Änderungsprotokoll | alle Änderungen mit Begründung | QM / Dokumentenlenkung |

---

## Teil B – Umgesetzte Änderungen und Begründung

### Was garantiert unverändert ist
- **Alle 87 Bewertungen** (F, IH, Q, AS, US, S, A, E, SF) sind übernommen; Verknüpfungen wie `=F44` bleiben Verknüpfungen (jetzt auf die neue Zeile).
- **RPZ und Intervall jeder Maßnahme sind identisch** – maschinell gegen die Vorversion geprüft (88 Zeilen, 0 Abweichungen; einzige Differenz: die neu angelegte Maßnahme 17.1 ohne Bewertung).
- **Jeder Originaltext** steht in der Spalte „Originaltext Vorversion“, jede alte Nummer in „Nr. Vorversion“ (aufklappbare Spaltengruppe rechts).
- Alle Skalentexte, Kommentare (Leitfragen, Schichtfaktor), Grenzwerte, Bilder, das Logo und die Fußzeilenangaben sind übernommen.

### Wesentliche Änderungen

| Bereich | Änderung | Warum besser |
|---|---|---|
| Methodik | Klasse **„kritisch“ (RPZ > 125 → 4 Wochen + Risikominderung)** neu; RPZ < 15 immer „gering“ | Jede vollständig bewertete Maßnahme bekommt ein Intervall. Aktuell keine Maßnahme betroffen. |
| Methodik | Grenzen 15 / 66 / 125 und Intervalle im Blatt **Parameter** | Eine Stelle zum Ändern, Formeln lesbar. |
| Methodik | Doku vereinheitlicht: RPZ = B × A × E × SF; Grenzen wie in der Formel | Keine Widersprüche mehr. |
| Methodik | Schichtfaktor zentral über Schichtmodell (Deckblatt), zeilenweise überschreibbar | Konsistent; Wechsel auf 2-Schicht verschiebt alle Intervalle automatisch. |
| Plan | **Risikoklasse mit Ampel** (gering/mittel/hoch/kritisch/unvollständig) | Für Produktion und Management sofort lesbar. |
| Plan | **Intervall festgelegt + Begründung** (Standard = berechnet; blau markiert bei Abweichung) | Herstellervorgaben, Gesetz und Erfahrung können das RPZ-Ergebnis begründet übersteuern – auditfest. |
| Plan | **Relevanz AS/Q/US** (≥ 4) | Sicherheits-/qualitätskritische Maßnahmen bleiben auch bei kleiner RPZ sichtbar. |
| Plan | **Instandsetzung bei Ausfall** (aus IH abgeleitet), **Ausführung durch**, **Plan-Zeit**, **Aufwand h/Jahr** | Zuständigkeit, Stillstandsrisiko und Kapazitätsbedarf werden planbar. |
| Plan | **Art der Maßnahme** (Inspektion, Funktionsprüfung, Messung, Schmierung, Reinigung, Wartung, Austausch) | Filter nach Tätigkeit, z. B. Schmierrundgang. |
| Plan | Warn-/Info-/Elektro-Symbole → Textspalte **Hinweis** | Filterbar, druckt zuverlässig. |
| Plan | **Sollzustände ergänzt** (grau kursiv = Vorschlag, freizugeben) | Objektive i.O./n.i.O.-Entscheidung. |
| Plan | Einheitliche Nummerierung, Positionsnummern in eigener Spalte, Formulierungen „Objekt + Tätigkeit“, Rechtschreibung | Professionell, eindeutig referenzierbar. |
| Plan | Excel-Tabelle mit Filter, fixierten Spalten, **Eingabeprüfung 1–5 mit Skalen-Hilfe**, Druck A3 quer | Weniger Fehler, Bewertung ohne Blattwechsel. |
| Plan | Bilder je Maßnahme in Spalte „Objekt / Bild“ | Bild steht fest bei der richtigen Maßnahme. |
| Neu | **Durchführungsnachweis** mit Fälligkeit, Status-Ampel (überfällig / bald fällig / in Ordnung), Befund | Nachweis nach IATF 16949 Kap. 8.5.1.5; Basis für PM-Compliance. |
| Neu | **Dashboard** | Management-Übersicht ohne Detailstudium. |
| Neu | **Offene Punkte** (13 Befunde mit Priorität, Verantwortlich, Termin, Status) | Lücken werden systematisch geschlossen statt vergessen. |
| Struktur | 8 Skalen-Blätter → 1 **Bewertungskatalog** (Matrix Kriterium × Stufe) inkl. Bewertungsregeln; Stufen zusätzlich als „sehr gering … sehr hoch“ | Vergleichbar, kompakt, keine Verwechslung mehr. |
| Struktur | Summenzeile → Dashboard-Kennzahlen | Nur aussagekräftige Kennzahlen (Ø RPZ korrekt: 28,5 statt 28,9). |
| Technik | `.xlsx` ohne Makro, tote externe Namen entfernt | Keine Sicherheits- und Verknüpfungswarnungen; Tabelle ersetzt das Makro. |
| Lenkung | Deckblatt mit Stammdaten, Freigabe (IH, Produktion, QM, AS/Umwelt), nächster Überprüfung, Historie | Gelenktes Dokument (IATF 16949 / ISO 9001 Kap. 7.5). |

Bei der Übernahme der Fotos gingen die **Pfeil-/Linienmarkierungen** verloren, die in der Vorversion auf vier Fotos (4.1.1, 6.1.2, 9.4.5, 21.3) gezeichnet waren. Ebenso entfiel eine Ovalmarkierung bei 4.1.4. Die Fotos selbst sind vollständig übernommen.

---

## Teil C – Vorschläge für zusätzliche Kriterien und Kennzahlen

**Zusätzliche Kriterien** (im Blatt „Weiterentwicklung“ mit Status):
1. Maschinenzustand bei Durchführung (Betrieb / Stillstand / LOTO).
2. Gesetzliche/normative Prüfgrundlage mit Frist (BetrSichV, DGUV V3, DGUV Regel 113-020, Druckgeräte).
3. Ersatzteil/Material (SAP-Nr.) und Wiederbeschaffungszeit als Risikofaktor.
4. Messmittel und Grenzwerte im Sollzustand (Drücke, Spiel, Temperaturen).
5. Verknüpfung zu PFMEA / Control Plan für Maßnahmen mit Q ≥ 4 (besondere Merkmale).
6. Zustandsorientierte Instandhaltung für A ≥ 4 und E ≥ 4 (Kandidaten zeigt die Risikomatrix im Dashboard).
7. Validierung von A aus der Störhistorie (SAP-PM, MTBF).

**Kennzahlen:** umgesetzt sind PM-Compliance, Überfällige, Anteil hoch/kritisch, Ø RPZ je Baugruppe, n.i.O.-Quote, Soll-Durchführungen und Aufwand pro Jahr. Vorgeschlagen werden OEE, MTBF, MTTR und der Anteil geplanter Instandhaltung (Datenquelle MES/SAP-PM). Diese Kennzahlen nennt IATF 16949 Kap. 8.5.1.5 ausdrücklich.

---

## Teil D – Nächste Schritte

1. Stammdaten auf dem Deckblatt ergänzen.
2. Die **Offenen Punkte** im Team abarbeiten, mit Priorität „hoch“ zuerst: A/E für 4.1.4, Bewertung 17.1, tägliche Not-Aus-Prüfung, gesetzliche Prüfungen.
3. Vorgeschlagene Sollzustände prüfen und mit Messwerten konkretisieren; danach die graue Kursivschrift auf normal umstellen.
4. „Ausführung durch“ und „Plan-Zeit“ pflegen, damit Aufwand und Kapazität aussagekräftig werden.
5. Freigabe durch Instandhaltung, Produktion, QM und Arbeitssicherheit/Umwelt, danach Stand 02 als gültig kennzeichnen.

*Technischer Hinweis:* Die Auswertung „max. RPZ je Baugruppe“ nutzt MAXIFS (Excel 2019 / Microsoft 365). Alle 2.225 Formeln wurden ohne Fehler neu berechnet.
