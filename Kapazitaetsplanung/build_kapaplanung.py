import sys, datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter as L

SRC, OUT = sys.argv[1], sys.argv[2]
wb = openpyxl.load_workbook(SRC)
orig = wb.active
orig.title = "Original 21.07.2025"

# ---------- Stil ----------
F = "Arial"
NAVY, BLUE_L, GREY_L, INPUT, SUGG = "1F3864", "D9E2F3", "F2F2F2", "FFF2CC", "EADCF8"
thin = Side(style="thin", color="BFBFBF")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
def font(**k):
    k.setdefault("name", F); k.setdefault("size", 10); return Font(**k)
def fill(c): return PatternFill("solid", fgColor=c)
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

def title(ws, text, sub=None, width_to="H"):
    ws["A1"] = text; ws["A1"].font = font(size=16, b=True, color=NAVY)
    ws.row_dimensions[1].height = 26
    if sub:
        ws["A2"] = sub; ws["A2"].font = font(size=9, i=True, color="595959")
    ws.sheet_view.showGridLines = False

def header(ws, row, labels, col=1, h=32):
    for i, lab in enumerate(labels):
        c = ws.cell(row, col + i, lab)
        c.font = font(b=True, color="FFFFFF"); c.fill = fill(NAVY)
        c.alignment = CENTER; c.border = BOX
    ws.row_dimensions[row].height = h

def section(ws, row, text, ncols=8, col=1):
    c = ws.cell(row, col, text); c.font = font(size=12, b=True, color=NAVY)
    for i in range(ncols):
        ws.cell(row, col + i).border = Border(bottom=Side(style="medium", color=NAVY))

def cell(ws, ref, v, fmt=None, inp=False, bold=False, sugg=False, align=None, color=None):
    c = ws[ref]; c.value = v
    c.font = font(b=bold, color=("0000FF" if inp else color or ("7030A0" if sugg else "000000")), i=sugg)
    if inp: c.fill = fill(INPUT)
    if sugg: c.fill = fill(SUGG)
    if fmt: c.number_format = fmt
    c.border = BOX
    c.alignment = align or Alignment(vertical="center", wrap_text=True)
    return c

def name(n, ref): wb.defined_names[n] = DefinedName(n, attr_text=ref)

# =====================================================================
# 1) Parameter
# =====================================================================
P = wb.create_sheet("Parameter")
title(P, "Parameter & Annahmen", "Alle Berechnungsgrundlagen an einer Stelle. Nur blaue Werte auf gelbem Grund ändern – alle Blätter rechnen automatisch neu.")
header(P, 4, ["Parameter", "Wert", "Einheit", "Quelle / Herleitung"])
params = [
    # name, label, value, fmt, unit, source, input
    ("P_Stand", "Planungsstand", datetime.date(2025, 7, 21), "DD.MM.YYYY", "Datum", "Original, Zelle F1", True),
    ("P_AT", "Arbeitstage pro Jahr", 250, "0", "AT/Jahr", "Original, Formel Std./Jahr = Std./Tag × 250", True),
    ("P_TageWoche", "Arbeitstage pro Woche", 5, "0", "AT/Woche", "Original, Formel Std./Woche = Std./Tag × 5", True),
    ("P_JahresStd", "Jahresarbeitsstunden je Mitarbeitende/r (1 FTE)", 1750, "#,##0", "Std./Jahr", "Original, Spaltenkopf „1 MA = 1750“", True),
    ("P_StdTag", "Stunden je FTE und Arbeitstag (abgeleitet)", "=P_JahresStd/P_AT", "0.00", "Std./AT", "Berechnet: 1.750 / 250 = 7 (im Original fest als „/7“ in D24 hinterlegt)", False),
    ("P_Urlaub", "Urlaubstage je Mitarbeitende/r", 38, "0", "Tage/Jahr", "Original, Bemerkung „38 Tage Urlaub / TZ / 220 AT“", True),
    ("P_BezugAT", "Bezugsarbeitstage für Urlaubsquote", 220, "0", "AT", "Original, Formel Urlaub = 38/220", True),
    ("P_Krank", "Krankenquote", 0.041, "0.0%", "%", "Original, „Statistik fz-IH“ (4,1 %)", True),
    ("P_WB", "Weiterbildungstage je Mitarbeitende/r (NEU)", 0, "0", "Tage/Jahr", "NEU – im Original nicht enthalten; 0 = keine Auswirkung auf das Ergebnis. Bei Bedarf eintragen.", True),
    ("P_Sonst", "Sonstige Ausfallzeiten (NEU)", 0, "0.0%", "%", "NEU – z. B. Betriebsversammlung, Arztbesuche, Gremien; 0 = keine Auswirkung.", True),
    ("P_RiskHoch", "Risikozahl ab der „hoch“ gilt", 15, "0", "Punkte", "Bewertungsmaßstab NEU: Auswirkung (1–5) × Wahrscheinlichkeit (1–5)", True),
    ("P_RiskMittel", "Risikozahl ab der „mittel“ gilt", 8, "0", "Punkte", "Bewertungsmaßstab NEU", True),
    ("P_DeckGruen", "Deckungsgrad Ampel grün ab", 1.0, "0%", "%", "Bewertungsmaßstab NEU: verfügbare FTE ÷ Brutto-Bedarf", True),
    ("P_DeckGelb", "Deckungsgrad Ampel gelb ab", 0.9, "0%", "%", "Bewertungsmaßstab NEU", True),
]
for i, (n, lab, v, fmt, unit, src, inp) in enumerate(params):
    r = 5 + i
    cell(P, f"A{r}", lab)
    cell(P, f"B{r}", v, fmt, inp=inp, align=Alignment(horizontal="right", vertical="center"))
    cell(P, f"C{r}", unit); cell(P, f"D{r}", src)
    name(n, f"Parameter!$B${r}")
pr = 5 + len(params) + 1
section(P, pr, "Bewertungsskalen (einheitlich für alle Blätter)", 4)
scales = [
    ("Auswirkung bei Nichterfüllung", "1 = gering (kein Einfluss auf Produktion/Qualität) · 2 = spürbar · 3 = deutlich (Verzögerungen, Mehraufwand) · 4 = hoch (Anlagenstillstand / Qualitätsrisiko) · 5 = kritisch (Sicherheit, Recht, Kunde)"),
    ("Eintrittswahrscheinlichkeit Engpass", "1 = sehr unwahrscheinlich · 2 = selten · 3 = gelegentlich · 4 = häufig · 5 = nahezu sicher"),
    ("Risikozahl / -klasse", "Auswirkung × Wahrscheinlichkeit (1–25). „hoch“ ≥ Schwelle hoch, „mittel“ ≥ Schwelle mittel, sonst „niedrig“."),
    ("Datenbasis", "Messung (Zeiterfassung/CMMS) · Erfahrungswert · Schätzung"),
    ("Bewertungsstatus (automatisch)", "bewertet = Stunden > 0 · „0 h – Plausibilität prüfen“ = bewusst 0 eingetragen · „offen – zu bewerten“ = noch kein Wert"),
]
for i, (a, b) in enumerate(scales):
    r = pr + 1 + i
    cell(P, f"A{r}", a, bold=True); P.merge_cells(f"B{r}:D{r}"); cell(P, f"B{r}", b)
    P.row_dimensions[r].height = 30
for col, w in zip("ABCD", (46, 14, 12, 80)): P.column_dimensions[col].width = w
P.freeze_panes = "A5"

# =====================================================================
# 2) Kapazitätsbedarf
# =====================================================================
K = wb.create_sheet("Kapazitätsbedarf")
title(K, "Kapazitätsbedarf BB2 Instandhaltung – Umsetzung",
      "Eingabe: Std. je Arbeitstag (Spalte G). Wochen-, Jahres- und FTE-Werte werden aus den Parametern berechnet. Lila = Erstvorschlag zur Validierung im Team.")
K["L2"] = "Stand:"; K["L2"].font = font(b=True); K["L2"].alignment = Alignment(horizontal="right")
K["M2"] = "=P_Stand"; K["M2"].number_format = "DD.MM.YYYY"; K["M2"].font = font(b=True, color="008000")
heads = ["Nr.", "Herkunft", "Tätigkeitsbereich", "Instandhaltungs-\nart (DIN 31051)", "Tätigkeit", "Bemerkung (Original)",
         "Std. /\nArbeitstag", "Std. /\nWoche", "Std. /\nJahr", "FTE\nnetto", "Anteil am\nGesamtbedarf", "Bewertungs-\nstatus",
         "Verantwortlich", "Schicht-\nübergreifend", "Datenbasis", "Auswirkung\n(1–5)", "Wahrschein-\nlichkeit (1–5)", "Risiko-\nzahl",
         "Risiko-\nklasse", "Ist-Std. / Jahr\n(Zeiterfassung)", "Abweichung\nIst vs. Plan", "Klärungsbedarf / Hinweis", "Kennzeichnung\nim Original", "Hilfe\nPareto", "Hilfe\nRisiko"]
HR = 5
K.cell(4, 7, "Kapazität (Plan)").font = font(b=True, color=NAVY)
K.cell(4, 13, "Organisation").font = font(b=True, color=NAVY)
K.cell(4, 16, "Risikobewertung (Vorschlag)").font = font(b=True, color=NAVY)
K.cell(4, 20, "Plan-Ist-Vergleich").font = font(b=True, color=NAVY)
header(K, HR, heads, h=44)

ET, IS, PV, VP, NA, QT, BB, FS, OA = ("Ersatzteilmanagement", "Instandsetzung & Reparatur", "Präventive Instandhaltung",
    "Verbesserung, Projekte & Technik", "Neuanlagen, Werkzeuge & Umzüge", "Qualifizierung & TPM",
    "Bereichs- & Standortbetreuung", "Facility Management & Sonderaufgaben", "Organisation & Administration")
CATS = [ET, IS, PV, VP, NA, QT, BB, FS, OA]
W_, I_, R_, V_, U_ = "Wartung", "Inspektion", "Instandsetzung", "Verbesserung", "Unterstützende Tätigkeit"
DIN = [W_, I_, R_, V_, U_]
# (Herkunft, Bereich, DIN, Tätigkeit, Bemerkung, Std/AT, Ausw, Wahrsch, Hinweis, Kennzeichnung)
O = "Original"; N = "Neu (Vorschlag)"
rows = [
 (O, ET, U_, "Ersatzteilmanagement – Planung", None, 1.5, 4, 3, None, "–"),
 (O, ET, U_, "Ersatzteilumlagerung und Sonstiges", "Lagerhaltung / Ersatzteilbeschaffung extern", 2, 3, 2, "Farbmarkierung im Original nicht erläutert – Bedeutung klären.", "grün hinterlegt, rote fette Schrift"),
 (O, IS, R_, "Komplexe Reparaturen / Generalüberholungen", None, 0.3, 4, 3, None, "–"),
 (O, IS, R_, "Reparatur / Überholung von Ersatzteilen", "Komplexe Ersatzteile (Druckübersetzer, Wasserventile …)", 0, 4, 3, "0 h: Wird die Aufbereitung extern vergeben oder fehlt Kapazität? Kosten-/Zeitvergleich intern vs. extern.", "rote Schrift"),
 (O, VP, V_, "Projekte und KVP (Lean-Projekte, z. B. SMED)", "Entwicklung, Berechnung, Bestellung", 10, 3, 3, "Größter Einzelposten (≈36 %). Empfehlung: Projektliste mit Stundenbudget hinterlegen.", "–"),
 (O, QT, U_, "Training / Schulung / Betreuung / TPM / WPS", None, 0.5, 3, 3, None, "–"),
 (O, BB, U_, "Auslandseinsätze / Auslandsbetreuung", None, 0.5, 2, 2, None, "–"),
 (O, NA, U_, "Anfahren von Neuanlagen", None, 0.5, 4, 3, None, "–"),
 (O, NA, U_, "Freiprüfung neuer Werkzeuge (Versuche)", None, 0.2, 4, 2, None, "–"),
 (O, VP, V_, "Konstruktion / 3D", None, 1, 2, 2, None, "–"),
 (O, NA, U_, "Maschinenumzüge", None, 0.1, 3, 2, None, "–"),
 (O, VP, R_, "CNC-Technik (z. B. Nachdrehen Messer)", None, 0, 2, 2, "0 h: Leistung extern bezogen oder nicht benötigt? Bitte bestätigen.", "–"),
 (O, FS, U_, "Sonderaufgaben (Lagersysteme …)", None, 0.5, 2, 2, None, "grün hinterlegt"),
 (O, FS, U_, "Facility Management", None, 6, 3, 2, "Zweitgrößter Posten (≈21 %). Abgrenzung zu Gebäudetechnik/Hausmeisterei prüfen.", "grün hinterlegt"),
 (O, OA, U_, "Planung und Dokumentation", None, 0, 4, 4, "0 h: Wartungsplanung, CMMS-Pflege und Dokumentation sind Grundlage für IATF 16949 (TPM) – Aufwand bewerten.", "orange hinterlegt, rote fette Schrift"),
 (O, BB, U_, "Betreuung Balg-Bereich", None, 2, 3, 2, "Farbmarkierung im Original nicht erläutert – Bedeutung klären.", "grün hinterlegt"),
 (O, PV, W_, "Wartung Betriebsmittel", None, 0, 5, 4, "0 h: Klären, ob die Wartung durch die Schicht-Instandhaltung abgedeckt ist. Falls nicht: höchstes Risiko.", "orange hinterlegt, rote fette Schrift"),
 (O, PV, W_, "Wartung allgemeiner Maschinen", None, 0, 5, 4, "0 h: wie Wartung Betriebsmittel – Zuständigkeit klären.", "rote fette Schrift"),
 (O, OA, U_, "Administration", None, 3, 2, 2, None, "grün hinterlegt, fette Schrift"),
 (N, IS, R_, "Störungsbehebung / ungeplante Instandsetzung", None, None, 5, 3, "Fehlt im Original. Nur relevant, falls das Umsetzungsteam Störungen mitbearbeitet (Abgrenzung zur Schicht-IH).", "–"),
 (N, PV, I_, "Inspektion & gesetzliche Prüfungen (DGUV V3, UVV, Krane, Druckbehälter)", None, None, 5, 3, "Fehlt im Original. Prüffristen sind rechtlich bindend.", "–"),
 (N, OA, U_, "Arbeitssicherheit (Gefährdungsbeurteilung, LOTO, Unterweisungen)", None, None, 5, 2, "Fehlt im Original.", "–"),
 (N, VP, V_, "Kennzahlen & Schwachstellenanalyse (MTBF, MTTR, Top-Störer)", None, None, 3, 3, "Fehlt im Original. Grundlage für gezielte KVP-Projekte.", "–"),
 (N, OA, U_, "Steuerung von Fremdfirmen / Dienstleistern", None, None, 3, 2, "Fehlt im Original (Angebote, Einweisung, Abnahme).", "–"),
 (N, QT, U_, "Audits & QM-Unterstützung (IATF 16949, ISO 14001/50001, 8D)", None, None, 4, 2, "Fehlt im Original.", "–"),
 (N, VP, V_, "Energiemanagement (z. B. Druckluft-Leckagen)", None, None, 2, 2, "Fehlt im Original.", "–"),
 (N, QT, U_, "Einarbeitung neuer Mitarbeitender / Wissenssicherung", None, None, 4, 3, "Fehlt im Original. Reduziert Abhängigkeit von Einzelpersonen.", "–"),
]
R0 = HR + 1
RN = R0 + len(rows) - 1
TOT = RN + 1
rng = lambda c: f"${c}${R0}:${c}${RN}"
for i, (her, cat, din, act, rem, h, aus, wa, hint, mark) in enumerate(rows):
    r = R0 + i
    new = her == N
    cell(K, f"A{r}", i + 1, "0", align=CENTER)
    cell(K, f"B{r}", her, sugg=new)
    cell(K, f"C{r}", cat); cell(K, f"D{r}", din); cell(K, f"E{r}", act, bold=False)
    cell(K, f"F{r}", rem)
    cell(K, f"G{r}", h, "0.0", inp=True, align=CENTER)
    cell(K, f"H{r}", f'=IF(G{r}="","",G{r}*P_TageWoche)', "0.0", align=CENTER)
    cell(K, f"I{r}", f'=IF(G{r}="","",G{r}*P_AT)', "#,##0", align=CENTER)
    cell(K, f"J{r}", f'=IF(G{r}="","",I{r}/P_JahresStd)', "0.00", align=CENTER)
    cell(K, f"K{r}", f'=IF(G{r}="","",I{r}/$I${TOT})', "0.0%", align=CENTER)
    cell(K, f"L{r}", f'=IF(G{r}="","offen – zu bewerten",IF(G{r}=0,"0 h – Plausibilität prüfen","bewertet"))', align=CENTER)
    cell(K, f"M{r}", None, inp=True); cell(K, f"N{r}", None, inp=True, align=CENTER); cell(K, f"O{r}", None, inp=True)
    cell(K, f"P{r}", aus, "0", sugg=True, align=CENTER); cell(K, f"Q{r}", wa, "0", sugg=True, align=CENTER)
    cell(K, f"R{r}", f'=IF(OR(P{r}="",Q{r}=""),"",P{r}*Q{r})', "0", align=CENTER)
    cell(K, f"S{r}", f'=IF(R{r}="","",IF(R{r}>=P_RiskHoch,"hoch",IF(R{r}>=P_RiskMittel,"mittel","niedrig")))', align=CENTER)
    cell(K, f"T{r}", None, "#,##0", inp=True, align=CENTER)
    cell(K, f"U{r}", f'=IF(OR(T{r}="",N(I{r})=0),"",T{r}/I{r}-1)', "+0%;-0%;0%", align=CENTER)
    cell(K, f"V{r}", hint); cell(K, f"W{r}", mark)
    cell(K, f"X{r}", f"=N(I{r})+(1000-ROW())/1000000", "0.000000")
    cell(K, f"Y{r}", f"=N(R{r})+(1000-ROW())/1000000", "0.000000")
    if new:
        for c in "ACDEF": K[f"{c}{r}"].fill = fill(GREY_L)
    K.row_dimensions[r].height = 30
# Summenzeile
for c in "ABCDEFGHIJKLMNOPQRSTUVW":
    x = K[f"{c}{TOT}"]; x.font = font(b=True); x.fill = fill(BLUE_L); x.border = BOX
    x.alignment = CENTER
K[f"E{TOT}"] = "Gesamt (netto)"; K[f"E{TOT}"].alignment = Alignment(vertical="center")
for c, fmt in (("G", "0.0"), ("H", "0.0"), ("I", "#,##0"), ("J", "0.00"), ("K", "0.0%"), ("T", "#,##0")):
    K[f"{c}{TOT}"] = f"=SUM({c}{R0}:{c}{RN})"; K[f"{c}{TOT}"].number_format = fmt
K[f"T{TOT}"] = f'=IF(COUNT(T{R0}:T{RN})=0,"",SUM(T{R0}:T{RN}))'
K[f"U{TOT}"] = f'=IF(T{TOT}="","",T{TOT}/I{TOT}-1)'; K[f"U{TOT}"].number_format = "+0%;-0%;0%"
K.row_dimensions[TOT].height = 22
# Kontrollzeile (Originalmethode D24 = B23/7)
c1 = TOT + 2
K[f"E{c1}"] = "Kontrolle: FTE nach Originalmethode (Std./Tag ÷ Std. je FTE und Tag)"
K[f"J{c1}"] = f"=G{TOT}/P_StdTag"; K[f"J{c1}"].number_format = "0.00"
K[f"K{c1}"] = f'=IF(ABS(J{c1}-J{TOT})<0.0001,"✔ identisch","✘ Abweichung")'
K[f"E{c1+1}"] = "Kontrolle: Abgleich mit Originalwert 21.07.2025 (4,01 FTE netto, 28,1 Std./Tag)"
K[f"J{c1+1}"] = 4.014285714285714; K[f"J{c1+1}"].number_format = "0.00"; K[f"J{c1+1}"].font = font(color="0000FF")
K[f"J{c1+1}"].comment = Comment("Hart hinterlegter Referenzwert aus dem Original (Zelle E24), nur zur Kontrolle.", "Überarbeitung")
K[f"K{c1+1}"] = f'=IF(ABS(J{c1+1}-J{TOT})<0.0001,"✔ unverändert","Δ "&TEXT(J{TOT}-J{c1+1},"+0.00;-0.00"))'
for r in (c1, c1 + 1):
    K[f"E{r}"].font = font(i=True, color="595959"); K[f"J{r}"].font = font(b=True, color=K[f"J{r}"].font.color)
    K[f"K{r}"].font = font(b=True, color="008000")
K[f"E{c1+3}"] = "Legende: blau auf gelb = Eingabe · schwarz = Formel · lila = Erstvorschlag (zu validieren) · grau hinterlegt = neu ergänzte Tätigkeit (0 Std. bis zur Bewertung → ändert das Ergebnis nicht)"
K[f"E{c1+3}"].font = font(size=9, i=True, color="595959")
widths = dict(A=5, B=11, C=24, D=16, E=44, F=30, G=9, H=9, I=9, J=8, K=10, L=17, M=15, N=11, O=12, P=10, Q=11, R=8, S=9, T=13, U=11, V=50, W=20, X=8, Y=8)
for c, w in widths.items(): K.column_dimensions[c].width = w
K.column_dimensions["X"].hidden = True; K.column_dimensions["Y"].hidden = True
K.freeze_panes = f"F{R0}"
K.auto_filter.ref = f"A{HR}:W{RN}"
# Datenüberprüfung
def dv(ws, formula, ref, prompt=None):
    d = DataValidation(type="list", formula1=formula, allow_blank=True)
    if prompt: d.promptTitle, d.prompt, d.showInputMessage = "Hinweis", prompt, True
    ws.add_data_validation(d); d.add(ref)
dv(K, '"Ja,Nein"', f"N{R0}:N{RN}")
dv(K, '"Messung,Erfahrungswert,Schätzung"', f"O{R0}:O{RN}")
dv(K, '"1,2,3,4,5"', f"P{R0}:Q{RN}", "Skala siehe Blatt Parameter")
dv(K, "=Lst_Bereich", f"C{R0}:C{RN}"); dv(K, "=Lst_DIN", f"D{R0}:D{RN}")
dv(K, "=Lst_Team", f"M{R0}:M{RN}")
dv(K, '"Original,Neu (Vorschlag)"', f"B{R0}:B{RN}")
num = DataValidation(type="decimal", operator="between", formula1="0", formula2="24", allow_blank=True,
                     errorTitle="Ungültig", error="Bitte Stunden je Arbeitstag zwischen 0 und 24 eingeben.")
K.add_data_validation(num); num.add(f"G{R0}:G{RN}")
# bedingte Formatierung
red, amber, green = fill("F8CBAD"), fill("FFE699"), fill("C6EFCE")
K.conditional_formatting.add(f"S{R0}:S{RN}", CellIsRule(operator="equal", formula=['"hoch"'], fill=red))
K.conditional_formatting.add(f"S{R0}:S{RN}", CellIsRule(operator="equal", formula=['"mittel"'], fill=amber))
K.conditional_formatting.add(f"S{R0}:S{RN}", CellIsRule(operator="equal", formula=['"niedrig"'], fill=green))
K.conditional_formatting.add(f"L{R0}:L{RN}", FormulaRule(formula=[f'LEFT(L{R0},1)="0"'], fill=amber))
K.conditional_formatting.add(f"L{R0}:L{RN}", FormulaRule(formula=[f'LEFT(L{R0},1)="o"'], fill=fill("E7E6E6")))
from openpyxl.formatting.rule import DataBarRule
K.conditional_formatting.add(f"I{R0}:I{RN}", DataBarRule(start_type="num", start_value=0, end_type="max", color="5B9BD5"))

# =====================================================================
# 3) Personalbedarf
# =====================================================================
PB = wb.create_sheet("Personalbedarf")
title(PB, "Personalbedarf & Kapazitätsabgleich", "Netto-Bedarf aus Blatt Kapazitätsbedarf + Verteilzeiten = Brutto-Bedarf; Abgleich mit verfügbarem Team.")
section(PB, 4, "A) Netto-Bedarf", 6)
header(PB, 5, ["Kennzahl", "Wert", "Einheit", "Wert Original\n21.07.2025", "Abgleich", "Bemerkung"])
cell(PB, "A6", "Summe Stunden je Arbeitstag"); cell(PB, "B6", f"=Kapazitätsbedarf!G{TOT}", "0.0", color="008000")
cell(PB, "C6", "Std./AT"); cell(PB, "D6", 28.1, "0.0", color="0000FF"); cell(PB, "E6", '=IF(ABS(B6-D6)<0.0001,"✔","Δ")', align=CENTER); cell(PB, "F6", "Original B23 „Gesamt“")
cell(PB, "A7", "Summe Stunden je Jahr"); cell(PB, "B7", f"=Kapazitätsbedarf!I{TOT}", "#,##0", color="008000"); cell(PB, "C7", "Std./Jahr")
cell(PB, "D7", None); cell(PB, "E7", None); cell(PB, "F7", "Im Original nicht summiert – ergänzt")
cell(PB, "A8", "Personalbedarf netto", bold=True); cell(PB, "B8", f"=Kapazitätsbedarf!J{TOT}", "0.00", bold=True, color="008000")
cell(PB, "C8", "FTE"); cell(PB, "D8", 4.014285714285714, "0.00", color="0000FF"); cell(PB, "E8", '=IF(ABS(B8-D8)<0.0001,"✔","Δ")', align=CENTER); cell(PB, "F8", "Original D24/E24 „Summe Direkte Mitarbeiter – netto“")

section(PB, 10, "B) Verteilzeiten (Ausfallzeiten) – Methode wie Original: Zuschlag auf Netto-Bedarf", 6)
header(PB, 11, ["Verteilzeit", "Anteil", "FTE", "Wert Original\n21.07.2025", "Abgleich", "Herleitung / Quelle"])
vt = [
    ("Urlaub", "=P_Urlaub/P_BezugAT", 0.6933766233766234, "38 Urlaubstage / 220 Arbeitstage (Original)"),
    ("Krankheit", "=P_Krank", 0.16458571428571428, "Statistik fz-IH (Original)"),
    ("Weiterbildung (NEU)", "=P_WB/P_BezugAT", None, "NEU – Parameter, derzeit 0"),
    ("Sonstige Ausfallzeiten (NEU)", "=P_Sonst", None, "NEU – Parameter, derzeit 0"),
]
for i, (lab, f, ov, src) in enumerate(vt):
    r = 12 + i
    cell(PB, f"A{r}", lab); cell(PB, f"B{r}", f, "0.0%"); cell(PB, f"C{r}", f"=$B$8*B{r}", "0.00")
    cell(PB, f"D{r}", ov, "0.00", color="0000FF")
    cell(PB, f"E{r}", f'=IF(D{r}="","",IF(ABS(C{r}-D{r})<0.0001,"✔","Δ"))', align=CENTER); cell(PB, f"F{r}", src)
cell(PB, "A16", "Summe Verteilzeiten", bold=True); cell(PB, "B16", "=SUM(B12:B15)", "0.0%", bold=True); cell(PB, "C16", "=SUM(C12:C15)", "0.00", bold=True)
cell(PB, "D16", 0.8579623376623378, "0.00", color="0000FF"); cell(PB, "E16", '=IF(ABS(C16-D16)<0.0001,"✔","Δ")', align=CENTER); cell(PB, "F16", "Original D28")
cell(PB, "A17", "Personalbedarf brutto", bold=True); cell(PB, "B17", None); cell(PB, "C17", "=B8+C16", "0.00", bold=True)
cell(PB, "D17", 4.872248051948052, "0.00", color="0000FF"); cell(PB, "E17", '=IF(ABS(C17-D17)<0.0001,"✔","Δ")', align=CENTER); cell(PB, "F17", "Original D29 „brutto“ (im Original ebenfalls als „Summe Direkte Mitarbeiter“ bezeichnet)")
for c in "ABCDEF": PB[f"{c}17"].fill = fill(BLUE_L)
PB["A18"] = "„Wert Original“: hart hinterlegte Referenzwerte aus der Datei vom 21.07.2025 – belegen, dass die Bewertung unverändert übernommen wurde."
PB["A18"].font = font(size=9, i=True, color="595959")

section(PB, 20, "C) Methodenvergleich (Sensitivität)", 6)
header(PB, 21, ["Methode", "Formel", "FTE brutto", "Differenz zu\nOriginal", "", "Einordnung"])
cell(PB, "A22", "1 – Zuschlagsmethode (Original)"); cell(PB, "B22", "netto × (1 + Σ Verteilzeit)")
cell(PB, "C22", "=C17", "0.00"); cell(PB, "D22", 0, "+0.00;-0.00;0.00"); cell(PB, "E22", None)
cell(PB, "F22", "Korrekt, wenn die 220 AT die tatsächlichen Anwesenheitstage sind.")
cell(PB, "A23", "2 – Verfügbarkeitsmethode"); cell(PB, "B23", "netto ÷ (1 − Σ Verteilzeit)")
cell(PB, "C23", "=B8/(1-B16)", "0.00"); cell(PB, "D23", "=C23-C22", "+0.00;-0.00;0.00"); cell(PB, "E23", None)
cell(PB, "F23", "Üblich, wenn Quoten auf die Soll-Arbeitszeit bezogen sind (z. B. Krankenquote). Konservativer Wert.")
for r in (22, 23): PB.row_dimensions[r].height = 30
PB["A24"] = "Empfehlung: Bezugsbasis der 220 AT klären und eine Methode verbindlich festlegen. Bis dahin gilt Methode 1 (Original)."
PB["A24"].font = font(size=9, i=True, color="595959")

section(PB, 26, "D) Kapazitätsabgleich – verfügbares Team", 6)
header(PB, 27, ["Nr. (Original)", "Name", "Verfügbare FTE", "Rolle / Funktion", "Kennzeichnung\nim Original", "Bemerkung"])
team = [(1, "Scheifele", "grün hinterlegt"), (2, "Baran", "grün hinterlegt"), (4, "Vischer", "grün hinterlegt"), (5, "Mrugalla", "blau hinterlegt")]
for i, (nr, nm, mk) in enumerate(team):
    r = 28 + i
    cell(PB, f"A{r}", nr, "0", align=CENTER); cell(PB, f"B{r}", nm)
    cell(PB, f"C{r}", 1.0, "0.00", inp=True, align=CENTER); cell(PB, f"D{r}", None, inp=True)
    cell(PB, f"E{r}", mk); cell(PB, f"F{r}", "Annahme 1,0 FTE (Vollzeit, voll im Umsetzungsteam) – bitte prüfen" if i == 0 else "Annahme 1,0 FTE – bitte prüfen")
for i in range(2):  # Reserve
    r = 32 + i
    cell(PB, f"A{r}", None, inp=True); cell(PB, f"B{r}", None, inp=True); cell(PB, f"C{r}", None, "0.00", inp=True)
    cell(PB, f"D{r}", None, inp=True); cell(PB, f"E{r}", None); cell(PB, f"F{r}", "freie Zeile (z. B. Neueinstellung, Leiharbeit)")
name("Lst_Team", "Personalbedarf!$B$28:$B$33")
cell(PB, "A34", "Verfügbare Kapazität", bold=True); cell(PB, "B34", None); cell(PB, "C34", "=SUM(C28:C33)", "0.00", bold=True)
cell(PB, "A35", "Brutto-Bedarf (Methode 1)", bold=True); cell(PB, "B35", None); cell(PB, "C35", "=C17", "0.00", bold=True)
cell(PB, "A36", "Über-/Unterdeckung", bold=True); cell(PB, "B36", None); cell(PB, "C36", "=C34-C35", "+0.00;-0.00;0.00", bold=True)
cell(PB, "A37", "Deckungsgrad", bold=True); cell(PB, "B37", None); cell(PB, "C37", "=IF(C35=0,\"\",C34/C35)", "0%", bold=True)
cell(PB, "A38", "Ampel", bold=True); cell(PB, "B38", None)
cell(PB, "C38", '=IF(C37="","",IF(C37>=P_DeckGruen,"gedeckt",IF(C37>=P_DeckGelb,"knapp","Unterdeckung")))', bold=True, align=CENTER)
for r in range(34, 39):
    for c in "DEF": PB[f"{c}{r}"].border = BOX
PB["D36"] = "negativ = fehlende Kapazität"; PB["D36"].font = font(size=9, i=True)
PB["D37"] = "verfügbare FTE ÷ Brutto-Bedarf"; PB["D37"].font = font(size=9, i=True)
PB.conditional_formatting.add("C38", CellIsRule(operator="equal", formula=['"gedeckt"'], fill=green))
PB.conditional_formatting.add("C38", CellIsRule(operator="equal", formula=['"knapp"'], fill=amber))
PB.conditional_formatting.add("C38", CellIsRule(operator="equal", formula=['"Unterdeckung"'], fill=red))
PB.conditional_formatting.add("C36", CellIsRule(operator="lessThan", formula=["0"], font=Font(name=F, b=True, color="C00000")))
PB["A39"] = "Im Original stand die Namensliste ohne Erläuterung unter dem Hinweis „*Schichtübergreifend“; Nr. 3 ist dort nicht vergeben."
PB["A39"].font = font(size=9, i=True, color="595959")
for c, w in zip("ABCDEF", (36, 26, 14, 18, 18, 70)): PB.column_dimensions[c].width = w
for r in (6, 7, 8, 12, 13, 14, 15, 16, 17): PB[f"F{r}"].alignment = WRAP

# Listen für Dropdowns (auf Parameter)
lr = pr + len(scales) + 3
section(P, lr, "Auswahllisten", 4)
cell(P, f"A{lr+1}", "Tätigkeitsbereiche", bold=True); cell(P, f"B{lr+1}", "DIN 31051", bold=True)
P.merge_cells(f"B{lr+1}:C{lr+1}")
for i, c in enumerate(CATS): cell(P, f"A{lr+2+i}", c)
for i, d in enumerate(DIN): P.merge_cells(f"B{lr+2+i}:C{lr+2+i}"); cell(P, f"B{lr+2+i}", d)
name("Lst_Bereich", f"Parameter!$A${lr+2}:$A${lr+1+len(CATS)}")
name("Lst_DIN", f"Parameter!$B${lr+2}:$B${lr+1+len(DIN)}")

# =====================================================================
# 4) Dashboard
# =====================================================================
D = wb.create_sheet("Dashboard")
title(D, "Dashboard Kapazitätsplanung BB2 Instandhaltung", "Automatische Auswertung – keine Eingaben auf diesem Blatt.")
D["J2"] = "Stand:"; D["J2"].font = font(b=True); D["J2"].alignment = Alignment(horizontal="right"); D["K2"] = "=P_Stand"; D["K2"].number_format = "DD.MM.YYYY"; D["K2"].font = font(b=True, color="008000")
kpis = [
    ("Bedarf netto", "=Personalbedarf!B8", "0.00 \"FTE\""),
    ("Bedarf brutto", "=Personalbedarf!C17", "0.00 \"FTE\""),
    ("Verfügbar", "=Personalbedarf!C34", "0.00 \"FTE\""),
    ("Über-/Unterdeckung", "=Personalbedarf!C36", "+0.00 \"FTE\";-0.00 \"FTE\""),
    ("Deckungsgrad", "=Personalbedarf!C37", "0%"),
    ("Präventivanteil\n(Wartung + Inspektion)", f'=IF(Kapazitätsbedarf!I{TOT}=0,0,(SUMIFS(Kapazitätsbedarf!{rng("I")},Kapazitätsbedarf!{rng("D")},"Wartung")+SUMIFS(Kapazitätsbedarf!{rng("I")},Kapazitätsbedarf!{rng("D")},"Inspektion"))/Kapazitätsbedarf!I{TOT})', "0%"),
]
TILES = [(1, 1), (2, 4), (6, 7), (8, 9), (10, 10), (11, 11)]
for (lab, f, fmt), (c0, c1_) in zip(kpis, TILES):
    if c1_ > c0:
        D.merge_cells(start_row=4, start_column=c0, end_row=4, end_column=c1_)
        D.merge_cells(start_row=5, start_column=c0, end_row=5, end_column=c1_)
    a = D.cell(4, c0, lab); a.font = font(size=9, b=True, color="FFFFFF"); a.fill = fill(NAVY); a.alignment = CENTER
    b = D.cell(5, c0, f); b.font = font(size=18, b=True, color=NAVY); b.fill = fill(BLUE_L); b.alignment = CENTER; b.number_format = fmt
D.row_dimensions[4].height = 30; D.row_dimensions[5].height = 40
D.conditional_formatting.add("H5", CellIsRule(operator="lessThan", formula=["0"], font=Font(name=F, size=18, b=True, color="C00000")))
D.conditional_formatting.add("J5", CellIsRule(operator="lessThan", formula=["P_DeckGelb"], fill=red))
D.conditional_formatting.add("J5", CellIsRule(operator="between", formula=["P_DeckGelb", "P_DeckGruen-0.0001"], fill=amber))
D.conditional_formatting.add("J5", CellIsRule(operator="greaterThanOrEqual", formula=["P_DeckGruen"], fill=green))
D.conditional_formatting.add("K5", CellIsRule(operator="lessThan", formula=["0.1"], fill=red))

# Status-Kennzahlen
section(D, 7, "Status der Bewertung", 4)
stat = [("Tätigkeiten gesamt", f"=COUNTA(Kapazitätsbedarf!{rng('E')})"),
        ("davon bewertet (> 0 Std.)", f'=COUNTIF(Kapazitätsbedarf!{rng("L")},"bewertet")'),
        ("0 Std. – Plausibilität prüfen", f'=COUNTIF(Kapazitätsbedarf!{rng("L")},"0 h*")'),
        ("offen – neue Ergänzungen", f'=COUNTIF(Kapazitätsbedarf!{rng("L")},"offen*")'),
        ("Risikoklasse „hoch“", f'=COUNTIF(Kapazitätsbedarf!{rng("S")},"hoch")'),
        ("Risikoklasse „mittel“", f'=COUNTIF(Kapazitätsbedarf!{rng("S")},"mittel")')]
for i, (a, f) in enumerate(stat):
    r = 8 + i
    D.merge_cells(f"A{r}:C{r}"); cell(D, f"A{r}", a); cell(D, f"D{r}", f, "0", bold=True, align=CENTER)
    for c in "BC": D[f"{c}{r}"].border = BOX

# Bereichsauswertung
section(D, 15, "Kapazität nach Tätigkeitsbereich", 4)
header(D, 16, ["Tätigkeitsbereich", "Std./Jahr", "FTE", "Anteil"], h=22)
for i, cat in enumerate(CATS):
    r = 17 + i
    cell(D, f"A{r}", cat)
    cell(D, f"B{r}", f"=SUMIFS(Kapazitätsbedarf!{rng('I')},Kapazitätsbedarf!{rng('C')},A{r})", "#,##0", align=CENTER)
    cell(D, f"C{r}", f"=B{r}/P_JahresStd", "0.00", align=CENTER)
    cell(D, f"D{r}", f"=IF(Kapazitätsbedarf!$I${TOT}=0,0,B{r}/Kapazitätsbedarf!$I${TOT})", "0.0%", align=CENTER)
CE = 17 + len(CATS)
cell(D, f"A{CE}", "Summe", bold=True); cell(D, f"B{CE}", f"=SUM(B17:B{CE-1})", "#,##0", bold=True, align=CENTER)
cell(D, f"C{CE}", f"=SUM(C17:C{CE-1})", "0.00", bold=True, align=CENTER); cell(D, f"D{CE}", f"=SUM(D17:D{CE-1})", "0.0%", bold=True, align=CENTER)

# DIN-Auswertung
d0 = CE + 2
section(D, d0, "Kapazität nach Instandhaltungsart (DIN 31051)", 4)
header(D, d0 + 1, ["Instandhaltungsart", "Std./Jahr", "FTE", "Anteil"], h=22)
for i, d in enumerate(DIN):
    r = d0 + 2 + i
    cell(D, f"A{r}", d)
    cell(D, f"B{r}", f"=SUMIFS(Kapazitätsbedarf!{rng('I')},Kapazitätsbedarf!{rng('D')},A{r})", "#,##0", align=CENTER)
    cell(D, f"C{r}", f"=B{r}/P_JahresStd", "0.00", align=CENTER)
    cell(D, f"D{r}", f"=IF(Kapazitätsbedarf!$I${TOT}=0,0,B{r}/Kapazitätsbedarf!$I${TOT})", "0.0%", align=CENTER)

# Pareto Top 10
section(D, 7, "Top 10 Tätigkeiten (Pareto)", 6, col=6)
header(D, 8, ["Rang", "Tätigkeit", "Std./Jahr", "FTE", "Anteil", "kumuliert"], col=6, h=22)
for k in range(1, 11):
    r = 8 + k
    idx = f"MATCH(LARGE(Kapazitätsbedarf!{rng('X')},{k}),Kapazitätsbedarf!{rng('X')},0)"
    cell(D, f"F{r}", k, "0", align=CENTER)
    cell(D, f"G{r}", f"=INDEX(Kapazitätsbedarf!{rng('E')},{idx})")
    cell(D, f"H{r}", f"=N(INDEX(Kapazitätsbedarf!{rng('I')},{idx}))", "#,##0", align=CENTER)
    cell(D, f"I{r}", f"=H{r}/P_JahresStd", "0.00", align=CENTER)
    cell(D, f"J{r}", f"=IF(Kapazitätsbedarf!$I${TOT}=0,0,H{r}/Kapazitätsbedarf!$I${TOT})", "0.0%", align=CENTER)
    cell(D, f"K{r}", f"=SUM($J$9:J{r})", "0.0%", align=CENTER)

# Top-Risiken
section(D, 20, "Top 8 Risiken (Vorschlag – zu validieren)", 6, col=6)
header(D, 21, ["Rang", "Tätigkeit", "Risiko-\nzahl", "Klasse", "Std./Jahr", "Status"], col=6, h=30)
for k in range(1, 9):
    r = 21 + k
    idx = f"MATCH(LARGE(Kapazitätsbedarf!{rng('Y')},{k}),Kapazitätsbedarf!{rng('Y')},0)"
    cell(D, f"F{r}", k, "0", align=CENTER)
    cell(D, f"G{r}", f"=INDEX(Kapazitätsbedarf!{rng('E')},{idx})")
    cell(D, f"H{r}", f"=INDEX(Kapazitätsbedarf!{rng('R')},{idx})", "0", align=CENTER)
    cell(D, f"I{r}", f"=INDEX(Kapazitätsbedarf!{rng('S')},{idx})", align=CENTER)
    cell(D, f"J{r}", f"=N(INDEX(Kapazitätsbedarf!{rng('I')},{idx}))", "#,##0", align=CENTER)
    cell(D, f"K{r}", f"=INDEX(Kapazitätsbedarf!{rng('L')},{idx})", align=CENTER)
D.conditional_formatting.add("I22:I29", CellIsRule(operator="equal", formula=['"hoch"'], fill=red))
D.conditional_formatting.add("I22:I29", CellIsRule(operator="equal", formula=['"mittel"'], fill=amber))
D.conditional_formatting.add("I22:I29", CellIsRule(operator="equal", formula=['"niedrig"'], fill=green))

# FTE-Vergleich (Diagrammdaten)
f0 = d0 + 2 + len(DIN) + 1
section(D, f0, "Bedarf vs. Verfügbarkeit", 4)
header(D, f0 + 1, ["Größe", "FTE"], h=22)
for i, (a, f) in enumerate([("Bedarf netto", "=A5"), ("Bedarf brutto (Methode 1, Original)", "=B5"), ("Bedarf brutto (Methode 2)", "=Personalbedarf!C23"), ("Verfügbar (Annahme)", "=F5")]):
    cell(D, f"A{f0+2+i}", a); cell(D, f"B{f0+2+i}", f, "0.00", align=CENTER)
for c, w in zip("ABCDEFGHIJKL", (34, 11, 9, 9, 3, 7, 46, 10, 10, 12, 22, 3)): D.column_dimensions[c].width = w
CH = f0 + 15
D[f"A{CH-1}"] = "Diagramme"; D[f"A{CH-1}"].font = font(size=12, b=True, color=NAVY)
def bar(title_, cats, vals, anchor, w=16, h=8, horiz=True, color="2E75B6", fmt="#,##0"):
    ch = BarChart(); ch.type = "bar" if horiz else "col"; ch.style = 10
    ch.title = title_; ch.legend = None
    ch.add_data(vals, titles_from_data=False); ch.set_categories(cats)
    ch.series[0].graphicalProperties.solidFill = color
    ch.series[0].graphicalProperties.line.solidFill = color
    ch.dataLabels = DataLabelList(showVal=True, showSerName=False, showCatName=False, showLegendKey=False, showPercent=False); ch.dataLabels.numFmt = fmt
    if horiz: ch.y_axis.scaling.orientation = "minMax"; ch.x_axis.scaling.orientation = "maxMin"
    ch.y_axis.majorGridlines = None; ch.y_axis.delete = True; ch.x_axis.delete = False
    ch.width, ch.height = w, h
    D.add_chart(ch, anchor)
bar("Std./Jahr nach Tätigkeitsbereich", Reference(D, min_col=1, min_row=17, max_row=CE - 1),
    Reference(D, min_col=2, min_row=17, max_row=CE - 1), f"A{CH+1}", w=18, h=9)
bar("Top 10 Tätigkeiten (Std./Jahr)", Reference(D, min_col=7, min_row=9, max_row=18),
    Reference(D, min_col=8, min_row=9, max_row=18), f"G{CH+1}", w=17, h=9)
bar("Personalbedarf vs. verfügbare Kapazität (FTE)", Reference(D, min_col=1, min_row=f0 + 2, max_row=f0 + 5),
    Reference(D, min_col=2, min_row=f0 + 2, max_row=f0 + 5), f"G{f0}", w=17, h=6.5, horiz=False, color="1F3864", fmt="0.00")

# Original-Diagramm im Archivblatt nachbilden (gleiche Datenquelle wie im Original)
pc = PieChart(); pc.title = "Std./Jahr"
pc.add_data(Reference(orig, min_col=4, min_row=4, max_row=18)); pc.set_categories(Reference(orig, min_col=1, min_row=4, max_row=18))
pc.dataLabels = DataLabelList(showPercent=True, showVal=False, showSerName=False, showCatName=False, showLegendKey=False); pc.width, pc.height = 20, 13
orig.add_chart(pc, "H2")
orig["H32"] = ("Archiv: unveränderte Kopie der Originaldatei. Hinweis: Das Original-Kreisdiagramm umfasst nur Zeilen 4–18; "
               "Balg-Bereich, Wartung und Administrativ (1.250 Std./Jahr) fehlen, daher sind die Prozentwerte überhöht.")
orig["H32"].font = font(size=9, i=True, color="C00000")
orig.sheet_view.tabSelected = False

# =====================================================================
# 5) Übersicht & Anleitung
# =====================================================================
U = wb.create_sheet("Übersicht", 0)
title(U, "Kapazitätsplanung BB2 Instandhaltung – Umsetzung", "Personalbedarfsermittlung für das Umsetzungsteam der Instandhaltung BB2")
U["A3"] = "=\"Planungsstand: \"&TEXT(P_Stand,\"DD.MM.YYYY\")&\"   ·   Verantwortlich: fz-pA / Mrugalla\""
U["A3"].font = font(b=True, color="595959")
section(U, 5, "Zweck", 3)
U["A6"] = ("Die Datei ermittelt, wie viele Mitarbeitende (FTE) das Umsetzungsteam der Instandhaltung BB2 benötigt. "
           "Grundlage ist der geschätzte tägliche Zeitbedarf je Tätigkeit. Daraus werden Jahresstunden, Netto-FTE und – über Urlaub "
           "und Krankheit – der Brutto-Personalbedarf berechnet und mit dem verfügbaren Team abgeglichen.")
U.merge_cells("A6:C6"); U["A6"].alignment = WRAP; U["A6"].font = font(); U.row_dimensions[6].height = 48
section(U, 8, "Ergebnis auf einen Blick", 3)
res = [("Personalbedarf netto", "=Personalbedarf!B8", "0.00 \"FTE\""), ("Personalbedarf brutto", "=Personalbedarf!C17", "0.00 \"FTE\""),
       ("Verfügbare Kapazität (Annahme)", "=Personalbedarf!C34", "0.00 \"FTE\""), ("Über-/Unterdeckung", "=Personalbedarf!C36", "+0.00 \"FTE\";-0.00 \"FTE\""),
       ("Bewertung", "=Personalbedarf!C38", None)]
for i, (a, f, fmt) in enumerate(res):
    r = 9 + i; cell(U, f"A{r}", a, bold=True); cell(U, f"B{r}", f, fmt, align=CENTER, bold=True)
U.conditional_formatting.add("B13", CellIsRule(operator="equal", formula=['"gedeckt"'], fill=green))
U.conditional_formatting.add("B13", CellIsRule(operator="equal", formula=['"knapp"'], fill=amber))
U.conditional_formatting.add("B13", CellIsRule(operator="equal", formula=['"Unterdeckung"'], fill=red))
section(U, 15, "Aufbau der Arbeitsmappe", 3)
header(U, 16, ["Blatt", "Inhalt", "Zielgruppe"], h=20)
sheets = [("Übersicht", "Zweck, Ergebnis, Bedienung, Legende", "alle"),
          ("Dashboard", "Kennzahlen, Pareto, Top-Risiken, Diagramme (automatisch)", "Management, Produktion"),
          ("Kapazitätsbedarf", "Tätigkeiten mit Zeitbedarf, Kategorisierung, Verantwortung, Risiko, Plan-Ist", "Instandhaltung (Eingabe)"),
          ("Personalbedarf", "Netto → Brutto, Methodenvergleich, Abgleich mit Team", "Instandhaltung, Management"),
          ("Parameter", "Alle Annahmen (Arbeitstage, Jahresstunden, Urlaub, Krankenquote) und Bewertungsskalen", "Instandhaltung, Controlling"),
          ("Analyse", "Stärken/Schwächen der Ausgangsdatei, Kennzahlenvorschläge, nächste Schritte", "alle"),
          ("Änderungsprotokoll", "Jede Änderung gegenüber dem Original mit Begründung", "QM, Revision"),
          ("Original 21.07.2025", "Unveränderte Archivkopie der Ausgangsdatei", "Nachvollziehbarkeit")]
for i, (a, b, c) in enumerate(sheets):
    r = 17 + i; cell(U, f"A{r}", a, bold=True); cell(U, f"B{r}", b); cell(U, f"C{r}", c)
    U[f"A{r}"].hyperlink = f"#'{a}'!A1"; U[f"A{r}"].font = font(b=True, color="0563C1", u="single")
section(U, 26, "Bedienung in 4 Schritten", 3)
steps = ["1. Blatt „Parameter“: Annahmen prüfen (Arbeitstage, Jahresstunden, Urlaub, Krankenquote).",
         "2. Blatt „Kapazitätsbedarf“: Std. je Arbeitstag (Spalte G) je Tätigkeit eintragen; neu ergänzte Tätigkeiten bewerten; Verantwortliche zuordnen.",
         "3. Risikovorschläge (lila) im Team validieren bzw. anpassen; Klärungsbedarf (Spalte V) abarbeiten.",
         "4. Blatt „Personalbedarf“: verfügbare FTE je Person bestätigen → Ergebnis erscheint auf Dashboard und Übersicht."]
for i, s in enumerate(steps):
    r = 27 + i; U.merge_cells(f"A{r}:C{r}"); U[f"A{r}"] = s; U[f"A{r}"].font = font(); U[f"A{r}"].alignment = WRAP; U.row_dimensions[r].height = 28
section(U, 32, "Legende", 3)
leg = [("Eingabefeld", "blaue Schrift auf gelbem Grund", INPUT, "0000FF", False),
       ("Formel / Berechnung", "schwarze Schrift – nicht überschreiben", None, "000000", False),
       ("Verknüpfung zu anderem Blatt", "grüne Schrift", None, "008000", False),
       ("Erstvorschlag (zu validieren)", "lila kursiv auf hellviolettem Grund", SUGG, "7030A0", True),
       ("Neu ergänzte Tätigkeit", "grau hinterlegt, Herkunft „Neu (Vorschlag)“", GREY_L, "000000", False)]
for i, (a, b, fc, col, it) in enumerate(leg):
    r = 33 + i
    c = cell(U, f"A{r}", a); c.font = font(color=col, i=it)
    if fc: c.fill = fill(fc)
    U.merge_cells(f"B{r}:C{r}"); cell(U, f"B{r}", b)
for c, w in zip("ABC", (38, 70, 28)): U.column_dimensions[c].width = w

# =====================================================================
# 6) Analyse
# =====================================================================
A = wb.create_sheet("Analyse")
title(A, "Analyse der Ausgangsdatei (Stand 21.07.2025)", "Grundlage der Überarbeitung – siehe auch Änderungsprotokoll")
blocks = [
 ("1. Stärken", [
  "Kompakt: vollständige Personalbedarfsrechnung auf einer Seite, druckfertig (Fußzeile, An-Seite-anpassen).",
  "Nachvollziehbare Logik: Std./Tag → Std./Woche → Std./Jahr → FTE, durchgängig per Formel.",
  "Breites Tätigkeitsspektrum (19 Positionen) inkl. Projekten/KVP, Neuanlagen, Facility Management.",
  "Trennung Netto-Bedarf und Verteilzeiten (Urlaub, Krankheit) mit Quellenangabe (Statistik fz-IH).",
  "Visualisierung der Stundenverteilung per Kreisdiagramm."]),
 ("2. Schwächen", [
  "Falsche Einheit im Spaltenkopf: „Dauer/Min./Teil“ – gerechnet wird mit Stunden je Arbeitstag (×250 = Std./Jahr).",
  "Spalte C (Std./Woche) ohne Überschrift.",
  "Parameter fest in Formeln (×5, ×250, /1750, /7, 38/220) – Änderungen fehleranfällig; 1.750 und 7 sind zwei getrennt gepflegte, voneinander abhängige Werte.",
  "Kreisdiagramm erfasst nur Zeilen 4–18: 1.250 Std./Jahr (Balg-Bereich, Wartung, Administration) fehlen, Prozentangaben dadurch verzerrt; 19 Segmente schwer lesbar.",
  "FTE je Zeile mit 1 Nachkommastelle → kleine Posten erscheinen als 0,0.",
  "Farbmarkierungen (grün, orange, rote Schrift) und Namensliste (Nr. 1, 2, 4, 5) ohne Legende; Hinweis „*Schichtübergreifend“ ohne Bezug (kein Stern in der Tabelle).",
  "„Summe Direkte Mitarbeiter“ doppelt verwendet (netto und brutto).",
  "Rechtschreibfehler (z. B. Überhollung, Konstrunktion, Managemant, Betreung, Ausladsbetreuung, Lagersystemme, Schcihtübegreifend).",
  "Kein Abgleich Bedarf ↔ verfügbares Personal; keine Verantwortlichkeiten, Prioritäten oder Risiken.",
  "Methodische Unschärfe: Verteilzeit als Zuschlag (× (1+q)) statt Verfügbarkeitsfaktor (÷ (1−q)) – Bezugsbasis der 220 AT unklar."]),
 ("3. Fehlende Inhalte", [
  "Präventive Instandhaltung mit 0 Std. bewertet (Wartung Betriebsmittel / allgemeine Maschinen) – Zuständigkeit offen.",
  "Keine Positionen für Störungsbehebung, Inspektion/gesetzliche Prüfungen (DGUV V3, UVV), Arbeitssicherheit, Fremdfirmensteuerung, Audits/QM, Energiemanagement, Einarbeitung/Wissenssicherung.",
  "Keine Weiterbildungstage oder sonstigen Ausfallzeiten in den Verteilzeiten.",
  "Keine Kennzahlen (Präventivquote, Deckungsgrad), kein Plan-Ist-Vergleich, keine Versionierung."]),
 ("4. Verbesserungspotenzial", [
  "Zentrale Parameter mit Quellenangabe → eine Änderung wirkt überall.",
  "Einheitliche Kategorisierung nach Tätigkeitsbereich und DIN 31051 (Wartung, Inspektion, Instandsetzung, Verbesserung).",
  "Risikobewertung je Tätigkeit (Auswirkung × Wahrscheinlichkeit) als Grundlage für Priorisierung.",
  "Kapazitätsabgleich mit Ampel, Dashboard mit Pareto und Top-Risiken.",
  "Plan-Ist-Spalte für die Validierung der Schätzwerte über die Zeiterfassung.",
  "Kontrollzeilen, die belegen, dass Ergebnisse zum Original identisch sind."]),
 ("5. Empfohlene neue Struktur", [
  "Übersicht → Dashboard → Kapazitätsbedarf (Eingabe) → Personalbedarf → Parameter → Analyse → Änderungsprotokoll → Original (Archiv).",
  "Eingaben nur auf Kapazitätsbedarf, Personalbedarf (Team-FTE) und Parameter; Dashboard und Übersicht rein automatisch."]),
 ("6. Zentrale Erkenntnisse aus den Daten", [
  "Netto-Bedarf 4,01 FTE (28,1 Std./Tag; 7.025 Std./Jahr), brutto 4,87 FTE – unverändert zum Original.",
  "Projekte/KVP (≈36 %) und Facility Management (≈21 %) binden zusammen > 55 % der Kapazität.",
  "Präventivanteil (Wartung + Inspektion) aktuell 0 % – für ein Instandhaltungsteam auffällig; Zuständigkeit klären.",
  "Bei 4 Personen à 1,0 FTE (Annahme) fehlen ≈0,87 FTE (Deckungsgrad ≈82 %); nach Verfügbarkeitsmethode ≈1,1 FTE.",
  "Acht ergänzte Tätigkeiten sind noch unbewertet – der tatsächliche Bedarf liegt voraussichtlich höher."]),
]
r = 4
for head, items in blocks:
    section(A, r, head, 4); r += 1
    for it in items:
        A.merge_cells(f"A{r}:D{r}"); A[f"A{r}"] = "•  " + it; A[f"A{r}"].font = font(); A[f"A{r}"].alignment = WRAP
        A.row_dimensions[r].height = 28 if len(it) > 110 else 16; r += 1
    r += 1
section(A, r, "7. Vorgeschlagene zusätzliche Kennzahlen (benötigen Daten aus CMMS / Zeiterfassung / BDE)", 4); r += 1
header(A, r, ["Kennzahl", "Berechnung", "Orientierungswert", "Datenquelle"], h=20); r += 1
kz = [("Deckungsgrad Personal", "verfügbare FTE ÷ Brutto-Bedarf", "≥ 100 %", "diese Datei"),
      ("Präventivquote", "Std. Wartung + Inspektion ÷ Gesamtstunden", "≥ 30–50 % (branchenüblich)", "diese Datei / CMMS"),
      ("Planungsgrad", "geplante IH-Stunden ÷ gesamte IH-Stunden", "≥ 80 %", "CMMS"),
      ("Wartungsplan-Erfüllung", "erledigte ÷ fällige Wartungsaufträge", "≥ 95 %", "CMMS"),
      ("MTBF / MTTR", "Betriebszeit ÷ Ausfälle / Reparaturzeit ÷ Ausfälle", "Trend steigend / fallend", "BDE, CMMS"),
      ("Technische Verfügbarkeit / OEE-Verlust durch IH", "Ausfallzeit IH ÷ geplante Laufzeit", "Ziel je Anlage", "BDE"),
      ("Plan-Ist-Abweichung Kapazität", "Ist-Std. ÷ Plan-Std. − 1", "± 10 %", "Zeiterfassung (Spalte T)"),
      ("Ersatzteilverfügbarkeit", "sofort verfügbare ÷ angeforderte Teile", "≥ 95 %", "ERP/Lager"),
      ("Überstundenquote IH", "Überstunden ÷ Sollstunden", "≤ 5 %", "Zeiterfassung"),
      ("Instandhaltungskostenquote", "IH-Kosten ÷ Wiederbeschaffungswert", "2–5 % p. a.", "Controlling")]
for row_ in kz:
    for j, v in enumerate(row_): cell(A, f"{'ABCD'[j]}{r}", v)
    r += 1
r += 1
section(A, r, "8. Nächste Schritte", 4); r += 1
for it in ["Bedeutung der Farbmarkierungen und Namensliste im Original klären (Spalte W / Blatt Personalbedarf).",
           "Zuständigkeit Wartung/Inspektion (Schicht-IH vs. Umsetzung) klären und Stunden eintragen.",
           "Acht ergänzte Tätigkeiten bewerten (oder als „nicht zutreffend“ löschen).",
           "Risikovorschläge im Team validieren; Verantwortliche und Datenbasis je Tätigkeit eintragen.",
           "Methode für Verteilzeiten festlegen (Bezugsbasis 220 AT).",
           "Quartalsweise Plan-Ist-Abgleich aus der Zeiterfassung (Spalte T)."]:
    A.merge_cells(f"A{r}:D{r}"); A[f"A{r}"] = "☐  " + it; A[f"A{r}"].font = font(); r += 1
for c, w in zip("ABCD", (40, 48, 26, 26)): A.column_dimensions[c].width = w

# =====================================================================
# 7) Änderungsprotokoll
# =====================================================================
C = wb.create_sheet("Änderungsprotokoll")
title(C, "Änderungsprotokoll", "Alle Änderungen gegenüber der Datei vom 21.07.2025 – Bewertungen (Stundenwerte) sind unverändert übernommen.")
header(C, 4, ["Nr.", "Bereich", "Original", "Neu", "Begründung"])
chg = [
 ("Struktur", "1 Blatt", "8 Blätter (Übersicht, Dashboard, Kapazitätsbedarf, Personalbedarf, Parameter, Analyse, Änderungsprotokoll, Original)", "Trennung von Eingabe, Berechnung und Auswertung; jede Zielgruppe findet ihren Einstieg."),
 ("Parameter", "250, 5, 1750, 7, 38/220, 4,1 % fest in Formeln", "Blatt „Parameter“ mit benannten Zellen und Quellen; 7 Std./Tag wird aus 1.750/250 abgeleitet", "Eine Änderung wirkt überall; keine widersprüchlichen Werte; Annahmen dokumentiert."),
 ("Einheit", "„Dauer/Min./Teil“", "„Std. / Arbeitstag“", "Die Formeln rechnen mit Stunden je Arbeitstag (× 250 = Std./Jahr). Der alte Kopf war irreführend."),
 ("Spaltenkopf", "Spalte C ohne Überschrift", "„Std. / Woche“", "Vollständige Beschriftung."),
 ("Spaltenkopf", "„1 MA = 1750“", "„FTE netto“ (Jahresstunden je FTE im Parameterblatt)", "Einheitlicher Begriff FTE; Parameter sichtbar statt im Kopf versteckt."),
 ("Bezeichnungen", "„Summe Direkte Mitarbeiter“ (2× für netto und brutto)", "„Personalbedarf netto“ / „Personalbedarf brutto“", "Eindeutigkeit."),
 ("Rechtschreibung", "Ersatzteilbeschafung, Überhollung, Ersatzeile, Druckübersetzter, Betreung, Ausladsbetreuung, Konstrunktion, Lagersystemme, Managemant, Dokumentationen, neu WKZ, Schcihtübegreifend", "korrigiert (Ersatzteilbeschaffung, Überholung, Ersatzteile, Druckübersetzer, Betreuung, Auslandsbetreuung, Konstruktion, Lagersysteme, Management, Dokumentation, neuer Werkzeuge, schichtübergreifend)", "Professionelles Erscheinungsbild."),
 ("Bezeichnungen", "„Administrativ“, „Wartung allg. Maschinen“, „Ersatzteilumlagerung u. sonstiges“", "„Administration“, „Wartung allgemeiner Maschinen“, „Ersatzteilumlagerung und Sonstiges“", "Einheitlicher Substantivstil, keine Abkürzungen."),
 ("Bewertungen", "19 Stundenwerte", "identisch übernommen (Spalte G); Kontrollzeilen belegen 28,1 Std./Tag, 4,01 FTE netto, 4,87 FTE brutto", "Kein Informationsverlust; Nachweis per Formel."),
 ("Bemerkungen", "Spalte F", "Spalte „Bemerkung (Original)“ – Wortlaut korrigiert, Inhalt unverändert", "Erhalt aller Informationen."),
 ("Farbmarkierungen", "grün / orange / rote Schrift ohne Legende", "Spalte „Kennzeichnung im Original“ (Text) + Klärungsbedarf", "Information bleibt erhalten, Bedeutung wird explizit geklärt statt erraten."),
 ("Namensliste", "Nr. 1, 2, 4, 5 mit Namen, ohne Überschrift", "Tabelle „Kapazitätsabgleich“ mit FTE-Annahme, Rolle, Kennzeichnung", "Ermöglicht Deckungsgrad-Berechnung; Annahme 1,0 FTE klar gekennzeichnet."),
 ("Hinweis", "„*Schichtübergreifend“", "Spalte „Schichtübergreifend (Ja/Nein)“ je Tätigkeit", "Der Hinweis hatte im Original keinen Bezug; jetzt je Tätigkeit pflegbar."),
 ("Kategorisierung", "–", "Tätigkeitsbereich (9 Bereiche) und Instandhaltungsart nach DIN 31051", "Gemeinsame Sprache für IH, Produktion, QM und Management; Auswertung nach Bereichen."),
 ("Neue Tätigkeiten", "–", "8 ergänzte Tätigkeiten ohne Stundenwert (Status „offen – zu bewerten“)", "Schließt Lücken (Störungen, Prüfpflichten, Arbeitssicherheit …); ändert das Ergebnis erst nach Bewertung."),
 ("Risiko", "–", "Auswirkung × Wahrscheinlichkeit, Risikoklasse mit Ampel (Erstvorschlag)", "Priorisierung, Begründung von Personalbedarf gegenüber Management."),
 ("Organisation", "–", "Verantwortlich, Datenbasis", "Klare Zuständigkeit, Transparenz über Schätzgüte."),
 ("Plan-Ist", "–", "Ist-Std./Jahr und Abweichung", "Validierung der Schätzwerte (KVP-Regelkreis)."),
 ("Verteilzeiten", "Urlaub, Krankheit", "+ Weiterbildung, Sonstige (Parameter = 0)", "Vollständigkeit ohne Änderung des Ergebnisses."),
 ("Methodik", "nur Zuschlagsmethode", "+ Verfügbarkeitsmethode als Sensitivität (≈5,10 FTE)", "Transparenz über methodische Bandbreite; Original bleibt Basis."),
 ("Kapazitätsabgleich", "–", "verfügbare FTE, Über-/Unterdeckung, Deckungsgrad, Ampel", "Kernaussage für das Management."),
 ("Diagramm", "Kreisdiagramm mit 15 von 19 Positionen", "Balkendiagramme: Bereiche, Top 10, Bedarf vs. Verfügbarkeit (alle Positionen)", "Korrekte und lesbare Darstellung."),
 ("Zahlenformat", "FTE mit 1 Nachkommastelle", "FTE mit 2 Nachkommastellen, Anteile in %", "Kleine Posten werden sichtbar."),
 ("Benutzerfreundlichkeit", "–", "Dropdowns, Eingabeprüfung (0–24 Std.), fixierte Kopfzeilen, Filter, Navigation per Link", "Weniger Eingabefehler, schnellere Bedienung."),
 ("Dateiformat", ".xlsm (enthielt keine Makros)", ".xlsx", "Keine Makro-Sicherheitswarnung; Funktionsumfang identisch."),
]
for i, row_ in enumerate(chg):
    r = 5 + i
    cell(C, f"A{r}", i + 1, "0", align=CENTER)
    for j, v in enumerate(row_): cell(C, f"{'BCDE'[j]}{r}", v, align=WRAP)
for c, w in zip("ABCDE", (5, 18, 42, 52, 55)): C.column_dimensions[c].width = w
C.freeze_panes = "A5"

# Reihenfolge & Druck
order = ["Übersicht", "Dashboard", "Kapazitätsbedarf", "Personalbedarf", "Parameter", "Analyse", "Änderungsprotokoll", "Original 21.07.2025"]
wb._sheets = [wb[n] for n in order]
wb.active = 0
for ws in wb.worksheets:
    ws.sheet_view.tabSelected = ws.title == "Übersicht"
    if ws.title == "Original 21.07.2025": continue
    ws.page_setup.orientation = "landscape"; ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0; ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.oddFooter.left.text = "&F – &A"; ws.oddFooter.center.text = "fz-pA / Mrugalla"; ws.oddFooter.right.text = "Seite &P / &N · &D"
    ws.print_options.horizontalCentered = True
K.print_title_rows = f"{HR}:{HR}"
K.page_setup.paperSize = K.PAPERSIZE_A3
K.print_area = f"A1:W{c1+3}"
for t, col in (("Übersicht", NAVY), ("Dashboard", NAVY), ("Kapazitätsbedarf", "0070C0"), ("Personalbedarf", "0070C0"),
               ("Parameter", "BF8F00"), ("Analyse", "7F7F7F"), ("Änderungsprotokoll", "7F7F7F"), ("Original 21.07.2025", "BFBFBF")):
    wb[t].sheet_properties.tabColor = col
wb.properties.title = "Kapazitätsplanung BB2 Instandhaltung – Umsetzung"
wb.save(OUT)
print("saved", OUT, "rows", R0, RN, TOT)
