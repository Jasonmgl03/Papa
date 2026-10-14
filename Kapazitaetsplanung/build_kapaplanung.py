"""Erzeugt die einfache Personalbedarfsrechnung aus der Originaldatei.

Aufruf: python build_kapaplanung.py <original.xlsm> <ziel.xlsx>
"""
import sys, datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList

SRC, OUT = sys.argv[1], sys.argv[2]
wb = openpyxl.load_workbook(SRC)
orig = wb.active
orig.title = "Original"
orig.sheet_view.tabSelected = False
ws = wb.create_sheet("Personalbedarf", 0)
wb.active = 0

NAVY, LIGHT, INPUT = "1F3864", "DDEBF7", "FFF2CC"
line = Side(style="thin", color="BFBFBF")
BOX = Border(left=line, right=line, top=line, bottom=line)
def f(size=11, b=False, color="000000"): return Font(name="Arial", size=size, b=b, color=color)
def fill(c): return PatternFill("solid", fgColor=c)
MID = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(vertical="center", wrap_text=True)
PLAIN = Alignment(vertical="center")

def put(ref, v, fmt=None, b=False, size=11, color="000000", bg=None, align=LEFT, border=True):
    c = ws[ref]; c.value = v; c.font = f(size, b, color); c.alignment = align
    if fmt: c.number_format = fmt
    if bg: c.fill = fill(bg)
    if border: c.border = BOX
    return c

ws.sheet_view.showGridLines = False
put("A1", "Personalbedarf Instandhaltung BB2 – Umsetzung", size=16, b=True, color=NAVY, border=False, align=PLAIN)
put("A2", "Wie viele Mitarbeiter brauchen wir, um diese Aufgaben zu erledigen?", size=11, color="595959", border=False, align=PLAIN)
put("E2", "Stand:", b=True, align=Alignment(horizontal="right"), border=False)
put("F2", datetime.date(2025, 7, 21), "DD.MM.YYYY", b=True, color="0000FF", bg=INPUT, align=PLAIN)

# ---- Aufgaben ----
for col, h in zip("ABCDEF", ["Nr.", "Aufgabe", "Stunden\npro Tag", "Stunden\npro Jahr", "Mitarbeiter", "Bemerkung"]):
    put(f"{col}4", h, b=True, color="FFFFFF", bg=NAVY, align=MID)
ws.row_dimensions[4].height = 32

tasks = [
    ("Ersatzteilmanagement – Planung", 1.5, None),
    ("Ersatzteilumlagerung und Sonstiges", 2, "Lagerhaltung / Ersatzteilbeschaffung extern"),
    ("Komplexe Reparaturen / Generalüberholungen", 0.3, None),
    ("Reparatur / Überholung von Ersatzteilen", 0, "Komplexe Ersatzteile (Druckübersetzer, Wasserventile …)"),
    ("Projekte und KVP (Lean-Projekte, z. B. SMED)", 10, "Entwicklung, Berechnung, Bestellung"),
    ("Training / Schulung / Betreuung / TPM / WPS", 0.5, None),
    ("Auslandseinsätze / Auslandsbetreuung", 0.5, None),
    ("Anfahren von Neuanlagen", 0.5, None),
    ("Freiprüfung neuer Werkzeuge (Versuche)", 0.2, None),
    ("Konstruktion / 3D", 1, None),
    ("Maschinenumzüge", 0.1, None),
    ("CNC-Technik (z. B. Nachdrehen Messer)", 0, None),
    ("Sonderaufgaben (Lagersysteme …)", 0.5, None),
    ("Facility Management", 6, None),
    ("Planung und Dokumentation", 0, None),
    ("Betreuung Balg-Bereich", 2, None),
    ("Wartung Betriebsmittel", 0, None),
    ("Wartung allgemeiner Maschinen", 0, None),
    ("Administration", 3, None),
    (None, None, "freie Zeile – weitere Aufgabe eintragen"),
    (None, None, "freie Zeile – weitere Aufgabe eintragen"),
    (None, None, "freie Zeile – weitere Aufgabe eintragen"),
]
R0 = 5
RN = R0 + len(tasks) - 1
SUM = RN + 1
for i, (name, h, rem) in enumerate(tasks):
    r = R0 + i
    bg = "F7F9FC" if i % 2 else None
    put(f"A{r}", i + 1, align=MID, bg=bg)
    put(f"B{r}", name, color="0000FF", bg=INPUT)
    put(f"C{r}", h, "0.0", color="0000FF", bg=INPUT, align=MID)
    put(f"D{r}", f'=IF(C{r}="","",C{r}*$I$5)', "#,##0", align=MID, bg=bg)
    put(f"E{r}", f'=IF(C{r}="","",D{r}/$I$6)', "0.00", align=MID, bg=bg)
    put(f"F{r}", rem, size=9, color="595959" if name else "A6A6A6", bg=bg)
    ws.row_dimensions[r].height = 20
put(f"B{SUM}", "Summe", b=True, bg=LIGHT)
put(f"A{SUM}", None, bg=LIGHT)
put(f"C{SUM}", f"=SUM(C{R0}:C{RN})", "0.0", b=True, bg=LIGHT, align=MID)
put(f"D{SUM}", f"=SUM(D{R0}:D{RN})", "#,##0", b=True, bg=LIGHT, align=MID)
put(f"E{SUM}", f"=SUM(E{R0}:E{RN})", "0.00", b=True, bg=LIGHT, align=MID)
put(f"F{SUM}", None, bg=LIGHT)
ws.row_dimensions[SUM].height = 22

# ---- Grundlagen (rechts) ----
put("H4", "Grundlagen", b=True, color="FFFFFF", bg=NAVY)
put("I4", None, bg=NAVY)
basics = [("Arbeitstage pro Jahr", 250, "0"),
          ("Stunden pro Mitarbeiter und Jahr", 1750, "#,##0"),
          ("Urlaubstage pro Mitarbeiter", 38, "0"),
          ("bezogen auf Arbeitstage", 220, "0"),
          ("Krankenquote", 0.041, "0.0%")]
for i, (lab, v, fmt) in enumerate(basics):
    r = 5 + i
    put(f"H{r}", lab); put(f"I{r}", v, fmt, color="0000FF", bg=INPUT, align=MID)
put("H10", "Gelbe Felder können geändert werden – alles andere rechnet automatisch.", size=9, color="595959", border=False, align=PLAIN)

# ---- Ergebnis (rechts) ----
put("H12", "Ergebnis", b=True, color="FFFFFF", bg=NAVY); put("I12", None, bg=NAVY)
res = [
    ("Mitarbeiter für die Aufgaben", f"=E{SUM}", "0.00", False),
    ('="+ Urlaub ("&TEXT(I7/I8,"0.0%")&")"', f"=E{SUM}*I7/I8", "0.00", False),
    ('="+ Krankheit ("&TEXT(I9,"0.0%")&")"', f"=E{SUM}*I9", "0.00", False),
]
for i, (lab, v, fmt, b) in enumerate(res):
    r = 13 + i
    put(f"H{r}", lab); put(f"I{r}", v, fmt, align=MID)
put("H16", "Benötigte Mitarbeiter gesamt", b=True, size=12, bg=LIGHT)
put("I16", "=SUM(I13:I15)", "0.0", b=True, size=16, color=NAVY, bg=LIGHT, align=MID)
ws.row_dimensions[16].height = 30
put("H17", "aufgerundet (Personen)", b=True)
put("I17", "=ROUNDUP(I16,0)", "0", b=True, size=12, align=MID)
put("H19", "Vorhandene Mitarbeiter", b=False)
put("I19", 4, "0.0", color="0000FF", bg=INPUT, align=MID)
put("H20", '=IF(I19>=I16,"Reserve","Es fehlen")', b=True)
put("I20", "=ABS(I19-I16)", "0.0", b=True, size=12, align=MID)
from openpyxl.formatting.rule import FormulaRule
ws.conditional_formatting.add("H20:I20", FormulaRule(formula=["$I$19<$I$16"], fill=fill("F8CBAD"), font=Font(name="Arial", b=True, color="9C0006")))
ws.conditional_formatting.add("H20:I20", FormulaRule(formula=["$I$19>=$I$16"], fill=fill("C6EFCE"), font=Font(name="Arial", b=True, color="006100")))
put("H21", "Original: 4 Personen (Scheifele, Baran, Vischer, Mrugalla)", size=9, color="595959", border=False, align=PLAIN)

# ---- Diagramm ----
ch = BarChart(); ch.type = "bar"; ch.style = 10
ch.title = "Mitarbeiter je Aufgabe"; ch.legend = None
ch.add_data(Reference(ws, min_col=5, min_row=R0, max_row=R0 + 18)); ch.set_categories(Reference(ws, min_col=2, min_row=R0, max_row=R0 + 18))
s = ch.series[0]; s.graphicalProperties.solidFill = "2E75B6"; s.graphicalProperties.line.solidFill = "2E75B6"
ch.dataLabels = DataLabelList(showVal=True, showSerName=False, showCatName=False, showLegendKey=False, showPercent=False)
ch.dataLabels.numFmt = "0.0"
ch.x_axis.scaling.orientation = "maxMin"; ch.y_axis.delete = True; ch.y_axis.majorGridlines = None
ch.gapWidth = 40; ch.width, ch.height = 22, 12
ws.add_chart(ch, f"A{SUM + 2}")

for col, w in zip("ABCDEFGHI", (5, 44, 10, 10, 12, 42, 3, 34, 12)):
    ws.column_dimensions[col].width = w
ws.freeze_panes = "A5"
ws.page_setup.orientation = "landscape"; ws.page_setup.paperSize = ws.PAPERSIZE_A4
ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 1; ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.oddFooter.left.text = "&F"; ws.oddFooter.center.text = "fz-pA / Mrugalla"; ws.oddFooter.right.text = "&D"
ws.sheet_properties.tabColor = NAVY
ws.sheet_view.tabSelected = True
wb.save(OUT)
print("ok", R0, RN, SUM)
