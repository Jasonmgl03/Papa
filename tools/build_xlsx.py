import json
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, CellIsRule
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter as L

import os
HERE = os.path.dirname(os.path.abspath(__file__))
D = json.load(open(os.path.join(HERE, "data.json"), encoding="utf-8"))
CL, MF, TM = D["CL"], D["MF"], D["TM"]
OUT = os.path.join(HERE, "..", "Instandhaltungs-Audit.xlsx")

F = "Arial"
NAVY = "0F172A"; BLUE = "1D4ED8"; GREY = "F1F4F8"; LINE = "D5DCE6"
INPUT = "FFF7CC"
RED, REDS = "C62F2F", "FBDADA"; AMB, AMBS = "B47400", "FCEBC2"; GRN, GRNS = "0F7A2F", "D3F0DB"; NAS = "E6E9EF"

def font(**k):
    k.setdefault("name", F); k.setdefault("size", 10); return Font(**k)
thin = Side(style="thin", color=LINE)
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
fill = lambda c: PatternFill("solid", fgColor=c)
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

wb = Workbook()
wb.remove(wb.active)

def title(ws, text, sub=None, width_to="J"):
    ws["A1"] = text; ws["A1"].font = font(size=16, bold=True, color=NAVY)
    if sub:
        ws["A2"] = sub; ws["A2"].font = font(size=10, color="5B6678", italic=True)
    ws.sheet_view.showGridLines = False

def header(ws, row, cols, start=1):
    for i, t in enumerate(cols):
        c = ws.cell(row=row, column=start + i, value=t)
        c.font = font(bold=True, color="FFFFFF"); c.fill = fill(NAVY); c.alignment = CENTER; c.border = BOX

def inp(c):
    c.fill = fill(INPUT); c.border = BOX; c.font = font(color="0000FF")

def cell(ws, ref, v, **k):
    c = ws[ref]; c.value = v; c.font = font(**k); return c

RATE = ["N.I.O", "bed. OK", "OK", "n. b."]

# ---------------------------------------------------------------- Übersicht (Einstellungen zuerst, damit Namen existieren)
ov = wb.create_sheet("Übersicht")
ins = wb.create_sheet("Anleitung")

# ---------------------------------------------------------------- Checklisten
SHEETS = {"allg": "Audit_Allgemein", "maschine": "Audit_Maschine", "werkzeug": "Audit_Werkzeuge", "kompakt": "Audit_Kompakt"}
FIRST = 12  # erste Datenzeile
meta = {}
for key in ["allg", "maschine", "werkzeug", "kompakt"]:
    c = CL[key]; name = SHEETS[key]
    ws = wb.create_sheet(name)
    title(ws, "Audit-Checkliste für: " + c["name"], "Vorlage: Blatt „%s“ · Bewertung: %s" % (c["sheet"], c["rateLabel"]))
    labels = [c["objLabel"], "Equipmentnummer", "Datum", "Auditor", "Abteilung / Bereich"]
    for i, lab in enumerate(labels):
        r = 4 + i
        cell(ws, "A%d" % r, lab, bold=True)
        ws.merge_cells("C%d:D%d" % (r, r))
        inp(ws["C%d" % r]); ws["D%d" % r].border = BOX
        if lab == "Datum":
            ws["C%d" % r].number_format = "DD.MM.YYYY"
    items = [(cat["name"], it) for cat in c["cats"] for it in cat["items"]]
    n = len(items); last = FIRST + n - 1
    E = "$E$%d:$E$%d" % (FIRST, last)
    # Ergebnisblock
    res = [("N.I.O", '=COUNTIF(%s,"N.I.O")' % E), ("bed. OK", '=COUNTIF(%s,"bed. OK")' % E), ("OK", '=COUNTIF(%s,"OK")' % E),
           ("n. b.", '=COUNTIF(%s,"n. b.")' % E), ("offen", "=%d-SUM(H4:H7)" % n)]
    cell(ws, "G3", "Ergebnis", bold=True, color=NAVY)
    for i, (lab, f) in enumerate(res):
        r = 4 + i
        cell(ws, "G%d" % r, lab); x = cell(ws, "H%d" % r, f, bold=True); x.alignment = Alignment(horizontal="center"); x.border = BOX; ws["G%d" % r].border = BOX
    cell(ws, "I4", "Erfüllungsgrad", bold=True)
    ws.merge_cells("I5:J6")
    x = cell(ws, "I5", '=IFERROR((H6+Gewicht_bedOK*H5)/(H4+H5+H6),"")', bold=True, size=20)
    x.number_format = "0%"; x.alignment = CENTER; x.border = BOX
    ws.merge_cells("I7:J7")
    x = cell(ws, "I7", '=IF(I5="","nicht bewertet",IF(I5>=Ampel_gruen,"gut",IF(I5>=Ampel_gelb,"mittel","kritisch")))', bold=True)
    x.alignment = CENTER; x.border = BOX
    for rng in ("I5", "I7"):
        ws.conditional_formatting.add(rng, FormulaRule(formula=['AND($I$5<>"",$I$5>=Ampel_gruen)'], fill=fill(GRNS), font=Font(color=GRN, bold=True)))
        ws.conditional_formatting.add(rng, FormulaRule(formula=['AND($I$5<>"",$I$5>=Ampel_gelb,$I$5<Ampel_gruen)'], fill=fill(AMBS), font=Font(color=AMB, bold=True)))
        ws.conditional_formatting.add(rng, FormulaRule(formula=['AND($I$5<>"",$I$5<Ampel_gelb)'], fill=fill(REDS), font=Font(color=RED, bold=True)))
    cell(ws, "G9", "Erfüllungsgrad = (OK + Gewicht × bed. OK) ÷ (N.I.O + bed. OK + OK); Gewicht und Ampelgrenzen auf „Übersicht“.", size=8, italic=True, color="5B6678")

    hdr = ["Nr.", "Kapitel", "Check-Punkt", "Normbezug" if key == "kompakt" else ("Check-Punkt (EN)" if key == "allg" else "Hinweis"),
           "Bewertung", "Maßnahmen / Bemerkung", "Verantwortlich", "Termin", "Erledigt", "Status", "Maßn.-Nr."]
    header(ws, FIRST - 1, hdr)
    for i, (catname, it) in enumerate(items):
        r = FIRST + i
        vals = [it[1], catname, it[2], it[3]]
        for j, v in enumerate(vals):
            x = ws.cell(row=r, column=j + 1, value=v); x.font = font(color="5B6678" if j == 3 else "000000", italic=(j == 3)); x.alignment = WRAP; x.border = BOX
        ws.cell(row=r, column=1).alignment = Alignment(horizontal="center", vertical="top")
        for col in (5, 6, 7, 8, 9):
            inp(ws.cell(row=r, column=col)); ws.cell(row=r, column=col).alignment = WRAP if col == 6 else CENTER
        ws.cell(row=r, column=8).number_format = "DD.MM.YYYY"
        x = ws.cell(row=r, column=10, value='=IF(OR(E{r}="N.I.O",E{r}="bed. OK"),IF(I{r}="ja","erledigt",IF(AND(H{r}<>"",H{r}<TODAY()),"überfällig","offen")),"")'.format(r=r))
        x.font = font(bold=True); x.alignment = CENTER; x.border = BOX
        x = ws.cell(row=r, column=11, value='=IF(OR(E{r}="N.I.O",E{r}="bed. OK"),{off}+COUNTIF($E${f}:E{r},"N.I.O")+COUNTIF($E${f}:E{r},"bed. OK"),"")'.format(r=r, f=FIRST, off="Offset_" + key))
        x.font = font(color="8A94A6", size=8); x.alignment = CENTER
        ws.row_dimensions[r].height = 30 if len(it[2]) < 70 else 45
    # Dropdowns
    dv = DataValidation(type="list", formula1='"%s"' % ",".join(RATE), allow_blank=True, errorTitle="Bewertung", error="Bitte N.I.O, bed. OK, OK oder n. b. wählen.")
    ws.add_data_validation(dv); dv.add("E%d:E%d" % (FIRST, last))
    dv2 = DataValidation(type="list", formula1='"ja,nein"', allow_blank=True)
    ws.add_data_validation(dv2); dv2.add("I%d:I%d" % (FIRST, last))
    dv3 = DataValidation(type="date", operator="greaterThan", formula1="36526", allow_blank=True, error="Bitte ein Datum eingeben (TT.MM.JJJJ).")
    ws.add_data_validation(dv3); dv3.add("H%d:H%d" % (FIRST, last)); dv3.add("C6")
    rng = "E%d:E%d" % (FIRST, last)
    for v, bg, fg in (("N.I.O", RED, "FFFFFF"), ("bed. OK", "F0A812", "1F1600"), ("OK", "13A13D", "FFFFFF"), ("n. b.", "8A94A6", "FFFFFF")):
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"%s"' % v], fill=fill(bg), font=Font(color=fg, bold=True)))
    srng = "J%d:J%d" % (FIRST, last)
    for v, bg, fg in (("überfällig", REDS, RED), ("offen", AMBS, AMB), ("erledigt", GRNS, GRN)):
        ws.conditional_formatting.add(srng, CellIsRule(operator="equal", formula=['"%s"' % v], fill=fill(bg), font=Font(color=fg, bold=True)))

    # Auswertung je Kapitel
    k0 = last + 3
    cell(ws, "A%d" % (k0 - 1), "Auswertung je Kapitel", bold=True, size=12, color=NAVY)
    header(ws, k0, ["", "Kapitel", "Punkte", "N.I.O", "bed. OK", "OK", "n. b.", "offen", "Erfüllungsgrad"])
    B = "$B$%d:$B$%d" % (FIRST, last)
    for i, cat in enumerate(c["cats"]):
        r = k0 + 1 + i
        ws.cell(row=r, column=2, value=cat["name"]).font = font()
        ws.cell(row=r, column=3, value='=COUNTIF(%s,B%d)' % (B, r))
        for j, v in enumerate(RATE):
            ws.cell(row=r, column=4 + j, value='=COUNTIFS(%s,$B%d,%s,"%s")' % (B, r, E, v))
        ws.cell(row=r, column=8, value="=C{r}-SUM(D{r}:G{r})".format(r=r))
        x = ws.cell(row=r, column=9, value='=IFERROR((F{r}+Gewicht_bedOK*E{r})/(D{r}+E{r}+F{r}),"")'.format(r=r)); x.number_format = "0%"
        for col in range(2, 10):
            ws.cell(row=r, column=col).border = BOX
            if col > 2: ws.cell(row=r, column=col).alignment = Alignment(horizontal="center"); ws.cell(row=r, column=col).font = font()
    kl = k0 + len(c["cats"])
    ws.conditional_formatting.add("I%d:I%d" % (k0 + 1, kl), FormulaRule(formula=['AND(I%d<>"",I%d>=Ampel_gruen)' % (k0 + 1, k0 + 1)], fill=fill(GRNS), font=Font(color=GRN, bold=True)))
    ws.conditional_formatting.add("I%d:I%d" % (k0 + 1, kl), FormulaRule(formula=['AND(I%d<>"",I%d>=Ampel_gelb,I%d<Ampel_gruen)' % (k0 + 1, k0 + 1, k0 + 1)], fill=fill(AMBS), font=Font(color=AMB, bold=True)))
    ws.conditional_formatting.add("I%d:I%d" % (k0 + 1, kl), FormulaRule(formula=['AND(I%d<>"",I%d<Ampel_gelb)' % (k0 + 1, k0 + 1)], fill=fill(REDS), font=Font(color=RED, bold=True)))
    ch = BarChart(); ch.type = "bar"; ch.style = 10
    ch.title = "Erfüllungsgrad je Kapitel"; ch.y_axis.scaling.min = 0; ch.y_axis.scaling.max = 1; ch.y_axis.number_format = "0%"
    ch.y_axis.majorGridlines = None
    ch.add_data(Reference(ws, min_col=9, min_row=k0, max_row=kl), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=2, min_row=k0 + 1, max_row=kl))
    ch.legend = None; ch.height = 7 + 0.4 * len(c["cats"]); ch.width = 16
    ch.x_axis.scaling.orientation = "maxMin"; ch.x_axis.delete = False; ch.y_axis.delete = False
    ch.series[0].graphicalProperties.solidFill = BLUE
    ch.dataLabels = DataLabelList(); ch.dataLabels.showVal = True; ch.dataLabels.showSerName = False; ch.dataLabels.showCatName = False; ch.dataLabels.showLegendKey = False
    ws.add_chart(ch, "B%d" % (kl + 3))

    widths = {"A": 7, "B": 22, "C": 52, "D": 32, "E": 12, "F": 38, "G": 18, "H": 12, "I": 10, "J": 12, "K": 9}
    for k, w in widths.items(): ws.column_dimensions[k].width = w
    ws.freeze_panes = "C%d" % FIRST
    ws.auto_filter.ref = "A%d:K%d" % (FIRST - 1, last)
    ws.print_title_rows = "%d:%d" % (FIRST - 1, FIRST - 1)
    ws.page_setup.orientation = "landscape"; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    meta[key] = dict(name=name, n=n, first=FIRST, last=last, short=c["short"], full=c["name"])

# ---------------------------------------------------------------- Übersicht
title(ov, "Instandhaltungs-Audit – Übersicht", "Auswertung aller Checklisten dieser Mappe. Gelbe Felder sind Eingaben.")
cell(ov, "A4", "Einstellungen", bold=True, size=12, color=NAVY)
sett = [("Firma / Werk / Abteilung", "", None, "Erscheint nur hier; frei ausfüllen."),
        ("Gewicht „bed. OK“", 0.5, "0%", "Annahme (Vorschlag): bed. OK zählt zur Hälfte. Die Excel-Vorlage enthält keine Formel."),
        ("Ampel grün ab", 0.85, "0%", "Annahme (Vorschlag), frei änderbar."),
        ("Ampel gelb ab", 0.60, "0%", "Annahme (Vorschlag), frei änderbar.")]
for i, (lab, v, nf, note) in enumerate(sett):
    r = 5 + i
    cell(ov, "A%d" % r, lab); x = ov["C%d" % r]; x.value = v; inp(x)
    if nf: x.number_format = nf
    x.comment = Comment(note, "IH-Audit")
    cell(ov, "D%d" % r, note, size=8, italic=True, color="5B6678")
wb.defined_names["Gewicht_bedOK"] = DefinedName("Gewicht_bedOK", attr_text="'Übersicht'!$C$6")
wb.defined_names["Ampel_gruen"] = DefinedName("Ampel_gruen", attr_text="'Übersicht'!$C$7")
wb.defined_names["Ampel_gelb"] = DefinedName("Ampel_gelb", attr_text="'Übersicht'!$C$8")

cell(ov, "A11", "Ergebnis je Checkliste", bold=True, size=12, color=NAVY)
header(ov, 12, ["Checkliste", "Objekt", "Datum", "Auditor", "Punkte", "N.I.O", "bed. OK", "OK", "n. b.", "offen", "Erfüllungsgrad", "Ampel", "Maßnahmen"])
for i, key in enumerate(["allg", "maschine", "werkzeug", "kompakt"]):
    m = meta[key]; r = 13 + i; s = "'%s'!" % m["name"] if " " in m["name"] else m["name"] + "!"
    vals = [m["short"], "=IF(%sC4=\"\",\"–\",%sC4)" % (s, s), "=IF(%sC6=\"\",\"–\",%sC6)" % (s, s), "=IF(%sC7=\"\",\"–\",%sC7)" % (s, s), m["n"],
            "=%sH4" % s, "=%sH5" % s, "=%sH6" % s, "=%sH7" % s, "=%sH8" % s, "=%sI5" % s, "=%sI7" % s,
            '=COUNTIF(%sJ%d:J%d,"<>")-COUNTIF(%sJ%d:J%d,"")' % (s, m["first"], m["last"], s, m["first"], m["last"])]
    for j, v in enumerate(vals):
        x = ov.cell(row=r, column=1 + j, value=v); x.border = BOX; x.font = font(bold=(j == 0), color="008000" if isinstance(v, str) and v.startswith("=") else "000000")
        if j >= 4: x.alignment = Alignment(horizontal="center")
    ov.cell(row=r, column=3).number_format = "DD.MM.YYYY"; ov.cell(row=r, column=3).alignment = Alignment(horizontal="center")
    ov.cell(row=r, column=11).number_format = "0%"
# Maßnahmen-Spalte: echte Zählung über Status-Spalte
for i, key in enumerate(["allg", "maschine", "werkzeug", "kompakt"]):
    m = meta[key]; r = 13 + i; s = m["name"] + "!"
    ov.cell(row=r, column=13, value='=COUNTIF(%sE%d:E%d,"N.I.O")+COUNTIF(%sE%d:E%d,"bed. OK")' % (s, m["first"], m["last"], s, m["first"], m["last"]))
r = 17
cell(ov, "A17", "Gesamt", bold=True)
for col in range(5, 11):
    ov.cell(row=r, column=col, value="=SUM(%s13:%s16)" % (L(col), L(col))).alignment = Alignment(horizontal="center")
x = ov.cell(row=r, column=11, value='=IFERROR((H17+Gewicht_bedOK*G17)/(F17+G17+H17),"")'); x.number_format = "0%"; x.alignment = Alignment(horizontal="center")
ov.cell(row=r, column=12, value='=IF(K17="","nicht bewertet",IF(K17>=Ampel_gruen,"gut",IF(K17>=Ampel_gelb,"mittel","kritisch")))').alignment = Alignment(horizontal="center")
ov.cell(row=r, column=13, value="=SUM(M13:M16)").alignment = Alignment(horizontal="center")
for col in range(1, 14):
    ov.cell(row=r, column=col).border = BOX; ov.cell(row=r, column=col).font = font(bold=True); ov.cell(row=r, column=col).fill = fill(GREY)
for rr in ("K13:L17",):
    ov.conditional_formatting.add(rr, FormulaRule(formula=['AND($K13<>"",$K13>=Ampel_gruen)'], fill=fill(GRNS), font=Font(color=GRN, bold=True)))
    ov.conditional_formatting.add(rr, FormulaRule(formula=['AND($K13<>"",$K13>=Ampel_gelb,$K13<Ampel_gruen)'], fill=fill(AMBS), font=Font(color=AMB, bold=True)))
    ov.conditional_formatting.add(rr, FormulaRule(formula=['AND($K13<>"",$K13<Ampel_gelb)'], fill=fill(REDS), font=Font(color=RED, bold=True)))

cell(ov, "A20", "Umsetzungsstand Maßnahmen", bold=True, size=12, color=NAVY)
header(ov, 21, ["Status", "Anzahl"])
for i, st in enumerate(["überfällig", "offen", "erledigt"]):
    r = 22 + i
    ov.cell(row=r, column=1, value=st).border = BOX
    f = "+".join('COUNTIF(%s!J%d:J%d,"%s")' % (meta[k]["name"], meta[k]["first"], meta[k]["last"], st) for k in meta)
    x = ov.cell(row=r, column=2, value="=" + f); x.border = BOX; x.alignment = Alignment(horizontal="center")
cell(ov, "D20", "Maschinenpark nach Kritikalität", bold=True, size=12, color=NAVY)
header(ov, 21, ["Klasse", "Anzahl"], start=4)
for i, (k, lab) in enumerate([("A", "A – kritisch"), ("B", "B – mittel"), ("C", "C – gering")]):
    r = 22 + i
    ov.cell(row=r, column=4, value=lab).border = BOX
    x = ov.cell(row=r, column=5, value='=COUNTIF(Maschinen!$V$5:$V$24,"%s*")' % k); x.border = BOX; x.alignment = Alignment(horizontal="center")
ov.cell(row=25, column=4, value="nicht eingestuft").border = BOX
x = ov.cell(row=25, column=5, value='=COUNTIF(Maschinen!$B$5:$B$24,"?*")-SUM(E22:E24)'); x.border = BOX; x.alignment = Alignment(horizontal="center")

def bar(ws, title_, data_ref, cat_ref, anchor, color=BLUE, pct=False, stacked=None, w=15, h=7.5, horizontal=False):
    ch = BarChart(); ch.type = "bar" if horizontal else "col"; ch.title = title_; ch.style = 10
    if stacked: ch.grouping = "stacked"; ch.overlap = 100
    for ref in data_ref: ch.add_data(ref, titles_from_data=True)
    ch.set_categories(cat_ref)
    if pct: ch.y_axis.scaling.min = 0; ch.y_axis.scaling.max = 1; ch.y_axis.number_format = "0%"
    ch.x_axis.delete = False; ch.y_axis.delete = False
    ch.width, ch.height = w, h
    if not stacked:
        ch.legend = None; ch.series[0].graphicalProperties.solidFill = color
        ch.dataLabels = DataLabelList(); ch.dataLabels.showVal = True; ch.dataLabels.showSerName = False; ch.dataLabels.showCatName = False; ch.dataLabels.showLegendKey = False
    ws.add_chart(ch, anchor); return ch

bar(ov, "Erfüllungsgrad je Checkliste", [Reference(ov, min_col=11, min_row=12, max_row=16)], Reference(ov, min_col=1, min_row=13, max_row=16), "A28", pct=True)
st = bar(ov, "Bewertungen je Checkliste", [Reference(ov, min_col=c_, min_row=12, max_row=16) for c_ in (6, 7, 8, 9)],
         Reference(ov, min_col=1, min_row=13, max_row=16), "G28", stacked=True, w=17)
for s_, col in zip(st.series, (RED, "F0A812", "13A13D", "8A94A6")):
    s_.graphicalProperties.solidFill = col; s_.graphicalProperties.line.solidFill = "FFFFFF"
st.legend.position = "b"
bar(ov, "Maßnahmen nach Status", [Reference(ov, min_col=2, min_row=21, max_row=24)], Reference(ov, min_col=1, min_row=22, max_row=24), "A44", color="B47400", w=11)
bar(ov, "Maschinen nach Kritikalität", [Reference(ov, min_col=5, min_row=21, max_row=25)], Reference(ov, min_col=4, min_row=22, max_row=25), "G44", color="5B6678", w=11)
for k, w in {"A": 18, "B": 30, "C": 12, "D": 18, "E": 9, "F": 8, "G": 9, "H": 8, "I": 8, "J": 8, "K": 14, "L": 13, "M": 12}.items():
    ov.column_dimensions[k].width = w

ov.print_area = "A1:M60"; ov.page_setup.orientation = "landscape"; ov.sheet_properties.pageSetUpPr.fitToPage = True; ov.page_setup.fitToHeight = 0
# Offsets für die globale Maßnahmen-Nummer
order = ["allg", "maschine", "werkzeug", "kompakt"]
cell(ov, "O11", "Hilfswerte (nicht ändern)", size=8, italic=True, color="8A94A6")
for i, k in enumerate(order):
    r = 12 + i
    ov.cell(row=r, column=15, value="Offset " + meta[k]["short"]).font = font(size=8, color="8A94A6")
    x = ov.cell(row=r, column=16, value=0 if i == 0 else "=P%d+M%d" % (r - 1, 13 + i - 1)); x.font = font(size=8, color="8A94A6")
    wb.defined_names["Offset_" + k] = DefinedName("Offset_" + k, attr_text="'Übersicht'!$P$%d" % r)

# ---------------------------------------------------------------- Maßnahmenplan (sammelt automatisch)
mp = wb.create_sheet("Maßnahmen", 1)
title(mp, "Maßnahmenplan", "Füllt sich automatisch aus allen Prüfpunkten mit N.I.O oder bed. OK. Eingaben bitte in den Audit-Blättern machen.")
cols = ["Nr.", "Checkliste", "Objekt", "Punkt", "Check-Punkt", "Bewertung", "Befund / Maßnahme", "Verantwortlich", "Termin", "Status"]
header(mp, 4, cols)
total_items = sum(meta[k]["n"] for k in order)
srcmap = {"Punkt": "A", "Check-Punkt": "C", "Bewertung": "E", "Befund / Maßnahme": "F", "Verantwortlich": "G", "Termin": "H", "Status": "J"}
for i in range(total_items):
    r = 5 + i; nr = i + 1
    mp.cell(row=r, column=1, value='=IF(%d<=Übersicht!$M$17,%d,"")' % (nr, nr))
    # Quelle bestimmen
    cond = lambda col: 'IF($A{r}="","",IF($A{r}<=Übersicht!$P$13,{a},IF($A{r}<=Übersicht!$P$14,{b},IF($A{r}<=Übersicht!$P$15,{c},{d}))))'
    def pick(sheetcol, as_text=True):
        parts = []
        for k in order:
            m = meta[k]
            idx = 'INDEX({s}!${c}${f}:${c}${l},MATCH($A{r},{s}!$K${f}:$K${l},0))'.format(s=m["name"], c=sheetcol, f=m["first"], l=m["last"], r=r)
            parts.append(('IF(%s="","",%s)' % (idx, idx)) if as_text else idx)
        return "=IFERROR(" + cond(None).format(r=r, a=parts[0], b=parts[1], c=parts[2], d=parts[3]) + ',"")'
    mp.cell(row=r, column=2, value="=IFERROR(" + cond(None).format(r=r, a='"%s"' % meta["allg"]["short"], b='"%s"' % meta["maschine"]["short"], c='"%s"' % meta["werkzeug"]["short"], d='"%s"' % meta["kompakt"]["short"]) + ',"")')
    objs = []
    for k in order:
        objs.append('IF(%s!$C$4="","–",%s!$C$4)' % (meta[k]["name"], meta[k]["name"]))
    mp.cell(row=r, column=3, value="=IFERROR(" + cond(None).format(r=r, a=objs[0], b=objs[1], c=objs[2], d=objs[3]) + ',"")')
    for j, lab in enumerate(cols[3:]):
        mp.cell(row=r, column=4 + j, value=pick(srcmap[lab]))
    mp.cell(row=r, column=9).number_format = "DD.MM.YYYY"
    for col in range(1, 11):
        x = mp.cell(row=r, column=col); x.border = BOX; x.alignment = WRAP if col in (5, 7) else Alignment(horizontal="center", vertical="top"); x.font = font()
rng = "A5:J%d" % (4 + total_items)
mp.conditional_formatting.add("J5:J%d" % (4 + total_items), CellIsRule(operator="equal", formula=['"überfällig"'], fill=fill(REDS), font=Font(color=RED, bold=True)))
mp.conditional_formatting.add("J5:J%d" % (4 + total_items), CellIsRule(operator="equal", formula=['"offen"'], fill=fill(AMBS), font=Font(color=AMB, bold=True)))
mp.conditional_formatting.add("J5:J%d" % (4 + total_items), CellIsRule(operator="equal", formula=['"erledigt"'], fill=fill(GRNS), font=Font(color=GRN, bold=True)))
mp.conditional_formatting.add("F5:F%d" % (4 + total_items), CellIsRule(operator="equal", formula=['"N.I.O"'], fill=fill(RED), font=Font(color="FFFFFF", bold=True)))
mp.conditional_formatting.add("F5:F%d" % (4 + total_items), CellIsRule(operator="equal", formula=['"bed. OK"'], fill=fill("F0A812"), font=Font(color="1F1600", bold=True)))
for k, w in {"A": 6, "B": 12, "C": 26, "D": 7, "E": 46, "F": 11, "G": 40, "H": 18, "I": 12, "J": 12}.items():
    mp.column_dimensions[k].width = w
mp.freeze_panes = "A5"; mp.auto_filter.ref = "A4:J%d" % (4 + total_items)
mp.page_setup.orientation = "landscape"; mp.sheet_properties.pageSetUpPr.fitToPage = True; mp.page_setup.fitToHeight = 0
mp.print_title_rows = "4:4"

# ---------------------------------------------------------------- Maschinen-Register
ms = wb.create_sheet("Maschinen")
title(ms, "Maschinen – Stammdaten", "Spalten nach Blatt „Checkliste_Maschine_Daten“ der Vorlage. Eine Zeile pro Anlage; Zeilen 5–8 aus der Vorlage übernommen.")
mh = ["Nr.", "Maschinenbezeichnung", "Equipmentnummer"] + ["%d %s" % (i + 1, f[1]) for i, f in enumerate(MF)] + ["Bemerkungen", "Quelle (Blatt der Vorlage)"]
header(ms, 4, mh)
ms.row_dimensions[4].height = 48
for i in range(20):
    r = 5 + i
    ms.cell(row=r, column=1, value=i + 1).alignment = Alignment(horizontal="center")
    for col in range(2, len(mh) + 1):
        x = ms.cell(row=r, column=col); inp(x); x.alignment = WRAP
    if i < len(TM):
        m = TM[i]
        ms.cell(row=r, column=2, value=m["name"]); ms.cell(row=r, column=3, value=m.get("equip") or None)
        notes = []
        for j, f in enumerate(MF):
            fv = m["fields"].get(f[0], {})
            if fv.get("v"): ms.cell(row=r, column=4 + j, value=fv["v"])
            if fv.get("n"): notes.append("%s: %s" % (f[1], fv["n"]))
        ms.cell(row=r, column=4 + len(MF), value="; ".join(notes) or None)
        ms.cell(row=r, column=5 + len(MF), value=m["source"])
    ms.cell(row=r, column=1).border = BOX
crit_col = L(4 + 16)  # Feld 17
assert crit_col == "T", crit_col
dvk = DataValidation(type="list", formula1='"A,B,C"', allow_blank=True)
ms.add_data_validation(dvk); dvk.add("%s5:%s24" % (crit_col, crit_col))
ms.column_dimensions["A"].width = 5; ms.column_dimensions["B"].width = 30; ms.column_dimensions["C"].width = 14
for col in range(4, len(mh) + 1):
    ms.column_dimensions[L(col)].width = 16
for col, w in ((4 + 10, 30), (4 + 21, 30), (4 + 25, 28), (4 + 26, 24), (4 + 27, 30), (4 + 28, 30), (5 + 28, 26)):
    ms.column_dimensions[L(col)].width = w
ms.freeze_panes = "C5"
for v, bg, fg in (("A", REDS, RED), ("B", AMBS, AMB), ("C", GRNS, GRN)):
    ms.conditional_formatting.add("%s5:%s24" % (crit_col, crit_col), CellIsRule(operator="equal", formula=['"%s"' % v], fill=fill(bg), font=Font(color=fg, bold=True)))

# Kritikalität-Formeln in der Übersicht auf richtige Spalte setzen
for i, k in enumerate("ABC"):
    ov.cell(row=22 + i, column=5, value='=COUNTIF(Maschinen!$%s$5:$%s$24,"%s*")' % (crit_col, crit_col, k))

# ---------------------------------------------------------------- Maschinen-Stammblatt (Druckansicht)
sb = wb.create_sheet("Stammblatt")
title(sb, "Checkliste Maschinendaten", "Maschine im gelben Feld auswählen – die Angaben kommen aus dem Blatt „Maschinen“.")
cell(sb, "A4", "Maschinenbezeichnung", bold=True)
sb.merge_cells("C4:D4"); inp(sb["C4"]); sb["C4"].value = TM[0]["name"]
dvs = DataValidation(type="list", formula1="=Maschinen!$B$5:$B$24", allow_blank=True); sb.add_data_validation(dvs); dvs.add("C4")
cell(sb, "A5", "Equipmentnummer", bold=True)
sb.merge_cells("C5:D5")
sb["C5"] = '=IFERROR(INDEX(Maschinen!$C$5:$C$24,MATCH($C$4,Maschinen!$B$5:$B$24,0))&"","")'
header(sb, 7, ["Nr.", "Check-Punkt", "Angaben", "Hinweis aus der Vorlage"])
for i, f in enumerate(MF):
    r = 8 + i
    sb.cell(row=r, column=1, value=i + 1).alignment = Alignment(horizontal="center", vertical="top")
    sb.cell(row=r, column=2, value=f[1]).font = font(bold=True)
    sb.cell(row=r, column=3, value='=IFERROR(INDEX(Maschinen!$%s$5:$%s$24,MATCH($C$4,Maschinen!$B$5:$B$24,0))&"","")' % (L(4 + i), L(4 + i)))
    sb.cell(row=r, column=4, value=f[2] or None).font = font(size=8, italic=True, color="5B6678")
    for col in range(1, 5):
        sb.cell(row=r, column=col).border = BOX; sb.cell(row=r, column=col).alignment = WRAP
        if col == 3: sb.cell(row=r, column=col).font = font(color="008000")
r = 8 + len(MF)
sb.cell(row=r, column=2, value="Bemerkungen").font = font(bold=True)
sb.cell(row=r, column=3, value='=IFERROR(INDEX(Maschinen!$%s$5:$%s$24,MATCH($C$4,Maschinen!$B$5:$B$24,0))&"","")' % (L(4 + len(MF)), L(4 + len(MF))))
sb.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
for col in range(1, 5): sb.cell(row=r, column=col).border = BOX; sb.cell(row=r, column=col).alignment = WRAP
for k, w in {"A": 6, "B": 38, "C": 40, "D": 44}.items(): sb.column_dimensions[k].width = w
sb.page_setup.fitToHeight = 1; sb.page_setup.fitToWidth = 1; sb.sheet_properties.pageSetUpPr.fitToPage = True

# ---------------------------------------------------------------- Wissen
wi = wb.create_sheet("Wissen")
title(wi, "Wissen: Instandhaltung", "Normen, Strategien und Kennzahlen, auf die sich die Checklisten beziehen. Quellen siehe Blatt „Anleitung“.")
r = 4
def block(head, cols, rows, widths=None):
    global r
    cell(wi, "A%d" % r, head, bold=True, size=12, color=NAVY); r += 1
    header(wi, r, cols); r += 1
    for row in rows:
        for j, v in enumerate(row):
            x = wi.cell(row=r, column=1 + j, value=v); x.alignment = WRAP; x.border = BOX; x.font = font(bold=(j == 0))
        r += 1
    r += 1
block("Die vier Grundmaßnahmen nach DIN 31051 (2019-06)", ["Maßnahme", "Ziel", "Beispiele", "Bezug Checkliste"], [
    ["Wartung", "Abbau des vorhandenen Abnutzungsvorrats verzögern", "Schmieren, Betriebsstoffe nachfüllen, Verschleißteile tauschen", "Schmierpläne, Wartungsnachweise"],
    ["Inspektion", "Ist-Zustand auf Konformität prüfen", "Messen, Beobachten, Funktionsprüfung", "Mechanik, Hydraulik, Pneumatik, Elektrik"],
    ["Instandsetzung", "Funktion einer fehlerhaften Einheit wiederherstellen", "Reparieren, Austauschen", "Störungen, ungeplante Reparaturen"],
    ["Verbesserung", "Zuverlässigkeit, Instandhaltbarkeit oder Sicherheit erhöhen", "Konstruktive Änderung, Beseitigen von Schwachstellen", "Schwachstellen- und Schadensanalyse"]])
block("Instandhaltungsstrategien (Begriffe nach DIN EN 13306)", ["Strategie", "Auslöser", "Stärken", "Grenzen", "Typischer Einsatz"], [
    ["Korrektiv (ausfallbedingt)", "Ausfall / Störung", "Abnutzungsvorrat voll genutzt, kein Planungsaufwand", "Ungeplante Stillstände, Folgeschäden", "C-Anlagen, redundante Einheiten"],
    ["Präventiv, vorausbestimmt", "Feste Intervalle (Zeit, Zyklen, Laufzeit)", "Gut planbar", "Teile ggf. zu früh getauscht", "Sicherheitsrelevante Teile, Herstellervorgaben, gesetzliche Prüfungen"],
    ["Zustandsorientiert", "Inspektion / Überwachung zeigt Abnutzung", "Eingriff erst bei Bedarf", "Messung, Grenzwerte und Inspektionsdisziplin nötig", "Lager, Spindeln, Hydraulik"],
    ["Voraussagend (predictive)", "Prognose aus Daten- und Trendanalyse", "Eingriff kurz vor erwartetem Ausfall", "Sensorik und Datenauswertung nötig", "A-Anlagen mit hohen Ausfallkosten"],
    ["TPM (Total Productive Maintenance)", "Ganzheitlich, Bediener übernehmen einfache Pflege", "Kontinuierliche Verbesserung der OEE", "Braucht Schulung und Führung", "Automobilindustrie (IATF 16949, 8.5.1.5)"]])
block("Normen und Vorschriften", ["Regelwerk", "Inhalt", "Bezug Checkliste"], [
    ["DIN 31051 (2019-06)", "Grundlagen der Instandhaltung: Wartung, Inspektion, Instandsetzung, Verbesserung", "Allgemein 9–14, Maschine 2–2.8"],
    ["DIN EN 13306", "Begriffe der Instandhaltung, Instandhaltungsarten und -strategien", "Allgemein 8"],
    ["DIN EN 15341", "Wesentliche Leistungskennzahlen (KPI) für die Instandhaltung (über 50 Kennzahlen)", "Allgemein 24, 37, 38"],
    ["IATF 16949:2016, 8.5.1.5", "Dokumentiertes TPM-System: Schlüsselanlagen, Ressourcen, Ersatzteilverfügbarkeit, Verpackung und Konservierung von Ausrüstung, Werkzeugen und Prüfmitteln, kundenspezifische Anforderungen, dokumentierte Instandhaltungsziele (z. B. OEE, MTBF, MTTR), regelmäßige Überprüfung, vorausschauende Methoden, periodische Überholung", "Kompakt N1–N2d"],
    ["ISO/TS 16949:2009, 7.5.1.4 und 6.4.2", "Vorgänger der IATF 16949 (2016 ersetzt). Die Normfragen im Kompakt-Audit stammen aus diesen Abschnitten", "Kompakt N1–N3"],
    ["BetrSichV § 10", "Instandhaltung von Arbeitsmitteln: sicherer Zustand während der Verwendungsdauer, Gefährdungsbeurteilung, Herstellerangaben, nur fachkundige und beauftragte Beschäftigte oder vergleichbar qualifizierte Fremdfirmen", "Allgemein 15, 16, 39"],
    ["BetrSichV § 14", "Prüfung von Arbeitsmitteln durch eine zur Prüfung befähigte Person", "Allgemein 12, 39"],
    ["TRBS 1112", "Gefährdungsbeurteilung und Schutzmaßnahmen bei Instandhaltungsarbeiten nach § 10 BetrSichV", "Allgemein 39"],
    ["DGUV Vorschrift 3", "Prüfung elektrischer Anlagen und Betriebsmittel. Richtwerte: ortsveränderliche Geräte 6 Monate (Baustellen 3), bei Fehlerquote < 2 % in Werkstätten bis 1 Jahr; ortsfeste Anlagen 4 Jahre. Maßgeblich ist die Gefährdungsbeurteilung", "Maschine 2.8"],
    ["AwSV", "WGK 1 schwach, WGK 2 deutlich, WGK 3 stark wassergefährdend; nicht eingestufte Stoffe gelten als WGK 3", "Stammblatt Feld 26"]])
block("Kritikalität A / B / C (übliche Einteilung – betriebsintern festlegen)", ["Klasse", "Beschreibung", "Strategie-Hinweis"], [
    ["A", "Ausfall stoppt die Produktion oder gefährdet Sicherheit/Qualität, keine Ersatzmaschine", "Vorbeugend oder zustandsorientiert, Ersatzteile auf Lager"],
    ["B", "Ausfall beeinträchtigt die Produktion, Ausweichen begrenzt möglich", "Planmäßige Wartung und Inspektion"],
    ["C", "Geringe Auswirkung oder Redundanz vorhanden", "Ausfallbedingte IH kann wirtschaftlich sein"]])
cell(wi, "A%d" % r, "Kennzahlen-Rechner", bold=True, size=12, color=NAVY); r += 1
cell(wi, "A%d" % r, "Gelbe Felder ändern. Die Startwerte sind Rechenbeispiele, keine Messwerte.", size=8, italic=True, color="5B6678"); r += 1
calc = [("Betriebszeit [h]", 720, None), ("Anzahl Ausfälle", 4, None), ("Reparaturzeit gesamt [h]", 12, None)]
c0 = r
for lab, v, _ in calc:
    cell(wi, "A%d" % r, lab); x = wi["B%d" % r]; x.value = v; inp(x); r += 1
outs = [("MTBF [h] = Betriebszeit ÷ Ausfälle", '=IFERROR(B%d/B%d,"")' % (c0, c0 + 1), "0.0"),
        ("MTTR [h] = Reparaturzeit ÷ Ausfälle", '=IFERROR(B%d/B%d,"")' % (c0 + 2, c0 + 1), "0.0"),
        ("Technische Verfügbarkeit = MTBF ÷ (MTBF + MTTR)", '=IFERROR(B%d/(B%d+B%d),"")' % (r, r, r + 1), "0.0%")]
for lab, f, nf in outs:
    cell(wi, "A%d" % r, lab, bold=True); x = wi["B%d" % r]; x.value = f; x.number_format = nf; x.font = font(bold=True); x.border = BOX; r += 1
r += 1
o0 = r
for lab, v in (("Verfügbarkeit", 0.90), ("Leistung", 0.95), ("Qualität", 0.99)):
    cell(wi, "A%d" % r, lab); x = wi["B%d" % r]; x.value = v; x.number_format = "0.0%"; inp(x); r += 1
cell(wi, "A%d" % r, "OEE = Verfügbarkeit × Leistung × Qualität", bold=True)
x = wi["B%d" % r]; x.value = "=B%d*B%d*B%d" % (o0, o0 + 1, o0 + 2); x.number_format = "0.0%"; x.font = font(bold=True); x.border = BOX
for k, w in {"A": 34, "B": 44, "C": 40, "D": 34, "E": 34}.items(): wi.column_dimensions[k].width = w

# ---------------------------------------------------------------- Anleitung
title(ins, "Anleitung", "Digitale Fassung von „Vorlage_Audit_Instandhaltung.xlsm“ – ohne Makros.")
lines = [
    ("So geht's", None),
    ("1.", "Kopfdaten im Audit-Blatt ausfüllen (Objekt, Equipmentnummer, Datum, Auditor, Bereich)."),
    ("2.", "In Spalte „Bewertung“ je Prüfpunkt aus der Liste wählen: N.I.O, bed. OK, OK oder n. b. (nicht bewertet / nicht zutreffend)."),
    ("3.", "Bei N.I.O und bed. OK: Befund/Maßnahme, Verantwortlich und Termin eintragen. Erledigt = ja setzt den Status auf „erledigt“."),
    ("4.", "Ergebnis, Auswertung je Kapitel und Diagramme rechnen sich automatisch. Alle Maßnahmen erscheinen im Blatt „Maßnahmen“."),
    ("5.", "Maschinen im Blatt „Maschinen“ pflegen; „Stammblatt“ zeigt eine ausgewählte Maschine druckfertig."),
    ("", ""),
    ("Farben", None),
    ("Gelbe Zellen, blaue Schrift", "Eingabefelder – nur hier eintragen."),
    ("Schwarze / grüne Schrift", "Formeln bzw. Verweise auf andere Blätter – nicht überschreiben."),
    ("", ""),
    ("Beispiel (nur Format)", "Bewertung: bed. OK · Maßnahme: Schmierplan an der Maschine aushängen · Verantwortlich: Meister IH · Termin: 31.12.2026 · Erledigt: nein"),
    ("", ""),
    ("Berechnung", None),
    ("Erfüllungsgrad", "(OK + Gewicht × bed. OK) ÷ (N.I.O + bed. OK + OK). „n. b.“ und leere Punkte zählen nicht. Die Vorlage hatte Zeilen „Sum/Total“ ohne Formel; Gewicht 50 % und Ampel (grün ab 85 %, gelb ab 60 %) sind Annahmen und auf „Übersicht“ änderbar."),
    ("Mehrere Maschinen", "Pro Mappe ist je Checkliste ein Audit vorgesehen. Für weitere Maschinen die Mappe kopieren (Datei speichern unter …)."),
    ("", ""),
    ("Hinweise zur Vorlage", None),
    ("Prüfpunkt 24", "Deutscher Text nennt 90 % technische Verfügbarkeit, die englische Fassung 80 %. Bitte den gültigen Zielwert klären."),
    ("Kompakt-Audit", "Normfragen stammen aus ISO/TS 16949:2009 (7.5.1.4, 6.4.2); diese wurde 2016 durch IATF 16949 ersetzt (heute 8.5.1.5)."),
    ("Korrekturen", "Rechtschreibfehler der Vorlage wurden korrigiert, Inhalte nicht verändert. Das Kompakt-Audit entspricht dem Blatt „Auditcheckliste_1“."),
    ("", ""),
    ("Quellen", None),
    ("DIN 31051", "https://www.baunormenlexikon.de/norm/din-31051/8d1293a4-de10-49d4-9fee-0b415e6545a1"),
    ("DIN 31051", "https://www.buerklin.com/de/elektronik-kompetenz/instandhaltung/din-31051-grundlagen-der-instandhaltung/"),
    ("DIN EN 13306", "https://ihb.tuev-media.de/docs/din_en_13306.html"),
    ("Kennzahlen", "https://blog.ccc-industriesoftware.de/kennzahlen-und-kpis-instandhaltung/"),
    ("MTTR / MTBF", "https://comain.cloud/glossar/mttr-mtbf"),
    ("OEE", "https://de.wikipedia.org/wiki/Gesamtanlageneffektivit%C3%A4t"),
    ("IATF 16949 TPM", "https://community.advisera.com/topic/total-productive-maintenance-in-iatf-16949/"),
    ("IATF 16949 8.5.1.5", "https://elsmar.com/elsmarqualityforum/threads/iatf-16949-section-8-5-1-5-total-productive-maintenance.70125/"),
    ("ISO/TS 16949 7.5.1.4", "https://elsmar.com/elsmarqualityforum/threads/ts-16949-clause-7-5-1-4-preventive-and-predictive-maintenance-requirements-scope.54706/"),
    ("BetrSichV", "https://www.gesetze-im-internet.de/betrsichv_2015/BJNR004910015.html"),
    ("TRBS 1112", "https://www.baua.de/DE/Angebote/Regelwerk/TRBS/pdf/TRBS-1112.pdf"),
    ("DGUV V3", "https://www.forum-verlag.com/fachwissen/elektrosicherheit-und-elektrotechnik/prueffristen-nach-dguv-v3/"),
    ("AwSV / WGK", "https://www.umweltbundesamt.de/system/files/medien/421/dokumente/info_awsv_dieter_wgk.pdf"),
]
r = 4
for a, b in lines:
    if b is None:
        cell(ins, "A%d" % r, a, bold=True, size=12, color=NAVY)
    else:
        cell(ins, "A%d" % r, a, bold=True); x = cell(ins, "B%d" % r, b); x.alignment = WRAP
        if b.startswith("http"): x.hyperlink = b; x.font = font(color="1D4ED8", underline="single")
        if a == "Gelbe Zellen, blaue Schrift": ins["A%d" % r].fill = fill(INPUT); ins["A%d" % r].font = font(bold=True, color="0000FF")
    r += 1
ins.column_dimensions["A"].width = 26; ins.column_dimensions["B"].width = 110

# Reihenfolge der Blätter
wb._sheets = [wb[n] for n in ["Übersicht", "Audit_Allgemein", "Audit_Maschine", "Audit_Werkzeuge", "Audit_Kompakt", "Maßnahmen", "Maschinen", "Stammblatt", "Wissen", "Anleitung"]]
ov.sheet_properties.tabColor = BLUE
for k in SHEETS.values(): wb[k].sheet_properties.tabColor = "13A13D"
wb["Maßnahmen"].sheet_properties.tabColor = "F0A812"
wb.active = 0
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if c.value is not None and c.font and c.font.name != F:
                c.font = Font(name=F, size=c.font.size or 10, bold=c.font.bold, italic=c.font.italic, color=c.font.color)
wb.save(OUT)
print("saved", OUT)
