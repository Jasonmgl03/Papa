# Build-Skripte

Die Excel-Fassung `Instandhaltungs-Audit.xlsx` wird aus den Checklisten der Web-App erzeugt,
damit beide Fassungen inhaltlich gleich bleiben.

```bash
node tools/extract_data.js     # Checklisten aus Instandhaltungs-Audit.html -> tools/data.json
python3 tools/build_xlsx.py    # erzeugt Instandhaltungs-Audit.xlsx (benötigt openpyxl)
```

Danach die Datei einmal in Excel oder LibreOffice öffnen und speichern, damit die Formelwerte
zwischengespeichert sind (Excel rechnet beim Öffnen ohnehin neu).
