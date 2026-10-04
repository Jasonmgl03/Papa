// Liest Checklisten, Stammblatt-Felder und Vorlagen-Maschinen aus Instandhaltungs-Audit.html
// und schreibt sie nach tools/data.json (Eingabe für build_xlsx.py).
const fs = require("fs"), path = require("path");
const root = path.join(__dirname, "..");
const s = fs.readFileSync(path.join(root, "Instandhaltungs-Audit.html"), "utf8");
const slice = (start, end) => { const a = s.indexOf(start) + start.length, b = s.indexOf(end); return s.slice(a, b).trim().replace(/;$/, ""); };
const CL = eval("(" + slice("const CHECKLISTS = ", "const TYPE_ORDER") + ")");
const MF = eval(slice("const MFIELDS = ", "/* Die drei"));
const TM = eval(slice("const TEMPLATE_MACHINES = ", "const DEFAULT_SETTINGS"));
fs.writeFileSync(path.join(__dirname, "data.json"), JSON.stringify({ CL, MF, TM }, null, 1));
console.log("tools/data.json geschrieben:", Object.keys(CL).join(", "), MF.length, "Felder,", TM.length, "Maschinen");
