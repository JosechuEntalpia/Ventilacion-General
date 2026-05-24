const express = require("express");
const cors = require("cors");
const fs = require("fs");
const path = require("path");
const PizZip = require("pizzip");

const app = express();
app.use(cors());
app.use(express.json({ limit: "5mb" }));

const PORT = process.env.PORT || 3000;

function n(v) {
  if (v === null || v === undefined || v === "") return 0;
  const x = Number(String(v).replace(",", "."));
  return Number.isFinite(x) ? x : 0;
}
function s(v) { return (v === null || v === undefined) ? "" : String(v); }
function roundExcel(v) {
  if (typeof v !== "number" || !Number.isFinite(v)) return v;
  return Math.round((v + Number.EPSILON) * 1000000000000) / 1000000000000;
}
function metPersona(ida, personas, met) {
  if (ida === "IDA 1") return personas * (met / 1.2) * 72;
  if (ida === "IDA 2") return personas * (met / 1.2) * 45;
  if (ida === "IDA 3") return personas * (met / 1.2) * 28.8;
  if (ida === "IDA 4") return personas * (met / 1.2) * 18;
  return "Error";
}
// Transcripción literal de la lógica Excel observada: la fórmula tiene dos ramas IDA 3 y no una rama real para IDA 4.
function metCO2FormulaExcel(ida, personas, met) {
  if (ida === "IDA 1") return 3.6 * personas * (met / 1.2) * ((0.0042 * met) / 350) * 1000000;
  if (ida === "IDA 2") return 3.6 * personas * (met / 1.2) * ((0.0042 * met) / 500) * 1000000;
  if (ida === "IDA 3") return 3.6 * personas * (met / 1.2) * ((0.0042 * met) / 800) * 1000000;
  if (ida === "IDA 3") return 3.6 * personas * (met / 1.2) * ((0.0042 * met) / 1200) * 1000000;
  return "Error";
}
function olfAreaSala1(ida, area, efVent) {
  if (ida === "IDA 1") return 3.6 * area * 10 * (0.1 / 0.8) * (1 / efVent);
  if (ida === "IDA 2") return 3.6 * area * 10 * (0.1 / 1.2) * (1 / efVent);
  if (ida === "IDA 3") return 3.6 * area * 10 * (0.1 / 2) * (1 / efVent);
  if (ida === "IDA 4") return 3.6 * area * 10 * (0.1 / 3) * (1 / efVent);
  return "Error";
}
function olfAreaSala2Fila19(ida, area, per12) {
  if (ida === "IDA 1") return 3.6 * area * 10 * (0.1 / 0.8) * (1 / per12);
  if (ida === "IDA 2") return 3.6 * area * 10 * (0.1 / 1.2) * (1 / per12);
  if (ida === "IDA 3") return 3.6 * area * 10 * (0.1 / 2) * (1 / per12);
  if (ida === "IDA 4") return 3.6 * area * 10 * (0.1 / 3) * (1 / per12);
  return "Error";
}
// Transcripción literal de la fórmula de F20: repite IDA 3 en la última rama.
function olfAreaSala2Fila20(ida, area, per12) {
  if (ida === "IDA 1") return 3.6 * area * 10 * (0.1 / 0.8) * (1 / per12);
  if (ida === "IDA 2") return 3.6 * area * 10 * (0.1 / 1.2) * (1 / per12);
  if (ida === "IDA 3") return 3.6 * area * 10 * (0.1 / 2) * (1 / per12);
  if (ida === "IDA 3") return 3.6 * area * 10 * (0.1 / 3) * (1 / per12);
  return "Error";
}
function calcular(input) {
  const v = { ...input };
  const ida01 = s(v.IDA01);
  const ida02 = s(v.IDA02);
  const area01 = n(v.Area01);
  const area02 = n(v.Area02);
  const efVent = n(v.EfVent);
  const per12 = n(v.Per12);

  v.MPersona11 = roundExcel(metPersona(ida01, n(v.Per11), n(v.Met11)));
  v.MCO211 = roundExcel(metCO2FormulaExcel(ida01, n(v.Per11), n(v.Met11)));
  v.MOlfArea11 = roundExcel(olfAreaSala1(ida01, area01, efVent));

  v.MPersona12 = roundExcel(metPersona(ida01, n(v.Per12), n(v.Met12)));
  v.MCO212 = roundExcel(metCO2FormulaExcel(ida01, n(v.Per12), n(v.Met12)));
  v.MOlfArea12 = roundExcel(olfAreaSala1(ida01, area01, efVent));

  v.MPersona21 = roundExcel(metPersona(ida02, n(v.Per21), n(v.Met21)));
  v.MCO221 = roundExcel(metCO2FormulaExcel(ida02, n(v.Per21), n(v.Met21)));
  v.MOlfArea21 = roundExcel(olfAreaSala2Fila19(ida02, area02, per12));

  v.MPersona22 = roundExcel(metPersona(ida02, n(v.Per22), n(v.Met22)));
  v.MCO222 = roundExcel(metCO2FormulaExcel(ida02, n(v.Per22), n(v.Met22)));
  v.MOlfArea22 = roundExcel(olfAreaSala2Fila20(ida02, area02, per12));

  // Estas variables existen en la hoja de variables, pero no tienen fórmula en la hoja de lógica.
  v.MOlfPer11 = v.MOlfPer11 || "";
  v.MOlfPer12 = v.MOlfPer12 || "";
  v.MOlfPer21 = v.MOlfPer21 || "";
  v.MOlfPer22 = v.MOlfPer22 || "";
  return v;
}
function escapeXml(str) {
  return String(str ?? "")
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;").replace(/'/g, "&apos;");
}
function generarDocx(variables) {
  const templatePath = path.join(__dirname, "plantilla.docx");
  const zip = new PizZip(fs.readFileSync(templatePath));
  const files = Object.keys(zip.files).filter(name => name.startsWith("word/") && name.endsWith(".xml"));
  for (const file of files) {
    let txt = zip.file(file).asText();
    for (const [key, value] of Object.entries(variables)) {
      const token = "{" + key + "}";
      txt = txt.split(token).join(escapeXml(value));
    }
    zip.file(file, txt);
  }
  return zip.generate({ type: "nodebuffer", compression: "DEFLATE" });
}

app.get("/", (req, res) => res.send("API Ventilación Genérica IEE funcionando"));
app.get("/api/health", (req, res) => res.json({ ok: true }));
app.post("/api/calcular", (req, res) => res.json({ variables: calcular(req.body || {}) }));
app.post("/api/informe-word", (req, res) => {
  const variables = calcular(req.body || {});
  const buffer = generarDocx(variables);
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
  res.setHeader("Content-Disposition", "attachment; filename=Ventilacion-Generica-Informe.docx");
  res.send(buffer);
});
app.listen(PORT, () => console.log("API funcionando en puerto " + PORT));
