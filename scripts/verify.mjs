// Verifica el pipeline real contra un archivo. Uso: node scripts/verify.mjs <ruta>
import fs from "node:fs";
import { readSheet, buildPunches } from "../src/lib/parseExcel.ts";
import { processRecords } from "../src/lib/processRecords.ts";

const path = process.argv[2];
const buf = fs.readFileSync(path);
const file = new File([buf], path.split(/[\\/]/).pop());

const sheet = await readSheet(file);
const m = sheet.suggested;
const L = (i) => (i < 0 ? "(ninguna)" : `${sheet.columns[i].letter}:${sheet.columns[i].label}`);
console.log("Detectado -> Nombre:", L(m.nameCol), "| Hora:", L(m.timeCol), "| Fecha:", L(m.dateCol),
  "| DNI:", L(m.dniCol), "| Sede:", L(m.sedeCol), "| Posición:", L(m.posicionCol));

const rows = processRecords(buildPunches(sheet.rows, m));
console.log(`Filas de resultado: ${rows.length}\n`);
for (const r of rows) {
  console.log(
    `${(r.dni || "—").padEnd(9)} ${r.nombre.padEnd(22)} ${r.fecha || "—"} | ${r.posicion.padEnd(20)} | ` +
    `ent ${r.entrada}@${(r.sedeEntrada || "—").padEnd(10)} | sal ${(r.salida || "—").padEnd(5)}@${(r.sedeSalida || "—").padEnd(10)} | tot ${r.total || "—"}`
  );
}
