// Genera un Excel de ejemplo con DATOS FICTICIOS que imita la arquitectura de un
// "Reporte de Eventos" real de fichadas (columnas dispersas, filas de título, una
// columna de fecha constante, columnas señuelo). Sirve como fixture de pruebas sin PII.
//   node scripts/gen-example.mjs
import * as XLSX from "xlsx";
import fs from "node:fs";

const W = 43; // ancho como el reporte real (A..AQ)
const blank = () => Array(W).fill("");
const rows = [];

// --- Filas de título / basura (0–4) ---
let r;
r = blank(); r[1] = "Desde: "; r[3] = "18/06/2026"; r[10] = "SECTOR"; r[20] = "Persona:"; r[29] = "Evento"; r[36] = "Reporte de Eventos"; rows.push(r);
r = blank(); r[1] = "Hasta:"; r[3] = "18/06/2026"; r[10] = "SUCURSAL"; r[20] = "Categoria"; r[32] = "Fecha"; rows.push(r);
r = blank(); r[10] = "COSTOS"; r[20] = "Registrador"; rows.push(r);
r = blank(); r[37] = "Hotel Ejemplo SRL"; rows.push(r);
r = blank(); r[37] = "30999999999"; rows.push(r);
// --- Fila de encabezados (5), dispersos como en el real ---
r = blank(); r[1] = "Fecha / hora"; r[7] = "Evento"; r[13] = "idFichada"; r[18] = "Persona"; r[22] = "SECTOR"; r[27] = "EmpresaVisita"; r[34] = "Visitado"; r[39] = "Registrador"; rows.push(r);

const DATE = new Date(2026, 5, 18); // 18/06/2026 (mes 5 = junio)
const SECTORS = ["Alimentos y bebidas", "Housekeeping", "Recepción", "Mantenimiento"];
const REGS = ["HTL_URBANO", "HTL_CITY", "HTL_9DEJULIO"];

// [Apellido, Nombre, [horarios del día], idxSector, idxRegistrador]
const EMPLEADOS = [
  ["González, María", ["08:01", "12:30", "17:05"], 0, 0], // 3 fichajes: intermedio ignorado
  ["Rodríguez, Juan", ["07:55", "16:40"], 0, 1],
  ["Fernández, Lucía", ["09:15"], 1, 0], // 1 fichaje: salida en blanco
  ["Pérez, Diego", ["08:10", "18:02"], 2, 2],
  ["López, Antonella", ["06:58", "15:03"], 1, 1],
  ["Gómez, Ángel", ["07:30", "11:00", "12:00", "16:30"], 3, 0], // 4 fichajes
  ["Díaz, Sofía", ["08:45", "17:36"], 2, 2],
  ["Martínez, Bruno", ["07:02", "16:37"], 3, 1],
  ["Sánchez, Camila", ["08:14", "18:09"], 1, 0],
  ["Romero, Tomás", ["09:00"], 0, 2], // 1 fichaje
  ["Torres, Valentina", ["08:36", "18:04"], 2, 1],
  ["Ramírez, Mateo", ["07:58", "17:02"], 0, 0],
  ["Flores, Julieta", ["08:39", "17:36"], 1, 2],
  ["Benítez, Nicolás", ["11:54", "13:00", "17:02"], 3, 1], // 3 fichajes
  ["Acosta, Florencia", ["08:01", "18:19"], 2, 0],
  ["Núñez, Joaquín", ["22:30"], 2, 1], // turno noche: entra de noche, 1 fichaje
];

let id = 40000000;
const eventos = [];
for (const [nombre, horas, si, ri] of EMPLEADOS) {
  for (const hora of horas) {
    eventos.push({ nombre, hora, sector: SECTORS[si], reg: REGS[ri] });
  }
}
// Ordenar por hora ascendente, como el reporte real.
eventos.sort((a, b) => a.hora.localeCompare(b.hora));

for (const e of eventos) {
  const row = blank();
  row[1] = new Date(DATE); // fecha (constante) como Date
  row[5] = e.hora;         // hora real como texto "HH:MM"
  row[7] = "Entrada ZK";
  row[13] = String(id++);
  row[18] = e.nombre;
  row[22] = e.sector;
  row[39] = e.reg;
  rows.push(row);
}
// Fila de pie como en el real.
r = blank(); r[1] = "Fecha de emision"; r[3] = "19/06/2026"; rows.push(r);

const ws = XLSX.utils.aoa_to_sheet(rows, { cellDates: true });
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
const out = "ejemplos/fichajes-ejemplo-zk.xlsx";
XLSX.writeFile(wb, out);
console.log(`OK -> ${out} (${eventos.length} eventos, ${EMPLEADOS.length} empleados)`);
