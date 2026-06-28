// Genera un Excel ficticio MULTI-DÍA con la arquitectura del reporte real de fichadas.
// Incluye: DNI (idFichada), Sede (Registrador), Posición (Sector), fecha por fila,
// un empleado que cambia de sede entre entrada y salida, y un turno noche que cruza
// la medianoche (queda como dos jornadas incompletas, según lo acordado).
//   node scripts/gen-example-multidia.mjs
import * as XLSX from "xlsx";

const W = 43;
const blank = () => Array(W).fill("");
const rows = [];

// Título / basura (0–4)
let r;
r = blank(); r[1] = "Desde: "; r[3] = "18/06/2026"; r[10] = "SECTOR"; r[20] = "Persona:"; r[36] = "Reporte de Eventos"; rows.push(r);
r = blank(); r[1] = "Hasta:"; r[3] = "19/06/2026"; r[10] = "SUCURSAL"; rows.push(r);
r = blank(); r[10] = "COSTOS"; rows.push(r);
r = blank(); r[37] = "Hotel Ejemplo SRL"; rows.push(r);
r = blank(); r[37] = "30999999999"; rows.push(r);
// Encabezados (5)
r = blank(); r[1] = "Fecha / hora"; r[7] = "Evento"; r[13] = "idFichada"; r[18] = "Persona"; r[22] = "SECTOR"; r[34] = "Visitado"; r[39] = "Registrador"; rows.push(r);

const URBANO = "HTL_URBANO", CITY = "HTL_CITY";
const AYB = "Alimentos y bebidas", HK = "Housekeeping", REC = "Recepción";

// Por empleado: DNI, Apellido Nombre, y por día una lista de eventos {hora, sede, sector}.
const EMP = [
  {
    dni: "40123456", nombre: "González, María",
    dias: {
      "18/06/2026": [{ h: "08:01", s: URBANO, p: HK }, { h: "12:30", s: URBANO, p: HK }, { h: "17:05", s: URBANO, p: HK }],
      "19/06/2026": [{ h: "07:58", s: URBANO, p: HK }, { h: "16:40", s: URBANO, p: HK }],
    },
  },
  {
    // Cambia de sede: entra en URBANO, sale en CITY.
    dni: "41234567", nombre: "Rodríguez, Juan",
    dias: {
      "18/06/2026": [{ h: "07:55", s: URBANO, p: AYB }, { h: "16:40", s: CITY, p: AYB }],
      "19/06/2026": [{ h: "08:05", s: CITY, p: AYB }, { h: "17:10", s: CITY, p: AYB }],
    },
  },
  {
    dni: "42345678", nombre: "Fernández, Lucía",
    dias: {
      "18/06/2026": [{ h: "09:15", s: CITY, p: REC }], // 1 fichaje: salida en blanco
      "19/06/2026": [{ h: "09:02", s: CITY, p: REC }, { h: "18:00", s: CITY, p: REC }],
    },
  },
  {
    dni: "43456789", nombre: "Pérez, Diego",
    dias: {
      "18/06/2026": [{ h: "08:10", s: URBANO, p: AYB }, { h: "18:02", s: URBANO, p: AYB }],
      "19/06/2026": [{ h: "08:14", s: URBANO, p: AYB }, { h: "18:09", s: URBANO, p: AYB }],
    },
  },
  {
    // Turno noche que cruza medianoche: 18 a la noche entra, 19 a la mañana sale.
    dni: "44567890", nombre: "Núñez, Joaquín",
    dias: {
      "18/06/2026": [{ h: "22:30", s: URBANO, p: REC }], // entra de noche (queda incompleto)
      "19/06/2026": [{ h: "06:05", s: URBANO, p: REC }], // sale a la mañana (queda incompleto)
    },
  },
  {
    dni: "45678901", nombre: "Torres, Valentina",
    dias: {
      "18/06/2026": [{ h: "08:36", s: CITY, p: HK }, { h: "18:04", s: CITY, p: HK }],
      "19/06/2026": [{ h: "08:40", s: CITY, p: HK }, { h: "17:50", s: CITY, p: HK }],
    },
  },
];

const eventos = [];
let id = 0;
for (const e of EMP) {
  for (const [fecha, lista] of Object.entries(e.dias)) {
    for (const ev of lista) {
      eventos.push({ fecha, ...ev, dni: e.dni, nombre: e.nombre });
    }
  }
}
// Ordenar por fecha y hora (como el reporte real).
eventos.sort((a, b) => a.fecha.localeCompare(b.fecha) || a.h.localeCompare(b.h));

for (const ev of eventos) {
  const [d, m, y] = ev.fecha.split("/").map(Number);
  const row = blank();
  row[1] = new Date(y, m - 1, d); // fecha como Date (a medianoche)
  row[5] = ev.h;
  row[7] = "Entrada ZK";
  row[13] = ev.dni;
  row[18] = ev.nombre;
  row[22] = ev.p;
  row[39] = ev.s;
  rows.push(row);
  id++;
}
r = blank(); r[1] = "Fecha de emision"; r[3] = "20/06/2026"; rows.push(r);

const ws = XLSX.utils.aoa_to_sheet(rows, { cellDates: true });
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
const out = "ejemplos/fichajes-ejemplo-multidia.xlsx";
XLSX.writeFile(wb, out);
console.log(`OK -> ${out} (${id} eventos, ${EMP.length} empleados, 2 días)`);
