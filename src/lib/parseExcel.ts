import * as XLSX from "xlsx";
import type { Punch, ParsedSheet, ColumnInfo, Mapping } from "../types";

/** Error con mensaje claro y orientado al usuario (en español). */
export class ArchivoInvalidoError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "ArchivoInvalidoError";
  }
}

/**
 * Intenta extraer una hora de un valor de celda.
 * Acepta "HH:MM", "H:MM", con segundos opcionales y AM/PM opcional, y también
 * fracciones de día de Excel (ej. 0.5 = 12:00).
 * Devuelve los minutos desde la medianoche (0–1439) o null si no es una hora.
 */
export function parseTime(value: unknown): number | null {
  if (value === null || value === undefined) return null;

  // Celda de fecha/hora real de Excel (objeto Date).
  if (value instanceof Date) {
    if (Number.isNaN(value.getTime())) return null;
    return value.getHours() * 60 + value.getMinutes();
  }

  // Fracción de día de Excel (cuando la celda viene como número 0–1).
  if (typeof value === "number") {
    return fractionToMinutes(value);
  }

  const text = String(value).trim();
  if (text === "") return null;

  // Buscar un patrón HH:MM(:SS)? dentro del texto (puede venir con fecha delante).
  const match = text.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(a\.?\s*m\.?|p\.?\s*m\.?)?/i);
  if (match) {
    let hours = parseInt(match[1], 10);
    const minutes = parseInt(match[2], 10);
    const meridiem = match[4]?.toLowerCase().replace(/[.\s]/g, "");

    if (meridiem === "pm" && hours < 12) hours += 12;
    if (meridiem === "am" && hours === 12) hours = 0;

    if (hours > 23 || minutes > 59) return null;
    return hours * 60 + minutes;
  }

  // Texto que es solo un número decimal < 1 -> fracción de día.
  if (/^0?\.\d+$/.test(text)) {
    return fractionToMinutes(parseFloat(text));
  }

  return null;
}

function fractionToMinutes(fraction: number): number | null {
  if (!Number.isFinite(fraction) || fraction < 0 || fraction >= 1) return null;
  return Math.round(fraction * 24 * 60);
}

/** ¿El valor parece un nombre? (tiene al menos una letra y no es una hora). */
export function isNameLike(value: unknown): boolean {
  const text = String(value ?? "").trim();
  if (text.length < 2) return false;
  if (parseTime(text) !== null) return false;
  return /[a-záéíóúñü]/i.test(text);
}

/** Convierte minutos desde la medianoche a "HH:MM". */
export function formatMinutes(minutes: number): string {
  const h = Math.floor(minutes / 60);
  const m = minutes % 60;
  return `${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}`;
}

/**
 * Intenta extraer una FECHA de un valor de celda y la devuelve como "YYYY-MM-DD".
 * Acepta objetos Date reales (descarta el epoch de Excel de las celdas de solo-hora),
 * y texto "DD/MM/YYYY" (formato argentino) o "YYYY-MM-DD", con hora opcional detrás.
 * Devuelve null si no es una fecha.
 */
export function parseDate(value: unknown): string | null {
  if (value instanceof Date) {
    if (Number.isNaN(value.getTime())) return null;
    return validYmd(value.getFullYear(), value.getMonth() + 1, value.getDate());
  }
  const text = String(value ?? "").trim();
  if (text === "") return null;

  // DD/MM/YYYY o DD-MM-YYYY o DD.MM.YYYY (día primero).
  let m = text.match(/^(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})\b/);
  if (m) {
    let y = parseInt(m[3], 10);
    if (y < 100) y += 2000;
    return validYmd(y, parseInt(m[2], 10), parseInt(m[1], 10));
  }
  // YYYY-MM-DD (ISO).
  m = text.match(/^(\d{4})[/.-](\d{1,2})[/.-](\d{1,2})\b/);
  if (m) {
    return validYmd(parseInt(m[1], 10), parseInt(m[2], 10), parseInt(m[3], 10));
  }
  return null;
}

/** Valida y formatea año/mes/día; descarta fuera de rango y el epoch de Excel. */
function validYmd(y: number, mo: number, d: number): string | null {
  if (y < 2000 || y > 2100 || mo < 1 || mo > 12 || d < 1 || d > 31) return null;
  return `${y}-${String(mo).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
}

/** Normaliza texto para comparar encabezados (minúsculas, sin acentos). */
function normalizar(s: string): string {
  return s
    .toLowerCase()
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .trim();
}

/** Primer índice de columna cuyo encabezado contiene alguna de las palabras clave. */
function findByHeader(columns: ColumnInfo[], keywords: string[]): number {
  for (const col of columns) {
    const label = normalizar(col.label);
    if (label && keywords.some((k) => label.includes(k))) return col.index;
  }
  return -1;
}

/**
 * Lee un archivo Excel/CSV y devuelve las filas crudas + metadatos por columna +
 * una sugerencia automática de qué columna es la hora y cuál el nombre.
 * No procesa todavía: el usuario puede confirmar o corregir las columnas.
 * Lanza ArchivoInvalidoError solo si el archivo es ilegible, vacío, o no tiene
 * ninguna columna con horas o ninguna con texto tipo nombre.
 */
export async function readSheet(file: File): Promise<ParsedSheet> {
  const buffer = await file.arrayBuffer();

  let workbook: XLSX.WorkBook;
  try {
    // codepage 65001 = UTF-8: evita que los CSV con tildes/ñ se lean como Latin-1.
    workbook = XLSX.read(buffer, { type: "array", cellDates: true, codepage: 65001 });
  } catch {
    throw new ArchivoInvalidoError(
      "No se pudo leer el archivo. Asegurate de que sea un Excel o CSV válido."
    );
  }

  const sheetName = workbook.SheetNames[0];
  if (!sheetName) {
    throw new ArchivoInvalidoError("El archivo no tiene ninguna hoja de datos.");
  }

  const sheet = workbook.Sheets[sheetName];
  // header:1 -> matriz de filas; raw:true conserva los Date de celdas de hora
  // y los strings "08:02" de los CSV tal cual (luego parseTime los normaliza).
  const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
    header: 1,
    raw: true,
    defval: "",
    blankrows: false,
  });

  if (rows.length === 0) {
    throw new ArchivoInvalidoError("El archivo está vacío.");
  }

  // Por cada columna: contar horas y nombres, y -clave- cuántos valores DISTINTOS
  // tiene. En reportes reales la fecha es una columna constante (1 distinto) que se
  // descarta sola; la hora varía mucho. Lo mismo separa la columna de nombres
  // (muchos distintos) de columnas tipo "Evento"/"Sector" (pocos valores repetidos).
  const ncols = rows.reduce((max, r) => Math.max(max, r.length), 0);
  const timeCount = new Array(ncols).fill(0);
  const nameCount = new Array(ncols).fill(0);
  const dateCount = new Array(ncols).fill(0);
  const timeDistinct: Set<number>[] = Array.from({ length: ncols }, () => new Set());
  const nameDistinct: Set<string>[] = Array.from({ length: ncols }, () => new Set());
  const dateDistinct: Set<string>[] = Array.from({ length: ncols }, () => new Set());

  for (const row of rows) {
    for (let c = 0; c < ncols; c++) {
      const cell = row[c];
      // La fecha se cuenta aparte (un Date a medianoche también es "hora 00:00").
      const d = parseDate(cell);
      if (d !== null) {
        dateCount[c]++;
        dateDistinct[c].add(d);
      }
      const t = parseTime(cell);
      if (t !== null) {
        timeCount[c]++;
        timeDistinct[c].add(t);
      } else if (isNameLike(cell)) {
        nameCount[c]++;
        nameDistinct[c].add(String(cell).trim().toLocaleLowerCase("es"));
      }
    }
  }

  // Hora: columna con más horas DISTINTAS (descarta la fecha constante).
  const heurTimeCol = bestColumn(timeDistinct.map((s) => s.size), timeCount);
  if (heurTimeCol === -1) {
    throw new ArchivoInvalidoError(
      "El archivo no contiene ninguna columna con horarios (HH:MM). No parece un registro de fichajes."
    );
  }
  const heurNameCol = bestColumn(
    nameDistinct.map((s, c) => (c === heurTimeCol ? 0 : s.size)),
    nameCount
  );

  // Encabezado: última fila no vacía anterior a la primera con hora válida.
  const firstDataRow = rows.findIndex((r) => parseTime(r[heurTimeCol]) !== null);
  let headerRow: unknown[] | null = null;
  for (let i = firstDataRow - 1; i >= 0; i--) {
    if (rows[i].some((c) => String(c ?? "").trim() !== "")) {
      headerRow = rows[i];
      break;
    }
  }

  const columns: ColumnInfo[] = [];
  for (let c = 0; c < ncols; c++) {
    const letter = XLSX.utils.encode_col(c);
    const headerLabel = String(headerRow?.[c] ?? "").trim();
    const samples: string[] = [];
    for (let i = firstDataRow < 0 ? 0 : firstDataRow; i < rows.length && samples.length < 3; i++) {
      const v = displayCell(rows[i][c]);
      if (v !== "") samples.push(v);
    }
    columns.push({
      index: c,
      letter,
      label: headerLabel || `Columna ${letter}`,
      samples,
      timeCount: timeCount[c],
      nameCount: nameCount[c],
      distinctTimes: timeDistinct[c].size,
      distinctNames: nameDistinct[c].size,
    });
  }

  // Campos nuevos (DNI/Sede/Posición/Fecha): se detectan por NOMBRE de encabezado,
  // porque Sede y Posición son ambas "texto con pocos valores" y la heurística no
  // las distingue. Si no hay encabezado que coincida, quedan sin asignar (opcionales).
  const timeCol = heurTimeCol;
  const nameByHeader = findByHeader(columns, ["persona", "nombre", "empleado", "apellido", "agente", "trabajador"]);
  const nameCol = nameByHeader >= 0 ? nameByHeader : heurNameCol;
  if (nameCol === -1) {
    throw new ArchivoInvalidoError(
      "El archivo no contiene ninguna columna con nombres. No parece un registro de fichajes."
    );
  }

  // Helper: match por encabezado, descartando colisiones con nombre/hora.
  const opt = (keywords: string[]): number => {
    const c = findByHeader(columns, keywords);
    return c === nameCol || c === timeCol ? -1 : c;
  };
  const dniCol = opt(["idfichada", "dni", "documento", "legajo", "cuil", "cuit"]);
  const sedeCol = opt(["registrador", "sede", "sucursal"]);
  const posicionCol = opt(["sector", "puesto", "posicion", "area", "cargo"]);

  // Fecha: primero por encabezado ("fecha"); si no, la columna con más fechas reales.
  let dateCol = findByHeader(columns, ["fecha", "dia"]);
  if (dateCol < 0 || dateCount[dateCol] === 0) {
    dateCol = bestColumn(dateCount, dateDistinct.map((s) => s.size));
  }
  if (dateCol === nameCol) dateCol = -1;

  return {
    rows,
    columns,
    suggested: { nameCol, timeCol, dateCol, dniCol, sedeCol, posicionCol },
    totalRows: rows.length,
  };
}

/** Texto legible de una celda (Date -> "HH:MM"). */
export function displayCell(value: unknown): string {
  if (value instanceof Date) {
    const t = parseTime(value);
    return t === null ? "" : formatMinutes(t);
  }
  return String(value ?? "").trim();
}

/**
 * Construye los fichajes a partir de las filas y la asignación de columnas.
 * Las filas de encabezado/basura se descartan solas: su celda de hora no parsea.
 */
export function buildPunches(rows: unknown[][], mapping: Mapping): Punch[] {
  const { timeCol, nameCol, dateCol, dniCol, sedeCol, posicionCol } = mapping;
  const punches: Punch[] = [];
  for (const row of rows) {
    const minutes = parseTime(row[timeCol]);
    const name = String(row[nameCol] ?? "").trim();
    if (minutes === null || name === "") continue;
    if (!isNameLike(name)) continue;
    punches.push({
      name,
      minutes,
      dni: dniCol >= 0 ? displayCell(row[dniCol]) : "",
      fecha: dateCol >= 0 ? parseDate(row[dateCol]) ?? "" : "",
      sede: sedeCol >= 0 ? displayCell(row[sedeCol]) : "",
      posicion: posicionCol >= 0 ? displayCell(row[posicionCol]) : "",
    });
  }
  return punches;
}

/** Índice de mayor valor en `primary`; desempata por `tiebreak`. -1 si todo es <= 0. */
function bestColumn(primary: number[], tiebreak: number[]): number {
  let idx = -1;
  let bestP = 0;
  let bestT = -Infinity;
  for (let i = 0; i < primary.length; i++) {
    const p = primary[i];
    if (p <= 0) continue;
    if (p > bestP || (p === bestP && tiebreak[i] > bestT)) {
      bestP = p;
      bestT = tiebreak[i];
      idx = i;
    }
  }
  return idx;
}
