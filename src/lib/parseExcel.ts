import * as XLSX from "xlsx";
import type { Punch, ParsedSheet, ColumnInfo } from "../types";

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
  const timeDistinct: Set<number>[] = Array.from({ length: ncols }, () => new Set());
  const nameDistinct: Set<string>[] = Array.from({ length: ncols }, () => new Set());

  for (const row of rows) {
    for (let c = 0; c < ncols; c++) {
      const cell = row[c];
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

  // Sugerencias automáticas.
  const suggestedTimeCol = bestColumn(timeDistinct.map((s) => s.size), timeCount);
  const suggestedNameCol = bestColumn(
    nameDistinct.map((s, c) => (c === suggestedTimeCol ? 0 : s.size)),
    nameCount
  );

  if (suggestedTimeCol === -1) {
    throw new ArchivoInvalidoError(
      "El archivo no contiene ninguna columna con horarios (HH:MM). No parece un registro de fichajes."
    );
  }
  if (suggestedNameCol === -1) {
    throw new ArchivoInvalidoError(
      "El archivo no contiene ninguna columna con nombres. No parece un registro de fichajes."
    );
  }

  // Detectar fila de encabezado para etiquetar columnas: la última fila no vacía
  // anterior a la primera fila que tiene una hora en la columna de hora sugerida.
  const firstDataRow = rows.findIndex((r) => parseTime(r[suggestedTimeCol]) !== null);
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

  return {
    rows,
    columns,
    suggestedTimeCol,
    suggestedNameCol,
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
 * Construye los fichajes a partir de las filas y las columnas elegidas.
 * Las filas de encabezado/basura se descartan solas: su celda de hora no parsea.
 */
export function buildPunches(
  rows: unknown[][],
  timeCol: number,
  nameCol: number
): Punch[] {
  const punches: Punch[] = [];
  for (const row of rows) {
    const minutes = parseTime(row[timeCol]);
    const name = String(row[nameCol] ?? "").trim();
    if (minutes === null || name === "") continue;
    if (!isNameLike(name)) continue;
    punches.push({ name, minutes });
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
