import * as XLSX from "xlsx";
import type { Punch } from "../types";

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
function isNameLike(value: unknown): boolean {
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
 * Lee un archivo Excel/CSV y devuelve la lista de fichajes.
 * Detecta automáticamente la columna de nombre y la de hora.
 * Lanza ArchivoInvalidoError si el archivo no contiene nombres + horarios.
 */
export async function parseExcel(file: File): Promise<Punch[]> {
  const buffer = await file.arrayBuffer();

  let workbook: XLSX.WorkBook;
  try {
    workbook = XLSX.read(buffer, { type: "array", cellDates: true });
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
  // header:1 -> matriz de filas; raw:false -> valores formateados como texto.
  const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
    header: 1,
    raw: false,
    defval: "",
    blankrows: false,
  });

  if (rows.length === 0) {
    throw new ArchivoInvalidoError("El archivo está vacío.");
  }

  // Detectar columnas contando, por columna, cuántas celdas son horas y cuántas nombres.
  const ncols = rows.reduce((max, r) => Math.max(max, r.length), 0);
  const timeCount = new Array(ncols).fill(0);
  const nameCount = new Array(ncols).fill(0);

  for (const row of rows) {
    for (let c = 0; c < ncols; c++) {
      const cell = row[c];
      if (parseTime(cell) !== null) timeCount[c]++;
      else if (isNameLike(cell)) nameCount[c]++;
    }
  }

  const timeCol = argMax(timeCount);
  if (timeCol === -1 || timeCount[timeCol] === 0) {
    throw new ArchivoInvalidoError(
      "El archivo no contiene horarios (HH:MM). No parece un registro de fichajes."
    );
  }

  // Mejor columna de nombre, distinta a la de hora.
  let nameCol = -1;
  let best = 0;
  for (let c = 0; c < ncols; c++) {
    if (c === timeCol) continue;
    if (nameCount[c] > best) {
      best = nameCount[c];
      nameCol = c;
    }
  }
  if (nameCol === -1) {
    throw new ArchivoInvalidoError(
      "El archivo no contiene nombres de empleados. No parece un registro de fichajes."
    );
  }

  // Construir los fichajes. Las filas de encabezado se descartan solas:
  // su celda de hora ("Hora") no parsea como horario.
  const punches: Punch[] = [];
  for (const row of rows) {
    const minutes = parseTime(row[timeCol]);
    const name = String(row[nameCol] ?? "").trim();
    if (minutes === null || name === "") continue;
    if (!isNameLike(name)) continue;
    punches.push({ name, minutes });
  }

  if (punches.length === 0) {
    throw new ArchivoInvalidoError(
      "No se encontró ningún fichaje válido (nombre + horario) en el archivo."
    );
  }

  return punches;
}

function argMax(arr: number[]): number {
  let idx = -1;
  let max = -Infinity;
  for (let i = 0; i < arr.length; i++) {
    if (arr[i] > max) {
      max = arr[i];
      idx = i;
    }
  }
  return idx;
}
