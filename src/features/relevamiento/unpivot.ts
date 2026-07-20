import * as XLSX from "xlsx";

/** Error con mensaje claro y orientado al usuario (en español). */
export class ArchivoInvalidoError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "ArchivoInvalidoError";
  }
}

/** Una fila de salida: una habitación con todos sus campos ya "aplanados". */
export interface RelevRow {
  /** Nombre de la pestaña de origen (ej. "Pintura general habt."). */
  categoria: string;
  /** Piso, derivado del número de habitación (101 → "1", 1001 → "10"). */
  piso: string;
  /** Número de habitación tal como aparece en el Excel. */
  habitacion: string;
  /** Campo → valor (ej. Detalle, Auditada, Realizado, Observación, …). */
  campos: Record<string, string>;
}

/** Resultado de despivotar una pestaña. */
export interface SheetResult {
  categoria: string;
  /** Columnas de campo detectadas, en orden (sin contar Categoría/Piso/Habitación). */
  columnas: string[];
  filas: RelevRow[];
}

/** Columnas fijas que llevan siempre las listas. */
export const COLUMNAS_BASE = ["Categoría", "Piso", "Habitación"] as const;

/** Orden preferido de las columnas de campo más comunes. El resto va detrás, alfabético. */
const ORDEN_CAMPOS = ["Detalle", "Auditada", "Realizado", "Observación"];

/**
 * ¿La celda parece un número de habitación? Aceptamos 3 dígitos (101–999) y
 * 4 dígitos para el piso 10 (1001–1099). Descarta 2 dígitos (conteos como "10").
 */
function esHabitacion(value: unknown): boolean {
  if (value === null || value === undefined) return false;
  if (typeof value === "number") {
    return Number.isInteger(value) && value >= 100 && value <= 9999 && esHabitacion(String(value));
  }
  const s = String(value).trim();
  return /^[1-9]\d{2,3}$/.test(s);
}

/** Piso a partir del número de habitación (todo menos los dos últimos dígitos). */
function pisoDe(hab: string): string {
  return hab.length > 2 ? hab.slice(0, hab.length - 2) : hab;
}

/** Formatea una fecha a "dd/mm/aaaa". */
function formatFecha(d: Date): string {
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  return `${dd}/${mm}/${d.getFullYear()}`;
}

/** Texto legible de una celda (Date → dd/mm/aaaa, resto → string recortado). */
function celdaTexto(value: unknown): string {
  if (value === null || value === undefined) return "";
  if (value instanceof Date) {
    return Number.isNaN(value.getTime()) ? "" : formatFecha(value);
  }
  return String(value).replace(/\s+/g, " ").trim();
}

/**
 * Normaliza la etiqueta de la columna A a un nombre de campo canónico.
 * Las variantes conocidas (auditada/o, realizado/a, observación…) se unifican;
 * cualquier otra etiqueta se conserva tal cual (capitalizada), sin perder datos.
 */
function normalizarEtiqueta(raw: string): string {
  const s = raw
    .toLowerCase()
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .trim();
  if (s.startsWith("detalle")) return "Detalle";
  if (s.startsWith("auditad") || s.startsWith("actualizad")) return "Auditada";
  if (s.startsWith("realizad")) return "Realizado";
  if (s.startsWith("observ")) return "Observación";
  if (s.startsWith("condicion")) return "Condición";
  if (s.startsWith("colocad")) return "Colocada";
  if (s.startsWith("cambio")) return "Cambio";
  // Etiqueta desconocida: la dejamos, capitalizando la primera letra.
  const limpio = raw.replace(/\s+/g, " ").trim();
  return limpio.charAt(0).toUpperCase() + limpio.slice(1).toLowerCase();
}

/** Agrega un valor a un campo, uniendo (sin duplicar) si ya había algo cargado. */
function agregarCampo(campos: Record<string, string>, campo: string, valor: string): void {
  if (!valor) return;
  const previo = campos[campo];
  if (!previo) {
    campos[campo] = valor;
  } else if (!previo.split(" / ").includes(valor)) {
    campos[campo] = `${previo} / ${valor}`;
  }
}

/** Todas las filas no vacías de una hoja como matriz (conserva Date). */
function hojaAMatriz(ws: XLSX.WorkSheet): unknown[][] {
  return XLSX.utils.sheet_to_json<unknown[]>(ws, {
    header: 1,
    raw: true,
    defval: "",
    blankrows: false,
  });
}

/** Índices de columna con número de habitación en una fila. */
function columnasHab(row: unknown[]): number[] {
  const cols: number[] = [];
  for (let c = 0; c < row.length; c++) if (esHabitacion(row[c])) cols.push(c);
  return cols;
}

/**
 * Despivota una hoja con formato de GRILLA horizontal: bloques de habitaciones
 * (una fila-cabecera con los números) y, debajo, filas de atributos hasta el
 * próximo bloque. La columna A puede traer la etiqueta del atributo (DETALLE,
 * AUDITADA, REALIZADO…) o no traerla (entonces el valor va a "Detalle").
 */
function despivotarGrilla(categoria: string, rows: unknown[][], headerRows: number[]): SheetResult {
  const filas: RelevRow[] = [];
  const columnasVistas = new Set<string>();

  for (let b = 0; b < headerRows.length; b++) {
    const hr = headerRows[b];
    const finBloque = b + 1 < headerRows.length ? headerRows[b + 1] : rows.length;

    // Mapa columna → habitación para este bloque.
    const mapa: { col: number; hab: string }[] = [];
    for (const c of columnasHab(rows[hr])) mapa.push({ col: c, hab: String(rows[hr][c]).trim() });
    const minCol = Math.min(...mapa.map((m) => m.col));
    // Si las habitaciones no arrancan en la col 0, la col 0 es la etiqueta.
    const labelCol = minCol >= 1 ? 0 : -1;

    // Una fila por habitación del bloque (aunque quede vacía).
    const registros = new Map<string, RelevRow>();
    for (const { hab } of mapa) {
      const fila: RelevRow = { categoria, piso: pisoDe(hab), habitacion: hab, campos: {} };
      registros.set(hab, fila);
      filas.push(fila);
    }

    // Filas de atributos del bloque.
    for (let r = hr + 1; r < finBloque; r++) {
      const row = rows[r];
      if (!row) continue;
      const etiquetaRaw = labelCol >= 0 ? celdaTexto(row[labelCol]) : "";
      const campo = etiquetaRaw ? normalizarEtiqueta(etiquetaRaw) : "Detalle";

      let algunValor = false;
      for (const { col, hab } of mapa) {
        const valor = celdaTexto(row[col]);
        if (valor) {
          agregarCampo(registros.get(hab)!.campos, campo, valor);
          algunValor = true;
        }
      }
      // Fila sin etiqueta y sin ningún valor: separador, se ignora.
      if (algunValor) columnasVistas.add(campo);
      else if (etiquetaRaw) columnasVistas.add(campo);
    }
  }

  return { categoria, columnas: ordenarColumnas(columnasVistas), filas };
}

/**
 * Despivota una hoja VERTICAL (tipo "Generales"): la habitación está en la
 * columna A y el detalle en la B. Una fila del Excel = una habitación.
 */
function despivotarVertical(categoria: string, rows: unknown[][]): SheetResult {
  const filas: RelevRow[] = [];
  for (const row of rows) {
    if (!esHabitacion(row[0])) continue;
    const hab = String(row[0]).trim();
    const detalle = celdaTexto(row[1]);
    const campos: Record<string, string> = {};
    if (detalle) campos["Detalle"] = detalle;
    filas.push({ categoria, piso: pisoDe(hab), habitacion: hab, campos });
  }
  return { categoria, columnas: ["Detalle"], filas };
}

/** Ordena las columnas de campo: primero las comunes, luego el resto alfabético. */
function ordenarColumnas(set: Set<string>): string[] {
  const cols = [...set];
  return cols.sort((a, b) => {
    const ia = ORDEN_CAMPOS.indexOf(a);
    const ib = ORDEN_CAMPOS.indexOf(b);
    if (ia !== -1 || ib !== -1) {
      return (ia === -1 ? 99 : ia) - (ib === -1 ? 99 : ib);
    }
    return a.localeCompare(b, "es");
  });
}

/** Une las columnas de campo de varias hojas en un único orden canónico. */
export function columnasCombinadas(resultados: SheetResult[]): string[] {
  const set = new Set<string>();
  for (const r of resultados) for (const c of r.columnas) set.add(c);
  return ordenarColumnas(set);
}

/** Despivota una sola hoja, eligiendo el modo (grilla u vertical) automáticamente. */
export function despivotarHoja(categoria: string, ws: XLSX.WorkSheet): SheetResult {
  const rows = hojaAMatriz(ws);
  // Fila-cabecera = fila con 3+ números de habitación (bloques horizontales).
  const headerRows: number[] = [];
  for (let i = 0; i < rows.length; i++) {
    if (columnasHab(rows[i]).length >= 3) headerRows.push(i);
  }
  if (headerRows.length > 0) return despivotarGrilla(categoria, rows, headerRows);

  // Sin bloques horizontales: ¿hay habitaciones en la columna A (vertical)?
  const habsColA = rows.filter((r) => esHabitacion(r[0])).length;
  if (habsColA >= 3) return despivotarVertical(categoria, rows);

  // No se reconoció el formato: hoja vacía de habitaciones.
  return { categoria, columnas: [], filas: [] };
}

/**
 * Despivota todo un archivo Excel: procesa cada pestaña y devuelve una lista por
 * categoría. Lanza ArchivoInvalidoError si el archivo es ilegible o no tiene
 * ninguna habitación reconocible en ninguna pestaña.
 */
export async function despivotarArchivo(file: File): Promise<SheetResult[]> {
  const buffer = await file.arrayBuffer();
  let workbook: XLSX.WorkBook;
  try {
    workbook = XLSX.read(buffer, { type: "array", cellDates: true, codepage: 65001 });
  } catch {
    throw new ArchivoInvalidoError(
      "No se pudo leer el archivo. Asegurate de que sea un Excel (.xlsx / .xls) válido."
    );
  }

  const resultados: SheetResult[] = [];
  for (const nombre of workbook.SheetNames) {
    const res = despivotarHoja(nombre.trim(), workbook.Sheets[nombre]);
    if (res.filas.length > 0) resultados.push(res);
  }

  if (resultados.length === 0) {
    throw new ArchivoInvalidoError(
      "No se reconoció ninguna habitación en el archivo. Esperaba una grilla de habitaciones (101, 102, …) por pestaña."
    );
  }
  return resultados;
}
