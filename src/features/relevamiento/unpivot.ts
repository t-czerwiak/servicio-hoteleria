import * as XLSX from "xlsx";

/** Error con mensaje claro y orientado al usuario (en español). */
export class ArchivoInvalidoError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "ArchivoInvalidoError";
  }
}

/** Estado de una habitación, derivado del COLOR de relleno de su celda en el Excel. */
export type Estado = "Bien" | "Más o menos" | "Mal" | "No revisado";

/** Una fila de salida: una habitación con todos sus campos ya "aplanados". */
export interface RelevRow {
  /** Nombre de la pestaña de origen (ej. "Pintura general habt."). */
  pestana: string;
  /** Piso, derivado del número de habitación (101 → "1", 1001 → "10"). */
  piso: string;
  /** Número de habitación tal como aparece en el Excel. */
  habitacion: string;
  /** Campo → valor (ej. Estado, Detalle, Auditada, Realizado, Observación, …). */
  campos: Record<string, string>;
}

/** Resultado de despivotar una pestaña. */
export interface SheetResult {
  /** Nombre de la pestaña de origen. */
  pestana: string;
  /**
   * Columnas de campo que esta pestaña realmente tiene con datos, en orden.
   * Se calculan de los valores reales: nunca incluye columnas vacías.
   */
  columnas: string[];
  filas: RelevRow[];
}

/** Columnas fijas que llevan siempre las listas. */
export const COLUMNAS_BASE = ["Pestaña", "Piso", "Habitación"] as const;

/** Orden preferido de las columnas de campo. Estado primero; el resto detrás, alfabético. */
const ORDEN_CAMPOS = ["Estado", "Detalle", "Auditada", "Realizado", "Observación"];

/**
 * Clasifica el color de relleno de una celda en un estado.
 * Verde = Bien, amarillo = Más o menos, rojo = Mal, gris/sin color = No revisado.
 * Se clasifica por canales RGB (no por un hex exacto) para tolerar variantes de tono.
 */
function estadoDeFill(estilo: unknown): Estado {
  const s = estilo as { patternType?: string; fgColor?: { rgb?: string } } | undefined;
  const rgb = s?.patternType === "solid" ? s?.fgColor?.rgb : undefined;
  if (!rgb) return "No revisado";
  const hex = String(rgb).slice(-6);
  const r = parseInt(hex.slice(0, 2), 16);
  const g = parseInt(hex.slice(2, 4), 16);
  const b = parseInt(hex.slice(4, 6), 16);
  if (Number.isNaN(r) || Number.isNaN(g) || Number.isNaN(b)) return "No revisado";
  // Blanco/gris/negro (canales parejos): sin estado.
  if (Math.max(r, g, b) - Math.min(r, g, b) < 40) return "No revisado";
  if (r > 150 && g < 110 && b < 110) return "Mal"; // rojo
  if (g > 110 && r < 160 && b < 150) return "Bien"; // verde
  if (r > 180 && g > 150) return "Más o menos"; // amarillo / ámbar
  return "No revisado";
}

/** Piso a partir del número de habitación (todo menos los dos últimos dígitos). */
function pisoDe(hab: string): string {
  return hab.length > 2 ? hab.slice(0, hab.length - 2) : hab;
}

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

/**
 * Lee una hoja como dos matrices alineadas a la grilla real (misma fila/columna):
 * `valores` con el contenido de cada celda y `estados` con el estado (color) de cada
 * celda. Mantener los índices reales permite mirar el color de la celda de cada
 * habitación (que es donde el Excel marca verde/amarillo/rojo).
 */
function leerHoja(ws: XLSX.WorkSheet): { valores: unknown[][]; estados: Estado[][] } {
  const ref = ws["!ref"];
  if (!ref) return { valores: [], estados: [] };
  const range = XLSX.utils.decode_range(ref);
  const valores: unknown[][] = [];
  const estados: Estado[][] = [];
  for (let r = range.s.r; r <= range.e.r; r++) {
    const filaV: unknown[] = [];
    const filaE: Estado[] = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      filaV.push(cell ? cell.v : "");
      filaE.push(cell ? estadoDeFill(cell.s) : "No revisado");
    }
    valores.push(filaV);
    estados.push(filaE);
  }
  return { valores, estados };
}

/** Índices de columna con número de habitación en una fila de valores. */
function columnasHab(row: unknown[]): number[] {
  const cols: number[] = [];
  for (let c = 0; c < row.length; c++) if (esHabitacion(row[c])) cols.push(c);
  return cols;
}

/** Ordena un conjunto de columnas: primero las comunes, luego el resto alfabético. */
function ordenarColumnas(set: Set<string>): string[] {
  return [...set].sort((a, b) => {
    const ia = ORDEN_CAMPOS.indexOf(a);
    const ib = ORDEN_CAMPOS.indexOf(b);
    if (ia !== -1 || ib !== -1) {
      return (ia === -1 ? 99 : ia) - (ib === -1 ? 99 : ib);
    }
    return a.localeCompare(b, "es");
  });
}

/**
 * Columnas de campo que aparecen con datos en las filas. Como `campos` solo
 * contiene valores no vacíos, esto garantiza que NO haya columnas vacías: cada
 * pestaña muestra únicamente las columnas que su formato realmente usó.
 */
function columnasDeFilas(filas: RelevRow[]): string[] {
  const set = new Set<string>();
  for (const f of filas) for (const k of Object.keys(f.campos)) set.add(k);
  return ordenarColumnas(set);
}

/**
 * La columna "Estado" solo tiene sentido si la pestaña usa colores. Si ninguna
 * habitación tiene color (todas "No revisado"), se quita para no ensuciar la lista.
 */
function limpiarEstadoSiNoAplica(filas: RelevRow[]): void {
  const hayColor = filas.some((f) => f.campos["Estado"] && f.campos["Estado"] !== "No revisado");
  if (!hayColor) for (const f of filas) delete f.campos["Estado"];
}

/**
 * Despivota una hoja con formato de GRILLA horizontal: bloques de habitaciones
 * (una fila-cabecera con los números) y, debajo, filas de atributos hasta el
 * próximo bloque. La columna A puede traer la etiqueta del atributo (DETALLE,
 * AUDITADA, REALIZADO…) o no traerla (entonces el valor va a "Detalle"). El estado
 * (color) se toma de la celda del número de habitación.
 */
function despivotarGrilla(
  pestana: string,
  valores: unknown[][],
  estados: Estado[][],
  headerRows: number[]
): SheetResult {
  const filas: RelevRow[] = [];

  for (let b = 0; b < headerRows.length; b++) {
    const hr = headerRows[b];
    const finBloque = b + 1 < headerRows.length ? headerRows[b + 1] : valores.length;

    // Mapa columna → habitación para este bloque.
    const mapa: { col: number; hab: string }[] = [];
    for (const c of columnasHab(valores[hr])) mapa.push({ col: c, hab: String(valores[hr][c]).trim() });
    const minCol = Math.min(...mapa.map((m) => m.col));
    // Si las habitaciones no arrancan en la col 0, la col 0 es la etiqueta.
    const labelCol = minCol >= 1 ? 0 : -1;

    // Una fila por habitación del bloque (aunque quede sin datos). El estado sale
    // del color de la propia celda del número de habitación.
    const registros = new Map<string, RelevRow>();
    for (const { col, hab } of mapa) {
      const fila: RelevRow = {
        pestana,
        piso: pisoDe(hab),
        habitacion: hab,
        campos: { Estado: estados[hr]?.[col] ?? "No revisado" },
      };
      registros.set(hab, fila);
      filas.push(fila);
    }

    // Filas de atributos del bloque.
    for (let r = hr + 1; r < finBloque; r++) {
      const row = valores[r];
      if (!row) continue;
      const etiquetaRaw = labelCol >= 0 ? celdaTexto(row[labelCol]) : "";
      const campo = etiquetaRaw ? normalizarEtiqueta(etiquetaRaw) : "Detalle";
      for (const { col, hab } of mapa) {
        agregarCampo(registros.get(hab)!.campos, campo, celdaTexto(row[col]));
      }
    }
  }

  limpiarEstadoSiNoAplica(filas);
  return { pestana, columnas: columnasDeFilas(filas), filas };
}

/**
 * Despivota una hoja VERTICAL (tipo "Generales"): la habitación está en la
 * columna A y el detalle en la B. Una fila del Excel = una habitación.
 */
function despivotarVertical(
  pestana: string,
  valores: unknown[][],
  estados: Estado[][]
): SheetResult {
  const filas: RelevRow[] = [];
  for (let r = 0; r < valores.length; r++) {
    const row = valores[r];
    if (!esHabitacion(row[0])) continue;
    const hab = String(row[0]).trim();
    const detalle = celdaTexto(row[1]);
    const campos: Record<string, string> = { Estado: estados[r]?.[0] ?? "No revisado" };
    if (detalle) campos["Detalle"] = detalle;
    filas.push({ pestana, piso: pisoDe(hab), habitacion: hab, campos });
  }
  limpiarEstadoSiNoAplica(filas);
  return { pestana, columnas: columnasDeFilas(filas), filas };
}

/** Despivota una sola hoja, eligiendo el modo (grilla o vertical) automáticamente. */
export function despivotarHoja(pestana: string, ws: XLSX.WorkSheet): SheetResult {
  const { valores, estados } = leerHoja(ws);
  // Fila-cabecera = fila con 3+ números de habitación (bloques horizontales).
  const headerRows: number[] = [];
  for (let i = 0; i < valores.length; i++) {
    if (columnasHab(valores[i]).length >= 3) headerRows.push(i);
  }
  if (headerRows.length > 0) return despivotarGrilla(pestana, valores, estados, headerRows);

  // Sin bloques horizontales: ¿hay habitaciones en la columna A (vertical)?
  const habsColA = valores.filter((r) => esHabitacion(r[0])).length;
  if (habsColA >= 3) return despivotarVertical(pestana, valores, estados);

  // No se reconoció el formato: hoja sin habitaciones.
  return { pestana, columnas: [], filas: [] };
}

/**
 * Despivota todo un archivo Excel: procesa cada pestaña y devuelve una lista por
 * pestaña. Lanza ArchivoInvalidoError si el archivo es ilegible o no tiene
 * ninguna habitación reconocible en ninguna pestaña.
 */
export async function despivotarArchivo(file: File): Promise<SheetResult[]> {
  const buffer = await file.arrayBuffer();
  let workbook: XLSX.WorkBook;
  try {
    // cellStyles: true → necesario para leer el color de relleno (el estado).
    workbook = XLSX.read(buffer, { type: "array", cellDates: true, cellStyles: true, codepage: 65001 });
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
