import JSZip from "jszip";
import { COLUMNAS_BASE, columnasCombinadas } from "./unpivot";
import type { RelevRow, SheetResult } from "./unpivot";

// BOM UTF-8: hace que Excel abra los CSV con tildes/ñ correctamente. Airtable
// también lo acepta sin problemas al importar.
const BOM = "﻿";

/** Escapa un valor para CSV (comillas, comas y saltos de línea). */
function escaparCSV(valor: string): string {
  if (/[",\n\r]/.test(valor)) {
    return `"${valor.replace(/"/g, '""')}"`;
  }
  return valor;
}

/** Valor de una columna para una fila (columnas base + campos despivotados). */
function valorCelda(fila: RelevRow, columna: string): string {
  switch (columna) {
    case "Categoría":
      return fila.categoria;
    case "Piso":
      return fila.piso;
    case "Habitación":
      return fila.habitacion;
    default:
      return fila.campos[columna] ?? "";
  }
}

/** Genera el texto CSV a partir de columnas y filas. */
function construirCSV(columnas: string[], filas: RelevRow[]): string {
  const lineas = [columnas.map(escaparCSV).join(",")];
  for (const fila of filas) {
    lineas.push(columnas.map((c) => escaparCSV(valorCelda(fila, c))).join(","));
  }
  return BOM + lineas.join("\r\n");
}

/** Dispara la descarga de un blob con el nombre indicado. */
function descargar(blob: Blob, nombre: string): void {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = nombre;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

/** Nombre de archivo seguro a partir de una categoría (para los CSV del ZIP). */
function nombreArchivo(categoria: string): string {
  const limpio = categoria
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/[^a-zA-Z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .toLowerCase();
  return `${limpio || "categoria"}.csv`;
}

/** Columnas completas (base + campos) para la TABLA COMBINADA. */
export function columnasCombinadasCompletas(resultados: SheetResult[]): string[] {
  return [...COLUMNAS_BASE, ...columnasCombinadas(resultados)];
}

/** Todas las filas de todas las categorías, en orden. */
export function filasCombinadas(resultados: SheetResult[]): RelevRow[] {
  return resultados.flatMap((r) => r.filas);
}

/**
 * Descarga UN CSV con todas las categorías juntas (columna "Categoría").
 * Ideal para importar como una sola tabla en Airtable.
 */
export function descargarCombinado(resultados: SheetResult[]): void {
  const columnas = columnasCombinadasCompletas(resultados);
  const csv = construirCSV(columnas, filasCombinadas(resultados));
  descargar(new Blob([csv], { type: "text/csv;charset=utf-8" }), "relevamiento-airtable.csv");
}

/**
 * Descarga un ZIP con un CSV por categoría (pestaña), para importar cada uno
 * como su propia tabla en Airtable.
 */
export async function descargarZipPorCategoria(resultados: SheetResult[]): Promise<void> {
  const zip = new JSZip();
  const usados = new Map<string, number>();

  for (const res of resultados) {
    const columnas = [...COLUMNAS_BASE, ...res.columnas];
    const csv = construirCSV(columnas, res.filas);
    // Evita colisiones si dos categorías generan el mismo nombre de archivo.
    let nombre = nombreArchivo(res.categoria);
    const n = usados.get(nombre) ?? 0;
    usados.set(nombre, n + 1);
    if (n > 0) nombre = nombre.replace(/\.csv$/, `-${n + 1}.csv`);
    zip.file(nombre, csv);
  }

  const blob = await zip.generateAsync({ type: "blob" });
  descargar(blob, "relevamiento-airtable-por-categoria.zip");
}
