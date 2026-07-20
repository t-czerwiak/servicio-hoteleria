import JSZip from "jszip";
import { COLUMNAS_BASE } from "./unpivot";
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
export function valorCelda(fila: RelevRow, columna: string): string {
  switch (columna) {
    case "Pestaña":
      return fila.pestana;
    case "Piso":
      return fila.piso;
    case "Habitación":
      return fila.habitacion;
    default:
      return fila.campos[columna] ?? "";
  }
}

/** Columnas completas de una pestaña: base + solo sus campos con datos. */
export function columnasDePestana(res: SheetResult): string[] {
  return [...COLUMNAS_BASE, ...res.columnas];
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

/** Nombre de archivo seguro a partir del nombre de la pestaña. */
function nombreArchivo(pestana: string): string {
  const limpio = pestana
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/[^a-zA-Z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .toLowerCase();
  return `${limpio || "pestana"}.csv`;
}

/**
 * Descarga el CSV de UNA sola pestaña, con solo sus columnas con datos.
 * Ideal para importar esa pestaña como una tabla en Airtable.
 */
export function descargarPestana(res: SheetResult): void {
  const csv = construirCSV(columnasDePestana(res), res.filas);
  descargar(new Blob([csv], { type: "text/csv;charset=utf-8" }), nombreArchivo(res.pestana));
}

/**
 * Descarga un ZIP con un CSV por pestaña (cada uno con solo sus columnas con
 * datos), para tener todas las listas por separado en un solo archivo.
 */
export async function descargarZipPorPestana(resultados: SheetResult[]): Promise<void> {
  const zip = new JSZip();
  const usados = new Map<string, number>();

  for (const res of resultados) {
    const csv = construirCSV(columnasDePestana(res), res.filas);
    // Evita colisiones si dos pestañas generan el mismo nombre de archivo.
    let nombre = nombreArchivo(res.pestana);
    const n = usados.get(nombre) ?? 0;
    usados.set(nombre, n + 1);
    if (n > 0) nombre = nombre.replace(/\.csv$/, `-${n + 1}.csv`);
    zip.file(nombre, csv);
  }

  const blob = await zip.generateAsync({ type: "blob" });
  descargar(blob, "relevamiento-airtable-por-pestana.zip");
}
