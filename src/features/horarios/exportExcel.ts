import * as XLSX from "xlsx";
import type { EmployeeRow, ResultColumn } from "./types";

/**
 * Genera y descarga un Excel con el resultado procesado.
 * Las columnas son las mismas que se muestran en pantalla (adaptativas según el
 * archivo de origen): DNI, Fecha, Posición y Sedes aparecen solo si se detectaron.
 */
export function exportToExcel(
  rows: EmployeeRow[],
  columns: ResultColumn[],
  fileName = "horarios-procesados.xlsx"
): void {
  const data = [
    columns.map((c) => c.label),
    ...rows.map((r) => columns.map((c) => String(r[c.key] ?? ""))),
  ];

  const sheet = XLSX.utils.aoa_to_sheet(data);
  sheet["!cols"] = columns.map((c) => ({ wch: Math.max(c.label.length + 2, 14) }));

  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, sheet, "Horarios");
  XLSX.writeFile(book, fileName);
}
