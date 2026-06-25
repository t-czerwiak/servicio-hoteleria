import * as XLSX from "xlsx";
import type { EmployeeRow } from "../types";

/**
 * Genera y descarga un Excel con el resultado procesado.
 * Columnas: Nombre completo | Entrada | Salida | Horas trabajadas.
 */
export function exportToExcel(rows: EmployeeRow[], fileName = "horarios-procesados.xlsx"): void {
  const data = [
    ["Nombre completo", "Entrada", "Salida", "Horas trabajadas"],
    ...rows.map((r) => [r.nombre, r.entrada, r.salida, r.total]),
  ];

  const sheet = XLSX.utils.aoa_to_sheet(data);
  // Anchos de columna para que se lea cómodo.
  sheet["!cols"] = [{ wch: 32 }, { wch: 12 }, { wch: 12 }, { wch: 16 }];

  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, sheet, "Horarios");
  XLSX.writeFile(book, fileName);
}
