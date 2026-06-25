import type { Punch, EmployeeRow } from "../types";
import { formatMinutes } from "./parseExcel";

interface Group {
  nombre: string;
  min: number;
  max: number;
  count: number;
}

/**
 * Procesa los fichajes según la lógica del sistema de la máquina:
 *  - Se agrupa por nombre completo del empleado (del día).
 *  - Hora más temprana = ENTRADA. Hora más tardía = SALIDA.
 *  - Los fichajes intermedios se ignoran.
 *  - Un solo fichaje = entrada registrada, salida sin registrar (celda en blanco).
 *  - Las horas totales solo se calculan cuando hay entrada Y salida.
 *  - No se inventan ni completan datos faltantes.
 */
export function processRecords(punches: Punch[]): EmployeeRow[] {
  const groups = new Map<string, Group>();

  for (const punch of punches) {
    // Clave normalizada para agrupar; se conserva el nombre original para mostrar.
    const key = punch.name.replace(/\s+/g, " ").trim().toLocaleLowerCase("es");
    const existing = groups.get(key);
    if (existing) {
      existing.min = Math.min(existing.min, punch.minutes);
      existing.max = Math.max(existing.max, punch.minutes);
      existing.count += 1;
    } else {
      groups.set(key, {
        nombre: punch.name.replace(/\s+/g, " ").trim(),
        min: punch.minutes,
        max: punch.minutes,
        count: 1,
      });
    }
  }

  const rows: EmployeeRow[] = [];
  for (const g of groups.values()) {
    const tieneSalida = g.count >= 2;
    const entrada = formatMinutes(g.min);
    const salida = tieneSalida ? formatMinutes(g.max) : "";
    // Total solo con entrada y salida presentes. Diferencia simple del día.
    const total = tieneSalida ? formatMinutes(g.max - g.min) : "";

    rows.push({
      nombre: g.nombre,
      entrada,
      salida,
      total,
      fichajes: g.count,
    });
  }

  // Orden alfabético por nombre (español).
  rows.sort((a, b) => a.nombre.localeCompare(b.nombre, "es"));
  return rows;
}
