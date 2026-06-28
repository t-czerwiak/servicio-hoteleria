import type { Punch, EmployeeRow } from "../types";
import { formatMinutes } from "./parseExcel";

interface Group {
  dni: string;
  nombre: string;
  fecha: string;
  count: number;
  /** Fichaje más temprano (entrada) y más tardío (salida) de la jornada. */
  entrada: Punch;
  salida: Punch;
}

/**
 * Procesa los fichajes según la lógica del sistema de la máquina:
 *  - Se agrupa por empleado y por JORNADA (DNI + fecha; si no hay DNI, por nombre;
 *    si no hay fecha, todo se trata como un único día).
 *  - Hora más temprana = ENTRADA. Hora más tardía = SALIDA.
 *  - La sede/posición de la entrada salen de su fichaje; la sede de la salida, del suyo
 *    (pueden ser distintas: alguien puede empezar en una sede y terminar en otra).
 *  - Los fichajes intermedios se ignoran.
 *  - Un solo fichaje = entrada registrada, salida sin registrar (celda en blanco).
 *  - Las horas totales solo se calculan cuando hay entrada Y salida.
 *  - No se inventan ni completan datos faltantes.
 */
export function processRecords(punches: Punch[]): EmployeeRow[] {
  const groups = new Map<string, Group>();

  for (const punch of punches) {
    const nombre = punch.name.replace(/\s+/g, " ").trim();
    // Identidad: DNI si lo hay, si no el nombre normalizado. Jornada = identidad + fecha.
    const identidad = punch.dni ? `dni:${punch.dni}` : `nom:${nombre.toLocaleLowerCase("es")}`;
    const key = `${identidad}|${punch.fecha}`;

    const existing = groups.get(key);
    if (existing) {
      existing.count += 1;
      if (punch.minutes < existing.entrada.minutes) existing.entrada = punch;
      if (punch.minutes > existing.salida.minutes) existing.salida = punch;
    } else {
      groups.set(key, {
        dni: punch.dni,
        nombre,
        fecha: punch.fecha,
        count: 1,
        entrada: punch,
        salida: punch,
      });
    }
  }

  const rows: EmployeeRow[] = [];
  for (const g of groups.values()) {
    const tieneSalida = g.count >= 2;
    rows.push({
      dni: g.dni,
      nombre: g.nombre,
      fecha: g.fecha,
      posicion: g.entrada.posicion,
      sedeEntrada: g.entrada.sede,
      entrada: formatMinutes(g.entrada.minutes),
      sedeSalida: tieneSalida ? g.salida.sede : "",
      salida: tieneSalida ? formatMinutes(g.salida.minutes) : "",
      total: tieneSalida ? formatMinutes(g.salida.minutes - g.entrada.minutes) : "",
      fichajes: g.count,
    });
  }

  // Orden por nombre y, dentro del mismo empleado, por fecha.
  rows.sort((a, b) => a.nombre.localeCompare(b.nombre, "es") || a.fecha.localeCompare(b.fecha));
  return rows;
}
