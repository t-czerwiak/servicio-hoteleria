import type { ParsedSheet, Punch, EmployeeRow, Mapping } from "../types";
import { isNameLike } from "./parseExcel";

export interface Aviso {
  /** "error" bloquea el procesamiento; "warn" solo advierte. */
  level: "error" | "warn" | "info";
  message: string;
}

/**
 * Valida la elección de columnas y el resultado, devolviendo avisos.
 * Filosofía: bloquear solo cuando es claramente inválido; en los casos dudosos,
 * advertir y dejar que el usuario corrija el mapeo de columnas.
 */
export function validar(
  sheet: ParsedSheet,
  mapping: Mapping,
  punches: Punch[],
  rows: EmployeeRow[]
): Aviso[] {
  const { timeCol, nameCol } = mapping;
  const avisos: Aviso[] = [];

  if (timeCol === nameCol) {
    avisos.push({
      level: "error",
      message: "La columna de hora y la de nombre no pueden ser la misma.",
    });
    return avisos;
  }

  if (punches.length === 0) {
    avisos.push({
      level: "error",
      message:
        "Con estas columnas no se encontró ningún fichaje válido (nombre + hora). Revisá el mapeo de columnas.",
    });
    return avisos;
  }

  // ¿Qué porcentaje de las filas produjo un fichaje válido? Si es muy bajo en un
  // archivo grande, probablemente las columnas elegidas no son las correctas.
  const ratio = punches.length / sheet.totalRows;
  if (sheet.totalRows >= 12 && ratio < 0.4) {
    avisos.push({
      level: "warn",
      message: `Solo ${punches.length} de ${sheet.totalRows} filas dieron un fichaje válido. Si esperabas más, revisá que las columnas de Hora y Nombre sean las correctas.`,
    });
  }

  // ¿La columna de "nombre" parece de personas? Una columna de personas tiene
  // muchos valores distintos (cada quien ficha ~2 veces); columnas tipo "Sector"
  // o "Evento" repiten pocos valores. Avisamos si hay menos distintos de lo esperado.
  const nameInfo = sheet.columns[nameCol];
  const distintosEsperados = Math.max(4, Math.floor(punches.length / 4));
  if (punches.length >= 12 && nameInfo && nameInfo.distinctNames < distintosEsperados) {
    avisos.push({
      level: "warn",
      message: `La columna de nombre ("${nameInfo.label}") tiene pocos valores distintos (${nameInfo.distinctNames}) para ${punches.length} fichajes. ¿Seguro que es la columna de personas y no algo como un sector o tipo de evento?`,
    });
  }

  // ¿La columna de "nombre" es en realidad numérica (IDs)?
  const nameSampleNumeric = (nameInfo?.samples ?? []).filter(
    (s) => s !== "" && !isNameLike(s)
  ).length;
  if (nameInfo && nameInfo.samples.length > 0 && nameSampleNumeric === nameInfo.samples.length) {
    avisos.push({
      level: "warn",
      message: `La columna de nombre ("${nameInfo.label}") no parece contener nombres de personas. Revisá el mapeo.`,
    });
  }

  // ¿La columna de "hora" es constante (1 sola hora distinta)? Probablemente sea la fecha.
  const timeInfo = sheet.columns[timeCol];
  if (timeInfo && timeInfo.distinctTimes <= 1 && punches.length > 3) {
    avisos.push({
      level: "warn",
      message: `La columna de hora ("${timeInfo.label}") tiene un único valor para todas las filas. ¿Elegiste la columna de fecha por error?`,
    });
  }

  // Informativo: jornadas con par incompleto (un solo fichaje -> salida en blanco).
  const incompletos = rows.filter((r) => r.fichajes < 2).length;
  if (incompletos > 0) {
    avisos.push({
      level: "info",
      message: `${incompletos} empleado(s) con un solo fichaje: se muestra la entrada y la salida queda en blanco (no se calculan horas).`,
    });
  }

  return avisos;
}
