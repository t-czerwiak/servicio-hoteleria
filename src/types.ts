/** Metadatos de una columna del Excel, para mostrar en el mapeo. */
export interface ColumnInfo {
  /** Índice de columna (0-based). */
  index: number;
  /** Letra de columna estilo Excel ("A", "B", …). */
  letter: string;
  /** Etiqueta detectada (del encabezado) o "Columna A". */
  label: string;
  /** Hasta 3 valores de ejemplo, para que el usuario reconozca la columna. */
  samples: string[];
  /** Cuántas celdas parsean como hora. */
  timeCount: number;
  /** Cuántas celdas parecen un nombre. */
  nameCount: number;
  /** Horas distintas (alto = probable columna de hora; 1 = fecha constante). */
  distinctTimes: number;
  /** Nombres distintos (alto = probable columna de personas). */
  distinctNames: number;
}

/**
 * Asignación de cada rol a un índice de columna. Nombre y Hora son obligatorios;
 * el resto es opcional (-1 = ninguna columna).
 */
export interface Mapping {
  nameCol: number;
  timeCol: number;
  dateCol: number;
  dniCol: number;
  sedeCol: number;
  posicionCol: number;
}

/** Resultado de leer la hoja: filas crudas + metadatos + sugerencia de columnas. */
export interface ParsedSheet {
  /** Todas las filas no vacías, con sus valores crudos (incluye objetos Date). */
  rows: unknown[][];
  /** Metadatos por columna. */
  columns: ColumnInfo[];
  /** Asignación de columnas sugerida automáticamente. */
  suggested: Mapping;
  /** Cantidad de filas no vacías. */
  totalRows: number;
}

/** Un fichaje individual leído del Excel (un evento de uso de la máquina). */
export interface Punch {
  /** Nombre completo tal como vino en el archivo (recortado). */
  name: string;
  /** Minutos desde la medianoche (0–1439), usado para comparar. */
  minutes: number;
  /** DNI/legajo del empleado, o "" si no hay columna. */
  dni: string;
  /** Fecha de la jornada en formato "YYYY-MM-DD", o "" si no hay columna. */
  fecha: string;
  /** Sede desde donde se registró este fichaje, o "". */
  sede: string;
  /** Posición/puesto registrado en este fichaje, o "". */
  posicion: string;
}

/** Fila de resultado: un empleado en una jornada, con entrada, salida y total. */
export interface EmployeeRow {
  /** DNI/legajo, o "" si no hay columna. */
  dni: string;
  /** Nombre completo del empleado. */
  nombre: string;
  /** Fecha de la jornada ("YYYY-MM-DD"), o "" si no hay columna. */
  fecha: string;
  /** Posición/puesto (tomada del fichaje de entrada), o "". */
  posicion: string;
  /** Sede del fichaje de entrada, o "". */
  sedeEntrada: string;
  /** Hora de entrada (HH:MM) o cadena vacía si no corresponde. */
  entrada: string;
  /** Sede del fichaje de salida, o "" (también vacía si solo fichó una vez). */
  sedeSalida: string;
  /** Hora de salida (HH:MM) o cadena vacía si solo fichó una vez. */
  salida: string;
  /** Horas trabajadas (HH:MM) o cadena vacía si falta entrada o salida. */
  total: string;
  /** Cantidad de fichajes registrados esa jornada (informativo). */
  fichajes: number;
}

/** Definición de una columna visible del resultado (tabla y export). */
export interface ResultColumn {
  key: keyof EmployeeRow;
  label: string;
  /** Si es una columna numérica/hora (alineación tabular). */
  num?: boolean;
}
