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

/** Resultado de leer la hoja: filas crudas + metadatos + sugerencia de columnas. */
export interface ParsedSheet {
  /** Todas las filas no vacías, con sus valores crudos (incluye objetos Date). */
  rows: unknown[][];
  /** Metadatos por columna. */
  columns: ColumnInfo[];
  /** Columna de hora sugerida automáticamente (-1 si ninguna). */
  suggestedTimeCol: number;
  /** Columna de nombre sugerida automáticamente (-1 si ninguna). */
  suggestedNameCol: number;
  /** Cantidad de filas no vacías. */
  totalRows: number;
}

/** Un fichaje individual leído del Excel (un evento de uso de la máquina). */
export interface Punch {
  /** Nombre completo tal como vino en el archivo (recortado). */
  name: string;
  /** Minutos desde la medianoche (0–1439), usado para comparar. */
  minutes: number;
}

/** Fila de resultado: un empleado con su entrada, salida y total. */
export interface EmployeeRow {
  /** Nombre completo del empleado. */
  nombre: string;
  /** Hora de entrada (HH:MM) o cadena vacía si no corresponde. */
  entrada: string;
  /** Hora de salida (HH:MM) o cadena vacía si solo fichó una vez. */
  salida: string;
  /** Horas trabajadas (HH:MM) o cadena vacía si falta entrada o salida. */
  total: string;
  /** Cantidad de fichajes registrados ese día (informativo). */
  fichajes: number;
}
