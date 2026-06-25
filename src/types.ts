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
