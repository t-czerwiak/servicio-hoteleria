import { useCallback, useMemo, useRef, useState } from "react";
import type { DragEvent } from "react";
import { readSheet, buildPunches, ArchivoInvalidoError } from "./lib/parseExcel";
import { processRecords } from "./lib/processRecords";
import { validar } from "./lib/validate";
import { exportToExcel } from "./lib/exportExcel";
import type { ColumnInfo, Mapping, ParsedSheet, ResultColumn } from "./types";
import "./App.css";

const EXTENSIONES_VALIDAS = [".xlsx", ".xls", ".csv"];

const MAPPING_VACIO: Mapping = {
  nameCol: -1,
  timeCol: -1,
  dateCol: -1,
  dniCol: -1,
  sedeCol: -1,
  posicionCol: -1,
};

/** Texto de una opción del desplegable de columnas. */
function opcionColumna(col: ColumnInfo): string {
  const muestra = col.samples[0] ? ` · ej: ${col.samples[0]}` : "";
  return `${col.letter} · ${col.label}${muestra}`;
}

/** Columnas visibles del resultado, según qué campos opcionales se detectaron. */
function buildColumns(m: Mapping): ResultColumn[] {
  const cols: ResultColumn[] = [];
  if (m.dniCol >= 0) cols.push({ key: "dni", label: "DNI" });
  cols.push({ key: "nombre", label: "Nombre completo" });
  if (m.dateCol >= 0) cols.push({ key: "fecha", label: "Fecha" });
  if (m.posicionCol >= 0) cols.push({ key: "posicion", label: "Posición" });
  cols.push({ key: "entrada", label: "Entrada", num: true });
  cols.push({ key: "salida", label: "Salida", num: true });
  if (m.sedeCol >= 0) cols.push({ key: "sedeEntrada", label: "Sede entrada" });
  if (m.sedeCol >= 0) cols.push({ key: "sedeSalida", label: "Sede salida" });
  cols.push({ key: "total", label: "Horas totales", num: true });
  return cols;
}

export default function App() {
  const [sheet, setSheet] = useState<ParsedSheet | null>(null);
  const [mapping, setMapping] = useState<Mapping>(MAPPING_VACIO);
  const [fileName, setFileName] = useState("");
  const [error, setError] = useState("");
  const [cargando, setCargando] = useState(false);
  const [arrastrando, setArrastrando] = useState(false);
  const [ajustando, setAjustando] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  const procesarArchivo = useCallback(async (file: File) => {
    setError("");
    setSheet(null);
    setAjustando(false);
    setFileName(file.name);

    const nombre = file.name.toLowerCase();
    if (!EXTENSIONES_VALIDAS.some((ext) => nombre.endsWith(ext))) {
      setError("Formato no soportado. Subí un archivo .xlsx, .xls o .csv.");
      return;
    }

    setCargando(true);
    try {
      const parsed = await readSheet(file);
      setSheet(parsed);
      setMapping(parsed.suggested);
    } catch (e) {
      if (e instanceof ArchivoInvalidoError) setError(e.message);
      else {
        setError("Ocurrió un error inesperado al procesar el archivo.");
        console.error(e);
      }
    } finally {
      setCargando(false);
    }
  }, []);

  const onInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) procesarArchivo(file);
    e.target.value = ""; // permitir resubir el mismo archivo
  };

  const onDrop = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setArrastrando(false);
    const file = e.dataTransfer.files?.[0];
    if (file) procesarArchivo(file);
  };

  const reiniciar = () => {
    setSheet(null);
    setError("");
    setFileName("");
    setAjustando(false);
  };

  const setRole = (role: keyof Mapping, val: number) =>
    setMapping((m) => ({ ...m, [role]: val }));

  // Cálculo en vivo: cambia al instante si el usuario corrige las columnas.
  const { rows, avisos, hayError, columns } = useMemo(() => {
    if (!sheet || mapping.timeCol < 0 || mapping.nameCol < 0) {
      return { rows: [], avisos: [], hayError: false, columns: [] as ResultColumn[] };
    }
    const punches = buildPunches(sheet.rows, mapping);
    const rows = processRecords(punches);
    const avisos = validar(sheet, mapping, punches, rows);
    const hayError = avisos.some((a) => a.level === "error");
    return { rows, avisos, hayError, columns: buildColumns(mapping) };
  }, [sheet, mapping]);

  const labelDe = (col: number) => sheet?.columns[col]?.label ?? "—";

  // Render de un desplegable de mapeo (opcional incluye la opción "(ninguna)").
  const campo = (id: string, etiqueta: string, role: keyof Mapping, opcional: boolean) => (
    <div className="mapeo__campo">
      <label htmlFor={id}>{etiqueta}</label>
      <select id={id} value={mapping[role]} onChange={(e) => setRole(role, Number(e.target.value))}>
        {opcional && <option value={-1}>(ninguna)</option>}
        {sheet?.columns.map((c) => (
          <option key={c.index} value={c.index}>
            {opcionColumna(c)}
          </option>
        ))}
      </select>
    </div>
  );

  return (
    <div className="page">
      <header className="header">
        <h1>Control de Horarios</h1>
        <p className="subtitulo">
          Subí el Excel de fichajes y obtené, por empleado y por día, la entrada, la
          salida, las horas trabajadas y desde qué sede fichó.
        </p>
      </header>

      <main>
        {!sheet && (
          <section aria-labelledby="subir-titulo" className="card card--subir">
            <h2 id="subir-titulo" className="sr-only">
              Subir archivo
            </h2>

            <div
              className={`dropzone${arrastrando ? " dropzone--activa" : ""}`}
              role="button"
              tabIndex={0}
              aria-describedby="dropzone-ayuda"
              onClick={() => inputRef.current?.click()}
              onKeyDown={(e) => {
                if (e.key === "Enter" || e.key === " ") {
                  e.preventDefault();
                  inputRef.current?.click();
                }
              }}
              onDragOver={(e) => {
                e.preventDefault();
                setArrastrando(true);
              }}
              onDragLeave={() => setArrastrando(false)}
              onDrop={onDrop}
            >
              <svg className="dropzone__icono" viewBox="0 0 24 24" aria-hidden="true">
                <path
                  d="M12 16V4m0 0l-4 4m4-4l4 4M4 16v2a2 2 0 002 2h12a2 2 0 002-2v-2"
                  fill="none"
                  stroke="currentColor"
                  strokeWidth="2"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                />
              </svg>
              <p className="dropzone__texto">
                <strong>Hacé clic para elegir</strong> o arrastrá el archivo acá
              </p>
              <p id="dropzone-ayuda" className="dropzone__formatos">
                Formatos aceptados: .xlsx, .xls, .csv — una sola hoja · uno o varios días
              </p>
            </div>

            <label htmlFor="archivo" className="sr-only">
              Seleccionar archivo Excel de fichajes
            </label>
            <input
              id="archivo"
              ref={inputRef}
              type="file"
              accept=".xlsx,.xls,.csv"
              onChange={onInputChange}
              className="sr-only"
            />

            <p className="privacidad">
              🔒 Tus datos se procesan únicamente en tu navegador. No se suben a ningún
              servidor ni se guardan en ninguna base de datos. Podés subir todos los
              archivos que quieras.
            </p>
          </section>
        )}

        <div aria-live="polite" className="estado">
          {cargando && <p className="info">Leyendo «{fileName}»…</p>}
          {error && (
            <p className="alerta" role="alert">
              ⚠️ {error}
            </p>
          )}
        </div>

        {sheet && !cargando && (
          <section aria-labelledby="resultado-titulo" className="card">
            <div className="resultado__cabecera">
              <h2 id="resultado-titulo">
                Resultado <span className="contador">({rows.length} filas)</span>
              </h2>
              <div className="acciones">
                {!hayError && (
                  <button
                    type="button"
                    className="boton boton--primario"
                    onClick={() => exportToExcel(rows, columns)}
                    disabled={rows.length === 0}
                  >
                    Descargar Excel
                  </button>
                )}
                <button type="button" className="boton boton--secundario" onClick={reiniciar}>
                  Procesar otro archivo
                </button>
              </div>
            </div>

            {/* Solo se muestran avisos relevantes (advertencias/errores), no info. */}
            {avisos.some((a) => a.level !== "info") && (
              <ul className="avisos" aria-live="polite">
                {avisos
                  .filter((a) => a.level !== "info")
                  .map((a, i) => (
                    <li key={i} className={`aviso aviso--${a.level}`}>
                      {a.level === "error" ? "⛔" : "⚠️"} {a.message}
                    </li>
                  ))}
              </ul>
            )}

            {/* Escape hatch discreto: ajustar columnas solo si hace falta. */}
            <p className="deteccion">
              Columnas usadas — Nombre: <strong>{labelDe(mapping.nameCol)}</strong>, Hora:{" "}
              <strong>{labelDe(mapping.timeCol)}</strong>.{" "}
              <button
                type="button"
                className="link"
                aria-expanded={ajustando}
                onClick={() => setAjustando((v) => !v)}
              >
                {ajustando ? "Ocultar ajuste" : "¿Columnas incorrectas? Ajustar"}
              </button>
            </p>

            {ajustando && (
              <div className="mapeo">
                {campo("col-nombre", "Nombre *", "nameCol", false)}
                {campo("col-hora", "Hora *", "timeCol", false)}
                {campo("col-fecha", "Fecha (día)", "dateCol", true)}
                {campo("col-dni", "DNI", "dniCol", true)}
                {campo("col-sede", "Sede", "sedeCol", true)}
                {campo("col-posicion", "Posición", "posicionCol", true)}
              </div>
            )}

            {!hayError &&
              (rows.length === 0 ? (
                <p className="info-card">No se encontraron empleados con fichajes válidos.</p>
              ) : (
                <div className="tabla-wrap">
                  <table className="tabla">
                    <caption className="sr-only">
                      Entrada, salida, horas trabajadas y sede por empleado y jornada
                    </caption>
                    <thead>
                      <tr>
                        {columns.map((c) => (
                          <th key={c.key} scope="col">
                            {c.label}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {rows.map((r) => (
                        <tr key={`${r.dni}|${r.fecha}|${r.nombre}`}>
                          {columns.map((c) => (
                            <td key={c.key} className={c.num ? "num" : undefined}>
                              {String(r[c.key] ?? "") || "—"}
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              ))}
          </section>
        )}
      </main>

      <footer className="footer">
        <p>
          Procesamiento transitorio en el navegador · Hora más temprana = entrada · Hora más
          tardía = salida · Los fichajes intermedios se ignoran.
        </p>
      </footer>
    </div>
  );
}
