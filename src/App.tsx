import { useCallback, useMemo, useRef, useState } from "react";
import type { DragEvent } from "react";
import { readSheet, buildPunches, ArchivoInvalidoError } from "./lib/parseExcel";
import { processRecords } from "./lib/processRecords";
import { validar } from "./lib/validate";
import { exportToExcel } from "./lib/exportExcel";
import type { ColumnInfo, ParsedSheet } from "./types";
import "./App.css";

const EXTENSIONES_VALIDAS = [".xlsx", ".xls", ".csv"];

/** Texto de una opción del desplegable de columnas. */
function opcionColumna(col: ColumnInfo): string {
  const muestra = col.samples[0] ? ` · ej: ${col.samples[0]}` : "";
  return `${col.letter} · ${col.label}${muestra}`;
}

export default function App() {
  const [sheet, setSheet] = useState<ParsedSheet | null>(null);
  const [timeCol, setTimeCol] = useState(-1);
  const [nameCol, setNameCol] = useState(-1);
  const [fileName, setFileName] = useState("");
  const [error, setError] = useState("");
  const [cargando, setCargando] = useState(false);
  const [arrastrando, setArrastrando] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  const procesarArchivo = useCallback(async (file: File) => {
    setError("");
    setSheet(null);
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
      setTimeCol(parsed.suggestedTimeCol);
      setNameCol(parsed.suggestedNameCol);
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
  };

  // Cálculo en vivo: cambia al instante si el usuario corrige las columnas.
  const { rows, avisos, hayError } = useMemo(() => {
    if (!sheet || timeCol < 0 || nameCol < 0) {
      return { rows: [], avisos: [], hayError: false };
    }
    const punches = buildPunches(sheet.rows, timeCol, nameCol);
    const rows = processRecords(punches);
    const avisos = validar(sheet, timeCol, nameCol, punches, rows);
    const hayError = avisos.some((a) => a.level === "error");
    return { rows, avisos, hayError };
  }, [sheet, timeCol, nameCol]);

  return (
    <div className="page">
      <header className="header">
        <h1>Control de Horarios</h1>
        <p className="subtitulo">
          Subí el Excel de fichajes del día y obtené la entrada, la salida y las horas
          trabajadas de cada empleado.
        </p>
      </header>

      <main>
        {!sheet && (
          <section aria-labelledby="subir-titulo" className="card">
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
                Formatos aceptados: .xlsx, .xls, .csv — una sola hoja
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
          <>
            {/* Panel de mapeo de columnas */}
            <section aria-labelledby="mapeo-titulo" className="card">
              <div className="resultado__cabecera">
                <h2 id="mapeo-titulo">Columnas detectadas</h2>
                <button type="button" className="boton boton--secundario" onClick={reiniciar}>
                  Procesar otro archivo
                </button>
              </div>
              <p className="ayuda-mapeo">
                Detecté automáticamente estas columnas en «{fileName}». Si no son las
                correctas, cambialas y la tabla se recalcula al instante.
              </p>

              <div className="mapeo">
                <div className="mapeo__campo">
                  <label htmlFor="col-nombre">Columna de Nombre</label>
                  <select
                    id="col-nombre"
                    value={nameCol}
                    onChange={(e) => setNameCol(Number(e.target.value))}
                  >
                    {sheet.columns.map((c) => (
                      <option key={c.index} value={c.index}>
                        {opcionColumna(c)}
                      </option>
                    ))}
                  </select>
                </div>
                <div className="mapeo__campo">
                  <label htmlFor="col-hora">Columna de Hora</label>
                  <select
                    id="col-hora"
                    value={timeCol}
                    onChange={(e) => setTimeCol(Number(e.target.value))}
                  >
                    {sheet.columns.map((c) => (
                      <option key={c.index} value={c.index}>
                        {opcionColumna(c)}
                      </option>
                    ))}
                  </select>
                </div>
              </div>

              {avisos.length > 0 && (
                <ul className="avisos" aria-live="polite">
                  {avisos.map((a, i) => (
                    <li key={i} className={`aviso aviso--${a.level}`}>
                      {a.level === "error" ? "⛔" : a.level === "warn" ? "⚠️" : "ℹ️"}{" "}
                      {a.message}
                    </li>
                  ))}
                </ul>
              )}
            </section>

            {/* Resultado */}
            {!hayError && (
              <section aria-labelledby="resultado-titulo" className="card">
                <div className="resultado__cabecera">
                  <h2 id="resultado-titulo">
                    Resultado <span className="contador">({rows.length} empleados)</span>
                  </h2>
                  <button
                    type="button"
                    className="boton boton--primario"
                    onClick={() => exportToExcel(rows)}
                    disabled={rows.length === 0}
                  >
                    Descargar Excel
                  </button>
                </div>

                {rows.length === 0 ? (
                  <p className="info-card">No se encontraron empleados con fichajes válidos.</p>
                ) : (
                  <div className="tabla-wrap">
                    <table className="tabla">
                      <caption className="sr-only">
                        Entrada, salida y horas trabajadas por empleado
                      </caption>
                      <thead>
                        <tr>
                          <th scope="col">Nombre completo</th>
                          <th scope="col">Entrada</th>
                          <th scope="col">Salida</th>
                          <th scope="col">Horas trabajadas</th>
                        </tr>
                      </thead>
                      <tbody>
                        {rows.map((r) => (
                          <tr key={r.nombre}>
                            <td>{r.nombre}</td>
                            <td className="num">{r.entrada || "—"}</td>
                            <td className="num">{r.salida || "—"}</td>
                            <td className="num">{r.total || "—"}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </section>
            )}
          </>
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
