import { useCallback, useMemo, useRef, useState } from "react";
import type { DragEvent } from "react";
import { despivotarArchivo, ArchivoInvalidoError } from "./unpivot";
import type { SheetResult } from "./unpivot";
import {
  descargarCombinado,
  descargarZipPorCategoria,
  columnasCombinadasCompletas,
  filasCombinadas,
} from "./exportAirtable";

const EXTENSIONES_VALIDAS = [".xlsx", ".xls"];
const MAX_PREVIEW = 50;

/**
 * Herramienta "Convertidor Relevamiento → Airtable": toma el Excel de relevamiento
 * de habitaciones (una grilla por pestaña) y lo aplana en listas importables a
 * Airtable, una fila por habitación. Todo en el navegador.
 */
export default function RelevamientoTool() {
  const [resultados, setResultados] = useState<SheetResult[] | null>(null);
  const [fileName, setFileName] = useState("");
  const [error, setError] = useState("");
  const [cargando, setCargando] = useState(false);
  const [arrastrando, setArrastrando] = useState(false);
  const [generandoZip, setGenerandoZip] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  const procesarArchivo = useCallback(async (file: File) => {
    setError("");
    setResultados(null);
    setFileName(file.name);

    const nombre = file.name.toLowerCase();
    if (!EXTENSIONES_VALIDAS.some((ext) => nombre.endsWith(ext))) {
      setError("Formato no soportado. Subí un Excel .xlsx o .xls con las pestañas del relevamiento.");
      return;
    }

    setCargando(true);
    try {
      const res = await despivotarArchivo(file);
      setResultados(res);
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
    e.target.value = "";
  };

  const onDrop = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setArrastrando(false);
    const file = e.dataTransfer.files?.[0];
    if (file) procesarArchivo(file);
  };

  const reiniciar = () => {
    setResultados(null);
    setError("");
    setFileName("");
  };

  const descargarZip = async () => {
    if (!resultados) return;
    setGenerandoZip(true);
    try {
      await descargarZipPorCategoria(resultados);
    } finally {
      setGenerandoZip(false);
    }
  };

  // Vista previa de la tabla combinada (primeras filas).
  const preview = useMemo(() => {
    if (!resultados) return null;
    const columnas = columnasCombinadasCompletas(resultados);
    const filas = filasCombinadas(resultados);
    return { columnas, filas, total: filas.length };
  }, [resultados]);

  return (
    <>
      <p className="herramienta__intro">
        Subí el Excel del relevamiento de habitaciones (una o varias pestañas) y obtené
        listas planas, una fila por habitación, listas para importar a Airtable.
      </p>

      {!resultados && (
        <section aria-labelledby="subir-rel-titulo" className="card card--subir">
          <h2 id="subir-rel-titulo" className="sr-only">
            Subir archivo de relevamiento
          </h2>

          <div
            className={`dropzone${arrastrando ? " dropzone--activa" : ""}`}
            role="button"
            tabIndex={0}
            aria-describedby="dropzone-rel-ayuda"
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
            <p id="dropzone-rel-ayuda" className="dropzone__formatos">
              Formato aceptado: .xlsx, .xls — una o varias pestañas de habitaciones
            </p>
          </div>

          <label htmlFor="archivo-rel" className="sr-only">
            Seleccionar Excel de relevamiento
          </label>
          <input
            id="archivo-rel"
            ref={inputRef}
            type="file"
            accept=".xlsx,.xls"
            onChange={onInputChange}
            className="sr-only"
          />

          <p className="privacidad">
            🔒 Tus datos se procesan únicamente en tu navegador. No se suben a ningún
            servidor ni se guardan en ninguna base de datos.
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

      {resultados && preview && !cargando && (
        <section aria-labelledby="resultado-rel-titulo" className="card">
          <div className="resultado__cabecera">
            <h2 id="resultado-rel-titulo">
              Resultado{" "}
              <span className="contador">
                ({resultados.length} categorías · {preview.total} habitaciones)
              </span>
            </h2>
            <div className="acciones">
              <button
                type="button"
                className="boton boton--primario"
                onClick={() => descargarCombinado(resultados)}
              >
                Descargar CSV combinado
              </button>
              <button
                type="button"
                className="boton boton--secundario"
                onClick={descargarZip}
                disabled={generandoZip}
              >
                {generandoZip ? "Generando ZIP…" : "Descargar ZIP por categoría"}
              </button>
              <button type="button" className="boton boton--secundario" onClick={reiniciar}>
                Procesar otro archivo
              </button>
            </div>
          </div>

          <ul className="avisos">
            <li className="aviso aviso--info">
              El <strong>CSV combinado</strong> es una sola tabla con la columna
              «Categoría»: importalo como una tabla en Airtable. El{" "}
              <strong>ZIP por categoría</strong> trae un CSV por pestaña, para importar
              cada categoría como su propia tabla.
            </li>
          </ul>

          <p className="deteccion">
            Categorías detectadas:{" "}
            <strong>{resultados.map((r) => r.categoria).join(", ")}</strong>.
          </p>

          <div className="tabla-wrap">
            <table className="tabla">
              <caption className="sr-only">
                Vista previa de la lista combinada, una fila por habitación
              </caption>
              <thead>
                <tr>
                  {preview.columnas.map((c) => (
                    <th key={c} scope="col">
                      {c}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {preview.filas.slice(0, MAX_PREVIEW).map((fila, i) => (
                  <tr key={`${fila.categoria}|${fila.habitacion}|${i}`}>
                    {preview.columnas.map((c) => {
                      const valor =
                        c === "Categoría"
                          ? fila.categoria
                          : c === "Piso"
                          ? fila.piso
                          : c === "Habitación"
                          ? fila.habitacion
                          : fila.campos[c] ?? "";
                      return (
                        <td key={c} className={c === "Piso" || c === "Habitación" ? "num" : undefined}>
                          {valor || "—"}
                        </td>
                      );
                    })}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {preview.total > MAX_PREVIEW && (
            <p className="info-card">
              Mostrando las primeras {MAX_PREVIEW} de {preview.total} filas. La descarga
              incluye todas.
            </p>
          )}
        </section>
      )}
    </>
  );
}
