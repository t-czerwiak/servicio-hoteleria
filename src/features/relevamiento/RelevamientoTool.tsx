import { useCallback, useMemo, useRef, useState } from "react";
import type { DragEvent } from "react";
import { despivotarArchivo, ArchivoInvalidoError } from "./unpivot";
import type { SheetResult } from "./unpivot";
import {
  descargarPestana,
  descargarZipPorPestana,
  columnasDePestana,
  valorCelda,
} from "./exportAirtable";

const EXTENSIONES_VALIDAS = [".xlsx", ".xls"];
const MAX_PREVIEW = 60;

/** Clase de color para el valor de la columna "Estado" (verde/amarillo/rojo/gris). */
const CLASE_ESTADO: Record<string, string> = {
  Bien: "est est--bien",
  "Más o menos": "est est--mas",
  Mal: "est est--mal",
  "No revisado": "est est--no",
};

/** Clase CSS de una celda de la vista previa según su columna y valor. */
function claseCelda(columna: string, valor: string): string | undefined {
  if (columna === "Estado") return CLASE_ESTADO[valor] ?? undefined;
  if (columna === "Piso" || columna === "Habitación") return "num";
  return undefined;
}

/**
 * Herramienta "Convertidor Relevamiento → Airtable": toma el Excel de relevamiento
 * de habitaciones (una grilla por pestaña) y lo aplana en listas importables a
 * Airtable, una fila por habitación. Cada pestaña se ve y se descarga por separado,
 * con solo las columnas que esa pestaña realmente usa. Todo en el navegador.
 */
export default function RelevamientoTool() {
  const [resultados, setResultados] = useState<SheetResult[] | null>(null);
  const [seleccion, setSeleccion] = useState(0);
  const [fileName, setFileName] = useState("");
  const [error, setError] = useState("");
  const [cargando, setCargando] = useState(false);
  const [arrastrando, setArrastrando] = useState(false);
  const [generandoZip, setGenerandoZip] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  const procesarArchivo = useCallback(async (file: File) => {
    setError("");
    setResultados(null);
    setSeleccion(0);
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
    setSeleccion(0);
    setError("");
    setFileName("");
  };

  const descargarZip = async () => {
    if (!resultados) return;
    setGenerandoZip(true);
    try {
      await descargarZipPorPestana(resultados);
    } finally {
      setGenerandoZip(false);
    }
  };

  // Pestaña actualmente seleccionada y sus columnas (solo las que tienen datos).
  const actual = resultados?.[seleccion] ?? null;
  const columnas = useMemo(() => (actual ? columnasDePestana(actual) : []), [actual]);

  return (
    <>
      <p className="herramienta__intro">
        Subí el Excel del reporte de habitaciones (una o varias pestañas). Cada pestaña
        se convierte en una lista plana —una fila por habitación— y la ves y descargás por
        separado, con solo las columnas que esa pestaña realmente tiene.
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

      {resultados && actual && !cargando && (
        <section aria-labelledby="resultado-rel-titulo" className="card">
          <div className="resultado__cabecera">
            <h2 id="resultado-rel-titulo">
              {resultados.length} pestaña{resultados.length !== 1 ? "s" : ""} detectada
              {resultados.length !== 1 ? "s" : ""}
            </h2>
            <div className="acciones">
              <button
                type="button"
                className="boton boton--secundario"
                onClick={descargarZip}
                disabled={generandoZip}
              >
                {generandoZip ? "Generando ZIP…" : "Descargar todas (ZIP)"}
              </button>
              <button type="button" className="boton boton--secundario" onClick={reiniciar}>
                Procesar otro archivo
              </button>
            </div>
          </div>

          {/* Selector de pestaña: se ve y se descarga una por una. */}
          <div className="selector-pestana">
            <label htmlFor="sel-pestana">Pestaña</label>
            <select
              id="sel-pestana"
              value={seleccion}
              onChange={(e) => setSeleccion(Number(e.target.value))}
            >
              {resultados.map((r, i) => (
                <option key={r.pestana + i} value={i}>
                  {r.pestana} ({r.filas.length} hab.)
                </option>
              ))}
            </select>
            <button
              type="button"
              className="boton boton--primario"
              onClick={() => descargarPestana(actual)}
            >
              Descargar esta pestaña (CSV)
            </button>
          </div>

          <p className="deteccion">
            Columnas de <strong>{actual.pestana}</strong>: {columnas.join(", ")}.
          </p>

          {columnas.includes("Estado") && (
            <p className="leyenda-estado">
              Estado (según el color en el Excel):
              <span className="est est--bien">Bien</span>
              <span className="est est--mas">Más o menos</span>
              <span className="est est--mal">Mal</span>
              <span className="est est--no">No revisado</span>
              <span className="leyenda-estado__nota">
                — en el CSV va como texto (en Airtable podés hacerlo un campo con colores).
              </span>
            </p>
          )}

          <div className="tabla-wrap">
            <table className="tabla">
              <caption className="sr-only">
                Vista previa de la pestaña {actual.pestana}, una fila por habitación
              </caption>
              <thead>
                <tr>
                  {columnas.map((c) => (
                    <th key={c} scope="col">
                      {c}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {actual.filas.slice(0, MAX_PREVIEW).map((fila, i) => (
                  <tr key={`${fila.habitacion}|${i}`}>
                    {columnas.map((c) => {
                      const valor = valorCelda(fila, c);
                      return (
                        <td key={c} className={claseCelda(c, valor)}>
                          {valor || "—"}
                        </td>
                      );
                    })}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {actual.filas.length > MAX_PREVIEW && (
            <p className="info-card">
              Mostrando las primeras {MAX_PREVIEW} de {actual.filas.length} filas. La
              descarga incluye todas.
            </p>
          )}
        </section>
      )}
    </>
  );
}
