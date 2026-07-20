import { useState } from "react";
import HorariosTool from "./features/horarios/HorariosTool";
import RelevamientoTool from "./features/relevamiento/RelevamientoTool";
import "./App.css";

type Herramienta = "horarios" | "relevamiento";

const TABS: { id: Herramienta; label: string }[] = [
  { id: "horarios", label: "Control de Horarios" },
  { id: "relevamiento", label: "Convertidor Relevamiento → Airtable" },
];

/**
 * Shell de "Servicio Hotelería": agrupa las herramientas internas en pestañas.
 * Cada herramienta vive en su propia carpeta bajo src/features/ y se entiende sola.
 */
export default function App() {
  const [activa, setActiva] = useState<Herramienta>("horarios");

  return (
    <div className="page">
      <header className="header">
        <h1>Servicio Hotelería</h1>
        <p className="subtitulo">
          Herramientas para pasar datos de hotelería más fácil: procesar fichajes y
          convertir relevamientos a listas para Airtable. Todo en tu navegador.
        </p>
      </header>

      <nav className="tabs" aria-label="Herramientas">
        {TABS.map((t) => (
          <button
            key={t.id}
            type="button"
            className={`tab${activa === t.id ? " tab--activa" : ""}`}
            aria-current={activa === t.id ? "page" : undefined}
            onClick={() => setActiva(t.id)}
          >
            {t.label}
          </button>
        ))}
      </nav>

      <main>{activa === "horarios" ? <HorariosTool /> : <RelevamientoTool />}</main>
    </div>
  );
}
