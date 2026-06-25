# Control de Horarios 🕒

Aplicación web que procesa los **fichajes de empleados** de un día desde un archivo
Excel/CSV y devuelve, por cada empleado, la **hora de entrada**, la **hora de salida** y
las **horas trabajadas**. Todo el procesamiento ocurre **en el navegador**: no se sube nada
a ningún servidor ni se guarda en ninguna base de datos.

> Proyecto de portfolio — React + TypeScript + Vite, desplegado en Firebase Hosting.

**🔗 App en vivo:** https://control-horarios-9cf4d.web.app

## ✨ Funcionalidades

- Subida de un archivo a la vez (`.xlsx`, `.xls`, `.csv`, una sola hoja).
- **Funciona con cualquier estructura**: detecta automáticamente las columnas de
  nombre y hora (sin importar el orden, encabezados dispersos o filas de título), y
  permite **confirmar o corregir** ese mapeo con dos desplegables. La tabla se
  recalcula al instante.
- **Validación inteligente**: rechaza archivos sin nombre+horario, y advierte cuando
  el mapeo parece incorrecto (ej. se eligió una columna de sector o de fecha).
- Procesamiento 100% en el cliente (uso libre, ilimitado, sin almacenamiento).
- Exportación del resultado a Excel.
- Interfaz accesible (etiquetas ARIA, navegación por teclado, foco visible).

### Formatos probados

- CSV/Excel simple (`Nombre`, `Hora`).
- Reportes reales de fichadas tipo ZK ("Reporte de Eventos"): decenas de columnas,
  celdas combinadas, una columna de fecha constante y columnas señuelo (Evento, Sector).
  Hay un ejemplo con datos ficticios en [`ejemplos/`](ejemplos/) (generado con
  [`scripts/gen-example.mjs`](scripts/gen-example.mjs)).

## 🧠 Lógica de procesamiento

La máquina de fichajes solo registra **cuándo se usó** (nombre + hora actual); no distingue
entrada de salida. Por eso, para cada empleado del día:

| Situación | Entrada | Salida | Horas |
| --- | --- | --- | --- |
| Hora más temprana | ✅ la más temprana | — | — |
| Hora más tardía | — | ✅ la más tardía | — |
| Fichajes intermedios | ❌ se ignoran | ❌ se ignoran | — |
| **1 solo fichaje** | ✅ entrada | 🔲 en blanco | 🔲 en blanco |
| **2 o más fichajes** | ✅ más temprana | ✅ más tardía | `salida − entrada` |

- Las horas se muestran en formato **HH:MM (24 h)**.
- **No se inventan ni completan datos faltantes.** Si falta entrada o salida, el total queda en blanco.

> ⚠️ **Turnos nocturnos (cruce de medianoche):** como los archivos son de un único día,
> un turno que entra a las 22:00 y sale a las 06:00 del día siguiente no puede resolverse
> de forma fiable con la regla más-temprana/más-tardía. Queda pendiente para una mejora futura.

## 🚀 Desarrollo

```bash
npm install
npm run dev      # servidor de desarrollo
npm run build    # build de producción en dist/
npm run preview  # previsualizar el build
```

## ☁️ Despliegue (Firebase Hosting)

```bash
npm run build
npx firebase-tools deploy
```

El sitio es 100% estático. Firebase se usa **solo para hosting**; no hay backend,
ni Firestore, ni autenticación.

## 🛠️ Stack

- React 18 + TypeScript
- Vite
- [SheetJS (`xlsx`)](https://sheetjs.com/) para leer y escribir Excel/CSV
- Firebase Hosting
