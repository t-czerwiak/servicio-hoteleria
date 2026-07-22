# Servicio Hotelería 🏨

Conjunto de **herramientas web para hotelería** que ayudan a pasar datos de un formato a
otro más fácil. Todo el procesamiento ocurre **en el navegador**: no se sube nada a ningún
servidor ni se guarda en ninguna base de datos.

La app agrupa las herramientas en **pestañas**:

1. **Control de Horarios** — procesa fichajes de empleados desde Excel/CSV.
2. **Convertidor de Reportes de Habitaciones → Airtable** — aplana un reporte de
   habitaciones (relevamiento, mantenimiento, inventario…) a listas importables a Airtable.

> React + TypeScript + Vite, desplegado en Firebase Hosting.

**🔗 App en vivo:** https://servicio-hoteleria.web.app
(también sigue disponible en https://control-horarios-9cf4d.web.app)

---

## 🕒 Herramienta 1 — Control de Horarios

Procesa los **fichajes de empleados** desde un Excel/CSV y devuelve, por cada empleado y
por jornada, la **hora de entrada**, la **salida**, las **horas trabajadas**, el **DNI**, la
**posición** y desde qué **sede** fichó la entrada y la salida.

- Subida de un archivo a la vez (`.xlsx`, `.xls`, `.csv`, una sola hoja, **uno o varios días**).
- **Detecta las columnas automáticamente** (nombre y hora por contenido; DNI / sede /
  posición / fecha por encabezado) y permite **corregir el mapeo** con desplegables.
- **Multi-día**: agrupa por **DNI + fecha**; la entrada y la salida pueden ser en sedes distintas.
- Exporta el resultado a Excel.

### Lógica de procesamiento

La máquina de fichajes solo registra **cuándo se usó** (nombre + hora); no distingue entrada
de salida. Por eso, para cada empleado del día:

| Situación | Entrada | Salida | Horas |
| --- | --- | --- | --- |
| Hora más temprana | ✅ la más temprana | — | — |
| Hora más tardía | — | ✅ la más tardía | — |
| Fichajes intermedios | ❌ se ignoran | ❌ se ignoran | — |
| **1 solo fichaje** | ✅ entrada | 🔲 en blanco | 🔲 en blanco |
| **2 o más fichajes** | ✅ más temprana | ✅ más tardía | `salida − entrada` |

**No se inventan ni completan datos faltantes.** Hay ejemplos ficticios en [`ejemplos/`](ejemplos/).

---

## 🏷️ Herramienta 2 — Convertidor de Reportes de Habitaciones → Airtable

Toma un Excel de **reporte de habitaciones** (una pestaña por tema: pintura, TV, colchón,
cerraduras, etc. — sea relevamiento, mantenimiento o inventario) y lo convierte en **listas
planas, una fila por habitación**, listas para importar a Airtable.

### El problema que resuelve

En el Excel de origen los datos **no** son listas: cada pestaña es una **grilla**. Las
habitaciones están en horizontal, en bloques de 9 por piso (101–109, 201–209, …), y debajo
de cada bloque hay una o varias filas de atributos (a veces etiquetadas en la columna A con
`DETALLE`, `AUDITADA`, `REALIZADO`, `OBSERVACION`…, a veces una sola fila de valor). La
pestaña `Generales` es vertical (habitación en columna A, detalle en la B).

Airtable necesita **una fila por registro**, así que la herramienta **despivota** esa grilla.

**Normalización a la grilla del hotel.** Las pestañas del Excel vienen con errores de carga
(typos en el número, habitaciones repetidas, números sueltos fuera de lugar, filas
desordenadas, o habitaciones que faltan). La herramienta conoce la grilla real del hotel —
**83 habitaciones**: pisos 1–8 con `X01`–`X09`, piso 9 con `901`–`907` y piso 10 con
`1001`–`1004`— y reconstruye cada habitación **por su posición de columna**, no por el
número escrito. Así corrige typos (ej. un `804` repetido que en realidad es `805`), rellena
las habitaciones que faltan, descarta números sueltos y evita duplicados. Resultado: **todas
las pestañas quedan con las mismas habitaciones**. (La `1002` no existe en el hotel —está por
error en la planilla— así que se excluye; el piso 10 queda con `1001`, `1003` y `1004`.)

### Cómo funciona

1. Subís el Excel (`.xlsx` / `.xls`) con una o varias pestañas.
2. Por cada pestaña, detecta el formato (grilla horizontal o vertical) y genera **una fila
   por habitación**, con las columnas base `Pestaña | Piso | Habitación` más **solo los
   campos que esa pestaña realmente usa** (Estado, Detalle, Auditada, Realizado,
   Observación… o cualquier otra etiqueta que aparezca). Como **cada pestaña tiene un
   formato distinto**, las columnas se calculan de los datos reales: **nunca hay columnas
   vacías**.
   - `Piso` se deriva del número de habitación (`101` → `1`, `1001` → `10`).
   - **`Estado`** se deriva del **color de relleno** de la celda de cada habitación en el
     Excel: verde → `Bien`, amarillo → `Más o menos`, rojo → `Mal`, gris/sin color →
     `No revisado`. Aparece solo en las pestañas que usan colores. En la vista previa se
     muestra pintado; en el CSV va como **texto** (en Airtable podés convertirlo en un campo
     "single select" con colores).
   - **`Estado N°`** es el mismo estado en número, para importar más fácil: `Bien = 5`,
     `Más o menos = 3`, `Mal = 1`, `No revisado` queda en blanco.
   - Las fechas se formatean `dd/mm/aaaa`. Los demás datos van **en crudo**, sin clasificar.
   - Se incluyen todas las habitaciones del bloque (para ver qué falta revisar).
3. **Se ve y se descarga una pestaña a la vez**: elegís la pestaña en un selector, ves su
   vista previa y descargás su CSV individual. Opcionalmente, **Descargar todas (ZIP)** trae
   un CSV por pestaña, por separado, en un solo archivo
   (`relevamiento-airtable-por-pestana.zip`).

Los CSV son **UTF-8** (con BOM, para que Excel muestre bien las tildes/ñ) y se importan
como una tabla por pestaña en Airtable.

---

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
npx firebase-tools deploy --only hosting:servicio-hoteleria
```

El sitio es 100% estático. Firebase se usa **solo para hosting**; no hay backend, ni
Firestore, ni autenticación.

## 🗂️ Estructura

```
src/
  App.tsx                     # shell con las pestañas + branding
  features/
    horarios/                 # Herramienta 1 (fichajes)
      HorariosTool.tsx, parseExcel.ts, processRecords.ts, validate.ts, exportExcel.ts
    relevamiento/             # Herramienta 2 (convertidor Airtable)
      RelevamientoTool.tsx, unpivot.ts, exportAirtable.ts
```

## 🛠️ Stack

- React 18 + TypeScript
- Vite
- [SheetJS (`xlsx`)](https://sheetjs.com/) para leer Excel/CSV
- [JSZip](https://stuk.github.io/jszip/) para el ZIP por categoría
- Firebase Hosting
