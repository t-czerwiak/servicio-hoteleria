// Genera archivos de PRUEBA DE RECHAZO:
//  1) un .xlsx con extensión válida pero contenido que NO son fichajes (sin horarios)
//  2) (el .txt incompatible se crea aparte, es solo texto)
//   node scripts/gen-invalidos.mjs
import * as XLSX from "xlsx";

// Una planilla de inventario: tiene nombres (productos) pero NINGUNA hora -> se rechaza.
const data = [
  ["Producto", "Categoría", "Stock", "Precio"],
  ["Detergente", "Limpieza", 120, 850.5],
  ["Toallas", "Blanco", 60, 3200],
  ["Café molido", "Insumos", 45, 5400],
  ["Jabón líquido", "Limpieza", 200, 990],
  ["Sábanas", "Blanco", 80, 7800],
  ["Azúcar", "Insumos", 150, 1200],
];

const ws = XLSX.utils.aoa_to_sheet(data);
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, "Inventario");
const out = "ejemplos/invalido-sin-fichajes.xlsx";
XLSX.writeFile(wb, out);
console.log(`OK -> ${out} (inventario, sin columna de horarios)`);
