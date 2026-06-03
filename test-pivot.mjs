import fs from 'fs';
import JSZip from 'jszip';
import { createNativePivotWorkbook } from './src/utils/nativePivotExcel.js';

// Datos similares al dataset real de actuaciones
const data = [
  { 'Año': 2020, 'Mes': 1, 'Nombre de la actuación': 'ACEPTACIÓN TOTAL DE CARGOS EN AUDIENCIA CONCENTRADA', 'Número de indicados afectados por cada actuación': 2 },
  { 'Año': 2020, 'Mes': 1, 'Nombre de la actuación': 'Activo por otra causa', 'Número de indicados afectados por cada actuación': 1 },
  { 'Año': 2020, 'Mes': 2, 'Nombre de la actuación': 'ACEPTACIÓN TOTAL DE CARGOS EN AUDIENCIA CONCENTRADA', 'Número de indicados afectados por cada actuación': 3 },
  { 'Año': 2020, 'Mes': 3, 'Nombre de la actuación': 'Archivo por conducta atípica', 'Número de indicados afectados por cada actuación': 1 },
  { 'Año': 2021, 'Mes': 1, 'Nombre de la actuación': 'ACEPTACIÓN TOTAL DE CARGOS EN AUDIENCIA CONCENTRADA', 'Número de indicados afectados por cada actuación': 5 },
  { 'Año': 2021, 'Mes': 5, 'Nombre de la actuación': 'Activo por otra causa', 'Número de indicados afectados por cada actuación': 2 },
  { 'Año': 2021, 'Mes': 5, 'Nombre de la actuación': 'Archivo por conducta atípica', 'Número de indicados afectados por cada actuación': 1 },
  { 'Año': 2022, 'Mes': 7, 'Nombre de la actuación': 'Audiencia de juicio oral', 'Número de indicados afectados por cada actuación': 4 },
  { 'Año': 2022, 'Mes': 8, 'Nombre de la actuación': 'Audiencia preparatoria', 'Número de indicados afectados por cada actuación': 2 },
  { 'Año': 2022, 'Mes': 8, 'Nombre de la actuación': 'Audiencia de juicio oral', 'Número de indicados afectados por cada actuación': 1 },
  { 'Año': 2023, 'Mes': 1, 'Nombre de la actuación': 'Audiencia preparatoria', 'Número de indicados afectados por cada actuación': 3 },
  { 'Año': 2023, 'Mes': 3, 'Nombre de la actuación': 'Audiencia de formulación de acusación', 'Número de indicados afectados por cada actuación': 2 },
  { 'Año': 2023, 'Mes': 3, 'Nombre de la actuación': 'Audiencia preparatoria', 'Número de indicados afectados por cada actuación': 1 },
  { 'Año': 2023, 'Mes': 6, 'Nombre de la actuación': 'Audiencia de juicio oral', 'Número de indicados afectados por cada actuación': 5 },
  { 'Año': 2023, 'Mes': 6, 'Nombre de la actuación': 'Audiencia preparatoria', 'Número de indicados afectados por cada actuación': 2 },
];

const structure = {
  headers: ['Año', 'Mes', 'Nombre de la actuación', 'Número de indicados afectados por cada actuación'],
  dimensions: { year: 'Año', month: 'Mes' },
};

const buf = await createNativePivotWorkbook(data, structure);
fs.writeFileSync('/tmp/test-pivot.xlsx', Buffer.from(buf));

// Extraer y mostrar XMLs clave
const zip = await JSZip.loadAsync(buf);
const pivotXml = await zip.file('xl/pivotTables/pivotTable1.xml').async('string');
const cacheDef = await zip.file('xl/pivotCache/pivotCacheDefinition1.xml').async('string');
const cacheRec = await zip.file('xl/pivotCache/pivotCacheRecords1.xml').async('string');

console.log('=== PIVOT TABLE XML ===');
console.log(pivotXml);
console.log('\n=== CACHE DEFINITION (primeros 2000 chars) ===');
console.log(cacheDef.slice(0, 2000));
console.log('\n=== CACHE RECORDS (primeros 2000 chars) ===');
console.log(cacheRec.slice(0, 2000));
