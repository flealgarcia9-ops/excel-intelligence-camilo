import * as XLSX from 'xlsx';

/* ── Helpers ──────────────────────────────────────── */
const round = (n, d = 4) => {
  if (typeof n !== 'number' || !isFinite(n)) return n;
  const f = Math.pow(10, d);
  return Math.round(n * f) / f;
};

const isValidNumber = (v) => v !== null && v !== undefined && v !== '' && !isNaN(parseFloat(v)) && isFinite(v);

/** Detecta si una columna está vacía o tiene solo valores nulos/undefined/'' */
const isEmptyColumn = (data, key) => {
  const nonEmpty = data.filter((row) => {
    const v = row[key];
    return v !== null && v !== undefined && v !== '';
  });
  return nonEmpty.length === 0;
};

/** Limpia nombres de columnas: quita __EMPTY_, normaliza */
const cleanColumnName = (name) => {
  if (!name) return '';
  // Remove SheetJS auto-generated __EMPTY_ prefixes (with or without digits)
  let cleaned = name.replace(/^__EMPTY(?:_\d+)?$/, '').replace(/^__EMPTY_/, '');
  // If still empty after cleaning, return empty (will be filtered out)
  if (!cleaned.trim()) return '';
  return cleaned.trim();
};

/**
 * Detecta la fila de encabezados en datos crudos de Excel.
 * Busca la primera fila que tenga múltiples columnas con nombres válidos.
 */
const detectHeaderRow = (rows) => {
  if (!rows || rows.length < 2) return 0;

  for (let i = 0; i < Math.min(rows.length, 30); i++) {
    const row = rows[i];
    if (!row) continue;

    const values = Array.isArray(row) ? row : Object.values(row);
    const nonNullValues = values.filter((v) => v !== null && v !== undefined && v !== '');

    // Skip rows with very few values (metadata rows)
    if (nonNullValues.length < 3) continue;

    // Check if this row looks like headers (strings, reasonable length)
    const stringValues = nonNullValues.filter((v) => typeof v === 'string');
    const avgLength = stringValues.reduce((s, v) => s + v.length, 0) / (stringValues.length || 1);

    // Headers typically have multiple columns, mostly strings, not too long
    if (nonNullValues.length >= 3 && stringValues.length >= nonNullValues.length * 0.5 && avgLength < 80) {
      return i;
    }
  }

  return 0;
};

/**
 * Parsea un archivo Excel/CSV y retorna datos por hoja.
 * Detecta automáticamente la fila de encabezados y limpia columnas vacías.
 */
export const parseExcelFile = (file) => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = e.target.result;
        const workbook = XLSX.read(data, { type: 'array', cellDates: true });
        const sheetNames = workbook.SheetNames;
        const sheetsData = {};

        sheetNames.forEach((name) => {
          const ws = workbook.Sheets[name];

          // First, get raw data with header:1 to inspect structure
          const rawRows = XLSX.utils.sheet_to_json(ws, { defval: null, header: 1 });

          if (rawRows.length === 0) {
            sheetsData[name] = [];
            return;
          }

          // Detect header row
          const headerRowIndex = detectHeaderRow(rawRows);
          const headers = rawRows[headerRowIndex] || [];

          // Clean headers
          const cleanHeaders = headers.map((h, i) => {
            if (h === null || h === undefined || h === '') return `Columna_${i + 1}`;
            const cleaned = cleanColumnName(String(h));
            return cleaned || `Columna_${i + 1}`;
          });

          // Build data rows
          const dataRows = [];
          for (let i = headerRowIndex + 1; i < rawRows.length; i++) {
            const row = rawRows[i];
            if (!row) continue;

            const obj = {};
            let hasValue = false;
            for (let j = 0; j < cleanHeaders.length; j++) {
              const key = cleanHeaders[j];
              const val = row[j] !== undefined ? row[j] : null;
              if (val !== null && val !== undefined && val !== '') {
                hasValue = true;
              }
              obj[key] = val;
            }

            if (hasValue) {
              dataRows.push(obj);
            }
          }

          // Filter out empty columns
          const validKeys = cleanHeaders.filter((k) => !isEmptyColumn(dataRows, k));

          // Final cleaned data
          const finalRows = dataRows.map((row) => {
            const out = {};
            for (const k of validKeys) {
              out[k] = row[k];
            }
            return out;
          });

          sheetsData[name] = finalRows;
        });

        resolve({ sheetNames, data: sheetsData });
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = (err) => reject(err);
    reader.readAsArrayBuffer(file);
  });
};

/**
 * Identifica columnas numéricas con muestreo robusto.
 * Ignora columnas vacías o con pocos valores numéricos.
 */
export const getNumericColumns = (data) => {
  if (!data?.length) return [];
  const sampleSize = Math.min(50, data.length);
  const sample = data.slice(0, sampleSize);
  const keys = Object.keys(sample[0] || {});

  return keys.filter((key) => {
    // Skip empty-named columns
    if (!key || key === '') return false;
    const numCount = sample.filter((row) => isValidNumber(row[key])).length;
    return numCount >= sampleSize * 0.4; // Lowered threshold
  });
};

/**
 * Identifica columnas de texto/etiquetas (categóricas).
 */
export const getLabelColumns = (data) => {
  if (!data?.length) return [];
  const numCols = new Set(getNumericColumns(data));
  return Object.keys(data[0] || {}).filter((k) => !numCols.has(k) && k && k !== '');
};

/**
 * Detecta columnas de fecha.
 */
export const getDateColumns = (data) => {
  if (!data?.length) return [];
  const keys = Object.keys(data[0] || {});
  const dateRe = /fecha|date|año|year|mes|month|dia|day/i;

  return keys.filter((key) => {
    if (!key) return false;
    if (dateRe.test(key)) return true;
    // Check if values look like dates
    const sample = data.slice(0, 20).map((r) => r[key]).filter((v) => v != null);
    if (sample.length === 0) return false;
    const dateLike = sample.filter((v) => {
      if (v instanceof Date) return true;
      if (typeof v === 'string') {
        // ISO date or common formats
        return /^\d{4}-\d{2}-\d{2}/.test(v) || /^\d{2}\/\d{2}\/\d{4}/.test(v);
      }
      return false;
    });
    return dateLike.length >= sample.length * 0.5;
  });
};

/**
 * Estadísticas completas para una columna numérica.
 * Usa sampling para datasets grandes (>100k) para evitar stack overflow.
 */
export const getColumnStats = (data, column) => {
  // For large datasets, sample to avoid performance issues
  const MAX_SAMPLE = 100000;
  let values;
  if (data.length > MAX_SAMPLE) {
    // Stratified sampling: take every Nth row
    const step = Math.ceil(data.length / MAX_SAMPLE);
    values = [];
    for (let i = 0; i < data.length; i += step) {
      const v = data[i][column];
      if (isValidNumber(v)) values.push(parseFloat(v));
    }
  } else {
    values = data
      .map((row) => row[column])
      .filter(isValidNumber)
      .map((v) => parseFloat(v));
  }

  if (values.length === 0) return null;

  const count = values.length;
  let sum = 0;
  for (let i = 0; i < count; i++) sum += values[i];
  const mean = sum / count;

  const sorted = [...values].sort((a, b) => a - b);
  const min = sorted[0];
  const max = sorted[count - 1];

  const mid = Math.floor(count / 2);
  const median = count % 2 !== 0 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;

  let variance = 0;
  for (let i = 0; i < count; i++) variance += (values[i] - mean) ** 2;
  variance /= count;
  const stdDev = Math.sqrt(variance);

  const q1 = sorted[Math.floor(count * 0.25)];
  const q3 = sorted[Math.floor(count * 0.75)];
  const iqr = q3 - q1;
  const cv = mean !== 0 ? (stdDev / Math.abs(mean)) * 100 : 0;

  let skewness = 0;
  let kurtosis = 0;
  if (stdDev > 0 && count > 3) {
    for (let i = 0; i < count; i++) {
      const z = (values[i] - mean) / stdDev;
      skewness += z ** 3;
      kurtosis += z ** 4;
    }
    skewness /= count;
    kurtosis = (kurtosis / count) - 3;
  }

  return {
    count, sum: round(sum), mean: round(mean), median: round(median),
    min: round(min), max: round(max), stdDev: round(stdDev),
    variance: round(variance), q1: round(q1), q3: round(q3),
    iqr: round(iqr), range: round(max - min), cv: round(cv),
    skewness: round(skewness), kurtosis: round(kurtosis),
  };
};

/**
 * Correlación de Pearson entre dos columnas.
 */
export const getCorrelation = (data, col1, col2) => {
  const pairs = data
    .map((row) => [parseFloat(row[col1]), parseFloat(row[col2])])
    .filter(([a, b]) => !isNaN(a) && !isNaN(b) && isFinite(a) && isFinite(b));

  if (pairs.length < 3) return null;

  const n = pairs.length;
  const mean1 = pairs.reduce((s, [a]) => s + a, 0) / n;
  const mean2 = pairs.reduce((s, [, b]) => s + b, 0) / n;

  let num = 0, den1 = 0, den2 = 0;
  for (const [a, b] of pairs) {
    const da = a - mean1;
    const db = b - mean2;
    num += da * db;
    den1 += da * da;
    den2 += db * db;
  }

  const denom = Math.sqrt(den1 * den2);
  return denom === 0 ? null : round(num / denom, 4);
};

/**
 * Genera datos de histograma.
 */
export const getHistogramData = (data, column, buckets = 10) => {
  const values = data
    .map((row) => parseFloat(row[column]))
    .filter((v) => !isNaN(v) && isFinite(v));

  if (values.length === 0) return null;

  const min = Math.min(...values);
  const max = Math.max(...values);
  const range = max - min || 1;
  const size = range / buckets;

  const counts = Array(buckets).fill(0);
  for (const v of values) {
    const idx = Math.min(Math.floor((v - min) / size), buckets - 1);
    counts[idx]++;
  }

  const labels = Array.from({ length: buckets }, (_, i) => {
    const lo = min + i * size;
    const hi = lo + size;
    return `${lo.toFixed(1)}–${hi.toFixed(1)}`;
  });

  return { labels, counts };
};

/**
 * Valores únicos de una columna.
 */
export const getUniqueValues = (data, column, limit = 100) => {
  const set = new Set();
  for (const row of data) {
    const v = row[column];
    if (v != null) set.add(String(v));
    if (set.size >= limit) break;
  }
  return [...set].sort();
};

/**
 * Análisis de calidad de datos.
 */
export const getDataQuality = (data) => {
  if (!data?.length) return null;

  const totalRows = data.length;
  const keys = Object.keys(data[0] || {});
  const totalCols = keys.length;
  const totalCells = totalRows * totalCols;

  let nullCount = 0;
  const nullsByCol = {};

  for (const key of keys) {
    nullsByCol[key] = 0;
    for (const row of data) {
      const v = row[key];
      if (v === null || v === undefined || v === '') {
        nullsByCol[key]++;
        nullCount++;
      }
    }
  }

  const completeness = totalCells > 0 ? ((totalCells - nullCount) / totalCells) * 100 : 0;

  return { totalRows, totalCols, completeness: round(completeness, 1), nullsByCol };
};

/* ═══════════════════════════════════════════════════
   AGGREGATION & GROUPING
   ═══════════════════════════════════════════════════ */

/**
 * Agrupa datos por una columna y calcula agregaciones.
 * Supports: sum, avg, count, min, max
 */
export const aggregateByDimension = (data, dimCol, metricCols, aggFn = 'sum') => {
  if (!data?.length || !dimCol) return null;

  const groups = {};
  for (const row of data) {
    const key = row[dimCol] != null ? String(row[dimCol]) : '(vacío)';
    if (!groups[key]) groups[key] = [];
    groups[key].push(row);
  }

  // Sort labels (numeric-aware)
  const labels = Object.keys(groups).sort((a, b) => {
    const na = parseFloat(a), nb = parseFloat(b);
    if (!isNaN(na) && !isNaN(nb)) return na - nb;
    return a.localeCompare(b);
  });

  const aggregated = {};
  for (const col of metricCols) {
    aggregated[col] = labels.map((label) => {
      const rows = groups[label];
      const vals = rows.map((r) => parseFloat(r[col])).filter((v) => !isNaN(v) && isFinite(v));
      if (vals.length === 0) return 0;
      if (aggFn === 'sum') return vals.reduce((a, b) => a + b, 0);
      if (aggFn === 'avg') return vals.reduce((a, b) => a + b, 0) / vals.length;
      if (aggFn === 'count') return vals.length;
      if (aggFn === 'max') return Math.max(...vals);
      if (aggFn === 'min') return Math.min(...vals);
      return vals.reduce((a, b) => a + b, 0);
    });
  }

  return { labels, aggregated, counts: labels.map((l) => groups[l].length) };
};

/**
 * Top N values for a dimension/metric pair.
 */
export const getTopN = (data, dimCol, metricCol, n = 10, agg = 'sum') => {
  const result = aggregateByDimension(data, dimCol, [metricCol], agg);
  if (!result) return null;

  const pairs = result.labels.map((l, i) => ({ label: l, value: result.aggregated[metricCol][i] }));
  pairs.sort((a, b) => b.value - a.value);
  return pairs.slice(0, n);
};

/**
 * Detecta la estructura del dataset para análisis inteligente.
 */
export const analyzeSheetStructure = (data) => {
  if (!data?.length) return null;

  const headers = Object.keys(data[0]);
  const numCols = new Set(getNumericColumns(data));
  const labelCols = getLabelColumns(data);
  const dateCols = getDateColumns(data);

  // Detect dimension columns
  const dimPatterns = {
    year:      /año|anio|year/i,
    month:     /mes|month/i,
    region:    /direccion|seccional|departamento|region|territorial/i,
    unit:      /nro.*unidad|unidad|unid|unit/i,
    name:      /nombre.*fiscal|nombre|name/i,
    type:      /tipo|type|categoria/i,
    active:    /activa|activo|active|estado/i,
    id:        /nro.*fiscal|fiscal|codigo|código/i,
  };

  const dimensions = {};
  for (const h of headers) {
    if (/^(II|III|IV|V)_/.test(h)) continue;
    for (const [key, re] of Object.entries(dimPatterns)) {
      if (re.test(h) && !dimensions[key]) {
        dimensions[key] = h;
      }
    }
  }

  // Heuristic year detection
  if (!dimensions.year) {
    const numericCols = getNumericColumns(data);
    for (const col of numericCols) {
      if (!col || typeof col !== 'string') continue;
      if (col.startsWith('II_') || col.startsWith('III_') || col.startsWith('IV_') || col.startsWith('V_')) continue;
      const sample = data.slice(0, 50).map((r) => (r ? parseFloat(r[col]) : NaN)).filter((v) => !isNaN(v));
      if (sample.length > 0 && sample.every((v) => (v >= 1970 && v <= 2100) || (v >= 0 && v <= 99))) {
        dimensions.year = col;
        break;
      }
    }
  }

  // Heuristic month detection
  if (!dimensions.month) {
    const numericCols = getNumericColumns(data);
    for (const col of numericCols) {
      if (!col || typeof col !== 'string') continue;
      if (col === dimensions.year) continue;
      const sample = data.slice(0, 50).map((r) => (r ? parseFloat(r[col]) : NaN)).filter((v) => !isNaN(v));
      if (sample.length > 0 && sample.every((v) => v >= 1 && v <= 12)) {
        dimensions.month = col;
        break;
      }
    }
  }

  // Detect column groups by prefix (Fiscalia style)
  const prefixGroups = {};
  const prefixLabels = {
    'II_':  'Indagaciones',
    'III_': 'Investigaciones',
    'IV_':  'Juicios',
    'V_':   'Querellas / Ley 600',
  };

  for (const h of headers) {
    if (!numCols.has(h)) continue;
    for (const prefix of Object.keys(prefixLabels)) {
      if (h.startsWith(prefix)) {
        if (!prefixGroups[prefix]) prefixGroups[prefix] = { label: prefixLabels[prefix], cols: [] };
        prefixGroups[prefix].cols.push(h);
        break;
      }
    }
  }

  // Detect KPI columns per group
  const kpiCols = {};
  for (const [prefix, group] of Object.entries(prefixGroups)) {
    kpiCols[prefix] = {
      label: group.label,
      entran:  group.cols.find((c) => /entran/i.test(c)) || null,
      salen:   group.cols.find((c) => /\bsalen\b/i.test(c)) || null,
      pasan:   group.cols.find((c) => /\bpasan\b$/i.test(c)) || null,
      mesAnt:  group.cols.find((c) => /mes.?anterior/i.test(c)) || null,
    };
  }

  // Simple stage columns
  const simpleStages = {};
  const stageNames = ['INDAGACIONES', 'INVESTIGACIONES', 'JUICIOS', 'QUERELLAS'];
  for (const s of stageNames) {
    if (headers.includes(s)) simpleStages[s] = s;
  }

  const hasGroups = Object.keys(prefixGroups).length > 0;
  const hasSimpleStages = Object.keys(simpleStages).length > 0;

  return {
    headers,
    dimensions,
    prefixGroups,
    kpiCols,
    simpleStages,
    hasGroups,
    hasSimpleStages,
    numericCols: [...numCols],
    labelCols,
    dateCols,
  };
};

/* ═══════════════════════════════════════════════════
   OUTLIER DETECTION
   ═══════════════════════════════════════════════════ */

/**
 * Detecta outliers en una columna numérica usando el método IQR.
 * Retorna { outliers: [{row, value, zScore}], bounds: {lower, upper} }
 */
export const detectOutliers = (data, column) => {
  if (!data?.length) return { outliers: [], bounds: { lower: 0, upper: 0 }, summary: null };
  const vals = data.map((r) => Number(r[column])).filter((v) => !isNaN(v));
  if (vals.length < 10) return { outliers: [], bounds: { lower: 0, upper: 0 }, summary: null };

  const sorted = [...vals].sort((a, b) => a - b);
  const q1 = sorted[Math.floor(sorted.length * 0.25)];
  const q3 = sorted[Math.floor(sorted.length * 0.75)];
  const iqr = q3 - q1;
  const lower = q1 - 1.5 * iqr;
  const upper = q3 + 1.5 * iqr;

  const mean = vals.reduce((a, b) => a + b, 0) / vals.length;
  const stdDev = Math.sqrt(vals.reduce((s, v) => s + Math.pow(v - mean, 2), 0) / vals.length);

  const outliers = [];
  data.forEach((row, idx) => {
    const v = Number(row[column]);
    if (isNaN(v)) return;
    if (v < lower || v > upper) {
      const zScore = stdDev > 0 ? Math.abs(v - mean) / stdDev : 0;
      outliers.push({ row: idx, value: v, zScore: round(zScore, 2) });
    }
  });

  // Sort by zScore descending
  outliers.sort((a, b) => b.zScore - a.zScore);

  return {
    outliers: outliers.slice(0, 50),
    bounds: { lower: round(lower, 2), upper: round(upper, 2) },
    summary: {
      total: vals.length,
      outlierCount: outliers.length,
      percentage: round((outliers.length / vals.length) * 100, 2),
      mean: round(mean, 2),
      stdDev: round(stdDev, 2),
      q1: round(q1, 2),
      q3: round(q3, 2),
      iqr: round(iqr, 2),
    },
  };
};

/**
 * Detecta anomalías en múltiples columnas numéricas.
 */
export const detectAnomalies = (data, columns = null) => {
  if (!data?.length) return [];
  const cols = columns || getNumericColumns(data);
  return cols.map((col) => {
    const result = detectOutliers(data, col);
    return { column: col, ...result };
  }).filter((r) => r.outliers.length > 0);
};

/* ═══════════════════════════════════════════════════
   EXPORT
   ═══════════════════════════════════════════════════ */

/**
 * Exporta datos a CSV (genérico, funciona con cualquier dataset).
 */
export const exportToCSV = (data, filename = 'datos.csv') => {
  if (!data?.length) return;

  const headers = Object.keys(data[0]);
  const bom = '\uFEFF';
  const head = headers.join(',');
  const body = data.map((r) => headers.map((h) => {
    const v = r[h];
    if (v == null) return '';
    const s = String(v);
    return s.includes(',') || s.includes('"') || s.includes('\n') ? `"${s.replace(/"/g, '""')}"` : s;
  }).join(',')).join('\n');

  const blob = new Blob([bom + head + '\n' + body], { type: 'text/csv;charset=utf-8;' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
  URL.revokeObjectURL(a.href);
};

/**
 * Exporta datos filtrados a Excel (.xlsx) con formato básico.
 * Incluye metadatos de filtros aplicados en una hoja separada.
 */
export const exportToExcel = (data, filename = 'datos.xlsx', filters = []) => {
  if (!data?.length) return;

  const headers = Object.keys(data[0]);
  const ws = XLSX.utils.json_to_sheet(data, { header: headers });

  // Auto-width columns
  const colWidths = headers.map((h) => {
    const maxData = Math.max(...data.map((r) => String(r[h] || '').length), h.length);
    return { wch: Math.min(maxData + 2, 50) };
  });
  ws['!cols'] = colWidths;

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Datos');

  // Add filters sheet if filters exist
  if (filters?.length > 0) {
    const filterData = filters.map((f) => ({
      Columna: f.column,
      Operador: f.op,
      Valor: f.value,
    }));
    const filterWs = XLSX.utils.json_to_sheet(filterData);
    filterWs['!cols'] = [{ wch: 20 }, { wch: 15 }, { wch: 20 }];
    XLSX.utils.book_append_sheet(wb, filterWs, 'Filtros Aplicados');
  }

  // Add summary sheet
  const summaryData = [
    { Métrica: 'Total de registros', Valor: data.length },
    { Métrica: 'Total de columnas', Valor: headers.length },
    { Métrica: 'Fecha de exportación', Valor: new Date().toLocaleString('es-CO') },
  ];
  const summaryWs = XLSX.utils.json_to_sheet(summaryData);
  summaryWs['!cols'] = [{ wch: 25 }, { wch: 30 }];
  XLSX.utils.book_append_sheet(wb, summaryWs, 'Resumen');

  XLSX.writeFile(wb, filename);
};

/**
 * Genera Excel estructurado (estilo Fiscalía).
 * Mantiene compatibilidad con el formato original.
 */
export const exportToCustomExcel = (data, filename, metadata = {}, filters = []) => {
  if (!data?.length) return;

  let filteredData = data;
  if (Array.isArray(filters) && filters.length > 0) {
    filteredData = data.filter((row) => {
      for (const f of filters) {
        if (!f || !f.column) continue;
        const { column, op = 'includes', value } = f;
        const v = row[column];
        const sv = v == null ? '' : String(v).toLowerCase();
        const cv = value == null ? '' : String(value).toLowerCase();

        if (op === 'includes') {
          if (!sv.includes(cv)) return false;
        } else if (op === 'equals') {
          if (sv !== cv) return false;
        } else if (op === '>') {
          const nv = parseFloat(v);
          const wv = parseFloat(value);
          if (!(isFinite(nv) && isFinite(wv) && nv > wv)) return false;
        } else if (op === '<') {
          const nv = parseFloat(v);
          const wv = parseFloat(value);
          if (!(isFinite(nv) && isFinite(wv) && nv < wv)) return false;
        } else {
          if (!sv.includes(cv)) return false;
        }
      }
      return true;
    });
  }

  const norm = (s) => s?.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase() || '';

  const dimCols = [
    { label: 'SECCIONAL', key: () => 'DIRECCIÓN SECCIONAL NORTE DE SANTANDER' },
    { label: 'UNIDAD', key: (r) => r['UNIDAD'] || r['UNIDAD FISCAL'] || Object.entries(r).find(([k]) => norm(k).includes('unidad'))?.[1] || '' },
    { label: 'Nro Fiscal', key: (r) => r['Nro Fiscal'] || r['CÓDIGO FISCAL'] || r['CODIGO FISCAL'] || Object.entries(r).find(([k]) => norm(k).includes('nro fiscal') || norm(k).includes('codigo fiscal'))?.[1] || '' },
    { label: 'Mes', key: (r) => r['Mes'] || r['MES'] || '' },
    { label: 'Año', key: (r) => r['Año'] || r['ANIO'] || '' },
    { label: 'Nombre Fiscal', key: (r) => r['Nombre Fiscal'] || r['NOMBRE FISCAL'] || Object.entries(r).find(([k]) => norm(k).includes('nombre fiscal'))?.[1] || '' },
  ];

  const groups = [
    { label: 'INDAGACIONES', prefix: 'II_', cols: ['Mes Anterior', 'Entran', 'Salen', 'Imputación', 'Pasan'] },
    { label: 'INVESTIGACIONES', prefix: 'III_', cols: ['Mes Anterior', 'Entran', 'Salen', 'Imputación', 'Pasan'] },
    { label: 'JUICIOS', prefix: 'IV_', cols: ['Mes Anterior', 'Entran', 'Salen', 'Imputación', 'Pasan'] },
    { label: 'QUERELLAS', prefix: 'V_', cols: ['Mes Anterior', 'Entran', 'Salen', 'Imputación', 'Pasan'] },
  ];

  const findCol = (row, prefix, sub) => {
    const keys = Object.keys(row);
    const target = norm(sub);
    const found = keys.find((k) => k.startsWith(prefix) && norm(k).includes(target));
    return found ? row[found] : 0;
  };

  let timeRange = '';
  try {
    const years = data.map((r) => parseInt(r['Año'] || r['ANIO'])).filter((y) => !isNaN(y));
    if (years.length) {
      const minYear = Math.min(...years) + 2000;
      const maxYear = Math.max(...years) + 2000;
      timeRange = ` ENTRE NOVIEMBRE DE ${minYear} A MAYO DE ${maxYear}`;
    }
  } catch (e) {}

  const title = metadata.reportTitle || `ESTADISTICA FISCALIA PRIMERA SECCIONAL DE OCAÑA${timeRange}`;
  const headerRow1 = Array(27).fill('');
  headerRow1[0] = title;

  const headerRow2 = [
    ...dimCols.map(() => ''),
    'INDAGACIONES', '', '', '', '',
    'INVESTIGACIONES', '', '', '', '',
    'JUICIOS', '', '', '', '',
    'QUERELLAS', '', '', '', '',
    'Total Carga',
  ];

  const headerRow3 = [
    ...dimCols.map((d) => d.label),
    ...groups.flatMap((g) => g.cols),
    'Total Carga',
  ];

  const rows = filteredData.map((r) => {
    const dims = dimCols.map((d) => d.key(r));
    const vals = groups.flatMap((g) => g.cols.map((c) => findCol(r, g.prefix, c)));
    let totalCarga = r['Total Carga'] || 0;
    if (!totalCarga) {
      totalCarga = groups.reduce((acc, g) => acc + (parseFloat(findCol(r, g.prefix, 'Pasan')) || 0), 0);
    }
    return [...dims, ...vals, totalCarga];
  });

  const aoa = [headerRow1, headerRow2, headerRow3, ...rows];
  aoa.push([]);
  aoa.push([metadata.preparedBy || 'JULIAN ROSENDO BERMON BENCARDINO']);
  aoa.push([metadata.preparerTitle || 'Asesor III - Fiscalía Seccional Norte de Santander']);
  aoa.push(['NOTA: Esta información fue consultada en su totalidad de los datos que reposan en la Dirección de estadística suministrada por los Despachos en la época.']);

  const ws = XLSX.utils.aoa_to_sheet(aoa);

  const merges = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 26 } },
    ...dimCols.map((_, i) => ({ s: { r: 1, c: i }, e: { r: 2, c: i } })),
    { s: { r: 1, c: 6 }, e: { r: 1, c: 10 } },
    { s: { r: 1, c: 11 }, e: { r: 1, c: 15 } },
    { s: { r: 1, c: 16 }, e: { r: 1, c: 20 } },
    { s: { r: 1, c: 21 }, e: { r: 1, c: 25 } },
    { s: { r: 1, c: 26 }, e: { r: 2, c: 26 } },
  ];
  ws['!merges'] = merges;

  const wscols = [
    { wch: 35 }, { wch: 15 }, { wch: 10 }, { wch: 5 }, { wch: 5 }, { wch: 30 },
    ...Array(20).fill({ wch: 10 }),
    { wch: 12 },
  ];
  ws['!cols'] = wscols;

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Estadísticas');
  XLSX.writeFile(wb, filename || `REPORTE_ESTADISTICO_${Date.now()}.xlsx`);
};
