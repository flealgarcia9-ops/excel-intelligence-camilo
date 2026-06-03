import JSZip from 'jszip';
import * as XLSX from 'xlsx';
import { buildPivotTable } from './excelParser.js';

const xmlEscape = (value) => String(value ?? '')
  .replace(/&/g, '&amp;')
  .replace(/</g, '&lt;')
  .replace(/>/g, '&gt;')
  .replace(/"/g, '&quot;')
  .replace(/'/g, '&apos;');

const isUsableNumber = (value) => {
  if (value === null || value === undefined || value === '') return false;
  const n = Number(String(value).replace(',', '.'));
  return Number.isFinite(n);
};

const findFirstHeader = (headers, patterns) => (
  headers.find((header) => patterns.some((pattern) => pattern.test(String(header))))
);

const getUniqueValues = (data, header) => {
  const seen = new Map();
  for (const row of data) {
    const value = row[header] ?? '(vacío)';
    const key = String(value);
    if (!seen.has(key)) seen.set(key, value);
  }
  return [...seen.values()].sort((a, b) => {
    const na = Number(a), nb = Number(b);
    if (Number.isFinite(na) && Number.isFinite(nb)) return na - nb;
    return String(a).localeCompare(String(b));
  });
};

export const inferNativePivotFields = (data, structure = {}) => {
  const headers = structure?.headers || Object.keys(data?.[0] || {});
  const numericHeaders = headers.filter((header) => data.some((row) => isUsableNumber(row[header])));
  const textHeaders = headers.filter((header) => !numericHeaders.includes(header));

  const actionField = findFirstHeader(headers, [
    /nombre.*actuaci[oó]n/i,
    /actuaci[oó]n/i,
    /decisi[oó]n/i,
    /estado/i,
    /tipo/i,
    /categor/i,
    /etapa/i,
    /tr[aá]mite/i,
  ]) || structure?.dimensions?.type || structure?.dimensions?.name || textHeaders[0] || headers[0];

  const yearField = structure?.dimensions?.year || findFirstHeader(headers, [/^año$/i, /^anio$/i, /year/i]);
  const monthField = structure?.dimensions?.month || findFirstHeader(headers, [/^mes$/i, /month/i]);

  // Detectar estructura típica de actuaciones (año + mes + nombre de actuación)
  // En este caso usar SUM de un campo sintético de conteo para evitar problemas con Excel
  // cuando el campo de valor también está en las filas
  const hasActionYearMonth = actionField && yearField && monthField && /nombre.*actuaci[oó]n/i.test(actionField);
  if (hasActionYearMonth) {
    return {
      headers: [...headers, '_Conteo'],
      rowFields: [yearField, actionField],
      columnField: monthField,
      valueField: '_Conteo',
      aggFn: 'sum',
    };
  }

  const numericValueField = findFirstHeader(numericHeaders, [
    /n[uú]mero.*indicados?.*afectados?/i,
    /indicados?.*afectados?/i,
    /afectados?/i,
    /n[uú]mero/i,
    /cantidad/i,
    /total/i,
  ]);
  const valueField = numericValueField || actionField || numericHeaders[0] || headers[0];
  const aggFn = numericHeaders.includes(valueField) ? 'sum' : 'count';

  const rowFields = [yearField, actionField].filter(Boolean).filter((field, index, arr) => arr.indexOf(field) === index);
  const columnField = monthField && !rowFields.includes(monthField) ? monthField : null;

  return {
    headers,
    rowFields: rowFields.length ? rowFields : [headers[0]],
    columnField,
    valueField,
    aggFn,
  };
};

const buildVisiblePivotAoA = (data, fields) => {
  const pivot = buildPivotTable(data, fields.rowFields, fields.columnField, fields.valueField, fields.aggFn);
  if (!pivot) return [['Tabla dinámica']];

  const valueLabel = `${fields.aggFn === 'sum' ? 'Suma' : 'Cuenta'} de ${fields.valueField}`;
  const header = fields.columnField
    ? [valueLabel, ...pivot.colLabels.map(String), 'Total general']
    : [valueLabel, 'Total general'];

  const rows = pivot.rowLabels.map((label) => {
    if (!fields.columnField) return [label, pivot.rowTotals[label]];
    return [
      label,
      ...pivot.colLabels.map((colLabel) => pivot.matrix[label][colLabel] || 0),
      pivot.rowTotals[label],
    ];
  });

  const total = fields.columnField
    ? ['Total general', ...pivot.colLabels.map((colLabel) => pivot.colTotals[colLabel] || 0), pivot.grandTotal]
    : ['Total general', pivot.grandTotal];

  return [
    [],
    [],
    header,
    ['Etiquetas de fila', ...header.slice(1)],
    ...rows,
    total,
  ];
};

const makeCacheFieldXml = (data, header) => {
  const values = getUniqueValues(data, header);
  const numeric = values.length > 0 && values.every(isUsableNumber);
  const items = values.map((value) => numeric
    ? `<n v="${xmlEscape(Number(String(value).replace(',', '.')))}"/>`
    : `<s v="${xmlEscape(value)}"/>`
  ).join('');

  const attrs = numeric
    ? ` containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" count="${values.length}"`
    : ` count="${values.length}"`;

  return `<cacheField name="${xmlEscape(header)}" numFmtId="0"><sharedItems${attrs}>${items}</sharedItems></cacheField>`;
};

const makeCacheRecordsXml = (data, headers) => {
  const indexes = Object.fromEntries(headers.map((header) => [
    header,
    new Map(getUniqueValues(data, header).map((value, index) => [String(value), index])),
  ]));

  const rows = data.map((row) => {
    const cells = headers.map((header) => `<x v="${indexes[header].get(String(row[header] ?? '(vacío)')) ?? 0}"/>`).join('');
    return `<r>${cells}</r>`;
  }).join('');

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" count="${data.length}">${rows}</pivotCacheRecords>`;
};

const makeItemsXml = (count) => `<items count="${Math.max(count, 1)}">${Array.from({ length: Math.max(count, 1) }, (_, index) => `<item x="${index}"/>`).join('')}</items>`;

const makePivotTableXml = (data, fields) => {
  const fieldIndex = (name) => fields.headers.indexOf(name);
  const rowIndexes = fields.rowFields.map(fieldIndex).filter((index) => index >= 0);
  const colIndexes = fields.columnField ? [fieldIndex(fields.columnField)].filter((index) => index >= 0) : [];
  const valueIndex = Math.max(fieldIndex(fields.valueField), 0);

  const pivotFields = fields.headers.map((header, index) => {
    const uniqueCount = getUniqueValues(data, header).length;
    const isRow = rowIndexes.includes(index);
    const isCol = colIndexes.includes(index);
    const isValue = index === valueIndex;

    let attrs = 'compact="0" showAll="0"';
    if (isRow) attrs += ' axis="axisRow"';
    if (isCol) attrs += ' axis="axisCol"';
    if (isValue) attrs += ' dataField="1"';
    if ((isRow || isCol) && !isValue) attrs += ' defaultSubtotal="0"';

    if ((isRow || isCol) && !isValue) {
      return `<pivotField ${attrs}>${makeItemsXml(uniqueCount)}</pivotField>`;
    }
    return `<pivotField ${attrs}/>`;
  }).join('');

  const rowFields = rowIndexes.length ? `<rowFields count="${rowIndexes.length}">${rowIndexes.map((index) => `<field x="${index}"/>`).join('')}</rowFields>` : '';
  const colFields = colIndexes.length ? `<colFields count="${colIndexes.length}">${colIndexes.map((index) => `<field x="${index}"/>`).join('')}</colFields>` : '';
  const columns = fields.columnField ? getUniqueValues(data, fields.columnField).length + 2 : 2;
  const rows = Math.min(getUniqueValues(data, fields.rowFields[0]).length + 6, Math.max(data.length + 5, 20));
  const locationRef = XLSX.utils.encode_range({ s: { r: 2, c: 0 }, e: { r: rows, c: Math.max(columns - 1, 1) } });

  const subtotal = fields.aggFn === 'sum' ? 'sum' : 'count';
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PivotTable1" cacheId="1" dataOnRows="0" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="0" dataCaption="Valores" grandTotalCaption="Total general" showDrill="1" useAutoFormatting="1" itemPrintTitles="1" indent="0" outline="1" outlineData="1" compact="1" compactData="1"><location ref="${locationRef}" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/><pivotFields count="${fields.headers.length}">${pivotFields}</pivotFields>${rowFields}${colFields}<dataFields count="1"><dataField fld="${valueIndex}" subtotal="${subtotal}"/></dataFields><pivotTableStyleInfo name="PivotStyleMedium9" showRowHeaders="1" showColHeaders="1" showRowStripes="1" showColStripes="0" showLastColumn="1"/></pivotTableDefinition>`;
};

const appendOverride = (xml, partName, contentType) => {
  if (xml.includes(`PartName="${partName}"`)) return xml;
  return xml.replace('</Types>', `<Override PartName="${partName}" ContentType="${contentType}"/></Types>`);
};

const appendRelationship = (xml, id, type, target) => {
  if (xml.includes(`Id="${id}"`) || xml.includes(`Target="${target}"`)) return xml;
  return xml.replace('</Relationships>', `<Relationship Id="${id}" Type="${type}" Target="${target}"/></Relationships>`);
};

const injectNativePivotParts = async (xlsxArray, data, fields) => {
  const zip = await JSZip.loadAsync(xlsxArray);
  const sourceRef = XLSX.utils.encode_range({
    s: { r: 0, c: 0 },
    e: { r: data.length, c: fields.headers.length - 1 },
  });

  let contentTypes = await zip.file('[Content_Types].xml').async('string');
  contentTypes = appendOverride(contentTypes, '/xl/pivotCache/pivotCacheDefinition1.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml');
  contentTypes = appendOverride(contentTypes, '/xl/pivotCache/pivotCacheRecords1.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml');
  contentTypes = appendOverride(contentTypes, '/xl/pivotTables/pivotTable1.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml');
  zip.file('[Content_Types].xml', contentTypes);

  let workbook = await zip.file('xl/workbook.xml').async('string');
  if (!workbook.includes('<pivotCaches>')) {
    workbook = workbook.replace('</workbook>', '<pivotCaches><pivotCache cacheId="1" r:id="rIdPivotCache1"/></pivotCaches></workbook>');
  }
  zip.file('xl/workbook.xml', workbook);

  let workbookRels = await zip.file('xl/_rels/workbook.xml.rels').async('string');
  workbookRels = appendRelationship(workbookRels, 'rIdPivotCache1', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition', 'pivotCache/pivotCacheDefinition1.xml');
  zip.file('xl/_rels/workbook.xml.rels', workbookRels);

  const cacheFields = fields.headers.map((header) => makeCacheFieldXml(data, header)).join('');
  zip.file('xl/pivotCache/pivotCacheDefinition1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" refreshOnLoad="1" recordCount="${data.length}" createdVersion="6"><cacheSource type="worksheet"><worksheetSource ref="${sourceRef}" sheet="Datos"/></cacheSource><cacheFields count="${fields.headers.length}">${cacheFields}</cacheFields></pivotCacheDefinition>`);
  zip.file('xl/pivotCache/pivotCacheRecords1.xml', makeCacheRecordsXml(data, fields.headers));
  zip.file('xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels', '<?xml version="1.0" encoding="UTF-8"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords1.xml"/></Relationships>');
  zip.file('xl/pivotTables/pivotTable1.xml', makePivotTableXml(data, fields));
  zip.file('xl/pivotTables/_rels/pivotTable1.xml.rels', '<?xml version="1.0" encoding="UTF-8"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition1.xml"/></Relationships>');

  let sheet = await zip.file('xl/worksheets/sheet1.xml').async('string');
  if (!sheet.includes('pivotTableDefinitions')) {
    sheet = sheet.replace('</worksheet>', '<pivotTableDefinitions count="1"><pivotTableDefinition r:id="rIdPivotTable1"/></pivotTableDefinitions></worksheet>');
  }
  zip.file('xl/worksheets/sheet1.xml', sheet);

  const sheetRelsPath = 'xl/worksheets/_rels/sheet1.xml.rels';
  const existingSheetRels = zip.file(sheetRelsPath)
    ? await zip.file(sheetRelsPath).async('string')
    : '<?xml version="1.0" encoding="UTF-8"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  zip.file(sheetRelsPath, appendRelationship(existingSheetRels, 'rIdPivotTable1', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable', '../pivotTables/pivotTable1.xml'));

  return zip.generateAsync({ type: 'arraybuffer', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
};

export const createNativePivotWorkbook = async (data, structure = {}) => {
  if (!data?.length) return null;

  const fields = inferNativePivotFields(data, structure);

  // Convertir valores numéricos-string a números reales para que XLSX los escriba como números
  // Añadir campo sintético de conteo (1 por fila) para evitar problemas de Excel con COUNT de campos de texto
  const sourceData = data.map((row) => {
    const record = Object.fromEntries(fields.headers.filter((h) => h !== '_Conteo').map((header) => {
      const val = row[header];
      if (val === null || val === undefined || val === '') return [header, ''];
      const num = Number(String(val).replace(',', '.'));
      return [header, Number.isFinite(num) ? num : val];
    }));
    if (fields.headers.includes('_Conteo')) {
      record._Conteo = 1;
    }
    return record;
  });
  const sourceWs = XLSX.utils.json_to_sheet(sourceData, { header: fields.headers });

  // Hoja de tabla dinámica: vacía, Excel la renderiza desde el XML nativo
  const pivotWs = XLSX.utils.aoa_to_sheet([]);

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, pivotWs, 'Tabla dinámica');
  XLSX.utils.book_append_sheet(wb, sourceWs, 'Datos');
  wb.Workbook = { Sheets: [{ Hidden: 0 }, { Hidden: 1 }] };

  const base = XLSX.write(wb, { bookType: 'xlsx', type: 'array', compression: true });
  return injectNativePivotParts(base, sourceData, fields);
};

export const exportNativePivotExcel = async (data, filename = `tabla_dinamica_${Date.now()}.xlsx`, structure = {}) => {
  const withPivot = await createNativePivotWorkbook(data, structure);
  if (!withPivot) return;

  const blob = new Blob([withPivot], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
};
