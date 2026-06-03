import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { createNativePivotWorkbook } from '../src/utils/nativePivotExcel';

describe('createNativePivotWorkbook', () => {
  it('injects native Excel pivot table and cache parts', async () => {
    const data = [
      { 'Año': 2020, 'Mes': 1, 'Nombre de la actuación': 'Actuación A', 'Etapa del caso': 'Activa' },
      { 'Año': 2020, 'Mes': 2, 'Nombre de la actuación': 'Actuación B', 'Etapa del caso': 'Activa' },
      { 'Año': 2021, 'Mes': 1, 'Nombre de la actuación': 'Actuación A', 'Etapa del caso': 'Cerrada' },
    ];
    const structure = {
      headers: ['Año', 'Mes', 'Nombre de la actuación', 'Etapa del caso'],
      dimensions: { year: 'Año', month: 'Mes' },
    };

    const workbook = await createNativePivotWorkbook(data, structure);
    const zip = await JSZip.loadAsync(workbook);

    expect(zip.file('xl/pivotTables/pivotTable1.xml')).toBeTruthy();
    expect(zip.file('xl/pivotCache/pivotCacheDefinition1.xml')).toBeTruthy();
    expect(zip.file('xl/pivotCache/pivotCacheRecords1.xml')).toBeTruthy();

    const workbookXml = await zip.file('xl/workbook.xml').async('string');
    const sheetXml = await zip.file('xl/worksheets/sheet1.xml').async('string');
    const pivotXml = await zip.file('xl/pivotTables/pivotTable1.xml').async('string');

    expect(workbookXml).toContain('<pivotCaches>');
    expect(workbookXml).toContain('state="hidden"');
    expect(sheetXml).toContain('pivotTableDefinitions');
    expect(pivotXml).toContain('Cuenta de Nombre de la actuación');
    expect(pivotXml).toContain('<rowFields count="2">');
    expect(pivotXml).toContain('<colFields count="1">');
  });
});
