import { describe, it, expect, vi } from 'vitest';
import { exportToCustomExcel, exportToExcel, detectOutliers, detectAnomalies } from '../src/utils/excelParser';
import * as XLSX from 'xlsx';

// Mock xlsx module used by exportToCustomExcel
vi.mock('xlsx', () => ({
  __esModule: true,
  default: {},
  utils: {
    aoa_to_sheet: vi.fn(() => ({})),
    json_to_sheet: vi.fn(() => ({})),
    book_new: vi.fn(() => ({})),
    book_append_sheet: vi.fn(),
  },
  writeFile: vi.fn(),
}));

describe('exportToCustomExcel', () => {
  it('generates fixed 27-column layout and calls writeFile', () => {
    const data = [
      {
        'DIRECCIÓN SECCIONAL NORTE DE SANTANDER': 'Dir 1',
        'UNIDAD': 'Unidad A',
        'Nro Fiscal': 'NF1',
        'Mes': '01',
        'Año': '2014',
        'Nombre Fiscal': 'Nombre Fiscal',
        // Group II
        'II_Mes Anterior': 1, 'II_Entran': 2, 'II_Salen': 3, 'II_Imputación': 4, 'II_Pasan': 5,
        // Group III
        'III_Mes Anterior': 6, 'III_Entran': 7, 'III_Salen': 8, 'III_Imputación': 9, 'III_Pasan': 10,
        // Group IV
        'IV_Mes Anterior': 11, 'IV_Entran': 12, 'IV_Salen': 13, 'IV_Imputación': 14, 'IV_Pasan': 15,
        // Group V
        'V_Mes Anterior': 16, 'V_Entran': 17, 'V_Salen': 18, 'V_Imputación': 19, 'V_Pasan': 20,
        'Total Carga': 999,
      },
    ];

    exportToCustomExcel(data, 'test.xlsx');
    expect(XLSX.utils.book_new).toHaveBeenCalled();
    expect(XLSX.writeFile).toHaveBeenCalledWith(expect.any(Object), 'test.xlsx');
    // Validate that the generated worksheet layout has 27 columns in header
    const aoa = XLSX.utils.aoa_to_sheet.mock.calls[0][0];
    expect(Array.isArray(aoa)).toBe(true);
    expect(aoa[0]).toHaveLength(27);
  });
});

describe('exportToExcel', () => {
  it('exports generic data with filters sheet', () => {
    const data = [
      { name: 'Alice', age: 30, score: 95 },
      { name: 'Bob', age: 25, score: 88 },
    ];
    const filters = [{ column: 'age', op: '>', value: '20' }];

    exportToExcel(data, 'generic.xlsx', filters);
    expect(XLSX.utils.json_to_sheet).toHaveBeenCalled();
    expect(XLSX.utils.book_append_sheet).toHaveBeenCalledTimes(4); // Datos, Filtros, Resumen + mock from previous test
    expect(XLSX.writeFile).toHaveBeenCalledWith(expect.any(Object), 'generic.xlsx');
  });
});

describe('detectOutliers', () => {
  it('detects outliers using IQR method', () => {
    const data = [
      { val: 10 }, { val: 12 }, { val: 11 }, { val: 13 }, { val: 10 },
      { val: 11 }, { val: 12 }, { val: 10 }, { val: 11 }, { val: 100 }, // outlier
    ];
    const result = detectOutliers(data, 'val');
    expect(result.outliers.length).toBeGreaterThan(0);
    expect(result.outliers[0].value).toBe(100);
    expect(result.summary).not.toBeNull();
    expect(result.summary.outlierCount).toBe(1);
  });

  it('returns empty for small datasets', () => {
    const data = [{ val: 1 }, { val: 2 }];
    const result = detectOutliers(data, 'val');
    expect(result.outliers).toHaveLength(0);
    expect(result.summary).toBeNull();
  });
});

describe('detectAnomalies', () => {
  it('detects anomalies across multiple columns', () => {
    const data = [
      { a: 10, b: 100 }, { a: 12, b: 110 }, { a: 11, b: 105 },
      { a: 10, b: 100 }, { a: 12, b: 110 }, { a: 11, b: 105 },
      { a: 10, b: 100 }, { a: 12, b: 110 }, { a: 11, b: 105 },
      { a: 10, b: 100 }, { a: 12, b: 110 }, { a: 1000, b: 120 }, // outlier in 'a'
    ];
    const result = detectAnomalies(data, ['a', 'b']);
    expect(result.length).toBeGreaterThan(0);
    const colA = result.find((r) => r.column === 'a');
    expect(colA).toBeDefined();
    expect(colA.outliers.length).toBeGreaterThan(0);
  });

  it('returns empty for clean data', () => {
    const data = [
      { a: 10 }, { a: 11 }, { a: 12 }, { a: 10 }, { a: 11 },
      { a: 12 }, { a: 10 }, { a: 11 }, { a: 12 }, { a: 10 },
    ];
    const result = detectAnomalies(data);
    expect(result).toHaveLength(0);
  });
});
