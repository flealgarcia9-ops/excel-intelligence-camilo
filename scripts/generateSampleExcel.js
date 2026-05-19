// Generates a sample Excel file using the project's exportToCustomExcel function
import { exportToCustomExcel } from '../src/utils/excelParser.js';

const data = [
  {
    'DIRECCIÓN SECCIONAL NORTE DE SANTANDER': 'Dir 1',
    'UNIDAD': 'Unidad A',
    'Nro Fiscal': 'NF1',
    'Mes': '01',
    'Año': '2014',
    'Nombre Fiscal': 'Nombre Fiscal',
    // II_ group
    'II_Mes Anterior': 1, 'II_Entran': 2, 'II_Salen': 3, 'II_Imputación': 4, 'II_Pasan': 5,
    // III_ group
    'III_Mes Anterior': 6, 'III_Entran': 7, 'III_Salen': 8, 'III_Imputación': 9, 'III_Pasan': 10,
    // IV_ group
    'IV_Mes Anterior': 11, 'IV_Entran': 12, 'IV_Salen': 13, 'IV_Imputación': 14, 'IV_Pasan': 15,
    // V_ group
    'V_Mes Anterior': 16, 'V_Entran': 17, 'V_Salen': 18, 'V_Imputación': 19, 'V_Pasan': 20,
    'Total Carga': 999,
  },
];

exportToCustomExcel(data, 'TEST_ESTATISTICA.xlsx', {
  preparedBy: 'TEST',
  preparerTitle: 'Analista',
});

console.log('Generated TEST_ESTATISTICA.xlsx');
