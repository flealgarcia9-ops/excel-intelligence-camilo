import * as XLSX from 'xlsx';
import { readFileSync } from 'fs';

const buf = readFileSync('./data.xlsx');
const wb = XLSX.read(buf);

for (const name of wb.SheetNames) {
  const ws = wb.Sheets[name];
  const json = XLSX.utils.sheet_to_json(ws, { defval: null });
  console.log(`\n=== Sheet: ${name} (${json.length} rows) ===`);
  const headers = Object.keys(json[0] || {});
  console.log('Headers:', headers.join(' | '));
  
  // Find columns that might contain text unit names
  for (const h of headers) {
    const sample = json.slice(0, 5).map(r => r[h]);
    const hasText = sample.some(v => typeof v === 'string' && v.length > 5);
    if (hasText) {
      const vals = new Set(json.map(r => r[h]).filter(v => v != null).map(String));
      console.log(`\n  TEXT COL "${h}" (${vals.size} unique):`);
      console.log(`    ${[...vals].slice(0, 30).join('\n    ')}`);
    }
  }
}
