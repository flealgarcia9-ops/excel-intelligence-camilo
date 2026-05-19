import React, { useState, useMemo } from 'react';
import { Search, Play, X, Database, Table, Filter, Terminal } from 'lucide-react';
import { getColumnStats, getUniqueValues } from '../utils/excelParser';

const fmt = (n) => {
  if (n == null || isNaN(n)) return '—';
  if (Math.abs(n) >= 1e9) return (n / 1e9).toFixed(2) + 'B';
  if (Math.abs(n) >= 1e6) return (n / 1e6).toFixed(2) + 'M';
  if (Math.abs(n) >= 1e3) return (n / 1e3).toFixed(1) + 'K';
  return n.toLocaleString(undefined, { maximumFractionDigits: 2 });
};

export default function QueryView({ sheetData }) {
  const [queryText, setQueryText] = useState('');
  const [results, setResults] = useState(null);
  const [error, setError] = useState(null);

  const headers = useMemo(() => sheetData?.length ? Object.keys(sheetData[0]) : [], [sheetData]);
  const numericCols = useMemo(() => {
    if (!sheetData?.length) return [];
    return headers.filter((h) => {
      const vals = sheetData.map((r) => r[h]).filter((v) => v != null && v !== '');
      if (vals.length === 0) return false;
      return vals.slice(0, 20).every((v) => !isNaN(parseFloat(v)) && isFinite(v));
    });
  }, [sheetData, headers]);

  const sampleQueries = [
    { label: 'Top 10 registros', query: 'SELECT * LIMIT 10' },
    { label: 'Contar registros', query: 'SELECT COUNT(*)' },
    { label: 'Suma de columna numérica', query: `SELECT SUM(${numericCols[0] || 'valor'})` },
    { label: 'Promedio', query: `SELECT AVG(${numericCols[0] || 'valor'})` },
    { label: 'Valores únicos', query: `SELECT DISTINCT ${headers[0] || 'columna'} LIMIT 20` },
    { label: 'Filtrar por texto', query: `SELECT * WHERE ${headers[0] || 'columna'} CONTAINS "texto"` },
  ];

  const executeQuery = () => {
    setError(null);
    setResults(null);

    if (!queryText.trim()) {
      setError('Escribe una consulta');
      return;
    }

    const q = queryText.trim().toUpperCase();
    let data = [...sheetData];

    try {
      // Parse simple query language
      // SELECT [columns] [WHERE condition] [ORDER BY column] [LIMIT n]

      // WHERE clause
      const whereMatch = queryText.match(/WHERE\s+(.+?)(?:\s+ORDER\s+BY|\s+LIMIT|$)/i);
      if (whereMatch) {
        const condition = whereMatch[1].trim();
        const colMatch = condition.match(/^(.+?)\s*(=|!=|<>|>=|<=|>||<|CONTAINS)\s*(.+)$/i);
        if (colMatch) {
          const col = colMatch[1].trim();
          const op = colMatch[2].trim().toUpperCase();
          let val = colMatch[3].trim();
          // Remove quotes
          val = val.replace(/^["']|["']$/g, '');

          data = data.filter((r) => {
            const cell = r[col];
            if (cell == null) return false;
            const cellStr = String(cell).toLowerCase();
            const valStr = val.toLowerCase();
            const cellNum = parseFloat(cell);
            const valNum = parseFloat(val);

            switch (op) {
              case '=': return cellStr === valStr || (!isNaN(cellNum) && !isNaN(valNum) && cellNum === valNum);
              case '!=':
              case '<>': return cellStr !== valStr;
              case '>': return !isNaN(cellNum) && !isNaN(valNum) && cellNum > valNum;
              case '<': return !isNaN(cellNum) && !isNaN(valNum) && cellNum < valNum;
              case '>=': return !isNaN(cellNum) && !isNaN(valNum) && cellNum >= valNum;
              case '<=': return !isNaN(cellNum) && !isNaN(valNum) && cellNum <= valNum;
              case 'CONTAINS': return cellStr.includes(valStr);
              default: return cellStr.includes(valStr);
            }
          });
        }
      }

      // ORDER BY clause
      const orderMatch = queryText.match(/ORDER\s+BY\s+(.+?)(?:\s+LIMIT|$)/i);
      if (orderMatch) {
        const orderParts = orderMatch[1].trim().split(/\s+/);
        const col = orderParts[0];
        const dir = orderParts[1]?.toUpperCase() === 'DESC' ? -1 : 1;
        data.sort((a, b) => {
          const av = a[col];
          const bv = b[col];
          if (av == null && bv == null) return 0;
          if (av == null) return 1;
          if (bv == null) return -1;
          const an = parseFloat(av);
          const bn = parseFloat(bv);
          if (!isNaN(an) && !isNaN(bn)) return (an - bn) * dir;
          return String(av).localeCompare(String(bv)) * dir;
        });
      }

      // LIMIT clause
      const limitMatch = queryText.match(/LIMIT\s+(\d+)/i);
      const limit = limitMatch ? parseInt(limitMatch[1]) : data.length;

      // SELECT clause - determine what to return
      const selectMatch = queryText.match(/SELECT\s+(.+?)(?:\s+WHERE|\s+ORDER\s+BY|\s+LIMIT|$)/i);
      let selectCols = headers;
      let isAggregate = false;
      let aggregateResult = null;

      if (selectMatch) {
        const selectPart = selectMatch[1].trim();
        if (selectPart === '*') {
          selectCols = headers;
        } else if (selectPart.toUpperCase().includes('COUNT(*)')) {
          isAggregate = true;
          aggregateResult = { label: 'COUNT(*)', value: data.length };
        } else if (selectPart.toUpperCase().includes('SUM(')) {
          isAggregate = true;
          const col = selectPart.match(/SUM\((.+?)\)/i)?.[1];
          const sum = data.reduce((s, r) => s + (parseFloat(r[col]) || 0), 0);
          aggregateResult = { label: `SUM(${col})`, value: sum };
        } else if (selectPart.toUpperCase().includes('AVG(')) {
          isAggregate = true;
          const col = selectPart.match(/AVG\((.+?)\)/i)?.[1];
          const vals = data.map((r) => parseFloat(r[col])).filter((v) => !isNaN(v));
          const avg = vals.length ? vals.reduce((s, v) => s + v, 0) / vals.length : 0;
          aggregateResult = { label: `AVG(${col})`, value: avg };
        } else if (selectPart.toUpperCase().includes('DISTINCT')) {
          const col = selectPart.match(/DISTINCT\s+(.+)/i)?.[1];
          const vals = [...new Set(data.map((r) => r[col]).filter(Boolean))];
          selectCols = [col];
          data = vals.slice(0, limit).map((v) => ({ [col]: v }));
        } else {
          selectCols = selectPart.split(',').map((c) => c.trim());
        }
      }

      if (isAggregate) {
        setResults({ type: 'aggregate', data: aggregateResult });
      } else {
        const limited = data.slice(0, limit);
        setResults({ type: 'table', data: limited, columns: selectCols });
      }
    } catch (err) {
      setError(`Error: ${err.message}`);
    }
  };

  return (
    <div className="query-view">
      {/* Query input */}
      <div className="panel" style={{ marginBottom: '1rem' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', marginBottom: '0.5rem' }}>
          <Terminal size={16} color="var(--accent)" />
          <h3 style={{ margin: 0, fontSize: '0.9rem' }}>Consulta SQL-like</h3>
        </div>
        <div style={{ display: 'flex', gap: '0.5rem' }}>
          <input
            type="text"
            className="filter__input"
            placeholder="SELECT * WHERE columna CONTAINS 'valor' LIMIT 10"
            value={queryText}
            onChange={(e) => setQueryText(e.target.value)}
            onKeyDown={(e) => e.key === 'Enter' && executeQuery()}
            style={{ flex: 1, fontFamily: 'monospace', fontSize: '0.85rem' }}
          />
          <button className="btn btn--primary" onClick={executeQuery}>
            <Play size={14} /> Ejecutar
          </button>
        </div>
        {error && (
          <div style={{ color: 'var(--red)', fontSize: '0.8rem', marginTop: '0.5rem' }}>{error}</div>
        )}
      </div>

      {/* Sample queries */}
      <div className="panel" style={{ marginBottom: '1rem' }}>
        <h4 style={{ margin: '0 0 0.5rem', fontSize: '0.8rem', color: 'var(--text-dim)' }}>Consultas de ejemplo</h4>
        <div style={{ display: 'flex', flexWrap: 'wrap', gap: '0.4rem' }}>
          {sampleQueries.map((sq) => (
            <button
              key={sq.label}
              className="btn btn--ghost btn--sm"
              onClick={() => setQueryText(sq.query)}
              title={sq.query}
            >
              {sq.label}
            </button>
          ))}
        </div>
      </div>

      {/* Schema reference */}
      <div className="panel" style={{ marginBottom: '1rem' }}>
        <h4 style={{ margin: '0 0 0.5rem', fontSize: '0.8rem', color: 'var(--text-dim)' }}>
          <Database size={12} /> Columnas disponibles ({headers.length})
        </h4>
        <div style={{ display: 'flex', flexWrap: 'wrap', gap: '0.3rem' }}>
          {headers.map((h) => (
            <button
              key={h}
              className="badge badge--dim"
              style={{ cursor: 'pointer', fontSize: '0.7rem' }}
              onClick={() => setQueryText((prev) => prev + (prev ? ', ' : '') + h)}
              title={`Click para insertar "${h}"`}
            >
              {h}
              {numericCols.includes(h) && <span style={{ color: 'var(--green)', marginLeft: '0.25rem' }}>●</span>}
            </button>
          ))}
        </div>
      </div>

      {/* Results */}
      {results && results.type === 'aggregate' && (
        <div className="panel" style={{ textAlign: 'center', padding: '2rem' }}>
          <h2 style={{ fontSize: '2rem', color: 'var(--accent)', margin: '0 0 0.5rem' }}>
            {fmt(results.data.value)}
          </h2>
          <p style={{ color: 'var(--text-dim)', fontSize: '0.85rem' }}>{results.data.label}</p>
        </div>
      )}

      {results && results.type === 'table' && (
        <div className="panel">
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '0.5rem' }}>
            <span style={{ fontSize: '0.8rem', color: 'var(--text-dim)' }}>
              <Table size={12} /> {results.data.length} resultado(s)
            </span>
          </div>
          <div className="tbl-wrap" style={{ maxHeight: '400px', overflow: 'auto' }}>
            <table className="tbl">
              <thead>
                <tr>
                  {results.columns.map((h) => <th key={h}>{h}</th>)}
                </tr>
              </thead>
              <tbody>
                {results.data.map((row, i) => (
                  <tr key={i}>
                    {results.columns.map((h) => (
                      <td key={h}>{row[h] != null ? String(row[h]).slice(0, 50) : '—'}</td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}
