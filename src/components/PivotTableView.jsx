import React, { useState, useMemo, useEffect } from 'react';
import {
  buildPivotTable,
  getNumericColumns,
  getLabelColumns,
} from '../utils/excelParser';
import {
  Table, LayoutGrid, ArrowUpDown, Sigma, Settings2,
  Download, Plus, X, GripVertical,
} from 'lucide-react';

const fmt = (n) => {
  if (n == null || isNaN(n)) return '—';
  if (Math.abs(n) >= 1e9) return (n / 1e9).toFixed(2) + 'B';
  if (Math.abs(n) >= 1e6) return (n / 1e6).toFixed(2) + 'M';
  if (Math.abs(n) >= 1e3) return (n / 1e3).toFixed(1) + 'K';
  return Number.isInteger(n) ? n.toLocaleString() : n.toLocaleString(undefined, { maximumFractionDigits: 2 });
};

const AGG_OPTIONS = [
  { key: 'sum', label: 'Suma' },
  { key: 'avg', label: 'Promedio' },
  { key: 'count', label: 'Conteo' },
  { key: 'max', label: 'Máximo' },
  { key: 'min', label: 'Mínimo' },
];

/* Detectar columnas por nombre flexible */
const findCol = (headers, patterns) => {
  const norm = (s) => s?.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z0-9]/g, '') || '';
  for (const h of headers) {
    const nh = norm(h);
    for (const p of patterns) {
      if (nh.includes(norm(p))) return h;
    }
  }
  return null;
};

export default function PivotTableView({ sheetData }) {
  const numCols = useMemo(() => getNumericColumns(sheetData), [sheetData]);
  const labelCols = useMemo(() => getLabelColumns(sheetData), [sheetData]);
  const allCols = useMemo(() => {
    const keys = Object.keys(sheetData?.[0] || {});
    return keys.filter((k) => k && !/^Columna_\d+$/i.test(k) && k.trim() !== '');
  }, [sheetData]);

  const [rowFields, setRowFields] = useState([]);
  const [colField, setColField] = useState('');
  const [valueField, setValueField] = useState('');
  const [aggFn, setAggFn] = useState('sum');

  /* defaults inteligentes */
  useEffect(() => {
    if (!allCols.length) return;

    const año = findCol(allCols, ['año', 'anio', 'year']);
    const mes = findCol(allCols, ['mes', 'month']);
    const nombreActuacion = findCol(allCols, ['nombre', 'actuacion', 'actuación', 'actuac']);

    const defaultRows = [];
    if (año) defaultRows.push(año);
    if (nombreActuacion) defaultRows.push(nombreActuacion);
    if (defaultRows.length === 0) defaultRows.push(allCols[0]);

    const defaultCol = mes || '';
    const defaultValue = nombreActuacion || numCols[0] || allCols[0];
    const isTextValue = labelCols.includes(defaultValue);
    const defaultAgg = isTextValue ? 'count' : 'sum';

    setRowFields(defaultRows);
    setColField(defaultCol);
    setValueField(defaultValue);
    setAggFn(defaultAgg);
  }, [sheetData, allCols, numCols, labelCols]);

  const pivot = useMemo(() => {
    if (!rowFields.length || !valueField) return null;
    return buildPivotTable(sheetData, rowFields, colField || null, valueField, aggFn);
  }, [sheetData, rowFields, colField, valueField, aggFn]);

  const addRowField = (field) => {
    if (!field || rowFields.includes(field)) return;
    setRowFields((prev) => [...prev, field]);
  };

  const removeRowField = (idx) => {
    setRowFields((prev) => prev.filter((_, i) => i !== idx));
  };

  const moveRowField = (idx, dir) => {
    setRowFields((prev) => {
      const next = [...prev];
      const swap = idx + dir;
      if (swap < 0 || swap >= next.length) return prev;
      [next[idx], next[swap]] = [next[swap], next[idx]];
      return next;
    });
  };

  const exportPivotCSV = () => {
    if (!pivot) return;
    const { rowLabels, colLabels, matrix, rowTotals, colTotals, grandTotal, rowFields: rfs } = pivot;
    const headers = [rfs.join(' + '), ...colLabels, 'Total'];
    const rows = rowLabels.map((r) => [
      r,
      ...colLabels.map((c) => matrix[r][c] ?? 0),
      rowTotals[r],
    ]);
    const totalRow = ['Total', ...colLabels.map((c) => colTotals[c]), grandTotal];

    const escape = (v) => {
      const s = String(v ?? '');
      return s.includes(',') || s.includes('"') || s.includes('\n') ? `"${s.replace(/"/g, '""')}"` : s;
    };

    const csv = [headers.map(escape).join(','), ...rows.map((r) => r.map(escape).join(',')), totalRow.map(escape).join(',')].join('\n');
    const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8;' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `tabla_dinamica_${Date.now()}.csv`;
    a.click();
    URL.revokeObjectURL(a.href);
  };

  const unusedCols = allCols.filter((c) => !rowFields.includes(c) && c !== colField && c !== valueField);

  return (
    <div className="pivot-view">
      {/* Controls */}
      <div className="pivot-controls">
        {/* Row fields */}
        <div className="pivot-field pivot-field--rows">
          <label><LayoutGrid size={12} /> Filas</label>
          <div className="pivot-row-chips">
            {rowFields.map((f, i) => (
              <span key={`${f}-${i}`} className="pivot-chip">
                <GripVertical size={10} style={{ cursor: 'grab', opacity: 0.6 }} />
                {f}
                <button className="pivot-chip-btn" onClick={() => removeRowField(i)} title="Quitar"><X size={10} /></button>
              </span>
            ))}
            {rowFields.length > 1 && (
              <button className="pivot-chip-btn pivot-chip-btn--move" onClick={() => moveRowField(0, 1)} title="Mover">
                <ArrowUpDown size={10} />
              </button>
            )}
            <select
              className="sel sel--compact pivot-add-field"
              value=""
              onChange={(e) => { if (e.target.value) { addRowField(e.target.value); e.target.value = ''; } }}
            >
              <option value="">+ Agregar fila…</option>
              {unusedCols.map((c) => <option key={c} value={c}>{c}</option>)}
            </select>
          </div>
        </div>

        <div className="pivot-field">
          <label><ArrowUpDown size={12} /> Columnas</label>
          <select value={colField} onChange={(e) => setColField(e.target.value)}>
            <option value="">— Ninguna —</option>
            {allCols.map((c) => <option key={c} value={c}>{c}</option>)}
          </select>
        </div>

        <div className="pivot-field">
          <label><Sigma size={12} /> Valores</label>
          <select value={valueField} onChange={(e) => setValueField(e.target.value)}>
            {allCols.map((c) => <option key={c} value={c}>{c}</option>)}
          </select>
        </div>

        <div className="pivot-field">
          <label><Settings2 size={12} /> Agregación</label>
          <select value={aggFn} onChange={(e) => setAggFn(e.target.value)}>
            {AGG_OPTIONS.map((o) => <option key={o.key} value={o.key}>{o.label}</option>)}
          </select>
        </div>

        <button className="btn btn--export" onClick={exportPivotCSV} disabled={!pivot}>
          <Download size={14} /> CSV
        </button>
      </div>

      {/* Table */}
      {!pivot ? (
        <div className="chart-area chart-area--empty">
          <Table size={28} strokeWidth={1.5} />
          <p>Seleccione campos para generar la tabla dinámica</p>
        </div>
      ) : pivot.rowLabels.length === 0 ? (
        <div className="chart-area chart-area--empty">
          <Table size={28} strokeWidth={1.5} />
          <p>No hay datos para mostrar con la configuración actual</p>
        </div>
      ) : (
        <div className="pivot-wrap">
          <table className="pivot-table">
            <thead>
              <tr>
                <th className="pivot-corner">
                  <span className="pivot-corner-label">{pivot.rowFields.join(' + ')}</span>
                  {colField && <span className="pivot-corner-sublabel">{colField}</span>}
                </th>
                {pivot.colLabels.map((c) => (
                  <th key={c} className="pivot-col-header" title={c}>
                    {c === '__total__' ? 'Total' : c}
                  </th>
                ))}
                <th className="pivot-total-header">Total</th>
              </tr>
            </thead>
            <tbody>
              {pivot.rowLabels.map((r) => (
                <tr key={r}>
                  <td className="pivot-row-header" title={r}>{r}</td>
                  {pivot.colLabels.map((c) => (
                    <td key={c} className="pivot-cell">
                      {fmt(pivot.matrix[r][c])}
                    </td>
                  ))}
                  <td className="pivot-row-total">{fmt(pivot.rowTotals[r])}</td>
                </tr>
              ))}
              <tr className="pivot-grand-row">
                <td className="pivot-row-header">Total</td>
                {pivot.colLabels.map((c) => (
                  <td key={c} className="pivot-col-total">{fmt(pivot.colTotals[c])}</td>
                ))}
                <td className="pivot-grand-total">{fmt(pivot.grandTotal)}</td>
              </tr>
            </tbody>
          </table>
        </div>
      )}

      {pivot && (
        <p className="pivot-meta">
          {pivot.rowLabels.length.toLocaleString()} filas × {pivot.colLabels.length.toLocaleString()} columnas
          {' · '}
          {AGG_OPTIONS.find((o) => o.key === aggFn)?.label} de <strong>{valueField}</strong>
          {colField && <> agrupado por <strong>{colField}</strong></>}
        </p>
      )}
    </div>
  );
}
