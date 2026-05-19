import React, { useState, useMemo, useCallback, useRef, useEffect } from 'react';
import {
  Chart as ChartJS, CategoryScale, LinearScale, BarElement, Title, Tooltip,
  Legend, PointElement, LineElement, ArcElement, RadialLinearScale, Filler,
} from 'chart.js';
import { Bar, Line, Doughnut, Radar, PolarArea } from 'react-chartjs-2';
import {
  getColumnStats, getNumericColumns, getLabelColumns, getCorrelation,
  getDataQuality, getUniqueValues, aggregateByDimension, exportToCSV,
  exportToExcel, detectAnomalies,
} from '../utils/excelParser';
import {
  LayoutDashboard, PieChart, LineChart, BarChart2, Table,
  Download, Search, Target, Hexagon, Filter, FileDown, X,
  ChevronLeft, ChevronRight, ChevronsLeft, ChevronsRight,
  Maximize2, Minimize2, Activity, TrendingUp, TrendingDown, Minus,
  Grid3X3, Info, Sparkles, SlidersHorizontal, Trash2, Plus,
  FileSpreadsheet, AlertCircle, Moon, Sun, Database, Terminal,
} from 'lucide-react';
import SmartOverview from './SmartOverview';
import QueryView from './QueryView';

ChartJS.register(
  CategoryScale, LinearScale, BarElement, PointElement, LineElement,
  ArcElement, RadialLinearScale, Filler, Title, Tooltip, Legend,
);

const CIRCULAR = new Set(['doughnut', 'polarArea']);
const MAX_SLICES = 20;
const MAX_CHART_POINTS = 300;
const PAGE_SIZE = 50;

const PALETTE = [
  'rgba(99,102,241,.75)', 'rgba(168,85,247,.75)', 'rgba(59,130,246,.75)',
  'rgba(34,197,94,.75)', 'rgba(245,158,11,.75)', 'rgba(239,68,68,.75)',
  'rgba(14,165,233,.75)', 'rgba(236,72,153,.75)', 'rgba(20,184,166,.75)',
  'rgba(251,146,60,.75)', 'rgba(139,92,246,.75)', 'rgba(6,182,212,.75)',
];
const BORDERS = PALETTE.map((c) => c.replace('.75', '1'));

const CHART_TYPES = [
  { key: 'bar', icon: BarChart2, label: 'Barras' },
  { key: 'line', icon: LineChart, label: 'Líneas' },
  { key: 'doughnut', icon: PieChart, label: 'Anillo' },
  { key: 'radar', icon: Target, label: 'Radar' },
  { key: 'polarArea', icon: Hexagon, label: 'Polar' },
];

const AGG_OPTIONS = [
  { key: 'sum', label: 'Suma' },
  { key: 'avg', label: 'Promedio' },
  { key: 'count', label: 'Conteo' },
  { key: 'max', label: 'Máximo' },
  { key: 'min', label: 'Mínimo' },
];

const fmt = (n) => {
  if (n == null || isNaN(n)) return '—';
  if (Math.abs(n) >= 1e9) return (n / 1e9).toFixed(2) + 'B';
  if (Math.abs(n) >= 1e6) return (n / 1e6).toFixed(2) + 'M';
  if (Math.abs(n) >= 1e3) return (n / 1e3).toFixed(1) + 'K';
  return n.toLocaleString(undefined, { maximumFractionDigits: 2 });
};

function AnimatedValue({ value, format = fmt }) {
  const [display, setDisplay] = useState(value);
  const prev = useRef(value);
  useEffect(() => {
    if (typeof value !== 'number' || typeof prev.current !== 'number') {
      setDisplay(value); prev.current = value; return;
    }
    const start = prev.current;
    const diff = value - start;
    if (Math.abs(diff) < 0.01) { setDisplay(value); prev.current = value; return; }
    const dur = 400; let t0;
    const step = (ts) => {
      if (!t0) t0 = ts;
      const p = Math.min((ts - t0) / dur, 1);
      const ease = 1 - Math.pow(1 - p, 3);
      setDisplay(start + diff * ease);
      if (p < 1) requestAnimationFrame(step);
    };
    requestAnimationFrame(step);
    prev.current = value;
  }, [value]);
  return <span>{format(display)}</span>;
}

function TrendBadge({ data, column }) {
  const trend = useMemo(() => {
    if (!data || data.length < 4) return null;
    const vals = data.map((r) => Number(r[column])).filter((v) => !isNaN(v));
    if (vals.length < 4) return null;
    const half = Math.floor(vals.length / 2);
    const first = vals.slice(0, half).reduce((a, b) => a + b, 0) / half;
    const second = vals.slice(half).reduce((a, b) => a + b, 0) / (vals.length - half);
    const pct = first !== 0 ? ((second - first) / Math.abs(first)) * 100 : 0;
    return { pct, dir: pct > 1 ? 'up' : pct < -1 ? 'down' : 'flat' };
  }, [data, column]);

  if (!trend) return null;
  const Icon = trend.dir === 'up' ? TrendingUp : trend.dir === 'down' ? TrendingDown : Minus;
  const cls = trend.dir === 'up' ? 'trend--up' : trend.dir === 'down' ? 'trend--down' : 'trend--flat';
  return (
    <span className={`trend-badge ${cls}`}>
      <Icon size={12} /> {trend.pct > 0 ? '+' : ''}{trend.pct.toFixed(1)}%
    </span>
  );
}

/* ═══════════════════════════════════════════════════
   FILTER CHIP
   ═══════════════════════════════════════════════════ */
function FilterChip({ filter, onRemove }) {
  const opLabels = { includes: 'contiene', equals: '=', '>': '>', '<': '<', startsWith: 'empieza con' };
  return (
    <span className="filter-chip">
      <strong>{filter.column}</strong> {opLabels[filter.op] || filter.op} <em>{filter.value}</em>
      <button onClick={onRemove} aria-label="Quitar filtro"><X size={12} /></button>
    </span>
  );
}

/* ═══════════════════════════════════════════════════
   ADVANCED FILTER BUILDER
   ═══════════════════════════════════════════════════ */
function FilterBuilder({ headers, onAdd, numericCols }) {
  const [column, setColumn] = useState('');
  const [op, setOp] = useState('includes');
  const [value, setValue] = useState('');

  const isNumeric = numericCols.includes(column);
  const ops = isNumeric
    ? [{ key: 'equals', label: '=' }, { key: '>', label: '>' }, { key: '<', label: '<' }, { key: 'includes', label: 'Contiene' }]
    : [{ key: 'includes', label: 'Contiene' }, { key: 'equals', label: 'Igual a' }, { key: 'startsWith', label: 'Empieza con' }];

  const handleAdd = () => {
    if (!column || !value) return;
    onAdd({ column, op, value });
    setValue('');
  };

  return (
    <div className="filter-builder">
      <select className="sel sel--compact" value={column} onChange={(e) => { setColumn(e.target.value); setOp('includes'); }}>
        <option value="">Columna…</option>
        {headers.map((h) => <option key={h} value={h}>{h}</option>)}
      </select>
      <select className="sel sel--compact" value={op} onChange={(e) => setOp(e.target.value)}>
        {ops.map((o) => <option key={o.key} value={o.key}>{o.label}</option>)}
      </select>
      <input
        className="filter__input"
        type="text"
        placeholder="Valor…"
        value={value}
        onChange={(e) => setValue(e.target.value)}
        onKeyDown={(e) => e.key === 'Enter' && handleAdd()}
      />
      <button className="btn btn--icon btn--primary" onClick={handleAdd} title="Agregar filtro">
        <Plus size={14} />
      </button>
    </div>
  );
}

/* ═══════════════════════════════════════════════════
   MAIN COMPONENT
   ═══════════════════════════════════════════════════ */
export default function AnalysisDashboard({ sheetData, workbook }) {
  const [cols, setCols] = useState([]);
  const [chart, setChart] = useState('bar');
  const [view, setView] = useState('overview');
  const [search, setSearch] = useState('');
  const [filters, setFilters] = useState([]);
  const [page, setPage] = useState(0);
  const [sortKey, setSortKey] = useState(null);
  const [sortDir, setSortDir] = useState('asc');
  const [fullscreen, setFullscreen] = useState(false);
  const [aggFn, setAggFn] = useState('sum');
  const [groupBy, setGroupBy] = useState('');
  const chartRef = useRef(null);

  const numCols = useMemo(() => getNumericColumns(sheetData), [sheetData]);
  const labelCols = useMemo(() => getLabelColumns(sheetData), [sheetData]);
  const headers = useMemo(() => sheetData?.length ? Object.keys(sheetData[0]) : [], [sheetData]);
  const quality = useMemo(() => getDataQuality(sheetData), [sheetData]);

  /* ── Reset on data change ── */
  useEffect(() => {
    setCols(numCols.length ? [numCols[0]] : []);
    setSearch(''); setFilters([]); setPage(0); setSortKey(null);
    setGroupBy(labelCols[0] || '');
    // Always start with overview (Resumen) view
    setView('overview');
  }, [sheetData, numCols, headers, labelCols]);

  useEffect(() => { setPage(0); }, [search, filters, sortKey, sortDir]);

  /* ── Column toggle ── */
  const toggleCol = useCallback((c) => {
    if (CIRCULAR.has(chart)) { setCols([c]); return; }
    setCols((prev) => prev.includes(c) ? (prev.length > 1 ? prev.filter((x) => x !== c) : prev) : [...prev, c]);
  }, [chart]);

  /* ── Sort ── */
  const handleSort = useCallback((key) => {
    setSortDir((d) => sortKey === key ? (d === 'asc' ? 'desc' : 'asc') : 'asc');
    setSortKey(key);
  }, [sortKey]);

  /* ── Filters ── */
  const addFilter = useCallback((f) => setFilters((prev) => [...prev, f]), []);
  const removeFilter = useCallback((idx) => setFilters((prev) => prev.filter((_, i) => i !== idx)), []);
  const clearFilters = useCallback(() => { setFilters([]); setSearch(''); }, []);

  /* ── Filtered rows ── */
  const rows = useMemo(() => {
    if (!sheetData?.length) return [];
    let result = sheetData;

    // Global search
    const term = search.trim();
    if (term) {
      const lo = term.toLowerCase();
      result = result.filter((r) =>
        Object.values(r).some((v) => v != null && String(v).toLowerCase().includes(lo))
      );
    }

    // Column filters
    for (const f of filters) {
      const { column, op, value } = f;
      const lv = value.toLowerCase();
      result = result.filter((r) => {
        const v = r[column];
        if (v == null) return false;
        const sv = String(v).toLowerCase();
        if (op === 'includes') return sv.includes(lv);
        if (op === 'equals') return sv === lv;
        if (op === 'startsWith') return sv.startsWith(lv);
        const nv = parseFloat(v);
        const wv = parseFloat(value);
        if (isNaN(nv) || isNaN(wv)) return false;
        if (op === '>') return nv > wv;
        if (op === '<') return nv < wv;
        return true;
      });
    }

    // Sort
    if (sortKey) {
      result = [...result].sort((a, b) => {
        const va = a[sortKey], vb = b[sortKey];
        if (va == null && vb == null) return 0;
        if (va == null) return 1;
        if (vb == null) return -1;
        const na = parseFloat(va), nb = parseFloat(vb);
        const cmp = (!isNaN(na) && !isNaN(nb)) ? na - nb : String(va).localeCompare(String(vb));
        return sortDir === 'asc' ? cmp : -cmp;
      });
    }

    return result;
  }, [sheetData, search, filters, sortKey, sortDir]);

  const totalPages = Math.max(1, Math.ceil(rows.length / PAGE_SIZE));
  const pagedRows = useMemo(() => rows.slice(page * PAGE_SIZE, (page + 1) * PAGE_SIZE), [rows, page]);

  /* ── Stats ── */
  const primary = cols[0];
  const stats = useMemo(() => primary && rows.length ? getColumnStats(rows, primary) : null, [rows, primary]);

  /* ── Correlation ── */
  const corrMatrix = useMemo(() => {
    if (numCols.length < 2) return null;
    const selected = numCols.slice(0, 10);
    const matrix = {};
    for (const c1 of selected) {
      matrix[c1] = {};
      for (const c2 of selected) {
        matrix[c1][c2] = c1 === c2 ? 1 : getCorrelation(rows, c1, c2);
      }
    }
    return { cols: selected, matrix };
  }, [numCols, rows]);

  /* ── Chart data (with aggregation support) ── */
  const chartData = useMemo(() => {
    if (!cols.length || !rows.length) return null;

    // If grouping is enabled and valid, aggregate
    if (groupBy && labelCols.includes(groupBy) && cols.length > 0) {
      const agg = aggregateByDimension(rows, groupBy, cols, aggFn);
      if (agg && agg.labels.length > 0) {
        const labels = agg.labels.slice(0, MAX_CHART_POINTS);
        return {
          labels,
          datasets: cols.map((c, i) => ({
            label: `${c} (${AGG_OPTIONS.find((a) => a.key === aggFn)?.label})`,
            data: agg.aggregated[c].slice(0, MAX_CHART_POINTS),
            backgroundColor: PALETTE[i % PALETTE.length],
            borderColor: BORDERS[i % BORDERS.length],
            borderWidth: 2,
            borderRadius: 4,
          })),
        };
      }
    }

    // Default: raw data with label detection
    const chartRows = rows.length <= MAX_CHART_POINTS ? rows : rows.filter((_, i) => {
      const step = Math.ceil(rows.length / MAX_CHART_POINTS);
      return i % step === 0;
    });

    const lk = labelCols.find((k) => k && rows[0][k] != null);
    const labels = chartRows.map((r, i) => (lk && r[lk] != null ? String(r[lk]) : `#${i + 1}`));

    if (CIRCULAR.has(chart)) {
      const vals = chartRows.map((r) => Number(r[primary]) || 0);
      if (vals.length > MAX_SLICES) {
        const pairs = labels.map((l, i) => ({ l, v: vals[i] })).sort((a, b) => b.v - a.v);
        const top = pairs.slice(0, MAX_SLICES);
        const rest = pairs.slice(MAX_SLICES).reduce((s, x) => s + x.v, 0);
        if (rest > 0) top.push({ l: 'Otros', v: rest });
        return { labels: top.map((x) => x.l), datasets: [{ data: top.map((x) => x.v), backgroundColor: PALETTE, borderColor: BORDERS, borderWidth: 1 }] };
      }
      return { labels, datasets: [{ data: vals, backgroundColor: PALETTE, borderColor: BORDERS, borderWidth: 1 }] };
    }

    return {
      labels,
      datasets: cols.map((c, i) => ({
        label: c,
        data: chartRows.map((r) => Number(r[c]) || 0),
        backgroundColor: chart === 'radar' ? PALETTE[i % PALETTE.length].replace('.75', '.15') : PALETTE[i % PALETTE.length],
        borderColor: BORDERS[i % BORDERS.length],
        borderWidth: 2,
        pointBackgroundColor: BORDERS[i % BORDERS.length],
        pointRadius: chartRows.length > 80 ? 0 : 3,
        fill: chart === 'radar' || chart === 'line',
        tension: 0.35,
      })),
    };
  }, [rows, cols, chart, primary, groupBy, aggFn, labelCols]);

  /* ── Chart options ── */
  const chartOpts = useMemo(() => {
    const radial = CIRCULAR.has(chart) || chart === 'radar';
    return {
      responsive: true,
      maintainAspectRatio: false,
      animation: { duration: 400, easing: 'easeOutQuart' },
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: {
          position: CIRCULAR.has(chart) ? 'right' : 'top',
          labels: { color: '#c9d1d9', font: { family: "'Inter',sans-serif", size: 11 }, usePointStyle: true, padding: 16, boxWidth: 8 },
        },
        tooltip: {
          backgroundColor: 'rgba(10,14,20,.96)', titleColor: '#f0f6fc', bodyColor: '#c9d1d9',
          borderColor: 'rgba(99,102,241,.3)', borderWidth: 1, padding: 12, cornerRadius: 10,
          callbacks: { label: (ctx) => ` ${ctx.dataset.label || ''}: ${fmt(ctx.parsed.y ?? ctx.parsed)}` },
        },
      },
      scales: radial
        ? { r: { ticks: { backdropColor: 'transparent', color: '#8b949e' }, grid: { color: 'rgba(255,255,255,.06)' }, angleLines: { color: 'rgba(255,255,255,.06)' }, pointLabels: { color: '#c9d1d9', font: { size: 10 } } } }
        : {
            y: { ticks: { color: '#8b949e', font: { size: 11 }, callback: (v) => fmt(v) }, grid: { color: 'rgba(255,255,255,.04)' }, beginAtZero: true },
            x: { ticks: { color: '#8b949e', font: { size: 10 }, maxRotation: 45, autoSkip: true, maxTicksLimit: 30 }, grid: { display: false } },
          },
    };
  }, [chart]);

  /* ── Downloads ── */
  const dlPNG = useCallback(() => {
    if (!chartRef.current) return;
    const a = document.createElement('a');
    a.download = `grafico_${cols.join('-')}_${Date.now()}.png`;
    a.href = chartRef.current.toBase64Image('image/png', 1);
    a.click();
  }, [cols]);

  const dlCSV = useCallback(() => {
    if (!rows.length) return;
    exportToCSV(rows, `datos_${Date.now()}.csv`);
  }, [rows]);

  const dlExcel = useCallback(() => {
    if (!rows.length) return;
    exportToExcel(rows, `datos_${Date.now()}.xlsx`, filters);
  }, [rows, filters]);

  if (!sheetData?.length) {
    return (
      <div className="panel empty-state">
        <LayoutDashboard size={48} strokeWidth={1} />
        <h3>Sin datos</h3>
        <p>Esta hoja no contiene datos para analizar.</p>
      </div>
    );
  }

  const isTable = view === 'table';
  const isCorr = view === 'correlation';
  const isOverview = view === 'overview';
  const isAnomalies = view === 'anomalies';
  const isQuery = view === 'query';
  const hasActiveFilters = search || filters.length > 0;

  /* ── Anomalies ── */
  const anomalies = useMemo(() => {
    if (!isAnomalies || !rows.length) return [];
    return detectAnomalies(rows, numCols.slice(0, 8));
  }, [isAnomalies, rows, numCols]);

  const ChartEl = () => {
    if (!chartData) return (
      <div className="chart-area chart-area--empty">
        <LayoutDashboard size={28} strokeWidth={1.5} />
        <p>Seleccione columnas numéricas para visualizar</p>
      </div>
    );
    const p = { ref: chartRef, data: chartData, options: chartOpts };
    switch (chart) {
      case 'line': return <Line {...p} />;
      case 'doughnut': return <Doughnut {...p} />;
      case 'radar': return <Radar {...p} />;
      case 'polarArea': return <PolarArea {...p} />;
      default: return <Bar {...p} />;
    }
  };

  return (
    <div className={`dash${fullscreen ? ' dash--fullscreen' : ''}`}>
      {/* ── SIDEBAR ── */}
      <aside className="sidebar">
        {/* View toggle */}
        <div className="panel" style={{ padding: '0.4rem' }}>
          <div className="vtoggle">
            <button className={`vtoggle__btn${isOverview ? ' vtoggle__btn--active' : ''}`} onClick={() => setView('overview')} title="Resumen general con KPIs y filtros">
              <Sparkles size={14} /> Resumen
            </button>
            <button className={`vtoggle__btn${view === 'chart' ? ' vtoggle__btn--active' : ''}`} onClick={() => setView('chart')} title="Gráficos interactivos">
              <BarChart2 size={14} /> Gráficos
            </button>
            <button className={`vtoggle__btn${view === 'table' ? ' vtoggle__btn--active' : ''}`} onClick={() => setView('table')} title="Tabla de datos con paginación">
              <Table size={14} /> Tabla
            </button>
            <button className={`vtoggle__btn${view === 'query' ? ' vtoggle__btn--active' : ''}`} onClick={() => setView('query')} title="Consulta SQL-like">
              <Terminal size={14} /> Query
            </button>
            <button className={`vtoggle__btn${isCorr ? ' vtoggle__btn--active' : ''}`} onClick={() => setView('correlation')} title="Matriz de correlación">
              <Grid3X3 size={14} /> Correlación
            </button>
            <button className={`vtoggle__btn${view === 'anomalies' ? ' vtoggle__btn--active' : ''}`} onClick={() => setView('anomalies')} title="Detección de anomalías">
              <AlertCircle size={14} /> Anomalías
            </button>
          </div>
        </div>

        {/* Quality */}
        {quality && (
          <div className="panel" style={{ padding: '0.6rem 0.8rem' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '0.4rem', marginBottom: '0.3rem' }}>
              <Activity size={12} color="var(--green)" />
              <span style={{ fontSize: '0.65rem', color: 'var(--text-dim)', fontWeight: 600 }}>CALIDAD DE DATOS</span>
            </div>
            <div className="quality-bar" style={{ height: '4px' }}>
              <div className="quality-bar__fill" style={{ width: `${quality.completeness}%`, background: quality.completeness > 90 ? 'var(--green)' : quality.completeness > 70 ? 'var(--amber)' : 'var(--red)' }} />
            </div>
            <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '0.65rem', color: 'var(--text-dim)', marginTop: '0.3rem' }}>
              <span>{quality.completeness.toFixed(0)}% completo</span>
              <span>{quality.totalRows.toLocaleString()} filas · {quality.totalCols} cols</span>
            </div>
          </div>
        )}

        {/* Metrics */}
        {!isOverview && (
          <div className="panel ctrl">
            <span className="ctrl__label"><Target size={12} /> Métricas</span>
            <div className="collist">
              {numCols.map((c) => {
                const on = cols.includes(c);
                const off = CIRCULAR.has(chart) && !on && cols.length >= 1;
                return (
                  <label key={c} className={`colitem${on ? ' colitem--on' : ''}`}>
                    <input type="checkbox" checked={on} disabled={off} onChange={() => toggleCol(c)} />
                    <span className="colitem__name" title={c}>{c}</span>
                    {on && <TrendBadge data={rows} column={c} />}
                  </label>
                );
              })}
              {numCols.length === 0 && <p className="notice">No se detectaron columnas numéricas.</p>}
            </div>
            {CIRCULAR.has(chart) && numCols.length > 0 && <p className="notice" style={{ marginTop: '0.4rem' }}>Solo 1 métrica en gráfico circular.</p>}
          </div>
        )}

        {/* Chart type */}
        {view === 'chart' && !isOverview && (
          <div className="panel ctrl">
            <span className="ctrl__label"><LayoutDashboard size={12} /> Tipo de Gráfico</span>
            <div className="chart-type-grid">
              {CHART_TYPES.map(({ key, icon: Icon, label }) => (
                <button key={key} className={`chart-type-btn${chart === key ? ' chart-type-btn--active' : ''}`} onClick={() => setChart(key)} title={label}>
                  <Icon size={16} /><span>{label}</span>
                </button>
              ))}
            </div>
          </div>
        )}

        {/* Group by (for charts) */}
        {view === 'chart' && labelCols.length > 0 && (
          <div className="panel ctrl">
            <span className="ctrl__label"><SlidersHorizontal size={12} /> Agrupar por</span>
            <select className="sel sel--compact" value={groupBy} onChange={(e) => setGroupBy(e.target.value)}>
              <option value="">Sin agrupar (datos raw)</option>
              {labelCols.map((h) => <option key={h} value={h}>{h}</option>)}
            </select>
            {groupBy && (
              <select className="sel sel--compact" value={aggFn} onChange={(e) => setAggFn(e.target.value)} style={{ marginTop: '0.5rem' }}>
                {AGG_OPTIONS.map((o) => <option key={o.key} value={o.key}>{o.label}</option>)}
              </select>
            )}
          </div>
        )}

        {/* Stats */}
        {stats && !isOverview && (
          <div className="panel ctrl" style={{ marginTop: 'auto' }}>
            <span className="ctrl__label"><Info size={12} /> Resumen — {primary}</span>
            <div className="stats">
              {[
                { label: 'Suma', val: stats.sum, accent: true },
                { label: 'Promedio', val: stats.mean },
                { label: 'Mediana', val: stats.median },
                { label: 'Mínimo', val: stats.min },
                { label: 'Máximo', val: stats.max },
                { label: 'Registros', val: stats.count },
                { label: 'σ Desv.', val: stats.stdDev },
                { label: 'CV %', val: stats.cv, suffix: '%' },
              ].map(({ label, val, accent, suffix }) => (
                <div className="stat" key={label}>
                  <span className="stat__label">{label}</span>
                  <span className={`stat__val${accent ? ' stat__val--accent' : ''}`}>
                    <AnimatedValue value={val} />{suffix || ''}
                  </span>
                </div>
              ))}
            </div>
          </div>
        )}
      </aside>

      {/* ── MAIN ── */}
      <div className={`panel main${isOverview ? ' main--overview' : ''}`}>
        {!isOverview && (
          <div className="main__head">
            <div>
              <h3 className="main__title">
                {isTable && <><Table size={20} color="var(--accent)" /> Explorador de Datos</>}
                {isCorr && <><Grid3X3 size={20} color="var(--accent)" /> Matriz de Correlación</>}
                {view === 'chart' && (
                  <>
                    {CHART_TYPES.find((t) => t.key === chart)?.icon &&
                      React.createElement(CHART_TYPES.find((t) => t.key === chart).icon, { size: 20, color: 'var(--accent)' })}
                    Visualización
                  </>
                )}
              </h3>
              <p className="main__subtitle">
                {rows.length.toLocaleString()} registro(s)
                {hasActiveFilters && <span className="badge badge--blue" style={{ marginLeft: '0.5rem' }}>filtrados</span>}
                {view === 'chart' && rows.length > MAX_CHART_POINTS && !groupBy && (
                  <span className="badge badge--dim" style={{ marginLeft: '0.5rem' }}>mostrando {Math.min(rows.length, MAX_CHART_POINTS)} pts</span>
                )}
                {groupBy && <span className="badge badge--purple" style={{ marginLeft: '0.5rem' }}>agrupado por {groupBy}</span>}
              </p>
            </div>

            <div className="main__actions">
              {view === 'chart' && (
                <button className="btn btn--icon" onClick={() => setFullscreen((f) => !f)} title={fullscreen ? 'Salir' : 'Pantalla completa'}>
                  {fullscreen ? <Minimize2 size={14} /> : <Maximize2 size={14} />}
                </button>
              )}
              {view === 'chart' && chartData && (
                <button className="btn btn--export" onClick={dlPNG}><Download size={14} /> PNG</button>
              )}
              <button className="btn btn--export" onClick={dlExcel}><FileSpreadsheet size={14} /> Excel</button>
              <button className="btn btn--export" onClick={dlCSV}><FileDown size={14} /> CSV</button>
            </div>
          </div>
        )}

        {/* Active filters bar */}
        {!isOverview && (
          <div className="active-filters-bar">
            <div className="active-filters">
              <Filter size={13} color="var(--text-dim)" />
              <input
                className="filter__input filter__input--global"
                type="search"
                placeholder="Buscar en todos los campos…"
                value={search}
                onChange={(e) => setSearch(e.target.value)}
              />
              <FilterBuilder headers={headers} onAdd={addFilter} numericCols={numCols} />
            </div>
            {hasActiveFilters && (
              <div className="filter-chips">
                {search && <FilterChip filter={{ column: 'Global', op: 'includes', value: search }} onRemove={() => setSearch('')} />}
                {filters.map((f, i) => <FilterChip key={i} filter={f} onRemove={() => removeFilter(i)} />)}
                <button className="btn btn--ghost btn--sm" onClick={clearFilters}><Trash2 size={12} /> Limpiar</button>
              </div>
            )}
          </div>
        )}

        {/* OVERVIEW */}
        {isOverview && <SmartOverview sheetData={sheetData} workbook={workbook} />}

        {/* CHART */}
        {view === 'chart' && <div className="chart-area"><ChartEl /></div>}

        {/* CORRELATION */}
        {isCorr && corrMatrix && (
          <div className="corr-container">
            <div className="corr-scroll">
              <table className="corr-table">
                <thead><tr><th></th>{corrMatrix.cols.map((c) => <th key={c} title={c}>{c.length > 12 ? c.slice(0, 12) + '…' : c}</th>)}</tr></thead>
                <tbody>
                  {corrMatrix.cols.map((c1) => (
                    <tr key={c1}>
                      <th title={c1}>{c1.length > 12 ? c1.slice(0, 12) + '…' : c1}</th>
                      {corrMatrix.cols.map((c2) => {
                        const v = corrMatrix.matrix[c1][c2];
                        const abs = v != null ? Math.abs(v) : 0;
                        const hue = v != null ? (v >= 0 ? 210 : 0) : 0;
                        return (
                          <td key={c2} className="corr-cell" style={{ background: v != null ? `hsla(${hue},70%,55%,${abs * 0.5})` : 'transparent' }} title={`${c1} ↔ ${c2}: ${v != null ? v.toFixed(4) : 'N/A'}`}>
                            {v != null ? v.toFixed(2) : '—'}
                          </td>
                        );
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <p className="corr-legend">
              <span style={{ color: 'hsl(210,70%,55%)' }}>■ Positiva</span>
              <span style={{ color: '#8b949e' }}>■ Sin correlación</span>
              <span style={{ color: 'hsl(0,70%,55%)' }}>■ Negativa</span>
            </p>
          </div>
        )}
        {isCorr && !corrMatrix && (
          <div className="chart-area chart-area--empty">
            <Grid3X3 size={28} strokeWidth={1.5} />
            <p>Se necesitan al menos 2 columnas numéricas</p>
          </div>
        )}

        {/* TABLE */}
        {isTable && (
          <>
            <div className="tbl-wrap">
              <table className="tbl">
                <thead>
                  <tr>
                    <th>#</th>
                    {headers.map((h) => (
                      <th key={h} onClick={() => handleSort(h)} title={`Ordenar por ${h}`}>
                        {h}
                        <span className={`sort-arrow${sortKey === h ? ' sort-arrow--active' : ''}`}>
                          {sortKey === h ? (sortDir === 'asc' ? ' ↑' : ' ↓') : ' ↕'}
                        </span>
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {pagedRows.length === 0 ? (
                    <tr><td colSpan={headers.length + 1} className="tbl__empty">
                      <Search size={22} style={{ opacity: 0.3, display: 'block', margin: '0 auto 0.5rem' }} />
                      Sin coincidencias.
                    </td></tr>
                  ) : (
                    pagedRows.map((row, i) => (
                      <tr key={i}>
                        <td>{page * PAGE_SIZE + i + 1}</td>
                        {headers.map((h) => {
                          const v = row[h];
                          const isNum = typeof v === 'number';
                          let txt = '—';
                          if (v != null) txt = isNum ? (v % 1 === 0 ? v.toLocaleString() : v.toLocaleString(undefined, { maximumFractionDigits: 2 })) : String(v);
                          return <td key={h} className={isNum ? 'num' : ''} title={v != null ? String(v) : ''}>{txt}</td>;
                        })}
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
            {rows.length > PAGE_SIZE && (
              <div className="pager">
                <button className="pager__btn" disabled={page === 0} onClick={() => setPage(0)}><ChevronsLeft size={14} /></button>
                <button className="pager__btn" disabled={page === 0} onClick={() => setPage((p) => p - 1)}><ChevronLeft size={14} /></button>
                <span className="pager__info">Página {page + 1} de {totalPages} · {rows.length.toLocaleString()} filas</span>
                <button className="pager__btn" disabled={page >= totalPages - 1} onClick={() => setPage((p) => p + 1)}><ChevronRight size={14} /></button>
                <button className="pager__btn" disabled={page >= totalPages - 1} onClick={() => setPage(totalPages - 1)}><ChevronsRight size={14} /></button>
              </div>
            )}
          </>
        )}

        {/* QUERY */}
        {isQuery && <QueryView sheetData={sheetData} />}

        {/* ANOMALIES */}
        {isAnomalies && (
          <div className="anomalies-container">
            {anomalies.length === 0 ? (
              <div className="chart-area chart-area--empty">
                <AlertCircle size={28} strokeWidth={1.5} />
                <p>No se detectaron anomalías significativas en los datos</p>
              </div>
            ) : (
              <>
                <div className="anomalies-summary">
                  <h3><AlertCircle size={16} /> Anomalías Detectadas</h3>
                  <p>Se encontraron valores atípicos usando el método IQR (rango intercuartílico)</p>
                </div>
                {anomalies.map((a) => (
                  <div key={a.column} className="anomaly-card">
                    <div className="anomaly-card__header">
                      <strong>{a.column}</strong>
                      <span className="anomaly-badge">{a.outliers.length} outliers ({a.summary.percentage}%)</span>
                    </div>
                    <div className="anomaly-card__stats">
                      <span>Media: {fmt(a.summary.mean)}</span>
                      <span>Desv. Est.: {fmt(a.summary.stdDev)}</span>
                      <span>Q1: {fmt(a.summary.q1)}</span>
                      <span>Q3: {fmt(a.summary.q3)}</span>
                      <span>Límite inf.: {fmt(a.bounds.lower)}</span>
                      <span>Límite sup.: {fmt(a.bounds.upper)}</span>
                    </div>
                    <div className="anomaly-card__outliers">
                      {a.outliers.slice(0, 10).map((o, i) => (
                        <span key={i} className="anomaly-chip" title={`Fila ${o.row + 1}, z-score: ${o.zScore}`}>
                          {fmt(o.value)}
                        </span>
                      ))}
                      {a.outliers.length > 10 && (
                        <span className="anomaly-chip anomaly-chip--more">+{a.outliers.length - 10} más</span>
                      )}
                    </div>
                  </div>
                ))}
              </>
            )}
          </div>
        )}
      </div>
    </div>
  );
}
