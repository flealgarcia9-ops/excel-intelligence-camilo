import React, { useMemo, useState, useRef } from 'react';
import * as XLSX from 'xlsx';
import {
  Chart as ChartJS, CategoryScale, LinearScale, BarElement, Title, Tooltip,
  Legend, PointElement, LineElement, ArcElement, Filler,
} from 'chart.js';
import { Bar, Line, Doughnut } from 'react-chartjs-2';
import {
  analyzeSheetStructure, aggregateByDimension, getTopN, getColumnStats,
} from '../utils/excelParser';
import {
  TrendingUp, TrendingDown, Minus, ArrowRight, Calendar,
  MapPin, Users, BarChart2, Download, Filter as FilterIcon, FileSpreadsheet, Activity,
} from 'lucide-react';

ChartJS.register(
  CategoryScale, LinearScale, BarElement, PointElement, LineElement,
  ArcElement, Filler, Title, Tooltip, Legend,
);

/* ── Palette ─────────────────────────────────────── */
const STAGE_COLORS = {
  'II_':  { bg: 'rgba(99,102,241,.75)',  border: 'rgba(99,102,241,1)',  soft: 'rgba(99,102,241,.12)',  label: 'Indagaciones' },
  'III_': { bg: 'rgba(168,85,247,.75)',   border: 'rgba(168,85,247,1)',  soft: 'rgba(168,85,247,.12)',  label: 'Investigaciones' },
  'IV_':  { bg: 'rgba(245,158,11,.75)',   border: 'rgba(245,158,11,1)',  soft: 'rgba(245,158,11,.12)', label: 'Juicios' },
  'V_':   { bg: 'rgba(34,197,94,.75)',    border: 'rgba(34,197,94,1)',   soft: 'rgba(34,197,94,.12)',  label: 'Querellas' },
};

const PALETTE = [
  'rgba(99,102,241,.75)', 'rgba(168,85,247,.75)', 'rgba(245,158,11,.75)',
  'rgba(34,197,94,.75)', 'rgba(239,68,68,.75)', 'rgba(14,165,233,.75)',
  'rgba(236,72,153,.75)', 'rgba(20,184,166,.75)',
];

/* ── Helpers ─────────────────────────────────────── */
const fmt = (n) => {
  if (n == null || isNaN(n)) return '—';
  if (Math.abs(n) >= 1e9) return (n / 1e9).toFixed(2) + 'B';
  if (Math.abs(n) >= 1e6) return (n / 1e6).toFixed(2) + 'M';
  if (Math.abs(n) >= 1e3) return (n / 1e3).toFixed(1) + 'K';
  return n.toLocaleString(undefined, { maximumFractionDigits: 0 });
};

const chartBase = {
  responsive: true,
  maintainAspectRatio: false,
  animation: { duration: 500 },
  interaction: { mode: 'index', intersect: false },
  plugins: {
    legend: { labels: { color: '#c9d1d9', font: { family: "'Inter',sans-serif", size: 11 }, usePointStyle: true, padding: 12, boxWidth: 8 } },
    tooltip: {
      backgroundColor: 'rgba(10,14,20,.96)', titleColor: '#f0f6fc', bodyColor: '#c9d1d9',
      borderColor: 'rgba(99,102,241,.3)', borderWidth: 1, padding: 10, cornerRadius: 8,
      callbacks: { label: (ctx) => ` ${ctx.dataset.label}: ${fmt(ctx.parsed.y ?? ctx.parsed)}` },
    },
  },
};

function KpiCard({ label, value, icon: Icon, color, detail, trend }) {
  const TrendIcon = trend > 1 ? TrendingUp : trend < -1 ? TrendingDown : Minus;
  const trendCls = trend > 1 ? 'kpi__trend--up' : trend < -1 ? 'kpi__trend--down' : 'kpi__trend--flat';

  return (
    <div className="kpi" style={{ '--kpi-color': color }}>
      <div className="kpi__header">
        <div className="kpi__icon">{Icon && <Icon size={20} />}</div>
        <span className="kpi__label">{label}</span>
      </div>
      <div className="kpi__value">{fmt(value)}</div>
      <div className="kpi__footer">
        {detail && <span className="kpi__detail">{detail}</span>}
        {trend != null && !isNaN(trend) && (
          <span className={`kpi__trend ${trendCls}`}>
            <TrendIcon size={12} /> {trend > 0 ? '+' : ''}{trend.toFixed(1)}%
          </span>
        )}
      </div>
    </div>
  );
}

function ChartCard({ title, children, onDownload }) {
  return (
    <div className="overview-chart-card">
      <div className="overview-chart-card__head">
        <h4>{title}</h4>
        {onDownload && (
          <button className="btn btn--icon" onClick={onDownload} title="Descargar PNG">
            <Download size={13} />
          </button>
        )}
      </div>
      <div className="overview-chart-card__body">
        {children}
      </div>
    </div>
  );
}

export default function SmartOverview({ sheetData, workbook }) {
  const [startYear, setStartYear] = useState('');
  const [startMonth, setStartMonth] = useState('');
  const [endYear, setEndYear] = useState('');
  const [endMonth, setEndMonth] = useState('');
  const [regionFilter, setRegionFilter] = useState('');
  const [fiscalFilter, setFiscalFilter] = useState('');
  const [unitFilter, setUnitFilter] = useState('');

  const yearChartRef = useRef(null);
  const stageChartRef = useRef(null);
  const topChartRef = useRef(null);

  /* ── Reset on data change ── */
  React.useEffect(() => {
    setStartYear(''); setStartMonth(''); setEndYear(''); setEndMonth('');
    setRegionFilter(''); setFiscalFilter(''); setUnitFilter('');
  }, [sheetData]);

  const months = useMemo(() => [
    { v: '1',  l: 'Enero' }, { v: '2',  l: 'Febrero' }, { v: '3',  l: 'Marzo' },
    { v: '4',  l: 'Abril' }, { v: '5',  l: 'Mayo' },    { v: '6',  l: 'Junio' },
    { v: '7',  l: 'Julio' }, { v: '8',  l: 'Agosto' },  { v: '9',  l: 'Septiembre' },
    { v: '10', l: 'Octubre' },{ v: '11', l: 'Noviembre' },{ v: '12', l: 'Diciembre' },
  ], []);

  /* ── Analyze structure ── */
  const structure = useMemo(() => analyzeSheetStructure(sheetData), [sheetData]);

  /* ── Unique dimension values ── */
  const yearsArr = useMemo(() => {
    const col = structure?.dimensions?.year;
    if (!sheetData?.length || !col) return [];
    const set = new Set();
    sheetData.forEach(r => {
      const v = r[col];
      if (v != null && v !== '') set.add(String(v));
    });
    return [...set].sort((a, b) => Number(a) - Number(b));
  }, [sheetData, structure]);

  const regionsArr = useMemo(() => {
    const col = structure?.dimensions?.region;
    if (!sheetData?.length || !col) return [];
    const set = new Set();
    sheetData.forEach(r => {
      const v = r[col];
      if (v != null && v !== '') set.add(String(v));
    });
    return [...set].sort();
  }, [sheetData, structure]);

  const unitsLookup = useMemo(() => {
    // Build a map from unit number -> unit description using the Unidades sheet
    const map = {};
    if (workbook?.data) {
      // Find the Unidades sheet (case-insensitive)
      const unidadesKey = workbook.sheetNames?.find(s => /unidad/i.test(s));
      if (unidadesKey && workbook.data[unidadesKey]) {
        for (const row of workbook.data[unidadesKey]) {
          // Find the number col and description col
          const numCol = Object.keys(row).find(k => /nro|numero|cod/i.test(k) && /unidad/i.test(k));
          const descCol = Object.keys(row).find(k => /descripcion|nombre|desc/i.test(k));
          if (numCol && descCol && row[numCol] != null && row[descCol] != null) {
            map[String(row[numCol])] = String(row[descCol]);
          }
        }
      }
    }
    return map;
  }, [workbook]);

  const unitsArr = useMemo(() => {
    const col = structure?.dimensions?.unit;
    if (!sheetData?.length || !col) return [];
    const set = new Set();
    sheetData.forEach(r => {
      const v = r[col];
      if (v != null && v !== '') set.add(String(v));
    });
    // Sort by description name if lookup is available, otherwise by number
    const arr = [...set];
    if (Object.keys(unitsLookup).length > 0) {
      arr.sort((a, b) => (unitsLookup[a] || a).localeCompare(unitsLookup[b] || b));
    } else {
      arr.sort((a, b) => Number(a) - Number(b));
    }
    return arr;
  }, [sheetData, structure, unitsLookup]);

  /* ── Filtered data ── */
  const filteredData = useMemo(() => {
    if (!sheetData?.length) return [];
    let d = sheetData;
    const { year: yCol, month: mCol, region: rCol, unit: uCol, id: iCol } = structure?.dimensions || {};

    // Date range
    if (yCol && (startYear || endYear)) {
      d = d.filter((r) => {
        const y = Number(r[yCol]);
        const m = (mCol && r[mCol] != null) ? Number(r[mCol]) || 1 : 1;
        
        if (startYear) {
          const sy = Number(startYear);
          const sm = startMonth ? Number(startMonth) : 1;
          if (y < sy || (y === sy && m < sm)) return false;
        }
        if (endYear) {
          const ey = Number(endYear);
          const em = endMonth ? Number(endMonth) : 12;
          if (y > ey || (y === ey && m > em)) return false;
        }
        return true;
      });
    }

    if (regionFilter && rCol) {
      d = d.filter((r) => String(r[rCol]) === regionFilter);
    }
    if (unitFilter && uCol) {
      const lo = unitFilter.toLowerCase();
      d = d.filter((r) => {
        const val = String(r[uCol]);
        const name = String(unitsLookup[val] || val).toLowerCase();
        return val.toLowerCase() === lo || name.includes(lo);
      });
    }
    if (fiscalFilter && iCol) {
      const lo = fiscalFilter.toLowerCase();
      d = d.filter((r) => String(r[iCol]).toLowerCase().includes(lo));
    }

    return d;
  }, [sheetData, startYear, startMonth, endYear, endMonth, regionFilter, unitFilter, fiscalFilter, structure, unitsLookup]);

  /* ── KPI data ── */
  const kpis = useMemo(() => {
    if (!structure?.hasGroups || !filteredData.length) return null;

    const result = {};
    for (const [prefix, info] of Object.entries(structure.kpiCols)) {
      const col = info.entran;
      if (!col) continue;

      const sum = filteredData.reduce((s, r) => s + (Number(r[col]) || 0), 0);

      // Trend
      let trend = null;
      if (structure?.dimensions?.year && !(startYear || endYear || startMonth || endMonth)) {
        const sorted = [...filteredData].sort((a, b) =>
          Number(a[structure?.dimensions?.year]) - Number(b[structure?.dimensions?.year])
        );
        const half = Math.floor(sorted.length / 2);
        const firstSum = sorted.slice(0, half).reduce((s, r) => s + (Number(r[col]) || 0), 0);
        const secondSum = sorted.slice(half).reduce((s, r) => s + (Number(r[col]) || 0), 0);
        trend = firstSum > 0 ? ((secondSum - firstSum) / firstSum) * 100 : 0;
      }

      const salenCol = info.salen;
      const salenSum = salenCol ? filteredData.reduce((s, r) => s + (Number(r[salenCol]) || 0), 0) : null;

      result[prefix] = {
        label: info.label,
        entran: sum,
        salen: salenSum,
        trend,
        color: STAGE_COLORS[prefix]?.border || '#6366f1',
      };
    }
    return result;
  }, [structure, filteredData, startYear, startMonth, endYear, endMonth]);

  /* ── Year trend chart ── */
  const yearTrendData = useMemo(() => {
    if (!structure?.hasGroups || !structure?.dimensions?.year) return null;

    const entranCols = Object.values(structure.kpiCols)
      .filter((k) => k.entran)
      .map((k) => k.entran);

    const agg = aggregateByDimension(
      (startYear || endYear) ? sheetData : filteredData,
      structure?.dimensions?.year,
      entranCols,
    );
    if (!agg) return null;

    const prefixEntries = Object.entries(structure.kpiCols).filter(([, k]) => k.entran);
    return {
      labels: agg.labels.map((l) => {
        const n = Number(l);
        return n < 100 ? `20${String(n).padStart(2, '0')}` : String(l);
      }),
      datasets: prefixEntries.map(([prefix, info]) => ({
        label: info.label,
        data: agg.aggregated[info.entran],
        borderColor: STAGE_COLORS[prefix]?.border || '#6366f1',
        backgroundColor: STAGE_COLORS[prefix]?.bg?.replace('.75', '.1') || 'rgba(99,102,241,.1)',
        borderWidth: 2.5,
        pointRadius: 4,
        pointBackgroundColor: STAGE_COLORS[prefix]?.border || '#6366f1',
        tension: 0.35,
        fill: true,
      })),
    };
  }, [structure, filteredData, sheetData, startYear, endYear]);

  const stageCompData = useMemo(() => {
    if (!kpis) return null;
    const entries = Object.entries(kpis);
    return {
      labels: entries.map(([, v]) => v.label),
      datasets: [
        {
          label: 'Entran',
          data: entries.map(([, v]) => v.entran),
          backgroundColor: entries.map(([p]) => STAGE_COLORS[p]?.bg || PALETTE[0]),
          borderColor: entries.map(([p]) => STAGE_COLORS[p]?.border || PALETTE[0]),
          borderWidth: 2,
          borderRadius: 6,
        },
        {
          label: 'Salen',
          data: entries.map(([, v]) => v.salen || 0),
          backgroundColor: entries.map(([p]) => STAGE_COLORS[p]?.bg?.replace('.75', '.35') || PALETTE[0]),
          borderColor: entries.map(([p]) => STAGE_COLORS[p]?.border || PALETTE[0]),
          borderWidth: 1,
          borderRadius: 6,
        },
      ],
    };
  }, [kpis]);

  const topUnitsData = useMemo(() => {
    if (!structure?.dimensions?.unit && !structure?.dimensions?.region) return null;
    const dimCol = structure?.dimensions?.unit || structure?.dimensions?.region;
    const firstEntran = Object.values(structure.kpiCols).find((k) => k.entran)?.entran;
    if (!firstEntran) return null;

    const top = getTopN(filteredData, dimCol, firstEntran, 8);
    if (!top) return null;

    return {
      labels: top.map((t) => t.label.length > 25 ? t.label.slice(0, 25) + '…' : t.label),
      datasets: [{
        label: 'Total',
        data: top.map((t) => t.value),
        backgroundColor: PALETTE.slice(0, top.length),
        borderColor: PALETTE.slice(0, top.length).map((c) => c.replace('.75', '1')),
        borderWidth: 1,
        borderRadius: 6,
      }],
    };
  }, [structure, filteredData]);

  const monthlyData = useMemo(() => {
    if (!structure?.hasGroups || !structure?.dimensions?.month) return null;
    const entranCols = Object.values(structure.kpiCols).filter((k) => k.entran).map((k) => k.entran);
    const agg = aggregateByDimension(filteredData, structure?.dimensions?.month, entranCols);
    if (!agg) return null;

    const monthNames = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];
    const prefixEntries = Object.entries(structure.kpiCols).filter(([, k]) => k.entran);

    return {
      labels: agg.labels.map((l) => {
        const n = parseInt(l);
        return n >= 1 && n <= 12 ? monthNames[n - 1] : l;
      }),
      datasets: prefixEntries.map(([prefix, info]) => ({
        label: info.label,
        data: agg.aggregated[info.entran],
        backgroundColor: STAGE_COLORS[prefix]?.bg || PALETTE[0],
        borderColor: STAGE_COLORS[prefix]?.border || PALETTE[0],
        borderWidth: 2,
        borderRadius: 4,
      })),
    };
  }, [structure, filteredData]);

  const quickStats = useMemo(() => {
    if (!filteredData.length) return null;
    return {
      totalRows: filteredData.length,
      uniqueUnits: structure?.dimensions?.unit
        ? new Set(filteredData.map((r) => r[structure?.dimensions?.unit]).filter(Boolean)).size
        : null,
      uniqueRegions: structure?.dimensions?.region
        ? new Set(filteredData.map((r) => r[structure?.dimensions?.region]).filter(Boolean)).size
        : null,
      yearRange: structure?.dimensions?.year
        ? (() => {
            const vals = filteredData.map((r) => Number(r[structure?.dimensions?.year])).filter((v) => !isNaN(v));
            if (vals.length === 0) return null;
            const min = Math.min(...vals);
            const max = Math.max(...vals);
            return { min: min < 100 ? 2000 + min : min, max: max < 100 ? 2000 + max : max };
          })()
        : null,
    };
  }, [filteredData, structure]);

  const dlPNG = (ref, name) => {
    if (!ref.current) return;
    const a = document.createElement('a');
    a.download = `${name}_${Date.now()}.png`;
    a.href = ref.current.toBase64Image('image/png', 1);
    a.click();
  };

  const barOpts = {
    ...chartBase,
    scales: {
      y: { ticks: { color: '#8b949e', callback: (v) => fmt(v) }, grid: { color: 'rgba(255,255,255,.04)' }, beginAtZero: true },
      x: { ticks: { color: '#8b949e', font: { size: 10 }, maxRotation: 45 }, grid: { display: false } },
    },
  };

  const lineOpts = {
    ...chartBase,
    scales: {
      y: { ticks: { color: '#8b949e', callback: (v) => fmt(v) }, grid: { color: 'rgba(255,255,255,.04)' }, beginAtZero: true },
      x: { ticks: { color: '#8b949e', font: { size: 11 } }, grid: { display: false } },
    },
  };

  /* ── Data Quality summary ── */
  const completeness = useMemo(() => {
    if (!sheetData?.length || !structure) return 0;
    const totalCells = sheetData.length * structure.headers.length;
    let nulls = 0;
    for (const r of sheetData) {
      for (const h of structure.headers) if (r[h] == null || r[h] === '') nulls++;
    }
    return ((totalCells - nulls) / totalCells) * 100;
  }, [sheetData, structure]);

  const statsText = useMemo(() => {
    if (quickStats?.yearRange) {
      return `${filteredData.length.toLocaleString()} registros · ${quickStats.yearRange.min}—${quickStats.yearRange.max}`;
    }
    return `${filteredData.length.toLocaleString()} registros`;
  }, [filteredData, quickStats]);

  if (!structure) {
    return (
      <div className="panel empty-state">
        <BarChart2 size={48} strokeWidth={1} />
        <h3>Sin estructura detectada</h3>
        <p>No se pudo analizar automáticamente esta hoja.</p>
      </div>
    );
  }

  /* ── Pre-compute generic stats (for non-Fiscalia data) ── */
  const genericStats = React.useMemo(() => {
    if (structure.hasGroups || structure.hasSimpleStages) return [];
    return structure.numericCols.slice(0, 6).map((col) => {
      const stats = getColumnStats(filteredData, col);
      return { col, stats };
    }).filter((s) => s.stats);
  }, [structure.hasGroups, structure.hasSimpleStages, structure.numericCols, filteredData]);

  // Generic data view: show column summary when no Fiscalia structure detected
  if (!structure.hasGroups && !structure.hasSimpleStages) {
    return (
      <div className="overview">
        <div className="panel" style={{ marginBottom: '1.5rem', padding: '0.75rem 1.25rem' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
            <Activity size={16} color="var(--accent)" />
            <h3 style={{ margin: 0, fontSize: '1rem' }}>Resumen General</h3>
          </div>
          <p style={{ color: 'var(--text-dim)', fontSize: '0.8rem', margin: '0.5rem 0 0' }}>
            {sheetData.length.toLocaleString()} filas · {structure.headers.length} columnas · {structure.numericCols.length} numéricas
          </p>
        </div>

        <div className="overview-generic">
          <p className="overview-generic__hint">
            Esta hoja tiene <strong>{structure.headers.length}</strong> columnas y <strong>{sheetData.length.toLocaleString()}</strong> filas.
          </p>
          {genericStats.length > 0 && (
            <div className="overview-generic__stats">
              {genericStats.map(({ col, stats }) => (
                <div key={col} className="overview-generic__stat-card">
                  <h4>{col.length > 30 ? col.slice(0, 30) + '…' : col}</h4>
                  <div className="stats" style={{ gridTemplateColumns: '1fr 1fr 1fr 1fr' }}>
                    <div className="stat"><span className="stat__label">Suma</span><span className="stat__val stat__val--accent">{fmt(stats.sum)}</span></div>
                    <div className="stat"><span className="stat__label">Promedio</span><span className="stat__val">{fmt(stats.mean)}</span></div>
                    <div className="stat"><span className="stat__label">Máx</span><span className="stat__val">{fmt(stats.max)}</span></div>
                    <div className="stat"><span className="stat__label">Min</span><span className="stat__val">{fmt(stats.min)}</span></div>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>

        {/* Show all column names for reference */}
        <div className="panel" style={{ marginTop: '1.5rem' }}>
          <h4 style={{ margin: '0 0 0.75rem', fontSize: '0.85rem', color: 'var(--text-dim)' }}>Columnas detectadas</h4>
          <div style={{ display: 'flex', flexWrap: 'wrap', gap: '0.4rem' }}>
            {structure.headers.map((h) => (
              <span key={h} className="badge badge--dim" style={{ fontSize: '0.7rem' }}>
                {h}
                {structure.numericCols.includes(h) && <span style={{ color: 'var(--green)', marginLeft: '0.25rem' }}>●</span>}
              </span>
            ))}
          </div>
        </div>
      </div>
    );
  }

  const handleExport = () => {
    const ws = XLSX.utils.json_to_sheet(filteredData, { header: structure.headers });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Resumen Filtrado');
    XLSX.writeFile(wb, `resumen_filtrado_${Date.now()}.xlsx`, { compression: true });
  };

  return (
    <div className="overview">
      {/* ── Quality Bar ── */}
      <div className="panel" style={{ marginBottom: '1.5rem', padding: '0.75rem 1.25rem' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', marginBottom: '0.4rem' }}>
          <Activity size={13} color="var(--green)" />
          <span className="overview-filters__label" style={{ fontSize: '0.65rem' }}>CALIDAD DE DATOS</span>
        </div>
        <div className="quality-bar" style={{ height: '6px' }}>
          <div className="quality-bar__fill" style={{ width: `${completeness}%`, background: 'var(--green)' }} />
        </div>
        <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '0.7rem', color: 'var(--text-dim)', marginTop: '0.4rem' }}>
           <span>{completeness.toFixed(0)}% completo</span>
           <span>{filteredData.length.toLocaleString()}</span>
        </div>
      </div>

      {/* ── Filters ── */}
      <div className="overview-filters">
        <div className="overview-filters__row">
          <FilterIcon size={16} color="var(--text-dim)" />
          
          {/* Desde */}
          <div className="filter-group">
            <span className="overview-filters__label">Desde:</span>
            <div className="filter-box">
               <select className="sel sel--compact" value={startYear} onChange={(e) => setStartYear(e.target.value)}>
                <option value="">Año…</option>
                {yearsArr.map((y) => <option key={y} value={String(y)}>{Number(y) < 100 ? `20${String(y).padStart(2,'0')}` : y}</option>)}
              </select>
              <span className="filter-box__sep"></span>
              <select className="sel sel--compact" value={startMonth} onChange={(e) => setStartMonth(e.target.value)}>
                <option value="">Mes…</option>
                {months.map((m) => <option key={m.v} value={m.v}>{m.l}</option>)}
              </select>
            </div>
          </div>

          {/* Hasta */}
          <div className="filter-group">
            <span className="overview-filters__label">Hasta:</span>
            <div className="filter-box">
               <select className="sel sel--compact" value={endYear} onChange={(e) => setEndYear(e.target.value)}>
                <option value="">Año…</option>
                {yearsArr.map((y) => <option key={y} value={String(y)}>{Number(y) < 100 ? `20${String(y).padStart(2,'0')}` : y}</option>)}
              </select>
              <span className="filter-box__sep"></span>
              <select className="sel sel--compact" value={endMonth} onChange={(e) => setEndMonth(e.target.value)}>
                <option value="">Mes…</option>
                {months.map((m) => <option key={m.v} value={m.v}>{m.l}</option>)}
              </select>
            </div>
          </div>
        </div>

        <div className="overview-filters__row">
          {/* Seccional */}
          <select className="sel sel--compact" value={regionFilter} onChange={(e) => setRegionFilter(e.target.value)} style={{ paddingLeft: 0, fontWeight: 600 }}>
            <option value="">Todas las seccionales</option>
            {regionsArr.map((r) => <option key={r} value={String(r)}>{r}</option>)}
          </select>

          {/* Unidad */}
          <div className="filter-group">
             <span className="overview-filters__label">Unidad:</span>
             <input 
               list="unidades-list"
               value={unitFilter} 
               onChange={(e) => setUnitFilter(e.target.value)} 
               className="filter__input" 
               placeholder="(Todas)"
               style={{ background: 'transparent', border: 'none', borderBottom: '1px solid var(--border)', borderRadius: 0, paddingLeft: 0, width: '200px' }}
             />
             <datalist id="unidades-list">
               {unitsArr.map((u) => (
                 <option key={u} value={unitsLookup[u] || String(u)} />
               ))}
             </datalist>
          </div>

          {/* Nro Fiscal */}
          <div className="filter-group">
             <span className="overview-filters__label">Nro fiscal:</span>
             <select 
               value={fiscalFilter} 
               onChange={(e) => setFiscalFilter(e.target.value)} 
               className="sel sel--compact" 
               style={{ paddingLeft: 0, width: '80px' }}
             >
               <option value="">(Todos)</option>
               <option value="1">1</option>
               <option value="3">3</option>
             </select>
          </div>

          <div style={{ marginLeft: 'auto', display: 'flex', gap: '0.5rem', alignItems: 'center' }}>
            {/* Export Button */}
            <button className="btn btn--export" onClick={handleExport} style={{ padding: '0.2rem 0.6rem', fontSize: '0.72rem' }}>
              <FileSpreadsheet size={13} /> Excel
            </button>

            {/* Stats Summary */}
            <div className="overview-filters__stats">
               {statsText}
            </div>
          </div>
        </div>
      </div>

      {/* ── KPI Cards ── */}
      {kpis && Object.keys(kpis).length > 0 && (
        <div className="kpi-grid">
          {Object.entries(kpis).map(([prefix, data]) => (
            <KpiCard
              key={prefix}
              label={data.label}
              value={data.entran}
              color={data.color}
              icon={
                prefix === 'II_' ? BarChart2 :
                prefix === 'III_' ? Users :
                prefix === 'IV_' ? MapPin :
                Calendar
              }
              detail={data.salen != null ? `Salen: ${fmt(data.salen)}` : null}
              trend={data.trend}
            />
          ))}
        </div>
      )}

      {/* ── Funnel summary ── */}
      {kpis && Object.keys(kpis).length > 1 && (
        <div className="funnel">
          {Object.entries(kpis).map(([prefix, data], i, arr) => (
            <React.Fragment key={prefix}>
              <div className="funnel__stage" style={{ '--stage-color': data.color }}>
                <span className="funnel__label">{data.label}</span>
                <span className="funnel__value">{fmt(data.entran)}</span>
              </div>
              {i < arr.length - 1 && <ArrowRight size={18} className="funnel__arrow" />}
            </React.Fragment>
          ))}
        </div>
      )}

      {/* ── Charts grid ── */}
      <div className="overview-charts">
        {yearTrendData && (
          <ChartCard title="Evolución por Año" onDownload={() => dlPNG(yearChartRef, 'tendencia_anual')}>
            <Line ref={yearChartRef} data={yearTrendData} options={lineOpts} />
          </ChartCard>
        )}
        {stageCompData && (
          <ChartCard title="Comparación por Etapa" onDownload={() => dlPNG(stageChartRef, 'comparacion_etapas')}>
            <Bar ref={stageChartRef} data={stageCompData} options={barOpts} />
          </ChartCard>
        )}
        {monthlyData && (
          <ChartCard title="Desglose Mensual">
            <Bar data={monthlyData} options={{ ...barOpts, plugins: { ...barOpts.plugins, legend: { ...barOpts.plugins.legend, position: 'bottom' } } }} />
          </ChartCard>
        )}
        {topUnitsData && (
          <ChartCard title="Top Unidades por Volumen" onDownload={() => dlPNG(topChartRef, 'top_unidades')}>
            <Bar ref={topChartRef} data={topUnitsData} options={{ ...barOpts, indexAxis: 'y', scales: { ...barOpts.scales, x: { ...barOpts.scales.x, grid: { color: 'rgba(255,255,255,.04)' } }, y: { ...barOpts.scales.y, grid: { display: false }, ticks: { color: '#c9d1d9', font: { size: 10 } } } } }} />
          </ChartCard>
        )}
      </div>

      {!structure.hasGroups && (
        <div className="overview-generic">
          <p className="overview-generic__hint">
            Esta hoja tiene {structure.headers.length} columnas y {filteredData.length.toLocaleString()} filas.
          </p>
          {structure.numericCols.length > 0 && (
            <div className="overview-generic__stats">
              {structure.numericCols.slice(0, 6).map((col) => {
                const stats = getColumnStats(filteredData, col);
                if (!stats) return null;
                return (
                  <div key={col} className="overview-generic__stat-card">
                    <h4>{col.length > 30 ? col.slice(0, 30) + '…' : col}</h4>
                    <div className="stats" style={{ gridTemplateColumns: '1fr 1fr 1fr' }}>
                      <div className="stat"><span className="stat__label">Suma</span><span className="stat__val stat__val--accent">{fmt(stats.sum)}</span></div>
                      <div className="stat"><span className="stat__label">Promedio</span><span className="stat__val">{fmt(stats.mean)}</span></div>
                      <div className="stat"><span className="stat__label">Máx</span><span className="stat__val">{fmt(stats.max)}</span></div>
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>
      )}
    </div>
  );
}
