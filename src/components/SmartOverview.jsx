import React, { useMemo, useState, useRef } from 'react';
import * as XLSX from 'xlsx';
import {
  Chart as ChartJS, CategoryScale, LinearScale, BarElement, Title, Tooltip,
  Legend, PointElement, LineElement, ArcElement, Filler,
} from 'chart.js';
import { Bar, Line, Doughnut } from 'react-chartjs-2';
import {
  analyzeSheetStructure, aggregateByDimension, getTopN, getColumnStats,
  getUniqueValues,
} from '../utils/excelParser';
import {
  TrendingUp, TrendingDown, Minus, ArrowRight, Calendar,
  MapPin, Users, BarChart2, Download, Filter as FilterIcon, FileSpreadsheet, Activity,
  Database, Hash, Type, Search, X,
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

/* ── Generic column card ── */
function ColumnStatCard({ col, stats, color }) {
  return (
    <div className="overview-generic__stat-card" style={{ borderLeft: `3px solid ${color}` }}>
      <h4>{col.length > 35 ? col.slice(0, 35) + '…' : col}</h4>
      <div className="stats" style={{ gridTemplateColumns: '1fr 1fr 1fr 1fr' }}>
        <div className="stat"><span className="stat__label">Suma</span><span className="stat__val stat__val--accent">{fmt(stats.sum)}</span></div>
        <div className="stat"><span className="stat__label">Prom</span><span className="stat__val">{fmt(stats.mean)}</span></div>
        <div className="stat"><span className="stat__label">Máx</span><span className="stat__val">{fmt(stats.max)}</span></div>
        <div className="stat"><span className="stat__label">Min</span><span className="stat__val">{fmt(stats.min)}</span></div>
      </div>
      <div className="stats" style={{ gridTemplateColumns: '1fr 1fr 1fr 1fr', marginTop: '0.4rem' }}>
        <div className="stat"><span className="stat__label">Mediana</span><span className="stat__val">{fmt(stats.median)}</span></div>
        <div className="stat"><span className="stat__label">σ Desv.</span><span className="stat__val">{fmt(stats.stdDev)}</span></div>
        <div className="stat"><span className="stat__label">CV %</span><span className="stat__val">{fmt(stats.cv)}%</span></div>
        <div className="stat"><span className="stat__label">Registros</span><span className="stat__val">{fmt(stats.count)}</span></div>
      </div>
    </div>
  );
}

export default function SmartOverview({ sheetData, workbook }) {
  const [activeFilters, setActiveFilters] = useState({});
  const [searchText, setSearchText] = useState('');

  const yearChartRef = useRef(null);
  const stageChartRef = useRef(null);
  const topChartRef = useRef(null);
  const genericChartRef = useRef(null);

  /* ── Reset on data change ── */
  React.useEffect(() => {
    setActiveFilters({});
    setSearchText('');
  }, [sheetData]);

  /* ── Analyze structure ── */
  const structure = useMemo(() => analyzeSheetStructure(sheetData), [sheetData]);

  /* ── Filtered data with dynamic filters ── */
  const filteredData = useMemo(() => {
    if (!sheetData?.length) return [];
    let d = sheetData;

    // Apply dynamic column filters
    for (const [col, val] of Object.entries(activeFilters)) {
      if (!val || val === '') continue;
      const lo = String(val).toLowerCase();
      d = d.filter((r) => {
        const cell = r[col];
        if (cell == null || cell === '') return false;
        return String(cell).toLowerCase().includes(lo);
      });
    }

    // Global search
    if (searchText) {
      const lo = searchText.toLowerCase();
      d = d.filter((r) =>
        Object.values(r).some((v) => v != null && String(v).toLowerCase().includes(lo))
      );
    }

    return d;
  }, [sheetData, activeFilters, searchText]);

  /* ── Detect filterable columns ── */
  const filterableCols = useMemo(() => {
    if (!structure?.labelCols) return [];
    return structure.labelCols.filter((col) => {
      const vals = getUniqueValues(sheetData, col);
      return vals.length > 1 && vals.length <= 200; // reasonable cardinality
    }).slice(0, 6);
  }, [structure, sheetData]);

  /* ── Generic charts for non-Fiscalia data ── */
  const genericTopData = useMemo(() => {
    if (!structure || structure.hasGroups || structure.hasSimpleStages) return null;
    if (!structure.numericCols.length || !structure.labelCols.length) return null;

    const numCol = structure.numericCols[0];
    const labelCol = structure.labelCols[0];
    const top = getTopN(filteredData, labelCol, numCol, 10);
    if (!top) return null;

    return {
      labels: top.map((t) => t.label.length > 25 ? t.label.slice(0, 25) + '…' : t.label),
      datasets: [{
        label: numCol,
        data: top.map((t) => t.value),
        backgroundColor: PALETTE.slice(0, top.length),
        borderColor: PALETTE.slice(0, top.length).map((c) => c.replace('.75', '1')),
        borderWidth: 1,
        borderRadius: 6,
      }],
    };
  }, [structure, filteredData]);

  const genericDistribution = useMemo(() => {
    if (!structure || structure.hasGroups || structure.hasSimpleStages) return null;
    if (!structure.numericCols.length) return null;

    const numCol = structure.numericCols[0];
    const stats = getColumnStats(filteredData, numCol);
    if (!stats) return null;

    // Create histogram bins
    const bins = 8;
    const min = stats.min;
    const max = stats.max;
    const range = max - min || 1;
    const binSize = range / bins;
    const counts = new Array(bins).fill(0);

    filteredData.forEach((r) => {
      const v = Number(r[numCol]);
      if (!isNaN(v)) {
        const idx = Math.min(Math.floor((v - min) / binSize), bins - 1);
        counts[idx]++;
      }
    });

    const labels = counts.map((_, i) => {
      const from = min + i * binSize;
      const to = min + (i + 1) * binSize;
      return `${fmt(from)}–${fmt(to)}`;
    });

    return {
      labels,
      datasets: [{
        label: 'Frecuencia',
        data: counts,
        backgroundColor: 'rgba(99,102,241,.6)',
        borderColor: 'rgba(99,102,241,1)',
        borderWidth: 1,
        borderRadius: 4,
      }],
    };
  }, [structure, filteredData]);

  /* ── Fiscalia-specific data ── */
  const months = useMemo(() => [
    { v: '1',  l: 'Enero' }, { v: '2',  l: 'Febrero' }, { v: '3',  l: 'Marzo' },
    { v: '4',  l: 'Abril' }, { v: '5',  l: 'Mayo' },    { v: '6',  l: 'Junio' },
    { v: '7',  l: 'Julio' }, { v: '8',  l: 'Agosto' },  { v: '9',  l: 'Septiembre' },
    { v: '10', l: 'Octubre' },{ v: '11', l: 'Noviembre' },{ v: '12', l: 'Diciembre' },
  ], []);

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
    const map = {};
    if (workbook?.data) {
      const unidadesKey = workbook.sheetNames?.find(s => /unidad/i.test(s));
      if (unidadesKey && workbook.data[unidadesKey]) {
        for (const row of workbook.data[unidadesKey]) {
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
    const arr = [...set];
    if (Object.keys(unitsLookup).length > 0) {
      arr.sort((a, b) => (unitsLookup[a] || a).localeCompare(unitsLookup[b] || b));
    } else {
      arr.sort((a, b) => Number(a) - Number(b));
    }
    return arr;
  }, [sheetData, structure, unitsLookup]);

  /* ── Fiscalia filtered data (with date range) ── */
  const fiscaliaFilteredData = useMemo(() => {
    if (!sheetData?.length) return [];
    let d = sheetData;
    const { year: yCol, month: mCol, region: rCol, unit: uCol, id: iCol } = structure?.dimensions || {};

    if (yCol && (activeFilters._startYear || activeFilters._endYear)) {
      d = d.filter((r) => {
        const y = Number(r[yCol]);
        const m = (mCol && r[mCol] != null) ? Number(r[mCol]) || 1 : 1;
        if (activeFilters._startYear) {
          const sy = Number(activeFilters._startYear);
          const sm = activeFilters._startMonth ? Number(activeFilters._startMonth) : 1;
          if (y < sy || (y === sy && m < sm)) return false;
        }
        if (activeFilters._endYear) {
          const ey = Number(activeFilters._endYear);
          const em = activeFilters._endMonth ? Number(activeFilters._endMonth) : 12;
          if (y > ey || (y === ey && m > em)) return false;
        }
        return true;
      });
    }

    if (activeFilters._region && rCol) {
      d = d.filter((r) => String(r[rCol]) === activeFilters._region);
    }
    if (activeFilters._unit && uCol) {
      const lo = activeFilters._unit.toLowerCase();
      d = d.filter((r) => {
        const val = String(r[uCol]);
        const name = String(unitsLookup[val] || val).toLowerCase();
        return val.toLowerCase() === lo || name.includes(lo);
      });
    }
    if (activeFilters._fiscal && iCol) {
      const lo = activeFilters._fiscal.toLowerCase();
      d = d.filter((r) => String(r[iCol]).toLowerCase().includes(lo));
    }

    // Also apply global search
    if (searchText) {
      const lo = searchText.toLowerCase();
      d = d.filter((r) =>
        Object.values(r).some((v) => v != null && String(v).toLowerCase().includes(lo))
      );
    }

    return d;
  }, [sheetData, structure, activeFilters, searchText, unitsLookup]);

  /* ── KPI data ── */
  const kpis = useMemo(() => {
    if (!structure?.hasGroups || !fiscaliaFilteredData.length) return null;

    const result = {};
    for (const [prefix, info] of Object.entries(structure.kpiCols)) {
      const col = info.entran;
      if (!col) continue;

      const sum = fiscaliaFilteredData.reduce((s, r) => s + (Number(r[col]) || 0), 0);

      let trend = null;
      if (structure?.dimensions?.year && !(activeFilters._startYear || activeFilters._endYear)) {
        const sorted = [...fiscaliaFilteredData].sort((a, b) =>
          Number(a[structure?.dimensions?.year]) - Number(b[structure?.dimensions?.year])
        );
        const half = Math.floor(sorted.length / 2);
        const firstSum = sorted.slice(0, half).reduce((s, r) => s + (Number(r[col]) || 0), 0);
        const secondSum = sorted.slice(half).reduce((s, r) => s + (Number(r[col]) || 0), 0);
        trend = firstSum > 0 ? ((secondSum - firstSum) / firstSum) * 100 : 0;
      }

      const salenCol = info.salen;
      const salenSum = salenCol ? fiscaliaFilteredData.reduce((s, r) => s + (Number(r[salenCol]) || 0), 0) : null;

      result[prefix] = {
        label: info.label,
        entran: sum,
        salen: salenSum,
        trend,
        color: STAGE_COLORS[prefix]?.border || '#6366f1',
      };
    }
    return result;
  }, [structure, fiscaliaFilteredData, activeFilters]);

  /* ── Charts ── */
  const yearTrendData = useMemo(() => {
    if (!structure?.hasGroups || !structure?.dimensions?.year) return null;

    const entranCols = Object.values(structure.kpiCols)
      .filter((k) => k.entran)
      .map((k) => k.entran);

    const agg = aggregateByDimension(
      (activeFilters._startYear || activeFilters._endYear) ? sheetData : fiscaliaFilteredData,
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
  }, [structure, fiscaliaFilteredData, sheetData, activeFilters]);

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

    const top = getTopN(fiscaliaFilteredData, dimCol, firstEntran, 8);
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
  }, [structure, fiscaliaFilteredData]);

  const monthlyData = useMemo(() => {
    if (!structure?.hasGroups || !structure?.dimensions?.month) return null;
    const entranCols = Object.values(structure.kpiCols).filter((k) => k.entran).map((k) => k.entran);
    const agg = aggregateByDimension(fiscaliaFilteredData, structure?.dimensions?.month, entranCols);
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
  }, [structure, fiscaliaFilteredData]);

  const quickStats = useMemo(() => {
    const data = structure?.hasGroups ? fiscaliaFilteredData : filteredData;
    if (!data.length) return null;
    return {
      totalRows: data.length,
      uniqueUnits: structure?.dimensions?.unit
        ? new Set(data.map((r) => r[structure?.dimensions?.unit]).filter(Boolean)).size
        : null,
      uniqueRegions: structure?.dimensions?.region
        ? new Set(data.map((r) => r[structure?.dimensions?.region]).filter(Boolean)).size
        : null,
      yearRange: structure?.dimensions?.year
        ? (() => {
            const vals = data.map((r) => Number(r[structure?.dimensions?.year])).filter((v) => !isNaN(v));
            if (vals.length === 0) return null;
            let min = Infinity, max = -Infinity;
            for (const v of vals) {
              if (v < min) min = v;
              if (v > max) max = v;
            }
            return { min: min < 100 ? 2000 + min : min, max: max < 100 ? 2000 + max : max };
          })()
        : null,
    };
  }, [filteredData, fiscaliaFilteredData, structure]);

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
    const data = structure?.hasGroups ? fiscaliaFilteredData : filteredData;
    if (quickStats?.yearRange) {
      return `${data.length.toLocaleString()} registros · ${quickStats.yearRange.min}—${quickStats.yearRange.max}`;
    }
    return `${data.length.toLocaleString()} registros`;
  }, [filteredData, fiscaliaFilteredData, quickStats, structure]);

  if (!structure) {
    return (
      <div className="panel empty-state">
        <BarChart2 size={48} strokeWidth={1} />
        <h3>Sin estructura detectada</h3>
        <p>No se pudo analizar automáticamente esta hoja.</p>
      </div>
    );
  }

  /* ── Pre-compute generic stats ── */
  const genericStats = useMemo(() => {
    if (structure.hasGroups || structure.hasSimpleStages) return [];
    return structure.numericCols.slice(0, 8).map((col) => {
      const stats = getColumnStats(filteredData, col);
      return { col, stats };
    }).filter((s) => s.stats);
  }, [structure.hasGroups, structure.hasSimpleStages, structure.numericCols, filteredData]);

  const isFiscalia = structure.hasGroups || structure.hasSimpleStages;
  const displayData = isFiscalia ? fiscaliaFilteredData : filteredData;

  const handleExport = () => {
    const ws = XLSX.utils.json_to_sheet(displayData, { header: structure.headers });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Resumen Filtrado');
    XLSX.writeFile(wb, `resumen_filtrado_${Date.now()}.xlsx`, { compression: true });
  };

  const setFilter = (key, val) => {
    setActiveFilters((prev) => ({ ...prev, [key]: val }));
  };

  const clearFilter = (key) => {
    setActiveFilters((prev) => {
      const next = { ...prev };
      delete next[key];
      return next;
    });
  };

  const clearAllFilters = () => {
    setActiveFilters({});
    setSearchText('');
  };

  const hasActiveFilters = Object.keys(activeFilters).some((k) => activeFilters[k] && activeFilters[k] !== '') || searchText;

  return (
    <div className="overview">
      {/* ── Header ── */}
      <div className="panel" style={{ marginBottom: '1.5rem', padding: '0.75rem 1.25rem' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', justifyContent: 'space-between' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
            <Activity size={16} color="var(--accent)" />
            <h3 style={{ margin: 0, fontSize: '1rem' }}>
              {isFiscalia ? 'Resumen Fiscalía' : 'Resumen General'}
            </h3>
          </div>
          <span style={{ color: 'var(--text-dim)', fontSize: '0.75rem' }}>
            {statsText} · {completeness.toFixed(0)}% completo
          </span>
        </div>
        <div className="quality-bar" style={{ height: '4px', marginTop: '0.5rem' }}>
          <div className="quality-bar__fill" style={{ width: `${completeness}%`, background: 'var(--green)' }} />
        </div>
      </div>

      {/* ── Search & Filters ── */}
      <div className="panel" style={{ marginBottom: '1.5rem', padding: '0.75rem 1.25rem' }}>
        <div style={{ display: 'flex', gap: '0.75rem', alignItems: 'center', flexWrap: 'wrap' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '0.4rem', flex: '1 1 200px', minWidth: '200px' }}>
            <Search size={14} color="var(--text-dim)" />
            <input
              type="search"
              className="filter__input"
              placeholder="Buscar en todos los campos..."
              value={searchText}
              onChange={(e) => setSearchText(e.target.value)}
              style={{ flex: 1 }}
            />
          </div>

          {/* Fiscalia-specific filters */}
          {isFiscalia && (
            <>
              {yearsArr.length > 0 && (
                <>
                  <select className="sel sel--compact" value={activeFilters._startYear || ''} onChange={(e) => setFilter('_startYear', e.target.value)}>
                    <option value="">Desde año</option>
                    {yearsArr.map((y) => <option key={y} value={String(y)}>{Number(y) < 100 ? `20${String(y).padStart(2,'0')}` : y}</option>)}
                  </select>
                  <select className="sel sel--compact" value={activeFilters._endYear || ''} onChange={(e) => setFilter('_endYear', e.target.value)}>
                    <option value="">Hasta año</option>
                    {yearsArr.map((y) => <option key={y} value={String(y)}>{Number(y) < 100 ? `20${String(y).padStart(2,'0')}` : y}</option>)}
                  </select>
                </>
              )}
              {regionsArr.length > 0 && (
                <select className="sel sel--compact" value={activeFilters._region || ''} onChange={(e) => setFilter('_region', e.target.value)}>
                  <option value="">Todas las seccionales</option>
                  {regionsArr.map((r) => <option key={r} value={String(r)}>{r}</option>)}
                </select>
              )}
            </>
          )}

          {/* Generic column filters */}
          {!isFiscalia && filterableCols.map((col) => (
            <select
              key={col}
              className="sel sel--compact"
              value={activeFilters[col] || ''}
              onChange={(e) => setFilter(col, e.target.value)}
            >
              <option value="">{col}</option>
              {getUniqueValues(sheetData, col).map((v) => (
                <option key={String(v)} value={String(v)}>{String(v).slice(0, 30)}</option>
              ))}
            </select>
          ))}

          <button className="btn btn--export" onClick={handleExport} style={{ marginLeft: 'auto' }}>
            <FileSpreadsheet size={13} /> Excel
          </button>
        </div>

        {/* Active filter chips */}
        {hasActiveFilters && (
          <div style={{ display: 'flex', gap: '0.4rem', marginTop: '0.5rem', flexWrap: 'wrap' }}>
            {searchText && (
              <span className="badge badge--blue" style={{ display: 'flex', alignItems: 'center', gap: '0.25rem' }}>
                Buscar: {searchText}
                <button onClick={() => setSearchText('')} style={{ background: 'none', border: 'none', color: 'inherit', cursor: 'pointer', padding: 0 }}><X size={10} /></button>
              </span>
            )}
            {Object.entries(activeFilters).filter(([, v]) => v && v !== '').map(([k, v]) => (
              <span key={k} className="badge badge--purple" style={{ display: 'flex', alignItems: 'center', gap: '0.25rem' }}>
                {k.startsWith('_') ? k.replace('_', '') : k}: {String(v).slice(0, 20)}
                <button onClick={() => clearFilter(k)} style={{ background: 'none', border: 'none', color: 'inherit', cursor: 'pointer', padding: 0 }}><X size={10} /></button>
              </span>
            ))}
            <button className="btn btn--ghost btn--sm" onClick={clearAllFilters}><X size={12} /> Limpiar</button>
          </div>
        )}
      </div>

      {/* ── GENERIC VIEW ── */}
      {!isFiscalia && (
        <>
          {/* Generic KPIs */}
          <div className="kpi-grid" style={{ gridTemplateColumns: 'repeat(auto-fit, minmax(160px, 1fr))' }}>
            <KpiCard label="Filas" value={displayData.length} icon={Database} color="#6366f1" />
            <KpiCard label="Columnas" value={structure.headers.length} icon={Type} color="#a855f7" />
            <KpiCard label="Numéricas" value={structure.numericCols.length} icon={Hash} color="#22c55e" />
            <KpiCard label="Categóricas" value={structure.labelCols.length} icon={MapPin} color="#f59e0b" />
          </div>

          {/* Generic charts */}
          <div className="overview-charts">
            {genericTopData && (
              <ChartCard title={`Top ${structure.labelCols[0]} por ${structure.numericCols[0]}`} onDownload={() => dlPNG(genericChartRef, 'top_valores')}>
                <Bar ref={genericChartRef} data={genericTopData} options={{ ...barOpts, indexAxis: 'y', scales: { ...barOpts.scales, x: { ...barOpts.scales.x, grid: { color: 'rgba(255,255,255,.04)' } }, y: { ...barOpts.scales.y, grid: { display: false }, ticks: { color: '#c9d1d9', font: { size: 10 } } } } }} />
              </ChartCard>
            )}
            {genericDistribution && (
              <ChartCard title={`Distribución de ${structure.numericCols[0]}`}>
                <Bar data={genericDistribution} options={barOpts} />
              </ChartCard>
            )}
          </div>

          {/* Generic column stats */}
          {genericStats.length > 0 && (
            <div className="overview-generic__stats">
              {genericStats.map(({ col, stats }, i) => (
                <ColumnStatCard key={col} col={col} stats={stats} color={PALETTE[i % PALETTE.length]} />
              ))}
            </div>
          )}

          {/* Column list */}
          <div className="panel" style={{ marginTop: '1.5rem' }}>
            <h4 style={{ margin: '0 0 0.75rem', fontSize: '0.85rem', color: 'var(--text-dim)' }}>
              Columnas detectadas ({structure.headers.length})
            </h4>
            <div style={{ display: 'flex', flexWrap: 'wrap', gap: '0.4rem' }}>
              {structure.headers.map((h) => (
                <span key={h} className="badge badge--dim" style={{ fontSize: '0.7rem' }}>
                  {h}
                  {structure.numericCols.includes(h) && <span style={{ color: 'var(--green)', marginLeft: '0.25rem' }}>● num</span>}
                  {structure.dateCols?.includes(h) && <span style={{ color: 'var(--accent)', marginLeft: '0.25rem' }}>● fecha</span>}
                </span>
              ))}
            </div>
          </div>
        </>
      )}

      {/* ── FISCALIA VIEW ── */}
      {isFiscalia && (
        <>
          {/* KPI Cards */}
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

          {/* Funnel */}
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

          {/* Charts */}
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
        </>
      )}
    </div>
  );
}
