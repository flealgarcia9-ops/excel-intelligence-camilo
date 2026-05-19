import React, { useState, useCallback, useRef, useEffect } from 'react';
import {
  UploadCloud, FileSpreadsheet, RefreshCw, X, Shield,
  Zap, BarChart3, AlertTriangle, CheckCircle2, Clock,
  Loader2, FileUp, Sun, Moon,
} from 'lucide-react';
import { useExcelParser } from './utils/useExcelParser';
import AnalysisDashboard from './components/AnalysisDashboard';

/* ── Toast notification ─────────────────────────── */
function Toast({ message, type = 'error', onClose }) {
  useEffect(() => {
    const t = setTimeout(onClose, 5000);
    return () => clearTimeout(t);
  }, [onClose]);

  const Icon = type === 'error' ? AlertTriangle : CheckCircle2;
  return (
    <div className={`toast toast--${type}`} role="alert">
      <Icon size={16} />
      <span>{message}</span>
      <button className="toast__close" onClick={onClose} aria-label="Cerrar"><X size={14} /></button>
    </div>
  );
}

/* ── Feature card ───────────────────────────────── */
// eslint-disable-next-line no-unused-vars
function FeatureCard({ icon: Icon, title, desc }) {
  return (
    <div className="feature-card">
      <div className="feature-card__icon"><Icon size={20} /></div>
      <h3 className="feature-card__title">{title}</h3>
      <p className="feature-card__desc">{desc}</p>
    </div>
  );
}

/* ── Progress bar ───────────────────────────────── */
function ProgressBar({ progress, stage }) {
  const pct = Math.min(Math.round(progress), 100);
  return (
    <div className="progress-wrapper">
      <div className="progress-meta">
        <span className="progress-stage">{stage}</span>
        <span className="progress-pct">{pct}%</span>
      </div>
      <div className="progress-bar">
        <div className="progress-bar__fill" style={{ width: `${pct}%` }} />
      </div>
    </div>
  );
}

/* ── Loading stage text based on progress ───────── */
function getStageText(progress) {
  if (progress < 15) return 'Leyendo archivo…';
  if (progress < 40) return 'Parseando hojas…';
  if (progress < 70) return 'Extrayendo datos…';
  if (progress < 90) return 'Analizando columnas…';
  return 'Finalizando…';
}

/* ═════════════════════════════════════════════════════
   APP
   ═════════════════════════════════════════════════════ */
export default function App() {
  const [fileInfo, setFileInfo] = useState(null);
  const [workbook, setWorkbook] = useState(null);
  const [toast, setToast] = useState(null);
  const [activeSheet, setActiveSheet] = useState(null);
  const [dragging, setDragging] = useState(false);
  const [darkMode, setDarkMode] = useState(() => {
    const saved = localStorage.getItem('ei-dark-mode');
    return saved ? saved === 'true' : true;
  });
  const inputRef = useRef(null);

  /* ── Dark mode sync ── */
  useEffect(() => {
    document.documentElement.classList.toggle('light-mode', !darkMode);
    localStorage.setItem('ei-dark-mode', darkMode);
  }, [darkMode]);

  const { parseFile, loading, progress } = useExcelParser();

  const showToast = useCallback((message, type = 'error') => setToast({ message, type, key: Date.now() }), []);

  /* ── Keyboard: Ctrl+O / Esc ──────────────────── */
  useEffect(() => {
    const handler = (e) => {
      if ((e.ctrlKey || e.metaKey) && e.key === 'o') {
        e.preventDefault();
        if (!workbook && !loading) inputRef.current?.click();
      }
      if (e.key === 'Escape' && workbook) reset();
    };
    window.addEventListener('keydown', handler);
    return () => window.removeEventListener('keydown', handler);
  }, [workbook, loading]); // eslint-disable-line

  /* ── Process file ────────────────────────────── */
  const processFile = useCallback(async (file) => {
    if (!file) return;
    if (!/\.(xlsx|xls|csv)$/i.test(file.name)) {
      showToast('Formato no soportado. Use archivos .xlsx, .xls o .csv');
      return;
    }
    if (file.size > 100 * 1024 * 1024) {
      showToast('El archivo excede el límite de 100 MB.');
      return;
    }

    setToast(null);
    setWorkbook(null);

    const sizeMB = (file.size / 1024 / 1024).toFixed(2);
    const ext = file.name.split('.').pop().toUpperCase();
    setFileInfo({ name: file.name, size: `${sizeMB} MB`, ext, lastModified: new Date(file.lastModified) });

    try {
      const data = await parseFile(file);

      setWorkbook(data);
      const firstSheet = data?.sheetNames?.[0] ?? null;
      setActiveSheet(firstSheet);

      let totalRows = 0;
      try {
        totalRows = Object.values(data.data || {}).reduce((s, d) => s + (Array.isArray(d) ? d.length : 0), 0);
      } catch (e) { console.warn(e); }

      showToast(`${data.sheetNames.length} hoja(s) · ${totalRows.toLocaleString()} registros cargados`, 'success');
    } catch (err) {
      console.error(err);
      showToast('Error al procesar el archivo. Verifique que no esté corrupto.');
    }
  }, [showToast, parseFile]);

  const onDragOver = useCallback((e) => { e.preventDefault(); setDragging(true); }, []);
  const onDragLeave = useCallback((e) => { e.preventDefault(); setDragging(false); }, []);
  const onDrop = useCallback((e) => { e.preventDefault(); setDragging(false); processFile(e.dataTransfer.files?.[0]); }, [processFile]);
  const onSelect = useCallback((e) => { processFile(e.target.files?.[0]); }, [processFile]);
  const reset = useCallback(() => {
    setFileInfo(null); setWorkbook(null); setActiveSheet(null); setToast(null);
    if (inputRef.current) inputRef.current.value = '';
  }, []);

  /* ── Render ─────────────────────────────────── */
  return (
    <div className="app">
      {/* Toasts */}
      <div className="toast-container">
        {toast && <Toast key={toast.key} message={toast.message} type={toast.type} onClose={() => setToast(null)} />}
      </div>

      {/* Header */}
      <header className="header">
        <div className="header__brand">
          <div className="header__logo"><BarChart3 size={24} /></div>
          <div>
            <h1 className="header__title">Excel Intelligence</h1>
            <p className="header__tagline">Análisis de Datos — v4.1.0</p>
          </div>
        </div>
        {workbook && (
          <div className="header__actions">
            <button className="btn btn--ghost" onClick={reset} title="Cargar otro archivo (Esc)">
              <RefreshCw size={14} /> Nuevo
            </button>
          </div>
        )}
      </header>

      {/* File info bar */}
      {workbook && fileInfo && (
        <div className="file-bar">
          <div className="file-bar__info" style={{ flex: 1 }}>
            <div className="file-bar__icon"><FileSpreadsheet size={22} /></div>
            <div>
              <strong className="file-bar__name">{fileInfo.name}</strong>
              <div className="file-bar__meta">
                <span className="badge badge--purple">{fileInfo.ext}</span>
                <span className="badge badge--dim">{fileInfo.size}</span>
                <span className="badge badge--dim">{workbook.sheetNames.length} hoja(s)</span>
                <span className="file-bar__date"><Clock size={11} /> {fileInfo.lastModified.toLocaleDateString('es')}</span>
              </div>
            </div>
          </div>

          <div className="header__sheet-sel">
            <span style={{ fontSize: '0.75rem', fontWeight: 600 }}>Seleccionar Hoja:</span>
            <select className="sel sel--compact" value={activeSheet || ''} onChange={(e) => setActiveSheet(e.target.value)} style={{ minWidth: '180px' }}>
              {workbook.sheetNames.map((s) => (
                <option key={s} value={s}>{s} ({workbook.data[s].length} filas)</option>
              ))}
            </select>
          </div>
        </div>
      )}

      {/* Upload */}
      {!workbook && !loading && (
        <section className="upload-section">
          <div
            className={`upload${dragging ? ' upload--dragging' : ''}`}
            onDragOver={onDragOver}
            onDragLeave={onDragLeave}
            onDrop={onDrop}
            onClick={() => inputRef.current?.click()}
            role="button"
            tabIndex={0}
            aria-label="Zona de carga de archivos"
            onKeyDown={(e) => e.key === 'Enter' && inputRef.current?.click()}
          >
            <input ref={inputRef} type="file" onChange={onSelect} accept=".xlsx,.xls,.csv" hidden />
            <div className="upload__glow" />
            <div className="upload__icon-wrap">
              <UploadCloud className="upload__icon" />
            </div>
            <h2 className="upload__title">Arrastra tu archivo aquí</h2>
            <p className="upload__hint">
              Soporte para <strong>.xlsx</strong>, <strong>.xls</strong> y <strong>.csv</strong> — hasta 100 MB
            </p>
            <button className="btn btn--primary" onClick={(e) => { e.stopPropagation(); inputRef.current?.click(); }}>
              <FileSpreadsheet size={16} /> Seleccionar archivo
            </button>
            <p className="upload__shortcut"><kbd>Ctrl</kbd> + <kbd>O</kbd></p>
          </div>

          <div className="features">
            <FeatureCard icon={Shield} title="100% Privado" desc="Todo el procesamiento ocurre localmente en tu navegador. Ningún dato sale de tu equipo." />
            <FeatureCard icon={Zap} title="Análisis Instantáneo" desc="Estadísticas, gráficos y tablas generadas al instante con análisis inteligente." />
            <FeatureCard icon={BarChart3} title="Visualización Pro" desc="5 tipos de gráficos, filtros dinámicos, exportación PNG/CSV y correlaciones." />
          </div>
        </section>
      )}

      {/* Loading */}
      {loading && (
        <div className="loading-card">
          <Loader2 size={40} className="loading-card__spinner" />
          <h3 className="loading-card__title">Analizando tu archivo</h3>
          <p className="loading-card__desc">Extrayendo hojas, columnas y estadísticas…</p>
          <ProgressBar progress={progress} stage={getStageText(progress)} />
        </div>
      )}

      {/* Dashboard */}
      {workbook && !loading && activeSheet && (
        <AnalysisDashboard sheetData={workbook.data[activeSheet]} workbook={workbook} />
      )}

      {/* Theme toggle */}
      <button
        className="theme-toggle"
        onClick={() => setDarkMode((d) => !d)}
        title={darkMode ? 'Cambiar a modo claro' : 'Cambiar a modo oscuro'}
        aria-label="Toggle theme"
      >
        {darkMode ? <Sun size={18} /> : <Moon size={18} />}
      </button>

      {/* Footer */}
      <footer className="footer">
        <p>Excel Intelligence Platform · Procesamiento 100% local · {new Date().getFullYear()}</p>
      </footer>
    </div>
  );
}
