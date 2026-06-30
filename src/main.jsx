import React, { useMemo, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';
import { LineChart, Line, XAxis, YAxis, Tooltip, ResponsiveContainer } from 'recharts';
import { BarChart3, FileSpreadsheet, UploadCloud, Trash2, Plus, CheckCircle2 } from 'lucide-react';
import './styles.css';

const STORAGE_KEY = 'bloomstats-v2';

const DEFAULT_ACCOUNTS = [
  { id: 'linkedin-bloom', platform: 'linkedin', name: 'Bloom Digitals LinkedIn' },
  { id: 'instagram-bloom', platform: 'instagram', name: '@scalewithbloom' },
  { id: 'youtube-bloom', platform: 'youtube', name: 'Bloom Partner' },
  { id: 'facebook-bloom', platform: 'facebook', name: 'Bloom Digitals' }
];

function today() {
  return new Date().toISOString().slice(0, 10);
}

function loadData() {
  try {
    const saved = JSON.parse(localStorage.getItem(STORAGE_KEY));
    if (saved?.accounts && saved?.metrics) return saved;
  } catch {}
  return { accounts: DEFAULT_ACCOUNTS, metrics: [] };
}

function parseNumber(value) {
  if (value === undefined || value === null || value === '') return 0;
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
  const clean = String(value).replace(/,/g, '').replace(/%/g, '').replace(/[^0-9.-]/g, '');
  return Number(clean) || 0;
}

function normalizeKey(key) {
  return String(key || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

function getValue(row, aliases) {
  const normalized = Object.fromEntries(Object.entries(row).map(([key, value]) => [normalizeKey(key), value]));
  for (const alias of aliases) {
    const value = normalized[normalizeKey(alias)];
    if (value !== undefined && value !== null && value !== '') return value;
  }
  return '';
}

function normalizeDate(value) {
  if (!value) return today();
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  if (typeof value === 'number') {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) return `${parsed.y}-${String(parsed.m).padStart(2, '0')}-${String(parsed.d).padStart(2, '0')}`;
  }
  const raw = String(value).trim();
  const date = new Date(raw);
  if (!Number.isNaN(date.getTime())) return date.toISOString().slice(0, 10);
  return raw;
}

function rowLooksUseful(row) {
  const values = Object.values(row).join(' ').toLowerCase();
  return values.includes('impression') || values.includes('click') || values.includes('post title') || values.includes('created date') || values.includes('engagement') || values.includes('alcance') || values.includes('impres');
}

function sheetToObjectsSmart(sheet) {
  const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  let headerIndex = matrix.findIndex((row) => rowLooksUseful(Object.fromEntries(row.map((value, index) => [index, value]))));
  if (headerIndex < 0) headerIndex = 0;
  const headers = matrix[headerIndex].map((cell, index) => String(cell || `Column ${index + 1}`).trim());
  return matrix.slice(headerIndex + 1).map((row) => {
    const object = {};
    headers.forEach((header, index) => { object[header] = row[index] ?? ''; });
    return object;
  }).filter((row) => Object.values(row).some((value) => value !== ''));
}

function normalizeMetricRow(row, accountId, source) {
  const likes = parseNumber(getValue(row, ['likes', 'reactions', 'reactions total', 'reactions (total)', 'me gusta']));
  const comments = parseNumber(getValue(row, ['comments', 'comments total', 'comments (total)', 'comentarios']));
  const reposts = parseNumber(getValue(row, ['reposts', 'reposts total', 'reposts (total)', 'shares', 'compartidos']));
  const clicks = parseNumber(getValue(row, ['clicks', 'clicks total', 'clicks (total)', 'clics']));
  const engagement = parseNumber(getValue(row, ['engagement', 'interacciones'])) || likes + comments + reposts + clicks;

  return {
    id: crypto.randomUUID(),
    accountId,
    source,
    date: normalizeDate(getValue(row, ['date', 'fecha', 'created date', 'posted date'])),
    title: String(getValue(row, ['post title', 'title', 'publicacion', 'publicación'])).slice(0, 180),
    link: String(getValue(row, ['post link', 'link', 'url'])),
    impressions: parseNumber(getValue(row, ['impressions', 'impressions total', 'impressions (total)', 'impresiones', 'impressions organic', 'impressions (organic)'])),
    reach: parseNumber(getValue(row, ['reach', 'alcance', 'unique impressions organic', 'unique impressions (organic)'])),
    views: parseNumber(getValue(row, ['views', 'vistas', 'offsite views'])),
    clicks,
    likes,
    comments,
    reposts,
    engagement,
    engagementRate: parseNumber(getValue(row, ['engagement rate', 'engagement rate total', 'engagement rate (total)', 'click through rate ctr', 'click through rate (ctr)']))
  };
}

function App() {
  const fileInputRef = useRef(null);
  const [data, setData] = useState(loadData);
  const [selectedAccount, setSelectedAccount] = useState(data.accounts[0]?.id || '');
  const [newAccount, setNewAccount] = useState('');
  const [newPlatform, setNewPlatform] = useState('linkedin');
  const [status, setStatus] = useState('Sube un archivo CSV, XLS o XLSX para empezar.');
  const [isImporting, setIsImporting] = useState(false);

  function save(next) {
    setData(next);
    localStorage.setItem(STORAGE_KEY, JSON.stringify(next));
  }

  const account = data.accounts.find((item) => item.id === selectedAccount);
  const metrics = useMemo(() => data.metrics.map((metric) => ({ ...metric, accountName: data.accounts.find((a) => a.id === metric.accountId)?.name || 'Cuenta' })), [data]);
  const totals = useMemo(() => metrics.reduce((acc, metric) => {
    acc.impressions += metric.impressions || 0;
    acc.reach += metric.reach || 0;
    acc.clicks += metric.clicks || 0;
    acc.engagement += metric.engagement || 0;
    acc.posts += metric.title ? 1 : 0;
    return acc;
  }, { impressions: 0, reach: 0, clicks: 0, engagement: 0, posts: 0 }), [metrics]);

  const topPosts = useMemo(() => metrics.filter((m) => m.title).sort((a, b) => (b.impressions + b.clicks * 8) - (a.impressions + a.clicks * 8)).slice(0, 8), [metrics]);

  function addAccount() {
    if (!newAccount.trim()) return;
    const next = { id: crypto.randomUUID(), platform: newPlatform, name: newAccount.trim() };
    save({ ...data, accounts: [...data.accounts, next] });
    setSelectedAccount(next.id);
    setNewAccount('');
  }

  function importRows(rows, source) {
    const normalized = rows.map((row) => normalizeMetricRow(row, selectedAccount, source)).filter((row) => row.impressions || row.reach || row.views || row.clicks || row.engagement || row.title);
    if (!normalized.length) {
      setStatus('No pude encontrar columnas de metricas en ese archivo. Revisa que tenga encabezados como Impressions, Clicks, Date o Post title.');
      return;
    }
    save({ ...data, metrics: [...data.metrics, ...normalized] });
    const postCount = normalized.filter((row) => row.title).length;
    setStatus(`Listo: importe ${normalized.length} registros${postCount ? ` y ${postCount} publicaciones` : ''} para ${account?.name || 'la cuenta seleccionada'}.`);
  }

  function importFile(file) {
    if (!file) return;
    if (!selectedAccount) {
      setStatus('Primero selecciona una cuenta.');
      return;
    }
    setIsImporting(true);
    setStatus(`Leyendo ${file.name}...`);
    const fileName = file.name.toLowerCase();

    if (fileName.endsWith('.csv')) {
      Papa.parse(file, {
        header: true,
        skipEmptyLines: true,
        complete: (results) => {
          importRows(results.data, 'csv');
          setIsImporting(false);
        },
        error: () => {
          setStatus('No pude leer el CSV.');
          setIsImporting(false);
        }
      });
      return;
    }

    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const workbook = XLSX.read(event.target.result, { type: 'array', cellDates: true });
        const rows = workbook.SheetNames.flatMap((sheetName) => sheetToObjectsSmart(workbook.Sheets[sheetName]));
        importRows(rows, fileName.endsWith('.xls') ? 'xls' : 'xlsx');
      } catch (error) {
        setStatus('No pude leer el Excel. Prueba exportarlo otra vez o convertirlo a CSV.');
      } finally {
        setIsImporting(false);
      }
    };
    reader.onerror = () => {
      setStatus('No pude abrir el archivo seleccionado.');
      setIsImporting(false);
    };
    reader.readAsArrayBuffer(file);
  }

  function clearMetrics() {
    save({ ...data, metrics: [] });
    setStatus('Metricas borradas. Puedes importar de nuevo.');
  }

  return (
    <main className="shell">
      <header className="topbar">
        <div>
          <p className="eyebrow">BloomStats</p>
          <h1>Dashboard de estadisticas sociales</h1>
          <p className="subtitle">Importa reportes reales de LinkedIn, Meta, YouTube o CSV sin conectar APIs.</p>
        </div>
        <button className="ghost" onClick={clearMetrics}><Trash2 size={16} /> limpiar datos</button>
      </header>

      <section className="statsGrid">
        <Card title="Impresiones" value={totals.impressions} />
        <Card title="Alcance" value={totals.reach} />
        <Card title="Clicks" value={totals.clicks} />
        <Card title="Interacciones" value={totals.engagement} />
      </section>

      <section className="mainGrid">
        <div className="panel importPanel">
          <div className="sectionTitle"><UploadCloud /><div><h2>Importar archivo</h2><p>Selecciona cuenta y sube CSV, XLS o XLSX.</p></div></div>
          <label>Cuenta</label>
          <select value={selectedAccount} onChange={(e) => setSelectedAccount(e.target.value)}>
            {data.accounts.map((item) => <option key={item.id} value={item.id}>{item.name} · {item.platform}</option>)}
          </select>
          <button className="uploadButton" disabled={isImporting} onClick={() => fileInputRef.current?.click()}>
            <FileSpreadsheet size={20} /> {isImporting ? 'Importando...' : 'Elegir CSV / XLS / XLSX'}
          </button>
          <input ref={fileInputRef} className="hiddenInput" type="file" accept=".csv,.xls,.xlsx" onChange={(event) => { const file = event.target.files?.[0]; importFile(file); event.target.value = ''; }} />
          <p className="notice"><CheckCircle2 size={16} /> {status}</p>
          <div className="accountCreator">
            <h3>Agregar cuenta</h3>
            <div className="row">
              <select value={newPlatform} onChange={(e) => setNewPlatform(e.target.value)}>
                <option value="linkedin">LinkedIn</option><option value="instagram">Instagram</option><option value="facebook">Facebook</option><option value="youtube">YouTube</option><option value="tiktok">TikTok</option>
              </select>
              <input value={newAccount} onChange={(e) => setNewAccount(e.target.value)} placeholder="Nombre de cuenta" />
              <button onClick={addAccount}><Plus size={16} /> agregar</button>
            </div>
          </div>
        </div>

        <div className="panel chartPanel">
          <div className="sectionTitle"><BarChart3 /><div><h2>Evolucion</h2><p>Impresiones y clicks por fecha.</p></div></div>
          {metrics.length ? <ResponsiveContainer width="100%" height={310}><LineChart data={metrics}><XAxis dataKey="date" /><YAxis /><Tooltip /><Line type="monotone" dataKey="impressions" strokeWidth={3} /><Line type="monotone" dataKey="clicks" strokeWidth={3} /></LineChart></ResponsiveContainer> : <Empty text="Aun no hay datos importados." />}
        </div>
      </section>

      <section className="panel">
        <div className="sectionTitle"><FileSpreadsheet /><div><h2>Publicaciones detectadas</h2><p>Ordenadas por rendimiento.</p></div></div>
        {topPosts.length ? <div className="table">{topPosts.map((post) => <div className="tableRow" key={post.id}><span className="postTitle">{post.title || post.accountName}</span><span>{post.date}</span><span>{post.impressions.toLocaleString('es-MX')} imp.</span><span>{post.clicks.toLocaleString('es-MX')} clicks</span><span>{post.engagement.toLocaleString('es-MX')} int.</span></div>)}</div> : <Empty text="Cuando subas un reporte de contenido, aqui apareceran las publicaciones." />}
      </section>
    </main>
  );
}

function Card({ title, value }) {
  return <div className="statCard"><span>{title}</span><strong>{Number(value || 0).toLocaleString('es-MX')}</strong></div>;
}

function Empty({ text }) {
  return <div className="emptyState">{text}</div>;
}

createRoot(document.getElementById('root')).render(<App />);
