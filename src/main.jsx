import React, { useMemo, useState } from 'react';
import { createRoot } from 'react-dom/client';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';
import { LineChart, Line, XAxis, YAxis, Tooltip, ResponsiveContainer } from 'recharts';
import { Upload, BarChart3, FileText, Sparkles, Plus, Trash2 } from 'lucide-react';
import './styles.css';

const STORAGE_KEY = 'bloomstats-v1';

const initialData = {
  accounts: [
    { id: crypto.randomUUID(), platform: 'instagram', name: '@scalewithbloom', color: '#f472b6' },
    { id: crypto.randomUUID(), platform: 'youtube', name: 'Bloom Partner', color: '#fb7185' },
    { id: crypto.randomUUID(), platform: 'facebook', name: 'Bloom Digitals', color: '#60a5fa' }
  ],
  metrics: []
};

function loadData() {
  const saved = localStorage.getItem(STORAGE_KEY);
  if (!saved) return initialData;
  try { return JSON.parse(saved); } catch { return initialData; }
}

function platformIcon(platform) {
  const labels = { instagram: 'IG', youtube: 'YT', facebook: 'FB', tiktok: 'TT', linkedin: 'IN' };
  return <span className="platformBadge">{labels[platform] || 'RS'}</span>;
}

function parseNumber(value) {
  if (value == null) return 0;
  const clean = String(value).replace(/,/g, '').replace(/[^0-9.-]/g, '');
  return Number(clean) || 0;
}

function normalizeMetricRow(row, accountId, source = 'archivo') {
  return {
    id: crypto.randomUUID(),
    accountId,
    date: row.date || row.fecha || row.Date || row.Fecha || new Date().toISOString().slice(0, 10),
    followers: parseNumber(row.followers || row.seguidores || row.Followers || row.Seguidores),
    reach: parseNumber(row.reach || row.alcance || row.Reach || row.Alcance),
    impressions: parseNumber(row.impressions || row.impresiones || row.Impressions || row.Impresiones),
    views: parseNumber(row.views || row.vistas || row.Views || row.Vistas),
    likes: parseNumber(row.likes || row['me gusta'] || row.Likes),
    comments: parseNumber(row.comments || row.comentarios || row.Comments || row.Comentarios),
    shares: parseNumber(row.shares || row.compartidos || row.Shares || row.Compartidos),
    clicks: parseNumber(row.clicks || row.clics || row.Clicks || row.Clics),
    engagement: parseNumber(row.engagement || row.interacciones || row.Engagement || row.Interacciones),
    source
  };
}

function extractFromText(text, accountId) {
  const lower = text.toLowerCase();
  const fields = ['followers', 'reach', 'impressions', 'views', 'likes', 'comments', 'shares', 'clicks', 'engagement'];
  const metric = { id: crypto.randomUUID(), accountId, date: new Date().toISOString().slice(0, 10), source: 'texto', notes: text.slice(0, 500) };
  fields.forEach((field) => {
    const spanish = {
      followers: 'seguidores', reach: 'alcance', impressions: 'impresiones', views: 'vistas', likes: 'me gusta', comments: 'comentarios', shares: 'compartidos', clicks: 'clics', engagement: 'interacciones'
    }[field];
    const rx = new RegExp(`(${field}|${spanish})\\D{0,25}([0-9][0-9.,]*)`, 'i');
    const match = lower.match(rx);
    metric[field] = match ? parseNumber(match[2]) : 0;
  });
  return metric;
}

function App() {
  const [data, setData] = useState(loadData);
  const [selectedAccount, setSelectedAccount] = useState(data.accounts[0]?.id || '');
  const [rawText, setRawText] = useState('');
  const [newAccount, setNewAccount] = useState({ name: '', platform: 'instagram' });

  function save(next) {
    setData(next);
    localStorage.setItem(STORAGE_KEY, JSON.stringify(next));
  }

  const enriched = useMemo(() => data.metrics.map((m) => ({ ...m, account: data.accounts.find((a) => a.id === m.accountId)?.name || 'Cuenta' })), [data]);
  const latest = enriched.at(-1);
  const totals = enriched.reduce((acc, m) => {
    acc.followers = Math.max(acc.followers, m.followers || 0);
    acc.reach += m.reach || 0;
    acc.views += m.views || 0;
    acc.engagement += m.engagement || 0;
    return acc;
  }, { followers: 0, reach: 0, views: 0, engagement: 0 });

  function addAccount() {
    if (!newAccount.name.trim()) return;
    save({ ...data, accounts: [...data.accounts, { id: crypto.randomUUID(), ...newAccount, color: '#a78bfa' }] });
    setNewAccount({ name: '', platform: 'instagram' });
  }

  function importText() {
    if (!rawText.trim() || !selectedAccount) return;
    const metric = extractFromText(rawText, selectedAccount);
    save({ ...data, metrics: [...data.metrics, metric] });
    setRawText('');
  }

  function importFile(file) {
    if (!file || !selectedAccount) return;
    const name = file.name.toLowerCase();
    if (name.endsWith('.csv')) {
      Papa.parse(file, {
        header: true,
        skipEmptyLines: true,
        complete: (results) => {
          const rows = results.data.map((row) => normalizeMetricRow(row, selectedAccount, 'csv'));
          save({ ...data, metrics: [...data.metrics, ...rows] });
        }
      });
      return;
    }

    const reader = new FileReader();
    reader.onload = (event) => {
      const workbook = XLSX.read(event.target.result, { type: 'array' });
      const firstSheet = workbook.SheetNames[0];
      const sheet = workbook.Sheets[firstSheet];
      const jsonRows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
      const rows = jsonRows.map((row) => normalizeMetricRow(row, selectedAccount, name.endsWith('.xls') ? 'xls' : 'xlsx'));
      save({ ...data, metrics: [...data.metrics, ...rows] });
    };
    reader.readAsArrayBuffer(file);
  }

  function deleteMetric(id) {
    save({ ...data, metrics: data.metrics.filter((m) => m.id !== id) });
  }

  return (
    <main className="app">
      <aside className="sidebar">
        <div className="brand"><Sparkles /> <span>BloomStats</span></div>
        <p>Analiza estadisticas de redes sin sincronizar APIs. Sube texto, CSV o Excel y deja que el dashboard haga la talacha.</p>
        <nav><a>Dashboard</a><a>Importar</a><a>Cuentas</a><a>Analisis</a></nav>
      </aside>
      <section className="content">
        <header className="hero">
          <div><p className="eyebrow">Marketing dashboard</p><h1>Metricas rapidas, cero datos inventados.</h1><p>Primera version: guarda datos en tu navegador. Siguiente paso: Supabase para nube.</p></div>
          <button onClick={() => save({ ...data, metrics: [] })}><Trash2 size={16}/> limpiar metricas</button>
        </header>
        <section className="cards">
          <Card title="Seguidores" value={totals.followers} />
          <Card title="Alcance" value={totals.reach} />
          <Card title="Vistas" value={totals.views} />
          <Card title="Interacciones" value={totals.engagement} />
        </section>
        <section className="grid">
          <div className="panel wide">
            <h2><BarChart3/> Evolucion</h2>
            {enriched.length ? <ResponsiveContainer width="100%" height={300}><LineChart data={enriched}><XAxis dataKey="date"/><YAxis/><Tooltip/><Line type="monotone" dataKey="reach" strokeWidth={3}/><Line type="monotone" dataKey="views" strokeWidth={3}/></LineChart></ResponsiveContainer> : <Empty text="Aun no hay metricas. Importa texto, CSV o Excel." />}
          </div>
          <div className="panel">
            <h2>Resumen IA</h2>
            <p>{latest ? `Ultima carga: ${latest.account}. Alcance ${latest.reach || 0}, vistas ${latest.views || 0} e interacciones ${latest.engagement || 0}.` : 'Cuando importes datos, aqui aparecera una lectura rapida del rendimiento.'}</p>
          </div>
        </section>
        <section className="grid">
          <div className="panel">
            <h2><Plus/> Cuentas</h2>
            <div className="row"><select value={newAccount.platform} onChange={(e)=>setNewAccount({...newAccount, platform:e.target.value})}><option>instagram</option><option>youtube</option><option>facebook</option><option>tiktok</option><option>linkedin</option></select><input placeholder="Nombre" value={newAccount.name} onChange={(e)=>setNewAccount({...newAccount, name:e.target.value})}/><button onClick={addAccount}>Agregar</button></div>
            <div className="accountList">{data.accounts.map(a => <div className="account" key={a.id}>{platformIcon(a.platform)} <span>{a.name}</span></div>)}</div>
          </div>
          <div className="panel">
            <h2><Upload/> Importar datos</h2>
            <select value={selectedAccount} onChange={(e)=>setSelectedAccount(e.target.value)}>{data.accounts.map(a => <option key={a.id} value={a.id}>{a.name}</option>)}</select>
            <textarea placeholder="Pega texto copiado de Insights. Ejemplo: Seguidores 320 Alcance 1500 Vistas 4200 Interacciones 180" value={rawText} onChange={(e)=>setRawText(e.target.value)} />
            <div className="row"><button onClick={importText}>Importar texto</button><label className="file"><FileText size={16}/> CSV / XLS / XLSX<input type="file" accept=".csv,.xls,.xlsx" onChange={(e)=>e.target.files?.[0] && importFile(e.target.files[0])}/></label></div>
          </div>
        </section>
        <section className="panel">
          <h2>Historial</h2>
          {enriched.length ? <div className="table">{enriched.slice().reverse().map(m => <div className="tr" key={m.id}><span>{m.date}</span><span>{m.account}</span><span>Alcance {m.reach || 0}</span><span>Vistas {m.views || 0}</span><button onClick={()=>deleteMetric(m.id)}>x</button></div>)}</div> : <Empty text="Todavia no hay registros." />}
        </section>
      </section>
    </main>
  );
}

function Card({ title, value }) { return <div className="card"><span>{title}</span><strong>{Number(value || 0).toLocaleString('es-MX')}</strong></div>; }
function Empty({ text }) { return <div className="empty">{text}</div>; }

createRoot(document.getElementById('root')).render(<App />);
