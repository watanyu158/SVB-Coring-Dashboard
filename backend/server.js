const express = require('express');
const XLSX    = require('xlsx');
const cors    = require('cors');
const https   = require('https');
const http    = require('http');
const path    = require('path');
const fs      = require('fs');

const app = express();
app.use(cors());
app.use(express.json());

const SHAREPOINT_URL = process.env.SHAREPOINT_URL || '';
const CACHE_PATH     = path.join(__dirname, 'coring_cache.xlsx');
const LOCAL_PATH     = path.join(__dirname, 'SAT Progress.xlsx');
const CACHE_TTL_MS   = 5 * 60 * 1000;

let cacheTime = 0;
let cachedData = null;

// ── Static fallback data (from Excel SVB-Coring sheet) ──────────────────────
const STATIC_DATA = {
  overall: {
    total_points: 39, cabling_done: 39, coring_done: 7,
    cabling_pct: 100.0, coring_pct: 17.9, combined_pct: 59.0,
    project_start: '2025-11-08', project_end: '2026-01-21',
    status: 'Cabling Complete / Coring In Progress'
  },
  locations: [
    { name:'Concourse C', plan:10, done:10, pct:100 },
    { name:'Concourse E', plan:10, done:10, pct:100 },
    { name:'Concourse D', plan:8,  done:8,  pct:100 },
    { name:'Concourse F', plan:6,  done:6,  pct:100 },
    { name:'Concourse G', plan:5,  done:5,  pct:100 },
  ],
  weekly: {
    labels: ['W.1','W.2','W.3','W.4','W.5','W.6','W.7','W.8','W.9','W.10',
             'W.11','W.12','W.13','W.14','W.15','W.16','W.17','W.18','W.19'],
    plan:   [2,2,7,3,4,6,2,0,2,6,5,0,0,0,0,0,0,0,0],
    actual: [2,2,7,3,4,6,2,0,2,6,5,null,null,null,null,null,null,null,null],
    cum_plan_pct: [2.6,5.1,14.1,17.9,23.1,30.8,33.3,33.3,35.9,43.6,51.3,51.3,51.3,52.6,55.1,56.4,57.7,59.0,61.5],
    cum_act_pct:  [2.6,5.1,14.1,17.9,23.1,30.8,33.3,33.3,35.9,43.6,51.3,51.3,51.3,52.6,55.1,56.4,57.7,59.0,59.0]
  },
  last_updated: new Date().toISOString()
};

function downloadFile(url, dest) {
  return new Promise((resolve, reject) => {
    const proto = url.startsWith('https') ? https : http;
    proto.get(url, { headers:{'User-Agent':'Mozilla/5.0'} }, res => {
      if ([301,302,303,307,308].includes(res.statusCode))
        return downloadFile(res.headers.location, dest).then(resolve).catch(reject);
      if (res.statusCode !== 200) return reject(new Error(`HTTP ${res.statusCode}`));
      const f = fs.createWriteStream(dest);
      res.pipe(f);
      f.on('finish', () => f.close(resolve));
      f.on('error', reject);
    }).on('error', reject);
  });
}

async function parseExcel(wb) {
  try {
    const ws = wb.Sheets['SVB-Coring'];
    if (!ws) return null;
    const rows = XLSX.utils.sheet_to_json(ws, { header:1 });
    
    let total=0, done=0;
    const locs = {};
    rows.slice(1).forEach(r => {
      if (!r[0] || typeof r[0] !== 'number') return;
      const loc = r[2]; const qty = r[5]||0; const ok = r[6]||0;
      total+=qty; done+=ok;
      if (loc) { if(!locs[loc]) locs[loc]={plan:0,done:0}; locs[loc].plan+=qty; locs[loc].done+=ok; }
    });

    return {
      overall: {
        total_points: total, cabling_done: done,
        cabling_pct: total>0?Math.round(done/total*1000)/10:0,
        status: done>=total ? 'Cabling Complete' : 'In Progress'
      },
      locations: Object.entries(locs).map(([name,d])=>({
        name, plan:d.plan, done:d.done,
        pct: d.plan>0?Math.round(d.done/d.plan*100):0
      })),
      last_updated: new Date().toISOString()
    };
  } catch(e) { console.error('Parse err:', e.message); return null; }
}

async function getData() {
  const now = Date.now();
  if (cachedData && (now - cacheTime) < CACHE_TTL_MS) return cachedData;

  // Try SharePoint first
  if (SHAREPOINT_URL) {
    try {
      await downloadFile(SHAREPOINT_URL, CACHE_PATH);
      const wb = XLSX.readFile(CACHE_PATH);
      const parsed = await parseExcel(wb);
      if (parsed) { cachedData = {...STATIC_DATA, ...parsed}; cacheTime = now; return cachedData; }
    } catch(e) { console.error('SharePoint err:', e.message); }
  }

  // Try local Excel
  const localPath = fs.existsSync(CACHE_PATH) ? CACHE_PATH : LOCAL_PATH;
  if (fs.existsSync(localPath)) {
    try {
      const wb = XLSX.readFile(localPath);
      const parsed = await parseExcel(wb);
      if (parsed) { cachedData = {...STATIC_DATA, ...parsed}; cacheTime = now; return cachedData; }
    } catch(e) { console.error('Local Excel err:', e.message); }
  }

  // Fallback: static data
  console.log('Using static data');
  cachedData = STATIC_DATA;
  cacheTime = now;
  return cachedData;
}

// ── Routes ────────────────────────────────────────────────────────────────────
app.get('/api/summary', async (req, res) => {
  try { res.json(await getData()); }
  catch(e) { res.json(STATIC_DATA); }
});

app.get('/api/health', (req, res) => res.json({ status:'ok', time: new Date().toISOString() }));

app.post('/api/cache/refresh', async (req, res) => {
  cacheTime = 0; cachedData = null;
  try { await getData(); res.json({ success:true }); }
  catch(e) { res.json({ success:false, error:e.message }); }
});

// ── Serve frontend ─────────────────────────────────────────────────────────
app.use(express.static(path.join(__dirname, '../frontend')));
app.get('*', (req, res) => res.sendFile(path.join(__dirname,'../frontend/index.html')));

const PORT = process.env.PORT || 3002;
app.listen(PORT, () => console.log(`Coring Dashboard running on port ${PORT}`));
