const express = require('express');
const XLSX    = require('xlsx');
const cors    = require('cors');
const path    = require('path');
const fs      = require('fs');

const app = express();
app.use(cors());
app.use(express.json());

const EXCEL_PATH = path.join(__dirname, 'SAT Progress.xlsx');
const CACHE_TTL  = 5 * 60 * 1000;
let cacheTime = 0, cachedData = null;

function toDate(v) {
  if (typeof v !== 'number' || v < 40000) return null;
  const d = new Date((v - 25569) * 86400000);
  return isNaN(d) ? null : d.toISOString().slice(0,10);
}

function getWeekKey(dateStr, projStart) {
  if (!dateStr || !projStart) return -1;
  const d = new Date(dateStr + 'T00:00:00');
  const s = new Date(projStart + 'T00:00:00');
  return Math.floor((d - s) / (7 * 86400000));
}

function parseExcel() {
  if (!fs.existsSync(EXCEL_PATH)) throw new Error('Excel not found');
  const wb = XLSX.readFile(EXCEL_PATH);
  const ws = wb.Sheets['SVB-Coring'];
  if (!ws) throw new Error('SVB-Coring sheet not found');
  const rows = XLSX.utils.sheet_to_json(ws, { header:1, defval:null });

  // หา proj_start จาก Cabling Date (S) ที่เร็วสุด
  let projStart = null, projEnd = null;
  const gates = [];
  const locMap = {};

  rows.slice(1).forEach(r => {
    if (!r[0] || typeof r[0] !== 'number') return;
    const gate    = String(r[1] || '').trim();
    const loc     = String(r[2] || '').trim();
    const cab_s   = toDate(r[3]);
    const cab_a   = toDate(r[4]);
    const qty     = Number(r[5]  || 0);
    const cab_d   = Number(r[6]  || 0);
    const cor_s   = toDate(r[8]);
    const cor_a   = toDate(r[9]);
    const cor_d   = Number(r[11] || 0);

    if (cab_s && (!projStart || cab_s < projStart)) projStart = cab_s;
    if (cor_s && (!projEnd   || cor_s > projEnd))   projEnd   = cor_s;

    if (loc) {
      if (!locMap[loc]) locMap[loc] = {cab_p:0,cab_d:0,cor_p:0,cor_d:0};
      locMap[loc].cab_p += qty;
      locMap[loc].cab_d += cab_d;
      locMap[loc].cor_p += qty;
      locMap[loc].cor_d += cor_d;
    }
    gates.push({ gate, loc, qty, cab_d, cor_d, cab_s, cab_a, cor_s, cor_a });
  });

  // ตัวเลขรวม
  const total    = gates.reduce((a,g) => a + g.qty, 0);
  const cab_done = gates.reduce((a,g) => a + g.cab_d, 0);
  const cor_done = gates.reduce((a,g) => a + g.cor_d, 0);
  const cab_pct  = total > 0 ? Math.round(cab_done / total * 1000) / 10 : 0;
  const cor_pct  = total > 0 ? Math.round(cor_done / total * 1000) / 10 : 0;
  const combined = Math.round((cab_pct + cor_pct) / 2 * 10) / 10;

  // Weekly — สร้าง labels จาก projStart ถึง projEnd+4 สัปดาห์
  const PS = new Date(projStart + 'T00:00:00');
  const PE = new Date((projEnd || projStart) + 'T00:00:00');
  const nWk = Math.ceil((PE - PS) / (7 * 86400000)) + 5;

  const wk_cab_p = new Array(nWk).fill(0);
  const wk_cab_a = new Array(nWk).fill(null);
  const wk_cor_p = new Array(nWk).fill(0);
  const wk_cor_a = new Array(nWk).fill(null);

  gates.forEach(g => {
    if (g.cab_s) { const w=getWeekKey(g.cab_s,projStart); if(w>=0&&w<nWk) wk_cab_p[w]+=g.qty; }
    if (g.cab_a) { const w=getWeekKey(g.cab_a,projStart); if(w>=0&&w<nWk) { wk_cab_a[w]=(wk_cab_a[w]||0)+g.cab_d; } }
    if (g.cor_s) { const w=getWeekKey(g.cor_s,projStart); if(w>=0&&w<nWk) wk_cor_p[w]+=g.qty; }
    if (g.cor_a) { const w=getWeekKey(g.cor_a,projStart); if(w>=0&&w<nWk) { wk_cor_a[w]=(wk_cor_a[w]||0)+g.cor_d; } }
  });

  // สร้าง week labels
  const wkLabels = Array.from({length:nWk},(_,i)=>`W.${i+1}`);
  const wkDates  = Array.from({length:nWk},(_,i)=>{
    const s=new Date(PS); s.setDate(s.getDate()+i*7);
    const e=new Date(s);  e.setDate(e.getDate()+6);
    return `${s.getDate()}/${s.getMonth()+1}-${e.getDate()}/${e.getMonth()+1}`;
  });

  // Cumulative %
  const cumCabPlan=[], cumCabAct=[], cumCorPlan=[], cumCorAct=[], cumCombPlan=[], cumCombAct=[];
  let sCabP=0,sCabA=0,sCorP=0,sCorA=0;
  wk_cab_p.forEach((_,i)=>{
    sCabP+=wk_cab_p[i]; sCorP+=wk_cor_p[i];
    const cabPct=total>0?Math.round(sCabP/total*10000)/100:0;
    const corPct=total>0?Math.round(sCorP/total*10000)/100:0;
    cumCabPlan.push(cabPct);
    cumCorPlan.push(corPct);
    cumCombPlan.push(Math.round((cabPct+corPct)/2*100)/100);
    if(wk_cab_a[i]!==null) sCabA+=wk_cab_a[i];
    if(wk_cor_a[i]!==null) sCorA+=wk_cor_a[i];
    const cabActPct=total>0?Math.round(sCabA/total*10000)/100:null;
    const corActPct=total>0?Math.round(sCorA/total*10000)/100:null;
    cumCabAct.push(wk_cab_a[i]!==null?cabActPct:null);
    cumCorAct.push(wk_cor_a[i]!==null?corActPct:null);
    cumCombAct.push((wk_cab_a[i]!==null||wk_cor_a[i]!==null)?Math.round((sCabA+sCorA)/total/2*10000)/100:null);
  });

  // Burndown
  const bdPlan = cumCombPlan.map(p => Math.round((100-p)*total/100*2)); // total*2 = combined
  const bdAct  = cumCombAct.map(p => p!==null ? Math.round((100-p)*total/100*2) : null);

  // Locations
  const locations = Object.entries(locMap).map(([n,v])=>({
    n, cab_p:v.cab_p, cab_d:v.cab_d, cor_p:v.cor_p, cor_d:v.cor_d,
    cab_pct: v.cab_p>0?Math.round(v.cab_d/v.cab_p*100):0,
    cor_pct: v.cor_p>0?Math.round(v.cor_d/v.cor_p*100):0,
  }));

  return {
    overall: {
      total, cab_done, cor_done, cab_pct, cor_pct, combined_pct:combined,
      proj_start:projStart, proj_end:projEnd,
    },
    weekly: {
      labels:wkLabels, dates:wkDates, n_wk:nWk,
      cab_plan:wk_cab_p, cab_act:wk_cab_a,
      cor_plan:wk_cor_p, cor_act:wk_cor_a,
      cum_cab_plan:cumCabPlan, cum_cab_act:cumCabAct,
      cum_cor_plan:cumCorPlan, cum_cor_act:cumCorAct,
      cum_comb_plan:cumCombPlan, cum_comb_act:cumCombAct,
    },
    burndown: { plan:bdPlan, act:bdAct },
    locations,
    gates: gates.map(g=>({
      gate:g.gate, loc:g.loc, qty:g.qty,
      cab:g.cab_d, cor:g.cor_d,
    })),
    last_updated: new Date().toISOString(),
  };
}

async function getData() {
  const now = Date.now();
  if (cachedData && (now - cacheTime) < CACHE_TTL) return cachedData;
  cachedData = parseExcel();
  cacheTime = now;
  return cachedData;
}

app.get('/api/summary', async (req, res) => {
  try { res.json(await getData()); }
  catch(e) { res.status(500).json({ error: e.message }); }
});

app.get('/api/health', (req, res) => res.json({ status:'ok' }));

app.post('/api/cache/refresh', (req, res) => {
  cacheTime = 0; cachedData = null;
  try { res.json({ success:true, data: parseExcel() }); }
  catch(e) { res.json({ success:false, error:e.message }); }
});

app.use(express.static(path.join(__dirname, '../frontend')));
app.get('*', (req, res) => res.sendFile(path.join(__dirname,'../frontend/index.html')));

const PORT = process.env.PORT || 3002;
app.listen(PORT, () => console.log(`Coring Dashboard on port ${PORT}`));
