import { useState, useCallback, useRef, useEffect, useMemo, Component } from "react";
import * as XLSX from "xlsx";

const LOGO_AIRBOX = "data:image/webp;base64,UklGRtoHAABXRUJQVlA4IM4HAADwJACdASqgAC8APlEgjUQjoiEUTP9EOAUEtgBm/OBdb/bP5B+tP5FfKpW37T9+uELMT2G/pftk953qJ/IH+O9wD9I/8R+Yf9A7hvmA/W39TPdf9B/2AfAP+vHWe+gR+2fpdftx8H/7NftZ8B37C3hp+b6P70d7Cco9kf+D/LTkb2vv81+UnCT8o/pH+l8HD9r9FdKDoAfyD+uee59D+db82/yH/l/yXwF/y/+s+mB67/3M9iL9bFNNAeXon1YodIzbPHghy5qHE8Aki9GXy7t6h89ySixxMBTkVF6JC3B6t8rkTnFF6fIUSHanGnEY1wirAPWMUgGFiDYgCbtunqz9t6rlhe+xSyumsUBhRDjD78rwkNOMUejVq5ypajusSYjDwEUpOaFUR5GcGaiKNwgAAP78KdH+MyKRcYN3gm7dADVWesqTNlLEDir+cBDRDNeFfmafO9nFo5aQ0O9esQMxxQjex+J7sOixidI7UQT3kbMtlTBqhpo/G//12bKJrgK5R312ZcQvEd7np5SAsfhgiEuIB8829zWOdV1YWuFNKixeY2c5aregcq5231JP4ujd+Awl2NF/xSfmtk6LIx9npB4sJn+q+sreCQD/383TttALAF/EuKX++pvWoOVHwtFdOO06PeoX7lg/IPuBDxSh5AAOXshXY0VNYS2FTYlvPfzi9sJtI86Yei9JplMcWmDJNGlA7WaXoDDj/q5n4iCZNM8rv3huVM+0heZb/8eSloMEDyRJJ5E5kvUZ6lwOtsYgK2MITFr4VUBjL8HfHfjG5VMt/HZNOJQPNw2fD7P13NQn7i0tTf8HYVqLOYk8OAg8+/5avQnShaRNawWcxT84wwwxCOVW6SiH44AtZXzfMGrDwabQhKJkc4IfxNfvJ9B+l+SX6DTaid/jElUaNsUP/GZE11aNkgyu3wXaI6+QFTG92SBy4H83LiBNqbhTZzdbfOns1R3MZIkSVpLHXZlMk7by/rAxVmfcaMQDI4U30dBDQFnXQ86xvECd+f7qYoeE3FEhseHJtZPQ6Jt8wrCwYaF8uH+zACj+gnPZJcQtQCW4DYbjd6MdYGmmRiVTCHQXjMfy9Oh99H/0Xd53rDnzHJ/WxOzzymUcWFYp/2q5JpbJU8kC0cI55RiaUi9eYZ8Z4057SF/H2GzIQWXlxoEpJy0WX133y9Xmj1TOIIfmCqoac2t0OGVoMpTe59zPEm7394KDgXF3TpeVHKw9xcvZPX0i4TubbHOEQ4cuUmD3c4CQem05x/9Yu14+M/X2OVHgnJ+BqXaSX5uYSZy5i6chHFCZ84KUuOzODz49us/T3VipjTracuhogYqnf60W9PI5gBDM8AgMlqkg3Sm0Xa/Ci9ZAVj+7LWjDLtNSMpR8jw3wyXLyOEHoT7Dw1bHkRlnEmaDyCjptQPUywqZs0BY9nIsrRcxqvyKzMF09IHWclL73uOmdfOiNIw59qD+l11NPDUmGn7t/XScfOu4WHnbTLHQxbp53//RloEYxVYwNcr/8Ln8gt66E1FA5kw9sC2UeMexPIZw7Ml14xoyA6ldqkoK8Bm6AsePn2cz7urwyxO+m/hQ/fzx3dzob8GydpB/7gVM9002F3OO7G19n82As73l+C4i5uLDXkjK8rN4gaeRH2GDu2uGgZ583z0rQLirjbecQNrr1rd2fOj5BI9fEUoxtVQmkl1BPy1cAUn/736U8wIA139fI4+XVrp/IPvDbgjZNsFVs7u2xGXcLhpffZuwmjn9LrAcGua7yMfgBcSQb+/ZB4McQ8C2p40wys2XuZE/g9rxAZXocWwpPON3ftOuEWyoqSicPRngRJ49Fn8N71BzP6dfVaVmb0yY/pxTmJDJfkwOhpXqZJ6Zdr3116A2wyC6SSfdPzj9nokBUYxgarH9DrjNwF6OuMN3enPuyI2UAfllO87FlNFIQcJkn1yvO4IJl1oIJeuOZxipmKDFBgCWqgvtqPb7mlhsLJvKDQaD///ni76Ndf1fAaQwOSX2ccc8sJKrut+CR5hqkWztO0koosGTbJNvugNnyfHfvsW2PDVlYlK3t8bZwHT7lIMwpk0AJ1Ll0v+Qv3QqMWnFC4Kwlp98wpZTURgPkSGF3cVSKtQijjJoMviLbeAybc4gqP9unRJaZWMoth8JhG+nUFXTbusq6Q4C8AOghCt8X6D5enUzp+PRT1FxYgjY/saiBReCncqkIuEDLbX//qqfeOdY0XtWF+0fyGX+9d/BIqerebQ//BhJ3n3+f7XLu5tTP9f6eK8AghP0L5FF/47Wn4D6iX1rVNEH+7lzFnBHRQIDPfRDpbPKNr0zXuvXcBSIHzF73pO7tQij+4KSlyb/JJ8WnwXb5KQPZ6ni2DYIXedAV1saUn+h72EkOjUzB7Q0z4Bhryge2919Gp5WqAQmTpW2YT++9R5EYRrRlhD5igEY4A660s1/v5uByLIOKaqfrLHw1C2FHx8SuZqhemgZzuwU2cD8cLzUU6x254l4drXxRajPi5QV7Teg2nfO3IPvquFaCP+NI0N5tHaCyGQZJJQ/5Vp/24sYgT1SUFSXy/U4gy42BBTtd4UWFZ5aA+179FmZRCOzG3gEcbHYvNk6YFCZDeOHpfZk3usCLAdH3J2If2kmx0lmEXyoDMuPU//BEZ/vB4U01OYS2AAA=";

// ── Error Boundary ────────────────────────────────────────────────────────────
class ErrorBoundary extends Component {
  constructor(props) { super(props); this.state = { error: null }; }
  static getDerivedStateFromError(error) { return { error }; }
  render() {
    if (this.state.error) {
      return (
        <div style={{padding:32,fontFamily:"monospace",color:"#F44336",background:"#fff"}}>
          <strong>Erro no dashboard:</strong><br/>
          {this.state.error.message}<br/>
          <button onClick={()=>this.setState({error:null})} style={{marginTop:12,padding:"6px 14px",cursor:"pointer"}}>
            Tentar novamente
          </button>
        </div>
      );
    }
    return this.props.children;
  }
}

// ── Paleta & cores ────────────────────────────────────────────────────────────
const PALETTE = [
  "#2E9BFF","#F5A623","#4CAF50","#E91E63","#9C27B0","#00BCD4","#FF5722","#795548",
  "#1DE9B6","#FF6D00","#8BC34A","#D500F9","#00BFA5","#FFD740","#304FFE","#607D8B",
  "#3F51B5","#00C853","#AA00FF","#0091EA","#64DD17","#FF6E40","#F9A825","#FF4081",
  "#651FFF","#40C4FF","#69F0AE","#CDDC39","#FF9800","#00ACC1","#E040FB","#F50057",
];
function varColor(idx, vi) { return PALETTE[(idx + vi * 8) % PALETTE.length]; }
const DASH_PATTERNS = [[], [6,3], [2,2], [9,3,2,3]];

// ── Variáveis ─────────────────────────────────────────────────────────────────
const VAR_GROUPS = [
  { label:"Potência", vars:[
    { key:"pac",  label:"pac (kW)",   unit:"kW"   },
    { key:"pdc",  label:"pdc (kW)",   unit:"kW"   },
    { key:"qac",  label:"qac (kVAr)", unit:"kVAr" },
  ]},
  { label:"Energia", vars:[
    { key:"dailyEnergyInjected", label:"Injetada (kWh)",  unit:"kWh" },
    { key:"dailyEnergyAbsorbed", label:"Absorvida (kWh)", unit:"kWh" },
  ]},
  { label:"Tensões DC", vars:[
    { key:"vdc",   label:"vdc (V)",   unit:"V" },
    { key:"vbus",  label:"vbus (V)",  unit:"V" },
    { key:"pvbus", label:"pvbus (V)", unit:"V" },
    { key:"nvbus", label:"nvbus (V)", unit:"V" },
  ]},
  { label:"Corrente DC", vars:[{ key:"idc", label:"idc (A)", unit:"A" }]},
  { label:"Tensões AC", vars:[
    { key:"vac1",  label:"vac1 (V)",  unit:"V" },
    { key:"vac2",  label:"vac2 (V)",  unit:"V" },
    { key:"vac3",  label:"vac3 (V)",  unit:"V" },
    { key:"vaux1", label:"vaux1 (V)", unit:"V" },
    { key:"vaux2", label:"vaux2 (V)", unit:"V" },
    { key:"vaux3", label:"vaux3 (V)", unit:"V" },
  ]},
  { label:"Correntes AC", vars:[
    { key:"iac1", label:"iac1 (A)", unit:"A" },
    { key:"iac2", label:"iac2 (A)", unit:"A" },
    { key:"iac3", label:"iac3 (A)", unit:"A" },
  ]},
  { label:"Frequência", vars:[{ key:"fac", label:"fac (Hz)", unit:"Hz" }]},
  { label:"Temperatura", vars:[
    { key:"tempColdCoolant", label:"Coolant (ºC)",    unit:"ºC" },
    { key:"tempStack",       label:"Stack (ºC)",       unit:"ºC" },
    { key:"tempAcCabinet",   label:"Cabinet AC (ºC)", unit:"ºC" },
    { key:"tempDcCabinet",   label:"Cabinet DC (ºC)", unit:"ºC" },
    { key:"tempInductor",    label:"Indutor (ºC)",     unit:"ºC" },
    { key:"tempOutside",     label:"Externa (ºC)",     unit:"ºC" },
  ]},
  { label:"Taxa de Saída", vars:[
    { key:"outHwRate",  label:"outHwRate (%)",  unit:"%" },
    { key:"outPacRate", label:"outPacRate (%)", unit:"%" },
    { key:"outQacRate", label:"outQacRate (%)", unit:"%" },
  ]},
  { label:"Isolamento", vars:[
    { key:"positiveRiso", label:"Riso+ (kΩ)", unit:"kΩ" },
    { key:"negativeRiso", label:"Riso- (kΩ)", unit:"kΩ" },
    { key:"gndBoardIgnd", label:"Ignd (A)",   unit:"A"  },
    { key:"totalCiso",    label:"Ciso (μF)",  unit:"μF" },
  ]},
];
const ALL_VARS = VAR_GROUPS.flatMap(g => g.vars);
const VAR_MAP  = Object.fromEntries(ALL_VARS.map(v => [v.key, v]));
const NONE     = "__none__";

// ── Disponibilidade: constantes ───────────────────────────────────────────────
const AVAIL_END       = "17:30";   // fim do janela de geração
const AVAIL_DEF_START = "05:30";   // início padrão (fallback)
const GEN_THRESHOLD   = 1.0;       // kW — mínimo p/ considerar "gerando"
const MIN_STOP_MINS   = 3;         // paradas ≤ 3 min são descartadas (nuvem/sol nascendo/pondo)
const END_TOLERANCE   = 5;         // ignora parada nos últimos N min do dia (sol se pondo cedo)

// ── Helpers de tempo ──────────────────────────────────────────────────────────
function timeToMins(t) {
  if (!t || typeof t !== "string") return 0;
  const [h="0", m="0"] = t.split(":");
  return parseInt(h,10)*60 + parseInt(m,10);
}
function minsToTime(m) {
  return `${String(Math.floor(m/60)).padStart(2,"0")}:${String(m%60).padStart(2,"0")}`;
}
function fmtMins(mins) {
  if (mins == null || isNaN(mins)) return "—";
  if (mins <= 0) return "0min";
  const h = Math.floor(mins/60);
  const m = Math.round(mins%60);
  if (h===0) return `${m}min`;
  if (m===0) return `${h}h`;
  return `${h}h${String(m).padStart(2,"0")}min`;
}

// ── Disponibilidade: motor de cálculo ────────────────────────────────────────

// Preenche timeline minuto a minuto, gaps herdam estado anterior
function buildTimeline(data, startMins, endMins) {
  const map = {};
  (data||[]).forEach(r => { if (r.time) map[r.time] = r.pac; });
  let lastPac = null;
  const tl = [];
  for (let m = startMins; m <= endMins; m++) {
    const t = minsToTime(m);
    if (map[t] !== undefined) {
      lastPac = map[t];
      tl.push({ time:t, pac:map[t], gap:false });
    } else {
      // Gap: se estava parado (≤0 ou null) → continua parado; senão → herda geração
      const inferred = (lastPac === null || lastPac <= 0) ? 0 : lastPac;
      tl.push({ time:t, pac:inferred, gap:true });
    }
  }
  return tl;
}

// Detecta intervalos de parada numa timeline
function findStopIntervals(timeline) {
  const intervals = [];
  let stopStart = null;
  for (let i = 0; i < timeline.length; i++) {
    const stopped = timeline[i].pac === null || timeline[i].pac <= 0;
    if (stopped && stopStart === null) {
      stopStart = i;
    } else if (!stopped && stopStart !== null) {
      const from = timeline[stopStart].time;
      const to   = timeline[i-1].time;
      intervals.push({ start:from, end:to, mins: timeToMins(to)-timeToMins(from)+1 });
      stopStart = null;
    }
  }
  if (stopStart !== null && timeline.length > 0) {
    const from = timeline[stopStart].time;
    const to   = timeline[timeline.length-1].time;
    intervals.push({ start:from, end:to, mins: timeToMins(to)-timeToMins(from)+1 });
  }
  return intervals;
}

// Detecta o início de cálculo de disponibilidade POR INVERSOR:
// — Se o inversor começou a gerar estritamente antes das 05:59 → conta a partir do seu início real
// — Se começou às 05:59 ou depois (ou nunca gerou) → conta desde 05:30 (indisponível)
function detectInverterStart(data) {
  const defMins = timeToMins(AVAIL_DEF_START); // 05:30
  const cutoff  = timeToMins("05:59");          // exclusivo: só antes de 05:59
  if (!data || !data.length) return defMins;
  const gen = data.find(r =>
    r.time >= AVAIL_DEF_START && r.time <= AVAIL_END &&
    r.pac !== null && r.pac > GEN_THRESHOLD
  );
  if (!gen) return defMins;
  const m = timeToMins(gen.time);
  return m < cutoff ? m : defMins;
}

// Mantido para compatibilidade com chamadas existentes (usa detectInverterStart internamente)
function detectSmartStart(invertersDataArr) {
  if (!invertersDataArr.length) return timeToMins(AVAIL_DEF_START);
  // Retorna o início mais cedo entre os inversores (cada um já calculado individualmente)
  const starts = invertersDataArr.map(detectInverterStart);
  return Math.min(...starts);
}

// Calcula disponibilidade de um inversor num dia
function calcDayAvail(data, smartStartMins) {
  const endMins      = timeToMins(AVAIL_END);
  const tl           = buildTimeline(data, smartStartMins, endMins);
  const rawIntervals = findStopIntervals(tl);

  // Tolerância de fim de dia: descarta parada que começa nos últimos END_TOLERANCE minutos
  // (sol se pondo mais cedo — não é falha do inversor)
  const endToleranceStart = endMins - END_TOLERANCE;
  const intervalsNoEndNoise = rawIntervals.filter(iv => {
    const ivStartMins = timeToMins(iv.start);
    return ivStartMins < endToleranceStart; // mantém só paradas que começaram antes da janela de pôr-do-sol
  });

  // Tolerância de paradas curtas: descarta eventos ≤ MIN_STOP_MINS min
  const intervals = intervalsNoEndNoise.filter(iv => iv.mins > MIN_STOP_MINS);

  const stoppedMins = intervals.reduce((s,iv) => s + iv.mins, 0);
  const stoppedH    = stoppedMins / 60;
  const avail       = isFinite(stoppedH) ? Math.max(0, Math.min(100, (12 - stoppedH) / 12 * 100)) : 0;
  return { availability:avail, stoppedMins, stoppedHours:stoppedH, intervals, rawIntervals };
}

function availColor(pct) {
  if (pct == null) return "var(--color-text-tertiary)";
  if (pct >= 90) return "#4CAF50";
  if (pct >= 70) return "#FF9800";
  return "#F44336";
}

// ── Helpers de parsing ────────────────────────────────────────────────────────
function datetimeToHHMM(val) {
  if (!val) return "00:00";
  if (val instanceof Date) {
    return `${String(val.getUTCHours()).padStart(2,"0")}:${String(val.getUTCMinutes()).padStart(2,"0")}`;
  }
  const s = String(val).trim();
  const part = s.includes(" ") ? s.split(" ")[1] : s;
  const [h="0", m="0"] = part.split(":");
  return `${String(h).padStart(2,"0")}:${String(m).padStart(2,"0")}`;
}
function valToDate(val) {
  if (!val) return null;
  if (val instanceof Date) {
    if (isNaN(val)) return null;
    return `${val.getUTCFullYear()}-${String(val.getUTCMonth()+1).padStart(2,"0")}-${String(val.getUTCDate()).padStart(2,"0")}`;
  }
  const s = String(val).trim();
  const m = s.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (m) return `${m[1]}-${m[2].padStart(2,"0")}-${m[3].padStart(2,"0")}`;
  return null;
}
function buildColMap(headers) {
  const map = {};
  ALL_VARS.forEach(({ key }) => {
    const found = headers.find(h => {
      if (!h) return false;
      const clean = String(h).toLowerCase()
        .replace(/\s*[\[(（].*?[\]）)]/g,"").replace(/[ºãçáéíóú°â]/gi,"").trim();
      return clean === key.toLowerCase();
    });
    if (found != null) map[key] = found;
  });
  return map;
}
function invKeyFromSheet(sheetName, filename) {
  const m = sheetName.match(/node[_\-]([A-Za-z0-9]+?)(?:[_\-]\d+)?$/i);
  if (m) return m[1];
  const base = filename.replace(/\.[^.]+$/, "").replace(/[_\-]\d+$/, "");
  const m2   = base.match(/node[_\-]([A-Za-z0-9]+)/i);
  return m2 ? m2[1] : base.slice(-20);
}
function displayNameFromFile(filename) {
  return filename.replace(/\.[^.]+$/, ""); // remove extension, keep full name
}

async function parseXLSXMultiSheet(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(e.target.result, { type:"array", cellDates:true });
        const results = [];
        wb.SheetNames.forEach(sheetName => {
          const ws   = wb.Sheets[sheetName];
          const rows = XLSX.utils.sheet_to_json(ws, { defval:null });
          if (!rows.length) return;
          const headers = Object.keys(rows[0]);
          const timeCol = headers.find(h => h.toLowerCase().trim()==="time") || headers[0];
          const colMap  = buildColMap(headers);
          let sheetDate = null;
          for (const row of rows) {
            const d = valToDate(row[timeCol]);
            if (d) { sheetDate = d; break; }
          }
          if (!sheetDate) return;
          const invKey     = invKeyFromSheet(sheetName, file.name);
          const displayName = displayNameFromFile(file.name);
          const data = rows.filter(r => r[timeCol]).map(row => {
            const entry = { time: datetimeToHHMM(row[timeCol]) };
            ALL_VARS.forEach(({ key }) => {
              const col = colMap[key];
              const raw = col != null ? row[col] : null;
              const v   = raw !== null && raw !== undefined
                ? (typeof raw==="number" ? raw : parseFloat(String(raw).replace(",",".")))
                : NaN;
              entry[key] = isFinite(v) ? v : null;
            });
            return entry;
          });
          results.push({ invKey, displayName, date:sheetDate, data });
        });
        resolve(results);
      } catch(err) { reject(err); }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function registerSmartTooltip() {
  if (window._smartTooltipRegistered) return;
  window._smartTooltipRegistered = true;
  window.Chart.Tooltip.positioners.smart = function(elements, pos) {
    if (!elements.length) return false;
    const ca = this.chart.chartArea;
    const cx = pos.x - ca.left,  cy = pos.y - ca.top;
    const cw = ca.right - ca.left, ch = ca.bottom - ca.top;
    const goRight  = cx < cw/2;
    const goBottom = cy < ch/2;
    return {
      x: goRight ? ca.right - 20 : ca.left + 20,
      y: goBottom ? ca.bottom - 20 : ca.top + 20,
      xAlign: goRight ? "right" : "left",
      yAlign: goBottom ? "bottom" : "top",
    };
  };
}

// ── Chart libs ────────────────────────────────────────────────────────────────
function loadChartLibs() {
  return new Promise(resolve => {
    if (window._chartReady) { resolve(); return; }
    const load = src => new Promise(res => {
      const s = document.createElement("script"); s.src=src; s.onload=res; document.head.appendChild(s);
    });
    load("https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js")
      .then(()=>load("https://cdnjs.cloudflare.com/ajax/libs/hammer.js/2.0.8/hammer.min.js"))
      .then(()=>load("https://cdnjs.cloudflare.com/ajax/libs/chartjs-plugin-zoom/2.0.1/chartjs-plugin-zoom.min.js"))
      .then(()=>{ window.Chart.register(window.ChartZoom); registerSmartTooltip(); window._chartReady=true; resolve(); });
  });
}

// ── Hook divisor arrastável ───────────────────────────────────────────────────
function useDivider(initial, min, max, dir="horizontal") {
  const [size, setSize] = useState(initial);
  const drag    = useRef(false);
  const start   = useRef({ pos:0, size:0 });
  const sizeRef = useRef(size);
  useEffect(() => { sizeRef.current = size; }, [size]);
  const onMouseDown = useCallback(e => {
    e.preventDefault();
    drag.current  = true;
    start.current = { pos: dir==="horizontal"?e.clientX:e.clientY, size: sizeRef.current };
    const onMove = ev => {
      if (!drag.current) return;
      const raw = (dir==="horizontal"?ev.clientX:ev.clientY) - start.current.pos;
      const delta = dir==="vertical" ? -raw : raw;
      setSize(Math.max(min, Math.min(max, start.current.size + delta)));
    };
    const onUp = () => { drag.current=false; window.removeEventListener("mousemove",onMove); window.removeEventListener("mouseup",onUp); };
    window.addEventListener("mousemove",onMove); window.addEventListener("mouseup",onUp);
  }, [min, max, dir]);   // ← size removed from deps; read via sizeRef instead
  return [size, onMouseDown];
}

// ── Gráfico de Variáveis ──────────────────────────────────────────────────────
function InverterChart({ inverters, varKeys, onPointClick }) {
  const canvasRef = useRef(null);
  const chartRef  = useRef(null);
  const [ready, setReady] = useState(!!window._chartReady);
  const cbRef    = useRef(onPointClick);
  useEffect(()=>{ cbRef.current=onPointClick; },[onPointClick]);
  useEffect(()=>{ if(window._chartReady){registerSmartTooltip();setReady(true);return;} loadChartLibs().then(()=>setReady(true)); },[]);

  useEffect(()=>{
    if(!ready||!canvasRef.current) return;
    if(chartRef.current){chartRef.current.destroy();chartRef.current=null;}
    const visible=inverters.filter(i=>i.visible);
    const av=varKeys.filter(k=>k!==NONE);
    if(!visible.length||!av.length) return;

    const timeSet=new Set();
    visible.forEach(inv=>inv.data.forEach(r=>timeSet.add(r.time)));
    const labels=Array.from(timeSet).sort();

    const units=[...new Set(av.map(k=>VAR_MAP[k]?.unit??""))];
    const axisFor=u=>units.indexOf(u)===0?"y1":"y2";
    const dual=units.length>1;

    const datasets=[];
    visible.forEach(inv=>{
      av.forEach((varKey,vi)=>{
        const info=VAR_MAP[varKey]||{unit:"",label:varKey};
        const color=varColor(inv.paletteIdx,vi);
        const byt={};
        inv.data.forEach(r=>{byt[r.time]=r[varKey];});
        datasets.push({
          label:`${inv.displayName||inv.name} — ${info.label}`,
          invId:inv.id,varKey,vi,
          data:labels.map(t=>{const v=byt[t];return(v!==undefined&&v!==null)?v:null;}),
          borderColor:color,backgroundColor:"transparent",
          borderWidth:1.8,borderDash:DASH_PATTERNS[vi],
          pointRadius:0,pointHoverRadius:5,tension:0.4,spanGaps:false,
          yAxisID:axisFor(info.unit),
        });
      });
    });

    const scales={
      x:{afterBuildTicks:axis=>{const tks=axis.ticks;if(!tks.length)return;const toMin=l=>(l&&l.length>=5)?parseInt(l.slice(0,2),10)*60+parseInt(l.slice(3,5),10):null;const f=toMin(axis.getLabelForValue(tks[0].value)),lt=toMin(axis.getLabelForValue(tks[tks.length-1].value));const span=(f!=null&&lt!=null)?Math.max(1,Math.abs(lt-f)):60;const steps=[1,2,5,10,15,20,30,60,120,180];let step=180;for(const s of steps){if(span/s<=14){step=s;break;}}axis.ticks=tks.filter(tk=>{const m=toMin(axis.getLabelForValue(tk.value));return m!=null&&m%step===0;});},
        ticks:{color:"#8595A6",font:{size:11},maxRotation:0,autoSkip:false,
          callback:function(value){return this.getLabelForValue(value)||"";}},
        grid:{color:"rgba(148,163,184,0.10)"},border:{color:"rgba(148,163,184,0.22)"}},
      y1:{position:"left",ticks:{color:"#8595A6",font:{size:11},maxTicksLimit:8},
        grid:{color:"rgba(148,163,184,0.10)"},border:{color:"rgba(148,163,184,0.22)"},
        title:{display:true,text:units[0]??"",color:"#A7B6C6",font:{size:11}}},
    };
    if(dual) scales.y2={position:"right",ticks:{color:"#666",font:{size:11},maxTicksLimit:8},
      grid:{drawOnChartArea:false},border:{color:"rgba(148,163,184,0.22)"},
      title:{display:true,text:units[1]??"",color:"#666",font:{size:11}}};

    chartRef.current=new window.Chart(canvasRef.current,{
      type:"line",data:{labels,datasets},
      options:{responsive:true,maintainAspectRatio:false,animation:false,
        interaction:{mode:"index",intersect:false},
        onClick:(_,els,ch)=>{if(!els.length)return;cbRef.current?.(ch.data.labels[els[0].index]);},
        plugins:{legend:{display:false},
          tooltip:{position:"smart",backgroundColor:"rgba(38,50,68,0.97)",titleColor:"#A7B6C6",bodyColor:"#EAF2FB",
            borderColor:"rgba(46,155,255,0.40)",borderWidth:1,padding:10,
            callbacks:{title:items=>items[0]?.label??"",
              label:ctx=>{const v=ctx.parsed.y;const info=VAR_MAP[ctx.dataset.varKey];
                return `  ${ctx.dataset.label}: ${v!==null?v.toFixed(3)+" "+(info?.unit??""):"—"}`;},
              labelTextColor:ctx=>ctx.dataset.borderColor}},
          zoom:{pan:{enabled:true,mode:"xy",threshold:10},
            zoom:{wheel:{enabled:true,mode:"xy"},pinch:{enabled:false},
              drag:{enabled:false},
              mode:"xy"}}},
        scales},
    });
    return()=>{if(chartRef.current){chartRef.current.destroy();chartRef.current=null;}};
  },[ready,inverters,varKeys]);

  return(
    <div style={{position:"relative",width:"100%",height:"100%"}}>
      {!ready&&<div style={{position:"absolute",inset:0,display:"flex",alignItems:"center",justifyContent:"center",color:"var(--color-text-tertiary)",fontSize:13}}>Carregando…</div>}
      <canvas ref={canvasRef} role="img" style={{cursor:"crosshair"}} onDoubleClick={()=>chartRef.current?.resetZoom()}/>
    </div>
  );
}

// ── Gráfico de Desbalanceamento ───────────────────────────────────────────────
function ImbalanceChart({ inverters, onPointClick }) {
  const canvasRef = useRef(null);
  const chartRef  = useRef(null);
  const [ready, setReady] = useState(!!window._chartReady);
  const cbRef = useRef(onPointClick);
  useEffect(()=>{cbRef.current=onPointClick;},[onPointClick]);
  useEffect(()=>{if(window._chartReady){registerSmartTooltip();setReady(true);return;}loadChartLibs().then(()=>setReady(true));},[]);

  useEffect(()=>{
    if(!ready||!canvasRef.current) return;
    if(chartRef.current){chartRef.current.destroy();chartRef.current=null;}
    const visible=inverters.filter(i=>i.visible);
    if(!visible.length) return;
    const timeSet=new Set();
    visible.forEach(inv=>inv.data.forEach(r=>timeSet.add(r.time)));
    const labels=Array.from(timeSet).sort();
    const datasets=visible.map(inv=>{
      const byt={};
      inv.data.forEach(r=>{
        if(r.pvbus!==null&&r.nvbus!==null){
          const s=r.pvbus+r.nvbus;byt[r.time]=s!==0?Math.abs((r.pvbus-r.nvbus)/s)*100:0;
        }
      });
      return{label:inv.displayName||inv.name,invId:inv.id,
        data:labels.map(t=>byt[t]!==undefined?byt[t]:null),
        borderColor:inv.color,backgroundColor:"transparent",
        borderWidth:1.8,pointRadius:0,pointHoverRadius:5,tension:0.4,spanGaps:false};
    });
    chartRef.current=new window.Chart(canvasRef.current,{
      type:"line",data:{labels,datasets},
      options:{responsive:true,maintainAspectRatio:false,animation:false,
        interaction:{mode:"index",intersect:false},
        onClick:(_,els,ch)=>{if(!els.length)return;cbRef.current?.(ch.data.labels[els[0].index]);},
        plugins:{legend:{display:false},
          tooltip:{position:"smart",backgroundColor:"rgba(38,50,68,0.97)",titleColor:"#A7B6C6",bodyColor:"#EAF2FB",borderColor:"rgba(46,155,255,0.40)",borderWidth:1,padding:10,
            callbacks:{title:items=>items[0]?.label??"",
              label:ctx=>{const v=ctx.parsed.y;return`  ${ctx.dataset.label}: ${v!==null?v.toFixed(2)+" %":"—"}`;},
              labelTextColor:ctx=>ctx.dataset.borderColor}},
          zoom:{pan:{enabled:true,mode:"xy",threshold:10},
            zoom:{wheel:{enabled:true,mode:"xy"},pinch:{enabled:false},
              drag:{enabled:false},
              mode:"xy"}}},
        scales:{
          x:{afterBuildTicks:axis=>{const tks=axis.ticks;if(!tks.length)return;const toMin=l=>(l&&l.length>=5)?parseInt(l.slice(0,2),10)*60+parseInt(l.slice(3,5),10):null;const f=toMin(axis.getLabelForValue(tks[0].value)),lt=toMin(axis.getLabelForValue(tks[tks.length-1].value));const span=(f!=null&&lt!=null)?Math.max(1,Math.abs(lt-f)):60;const steps=[1,2,5,10,15,20,30,60,120,180];let step=180;for(const s of steps){if(span/s<=14){step=s;break;}}axis.ticks=tks.filter(tk=>{const m=toMin(axis.getLabelForValue(tk.value));return m!=null&&m%step===0;});},
            ticks:{color:"#8595A6",font:{size:11},maxRotation:0,autoSkip:false,
              callback:function(value){return this.getLabelForValue(value)||"";}},
            grid:{color:"rgba(148,163,184,0.10)"},border:{color:"rgba(148,163,184,0.22)"}},
          y:{position:"left",min:0,max:100,
            ticks:{color:"#8595A6",font:{size:11},maxTicksLimit:8,callback:v=>v.toFixed(1)+"%"},
            grid:{color:"rgba(148,163,184,0.10)"},border:{color:"rgba(148,163,184,0.22)"},
            title:{display:true,text:"Desbalanceamento (%)",color:"#A7B6C6",font:{size:11}}}}},
    });
    return()=>{if(chartRef.current){chartRef.current.destroy();chartRef.current=null;}};
  },[ready,inverters]);

  return(
    <div style={{position:"relative",width:"100%",height:"100%"}}>
      {!ready&&<div style={{position:"absolute",inset:0,display:"flex",alignItems:"center",justifyContent:"center",color:"var(--color-text-tertiary)",fontSize:13}}>Carregando…</div>}
      <canvas ref={canvasRef} role="img" style={{cursor:"crosshair"}} onDoubleClick={()=>chartRef.current?.resetZoom()}/>
    </div>
  );
}

// ── Painel de Disponibilidade ─────────────────────────────────────────────────
const MONTH_FULL = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"];

function invName(sn, map) { return (map&&map[sn]) || `SN ${String(sn).slice(-4)}`; }
function invColor(idx) { return PALETTE[idx % PALETTE.length]; }
function useInverterMap() {
  const [map, setMap] = useState({});
  useEffect(()=>{
    let cancelled=false;
    fetch("/api/inverter-map").then(r=>r.ok?r.json():{}).then(m=>{ if(!cancelled) setMap(m); }).catch(()=>{});
    return ()=>{ cancelled=true; };
  },[]);
  return map;
}

function AvailabilityPanel() {
  const invMap = useInverterMap();
  const dates = useMemo(()=>last30Dates(),[]);
  const today = dates[dates.length-1];
  const [viewMode, setViewMode] = useState("daily");
  const [selDate, setSelDate] = useState(today);
  const [selMonth, setSelMonth] = useState(today.slice(0,6));
  const [sampleCache, setSampleCache] = useState({}); // { [ymd]: {[sn]: {time,pac}[]} }
  const [loading, setLoading] = useState(false);
  const [progress, setProgress] = useState({loaded:0,total:0});
  const [fatalError, setFatalError] = useState(null);

  const months = useMemo(()=>[...new Set(dates.map(d=>d.slice(0,6)))].sort(),[dates]);
  const monthIdx = months.indexOf(selMonth);

  const targetDates = useMemo(()=>
    viewMode==="daily" ? [selDate] : dates.filter(d=>d.startsWith(selMonth))
  ,[viewMode, selDate, selMonth, dates]);

  useEffect(()=>{
    let cancelled = false;
    const missing = targetDates.filter(d=>!sampleCache[d]);
    if (!missing.length) return;
    setLoading(true);
    setProgress({loaded:0, total:missing.length});
    (async()=>{
      let loaded=0;
      for (const d of missing) {
        if (cancelled) return;
        try {
          const res = await fetch(`/api/availability?date=${d}`);
          if (!res.ok) {
            const body = await res.json().catch(()=>({}));
            throw new Error(body.error || `Erro ${res.status} ao buscar ${d}`);
          }
          const data = await res.json();
          if (cancelled) return;
          setSampleCache(prev=>({...prev,[d]:data}));
          const cacheStatus = res.headers.get("X-Cache");
          loaded++;
          setProgress({loaded, total:missing.length});
          if (cacheStatus!=="HIT" && d!==missing[missing.length-1]) {
            await new Promise(r=>setTimeout(r,3500));
          }
        } catch(err) {
          if (loaded===0) setFatalError(err.message);
          loaded++;
          setProgress({loaded, total:missing.length});
        }
      }
      if (!cancelled) setLoading(false);
    })();
    return ()=>{ cancelled=true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  },[targetDates]);

  // Início calculado por inversor individualmente (sem smart start global)

  // Disponibilidade diária
  const dailyAvail = useMemo(()=>{
    const samples = sampleCache[selDate];
    if (!samples) return [];
    return Object.entries(samples)
      .map(([sn,data])=>({ sn, name:invName(sn,invMap), ...calcDayAvail(data, detectInverterStart(data)) }))
      .sort((a,b)=>a.name.localeCompare(b.name,undefined,{numeric:true}))
      .map((inv,i)=>({...inv, color:invColor(i)}));
  },[sampleCache, selDate, invMap]);

  // Disponibilidade mensal
  const monthlyAvail = useMemo(()=>{
    const monthDates = dates.filter(d=>d.startsWith(selMonth));
    const bySn = {};
    monthDates.forEach(d=>{
      const samples = sampleCache[d];
      if (!samples) return;
      Object.entries(samples).forEach(([sn,data])=>{
        const res = calcDayAvail(data, detectInverterStart(data));
        (bySn[sn] ??= []).push({date:d, availability:res.availability, stoppedMins:res.stoppedMins, stoppedHours:res.stoppedHours});
      });
    });
    return Object.entries(bySn)
      .map(([sn,dayBreakdown])=>{
        const totalStoppedH = dayBreakdown.reduce((s,x)=>s+x.stoppedHours,0);
        const predicted = dayBreakdown.length*12;
        const avail = predicted>0 ? Math.max(0,Math.min(100,(predicted-totalStoppedH)/predicted*100)) : 0;
        const nm = invName(sn,invMap);
        return { sn, name:nm, invKey:sn, displayName:nm,
          availability:avail, stoppedHours:totalStoppedH, days:dayBreakdown.length, dayBreakdown };
      })
      .sort((a,b)=>a.name.localeCompare(b.name,undefined,{numeric:true}))
      .map((inv,i)=>({...inv, color:invColor(i)}));
  },[sampleCache, selMonth, dates, invMap]);

  const items = viewMode==="daily" ? dailyAvail : monthlyAvail;
  const dateIdx = dates.indexOf(selDate);
  const goPrevDay = () => { if(dateIdx>0) setSelDate(dates[dateIdx-1]); };
  const goNextDay = () => { if(dateIdx<dates.length-1) setSelDate(dates[dateIdx+1]); };
  const noData = !items.length && !loading;

  return(
    <div style={{flex:1,display:"flex",flexDirection:"column",overflow:"hidden",padding:"10px 16px 8px",minHeight:0}}>
      {/* ── Header ── */}
      <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:10,flexShrink:0,flexWrap:"wrap"}}>
        {/* Toggle */}
        <div style={{display:"flex",gap:2,background:"var(--color-background-secondary)",padding:3,borderRadius:8,flexShrink:0}}>
          {[{id:"daily",icon:"ti-calendar-day",label:"Diária"},{id:"monthly",icon:"ti-calendar-month",label:"Mensal"}].map(v=>(
            <button key={v.id} onClick={()=>setViewMode(v.id)}
              style={{padding:"5px 14px",fontSize:13,cursor:"pointer",borderRadius:6,border:"none",
                background:viewMode===v.id?"#2E9BFF":"transparent",
                color:viewMode===v.id?"#fff":"var(--color-text-secondary)",
                fontWeight:viewMode===v.id?600:400,display:"flex",alignItems:"center",gap:5,transition:"all 0.15s"}}>
              <i className={`ti ${v.icon}`} style={{fontSize:13}}/>{v.label}
            </button>
          ))}
        </div>

        {/* Contexto */}
        {viewMode==="daily"&&(
          <div style={{display:"flex",alignItems:"center",gap:4}}>
            <button onClick={goPrevDay} disabled={dateIdx<=0}
              style={{background:"none",border:"none",cursor:dateIdx>0?"pointer":"default",fontSize:15,
                color:dateIdx>0?"var(--color-text-secondary)":"var(--color-text-tertiary)",padding:"2px 4px"}}>
              <i className="ti ti-chevron-left"/>
            </button>
            <span style={{fontSize:13,fontWeight:600,minWidth:88,textAlign:"center",fontFamily:"var(--font-mono)"}}>
              {fmtYmd(selDate)}
            </span>
            <button onClick={goNextDay} disabled={dateIdx>=dates.length-1}
              style={{background:"none",border:"none",cursor:dateIdx<dates.length-1?"pointer":"default",fontSize:15,
                color:dateIdx<dates.length-1?"var(--color-text-secondary)":"var(--color-text-tertiary)",padding:"2px 4px"}}>
              <i className="ti ti-chevron-right"/>
            </button>
          </div>
        )}
        {viewMode==="monthly"&&(
          <div style={{display:"flex",alignItems:"center",gap:4}}>
            <button onClick={()=>{if(monthIdx>0)setSelMonth(months[monthIdx-1]);}}
              disabled={monthIdx<=0}
              style={{background:"none",border:"none",cursor:monthIdx>0?"pointer":"default",
                fontSize:14,color:monthIdx>0?"var(--color-text-secondary)":"var(--color-text-tertiary)",padding:"2px 4px"}}>
              <i className="ti ti-chevron-left"/>
            </button>
            <span style={{fontSize:13,fontWeight:600,minWidth:130,textAlign:"center"}}>
              {selMonth?`${MONTH_FULL[parseInt(selMonth.slice(4,6))-1]} ${selMonth.slice(0,4)}`:"—"}
            </span>
            <button onClick={()=>{if(monthIdx<months.length-1)setSelMonth(months[monthIdx+1]);}}
              disabled={monthIdx>=months.length-1}
              style={{background:"none",border:"none",cursor:monthIdx<months.length-1?"pointer":"default",
                fontSize:14,color:monthIdx<months.length-1?"var(--color-text-secondary)":"var(--color-text-tertiary)",padding:"2px 4px"}}>
              <i className="ti ti-chevron-right"/>
            </button>
            {monthlyAvail.length>0&&<span style={{fontSize:13,color:"var(--color-text-tertiary)"}}>{monthlyAvail[0]?.days}d</span>}
          </div>
        )}

        {loading&&(
          <div style={{display:"flex",alignItems:"center",gap:6,fontSize:13,color:"var(--color-text-tertiary)"}}>
            <div style={{width:80,height:5,borderRadius:3,background:"var(--color-background-secondary)",overflow:"hidden"}}>
              <div style={{width:`${(progress.loaded/Math.max(1,progress.total)*100).toFixed(0)}%`,height:"100%",
                background:"#2E9BFF",transition:"width 0.3s"}}/>
            </div>
            <span>carregando {progress.loaded}/{progress.total}</span>
          </div>
        )}

      </div>

      {fatalError?(
        <div style={{flex:1,display:"flex",alignItems:"center",justifyContent:"center",
          color:"var(--color-text-tertiary)",fontSize:13,textAlign:"center"}}>
          <div>
            <i className="ti ti-plug-connected-x" style={{fontSize:40,display:"block",marginBottom:8}}/>
            {fatalError}
          </div>
        </div>
      ):noData?(
        <div style={{flex:1,display:"flex",alignItems:"center",justifyContent:"center",
          color:"var(--color-text-tertiary)",fontSize:13,textAlign:"center"}}>
          <div>
            <i className="ti ti-chart-bar-off" style={{fontSize:40,display:"block",marginBottom:8}}/>
            {viewMode==="daily"?"Sem dados para este dia":"Nenhum dado para o mês selecionado"}
          </div>
        </div>
      ):(
        <div style={{flex:1,display:"flex",gap:0,overflow:"hidden",minHeight:0}}>

          {/* ── Coluna 1: Barras de disponibilidade ── */}
          <div style={{flex:"0 0 48%",display:"flex",flexDirection:"column",overflow:"hidden",paddingRight:12}}>
            <div style={{fontSize:13,fontWeight:600,color:"var(--color-text-tertiary)",textTransform:"uppercase",
              letterSpacing:"0.05em",marginBottom:8,flexShrink:0}}>
              {viewMode==="daily"?"Disponibilidade Diária (12h previstas)":"Disponibilidade Mensal"}
            </div>

            {/* Cabeçalho das linhas de referência */}
            <div style={{position:"relative",marginLeft:140,marginRight:56,height:14,flexShrink:0}}>
              {[0,25,50,75,90,100].map(p=>(
                <span key={p} style={{position:"absolute",left:`${p}%`,transform:"translateX(-50%)",
                  fontSize:13,color:"var(--color-text-tertiary)"}}>{p}%</span>
              ))}
            </div>

            <div style={{overflowY:"auto",flex:1,paddingRight:4}}>
              {items.map((inv)=>(
                <div key={inv.invKey||inv.id} style={{marginBottom:viewMode==="monthly"?4:8}}>
                  <div style={{display:"flex",alignItems:"center",gap:8}}>
                    {/* Label inversor */}
                    <div style={{width:140,flexShrink:0,display:"flex",alignItems:"center",gap:5}}>
                      <span style={{width:8,height:8,borderRadius:"50%",background:inv.color,flexShrink:0}}></span>
                      <span style={{fontSize:13,fontWeight:500,overflow:"hidden",textOverflow:"ellipsis",
                        whiteSpace:"nowrap",color:"var(--color-text-primary)"}}
                        title={inv.displayName||inv.name}>{inv.displayName||inv.name}</span>
                    </div>
                    {/* Barra */}
                    <div style={{flex:1,position:"relative",height:viewMode==="monthly"?16:24,background:"var(--color-background-secondary)",
                      borderRadius:5,overflow:"hidden"}}>
                      {[25,50,75].map(p=>(
                        <div key={p} style={{position:"absolute",left:`${p}%`,top:0,bottom:0,
                          width:1,background:"rgba(0,0,0,0.08)",zIndex:1,pointerEvents:"none"}}></div>
                      ))}
                      {/* Linha 90% */}
                      <div style={{position:"absolute",left:"90%",top:0,bottom:0,
                        width:1.5,background:"rgba(76,175,80,0.5)",zIndex:1,pointerEvents:"none"}}></div>
                      <div style={{
                        position:"absolute",left:0,top:0,bottom:0,
                        width:`${Math.min(100,inv.availability??0).toFixed(1)}%`,
                        background:availColor(inv.availability),opacity:0.88,
                        borderRadius:5,zIndex:2,display:"flex",alignItems:"center",
                        paddingLeft:8,transition:"width 0.4s ease",
                      }}>
                        {inv.availability>=18&&(
                          <span style={{fontSize:13,color:"#fff",fontWeight:700,whiteSpace:"nowrap"}}>
                            {(inv.availability??0).toFixed(1)}%
                          </span>
                        )}
                      </div>
                    </div>
                    {/* Valor & parado */}
                    <div style={{width:72,flexShrink:0,textAlign:"right"}}>
                      {inv.availability<18&&(
                        <div style={{fontSize:13,fontWeight:700,color:availColor(inv.availability),lineHeight:1.2}}>
                          {(inv.availability??0).toFixed(1)}%
                        </div>
                      )}
                      <div style={{fontSize:13,color:"var(--color-text-tertiary)",lineHeight:1.3}}>
                        {fmtMins(
                          viewMode==="daily"
                            ? (inv.stoppedMins ?? 0)
                            : Math.round((inv.stoppedHours ?? 0) * 60)
                        )} parado
                      </div>
                    </div>
                  </div>

                  {/* Breakdown diário (mensal) */}
                  {viewMode==="monthly"&&inv.dayBreakdown&&(
                    <div style={{marginLeft:148,marginTop:2,display:"flex",flexWrap:"wrap",gap:2}}>
                      {inv.dayBreakdown.map(db=>(
                        <span key={db.date} title={`${db.date}: ${(db.availability??0).toFixed(1)}% (${fmtMins(db.stoppedMins)} parado)`}
                          style={{fontSize:13,padding:"1px 5px",borderRadius:3,cursor:"default",
                            background:`${availColor(db.availability)}1A`,
                            color:availColor(db.availability),
                            border:`1px solid ${availColor(db.availability)}44`}}>
                          {db.date.slice(6,8)} {(db.availability??0).toFixed(0)}%
                        </span>
                      ))}
                    </div>
                  )}
                </div>
              ))}
            </div>
          </div>

          {/* ── Divisor ── */}
          <div style={{width:1,background:"var(--color-border-tertiary)",flexShrink:0,margin:"0 2px"}}></div>

          {/* ── Coluna 2: Intervalos de parada (diária) ou resumo mensal ── */}
          <div style={{flex:1,display:"flex",flexDirection:"column",overflow:"hidden",paddingLeft:12}}>
            {viewMode==="daily"?(
              <>
                <div style={{fontSize:13,fontWeight:600,color:"var(--color-text-tertiary)",textTransform:"uppercase",
                  letterSpacing:"0.05em",marginBottom:8,flexShrink:0}}>
                  Intervalos sem geração (05:30 – 17:30)
                </div>
                <div style={{overflowY:"auto",flex:1,display:"flex",flexDirection:"column",gap:8}}>
                  {dailyAvail.map(inv=>(
                    <div key={inv.id||inv.invKey}
                      style={{background:"var(--color-background-secondary)",borderRadius:8,
                        padding:"8px 10px",borderLeft:`3px solid ${availColor(inv.availability)}`}}>
                      <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:inv.intervals.length?5:0}}>
                        <span style={{width:7,height:7,borderRadius:"50%",background:inv.color,flexShrink:0}}></span>
                        <span style={{fontSize:13,fontWeight:600,color:"var(--color-text-primary)",
                          overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",flex:1}}
                          title={inv.displayName||inv.name}>{inv.displayName||inv.name}</span>
                        <span style={{fontSize:13,fontWeight:700,color:availColor(inv.availability),flexShrink:0}}>
                          {(inv.availability??0).toFixed(1)}%
                        </span>
                      </div>
                      {!inv.intervals.length?(
                        <div style={{fontSize:13,color:"#4CAF50",display:"flex",alignItems:"center",gap:4}}>
                          <i className="ti ti-check" style={{fontSize:13}}/>Nenhuma parada relevante
                          {inv.rawIntervals?.length>0&&(
                            <span style={{fontSize:13,color:"var(--color-text-tertiary)",marginLeft:4}}>
                              ({inv.rawIntervals.length} ignorada(s): ≤{MIN_STOP_MINS}min ou fim do dia)
                            </span>
                          )}
                        </div>
                      ):(
                        <div style={{display:"flex",flexDirection:"column",gap:2}}>
                          {inv.intervals.map((iv,idx)=>(
                            <div key={idx} style={{display:"flex",alignItems:"center",gap:6,fontSize:13,
                              color:"var(--color-text-secondary)"}}>
                              <i className="ti ti-clock-off" style={{fontSize:13,color:"#F44336",flexShrink:0}}/>
                              <span style={{fontFamily:"var(--font-mono)"}}>{iv.start} – {iv.end}</span>
                              <span style={{color:"var(--color-text-tertiary)",flexShrink:0}}>({fmtMins(iv.mins)})</span>
                            </div>
                          ))}
                          <div style={{fontSize:13,color:"var(--color-text-tertiary)",marginTop:3,paddingTop:3,
                            borderTop:"0.5px solid var(--color-border-tertiary)",display:"flex",justifyContent:"space-between"}}>
                            <span>Total parado: <strong style={{color:availColor(inv.availability)}}>{fmtMins(inv.stoppedMins)}</strong></span>
                          {inv.rawIntervals?.length > inv.intervals.length && (
                              <span style={{fontSize:13,color:"var(--color-text-tertiary)"}}>
                                +{inv.rawIntervals.length - inv.intervals.length} ignorada(s) (≤{MIN_STOP_MINS}min ou fim do dia)
                              </span>
                            )}
                          </div>
                        </div>
                      )}
                    </div>
                  ))}
                </div>
              </>
            ):(
              <>
                <div style={{fontSize:13,fontWeight:600,color:"var(--color-text-tertiary)",textTransform:"uppercase",
                  letterSpacing:"0.05em",marginBottom:8,flexShrink:0}}>
                  Resumo Mensal
                </div>
                <div style={{overflowY:"auto",flex:1}}>
                  <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                    <thead style={{position:"sticky",top:0,zIndex:1,background:"var(--color-background-primary)"}}>
                      <tr>
                        {["Inversor","Dias","Total parado","Disp. Mensal"].map(h=>(
                          <th key={h} style={{padding:"3px 8px",textAlign:h==="Inversor"?"left":"right",
                            fontWeight:600,fontSize:13,color:"var(--color-text-secondary)",
                            borderBottom:"0.5px solid var(--color-border-tertiary)"}}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {monthlyAvail.map((inv,i)=>(
                        <tr key={inv.invKey} style={{background:i%2===1?"var(--color-background-secondary)":"transparent"}}>
                          <td style={{padding:"2px 8px"}}>
                            <div style={{display:"flex",alignItems:"center",gap:5}}>
                              <span style={{width:7,height:7,borderRadius:"50%",background:inv.color,flexShrink:0}}></span>
                              <span style={{fontSize:13,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",maxWidth:120}}
                                title={inv.displayName}>{inv.displayName}</span>
                            </div>
                          </td>
                          <td style={{padding:"2px 8px",textAlign:"right",fontSize:13,color:"var(--color-text-tertiary)"}}>{inv.days}</td>
                          <td style={{padding:"2px 8px",textAlign:"right",fontSize:13,fontFamily:"var(--font-mono)",color:"var(--color-text-secondary)"}}>
                            {fmtMins(Math.round(inv.stoppedHours*60))}
                          </td>
                          <td style={{padding:"2px 8px",textAlign:"right",fontSize:13,fontWeight:700,
                            color:availColor(inv.availability),fontFamily:"var(--font-mono)"}}>
                            {(inv.availability??0).toFixed(1)}%
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </>
            )}
          </div>
        </div>
      )}
    </div>
  );
}

// ── Ordenação de tabelas ──────────────────────────────────────────────────────
// Hook: retorna [sortKey, sortDir, toggle(key), sortFn]
// Ciclo: asc(↑) → desc(↓) → alpha(AZ) → asc
function useSortableTable(defaultKey="", defaultDir="asc") {
  const [sortKey, setSortKey] = useState(defaultKey);
  const [sortDir, setSortDir] = useState(defaultDir);
  const toggle = useCallback(key => {
    setSortKey(prev => {
      if (prev !== key) { setSortDir("asc"); return key; }
      setSortDir(d => d==="asc" ? "desc" : d==="desc" ? "alpha" : "asc");
      return key;
    });
  }, []);
  return { sortKey, sortDir, toggle };
}

function applySortDir(rows, sortKey, sortDir) {
  if (!sortKey) return rows;
  return [...rows].sort((a, b) => {
    if (sortDir === "alpha") {
      const an = String(a.name ?? a.invName ?? "");
      const bn = String(b.name ?? b.invName ?? "");
      return an.localeCompare(bn, undefined, {numeric:true, sensitivity:"base"});
    }
    const av = a[sortKey], bv = b[sortKey];
    if (av == null && bv == null) return 0;
    if (av == null) return 1;
    if (bv == null) return -1;
    const cmp = typeof av==="number" && typeof bv==="number"
      ? av - bv
      : String(av).localeCompare(String(bv), undefined, {numeric:true, sensitivity:"base"});
    return sortDir === "asc" ? cmp : -cmp;
  });
}

function SortTh({ label, sortKey, col, sortDir, onSort, align="right" }) {
  const active = sortKey === col;
  const icon   = !active ? "↕" : sortDir==="asc" ? "↑" : sortDir==="desc" ? "↓" : "AZ";
  const tip    = !active ? "1°clique: menor→maior"
    : sortDir==="asc"  ? "2°clique: maior→menor"
    : sortDir==="desc" ? "3°clique: ordem A→Z / 0→9"
    : "1°clique: menor→maior";
  return (
    <th onClick={() => onSort(col)} title={tip}
      style={{padding:"4px 6px", textAlign:align, fontWeight:600, fontSize:11,
        color: active ? "var(--color-text-primary)" : "var(--color-text-secondary)",
        borderBottom:"0.5px solid var(--color-border-tertiary)",
        cursor:"pointer", userSelect:"none", whiteSpace:"nowrap", transition:"color 0.12s"}}>
      {label}
      <span style={{marginLeft:3, fontSize:11, opacity:active?1:0.25,
        color:active?"#2E9BFF":"inherit", fontFamily:"monospace"}}>
        {icon}
      </span>
    </th>
  );
}

// ── VarSelect ─────────────────────────────────────────────────────────────────
function VarSelect({ id, value, onChange, label, swatchColor, dashPattern, optional }) {
  return(
    <div style={{display:"flex",alignItems:"center",gap:5}}>
      <svg width="18" height="10" style={{flexShrink:0}}>
        <line x1="0" y1="5" x2="18" y2="5" stroke={swatchColor} strokeWidth="2.5"
          strokeDasharray={(dashPattern||[]).join(",")}/>
      </svg>
      <label htmlFor={id} style={{fontSize:13,color:"var(--color-text-secondary)",flexShrink:0}}>{label}</label>
      <select id={id} value={value} onChange={e=>onChange(e.target.value)}
        style={{fontSize:12,maxWidth:170,padding:"3px 8px",borderRadius:7,cursor:"pointer",outline:"none",
          background:"#1A2433",color:"#E6EDF5",
          border:`1px solid ${value!==NONE?swatchColor+"99":"rgba(255,255,255,0.12)"}`}}>
        {optional&&<option value={NONE} style={{background:"#18222F",color:"#E6EDF5"}}>— nenhuma —</option>}
        {VAR_GROUPS.map(g=>(
          <optgroup key={g.label} label={g.label} style={{background:"#18222F",color:"#9FB0C0"}}>
            {g.vars.map(v=><option key={v.key} value={v.key} style={{background:"#18222F",color:"#E6EDF5"}}>{v.label}</option>)}
          </optgroup>
        ))}
      </select>
    </div>
  );
}

// ── Calendário ────────────────────────────────────────────────────────────────
const WD=["D","S","T","Q","Q","S","S"];
function MiniCalendar({ datesSet, selectedDate, onSelect, invColors }) {
  const today=new Date();
  const [vy,setVy]=useState(()=>selectedDate?parseInt(selectedDate.slice(0,4)):today.getFullYear());
  const [vm,setVm]=useState(()=>selectedDate?parseInt(selectedDate.slice(5,7))-1:today.getMonth());

  // Segue o selectedDate se mudar para um mês diferente do visível
  useEffect(()=>{
    if(!selectedDate) return;
    const y=parseInt(selectedDate.slice(0,4));
    const m=parseInt(selectedDate.slice(5,7))-1;
    if(y!==vy||m!==vm){ setVy(y); setVm(m); }
  },[selectedDate]); // eslint-disable-line
  const first=new Date(vy,vm,1).getDay();
  const days=new Date(vy,vm+1,0).getDate();
  const cells=[...Array(first).fill(null),...Array.from({length:days},(_,i)=>i+1)];
  while(cells.length%7) cells.push(null);
  const todayStr=`${today.getFullYear()}-${String(today.getMonth()+1).padStart(2,"0")}-${String(today.getDate()).padStart(2,"0")}`;
  const MSHORT=["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"];
  return(
    <div style={{padding:"8px 10px"}}>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:5}}>
        <button onClick={()=>{if(vm===0){setVy(y=>y-1);setVm(11);}else setVm(m=>m-1);}}
          style={{background:"none",border:"none",cursor:"pointer",fontSize:14,color:"var(--color-text-secondary)",padding:"2px 4px"}}>
          <i className="ti ti-chevron-left"/>
        </button>
        <span style={{fontWeight:600,fontSize:13,color:"var(--color-text-primary)"}}>{MSHORT[vm]} {vy}</span>
        <button onClick={()=>{if(vm===11){setVy(y=>y+1);setVm(0);}else setVm(m=>m+1);}}
          style={{background:"none",border:"none",cursor:"pointer",fontSize:14,color:"var(--color-text-secondary)",padding:"2px 4px"}}>
          <i className="ti ti-chevron-right"/>
        </button>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:1}}>
        {WD.map((w,i)=>(
          <div key={i} style={{textAlign:"center",fontSize:8,fontWeight:700,color:"var(--color-text-tertiary)",padding:"2px 0",textTransform:"uppercase"}}>{w}</div>
        ))}
        {cells.map((day,i)=>{
          if(!day) return<div key={`e${i}`}></div>;
          const ds=`${vy}-${String(vm+1).padStart(2,"0")}-${String(day).padStart(2,"0")}`;
          const has=datesSet.has(ds),sel=ds===selectedDate,tod=ds===todayStr;
          const dots=invColors[ds]||[];
          return(
            <div key={ds} onClick={()=>has&&onSelect(ds)}
              style={{textAlign:"center",borderRadius:5,padding:"3px 1px 2px",cursor:has?"pointer":"default",
                background:sel?"#2E9BFF":tod?"rgba(46,155,255,0.1)":has?"rgba(0,0,0,0.035)":"transparent",
                border:tod&&!sel?"1px solid rgba(46,155,255,0.35)":"1px solid transparent",
                transition:"background 0.1s"}}
              onMouseEnter={e=>{if(has&&!sel)e.currentTarget.style.background="rgba(46,155,255,0.18)";}}
              onMouseLeave={e=>{if(has&&!sel)e.currentTarget.style.background=has?"rgba(0,0,0,0.035)":"transparent";}}>
              <div style={{fontSize:13,fontWeight:sel||tod?700:has?500:400,lineHeight:1.2,
                color:sel?"#fff":tod?"#1B6FC9":has?"var(--color-text-primary)":"var(--color-text-tertiary)"}}>{day}</div>
              {has&&dots.length>0&&(
                <div style={{display:"flex",justifyContent:"center",gap:1,marginTop:1}}>
                  {dots.slice(0,4).map((c,ii)=>(
                    <span key={ii} style={{width:3,height:3,borderRadius:"50%",
                      background:sel?"rgba(255,255,255,0.8)":c,display:"block"}}></span>
                  ))}
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}

// ── Dashboard principal ───────────────────────────────────────────────────────
function App() {
  const [allData,      setAllData]      = useState({});
  const [selectedDate, setSelectedDate] = useState(null);
  const [vars,         setVars]         = useState(["pac",NONE,NONE,NONE]);
  const [selectedTime, setSelectedTime] = useState(null);
  const [activeTab,    setActiveTab]    = useState("vars");
  const [isDragging,   setIsDragging]   = useState(false);
  const [loading,      setLoading]      = useState(false);
  const [errors,       setErrors]       = useState([]);
  const [showStats,    setShowStats]    = useState(true);
  const [hidden,       setHidden]       = useState(new Set());
  const fileRef = useRef();

  const [sideW,  onSideDrag]   = useDivider(255,180,420,"horizontal");
  const [statsH, onStatsDrag]  = useDivider(190,80,440,"vertical");

  // Ícone da aba (favicon) + título — tema solar/performance
  useEffect(()=>{
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32">
      <rect width="32" height="32" rx="7" fill="#0F1722"/>
      <g stroke="#FFC42E" stroke-width="2" stroke-linecap="round">
        <line x1="16" y1="3" x2="16" y2="6.5"/>
        <line x1="16" y1="16.5" x2="16" y2="20"/>
        <line x1="7.7" y1="11.7" x2="10.2" y2="14.2"/>
        <line x1="21.8" y1="11.7" x2="19.3" y2="14.2"/>
        <line x1="3.5" y1="11.5" x2="7" y2="11.5"/>
        <line x1="25" y1="11.5" x2="28.5" y2="11.5"/>
        <line x1="7.7" y1="3.3" x2="10.2" y2="5.8"/>
        <line x1="21.8" y1="3.3" x2="19.3" y2="5.8"/>
      </g>
      <circle cx="16" cy="11.5" r="4.3" fill="#FFC42E"/>
      <polyline points="4,28 11,22 16,25 28,15" fill="none" stroke="#2E9BFF"
        stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round"/>
      <circle cx="28" cy="15" r="2.1" fill="#2E9BFF"/>
    </svg>`;
    const href = "data:image/svg+xml," + encodeURIComponent(svg);
    let link = document.querySelector("link[rel~='icon']");
    if(!link){ link = document.createElement("link"); link.rel = "icon"; document.head.appendChild(link); }
    link.type = "image/svg+xml";
    link.href = href;
    const prevTitle = document.title;
    document.title = "Airbox · Monitoramento Solar";
    return ()=>{ document.title = prevTitle; };
  },[]);

  // Datas disponíveis
  const datesSet = useMemo(()=>{
    const s=new Set();
    Object.values(allData).forEach(inv=>Object.keys(inv.dates).forEach(d=>s.add(d)));
    return s;
  },[allData]);
  const sortedDates=useMemo(()=>[...datesSet].sort(),[datesSet]);

  // Cores por data para calendário
  const calDotColors=useMemo(()=>{
    const map={};
    sortedDates.forEach(d=>{
      map[d]=Object.values(allData).filter(inv=>inv.dates[d]).map(inv=>inv.color);
    });
    return map;
  },[allData,sortedDates]);

  // Inversores do dia selecionado
  const inverters=useMemo(()=>{
    if(!selectedDate) return[];
    return Object.values(allData)
      .filter(inv=>inv.dates[selectedDate])
      .map(inv=>({
        id:`${inv.invKey}-${selectedDate}`,
        invKey:inv.invKey,
        name:inv.displayName||inv.invKey,
        displayName:inv.displayName||inv.invKey,
        color:inv.color,paletteIdx:inv.paletteIdx,
        visible:!hidden.has(inv.invKey),
        data:inv.dates[selectedDate],
      }));
  },[allData,selectedDate,hidden]);

  // Carregar arquivos
  const addFiles=useCallback(async(files)=>{
    setLoading(true);setErrors([]);
    const errs=[];
    for(const file of Array.from(files)){
      const ext=file.name.split(".").pop().toLowerCase();
      if(ext!=="xlsx"&&ext!=="xls"){errs.push(`"${file.name}": use .xlsx`);continue;}
      try{
        const sheets=await parseXLSXMultiSheet(file);
        if(!sheets.length){errs.push(`"${file.name}": nenhuma planilha com data válida`);continue;}
        setAllData(prev=>{
          const next={...prev};
          sheets.forEach(({invKey,displayName,date,data})=>{
            if(!next[invKey]){
              const idx=Object.keys(next).length%PALETTE.length;
              next[invKey]={invKey,displayName,color:varColor(idx,0),paletteIdx:idx,dates:{}};
            }
            next[invKey]={...next[invKey],dates:{...next[invKey].dates,[date]:data}};
          });
          return next;
        });
        // Navega sempre para a data mais recente nos arquivos carregados
        const newDates = [...new Set(sheets.map(s=>s.date).filter(Boolean))].sort();
        if (newDates.length) setSelectedDate(prev => {
          // Mantém data mais recente entre a atual e as novas
          return (!prev || newDates.at(-1) > prev) ? newDates.at(-1) : prev;
        });
      }catch(err){errs.push(`"${file.name}": ${err.message}`);}
    }
    if(errs.length) setErrors(errs);
    setLoading(false);
  },[]);

  const onDrop=useCallback(e=>{e.preventDefault();setIsDragging(false);addFiles(e.dataTransfer.files);},[addFiles]);

  // Navegação de datas
  const curIdx=selectedDate?sortedDates.indexOf(selectedDate):-1;
  const prevDate=()=>{if(curIdx>0){setSelectedDate(sortedDates[curIdx-1]);setSelectedTime(null);}};
  const nextDate=()=>{if(curIdx<sortedDates.length-1){setSelectedDate(sortedDates[curIdx+1]);setSelectedTime(null);}};

  // Visibilidade
  const toggleInv=invKey=>setHidden(prev=>{const n=new Set(prev);n.has(invKey)?n.delete(invKey):n.add(invKey);return n;});
  const toggleAll=()=>{
    const ks=Object.keys(allData);
    setHidden(ks.some(k=>!hidden.has(k))?new Set(ks):new Set());
  };
  const clearAll=()=>{setAllData({});setSelectedDate(null);setSelectedTime(null);setHidden(new Set());setErrors([]);};
  const toggleLegend=invKey=>{
    const visKeys=Object.keys(allData).filter(k=>!hidden.has(k));
    if(hidden.has(invKey)){setHidden(prev=>{const n=new Set(prev);n.delete(invKey);return n;});}
    else if(visKeys.length>1){setHidden(new Set(Object.keys(allData).filter(k=>k!==invKey)));}
    else{setHidden(new Set());}
  };

  const setVar=(vi,key)=>setVars(prev=>prev.map((k,i)=>i===vi?key:k));
  const visible       = inverters.filter(i=>i.visible);
  const activeVars    = vars.filter(k=>k!==NONE);
  const selectorColors= [0,1,2,3].map(vi=>varColor(0,vi));
  const activeUnits   = [...new Set(activeVars.map(k=>VAR_MAP[k]?.unit??""))];
  const dualAxis      = activeUnits.length>1;
  const hasData       = Object.keys(allData).length>0;

  // Estatísticas (só nas abas var/imbalance)
  const stats=useMemo(()=>inverters.map(inv=>({
    ...inv,
    varStats:activeVars.map(key=>{
      const vals=(inv.data||[]).map(r=>r[key]).filter(v=>v!==null&&isFinite(v));
      if(!vals.length) return{key,min:null,max:null,avg:null};
      return{key,min:Math.min(...vals),max:Math.max(...vals),avg:vals.reduce((a,b)=>a+b,0)/vals.length};
    }),
  })),[inverters,activeVars.join(",")]);

  // Snapshot
  const snapshot=useMemo(()=>{
    if(!selectedTime||!inverters.length||!activeVars.length) return null;
    return inverters.map(inv=>{
      const row=(inv.data||[]).find(r=>r.time===selectedTime);
      return{...inv,values:activeVars.map(key=>({key,value:row?row[key]:null}))};
    });
  },[selectedTime,inverters,activeVars.join(",")]);

  // Desbalanceamento
  const IMBST="06:30",IMBEN="17:00";
  const imbalanceData=useMemo(()=>inverters.map(inv=>{
    let atPoint=null;
    if(selectedTime){
      const row=(inv.data||[]).find(r=>r.time===selectedTime);
      if(row&&row.pvbus!==null&&row.nvbus!==null){
        const s=row.pvbus+row.nvbus;atPoint=s!==0?Math.abs((row.pvbus-row.nvbus)/s)*100:0;
      }
    }
    const vals=(inv.data||[]).filter(r=>r.time>=IMBST&&r.time<=IMBEN&&r.pvbus!==null&&r.nvbus!==null)
      .map(r=>{const s=r.pvbus+r.nvbus;return s!==0?Math.abs((r.pvbus-r.nvbus)/s)*100:0;});
    return{...inv,atPoint,min:vals.length?Math.min(...vals):null,max:vals.length?Math.max(...vals):null,
      avg:vals.length?vals.reduce((a,b)=>a+b,0)/vals.length:null};
  }),[selectedTime,inverters]);

  const fmt=v=>v!==null?Number(v).toFixed(3):"—";
  const fmtPct=v=>v!==null?v.toFixed(2)+" %":"—";
  const alertC=v=>v===null?"var(--color-text-tertiary)":v<35?"#4CAF50":v<70?"#FF9800":"#F44336";

  const showFileTab = activeTab==="vars" || activeTab==="imbalance";
  const showBottomPanel = showFileTab && stats.length>0;

  return(
    <div style={{display:"flex",flexDirection:"column",height:"100vh",width:"100vw",overflow:"hidden",fontFamily:"var(--font-sans)",
      background:"#0F1722",
      "--color-background-primary":"#0F1722",
      "--color-background-secondary":"#18222F",
      "--color-background-danger":"rgba(244,63,94,0.14)",
      "--color-text-primary":"#E6EDF5",
      "--color-text-secondary":"#9FB0C0",
      "--color-text-tertiary":"#647386",
      "--color-text-danger":"#FF6B6B",
      "--color-border-tertiary":"rgba(255,255,255,0.07)",
      "--color-border-secondary":"rgba(255,255,255,0.13)"}}>

      {/* ── Barras de rolagem: ocultas por padrão, aparecem ao passar o cursor ── */}
      <style dangerouslySetInnerHTML={{__html:`
        ::-webkit-scrollbar{width:10px;height:10px;}
        ::-webkit-scrollbar-track{background:transparent;}
        ::-webkit-scrollbar-corner{background:transparent;}
        ::-webkit-scrollbar-thumb{
          background-color:transparent;border-radius:8px;
          border:2px solid transparent;background-clip:content-box;
          transition:background-color .2s ease;
        }
        :hover::-webkit-scrollbar-thumb{background-color:rgba(148,163,184,.32);background-clip:content-box;}
        ::-webkit-scrollbar-thumb:hover{background-color:rgba(148,163,184,.60);background-clip:content-box;}
        *{scrollbar-width:thin;scrollbar-color:transparent transparent;}
        *:hover{scrollbar-color:rgba(148,163,184,.32) transparent;}
      `}}/>

      {/* ── Header ── */}
      <header style={{
        display:"flex",alignItems:"center",justifyContent:"space-between",
        padding:"0 22px",flexShrink:0,
        height:52,
        background:"linear-gradient(180deg,#16212F 0%,#101925 100%)",
        borderBottom:"1px solid rgba(46,155,255,0.45)",
        boxShadow:"0 1px 0 rgba(46,155,255,0.15), 0 6px 18px -8px rgba(0,0,0,0.6)",
      }}>
        {/* Logo */}
        <img src={LOGO_AIRBOX} alt="Airbox" style={{height:34,objectFit:"contain",display:"block"}}/>

        {/* Nome da usina */}
        <div style={{display:"flex",flexDirection:"column",alignItems:"center",lineHeight:1.3}}>
          <span style={{fontSize:14,fontWeight:700,color:"#EAF2FB",letterSpacing:"0.08em",textTransform:"uppercase"}}>
            UFV — Serra do Mato
          </span>
          <span style={{fontSize:13,color:"#5E97D6",letterSpacing:"0.05em"}}>
            Usina Trairi · CE
          </span>
        </div>

        {/* Info direita */}
        <div style={{display:"flex",alignItems:"center",gap:16}}>
          <div style={{textAlign:"right"}}>
            <div style={{fontSize:13,color:"rgba(159,176,192,0.6)",letterSpacing:"0.03em"}}>Monitoramento</div>
            <div style={{fontSize:13,fontWeight:600,color:"rgba(230,237,245,0.9)"}}>Inversores Fotovoltaicos</div>
          </div>
          <div style={{width:1,height:28,background:"rgba(255,255,255,0.12)"}}></div>
          <div style={{display:"flex",flexDirection:"column",alignItems:"flex-end"}}>
            <span style={{fontSize:13,color:"rgba(159,176,192,0.55)"}}>inversores carregados</span>
            <span style={{fontSize:17,fontWeight:700,color:"#2E9BFF",textShadow:"0 0 12px rgba(46,155,255,0.45)"}}>{Object.keys(allData).length}</span>
          </div>
        </div>
      </header>

      {/* ── Corpo principal ── */}
      <div style={{display:"flex",flex:1,overflow:"hidden",minHeight:0}}>

      {/* ── Sidebar (só nas abas com upload de .xlsx) ── */}
      {showFileTab&&(<>
      <aside style={{width:sideW,flexShrink:0,display:"flex",flexDirection:"column",
        background:"var(--color-background-primary)",
        borderRight:"0.5px solid var(--color-border-tertiary)",
        overflow:"hidden",minWidth:180,maxWidth:420}}>

        {/* Inversores — flex:1, lista scrollável */}
        <div style={{flex:1,display:"flex",flexDirection:"column",minHeight:0,
          overflow:"hidden",borderBottom:"0.5px solid var(--color-border-tertiary)"}}>
          <div style={{padding:"8px 10px 5px",flexShrink:0,
            display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <span style={{fontWeight:600,fontSize:13,color:"var(--color-text-primary)"}}>Inversores</span>
            {hasData&&(
              <div style={{display:"flex",gap:5}}>
                <button onClick={toggleAll} style={{fontSize:13,padding:"2px 7px",cursor:"pointer",
                  borderRadius:5,border:"0.5px solid var(--color-border-secondary)",
                  background:"var(--color-background-secondary)"}}>
                  {visible.length>0?"Ocultar":"Mostrar"}
                </button>
                <button onClick={clearAll} style={{fontSize:13,padding:"2px 7px",cursor:"pointer",
                  borderRadius:5,border:"0.5px solid var(--color-border-secondary)",
                  background:"var(--color-background-secondary)",color:"var(--color-text-danger)"}}>
                  Limpar
                </button>
              </div>
            )}
          </div>
          <div style={{flex:1,overflowY:"auto",padding:"0 6px 6px"}}>
            {hasData?Object.values(allData).map(inv=>{
              const isHidden=hidden.has(inv.invKey);
              const nDays=Object.keys(inv.dates).length;
              return(
                <div key={inv.invKey} onClick={()=>toggleLegend(inv.invKey)}
                  title={isHidden?"Ativar":Object.keys(allData).filter(k=>!hidden.has(k)).length>1?"Isolar":"Mostrar todos"}
                  style={{display:"flex",alignItems:"center",gap:6,padding:"3px 5px",borderRadius:5,
                    cursor:"pointer",opacity:isHidden?0.3:1,
                    background:isHidden?"transparent":"rgba(0,0,0,0.025)",transition:"opacity 0.15s"}}>
                  <div style={{display:"flex",flexDirection:"column",gap:1.5,flexShrink:0}}>
                    {(activeVars.length>0?activeVars:["pac"]).slice(0,2).map((_,vi)=>(
                      <svg key={vi} width="18" height="6" style={{flexShrink:0}}>
                        <line x1="0" y1="3" x2="18" y2="3" stroke={varColor(inv.paletteIdx,vi)}
                          strokeWidth="2" strokeDasharray={DASH_PATTERNS[vi].join(",")}/>
                      </svg>
                    ))}
                  </div>
                  <span style={{flex:1,fontSize:13,fontWeight:500,overflow:"hidden",
                    textOverflow:"ellipsis",whiteSpace:"nowrap",
                    color:isHidden?"var(--color-text-tertiary)":"var(--color-text-primary)"}}
                    title={inv.displayName||inv.invKey}>{inv.displayName||inv.invKey}</span>
                  <span style={{fontSize:13,color:"var(--color-text-tertiary)",flexShrink:0}}>{nDays}d</span>
                  <button onClick={e=>{e.stopPropagation();toggleInv(inv.invKey);}}
                    style={{background:"none",border:"none",cursor:"pointer",fontSize:13,
                      color:"var(--color-text-tertiary)",padding:1,lineHeight:1,flexShrink:0}}>
                    <i className={`ti ti-eye${isHidden?"-off":""}`}/>
                  </button>
                </div>
              );
            }):(
              <div style={{padding:"14px 6px",fontSize:13,color:"var(--color-text-tertiary)",
                textAlign:"center",lineHeight:1.7}}>
                Arraste um .xlsx<br/>para começar
              </div>
            )}
          </div>
        </div>

        {/* ── Calendário (fixo) ── */}
        <div style={{flexShrink:0,borderBottom:"0.5px solid var(--color-border-tertiary)"}}>
          {!hasData?(
            <div style={{padding:"8px 12px",textAlign:"center",
              color:"var(--color-text-tertiary)",fontSize:13,lineHeight:1.5}}>
              Cada .xlsx = 1 inversor<br/>cada planilha = 1 dia
            </div>
          ):(
            <MiniCalendar datesSet={datesSet} selectedDate={selectedDate}
              onSelect={d=>{setSelectedDate(d);setSelectedTime(null);}}
              invColors={calDotColors}/>
          )}
        </div>

        {/* ── Upload (fixo no fundo) ── */}
        <div style={{flexShrink:0}}>
          {/* Inversores */}
          <div onDragOver={e=>{e.preventDefault();setIsDragging(true);}}
            onDragLeave={()=>setIsDragging(false)} onDrop={onDrop}
            onClick={()=>fileRef.current?.click()}
            style={{margin:"6px 8px 3px",padding:"7px 8px",textAlign:"center",cursor:"pointer",
              border:`1.5px dashed ${isDragging?"#2E9BFF":"var(--color-border-secondary)"}`,
              borderRadius:6,
              background:isDragging?"rgba(46,155,255,0.06)":"var(--color-background-secondary)",
              transition:"all 0.15s"}}>
            <div style={{display:"flex",alignItems:"center",justifyContent:"center",gap:6}}>
              <i className="ti ti-upload" style={{fontSize:14,color:"#2E9BFF"}}/>
              <span style={{fontSize:13,color:"var(--color-text-secondary)"}}>
                {loading?"⏳ Carregando…":"Inversores .xlsx"}
              </span>
            </div>
            <input ref={fileRef} type="file" accept=".xlsx,.xls" multiple
              style={{display:"none"}} onChange={e=>addFiles(e.target.files)}/>
          </div>
          {errors.length>0&&(
            <div style={{margin:"0 8px 5px",padding:"5px 8px",borderRadius:5,fontSize:13,
              background:"var(--color-background-danger)",color:"var(--color-text-danger)",lineHeight:1.5}}>
              {errors.map((e,i)=><div key={i}>⚠ {e}</div>)}
            </div>
          )}
        </div>
      </aside>

            {/* Divisor lateral */}
      <div onMouseDown={onSideDrag}
        style={{width:5,cursor:"col-resize",flexShrink:0,background:"transparent",
          borderRight:"1px solid var(--color-border-tertiary)",
          display:"flex",alignItems:"center",justifyContent:"center"}}
        onMouseEnter={e=>e.currentTarget.style.background="rgba(46,155,255,0.18)"}
        onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
        <div style={{width:2,height:40,borderRadius:2,background:"rgba(148,163,184,0.22)",pointerEvents:"none"}}></div>
      </div>
      </>)}

      {/* ── Área principal ── */}
      <main style={{flex:1,display:"flex",flexDirection:"column",overflow:"hidden",minWidth:0,minHeight:0}}>

        {/* ── Barra de navegação + abas ── */}
        <div style={{borderBottom:"1px solid rgba(255,255,255,0.06)",
          background:"#131C28",flexShrink:0}}>

          {/* Linha 1: navegação + abas */}
          <div style={{display:"flex",alignItems:"center",padding:"0 14px",
            borderBottom:"0.5px solid var(--color-border-tertiary)",minHeight:38}}>

            {/* Setas de dia (só nas abas com upload de .xlsx) */}
            {showFileTab&&(
            <div style={{display:"flex",alignItems:"center",gap:2,marginRight:14,flexShrink:0}}>
              <button onClick={prevDate} disabled={curIdx<=0}
                style={{background:"none",border:"none",cursor:curIdx>0?"pointer":"default",fontSize:15,
                  color:curIdx>0?"var(--color-text-secondary)":"var(--color-text-tertiary)",
                  padding:"3px 5px",borderRadius:4,lineHeight:1,transition:"background 0.1s"}}
                onMouseEnter={e=>{if(curIdx>0)e.currentTarget.style.background="var(--color-background-secondary)";}}
                onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
                <i className="ti ti-chevron-left"/>
              </button>
              <div style={{minWidth:118,textAlign:"center",padding:"3px 8px",borderRadius:5,
                background:"var(--color-background-secondary)",fontSize:13,fontWeight:600,
                color:selectedDate?"var(--color-text-primary)":"var(--color-text-tertiary)"}}>
                {selectedDate?(()=>{
                  const[y,m,d]=selectedDate.split("-");
                  return`${d}/${m}/${y}`;
                })():"Selecione um dia"}
              </div>
              <button onClick={nextDate} disabled={curIdx>=sortedDates.length-1}
                style={{background:"none",border:"none",cursor:curIdx<sortedDates.length-1?"pointer":"default",fontSize:15,
                  color:curIdx<sortedDates.length-1?"var(--color-text-secondary)":"var(--color-text-tertiary)",
                  padding:"3px 5px",borderRadius:4,lineHeight:1,transition:"background 0.1s"}}
                onMouseEnter={e=>{if(curIdx<sortedDates.length-1)e.currentTarget.style.background="var(--color-background-secondary)";}}
                onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
                <i className="ti ti-chevron-right"/>
              </button>
              {sortedDates.length>0&&(
                <span style={{fontSize:13,color:"var(--color-text-tertiary)",marginLeft:3}}>
                  {curIdx>=0?curIdx+1:0}/{sortedDates.length}
                </span>
              )}
            </div>
            )}

            {/* Abas */}
            {[
              {id:"vars",      icon:"ti-chart-line",        label:"Variáveis"},
              {id:"imbalance", icon:"ti-arrows-left-right", label:"Desbalanceamento"},
              {id:"avail",     icon:"ti-chart-bar",         label:"Disponibilidade"},
              {id:"gen",       icon:"ti-bolt",              label:"Geração"},
              {id:"combiners", icon:"ti-plug-connected",    label:"Combiners"},
            ].map(tab=>(
              <button key={tab.id} onClick={()=>setActiveTab(tab.id)}
                style={{padding:"7px 13px",fontSize:13,cursor:"pointer",border:"none",
                  borderBottom:activeTab===tab.id?"2.5px solid #2E9BFF":"2.5px solid transparent",
                  background:"transparent",
                  color:activeTab===tab.id?"#2E9BFF":"var(--color-text-secondary)",
                  fontWeight:activeTab===tab.id?600:400,
                  filter:activeTab===tab.id?"drop-shadow(0 1px 6px rgba(46,155,255,0.5))":"none",
                  display:"flex",alignItems:"center",gap:5,borderRadius:0,flexShrink:0,transition:"all 0.15s"}}>
                <i className={`ti ${tab.icon}`} style={{fontSize:13}}/>{tab.label}
              </button>
            ))}

            <div style={{marginLeft:"auto",fontSize:13,color:"var(--color-text-tertiary)",flexShrink:0,display:"flex",alignItems:"center",gap:10}}>
              <span>{visible.length} visível(is)</span>
              {selectedTime&&(
                <span style={{color:"#2E9BFF",display:"inline-flex",alignItems:"center",gap:4}}>
                  <i className="ti ti-map-pin" style={{fontSize:13}}/><strong>{selectedTime}</strong>
                  <button onClick={()=>setSelectedTime(null)}
                    style={{background:"none",border:"none",cursor:"pointer",fontSize:13,color:"#2E9BFF",padding:0,lineHeight:1,marginLeft:1}}>
                    <i className="ti ti-x"/>
                  </button>
                </span>
              )}
            </div>
          </div>

          {/* Linha 2: seletores de variável (só na aba vars) */}
          {activeTab==="vars"&&(
            <div style={{padding:"6px 14px",display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
              {[0,1,2,3].map(vi=>(
                <VarSelect key={vi} id={`var${vi+1}sel`} value={vars[vi]}
                  onChange={key=>setVar(vi,key)} label={`V${vi+1}:`}
                  swatchColor={selectorColors[vi]} dashPattern={DASH_PATTERNS[vi]} optional={vi>0}/>
              ))}
              {dualAxis&&<span style={{fontSize:13,padding:"3px 8px",borderRadius:10,
                background:"rgba(46,155,255,0.1)",color:"#1B6FC9",flexShrink:0}}>eixo duplo</span>}
            </div>
          )}
        </div>

        {/* ── Área do gráfico / disponibilidade ── */}
        <div style={{flex:1,margin:"12px 14px 8px",
          padding:!showFileTab?"0":"14px 16px 8px",minHeight:0,overflow:"hidden",
          display:"flex",flexDirection:"column",
          background:"#131C28",borderRadius:14,border:"1px solid rgba(255,255,255,0.06)",
          boxShadow:"0 12px 32px -16px rgba(0,0,0,0.75)"}}>
          {activeTab==="combiners"?(
            <CombinerPanel/>
          ):activeTab==="avail"?(
            <AvailabilityPanel/>
          ):activeTab==="gen"?(
            <GenerationPanel/>
          ):!selectedDate||inverters.length===0?(
            <div style={{flex:1,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",
              color:"var(--color-text-tertiary)",gap:14,textAlign:"center"}}>
              <i className="ti ti-chart-line" style={{fontSize:56}}/>
              <div style={{fontSize:15,fontWeight:500}}>
                {!hasData?"Carregue um arquivo .xlsx":"Selecione um dia no calendário"}
              </div>
              <div style={{fontSize:13,maxWidth:380,lineHeight:1.9,color:"var(--color-text-tertiary)"}}>
                {!hasData?<>Arraste um <strong>.xlsx</strong> na sidebar.<br/>Cada planilha dentro = um dia diferente.<br/>Vários arquivos = vários inversores.</>
                  :<>Use as setas ← → ou clique no calendário.<br/>Todos os inversores com dados naquele dia são sobrepostos.</>}
              </div>
            </div>
          ):activeTab==="vars"?(
            <InverterChart inverters={inverters} varKeys={vars} onPointClick={setSelectedTime}/>
          ):(
            <ImbalanceChart inverters={inverters} onPointClick={setSelectedTime}/>
          )}
        </div>

        {/* ── Divisor horizontal ── */}
        {showBottomPanel&&showStats&&(
          <div onMouseDown={onStatsDrag}
            style={{height:5,cursor:"row-resize",flexShrink:0,background:"transparent",
              borderTop:"1px solid var(--color-border-tertiary)",
              display:"flex",alignItems:"center",justifyContent:"center"}}
            onMouseEnter={e=>e.currentTarget.style.background="rgba(46,155,255,0.18)"}
            onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
            <div style={{height:2,width:40,borderRadius:2,background:"rgba(148,163,184,0.22)",pointerEvents:"none"}}></div>
          </div>
        )}

        {/* ── Painel inferior (oculto na aba disponibilidade) ── */}
        {showBottomPanel&&(
          <div style={{height:showStats?statsH:"auto",minHeight:showStats?80:"auto",
            borderTop:showStats?"none":"0.5px solid var(--color-border-tertiary)",
            display:"flex",flexDirection:"column",flexShrink:0,overflow:"hidden"}}>

            <div style={{padding:"5px 14px 0",display:"flex",alignItems:"center",justifyContent:"space-between",
              borderBottom:"0.5px solid var(--color-border-tertiary)",flexShrink:0}}>
              <div style={{display:"flex",alignItems:"center",gap:8}}>
                <span style={{fontSize:13,fontWeight:600,color:"var(--color-text-secondary)"}}>Análise</span>
                {inverters.some(i=>!i.visible)&&(
                  <span style={{fontSize:13,padding:"1px 6px",borderRadius:10,
                    background:"rgba(148,163,184,0.10)",color:"var(--color-text-tertiary)"}}>
                    {inverters.filter(i=>!i.visible).length} oculto(s) — dados mantidos
                  </span>
                )}
              </div>
              <button onClick={()=>setShowStats(s=>!s)}
                style={{background:"rgba(255,255,255,0.06)",border:"1px solid rgba(255,255,255,0.08)",cursor:"pointer",
                  width:22,height:22,borderRadius:6,display:"flex",alignItems:"center",justifyContent:"center",
                  color:"var(--color-text-secondary)",fontSize:13,lineHeight:1,flexShrink:0}}>
                <i className={`ti ti-chevron-${showStats?"down":"up"}`}/>
              </button>
            </div>

            {showStats&&(
              <div style={{flex:1,display:"flex",gap:10,padding:"10px 14px 12px",overflow:"hidden",minHeight:0}}>

                {/* Col 1 – Estatísticas */}
                <div style={{flex:"0 0 38%",minWidth:220,
                  display:"flex",flexDirection:"column",overflow:"hidden",
                  background:"#161F2C",borderRadius:12,border:"1px solid rgba(255,255,255,0.06)"}}>
                  <div style={{padding:"8px 13px 4px",fontSize:13,fontWeight:600,color:"var(--color-text-tertiary)",
                    textTransform:"uppercase",letterSpacing:"0.05em",flexShrink:0}}>
                    Estatísticas — {activeVars.map(k=>VAR_MAP[k]?.label).join(" · ")}
                  </div>
                  <StatsTable stats={stats} activeVars={activeVars} fmt={fmt}/>
                </div>

                {/* Col 2 – Snapshot */}
                <div style={{flex:"0 0 27%",minWidth:160,
                  display:"flex",flexDirection:"column",overflow:"hidden",
                  background:"#161F2C",borderRadius:12,border:"1px solid rgba(255,255,255,0.06)"}}>
                  <div style={{padding:"8px 13px 4px",fontSize:13,fontWeight:600,color:"var(--color-text-tertiary)",
                    textTransform:"uppercase",letterSpacing:"0.05em",flexShrink:0,display:"flex",alignItems:"center",gap:5}}>
                    <i className="ti ti-map-pin" style={{fontSize:13,color:"#2E9BFF"}}/>
                    Ponto {selectedTime&&<span style={{color:"#2E9BFF",fontWeight:700,marginLeft:3}}>{selectedTime}</span>}
                  </div>
                  {!selectedTime?(
                    <div style={{flex:1,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",
                      color:"var(--color-text-tertiary)",fontSize:13,textAlign:"center",padding:12}}>
                      <i className="ti ti-hand-click" style={{fontSize:24,marginBottom:5}}/>
                      Clique no gráfico para inspecionar
                    </div>
                  ):(
                    <SnapshotTable snapshot={snapshot} activeVars={activeVars}/>
                  )}
                </div>

                {/* Col 3 – Desbalanceamento */}
                <div style={{flex:1,minWidth:160,display:"flex",flexDirection:"column",overflow:"hidden",
                  background:"#161F2C",borderRadius:12,border:"1px solid rgba(255,255,255,0.06)"}}>
                  <div style={{padding:"8px 13px 3px",fontSize:13,fontWeight:600,color:"var(--color-text-tertiary)",
                    textTransform:"uppercase",letterSpacing:"0.05em",flexShrink:0,display:"flex",alignItems:"center",gap:5}}>
                    <i className="ti ti-arrows-left-right" style={{fontSize:13,color:"#2E9BFF"}}/>Desbalanceamento pvbus/nvbus
                  </div>
                  {imbalanceData.every(d=>d.min===null)?(
                    <div style={{flex:1,display:"flex",alignItems:"center",justifyContent:"center",
                      color:"var(--color-text-tertiary)",fontSize:13,textAlign:"center",padding:10}}>
                      <div><i className="ti ti-alert-circle" style={{fontSize:20,display:"block",marginBottom:4}}/>Sem pvbus/nvbus</div>
                    </div>
                  ):(
                    <ImbalanceTable imbalanceData={imbalanceData} selectedTime={selectedTime} alertC={alertC} fmtPct={fmtPct}/>
                  )}
                </div>
              </div>
            )}
          </div>
        )}
      </main>
    </div>
    </div>
  );
}

// ── Tabela de Estatísticas (sortable) ────────────────────────────────────────
function StatsTable({ stats, activeVars, fmt }) {
  const { sortKey, sortDir, toggle } = useSortableTable("name","asc");
  const flat = useMemo(() => {
    const rows = stats.flatMap(s =>
      s.varStats.map((vs, vi) => ({
        name: s.displayName||s.name,
        varLabel: VAR_MAP[vs.key]?.label??vs.key,
        min: vs.min, max: vs.max, avg: vs.avg,
        s, vs, vi, isHid: !s.visible,
      }))
    );
    return applySortDir(rows, sortKey, sortDir);
  }, [stats, sortKey, sortDir, activeVars.join(",")]);

  return (
    <div style={{overflowY:"auto",overflowX:"hidden",flex:1,padding:"0 12px 8px"}}>
      <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
        <thead style={{position:"sticky",top:0,zIndex:1,background:"#161F2C"}}>
          <tr>
            <SortTh label="Inversor" sortKey={sortKey} col="name"     sortDir={sortDir} onSort={toggle} align="left"/>
            <SortTh label="Variável" sortKey={sortKey} col="varLabel" sortDir={sortDir} onSort={toggle} align="left"/>
            <SortTh label="Mín"      sortKey={sortKey} col="min"      sortDir={sortDir} onSort={toggle}/>
            <SortTh label="Máx"      sortKey={sortKey} col="max"      sortDir={sortDir} onSort={toggle}/>
            <SortTh label="Média"    sortKey={sortKey} col="avg"      sortDir={sortDir} onSort={toggle}/>
          </tr>
        </thead>
        <tbody>
          {flat.map((row, idx) => {
            const { s, vs, vi, isHid } = row;
            const info=VAR_MAP[vs.key], color=varColor(s.paletteIdx,vi), dash=DASH_PATTERNS[vi];
            return (
              <tr key={`${s.id}-${vs.key}-${idx}`}
                style={{background:idx%2===1?"var(--color-background-secondary)":"transparent",opacity:isHid?0.45:1}}>
                <td style={{padding:"3px 6px",whiteSpace:"nowrap"}}>
                  <div style={{display:"flex",alignItems:"center",gap:6}}>
                    <span style={{width:8,height:8,borderRadius:"50%",background:s.color,flexShrink:0,
                      outline:isHid?`2px dashed ${s.color}`:`2px solid ${s.color}44`,outlineOffset:1}}></span>
                    <span style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",maxWidth:110,
                      fontSize:11,fontWeight:500,color:isHid?"var(--color-text-tertiary)":"var(--color-text-primary)"}}
                      title={s.displayName||s.name}>
                      {s.displayName||s.name}
                      {isHid&&<span style={{fontSize:11,marginLeft:3,color:"var(--color-text-tertiary)"}}>(oculto)</span>}
                    </span>
                  </div>
                </td>
                <td style={{padding:"3px 6px",whiteSpace:"nowrap"}}>
                  <div style={{display:"flex",alignItems:"center",gap:5}}>
                    <svg width="16" height="8" style={{flexShrink:0}}>
                      <line x1="0" y1="4" x2="16" y2="4" stroke={color} strokeWidth="2.5" strokeDasharray={dash.join(",")}/>
                    </svg>
                    <span style={{fontSize:11,color,whiteSpace:"nowrap"}}>{info?.label??vs.key}</span>
                  </div>
                </td>
                {[vs.min,vs.max,vs.avg].map((v,j)=>(
                  <td key={j} style={{padding:"3px 6px",textAlign:"right",fontFamily:"var(--font-mono)",
                    fontSize:11,color:isHid?"var(--color-text-tertiary)":"var(--color-text-primary)",whiteSpace:"nowrap"}}>
                    {fmt(v)}{v!==null&&<span style={{color:"var(--color-text-tertiary)",marginLeft:2,fontSize:10}}>{info?.unit}</span>}
                  </td>
                ))}
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

// ── Tabela de Snapshot (sortable) ────────────────────────────────────────────
function SnapshotTable({ snapshot, activeVars }) {
  const { sortKey, sortDir, toggle } = useSortableTable("name","asc");
  const rows = useMemo(() => {
    if (!snapshot) return [];
    const base = snapshot.map(s => {
      const obj = { name: s.displayName||s.name, s };
      activeVars.forEach((k,vi) => { obj[k] = s.values[vi]?.value ?? null; });
      return obj;
    });
    return applySortDir(base, sortKey, sortDir);
  }, [snapshot, sortKey, sortDir, activeVars.join(",")]);

  return (
    <div style={{overflowY:"auto",overflowX:"hidden",flex:1,padding:"0 12px 8px"}}>
      <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
        <thead style={{position:"sticky",top:0,zIndex:1,background:"#161F2C"}}>
          <tr>
            <SortTh label="Inversor" sortKey={sortKey} col="name" sortDir={sortDir} onSort={toggle} align="left"/>
            {activeVars.map((k)=>(
              <SortTh key={k} label={VAR_MAP[k]?.label??k} sortKey={sortKey} col={k} sortDir={sortDir} onSort={toggle}/>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((row,si)=>{
            const s=row.s;
            return (
              <tr key={s.id} style={{background:si%2===1?"var(--color-background-secondary)":"transparent",opacity:s.visible?1:0.45}}>
                <td style={{padding:"3px 6px"}}>
                  <div style={{display:"flex",alignItems:"center",gap:6}}>
                    <span style={{width:8,height:8,borderRadius:"50%",background:s.color,flexShrink:0}}></span>
                    <span style={{fontSize:11,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",maxWidth:90,
                      color:s.visible?"var(--color-text-primary)":"var(--color-text-tertiary)"}}
                      title={s.displayName||s.name}>{s.displayName||s.name}</span>
                  </div>
                </td>
                {activeVars.map((k,vi)=>{
                  const v=s.values[vi]?.value;
                  return (
                    <td key={k} style={{padding:"3px 6px",textAlign:"right",fontFamily:"var(--font-mono)",
                      fontSize:11,whiteSpace:"nowrap",
                      color:v!==null&&v!==undefined?varColor(s.paletteIdx,vi):"var(--color-text-tertiary)"}}>
                      {v!==null&&v!==undefined?<>{v.toFixed(3)}<span style={{color:"var(--color-text-tertiary)",marginLeft:3,fontSize:11}}>{VAR_MAP[k]?.unit}</span></>:"—"}
                    </td>
                  );
                })}
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

// ── Tabela de Desbalanceamento (sortable) ────────────────────────────────────
function ImbalanceTable({ imbalanceData, selectedTime, alertC, fmtPct }) {
  const { sortKey, sortDir, toggle } = useSortableTable("name","asc");
  const rows = useMemo(() => {
    const base = imbalanceData.map(d => ({
      name: d.displayName||d.name, atPoint: d.atPoint,
      min: d.min, max: d.max, avg: d.avg, d,
    }));
    return applySortDir(base, sortKey, sortDir);
  }, [imbalanceData, sortKey, sortDir]);

  return (
    <div style={{overflowY:"auto",overflowX:"hidden",flex:1,padding:"0 12px 8px"}}>
      <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
        <thead style={{position:"sticky",top:0,zIndex:1,background:"#161F2C"}}>
          <tr>
            <SortTh label="Inversor" sortKey={sortKey} col="name"    sortDir={sortDir} onSort={toggle} align="left"/>
            <SortTh label={selectedTime?`Às ${selectedTime}`:"—"} sortKey={sortKey} col="atPoint" sortDir={sortDir} onSort={toggle}/>
            <SortTh label="Mín"      sortKey={sortKey} col="min"     sortDir={sortDir} onSort={toggle}/>
            <SortTh label="Máx"      sortKey={sortKey} col="max"     sortDir={sortDir} onSort={toggle}/>
            <SortTh label="Média"    sortKey={sortKey} col="avg"     sortDir={sortDir} onSort={toggle}/>
          </tr>
        </thead>
        <tbody>
          {rows.map((row,i)=>{
            const d=row.d;
            return (
              <tr key={d.id} style={{background:i%2===1?"var(--color-background-secondary)":"transparent",opacity:d.visible?1:0.45}}>
                <td style={{padding:"3px 6px"}}>
                  <div style={{display:"flex",alignItems:"center",gap:6}}>
                    <span style={{width:8,height:8,borderRadius:"50%",background:d.color,flexShrink:0}}></span>
                    <span style={{fontSize:11,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",maxWidth:80,
                      color:d.visible?"var(--color-text-primary)":"var(--color-text-tertiary)"}}
                      title={d.displayName||d.name}>{d.displayName||d.name}</span>
                  </div>
                </td>
                <td style={{padding:"3px 6px",textAlign:"right",fontFamily:"var(--font-mono)",fontSize:11,
                  fontWeight:600,color:selectedTime?alertC(d.atPoint):"var(--color-text-tertiary)",whiteSpace:"nowrap"}}>
                  {selectedTime?fmtPct(d.atPoint):"—"}
                </td>
                {[d.min,d.max,d.avg].map((v,j)=>(
                  <td key={j} style={{padding:"3px 6px",textAlign:"right",fontFamily:"var(--font-mono)",
                    fontSize:11,color:alertC(v),whiteSpace:"nowrap"}}>{fmtPct(v)}</td>
                ))}
              </tr>
            );
          })}
        </tbody>
      </table>
      <div style={{marginTop:5,display:"flex",gap:12,flexWrap:"wrap",fontSize:11,
        color:"var(--color-text-tertiary)",paddingLeft:2}}>
        <span><span style={{color:"#4CAF50"}}>●</span> &lt;35% normal</span>
        <span><span style={{color:"#FF9800"}}>●</span> 35–70% atenção</span>
        <span><span style={{color:"#F44336"}}>●</span> &gt;70% crítico</span>
      </div>
    </div>
  );
}

// ── Painel de Geração ─────────────────────────────────────────────────────────
const GEN_PERIOD_LABELS = { daily:"Diário", weekly:"7 dias", monthly:"Mensal" };

function GenerationPanel() {
  const invMap = useInverterMap();
  const dates = useMemo(()=>last30Dates(),[]);
  const today = dates[dates.length-1];
  const [period, setPeriod] = useState("daily");
  const [unit, setUnit]     = useState("MWh");
  const [selDate, setSelDate] = useState(today);
  const [rows, setRows] = useState([]); // [{sn,date,eInjection,eAbsorption}]
  const [loading, setLoading] = useState(true);
  const [fatalError, setFatalError] = useState(null);
  const uF  = unit==="kWh" ? 1000 : 1;                 // fator de conversão (gen é calculada em MWh)
  const uDec = unit==="kWh" ? 0 : 3;                   // casas decimais p/ total/média
  const uDecR = unit==="kWh" ? 0 : 1;                  // casas decimais p/ rankings
  const fmtU = (mwh,dec)=> (mwh*uF).toLocaleString("pt-BR",{minimumFractionDigits:dec,maximumFractionDigits:dec});
  const [chartReady, setChartReady] = useState(!!window._chartReady);
  const canvasRef = useRef(null);
  const chartRef  = useRef(null);

  useEffect(()=>{ if(window._chartReady){setChartReady(true);return;} loadChartLibs().then(()=>setChartReady(true)); },[]);

  // Uma única chamada cobre os últimos 30 dias — groupbyday já devolve o range inteiro
  useEffect(()=>{
    let cancelled = false;
    (async()=>{
      setLoading(true);
      try {
        const res = await fetch(`/api/generation?from=${dates[0]}&to=${dates[dates.length-1]}`);
        if (!res.ok) {
          const body = await res.json().catch(()=>({}));
          throw new Error(body.error || `Erro ${res.status}`);
        }
        const data = await res.json();
        if (!cancelled) setRows(data);
      } catch(err) {
        if (!cancelled) setFatalError(err.message);
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();
    return ()=>{ cancelled=true; };
  },[dates]);

  const dateIdx = dates.indexOf(selDate);
  const goPrevDay = () => { if(dateIdx>0) setSelDate(dates[dateIdx-1]); };
  const goNextDay = () => { if(dateIdx<dates.length-1) setSelDate(dates[dateIdx+1]); };

  // Agrega geração (EInjection, kWh→MWh) por inversor e período, a partir do cache já baixado
  const genData = useMemo(()=>{
    if(!rows.length) return [];
    let periodDates;
    if(period==="daily"){
      periodDates=new Set([selDate]);
    } else if(period==="weekly"){
      periodDates=new Set(dates.slice(Math.max(0,dateIdx-6),dateIdx+1));
    } else {
      const ym=selDate.slice(0,6);
      periodDates=new Set(dates.filter(d=>d.startsWith(ym)));
    }
    const bySn={};
    rows.forEach(r=>{
      if(!periodDates.has(r.date.replaceAll("-",""))) return;
      (bySn[r.sn] ??= []).push(r);
    });
    return Object.entries(bySn)
      .map(([sn,recs])=>{
        const pos = invMap[sn]; // ex.: "3.4.1"
        return {
          invKey:sn, name:invName(sn,invMap),
          gen: recs.reduce((s,r)=>s+(r.eInjection||0),0)/1000, // kWh → MWh
          group: pos ? pos.slice(-1) : null, // "1" (17 combiners) ou "2" (16 combiners)
        };
      })
      .sort((a,b)=>a.name.localeCompare(b.name,undefined,{numeric:true}))
      .map((d,i)=>({...d, color:invColor(i)}));
  },[rows,period,selDate,dateIdx,dates,invMap]);

  // Agrupa por posição: termina em 1 → inversor com 17 combiners; termina em 2 → 16 combiners
  const groups = useMemo(()=>{
    const g1=genData.filter(d=>d.group==="1").sort((a,b)=>b.gen-a.gen);
    const g2=genData.filter(d=>d.group==="2").sort((a,b)=>b.gen-a.gen);
    const other=genData.filter(d=>d.group!=="1"&&d.group!=="2").sort((a,b)=>b.gen-a.gen);
    return {g1,g2,other};
  },[genData]);

  // Δ% entre maior e menor geração de um grupo
  const spreadPct = arr => {
    if(arr.length<2) return null;
    const max=arr[0].gen, min=arr[arr.length-1].gen;
    return min>0 ? (max-min)/min*100 : null;
  };

  // Render bar chart via Chart.js
  useEffect(()=>{
    if(!chartReady||!canvasRef.current||!genData.length) return;
    if(chartRef.current){chartRef.current.destroy();chartRef.current=null;}
    // Ordem: grupo 1 (ranqueado) · espaço · grupo 2 (ranqueado) · outros
    const ordered=[];
    groups.g1.forEach(d=>ordered.push(d));
    if(groups.g1.length&&(groups.g2.length||groups.other.length)) ordered.push({spacer:true,name:"",color:"#00000000"});
    groups.g2.forEach(d=>ordered.push(d));
    if(groups.g2.length&&groups.other.length) ordered.push({spacer:true,name:" ",color:"#00000000"});
    groups.other.forEach(d=>ordered.push(d));
    const labels=ordered.map(d=>d.name);
    const values=ordered.map(d=>d.spacer?null:d.gen*uF);
    const colors=ordered.map(d=>d.color);
    chartRef.current=new window.Chart(canvasRef.current,{
      type:"bar",
      data:{
        labels,
        datasets:[{
          label:"Geração (MWh)",
          data:values,
          backgroundColor:colors.map(c=>c+"cc"),
          borderColor:colors,
          borderWidth:1.5,
          borderRadius:4,
        }]
      },
      options:{
        indexAxis:"y",  // barras horizontais — mais legível com muitos inversores
        responsive:true,maintainAspectRatio:false,animation:false,
        plugins:{
          legend:{display:false},
          tooltip:{
            backgroundColor:"rgba(38,50,68,0.97)",titleColor:"#A7B6C6",bodyColor:"#EAF2FB",
            borderColor:"#e0e0e0",borderWidth:1,padding:10,
            callbacks:{
              label:ctx=>`  ${ctx.parsed.x.toLocaleString("pt-BR",{maximumFractionDigits:uDecR})} ${unit}`,
            }
          }
        },
        scales:{
          x:{
            ticks:{color:"#8595A6",font:{size:11}},
            grid:{color:"rgba(148,163,184,0.10)"},border:{color:"rgba(148,163,184,0.22)"},
            title:{display:true,text:unit,color:"#A7B6C6",font:{size:11}},
          },
          y:{
            ticks:{color:"#8595A6",font:{size:11}},
            grid:{color:"rgba(0,0,0,0.04)"},border:{color:"rgba(0,0,0,0.12)"},
          },
        },
      },
    });
    return()=>{if(chartRef.current){chartRef.current.destroy();chartRef.current=null;}};
  },[chartReady,groups,unit]);

  const total   = genData.reduce((s,d)=>s+d.gen,0);
  const avg     = genData.length>0 ? total/genData.length : 0;

  return(
    <div style={{flex:1,display:"flex",flexDirection:"column",overflow:"hidden",padding:"10px 16px 8px",minHeight:0}}>
      {/* Header */}
      <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:10,flexShrink:0,flexWrap:"wrap"}}>
        {/* Toggle período */}
        <div style={{display:"flex",gap:2,background:"var(--color-background-secondary)",padding:3,borderRadius:8,flexShrink:0}}>
          {["daily","weekly","monthly"].map(p=>(
            <button key={p}
              onClick={()=>setPeriod(p)}
              onDoubleClick={()=>{
                setPeriod(p);
                if(p==="daily"||p==="weekly") setSelDate(today);
              }}
              title={p==="daily"?"Clique: período diário · 2× clique: ir ao dia mais recente"
                    :p==="weekly"?"Clique: 7 dias · 2× clique: 7 dias até o mais recente"
                    :GEN_PERIOD_LABELS[p]}
              style={{padding:"5px 14px",fontSize:13,cursor:"pointer",borderRadius:6,border:"none",
                background:period===p?"#2E9BFF":"transparent",
                color:period===p?"#fff":"var(--color-text-secondary)",
                fontWeight:period===p?600:400,transition:"all 0.15s"}}>
              {GEN_PERIOD_LABELS[p]}
            </button>
          ))}
        </div>
        {/* Toggle unidade */}
        <div style={{display:"flex",gap:2,background:"var(--color-background-secondary)",padding:3,borderRadius:8,flexShrink:0}}>
          {["MWh","kWh"].map(u=>(
            <button key={u} onClick={()=>setUnit(u)}
              title={`Exibir geração em ${u}`}
              style={{padding:"5px 12px",fontSize:13,cursor:"pointer",borderRadius:6,border:"none",
                background:unit===u?"#2E9BFF":"transparent",
                color:unit===u?"#fff":"var(--color-text-secondary)",
                fontWeight:unit===u?600:400,transition:"all 0.15s"}}>
              {u}
            </button>
          ))}
        </div>
        {(period==="daily"||period==="weekly")&&(
          <div style={{display:"flex",alignItems:"center",gap:4}}>
            <button onClick={goPrevDay} disabled={dateIdx<=0}
              style={{background:"none",border:"none",cursor:dateIdx>0?"pointer":"default",fontSize:15,
                color:dateIdx>0?"var(--color-text-secondary)":"var(--color-text-tertiary)",padding:"2px 4px"}}>
              <i className="ti ti-chevron-left"/>
            </button>
            <span style={{fontSize:13,fontWeight:600,minWidth:88,textAlign:"center",fontFamily:"var(--font-mono)"}}>
              {fmtYmd(selDate)}
            </span>
            <button onClick={goNextDay} disabled={dateIdx>=dates.length-1}
              style={{background:"none",border:"none",cursor:dateIdx<dates.length-1?"pointer":"default",fontSize:15,
                color:dateIdx<dates.length-1?"var(--color-text-secondary)":"var(--color-text-tertiary)",padding:"2px 4px"}}>
              <i className="ti ti-chevron-right"/>
            </button>
          </div>
        )}
        <div style={{fontSize:13,color:"var(--color-text-secondary)"}}>
          <strong>{genData.length}</strong> inversores
          · Total: <strong style={{color:"#2E9BFF"}}>{fmtU(total||0,uDec)} {unit}</strong>
          · Média: <strong>{fmtU(avg||0,uDec)} {unit}</strong>
        </div>
      </div>

      {/* Corpo: gráfico + rankings */}
      <div style={{flex:1,display:"flex",gap:0,overflow:"hidden",minHeight:0}}>

        {/* ── Gráfico de barras horizontais ── */}
        <div style={{flex:1,minWidth:0,paddingRight:12,display:"flex",flexDirection:"column",overflow:"hidden"}}>
          <div style={{flex:1,minHeight:0,position:"relative"}}>
            {fatalError?(
              <div style={{height:"100%",display:"flex",alignItems:"center",justifyContent:"center",
                color:"var(--color-text-tertiary)",fontSize:13,textAlign:"center"}}>
                <div><i className="ti ti-plug-connected-x" style={{fontSize:40,display:"block",marginBottom:8}}/>{fatalError}</div>
              </div>
            ):!chartReady||loading||!genData.length?(
              <div style={{height:"100%",display:"flex",alignItems:"center",justifyContent:"center",
                color:"var(--color-text-tertiary)",fontSize:13}}>
                {loading?"Carregando…":!genData.length?"Sem dados para o período":"Carregando…"}
              </div>
            ):(
              <canvas ref={canvasRef} role="img"/>
            )}
          </div>
        </div>

        {/* ── Rankings por terminação (Final 1 / Final 2) ── */}
        <div style={{width:232,flexShrink:0,display:"flex",flexDirection:"column",gap:10,
          paddingLeft:12,borderLeft:"0.5px solid var(--color-border-tertiary)",overflowY:"auto"}}>

          {[{key:"g1",label:"17 combiners (·1)",arr:groups.g1},{key:"g2",label:"16 combiners (·2)",arr:groups.g2}].map(({key,label,arr})=>{
            if(!arr.length) return null;
            const gmax=arr[0]?.gen||0;
            const gavg=arr.reduce((s,x)=>s+x.gen,0)/arr.length;
            return(
              <div key={key} style={{flexShrink:0}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:4,gap:6}}>
                  <span style={{fontSize:11,fontWeight:600,color:"var(--color-text-tertiary)",
                    textTransform:"uppercase",letterSpacing:"0.04em"}}>{label}</span>
                  <span title="Média de geração do grupo"
                    style={{fontSize:12,fontFamily:"var(--font-mono)",fontWeight:700,
                      color:"#A7B6C6",background:"var(--color-background-secondary)",
                      padding:"2px 8px",borderRadius:5,flexShrink:0}}>
                    Média {fmtU(gavg,uDecR)} {unit}
                  </span>
                </div>
                {arr.map((d,i)=>{
                  const pct=gmax>0?d.gen/gmax*100:0;
                  const dp=gavg>0?(d.gen-gavg)/gavg*100:0;
                  const dpColor=dp>=0?"#4CAF50":"#F44336";
                  return(
                    <div key={d.invKey} style={{position:"relative",display:"flex",alignItems:"center",gap:6,
                      padding:"3px 7px",marginBottom:2,borderRadius:5,overflow:"hidden",
                      background:"var(--color-background-secondary)"}}>
                      <div style={{position:"absolute",left:0,top:0,bottom:0,width:`${pct}%`,
                        background:d.color,opacity:0.18}}></div>
                      <span style={{position:"relative",fontSize:11,color:"var(--color-text-tertiary)",fontWeight:700,
                        minWidth:15,textAlign:"center"}}>#{i+1}</span>
                      <span style={{position:"relative",width:8,height:8,borderRadius:"50%",background:d.color,flexShrink:0}}></span>
                      <span style={{position:"relative",fontSize:12,fontWeight:600,flex:1,overflow:"hidden",
                        textOverflow:"ellipsis",whiteSpace:"nowrap",color:"var(--color-text-primary)"}}
                        title={d.name}>{d.name}</span>
                      <span style={{position:"relative",fontSize:12,fontFamily:"var(--font-mono)",fontWeight:700,
                        color:"#2E9BFF",flexShrink:0,whiteSpace:"nowrap"}}>
                        {fmtU(d.gen??0,uDecR)}<span style={{fontSize:10,color:"var(--color-text-tertiary)",marginLeft:2}}>{unit}</span>
                      </span>
                      <span title="Diferença em relação à média do grupo"
                        style={{position:"relative",fontSize:11,fontFamily:"var(--font-mono)",fontWeight:700,
                          color:dpColor,minWidth:54,textAlign:"right",flexShrink:0}}>
                        {dp>=0?"+":""}{dp.toFixed(2)}%
                      </span>
                    </div>
                  );
                })}
              </div>
            );
          })}

          {groups.other.length>0&&(
            <div style={{flexShrink:0}}>
              {groups.other.map((d,i)=>{
                const gmax=groups.other[0]?.gen||0;
                const pct=gmax>0?d.gen/gmax*100:0;
                return(
                  <div key={d.invKey} style={{position:"relative",display:"flex",alignItems:"center",gap:6,
                    padding:"3px 7px",marginBottom:2,borderRadius:5,overflow:"hidden",
                    background:"var(--color-background-secondary)"}}>
                    <div style={{position:"absolute",left:0,top:0,bottom:0,width:`${pct}%`,
                      background:d.color,opacity:0.18}}></div>
                    <span style={{position:"relative",fontSize:11,color:"var(--color-text-tertiary)",fontWeight:700,
                      minWidth:15,textAlign:"center"}}>#{i+1}</span>
                    <span style={{position:"relative",width:8,height:8,borderRadius:"50%",background:d.color,flexShrink:0}}></span>
                    <span style={{position:"relative",fontSize:12,fontWeight:600,flex:1,overflow:"hidden",
                      textOverflow:"ellipsis",whiteSpace:"nowrap",color:"var(--color-text-primary)"}}
                      title={d.name}>{d.name}</span>
                    <span style={{position:"relative",fontSize:12,fontFamily:"var(--font-mono)",fontWeight:700,
                      color:"#2E9BFF",flexShrink:0,whiteSpace:"nowrap"}}>{fmtU(d.gen??0,uDecR)}
                      <span style={{fontSize:10,color:"var(--color-text-tertiary)",marginLeft:2}}>{unit}</span></span>
                  </div>
                );
              })}
            </div>
          )}

        </div>
      </div>
    </div>
  );
}

// ── Combiners (API INGECON SUN Monitor — corrente DC por combiner) ───────────
const SB_CACHE_PREFIX  = "ingecon_sb_v1_";
const SB_FETCH_SPACING_MS = 3500; // espaça chamadas p/ respeitar limite de 20 req distintas/min da API

function ymd(d) {
  return `${d.getFullYear()}${String(d.getMonth()+1).padStart(2,"0")}${String(d.getDate()).padStart(2,"0")}`;
}
function last30Dates() {
  const arr = [];
  const now = new Date();
  for (let i=29; i>=0; i--) {
    const d = new Date(now); d.setDate(d.getDate()-i);
    arr.push(ymd(d));
  }
  return arr;
}
function fmtYmd(yyyymmdd) {
  if (!yyyymmdd) return "—";
  return `${yyyymmdd.slice(6,8)}/${yyyymmdd.slice(4,6)}/${yyyymmdd.slice(0,4)}`;
}

// Cada GId termina em ".1ST" ou ".2ST" — caixas de posição ímpar (1) têm 17 entradas
// reais conectadas, caixas de posição par (2) têm só 16.
function sbChannelCount(gid) {
  const m = String(gid).match(/\.(\d+)ST$/i);
  if (!m) return 17;
  return (parseInt(m[1],10) % 2 === 1) ? 17 : 16;
}

// Agrupa as leituras por GId (nunca por SN — SN é do concentrador e é igual p/ toda a planta).
function groupStringboxByGId(records) {
  const byGid = {};
  (records||[]).forEach(r=>{
    if (!r.GId) return;
    const time = String(r.DateTime||"").slice(11,16);
    const count = sbChannelCount(r.GId);
    if (!byGid[r.GId]) byGid[r.GId] = [];
    byGid[r.GId].push({ time, idc:(r.Idc||[]).slice(0,count) });
  });
  Object.values(byGid).forEach(arr=>arr.sort((a,b)=>a.time.localeCompare(b.time)));
  return byGid;
}

// "SM3/INV3.1.1ST" → { inverter:"SDM3" (grupo/optgroup), pos:"3.1.1", label:"INV3.1.1" (item) }
function combinerMeta(gid) {
  const m = String(gid).match(/INV(\d+)\.(\d+\.\d+)ST$/i);
  if (!m) return { inverter:gid, pos:gid, label:gid };
  const inverter = `SDM${m[1]}`;
  const pos = `${m[1]}.${m[2]}`;
  return { inverter, pos, label:`INV${pos}` };
}

async function fetchStringboxDay(date, isToday, attempt=0) {
  if (!isToday) {
    const cached = localStorage.getItem(SB_CACHE_PREFIX+date);
    if (cached) {
      try { return JSON.parse(cached); } catch { /* cache corrompido, refaz a busca */ }
    }
  }
  const res = await fetch(`/api/stringbox?date=${date}`);
  if (res.status===429 && attempt<1) {
    await new Promise(r=>setTimeout(r,8000));
    return fetchStringboxDay(date, isToday, attempt+1);
  }
  if (!res.ok) {
    const body = await res.json().catch(()=>({}));
    throw new Error(body.error || `Erro ${res.status} ao buscar ${date}`);
  }
  const records = await res.json();
  const parsed = groupStringboxByGId(records);
  if (!isToday) {
    try { localStorage.setItem(SB_CACHE_PREFIX+date, JSON.stringify(parsed)); } catch { /* quota cheia — segue sem cache */ }
  }
  return parsed;
}

function CombinerChart({ records, onPointClick }) {
  const canvasRef = useRef(null);
  const chartRef  = useRef(null);
  const cbRef     = useRef(onPointClick);
  const [ready, setReady] = useState(!!window._chartReady);
  useEffect(()=>{ cbRef.current=onPointClick; },[onPointClick]);
  useEffect(()=>{ if(window._chartReady){registerSmartTooltip();setReady(true);return;} loadChartLibs().then(()=>setReady(true)); },[]);

  useEffect(()=>{
    if(!ready||!canvasRef.current) return;
    if(chartRef.current){chartRef.current.destroy();chartRef.current=null;}
    if(!records.length) return;

    const labels = records.map(r=>r.time);
    const channelCount = records.reduce((max,r)=>Math.max(max,r.idc.length),0);
    const datasets = Array.from({length:channelCount},(_,i)=>({
      label:`Entrada ${i+1}`,
      data: records.map(r=>{ const v=r.idc[i]; return (v!=null&&isFinite(v))?v:null; }),
      borderColor: PALETTE[i % PALETTE.length], backgroundColor:"transparent",
      borderWidth:1.5, pointRadius:0, pointHoverRadius:4, tension:0.35, spanGaps:false,
    }));

    chartRef.current = new window.Chart(canvasRef.current,{
      type:"line", data:{labels,datasets},
      options:{
        responsive:true, maintainAspectRatio:false, animation:false,
        interaction:{mode:"index",intersect:false},
        onClick:(_,els,ch)=>{if(!els.length)return;cbRef.current?.(ch.data.labels[els[0].index]);},
        plugins:{
          legend:{display:true,position:"bottom",labels:{color:"#8595A6",boxWidth:10,font:{size:10}}},
          tooltip:{position:"smart",backgroundColor:"rgba(38,50,68,0.97)",titleColor:"#A7B6C6",bodyColor:"#EAF2FB",
            borderColor:"rgba(46,155,255,0.40)",borderWidth:1,padding:10,
            callbacks:{title:items=>items[0]?.label??"",
              label:ctx=>`  ${ctx.dataset.label}: ${ctx.parsed.y!=null?ctx.parsed.y.toFixed(2)+" A":"—"}`,
              labelTextColor:ctx=>ctx.dataset.borderColor}},
          zoom:{pan:{enabled:true,mode:"xy",threshold:10},
            zoom:{wheel:{enabled:true,mode:"xy"},pinch:{enabled:false},drag:{enabled:false},mode:"xy"}},
        },
        scales:{
          x:{afterBuildTicks:axis=>{const tks=axis.ticks;if(!tks.length)return;const toMin=l=>(l&&l.length>=5)?parseInt(l.slice(0,2),10)*60+parseInt(l.slice(3,5),10):null;const f=toMin(axis.getLabelForValue(tks[0].value)),lt=toMin(axis.getLabelForValue(tks[tks.length-1].value));const span=(f!=null&&lt!=null)?Math.max(1,Math.abs(lt-f)):60;const steps=[1,2,5,10,15,20,30,60,120,180];let step=180;for(const s of steps){if(span/s<=14){step=s;break;}}axis.ticks=tks.filter(tk=>{const m=toMin(axis.getLabelForValue(tk.value));return m!=null&&m%step===0;});},
            ticks:{color:"#8595A6",font:{size:11},maxRotation:0,autoSkip:false,callback:function(value){return this.getLabelForValue(value)||"";}},
            grid:{color:"rgba(148,163,184,0.10)"},border:{color:"rgba(148,163,184,0.22)"}},
          y:{position:"left",ticks:{color:"#8595A6",font:{size:11},maxTicksLimit:8,callback:v=>v.toFixed(0)+"A"},
            grid:{color:"rgba(148,163,184,0.10)"},border:{color:"rgba(148,163,184,0.22)"},
            title:{display:true,text:"Corrente DC (A)",color:"#A7B6C6",font:{size:11}}},
        },
      },
    });
    return()=>{if(chartRef.current){chartRef.current.destroy();chartRef.current=null;}};
  },[ready,records]);

  return(
    <div style={{position:"relative",width:"100%",height:"100%"}}>
      {(!ready||!records.length)&&(
        <div style={{position:"absolute",inset:0,display:"flex",alignItems:"center",justifyContent:"center",
          color:"var(--color-text-tertiary)",fontSize:13}}>
          {!ready?"Carregando…":"Sem dados para este dia"}
        </div>
      )}
      <canvas ref={canvasRef} role="img" style={{cursor:"crosshair"}} onDoubleClick={()=>chartRef.current?.resetZoom()}/>
    </div>
  );
}

function CombinerPanel() {
  const dates = useMemo(()=>last30Dates(),[]);
  const today = dates[dates.length-1];
  const [dayData, setDayData]   = useState({});   // { [date]: { [GId]: {time,idc[]}[] } }
  const [progress, setProgress] = useState({loaded:0,total:dates.length});
  const [selDate, setSelDate]   = useState(today);
  const [selGid, setSelGid]     = useState(null);
  const [fatalError, setFatalError] = useState(null);
  const [selectedTime, setSelectedTime] = useState(null);

  useEffect(()=>{ setSelectedTime(null); },[selDate]);

  useEffect(()=>{
    // Flag local por execução do efeito — um useRef compartilhado não isolaria
    // corretamente remounts (StrictMode em dev, ou trocar de aba e voltar rápido):
    // o cleanup de uma execução antiga acabaria cancelando a corrida da nova.
    let cancelled = false;
    const order = dates.slice().reverse(); // hoje primeiro (dado mais recente possível), depois regride no tempo
    setProgress({loaded:0,total:order.length});

    (async () => {
      let loaded = 0;
      for (const date of order) {
        if (cancelled) return;
        const isToday = date === today;
        try {
          const parsed = await fetchStringboxDay(date, isToday);
          if (cancelled) return;
          setDayData(prev=>({...prev,[date]:parsed}));
        } catch (err) {
          if (loaded===0) setFatalError(err.message);
        }
        loaded++;
        setProgress({loaded, total:order.length});
        if (date !== order[order.length-1]) {
          await new Promise(r=>setTimeout(r, SB_FETCH_SPACING_MS));
        }
      }
    })();

    return ()=>{ cancelled = true; };
  },[dates, today]);

  useEffect(()=>{ if(Object.keys(dayData).length>0 && fatalError) setFatalError(null); },[dayData, fatalError]);

  const allGids = useMemo(()=>{
    const s = new Set();
    Object.values(dayData).forEach(byGid=>Object.keys(byGid).forEach(g=>s.add(g)));
    return [...s].sort();
  },[dayData]);

  useEffect(()=>{ if(!selGid && allGids.length) setSelGid(allGids[0]); },[allGids, selGid]);

  const gidsByInverter = useMemo(()=>{
    const map = {};
    allGids.forEach(g=>{
      const {inverter} = combinerMeta(g);
      if (!map[inverter]) map[inverter] = [];
      map[inverter].push(g);
    });
    return map;
  },[allGids]);

  const dateIdx = dates.indexOf(selDate);
  const goPrev = () => { if(dateIdx>0) setSelDate(dates[dateIdx-1]); };
  const goNext = () => { if(dateIdx<dates.length-1) setSelDate(dates[dateIdx+1]); };

  const dateLoaded = !!dayData[selDate];
  const records = (selGid && dayData[selDate]?.[selGid]) || [];

  const summary = useMemo(()=>{
    if (!records.length) return null;
    const last = records[records.length-1];
    const vals = last.idc.map((v,i)=>({i,v})).filter(x=>x.v!=null&&isFinite(x.v));
    if (!vals.length) return null;
    const min = vals.reduce((a,b)=>b.v<a.v?b:a);
    const max = vals.reduce((a,b)=>b.v>a.v?b:a);
    return { time:last.time, min, max };
  },[records]);

  // Ranking de todos os 30 combinadores no horário clicado, separado por 16 vs 17 entradas
  // (comparar direto entre grupos seria injusto — um combinador de 17 soma mais só por ter mais canais).
  const rankingAnalysis = useMemo(()=>{
    if (!selectedTime) return null;
    const byGid = dayData[selDate];
    if (!byGid) return null;
    const rows = Object.entries(byGid).map(([gid,recs])=>{
      const rec = recs.find(r=>r.time===selectedTime);
      if (!rec) return null;
      const vals = rec.idc.filter(v=>v!=null&&isFinite(v));
      if (!vals.length) return null;
      return { gid, label:combinerMeta(gid).label, channelCount:rec.idc.length,
        avg: vals.reduce((s,v)=>s+v,0)/vals.length, idc:rec.idc };
    }).filter(Boolean);

    const worstEntradas = row => !row ? [] : row.idc
      .map((v,i)=>({i,v}))
      .filter(x=>x.v!=null&&isFinite(x.v))
      .sort((a,b)=>a.v-b.v)
      .slice(0,4);

    const cohort16 = rows.filter(r=>r.channelCount===16).sort((a,b)=>a.avg-b.avg);
    const cohort17 = rows.filter(r=>r.channelCount===17).sort((a,b)=>a.avg-b.avg);
    return {
      cohort16, cohort17,
      worst16: cohort16[0] ? { ...cohort16[0], worstEntradas:worstEntradas(cohort16[0]) } : null,
      worst17: cohort17[0] ? { ...cohort17[0], worstEntradas:worstEntradas(cohort17[0]) } : null,
    };
  },[selectedTime, dayData, selDate]);

  return(
    <div style={{flex:1,display:"flex",flexDirection:"column",overflow:"hidden",padding:"10px 16px 8px",minHeight:0}}>
      {/* Header */}
      <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:10,flexShrink:0,flexWrap:"wrap"}}>

        {/* Navegação de data (últimos 30 dias) */}
        <div style={{display:"flex",alignItems:"center",gap:4}}>
          <button onClick={goPrev} disabled={dateIdx<=0}
            style={{background:"none",border:"none",cursor:dateIdx>0?"pointer":"default",fontSize:15,
              color:dateIdx>0?"var(--color-text-secondary)":"var(--color-text-tertiary)",padding:"2px 4px"}}>
            <i className="ti ti-chevron-left"/>
          </button>
          <span style={{fontSize:13,fontWeight:600,minWidth:88,textAlign:"center",fontFamily:"var(--font-mono)"}}>
            {fmtYmd(selDate)}
          </span>
          <button onClick={goNext} disabled={dateIdx>=dates.length-1}
            style={{background:"none",border:"none",cursor:dateIdx<dates.length-1?"pointer":"default",fontSize:15,
              color:dateIdx<dates.length-1?"var(--color-text-secondary)":"var(--color-text-tertiary)",padding:"2px 4px"}}>
            <i className="ti ti-chevron-right"/>
          </button>
          {!dateLoaded&&(
            <span style={{fontSize:13,color:"var(--color-text-tertiary)",display:"flex",alignItems:"center",gap:4}}>
              <i className="ti ti-loader-2" style={{fontSize:13}}/> carregando…
            </span>
          )}
        </div>

        {/* Seletor de combinador */}
        <select value={selGid||""} onChange={e=>setSelGid(e.target.value)} disabled={!allGids.length}
          style={{fontSize:12,maxWidth:220,padding:"5px 10px",borderRadius:7,cursor:"pointer",outline:"none",
            background:"#1A2433",color:"#E6EDF5",border:"1px solid rgba(255,255,255,0.12)"}}>
          {Object.entries(gidsByInverter).map(([inv,gids])=>(
            <optgroup key={inv} label={inv} style={{background:"#18222F",color:"#9FB0C0"}}>
              {gids.map(g=>(
                <option key={g} value={g} style={{background:"#18222F",color:"#E6EDF5"}}>
                  {combinerMeta(g).label}
                </option>
              ))}
            </optgroup>
          ))}
        </select>

        {/* Progresso do backfill de 30 dias */}
        {progress.loaded<progress.total&&(
          <div style={{display:"flex",alignItems:"center",gap:6,fontSize:13,color:"var(--color-text-tertiary)"}}>
            <div style={{width:80,height:5,borderRadius:3,background:"var(--color-background-secondary)",overflow:"hidden"}}>
              <div style={{width:`${(progress.loaded/progress.total*100).toFixed(0)}%`,height:"100%",
                background:"#2E9BFF",transition:"width 0.3s"}}/>
            </div>
            <span>histórico {progress.loaded}/{progress.total}</span>
          </div>
        )}

        {/* Resumo da última leitura do dia selecionado */}
        {summary&&(
          <div style={{marginLeft:"auto",display:"flex",alignItems:"center",gap:14,fontSize:13,
            color:"var(--color-text-secondary)",flexWrap:"wrap"}}>
            <span>Última leitura <strong style={{fontFamily:"var(--font-mono)"}}>{summary.time}</strong></span>
            <span style={{color:"#4CAF50"}}>▲ Entrada {summary.max.i+1} ({summary.max.v.toFixed(1)}A)</span>
            <span style={{color:"#F44336"}}>▼ Entrada {summary.min.i+1} ({summary.min.v.toFixed(1)}A)</span>
            {selectedTime&&(
              <span style={{color:"#2E9BFF",display:"inline-flex",alignItems:"center",gap:4}}>
                <i className="ti ti-map-pin" style={{fontSize:13}}/><strong>{selectedTime}</strong>
                <button onClick={()=>setSelectedTime(null)}
                  style={{background:"none",border:"none",cursor:"pointer",fontSize:13,color:"#2E9BFF",padding:0,lineHeight:1,marginLeft:1}}>
                  <i className="ti ti-x"/>
                </button>
              </span>
            )}
          </div>
        )}
      </div>

      {/* Gráfico */}
      <div style={{flex:selectedTime?"1 1 58%":"1 1 auto",minHeight:0}}>
        {fatalError?(
          <div style={{height:"100%",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",
            gap:10,color:"var(--color-text-tertiary)",textAlign:"center"}}>
            <i className="ti ti-plug-connected-x" style={{fontSize:40}}/>
            <div style={{fontSize:13,maxWidth:380}}>{fatalError}</div>
          </div>
        ):!allGids.length?(
          <div style={{height:"100%",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",
            gap:10,color:"var(--color-text-tertiary)",textAlign:"center"}}>
            <i className="ti ti-plug-connected" style={{fontSize:40}}/>
            <div style={{fontSize:13}}>Carregando combiners da API…</div>
          </div>
        ):(
          <CombinerChart records={records} onPointClick={setSelectedTime}/>
        )}
      </div>

      {/* Análise: ranking dos 30 combinadores no horário clicado, separado por 16 vs 17 entradas */}
      {!selectedTime?(
        <div style={{flexShrink:0,paddingTop:8,fontSize:13,color:"var(--color-text-tertiary)",textAlign:"center"}}>
          <i className="ti ti-hand-click" style={{marginRight:5}}/>Clique num ponto do gráfico para comparar todos os combinadores nesse horário
        </div>
      ):rankingAnalysis&&(
        <div style={{flex:"0 0 38%",display:"flex",gap:12,paddingTop:10,marginTop:8,minHeight:0,
          borderTop:"0.5px solid var(--color-border-tertiary)"}}>
          {[{title:"16 entradas",arr:rankingAnalysis.cohort16,worst:rankingAnalysis.worst16},
            {title:"17 entradas",arr:rankingAnalysis.cohort17,worst:rankingAnalysis.worst17}].map(({title,arr,worst})=>(
            <div key={title} style={{flex:1,display:"flex",flexDirection:"column",overflow:"hidden",minWidth:0}}>
              <div style={{fontSize:13,fontWeight:600,color:"var(--color-text-tertiary)",textTransform:"uppercase",
                letterSpacing:"0.05em",marginBottom:6,flexShrink:0}}>
                Combinadores de {title} — piores às {selectedTime}
              </div>
              {worst&&(
                <div style={{flexShrink:0,marginBottom:6,padding:"6px 8px",borderRadius:6,
                  background:"rgba(244,67,54,0.08)",border:"1px solid rgba(244,67,54,0.28)"}}>
                  <div style={{fontSize:13,color:"var(--color-text-secondary)"}}>
                    Pior combinador: <strong style={{color:"#F44336"}}>{worst.label}</strong>
                    <span style={{color:"var(--color-text-tertiary)"}}> ({worst.avg.toFixed(2)} A méd.)</span>
                  </div>
                  <div style={{fontSize:13,color:"var(--color-text-tertiary)",marginTop:2}}>
                    4 piores entradas:{" "}
                    {worst.worstEntradas.map((e,i)=>(
                      <span key={e.i} style={{color:"var(--color-text-secondary)"}}>
                        {i>0&&", "}Entrada {e.i+1} ({e.v.toFixed(2)}A)
                      </span>
                    ))}
                  </div>
                </div>
              )}
              <div style={{overflowY:"auto",flex:1,display:"flex",flexDirection:"column",gap:3}}>
                {!arr.length?(
                  <div style={{fontSize:13,color:"var(--color-text-tertiary)"}}>Sem dados neste horário</div>
                ):arr.map((r,i)=>{
                  const worst = i===0;
                  const isSelected = r.gid===selGid;
                  return(
                    <div key={r.gid} onClick={()=>setSelGid(r.gid)}
                      style={{display:"flex",alignItems:"center",gap:8,padding:"3px 8px",borderRadius:5,cursor:"pointer",
                        background:isSelected?"rgba(46,155,255,0.14)":worst?"rgba(244,67,54,0.10)":"var(--color-background-secondary)",
                        border:isSelected?"1px solid rgba(46,155,255,0.45)":"1px solid transparent"}}>
                      <span style={{fontSize:13,color:"var(--color-text-tertiary)",fontWeight:700,minWidth:16,textAlign:"center"}}>#{i+1}</span>
                      <span style={{fontSize:13,fontWeight:600,flex:1,overflow:"hidden",textOverflow:"ellipsis",
                        whiteSpace:"nowrap",color:"var(--color-text-primary)"}} title={r.label}>{r.label}</span>
                      <span style={{fontSize:13,fontFamily:"var(--font-mono)",fontWeight:700,flexShrink:0,
                        color:worst?"#F44336":"var(--color-text-secondary)"}}>{r.avg.toFixed(2)} A</span>
                    </div>
                  );
                })}
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

export default function WrappedApp() { return <ErrorBoundary><App/></ErrorBoundary>; }
