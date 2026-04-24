/**
 * SGO — charts.js
 * Sparklines, gráfico diário (Chart.js) e ranking de técnicos.
 */

import { state, DN } from './state.js';
import { isMobile } from './ui.js';

// ── Sparkline ──────────────────────────────────────────────
export function drawSparkline(id, data, color) {
  const c = document.getElementById(id);
  if (!c) return;
  const ctx = c.getContext('2d');
  const W = c.offsetWidth || 60, H = c.offsetHeight || 26;
  c.width = W * 2; c.height = H * 2; ctx.scale(2, 2);
  ctx.clearRect(0, 0, W, H);
  if (!data || data.length < 2) return;

  const mn = Math.min(...data), mx = Math.max(...data), rng = mx - mn || 1, pad = 3;
  const pts = data.map((v, i) => ({
    x: pad + (i / (data.length - 1)) * (W - pad * 2),
    y: H - pad - (v - mn) / rng * (H - pad * 2)
  }));

  const grad = ctx.createLinearGradient(0, 0, 0, H);
  grad.addColorStop(0, color + '44');
  grad.addColorStop(1, color + '00');

  ctx.beginPath();
  ctx.moveTo(pts[0].x, pts[0].y);
  pts.slice(1).forEach(p => ctx.lineTo(p.x, p.y));
  ctx.lineTo(pts[pts.length - 1].x, H);
  ctx.lineTo(pts[0].x, H);
  ctx.closePath();
  ctx.fillStyle = grad; ctx.fill();

  ctx.beginPath();
  ctx.moveTo(pts[0].x, pts[0].y);
  pts.slice(1).forEach(p => ctx.lineTo(p.x, p.y));
  ctx.strokeStyle = color; ctx.lineWidth = 1.5; ctx.stroke();
}

// ── Daily Bar Chart ────────────────────────────────────────
export function drawDailyChart(dailyArr) {
  const canvas = document.getElementById('dailyChart');
  if (!canvas) return;

  if (state.dailyChartInstance) {
    state.dailyChartInstance.destroy();
    state.dailyChartInstance = null;
  }

  const dark = state.isDark;
  const maxVal = Math.max(...dailyArr);
  const colors = dailyArr.map(v => {
    if (v === 0) return dark ? 'rgba(48,54,61,.4)' : 'rgba(228,232,239,.6)';
    const pct = v / maxVal;
    if (pct > 0.75) return 'rgba(22,163,74,.75)';
    if (pct > 0.45) return 'rgba(29,111,232,.7)';
    if (pct > 0.20) return 'rgba(202,138,4,.7)';
    return 'rgba(220,38,38,.7)';
  });

  state.dailyChartInstance = new Chart(canvas, {
    type: 'bar',
    data: {
      labels: dailyArr.map((_, i) => i + 1),
      datasets: [{ data: dailyArr, backgroundColor: colors, borderRadius: 3, borderSkipped: false }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: { label: ctx => `${ctx.raw} O.S.` },
          bodyFont: { family: 'Geist Mono' },
          titleFont: { family: 'Geist Mono', size: 10 }
        }
      },
      scales: {
        x: {
          ticks: { font: { family: 'Geist Mono', size: isMobile() ? 7 : 8 }, color: dark ? '#484F58' : '#9CA3AF', maxRotation: 0 },
          grid: { display: false }, border: { display: false }
        },
        y: {
          ticks: { font: { family: 'Geist Mono', size: isMobile() ? 7 : 8 }, color: dark ? '#484F58' : '#9CA3AF', maxTicksLimit: 5 },
          grid: { color: dark ? 'rgba(48,54,61,.5)' : 'rgba(228,232,239,.8)' }, border: { display: false }
        }
      }
    }
  });
}

// ── Ranking ────────────────────────────────────────────────
export function buildRanking(filtered) {
  const techTotals = {};
  filtered.forEach(i => { techTotals[i.techKey] = (techTotals[i.techKey] || 0) + 1; });

  const sorted = Object.entries(techTotals).sort((a, b) => b[1] - a[1]).slice(0, 8);
  if (!sorted.length) return;

  const maxVal    = sorted[0][1];
  const medals    = ['🥇', '🥈', '🥉'];
  const posClass  = ['gold', 'silver', 'bronze'];

  const rb = document.getElementById('rankBody');
  if (!rb) return;

  rb.innerHTML = sorted.map(([tk, total], i) => {
    const nome = state.teamData[tk]?.originalName || tk;
    const base = state.teamData[tk]?.base || '';
    const pct  = Math.round(total / maxVal * 100);
    const pos  = i < 3
      ? `<span class="rank-pos ${posClass[i]}">${medals[i]}</span>`
      : `<span class="rank-pos">${i + 1}</span>`;
    return `<div class="rank-item">
      ${pos}
      <div style="flex:1;min-width:0;">
        <div class="rank-name">${nome}</div>
        <div class="rank-base">${base}</div>
      </div>
      <div class="rank-bar-wrap"><div class="rank-bar-fill" style="width:${pct}%;"></div></div>
      <div class="rank-val">${total}</div>
    </div>`;
  }).join('');

  const rt = document.getElementById('rankTotal');
  if (rt) rt.textContent = `${sorted.length} técnicos`;
}

// ── KPIs ──────────────────────────────────────────────────
export function updateDashboardStats(filtered) {
  const grid = document.getElementById('kpiGrid');
  if (grid) grid.style.display = 'grid';

  const tot = filtered.length;
  document.getElementById('dashTotalOs').textContent = tot || '—';

  if (!tot) {
    ['dashActiveTechs', 'dashAvgTech', 'dashTopBase', 'dashCriticalBase']
      .forEach(id => document.getElementById(id).textContent = '—');
    return;
  }

  const ts = new Set(filtered.map(i => i.techKey));
  document.getElementById('dashActiveTechs').textContent = ts.size;
  document.getElementById('dashAvgTech').textContent = Math.round(tot / ts.size);

  const [ys, ms] = state.activeMonthYear.split('-');
  const dIM = new Date(parseInt(ys), parseInt(ms), 0).getDate();
  const dayMap = {};
  filtered.forEach(i => { dayMap[i.day] = (dayMap[i.day] || 0) + 1; });
  const dailyArr = Array.from({ length: dIM }, (_, i) => dayMap[i + 1] || 0);

  setTimeout(() => {
    drawSparkline('spark0', dailyArr, '#E07B1F');
    const tArr = Array.from(ts).map(tk => {
      let c = 0; filtered.forEach(i => { if (i.techKey === tk) c++; }); return c;
    }).sort((a, b) => a - b);
    drawSparkline('spark1', tArr, '#1D6FE8');
    drawSparkline('spark2', dailyArr.map(v => v > 0 ? Math.round(v / ts.size) : 0), '#8B5CF6');
    drawDailyChart(dailyArr);
    buildRanking(filtered);
  }, 100);

  // Top / Bottom bases
  const bs = {};
  filtered.forEach(i => {
    if (!bs[i.cidade]) bs[i.cidade] = { total: 0, techs: new Set() };
    bs[i.cidade].total++;
    bs[i.cidade].techs.add(i.techKey);
  });
  const arr = Object.keys(bs)
    .map(b => ({ nome: b, media: bs[b].total / bs[b].techs.size }))
    .sort((a, b) => b.media - a.media);

  if (arr.length) {
    document.getElementById('dashTopBase').textContent = arr[0].nome;
    document.getElementById('dashCriticalBase').textContent = arr[arr.length - 1].nome;
    const tb = document.getElementById('dashTopBadge');
    const cb = document.getElementById('dashCritBadge');
    if (tb) { tb.style.display = 'inline-flex'; tb.textContent = '▲ ' + Math.round(arr[0].media); }
    if (cb) { cb.style.display = 'inline-flex'; cb.textContent = '▼ ' + Math.round(arr[arr.length - 1].media); }
  }

  const ar = document.getElementById('analysisRow');
  if (ar) ar.style.display = 'grid';
  const cl = document.getElementById('chartMonthLabel');
  if (cl) cl.textContent = state.activeMonthYear;
}
