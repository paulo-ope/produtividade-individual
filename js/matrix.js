/**
 * SGO — matrix.js
 * Geração da matriz operacional e exportação para Excel.
 */

import { state, DN, TEAM_TYPES } from './state.js';
import { isMobile } from './ui.js';

// ── Helpers de cor ─────────────────────────────────────────
export function vC(v, m) {
  if (!v || v === 0) return 'vd';
  if (v >= m)     return 'vg';
  if (v >= m - 1) return 'vb';
  if (v >= m - 2) return 'vy';
  return 'vr';
}

export function tC(v, ce, cb, cm) {
  if (v >= ce) return 'tg';
  if (v >= cb) return 'tb';
  if (v >= cm) return 'ty';
  return 'tr';
}

// ── Construir mapa de OS por cidade/técnico ────────────────
function buildCityMap(filtered) {
  const cm = {};
  filtered.forEach(i => {
    if (!cm[i.cidade]) cm[i.cidade] = {};
    if (!cm[i.cidade][i.techKey]) cm[i.cidade][i.techKey] = { dias: {}, total: 0, tipo: i.tipo };
    cm[i.cidade][i.techKey].dias[i.day] = (cm[i.cidade][i.techKey].dias[i.day] || 0) + 1;
    cm[i.cidade][i.techKey].total++;
  });
  return cm;
}

// ── Gerar cabeçalho de semanas ─────────────────────────────
function buildWeekHeader(dIM, fdw) {
  let r0 = [], r1 = [], r2 = [];
  let cw = 1, cdw = 0, dic = 0;

  if (fdw > 0) {
    r0.push({ colspan: 7, label: 'SEM 1', cls: 'week-th' });
    r0.push({ colspan: 1, label: '', cls: 'week-th' });
    for (let i = 0; i < fdw; i++) {
      const wk = cdw === 0 || cdw === 6;
      r1.push({ label: '—', wk }); r2.push({ label: DN[cdw], wk });
      cdw++; dic++;
    }
  }

  for (let d = 1; d <= dIM; d++) {
    if (cdw > 6) {
      r1.push({ label: 'TOT', cls: 'ctot' }); r2.push({ label: '' });
      cdw = 0; cw++; dic = 0;
    }
    if (dic === 0) {
      r0.push({ colspan: 7, label: `SEM ${cw}`, cls: 'week-th' });
      r0.push({ colspan: 1, label: '', cls: 'week-th' });
    }
    const wk = cdw === 0 || cdw === 6;
    r1.push({ label: d, wk }); r2.push({ label: DN[cdw], wk });
    cdw++; dic++;
  }

  if (dic > 0) {
    for (let i = 0; i < 7 - dic; i++) {
      const wk = cdw === 0 || cdw === 6;
      r1.push({ label: '—', wk }); r2.push({ label: '' });
      cdw++;
    }
    r1.push({ label: 'TOT', cls: 'ctot' }); r2.push({ label: '' });
  }

  return { r0, r1, r2, finalCdw: cdw, finalDic: dic };
}

// ── Renderizar string do cabeçalho HTML ────────────────────
function renderHeaderHTML(dIM, fdw) {
  let r0 = `<tr><th rowspan="3" class="cn" style="background:var(--mat-hd);color:#8B949E;text-align:left!important;">Equipe / Técnico</th>`;
  let r1 = `<tr>`, r2 = `<tr>`;
  let cw = 1, cdw = 0, dic = 0;

  if (fdw > 0) {
    r0 += `<th colspan="7" class="week-th">SEM 1</th><th class="week-th"></th>`;
    for (let i = 0; i < fdw; i++) {
      const wk = cdw === 0 || cdw === 6;
      r1 += `<th class="${wk ? 'wknd' : ''}">—</th>`;
      r2 += `<th class="${wk ? 'wknd' : ''}" style="font-size:7px;">${DN[cdw]}</th>`;
      cdw++; dic++;
    }
  }

  for (let d = 1; d <= dIM; d++) {
    if (cdw > 6) {
      r1 += `<th class="ctot" style="background:var(--mat-hd);border-left:1px solid rgba(255,255,255,.06)!important;">TOT</th>`;
      r2 += `<th style="background:var(--mat-hd);"></th>`;
      cdw = 0; cw++; dic = 0;
    }
    if (dic === 0) {
      r0 += `<th colspan="7" class="week-th">SEM ${cw}</th><th class="week-th"></th>`;
    }
    const wk = cdw === 0 || cdw === 6;
    r1 += `<th class="${wk ? 'wknd' : ''}">${d}</th>`;
    r2 += `<th class="${wk ? 'wknd' : ''}" style="font-size:7px;letter-spacing:.03em;">${DN[cdw]}</th>`;
    cdw++; dic++;
  }

  if (dic > 0) {
    for (let i = 0; i < 7 - dic; i++) {
      const wk = cdw === 0 || cdw === 6;
      r1 += `<th class="${wk ? 'wknd' : ''}">—</th>`;
      r2 += `<th class="${wk ? 'wknd' : ''}"></th>`;
      cdw++;
    }
    r1 += `<th class="ctot" style="background:var(--mat-hd);">TOT</th>`;
    r2 += `<th style="background:var(--mat-hd);"></th>`;
  }

  r0 += `<th rowspan="3" class="ctos" style="background:var(--mat-wk);color:#93C5FD;min-width:50px;">TOTAL<br>O.S.</th>`;
  r0 += `<th colspan="3" class="week-th">CAP.</th></tr>`;
  r1 += `<th class="ccap" style="color:rgba(22,163,74,.5)!important;background:var(--mat-hd);">Exc.</th><th class="ccap" style="color:rgba(29,111,232,.5)!important;background:var(--mat-hd);">Bom</th><th class="ccap" style="color:rgba(202,138,4,.5)!important;background:var(--mat-hd);">Med.</th></tr>`;
  r2 += `<th style="background:var(--mat-hd);"></th><th style="background:var(--mat-hd);"></th><th style="background:var(--mat-hd);"></th></tr>`;

  return r0 + r1 + r2;
}

// ── Linha de técnico HTML ──────────────────────────────────
function renderTechRow(tk, data, dIM, fdw) {
  const meta  = state.appSettings.metasDiarias[data.tipo] || 5;
  const nome  = state.teamData[tk].originalName;
  const tipo  = TEAM_TYPES[data.tipo] || data.tipo;
  const bm    = data.tipo === 'TECNICO 12/36H' ? 15 : 24;
  const ce    = meta * bm, cb = (meta - 1) * bm, cm2 = (meta - 2) * bm;

  let row = `<tr><td class="cn"><div class="cn-name">${nome}</div><span class="cn-type">${tipo}</span></td>`;
  let wt = 0, cur = 0;

  for (let i = 0; i < fdw; i++) { row += `<td class="${cur === 0 || cur === 6 ? 'cwknd' : ''}"></td>`; cur++; }

  for (let d = 1; d <= dIM; d++) {
    if (cur > 6) { row += `<td class="ctot">${wt > 0 ? wt : ''}</td>`; wt = 0; cur = 0; }
    const v  = data.dias[d] || 0; wt += v;
    const vc = vC(v, meta);
    const wk = cur === 0 || cur === 6;
    row += `<td class="${vc}${wk ? ' cwknd' : ''}">${v > 0 ? v : ''}</td>`;
    cur++;
  }

  if (cur > 0) {
    for (let i = 0; i < 7 - cur; i++) row += `<td class="${cur === 0 || cur === 6 ? 'cwknd' : ''}"></td>`;
    row += `<td class="ctot">${wt > 0 ? wt : ''}</td>`;
  }

  row += `<td class="ctos ${tC(data.total, ce, cb, cm2)}">${data.total}</td>`;
  row += `<td class="ccap cg">${ce}</td><td class="ccap cb">${cb > 0 ? cb : 0}</td><td class="ccap cy">${cm2 > 0 ? cm2 : 0}</td></tr>`;
  return row;
}

// ── Gerar Matriz ───────────────────────────────────────────
export function generateMatrix(filtered) {
  const wrapper = document.getElementById('matrixWrapper');

  if (!filtered.length) {
    wrapper.innerHTML = `<div style="padding:40px;text-align:center;font-family:var(--mono);font-size:11px;color:var(--text-3);background:var(--card-bg);border:1px solid var(--card-border);border-radius:var(--r-xl);">Nenhum dado para os filtros selecionados.</div>`;
    return;
  }

  const [ys, ms]  = state.activeMonthYear.split('-');
  const year      = parseInt(ys), mIdx = parseInt(ms) - 1;
  const dIM       = new Date(year, mIdx + 1, 0).getDate();
  const fdw       = new Date(year, mIdx, 1).getDay();
  const cityMap   = buildCityMap(filtered);

  let html = '';
  Object.keys(cityMap).sort().forEach(cidade => {
    if (state.selectedCityTab !== 'ALL' && cidade !== state.selectedCityTab) return;
    const techs  = cityMap[cidade];
    const sorted = Object.entries(techs).sort((a, b) => b[1].total - a[1].total);
    const totOS  = Object.values(techs).reduce((s, t) => s + t.total, 0);

    const thead = renderHeaderHTML(dIM, fdw);
    const tbody = sorted.map(([tk, data]) => renderTechRow(tk, data, dIM, fdw)).join('');

    html += `<div class="mat-block">
      <div class="mat-hdr">
        <div class="mat-dot"></div>
        <div class="mat-city">${cidade}</div>
        <div class="mat-badge">${sorted.length} técnicos · ${totOS} O.S.</div>
      </div>
      <div class="mat-scroll">
        <table class="mat-table"><thead>${thead}</thead><tbody>${tbody}</tbody></table>
      </div>
    </div>`;
  });

  wrapper.innerHTML = html;
  if (isMobile()) wrapper.classList.add('compact');
}

// ── Export Excel ───────────────────────────────────────────
export function exportToExcel() {
  if (!state.globalRawData.length) return alert('Sem dados para exportar.');

  const [ys, ms] = state.activeMonthYear.split('-');
  const year = parseInt(ys), mIdx = parseInt(ms) - 1;
  const dIM  = new Date(year, mIdx + 1, 0).getDate();
  const fdw  = new Date(year, mIdx, 1).getDay();
  const fc   = state.selectedCityTab;
  const ft   = document.getElementById('filterType')?.value || 'ALL';

  const filtered = state.globalRawData.filter(i =>
    i.monthStr === state.activeMonthYear &&
    (fc === 'ALL' || i.cidade === fc) &&
    (ft === 'ALL' || i.tipo === ft)
  );
  if (!filtered.length) return alert('Sem dados nos filtros selecionados.');

  const cityMap = buildCityMap(filtered);
  const wb = XLSX.utils.book_new();

  const XL = {
    navy:'FF0D1117', navyMid:'FF1A2236', navyLight:'FF243047', navyTxt:'FFCBD5E1',
    weekTxt:'FF93C5FD', white:'FFFFFFFF', offWhite:'FFFAFBFC', offWhite2:'FFF4F6F9',
    border:'FF000000',
    greenBg:'FF16A34A', greenFg:'FFFFFFFF',
    blueBg:'FF1D6FE8',  blueFg:'FFFFFFFF',
    yellowBg:'FFCA8A04',yellowFg:'FFFFFFFF',
    redBg:'FFDC2626',   redFg:'FFFFFFFF',
    wkndBg:'FFF1F5F9', totBg:'FFEFF3F9', totFg:'FF64748B',
    accentBg:'FFFEF3E2', accentFg:'FFE07B1F'
  };

  const cs = (fill, font, bold = false, align = 'center', sz = 9) => ({
    fill: fill ? { fgColor: { rgb: fill } } : {},
    font: { color: { rgb: font || XL.navy }, bold, name: 'Calibri', sz },
    border: {
      top:    { style: 'thin', color: { rgb: XL.border } },
      bottom: { style: 'thin', color: { rgb: XL.border } },
      left:   { style: 'thin', color: { rgb: XL.border } },
      right:  { style: 'thin', color: { rgb: XL.border } }
    },
    alignment: { horizontal: align, vertical: 'center', wrapText: false }
  });

  Object.keys(cityMap).sort().forEach(cidade => {
    const techs  = cityMap[cidade];
    const sorted = Object.entries(techs).sort((a, b) => b[1].total - a[1].total);

    let wsData = [], merges = [], r0 = ['Técnico'], r1 = [''], r2 = [''];
    let cw = 1, cdw = 0, dic = 0, ci = 1;

    if (fdw > 0) {
      r0.push('SEM 1'); merges.push({ s: { r: 0, c: ci }, e: { r: 0, c: ci + 6 } });
      for (let i = 0; i < 6; i++) r0.push(''); r0.push('');
      for (let i = 0; i < fdw; i++) { r1.push('—'); r2.push(DN[cdw]); cdw++; dic++; ci++; }
    }

    for (let d = 1; d <= dIM; d++) {
      if (cdw > 6) { r1.push('TOT'); r2.push(''); ci++; cdw = 0; cw++; dic = 0; }
      if (dic === 0) {
        r0.push(`SEM ${cw}`); merges.push({ s: { r: 0, c: ci }, e: { r: 0, c: ci + 6 } });
        for (let i = 0; i < 6; i++) r0.push(''); r0.push('');
      }
      r1.push(d); r2.push(DN[cdw]); cdw++; dic++; ci++;
    }

    if (dic > 0) {
      let rem = 7 - dic;
      for (let i = 0; i < rem; i++) { r1.push('—'); r2.push(DN[cdw]); cdw++; ci++; }
      r1.push('TOT'); r2.push(''); ci++;
    }

    r0.push('TOTAL'); merges.push({ s: { r: 0, c: ci }, e: { r: 2, c: ci } });
    r0.push('CAPACIDADE'); merges.push({ s: { r: 0, c: ci + 1 }, e: { r: 0, c: ci + 3 } });
    r0.push(''); r0.push('');
    r1.push(''); r1.push('Excelente'); r1.push('Bom'); r1.push('Mediano');
    r2.push(''); r2.push(''); r2.push(''); r2.push('');

    wsData.push(r0); wsData.push(r1); wsData.push(r2);

    let mnl = 18;
    sorted.forEach(([tk, data]) => {
      const meta = state.appSettings.metasDiarias[data.tipo] || 5;
      const nd   = state.teamData[tk].originalName;
      if (nd.length > mnl) mnl = nd.length;

      let row = [nd], wt = 0, cur = 0;
      for (let i = 0; i < fdw; i++) { row.push(''); cur++; }
      for (let d = 1; d <= dIM; d++) {
        if (cur > 6) { row.push(wt > 0 ? wt : ''); wt = 0; cur = 0; }
        const v = data.dias[d] || 0; wt += v; row.push(v > 0 ? v : ''); cur++;
      }
      if (cur > 0) {
        let rem = 7 - cur;
        for (let i = 0; i < rem; i++) row.push('');
        row.push(wt > 0 ? wt : '');
      }
      const bm   = data.tipo === 'TECNICO 12/36H' ? 15 : 24;
      const ce   = meta * bm, cb = (meta - 1) * bm, cm2 = (meta - 2) * bm;
      row.push(data.total); row.push(ce); row.push(cb > 0 ? cb : 0); row.push(cm2 > 0 ? cm2 : 0);
      wsData.push(row);
    });

    const ws = XLSX.utils.aoa_to_sheet(wsData);
    ws['!merges'] = merges;

    let cols = [{ wch: mnl + 4 }];
    for (let i = 1; i < r0.length; i++) {
      if      (r0[i] === 'TOTAL')   cols.push({ wch: 9 });
      else if (['Excelente','Bom','Mediano'].includes(r1[i])) cols.push({ wch: 10 });
      else if (r1[i] === 'TOT')     cols.push({ wch: 5 });
      else                          cols.push({ wch: 4.5 });
    }
    ws['!cols'] = cols;
    ws['!rows'] = [{ hpt: 22 }, { hpt: 17 }, { hpt: 13 }];

    const range = XLSX.utils.decode_range(ws['!ref']);
    const tci   = r0.indexOf('TOTAL');

    for (let R = range.s.r; R <= range.e.r; ++R) {
      for (let C = range.s.c; C <= range.e.c; ++C) {
        const ca = XLSX.utils.encode_cell({ r: R, c: C });
        if (!ws[ca]) ws[ca] = { v: '', t: 's' };
        if (!ws[ca].s) ws[ca].s = {};

        const isTot = r1[C] === 'TOT';
        const isCap = ['Excelente','Bom','Mediano'].includes(r1[C]);

        if (R === 0) {
          const isSem = r0[C]?.startsWith('SEM');
          ws[ca].s = cs(isSem ? XL.navyLight : XL.navyMid, isSem ? XL.weekTxt : XL.navyTxt, true, 'center', 9);
        } else if (R === 1) {
          if      (r1[C] === 'Excelente') ws[ca].s = cs('FF00B050', 'FFFFFFFF', true, 'center', 10);
          else if (r1[C] === 'Bom')       ws[ca].s = cs('FF0000FF', 'FFFFFFFF', true, 'center', 10);
          else if (r1[C] === 'Mediano')   ws[ca].s = cs('FFCA8A04', 'FFFFFFFF', true, 'center', 10);
          else                            ws[ca].s = cs(XL.navyMid, XL.navyTxt, true, 'center', 9);
        } else if (R === 2) {
          ws[ca].s = cs(XL.navyMid, 'FF5A6478', false, 'center', 8);
        } else if (C === 0) {
          ws[ca].s = cs(XL.white, XL.navy, true, 'left', 10);
          ws[ca].s.font.sz = 10;
        } else if (C === tci && R > 2) {
          const vt  = Number(ws[ca].v || 0);
          const vce = Number(ws[XLSX.utils.encode_cell({ r: R, c: tci + 1 })].v || 0);
          const vcb = Number(ws[XLSX.utils.encode_cell({ r: R, c: tci + 2 })].v || 0);
          const vcm = Number(ws[XLSX.utils.encode_cell({ r: R, c: tci + 3 })].v || 0);
          let bg = XL.redBg, fg = XL.redFg;
          if (vt >= vce) { bg = XL.greenBg;  fg = XL.greenFg; }
          else if (vt >= vcb) { bg = XL.blueBg; fg = XL.blueFg; }
          else if (vt >= vcm) { bg = XL.yellowBg; fg = XL.yellowFg; }
          ws[ca].s = cs(bg, fg, true, 'center', 11);
        } else if (isCap) {
          ws[ca].s = cs(XL.white, 'FF9CA3AF', false, 'center', 10);
        } else if (isTot) {
          ws[ca].s = cs(XL.totBg, XL.totFg, true, 'center', 9);
        } else if (R > 2) {
          const v      = Number(ws[ca].v || 0);
          const rowIdx = R - 3;
          if (rowIdx < sorted.length) {
            const meta = state.appSettings.metasDiarias[sorted[rowIdx][1].tipo] || 5;
            let bg = XL.white, fg = 'FF9CA3AF', bold = false;
            if (v > 0) {
              if      (v >= meta)     { bg = XL.greenBg;  fg = XL.greenFg;  bold = true; }
              else if (v >= meta - 1) { bg = XL.blueBg;   fg = XL.blueFg;   bold = true; }
              else if (v >= meta - 2) { bg = XL.yellowBg; fg = XL.yellowFg; bold = true; }
              else                    { bg = XL.redBg;    fg = XL.redFg;    bold = true; }
            }
            ws[ca].s = cs(bg, fg, bold, 'center', 9);
          }
        }
      }
    }

    const rowH = ws['!rows'] || [];
    for (let R = 3; R <= range.e.r; R++) rowH.push({ hpt: 15 });
    ws['!rows'] = rowH;
    ws['!freeze'] = { xSplit: 1, ySplit: 3 };

    XLSX.utils.book_append_sheet(wb, ws, cidade.substring(0, 31));
  });

  XLSX.writeFile(wb, `SGO_Matriz_${state.activeMonthYear}.xlsx`);
}
