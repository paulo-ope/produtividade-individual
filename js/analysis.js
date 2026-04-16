/**
 * SGO — analysis.js
 * Análise operacional local e integração com Gemini API.
 */

import { state, DN, TEAM_TYPES } from './state.js';
import { loadGeminiKey, saveGeminiKey } from './storage.js';
import { switchTab } from './ui.js';

// ── Análise local ──────────────────────────────────────────
export function buildOperationalAnalysis(filtered) {
  if (!filtered.length) return;

  const techTotals = {}, techDias = {};
  filtered.forEach(i => {
    techTotals[i.techKey] = (techTotals[i.techKey] || 0) + 1;
    if (!techDias[i.techKey]) techDias[i.techKey] = new Set();
    techDias[i.techKey].add(i.day);
  });

  const sorted     = Object.entries(techTotals).sort((a, b) => b[1] - a[1]);
  const totalOS    = filtered.length;
  const totalTechs = sorted.length;
  const avgOS      = totalOS / totalTechs;
  const best       = sorted[0];
  const worst      = sorted[sorted.length - 1];

  const bestNome  = state.teamData[best[0]]?.originalName || best[0];
  const worstNome = state.teamData[worst[0]]?.originalName || worst[0];
  const bestMeta  = state.appSettings.metasDiarias[state.teamData[best[0]]?.tipo || 'INSTALAÇÃO CIDADE'] || 5;
  const bestCap   = best[1] / (bestMeta * 24) * 100;

  const [ys, ms] = state.activeMonthYear.split('-');
  const dIM = new Date(parseInt(ys), parseInt(ms), 0).getDate();

  const dayMap = {};
  filtered.forEach(i => { dayMap[i.day] = (dayMap[i.day] || 0) + 1; });
  const dailyArr = Array.from({ length: dIM }, (_, i) => dayMap[i + 1] || 0);
  const workDays = dailyArr.filter(v => v > 0);
  const weakIdx  = dailyArr.indexOf(Math.min(...workDays.filter(v => v > 0)));
  const weakDay  = weakIdx >= 0
    ? `Dia ${weakIdx + 1} (${DN[new Date(parseInt(ys), parseInt(ms) - 1, weakIdx + 1).getDay()]})`
    : '—';
  const weakVal = weakIdx >= 0 ? dailyArr[weakIdx] : 0;

  const abovePct = sorted.filter(([k, v]) => {
    const m = state.appSettings.metasDiarias[state.teamData[k]?.tipo || 'INSTALAÇÃO CIDADE'] || 5;
    return v / (m * 24) >= 0.8;
  });
  const effPct = Math.round(abovePct.length / totalTechs * 100);

  // KPI cards
  document.getElementById('aEfic').textContent    = effPct + '%';
  document.getElementById('aEficSub').textContent = `${abovePct.length} de ${totalTechs} técnicos acima de 80%`;
  document.getElementById('aBest').textContent    = bestNome;
  document.getElementById('aBestSub').textContent = `${best[1]} OS · ${Math.round(bestCap)}% cap.`;
  document.getElementById('aWorst').textContent   = worstNome;
  document.getElementById('aWorstSub').textContent= `${worst[1]} OS · ${Math.round(worst[1] / avgOS * 100)}% da média`;
  document.getElementById('aWeakDay').textContent = weakDay;
  document.getElementById('aWeakDaySub').textContent = `${weakVal} OS · menor produção`;

  // Positivos
  const topList = sorted.slice(0, 5).map(([k, v]) => {
    const m   = state.appSettings.metasDiarias[state.teamData[k]?.tipo || 'INSTALAÇÃO CIDADE'] || 5;
    const cap = Math.round(v / (m * 24) * 100);
    const nm  = state.teamData[k]?.originalName || k;
    return `<div class="obs-item"><div class="obs-dot g"></div><span><b>${nm}</b> — ${v} OS (${cap}% cap.)</span></div>`;
  }).join('');
  document.getElementById('obsPositivos').innerHTML = topList || '<div class="empty" style="padding:12px;">Sem dados.</div>';

  // Atenção
  const botList = sorted.filter(([k, v]) => v < avgOS * 0.6).slice(0, 5).map(([k, v]) => {
    const nm   = state.teamData[k]?.originalName || k;
    const diff = Math.round((1 - v / avgOS) * 100);
    return `<div class="obs-item"><div class="obs-dot r"></div><span><b>${nm}</b> — ${v} OS (${diff}% abaixo da média)</span></div>`;
  }).join('');
  document.getElementById('obsAtencao').innerHTML = botList ||
    '<div class="obs-item"><div class="obs-dot g"></div><span>Todos acima de 60% da média!</span></div>';

  // Rápidas
  const diasComOS  = workDays.length;
  const diasUteis  = dailyArr.filter((_, i) => {
    const dw = new Date(parseInt(ys), parseInt(ms) - 1, i + 1).getDay();
    return dw !== 0 && dw !== 6;
  }).length;
  const taxaPresenca  = Math.round(diasComOS / diasUteis * 100);
  const mediaDiaUtil  = Math.round(totalOS / diasComOS);

  const obs = [
    `Taxa de presença: <b>${taxaPresenca}%</b> dos dias úteis`,
    `Média de ${mediaDiaUtil} OS por dia produtivo`,
    sorted.length > 0 ? `Variação: <b>${best[1] - worst[1]} OS</b> entre melhor e pior` : null
  ].filter(Boolean).map(o =>
    `<div class="obs-item"><div class="obs-dot y"></div><span>${o}</span></div>`
  ).join('');
  document.getElementById('obsRapidas').innerHTML = obs;

  // Sugestões
  const sug = [];
  if (worst[1] < avgOS * 0.5) sug.push(`Investigar <b>${worstNome}</b> — abaixo de 50% da média`);
  if (effPct < 50)             sug.push('Revisar metas — menos de 50% atingiu 80% da capacidade');
  if (weakVal < avgOS * 0.3)   sug.push(`Analisar dia fraco (${weakDay})`);
  sug.push('Considerar redistribuição entre polos com desempenhos desiguais');

  document.getElementById('obsSugestoes').innerHTML = sug.slice(0, 3).map(s =>
    `<div class="obs-item"><div class="obs-dot y"></div><span>${s}</span></div>`
  ).join('');

  // Badges
  document.getElementById('badgeAnalise')?.classList.remove('hidden');
  document.getElementById('mobBadgeAnalise')?.classList.remove('hidden');
}

// ── IA Local — Classificação de técnicos ──────────────────
export function evaluateTechsByAI() {
  if (!Object.keys(state.globalTechStats).length) return;
  state.pendingTechs          = [];
  state.reclassifySuggestions = [];

  const cache = {};
  const matchTeam = nome => {
    if (cache[nome] !== undefined) return cache[nome];
    for (const k in state.teamData) {
      if (nome === k || nome.includes(k) || k.includes(nome)) { cache[nome] = k; return k; }
    }
    cache[nome] = null;
    return null;
  };

  Object.keys(state.globalTechStats).forEach(nome => {
    const s   = state.globalTechStats[nome];
    const pR  = s.rural / s.total;
    const da  = [...s.days].sort((a, b) => a - b);
    const dt  = da.length;
    let d2 = 0;
    for (let k = 1; k < da.length; k++) if (da[k] - da[k - 1] === 2) d2++;
    const isP = dt >= 4 && d2 / (dt - 1 || 1) >= 0.5;

    let sug = 'INSTALAÇÃO CIDADE';
    if (isP)           sug = 'TECNICO 12/36H';
    else if (pR >= 0.6) sug = 'RURAL';
    else if (s.total / dt > 7.5) sug = 'SUPORTE MOTO';

    const kc = matchTeam(nome);
    if (!kc && s.total > 1) {
      state.pendingTechs.push({ nome, total: s.total, suggestedType: sug });
    } else if (kc && s.total >= 5) {
      const ct = state.teamData[kc].tipo || 'INSTALAÇÃO CIDADE';
      if (ct !== sug) state.reclassifySuggestions.push({ key: kc, nome: state.teamData[kc].originalName, currentType: ct, suggestedType: sug });
    }
  });
}

// ── Render pendentes e reclassificação ─────────────────────
export function renderPendentes(onAddTech) {
  const badge   = (id, n) => {
    const el = document.getElementById(id);
    if (!el) return;
    if (n > 0) { el.textContent = n; el.classList.remove('hidden'); }
    else       { el.classList.add('hidden'); }
  };

  badge('badgePendentes',        state.pendingTechs.length);
  badge('mobBadgePendentes',     state.pendingTechs.length);
  badge('badgeReclassificacao',  state.reclassifySuggestions.length);
  badge('mobBadgeReclass',       state.reclassifySuggestions.length);

  // Tabela pendentes
  const tb = document.getElementById('pendentesTableBody');
  if (tb) {
    tb.innerHTML = state.pendingTechs.length
      ? state.pendingTechs.sort((a, b) => b.total - a.total).map(n => {
          const sn = n.nome.replace(/"/g, '&quot;').replace(/'/g, "\\'");
          return `<tr>
            <td class="td-nm">${n.nome}</td>
            <td class="td-ct">${n.total}</td>
            <td class="td-tp">${TEAM_TYPES[n.suggestedType] || 'Outros'}</td>
            <td class="td-r">
              <button class="btn btn-accent" style="font-size:10px;padding:6px 11px;"
                data-nome="${sn}" data-type="${n.suggestedType}"
                onclick="App.openAddTechModal(this.getAttribute('data-nome'),'',this.getAttribute('data-type'))">
                + Adicionar
              </button>
            </td>
          </tr>`;
        }).join('')
      : `<tr><td colspan="4" class="empty">Nenhum pendente.</td></tr>`;
  }

  // Tabela reclassificação
  const tr = document.getElementById('reclassifyTableBody');
  if (tr) {
    tr.innerHTML = state.reclassifySuggestions.length
      ? state.reclassifySuggestions.map(r =>
          `<tr>
            <td class="td-nm">${r.nome}</td>
            <td class="td-st">${TEAM_TYPES[r.currentType]}</td>
            <td class="td-pu">${TEAM_TYPES[r.suggestedType]}</td>
            <td class="td-r">
              <button class="btn btn-purple" style="font-size:10px;padding:6px 11px;"
                onclick="App.acceptReclassification('${r.key}','${r.suggestedType}')">
                ✓ Aceitar
              </button>
            </td>
          </tr>`
        ).join('')
      : `<tr><td colspan="4" class="empty">Classificações corretas.</td></tr>`;
  }
}

// ── Gemini ─────────────────────────────────────────────────
export function saveGeminiKeyUI() {
  const k = document.getElementById('geminiApiKey')?.value.trim();
  if (!k) return alert('Cole a chave de API.');
  saveGeminiKey(k);
  alert('✓ Chave salva. Volte à Matriz e clique em Analisar IA.');
}

function buildGeminiContext() {
  if (!state.globalRawData.length) return null;
  const fc = state.selectedCityTab;
  const ft = document.getElementById('filterType')?.value || 'ALL';
  const filtered = state.globalRawData.filter(i =>
    i.monthStr === state.activeMonthYear &&
    (fc === 'ALL' || i.cidade === fc) &&
    (ft === 'ALL' || i.tipo === ft)
  );
  if (!filtered.length) return null;

  const cs = {};
  filtered.forEach(i => {
    if (!cs[i.cidade]) cs[i.cidade] = { total: 0, techs: {} };
    cs[i.cidade].total++;
    if (!cs[i.cidade].techs[i.techKey]) cs[i.cidade].techs[i.techKey] = { total: 0, tipo: i.tipo, dias: new Set() };
    cs[i.cidade].techs[i.techKey].total++;
    cs[i.cidade].techs[i.techKey].dias.add(i.day);
  });

  const m = state.appSettings.metasDiarias;
  let ctx = `Período: ${state.activeMonthYear}\nTotal O.S.: ${filtered.length}\n`;
  ctx += `Metas: Inst.Cidade=${m["INSTALAÇÃO CIDADE"]}, Plantão=${m["TECNICO 12/36H"]}, Suporte=${m["SUPORTE MOTO"]}, Rural=${m["RURAL"]}, FazTudo=${m["FAZ TUDO"]}\n\n`;

  Object.keys(cs).sort().forEach(cidade => {
    const c  = cs[cidade];
    const tl = Object.entries(c.techs);
    ctx += `=== ${cidade} ===\nTotal: ${c.total} | Técnicos: ${tl.length} | Média: ${(c.total / tl.length).toFixed(1)}\n`;
    tl.sort((a, b) => b[1].total - a[1].total).forEach(([k, d]) => {
      const meta = state.appSettings.metasDiarias[d.tipo] || 5;
      const bm   = d.tipo === 'TECNICO 12/36H' ? 15 : 24;
      const ce   = meta * bm;
      const dt   = d.dias.size;
      ctx += `  • ${state.teamData[k]?.originalName || k} (${TEAM_TYPES[d.tipo] || d.tipo}): ${d.total} OS / ${dt} dias | média/dia: ${dt > 0 ? (d.total / dt).toFixed(1) : 0} | meta: ${meta}/dia | cap: ${Math.round(d.total / ce * 100)}%\n`;
    });
    ctx += '\n';
  });

  return ctx;
}

export async function analyzeWithGemini() {
  const ak = loadGeminiKey();
  if (!ak) { alert('Configure a API Key na aba Base de Equipes.'); switchTab('config'); return; }

  const ctx = buildGeminiContext();
  if (!ctx) { alert('Carregue uma planilha primeiro.'); return; }

  const btn    = document.getElementById('btnGeminiAnalyze');
  const panel  = document.getElementById('geminiPanel');
  const content= document.getElementById('geminiInsightContent');

  if (btn) { btn.disabled = true; btn.textContent = '⏳ Analisando...'; }
  panel.style.display = 'block';
  content.innerHTML = `<div style="display:flex;align-items:center;gap:10px;color:var(--purple);font-family:var(--mono);font-size:11px;"><div class="ld-ring" style="width:18px;height:18px;border-width:2px;border-top-color:var(--purple);"></div>Processando com Gemini 2.5 Flash...</div>`;
  panel.scrollIntoView({ behavior: 'smooth', block: 'start' });

  const prompt = `Você é um analista de produtividade de operações técnicas de provedor de internet. Analise os dados e gere um relatório objetivo. Responda em português do Brasil.\n\nRetorne apenas:\n\n1️⃣ Ranking de produtividade (Top 5 técnicos com OS e % capacidade)\n\n2️⃣ TOP 5 Técnicos abaixo da média (nome, OS, quanto abaixo)\n\n3️⃣ Observações rápidas (máx 3 pontos)\n\n4️⃣ Sugestões de melhoria (máx 5 ações concretas)\n\nUse **negrito** para nomes. Seja direto e objetivo.\n\nDADOS:\n${ctx}`;

  try {
    const res = await fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${ak}`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          contents: [{ parts: [{ text: prompt }] }],
          generationConfig: { temperature: 0.3, maxOutputTokens: 50000 }
        })
      }
    );

    if (!res.ok) {
      const e = await res.json();
      throw new Error(e?.error?.message || `HTTP ${res.status}`);
    }

    const data = await res.json();
    const text = data?.candidates?.[0]?.content?.parts?.[0]?.text;
    if (!text) throw new Error('Resposta vazia.');

    const fmt = text
      .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
      .replace(/\*(.*?)\*/g, '<em style="color:var(--text-2)">$1</em>')
      .split('\n\n').filter(p => p.trim())
      .map(p => {
        if (p.includes('\n• ') || p.includes('\n- ')) {
          return p.split('\n').map(l => {
            if (l.startsWith('• ') || l.startsWith('- '))
              return `<div class="gb"><span class="gb-dot">›</span><span>${l.slice(2)}</span></div>`;
            return `<p style="margin-bottom:3px;">${l}</p>`;
          }).join('');
        }
        return `<p style="margin-bottom:10px;">${p.replace(/\n/g, '<br>')}</p>`;
      }).join('');

    content.innerHTML = fmt;
    document.getElementById('geminiTimestamp').textContent = 'Gerado em ' + new Date().toLocaleString('pt-BR');

  } catch (err) {
    content.innerHTML = `<div style="color:var(--red);font-family:var(--mono);font-size:11px;"><strong style="display:block;margin-bottom:5px;">Erro ao chamar Gemini</strong>${err.message}</div>`;
  } finally {
    if (btn) {
      btn.disabled = false;
      btn.innerHTML = `<svg width="11" height="11" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9.663 17h4.673M12 3v1m6.364 1.636l-.707.707M21 12h-1M4 12H3m3.343-5.657l-.707-.707m2.828 9.9a5 5 0 117.072 0l-.548.547A3.374 3.374 0 0014 18.469V19a2 2 0 11-4 0v-.531c0-.895-.356-1.754-.988-2.386l-.548-.547z"/></svg> Analisar IA`;
    }
  }
}
