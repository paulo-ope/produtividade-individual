/**
 * SGO — team.js
 * CRUD de colaboradores, modal de cadastro, listagem por grupo.
 */

import { state, TEAM_TYPES } from './state.js';
import { saveTeam } from './storage.js';

export function limparNome(n) {
  if (!n) return '';
  return String(n)
    .split('-')[0].split('>')[0].split('(')[0].split('/')[0]
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ').trim().toUpperCase();
}

// ── Modal ──────────────────────────────────────────────────
export function openAddTechModal(name = '', city = '', tipo = 'INSTALAÇÃO CIDADE') {
  document.getElementById('modalTitle').textContent = 'Adicionar Colaborador';
  document.getElementById('modTechName').value = name;
  document.getElementById('modTechName').disabled = false;
  document.getElementById('modTechCity').value = city;
  document.getElementById('modTechType').value = tipo;
  document.getElementById('techModal').style.display = 'flex';
}

export function editTech(key) {
  const t = state.teamData[key];
  if (!t) return;
  document.getElementById('modalTitle').textContent = 'Editar Colaborador';
  document.getElementById('modTechName').value = t.originalName;
  document.getElementById('modTechName').disabled = true;
  document.getElementById('modTechCity').value = t.base;
  document.getElementById('modTechType').value = t.tipo || 'INSTALAÇÃO CIDADE';
  document.getElementById('techModal').style.display = 'flex';
}

export function closeTechModal() {
  document.getElementById('techModal').style.display = 'none';
}

export function saveTechForm(onSuccess) {
  const name = document.getElementById('modTechName').value.trim().toUpperCase();
  const base = document.getElementById('modTechCity').value.trim().toUpperCase() || 'BASE NÃO DEFINIDA';
  const tipo = document.getElementById('modTechType').value;
  if (!name) return alert('Nome obrigatório.');
  state.teamData[limparNome(name)] = { originalName: name, base, tipo };
  saveTeam();
  closeTechModal();
  if (onSuccess) onSuccess();
}

export function deleteTech(key, onSuccess) {
  if (!confirm('Remover colaborador?')) return;
  delete state.teamData[key];
  saveTeam();
  if (onSuccess) onSuccess();
}

export function acceptReclassification(key, tipo, onSuccess) {
  if (!state.teamData[key]) return;
  state.teamData[key].tipo = tipo;
  saveTeam();
  alert('Reclassificado: ' + (TEAM_TYPES[tipo] || tipo));
  if (onSuccess) onSuccess();
}

// ── Renderização ───────────────────────────────────────────
export function renderTeamTable() {
  const ft  = limparNome(document.getElementById('searchTech')?.value || '');
  const fc  = document.getElementById('filterTeamCity')?.value || 'ALL';
  const grp = {};

  Object.values(state.teamData).forEach(t => {
    if (ft && !limparNome(t.originalName).includes(ft)) return;
    if (fc !== 'ALL' && t.base !== fc) return;
    if (!grp[t.base]) grp[t.base] = [];
    grp[t.base].push(t);
  });

  const container = document.getElementById('teamListContainer');
  if (!container) return;

  if (!Object.keys(grp).length) {
    container.innerHTML = '<div class="empty">Nenhum colaborador encontrado.</div>';
    return;
  }

  container.innerHTML = Object.keys(grp).sort().map(base => {
    const members = grp[base].sort((a, b) => a.originalName.localeCompare(b.originalName));
    const rows = members.map(t => {
      const key = limparNome(t.originalName);
      return `<div class="t-row">
        <span class="t-name">${t.originalName}</span>
        <span class="t-type">${TEAM_TYPES[t.tipo] || t.tipo}</span>
        <div style="display:flex;gap:8px;">
          <button class="lbtn e" onclick="App.editTech('${key}')">Editar</button>
          <button class="lbtn d" onclick="App.deleteTech('${key}')">Remover</button>
        </div>
      </div>`;
    }).join('');
    return `<div class="team-grp">
      <div class="tg-hdr">
        <span class="tg-name">${base}</span>
        <span class="tg-cnt">${members.length} membros</span>
      </div>
      ${rows}
    </div>`;
  }).join('');
}

// ── Filtros de regionais ───────────────────────────────────
export function populateFilters() {
  const bases = [...new Set(Object.values(state.teamData).map(t => t.base))].sort();

  // City tabs (dashboard)
  const ct = document.getElementById('cityTabsContainer');
  if (ct) {
    ct.innerHTML = `<button class="city-tab active" data-city="ALL" onclick="App.selectCityTab('ALL')">Todas</button>`;
    bases.forEach(b => {
      ct.innerHTML += `<button class="city-tab" data-city="${b}" onclick="App.selectCityTab('${b}')">${b}</button>`;
    });
    if (state.selectedCityTab !== 'ALL' && !bases.includes(state.selectedCityTab)) {
      state.selectedCityTab = 'ALL';
    }
    document.querySelectorAll('.city-tab').forEach(t =>
      t.classList.toggle('active', t.dataset.city === state.selectedCityTab)
    );
  }

  // Filter select (config)
  const fc = document.getElementById('filterTeamCity');
  if (fc) {
    const prev = fc.value;
    fc.innerHTML = '<option value="ALL">Todas as Regionais</option>';
    bases.forEach(b => fc.innerHTML += `<option value="${b}">${b}</option>`);
    fc.value = bases.includes(prev) ? prev : 'ALL';
  }

  // Datalist
  const dl = document.getElementById('regionaisDatalist');
  if (dl) {
    dl.innerHTML = '';
    bases.forEach(b => dl.innerHTML += `<option value="${b}">`);
  }
}

export function selectCityTab(city) {
  state.selectedCityTab = city;
  document.querySelectorAll('.city-tab').forEach(t =>
    t.classList.toggle('active', t.dataset.city === city)
  );
}
