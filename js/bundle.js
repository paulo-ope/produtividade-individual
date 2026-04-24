/**
 * SGO — bundle.js
 * Todos os módulos concatenados — funciona com file:// e servidores.
 * Ordem: state → storage → ui → team → charts → matrix → analysis → upload → app
 */

/* ═══════════════════════════════════════════════════════════
   STATE — constantes e estado global
   ═══════════════════════════════════════════════════════════ */
const TEAM_TYPES = {
  "INSTALAÇÃO CIDADE": "Cidade",
  "TECNICO 12/36H":    "Plantão 12/36H",
  "SUPORTE MOTO":      "Suporte Moto",
  "SUPORTE CARRO":     "Suporte Carro",
  "RURAL":             "Rural",
  "FAZ TUDO":          "Faz Tudo",
  "AUXILIAR":          "Auxiliar"
};

const DEFAULT_SETTINGS = {
  metasDiarias: {
    "INSTALAÇÃO CIDADE": 5,
    "TECNICO 12/36H":    7,
    "SUPORTE MOTO":      7,
    "SUPORTE CARRO":     7,
    "RURAL":             5,
    "FAZ TUDO":          6
  },
  metaSabadoPct: 50
};

const DN = ['Dom','Seg','Ter','Qua','Qui','Sex','Sáb'];

const PAGE_TITLES = {
  dashboard:       'Dashboard',
  analise:         'Inteligência',
  recorrencia:     'Recorrência',
  pendentes:       'Novos Colaboradores',
  reclassificacao: 'Validações',
  capacidades:     'Performance Regional',
  moto:            'Redistribuição Moto',
  config:          'Equipes'
};

const SUPORTE_MOTO_MAP = {
  "Suporte Externo Fibra Urbana": new Set([
    "SUP: Atualização de configuração de equipamento ((Fibra))",
    "SUP: Troca de Roteador/ONU (Fibra)",
    "SUP: Problemas na rede interna do cliente (Fibra)",
    "SUP: Alteração de Senha/Wifi (Fibra)",
    "SUP: Fonte (Fibra)",
    "SUP: Equipamentos travados (Fibra)",
    "SUP: IPTV não licenciado (Fibra)",
    "SUP: Equipamento resetado (Fibra)"
  ]),
  "Remoção de Equipamentos": new Set([
    "Equipamentos: Removidos",
    "Equipamentos: Não Removidos",
    "Equipamentos: Parcialmente Removidos"
  ]),
  "Sem Conexão Fibra Urbana": new Set([
    "SUP: Troca de Roteador/ONU (Fibra)",
    "SUP: Atualização de configuração de equipamento ((Fibra))",
    "SUP: Problemas na rede interna do cliente (Fibra)",
    "SUP: Fonte (Fibra)",
    "SUP: Equipamentos travados (Fibra)",
    "SUP: Equipamento resetado (Fibra)",
    "SUP: Cabo Lan invertido (Fibra)"
  ]),
  "Remoção de Flashman": new Set([
    "Remoção de Flashman: Concluída",
    "Remoção de Flashman: Não concluída / Trocado equipamento",
    "Remoção de Flashman: Cliente ausente"
  ]),
  "Troca de Equipamentos": new Set([
    "SUP: Troca de Roteador/ONU (Fibra)",
    "Equipamento não Trocado sem Necessidade de Substituição"
  ]),
  "Suporte Streaming/Apps": new Set([
    "Apps: Canais em funcionamento",
    "Apps: TV / Aparelho não compatível",
    "Apps: Max em funcionamento",
    "Apps: Max + Canais em funcionamento"
  ]),
  "Suporte Prioritário": new Set([
    "SUP: Troca de Roteador/ONU (Fibra)",
    "SUP: Atualização de configuração de equipamento ((Fibra))"
  ]),
  "Suporte Externo Fibra Rural": new Set([
    "SUP: Atualização de configuração de equipamento ((Fibra))",
    "SUP: Troca de Roteador/ONU (Fibra)",
    "SUP: Conector Danificado (Fibra)",
    "SUP: IPTV não licenciado (Fibra)",
    "SUP: Problemas na rede interna do cliente (Fibra)"
  ]),
  "Ativação de Login Presencial": new Set([
    "Login ativado - Equipamento original mantido",
    "Login ativado - Com substituição de equipamento"
  ]),
  "Sem Conexão Fibra Rural": new Set([
    "SUP: Troca de Roteador/ONU (Fibra)",
    "SUP: Fonte (Fibra)",
    "SUP: Atualização de configuração de equipamento ((Fibra))"
  ]),
  "Recuperação de Equipamento por Cobrança": new Set([
    "Equipamentos: Removidos"
  ])
};

const state = {
  appSettings:           {},
  teamData:              {},
  rawExcelCache:         [],
  globalRawData:         [],
  globalRawDataByMonth:  {},
  globalTechStats:       {},
  uploadMeta:            { hasContrato: false, hasCliente: false, hasLogin: false, availableMonths: [] },
  recurrenceSettings:    { diasAnalise: 30, periodoBuscaDias: 90, minimoOsParaReincidencia: 2 },
  recurrenceAnalysis:    null,
  recurrenceAnalysisKey: '',
  pendingTechs:          [],
  reclassifySuggestions: [],
  activeMonthYear:       "",
  workerBlobUrl:         null,
  selectedCityTab:       "ALL",
  isCompactMode:         false,
  dailyChartInstance:    null,
  isDark:                false,
  showCap:               true,
  showRank:              true,
  motoOportunidades:     [],
  chartDOW:              [0,1,2,3,4,5,6],
  currentFiltered:       [],
  filterViewCache:       {},
  currentFilterMeta:     { cityItems: {}, techItems: {}, cityTmaStats: {}, techTmaStats: {} },
  currentDayDetails:     [],
  currentDayDetailsDay:  null
};

/* ═══════════════════════════════════════════════════════════
   STORAGE — localStorage e import/export
   ═══════════════════════════════════════════════════════════ */
const STORAGE_KEYS = {
  settings: 'sgo_settings_pro_v3',
  team:     'sgo_team_pro_v3',
  dark:     'sgo_dark'
};

const UPLOAD_DB = 'sgo_upload_cache';
const UPLOAD_STORE = 'files';
const UPLOAD_KEY = 'last_upload';

const TYPE_MIGRATION = {
  COMERCIAL:    'INSTALAÇÃO CIDADE',
  PLANTAO:      'TECNICO 12/36H',
  SUPORTE:      'SUPORTE MOTO',
  SUPORTE_CARRO:'SUPORTE CARRO',
  FAZ_TUDO:     'FAZ TUDO'
};

function migrateTypes(teamObj) {
  let changed = false;
  Object.values(teamObj).forEach(t => {
    const nt = TYPE_MIGRATION[t.tipo];
    if (nt) { t.tipo = nt; changed = true; }
  });
  return changed;
}

function initLocalStorage() {
  const rawSettings = localStorage.getItem(STORAGE_KEYS.settings);
  if (rawSettings) {
    state.appSettings = JSON.parse(rawSettings);
    if (state.appSettings.metaSabadoPct === undefined) state.appSettings.metaSabadoPct = 50;
  } else {
    const old = localStorage.getItem('sgo_settings_pro_v2');
    if (old) {
      const p = JSON.parse(old);
      state.appSettings = { metasDiarias: {
        "INSTALAÇÃO CIDADE": p.metasDiarias.COMERCIAL || 5,
        "TECNICO 12/36H":    p.metasDiarias.PLANTAO   || 4,
        "SUPORTE MOTO":      p.metasDiarias.SUPORTE    || 8,
        "RURAL":             p.metasDiarias.RURAL      || 3,
        "FAZ TUDO":          p.metasDiarias.FAZ_TUDO   || 5
      }};
    } else {
      state.appSettings = JSON.parse(JSON.stringify(DEFAULT_SETTINGS));
    }
  }
  const rawTeam = localStorage.getItem(STORAGE_KEYS.team) || localStorage.getItem('sgo_team_pro_v2');
  state.teamData = rawTeam ? JSON.parse(rawTeam) : {};
  if (migrateTypes(state.teamData)) saveTeam();
}

function saveSettings() { localStorage.setItem(STORAGE_KEYS.settings, JSON.stringify(state.appSettings)); }
function saveTeam()     { localStorage.setItem(STORAGE_KEYS.team,     JSON.stringify(state.teamData)); }
function saveDarkPref(v)  { localStorage.setItem(STORAGE_KEYS.dark, v ? '1' : '0'); }
function loadDarkPref()   { return localStorage.getItem(STORAGE_KEYS.dark) === '1'; }

function openUploadDb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(UPLOAD_DB, 1);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(UPLOAD_STORE)) db.createObjectStore(UPLOAD_STORE);
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function saveCachedUpload(payload) {
  const db = await openUploadDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(UPLOAD_STORE, 'readwrite');
    tx.objectStore(UPLOAD_STORE).put(payload, UPLOAD_KEY);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

async function loadCachedUpload() {
  const db = await openUploadDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(UPLOAD_STORE, 'readonly');
    const req = tx.objectStore(UPLOAD_STORE).get(UPLOAD_KEY);
    req.onsuccess = () => resolve(req.result || null);
    req.onerror = () => reject(req.error);
  });
}

function exportTeamData() {
  if (!Object.keys(state.teamData).length) return showToast('Sem equipes cadastradas.','warning');
  const a = document.createElement('a');
  a.href = 'data:text/json;charset=utf-8,' + encodeURIComponent(JSON.stringify(state.teamData, null, 2));
  a.download = 'equipes_sgo.json';
  a.click();
}

function importTeamData(event) {
  const file = event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = ev => {
    try {
      const raw = JSON.parse(ev.target.result);
      const TIPOS_VALIDOS = Object.keys(TEAM_TYPES);
      let imported = 0, warnings = [];
      const data = {};
      Object.entries(raw).forEach(([key, val]) => {
        if (!val || typeof val !== 'object') return;
        const nome = val.originalName || val.nome || key;
        const base = (val.base || 'BASE NÃO DEFINIDA').toUpperCase();
        let tipo = val.tipo || 'INSTALAÇÃO CIDADE';
        const migrated = TYPE_MIGRATION[tipo];
        if (migrated) tipo = migrated;
        if (!TIPOS_VALIDOS.includes(tipo)) {
          warnings.push(`"${nome}": tipo "${tipo}" desconhecido → definido como Cidade`);
          tipo = 'INSTALAÇÃO CIDADE';
        }
        data[key] = { originalName: nome, base, tipo };
        imported++;
      });
      migrateTypes(data);
      state.teamData = data;
      saveTeam();
      let msg = `✅ ${imported} colaboradores importados com sucesso!`;
      const bases = {};
      Object.values(data).forEach(t => { bases[t.base] = (bases[t.base]||0)+1; });
      msg += '\n\nPolos: ' + Object.entries(bases).sort((a,b)=>b[1]-a[1]).map(([b,n])=>`${b} (${n})`).join(', ');
      if (warnings.length) msg += `\n\n⚠️ ${warnings.length} aviso(s): ` + warnings.slice(0,3).join('; ');
      showToast(msg,'success',5000);
      syncEcosystem();
    } catch(e) { showToast('Arquivo JSON inválido: ' + e.message,'error'); }
  };
  reader.readAsText(file);
}

/* ═══════════════════════════════════════════════════════════
   UI — navegação, tema, relógio, sheet mobile
   ═══════════════════════════════════════════════════════════ */
function isMobile() { return window.innerWidth <= 768; }

function initClock() {
  const tick = () => {
    const el = document.getElementById('clockEl');
    if (el) el.textContent = new Date().toLocaleTimeString('pt-BR', { hour:'2-digit', minute:'2-digit', second:'2-digit' });
  };
  setInterval(tick, 1000); tick();
}

function applyDarkTheme(isDark) {
  document.documentElement.setAttribute('data-theme', isDark ? 'dark' : 'light');
  ['darkBtn','darkBtnMob'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.textContent = isDark ? '☀️' : '🌙';
  });
  const sbBtn = document.querySelector('.sb-theme-btn');
  if(sbBtn) sbBtn.textContent = isDark ? '☀️ Tema' : '🌙 Tema';
  const darkToggle = document.getElementById('cfgDarkToggle');
  if(darkToggle) darkToggle.checked = isDark;
}

function toggleDark(callback) {
  state.isDark = !state.isDark;
  applyDarkTheme(state.isDark);
  saveDarkPref(state.isDark);
  if (state.dailyChartInstance) { state.dailyChartInstance.destroy(); state.dailyChartInstance = null; }
  if (callback) callback();
}

function switchTab(tab) {
  ['dashboard','analise','recorrencia','pendentes','reclassificacao','capacidades','moto','config'].forEach(t => {
    document.getElementById('view-' + t)?.classList.remove('active');
    document.getElementById('tab-' + t)?.classList.remove('active');
    document.getElementById('mob-tab-' + t)?.classList.remove('active');
    document.getElementById('sb-tab-' + t)?.classList.remove('active');
  });
  document.getElementById('view-' + tab)?.classList.add('active');
  document.getElementById('tab-' + tab)?.classList.add('active');
  document.getElementById('mob-tab-' + tab)?.classList.add('active');
  document.getElementById('sb-tab-' + tab)?.classList.add('active');
  const title = PAGE_TITLES[tab] || tab;
  const pt = document.getElementById('mobPageTitle');
  if (pt) pt.textContent = title;
  const nt = document.getElementById('nbPageTitle');
  if (nt) nt.textContent = title;
  window.scrollTo({ top: 0, behavior: 'smooth' });
  if(tab === 'recorrencia') renderRecorrenciaClientes();
  if(tab === 'capacidades') { renderCapacidades(); syncCapacidadesContext(); }
  if(tab === 'moto') renderMotoOportunidades();
}

function toggleDotsMenuSyncDark() {
  setTimeout(() => syncDotsMenuState(), 50);
}

function syncSidebarState(open) {
  const menuBtn = document.querySelector('.nb-menu-btn');
  const sidebar = document.getElementById('sidebar');
  const overlay = document.getElementById('sidebarOverlay');
  if (menuBtn) menuBtn.setAttribute('aria-expanded', open ? 'true' : 'false');
  if (sidebar) sidebar.setAttribute('aria-hidden', open ? 'false' : 'true');
  if (overlay) overlay.setAttribute('aria-hidden', open ? 'false' : 'true');
}

function toggleSidebar() {
  const sb = document.getElementById('sidebar');
  const ov = document.getElementById('sidebarOverlay');
  if(!sb) return;
  const open = sb.classList.toggle('open');
  if(ov) ov.classList.toggle('open', open);
  syncSidebarState(open);
}
function closeSidebar() {
  document.getElementById('sidebar')?.classList.remove('open');
  document.getElementById('sidebarOverlay')?.classList.remove('open');
  syncSidebarState(false);
}
function toggleDotsMenu() {
  const m = document.getElementById('nbDotsMenu');
  if(!m) return;
  m.classList.toggle('open');
  if(m.classList.contains('open')){
    syncDotsMenuState();
    setTimeout(()=>document.addEventListener('click', closeDotsOutside, {once:true}), 10);
  }
}
function closeDotsMenu() {
  document.getElementById('nbDotsMenu')?.classList.remove('open');
}
function closeDotsOutside(e) {
  if(!e.target.closest('.nb-dots-wrap')) closeDotsMenu();
}

function syncDotsMenuState() {
  const dark = document.getElementById('cfgDarkToggle');
  if(dark) dark.checked = state.isDark;
  const compact = document.getElementById('cfgCompactToggle');
  if(compact) compact.checked = state.isCompactMode;
}

function toggleShowCap() {
  state.showCap = document.getElementById('cfgShowCap')?.checked !== false;
  applyCapVisibility();
  try { localStorage.setItem('sgo_show_cap', state.showCap ? '1' : '0'); } catch(e){}
}

function applyCapVisibility() {
  const show = state.showCap !== false;
  document.querySelectorAll('.ccap, .th-cap-group').forEach(el => {
    el.style.display = show ? '' : 'none';
  });
}

function toggleShowRank() {
  const show = document.getElementById('cfgShowRank')?.checked !== false;
  const ar = document.getElementById('analysisRow');
  if(ar) {
    const rankCard = ar.querySelector('.rank-card');
    if(rankCard) rankCard.style.display = show ? '' : 'none';
  }
  try { localStorage.setItem('sgo_show_rank', show ? '1' : '0'); } catch(e){}
}

function initNbClock() {
  const tick = () => {
    const el = document.getElementById('nbClockEl');
    if(el) el.textContent = new Date().toLocaleTimeString('pt-BR', {hour:'2-digit',minute:'2-digit',second:'2-digit'});
  };
  setInterval(tick, 1000); tick();
}

function openMobileFilterSheet() {
  const s = document.getElementById('mobileFilterSheet');
  s.style.display = 'block';
  setTimeout(() => s.classList.add('open'), 10);
}
function closeMobileFilterSheet() {
  const s = document.getElementById('mobileFilterSheet');
  s.classList.remove('open');
  setTimeout(() => s.style.display = 'none', 300);
}
function initMobileSheetSwipe() {
  const s = document.getElementById('mobileFilterSheet');
  if (!s) return;
  let startY = 0;
  s.addEventListener('touchstart', e => { startY = e.touches[0].clientY; }, { passive: true });
  s.addEventListener('touchmove',  e => { if (e.touches[0].clientY - startY > 80) closeMobileFilterSheet(); }, { passive: true });
}

function toggleCompactMode() {
  state.isCompactMode = !state.isCompactMode;
  document.getElementById('matrixWrapper')?.classList.toggle('compact', state.isCompactMode);
  const lbl = state.isCompactMode ? 'Padrão' : 'Compacto';
  const btn = document.getElementById('btnCompact'); if (btn) btn.textContent = lbl;
  const mbl = document.getElementById('mobCompactLbl'); if (mbl) mbl.textContent = lbl;
}


/* ═══════════════════════════════════════════════════════════
   TOAST + CONFIRM CUSTOMIZADOS
   ═══════════════════════════════════════════════════════════ */
function showToast(msg, type='success', duration=3500) {
  let container = document.getElementById('toastContainer');
  if (!container) {
    container = document.createElement('div');
    container.id = 'toastContainer';
    container.style.cssText = 'position:fixed;top:16px;right:16px;z-index:9999;display:flex;flex-direction:column;gap:8px;pointer-events:none;';
    document.body.appendChild(container);
  }
  const icons = { success:'✅', error:'❌', warning:'⚠️', info:'ℹ️' };
  const colors = { success:'#16A34A','error':'#DC2626',warning:'#D97706',info:'#2563EB' };
  const t = document.createElement('div');
  t.style.cssText = `
    background:#1E2D45;color:#E8F4FF;
    border:1px solid ${colors[type]}44;border-left:3px solid ${colors[type]};
    border-radius:10px;padding:11px 16px;min-width:240px;max-width:360px;
    font-family:var(--font);font-size:13px;font-weight:500;line-height:1.4;
    box-shadow:0 8px 32px rgba(0,0,0,.4);
    display:flex;align-items:flex-start;gap:10px;pointer-events:all;
    animation:toastIn .2s ease;
  `;
  t.innerHTML = `<span style="font-size:15px;flex-shrink:0;margin-top:1px;">${icons[type]}</span><span>${msg}</span>`;
  container.appendChild(t);
  setTimeout(() => {
    t.style.animation = 'toastOut .2s ease forwards';
    setTimeout(() => t.remove(), 200);
  }, duration);
}

function showConfirm(msg, onConfirm) {
  let overlay = document.getElementById('confirmOverlay');
  if (overlay) overlay.remove();
  overlay = document.createElement('div');
  overlay.id = 'confirmOverlay';
  overlay.style.cssText = 'position:fixed;inset:0;z-index:9998;background:rgba(0,0,0,.55);display:flex;align-items:center;justify-content:center;backdrop-filter:blur(4px);animation:fadeIn .15s ease;';
  overlay.innerHTML = `
    <div style="background:#1E2D45;border:1px solid #2A3D58;border-radius:16px;padding:24px 28px;max-width:360px;width:90%;box-shadow:0 16px 48px rgba(0,0,0,.5);animation:scaleIn .15s ease;">
      <div style="font-size:15px;font-weight:600;color:#E8F4FF;margin-bottom:8px;">Confirmar</div>
      <div style="font-size:13px;color:#94A3B8;margin-bottom:20px;line-height:1.5;">${msg}</div>
      <div style="display:flex;gap:10px;justify-content:flex-end;">
        <button id="confirmNo" style="padding:8px 18px;border-radius:8px;border:1px solid #2A3D58;background:transparent;color:#94A3B8;font-size:13px;font-weight:600;cursor:pointer;">Cancelar</button>
        <button id="confirmYes" style="padding:8px 18px;border-radius:8px;border:none;background:#DC2626;color:#fff;font-size:13px;font-weight:700;cursor:pointer;">Confirmar</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  document.getElementById('confirmNo').onclick  = () => overlay.remove();
  document.getElementById('confirmYes').onclick = () => { overlay.remove(); onConfirm(); };
  overlay.onclick = e => { if (e.target === overlay) overlay.remove(); };
}

function showLoading(v) { document.getElementById('loadingOverlay').style.display = v ? 'flex' : 'none'; }
function setStatus(text) { const s = document.getElementById('statusText'); if (s) s.textContent = text; }

function syncMobileFilters() {
  const ft = document.getElementById('filterType');
  const fm = document.getElementById('filterTypeMob');
  if (fm && ft) ft.value = fm.value;
}

/* ═══════════════════════════════════════════════════════════
   TEAM — CRUD de colaboradores
   ═══════════════════════════════════════════════════════════ */
function limparNome(n) {
  if (!n) return '';
  return String(n).split('-')[0].split('>')[0].split('(')[0].split('/')[0]
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, ' ').trim().toUpperCase();
}

function openAddTechModal(name, city, tipo) {
  name = name || ''; city = city || ''; tipo = tipo || 'INSTALAÇÃO CIDADE';
  document.getElementById('modalTitle').textContent = 'Adicionar Colaborador';
  document.getElementById('modTechName').value = name;
  document.getElementById('modTechName').disabled = false;
  document.getElementById('modTechCity').value = city;
  document.getElementById('modTechType').value = tipo;
  document.getElementById('techModal').style.display = 'flex';
}

function editTech(key) {
  const t = state.teamData[key]; if (!t) return;
  document.getElementById('modalTitle').textContent = 'Editar Colaborador';
  document.getElementById('modTechName').value = t.originalName;
  document.getElementById('modTechName').disabled = true;
  document.getElementById('modTechCity').value = t.base;
  document.getElementById('modTechType').value = t.tipo || 'INSTALAÇÃO CIDADE';
  document.getElementById('techModal').style.display = 'flex';
}

function closeTechModal() { document.getElementById('techModal').style.display = 'none'; }

function saveTechForm() {
  const name = document.getElementById('modTechName').value.trim().toUpperCase();
  const base = document.getElementById('modTechCity').value.trim().toUpperCase() || 'BASE NÃO DEFINIDA';
  const tipo = document.getElementById('modTechType').value;
  if (!name) return showToast('Nome obrigatório.','warning');
  state.teamData[limparNome(name)] = { originalName: name, base, tipo };
  saveTeam(); closeTechModal(); syncEcosystem();
}

function deleteTech(key) {
  showConfirm('Remover o colaborador permanentemente?', () => { delete state.teamData[key]; saveTeam(); syncEcosystem(); showToast('Colaborador removido.','info'); });
}

function acceptReclassification(key, tipo) {
  if (!state.teamData[key]) return;
  state.teamData[key].tipo = tipo; saveTeam();
  showToast('✓ Reclassificado para: <b>' + (TEAM_TYPES[tipo] || tipo) + '</b>','success');
  syncEcosystem();
}

function renderTeamTable() {
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
    container.innerHTML = '<div class="empty" style="padding:40px;text-align:center;">Nenhum colaborador encontrado.</div>';
    return;
  }
  const TIPO_COLORS = {
    'INSTALAÇÃO CIDADE':'#2563EB','TECNICO 12/36H':'#6D28D9',
    'SUPORTE MOTO':'#059669','SUPORTE CARRO':'#0891B2',
    'RURAL':'#92400E','FAZ TUDO':'#9333EA','AUXILIAR':'#475569'
  };
  const techStats = {};
  const monthRows = state.globalRawDataByMonth?.[state.activeMonthYear] || [];
  monthRows.forEach(i=>{
    techStats[i.techKey]=(techStats[i.techKey]||0)+1;
  });
  container.innerHTML = Object.keys(grp).sort().map(base => {
    const members = grp[base].sort((a,b) => a.originalName.localeCompare(b.originalName));
    const tipoCount = {};
    members.forEach(t => { tipoCount[t.tipo]=(tipoCount[t.tipo]||0)+1; });
    const tipoTags = Object.entries(tipoCount).map(([tipo,n]) => {
      const lbl = TEAM_TYPES[tipo]||tipo;
      return `<span class="tg-tipo-tag" style="background:rgba(255,255,255,0.06);color:#7A92AA;border-color:rgba(255,255,255,0.10);">${lbl} <b>${n}</b></span>`;
    }).join('');
    return `<div class="team-grp">
      <div class="tg-hdr">
        <div>
          <span class="tg-name">${base}</span>
          <div class="tg-tipo-tags">${tipoTags}</div>
        </div>
        <span class="tg-cnt">${members.length} membros</span>
      </div>
      <div class="t-list">
        ${members.map(t => {
          const key = limparNome(t.originalName).replace(/'/g,"\'");
          const lbl = TEAM_TYPES[t.tipo]||t.tipo;
          const os  = techStats[limparNome(t.originalName)] || 0;
          const isAux = t.tipo==='AUXILIAR';
          const initials = t.originalName.split(' ').filter(w=>w.length>2).slice(0,2).map(w=>w[0]).join('');
          return `<div class="t-row">
            <div class="t-avatar" style="background:rgba(255,255,255,0.06);color:#7A92AA;border:1px solid rgba(255,255,255,0.10);">${initials}</div>
            <div class="t-info">
              <div class="t-name">${t.originalName}</div>
              <span class="t-badge" style="background:rgba(255,255,255,0.05);color:#7A92AA;border-color:rgba(255,255,255,0.10);">${lbl}${isAux?' · sem meta':''}</span>
            </div>
            ${os > 0 ? `<div class="t-os-count">${os}<span>OS</span></div>` : '<div class="t-os-empty">—</div>'}
            <div class="t-actions">
              <button class="lbtn e" onclick="editTech('${key}')" title="Editar">
                <svg width="13" height="13" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z"/></svg>
              </button>
              <button class="lbtn d" onclick="deleteTech('${key}')" title="Remover">
                <svg width="13" height="13" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"/></svg>
              </button>
            </div>
          </div>`;
        }).join('')}
      </div>
    </div>`;
  }).join('');
}
function populateFilters() {
  const bases = [...new Set(Object.values(state.teamData).map(t => t.base))].sort();
  const ct = document.getElementById('cityTabsContainer');
  if (ct) {
    ct.innerHTML = `<button class="city-tab active" data-city="ALL" onclick="selectCityTab('ALL')">Todas</button>`;
    bases.forEach(b => ct.innerHTML += `<button class="city-tab" data-city="${b}" onclick="selectCityTab('${b}')">${b}</button>`);
    if (state.selectedCityTab !== 'ALL' && !bases.includes(state.selectedCityTab)) state.selectedCityTab = 'ALL';
    document.querySelectorAll('.city-tab').forEach(t => t.classList.toggle('active', t.dataset.city === state.selectedCityTab));
  }
  const fc = document.getElementById('filterTeamCity');
  if (fc) {
    const prev = fc.value;
    fc.innerHTML = '<option value="ALL">Todas as Regionais</option>';
    bases.forEach(b => fc.innerHTML += `<option value="${b}">${b}</option>`);
    fc.value = bases.includes(prev) ? prev : 'ALL';
  }
  const dl = document.getElementById('regionaisDatalist');
  if (dl) { dl.innerHTML = ''; bases.forEach(b => dl.innerHTML += `<option value="${b}">`); }
}

function selectCityTab(city) {
  state.selectedCityTab = city;
  document.querySelectorAll('.city-tab').forEach(t => t.classList.toggle('active', t.dataset.city === city));
  applyFilters();
}

/* ═══════════════════════════════════════════════════════════
   CHARTS — sparklines, gráfico diário, ranking, KPIs
   ═══════════════════════════════════════════════════════════ */
function drawSparkline(id, data, color) {
  const c = document.getElementById(id); if (!c) return;
  const ctx = c.getContext('2d');
  const W = c.offsetWidth || 60, H = c.offsetHeight || 26;
  c.width = W * 2; c.height = H * 2; ctx.scale(2, 2);
  ctx.clearRect(0, 0, W, H);
  if (!data || data.length < 2) return;
  const mn = Math.min(...data), mx = Math.max(...data), rng = mx - mn || 1, pad = 3;
  const pts = data.map((v, i) => ({ x: pad + (i/(data.length-1))*(W-pad*2), y: H-pad-(v-mn)/rng*(H-pad*2) }));
  const grad = ctx.createLinearGradient(0,0,0,H);
  grad.addColorStop(0, color+'44'); grad.addColorStop(1, color+'00');
  ctx.beginPath(); ctx.moveTo(pts[0].x, pts[0].y);
  pts.slice(1).forEach(p => ctx.lineTo(p.x, p.y));
  ctx.lineTo(pts[pts.length-1].x, H); ctx.lineTo(pts[0].x, H); ctx.closePath();
  ctx.fillStyle = grad; ctx.fill();
  ctx.beginPath(); ctx.moveTo(pts[0].x, pts[0].y);
  pts.slice(1).forEach(p => ctx.lineTo(p.x, p.y));
  ctx.strokeStyle = color; ctx.lineWidth = 1.5; ctx.stroke();
}

function renderAdvancedDailyChart() {
  const canvas = document.getElementById('dailyChart');
  const dowContainer = document.getElementById('chartDowFilters');
  const kpiContainer = document.getElementById('chartKpis');
  if (!canvas || !state.currentFiltered) return;

  const filtered = state.currentFiltered;
  const [ys,ms] = state.activeMonthYear.split('-');
  if (!ys || !ms) return;
  const year = parseInt(ys), month = parseInt(ms)-1;
  const dIM = new Date(year, month+1, 0).getDate();

  const dayMap = {};
  const activeTechKeys = new Set();
  filtered.forEach(i => {
     dayMap[i.day] = (dayMap[i.day] || 0) + 1;
     activeTechKeys.add(i.techKey);
  });

  const capMedianaArr = [];
  const capBoaArr = [];
  const capExcelenteArr = [];
  for(let d=1; d<=dIM; d++) {
     let capMedDia = 0;
     let capBoaDia = 0;
     let capExcDia = 0;
     const dow = new Date(year, month, d).getDay();
     let equipes1236 = 0;
     let meta1236 = 0;
     activeTechKeys.forEach(tk => {
        const td = state.teamData[tk];
        if(!td || td.tipo === 'AUXILIAR') return;
        const baseMeta = state.appSettings.metasDiarias[td.tipo] || 5;
        if(td.tipo === 'TECNICO 12/36H') {
          equipes1236++;
          meta1236 = baseMeta;
          return;
        }
        const metaDia = getMetaForDay(baseMeta, dow, td.tipo);
        if(metaDia > 0) {
          capExcDia += metaDia;
          capBoaDia += Math.max(0, metaDia - 1);
          capMedDia += Math.max(0, metaDia - 2);
        }
     });
     if(equipes1236 > 0) {
       const equipesAtivas1236 = Math.round(equipes1236 / 2);
       capExcDia += equipesAtivas1236 * meta1236;
       capBoaDia += equipesAtivas1236 * Math.max(0, meta1236 - 1);
       capMedDia += equipesAtivas1236 * Math.max(0, meta1236 - 2);
     }
     capMedianaArr.push(capMedDia);
     capBoaArr.push(capBoaDia);
     capExcelenteArr.push(capExcDia);
  }

  const labels = [];
  const dataOs = [];
  const dataCap = [];
  const colors = [];
  let totalOsView = 0;
  let peakOs = 0; let peakDay = 0; let peakDow = '';
  let sumCap = 0;

  const dark = state.isDark;

  for(let d=1; d<=dIM; d++) {
     const dow = new Date(year, month, d).getDay();
     if(!state.chartDOW.includes(dow)) continue;

     labels.push([String(d), DN[dow]]);
     const v = dayMap[d] || 0;
     const cMed = capMedianaArr[d-1];
     const cBoa = capBoaArr[d-1];
     const cExc = capExcelenteArr[d-1];

     dataOs.push(v);
     dataCap.push(cMed);
     totalOsView += v;
     sumCap += cMed;
     if(v > peakOs) { peakOs = v; peakDay = d; peakDow = DN[dow]; }

     let bgCol;
     if (v === 0) bgCol = dark ? 'rgba(48,54,61,.4)' : 'rgba(228,232,239,.6)';
     else if (v >= cExc) bgCol = 'rgba(29,111,232,.9)';
     else if (v >= cBoa) bgCol = 'rgba(22,163,74,.85)';
     else if (v >= cMed) bgCol = 'rgba(202,138,4,.85)';
     else bgCol = 'rgba(220,38,38,.85)';
     colors.push(bgCol);
  }

  if(dowContainer) {
     const diasNomes = ['Dom','Seg','Ter','Qua','Qui','Sex','Sáb'];
     dowContainer.innerHTML = diasNomes.map((n, i) => {
        const isActive = state.chartDOW.includes(i);
        const bg = isActive ? 'var(--accent)' : 'transparent';
        const col = isActive ? '#fff' : 'var(--text-tertiary)';
        const bd = isActive ? 'var(--accent)' : 'var(--surface-border)';
        return `<button onclick="App.toggleChartDOW(${i})" style="background:${bg}; color:${col}; border:1px solid ${bd}; border-radius:6px; font-size:10px; padding:3px 10px; cursor:pointer; font-family:var(--font); font-weight:700; transition:all 0.2s;">${n}</button>`;
     }).join('');
  }

  if(kpiContainer) {
     const avgOs = labels.length ? Math.round(totalOsView / labels.length) : 0;
     const avgCap = labels.length ? Math.round(sumCap / labels.length) : 0;
     kpiContainer.innerHTML = `
       <div style="display:flex; flex-direction:column; padding-right:16px; border-right:1px solid var(--surface-border);">
         <span style="font-size:18px; font-weight:900; color:var(--text-primary); line-height:1;">${avgOs} <span style="font-size:11px;color:var(--text-tertiary);font-weight:700;">/ ${avgCap}</span></span>
         <span style="font-size:9px; color:var(--text-tertiary); text-transform:uppercase; font-family:var(--mono); margin-top:4px;">Média O.S. vs Cap.</span>
       </div>
       <div style="display:flex; flex-direction:column; padding-right:16px; border-right:1px solid var(--surface-border);">
       <span style="font-size:18px; font-weight:900; color:var(--text-primary); line-height:1;">${peakOs} <span style="font-size:11px;color:var(--text-tertiary);font-weight:700;">(Dia ${peakDay} - ${peakDow})</span></span>
         <span style="font-size:9px; color:var(--text-tertiary); text-transform:uppercase; font-family:var(--mono); margin-top:4px;">Pico de Produção</span>
       </div>
       <div style="display:flex; flex-direction:column;">
         <span style="font-size:18px; font-weight:900; color:var(--accent); line-height:1;">${totalOsView}</span>
         <span style="font-size:9px; color:var(--text-tertiary); text-transform:uppercase; font-family:var(--mono); margin-top:4px;">Total no Filtro</span>
       </div>
     `;
  }

  if (state.dailyChartInstance) { state.dailyChartInstance.destroy(); state.dailyChartInstance = null; }

  const capBorder = dark ? 'rgba(148, 163, 184, 0.5)' : 'rgba(100, 116, 139, 0.4)';
  const capBg = dark ? 'rgba(148, 163, 184, 0.15)' : 'rgba(100, 116, 139, 0.1)';

  state.dailyChartInstance = new Chart(canvas, {
    data: {
      labels: labels,
      datasets:[
        { type: 'line', label: 'Meta de Capacidade', data: dataCap, borderColor: capBorder, backgroundColor: capBg, borderWidth: 2, borderDash: [], stepped: 'middle', pointRadius: 0, pointHoverRadius: 4, fill: true },
        { type: 'bar', label: 'O.S. Entregues', data: dataOs, backgroundColor: colors, borderRadius: 6, borderSkipped: false, barPercentage: 0.7, categoryPercentage: 0.85 }
      ]
    },
    options: { responsive: true, maintainAspectRatio: false, interaction: { mode: 'index', intersect: false }, plugins: { legend:{display:false}, tooltip:{ bodyFont:{family:'Geist Mono', size:11}, titleFont:{family:'Geist Mono',size:11}, callbacks: { title: ctx => 'Dia ' + String(ctx[0].label).replace(',', ' (') + ')', label: ctx => `${ctx.dataset.label}: ${ctx.raw}` } } }, scales: { x: { ticks:{font:{family:'Geist Mono',size:isMobile()?8:9},color:dark?'#7CA8CC':'#9CA3AF',maxRotation:0}, grid:{display:false}, border:{display:false} }, y: { ticks:{font:{family:'Geist Mono',size:isMobile()?8:9},color:dark?'#7CA8CC':'#9CA3AF',maxTicksLimit:5}, grid:{color:dark?'rgba(48,54,61,.5)':'rgba(228,232,239,.8)'}, border:{display:false} } } }
  });
}

function buildRanking(filtered) {
  const totals = {};
  filtered.forEach(i => { totals[i.techKey] = (totals[i.techKey]||0)+1; });
  const ranked = Object.entries(totals).map(([tk,total]) => {
    const td   = state.teamData[tk];
    if(!td || td.tipo==='AUXILIAR') return null;
    const meta = state.appSettings.metasDiarias[td.tipo]||5;
    const bm   = td.tipo==='TECNICO 12/36H'?15:24;
    const capExc = meta * bm;
    const pctCap = capExc>0 ? Math.round(total/capExc*100) : 0;
    return {tk, total, pctCap, capExc};
  }).filter(Boolean).sort((a,b)=>b.pctCap-a.pctCap).slice(0,8);

  if (!ranked.length) return;
  const maxPct = ranked[0].pctCap || 1;
  const medals = ['🥇','🥈','🥉'], posClass = ['gold','silver','bronze'];
  const rb = document.getElementById('rankBody'); if (!rb) return;

  rb.innerHTML = ranked.map(({tk,total,pctCap,capExc},i) => {
    const nome = state.teamData[tk]?.originalName || tk;
    const base = state.teamData[tk]?.base || '';
    const tipo = TEAM_TYPES[state.teamData[tk]?.tipo]||'';
    const barW = Math.round(pctCap/maxPct*100);
    const col  = pctCap>=100?'#60A5FA':pctCap>=80?'#4ADE80':pctCap>=60?'#FCD34D':'#F87171';
    const pos  = i<3 ? `<span class="rank-pos ${posClass[i]}">${medals[i]}</span>` : `<span class="rank-pos">${i+1}</span>`;
    return `<div class="rank-item">
      ${pos}
      <div style="flex:1;min-width:0;">
        <div class="rank-name">${nome}</div>
        <div class="rank-base">${base} · ${tipo}</div>
      </div>
      <div class="rank-bar-wrap">
        <div class="rank-bar-fill" style="width:${barW}%;background:${col};"></div>
      </div>
      <div class="rank-val" style="color:${col};">${pctCap}%<span style="display:block;font-size:8px;color:#4A7A9B;text-align:right;">${total} OS</span></div>
    </div>`;
  }).join('');
  const rt = document.getElementById('rankTotal'); if (rt) rt.textContent = `${ranked.length} técnicos`;
}

function updateDashboardStats(filtered) {
  state.currentFiltered = filtered;
  const filterMeta = state.currentFilterMeta || {};
  const grid = document.getElementById('kpiGrid'); if (grid) grid.style.display = 'grid';
  const tot = filtered.length;
  document.getElementById('dashTotalOs').textContent = tot || '—';
  if (!tot) { ['dashActiveTechs','dashAvgTech','dashTopBase','dashCriticalBase'].forEach(id=>document.getElementById(id).textContent='—'); return; }
  const techTotals = filterMeta.techTotals || {};
  const techKeys = Object.keys(techTotals);
  const activeTechCount = techKeys.length;
  document.getElementById('dashActiveTechs').textContent = activeTechCount;
  document.getElementById('dashAvgTech').textContent = Math.round(tot / Math.max(activeTechCount, 1));
  const [ys,ms] = state.activeMonthYear.split('-');
  const dIM = new Date(parseInt(ys), parseInt(ms), 0).getDate();
  const dayMap = filterMeta.dayMap || {};
  const dailyArr = Array.from({length:dIM},(_,i)=>dayMap[i+1]||0);
  setTimeout(()=>{
    drawSparkline('spark0', dailyArr, '#E07B1F');
    const tArr = Object.values(techTotals).sort((a,b)=>a-b);
    drawSparkline('spark1', tArr, '#1D6FE8');
    drawSparkline('spark2', dailyArr.map(v=>v>0?Math.round(v/Math.max(activeTechCount, 1)):0), '#8B5CF6');
    renderAdvancedDailyChart();
    buildRanking(filtered);
  }, 100);
  const cityMap = filterMeta.cityMap || {};
  const arr = Object.keys(cityMap).map(b => {
    const techs = cityMap[b];
    const totals = Object.values(techs);
    const totalOs = totals.reduce((sum, tech) => sum + tech.total, 0);
    return { nome: b, media: totalOs / Math.max(totals.length, 1) };
  }).sort((a,b)=>b.media-a.media);
  if (arr.length) {
    document.getElementById('dashTopBase').textContent = arr[0].nome;
    const tb = document.getElementById('dashTopBadge');
    if(tb){tb.style.display='inline-flex';tb.textContent='▲ '+Math.round(arr[0].media);}
    const critCard = document.getElementById('dashCriticalBase');
    const cb2 = document.getElementById('dashCritBadge');
    const critKpi = critCard?.closest('.kpi');
    if(arr.length > 1){
      if(critCard) critCard.textContent = arr[arr.length-1].nome;
      if(cb2){cb2.style.display='inline-flex';cb2.textContent='▼ '+Math.round(arr[arr.length-1].media);}
      if(critKpi) critKpi.style.display='';
    } else {
      if(critKpi) critKpi.style.display='none';
    }
  }
  const ar = document.getElementById('analysisRow'); if (ar) ar.style.display = 'grid';
  const cl = document.getElementById('chartMonthLabel'); if (cl) cl.textContent = state.activeMonthYear;
}

/* ═══════════════════════════════════════════════════════════
   MATRIX — geração da tabela na interface
   ═══════════════════════════════════════════════════════════ */
function getMetaForDay(baseMeta, dayOfWeek, tipo) {
  if (dayOfWeek === 0) {
    if (tipo === 'TECNICO 12/36H') return baseMeta;
    return 0;
  }
  if (dayOfWeek === 6) {
    if (tipo === 'TECNICO 12/36H') return baseMeta;
    const pct = (state.appSettings.metaSabadoPct || 50) / 100;
    return Math.max(1, Math.round(baseMeta * pct));
  }
  return baseMeta;
}
function vC(v, m, dayOfWeek, tipo) {
  if (!v || v===0) return 'vd';
  const mEff = (dayOfWeek !== undefined) ? getMetaForDay(m, dayOfWeek, tipo) : m;
  if (mEff === 0) return v > 0 ? 'vg' : 'vd';
  if (v>=mEff)     return 'vg';
  if (v>=mEff-1)   return 'vb';
  if (v>=mEff-2)   return 'vy';
  return 'vr';
}
function tC(v, ce, cb, cm) {
  if (v>=ce) return 'tg';
  if (v>=cb) return 'tb';
  if (v>=cm) return 'ty';
  return 'tr';
}

function formatMinutesCompact(minutes) {
  if (minutes == null || isNaN(minutes)) return '--';
  if (minutes < 60) return `${Math.round(minutes)}m`;
  const hours = Math.floor(minutes / 60);
  const mins = Math.round(minutes % 60);
  return mins ? `${hours}h ${mins}m` : `${hours}h`;
}

function getOsTmaMinutes(os) {
  const dIni = toDateTime(os?.dtInicio);
  const dFin = toDateTime(os?.dtFinal);
  if (!dIni || !dFin || dFin < dIni) return null;
  const tmaMin = (dFin - dIni) / 60000;
  return tmaMin < 1440 ? tmaMin : null;
}

function getNumericStats(values) {
  const valid = (values || []).filter(v => v != null && !isNaN(v));
  if (!valid.length) {
    return {
      count: 0,
      totalMinutes: 0,
      avgMinutes: null,
      medianMinutes: null,
      minMinutes: null,
      maxMinutes: null,
      p90Minutes: null,
      above120Count: 0,
      above120Pct: 0
    };
  }

  const sorted = [...valid].sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  const totalMinutes = valid.reduce((sum, value) => sum + value, 0);
  const medianMinutes = sorted.length % 2 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
  const p90Index = Math.min(sorted.length - 1, Math.ceil(sorted.length * 0.9) - 1);
  const above120Count = valid.filter(v => v >= 120).length;

  return {
    count: valid.length,
    totalMinutes,
    avgMinutes: totalMinutes / valid.length,
    medianMinutes,
    minMinutes: sorted[0],
    maxMinutes: sorted[sorted.length - 1],
    p90Minutes: sorted[p90Index],
    above120Count,
    above120Pct: Math.round((above120Count / valid.length) * 100)
  };
}

function getTmaStats(items) {
  return getNumericStats((items || []).map(getOsTmaMinutes));
}

function getTeamDayTmaStats(items) {
  const buckets = {};
  let totalMinutes = 0;
  (items || []).forEach(item => {
    const tmaMin = getOsTmaMinutes(item);
    if (tmaMin == null || isNaN(tmaMin)) return;
    totalMinutes += tmaMin;
    const dateKey = item.dateStr || `${item.monthStr || ''}-${String(item.day || '').padStart(2, '0')}`;
    const teamKey = item.techKey || item.nomeOriginal || item.nome || 'SEM_EQUIPE';
    const bucketKey = `${dateKey}::${teamKey}`;
    if (!buckets[bucketKey]) buckets[bucketKey] = { total: 0, count: 0 };
    buckets[bucketKey].total += tmaMin;
    buckets[bucketKey].count++;
  });

  const teamDayAverages = Object.values(buckets).map(bucket => bucket.total / bucket.count);
  return {
    ...getNumericStats(teamDayAverages),
    bucketCount: teamDayAverages.length,
    totalMinutes,
    monthlyMinutesPerTeamDay: teamDayAverages.length ? totalMinutes / teamDayAverages.length : null
  };
}

function getResolvedTeamType(item) {
  return state.teamData[item?.techKey]?.tipo || item?.tipo || '';
}

function getDatePartsForItem(item) {
  const day = Number(item?.day);
  if (!day) return null;
  const dateKey = item.dateStr || `${item.monthStr || ''}-${String(day).padStart(2, '0')}`;
  const date = toDateTime(`${item.monthStr || ''}-${String(day).padStart(2, '0')} 00:00`);
  return { day, dateKey, date };
}

function getWeightedTeamDayTmaStats(items, getWeight) {
  const buckets = {};
  let totalMinutes = 0;

  (items || []).forEach(item => {
    const tmaMin = getOsTmaMinutes(item);
    if (tmaMin == null || isNaN(tmaMin)) return;

    const parts = getDatePartsForItem(item);
    if (!parts) return;

    const teamKey = item.techKey || item.nomeOriginal || item.nome || 'SEM_EQUIPE';
    const resolvedType = getResolvedTeamType(item);
    const dow = parts.date ? parts.date.getDay() : null;
    const weight = getWeight({ item, teamKey, resolvedType, dow, dateKey: parts.dateKey });
    if (!weight || weight <= 0) return;

    totalMinutes += tmaMin;
    const bucketKey = `${parts.dateKey}::${teamKey}`;
    if (!buckets[bucketKey]) buckets[bucketKey] = { total: 0, count: 0, weight };
    buckets[bucketKey].total += tmaMin;
    buckets[bucketKey].count++;
  });

  const averages = Object.values(buckets).map(bucket => bucket.total / bucket.count);
  const equivalentBucketCount = Object.values(buckets).reduce((sum, bucket) => sum + (bucket.weight || 0), 0);

  return {
    ...getNumericStats(averages),
    bucketCount: Object.keys(buckets).length,
    equivalentBucketCount,
    totalMinutes,
    monthlyMinutesPerTeamDay: equivalentBucketCount ? totalMinutes / equivalentBucketCount : null
  };
}

function getTeamDayTmaSplitStats(items) {
  return {
    weekday: getWeightedTeamDayTmaStats(items, ({ dow }) => {
      if (dow == null) return 1;
      return dow >= 1 && dow <= 5 ? 1 : 0;
    }),
    weekend: getWeightedTeamDayTmaStats(items, ({ resolvedType, dow }) => {
      if (dow === 6) {
        return resolvedType === 'TECNICO 12/36H' ? 1 : 0.5;
      }
      if (dow === 0) {
        return resolvedType === 'TECNICO 12/36H' ? 1 : 0;
      }
      return 0;
    })
  };
}

function getTmaAlertThresholdMinutes(items, techKey) {
  const supportTypes = new Set(['SUPORTE MOTO', 'SUPORTE CARRO']);
  const explicitType = techKey ? state.teamData[techKey]?.tipo : null;
  if (supportTypes.has(explicitType)) return 60;

  const types = [...new Set((items || []).map(item => item?.tipo).filter(Boolean))];
  return types.length === 1 && supportTypes.has(types[0]) ? 60 : 120;
}

function getFilterViewCacheKey(monthValue, cityValue, typeValue) {
  return `${monthValue || ''}__${cityValue || 'ALL'}__${typeValue || 'ALL'}`;
}

function buildCurrentFilterMeta(filtered) {
  const cityItems = {};
  const techItems = {};
  const cityMap = {};
  const techTotals = {};
  const auxTotals = {};
  const techDays = {};
  const citySummary = {};
  const dayMap = {};
  const activeTechKeys = new Set();

  (filtered || []).forEach(item => {
    if (!cityItems[item.cidade]) cityItems[item.cidade] = [];
    cityItems[item.cidade].push(item);

    const techKey = `${item.cidade}::${item.techKey}`;
    if (!techItems[techKey]) techItems[techKey] = [];
    techItems[techKey].push(item);

    if (!cityMap[item.cidade]) cityMap[item.cidade] = {};
    if (!cityMap[item.cidade][item.techKey]) cityMap[item.cidade][item.techKey] = { dias: {}, total: 0, tipo: item.tipo };
    cityMap[item.cidade][item.techKey].dias[item.day] = (cityMap[item.cidade][item.techKey].dias[item.day] || 0) + 1;
    cityMap[item.cidade][item.techKey].total++;
    if (state.teamData[item.techKey]?.tipo) cityMap[item.cidade][item.techKey].tipo = state.teamData[item.techKey].tipo;

    const resolvedType = state.teamData[item.techKey]?.tipo || item.tipo || 'INSTALAÇÃO CIDADE';
    if (resolvedType === 'AUXILIAR') auxTotals[item.techKey] = (auxTotals[item.techKey] || 0) + 1;
    else techTotals[item.techKey] = (techTotals[item.techKey] || 0) + 1;

    if (!techDays[item.techKey]) techDays[item.techKey] = new Set();
    techDays[item.techKey].add(item.day);

    if (!citySummary[item.cidade]) citySummary[item.cidade] = { total: 0, techs: new Set() };
    citySummary[item.cidade].total++;
    citySummary[item.cidade].techs.add(item.techKey);

    dayMap[item.day] = (dayMap[item.day] || 0) + 1;
    activeTechKeys.add(item.techKey);
  });

  const cityTmaStats = {};
  Object.keys(cityItems).forEach(cidade => {
    cityTmaStats[cidade] = getTmaStats(cityItems[cidade]);
  });

  const techTmaStats = {};
  Object.keys(techItems).forEach(key => {
    techTmaStats[key] = getTmaStats(techItems[key]);
  });

  state.currentFiltered = filtered || [];
  const meta = {
    cityItems,
    techItems,
    cityTmaStats,
    techTmaStats,
    cityMap,
    techTotals,
    auxTotals,
    techDays,
    citySummary,
    dayMap,
    activeTechKeys
  };
  state.currentFilterMeta = meta;
  return meta;
}

function buildCityMap(filtered) {
  const cm = {};
  filtered.forEach(i=>{
    if(!cm[i.cidade])cm[i.cidade]={};
    if(!cm[i.cidade][i.techKey])cm[i.cidade][i.techKey]={dias:{},total:0,tipo:i.tipo};
    cm[i.cidade][i.techKey].dias[i.day]=(cm[i.cidade][i.techKey].dias[i.day]||0)+1;
    cm[i.cidade][i.techKey].total++;
    if(state.teamData[i.techKey]?.tipo) cm[i.cidade][i.techKey].tipo=state.teamData[i.techKey].tipo;
  });
  return cm;
}

function renderHeaderHTML(dIM, fdw, cidade) {
  const wkBg = ['#0F2540','#162D50','#0F2540','#162D50'];
  let r0=`<tr><th rowspan="3" class="cn th-name-col">Equipe / Técnico</th>`;
  let r1=`<tr>`, r2=`<tr>`;
  let cw=1, cdw=0, dic=0;
  if(fdw>0){
    r0+=`<th colspan="7" class="week-th" style="background:${wkBg[0]};">SEM 1</th><th class="week-th week-sep" style="background:${wkBg[0]};"></th>`;
    for(let i=0;i<fdw;i++){
      const wk=cdw===0||cdw===6;
      r1+=`<th class="${wk?'wknd':''}" style="background:${wkBg[0]};">—</th>`;
      r2+=`<th class="${wk?'wknd':''}" style="background:${wkBg[0]};font-size:7px;color:${wk?'#7BA7CC':'#64748B'};">${DN[cdw]}</th>`;
      cdw++;dic++;
    }
  }
  const cityStr = cidade ? `,'${cidade.replace(/'/g, "\\\\'")}'` : ",''";
  for(let d=1;d<=dIM;d++){
    if(cdw>6){
      const bg=wkBg[(cw-1)%4];
      r1+=`<th class="ctot week-sep" style="background:#0A1628;border-left:2px solid rgba(0,180,255,.25)!important;">TOT</th>`;
      r2+=`<th style="background:#0A1628;border-left:2px solid rgba(0,180,255,.25)!important;"></th>`;
      cdw=0;cw++;dic=0;
    }
    if(dic===0){
      const bg=wkBg[(cw-1)%4];
      r0+=`<th colspan="7" class="week-th" style="background:${bg};">SEM ${cw}</th><th class="week-th week-sep" style="background:${bg};"></th>`;
    }
    const bg=wkBg[(cw-1)%4];
    const wk=cdw===0||cdw===6;
    r1+=`<th class="${wk?'wknd':''} clickable-day" style="background:${bg};cursor:pointer;transition:background 0.2s;" onmouseover="this.style.filter='brightness(1.5)'" onmouseout="this.style.filter='none'" onclick="App.showDayDetails(${d}${cityStr})" title="Ver O.S. do dia ${d}">${d}</th>`;
    r2+=`<th class="${wk?'wknd':''}" style="background:${bg};font-size:7px;color:${wk?'#7BA7CC':'#7AA8CC'};">${DN[cdw]}</th>`;
    cdw++;dic++;
  }
  if(dic>0){
    const bg=wkBg[(cw-1)%4];
    for(let i=0;i<7-dic;i++){
      const wk=cdw===0||cdw===6;
      r1+=`<th class="${wk?'wknd':''}" style="background:${bg};">—</th>`;
      r2+=`<th class="${wk?'wknd':''}" style="background:${bg};"></th>`;
      cdw++;
    }
    r1+=`<th class="ctot week-sep" style="background:#0A1628;border-left:2px solid rgba(0,180,255,.25)!important;">TOT</th>`;
    r2+=`<th style="background:#0A1628;border-left:2px solid rgba(0,180,255,.25)!important;"></th>`;
  }
  r0+=`<th rowspan="3" class="ctos th-total">TOTAL<br><span style="font-size:16px;font-weight:900;letter-spacing:-.02em;">O.S.</span></th>`;
  r0+=`<th colspan="3" class="week-th th-cap-group" style="background:#0F2540;">CAP.</th></tr>`;
  r1+=`<th class="ccap" style="color:#60A5FA;background:#0A1628;font-size:9px;">Exc.</th><th class="ccap" style="color:#4ADE80;background:#0A1628;font-size:9px;">Bom</th><th class="ccap" style="color:#FCD34D;background:#0A1628;font-size:9px;">Med.</th></tr>`;
  r2+=`<th style="background:#0A1628;"></th><th style="background:#0A1628;"></th><th style="background:#0A1628;"></th></tr>`;
  return r0+r1+r2;
}

function renderTechRow(tk, data, dIM, fdw, cidade) {
  const isAux = data.tipo === 'AUXILIAR';
  const meta = isAux ? 0 : (state.appSettings.metasDiarias[data.tipo]||5);
  const nome = state.teamData[tk].originalName;
  const tipo = TEAM_TYPES[data.tipo]||data.tipo;
  const bm   = data.tipo==='TECNICO 12/36H'?15:24;
  const ce=meta*bm, cb=Math.max(0,(meta-1)*bm), cm2=Math.max(0,(meta-2)*bm);
  const auxBadge = isAux && data.total > 0 ? `<span class="aux-badge">✓ ${data.total} OS</span>` : '';
  const initials = nome.split(' ').filter(w=>w.length>2).slice(0,2).map(w=>w[0]).join('');
  const tipoColors = {
    'INSTALAÇÃO CIDADE':'#2563EB',
    'TECNICO 12/36H'   :'#6D28D9',
    'SUPORTE MOTO'     :'#059669',
    'SUPORTE CARRO'    :'#0891B2',
    'RURAL'            :'#92400E',
    'FAZ TUDO'         :'#9333EA',
    'AUXILIAR'         :'#64748B'
  };
  const badgeColor = tipoColors[data.tipo] || '#64748B';
  const cityStr = cidade ? `,'${cidade.replace(/'/g, "\\\\'")}'` : ",''";
  const tkStr = tk ? `,'${tk.replace(/'/g, "\\\\'")}'` : '';
  const tmaStats = state.currentFilterMeta?.techTmaStats?.[`${cidade}::${tk}`] || getTmaStats([]);
  let row=`<tr class="mat-row${isAux?' aux-row':''}"><td class="cn">
    <div class="cn-inner">
      <div class="cn-avatar" style="background:rgba(255,255,255,0.06);color:#7A92AA;border:1px solid rgba(255,255,255,0.10);">${initials}</div>
      <div class="cn-info">
        <div class="cn-name">${nome}${auxBadge}</div>
        <span class="cn-badge" style="background:rgba(255,255,255,0.05);color:#7A92AA;border:1px solid rgba(255,255,255,0.10);">${tipo}</span>
      </div>
    </div>
  </td>`;
  let wt=0, cur=0;
  for(let i=0;i<fdw;i++){
    const dow=i;
    row+=`<td class="${(dow===0||dow===6)?'cwknd':''}"></td>`;
    cur++;
  }
  for(let d=1;d<=dIM;d++){
    if(cur>6){row+=`<td class="ctot">${wt>0?wt:''}</td>`;wt=0;cur=0;}
    const v=data.dias[d]||0; wt+=v;
    const dow=(fdw+d-1)%7;
    const wk=dow===0||dow===6;
    let vc;
    if(isAux){ vc = v>0?'vg':'vd'; }
    else { vc=vC(v,meta,dow,data.tipo); }
    row+=`<td class="${vc}${wk?' cwknd':''}${isAux?' aux-cell':''}" ${v>0?`style="cursor:pointer;transition:transform 0.1s;" onmouseover="this.style.transform='scale(1.15)'" onmouseout="this.style.transform='none'" onclick="App.showDayDetails(${d}${cityStr}${tkStr})" title="Ver ${v} O.S. do dia ${d}"`:''}>${v>0?v:''}</td>`;
    cur++;
  }
  if(cur>0){for(let i=0;i<7-cur;i++)row+=`<td></td>`;row+=`<td class="ctot">${wt>0?wt:''}</td>`;}
  const onClickTot = `style="cursor:pointer;transition:opacity 0.2s;" onmouseover="this.style.opacity='0.75'" onmouseout="this.style.opacity='1'" onclick="App.showDayDetails(null${cityStr}${tkStr})" title="Ver todas as ${data.total} O.S. no mês"`;
  if(isAux){
    row+=`<td class="ctos ${data.total>0?'tg':'vd'}" ${data.total>0?onClickTot:''}>${data.total||'—'}</td>`;
    row+=`<td class="ccap" colspan="3" style="color:var(--text-3);font-size:10px;text-align:center;font-style:italic;">sem meta</td></tr>`;
  } else {
    row+=`<td class="ctos ${tC(data.total,ce,cb,cm2)}" ${data.total>0?onClickTot:''}>${data.total}</td>`;
    row+=`<td class="ccap cg">${ce}</td><td class="ccap cb">${cb>0?cb:0}</td><td class="ccap cy">${cm2>0?cm2:0}</td></tr>`;
  }
  return row;
}

function generateMatrix(filtered, cacheKey) {
  const wrapper = document.getElementById('matrixWrapper');
  const cachedView = cacheKey ? state.filterViewCache[cacheKey] : null;
  if (cachedView?.matrixHtml) {
    wrapper.innerHTML = cachedView.matrixHtml;
    if(isMobile())wrapper.classList.add('compact');
    setTimeout(applyCapVisibility, 0);
    return;
  }
  if (!filtered.length) {
    wrapper.innerHTML=`<div style="padding:40px;text-align:center;font-family:var(--mono);font-size:11px;color:var(--text-3);background:var(--card-bg);border:1px solid var(--card-border);border-radius:var(--r-xl);">Nenhum dado para os filtros selecionados.</div>`;
    return;
  }
  const [ys,ms]=state.activeMonthYear.split('-');
  const year=parseInt(ys), mIdx=parseInt(ms)-1;
  const dIM=new Date(year,mIdx+1,0).getDate(), fdw=new Date(year,mIdx,1).getDay();
  const cityMap=state.currentFilterMeta?.cityMap || buildCityMap(filtered);
  let html='';
  Object.keys(cityMap).sort().forEach(cidade=>{
    if(state.selectedCityTab!=='ALL'&&cidade!==state.selectedCityTab)return;
    const techs=cityMap[cidade];
    const sorted=Object.entries(techs).sort((a,b)=>b[1].total-a[1].total);
    const totOS=Object.values(techs).reduce((s,t)=>s+t.total,0);
    const thead=renderHeaderHTML(dIM,fdw,cidade);
    const tbody=sorted.map(([tk,data])=>renderTechRow(tk,data,dIM,fdw,cidade)).join('');
    
    let capExc=0, capBom=0, capMed=0;
    sorted.forEach(([tk, d]) => {
      if(d.tipo === 'AUXILIAR') return;
      const m  = state.appSettings.metasDiarias[d.tipo] || 5;
      const bm = d.tipo === 'TECNICO 12/36H' ? 15 : 24;
      capExc += m * bm;
      capBom += (m - 1) * bm;
      capMed += (m - 2) * bm;
    });
    capBom = Math.max(0, capBom);
    capMed = Math.max(0, capMed);

    const pctExc = capExc > 0 ? Math.round(totOS / capExc * 100) : 0;
    const pctBom = capBom > 0 ? Math.round(totOS / capBom * 100) : 0;
    const pctMed = capMed > 0 ? Math.round(totOS / capMed * 100) : 0;

    const pctColor = pctMed >= 100 ? '#60A5FA' : pctMed >= 80 ? '#4ADE80' : pctMed >= 60 ? '#FCD34D' : '#F87171';
    const pctBomColor = pctBom >= 100 ? '#60A5FA' : pctBom >= 80 ? '#4ADE80' : pctBom >= 60 ? '#FCD34D' : '#F87171';

    html+=`<div class="mat-block">
      <div class="mat-hdr">
        <div class="mat-hdr-left">
          <div class="mat-dot"></div>
          <div class="mat-city">${cidade}</div>
        </div>
        <div class="mat-hdr-kpis">
          <div class="mat-kpi"><span class="mat-kpi-icon">👷</span><div><div class="mat-kpi-val">${sorted.length}</div><div class="mat-kpi-lbl">Técnicos</div></div></div>
          <div class="mat-kpi-sep"></div>
          <div class="mat-kpi" style="cursor:pointer;" onclick="App.showDayDetails(null,'${cidade.replace(/'/g, "\\'")}')" title="Ver detalhamento mensal da regional ${cidade}"><span class="mat-kpi-icon">📋</span><div><div class="mat-kpi-val">${totOS}</div><div class="mat-kpi-lbl">O.S. Entregues</div></div></div>
          <div class="mat-kpi-sep"></div>
          <div class="mat-kpi mat-kpi-cap">
            <span class="mat-kpi-icon">🎯</span>
            <div>
              <div class="mat-kpi-val" style="color:${pctColor};font-size:16px;">${pctMed}%</div>
              <div class="mat-kpi-lbl">Entrega / Cap.Med</div>
            </div>
          </div>
          <div class="mat-kpi-sep"></div>
          <div class="mat-kpi mat-kpi-cap">
            <span class="mat-kpi-icon">⭐</span>
            <div>
              <div class="mat-kpi-val" style="color:${pctBomColor};font-size:16px;">${pctBom}%</div>
              <div class="mat-kpi-lbl">Entrega / Cap.Bom</div>
            </div>
          </div>
          <div class="mat-kpi-sep"></div>
          <div class="mat-kpi mat-kpi-caps-detail">
            <span class="mat-kpi-icon">📊</span>
            <div>
              <div style="display:flex;gap:8px;align-items:center;">
                <span style="font-family:var(--mono);font-size:10px;color:#60A5FA;font-weight:700;">EXC ${capExc}</span>
                <span style="font-family:var(--mono);font-size:10px;color:#4ADE80;font-weight:700;">BOM ${capBom}</span>
                <span style="font-family:var(--mono);font-size:10px;color:#FCD34D;font-weight:700;">MED ${capMed}</span>
              </div>
              <div class="mat-kpi-lbl">Capacidades OS</div>
            </div>
          </div>
        </div>
      </div>
      <div class="mat-scroll"><table class="mat-table"><thead>${thead}</thead><tbody>${tbody}</tbody></table></div>
    </div>`; 
  });
  wrapper.innerHTML=html;
  if (cacheKey) {
    if (!state.filterViewCache[cacheKey]) state.filterViewCache[cacheKey] = {};
    state.filterViewCache[cacheKey].matrixHtml = html;
  }
  if(isMobile())wrapper.classList.add('compact');
  setTimeout(applyCapVisibility, 0);
}

/* ═══════════════════════════════════════════════════════════
   CAPACIDADES — aba de resumo por regional
   ═══════════════════════════════════════════════════════════ */
function renderCapacidades() {
  const container = document.getElementById('capSection');
  if(!container) return;
  if(!state.globalRawData.length) {
    container.innerHTML = '<div class="empty" style="padding:60px;text-align:center;">Carregue uma planilha para ver as capacidades.</div>';
    return;
  }
  const fc  = state.selectedCityTab;
  const filtered = state.currentFiltered || [];
  const cityMap = state.currentFilterMeta?.cityMap || buildCityMap(filtered);

  function capColors(p) {
    if(p>=100) return {bg:'rgba(96,165,250,.15)',border:'#60A5FA',text:'#60A5FA',badge:'#1E3A8A'};
    if(p>=80)  return {bg:'rgba(74,222,128,.18)',border:'#4ADE80',text:'#4ADE80',badge:'#14532D'};
    if(p>=60)  return {bg:'rgba(252,211,77,.14)',border:'#FCD34D',text:'#FCD34D',badge:'#78350F'};
    return       {bg:'rgba(248,113,113,.14)',border:'#F87171',text:'#F87171',badge:'#7F1D1D'};
  }
  function pctBadge(p) {
    const c = capColors(p);
    return `<span style="display:inline-flex;align-items:center;padding:2px 7px;border-radius:5px;font-family:var(--mono);font-size:10px;font-weight:800;background:${c.bg};border:1px solid ${c.border};color:${c.text};">${p}%</span>`;
  }

  let totTechs=0, totOS=0, totExc=0, totBom=0, totMed=0;
  const rows = [];

  Object.keys(cityMap).sort().forEach(cidade => {
    const techs = cityMap[cidade];
    const allEntries = Object.entries(techs);
    const entregues = allEntries.reduce((s,[,d])=>s+d.total,0);
    const byTipo = {};
    allEntries.forEach(([tk,d]) => {
      if(d.tipo==='AUXILIAR') return;
      if(!byTipo[d.tipo]) byTipo[d.tipo]={techs:0,os:0,capExc:0,capBom:0,capMed:0};
      const m  = state.appSettings.metasDiarias[d.tipo]||5;
      const bm = d.tipo==='TECNICO 12/36H'?15:24;
      byTipo[d.tipo].techs++;
      byTipo[d.tipo].os += d.total;
      byTipo[d.tipo].capExc += m*bm;
      byTipo[d.tipo].capBom += Math.max(0,(m-1)*bm);
      byTipo[d.tipo].capMed += Math.max(0,(m-2)*bm);
    });

    let capExc=0,capBom=0,capMed=0,nonAux=0;
    Object.values(byTipo).forEach(t=>{capExc+=t.capExc;capBom+=t.capBom;capMed+=t.capMed;nonAux+=t.techs;});

    const pExc = capExc>0?Math.round(entregues/capExc*100):0;
    const pBom = capBom>0?Math.round(entregues/capBom*100):0;
    const pMed = capMed>0?Math.round(entregues/capMed*100):0;

    rows.push({cidade,nonAux,entregues,capExc,capBom,capMed,pExc,pBom,pMed,byTipo});
    totTechs+=nonAux; totOS+=entregues; totExc+=capExc; totBom+=capBom; totMed+=capMed;
  });

  const tpExc=totExc>0?Math.round(totOS/totExc*100):0;
  const tpBom=totBom>0?Math.round(totOS/totBom*100):0;
  const tpMed=totMed>0?Math.round(totOS/totMed*100):0;

  const TIPO_COLORS = {
    'INSTALAÇÃO CIDADE':'#2563EB','TECNICO 12/36H':'#6D28D9',
    'SUPORTE MOTO':'#059669','SUPORTE CARRO':'#0891B2',
    'RURAL':'#92400E','FAZ TUDO':'#9333EA'
  };
  const TIPO_LABEL = {
    'INSTALAÇÃO CIDADE':'Cidade','TECNICO 12/36H':'Plantão 12/36H',
    'SUPORTE MOTO':'Sup. Moto','SUPORTE CARRO':'Sup. Carro',
    'RURAL':'Rural','FAZ TUDO':'Faz Tudo'
  };
  const maxCap = Math.max(...rows.map(r=>r.capExc), 1);

  const filialCards = rows.map(r => {
    const c = capColors(r.pMed);
    const tipoCards = Object.entries(r.byTipo).sort((a,b)=>b[1].capExc-a[1].capExc).map(([tipo,t]) => {
      const col = TIPO_COLORS[tipo]||'#64748B';
      const lbl = TIPO_LABEL[tipo]||tipo;
      
      const maxTarget = t.capExc || 1;
      const pctMed = Math.min((t.capMed / maxTarget) * 100, 100);
      const pctBom = Math.min((t.capBom / maxTarget) * 100, 100);
      const pctOS = Math.min((t.os / maxTarget) * 100, 100);
      const isOver = t.os > maxTarget;

      return `<div class="cap-tipo-card" style="border-left:4px solid ${col};background:var(--surface-hover, rgba(255,255,255,0.03));padding:14px;border-radius:8px;margin-bottom:10px;">
        <div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:14px;">
          <div>
            <div style="font-size:14px;font-weight:800;color:var(--text-primary);letter-spacing:0.02em;">${lbl}</div>
            <div style="font-size:12px;color:var(--text-secondary);margin-top:4px;">
              <span style="background:rgba(255,255,255,0.1);padding:2px 6px;border-radius:4px;">${t.techs} técnicos</span>
            </div>
          </div>
          <div style="text-align:right;">
            <div style="font-size:20px;font-weight:900;color:${col};line-height:1;">${t.os}</div>
            <div style="font-size:10px;color:var(--text-secondary);text-transform:uppercase;letter-spacing:0.05em;margin-top:4px;">OS Entregues</div>
          </div>
        </div>
        
        <div style="position:relative;height:14px;background:var(--surface-border, rgba(255,255,255,0.1));border-radius:7px;overflow:hidden;margin-bottom:10px;">
          <!-- Barra de Progresso -->
          <div style="position:absolute;left:0;top:0;bottom:0;width:${pctOS}%;background:${col};z-index:1;transition:width 0.5s ease;border-radius:7px;"></div>
          ${isOver ? `<div style="position:absolute;right:0;top:0;bottom:0;width:6px;background:#fff;z-index:3;" title="Superou a meta excelente!"></div>` : ''}
          
          <!-- Marcadores de Metas -->
          <div style="position:absolute;left:${pctMed}%;top:0;bottom:0;width:2px;background:rgba(255,255,255,0.7);z-index:2;" title="Meta Mediana: ${t.capMed}"></div>
          <div style="position:absolute;left:${pctBom}%;top:0;bottom:0;width:2px;background:rgba(255,255,255,0.7);z-index:2;" title="Meta Boa: ${t.capBom}"></div>
        </div>
        
        <div style="display:flex;justify-content:space-between;font-size:11px;font-family:var(--mono);">
          <div style="display:flex;flex-direction:column;align-items:flex-start;" title="Meta Mediana">
            <span style="color:#FCD34D;font-weight:800;letter-spacing:0.05em;">MED</span>
            <span style="color:var(--text-secondary);margin-top:2px;">${t.capMed}</span>
          </div>
          <div style="display:flex;flex-direction:column;align-items:center;" title="Meta Boa">
            <span style="color:#4ADE80;font-weight:800;letter-spacing:0.05em;">BOM</span>
            <span style="color:var(--text-secondary);margin-top:2px;">${t.capBom}</span>
          </div>
          <div style="display:flex;flex-direction:column;align-items:flex-end;" title="Meta Excelente">
            <span style="color:#60A5FA;font-weight:800;letter-spacing:0.05em;">EXC</span>
            <span style="color:var(--text-secondary);margin-top:2px;">${t.capExc}</span>
          </div>
        </div>
      </div>`;
    }).join('');

    const wBar = Math.round(r.capExc/maxCap*100);
    const wEnt = Math.round(r.entregues/maxCap*100);

    return `<div class="cap-filial-card">
      <div class="cap-filial-hdr">
        <div class="cap-filial-left">
          <div class="cap-filial-dot" style="background:${c.border};box-shadow:0 0 6px ${c.border};"></div>
          <div>
            <div class="cap-filial-nome">${r.cidade}</div>
            <div class="cap-filial-sub">${r.nonAux} técnicos · ${r.entregues} OS</div>
          </div>
        </div>
        <div class="cap-filial-badges">
          <div class="cap-filial-badge-group">
            <span style="font-family:var(--mono);font-size:8px;color:#60A5FA;margin-right:3px;">EXC</span>${pctBadge(r.pExc)}
          </div>
          <div class="cap-filial-badge-group">
            <span style="font-family:var(--mono);font-size:8px;color:#4ADE80;margin-right:3px;">BOM</span>${pctBadge(r.pBom)}
          </div>
          <div class="cap-filial-badge-group">
            <span style="font-family:var(--mono);font-size:8px;color:#FCD34D;margin-right:3px;">MED</span>${pctBadge(r.pMed)}
          </div>
        </div>
      </div>
      <div class="cap-filial-bar-wrap">
        <div style="height:6px;background:rgba(74,222,128,.12);border-radius:4px;position:relative;overflow:hidden;">
          <div style="position:absolute;inset:0;width:${wBar}%;max-width:100%;background:rgba(74,222,128,.2);"></div>
          <div style="position:absolute;inset:0;width:${wEnt}%;max-width:100%;background:${c.border};border-radius:4px;transition:width .5s ease;"></div>
        </div>
      </div>
      <div class="cap-filial-tipos">${tipoCards}</div>
    </div>`;
  }).join('');

  const tableRows = rows.map(r => `
    <tr class="cap-row">
      <td class="cap-cidade">${r.cidade}</td>
      <td class="cap-num">${r.nonAux}</td>
      <td class="cap-num" style="color:#93C5FD;font-weight:700;">${r.entregues}</td>
      <td class="cap-num" style="color:#60A5FA;">${r.capExc}</td>
      <td class="cap-num" style="color:#4ADE80;">${r.capBom}</td>
      <td class="cap-num" style="color:#FCD34D;">${r.capMed}</td>
      <td style="padding:8px 10px;text-align:center;">${pctBadge(r.pExc)}</td>
      <td style="padding:8px 10px;text-align:center;">${pctBadge(r.pBom)}</td>
      <td style="padding:8px 10px;text-align:center;">${pctBadge(r.pMed)}</td>
    </tr>`).join('');

  const totalLabel = fc === 'ALL' ? 'TOTAL GERAL' : `TOTAL ${fc}`;
  const mainTitle = fc === 'ALL' ? 'Capacidade por Filial e Tipo de Equipe' : `Capacidade da Regional ${fc}`;
  const sideTitle = fc === 'ALL' ? 'Resumo Geral' : `Resumo da Regional ${fc}`;
  const totalRow = `
    <tr class="cap-row cap-total">
      <td class="cap-cidade">${totalLabel}</td>
      <td class="cap-num">${totTechs}</td>
      <td class="cap-num" style="color:#93C5FD;font-weight:800;">${totOS}</td>
      <td class="cap-num" style="color:#60A5FA;font-weight:700;">${totExc}</td>
      <td class="cap-num" style="color:#4ADE80;font-weight:700;">${totBom}</td>
      <td class="cap-num" style="color:#FCD34D;font-weight:700;">${totMed}</td>
      <td style="padding:8px 10px;text-align:center;">${pctBadge(tpExc)}</td>
      <td style="padding:8px 10px;text-align:center;">${pctBadge(tpBom)}</td>
      <td style="padding:8px 10px;text-align:center;">${pctBadge(tpMed)}</td>
    </tr>`;

  container.innerHTML = `
    <div class="cap-layout">
      <div class="cap-col-main">
        <div class="cap-section-title">📍 Capacidade por Filial e Tipo de Equipe</div>
        <div class="cap-filiais-grid">${filialCards}</div>
      </div>
      <div class="cap-col-side">
        <div class="cap-section-title" style="display:flex;align-items:center;justify-content:space-between;">
          <span>📊 Resumo Geral</span>
          <button class="btn btn-outline" onclick="App.exportToExcel()" style="font-size:10px;padding:4px 10px;gap:5px;">
            <svg width="12" height="12" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"/></svg>
            Excel
          </button>
        </div>
        <div class="cap-table-wrap">
          <table class="cap-table">
            <thead>
              <tr>
                <th class="cap-th-cidade" rowspan="2">Regional</th>
                <th class="cap-th-num" rowspan="2">Téc.</th>
                <th class="cap-th-os" rowspan="2">O.S.</th>
                <th colspan="3" class="cap-th-caps">Capacidade</th>
                <th colspan="3" class="cap-th-pcts">% Entrega</th>
              </tr>
              <tr>
                <th class="cap-th-exc">EXC</th>
                <th class="cap-th-bom">BOM</th>
                <th class="cap-th-med">MED</th>
                <th class="cap-th-exc">EXC</th>
                <th class="cap-th-bom">BOM</th>
                <th class="cap-th-med">MED</th>
              </tr>
            </thead>
            <tbody>${tableRows}${totalRow}</tbody>
          </table>
        </div>
      </div>
    </div>`;
}

/* ═══════════════════════════════════════════════════════════
   EXPORTAÇÃO EXCEL - DESIGN HEATMAP (IMAGEM 2)
   ═══════════════════════════════════════════════════════════ */
function exportToExcel() {
  if(!state.globalRawData.length)return alert('Sem dados para exportar.');
  const [ys,ms]=state.activeMonthYear.split('-');
  const year=parseInt(ys),mIdx=parseInt(ms)-1;
  const dIM=new Date(year,mIdx+1,0).getDate(),fdw=new Date(year,mIdx,1).getDay();
  const fc=state.selectedCityTab, ft=document.getElementById('filterType')?.value||'ALL';
  const filtered=state.globalRawData.filter(i=>i.monthStr===state.activeMonthYear&&(fc==='ALL'||i.cidade===fc)&&(ft==='ALL'||i.tipo===ft));
  if(!filtered.length)return alert('Sem dados nos filtros selecionados.');
  const cityMap=buildCityMap(filtered);
  const wb=XLSX.utils.book_new();
  const XL={navy:'FF0D1117',navyMid:'FF1A2236',navyLight:'FF243047',navyTxt:'FFCBD5E1',weekTxt:'FF93C5FD',white:'FFFFFFFF',border:'FF000000',greenBg:'FF16A34A',greenFg:'FFFFFFFF',blueBg:'FF1D6FE8',blueFg:'FFFFFFFF',yellowBg:'FFCA8A04',yellowFg:'FFFFFFFF',redBg:'FFDC2626',redFg:'FFFFFFFF',totBg:'FFEFF3F9',totFg:'FF64748B'};
  const cs=(fill,font,bold=false,align='center',sz=9)=>({fill:fill?{fgColor:{rgb:fill}}:{},font:{color:{rgb:font||XL.navy},bold,name:'Calibri',sz},border:{top:{style:'thin',color:{rgb:XL.border}},bottom:{style:'thin',color:{rgb:XL.border}},left:{style:'thin',color:{rgb:XL.border}},right:{style:'thin',color:{rgb:XL.border}}},alignment:{horizontal:align,vertical:'center',wrapText:false}});
  
  Object.keys(cityMap).sort().forEach(cidade=>{
    const techs=cityMap[cidade];
    const sorted=Object.entries(techs).sort((a,b)=>b[1].total-a[1].total);
    let wsData=[],merges=[],r0=['Técnico','Modelo de Equipe'],r1=['',''],r2=['',''];
    let cw=1,cdw=0,dic=0,ci=2; 
    if(fdw>0){r0.push('SEM 1');merges.push({s:{r:0,c:ci},e:{r:0,c:ci+6}});for(let i=0;i<6;i++)r0.push('');r0.push('');for(let i=0;i<fdw;i++){r1.push('—');r2.push(DN[cdw]);cdw++;dic++;ci++;}}
    for(let d=1;d<=dIM;d++){
      if(cdw>6){r1.push('TOT');r2.push('');ci++;cdw=0;cw++;dic=0;}
      if(dic===0){r0.push(`SEM ${cw}`);merges.push({s:{r:0,c:ci},e:{r:0,c:ci+6}});for(let i=0;i<6;i++)r0.push('');r0.push('');}
      r1.push(d);r2.push(DN[cdw]);cdw++;dic++;ci++;
    }
    if(dic>0){let rem=7-dic;for(let i=0;i<rem;i++){r1.push('—');r2.push(DN[cdw]);cdw++;ci++;}r1.push('TOT');r2.push('');ci++;}
    r0.push('TOTAL');merges.push({s:{r:0,c:ci},e:{r:2,c:ci}});
    r0.push('CAPACIDADE');merges.push({s:{r:0,c:ci+1},e:{r:0,c:ci+3}});r0.push('');r0.push('');
    r1.push('');r1.push('Excelente');r1.push('Bom');r1.push('Mediano');
    r2.push('');r2.push('');r2.push('');r2.push('');
    wsData.push(r0);wsData.push(r1);wsData.push(r2);
    let mnl=18;
    const firstDowXL = fdw; 
    sorted.forEach(([tk,data])=>{
      const baseMeta=state.appSettings.metasDiarias[data.tipo]||5;
      const nd=state.teamData[tk].originalName;if(nd.length>mnl)mnl=nd.length;
      const tipoLabel=TEAM_TYPES[data.tipo]||data.tipo;
      let row=[nd, tipoLabel],wt=0,cur=0;
      for(let i=0;i<fdw;i++){row.push('');cur++;}
      for(let d=1;d<=dIM;d++){if(cur>6){row.push(wt>0?wt:'');wt=0;cur=0;}const v=data.dias[d]||0;wt+=v;row.push(v>0?v:'');cur++;}
      if(cur>0){for(let i=0;i<7-cur;i++)row.push('');row.push(wt>0?wt:'');}
      const bm=data.tipo==='TECNICO 12/36H'?15:24;const ce=baseMeta*bm,cb=(baseMeta-1)*bm,cm2=(baseMeta-2)*bm;
      row.push(data.total);row.push(ce);row.push(cb>0?cb:0);row.push(cm2>0?cm2:0);
      wsData.push(row);
    });
    const ws=XLSX.utils.aoa_to_sheet(wsData);ws['!merges']=merges;
    let cols=[{wch:mnl+4},{wch:18}]; 
    for(let i=2;i<r0.length;i++){if(r0[i]==='TOTAL')cols.push({wch:9});else if(['Excelente','Bom','Mediano'].includes(r1[i]))cols.push({wch:10});else if(r1[i]==='TOT')cols.push({wch:5});else cols.push({wch:4.5});}
    ws['!cols']=cols;ws['!rows']=[{hpt:22},{hpt:17},{hpt:13}];
    const range=XLSX.utils.decode_range(ws['!ref']);const tci=r0.indexOf('TOTAL');
    
    for(let R=range.s.r;R<=range.e.r;++R){for(let C=range.s.c;C<=range.e.c;++C){
      const ca=XLSX.utils.encode_cell({r:R,c:C});
      if(!ws[ca])ws[ca]={v:'',t:'s'};if(!ws[ca].s)ws[ca].s={};
      const isTot=r1[C]==='TOT',isCap=['Excelente','Bom','Mediano'].includes(r1[C]);
      
      // ESTILO CORRIGIDO SEM ROXO E SEM ERRO DE SINTAXE
      if(R===0){ws[ca].s=cs(r0[C]?.startsWith('SEM')?XL.navyLight:XL.navyMid,r0[C]?.startsWith('SEM')?XL.weekTxt:XL.navyTxt,true,'center',9);}
      else if(R===1){if(r1[C]==='Excelente')ws[ca].s=cs('FF0000FF','FFFFFFFF',true,'center',10);else if(r1[C]==='Bom')ws[ca].s=cs('FF00B050','FFFFFFFF',true,'center',10);else if(r1[C]==='Mediano')ws[ca].s=cs('FFCA8A04','FFFFFFFF',true,'center',10);else ws[ca].s=cs(XL.navyMid,XL.navyTxt,true,'center',9);}
      else if(R===2){ws[ca].s=cs(XL.navyMid,'FF5A6478',false,'center',8);}
      else if(C===0){ws[ca].s=cs(XL.white,XL.navy,true,'left',10);}
      else if(C===1){ws[ca].s=cs(XL.white,XL.navy,false,'left',9);} // Letra padronizada (sem roxo)
      else if(C===tci&&R>2){const vt=Number(ws[ca].v||0),vce=Number(ws[XLSX.utils.encode_cell({r:R,c:tci+1})]?.v||0),vcb=Number(ws[XLSX.utils.encode_cell({r:R,c:tci+2})]?.v||0),vcm=Number(ws[XLSX.utils.encode_cell({r:R,c:tci+3})]?.v||0);let bg=XL.redBg,fg=XL.redFg;if(vt>=vce){bg=XL.blueBg;fg=XL.blueFg;}else if(vt>=vcb){bg=XL.greenBg;fg=XL.greenFg;}else if(vt>=vcm){bg=XL.yellowBg;fg=XL.yellowFg;}ws[ca].s=cs(bg,fg,true,'center',11);}
      else if(isCap){ws[ca].s=cs(XL.white,'FF9CA3AF',false,'center',10);}
      else if(isTot){ws[ca].s=cs(XL.totBg,XL.totFg,true,'center',9);}
      else if(R>2){
        const v=Number(ws[ca].v||0),rowIdx=R-3;
        if(rowIdx<sorted.length){
          const rowTipo=sorted[rowIdx][1].tipo;
          const baseMeta=state.appSettings.metasDiarias[rowTipo]||5;
          const dayNum=typeof r1[C]==='number'?r1[C]:null;
          let metaEff=baseMeta;
          if(dayNum!==null){
            const dow=(firstDowXL+dayNum-1)%7;
            metaEff=getMetaForDay(baseMeta,dow,rowTipo);
          }
          let bg=XL.white,fg='FF9CA3AF',bold=false;
          if(v>0){
            if(metaEff===0){bg=XL.blueBg;fg=XL.blueFg;bold=true;} 
            else if(v>=metaEff){bg=XL.blueBg;fg=XL.blueFg;bold=true;}
            else if(v>=metaEff-1){bg=XL.greenBg;fg=XL.greenFg;bold=true;}
            else if(v>=metaEff-2){bg=XL.yellowBg;fg=XL.yellowFg;bold=true;}
            else{bg=XL.redBg;fg=XL.redFg;bold=true;}
          }
          ws[ca].s=cs(bg,fg,bold,'center',9);
        }
      }
    }}
    const rowH=ws['!rows']||[];for(let R=3;R<=range.e.r;R++)rowH.push({hpt:15});
    ws['!rows']=rowH;ws['!freeze']={xSplit:1,ySplit:3};
    XLSX.utils.book_append_sheet(wb,ws,cidade.substring(0,31));
  });

  // --- ABA RAIO-X DIAGNOSTICOS ---
  const resumoData = [['Modelo de Equipe', 'Assunto', 'Diagnóstico', 'Qtd O.S.', '% na Equipe']];
  const statsEquipe = {};
  
  filtered.forEach(os => {
    const tipo = os.tipo || 'DESCONHECIDO';
    if (!statsEquipe[tipo]) statsEquipe[tipo] = { total: 0, assuntos: {} };
    statsEquipe[tipo].total++;
    
    const ass = os.assunto || 'Sem Assunto';
    if (!statsEquipe[tipo].assuntos[ass]) statsEquipe[tipo].assuntos[ass] = { total: 0, diag: {} };
    statsEquipe[tipo].assuntos[ass].total++;
    
    const diag = os.diagnostico || 'Sem Diagnóstico';
    statsEquipe[tipo].assuntos[ass].diag[diag] = (statsEquipe[tipo].assuntos[ass].diag[diag] || 0) + 1;
  });

  Object.keys(statsEquipe).sort().forEach(tipo => {
    const tData = statsEquipe[tipo];
    Object.keys(tData.assuntos).sort((a,b) => tData.assuntos[b].total - tData.assuntos[a].total).forEach(ass => {
      const aData = tData.assuntos[ass];
      Object.keys(aData.diag).sort((a,b) => aData.diag[b] - aData.diag[a]).forEach(diag => {
        const qtd = aData.diag[diag];
        const pctEquipe = ((qtd / tData.total) * 100).toFixed(1) + '%';
        resumoData.push([TEAM_TYPES[tipo] || tipo, ass, diag, qtd, pctEquipe]);
      });
    });
  });

  const wsResumo = XLSX.utils.aoa_to_sheet(resumoData);
  wsResumo['!cols'] = [{wch: 20}, {wch: 55}, {wch: 55}, {wch: 12}, {wch: 15}];
  
  const rangeResumo = XLSX.utils.decode_range(wsResumo['!ref']);
  for(let R = rangeResumo.s.r; R <= rangeResumo.e.r; ++R) {
    for(let C = rangeResumo.s.c; C <= rangeResumo.e.c; ++C) {
      const cellRef = XLSX.utils.encode_cell({r: R, c: C});
      if(!wsResumo[cellRef]) continue;
      if (R === 0) {
        wsResumo[cellRef].s = cs(XL.navyMid, XL.white, true, 'center', 10);
      } else {
        let align = (C === 3 || C === 4) ? 'right' : 'left';
        let bold = (C === 0); 
        wsResumo[cellRef].s = cs(XL.white, XL.navy, bold, align, 9);
      }
    }
  }
  wsResumo['!autofilter'] = { ref: wsResumo['!ref'] };
  XLSX.utils.book_append_sheet(wb, wsResumo, 'Raio-X Diagnosticos');
  
  XLSX.writeFile(wb,`SGO_Matriz_${state.activeMonthYear}.xlsx`);
}

/* ═══════════════════════════════════════════════════════════
   ANALYSIS — análise local e Gemini
   ═══════════════════════════════════════════════════════════ */
function buildOperationalAnalysis(filtered) {
  if (!filtered.length) return;
  const meta = state.currentFilterMeta || {};
  const techTotalsFast = meta.techTotals || {};
  const techDiasFast = meta.techDays || {};
  const auxTotalsFast = meta.auxTotals || {};
  const sortedFast=Object.entries(techTotalsFast).sort((a,b)=>b[1]-a[1]);
  if(sortedFast.length){
    const totalOS=filtered.length, totalTechs=sortedFast.length, avgOS=totalOS/totalTechs;
    const best=sortedFast[0], worst=sortedFast[sortedFast.length-1];
    const bestNome=state.teamData[best[0]]?.originalName||best[0];
    const worstNome=state.teamData[worst[0]]?.originalName||worst[0];
    const [ys,ms]=state.activeMonthYear.split('-');
    const dIM=new Date(parseInt(ys),parseInt(ms),0).getDate();
    const dayMapFast=meta.dayMap || {};
    const dailyArr=Array.from({length:dIM},(_,i)=>dayMapFast[i+1]||0);
    const diasUteisProd=dailyArr.filter((v,i)=>{const dw=new Date(parseInt(ys),parseInt(ms)-1,i+1).getDay();return dw>=1&&dw<=5&&v>0;});
    const diasSabProd=dailyArr.filter((v,i)=>{const dw=new Date(parseInt(ys),parseInt(ms)-1,i+1).getDay();return dw===6&&v>0;});
    const workDays=dailyArr.filter(v=>v>0);
    const nonSunArr=dailyArr.filter((v,i)=>{const dw=new Date(parseInt(ys),parseInt(ms)-1,i+1).getDay();return dw!==0;});
    const nonSunMin=Math.min(...nonSunArr.filter(v=>v>0));
    const weakIdx=dailyArr.findIndex((v,i)=>{const dw=new Date(parseInt(ys),parseInt(ms)-1,i+1).getDay();return v===nonSunMin&&dw!==0;});
    const weakDay=weakIdx>=0?`Dia ${weakIdx+1} (${DN[new Date(parseInt(ys),parseInt(ms)-1,weakIdx+1).getDay()]})`:'—';
    const weakVal=weakIdx>=0?dailyArr[weakIdx]:0;
    const abovePct=sortedFast.filter(([k,v])=>{
      const tipo=state.teamData[k]?.tipo||'INSTALAÇÃO CIDADE';
      if(tipo==='AUXILIAR')return false;
      const m=state.appSettings.metasDiarias[tipo]||5;
      const bm=tipo==='TECNICO 12/36H'?15:24;
      return v/(m*bm)>=0.8;
    });
    const effPct=totalTechs>0?Math.round(abovePct.length/totalTechs*100):0;
    const bestMeta=state.appSettings.metasDiarias[state.teamData[best[0]]?.tipo||'INSTALAÇÃO CIDADE']||5;
    const bestBm=state.teamData[best[0]]?.tipo==='TECNICO 12/36H'?15:24;
    const bestCap=Math.round(best[1]/(bestMeta*bestBm)*100);
    document.getElementById('aEfic').textContent=effPct+'%';
    document.getElementById('aEficSub').textContent=`${abovePct.length} de ${totalTechs} técnicos acima de 80% da capacidade`;
    document.getElementById('aBest').textContent=bestNome;
    document.getElementById('aBestSub').textContent=`${best[1]} OS · ${bestCap}% da cap.`;
    document.getElementById('aWorst').textContent=worstNome;
    document.getElementById('aWorstSub').textContent=`${worst[1]} OS · ${Math.round(worst[1]/avgOS*100)}% da média geral`;
    document.getElementById('aWeakDay').textContent=weakDay;
    document.getElementById('aWeakDaySub').textContent=`${weakVal} OS · dia de menor produção`;
    const topList=sortedFast.slice(0,5).map(([k,v])=>{
      const tipo=state.teamData[k]?.tipo||'INSTALAÇÃO CIDADE';
      const m=state.appSettings.metasDiarias[tipo]||5;
      const bm=tipo==='TECNICO 12/36H'?15:24;
      const cap=Math.round(v/(m*bm)*100);
      const tipoLabel=TEAM_TYPES[tipo]||tipo;
      return`<div class="obs-item"><div class="obs-dot g"></div><span><b>${state.teamData[k]?.originalName||k}</b> <span class="obs-type-tag">${tipoLabel}</span> — ${v} OS <span class="obs-cap-tag ${cap>=100?'cap-ex':cap>=80?'cap-bom':''}">${cap}%</span></span></div>`;
    }).join('');
    document.getElementById('obsPositivos').innerHTML=topList||'<div class="empty" style="padding:12px;">Sem dados.</div>';
    const botList=sortedFast.filter(([k,v])=>v<avgOS*0.6).slice(0,5).map(([k,v])=>{
      const diff=Math.round((1-v/avgOS)*100);
      const tipo=state.teamData[k]?.tipo||'INSTALAÇÃO CIDADE';
      const tipoLabel=TEAM_TYPES[tipo]||tipo;
      const dias=techDiasFast[k]?.size||0;
      return`<div class="obs-item"><div class="obs-dot r"></div><span><b>${state.teamData[k]?.originalName||k}</b> <span class="obs-type-tag">${tipoLabel}</span> — ${v} OS em ${dias} dias (${diff}% abaixo da média)</span></div>`;
    }).join('');
    document.getElementById('obsAtencao').innerHTML=botList||'<div class="obs-item"><div class="obs-dot g"></div><span>Todos os técnicos acima de 60% da média!</span></div>';
    const auxList=Object.entries(auxTotalsFast).sort((a,b)=>b[1]-a[1]);
    const diasComOS=workDays.length;
    const diasUteisTotal=dailyArr.filter((_,i)=>{const dw=new Date(parseInt(ys),parseInt(ms)-1,i+1).getDay();return dw!==0&&dw!==6;}).length;
    const mediaUteis=diasUteisProd.length>0?Math.round(diasUteisProd.reduce((a,b)=>a+b,0)/diasUteisProd.length):0;
    const mediaSab=diasSabProd.length>0?Math.round(diasSabProd.reduce((a,b)=>a+b,0)/diasSabProd.length):0;
    const obs=[];
    obs.push(`Presença em <b>${Math.round(diasComOS/diasUteisTotal*100)}%</b> dos dias úteis (${diasUteisTotal} dias)`);
    obs.push(`Média de <b>${mediaUteis} OS/dia</b> em dias úteis${mediaSab>0?' · <b>'+mediaSab+' OS/dia</b> aos sábados':''}`);
    if(sortedFast.length>1) obs.push(`Amplitude: <b>${best[1]-worst[1]} OS</b> entre o melhor e o pior técnico`);
    if(auxList.length>0) obs.push(`<b>${auxList.length}</b> auxiliar(es) com OS finalizada(s) este período`);
    const diasSemOS=dailyArr.filter((v,i)=>{const dw=new Date(parseInt(ys),parseInt(ms)-1,i+1).getDay();return v===0&&dw>=1&&dw<=5;}).length;
    if(diasSemOS>0) obs.push(`<b>${diasSemOS}</b> dias úteis sem nenhuma OS registrada`);
    document.getElementById('obsRapidas').innerHTML=obs.map(o=>`<div class="obs-item"><div class="obs-dot y"></div><span>${o}</span></div>`).join('');
    const sug=[];
    if(worst[1]<avgOS*0.5) sug.push(`Verificar situação de <b>${worstNome}</b> — produção abaixo de 50% da média da equipe`);
    if(effPct<40) sug.push(`Apenas <b>${effPct}%</b> da equipe atingiu 80%+ da capacidade — revisar distribuição de OS`);
    else if(effPct<60) sug.push(`<b>${100-effPct}%</b> da equipe abaixo de 80% — identificar gargalos de atendimento`);
    if(weakVal<avgOS*0.3&&weakIdx>=0) sug.push(`Investigar a baixa produção no <b>${weakDay}</b> — possível falta de escala ou OS`);
    if(diasSemOS>2) sug.push(`Revisar escalonamento — <b>${diasSemOS} dias úteis</b> sem OS registrada`);
    const cidadesArr=Object.entries(meta.citySummary || {}).map(([c,d])=>({nome:c,media:d.total/d.techs.size}));
    if(cidadesArr.length>1){cidadesArr.sort((a,b)=>b.media-a.media);const diff=Math.round(cidadesArr[0].media-cidadesArr[cidadesArr.length-1].media);if(diff>20)sug.push(`Desbalanceamento entre polos: <b>${cidadesArr[0].nome}</b> produz ${diff} OS/técnico a mais que <b>${cidadesArr[cidadesArr.length-1].nome}</b>`);}
    if(!sug.length) sug.push('Operação estável — manter monitoramento semanal de produtividade');
    document.getElementById('obsSugestoes').innerHTML=sug.slice(0,4).map(s=>`<div class="obs-item"><div class="obs-dot b"></div><span>${s}</span></div>`).join('');
    document.getElementById('badgeAnalise')?.classList.remove('hidden');
    document.getElementById('mobBadgeAnalise')?.classList.remove('hidden');
    return;
  }
  const techTotals={}, techDias={}, auxTotals={};

  filtered.forEach(i=>{
    const tipo = state.teamData[i.techKey]?.tipo||'INSTALAÇÃO CIDADE';

    if(tipo==='AUXILIAR'){
      auxTotals[i.techKey]=(auxTotals[i.techKey]||0)+1;
    } else {
      techTotals[i.techKey]=(techTotals[i.techKey]||0)+1;
    }
    if(!techDias[i.techKey])techDias[i.techKey]=new Set();
    techDias[i.techKey].add(i.day);
  });
  const sorted=Object.entries(techTotals).sort((a,b)=>b[1]-a[1]);
  if(!sorted.length){document.getElementById('badgeAnalise')?.classList.remove('hidden');return;}
  const totalOS=filtered.length, totalTechs=sorted.length, avgOS=totalOS/totalTechs;
  const best=sorted[0], worst=sorted[sorted.length-1];
  const bestNome=state.teamData[best[0]]?.originalName||best[0];
  const worstNome=state.teamData[worst[0]]?.originalName||worst[0];
  const [ys,ms]=state.activeMonthYear.split('-');
  const dIM=new Date(parseInt(ys),parseInt(ms),0).getDate();
  const dayMap={};filtered.forEach(i=>{dayMap[i.day]=(dayMap[i.day]||0)+1;});
  const dailyArr=Array.from({length:dIM},(_,i)=>dayMap[i+1]||0);
  const diasUteisProd=dailyArr.filter((v,i)=>{const dw=new Date(parseInt(ys),parseInt(ms)-1,i+1).getDay();return dw>=1&&dw<=5&&v>0;});
  const diasSabProd=dailyArr.filter((v,i)=>{const dw=new Date(parseInt(ys),parseInt(ms)-1,i+1).getDay();return dw===6&&v>0;});
  const workDays=dailyArr.filter(v=>v>0);
  const nonSunArr=dailyArr.filter((v,i)=>{const dw=new Date(parseInt(ys),parseInt(ms)-1,i+1).getDay();return dw!==0;});
  const nonSunMin=Math.min(...nonSunArr.filter(v=>v>0));
  const weakIdx=dailyArr.findIndex((v,i)=>{const dw=new Date(parseInt(ys),parseInt(ms)-1,i+1).getDay();return v===nonSunMin&&dw!==0;});
  const weakDay=weakIdx>=0?`Dia ${weakIdx+1} (${DN[new Date(parseInt(ys),parseInt(ms)-1,weakIdx+1).getDay()]})`:'—';
  const weakVal=weakIdx>=0?dailyArr[weakIdx]:0;
  const abovePct=sorted.filter(([k,v])=>{
    const tipo=state.teamData[k]?.tipo||'INSTALAÇÃO CIDADE';
    if(tipo==='AUXILIAR')return false;
    const m=state.appSettings.metasDiarias[tipo]||5;
    const bm=tipo==='TECNICO 12/36H'?15:24;
    return v/(m*bm)>=0.8;
  });
  const effPct=totalTechs>0?Math.round(abovePct.length/totalTechs*100):0;
  const bestMeta=state.appSettings.metasDiarias[state.teamData[best[0]]?.tipo||'INSTALAÇÃO CIDADE']||5;
  const bestBm=state.teamData[best[0]]?.tipo==='TECNICO 12/36H'?15:24;
  const bestCap=Math.round(best[1]/(bestMeta*bestBm)*100);
  document.getElementById('aEfic').textContent=effPct+'%';
  document.getElementById('aEficSub').textContent=`${abovePct.length} de ${totalTechs} técnicos acima de 80% da capacidade`;
  document.getElementById('aBest').textContent=bestNome;
  document.getElementById('aBestSub').textContent=`${best[1]} OS · ${bestCap}% da cap.`;
  document.getElementById('aWorst').textContent=worstNome;
  document.getElementById('aWorstSub').textContent=`${worst[1]} OS · ${Math.round(worst[1]/avgOS*100)}% da média geral`;
  document.getElementById('aWeakDay').textContent=weakDay;
  document.getElementById('aWeakDaySub').textContent=`${weakVal} OS · dia de menor produção`;
  const topList=sorted.slice(0,5).map(([k,v])=>{
    const tipo=state.teamData[k]?.tipo||'INSTALAÇÃO CIDADE';
    const m=state.appSettings.metasDiarias[tipo]||5;
    const bm=tipo==='TECNICO 12/36H'?15:24;
    const cap=Math.round(v/(m*bm)*100);
    const tipoLabel=TEAM_TYPES[tipo]||tipo;
    return`<div class="obs-item"><div class="obs-dot g"></div><span><b>${state.teamData[k]?.originalName||k}</b> <span class="obs-type-tag">${tipoLabel}</span> — ${v} OS <span class="obs-cap-tag ${cap>=100?'cap-ex':cap>=80?'cap-bom':''}">${cap}%</span></span></div>`;
  }).join('');
  document.getElementById('obsPositivos').innerHTML=topList||'<div class="empty" style="padding:12px;">Sem dados.</div>';
  const botList=sorted.filter(([k,v])=>v<avgOS*0.6).slice(0,5).map(([k,v])=>{
    const diff=Math.round((1-v/avgOS)*100);
    const tipo=state.teamData[k]?.tipo||'INSTALAÇÃO CIDADE';
    const tipoLabel=TEAM_TYPES[tipo]||tipo;
    const dias=techDias[k]?.size||0;
    return`<div class="obs-item"><div class="obs-dot r"></div><span><b>${state.teamData[k]?.originalName||k}</b> <span class="obs-type-tag">${tipoLabel}</span> — ${v} OS em ${dias} dias (${diff}% abaixo da média)</span></div>`;
  }).join('');
  document.getElementById('obsAtencao').innerHTML=botList||'<div class="obs-item"><div class="obs-dot g"></div><span>Todos os técnicos acima de 60% da média!</span></div>';
  const auxList=Object.entries(auxTotals).sort((a,b)=>b[1]-a[1]);
  const diasComOS=workDays.length;
  const diasUteisTotal=dailyArr.filter((_,i)=>{const dw=new Date(parseInt(ys),parseInt(ms)-1,i+1).getDay();return dw!==0&&dw!==6;}).length;
  const mediaUteis=diasUteisProd.length>0?Math.round(diasUteisProd.reduce((a,b)=>a+b,0)/diasUteisProd.length):0;
  const mediaSab=diasSabProd.length>0?Math.round(diasSabProd.reduce((a,b)=>a+b,0)/diasSabProd.length):0;
  const obs=[];
  obs.push(`Presença em <b>${Math.round(diasComOS/diasUteisTotal*100)}%</b> dos dias úteis (${diasUteisTotal} dias)`);
  obs.push(`Média de <b>${mediaUteis} OS/dia</b> em dias úteis${mediaSab>0?' · <b>'+mediaSab+' OS/dia</b> aos sábados':''}`);
  if(sorted.length>1) obs.push(`Amplitude: <b>${best[1]-worst[1]} OS</b> entre o melhor e o pior técnico`);
  if(auxList.length>0) obs.push(`<b>${auxList.length}</b> auxiliar(es) com OS finalizada(s) este período`);
  const diasSemOS=dailyArr.filter((v,i)=>{const dw=new Date(parseInt(ys),parseInt(ms)-1,i+1).getDay();return v===0&&dw>=1&&dw<=5;}).length;
  if(diasSemOS>0) obs.push(`<b>${diasSemOS}</b> dias úteis sem nenhuma OS registrada`);
  document.getElementById('obsRapidas').innerHTML=obs.map(o=>`<div class="obs-item"><div class="obs-dot y"></div><span>${o}</span></div>`).join('');
  const sug=[];
  if(worst[1]<avgOS*0.5) sug.push(`Verificar situação de <b>${worstNome}</b> — produção abaixo de 50% da média da equipe`);
  if(effPct<40) sug.push(`Apenas <b>${effPct}%</b> da equipe atingiu 80%+ da capacidade — revisar distribuição de OS`);
  else if(effPct<60) sug.push(`<b>${100-effPct}%</b> da equipe abaixo de 80% — identificar gargalos de atendimento`);
  if(weakVal<avgOS*0.3&&weakIdx>=0) sug.push(`Investigar a baixa produção no <b>${weakDay}</b> — possível falta de escala ou OS`);
  if(diasSemOS>2) sug.push(`Revisar escalonamento — <b>${diasSemOS} dias úteis</b> sem OS registrada`);
  const cidadesMap={};filtered.forEach(i=>{if(!cidadesMap[i.cidade])cidadesMap[i.cidade]={total:0,techs:new Set()};cidadesMap[i.cidade].total++;cidadesMap[i.cidade].techs.add(i.techKey);});
  const cidadesArr=Object.entries(cidadesMap).map(([c,d])=>({nome:c,media:d.total/d.techs.size}));
  if(cidadesArr.length>1){cidadesArr.sort((a,b)=>b.media-a.media);const diff=Math.round(cidadesArr[0].media-cidadesArr[cidadesArr.length-1].media);if(diff>20)sug.push(`Desbalanceamento entre polos: <b>${cidadesArr[0].nome}</b> produz ${diff} OS/técnico a mais que <b>${cidadesArr[cidadesArr.length-1].nome}</b>`);}
  if(!sug.length) sug.push('Operação estável — manter monitoramento semanal de produtividade');
  document.getElementById('obsSugestoes').innerHTML=sug.slice(0,4).map(s=>`<div class="obs-item"><div class="obs-dot b"></div><span>${s}</span></div>`).join('');

  document.getElementById('badgeAnalise')?.classList.remove('hidden');
  document.getElementById('mobBadgeAnalise')?.classList.remove('hidden');
}

function showDayDetails(day, cidade, techKey) {
  if (!state.globalRawData.length) return;

  const fc = state.selectedCityTab;
  const ft = document.getElementById('filterType')?.value || 'ALL';

  let filtered = state.globalRawData.filter(i =>
    i.monthStr === state.activeMonthYear &&
    (day === null || i.day === day) &&
    (fc === 'ALL' || i.cidade === fc) &&
    (ft === 'ALL' || i.tipo === ft)
  );

  if (cidade) {
    filtered = filtered.filter(i => i.cidade === cidade);
  }
  if (techKey) {
    filtered = filtered.filter(i => i.techKey === techKey);
  }

  if (!filtered.length) {
    showToast('Nenhuma O.S. encontrada para este dia.', 'info');
    return;
  }

  state.currentDayDetails = filtered;
  state.currentDayDetailsDay = day;

  let overlay = document.getElementById('dayDetailsModalOverlay');
  if (overlay) overlay.remove();
  overlay = document.createElement('div');
  overlay.id = 'dayDetailsModalOverlay';
  overlay.style.cssText = 'position:fixed;inset:0;z-index:9999;background:rgba(15,23,42,.6);display:flex;align-items:center;justify-content:center;backdrop-filter:blur(8px);animation:fadeIn .15s ease;padding:16px;';

  const formatTMA = (m) => m == null || isNaN(m) ? '—' : (m < 60 ? `${Math.round(m)}m` : `${Math.floor(m/60)}h ${Math.round(m%60)}m`);
  let totalTma = 0, countTma = 0;

  filtered.forEach(os => {
    const dIni = toDateTime(os.dtInicio), dFin = toDateTime(os.dtFinal);
    if (dIni && dFin && dFin >= dIni) {
      const tmaMin = (dFin - dIni) / 60000;
      if (tmaMin < 1440) { totalTma += tmaMin; countTma++; }
    }
  });

  const avgTmaStr = countTma > 0 ? formatTMA(totalTma / countTma) : '—';

  let subtitle = `Regional: ${escapeHtml(cidade || 'Todas')}`;
  if (techKey) subtitle += ` · Técnico: ${escapeHtml(state.teamData[techKey]?.originalName || techKey)}`;

  overlay.innerHTML = `
    <div style="background:var(--surface-card);border:1px solid var(--surface-border);border-radius:16px;max-width:950px;width:100%;max-height:85vh;display:flex;flex-direction:column;box-shadow:0 16px 48px rgba(0,0,0,.5);animation:scaleIn .15s ease;overflow:hidden;">
      <div style="padding:20px 24px;border-bottom:1px solid var(--surface-border);display:flex;justify-content:space-between;align-items:center;background:var(--surface-card2);flex-wrap:wrap;gap:12px;">
        <div><div style="font-size:16px;font-weight:800;color:var(--text-primary);">Detalhamento de O.S. — ${day !== null ? `Dia ${day}` : 'Mês Completo'}</div><div style="font-size:12px;color:var(--text-tertiary);margin-top:2px;">${subtitle} · <b style="color:var(--text-secondary);">${filtered.length} O.S.</b> · TMA: <b style="color:var(--blue);">${avgTmaStr}</b></div></div>
        <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;">
          <input type="text" id="dayDetailsSearch" placeholder="🔍 Buscar cliente, OS, assunto..." oninput="App.filterDayDetails()" style="padding:7px 12px;border-radius:8px;border:1px solid var(--surface-border);background:var(--surface-card);color:var(--text-primary);font-size:12px;font-family:var(--font);outline:none;min-width:220px;">
          <select id="dayDetailsSort" onchange="App.filterDayDetails()" style="padding:7px 12px;border-radius:8px;border:1px solid var(--surface-border);background:var(--surface-card);color:var(--text-primary);font-size:12px;font-family:var(--font);outline:none;">
            <option value="crono">Ordem Cronológica</option>
            <option value="assunto">Por Assunto</option>
            <option value="tma">Maior TMA</option>
          </select>
          <button id="closeDayDetailsModal" style="background:var(--hover-white-04);border:1px solid var(--surface-border);border-radius:8px;color:var(--text-tertiary);font-size:14px;cursor:pointer;width:32px;height:32px;display:flex;align-items:center;justify-content:center;transition:all var(--dur-fast);">✖</button>
        </div>
      </div>
      <div style="overflow-y:auto;padding:0;flex:1;"><table style="width:100%;border-collapse:collapse;text-align:left;"><thead style="background:var(--surface-card2);position:sticky;top:0;box-shadow:0 1px 2px rgba(0,0,0,0.05);z-index:2;"><tr><th style="padding:12px;font-size:10px;color:var(--text-tertiary);text-transform:uppercase;font-family:var(--mono);">Protocolo</th>${day === null ? `<th style="padding:12px;font-size:10px;color:var(--text-tertiary);text-transform:uppercase;font-family:var(--mono);">Data</th>` : ''}<th style="padding:12px;font-size:10px;color:var(--text-tertiary);text-transform:uppercase;font-family:var(--mono);">Técnico</th><th style="padding:12px;font-size:10px;color:var(--text-tertiary);text-transform:uppercase;font-family:var(--mono);">Assunto & Diagnóstico</th><th style="padding:12px;font-size:10px;color:var(--text-tertiary);text-transform:uppercase;font-family:var(--mono);">TMA</th><th style="padding:12px;font-size:10px;color:var(--text-tertiary);text-transform:uppercase;font-family:var(--mono);">Cliente/Login</th></tr></thead><tbody id="dayDetailsBody"></tbody></table></div>
    </div>`;
  document.body.appendChild(overlay);
  const closeBtn = document.getElementById('closeDayDetailsModal');
  closeBtn.onclick = () => overlay.remove();
  closeBtn.onmouseover = () => { closeBtn.style.background = 'var(--hover-white-08)'; closeBtn.style.color = 'var(--text-primary)'; };
  closeBtn.onmouseout = () => { closeBtn.style.background = 'var(--hover-white-04)'; closeBtn.style.color = 'var(--text-tertiary)'; };
  overlay.onclick = e => { if (e.target === overlay) overlay.remove(); };

  filterDayDetails();
}

function filterDayDetails() {
  const tbody = document.getElementById('dayDetailsBody');
  if (!tbody) return;
  const search = (document.getElementById('dayDetailsSearch')?.value || '').toLowerCase();
  const sort = document.getElementById('dayDetailsSort')?.value || 'crono';
  const day = state.currentDayDetailsDay;

  let list = [...state.currentDayDetails];

  if (search) {
    list = list.filter(os =>
      (os.osId || '').toLowerCase().includes(search) ||
      (os.cliente || '').toLowerCase().includes(search) ||
      (os.login || '').toLowerCase().includes(search) ||
      (os.assunto || '').toLowerCase().includes(search) ||
      (os.diagnostico || '').toLowerCase().includes(search) ||
      (os.nomeOriginal || os.nome || '').toLowerCase().includes(search)
    );
  }

  if (sort === 'crono') {
    list.sort((a, b) => (day === null ? a.day - b.day : 0) || (a.nomeOriginal || a.nome).localeCompare(b.nomeOriginal || b.nome));
  } else if (sort === 'assunto') {
    list.sort((a, b) => (a.assunto || '').localeCompare(b.assunto || '') || (day === null ? a.day - b.day : 0));
  } else if (sort === 'tma') {
    list.sort((a, b) => {
      const dIniA = toDateTime(a.dtInicio), dFinA = toDateTime(a.dtFinal);
      const dIniB = toDateTime(b.dtInicio), dFinB = toDateTime(b.dtFinal);
      const tA = (dIniA && dFinA && dFinA >= dIniA) ? (dFinA - dIniA) : -1;
      const tB = (dIniB && dFinB && dFinB >= dIniB) ? (dFinB - dIniB) : -1;
      return tB - tA;
    });
  }

  const getIcon = (txt) => {
    const t = (txt||'').toLowerCase();
    if(t.includes('troca') || t.includes('remoção')) return '🔄';
    if(t.includes('conexão') || t.includes('sinal') || t.includes('fibra')) return '📡';
    if(t.includes('login') || t.includes('senha')) return '🔑';
    if(t.includes('apps') || t.includes('streaming')) return '📺';
    return '🔧';
  };

  const formatTMA = (m) => m == null || isNaN(m) ? '—' : (m < 60 ? `${Math.round(m)}m` : `${Math.floor(m/60)}h ${Math.round(m%60)}m`);

  tbody.innerHTML = list.length ? list.map((os, idx) => {
    const dIni = toDateTime(os.dtInicio), dFin = toDateTime(os.dtFinal);
    let tmaMin = null;
    if (dIni && dFin && dFin >= dIni) { tmaMin = (dFin - dIni) / 60000; if (tmaMin >= 1440) tmaMin = null; }
    return `
    <tr style="border-bottom:1px solid var(--surface-border);background:${idx % 2 === 0 ? 'transparent' : 'var(--surface-card2)'}">
      <td style="padding:10px 12px;font-family:var(--mono);font-size:11px;color:var(--accent);white-space:nowrap;">${escapeHtml(os.osId)}</td>
      ${day === null ? `<td style="padding:10px 12px;font-family:var(--mono);font-size:11px;color:var(--text-secondary);white-space:nowrap;">${os.day}/${os.monthStr.split('-')[1]}</td>` : ''}
      <td style="padding:10px 12px;font-size:12px;white-space:nowrap;">
        <b style="color:var(--text-primary);display:block;">${escapeHtml(os.nomeOriginal || os.nome)}</b>
        <span style="font-size:10px;color:var(--text-tertiary);font-family:var(--mono);">${escapeHtml(TEAM_TYPES[os.tipo] || os.tipo)}</span>
      </td>
      <td style="padding:10px 12px;font-size:11px;line-height:1.4;">
        <span style="color:var(--text-secondary);">${escapeHtml(os.assunto)}</span><br>
        <span style="font-size:10px;color:var(--text-tertiary);font-style:italic;">${escapeHtml(os.diagnostico)}</span>
      </td>
      <td style="padding:10px 12px;font-family:var(--mono);font-size:11px;color:var(--blue);font-weight:700;">${tmaMin !== null ? formatTMA(tmaMin) : '—'}</td>
      <td style="padding:10px 12px;font-size:11px;color:var(--text-secondary);">${escapeHtml(os.cliente || os.login)}</td>
    </tr>
  `}).join('') : `<tr><td colspan="${day === null ? '6' : '5'}" style="padding:32px;text-align:center;color:var(--text-tertiary);font-size:12px;">Nenhuma O.S. encontrada para esta busca.</td></tr>`;
}

function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function formatDateDisplay(dateStr) {
  if (!dateStr) return '—';
  const dt = new Date(dateStr);
  if (Number.isNaN(dt.getTime())) return String(dateStr);
  return dt.toLocaleDateString('pt-BR');
}

function getMostFrequentEntry(mapObj) {
  const entries = Object.entries(mapObj || {});
  if (!entries.length) return '';
  entries.sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]));
  return entries[0][0];
}

function normalizeText(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toUpperCase();
}

function normalizeStatus(value) {
  return normalizeText(value);
}

function toDateTime(value) {
  if (!value) return null;
  const dt = new Date(value);
  return Number.isNaN(dt.getTime()) ? null : dt;
}

function formatDateTimeDisplay(value) {
  const dt = toDateTime(value);
  if (!dt) return '—';
  return dt.toLocaleString('pt-BR');
}

function roundToOne(value) {
  return Math.round((Number(value) || 0) * 10) / 10;
}

function resetRecurrenceCache() {
  state.recurrenceAnalysis = null;
  state.recurrenceAnalysisKey = '';
}

function diffDays(a, b) {
  const aTime = toDateTime(a)?.getTime();
  const bTime = toDateTime(b)?.getTime();
  if (!aTime || !bTime) return 0;
  return Math.round((bTime - aTime) / 86400000);
}

function getRecurrenceSourceRows() {
  if (state.rawExcelCache && state.rawExcelCache.length) return state.rawExcelCache;
  if (state.globalRawData && state.globalRawData.length) return state.globalRawData;
  return [];
}

function getRecurrenceReferenceEndDate() {
  const rowsInMonth = getRecurrenceSourceRows().filter(os => os.monthStr === state.activeMonthYear);
  const datedRows = rowsInMonth
    .map(os => toDateTime(os.dateTimeStr || os.dateStr))
    .filter(Boolean)
    .sort((a, b) => a - b);
  if (datedRows.length) return datedRows[datedRows.length - 1];
  if (state.activeMonthYear) {
    const [year, month] = state.activeMonthYear.split('-').map(Number);
    return new Date(year, month, 0, 23, 59, 59, 999);
  }
  return new Date();
}

function getSelectedRecurrenceTechnicians() {
  const byKey = new Set(Object.keys(state.teamData));
  const byOriginal = new Set(Object.values(state.teamData).map(t => normalizeText(t.originalName)));
  return {
    byKey,
    byOriginal,
    hasTeamFilter: byKey.size > 0 || byOriginal.size > 0,
    list: Object.values(state.teamData).map(t => t.originalName).sort()
  };
}

function getRecurrenceIdentityStrict(os) {
  const login = String(os.login || '').trim();
  if (login) {
    return {
      key: `LOGIN:${login.toUpperCase()}`,
      login,
      cliente: String(os.cliente || '').trim()
    };
  }
  const cliente = String(os.cliente || '').trim();
  if (cliente) {
    return {
      key: `CLIENTE:${normalizeText(cliente)}`,
      login: '',
      cliente
    };
  }
  return null;
}

function classifyRecurrence(totalOS, diasEntre, sameSubjectRepeated) {
  if (totalOS >= 4 && diasEntre <= 15) return 'CRITICA';
  if (totalOS >= 3 && diasEntre <= 20) return 'ALTA';
  if (totalOS >= 2 && sameSubjectRepeated) return 'ALTA';
  if (totalOS >= 2 && diasEntre <= 30) return 'MEDIA';
  if (totalOS >= 2 && diasEntre > 30) return 'BAIXA';
  return 'BAIXA';
}

function getAlertMeta(classificacao) {
  if (classificacao === 'CRITICA') return { tipo: 'CRITICO', prioridade: 1 };
  if (classificacao === 'ALTA') return { tipo: 'ALERTA', prioridade: 2 };
  if (classificacao === 'MEDIA') return { tipo: 'ATENCAO', prioridade: 3 };
  return { tipo: 'OBSERVACAO', prioridade: 4 };
}

function getClassificationLabel(classificacao) {
  const labels = { CRITICA: 'CRITICA', ALTA: 'ALTA', MEDIA: 'MEDIA', BAIXA: 'BAIXA' };
  return labels[classificacao] || classificacao;
}

function getExcludedRecurrenceSubjects() {
  return [
    "Instalação Fibra Urbana",
    "Instalação Fibra Rural",
    "Instalação Rádio",
    "Retorno de Instalação Fibra Urbana",
    "Retorno de Instalação Fibra Rural",
    "Retorno de Instalação Rádio",
    "Alteração de Endereço Fibra Urbana",
    "Alteração de Endereço Fibra Rural",
    "Alteração de Endereço Rádio",
    "Retorno de Alteração de Endereço Fibra Urbana",
    "Retorno de Alteração de Endereço Fibra Rural",
    "Retorno de Alteração de Endereço Rádio",
    "Alteração da Tecnologia para Fibra",
    "Alteração da Tecnologia para Rádio",
    "Retorno de Alteração de Tecnologia para Fibra",
    "Retorno de Alteração de Tecnologia Rádio para Fibra",
    "Retorno de Alteração de Tecnologia Fibra para Rádio",
    "Remoção de Flashman",
    "Manutenção Preventiva Operacional",
    "Viabilidade"
  ].map(normalizeText);
}

function isExcludedRecurrenceSubject(subject, excludedSubjects) {
  const assuntoNorm = normalizeText(subject || '');
  return excludedSubjects.some(ex => assuntoNorm.includes(ex));
}

function isExcludedRecurrenceDiagnosis(diagnosis) {
  const diagnosticoNorm = normalizeText(diagnosis || '');
  if (!diagnosticoNorm) return false;
  return diagnosticoNorm.includes('INSTALACAO') || diagnosticoNorm.includes('PASSAGEM DE CABO');
}

function isViabilitySubject(subject) {
  return normalizeText(subject || '').includes('VIABILIDADE');
}

function buildRecurrenceRecommendations(analysis) {
  const recs = new Set();
  const assuntos = analysis.reincidentes.map(item => normalizeText(item.assunto_principal));
  if (assuntos.some(a => a.includes('CONECTOR') || a.includes('CABO'))) {
    recs.add('Revisar procedimento de instalação de conectores');
  }
  if (assuntos.some(a => a.includes('ROTEADOR') || a.includes('ONU'))) {
    recs.add('Verificar lote de equipamentos e firmware');
  }
  if ((analysis.resumo?.percentual_reincidencia || 0) > 30) {
    recs.add('Implementar checklist pos-atendimento');
  }
  if ((analysis.resumo?.media_dias_entre_os || 0) < 10 && analysis.resumo?.media_dias_entre_os > 0) {
    recs.add('Otimizar agendamento e preventivas');
  }
  if (!recs.size) {
    recs.add('Manter monitoramento continuo dos clientes reincidentes');
  }
  return Array.from(recs);
}

function buildRecurrenceAnalysisData() {
  const settings = state.recurrenceSettings || { diasAnalise: 30, periodoBuscaDias: 90, minimoOsParaReincidencia: 2 };
  const selectedTechs = getSelectedRecurrenceTechnicians();
  const sourceRows = getRecurrenceSourceRows();
  if (state.uploadMeta && state.uploadMeta.hasStatus === false) {
    return {
      success: false,
      message: 'A planilha carregada nao possui coluna de status. A analise exige Status = Finalizada.'
    };
  }

  const EXCLUDED_RECURRENCE_SUBJECTS = [
    "Instalação Fibra Urbana",
    "Instalação Fibra Rural",
    "Instalação Rádio",
    "Retorno de Instalação Fibra Urbana",
    "Retorno de Instalação Fibra Rural",
    "Retorno de Instalação Rádio",
    "Alteração de Endereço Fibra Urbana",
    "Alteração de Endereço Fibra Rural",
    "Alteração de Endereço Rádio",
    "Retorno de Alteração de Endereço Fibra Urbana",
    "Retorno de Alteração de Endereço Fibra Rural",
    "Retorno de Alteração de Endereço Rádio",
    "Alteração da Tecnologia para Fibra",
    "Alteração da Tecnologia para Rádio",
    "Retorno de Alteração de Tecnologia para Fibra",
    "Retorno de Alteração de Tecnologia Rádio para Fibra",
    "Retorno de Alteração de Tecnologia Fibra para Rádio",
    "Remoção de Flashman",
    "Manutenção Preventiva Operacional",
    "Viabilidade"
  ].map(normalizeText);

  const cacheKey = JSON.stringify({
    month: state.activeMonthYear,
    diasAnalise: settings.diasAnalise,
    periodoBuscaDias: settings.periodoBuscaDias,
    minimoOsParaReincidencia: settings.minimoOsParaReincidencia,
    rawCount: sourceRows.length,
    teamCount: Object.keys(state.teamData).length
  });
  if (state.recurrenceAnalysis && state.recurrenceAnalysisKey === cacheKey) {
    return state.recurrenceAnalysis;
  }

  const allRows = sourceRows.map(os => ({
    ...os,
    _date: toDateTime(os.dateTimeStr || os.dateStr),
    _techKey: limparNome(os.nomeOriginal || os.nome || '')
  })).filter(os => os._date);

  const referenceEnd = getRecurrenceReferenceEndDate();
  const periodStart = new Date(referenceEnd);
  periodStart.setDate(periodStart.getDate() - settings.periodoBuscaDias);

  const rowsFinalizadas = allRows.filter(os => normalizeStatus(os.status) === 'FINALIZADA');
  
  // NOVO: Removemos as O.S. internas. Apenas OS finalizadas por equipes cadastradas seguem para o cálculo.
  const rowsEquipesCadastradas = selectedTechs.hasTeamFilter
    ? rowsFinalizadas.filter(os => selectedTechs.byKey.has(os._techKey) || selectedTechs.byOriginal.has(normalizeText(os.nomeOriginal || os.nome || '')))
    : rowsFinalizadas;
  const rowsInSelectedMonth = rowsEquipesCadastradas.filter(os => os.monthStr === state.activeMonthYear && String(os.filial).trim() !== '5');
  const seedRows = rowsInSelectedMonth.filter(os => !isViabilitySubject(os.assunto));

  const seedGroups = {};
  seedRows.forEach(os => {
    const identity = getRecurrenceIdentityStrict(os);
    if (!identity) return;
    if (!seedGroups[identity.key]) seedGroups[identity.key] = [];
    seedGroups[identity.key].push(os);
  });

  const rowsByIdentity = {};
  rowsInSelectedMonth.forEach(os => {
    const identity = getRecurrenceIdentityStrict(os);
    if (!identity) return;
    if (!rowsByIdentity[identity.key]) rowsByIdentity[identity.key] = [];
    rowsByIdentity[identity.key].push(os);
  });
  Object.keys(rowsByIdentity).forEach(identityKey => {
    rowsByIdentity[identityKey].sort((a, b) => a._date - b._date);
  });

  const totalClientesAnalisados = Object.keys(seedGroups).length;
  const reincidentes = [];
  const assuntoStats = {};
  const tecnicoStats = {};
  let totalOsReincidentes = 0;
  let somaMediasDias = 0;

  Object.entries(seedGroups).forEach(([identityKey, rows]) => {
    const seedSorted = rows.slice().sort((a, b) => a._date - b._date);
    const primeiraOS = seedSorted[0];
    if (!primeiraOS) return;
    const identity = getRecurrenceIdentityStrict(primeiraOS);
    if (!identity) return;
    const filial = primeiraOS.filial || 'NÃO INFORMADA';

    const dataLimite = new Date(primeiraOS._date);
    dataLimite.setDate(dataLimite.getDate() + settings.diasAnalise);

    const allRowsForClient = (rowsByIdentity[identityKey] || []).filter(os =>
      os._date >= primeiraOS._date && os._date <= dataLimite
    );

    const validRowsForClient = [];
    for (let index = 0; index < allRowsForClient.length; index++) {
      const os = allRowsForClient[index];
      if (index === 0 || !isExcludedRecurrenceSubject(os.assunto, EXCLUDED_RECURRENCE_SUBJECTS)) {
        validRowsForClient.push(os);
      }
    }

    if (validRowsForClient.length < settings.minimoOsParaReincidencia) return;
    const statsRowsForClient = validRowsForClient.filter(os => (
      !isExcludedRecurrenceSubject(os.assunto, EXCLUDED_RECURRENCE_SUBJECTS) &&
      !isExcludedRecurrenceDiagnosis(os.diagnostico)
    ));
    const rowsForSubjectAndDiagnosis = statsRowsForClient;

    const ultimaOS = validRowsForClient[validRowsForClient.length - 1];
    const diasEntre = diffDays(primeiraOS._date, ultimaOS._date);
    const intervalos = [];
    for (let i = 1; i < validRowsForClient.length; i++) {
      intervalos.push(diffDays(validRowsForClient[i - 1]._date, validRowsForClient[i]._date));
    }
    const mediaDias = intervalos.length ? roundToOne(intervalos.reduce((sum, val) => sum + val, 0) / intervalos.length) : 0;
    const assuntos = rowsForSubjectAndDiagnosis.map(os => os.assunto || 'Sem assunto');
    const assuntoCounts = {};
    const diagnosticoCounts = {};
    const diagStats = {};
    assuntos.forEach(assunto => { assuntoCounts[assunto] = (assuntoCounts[assunto] || 0) + 1; });
    rowsForSubjectAndDiagnosis.forEach((os, idx) => {
      const d = os.diagnostico || 'Sem diagnostico';
      diagnosticoCounts[d] = (diagnosticoCounts[d] || 0) + 1;
      if (!diagStats[d]) diagStats[d] = { count: 0, lastIdx: idx };
      diagStats[d].count++;
      diagStats[d].lastIdx = idx; // Sempre grava o índice mais recente
    });
    const diagEntries = Object.entries(diagStats).sort((a, b) => b[1].count - a[1].count || b[1].lastIdx - a[1].lastIdx);
    const assuntoPrincipal = getMostFrequentEntry(assuntoCounts) || 'Sem assunto';
    let pIdx = 0;
    if (diagEntries.length > 1 && normalizeText(diagEntries[0][0]).includes('DUPLICADA')) pIdx = 1;
    const diagnosticoPrincipal = diagEntries.length ? diagEntries[pIdx][0] : 'Sem diagnostico';
    const sameSubjectRepeated = new Set(assuntos.map(normalizeText)).size === 1;
    const classificacao = classifyRecurrence(validRowsForClient.length, diasEntre, sameSubjectRepeated);
    const originTech = String(primeiraOS.nomeOriginal || primeiraOS.nome || '—').trim();
    const tecnicosEnvolvidos = [originTech];
    
    const subjectOrigin = normalizeText(primeiraOS.assunto || '');
    const isAtivacao = subjectOrigin.includes('INSTALACAO') ||
                       subjectOrigin.includes('ALTERACAO DE ENDERECO') ||
                       subjectOrigin.includes('ALTERACAO DA TECNOLOGIA') ||
                       subjectOrigin.includes('MUDANCA') ||
                       subjectOrigin.includes('VIABILIDADE');
    const origemReincidencia = isAtivacao ? 'ATIVAÇÃO' : 'SUPORTE';

    const historicoOS = validRowsForClient.map((os, idx) => ({
      id: os.osId || '—',
      data: formatDateTimeDisplay(os._date),
      assunto: os.assunto || '',
      diagnostico: os.diagnostico || 'Sem diagnostico',
      tecnico: String(os.nomeOriginal || os.nome || '—').trim(),
      tipo_equipe: TEAM_TYPES[state.teamData[os._techKey]?.tipo] || state.teamData[os._techKey]?.tipo || '—',
      contrato: os.contrato || '—',
      dias_apos_primeira: diffDays(primeiraOS._date, os._date),
      is_origem: idx === 0,
      _uid: os._uid
    }));
    const diagnosticosResumo = Object.entries(diagnosticoCounts)
      .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))
      .map(([diagnostico, quantidade]) => ({ diagnostico, quantidade }));

    const reincidente = {
      login: identity.login || '',
      cliente: identity.cliente || '',
      filial: filial,
      origem_reincidencia: origemReincidencia,
      primeira_os: {
        id: primeiraOS.osId || '—',
        data: formatDateTimeDisplay(primeiraOS._date),
        assunto: primeiraOS.assunto || '',
        tecnico: String(primeiraOS.nomeOriginal || primeiraOS.nome || '—').trim()
      },
      total_os_periodo: validRowsForClient.length,
      dias_entre_primeira_e_ultima: diasEntre,
      media_dias_entre_os: mediaDias,
      classificacao,
      assunto_principal: assuntoPrincipal,
      diagnostico_principal: diagnosticoPrincipal,
      diagnosticos_resumo: diagnosticosResumo,
      tecnicos_envolvidos: tecnicosEnvolvidos,
      historico_os: historicoOS
    };

    reincidentes.push(reincidente);
    totalOsReincidentes += validRowsForClient.length;
    somaMediasDias += mediaDias;

    rowsForSubjectAndDiagnosis.forEach(os => {
      const assuntoKey = os.assunto || 'Sem assunto';
      if (!assuntoStats[assuntoKey]) assuntoStats[assuntoKey] = { quantidade: 0, intervals: [] };
      assuntoStats[assuntoKey].quantidade++;
    });

    Object.entries(assuntoCounts).forEach(([assunto]) => {
      const rowsMesmoAssunto = rowsForSubjectAndDiagnosis.filter(os => (os.assunto || 'Sem assunto') === assunto);
      for (let i = 1; i < rowsMesmoAssunto.length; i++) {
        assuntoStats[assunto].intervals.push(diffDays(rowsMesmoAssunto[i - 1]._date, rowsMesmoAssunto[i]._date));
      }
    });

    const techName = originTech;
    if (!tecnicoStats[techName]) tecnicoStats[techName] = { total_os_reincidentes: 0, clientes: new Set() };
    tecnicoStats[techName].total_os_reincidentes += validRowsForClient.length;
    tecnicoStats[techName].clientes.add(identityKey);
  });

  reincidentes.sort((a, b) => {
    const priorityDiff = (getAlertMeta(a.classificacao).prioridade - getAlertMeta(b.classificacao).prioridade);
    if (priorityDiff !== 0) return priorityDiff;
    return b.total_os_periodo - a.total_os_periodo || a.cliente.localeCompare(b.cliente);
  });

  const resumo = {
    total_clientes_analisados: totalClientesAnalisados,
    total_reincidentes: reincidentes.length,
    percentual_reincidencia: totalClientesAnalisados ? roundToOne((reincidentes.length / totalClientesAnalisados) * 100) : 0,
    media_dias_entre_os: reincidentes.length ? roundToOne(somaMediasDias / reincidentes.length) : 0
  };

  const estatisticasPorAssunto = Object.entries(assuntoStats)
    .map(([assunto, data]) => ({
      assunto,
      quantidade: data.quantidade,
      percentual: totalOsReincidentes ? roundToOne((data.quantidade / totalOsReincidentes) * 100) : 0,
      tempo_medio_resolucao_dias: data.intervals.length ? roundToOne(data.intervals.reduce((sum, val) => sum + val, 0) / data.intervals.length) : 0
    }))
    .sort((a, b) => b.quantidade - a.quantidade || a.assunto.localeCompare(b.assunto));

  const estatisticasPorTecnico = Object.entries(tecnicoStats)
    .map(([tecnico, data]) => ({
      tecnico,
      total_os_reincidentes: data.total_os_reincidentes,
      clientes_reincidentes_atendidos: data.clientes.size,
      taxa_envolvimento: totalOsReincidentes ? roundToOne((data.total_os_reincidentes / totalOsReincidentes) * 100) : 0
    }))
    .sort((a, b) => b.total_os_reincidentes - a.total_os_reincidentes || a.tecnico.localeCompare(b.tecnico));

  const alertas = reincidentes.map(item => {
    const alertMeta = getAlertMeta(item.classificacao);
    let mensagem = `${item.total_os_periodo} OS em ${item.dias_entre_primeira_e_ultima} dias - Monitoramento recomendado`;
    if (alertMeta.tipo === 'CRITICO') {
      mensagem = `${item.total_os_periodo} OS em ${item.dias_entre_primeira_e_ultima} dias - Analise de causa raiz URGENTE`;
    } else if (alertMeta.tipo === 'ALERTA') {
      mensagem = `${item.total_os_periodo} OS em ${item.dias_entre_primeira_e_ultima} dias - Acompanhamento prioritario`;
    } else if (alertMeta.tipo === 'ATENCAO') {
      mensagem = `${item.total_os_periodo} OS em ${item.dias_entre_primeira_e_ultima} dias - Monitorar proximas OS`;
    }
    return {
      tipo: alertMeta.tipo,
      cliente: item.login || item.cliente,
      mensagem,
      prioridade: alertMeta.prioridade
    };
  }).sort((a, b) => a.prioridade - b.prioridade || a.cliente.localeCompare(b.cliente));

  const analysis = {
    success: true,
    filtro_aplicado: {
      tecnicos: selectedTechs.list,
      dias_analise: settings.diasAnalise,
      periodo_busca_dias: settings.periodoBuscaDias,
      minimo_os_para_reincidencia: settings.minimoOsParaReincidencia,
      mes_referencia: state.activeMonthYear,
      data_fim_referencia: referenceEnd.toISOString()
    },
    processado_em: new Date().toISOString(),
    resumo,
    reincidentes,
    estatisticas_por_assunto: estatisticasPorAssunto,
    estatisticas_por_tecnico: estatisticasPorTecnico,
    alertas,
    recomendacoes: buildRecurrenceRecommendations({ resumo, reincidentes })
  };

  state.recurrenceAnalysis = analysis;
  state.recurrenceAnalysisKey = cacheKey;
  return analysis;
}

function getRecorrenciaViewData() {
  return buildRecurrenceAnalysisData();
}

let _lastRecurrenceRenderState = '';
function renderRecorrenciaClientes() {
  const container = document.getElementById('recorrenciaSection');
  if (!container) return;

  if (!getRecurrenceSourceRows().length) {
    container.innerHTML = '<div class="empty-state-card">Carregue uma planilha para analisar reincidencia por login ou cliente.</div>';
    return;
  }

  const filialSelectEl = document.getElementById('recFilialFilter');
  const origemSelectEl = document.getElementById('recOrigemFilter');
  const classSelectEl = document.getElementById('recClassFilter');
  const searchInputEl = document.getElementById('recSearchFilter');
  const techInputEl = document.getElementById('recTecnicoFilter');

  const currentStateStr = JSON.stringify({
    key: state.recurrenceAnalysisKey,
    f: filialSelectEl ? filialSelectEl.value : 'ALL',
    o: origemSelectEl ? origemSelectEl.value : 'ALL',
    c: classSelectEl ? classSelectEl.value : 'ALL',
    t: techInputEl ? techInputEl.value.trim() : 'ALL',
    s: searchInputEl ? searchInputEl.value : ''
  });

  if (_lastRecurrenceRenderState === currentStateStr && container.querySelector('.analise-grid')) {
    return; 
  }

  const analysis = getRecorrenciaViewData();
  if (!analysis?.success) {
    container.innerHTML = `<div class="empty-state-card">${escapeHtml(analysis?.message || 'Nao foi possivel processar a analise.')}</div>`;
    return;
  }

  _lastRecurrenceRenderState = currentStateStr;

  const settings = state.recurrenceSettings;
  const reincidentes = analysis.reincidentes || [];
  const alertas = analysis.alertas || [];
  const estatisticasPorAssunto = analysis.estatisticas_por_assunto || [];
  const estatisticasPorTecnico = analysis.estatisticas_por_tecnico || [];
  const resumo = analysis.resumo || {};

  const filiais = [...new Set(reincidentes.map(i => i.filial).filter(f => f && f !== 'NÃO INFORMADA'))].sort();
  
  const selectedFilial = filialSelectEl ? filialSelectEl.value : 'ALL';
  const selectedOrigem = origemSelectEl ? origemSelectEl.value : 'ALL';
  const selectedClass = classSelectEl ? classSelectEl.value : 'ALL';
  let selectedTech = techInputEl ? techInputEl.value.trim() : 'ALL';
  if (selectedTech === '') selectedTech = 'ALL';
  const searchVal = searchInputEl ? searchInputEl.value.trim().toLowerCase() : '';

  const preFilteredReincidentes = reincidentes.filter(i => {
    if (selectedFilial !== 'ALL' && i.filial !== selectedFilial) return false;
    if (selectedOrigem !== 'ALL' && i.origem_reincidencia !== selectedOrigem) return false;
    if (selectedClass !== 'ALL' && i.classificacao !== selectedClass) return false;
    if (searchVal) {
      const term = searchVal;
      const inLogin = (i.login || '').toLowerCase().includes(term);
      const inCliente = (i.cliente || '').toLowerCase().includes(term);
      const inAssunto = (i.assunto_principal || '').toLowerCase().includes(term);
      const inDiag = (i.diagnostico_principal || '').toLowerCase().includes(term);
      const inHistory = (i.historico_os || []).some(os =>
        (os.assunto || '').toLowerCase().includes(term) ||
        (os.diagnostico || '').toLowerCase().includes(term)
      );
      if (!inLogin && !inCliente && !inAssunto && !inDiag && !inHistory) return false;
    }
    return true;
  });

  // Atualiza as opções do Datalist apenas com os técnicos que sobraram na filtragem
  const tecnicosOrigem = [...new Set(preFilteredReincidentes.map(i => i.primeira_os.tecnico).filter(t => t && t !== '—'))].sort();
  if (selectedTech !== 'ALL' && !tecnicosOrigem.includes(selectedTech)) {
    selectedTech = 'ALL';
  }

  const filteredReincidentes = preFilteredReincidentes.filter(i => {
    if (selectedTech !== 'ALL' && i.primeira_os.tecnico !== selectedTech) return false;
    return true;
  });
  const filteredAlertas = alertas.filter(a => filteredReincidentes.some(r => (r.login || r.cliente) === a.cliente));

  const alertsHtml = filteredAlertas.length
    ? filteredAlertas.slice(0, 8).map(alerta => `<div class="obs-item"><div class="obs-dot ${alerta.prioridade <= 2 ? 'r' : alerta.prioridade === 3 ? 'y' : 'b'}"></div><span><b>${escapeHtml(alerta.tipo)}</b> · ${escapeHtml(alerta.cliente)} — ${escapeHtml(alerta.mensagem)}</span></div>`).join('')
    : '<div class="empty">Nenhum alerta gerado.</div>';

  const filialStats = {};
  filteredReincidentes.forEach(item => {
    const f = item.filial || 'NÃO INFORMADA';
    if (!filialStats[f]) filialStats[f] = { ativacao: 0, suporte: 0, total: 0 };
    if (item.origem_reincidencia === 'ATIVAÇÃO') filialStats[f].ativacao++;
    else filialStats[f].suporte++;
    filialStats[f].total++;
  });
  
  const filialStatsHtml = Object.keys(filialStats).length
    ? Object.entries(filialStats).sort((a, b) => b[1].total - a[1].total).map(([f, data]) => `
      <div style="display:flex;justify-content:space-between;align-items:center;padding:8px 0;border-bottom:1px dashed var(--surface-border);font-size:11px;">
        <span style="font-weight:600;color:var(--text-secondary);flex:1;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;padding-right:8px;" title="${escapeHtml(f)}">${escapeHtml(f)}</span>
        <div style="display:flex;gap:12px;font-family:var(--mono);text-align:center;flex-shrink:0;">
          <div title="Após Ativação" style="min-width:36px;"><span style="display:block;font-size:8px;color:var(--text-tertiary);">ATIV.</span><span style="color:#0ea5e9;font-weight:700;">${data.ativacao}</span></div>
          <div title="Após Suporte" style="min-width:36px;"><span style="display:block;font-size:8px;color:var(--text-tertiary);">SUP.</span><span style="color:#f59e0b;font-weight:700;">${data.suporte}</span></div>
          <div title="Total" style="min-width:36px;"><span style="display:block;font-size:8px;color:var(--text-tertiary);">TOTAL</span><span style="color:var(--text-primary);font-weight:800;">${data.total}</span></div>
        </div>
      </div>`).join('')
    : '<div class="empty">Sem dados de filiais para os filtros atuais.</div>';

  // NOVO: Recalcular as estatísticas de assuntos e técnicos em tempo real baseando-se nos filtros da tela
  const filteredAssuntoStats = {};
  const filteredTecnicoStats = {};
  const filteredDiagnosticoStats = {};
  const excludedRecurrenceSubjects = getExcludedRecurrenceSubjects();
  let filteredTotalOs = 0;
  let filteredTotalRetrabalho = 0;

  filteredReincidentes.forEach(item => {
    filteredTotalOs += item.historico_os.length;
    const assuntoGroups = {};

    item.historico_os.forEach(os => {
      const assuntoKey = os.assunto || 'Sem assunto';
      if (!filteredAssuntoStats[assuntoKey]) filteredAssuntoStats[assuntoKey] = { quantidade: 0, intervals: [] };
      filteredAssuntoStats[assuntoKey].quantidade++;
      
      if (!assuntoGroups[assuntoKey]) assuntoGroups[assuntoKey] = [];
      assuntoGroups[assuntoKey].push(os);

      const techName = os.tecnico || '—';
      if (!filteredTecnicoStats[techName]) filteredTecnicoStats[techName] = { total_os: 0, clientes: new Set() };
      filteredTecnicoStats[techName].total_os++;
      filteredTecnicoStats[techName].clientes.add(item.login || item.cliente);

      const diagKey = os.diagnostico || 'Sem diagnóstico';
      if (!filteredDiagnosticoStats[diagKey]) filteredDiagnosticoStats[diagKey] = { quantidade: 0 };
      filteredDiagnosticoStats[diagKey].quantidade++;
      // Contabiliza o diagnóstico apenas para O.S. de retorno (ignora a Origem/Instalação)
      if (!os.is_origem) {
        filteredTotalRetrabalho++;
        const diagKey = os.diagnostico || 'Sem diagnóstico';
        if (!filteredDiagnosticoStats[diagKey]) filteredDiagnosticoStats[diagKey] = { quantidade: 0 };
        filteredDiagnosticoStats[diagKey].quantidade++;
      }
    });

    Object.entries(assuntoGroups).forEach(([assuntoKey, rows]) => {
      for (let i = 1; i < rows.length; i++) {
        const diff = rows[i].dias_apos_primeira - rows[i - 1].dias_apos_primeira;
        filteredAssuntoStats[assuntoKey].intervals.push(diff);
      }
    });
  });

  const calcEstatisticasPorAssunto = Object.entries(filteredAssuntoStats)
    .map(([assunto, data]) => ({ assunto, quantidade: data.quantidade, percentual: filteredTotalOs ? roundToOne((data.quantidade / filteredTotalOs) * 100) : 0, tempo_medio_resolucao_dias: data.intervals.length ? roundToOne(data.intervals.reduce((s, v) => s + v, 0) / data.intervals.length) : 0 }))
    .sort((a, b) => b.quantidade - a.quantidade || a.assunto.localeCompare(b.assunto));

  const calcEstatisticasPorTecnico = Object.entries(filteredTecnicoStats)
    .map(([tecnico, data]) => ({ tecnico, total_os_reincidentes: data.total_os, clientes_reincidentes_atendidos: data.clientes.size, taxa_envolvimento: filteredTotalOs ? roundToOne((data.total_os / filteredTotalOs) * 100) : 0 }))
    .sort((a, b) => b.total_os_reincidentes - a.total_os_reincidentes || a.tecnico.localeCompare(b.tecnico));

  const recurrenceDiagnosticoStats = {};
  let recurrenceDiagnosticoTotal = 0;
  filteredReincidentes.forEach(item => {
    (item.historico_os || []).forEach(os => {
      if (os.is_origem) return;
      if (isExcludedRecurrenceSubject(os.assunto, excludedRecurrenceSubjects)) return;
      if (isExcludedRecurrenceDiagnosis(os.diagnostico)) return;
      const diagnosticoKey = os.diagnostico || 'Sem diagnostico';
      if (!recurrenceDiagnosticoStats[diagnosticoKey]) recurrenceDiagnosticoStats[diagnosticoKey] = { quantidade: 0 };
      recurrenceDiagnosticoStats[diagnosticoKey].quantidade++;
      recurrenceDiagnosticoTotal++;
    });
  });

  const calcEstatisticasPorDiagnostico = Object.entries(recurrenceDiagnosticoStats)
    .map(([diagnostico, data]) => ({ diagnostico, quantidade: data.quantidade, percentual: recurrenceDiagnosticoTotal ? roundToOne((data.quantidade / recurrenceDiagnosticoTotal) * 100) : 0 }))
    .sort((a, b) => b.quantidade - a.quantidade || a.diagnostico.localeCompare(b.diagnostico));

  const assuntoHtml = calcEstatisticasPorAssunto.length
    ? calcEstatisticasPorAssunto.slice(0, 8).map(item => `<div class="obs-item"><div class="obs-dot y"></div><span><b>${escapeHtml(item.assunto)}</b> · ${item.quantidade} OS · ${item.percentual}% · media ${item.tempo_medio_resolucao_dias}d</span></div>`).join('')
    : '<div class="empty">Sem estatisticas por assunto.</div>';

  const diagnosticoHtml = calcEstatisticasPorDiagnostico.length
    ? calcEstatisticasPorDiagnostico.slice(0, 8).map(item => `<div class="obs-item"><div class="obs-dot r"></div><span style="display:flex;align-items:center;width:100%;overflow:hidden;"><b style="flex:1;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;padding-right:6px;" title="${escapeHtml(item.diagnostico)}">${escapeHtml(item.diagnostico)}</b><span style="white-space:nowrap;flex-shrink:0;">· ${item.quantidade} OS · ${item.percentual}%</span></span></div>`).join('')
    : '<div class="empty">Sem estatisticas por diagnostico.</div>';

  const tecnicoHtml = calcEstatisticasPorTecnico.length
    ? calcEstatisticasPorTecnico.slice(0, 8).map(item => `<div class="obs-item"><div class="obs-dot b"></div><span><b>${escapeHtml(item.tecnico)}</b> · ${item.total_os_reincidentes} OS · ${item.clientes_reincidentes_atendidos} clientes · ${item.taxa_envolvimento}%</span></div>`).join('')
    : '<div class="empty">Sem estatisticas por tecnico.</div>';

  const rowsHtml = filteredReincidentes.length
    ? filteredReincidentes.map(item => `
      <tr>
        <td class="td-nm">
          <div style="display:flex;flex-direction:column;gap:6px;max-width:200px;">
            <strong style="line-height:1.35;">${escapeHtml(item.login || item.cliente || 'Sem identificador')}</strong>
            <span style="font-size:11px;color:var(--text-tertiary);line-height:1.35;">${escapeHtml(item.cliente || 'Cliente nao informado')}</span>
          </div>
        </td>
        <td><span class="rec-chip rec-chip-muted">${escapeHtml(item.filial)}</span></td>
        <td class="td-r"><span class="rec-chip rec-chip-${item.classificacao === 'CRITICA' ? 'danger' : item.classificacao === 'ALTA' ? 'warning' : 'info'}">${escapeHtml(getClassificationLabel(item.classificacao))}</span></td>
        <td class="td-r">${item.total_os_periodo}</td>
        <td class="td-r">${item.dias_entre_primeira_e_ultima}</td>
        <td class="td-r">${item.media_dias_entre_os}</td>
        <td>${escapeHtml(item.assunto_principal)}</td>
        <td style="max-width:220px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;" title="${escapeHtml(item.diagnostico_principal || '—')}">
          <span style="font-size:11px;font-weight:700;color:var(--text-primary);">${escapeHtml(item.diagnostico_principal || '—')}</span>
        </td>
        <td style="white-space:nowrap;">
          <div style="display:flex;flex-direction:column;gap:6px;">
            <span>${escapeHtml(item.primeira_os.data)}</span>
            <span class="rec-chip ${item.origem_reincidencia === 'ATIVAÇÃO' ? 'rec-chip-info' : 'rec-chip-warning'}" style="font-size:9px;padding:2px 6px;width:fit-content;">Após ${item.origem_reincidencia}</span>
          </div>
        </td>
        <td style="text-align:center; position:sticky; right:0; background:var(--surface-card); box-shadow:-2px 0 5px rgba(0,0,0,0.05); z-index:1;">
          <button class="btn btn-outline" style="font-size:10px;padding:4px 8px;" onclick="App.showRecurrenceDetails(decodeURIComponent('${encodeURIComponent(item.login || item.cliente)}'))">Histórico</button>
        </td>
      </tr>`).join('')
    : '<tr><td colspan="10" class="empty">Nenhum cliente reincidente encontrado com os filtros atuais.</td></tr>';

  container.innerHTML = `
    <div class="analise-grid">
      <div class="analise-card">
        <div class="analise-label">Clientes Analisados</div>
        <div class="analise-value">${resumo.total_clientes_analisados || 0}</div>
        <div class="analise-sub">Logins com OS finalizadas por tecnicos cadastrados</div>
      </div>
      <div class="analise-card">
        <div class="analise-label">Reincidentes</div>
        <div class="analise-value">${filteredReincidentes.length || 0}</div>
        <div class="analise-sub">${resumo.percentual_reincidencia || 0}% da base analisada</div>
      </div>
      <div class="analise-card">
        <div class="analise-label">Media Entre OS</div>
        <div class="analise-value">${resumo.media_dias_entre_os || 0}d</div>
        <div class="analise-sub">Media das reincidencias detectadas</div>
      </div>
      <div class="analise-card">
        <div class="analise-label">Periodo</div>
        <div class="analise-value">${settings.periodoBuscaDias}d</div>
        <div class="analise-sub">Janela de busca ate ${escapeHtml(formatDateDisplay(analysis.filtro_aplicado.data_fim_referencia))}</div>
      </div>
    </div>

    <div class="dc">
      <div class="dc-hdr"><span class="dc-ttl">Parametros da Analise</span></div>
      <div style="padding:16px 20px;">
        <div class="form-row">
          <div class="fgrp">
            <label class="flbl">Mes de Referencia</label>
            <input type="month" class="fctl" value="${escapeHtml(state.activeMonthYear)}" onchange="App.changeActiveMonth(this.value)">
          </div>
          <div class="fgrp">
            <label class="flbl">Dias de Analise</label>
            <input type="number" class="fctl" min="1" value="${settings.diasAnalise}" onchange="App.updateRecurrenceSetting('diasAnalise', this.value)">
          </div>
          <div class="fgrp">
            <label class="flbl">Periodo de Busca</label>
            <input type="number" class="fctl" min="1" value="${settings.periodoBuscaDias}" onchange="App.updateRecurrenceSetting('periodoBuscaDias', this.value)">
          </div>
        </div>
        <div style="display:grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap:14px; margin-top:14px;">
          <div class="fgrp">
            <label class="flbl">Minimo para Reincidencia</label>
            <input type="number" class="fctl" min="2" value="${settings.minimoOsParaReincidencia}" onchange="App.updateRecurrenceSetting('minimoOsParaReincidencia', this.value)">
          </div>
          <div class="fgrp">
            <label class="flbl">Filial</label>
            <select id="recFilialFilter" class="fctl" onchange="App.renderRecorrenciaClientes()">
              <option value="ALL">Todas as Filiais</option>
              ${filiais.map(f => `<option value="${escapeHtml(f)}" ${selectedFilial === f ? 'selected' : ''}>${escapeHtml(f)}</option>`).join('')}
            </select>
          </div>
          <div class="fgrp">
            <label class="flbl">Origem da Falha</label>
            <select id="recOrigemFilter" class="fctl" onchange="App.renderRecorrenciaClientes()">
              <option value="ALL">Todas as Origens</option>
              <option value="ATIVAÇÃO" ${selectedOrigem === 'ATIVAÇÃO' ? 'selected' : ''}>Após Ativação</option>
              <option value="SUPORTE" ${selectedOrigem === 'SUPORTE' ? 'selected' : ''}>Após Suporte</option>
            </select>
          </div>
        <div class="fgrp">
          <label class="flbl">Classificação</label>
          <select id="recClassFilter" class="fctl" onchange="App.renderRecorrenciaClientes()">
            <option value="ALL">Todas as Classificações</option>
            <option value="CRITICA" ${selectedClass === 'CRITICA' ? 'selected' : ''}>Crítica</option>
            <option value="ALTA" ${selectedClass === 'ALTA' ? 'selected' : ''}>Alta</option>
            <option value="MEDIA" ${selectedClass === 'MEDIA' ? 'selected' : ''}>Média</option>
            <option value="BAIXA" ${selectedClass === 'BAIXA' ? 'selected' : ''}>Baixa</option>
          </select>
        </div>
        <div class="fgrp">
          <label class="flbl">Técnico Origem</label>
          <input type="text" id="recTecnicoFilter" class="fctl" list="recTecnicosList" placeholder="Todos ou digite o nome..." value="${escapeHtml(selectedTech !== 'ALL' ? selectedTech : '')}" onchange="App.renderRecorrenciaClientes()">
          <datalist id="recTecnicosList">
            ${tecnicosOrigem.map(t => `<option value="${escapeHtml(t)}"></option>`).join('')}
          </datalist>
        </div>
        <div class="fgrp" style="grid-column: 1 / -1;">
          <label class="flbl">Busca Livre (Pressione Enter para confirmar)</label>
          <input type="text" id="recSearchFilter" class="fctl" placeholder="Buscar por Login, Cliente, Assunto ou Diagnóstico..." value="${escapeHtml(searchVal)}" onkeydown="if(event.key==='Enter') App.renderRecorrenciaClientes()">
        </div>
        </div>
        <div class="btn-row" style="margin-top:14px;">
          <button class="btn btn-outline" onclick="App.exportRecorrenciaExcel()">Exportar Excel</button>
          <button class="btn btn-outline" onclick="App.exportRecorrenciaJson()">Exportar JSON</button>
          <button class="btn btn-ghost" onclick="App.resetRecorrenciaFiltros()">Limpar Filtros</button>
        </div>
      </div>
    </div>

    <div class="obs-grid">
      <div class="obs-card">
        <div class="obs-title">Alertas</div>
        <div>${alertsHtml}</div>
      </div>
      <div class="obs-card">
        <div class="obs-title">Origem por Filial</div>
        <div style="max-height:220px;overflow-y:auto;padding-right:4px;">${filialStatsHtml}</div>
      </div>
    </div>

    <div class="obs-grid" style="grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));">
      <div class="obs-card">
        <div class="obs-title">Top Assuntos</div>
        <div>${assuntoHtml}</div>
      </div>
      <div class="obs-card">
        <div class="obs-title">Top Diagnósticos</div>
        <div>${diagnosticoHtml}</div>
      </div>
      <div class="obs-card">
        <div class="obs-title">Top Origem de Retrabalho</div>
        <div>${tecnicoHtml}</div>
      </div>
    </div>

    <div class="dc">
      <div class="dc-hdr"><span class="dc-ttl">Clientes Reincidentes</span></div>
      <div class="table-responsive" style="max-height:65vh; overflow:auto;">
        <table class="dtable">
          <thead style="position:sticky; top:0; z-index:10; background:var(--surface-card); box-shadow:0 2px 4px rgba(0,0,0,0.05);">
            <tr>
              <th>Login / Cliente</th>
              <th>Filial</th>
              <th>Classificacao</th>
              <th>Total OS</th>
              <th>Dias 1a-Ultima</th>
              <th>Media Entre OS</th>
              <th>Assunto Principal</th>
              <th>Diagnóstico Principal</th>
              <th>Primeira OS</th>
              <th style="text-align:center; position:sticky; right:0; background:var(--surface-card); z-index:11; box-shadow:-2px 0 4px rgba(0,0,0,0.05);">Ações</th>
            </tr>
          </thead>
          <tbody>${rowsHtml}</tbody>
        </table>
      </div>
    </div>`;
}

function showRecurrenceDetails(loginOrCliente) {
  const analysis = state.recurrenceAnalysis;
  if (!analysis || !analysis.reincidentes) return;
  const item = analysis.reincidentes.find(r => (r.login || r.cliente) === loginOrCliente);
  if (!item) return;

  let overlay = document.getElementById('recurrenceModalOverlay');
  if (overlay) overlay.remove();
  overlay = document.createElement('div');
  overlay.id = 'recurrenceModalOverlay';
  overlay.style.cssText = 'position:fixed;inset:0;z-index:9999;background:rgba(15,23,42,.6);display:flex;align-items:center;justify-content:center;backdrop-filter:blur(8px);animation:fadeIn .15s ease;padding:16px;';

  const rows = item.historico_os.map((os, idx) => `
    <tr style="border-bottom:1px solid var(--surface-border);background:${idx === 0 ? 'var(--info-bg)' : 'transparent'}">
      <td style="padding:10px 12px;font-family:var(--mono);font-size:11px;color:var(--accent);white-space:nowrap;">${escapeHtml(os.id)}</td>
      <td style="padding:10px 12px;font-size:11px;white-space:nowrap;">${escapeHtml(os.data)}</td>
      <td style="padding:10px 12px;font-size:11px;line-height:1.4;"><b style="color:var(--text-primary);">${escapeHtml(os.assunto)}</b><br><span style="color:var(--text-tertiary);">${escapeHtml(os.diagnostico)}</span></td>
      <td style="padding:10px 12px;font-size:11px;color:var(--text-secondary);">${escapeHtml(os.tecnico)}</td>
      <td style="padding:10px 12px;font-size:11px;text-align:center;">${!os.is_origem ? `<span style="background:var(--danger-bg);color:var(--danger);padding:3px 6px;border-radius:4px;font-weight:bold;">+${os.dias_apos_primeira} dias</span>` : '<span style="color:var(--text-tertiary);font-weight:bold;">ORIGEM</span>'}</td>
    </tr>
  `).join('');

  overlay.innerHTML = `
    <div style="background:var(--surface-card);border:1px solid var(--surface-border);border-radius:16px;max-width:850px;width:100%;max-height:85vh;display:flex;flex-direction:column;box-shadow:0 16px 48px rgba(0,0,0,.5);animation:scaleIn .15s ease;overflow:hidden;">
      <div style="padding:20px 24px;border-bottom:1px solid var(--surface-border);display:flex;justify-content:space-between;align-items:center;background:var(--surface-card2);">
        <div>
          <div style="font-size:16px;font-weight:800;color:var(--text-primary);">Histórico de O.S. — Reincidência</div>
          <div style="font-size:12px;color:var(--text-tertiary);margin-top:2px;">Cliente: <b style="color:var(--text-secondary);">${escapeHtml(item.login || item.cliente)}</b></div>
        </div>
        <button id="closeRecModal" style="background:var(--hover-white-04);border:1px solid var(--surface-border);border-radius:8px;color:var(--text-tertiary);font-size:14px;cursor:pointer;width:32px;height:32px;display:flex;align-items:center;justify-content:center;transition:all var(--dur-fast);">✖</button>
      </div>
      <div style="overflow-y:auto;padding:0;flex:1;">
        <table style="width:100%;border-collapse:collapse;text-align:left;">
          <thead style="background:var(--surface-card2);position:sticky;top:0;box-shadow:0 1px 2px rgba(0,0,0,0.05);z-index:2;">
            <tr>
              <th style="padding:12px;font-size:10px;color:var(--text-tertiary);text-transform:uppercase;font-family:var(--mono);">Protocolo</th>
              <th style="padding:12px;font-size:10px;color:var(--text-tertiary);text-transform:uppercase;font-family:var(--mono);">Data</th>
              <th style="padding:12px;font-size:10px;color:var(--text-tertiary);text-transform:uppercase;font-family:var(--mono);">Assunto & Diagnóstico</th>
              <th style="padding:12px;font-size:10px;color:var(--text-tertiary);text-transform:uppercase;font-family:var(--mono);">Técnico Envolvido</th>
              <th style="padding:12px;font-size:10px;color:var(--text-tertiary);text-transform:uppercase;font-family:var(--mono);text-align:center;">Intervalo</th>
            </tr>
          </thead>
          <tbody>${rows}</tbody>
        </table>
      </div>
    </div>`;

  document.body.appendChild(overlay);
  const closeBtn = document.getElementById('closeRecModal');
  closeBtn.onclick = () => overlay.remove();
  closeBtn.onmouseover = () => { closeBtn.style.background = 'var(--hover-white-08)'; closeBtn.style.color = 'var(--text-primary)'; };
  closeBtn.onmouseout = () => { closeBtn.style.background = 'var(--hover-white-04)'; closeBtn.style.color = 'var(--text-tertiary)'; };
  overlay.onclick = e => { if (e.target === overlay) overlay.remove(); };
}

function exportRecorrenciaExcel() {
  const analysis = getRecorrenciaViewData();
  if (!analysis?.success) return alert(analysis?.message || 'Sem dados para exportar.');
  if (!analysis.reincidentes.length) return alert('Nenhum reincidente encontrado para exportar.');

  const filialSelectEl = document.getElementById('recFilialFilter');
  const selectedFilial = filialSelectEl ? filialSelectEl.value : 'ALL';
  const origemSelectEl = document.getElementById('recOrigemFilter');
  const selectedOrigem = origemSelectEl ? origemSelectEl.value : 'ALL';
  const classSelectEl = document.getElementById('recClassFilter');
  const selectedClass = classSelectEl ? classSelectEl.value : 'ALL';
  const techInputEl = document.getElementById('recTecnicoFilter');
  let selectedTech = techInputEl ? techInputEl.value.trim() : 'ALL';
  if (selectedTech === '') selectedTech = 'ALL';
  const searchInputEl = document.getElementById('recSearchFilter');
  const searchVal = searchInputEl ? searchInputEl.value.trim().toLowerCase() : '';

  const filteredReincidentes = analysis.reincidentes.filter(i => {
    if (selectedFilial !== 'ALL' && i.filial !== selectedFilial) return false;
    if (selectedOrigem !== 'ALL' && i.origem_reincidencia !== selectedOrigem) return false;
    if (selectedClass !== 'ALL' && i.classificacao !== selectedClass) return false;
    if (selectedTech !== 'ALL' && i.primeira_os.tecnico !== selectedTech) return false;
    if (searchVal) {
      const term = searchVal;
      const inLogin = (i.login || '').toLowerCase().includes(term);
      const inCliente = (i.cliente || '').toLowerCase().includes(term);
      const inAssunto = (i.assunto_principal || '').toLowerCase().includes(term);
      const inDiag = (i.diagnostico_principal || '').toLowerCase().includes(term);
      if (!inLogin && !inCliente && !inAssunto && !inDiag) return false;
    }
    return true;
  });
  const filteredAlertas = analysis.alertas.filter(a => filteredReincidentes.some(r => (r.login || r.cliente) === a.cliente));

  if (!filteredReincidentes.length) return alert('Nenhum reincidente com os filtros atuais.');

  const wb = XLSX.utils.book_new();
  const reincidentesData = [
    ['Login', 'Cliente', 'Filial', 'Origem da Falha', 'Classificacao', 'Total OS', 'Dias 1a-Ultima', 'Media Entre OS', 'Assunto Principal', 'Diagnostico Principal', 'Qtd Diagnosticos Unicos', 'Tecnicos Envolvidos'],
    ...filteredReincidentes.map(item => [
      item.login,
      item.cliente,
      item.filial,
      `Apos ${item.origem_reincidencia}`,
      getClassificationLabel(item.classificacao),
      item.total_os_periodo,
      item.dias_entre_primeira_e_ultima,
      item.media_dias_entre_os,
      item.assunto_principal,
      item.diagnostico_principal,
      item.diagnosticos_resumo?.length || 0,
      item.tecnicos_envolvidos.join(', ')
    ])
  ];
  const wsReincidentes = XLSX.utils.aoa_to_sheet(reincidentesData);
  wsReincidentes['!cols'] = [{ wch: 24 }, { wch: 32 }, { wch: 24 }, { wch: 16 }, { wch: 14 }, { wch: 10 }, { wch: 14 }, { wch: 14 }, { wch: 28 }, { wch: 36 }, { wch: 18 }, { wch: 42 }];
  XLSX.utils.book_append_sheet(wb, wsReincidentes, 'Reincidentes');

  const historicoData = [
    ['Login', 'Cliente', 'Contrato', 'Filial', 'Origem da Falha', 'Tipo Visita', 'Protocolo', 'Data', 'Assunto', 'Diagnostico', 'Tecnico', 'Equipe', 'Intervalo (Dias)'],
    ...filteredReincidentes.flatMap(item =>
      item.historico_os.map(os => [
        item.login,
        item.cliente,
        os.contrato,
        item.filial,
        `Apos ${item.origem_reincidencia}`,
        os.is_origem ? 'ORIGEM' : 'REINCIDENCIA',
        os.id,
        os.data,
        os.assunto,
        os.diagnostico,
        os.tecnico,
        os.tipo_equipe,
        os.is_origem ? '—' : `+${os.dias_apos_primeira} dias`
      ])
    )
  ];
  const wsHistorico = XLSX.utils.aoa_to_sheet(historicoData);
  wsHistorico['!cols'] = [{ wch: 20 }, { wch: 32 }, { wch: 18 }, { wch: 24 }, { wch: 16 }, { wch: 16 }, { wch: 14 }, { wch: 18 }, { wch: 32 }, { wch: 42 }, { wch: 24 }, { wch: 18 }, { wch: 16 }];
  XLSX.utils.book_append_sheet(wb, wsHistorico, 'Historico');

  const diagnosticosData = [
    ['Login', 'Cliente', 'Filial', 'Diagnostico', 'Quantidade', 'Classificacao', 'Assunto Principal'],
    ...filteredReincidentes.flatMap(item =>
      (item.diagnosticos_resumo || []).map(diag => [
        item.login,
        item.cliente,
        item.filial,
        diag.diagnostico,
        diag.quantidade,
        getClassificationLabel(item.classificacao),
        item.assunto_principal
      ])
    )
  ];
  const wsDiagnosticos = XLSX.utils.aoa_to_sheet(diagnosticosData);
  wsDiagnosticos['!cols'] = [{ wch: 24 }, { wch: 32 }, { wch: 24 }, { wch: 42 }, { wch: 12 }, { wch: 14 }, { wch: 28 }];
  XLSX.utils.book_append_sheet(wb, wsDiagnosticos, 'Diagnosticos');

  const alertasData = [
    ['Tipo', 'Cliente', 'Mensagem', 'Prioridade'],
    ...filteredAlertas.map(item => [item.tipo, item.cliente, item.mensagem, item.prioridade])
  ];
  const wsAlertas = XLSX.utils.aoa_to_sheet(alertasData);
  wsAlertas['!cols'] = [{ wch: 12 }, { wch: 24 }, { wch: 60 }, { wch: 10 }];
  XLSX.utils.book_append_sheet(wb, wsAlertas, 'Alertas');

  XLSX.writeFile(wb, `SGO_Recorrencia_${state.activeMonthYear}.xlsx`);
}

function exportRecorrenciaJson() {
  const analysis = getRecorrenciaViewData();
  if (!analysis?.success) return alert(analysis?.message || 'Sem dados para exportar.');
  const a = document.createElement('a');
  a.href = 'data:text/json;charset=utf-8,' + encodeURIComponent(JSON.stringify(analysis, null, 2));
  a.download = `SGO_Recorrencia_${state.activeMonthYear}.json`;
  a.click();
}

function updateRecurrenceSetting(key, value) {
  const numericValue = Math.max(key === 'minimoOsParaReincidencia' ? 2 : 1, parseInt(value, 10) || 0);
  resetRecurrenceCache();
  state.recurrenceSettings[key] = numericValue;
  renderRecorrenciaClientes();
}

function resetRecorrenciaFiltros() {
  ['recFilialFilter', 'recOrigemFilter', 'recClassFilter'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = 'ALL';
  });
  const techEl = document.getElementById('recTecnicoFilter');
  if (techEl) techEl.value = '';
  const searchEl = document.getElementById('recSearchFilter');
  if (searchEl) searchEl.value = '';
  renderRecorrenciaClientes();
}

function evaluateTechsByAI() {
  if(!Object.keys(state.globalTechStats).length)return;
  state.pendingTechs=[]; state.reclassifySuggestions=[];
  const resolver = buildTeamResolver();
  const cache={};
  const matchTeam=nome=>{const team=resolveTeamData(nome,resolver,cache);return team?team.techKey:null;};
  Object.keys(state.globalTechStats).forEach(nome=>{
    const s=state.globalTechStats[nome];
    const pR=s.rural/s.total, da=[...s.days].sort((a,b)=>a-b), dt=da.length;
    let d2=0;for(let k=1;k<da.length;k++)if(da[k]-da[k-1]===2)d2++;
    const isP=dt>=4&&d2/(dt-1||1)>=0.5;
    let sug='INSTALAÇÃO CIDADE';
    if(isP)sug='TECNICO 12/36H';else if(pR>=0.6)sug='RURAL';else if(s.total/dt>7.5)sug='SUPORTE MOTO';
    const kc=matchTeam(nome);
    if(!kc&&s.total>1)state.pendingTechs.push({nome,total:s.total,suggestedType:sug});
    else if(kc&&s.total>=5){const ct=state.teamData[kc].tipo||'INSTALAÇÃO CIDADE';if(ct!==sug)state.reclassifySuggestions.push({key:kc,nome:state.teamData[kc].originalName,currentType:ct,suggestedType:sug});}
  });
  renderPendentes();
}

function renderPendentes() {
  const badge=(id,n)=>{const el=document.getElementById(id);if(!el)return;if(n>0){el.textContent=n;el.classList.remove('hidden');}else el.classList.add('hidden');};
  badge('badgePendentes',state.pendingTechs.length);badge('mobBadgePendentes',state.pendingTechs.length);
  badge('badgeReclassificacao',state.reclassifySuggestions.length);badge('mobBadgeReclass',state.reclassifySuggestions.length);
  const tb=document.getElementById('pendentesTableBody');
  if(tb)tb.innerHTML=state.pendingTechs.length?state.pendingTechs.sort((a,b)=>b.total-a.total).map(n=>{const sn=n.nome.replace(/"/g,'&quot;').replace(/'/g,"\\'");return`<tr><td class="td-nm">${n.nome}</td><td class="td-ct">${n.total}</td><td class="td-tp">${TEAM_TYPES[n.suggestedType]||'Outros'}</td><td class="td-r"><button class="btn btn-accent" style="font-size:10px;padding:6px 11px;" data-nome="${sn}" data-type="${n.suggestedType}" onclick="openAddTechModal(this.getAttribute('data-nome'),'',this.getAttribute('data-type'))">+ Adicionar</button></td></tr>`;}).join(''):`<tr><td colspan="4" class="empty">Nenhum pendente.</td></tr>`;
  const tr=document.getElementById('reclassifyTableBody');
  if(tr)tr.innerHTML=state.reclassifySuggestions.length?state.reclassifySuggestions.map(r=>`<tr><td class="td-nm">${r.nome}</td><td class="td-st">${TEAM_TYPES[r.currentType]}</td><td class="td-pu">${TEAM_TYPES[r.suggestedType]}</td><td class="td-r"><button class="btn btn-purple" style="font-size:10px;padding:6px 11px;" onclick="acceptReclassification('${r.key}','${r.suggestedType}')">✓ Aceitar</button></td></tr>`).join(''):`<tr><td colspan="4" class="empty">Classificações corretas.</td></tr>`;
}

/* ═══════════════════════════════════════════════════════════
   UPLOAD — worker de parsing e pipeline de dados
   ═══════════════════════════════════════════════════════════ */
const WORKER_CODE = `
importScripts('https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.min.js');
function cleanStr(s){if(!s)return"";return String(s).split('-')[0].split('>')[0].split('(')[0].split('/')[0].normalize("NFD").replace(/[\\u0300-\\u036f]/g,"").replace(/\\s+/g,' ').trim().toUpperCase();}
function extractDate(v){
  if(!v)return null;
  let s=String(v).trim();
  if(!s)return null;
  if(typeof v==='number'||(!isNaN(Number(s)) && !s.includes(':'))){let d=new Date(Math.round((Number(s)-25569)*86400*1000));d.setMinutes(d.getMinutes()+d.getTimezoneOffset());return d;}
  let b=s.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?/);
  if(b){let y=b[3].length===2?parseInt('20'+b[3]):parseInt(b[3]);let H=b[4]?parseInt(b[4],10):0;let M=b[5]?parseInt(b[5],10):0;let S=b[6]?parseInt(b[6],10):0;return new Date(y,parseInt(b[2],10)-1,parseInt(b[1],10),H,M,S);}
  let i=s.match(/(\\d{4})-(\\d{1,2})-(\\d{1,2})(?:T|\\s+)(\\d{1,2}):(\\d{1,2})(?::(\\d{1,2}))?/);
  if(i){let H=i[4]?parseInt(i[4],10):0;let M=i[5]?parseInt(i[5],10):0;let S=i[6]?parseInt(i[6],10):0;return new Date(parseInt(i[1]),parseInt(i[2])-1,parseInt(i[3]),H,M,S);}
  return null;
}
self.onmessage=function(e){try{const{fileData}=e.data;const wb=XLSX.read(fileData,{type:'array',raw:true});const rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1,defval:""});if(rows.length<=1)throw new Error("Arquivo sem dados");const hdrs=rows[0].map(h=>String(h).normalize("NFD").replace(/[\\u0300-\\u036f]/g,"").trim().toLowerCase());
let iD=hdrs.findIndex(h=>h.includes('fechamento')||h.includes('conclusao'));
if(iD===-1) iD=hdrs.findIndex(h=>h==='data'||h.includes('data/hora')||h.includes('abertura'));
const iR=hdrs.findIndex(h=>h.includes('colaborador')||h.includes('responsavel')||h.includes('tecnico')||h.includes('executor'));
const iA=hdrs.findIndex(h=>h.includes('assunto'));
const iDg=hdrs.findIndex(h=>h.includes('diagnostico')||h.includes('diagnóstico'));
const iCt=hdrs.findIndex(h=>h.includes('contrato'));
const iCl=hdrs.findIndex(h=>h.includes('cliente')||h.includes('assinante')||h.includes('nome do cliente')||h.includes('nome cliente'));
const iLg=hdrs.findIndex(h=>h.includes('login')||h.includes('usuario')||h.includes('usuário'));
const iSt=hdrs.findIndex(h=>h.includes('status')||h.includes('situacao')||h.includes('situação'));
const iOs=hdrs.findIndex(h=>h==='id'||h==='os'||h.includes('protocolo')||h.includes('ticket')||h==='numero'||h==='nº'); 
const iFi=hdrs.findIndex(h=>h.includes('filial')||h==='empresa'||h==='unidade');
const iIn=hdrs.findIndex(h=>h==='início'||h==='inicio');
const iFn=hdrs.findIndex(h=>h==='final');
if(iD===-1||iR===-1)throw new Error("Colunas 'Responsável' e/ou 'Data' não encontradas.");let pm={},ts={},vr=[];

// NOVO: Adicionei CHÁCARA e KM na busca para varrer endereços/bairros
const rR=/(RURAL|FAZENDA|S[IÍ]TIO|LINHA |GLEBA|PROJETO|CH[AÁ]CARA|KM \\d)/i; 

const mapFiliais={"6":"UNI - JI PARANA","7":"UNI - MACHADINHO DOESTE","8":"UNI - ROLIM DE MOURA","9":"UNI - JARU","10":"UNI - OURO PRETO DOESTE","11":"UNI - NOVA BRASILANDIA DOESTE","12":"UNI - PRESIDENTE MEDICI","13":"UNI - SAO FELIPE DOESTE","14":"UNI - ALVORADA DOESTE","15":"UNI - ALTA FLORESTA DOESTE","16":"UNI - SAO MIGUEL DO GUAPORE","17":"UNI - SERINGUEIRAS","18":"UNI - SAO FRANCISCO DO GUAPORE"};
function hasRuralHint(row){for(let j=0;j<row.length;j++){const cell=row[j];if(cell&&rR.test(String(cell)))return true;}return false;}
for(let i=1;i<rows.length;i++){const row=rows[i],dt=row[iD],rp=row[iR];if(!dt||!rp||String(rp).toLowerCase().includes("filtros"))continue;const d=extractDate(dt);if(!d)continue;const nm=cleanStr(rp);const mk=d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,'0');pm[mk]=(pm[mk]||0)+1;
const localDateStr=d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,'0')+"-"+String(d.getDate()).padStart(2,'0');

// Lê a linha inteira da planilha (incluindo endereço/bairro) e testa as palavras-chave rurais
let isR=hasRuralHint(row); 

if(!ts[nm])ts[nm]={total:0,rural:0,days:new Set()};ts[nm].total++;if(isR)ts[nm].rural++;ts[nm].days.add(d.getDate());const assunto=iA>=0?String(row[iA]||'').trim():'';const diagnostico=iDg>=0?String(row[iDg]||'').trim():'';
const contrato=iCt>=0?String(row[iCt]||'').trim():'';
const cliente=iCl>=0?String(row[iCl]||'').trim():'';
const login=iLg>=0?String(row[iLg]||'').trim():'';
const status=iSt>=0?String(row[iSt]||'').trim():'';
const osId=iOs>=0?String(rows[i][iOs]||'').trim():'—'; 
const rawFilial=iFi>=0?String(row[iFi]||'').trim():'';
const filial=mapFiliais[rawFilial]||rawFilial||'NÃO INFORMADA';
const dtInicio = iIn>=0 ? extractDate(row[iIn]) : null;
const dtFinal = iFn>=0 ? extractDate(row[iFn]) : null;

// NOVO: Passa o isRural para frente junto com a O.S. (Corrigido para usar fuso local na dateStr)
vr.push({nome:nm,nomeOriginal:String(rp||'').trim(),day:d.getDate(),monthStr:mk,dateStr:localDateStr,dateTimeStr:d.toISOString(),assunto,diagnostico,contrato,cliente,login,status,osId,isRural:isR,filial,dtInicio:dtInicio?dtInicio.toISOString():null,dtFinal:dtFinal?dtFinal.toISOString():null}); 
}if(!vr.length)throw new Error("Nenhuma OS válida.");const months=Object.keys(pm).sort();let am=months[months.length-1];let ss={};for(let k in ts)ss[k]={total:ts[k].total,rural:ts[k].rural,days:Array.from(ts[k].days)};self.postMessage({success:true,activeMonth:am,allOS:vr,techStats:ss,uploadMeta:{hasContrato:iCt>=0,hasCliente:iCl>=0,hasLogin:iLg>=0,hasStatus:iSt>=0,availableMonths:months}});}catch(err){self.postMessage({success:false,error:err.message});}};
`;

function initWorker() {
  const blob = new Blob([WORKER_CODE], { type: 'application/javascript' });
  state.workerBlobUrl = URL.createObjectURL(blob);
}

function buildTeamResolver() {
  const exact = Object.create(null);
  const partial = [];
  Object.entries(state.teamData).forEach(([key, info]) => {
    const resolved = {
      techKey: key,
      cidade: info.base,
      tipo: info.tipo || 'INSTALAÇÃO CIDADE'
    };
    resolved.tipo = info.tipo || Object.keys(TEAM_TYPES)[0];
    exact[key] = resolved;
    partial.push([key, resolved]);
  });
  partial.sort((a, b) => b[0].length - a[0].length);
  return { exact, partial };
}

function resolveTeamData(nome, resolver, cache) {
  if (Object.prototype.hasOwnProperty.call(cache, nome)) return cache[nome];
  const exactMatch = resolver.exact[nome];
  if (exactMatch) {
    cache[nome] = exactMatch;
    return exactMatch;
  }
  for (let i = 0; i < resolver.partial.length; i++) {
    const [key, resolved] = resolver.partial[i];
    if (nome.includes(key) || key.includes(nome)) {
      cache[nome] = resolved;
      return resolved;
    }
  }
  cache[nome] = null;
  return null;
}

function applyUploadedData(payload) {
  if (!payload?.allOS?.length) return false;
  resetRecurrenceCache();
  state.filterViewCache = {};
  
  // NOVO: Garantir um ID único por linha para impedir falsas colisões em O.S. genéricas
  payload.allOS.forEach((os, idx) => {
    if (!os._uid) os._uid = 'os_' + idx + '_' + Math.random().toString(36).substr(2, 5);
  });

  state.activeMonthYear = payload.activeMonth;
  state.rawExcelCache = payload.allOS;
  state.globalRawDataByMonth = {};
  state.globalTechStats = payload.techStats || {};
  state.uploadMeta = payload.uploadMeta || {
    hasContrato: state.rawExcelCache.some(os => !!os.contrato),
    hasCliente: state.rawExcelCache.some(os => !!os.cliente),
    hasLogin: state.rawExcelCache.some(os => !!os.login),
    hasStatus: state.rawExcelCache.some(os => !!os.status),
    availableMonths: [...new Set(state.rawExcelCache.map(os => os.monthStr).filter(Boolean))].sort()
  };
  const availableMonths = state.uploadMeta.availableMonths?.length
    ? state.uploadMeta.availableMonths.slice().sort()
    : [...new Set(state.rawExcelCache.map(os => os.monthStr).filter(Boolean))].sort();
  if (!availableMonths.includes(state.activeMonthYear) && availableMonths.length) {
    state.activeMonthYear = availableMonths[availableMonths.length - 1];
  }
  ['filterMonth','filterMonthMob'].forEach(id => {
    const el = document.getElementById(id);
    if (el) {
      el.value = state.activeMonthYear;
      if (availableMonths.length) {
        el.min = availableMonths[0];
        el.max = availableMonths[availableMonths.length - 1];
      }
    }
  });
  const cityTabs = document.getElementById('cityTabsWrapper');
  if (cityTabs) cityTabs.style.display = 'block';
  const searchBar = document.getElementById('searchBar');
  if (searchBar) searchBar.style.display = 'flex';
  const mobActions = document.getElementById('mobActions');
  if (mobActions) mobActions.style.display = 'flex';
  setStatus('Dados · ' + state.activeMonthYear);
  syncEcosystem();
  return true;
}

function changeActiveMonth(monthValue) {
  if (!monthValue || !state.rawExcelCache.length) return;
  const availableMonths = state.uploadMeta?.availableMonths?.length
    ? state.uploadMeta.availableMonths
    : [...new Set(state.rawExcelCache.map(os => os.monthStr).filter(Boolean))].sort();
  if (!availableMonths.includes(monthValue)) {
    showToast('Esse mês não está disponível na planilha carregada.', 'warning');
    ['filterMonth','filterMonthMob'].forEach(id => {
      const el = document.getElementById(id);
      if (el) el.value = state.activeMonthYear;
    });
    return;
  }
  
  showLoading(true);
  setTimeout(() => {
    resetRecurrenceCache();
    state.activeMonthYear = monthValue;
    ['filterMonth','filterMonthMob'].forEach(id => {
      const el = document.getElementById(id);
      if (el) el.value = monthValue;
    });
    setStatus('Dados · ' + state.activeMonthYear);
    applyFilters();
    setTimeout(() => showLoading(false), 100);
  }, 50);
}

async function restoreCachedUpload() {
  try {
    const cached = await loadCachedUpload();
    if (!cached?.allOS?.length) return;
    applyUploadedData(cached);
  } catch (err) {
    console.warn('[SGO] Falha ao restaurar upload em cache:', err);
  }
}


function parseFileWithWorker(event) {
  const file = event.target.files[0]; if (!file) return;
  showLoading(true);
  const reader = new FileReader();
  reader.onload = ev => {
    const ab = ev.target.result;
    const w  = new Worker(state.workerBlobUrl);
    w.onmessage = async msg => {
      showLoading(false);
      ['fileInput','fileInputMob'].forEach(id=>{ const el=document.getElementById(id); if(el)el.value=''; });
      if (msg.data.success) {
        applyUploadedData(msg.data);
        try {
          await saveCachedUpload({
            fileName: file.name,
            savedAt: Date.now(),
            activeMonth: msg.data.activeMonth,
            allOS: msg.data.allOS,
            techStats: msg.data.techStats,
            uploadMeta: msg.data.uploadMeta
          });
        } catch (err) {
          console.warn('[SGO] Falha ao salvar upload em cache:', err);
        }
      } else {
        showToast('Erro ao processar planilha: ' + msg.data.error, 'error');
      }
      w.terminate();
    };
    w.postMessage({ fileData: ab }, [ab]);
  };
  reader.readAsArrayBuffer(file);
}

function rebuildGlobalRawData() {
  if (!state.rawExcelCache || !state.rawExcelCache.length) return;
  resetRecurrenceCache();
  state.filterViewCache = {};
  const resolver = buildTeamResolver();
  const resolvedNames = Object.create(null);
  const enrichedRows = [];
  const byMonth = Object.create(null);
  state.rawExcelCache.forEach(os => {
    const team = resolveTeamData(os.nome || '', resolver, resolvedNames);
    if (!team) return;
    const row = {
      _uid: os._uid,
      techKey: team.techKey,
      cidade: team.cidade,
      tipo: team.tipo,
      day: os.day,
      monthStr: os.monthStr,
      dateStr: os.dateStr || '',
      dateTimeStr: os.dateTimeStr || os.dateStr || '',
      contrato: os.contrato || '',
      cliente: os.cliente || '',
      login: os.login || '',
      status: os.status || '',
      nome: os.nome || '',
      nomeOriginal: os.nomeOriginal || os.nome || '',
      assunto: os.assunto,
      diagnostico: os.diagnostico,
      osId: os.osId,
      isRural: os.isRural,
      filial: os.filial || 'NÃO INFORMADA',
      dtInicio: os.dtInicio,
      dtFinal: os.dtFinal
    };
    enrichedRows.push(row);
    if (!byMonth[row.monthStr]) byMonth[row.monthStr] = [];
    byMonth[row.monthStr].push(row);
  });
  state.globalRawData = enrichedRows;
  state.globalRawDataByMonth = byMonth;
  applyFilters();
  return;
  state.rawExcelCache.forEach(os => {
    let tk = cache[os.nome];
    if (tk === undefined) {
      tk = null;
      for (const k in state.teamData) { if (os.nome===k||os.nome.includes(k)||k.includes(os.nome)){tk=k;break;} }
      cache[os.nome] = tk;
    }
    if (tk) {
      const td = state.teamData[tk];
      state.globalRawData.push({ 
        _uid: os._uid,
        techKey: tk, 
        cidade: td.base, 
        tipo: td.tipo || 'INSTALAÇÃO CIDADE', 
        day: os.day, 
        monthStr: os.monthStr,
        dateStr: os.dateStr || '',
        dateTimeStr: os.dateTimeStr || os.dateStr || '',
        contrato: os.contrato || '',
        cliente: os.cliente || '',
        login: os.login || '',
        status: os.status || '',
        nome: os.nome || '',
        nomeOriginal: os.nomeOriginal || os.nome || '',
        assunto: os.assunto,           
        diagnostico: os.diagnostico,   
        osId: os.osId,                 
        isRural: os.isRural,
        filial: os.filial || 'NÃO INFORMADA',
        dtInicio: os.dtInicio,
        dtFinal: os.dtFinal
      });
    }
  });
  applyFilters();
}

function applyFilters() {
  if (!state.globalRawData.length) {
    const w=document.getElementById('matrixWrapper');
    if(w)w.innerHTML=`<div style="padding:40px;text-align:center;font-family:var(--mono);font-size:11px;color:var(--text-3);background:var(--card-bg);border:1px solid var(--card-border);border-radius:var(--r-xl);">Faça o upload da planilha para gerar a matriz.</div>`;
    if (document.getElementById('view-recorrencia')?.classList.contains('active')) {
      renderRecorrenciaClientes();
    }
    return;
  }
  const ft=document.getElementById('filterType')?.value||'ALL';
  const fm=document.getElementById('filterTypeMob'); if(fm)fm.value=ft;
  const fc=state.selectedCityTab;
  const monthRows = state.globalRawDataByMonth?.[state.activeMonthYear] || [];
  const filtered = (fc === 'ALL' && ft === 'ALL')
    ? monthRows
    : monthRows.filter(i => (fc === 'ALL' || i.cidade === fc) && (ft === 'ALL' || i.tipo === ft));
  buildCurrentFilterMeta(filtered);
  updateDashboardStats(filtered);
  generateMatrix(filtered);
  document.getElementById('legendBar').style.display='flex';
  document.getElementById('matrixSection').style.display='block';
  const capSection = document.getElementById('capSection');
  if(capSection && document.getElementById('view-capacidades')?.classList.contains('active')) {
    renderCapacidades();
    syncCapacidadesContext();
  }
  setTimeout(() => {
    buildOperationalAnalysis(filtered);
    if (document.getElementById('view-recorrencia')?.classList.contains('active')) {
      getRecorrenciaViewData();
      renderRecorrenciaClientes();
    }
  }, 50);
}

function syncEcosystem() {
  rebuildGlobalRawData();
  populateFilters();
  renderTeamTable();
  evaluateTechsByAI();
}

let _searchTimeout = null;
function filterMatrixBySearch() {
  clearTimeout(_searchTimeout);
  _searchTimeout = setTimeout(() => {
    const q=document.getElementById('techSearchInput')?.value.trim().toLowerCase()||'';
    const rows=document.querySelectorAll('#matrixWrapper tbody tr');
    let count=0;
    rows.forEach(r=>{ const name=r.querySelector('.cn-name'); if(!name){r.style.display='';return;} const match=!q||name.textContent.toLowerCase().includes(q); r.style.display=match?'':'none'; if(match)count++; });
    const sc=document.getElementById('searchCount'); if(sc)sc.textContent=q?`${count} resultado${count!==1?'s':''}`:'';
  }, 120);
}

function clearSearch() { const i=document.getElementById('techSearchInput'); if(i)i.value=''; filterMatrixBySearch(); }

/* ═══════════════════════════════════════════════════════════
   SETTINGS UI
   ═══════════════════════════════════════════════════════════ */
function loadSettingsToUI() {
  const m = state.appSettings.metasDiarias;
  document.getElementById('cfgMetaDiaComercial').value  = m['INSTALAÇÃO CIDADE'] || 5;
  document.getElementById('cfgMetaDiaPlantao').value    = m['TECNICO 12/36H']    || 7;
  document.getElementById('cfgMetaDiaSuporte').value    = m['SUPORTE MOTO']      || 9;
  document.getElementById('cfgMetaDiaSupCarro').value   = m['SUPORTE CARRO']     || 9;
  document.getElementById('cfgMetaDiaRural').value      = m['RURAL']             || 5;
  document.getElementById('cfgMetaDiaFazTudo').value    = m['FAZ TUDO']          || 6;
  const sabEl = document.getElementById('cfgMetaSabadoPct');
  if(sabEl) sabEl.value = state.appSettings.metaSabadoPct || 50;
}

function saveGlobalSettings() {
  const m = state.appSettings.metasDiarias;
  m['INSTALAÇÃO CIDADE'] = parseInt(document.getElementById('cfgMetaDiaComercial').value)||5;
  m['TECNICO 12/36H']    = parseInt(document.getElementById('cfgMetaDiaPlantao').value)  ||7;
  m['SUPORTE MOTO']      = parseInt(document.getElementById('cfgMetaDiaSuporte').value)  ||9;
  m['SUPORTE CARRO']     = parseInt(document.getElementById('cfgMetaDiaSupCarro').value) ||9;
  m['RURAL']             = parseInt(document.getElementById('cfgMetaDiaRural').value)    ||5;
  m['FAZ TUDO']          = parseInt(document.getElementById('cfgMetaDiaFazTudo').value)  ||6;
  const sabEl = document.getElementById('cfgMetaSabadoPct');
  if(sabEl) state.appSettings.metaSabadoPct = parseInt(sabEl.value)||50;
  saveSettings();
  state.filterViewCache = {};
  if (state.globalRawData.length) applyFilters();
  showToast('Metas salvas com sucesso.','success');
}

/* ═══════════════════════════════════════════════════════════
   SUPORTE MOTO — cruzamento de oportunidades
   ═══════════════════════════════════════════════════════════ */
function calcMotoOportunidades() {
  state.motoOportunidades = [];
  if (!state.rawExcelCache.length) return;

  const motoKeys = new Set(Object.keys(state.teamData).filter(k => state.teamData[k].tipo === 'SUPORTE MOTO' || state.teamData[k].tipo === 'SUPORTE CARRO'));
  const byRegional = {};

  state.rawExcelCache
    .filter(os => os.monthStr === state.activeMonthYear)
    .forEach(os => {
      const assuntoSet = SUPORTE_MOTO_MAP[os.assunto];
      if (!assuntoSet || !assuntoSet.has(os.diagnostico)) return; 

      let regional = null, tipoReal = null, feitaPorMoto = false;
      
      for (const k in state.teamData) {
        if (os.nome === k || os.nome.includes(k) || k.includes(os.nome)) {
          regional = state.teamData[k].base;
          tipoReal = state.teamData[k].tipo;
          feitaPorMoto = motoKeys.has(k);
          break;
        }
      }

      // REGRA 1: Ignora quem não está na base do SGO
      if (!regional || !tipoReal) return;

      // REGRA 2: Ignora a equipe "FAZ TUDO"
      if (tipoReal === 'FAZ TUDO') return;

      // REGRA 3: Ignora O.S. Rural
      // Bloqueia se o tipo do técnico for Rural, se o Assunto contiver Rural, 
      // OU se a leitura do endereço detectou palavras como "Linha", "Sítio", "Gleba", etc.
      if (tipoReal === 'RURAL' || os.assunto.toUpperCase().includes('RURAL') || os.isRural) return;

      if (!byRegional[regional]) byRegional[regional] = {
        total: 0, jaMoto: 0, oportunidade: 0,
        porAssunto: {}, porTipo: {}, oportunidadesDetalhadas: []
      };

      const r = byRegional[regional];
      r.total++;
      
      if (feitaPorMoto) { 
        r.jaMoto++; 
      } else { 
        r.oportunidade++; 
        r.oportunidadesDetalhadas.push({
          id: os.osId || '—',
          tecnicoAtual: os.nome,
          tipoTecnico: TEAM_TYPES[tipoReal] || tipoReal,
          assunto: os.assunto,
          diagnostico: os.diagnostico,
          data: os.day
        });
        r.porTipo[tipoReal] = (r.porTipo[tipoReal] || 0) + 1;
      }
      r.porAssunto[os.assunto] = (r.porAssunto[os.assunto] || 0) + 1;
    });

  state.motoOportunidades = Object.entries(byRegional)
    .map(([regional, d]) => ({ regional, ...d }))
    .sort((a, b) => b.oportunidade - a.oportunidade);
}

// Estado dos filtros de moto
const motoFiltros = { regional: 'ALL', assunto: 'ALL', tipo: 'ALL', busca: '' };

function showMotoDetails(regional) {
  const data = state.motoOportunidades.find(r => r.regional === regional);
  if(!data) return;

  const getIcon = (txt) => {
    const t = txt.toLowerCase();
    if(t.includes('troca') || t.includes('remoção')) return '🔄';
    if(t.includes('conexão') || t.includes('sinal') || t.includes('fibra')) return '📡';
    if(t.includes('apps') || t.includes('streaming')) return '📺';
    if(t.includes('login') || t.includes('senha')) return '🔑';
    return '🔧';
  };

  // Monta opções únicas de assunto e tipo para filtros do modal
  const assuntos = [...new Set(data.oportunidadesDetalhadas.map(o => o.assunto))].sort();
  const tipos    = [...new Set(data.oportunidadesDetalhadas.map(o => o.tipoTecnico))].sort();

  const modalBox = document.getElementById('motoModal');
  document.getElementById('motoModalTitle').textContent = `${regional} — ${data.oportunidade} O.S. elegíveis`;

  // Injeta filtros no modal
  const filterBar = `
    <div id="motoModalFilters" style="display:flex;flex-wrap:wrap;gap:8px;padding:12px 20px;background:var(--surface-card2);border-bottom:1px solid var(--surface-border);">
      <input type="text" id="motoModalBusca" placeholder="🔍 Buscar técnico ou protocolo..." oninput="App.filterMotoModal('${regional}')"
        style="flex:1;min-width:180px;padding:7px 12px;border-radius:8px;border:1px solid var(--surface-border);background:var(--surface-card);color:var(--text-primary);font-size:12px;font-family:var(--font);outline:none;">
      <select id="motoModalFiltroAssunto" onchange="App.filterMotoModal('${regional}')"
        style="padding:7px 12px;border-radius:8px;border:1px solid var(--surface-border);background:var(--surface-card);color:var(--text-primary);font-size:12px;font-family:var(--font);min-width:160px;">
        <option value="ALL">Todos os Assuntos</option>
        ${assuntos.map(a=>`<option value="${a}">${a}</option>`).join('')}
      </select>
      <select id="motoModalFiltroTipo" onchange="App.filterMotoModal('${regional}')"
        style="padding:7px 12px;border-radius:8px;border:1px solid var(--surface-border);background:var(--surface-card);color:var(--text-primary);font-size:12px;font-family:var(--font);min-width:140px;">
        <option value="ALL">Todos os Tipos</option>
        ${tipos.map(t=>`<option value="${t}">${t}</option>`).join('')}
      </select>
      <select id="motoModalOrdem" onchange="App.filterMotoModal('${regional}')"
        style="padding:7px 12px;border-radius:8px;border:1px solid var(--surface-border);background:var(--surface-card);color:var(--text-primary);font-size:12px;font-family:var(--font);">
        <option value="dia">Ordenar por Dia</option>
        <option value="tecnico">Ordenar por Técnico</option>
        <option value="assunto">Ordenar por Assunto</option>
      </select>
      <span id="motoModalCount" style="padding:7px 10px;font-family:var(--mono);font-size:11px;color:var(--text-tertiary);align-self:center;"></span>
    </div>`;

  // Injeta no modal (substitui a div de filtros se já existir)
  const existingFilters = document.getElementById('motoModalFilters');
  if (existingFilters) existingFilters.outerHTML = filterBar;
  else {
    const titleEl = document.querySelector('#motoModal .modal-title');
    if (titleEl) titleEl.insertAdjacentHTML('afterend', filterBar);
  }

  renderMotoModalBody(regional, data);
  modalBox.style.display = 'flex';
}

function renderMotoModalBody(regional, data) {
  const busca   = (document.getElementById('motoModalBusca')?.value || '').toLowerCase();
  const assunto = document.getElementById('motoModalFiltroAssunto')?.value || 'ALL';
  const tipo    = document.getElementById('motoModalFiltroTipo')?.value    || 'ALL';
  const ordem   = document.getElementById('motoModalOrdem')?.value         || 'dia';

  const getIcon = (txt) => {
    const t = txt.toLowerCase();
    if(t.includes('troca') || t.includes('remoção')) return '🔄';
    if(t.includes('conexão') || t.includes('sinal') || t.includes('fibra')) return '📡';
    if(t.includes('apps') || t.includes('streaming')) return '📺';
    if(t.includes('login') || t.includes('senha')) return '🔑';
    return '🔧';
  };

  let lista = [...data.oportunidadesDetalhadas];
  if (assunto !== 'ALL') lista = lista.filter(o => o.assunto === assunto);
  if (tipo    !== 'ALL') lista = lista.filter(o => o.tipoTecnico === tipo);
  if (busca)             lista = lista.filter(o =>
    o.tecnicoAtual.toLowerCase().includes(busca) || o.id.toLowerCase().includes(busca) || o.diagnostico.toLowerCase().includes(busca)
  );

  if      (ordem === 'dia')     lista.sort((a,b) => a.data - b.data);
  else if (ordem === 'tecnico') lista.sort((a,b) => a.tecnicoAtual.localeCompare(b.tecnicoAtual));
  else if (ordem === 'assunto') lista.sort((a,b) => a.assunto.localeCompare(b.assunto));

  const count = document.getElementById('motoModalCount');
  if (count) count.textContent = `${lista.length} O.S.`;

  const tbody = document.getElementById('motoModalBody');
  tbody.innerHTML = lista.length ? lista.map(os => `
    <tr style="border-bottom:1px solid var(--surface-border);">
      <td style="font-family:var(--mono);font-size:11px;color:var(--accent);padding:10px 12px;white-space:nowrap;">${os.id}</td>
      <td style="font-size:12px;padding:10px 12px;">
        <b style="color:var(--text-primary);display:block;">${os.tecnicoAtual}</b>
        <span style="font-size:10px;color:var(--text-tertiary);font-family:var(--mono);">${os.tipoTecnico}</span>
      </td>
      <td style="font-size:11px;padding:10px 12px;line-height:1.5;">
        <span style="color:var(--text-secondary);">${os.assunto}</span><br>
        <span style="font-size:10px;color:var(--text-tertiary);font-style:italic;">${os.diagnostico}</span>
      </td>
      <td style="font-family:var(--mono);font-size:11px;text-align:center;padding:10px 12px;color:var(--text-tertiary);">Dia ${os.data}</td>
    </tr>
  `).join('') : `<tr><td colspan="4" class="empty">Nenhuma O.S. com esses filtros.</td></tr>`;
}

function filterMotoModal(regional) {
  const data = state.motoOportunidades.find(r => r.regional === regional);
  if (data) renderMotoModalBody(regional, data);
}

function closeMotoDetails() {
  document.getElementById('motoModal').style.display = 'none';
}

function renderMotoOportunidades() {
  const container = document.getElementById('motoSection');
  if (!container) return;
  if (!state.rawExcelCache.length) {
    container.innerHTML = '<div class="empty" style="padding:60px;text-align:center;">Carregue uma planilha para ver as oportunidades do Suporte Moto.</div>';
    return;
  }
  calcMotoOportunidades();
  const data = state.motoOportunidades;
  if (!data.length) {
    container.innerHTML = '<div class="empty" style="padding:60px;text-align:center;">Nenhuma O.S. elegível para Suporte Moto encontrada.</div>';
    return;
  }

  const totalOS     = data.reduce((s, r) => s + r.total, 0);
  const totalOport  = data.reduce((s, r) => s + r.oportunidade, 0);
  const totalJaMoto = data.reduce((s, r) => s + r.jaMoto, 0);
  const pctOport    = totalOS > 0 ? Math.round(totalOport / totalOS * 100) : 0;

  // Monta opções únicas para os filtros globais da aba
  const todasRegionais = data.map(r => r.regional);
  const todosAssuntos  = [...new Set(data.flatMap(r => Object.keys(r.porAssunto)))].sort();
  const todosTipos     = [...new Set(data.flatMap(r => Object.keys(r.porTipo)))].sort();

  const getIcon = (txt) => {
    const t = txt.toLowerCase();
    if(t.includes('troca') || t.includes('remoção')) return '🔄';
    if(t.includes('conexão') || t.includes('sinal') || t.includes('fibra')) return '📡';
    if(t.includes('login') || t.includes('senha')) return '🔑';
    if(t.includes('apps') || t.includes('streaming')) return '📺';
    return '🔧';
  };

  // ── Filtros globais da aba ─────────────────────────────
  const filtersHTML = `
    <div style="background:var(--surface-card);border:1px solid var(--surface-border);border-radius:14px;padding:16px 20px;margin-bottom:18px;display:flex;flex-wrap:wrap;gap:10px;align-items:center;">
      <span style="font-family:var(--mono);font-size:9px;font-weight:700;letter-spacing:.12em;text-transform:uppercase;color:var(--text-tertiary);margin-right:4px;">Filtros</span>
      <select id="motoFiltroRegional" onchange="App.applyMotoFiltros()"
        style="padding:7px 12px;border-radius:8px;border:1px solid var(--surface-border);background:var(--surface-card2);color:var(--text-primary);font-size:12px;font-family:var(--font);">
        <option value="ALL">Todas as Regionais</option>
        ${todasRegionais.map(r=>`<option value="${r}">${r}</option>`).join('')}
      </select>
      <select id="motoFiltroAssunto" onchange="App.applyMotoFiltros()"
        style="padding:7px 12px;border-radius:8px;border:1px solid var(--surface-border);background:var(--surface-card2);color:var(--text-primary);font-size:12px;font-family:var(--font);min-width:180px;">
        <option value="ALL">Todos os Assuntos</option>
        ${todosAssuntos.map(a=>`<option value="${a}">${a}</option>`).join('')}
      </select>
      <select id="motoFiltroTipo" onchange="App.applyMotoFiltros()"
        style="padding:7px 12px;border-radius:8px;border:1px solid var(--surface-border);background:var(--surface-card2);color:var(--text-primary);font-size:12px;font-family:var(--font);">
        <option value="ALL">Todos os Tipos de Equipe</option>
        ${todosTipos.map(t=>`<option value="${t}">${TEAM_TYPES[t]||t}</option>`).join('')}
      </select>
      <select id="motoFiltroOrdem" onchange="App.applyMotoFiltros()"
        style="padding:7px 12px;border-radius:8px;border:1px solid var(--surface-border);background:var(--surface-card2);color:var(--text-primary);font-size:12px;font-family:var(--font);">
        <option value="oportunidade">Maior Oportunidade</option>
        <option value="pct">Maior % p/ Moto</option>
        <option value="jaMoto">Maior Já no Moto</option>
        <option value="total">Maior Total Elegível</option>
      </select>
      <button class="btn btn-ghost" onclick="App.resetMotoFiltros()" style="font-size:11px;padding:6px 12px;">Limpar</button>
      <span id="motoFiltrosCount" style="font-family:var(--mono);font-size:11px;color:var(--text-tertiary);"></span>
      <button class="btn btn-outline" onclick="App.exportMotoExcel()" style="font-size:11px;padding:6px 12px;margin-left:auto;display:flex;align-items:center;gap:4px;"><svg width="14" height="14" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"/></svg> Exportar Excel</button>
    </div>`;

  // ── KPIs de sumário ────────────────────────────────────
  const sumario = `<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:12px;margin-bottom:18px;">
    <div class="kpi" style="padding:16px;border-left:4px solid var(--accent);">
      <div class="kpi-val" style="color:var(--accent);font-size:22px;">${totalOS}</div>
      <div class="kpi-lbl">OS Elegíveis Total</div>
    </div>
    <div class="kpi" style="padding:16px;border-left:4px solid #EF4444;">
      <div class="kpi-val" style="color:#EF4444;font-size:22px;">${totalOport}</div>
      <div class="kpi-lbl">Oportunidades</div>
    </div>
    <div class="kpi" style="padding:16px;border-left:4px solid #10B981;">
      <div class="kpi-val" style="color:#10B981;font-size:22px;">${totalJaMoto}</div>
      <div class="kpi-lbl">Já no Sup. Rápido</div>
    </div>
    <div class="kpi" style="padding:16px;border-left:4px solid #F59E0B;">
      <div class="kpi-val" style="color:#F59E0B;font-size:22px;">${pctOport}%</div>
      <div class="kpi-lbl">% de Redistribuição</div>
    </div>
  </div>`;

  container.innerHTML = filtersHTML + sumario + `<div id="motoCardsContainer"></div>`;
  applyMotoFiltros();
}

function applyMotoFiltros() {
  const regional = document.getElementById('motoFiltroRegional')?.value || 'ALL';
  const assunto  = document.getElementById('motoFiltroAssunto')?.value  || 'ALL';
  const tipo     = document.getElementById('motoFiltroTipo')?.value     || 'ALL';
  const ordem    = document.getElementById('motoFiltroOrdem')?.value    || 'oportunidade';

  const getIcon = (txt) => {
    const t = txt.toLowerCase();
    if(t.includes('troca') || t.includes('remoção')) return '🔄';
    if(t.includes('conexão') || t.includes('sinal') || t.includes('fibra')) return '📡';
    if(t.includes('login') || t.includes('senha')) return '🔑';
    if(t.includes('apps') || t.includes('streaming')) return '📺';
    return '🔧';
  };

  let dados = state.motoOportunidades.map(r => {
    // Filtra oportunidadesDetalhadas por assunto e tipo
    let opsFiltradas = [...r.oportunidadesDetalhadas];
    if (assunto !== 'ALL') opsFiltradas = opsFiltradas.filter(o => o.assunto === assunto);
    if (tipo    !== 'ALL') opsFiltradas = opsFiltradas.filter(o => o.tipoTecnico === (TEAM_TYPES[tipo]||tipo) || o.tipoTecnico === tipo);

    // Recalcula porAssunto e porTipo filtrados
    const porAssuntoF = {};
    const porTipoF    = {};
    opsFiltradas.forEach(o => {
      porAssuntoF[o.assunto]    = (porAssuntoF[o.assunto]    || 0) + 1;
      porTipoF[o.tipoTecnico]   = (porTipoF[o.tipoTecnico]   || 0) + 1;
    });

    return { ...r, oportunidade: opsFiltradas.length, porAssunto: porAssuntoF, porTipo: porTipoF, _opsFiltradas: opsFiltradas };
  });

  if (regional !== 'ALL') dados = dados.filter(r => r.regional === regional);
  dados = dados.filter(r => r.oportunidade > 0 || r.jaMoto > 0);

  if      (ordem === 'oportunidade') dados.sort((a,b) => b.oportunidade - a.oportunidade);
  else if (ordem === 'pct')          dados.sort((a,b) => (b.oportunidade/Math.max(b.total,1)) - (a.oportunidade/Math.max(a.total,1)));
  else if (ordem === 'jaMoto')       dados.sort((a,b) => b.jaMoto - a.jaMoto);
  else if (ordem === 'total')        dados.sort((a,b) => b.total - a.total);

  const countEl = document.getElementById('motoFiltrosCount');
  if (countEl) countEl.textContent = `${dados.length} ${dados.length !== 1 ? 'regionais' : 'regional'} · ${dados.reduce((s,r)=>s+r.oportunidade,0)} OS`;

  const container = document.getElementById('motoCardsContainer');
  if (!container) return;

  if (!dados.length) {
    container.innerHTML = '<div class="empty" style="padding:40px;text-align:center;">Nenhuma regional com dados para esses filtros.</div>';
    return;
  }

  container.innerHTML = dados.map(r => {
    const totalFiltrado = r._opsFiltradas.length + r.jaMoto;
    const pct = totalFiltrado > 0 ? Math.round(r.oportunidade / totalFiltrado * 100) : 0;
    const barColor = pct >= 60 ? '#EF4444' : pct >= 30 ? '#F59E0B' : '#10B981';

    const assuntoRows = Object.entries(r.porAssunto).sort((a,b) => b[1]-a[1]).map(([a,n]) =>
      `<div style="display:flex;justify-content:space-between;align-items:center;padding:6px 0;border-bottom:1px dashed var(--surface-border);font-size:11px;">
        <span style="color:var(--text-secondary);display:flex;gap:6px;align-items:center;">${a}</span>
        <span style="font-family:var(--mono);font-weight:700;color:var(--text-primary);background:var(--surface-border);padding:2px 7px;border-radius:4px;">${n}</span>
      </div>`
    ).join('');

    const tipoRows = Object.entries(r.porTipo).sort((a,b) => b[1]-a[1]).map(([t,n]) =>
      `<span style="display:inline-flex;align-items:center;gap:4px;padding:3px 8px;border-radius:6px;font-size:10px;font-weight:600;background:rgba(255,255,255,0.05);color:var(--text-secondary);border:1px solid var(--surface-border);margin:2px;">${t} · ${n}</span>`
    ).join('');

    return `<div class="dc" style="margin-bottom:14px;border:1px solid ${barColor}35;box-shadow:0 4px 12px ${barColor}08;">
      <div class="dc-hdr" style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px;padding-bottom:12px;border-bottom:1px solid var(--surface-border);">
        <span class="dc-ttl" style="font-size:14px;display:flex;align-items:center;gap:8px;">
          <div style="width:9px;height:9px;border-radius:50%;background:${barColor};box-shadow:0 0 6px ${barColor};flex-shrink:0;"></div>
          ${r.regional}
        </span>
        <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;">
          <span style="font-family:var(--mono);font-size:10px;color:var(--text-tertiary);">${r.total} elegíveis</span>
          <span style="font-family:var(--mono);font-size:11px;font-weight:800;padding:3px 9px;border-radius:6px;background:${barColor}15;color:${barColor};border:1px solid ${barColor}40;">${pct}% → Moto</span>
          ${r.oportunidade > 0 ? `<button class="btn btn-outline" onclick="App.showMotoDetails('${r.regional}')" style="font-size:10px;padding:4px 10px;border-color:${barColor}50;color:${barColor};">
            Ver ${r.oportunidade} OS ›
          </button>` : ''}
        </div>
      </div>
      <div style="padding:14px 16px;">
        <div style="display:flex;gap:24px;margin-bottom:12px;flex-wrap:wrap;">
          <div style="text-align:center;">
            <div style="font-family:var(--mono);font-size:22px;font-weight:900;color:${barColor};">${r.oportunidade}</div>
            <div style="font-size:9px;color:var(--text-tertiary);text-transform:uppercase;letter-spacing:.5px;margin-top:2px;">Oportunidades</div>
          </div>
          <div style="text-align:center;">
            <div style="font-family:var(--mono);font-size:22px;font-weight:900;color:#10B981;">${r.jaMoto}</div>
            <div style="font-size:9px;color:var(--text-tertiary);text-transform:uppercase;letter-spacing:.5px;margin-top:2px;">Já no Sup. Rápido</div>
          </div>
        </div>
        <div style="height:6px;background:var(--surface-border);border-radius:4px;margin-bottom:12px;overflow:hidden;">
          <div style="height:100%;width:${pct}%;background:${barColor};border-radius:4px;transition:width .5s ease-out;"></div>
        </div>
        ${tipoRows ? `<div style="margin-bottom:10px;">${tipoRows}</div>` : ''}
        <div style="background:var(--surface-card2);border-radius:8px;padding:0 10px;">${assuntoRows}</div>
      </div>
    </div>`;
  }).join('');
}

function resetMotoFiltros() {
  ['motoFiltroRegional','motoFiltroAssunto','motoFiltroTipo','motoFiltroOrdem'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = id === 'motoFiltroOrdem' ? 'oportunidade' : 'ALL';
  });
  applyMotoFiltros();
}

function exportMotoExcel() {
  if (!state.motoOportunidades || !state.motoOportunidades.length) return alert('Sem dados de redistribuição para exportar.');
  
  const regional = document.getElementById('motoFiltroRegional')?.value || 'ALL';
  const assunto  = document.getElementById('motoFiltroAssunto')?.value  || 'ALL';
  const tipo     = document.getElementById('motoFiltroTipo')?.value     || 'ALL';

  let dados = state.motoOportunidades.map(r => {
    let opsFiltradas = [...r.oportunidadesDetalhadas];
    if (assunto !== 'ALL') opsFiltradas = opsFiltradas.filter(o => o.assunto === assunto);
    if (tipo    !== 'ALL') opsFiltradas = opsFiltradas.filter(o => o.tipoTecnico === (TEAM_TYPES[tipo]||tipo) || o.tipoTecnico === tipo);
    return { ...r, oportunidade: opsFiltradas.length, _opsFiltradas: opsFiltradas };
  });

  if (regional !== 'ALL') dados = dados.filter(r => r.regional === regional);
  dados = dados.filter(r => r.oportunidade > 0 || r.jaMoto > 0);
  dados.sort((a,b) => b.oportunidade - a.oportunidade);

  const wb = XLSX.utils.book_new();

  const resumoData = [
    ['Regional', 'Total Elegíveis', 'Já no Sup. Rápido', 'Oportunidades (Outras Equipes)', '% Oportunidade'],
    ...dados.map(r => {
      const totalFiltrado = r._opsFiltradas.length + r.jaMoto;
      const pct = totalFiltrado > 0 ? Math.round(r.oportunidade / totalFiltrado * 100) : 0;
      return [r.regional, r.total, r.jaMoto, r.oportunidade, `${pct}%`];
    })
  ];
  const wsResumo = XLSX.utils.aoa_to_sheet(resumoData);
  wsResumo['!cols'] = [{wch: 24}, {wch: 16}, {wch: 18}, {wch: 30}, {wch: 16}];
  XLSX.utils.book_append_sheet(wb, wsResumo, 'Resumo por Regional');

  const detalhadoData = [
    ['Regional', 'Protocolo', 'Data', 'Técnico Atual', 'Tipo Equipe', 'Assunto', 'Diagnóstico'],
    ...dados.flatMap(r => r._opsFiltradas.map(os => [r.regional, os.id, os.data, os.tecnicoAtual, os.tipoTecnico, os.assunto, os.diagnostico]))
  ];
  const wsDetalhado = XLSX.utils.aoa_to_sheet(detalhadoData);
  wsDetalhado['!cols'] = [{wch: 24}, {wch: 16}, {wch: 10}, {wch: 24}, {wch: 18}, {wch: 32}, {wch: 42}];
  XLSX.utils.book_append_sheet(wb, wsDetalhado, 'Oportunidades Detalhadas');

  XLSX.writeFile(wb, `SGO_Oportunidades_Moto_${state.activeMonthYear}.xlsx`);
}

function renderTechRow(tk, data, dIM, fdw, cidade) {
  const isAux = data.tipo === 'AUXILIAR';
  const meta = isAux ? 0 : (state.appSettings.metasDiarias[data.tipo]||5);
  const nome = state.teamData[tk].originalName;
  const tipo = TEAM_TYPES[data.tipo]||data.tipo;
  const ce = meta * (data.tipo==='TECNICO 12/36H' ? 15 : 24);
  const cb = Math.max(0, (meta - 1) * (data.tipo==='TECNICO 12/36H' ? 15 : 24));
  const cm2 = Math.max(0, (meta - 2) * (data.tipo==='TECNICO 12/36H' ? 15 : 24));
  const auxBadge = isAux && data.total > 0 ? `<span class="aux-badge">âœ“ ${data.total} OS</span>` : '';
  const initials = nome.split(' ').filter(w => w.length > 2).slice(0, 2).map(w => w[0]).join('');
  const cityStr = cidade ? `,'${cidade.replace(/'/g, "\\\\'")}'` : ",''";
  const tkStr = tk ? `,'${tk.replace(/'/g, "\\\\'")}'` : '';
  const monthlyItems = state.globalRawData.filter(i =>
    i.monthStr === state.activeMonthYear &&
    i.cidade === cidade &&
    i.techKey === tk
  );
  const tmaStats = getTmaStats(monthlyItems);

  let row = `<tr class="mat-row${isAux?' aux-row':''}"><td class="cn">
    <div class="cn-inner">
      <div class="cn-avatar" style="background:rgba(255,255,255,0.06);color:#7A92AA;border:1px solid rgba(255,255,255,0.10);">${initials}</div>
      <div class="cn-info">
        <div class="cn-name">${nome}${auxBadge}</div>
        <span class="cn-badge" style="background:rgba(255,255,255,0.05);color:#7A92AA;border:1px solid rgba(255,255,255,0.10);">${tipo}</span>
      </div>
    </div>
  </td>`;

  let wt = 0, cur = 0;
  for (let i = 0; i < fdw; i++) {
    const dow = i;
    row += `<td class="${(dow===0||dow===6)?'cwknd':''}"></td>`;
    cur++;
  }
  for (let d = 1; d <= dIM; d++) {
    if (cur > 6) { row += `<td class="ctot">${wt>0?wt:''}</td>`; wt = 0; cur = 0; }
    const v = data.dias[d] || 0; wt += v;
    const dow = (fdw + d - 1) % 7;
    const wk = dow === 0 || dow === 6;
    const vc = isAux ? (v > 0 ? 'vg' : 'vd') : vC(v, meta, dow, data.tipo);
    row += `<td class="${vc}${wk?' cwknd':''}${isAux?' aux-cell':''}" ${v>0?`style="cursor:pointer;transition:transform 0.1s;" onmouseover="this.style.transform='scale(1.15)'" onmouseout="this.style.transform='none'" onclick="App.showDayDetails(${d}${cityStr}${tkStr})" title="Ver ${v} O.S. do dia ${d}"`:''}>${v>0?v:''}</td>`;
    cur++;
  }
  if (cur > 0) {
    for (let i = 0; i < 7 - cur; i++) row += `<td></td>`;
    row += `<td class="ctot">${wt>0?wt:''}</td>`;
  }

  const onClickTot = `style="cursor:pointer;transition:opacity 0.2s;" onmouseover="this.style.opacity='0.75'" onmouseout="this.style.opacity='1'" onclick="App.showDayDetails(null${cityStr}${tkStr})" title="Ver todas as ${data.total} O.S. no mes${tmaStats.count ? ` | TMA medio ${formatMinutesCompact(tmaStats.avgMinutes)} | TMA somado ${formatMinutesCompact(tmaStats.totalMinutes)}` : ''}"`;
  if (isAux) {
    row += `<td class="ctos ${data.total>0?'tg':'vd'}" ${data.total>0?onClickTot:''}>${data.total||'--'}</td>`;
    row += `<td class="ccap" colspan="3" style="color:var(--text-3);font-size:10px;text-align:center;font-style:italic;">sem meta</td></tr>`;
  } else {
    row += `<td class="ctos ${tC(data.total, ce, cb, cm2)}" ${data.total>0?onClickTot:''}>${data.total}</td>`;
    row += `<td class="ccap cg">${ce}</td><td class="ccap cb">${cb>0?cb:0}</td><td class="ccap cy">${cm2>0?cm2:0}</td></tr>`;
  }
  return row;
}

function showDayDetails(day, cidade, techKey) {
  if (!state.globalRawData.length) return;

  let filtered = state.currentFiltered || [];
  if (cidade && techKey) {
    filtered = state.currentFilterMeta?.techItems?.[`${cidade}::${techKey}`] || [];
  } else if (cidade) {
    filtered = state.currentFilterMeta?.cityItems?.[cidade] || [];
  }
  if (day !== null) filtered = filtered.filter(i => i.day === day);
  if (!filtered.length) {
    showToast('Nenhuma O.S. encontrada para este dia.', 'info');
    return;
  }

  state.currentDayDetails = filtered;
  state.currentDayDetailsDay = day;

  let overlay = document.getElementById('dayDetailsModalOverlay');
  if (overlay) overlay.remove();
  overlay = document.createElement('div');
  overlay.id = 'dayDetailsModalOverlay';
  overlay.style.cssText = 'position:fixed;inset:0;z-index:9999;background:rgba(15,23,42,.6);display:flex;align-items:center;justify-content:center;backdrop-filter:blur(8px);animation:fadeIn .15s ease;padding:16px;';

  const tmaStats = getTmaStats(filtered);
  const teamDayTmaSplitStats = getTeamDayTmaSplitStats(filtered);
  const isSingleDayDetail = day !== null;
  const validTmaValues = filtered.map(getOsTmaMinutes).filter(v => v != null && !isNaN(v));
  const thresholdMinutes = getTmaAlertThresholdMinutes(filtered, techKey);
  const aboveThresholdCount = validTmaValues.filter(v => v >= thresholdMinutes).length;
  const aboveThresholdPct = tmaStats.count ? Math.round((aboveThresholdCount / tmaStats.count) * 100) : 0;
  const minTmaStr = tmaStats.count ? formatMinutesCompact(tmaStats.minMinutes) : '--';
  const activeTmaDays = new Set(filtered.filter(item => getOsTmaMinutes(item) != null).map(item => item.dateStr || `${item.monthStr}-${String(item.day).padStart(2, '0')}`)).size;
  const invalidTmaCount = Math.max(filtered.length - tmaStats.count, 0);
  const avgTmaStr = tmaStats.count ? formatMinutesCompact(tmaStats.avgMinutes) : '--';
  const sumTmaStr = tmaStats.count ? formatMinutesCompact(tmaStats.totalMinutes) : '--';
  const dailyAvgTmaStr = activeTmaDays > 0 ? formatMinutesCompact(tmaStats.totalMinutes / activeTmaDays) : '--';
  const medianTmaStr = tmaStats.count ? formatMinutesCompact(tmaStats.medianMinutes) : '--';
  const weekdayDivisor = isSingleDayDetail
    ? teamDayTmaSplitStats.weekday.bucketCount
    : teamDayTmaSplitStats.weekday.equivalentBucketCount;
  const weekendDivisor = isSingleDayDetail
    ? teamDayTmaSplitStats.weekend.bucketCount
    : teamDayTmaSplitStats.weekend.equivalentBucketCount;
  const weekdayTeamDayAvgStr = weekdayDivisor ? formatMinutesCompact(teamDayTmaSplitStats.weekday.totalMinutes / weekdayDivisor) : '--';
  const weekendTeamDayAvgStr = weekendDivisor ? formatMinutesCompact(teamDayTmaSplitStats.weekend.totalMinutes / weekendDivisor) : '--';
  const weekdayScopeMeta = isSingleDayDetail
    ? `${formatMinutesCompact(teamDayTmaSplitStats.weekday.totalMinutes)} divididos por ${weekdayDivisor || 0} equipes ativas no dia`
    : `${formatMinutesCompact(teamDayTmaSplitStats.weekday.totalMinutes)} divididos por ${weekdayDivisor || 0} dias-equipe uteis`;
  const weekendScopeMeta = isSingleDayDetail
    ? `${formatMinutesCompact(teamDayTmaSplitStats.weekend.totalMinutes)} divididos por ${weekendDivisor || 0} equipes ativas no dia`
    : `${formatMinutesCompact(teamDayTmaSplitStats.weekend.totalMinutes)} divididos por ${weekendDivisor || 0} dias-equipe no fim de semana · sabado normal vale meio dia`;
  let subtitle = `Regional: ${escapeHtml(cidade || 'Todas')}`;
  if (techKey) subtitle += ` | Tecnico: ${escapeHtml(state.teamData[techKey]?.originalName || techKey)}`;

  overlay.innerHTML = `
    <div style="background:var(--surface-card);border:1px solid var(--surface-border);border-radius:16px;max-width:950px;width:100%;max-height:85vh;display:flex;flex-direction:column;box-shadow:0 16px 48px rgba(0,0,0,.5);animation:scaleIn .15s ease;overflow:hidden;">
      <div class="tma-modal-header">
        <div class="tma-modal-summary">
          <div class="tma-modal-title-row">
            <div style="font-size:16px;font-weight:800;color:var(--text-primary);">Detalhamento de O.S. - ${day !== null ? `Dia ${day}` : 'Mes Completo'}</div>
            <button id="closeDayDetailsModal" class="tma-modal-close" type="button">X</button>
          </div>
          <div style="font-size:12px;color:var(--text-tertiary);margin-top:2px;">${subtitle} | <b style="color:var(--text-secondary);">${filtered.length} O.S.</b> analisadas</div>
          <div class="tma-highlight-grid">
            <div class="tma-highlight-card is-blue">
              <div class="tma-highlight-label">OS com TMA valido</div>
              <div class="tma-highlight-value">${tmaStats.count}</div>
              <div class="tma-highlight-meta">${invalidTmaCount} OS ficaram de fora do calculo</div>
            </div>
            <div class="tma-highlight-card is-indigo">
              <div class="tma-highlight-label">Media por OS</div>
              <div class="tma-highlight-value">${avgTmaStr}</div>
              <div class="tma-highlight-meta">Soma dos TMAs dividida por ${tmaStats.count} OS validas</div>
            </div>
            <div class="tma-highlight-card is-cyan">
              <div class="tma-highlight-label">${isSingleDayDetail ? 'Media por equipe no dia util' : 'Media por equipe em dia util'}</div>
              <div class="tma-highlight-value">${weekdayTeamDayAvgStr}</div>
              <div class="tma-highlight-meta">${weekdayScopeMeta}</div>
            </div>
            <div class="tma-highlight-card is-green">
              <div class="tma-highlight-label">Media por equipe no fim de semana</div>
              <div class="tma-highlight-value">${weekendTeamDayAvgStr}</div>
              <div class="tma-highlight-meta">${weekendScopeMeta}</div>
            </div>
            <div class="tma-highlight-card is-warning">
              <div class="tma-highlight-label">OS acima de ${thresholdMinutes >= 120 ? '2h' : '1h'}</div>
              <div class="tma-highlight-value">${tmaStats.count ? `${aboveThresholdPct}% acima de ${thresholdMinutes >= 120 ? '2h' : '1h'}` : '--'}</div>
              <div class="tma-highlight-meta">${aboveThresholdCount} OS passaram desse tempo no periodo</div>
            </div>
          </div>
        </div>
        <div class="tma-modal-toolbar">
          <input type="text" id="dayDetailsSearch" placeholder="Buscar cliente, OS, assunto..." oninput="App.filterDayDetails()" style="padding:7px 12px;border-radius:8px;border:1px solid var(--surface-border);background:var(--surface-card);color:var(--text-primary);font-size:12px;font-family:var(--font);outline:none;min-width:220px;">
          <select id="dayDetailsSort" onchange="App.filterDayDetails()" style="padding:7px 12px;border-radius:8px;border:1px solid var(--surface-border);background:var(--surface-card);color:var(--text-primary);font-size:12px;font-family:var(--font);outline:none;">
            <option value="crono">Ordem Cronologica</option>
            <option value="assunto">Por Assunto</option>
            <option value="tma">Maior TMA</option>
          </select>
          <button id="openTmaHelpPanel" class="tma-help-btn" type="button" title="Entender como esses calculos funcionam">?</button>
        </div>
      </div>
      <div style="overflow-y:auto;padding:0;flex:1;"><table style="width:100%;border-collapse:collapse;text-align:left;"><thead style="background:var(--surface-card2);position:sticky;top:0;box-shadow:0 1px 2px rgba(0,0,0,0.05);z-index:2;"><tr><th style="padding:12px;font-size:10px;color:var(--text-tertiary);text-transform:uppercase;font-family:var(--mono);">Protocolo</th>${day === null ? `<th style="padding:12px;font-size:10px;color:var(--text-tertiary);text-transform:uppercase;font-family:var(--mono);">Data</th>` : ''}<th style="padding:12px;font-size:10px;color:var(--text-tertiary);text-transform:uppercase;font-family:var(--mono);">Tecnico</th><th style="padding:12px;font-size:10px;color:var(--text-tertiary);text-transform:uppercase;font-family:var(--mono);">Assunto / Diagnostico</th><th style="padding:12px;font-size:10px;color:var(--text-tertiary);text-transform:uppercase;font-family:var(--mono);">TMA</th><th style="padding:12px;font-size:10px;color:var(--text-tertiary);text-transform:uppercase;font-family:var(--mono);">Cliente/Login</th></tr></thead><tbody id="dayDetailsBody"></tbody></table></div>
    </div>`;
  document.body.appendChild(overlay);
  const closeBtn = document.getElementById('closeDayDetailsModal');
  const helpBtn = document.getElementById('openTmaHelpPanel');
  const closeDayDetailsOverlay = () => {
    overlay.remove();
    document.getElementById('tmaHelpOverlay')?.remove();
  };
  closeBtn.onclick = closeDayDetailsOverlay;
  closeBtn.onmouseover = () => { closeBtn.style.background = 'var(--hover-white-08)'; closeBtn.style.color = 'var(--text-primary)'; };
  closeBtn.onmouseout = () => { closeBtn.style.background = 'var(--hover-white-04)'; closeBtn.style.color = 'var(--text-tertiary)'; };
  if (helpBtn) helpBtn.onclick = () => showTmaHelpPanel();
  overlay.onclick = e => { if (e.target === overlay) closeDayDetailsOverlay(); };
  filterDayDetails();
}

function showTmaHelpPanel() {
  let overlay = document.getElementById('tmaHelpOverlay');
  if (overlay) overlay.remove();

  overlay = document.createElement('div');
  overlay.id = 'tmaHelpOverlay';
  overlay.className = 'tma-help-overlay';
  overlay.innerHTML = `
    <div class="tma-help-panel">
      <div class="tma-help-header">
        <div>
          <div class="tma-help-eyebrow">Como ler os cards</div>
          <div class="tma-help-title">Explicacao dos calculos de TMA</div>
        </div>
        <button id="closeTmaHelpPanel" class="tma-help-close" type="button">X</button>
      </div>
      <div class="tma-help-body">
        <div class="tma-help-card">
          <div class="tma-help-card-title">OS com TMA valido</div>
          <div class="tma-help-card-text">Conta quantas O.S. tinham data de inicio e fim validas para entrar no calculo.</div>
        </div>
        <div class="tma-help-card">
          <div class="tma-help-card-title">Media por OS</div>
          <div class="tma-help-card-text">Soma todos os TMAs validos e divide pela quantidade de O.S. validas.</div>
        </div>
        <div class="tma-help-card">
          <div class="tma-help-card-title">Media por equipe em dia util</div>
          <div class="tma-help-card-text">Soma o TMA dos dias uteis e divide pela quantidade de dias-equipe uteis.</div>
          <div class="tma-help-card-note">Um dia-equipe util = uma equipe com atendimento em um dia de segunda a sexta.</div>
        </div>
        <div class="tma-help-card">
          <div class="tma-help-card-title">Media por equipe no fim de semana</div>
          <div class="tma-help-card-text">Soma o TMA do fim de semana e divide pelos dias-equipe considerados no sabado e domingo.</div>
          <div class="tma-help-card-note">No consolidado do mes: sabado de equipe normal vale meio dia e plantao 12/36 vale um dia inteiro no sabado e domingo. Quando voce detalha um unico dia, o calculo divide pelas equipes ativas naquele dia, sem dobrar o sabado.</div>
        </div>
        <div class="tma-help-card">
          <div class="tma-help-card-title">OS acima do limite</div>
          <div class="tma-help-card-text">Mostra o percentual de O.S. que passaram de 2h, ou 1h para equipes de suporte.</div>
        </div>
      </div>
    </div>`;

  document.body.appendChild(overlay);
  const close = document.getElementById('closeTmaHelpPanel');
  if (close) close.onclick = () => overlay.remove();
  overlay.onclick = e => { if (e.target === overlay) overlay.remove(); };
}

function filterDayDetails() {
  const tbody = document.getElementById('dayDetailsBody');
  if (!tbody) return;
  const search = (document.getElementById('dayDetailsSearch')?.value || '').toLowerCase();
  const sort = document.getElementById('dayDetailsSort')?.value || 'crono';
  const day = state.currentDayDetailsDay;

  let list = [...state.currentDayDetails];
  if (search) {
    list = list.filter(os =>
      (os.osId || '').toLowerCase().includes(search) ||
      (os.cliente || '').toLowerCase().includes(search) ||
      (os.login || '').toLowerCase().includes(search) ||
      (os.assunto || '').toLowerCase().includes(search) ||
      (os.diagnostico || '').toLowerCase().includes(search) ||
      (os.nomeOriginal || os.nome || '').toLowerCase().includes(search)
    );
  }

  if (sort === 'crono') {
    list.sort((a, b) => (day === null ? a.day - b.day : 0) || (a.nomeOriginal || a.nome).localeCompare(b.nomeOriginal || b.nome));
  } else if (sort === 'assunto') {
    list.sort((a, b) => (a.assunto || '').localeCompare(b.assunto || '') || (day === null ? a.day - b.day : 0));
  } else if (sort === 'tma') {
    list.sort((a, b) => (getOsTmaMinutes(b) ?? -1) - (getOsTmaMinutes(a) ?? -1));
  }

  tbody.innerHTML = list.length ? list.map((os, idx) => {
    const tmaMin = getOsTmaMinutes(os);
    return `
    <tr style="border-bottom:1px solid var(--surface-border);background:${idx % 2 === 0 ? 'transparent' : 'var(--surface-card2)'}">
      <td style="padding:10px 12px;font-family:var(--mono);font-size:11px;color:var(--accent);white-space:nowrap;">${escapeHtml(os.osId)}</td>
      ${day === null ? `<td style="padding:10px 12px;font-family:var(--mono);font-size:11px;color:var(--text-secondary);white-space:nowrap;">${os.day}/${os.monthStr.split('-')[1]}</td>` : ''}
      <td style="padding:10px 12px;font-size:12px;white-space:nowrap;">
        <b style="color:var(--text-primary);display:block;">${escapeHtml(os.nomeOriginal || os.nome)}</b>
        <span style="font-size:10px;color:var(--text-tertiary);font-family:var(--mono);">${escapeHtml(TEAM_TYPES[os.tipo] || os.tipo)}</span>
      </td>
      <td style="padding:10px 12px;font-size:11px;line-height:1.4;">
        <span style="color:var(--text-secondary);">${escapeHtml(os.assunto)}</span><br>
        <span style="font-size:10px;color:var(--text-tertiary);font-style:italic;">${escapeHtml(os.diagnostico)}</span>
      </td>
      <td style="padding:10px 12px;font-family:var(--mono);font-size:11px;color:var(--blue);font-weight:700;">${tmaMin !== null ? formatMinutesCompact(tmaMin) : '--'}</td>
      <td style="padding:10px 12px;font-size:11px;color:var(--text-secondary);">${escapeHtml(os.cliente || os.login)}</td>
    </tr>
  `}).join('') : `<tr><td colspan="${day === null ? '6' : '5'}" style="padding:32px;text-align:center;color:var(--text-tertiary);font-size:12px;">Nenhuma O.S. encontrada para esta busca.</td></tr>`;
}

function syncCapacidadesContext() {
  const capSection = document.getElementById('capSection');
  if (!capSection) return;

  const titles = capSection.querySelectorAll('.cap-section-title');
  const totalLabel = fc => fc === 'ALL' ? 'TOTAL GERAL' : `TOTAL ${fc}`;
  const city = state.selectedCityTab || 'ALL';

  if (titles[0]) titles[0].textContent = city === 'ALL' ? 'Capacidade por Filial e Tipo de Equipe' : `Capacidade da Regional ${city}`;
  const sideTitleSpan = titles[1]?.querySelector('span');
  if (sideTitleSpan) sideTitleSpan.textContent = city === 'ALL' ? 'Resumo Geral' : `Resumo da Regional ${city}`;

  const totalCell = capSection.querySelector('.cap-total .cap-cidade');
  if (totalCell) totalCell.textContent = totalLabel(city);
}

window.App = {
  switchTab,
  toggleDark: () => toggleDark(() => { if (state.globalRawData.length) applyFilters(); }),
  toggleSidebar,
  closeSidebar,
  toggleDotsMenu,
  closeDotsMenu,
  toggleShowCap,
  toggleShowRank,
  syncDotsMenuState,
  openMobileFilterSheet,
  closeMobileFilterSheet,
  syncMobileFilters,
  toggleCompactMode,
  parseFileWithWorker,
  changeActiveMonth,
  applyFilters,
  filterMatrixBySearch,
  clearSearch,
  exportToExcel,
  exportRecorrenciaExcel,
  exportRecorrenciaJson,
  exportMotoExcel,
  updateRecurrenceSetting,
  renderRecorrenciaClientes,
  renderCapacidades,
  renderMotoOportunidades,
  toggleChartDOW: (dow) => {
    const idx = state.chartDOW.indexOf(dow);
    if (idx > -1) {
      if (state.chartDOW.length > 1) state.chartDOW.splice(idx, 1);
    } else {
      state.chartDOW.push(dow);
    }
    renderAdvancedDailyChart();
  },
  showDayDetails,
  filterDayDetails,
  showRecurrenceDetails,
  resetRecorrenciaFiltros,
  selectCityTab: city => { selectCityTab(city); applyFilters(); },
  openAddTechModal,
  editTech,
  closeTechModal,
  saveTechForm:           () => saveTechForm(syncEcosystem),
  deleteTech:             key => deleteTech(key, syncEcosystem),
  deleteTechConfirm:      key => showConfirm(`Remover <b>${state.teamData[key]?.originalName||key}</b>?`, () => { deleteTech(key); }),
  acceptReclassification: (key, tipo) => acceptReclassification(key, tipo, syncEcosystem),
  filterTeamTable:        renderTeamTable,
  exportTeamData,
  importTeamData:         event => importTeamData(event, syncEcosystem),
  saveGlobalSettings,
  showMotoDetails,
  filterMotoModal,
  applyMotoFiltros,
  resetMotoFiltros,
  closeMotoDetails
};

/* ═══════════════════════════════════════════════════════════
   INIT
   ═══════════════════════════════════════════════════════════ */
document.addEventListener('DOMContentLoaded', async () => {
  try {
    state.isDark = loadDarkPref();
    if (state.isDark) applyDarkTheme(true);

    initLocalStorage();

    if (!state.appSettings.metaSabadoPct) state.appSettings.metaSabadoPct = 50;
    if (!state.appSettings.metasDiarias['SUPORTE CARRO']) state.appSettings.metasDiarias['SUPORTE CARRO'] = 9;
    try {
      state.showCap  = localStorage.getItem('sgo_show_cap')  !== '0';
      state.showRank = localStorage.getItem('sgo_show_rank') !== '0';
    } catch(e) { state.showCap = true; state.showRank = true; }

    initClock();
    if (typeof initNbClock === 'function') initNbClock();
    initWorker();
    initMobileSheetSwipe();
    loadSettingsToUI();
    renderTeamTable();
    populateFilters();
    await restoreCachedUpload();

    const pm = document.getElementById('pageMetaDate');
    if (pm) pm.textContent = new Date().toLocaleDateString('pt-BR', { weekday:'long', day:'2-digit', month:'long', year:'numeric' });
    if (isMobile()) state.isCompactMode = true;

  } catch(err) {
    console.error('[SGO] Erro na inicialização:', err);
    if (typeof showToast === 'function') showToast('Erro ao iniciar: ' + err.message, 'error', 8000);
  }
});


