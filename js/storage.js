/**
 * SGO — storage.js
 * Leitura e escrita no localStorage, importação e exportação de JSON.
 */

import { state, DEFAULT_SETTINGS, TEAM_TYPES } from './state.js';

const KEYS = {
  settings: 'sgo_settings_pro_v3',
  team:     'sgo_team_pro_v3',
  gemini:   'sgo_gemini_key',
  dark:     'sgo_dark'
};

// Mapeamento de tipos legados → novos
const TYPE_MIGRATION = {
  COMERCIAL: 'INSTALAÇÃO CIDADE',
  PLANTAO:   'TECNICO 12/36H',
  SUPORTE:   'SUPORTE MOTO',
  FAZ_TUDO:  'FAZ TUDO'
};

function migrateTypes(teamObj) {
  let changed = false;
  Object.values(teamObj).forEach(t => {
    const newType = TYPE_MIGRATION[t.tipo];
    if (newType) { t.tipo = newType; changed = true; }
  });
  return changed;
}

export function initLocalStorage() {
  // Settings
  const rawSettings = localStorage.getItem(KEYS.settings);
  if (rawSettings) {
    state.appSettings = JSON.parse(rawSettings);
  } else {
    const oldSettings = localStorage.getItem('sgo_settings_pro_v2');
    if (oldSettings) {
      const p = JSON.parse(oldSettings);
      state.appSettings = {
        metasDiarias: {
          "INSTALAÇÃO CIDADE": p.metasDiarias.COMERCIAL || 5,
          "TECNICO 12/36H":    p.metasDiarias.PLANTAO   || 4,
          "SUPORTE MOTO":      p.metasDiarias.SUPORTE    || 8,
          "RURAL":             p.metasDiarias.RURAL      || 3,
          "FAZ TUDO":          p.metasDiarias.FAZ_TUDO   || 5
        }
      };
    } else {
      state.appSettings = JSON.parse(JSON.stringify(DEFAULT_SETTINGS));
    }
  }

  // Team data
  const rawTeam = localStorage.getItem(KEYS.team) || localStorage.getItem('sgo_team_pro_v2');
  state.teamData = rawTeam ? JSON.parse(rawTeam) : {};
  if (migrateTypes(state.teamData)) {
    localStorage.setItem(KEYS.team, JSON.stringify(state.teamData));
  }
}

export function saveSettings() {
  localStorage.setItem(KEYS.settings, JSON.stringify(state.appSettings));
}

export function saveTeam() {
  localStorage.setItem(KEYS.team, JSON.stringify(state.teamData));
}

export function saveGeminiKey(key) {
  localStorage.setItem(KEYS.gemini, key);
}

export function loadGeminiKey() {
  return localStorage.getItem(KEYS.gemini) || '';
}

export function saveDarkPref(isDark) {
  localStorage.setItem(KEYS.dark, isDark ? '1' : '0');
}

export function loadDarkPref() {
  return localStorage.getItem(KEYS.dark) === '1';
}

// ── Import / Export JSON de equipes ───────────────────────
export function exportTeamData() {
  if (!Object.keys(state.teamData).length) return alert('Sem equipes cadastradas.');
  const a = document.createElement('a');
  a.href = 'data:text/json;charset=utf-8,' + encodeURIComponent(JSON.stringify(state.teamData, null, 2));
  a.download = 'equipes_sgo.json';
  a.click();
}

export function importTeamData(event, onSuccess) {
  const file = event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = ev => {
    try {
      const data = JSON.parse(ev.target.result);
      migrateTypes(data);
      state.teamData = data;
      saveTeam();
      alert('Importado com sucesso!');
      if (onSuccess) onSuccess();
    } catch {
      alert('Arquivo JSON inválido.');
    }
  };
  reader.readAsText(file);
}
