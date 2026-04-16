/**
 * SGO — app.js
 * Ponto de entrada principal. Inicialização e exposição global de handlers.
 */

import { state }                        from './state.js';
import { initLocalStorage, saveSettings, loadGeminiKey, exportTeamData, importTeamData } from './storage.js';
import { isMobile, initClock, applyDarkTheme, toggleDark, switchTab,
         openMobileFilterSheet, closeMobileFilterSheet, initMobileSheetSwipe,
         toggleCompactMode, syncMobileFilters }                              from './ui.js';
import { populateFilters, renderTeamTable, selectCityTab,
         openAddTechModal, editTech, closeTechModal, saveTechForm, deleteTech,
         acceptReclassification, limparNome }                                from './team.js';
import { initWorker, parseFileWithWorker, applyFilters, syncEcosystem,
         filterMatrixBySearch, clearSearch }                                  from './upload.js';
import { exportToExcel }                                                      from './matrix.js';
import { saveGeminiKeyUI, analyzeWithGemini }                                from './analysis.js';
import { loadDarkPref, saveDarkPref }                                         from './storage.js';

// ── Inicialização ─────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  // Tema
  state.isDark = loadDarkPref();
  if (state.isDark) applyDarkTheme(true);

  // Core
  initLocalStorage();
  initClock();
  initWorker();
  initMobileSheetSwipe();

  // Carregar UI
  loadSettingsToUI();
  renderTeamTable();
  populateFilters();

  // Gemini key
  const geminiInput = document.getElementById('geminiApiKey');
  if (geminiInput) geminiInput.value = loadGeminiKey();

  // Data atual no header
  const pm = document.getElementById('pageMetaDate');
  if (pm) pm.textContent = new Date().toLocaleDateString('pt-BR', {
    weekday: 'long', day: '2-digit', month: 'long', year: 'numeric'
  });

  // Auto-compact no mobile
  if (isMobile()) state.isCompactMode = true;
});

// ── Settings UI ───────────────────────────────────────────
function loadSettingsToUI() {
  const m = state.appSettings.metasDiarias;
  document.getElementById('cfgMetaDiaComercial').value = m['INSTALAÇÃO CIDADE'] || 5;
  document.getElementById('cfgMetaDiaPlantao').value   = m['TECNICO 12/36H']    || 4;
  document.getElementById('cfgMetaDiaSuporte').value   = m['SUPORTE MOTO']      || 8;
  document.getElementById('cfgMetaDiaRural').value     = m['RURAL']             || 3;
  document.getElementById('cfgMetaDiaFazTudo').value   = m['FAZ TUDO']          || 5;
}

function saveGlobalSettings() {
  const m = state.appSettings.metasDiarias;
  m['INSTALAÇÃO CIDADE'] = parseInt(document.getElementById('cfgMetaDiaComercial').value) || 5;
  m['TECNICO 12/36H']    = parseInt(document.getElementById('cfgMetaDiaPlantao').value)   || 4;
  m['SUPORTE MOTO']      = parseInt(document.getElementById('cfgMetaDiaSuporte').value)   || 8;
  m['RURAL']             = parseInt(document.getElementById('cfgMetaDiaRural').value)     || 3;
  m['FAZ TUDO']          = parseInt(document.getElementById('cfgMetaDiaFazTudo').value)   || 5;
  saveSettings();
  if (state.globalRawData.length) applyFilters();
  alert('Metas salvas com sucesso.');
}

// ── Namespace global (chamado pelo HTML) ───────────────────
window.App = {
  // Nav
  switchTab,
  toggleDark:              () => toggleDark(() => { if (state.globalRawData.length) applyFilters(); }),

  // Mobile
  openMobileFilterSheet,
  closeMobileFilterSheet,
  syncMobileFilters,
  toggleCompactMode,

  // Upload
  parseFileWithWorker,
  applyFilters,
  filterMatrixBySearch,
  clearSearch,
  exportToExcel,

  // City tabs
  selectCityTab:           city => { selectCityTab(city); applyFilters(); },

  // Team CRUD
  openAddTechModal,
  editTech,
  closeTechModal,
  saveTechForm:            () => saveTechForm(syncEcosystem),
  deleteTech:              key => deleteTech(key, syncEcosystem),
  acceptReclassification:  (key, tipo) => acceptReclassification(key, tipo, syncEcosystem),
  filterTeamTable:         renderTeamTable,
  exportTeamData,
  importTeamData:          event => importTeamData(event, syncEcosystem),

  // Settings
  saveGlobalSettings,

  // Gemini / IA
  saveGeminiKey:           saveGeminiKeyUI,
  analyzeWithGemini,
};

// Retrocompat: algumas funções ainda chamadas por onclick="func()" direto no HTML legado
window.switchTab              = App.switchTab;
window.toggleDark             = App.toggleDark;
window.openMobileFilterSheet  = App.openMobileFilterSheet;
window.closeMobileFilterSheet = App.closeMobileFilterSheet;
window.syncMobileFilters      = App.syncMobileFilters;
window.toggleCompactMode      = App.toggleCompactMode;
window.applyFilters           = App.applyFilters;
window.filterMatrixBySearch   = App.filterMatrixBySearch;
window.clearSearch            = App.clearSearch;
window.exportToExcel          = App.exportToExcel;
window.parseFileWithWorker    = App.parseFileWithWorker;
window.openAddTechModal       = App.openAddTechModal;
window.closeTechModal         = App.closeTechModal;
window.saveTechForm           = App.saveTechForm;
window.deleteTech             = App.deleteTech;
window.editTech               = App.editTech;
window.acceptReclassification = App.acceptReclassification;
window.filterTeamTable        = App.filterTeamTable;
window.exportTeamData         = App.exportTeamData;
window.importTeamData         = App.importTeamData;
window.saveGlobalSettings     = App.saveGlobalSettings;
window.saveGeminiKey          = App.saveGeminiKey;
window.analyzeWithGemini      = App.analyzeWithGemini;
