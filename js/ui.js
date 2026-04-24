/**
 * SGO — ui.js
 * Navegação entre abas, tema escuro, relógio, sheet mobile.
 */

import { state, PAGE_TITLES } from './state.js';
import { saveDarkPref } from './storage.js';

// ── Detectar mobile ────────────────────────────────────────
export function isMobile() {
  return window.innerWidth <= 768;
}

// ── Relógio ────────────────────────────────────────────────
export function initClock() {
  const tick = () => {
    const el = document.getElementById('clockEl');
    if (el) el.textContent = new Date().toLocaleTimeString('pt-BR', {
      hour: '2-digit', minute: '2-digit', second: '2-digit'
    });
  };
  setInterval(tick, 1000);
  tick();
}

// ── Tema escuro ────────────────────────────────────────────
export function applyDarkTheme(isDark) {
  document.documentElement.setAttribute('data-theme', isDark ? 'dark' : 'light');
  ['darkBtn', 'darkBtnMob'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.textContent = isDark ? '☀️' : '🌙';
  });
}

export function toggleDark(onToggle) {
  state.isDark = !state.isDark;
  applyDarkTheme(state.isDark);
  saveDarkPref(state.isDark);
  if (onToggle) onToggle();
}

// ── Troca de aba ───────────────────────────────────────────
export function switchTab(tab) {
  const TABS = ['dashboard', 'analise', 'recorrencia', 'pendentes', 'reclassificacao', 'capacidades', 'moto', 'config'];

  TABS.forEach(t => {
    document.getElementById(`view-${t}`)?.classList.remove('active');
    document.getElementById(`tab-${t}`)?.classList.remove('active');
    document.getElementById(`mob-tab-${t}`)?.classList.remove('active');
  });

  document.getElementById(`view-${tab}`)?.classList.add('active');
  document.getElementById(`tab-${tab}`)?.classList.add('active');
  document.getElementById(`mob-tab-${tab}`)?.classList.add('active');

  const pt = document.getElementById('mobPageTitle');
  if (pt) pt.textContent = PAGE_TITLES[tab] || tab;

  if (isMobile()) window.scrollTo({ top: 0, behavior: 'smooth' });
}

// ── Mobile Filter Sheet ────────────────────────────────────
export function openMobileFilterSheet() {
  const sheet = document.getElementById('mobileFilterSheet');
  sheet.style.display = 'block';
  setTimeout(() => sheet.classList.add('open'), 10);
}

export function closeMobileFilterSheet() {
  const sheet = document.getElementById('mobileFilterSheet');
  sheet.classList.remove('open');
  setTimeout(() => sheet.style.display = 'none', 300);
}

export function initMobileSheetSwipe() {
  let startY = 0;
  const sheet = document.getElementById('mobileFilterSheet');
  sheet.addEventListener('touchstart', e => { startY = e.touches[0].clientY; }, { passive: true });
  sheet.addEventListener('touchmove',  e => {
    if (e.touches[0].clientY - startY > 80) closeMobileFilterSheet();
  }, { passive: true });
}

// ── Modo compacto ──────────────────────────────────────────
export function toggleCompactMode() {
  state.isCompactMode = !state.isCompactMode;
  document.getElementById('matrixWrapper')?.classList.toggle('compact', state.isCompactMode);
  const label = state.isCompactMode ? 'Padrão' : 'Compacto';
  const btn = document.getElementById('btnCompact');
  if (btn) btn.textContent = label;
  const mbl = document.getElementById('mobCompactLbl');
  if (mbl) mbl.textContent = label;
}

// ── Loading overlay ────────────────────────────────────────
export function showLoading(visible) {
  document.getElementById('loadingOverlay').style.display = visible ? 'flex' : 'none';
}

// ── Status bar ─────────────────────────────────────────────
export function setStatus(text) {
  const s = document.getElementById('statusText');
  const ms = document.getElementById('mobStatusText');
  if (s) s.textContent = text;
  if (ms) ms.textContent = text;
}

// ── Sync mobile filter type ────────────────────────────────
export function syncMobileFilters() {
  const ft = document.getElementById('filterType');
  const ftm = document.getElementById('filterTypeMob');
  if (ftm && ft) ft.value = ftm.value;
}
