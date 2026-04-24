/**
 * SGO — state.js
 * Estado global e constantes da aplicação.
 */

export const TEAM_TYPES = {
  "INSTALAÇÃO CIDADE": "Instalação Cidade",
  "TECNICO 12/36H":    "Plantão 12×36",
  "SUPORTE MOTO":      "Suporte Moto",
  "RURAL":             "Rural",
  "FAZ TUDO":          "Faz Tudo"
};

export const DEFAULT_SETTINGS = {
  metasDiarias: {
    "INSTALAÇÃO CIDADE": 5,
    "TECNICO 12/36H":    4,
    "SUPORTE MOTO":      8,
    "RURAL":             3,
    "FAZ TUDO":          5
  }
};

export const DN = ['Dom','Seg','Ter','Qua','Qui','Sex','Sáb'];

export const PAGE_TITLES = {
  dashboard:       'Dashboard',
  analise:         'Análise Operacional',
  pendentes:       'Novos Colaboradores',
  reclassificacao: 'Reclassificar IA',
  config:          'Base de Equipes'
};

// ── Estado mutável ────────────────────────────────────────
export const state = {
  appSettings:          {},
  teamData:             {},
  rawExcelCache:        [],
  globalRawData:        [],
  globalRawDataByMonth: {},
  globalTechStats:      {},
  pendingTechs:         [],
  reclassifySuggestions:[],
  activeMonthYear:      "",
  workerBlobUrl:        null,
  selectedCityTab:      "ALL",
  isCompactMode:        false,
  dailyChartInstance:   null,
  isDark:               false
};
