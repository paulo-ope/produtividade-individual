/**
 * SGO — upload.js
 * Worker de parsing de planilha, pipeline de dados e filtros.
 */

import { state, TEAM_TYPES } from './state.js';
import { limparNome } from './team.js';
import { showLoading, setStatus, isMobile } from './ui.js';
import { updateDashboardStats } from './charts.js';
import { generateMatrix } from './matrix.js';
import { evaluateTechsByAI, renderPendentes, buildOperationalAnalysis } from './analysis.js';
import { populateFilters, renderTeamTable } from './team.js';

// ── Worker inline (XLSX no thread separado) ───────────────
const WORKER_CODE = `
importScripts('https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.min.js');

function cleanStr(s){
  if(!s)return"";
  return String(s).split('-')[0].split('>')[0].split('(')[0].split('/')[0]
    .normalize("NFD").replace(/[\\u0300-\\u036f]/g,"")
    .replace(/\\s+/g,' ').trim().toUpperCase();
}

function extractDate(v){
  if(!v)return null;
  let s=String(v).trim();
  if(!s)return null;
  if(typeof v==='number'||(!isNaN(Number(s)) && !s.includes(':'))){
    let d=new Date(Math.round((Number(s)-25569)*86400*1000));
    d.setMinutes(d.getMinutes()+d.getTimezoneOffset());
    return d;
  }
  let b=s.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?/);
  if(b){
    let y=b[3].length===2?parseInt('20'+b[3]):parseInt(b[3]);
    let H=b[4]?parseInt(b[4],10):0;
    let M=b[5]?parseInt(b[5],10):0;
    let S=b[6]?parseInt(b[6],10):0;
    return new Date(y,parseInt(b[2],10)-1,parseInt(b[1],10),H,M,S);
  }
  let i=s.match(/(\\d{4})-(\\d{1,2})-(\\d{1,2})(?:T|\\s+)(\\d{1,2}):(\\d{1,2})(?::(\\d{1,2}))?/);
  if(i){
    let H=i[4]?parseInt(i[4],10):0;
    let M=i[5]?parseInt(i[5],10):0;
    let S=i[6]?parseInt(i[6],10):0;
    return new Date(parseInt(i[1]),parseInt(i[2])-1,parseInt(i[3]),H,M,S);
  }
  return null;
}

self.onmessage=function(e){
  try{
    const{fileData}=e.data;
    const wb=XLSX.read(fileData,{type:'array',raw:true});
    const rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1,defval:""});
    if(rows.length<=1)throw new Error("Arquivo sem dados");

    const hdrs=rows[0].map(h=>String(h).normalize("NFD").replace(/[\\u0300-\\u036f]/g,"").trim().toLowerCase());
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
    if(iD===-1||iR===-1)throw new Error("Colunas 'Responsável' e/ou 'Data' não encontradas.");

    let pm={},ts={},vr=[];
    const rR=/(RURAL|FAZENDA|S[IÍ]TIO|LINHA |GLEBA|PROJETO|CH[AÁ]CARA|KM \\d)/i; 

    for(let i=1;i<rows.length;i++){
      const dt=rows[i][iD],rp=rows[i][iR];
      if(!dt||!rp||String(rp).toLowerCase().includes("filtros"))continue;
      const d=extractDate(dt);
      if(!d)continue;
      const nm=cleanStr(rp);
      const mk=d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,'0');
      pm[mk]=(pm[mk]||0)+1;
      const localDateStr=d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,'0')+"-"+String(d.getDate()).padStart(2,'0');
      let isR=rR.test(rows[i].join(" "));
      if(!ts[nm])ts[nm]={total:0,rural:0,days:new Set()};
      ts[nm].total++;
      if(isR)ts[nm].rural++;
      ts[nm].days.add(d.getDate());
      const assunto=iA>=0?String(rows[i][iA]||'').trim():'';const diagnostico=iDg>=0?String(rows[i][iDg]||'').trim():'';
      const contrato=iCt>=0?String(rows[i][iCt]||'').trim():'';
      const cliente=iCl>=0?String(rows[i][iCl]||'').trim():'';
      const login=iLg>=0?String(rows[i][iLg]||'').trim():'';
      const status=iSt>=0?String(rows[i][iSt]||'').trim():'';
      const osId=iOs>=0?String(rows[i][iOs]||'').trim():'—'; 
      const rawFilial=iFi>=0?String(rows[i][iFi]||'').trim():'';
      const mapFiliais={"6":"UNI - JI PARANA","7":"UNI - MACHADINHO DOESTE","8":"UNI - ROLIM DE MOURA","9":"UNI - JARU","10":"UNI - OURO PRETO DOESTE","11":"UNI - NOVA BRASILANDIA DOESTE","12":"UNI - PRESIDENTE MEDICI","13":"UNI - SAO FELIPE DOESTE","14":"UNI - ALVORADA DOESTE","15":"UNI - ALTA FLORESTA DOESTE","16":"UNI - SAO MIGUEL DO GUAPORE","17":"UNI - SERINGUEIRAS","18":"UNI - SAO FRANCISCO DO GUAPORE"};
      const filial=mapFiliais[rawFilial]||rawFilial||'NÃO INFORMADA';
      const dtInicio = iIn>=0 ? extractDate(rows[i][iIn]) : null;
      const dtFinal = iFn>=0 ? extractDate(rows[i][iFn]) : null;
      vr.push({nome:nm,nomeOriginal:String(rp||'').trim(),day:d.getDate(),monthStr:mk,dateStr:localDateStr,dateTimeStr:d.toISOString(),assunto,diagnostico,contrato,cliente,login,status,osId,isRural:isR,filial,dtInicio:dtInicio?dtInicio.toISOString():null,dtFinal:dtFinal?dtFinal.toISOString():null}); 
    }

    if(!vr.length)throw new Error("Nenhuma OS válida.");
    const months=Object.keys(pm).sort();
    let am=months[months.length-1];
    let ss={};
    for(let k in ts)ss[k]={total:ts[k].total,rural:ts[k].rural,days:Array.from(ts[k].days)};
    self.postMessage({success:true,activeMonth:am,allOS:vr,techStats:ss,uploadMeta:{hasContrato:iCt>=0,hasCliente:iCl>=0,hasLogin:iLg>=0,hasStatus:iSt>=0,availableMonths:months}});
  }catch(err){
    self.postMessage({success:false,error:err.message});
  }
};
`;

export function initWorker() {
  const blob = new Blob([WORKER_CODE], { type: 'application/javascript' });
  state.workerBlobUrl = URL.createObjectURL(blob);
}

// ── Upload de arquivo ─────────────────────────────────────
export function parseFileWithWorker(event) {
  const file = event.target.files[0];
  if (!file) return;
  showLoading(true);

  const reader = new FileReader();
  reader.onload = ev => {
    const ab = ev.target.result;
    const w  = new Worker(state.workerBlobUrl);

    w.onmessage = msg => {
      showLoading(false);
      ['fileInput', 'fileInputMob'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.value = '';
      });

      if (msg.data.success) {
        applyUploadedData(msg.data);
      } else {
        showToast('Erro ao processar planilha: ' + msg.data.error, 'error');
      }
      w.terminate();
    };

    w.postMessage({ fileData: ab }, [ab]);
  };
  reader.readAsArrayBuffer(file);
}

function applyUploadedData(payload) {
  if (!payload?.allOS?.length) return false;

  // NOVO: Garantir um ID único por linha para impedir falsas colisões em O.S. genéricas
  payload.allOS.forEach((os, idx) => {
    if (!os._uid) os._uid = 'os_' + idx + '_' + Math.random().toString(36).substr(2, 5);
  });

  state.activeMonthYear = payload.activeMonth;
  state.rawExcelCache = payload.allOS;
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
  ['filterMonth', 'filterMonthMob'].forEach(id => {
    const el = document.getElementById(id);
    if (el) {
      el.value = state.activeMonthYear;
      if (availableMonths.length) {
        el.min = availableMonths[0];
        el.max = availableMonths[availableMonths.length - 1];
      }
    }
  });
  document.getElementById('cityTabsWrapper').style.display = 'block';
  document.getElementById('searchBar').style.display = 'flex';
  const ma = document.getElementById('mobActions');
  if (ma) ma.style.display = 'flex';
  setStatus('Dados · ' + state.activeMonthYear);
  syncEcosystem();
  return true;
}

export function changeActiveMonth(monthValue) {
  if (!monthValue || !state.rawExcelCache.length) return;
  
  showLoading(true);
  setTimeout(() => {
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

// ── Rebuild dados brutos ───────────────────────────────────
export function rebuildGlobalRawData() {
  if (!state.rawExcelCache.length) return;
  state.globalRawData = [];
  const cache = {};

  state.rawExcelCache.forEach(os => {
    let tk = cache[os.nome];
    if (tk === undefined) {
      tk = null;
      for (const k in state.teamData) {
        if (os.nome === k || os.nome.includes(k) || k.includes(os.nome)) { tk = k; break; }
      }
      cache[os.nome] = tk;
    }

    const osData = {
        _uid: os._uid,
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

    if (tk) {
      state.globalRawData.push({
        ...osData,
        techKey:  tk,
        cidade:   state.teamData[tk].base,
        tipo:     state.teamData[tk].tipo || 'INSTALAÇÃO CIDADE'
      });
    } else {
      state.globalRawData.push({
        ...osData,
        techKey: os.nome,
        cidade: 'PENDENTE',
        tipo: 'INSTALAÇÃO CIDADE'
      });
    }
  });

  applyFilters();
}

// ── Aplicar filtros ────────────────────────────────────────
export function applyFilters() {
  if (!state.globalRawData.length) {
    const wrapper = document.getElementById('matrixWrapper');
    if (wrapper) wrapper.innerHTML = `<div style="padding:40px;text-align:center;font-family:var(--mono);font-size:11px;color:var(--text-3);background:var(--card-bg);border:1px solid var(--card-border);border-radius:var(--r-xl);">Faça o upload da planilha para gerar a matriz.</div>`;
    return;
  }

  // Sincronizar filtro mobile
  const ft  = document.getElementById('filterType')?.value || 'ALL';
  const ftm = document.getElementById('filterTypeMob');
  if (ftm) ftm.value = ft;

  const fc = state.selectedCityTab;
  const filtered = state.globalRawData.filter(i =>
    i.monthStr === state.activeMonthYear &&
    (fc === 'ALL' || i.cidade === fc) &&
    (ft === 'ALL' || i.tipo === ft)
  );

  updateDashboardStats(filtered);
  buildOperationalAnalysis(filtered);
  generateMatrix(filtered);

  document.getElementById('legendBar').style.display = 'flex';
  document.getElementById('matrixSection').style.display = 'block';
  document.getElementById('btnGeminiAnalyze')?.classList.remove('hidden');
}

// ── Sync completo do ecossistema ───────────────────────────
export function syncEcosystem() {
  populateFilters();
  renderTeamTable();
  evaluateTechsByAI();
  renderPendentes();
  rebuildGlobalRawData();
}

// ── Busca na matriz ────────────────────────────────────────
export function filterMatrixBySearch() {
  const q    = document.getElementById('techSearchInput')?.value.trim().toLowerCase() || '';
  const rows = document.querySelectorAll('#matrixWrapper tbody tr');
  let count  = 0;

  rows.forEach(r => {
    const name = r.querySelector('.cn-name');
    if (!name) { r.style.display = ''; return; }
    const match = !q || name.textContent.toLowerCase().includes(q);
    r.style.display = match ? '' : 'none';
    if (match) count++;
  });

  const sc = document.getElementById('searchCount');
  if (sc) sc.textContent = q ? `${count} resultado${count !== 1 ? 's' : ''}` : '';
}

export function clearSearch() {
  const input = document.getElementById('techSearchInput');
  if (input) input.value = '';
  filterMatrixBySearch();
}
