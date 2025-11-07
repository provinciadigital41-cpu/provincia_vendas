'use strict';

/**
 * server.js — Provincia Vendas (Pipefy + D4Sign via secure.d4sign.com.br)
 * Node 18+ (fetch global)
 */

const express = require('express');
const crypto = require('crypto');

// compression opcional
let compression = null;
try { compression = require('compression'); } catch {}

let undiciAgent = null;
try {
  const { Agent, setGlobalDispatcher } = require('undici');
  undiciAgent = new Agent({
    keepAliveTimeout: 30_000,
    keepAliveMaxTimeout: 60_000,
    connections: 16,
    pipelining: 1
  });
  setGlobalDispatcher(undiciAgent);
} catch {}

const app = express();
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
if (compression) app.use(compression({ threshold: 1024 }));
app.set('trust proxy', true);

app.use((req, _res, next) => {
  console.log(`[${new Date().toISOString()}] ${req.method} ${req.originalUrl} ua="${req.get('user-agent')}" ip=${req.ip}`);
  next();
});

/* =========================
 * ENV
 * =======================*/
let {
  PORT,
  PUBLIC_BASE_URL,
  PUBLIC_LINK_SECRET,

  PIPE_API_KEY,
  PIPE_GRAPHQL_ENDPOINT,

  FIELD_ID_CONNECT_MARCA_NOME,
  FIELD_ID_CONNECT_CLASSE,
  FIELD_ID_CONNECT_CLASSES,
  MARCAS_TABLE_ID,
  CONTACTS_TABLE_ID,
  MARCAS2_TABLE_ID,
  CLASSES_TABLE_ID,

  PIPEFY_FIELD_LINK_CONTRATO,
  NOVO_PIPE_ID,
  FASE_VISITA_ID,
  PHASE_ID_CONTRATO_ENVIADO,

  D4SIGN_TOKEN,
  D4SIGN_CRYPT_KEY,
  TEMPLATE_UUID_CONTRATO,

  EMAIL_ASSINATURA_EMPRESA,

  COFRE_UUID_EDNA,
  COFRE_UUID_GREYCE,
  COFRE_UUID_MARIANA,
  COFRE_UUID_VALDEIR,
  COFRE_UUID_DEBORA,
  COFRE_UUID_MAYKON,
  COFRE_UUID_JEFERSON,
  COFRE_UUID_RONALDO,
  COFRE_UUID_BRENDA,
  COFRE_UUID_MAURO,

  DEFAULT_COFRE_UUID
} = process.env;

PORT = PORT || 3000;
PIPE_GRAPHQL_ENDPOINT = PIPE_GRAPHQL_ENDPOINT || 'https://api.pipefy.com/graphql';
PIPEFY_FIELD_LINK_CONTRATO = PIPEFY_FIELD_LINK_CONTRATO || 'd4_contrato';
FIELD_ID_CONNECT_MARCA_NOME = FIELD_ID_CONNECT_MARCA_NOME || 'marcas_1';
FIELD_ID_CONNECT_CLASSE     = FIELD_ID_CONNECT_CLASSE     || 'marcas_2';
CLASSES_TABLE_ID            = CLASSES_TABLE_ID || '306521337';

if (!PUBLIC_BASE_URL || !PUBLIC_LINK_SECRET) console.warn('[AVISO] Configure PUBLIC_BASE_URL e PUBLIC_LINK_SECRET');
if (!PIPE_API_KEY) console.warn('[AVISO] PIPE_API_KEY ausente');
if (!D4SIGN_TOKEN || !D4SIGN_CRYPT_KEY) console.warn('[AVISO] D4SIGN_TOKEN / D4SIGN_CRYPT_KEY ausentes');

const COFRES_UUIDS = {
  'EDNA BERTO DA SILVA': COFRE_UUID_EDNA,
  'Greyce Maria Candido Souza': COFRE_UUID_GREYCE,
  'mariana cristina de oliveira': COFRE_UUID_MARIANA,
  'Valdeir Almedia': COFRE_UUID_VALDEIR,
  'Débora Gonçalves': COFRE_UUID_DEBORA,
  'Maykon Campos': COFRE_UUID_MAYKON,
  'Jeferson Andrade Siqueira': COFRE_UUID_JEFERSON,
  'RONALDO SCARIOT DA SILVA': COFRE_UUID_RONALDO,
  'BRENDA ROSA DA SILVA': COFRE_UUID_BRENDA,
  'Mauro Furlan Neto': COFRE_UUID_MAURO
};

/* =========================
 * Cache leve e índices
 * =======================*/
const cache = new Map();
function cacheGet(k){ const hit = cache.get(k); if (!hit) return null; if (hit.exp < Date.now()){ cache.delete(k); return null; } return hit.val; }
function cacheSet(k, val, ttlMs=5*60*1000){ cache.set(k, { val, exp: Date.now()+ttlMs }); }
const titleIndex = new Map();
function idxKey(tableId){ return `idx:${tableId}`; }
function titleNorm(s){ return String(s||'').trim().toLowerCase(); }
function idxSet(tableId, title, id){
  if (!tableId || !title || !id) return;
  const k = idxKey(tableId);
  const m = titleIndex.get(k) || new Map();
  m.set(titleNorm(title), String(id));
  titleIndex.set(k, m);
}
function idxGet(tableId, title){
  if (!tableId || !title) return null;
  const m = titleIndex.get(idxKey(tableId));
  return m ? m.get(titleNorm(title)) : null;
}

/* =========================
 * Helpers
 * =======================*/
function onlyDigits(s){ return String(s||'').replace(/\D/g,''); }
function normalizePhone(s){ return onlyDigits(s); }
function toBRL(n){ return isNaN(n)?'':Number(n).toLocaleString('pt-BR',{style:'currency',currency:'BRL'}); }
function parseNumberBR(v){
  if (v==null) return NaN;
  const s = String(v).trim();
  if (!s) return NaN;
  if (/^\d{1,3}(\.\d{3})*,\d{2}$/.test(s)) return Number(s.replace(/\./g,'').replace(',','.'));
  if (/^\d+(\.\d+)?$/.test(s)) return Number(s);
  return Number(s.replace(/[^\d.,-]/g,'').replace(/\./g,'').replace(',','.'));
}
function onlyNumberBR(s){ const n = parseNumberBR(s); return isNaN(n)? 0 : n; }
const MESES_PT = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
function monthNamePt(mIndex1to12) { return MESES_PT[(Math.max(1, Math.min(12, Number(mIndex1to12))) - 1)]; }
function parsePipeDateToDate(value){
  if (!value) return null;
  const s = String(value).trim();
  const m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (m){
    const dd = Number(m[1]), mm = Number(m[2]), yyyy = Number(m[3]);
    const d = new Date(yyyy, mm-1, dd);
    return isNaN(d) ? null : d;
  }
  const d = new Date(s);
  return isNaN(d) ? null : d;
}
function fmtDMY2(value){
  const d = value instanceof Date ? value : parsePipeDateToDate(value);
  if (!d) return '';
  const dd = String(d.getDate()).padStart(2,'0');
  const mm = String(d.getMonth()+1).padStart(2,'0');
  const yy = String(d.getFullYear()).slice(-2);
  return `${dd}/${mm}/${yy}`;
}
async function fetchWithRetry(url, init={}, opts={}){
  const attempts = opts.attempts ?? 2;
  const baseDelayMs = opts.baseDelayMs ?? 300;
  const timeoutMs = opts.timeoutMs ?? 10_000;

  for (let i=0;i<attempts;i++){
    try{
      const ctrl = new AbortController();
      const t = setTimeout(()=>ctrl.abort(), timeoutMs);
      const res = await fetch(url, { ...init, signal: ctrl.signal });
      clearTimeout(t);
      if (!res.ok && i < attempts-1) {
        await new Promise(r => setTimeout(r, baseDelayMs * (i+1)));
        continue;
      }
      return res;
    } catch(e){
      if (i === attempts-1) throw e;
      await new Promise(r => setTimeout(r, baseDelayMs * (i+1)));
    }
  }
  throw new Error('fetchWithRetry: esgotou tentativas');
}

/* =========================
 * Token público
 * =======================*/
function makeLeadToken(payload){
  const body = Buffer.from(JSON.stringify(payload)).toString('base64url');
  const sig = crypto.createHmac('sha256', PUBLIC_LINK_SECRET).update(body).digest('base64url');
  return `${body}.${sig}`;
}
function parseLeadToken(token){
  const [body, sig] = String(token||'').split('.');
  if (!body || !sig) throw new Error('token inválido');
  const expected = crypto.createHmac('sha256', PUBLIC_LINK_SECRET).update(body).digest('base64url');
  if (sig !== expected) throw new Error('assinatura inválida');
  const json = JSON.parse(Buffer.from(body,'base64url').toString('utf8'));
  if (!json.cardId) throw new Error('payload inválido');
  return json;
}

/* =========================
 * Pipefy GraphQL
 * =======================*/
async function gql(query, variables){
  const r = await fetchWithRetry(PIPE_GRAPHQL_ENDPOINT, {
    method: 'POST',
    headers: { 'Authorization': `Bearer ${PIPE_API_KEY}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ query, variables })
  }, { attempts: 2, baseDelayMs: 300, timeoutMs: 8000 });
  const j = await r.json();
  if (!r.ok || j.errors) throw new Error(`Pipefy GQL: ${r.status} ${JSON.stringify(j.errors||{})}`);
  return j.data;
}
async function getCard(cardId){
  const ck = `card:${cardId}`;
  const hit = cacheGet(ck);
  if (hit) return hit;
  const data = await gql(`query($id: ID!){
    card(id:$id){
      id title
      current_phase{ id name }
      pipe{ id name }
      fields{ name value field{ id type } }
      assignees{ name email }
    }
  }`, { id: cardId });
  const card = data.card;
  cacheSet(ck, card, 60_000);
  return card;
}
async function updateCardField(cardId, fieldId, newValue){
  await gql(`mutation($input: UpdateCardFieldInput!){
    updateCardField(input:$input){ card{ id } }
  }`, { input: { card_id: Number(cardId), field_id: fieldId, new_value: newValue } });
}
async function getTableRecord(recordId){
  const ck = `rec:${recordId}`;
  const hit = cacheGet(ck); if (hit) return hit;
  const data = await gql(`query($id: ID!){
    table_record(id:$id){
      id title
      record_fields{ name value field{ id type } }
    }
  }`, { id: recordId });
  const rec = data.table_record;
  cacheSet(ck, rec);
  return rec;
}
async function getCardShallow(cardId){
  const ck = `cardsh:${cardId}`;
  const hit = cacheGet(ck); if (hit) return hit;
  const data = await gql(`query($id: ID!){
    card(id:$id){
      id title
      fields{ name value field{ id type } }
    }
  }`, { id: cardId });
  const c = data.card;
  cacheSet(ck, c);
  return c;
}

/**
 * Abre um ID que pode ser table_record ou card.
 * Retorna { kind: 'record'|'card', node }
 */
async function getAnyNode(id){
  try {
    const rec = await getTableRecord(id);
    if (rec && rec.id) return { kind: 'record', node: rec };
  } catch {}
  try {
    const c = await getCardShallow(id);
    if (c && c.id) return { kind: 'card', node: c };
  } catch {}
  return { kind: null, node: null };
}

async function listTableRecords(tableId, first=200, after=null){
  const data = await gql(`query($tableId: ID!, $first: Int!, $after: String){
    table(id:$tableId){
      id
      table_records(first:$first, after:$after){
        edges{ node{ id title record_fields{ name value field{ id type } } } }
        pageInfo{ hasNextPage endCursor }
      }
    }
  }`, { tableId, first, after });
  const edges = data?.table?.table_records?.edges || [];
  for (const e of edges){ idxSet(tableId, e.node.title, e.node.id); }
  const pageInfo = data?.table?.table_records?.pageInfo || {};
  return { records: edges.map(e=>e.node), pageInfo };
}

/* =========================
 * Parsing de campos do card
 * =======================*/
function toById(card){
  const by={}; for (const f of card?.fields||[]) if (f?.field?.id) by[f.field.id]=f.value;
  return by;
}
function getByName(card, nameSub){
  const t = String(nameSub).toLowerCase();
  const f = (card.fields||[]).find(ff=> String(ff?.name||'').toLowerCase().includes(t));
  return f?.value || '';
}
function getFirstByNames(card, arr){
  for (const k of arr){ const v = getByName(card, k); if (v) return v; }
  return '';
}
function parseMaybeJsonArray(v){
  try { return Array.isArray(v)? v : JSON.parse(v); }
  catch { return v? [String(v)] : []; }
}

/* =========================
 * Conexões: extrair IDs
 * =======================*/
function extractConnectedRecordIds(value){
  const arr = parseMaybeJsonArray(value);
  const out = [];
  for (const item of arr){
    if (!item) continue;
    if (typeof item === 'object'){
      const cand = item.id || item.record_id || item.recordId || item.value || (item.node && item.node.id);
      if (cand && /^\d+$/.test(String(cand))) { out.push(String(cand)); continue; }
    }
    const s = String(item);
    if (/^\d+$/.test(s)) { out.push(s); continue; }
    const m = s.match(/(^|[^\d])(\d{3,})([^\d]|$)/);
    if (m) out.push(String(m[2]));
  }
  return [...new Set(out)];
}
function checklistToText(v) {
  const arr = parseMaybeJsonArray(v);
  return Array.isArray(arr) ? arr.join(', ') : String(v || '');
}

/* =========================
 * Normalização para ler campos tanto de record quanto de card
 * =======================*/
function getRecordFieldsLike(anyNode){
  // record: node.record_fields; card: node.fields
  if (!anyNode) return [];
  if (anyNode.record_fields) return anyNode.record_fields;
  if (anyNode.fields) {
    // já tem {name, value, field:{id,type}}
    return anyNode.fields.map(f => ({ name: f.name, value: f.value, field: { id: f.field?.id, type: f.field?.type }}));
  }
  return [];
}

/* =========================
 * Contato — extração
 * =======================*/
function extractNameEmailPhoneFromRecord(record){
  let nome='', email='', telefone='';

  const isNomeContato = (f)=>{
    const tId = (f?.field?.id||'').toLowerCase();
    const lbl = String(f?.name||'').toLowerCase();
    const type = (f?.field?.type||'').toLowerCase();
    if (lbl.includes('nome da marca') || tId.includes('nome_da_marca') || lbl.includes('marca')) return false;
    if (tId === 'nome_do_contato') return true;
    if ((lbl.includes('nome') && lbl.includes('contat')) || lbl === 'nome do contato') return true;
    if (type === 'short_text' && lbl.includes('contat') && lbl.includes('nome')) return true;
    return false;
  };

  for (const f of record?.record_fields||[]){
    const t = (f?.field?.type||'').toLowerCase();
    const id = (f?.field?.id||'').toLowerCase();
    const label = (f?.name||'').toLowerCase();
    if (!nome && isNomeContato(f)) nome = String(f.value||'');
    if (!email && (t==='email' || id==='email_do_contato' || label.includes('email'))) email = String(f.value||'');
    if (!telefone && (t==='phone' || id==='telefone_do_contato' || label.includes('telefone') || label.includes('whats') || label.includes('celular'))) telefone = String(f.value||'');
  }
  return { nome, email, telefone };
}
function extractNameEmailPhoneFromCard(card){
  let nome='', email='', telefone='';
  for (const f of (card?.fields||[])){
    const t = (f?.field?.type||'').toLowerCase();
    const id = (f?.field?.id||'').toLowerCase();
    const label = String(f?.name||'').toLowerCase();

    if (!nome){
      if (id==='nome_do_contato' || label==='nome do contato' || (label.includes('nome') && label.includes('contat'))) {
        nome = String(f.value||'');
      }
    }
    if (!email && (t==='email' || id==='email_do_contato' || label.includes('email'))) {
      email = String(f.value||'');
    }
    if (!telefone && (t==='phone' || id==='telefone_do_contato' || label.includes('telefone') || label.includes('whats') || label.includes('celular'))) {
      telefone = String(f.value||'');
    }
  }
  return { nome, email, telefone };
}

/* =========================
 * Marca — resolver a partir do card principal
 * (aceita tanto table_record quanto card)
 * =======================*/
async function resolveMarcaRecordFromCard(card){
  const by = toById(card);
  const v = by[FIELD_ID_CONNECT_MARCA_NOME];
  const ids = extractConnectedRecordIds(v);
  if (ids.length){
    // pega o primeiro
    return await getAnyNode(ids[0]); // {kind,node}
  }

  // Caso seja texto (título) — tenta achar em tabela de marcas
  const first = parseMaybeJsonArray(v)[0];
  if (!first) return null;

  if (/^\d+$/.test(String(first))){
    const any = await getAnyNode(String(first));
    return any?.kind ? any : null;
  }

  if (!MARCAS_TABLE_ID) return null;

  const idByIdx = idxGet(MARCAS_TABLE_ID, first);
  if (idByIdx) { try { const rec = await getTableRecord(idByIdx); return { kind:'record', node: rec }; } catch {} }

  let after=null;
  for (let i=0;i<20;i++){
    const {records, pageInfo} = await listTableRecords(MARCAS_TABLE_ID, 200, after);
    const hit = records.find(r=> titleNorm(r.title) === titleNorm(first));
    if (hit) return { kind:'record', node: hit };
    if (!pageInfo?.hasNextPage) break;
    after = pageInfo.endCursor || null;
  }
  return null;
}

/* =========================
 * Contato vindo da Marca (record ou card)
 * =======================*/
async function resolveContatoFromMarcaRecord(marcaAny){
  if (!marcaAny || !marcaAny.node) return { nome:'', email:'', telefone:'' };
  const node = marcaAny.node;

  // Procura dentro da marca um conector que leve a contato
  const fieldsLike = getRecordFieldsLike(node);
  const connField = fieldsLike.find(f=>{
    const t = String(f?.field?.type||'').toLowerCase();
    const lbl = String(f?.name||'').toLowerCase();
    const id = String(f?.field?.id||'').toLowerCase();
    return (t==='connector' || t==='table_connection' || lbl.includes('contato') || id.includes('contato'));
  });

  if (connField && connField.value) {
    const recIds = extractConnectedRecordIds(connField.value);
    for (const id of recIds) {
      try {
        const any = await getAnyNode(id);
        if (any.kind === 'record') {
          const ex = extractNameEmailPhoneFromRecord(any.node);
          if (ex.email || ex.telefone || ex.nome) return ex;
        } else if (any.kind === 'card') {
          const ex = extractNameEmailPhoneFromCard(any.node);
          if (ex.email || ex.telefone || ex.nome) return ex;
        }
      } catch {}
    }
  }

  // fallback: extrai direto do próprio nó
  if (marcaAny.kind === 'record') return extractNameEmailPhoneFromRecord(node);
  if (marcaAny.kind === 'card')   return extractNameEmailPhoneFromCard(node);
  return { nome:'', email:'', telefone:'' };
}

/* =========================
 * Classe
 * =======================*/
async function resolveClasseFromLabelOnCard(card){
  const f = (card.fields||[]).find(ff=>{
    const isConn = (ff?.field?.type==='connector' || ff?.field?.type==='table_connection');
    const idOk = String(ff?.field?.id||'').toLowerCase()==='classes_inpi';
    const lblOk = String(ff?.name||'').toLowerCase().includes('classes inpi');
    return isConn && (idOk||lblOk);
  });
  if (!f||!f.value) return '';
  const recIds = extractConnectedRecordIds(f.value);
  if (recIds.length){
    try { const rec = await getTableRecord(recIds[0]); return rec?.title || ''; } catch {}
  }
  const first = parseMaybeJsonArray(f.value)[0];
  return String(first||'');
}
function normalizeClasseToNumbersOnly(classeStr){
  if (!classeStr) return '';
  const nums = String(classeStr).match(/\d+/g) || [];
  return nums.join(', ');
}
async function resolveClasseFromCard(card, marcaAny){
  const fromCard = await resolveClasseFromLabelOnCard(card);
  if (fromCard) return normalizeClasseToNumbersOnly(fromCard);

  const by = toById(card);

  if (FIELD_ID_CONNECT_CLASSES){
    const v = by[FIELD_ID_CONNECT_CLASSES];
    const recIds = extractConnectedRecordIds(v);
    if (recIds.length){
      try { const rec = await getTableRecord(recIds[0]); return normalizeClasseToNumbersOnly(rec?.title||''); } catch{}
    } else {
      const first = parseMaybeJsonArray(v)[0];
      if (first) return normalizeClasseToNumbersOnly(String(first));
    }
  }

  const v2 = by[FIELD_ID_CONNECT_CLASSE];
  const recIds2 = extractConnectedRecordIds(v2);
  if (recIds2.length){
    try{
      const rec = await getTableRecord(recIds2[0]);
      const classeField = (rec.record_fields||[]).find(f => String(f?.name||'').toLowerCase().includes('classe'));
      if (classeField?.value){
        const innerIds = extractConnectedRecordIds(classeField.value);
        if (innerIds.length){ const recClasse = await getTableRecord(innerIds[0]); return normalizeClasseToNumbersOnly(recClasse?.title || ''); }
        const val = classeField.value;
        if (Array.isArray(val)) return normalizeClasseToNumbersOnly(String(val[0]||''));
        try { const a = JSON.parse(val); if (Array.isArray(a)&&a.length) return normalizeClasseToNumbersOnly(String(a[0]||'')); } catch { if (val) return normalizeClasseToNumbersOnly(String(val)); }
      }
      if (rec?.title) return normalizeClasseToNumbersOnly(String(rec.title));
    } catch {}
  } else {
    const first2 = parseMaybeJsonArray(v2)[0];
    if (first2) return normalizeClasseToNumbersOnly(String(first2));
  }

  if (marcaAny && marcaAny.node){
    const fieldsLike = getRecordFieldsLike(marcaAny.node);
    const classeField = fieldsLike.find(f=> String(f?.name||'').toLowerCase().includes('classe'));
    if (classeField?.value) return normalizeClasseToNumbersOnly(String(classeField.value));
  }
  return '';
}

/* =========================
 * Documento
 * =======================*/
function pickDocumento(card){
  const prefer = ['cpf','cnpj','documento','doc','cpf/cnpj','cnpj/cpf'];
  for (const k of prefer){
    const v = getFirstByNames(card, [k]);
    const d = onlyDigits(v);
    if (d.length===11) return { tipo:'CPF', valor:v };
    if (d.length===14) return { tipo:'CNPJ', valor:v };
  }
  const by = toById(card);
  const cnpjStart = by['cnpj'] || getFirstByNames(card, ['cnpj']);
  if (cnpjStart) return { tipo:'CNPJ', valor:cnpjStart };
  return { tipo:'', valor:'' };
}

/* =========================
 * Assignee → cofre
 * =======================*/
function stripDiacritics(s){ return String(s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,''); }
function normalizeName(s){ return stripDiacritics(String(s||'').trim()).toLowerCase(); }
function extractAssigneeNames(raw){
  const out=[]; const push=v=>{ if(v) out.push(String(v)); }; const tryParse=v=>{ if(typeof v==='string'){ try{return JSON.parse(v);} catch{return v;} } return v; };
  const val = tryParse(raw);
  if (Array.isArray(val)){ for (const it of val) push(typeof it==='string'? it : (it?.name||it?.username||it?.email||it?.value)); }
  else if (typeof val==='object'&&val){ push(val.name||val.username||val.email||val.value); }
  else if (typeof val==='string'){ const m = val.match(/^\s*\[.*\]\s*$/)? tryParse(val) : null; if (m && Array.isArray(m)) m.forEach(x=>push(typeof x==='string'? x : (x?.name||x?.email))); else push(val); }
  return [...new Set(out.filter(Boolean))];
}
function resolveCofreUuidByCard(card){
  const by = toById(card);
  const candidatosBrutos = [];
  if (by['vendedor_respons_vel']) candidatosBrutos.push(by['vendedor_respons_vel']);
  if (by['vendedor_respons_vel_1']) candidatosBrutos.push(by['vendedor_respons_vel_1']);
  if (by['respons_vel_5']) candidatosBrutos.push(by['respons_vel_5']);
  if (by['representante']) candidatosBrutos.push(by['representante']);
  const nomesOuEmails = candidatosBrutos.flatMap(extractAssigneeNames);
  const normKeys = Object.keys(COFRES_UUIDS||{}).reduce((acc,k)=>{ acc[normalizeName(k)] = COFRES_UUIDS[k]; return acc; },{});
  for (const s of nomesOuEmails){ const n=normalizeName(s); if (normKeys[n]) return normKeys[n]; }
  if (DEFAULT_COFRE_UUID) return DEFAULT_COFRE_UUID;
  return null;
}

/* =========================
 * Taxa
 * =======================*/
function computeValorTaxaBRLFromFaixa(d){
  let valorTaxaSemRS = '';
  const taxa = String(d.taxa_faixa||'');
  if (taxa.includes('440')) valorTaxaSemRS = '440,00';
  else if (taxa.includes('880')) valorTaxaSemRS = '880,00';
  return valorTaxaSemRS ? `R$ ${valorTaxaSemRS}` : '';
}

/* =========================
 * Montagem de dados
 * =======================*/
function pickParcelas(card){
  const by = toById(card);
  let raw = by['sele_o_de_lista'] || by['quantidade_de_parcelas'] || by['numero_de_parcelas'] || '';
  if (!raw) raw = getFirstByNames(card, ['parcelas','quantidade de parcelas','nº parcelas']);
  const m = String(raw||'').match(/(\d+)/);
  return m ? m[1] : '1';
}
function pickValorAssessoria(card){
  const by = toById(card);
  let raw = by['valor_da_assessoria'] || by['valor_assessoria'] || '';
  if (!raw) raw = getFirstByNames(card, ['valor da assessoria','assessoria']);
  if (!raw){
    const hit = (card.fields||[]).find(f => String(f?.field?.type||'').toLowerCase()==='currency'); raw = hit?.value||'';
  }
  const n = parseNumberBR(raw);
  return isNaN(n)? null : n;
}

async function resolveContatoFromCardConnections(card){
  const out = { nome:'', email:'', telefone:'' };

  const connectors = (card.fields||[]).filter(f => {
    const type = String(f?.field?.type||'').toLowerCase();
    return type==='connector' || type==='table_connection';
  });

  const primary = connectors.filter(f => String(f?.field?.id||'').toLowerCase()==='contato');
  const secondary = connectors.filter(f => {
    const idLc = String(f?.field?.id||'').toLowerCase();
    const nameLc = String(f?.name||'').toLowerCase();
    return idLc!=='contato' && (['contato','cliente','respons','contratante'].some(k => idLc.includes(k) || nameLc.includes(k)));
  });

  const ordered = [...primary, ...secondary];

  for (const f of ordered) {
    const recIds = extractConnectedRecordIds(f.value);
    for (const id of recIds) {
      try {
        const any = await getAnyNode(id);
        if (any.kind === 'record') {
          const { nome, email, telefone } = extractNameEmailPhoneFromRecord(any.node);
          if (nome && !out.nome) out.nome = nome;
          if (email && !out.email) out.email = email;
          if (telefone && !out.telefone) out.telefone = telefone;
        } else if (any.kind === 'card') {
          const { nome, email, telefone } = extractNameEmailPhoneFromCard(any.node);
          if (nome && !out.nome) out.nome = nome;
          if (email && !out.email) out.email = email;
          if (telefone && !out.telefone) out.telefone = telefone;
        }
        if (out.email && out.telefone && out.nome) return out;
      } catch {}
    }
  }
  return out;
}

async function montarDados(card){
  const by = toById(card);

  const [marcaAny, contatoDireto] = await Promise.all([
    resolveMarcaRecordFromCard(card),     // << agora existe
    resolveContatoFromCardConnections(card)
  ]);

  const classePromise = resolveClasseFromCard(card, marcaAny);
  const contatoFromMarcaPromise = resolveContatoFromMarcaRecord(marcaAny);

  let tipoMarca = checklistToText(
    by['tipo_de_marca'] ||
    by['checklist_vertical'] ||
    getFirstByNames(card, ['tipo de marca'])
  );

  if (!tipoMarca) {
    const v2 = by[FIELD_ID_CONNECT_CLASSE];
    const first2 = parseMaybeJsonArray(v2)[0];
    if (first2) {
      try {
        let rec = null;
        if (/^\d+$/.test(String(first2))) {
          rec = await getTableRecord(String(first2));
        } else if (MARCAS2_TABLE_ID) {
          const idIdx = idxGet(MARCAS2_TABLE_ID, String(first2));
          if (idIdx) rec = await getTableRecord(idIdx);
          else {
            let after = null;
            for (let i = 0; i < 10 && !rec; i++) {
              const { records, pageInfo } = await listTableRecords(MARCAS2_TABLE_ID, 200, after);
              rec = records.find(r => titleNorm(r.title)===titleNorm(String(first2)));
              if (!pageInfo?.hasNextPage) break;
              after = pageInfo.endCursor || null;
            }
          }
        }
        if (rec) {
          const fTipo = (rec.record_fields || []).find(f =>
            String(f?.field?.id || '').toLowerCase() === 'tipo_de_marca' ||
            String(f?.name || '').toLowerCase().includes('tipo de marca')
          );
          if (fTipo) tipoMarca = checklistToText(fTipo.value);
        }
      } catch {}
    }
  }

  const [classe, contatoFromMarca] = await Promise.all([classePromise, contatoFromMarcaPromise]);

  const contatoNome     = contatoDireto.nome     || contatoFromMarca.nome     || getFirstByNames(card, ['nome do contato','contratante','responsável legal','responsavel legal']);
  const contatoEmail    = contatoDireto.email    || contatoFromMarca.email    || getFirstByNames(card, ['email','e-mail']);
  const contatoTelefone = contatoDireto.telefone || contatoFromMarca.telefone || getFirstByNames(card, ['telefone','celular','whatsapp','whats']);

  const doc = pickDocumento(card);
  const cpfDoc  = doc.tipo==='CPF'?  doc.valor : '';
  const cnpjDoc = doc.tipo==='CNPJ'? doc.valor : '';
  const cpfCampo  = by['cpf']    || '';
  const cnpjCampo = by['cnpj_1'] || '';

  const nParcelas = pickParcelas(card);
  const valorAssessoria = pickValorAssessoria(card);
  const formaAss = by['copy_of_tipo_de_pagamento'] || getFirstByNames(card, ['tipo de pagamento assessoria']) || '';

  const serv1 = getFirstByNames(card, ['serviços de contratos','serviços contratados','serviços']);
  const temMarca = Boolean(getFirstByNames(card, ['marca']) || by['marca'] || by['marcas_1'] || card.title);
  const qtdMarca = temMarca ? '1' : '';
  const servicos = [serv1].filter(Boolean);

  const taxaFaixaRaw = by['taxa'] || getFirstByNames(card, ['taxa']);
  const valorTaxaBRL = computeValorTaxaBRLFromFaixa({ taxa_faixa: taxaFaixaRaw });
  const formaPagtoTaxa = by['tipo_de_pagamento'] || '';
  const dataPagtoTaxa = fmtDMY2(by['data_de_pagamento_taxa'] || '');

  const cepCnpj    = by['cep_do_cnpj']     || '';
  const ruaCnpj    = by['rua_av_do_cnpj']  || '';
  const bairroCnpj = by['bairro_do_cnpj']  || '';
  const cidadeCnpj = by['cidade_do_cnpj']  || '';
  const ufCnpj     = by['estado_do_cnpj']  || '';
  const numeroCnpj = by['n_mero_1']        || getFirstByNames(card, ['numero','número','nº']) || '';

  const vendedor = extractAssigneeNames(by['vendedor_respons_vel'] || by['vendedor_respons_vel_1'] || by['respons_vel_5'])[0] || '';

  const riscoMarca = by['risco_da_marca'] || '';
  const nacionalidade = by['nacionalidade'] || '';
  const selecaoCnpjOuCpf = by['cnpj_ou_cpf'] || '';
  const estadoCivil = by['estado_civ_l'] || '';
  const dataPagtoAssessoria = fmtDMY2(by['data_de_pagamento_assessoria'] || '');

  // Campo de texto longo para o template ${"marcas-espec"}
  const marcasEspec = by['classe'] || getByName(card, 'classes e especificações') || '';

  return {
    cardId: card.id,
    titulo: card.title,

    nome: contatoNome || (by['r_social_ou_n_completo']||''),
    cpf: cpfDoc, 
    cnpj: cnpjDoc,
    rg: by['rg'] || '',
    estado_civil: estadoCivil,

    cpf_campo: cpfCampo,
    cnpj_campo: cnpjCampo,

    email: contatoEmail || '',
    telefone: contatoTelefone || '',

    classe: classe || '',
    marcas_espec: marcasEspec || '',
    qtd_marca: qtdMarca,
    tipo_marca: tipoMarca || '',

    servicos,
    parcelas: nParcelas,
    valor_total: valorAssessoria ? toBRL(valorAssessoria) : '',
    forma_pagto_assessoria: formaAss,
    data_pagto_assessoria: dataPagtoAssessoria,

    taxa_faixa: taxaFaixaRaw || '',
    valor_taxa_brl: valorTaxaBRL,
    forma_pagto_taxa: formaPagtoTaxa,
    data_pagto_taxa: dataPagtoTaxa,

    cep_cnpj: cepCnpj,
    rua_cnpj: ruaCnpj,
    bairro_cnpj: bairroCnpj,
    cidade_cnpj: cidadeCnpj,
    uf_cnpj: ufCnpj,
    numero_cnpj: numeroCnpj,

    risco_marca: riscoMarca,
    nacionalidade,
    selecao_cnpj_ou_cpf: selecaoCnpjOuCpf,

    vendedor
  };
}

/* =========================
 * Template ADD
 * =======================*/
function montarADDWord(d, nowInfo){
  const valorTotalNum = onlyNumberBR(d.valor_total);
  const parcelaNum = parseInt(String(d.parcelas||'1'),10)||1;
  const valorParcela = parcelaNum>0 ? valorTotalNum/parcelaNum : 0;

  const rua    = d.rua_cnpj || '';
  const bairro = d.bairro_cnpj || '';
  const numero = d.numero_cnpj || '';
  const cidade = d.cidade_cnpj || '';
  const uf     = d.uf_cnpj || '';
  const cep    = d.cep_cnpj || '';

  const valorDaTaxa = d.valor_taxa_brl || '';
  const formaDaTaxa = d.forma_pagto_taxa || '';
  const dataDaTaxa  = d.data_pagto_taxa || '';

  const dia = String(nowInfo.dia).padStart(2,'0');
  const mesNum = String(nowInfo.mes).padStart(2,'0');
  const ano = String(nowInfo.ano);
  const mesExtenso = monthNamePt(nowInfo.mes);

  const baseVars = {
    contratante_1: d.nome || '',
    cpf: d.cpf || '',
    cnpj: d.cnpj || '',
    rg: d.rg || '',
    'Estado Civíl': d.estado_civil || '',
    'Estado Civil': d.estado_civil || '',

    'CPF/CNPJ': d.selecao_cnpj_ou_cpf || '',
    'CPF': d.cpf_campo || '',
    'CNPJ': d.cnpj_campo || '',

    rua,
    bairro,
    numero,
    nome_da_cidade: cidade,
    cidade,
    uf,
    cep,

    'E-mail': d.email || '',
    'Telefone': d.telefone || '',
    telefone: d.telefone || '',

    nome_da_marca: d.titulo || '',
    classe: d.classe || '',
    'Quantidade depósitos/processos de MARCA': d.qtd_marca || '',
    'tipo de marca': d.tipo_marca || '',
    risco_da_marca: d.risco_marca || '',

    'marcas-espec': d.marcas_espec || '',

    Nacionalidade: d.nacionalidade || '',

    numero_de_parcelas_da_assessoria: String(d.parcelas||'1'),
    valor_da_parcela_da_assessoria: toBRL(valorParcela),
    forma_de_pagamento_da_assessoria: d.forma_pagto_assessoria || '',
    data_de_pagamento_da_assessoria: d.data_pagto_assessoria || '',
    'Data de pagamento da Assessoria': d.data_pagto_assessoria || '',

    valor_da_taxa: valorDaTaxa,
    forma_de_pagamento_da_taxa: formaDaTaxa,
    data_de_pagamento_da_taxa: dataDaTaxa,
    'Valor da Taxa': valorDaTaxa,
    'Forma de pagamento da Taxa': formaDaTaxa,
    'Data de pagamento da Taxa': dataDaTaxa,

    dia,
    mes: mesNum,
    ano,
    mes_extenso: mesExtenso,
    'Mês': mesExtenso,
    'Mes': mesExtenso,

    TEMPLATE_UUID_CONTRATO: TEMPLATE_UUID_CONTRATO || ''
  };

  return baseVars;
}

function montarSigners(d){
  const list = [];
  if (d.email) list.push({ email: d.email, name: d.nome || d.titulo || d.email, act:'1', foreign:'0', send_email:'1' });
  if (EMAIL_ASSINATURA_EMPRESA) list.push({ email: EMAIL_ASSINATURA_EMPRESA, name: 'Empresa', act:'1', foreign:'0', send_email:'1' });
  const seen={}; return list.filter(s => (seen[s.email.toLowerCase()]? false : (seen[s.email.toLowerCase()]=true)));
}

/* =========================
 * Locks e preflight
 * =======================*/
const locks = new Set();
function acquireLock(key){ if (locks.has(key)) return false; locks.add(key); return true; }
function releaseLock(key){ locks.delete(key); }
async function preflightDNS(){}

/* =========================
 * D4Sign
 * =======================*/
async function makeDocFromWordTemplate(tokenAPI, cryptKey, uuidSafe, templateId, title, varsObj) {
  const base = 'https://secure.d4sign.com.br';
  const url = new URL(`/api/v1/documents/${uuidSafe}/makedocumentbytemplateword`, base);
  url.searchParams.set('tokenAPI', tokenAPI);
  url.searchParams.set('cryptKey', cryptKey);
  const body = { name_document: title, templates: { [templateId]: varsObj } };
  const res = await fetchWithRetry(url.toString(), {
    method: 'POST', headers: { 'Accept': 'application/json', 'Content-Type': 'application/json' },
    body: JSON.stringify(body)
  }, { attempts: 2, baseDelayMs: 300, timeoutMs: 10_000 });
  const text = await res.text();
  let json; try { json = JSON.parse(text); } catch { json = null; }
  if (!res.ok || !(json && (json.uuid || json.uuid_document))) {
    console.error('[ERRO D4SIGN WORD]', res.status, text);
    throw new Error(`Falha D4Sign(WORD): ${res.status}`);
  }
  return json.uuid || json.uuid_document;
}
async function cadastrarSignatarios(tokenAPI, cryptKey, uuidDocument, signers) {
  const base = 'https://secure.d4sign.com.br';
  const url = new URL(`/api/v1/documents/${uuidDocument}/createlist`, base);
  url.searchParams.set('tokenAPI', tokenAPI);
  url.searchParams.set('cryptKey', cryptKey);
  const body = { signers: signers.map(s => ({ email: s.email, name: s.name, act: s.act || '1', foreign: s.foreign || '0', send_email: s.send_email || '1' })) };
  const res = await fetchWithRetry(url.toString(), {
    method: 'POST', headers: { 'Accept': 'application/json', 'Content-Type': 'application/json' },
    body: JSON.stringify(body)
  }, { attempts: 2, baseDelayMs: 300, timeoutMs: 10_000 });
  const text = await res.text();
  if (!res.ok) { console.error('[ERRO D4SIGN createlist]', res.status, text); throw new Error(`Falha ao cadastrar signatários: ${res.status}`); }
  return text;
}
async function getDownloadUrl(tokenAPI, cryptKey, uuidDocument, { type = 'PDF', language = 'pt' } = {}) {
  const base = 'https://secure.d4sign.com.br';
  const url = new URL(`/api/v1/documents/${uuidDocument}/download`, base);
  url.searchParams.set('tokenAPI', tokenAPI);
  url.searchParams.set('cryptKey', cryptKey);
  const body = { type, language, document: 'false' };
  const res = await fetchWithRetry(url.toString(), {
    method: 'POST',
    headers: { 'Accept': 'application/json', 'Content-Type': 'application/json' },
    body: JSON.stringify(body)
  }, { attempts: 2, baseDelayMs: 300, timeoutMs: 10_000 });
  const text = await res.text();
  let json; try { json = JSON.parse(text); } catch { json = null; }
  if (!res.ok || !json?.url) {
    console.error('[ERRO D4SIGN download]', res.status, text);
    throw new Error(`Falha ao gerar URL de download: ${res.status}`);
  }
  return json;
}
async function sendToSigner(tokenAPI, cryptKey, uuidDocument, {
  message = '',
  skip_email = '0',
  workflow = '0'
} = {}) {
  const base = 'https://secure.d4sign.com.br';
  const url = new URL(`/api/v1/documents/${uuidDocument}/sendtosigner`, base);
  url.searchParams.set('cryptKey', cryptKey);
  const body = { message, skip_email, workflow, tokenAPI };
  const res = await fetchWithRetry(url.toString(), {
    method: 'POST',
    headers: { 'Accept': 'application/json', 'Content-Type': 'application/json' },
    body: JSON.stringify(body)
  }, { attempts: 2, baseDelayMs: 300, timeoutMs: 10_000 });
  const text = await res.text();
  if (!res.ok) {
    console.error('[ERRO D4SIGN sendtosigner]', res.status, text);
    throw new Error(`Falha ao enviar para assinatura: ${res.status}`);
  }
  return text;
}

/* =========================
 * Move fase
 * =======================*/
async function moveCardToPhaseSafe(cardId, phaseId){
  if (!phaseId) return;
  await gql(`mutation($input: MoveCardToPhaseInput!){
    moveCardToPhase(input:$input){ card{ id } }
  }`, { input: { card_id: Number(cardId), destination_phase_id: Number(phaseId) } }).catch(e=>{
    console.warn('[WARN] moveCardToPhaseSafe', e.message||e);
  });
}

/* =========================
 * Rotas
 * =======================*/
app.get('/lead/:token', async (req, res) => {
  try {
    const { cardId } = parseLeadToken(req.params.token);

    const card = await getCard(cardId);
    const d = await montarDados(card);

    const html = `
<!doctype html><html lang="pt-BR"><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Revisar contrato</title>
<style>
  body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Ubuntu,Arial,sans-serif;margin:0;background:#f7f7f7;color:#111}
  .wrap{max-width:920px;margin:24px auto;padding:0 16px}
  .card{background:#fff;border-radius:14px;box-shadow:0 4px 16px rgba(0,0,0,.08);padding:24px;margin-bottom:16px}
  h1{font-size:22px;margin:0 0 12px}
  h2{font-size:16px;margin:24px 0 8px}
  .grid{display:grid;grid-template-columns:1fr 1fr;gap:12px}
  .grid3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px}
  .btn{display:inline-block;padding:12px 18px;border-radius:10px;text-decoration:none;border:0;background:#111;color:#fff;font-weight:600;cursor:pointer}
  .muted{color:#666}
  .label{font-weight:700}
  .tag{display:inline-block;background:#111;color:#fff;border-radius:8px;padding:4px 8px;font-size:12px;margin-left:8px}
</style>
<div class="wrap">
  <div class="card">
    <h1>Revisar dados do contrato <span class="tag">Card #${card.id}</span></h1>

    <h2>Contratante</h2>
    <div class="grid">
      <div><div class="label">Nome</div><div>${d.nome||'-'}</div></div>
      <div><div class="label">Nacionalidade</div><div>${d.nacionalidade||'-'}</div></div>
      <div><div class="label">Estado Civíl</div><div>${d.estado_civil||'-'}</div></div>
      <div><div class="label">CPF/CNPJ (seleção)</div><div>${d.selecao_cnpj_ou_cpf||'-'}</div></div>
      <div><div class="label">CPF (campo)</div><div>${d.cpf_campo||'-'}</div></div>
      <div><div class="label">CNPJ (campo)</div><div>${d.cnpj_campo||'-'}</div></div>
      <div><div class="label">RG</div><div>${d.rg||'-'}</div></div>
    </div>

    <h2>Contato</h2>
    <div class="grid">
      <div><div class="label">E-mail</div><div>${d.email||'-'}</div></div>
      <div><div class="label">Telefone</div><div>${d.telefone||'-'}</div></div>
    </div>

    <h2>Marca</h2>
    <div class="grid3">
      <div><div class="label">Nome da marca</div><div>${d.titulo||'-'}</div></div>
      <div><div class="label">Classes</div><div>${d.classe||'-'}</div></div>
      <div><div class="label">CLASSES E ESPECIFICAÇÕES</div><div>${(d.marcas_espec||'').replace(/\n/g,'<br>')||'-'}</div></div>
      <div><div class="label">Risco da marca</div><div>${d.risco_marca||'-'}</div></div>
      <div><div class="label">Qtd. de marcas</div><div>${d.qtd_marca||'0'}</div></div>
    </div>

    <h2>Serviços</h2>
    <div>${(d.servicos||[]).join(', ') || '-'}</div>

    <h2>Remuneração — Assessoria</h2>
    <div class="grid3">
      <div><div class="label">Valor total</div><div>${d.valor_total||'-'}</div></div>
      <div><div class="label">Parcelas</div><div>${String(d.parcelas||'1')}</div></div>
      <div><div class="label">Forma de pagamento</div><div>${d.forma_pagto_assessoria||'-'}</div></div>
      <div><div class="label">Data de pagamento</div><div>${d.data_pagto_assessoria||'-'}</div></div>
    </div>

    <h2>Taxa</h2>
    <div class="grid3">
      <div><div class="label">Valor da Taxa</div><div>${d.valor_taxa_brl || '-'}</div></div>
      <div><div class="label">Forma de pagamento</div><div>${d.forma_pagto_taxa || '-'}</div></div>
      <div><div class="label">Data de pagamento</div><div>${d.data_pagto_taxa || '-'}</div></div>
    </div>

    <form method="POST" action="/lead/${encodeURIComponent(req.params.token)}/generate" style="margin-top:24px">
      <button class="btn" type="submit">Gerar contrato</button>
    </form>
    <p class="muted" style="margin-top:12px">Ao clicar, o documento será criado no D4Sign e o card poderá ser movido para Contrato enviado.</p>
  </div>
</div>
`;
    res.setHeader('content-type','text/html; charset=utf-8');
    return res.status(200).send(html);
  } catch (e) {
    console.error('[ERRO /lead]', e.message||e);
    return res.status(400).send('Link inválido ou expirado.');
  }
});

app.post('/lead/:token/generate', async (req, res) => {
  try {
    const { cardId } = parseLeadToken(req.params.token);
    const lockKey = `lead:${cardId}`;
    if (!acquireLock(lockKey)) return res.status(200).send('Processando, tente novamente em instantes.');

    preflightDNS().catch(()=>{});

    const card = await getCard(cardId);
    const d = await montarDados(card);

    const now = new Date();
    const nowInfo = { dia: now.getDate(), mes: now.getMonth()+1, ano: now.getFullYear() };
    const add = montarADDWord(d, nowInfo);
    const signers = montarSigners(d);

    const uuidSafe = COFRES_UUIDS[d.vendedor] || DEFAULT_COFRE_UUID;
    if (!uuidSafe) throw new Error(`Cofre não configurado para vendedor: ${d.vendedor}`);

    const uuidDoc = await makeDocFromWordTemplate(D4SIGN_TOKEN, D4SIGN_CRYPT_KEY, uuidSafe, TEMPLATE_UUID_CONTRATO, card.title, add);

    await Promise.all([
      cadastrarSignatarios(D4SIGN_TOKEN, D4SIGN_CRYPT_KEY, uuidDoc, signers).catch(e => console.error('createlist erro', e)),
      moveCardToPhaseSafe(card.id, PHASE_ID_CONTRATO_ENVIADO).catch(e => console.warn('move phase warn', e))
    ]);

    releaseLock(lockKey);

    const token = req.params.token;
    const html = `
<!doctype html><meta charset="utf-8"><title>Contrato gerado</title>
<style>
  body{font-family:system-ui;display:grid;place-items:center;min-height:100vh;background:#f7f7f7;color:#111;margin:0}
  .box{background:#fff;padding:24px;border-radius:14px;box-shadow:0 4px 16px rgba(0,0,0,.08);max-width:640px;width:92%}
  h2{margin:0 0 12px}
  .row{display:flex;gap:12px;flex-wrap:wrap;margin-top:12px}
  .btn{display:inline-block;padding:12px 16px;border-radius:10px;text-decoration:none;border:0;background:#111;color:#fff;font-weight:600}
  .muted{color:#666}
</style>
<div class="box">
  <h2>Contrato gerado com sucesso</h2>
  <p class="muted">UUID do documento: ${uuidDoc}</p>
  <div class="row">
    <a class="btn" href="/lead/${encodeURIComponent(token)}/doc/${encodeURIComponent(uuidDoc)}/download" target="_blank" rel="noopener">Baixar PDF</a>
    <form method="POST" action="/lead/${encodeURIComponent(token)}/doc/${encodeURIComponent(uuidDoc)}/send" style="display:inline">
      <button class="btn" type="submit">Enviar para assinatura</button>
    </form>
    <a class="btn" href="${PUBLIC_BASE_URL}/lead/${encodeURIComponent(token)}">Voltar</a>
  </div>
</div>`;
    return res.status(200).send(html);

  } catch (e) {
    console.error('[ERRO LEAD-GENERATE]', e.message || e);
    return res.status(400).send('Falha ao gerar o contrato.');
  }
});

app.get('/lead/:token/doc/:uuid/download', async (req, res) => {
  try {
    const { cardId } = parseLeadToken(req.params.token);
    if (!cardId) throw new Error('token inválido');
    const uuidDoc = req.params.uuid;

    const { url: downloadUrl } = await getDownloadUrl(D4SIGN_TOKEN, D4SIGN_CRYPT_KEY, uuidDoc, { type: 'PDF', language: 'pt' });
    return res.redirect(302, downloadUrl);
  } catch (e) {
    console.error('[ERRO lead download]', e.message || e);
    return res.status(400).send('Falha ao gerar link de download.');
  }
});

app.post('/lead/:token/doc/:uuid/send', async (req, res) => {
  try {
    const { cardId } = parseLeadToken(req.params.token);
    if (!cardId) throw new Error('token inválido');
    const uuidDoc = req.params.uuid;

    await sendToSigner(D4SIGN_TOKEN, D4SIGN_CRYPT_KEY, uuidDoc, {
      message: 'Olá! Há um documento aguardando sua assinatura.',
      skip_email: '0',
      workflow: '0'
    });

    const okHtml = `
<!doctype html><meta charset="utf-8"><title>Documento enviado</title>
<style>body{font-family:system-ui;display:grid;place-items:center;height:100vh;background:#f7f7f7} .box{background:#fff;padding:24px;border-radius:14px;box-shadow:0 4px 16px rgba(0,0,0,.08);max-width:560px}</style>
<div class="box">
  <h2>Documento enviado para assinatura</h2>
  <p>Os signatários foram notificados.</p>
  <p><a href="${PUBLIC_BASE_URL}/lead/${encodeURIComponent(req.params.token)}">Voltar</a></p>
</div>`;
    return res.status(200).send(okHtml);

  } catch (e) {
    console.error('[ERRO sendtosigner]', e.message || e);
    return res.status(400).send('Falha ao enviar para assinatura.');
  }
});

/* =========================
 * Link no Pipefy
 * =======================*/
app.post('/novo-pipe/criar-link-confirmacao', async (req, res) => {
  try {
    const cardId = req.body.cardId || req.body.card_id || req.query.cardId || req.query.card_id;
    if (!cardId) return res.status(400).json({ error: 'cardId é obrigatório' });

    const card = await getCard(cardId);
    if (NOVO_PIPE_ID && String(card?.pipe?.id)!==String(NOVO_PIPE_ID)) {
      return res.status(400).json({ error: 'Card não pertence ao pipe configurado' });
    }
    if (FASE_VISITA_ID && String(card?.current_phase?.id)!==String(FASE_VISITA_ID)) {
      return res.status(400).json({ error: 'Card não está na fase esperada' });
    }

    const token = makeLeadToken({ cardId: String(cardId), ts: Date.now() });
    const url = `${PUBLIC_BASE_URL.replace(/\/+$/,'')}/lead/${encodeURIComponent(token)}`;

    await updateCardField(cardId, PIPEFY_FIELD_LINK_CONTRATO, url);

    return res.json({ ok:true, link:url });
  } catch (e) {
    console.error('[ERRO criar-link]', e.message||e);
    return res.status(500).json({ error: String(e.message||e) });
  }
});
app.get('/novo-pipe/criar-link-confirmacao', async (req,res)=>{
  req.body = req.body || {};
  req.body.cardId = req.query.cardId || req.query.card_id;
  return app._router.handle(req, res, ()=>{});
});

/* =========================
 * Debug
 * =======================*/
app.get('/_echo/*', (req, res) => {
  res.json({
    method: req.method,
    originalUrl: req.originalUrl,
    path: req.path,
    baseUrl: req.baseUrl,
    host: req.get('host'),
    href: `${req.protocol}://${req.get('host')}${req.originalUrl}`,
    headers: req.headers,
    query: req.query,
  });
});
app.get('/debug/card', async (req,res)=>{
  try{
    const { cardId } = req.query; if (!cardId) return res.status(400).send('cardId obrigatório');
    const card = await getCard(cardId);
    res.json({
      id: card.id, title: card.title, pipe: card.pipe, phase: card.current_phase,
      fields: (card.fields||[]).map(f => ({ name:f.name, id:f.field?.id, type:f.field?.type, value:f.value }))
    });
  }catch(e){ res.status(500).json({ error:String(e.message||e) }); }
});
app.get('/health', (_req,res)=> res.json({ ok:true }));

/* =========================
 * Start
 * =======================*/
app.listen(PORT, () => {
  console.log(`Servidor rodando na porta ${PORT}`);
  const list=[];
  app._router.stack.forEach(m=>{
    if (m.route && m.route.path){
      const methods = Object.keys(m.route.methods).map(x=>x.toUpperCase()).join(',');
      list.push(`${methods} ${m.route.path}`);
    } else if (m.name==='router' && m.handle?.stack){
      m.handle.stack.forEach(h=>{
        const route = h.route;
        if (route){
          const methods = Object.keys(route.methods).map(x=>x.toUpperCase()).join(',');
          list.push(`${methods} ${route.path}`);
        }
      });
    }
  });
  console.log('[rotas-registradas]'); list.sort().forEach(r=>console.log('  -', r));
});

/**
 * ENV
 *
 * PUBLIC_BASE_URL=https://seu-dominio.com
 * PUBLIC_LINK_SECRET=um-segredo-forte
 *
 * PIPE_API_KEY=...
 * PIPE_GRAPHQL_ENDPOINT=https://api.pipefy.com/graphql
 *
 * PIPEFY_FIELD_LINK_CONTRATO=d4_contrato
 * FIELD_ID_CONNECT_MARCA_NOME=marcas_1
 * FIELD_ID_CONNECT_CLASSE=marcas_2
 * FIELD_ID_CONNECT_CLASSES=classes_inpi
 * MARCAS_TABLE_ID=MmqLNaPk
 * CONTACTS_TABLE_ID=Pbp9mARx
 * MARCAS2_TABLE_ID=tnDAtg7l
 * CLASSES_TABLE_ID=306521337
 *
 * D4SIGN_TOKEN=...
 * D4SIGN_CRYPT_KEY=...
 * TEMPLATE_UUID_CONTRATO=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
 *
 * NOVO_PIPE_ID=306505295
 * FASE_VISITA_ID=339299691
 * PHASE_ID_CONTRATO_ENVIADO=XXXXXXXX
 *
 * EMAIL_ASSINATURA_EMPRESA=contratos@empresa.com.br
 *
 * COFRE_UUID_EDNA=...
 * COFRE_UUID_GREYCE=...
 * COFRE_UUID_MARIANA=...
 * COFRE_UUID_VALDEIR=...
 * COFRE_UUID_DEBORA=...
 * COFRE_UUID_MAYKON=...
 * COFRE_UUID_JEFERSON=...
 * COFRE_UUID_RONALDO=...
 * COFRE_UUID_BRENDA=...
 * COFRE_UUID_MAURO=...
 * DEFAULT_COFRE_UUID=...
 */
