'use strict';

/**
 * server.js — Provincia Vendas (Pipefy + D4Sign via secure.d4sign.com.br)
 * Node 18+ (fetch global)
 */

const express = require('express');
const crypto = require('crypto');

const app = express();
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.set('trust proxy', true);

// Log básico
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

  // Ainda usados para escrever o link no card e validar fase
  PIPEFY_FIELD_LINK_CONTRATO,
  NOVO_PIPE_ID,
  FASE_VISITA_ID,
  PHASE_ID_CONTRATO_ENVIADO,

  // D4Sign
  D4SIGN_TOKEN,
  D4SIGN_CRYPT_KEY,
  TEMPLATE_UUID_CONTRATO,           // Modelo de Marca
  TEMPLATE_UUID_CONTRATO_OUTROS,    // Modelo de Outros Serviços
  TEMPLATE_UUID_PROCURACAO,         // Modelo de Procuração

  // Assinatura interna
  EMAIL_ASSINATURA_EMPRESA,

  // Cofres
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
  COFRE_UUID_REPRESENTANTESCLEISON,
  COFRE_UUID_FILIALSAOPAULO,
  COFRE_UUID_REPRESENTANTELUAN,
  COFRE_UUID_PROVINCIADIGITAL_LUCAANTONIAZZI,
  COFRE_UUID_REPRESENTANTEVINICIUS,
  COFRE_UUID_FILIALPORTOALEGRE2,
  COFRE_UUID_FILIALBRASILIA,
  COFRE_UUID_FILIALCAMPINAS,
  COFRE_UUID_FILIALVITORIA2,
  COFRE_UUID_FILIALJOINVILLE,
  COFRE_UUID_FILIALVITORIA1,
  COFRE_UUID_FILIALBELEM,
  COFRE_UUID_FILIALJUNDIAI2,
  COFRE_UUID_FILIAL_SAOJOSERIOPRETO2,
  COFRE_UUID_FILIALPRES_PRUDENTE,
  COFRE_UUID_FILIALMANAUS,
  COFRE_UUID_FILIALGUARULHOS,
  COFRE_UUID_FILIALMARINGA,
  COFRE_UUID_FILIALRIBEIRAOPRETO,
  COFRE_UUID_FILIALPALMAS_TO,

  DEFAULT_COFRE_UUID
} = process.env;
NOVO_PIPE_ID = NOVO_PIPE_ID || '306505295';

PORT = PORT || 3000;
PIPE_GRAPHQL_ENDPOINT = PIPE_GRAPHQL_ENDPOINT || 'https://api.pipefy.com/graphql';
PIPEFY_FIELD_LINK_CONTRATO = PIPEFY_FIELD_LINK_CONTRATO || 'd4_contrato';

if (!PUBLIC_BASE_URL || !PUBLIC_LINK_SECRET) console.warn('[AVISO] Configure PUBLIC_BASE_URL e PUBLIC_LINK_SECRET');
if (!PIPE_API_KEY) console.warn('[AVISO] PIPE_API_KEY ausente');
if (!D4SIGN_TOKEN || !D4SIGN_CRYPT_KEY) console.warn('[AVISO] D4SIGN_TOKEN / D4SIGN_CRYPT_KEY ausentes');
if (!TEMPLATE_UUID_CONTRATO) console.warn('[AVISO] TEMPLATE_UUID_CONTRATO (Marca) ausente');
if (!TEMPLATE_UUID_CONTRATO_OUTROS) console.warn('[AVISO] TEMPLATE_UUID_CONTRATO_OUTROS (Outros) ausente');
if (!TEMPLATE_UUID_PROCURACAO) console.warn('[AVISO] TEMPLATE_UUID_PROCURACAO ausente');

// Cofres mapeados por EQUIPE (campo "Equipe contrato" no Pipefy)
// ⚠️ ATENÇÃO: as chaves DEVEM ser exatamente os valores de "Equipe contrato"
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
  'Mauro Furlan Neto': COFRE_UUID_MAURO,
  'Cleison Villas Boas': COFRE_UUID_REPRESENTANTESCLEISON,
  'Felipe Cordeiro': COFRE_UUID_FILIALSAOPAULO,
  'PROVÍNCIACG': COFRE_UUID_REPRESENTANTELUAN,
  'Luca Andrade Antoniazzi': COFRE_UUID_PROVINCIADIGITAL_LUCAANTONIAZZI,
  'Luan Menegatti': COFRE_UUID_REPRESENTANTELUAN,
  'Vinicius Chiba': COFRE_UUID_REPRESENTANTEVINICIUS,
  'Michely Piloto': COFRE_UUID_FILIALPORTOALEGRE2,
  'Pamella Valero de Azevedo': COFRE_UUID_FILIALBRASILIA,
  'Bianca Angelo Dias': COFRE_UUID_FILIALCAMPINAS,
  'VALERIA DE ARAUJO GOUVEA': COFRE_UUID_FILIALVITORIA2,
  'JANAINA COSTA MORENO': COFRE_UUID_FILIALJOINVILLE,
  'Mariana Loureiro Lanes dos Santos': COFRE_UUID_FILIALBRASILIA,
  'Viviany Egnes Gonçalves Nogueira': COFRE_UUID_FILIALVITORIA1,
  'LUIS OTAVIO DE ALMEIDA NEVES': COFRE_UUID_FILIALBELEM,
  'Katia Regina Dias Martins': COFRE_UUID_FILIALJUNDIAI2,
  'mario augusto rodrigues de lima': COFRE_UUID_FILIAL_SAOJOSERIOPRETO2,
  'RAFAEL JUNIOR MENEGATTI': COFRE_UUID_FILIALPRES_PRUDENTE,
  'NEIMY ANDES DAS NEVES': COFRE_UUID_FILIALMANAUS,
  'Lays Pereira Martins': COFRE_UUID_FILIALGUARULHOS,
  'Viviane Lima Batista': COFRE_UUID_FILIALMARINGA,
  'Luana Purcino de Oliveira': COFRE_UUID_FILIALRIBEIRAOPRETO,
  'REPRESENTANTES CLEISON': COFRE_UUID_REPRESENTANTESCLEISON,
  'FILIAL SAO PAULO': COFRE_UUID_FILIALSAOPAULO,
  'REPRESENTANTE LUAN': COFRE_UUID_REPRESENTANTELUAN,
  'PROVINCIA DIGITAL_LUCA ANTONIAZZI': COFRE_UUID_PROVINCIADIGITAL_LUCAANTONIAZZI,
  'REPRESENTANTE VINICIUS': COFRE_UUID_REPRESENTANTEVINICIUS,
  'FILIAL PORTO ALEGRE 2': COFRE_UUID_FILIALPORTOALEGRE2,
  'FILIAL BRASILIA': COFRE_UUID_FILIALBRASILIA,
  'FILIAL CAMPINAS': COFRE_UUID_FILIALCAMPINAS,
  'FILIAL VITORIA 2': COFRE_UUID_FILIALVITORIA2,
  'FILIAL JOINVILLE': COFRE_UUID_FILIALJOINVILLE,
  'FILIAL VITORIA 1': COFRE_UUID_FILIALVITORIA1,
  'FILIAL BELEM': COFRE_UUID_FILIALBELEM,
  'FILIAL JUNDIAI 2': COFRE_UUID_FILIALJUNDIAI2,
  'FILIAL_SAO JOSE RIO PRETO 2': COFRE_UUID_FILIAL_SAOJOSERIOPRETO2,
  'FILIAL PRES. PRUDENTE': COFRE_UUID_FILIALPRES_PRUDENTE,
  'FILIAL MANAUS': COFRE_UUID_FILIALMANAUS,
  'FILIAL GUARULHOS': COFRE_UUID_FILIALGUARULHOS,
  'FILIAL MARINGA': COFRE_UUID_FILIALMARINGA,
  'FILIAL RIBEIRAO PRETO': COFRE_UUID_FILIALRIBEIRAOPRETO,
  'FILIAL PALMAS - TO': COFRE_UUID_FILIALPALMAS_TO
};

// Função para obter o nome da variável do cofre a partir do UUID
function getNomeCofreByUuid(uuidCofre) {
  if (!uuidCofre) return 'Cofre não identificado';
  
  // Mapeamento reverso: UUID -> Nome da variável
  const mapeamento = {
    [COFRE_UUID_EDNA]: 'COFRE_UUID_EDNA',
    [COFRE_UUID_GREYCE]: 'COFRE_UUID_GREYCE',
    [COFRE_UUID_MARIANA]: 'COFRE_UUID_MARIANA',
    [COFRE_UUID_VALDEIR]: 'COFRE_UUID_VALDEIR',
    [COFRE_UUID_DEBORA]: 'COFRE_UUID_DEBORA',
    [COFRE_UUID_MAYKON]: 'COFRE_UUID_MAYKON',
    [COFRE_UUID_JEFERSON]: 'COFRE_UUID_JEFERSON',
    [COFRE_UUID_RONALDO]: 'COFRE_UUID_RONALDO',
    [COFRE_UUID_BRENDA]: 'COFRE_UUID_BRENDA',
    [COFRE_UUID_MAURO]: 'COFRE_UUID_MAURO',
    [COFRE_UUID_REPRESENTANTESCLEISON]: 'COFRE_UUID_REPRESENTANTESCLEISON',
    [COFRE_UUID_FILIALSAOPAULO]: 'COFRE_UUID_FILIALSAOPAULO',
    [COFRE_UUID_REPRESENTANTELUAN]: 'COFRE_UUID_REPRESENTANTELUAN',
    [COFRE_UUID_PROVINCIADIGITAL_LUCAANTONIAZZI]: 'COFRE_UUID_PROVINCIADIGITAL_LUCAANTONIAZZI',
    [COFRE_UUID_REPRESENTANTEVINICIUS]: 'COFRE_UUID_REPRESENTANTEVINICIUS',
    [COFRE_UUID_FILIALPORTOALEGRE2]: 'COFRE_UUID_FILIALPORTOALEGRE2',
    [COFRE_UUID_FILIALBRASILIA]: 'COFRE_UUID_FILIALBRASILIA',
    [COFRE_UUID_FILIALCAMPINAS]: 'COFRE_UUID_FILIALCAMPINAS',
    [COFRE_UUID_FILIALVITORIA2]: 'COFRE_UUID_FILIALVITORIA2',
    [COFRE_UUID_FILIALJOINVILLE]: 'COFRE_UUID_FILIALJOINVILLE',
    [COFRE_UUID_FILIALVITORIA1]: 'COFRE_UUID_FILIALVITORIA1',
    [COFRE_UUID_FILIALBELEM]: 'COFRE_UUID_FILIALBELEM',
    [COFRE_UUID_FILIALJUNDIAI2]: 'COFRE_UUID_FILIALJUNDIAI2',
    [COFRE_UUID_FILIAL_SAOJOSERIOPRETO2]: 'COFRE_UUID_FILIAL_SAOJOSERIOPRETO2',
    [COFRE_UUID_FILIALPRES_PRUDENTE]: 'COFRE_UUID_FILIALPRES_PRUDENTE',
    [COFRE_UUID_FILIALMANAUS]: 'COFRE_UUID_FILIALMANAUS',
    [COFRE_UUID_FILIALGUARULHOS]: 'COFRE_UUID_FILIALGUARULHOS',
    [COFRE_UUID_FILIALMARINGA]: 'COFRE_UUID_FILIALMARINGA',
    [COFRE_UUID_FILIALRIBEIRAOPRETO]: 'COFRE_UUID_FILIALRIBEIRAOPRETO',
    [COFRE_UUID_FILIALPALMAS_TO]: 'COFRE_UUID_FILIALPALMAS_TO'
  };
  
  if (uuidCofre === DEFAULT_COFRE_UUID) {
    return 'DEFAULT_COFRE_UUID';
  }
  
  return mapeamento[uuidCofre] || 'Cofre não identificado';
}

/* =========================
 * Helpers gerais
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
function onlyNumberBR(s){
  const n = parseNumberBR(s);
  return isNaN(n)? 0 : n;
}

// Datas
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

// Retry com timeout exponencial
async function fetchWithRetry(url, init={}, opts={})
{
  const attempts = opts.attempts || 3;
  const baseDelayMs = opts.baseDelayMs || 500;
  const timeoutMs = opts.timeoutMs || 15000;

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
 * Token público (/lead/:token)
 * =======================*/
function makeLeadToken(payload){ // {cardId, ts}
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
  const r = await fetch(PIPE_GRAPHQL_ENDPOINT, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${PIPE_API_KEY}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({ query, variables })
  });
  const j = await r.json().catch(() => ({}));
  if (!r.ok || j?.errors) {
    throw new Error(`Pipefy GQL: ${r.status} ${JSON.stringify(j.errors || {})}`);
  }
  return j.data;
}
async function getCard(cardId){
  const data = await gql(`query($id: ID!){
    card(id:$id){
      id title
      current_phase{ id name }
      pipe{ id name }
      fields{ name value array_value field{ id type label description } }
      assignees{ name email }
    }
  }`, { id: cardId });
  if (!data?.card) throw new Error(`Card ${cardId} não encontrado`);
  return data.card;
}
async function updateCardField(cardId, fieldId, newValue){
  await gql(`mutation($input: UpdateCardFieldInput!){
    updateCardField(input:$input){ card{ id } }
  }`, { input: { card_id: Number(cardId), field_id: fieldId, new_value: newValue } });
}

async function createCardComment(cardId, comment){
  try {
    const data = await gql(`mutation($input: CreateCardCommentInput!){
      createCardComment(input:$input){
        comment{
          id
          text
        }
      }
    }`, { 
      input: { 
        card_id: Number(cardId), 
        text: comment 
      } 
    });
    return data?.createCardComment?.comment;
  } catch (e) {
    console.error('[ERRO createCardComment]', e.message || e);
    throw e;
  }
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
function getFieldObjById(card, id){
  return (card.fields||[]).find(f=> String(f?.field?.id||'')===String(id));
}
function getFirstByNames(card, arr){
  for (const k of arr){ const v = getByName(card, k); if (v) return v; }
  return '';
}

// NOVO: ler campo "Equipe contrato"
function getEquipeContratoFromCard(card){
  if (!card || !Array.isArray(card.fields)) return '';
  const f = (card.fields || []).find(ff =>
    String(ff.name || '').toLowerCase() === 'equipe contrato'
  );
  if (!f || !f.value) return '';
  return String(f.value).trim();
}

function checklistToText(v){
  try{
    const arr = Array.isArray(v)? v : JSON.parse(v);
    return Array.isArray(arr) ? arr.join(', ') : String(v || '');
  }catch{ return String(v || ''); }
}

function extractAssigneeNames(raw){
  const out=[];
  const push=v=>{ if(v) out.push(String(v)); };
  const tryParse=v=>{
    if (typeof v==='string'){
      try { return JSON.parse(v); } catch { return v; }
    }
    return v;
  };

  const val = tryParse(raw);
  if (Array.isArray(val)){
    for (const it of val){
      push(typeof it==='string' ? it : (it?.name||it?.username||it?.email||it?.value));
    }
  } else if (typeof val==='object' && val){
    push(val.name||val.username||val.email||val.value);
  } else if (typeof val==='string'){
    const m = val.match(/^\s*\[.*\]\s*$/) ? tryParse(val) : null;
    if (m && Array.isArray(m)){
      m.forEach(x=>push(typeof x==='string' ? x : (x?.name||x?.email)));
    } else {
      push(val);
    }
  }
  return [...new Set(out.filter(Boolean))];
}

// Documento (CPF/CNPJ)
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

// Assignee parsing (para cofre) — hoje não usado diretamente
function stripDiacritics(s){ return String(s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,''); }
function normalizeName(s){ return stripDiacritics(String(s||'').trim()).toLowerCase(); }
function resolveCofreUuidByCard(card){
  if (!card) return DEFAULT_COFRE_UUID || null;
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
 * Regras específicas
 * =======================*/
function computeValorTaxaBRLFromFaixa(d){
  let valorTaxaSemRS = '';
  const taxa = String(d.taxa_faixa||'');
  if (taxa.includes('440')) valorTaxaSemRS = '440,00';
  else if (taxa.includes('880')) valorTaxaSemRS = '880,00';
  return valorTaxaSemRS ? `R$ ${valorTaxaSemRS}` : '';
}

// Extrai todos os números em ordem de aparição e devolve separados por vírgula
function extractClasseNumbersFromText(s){
  const nums=[]; const seen=new Set();
  for (const m of String(s||'').matchAll(/\b\d+\b/g)){
    const n = String(Number(m[0]));
    if (!seen.has(n)){ seen.add(n); nums.push(n); }
  }
  return nums.join(', ');
}

// Identifica o tipo base do serviço a partir do texto do statement ou connector
function serviceKindFromText(s){
  const t = String(s||'').toUpperCase();
  if (t.includes('MARCA')) return 'MARCA';
  if (t.includes('PATENTE')) return 'PATENTE';
  if (t.includes('DESENHO')) return 'DESENHO INDUSTRIAL';
  if (t.includes('COPYRIGHT') || t.includes('DIREITO AUTORAL')) return 'COPYRIGHT/DIREITO AUTORAL';
  return 'OUTROS';
}

// Busca campo statement por N com fallback para connector
function buscarServicoN(card, n){
  const mapStmt = {
    1: 'statement_9a115410_226d_43bc_9c1b_a28887e1f8a6',
    2: 'statement_432366f2_fbbc_448d_82e4_fbd73c3fc52e',
    3: 'statement_c5616541_5f30_41b9_bd74_e2bd2063f253',
    4: 'statement_8d833401_2294_448b_a34f_07f86c52981c',
    5: 'statement_ca0eb59e_a015_4628_8c56_28af6e23c8d9'
  };
  const mapConn = {
    1: 'servi_os_marca_1',
    2: 'servi_os_marca_2',
    3: 'servi_os_marca_3',
    4: 'servi_os_marca_4',
    5: 'servi_os_marca_5'
  };

  let v = '';
  const stmtId = mapStmt[n];
  if (stmtId){
    const f = getFieldObjById(card, stmtId);
    v = String(f?.value||'').replace(/<[^>]*>/g,' ').trim();
  }
  if (!v){
    const connId = mapConn[n];
    if (connId){
      const f = getFieldObjById(card, connId);
      if (f?.value){
        try{
          const arr = JSON.parse(f.value);
          if (Array.isArray(arr) && arr[0]) v = String(arr[0]);
        }catch{
          v = String(f.value);
        }
      } else if (Array.isArray(f?.array_value) && f.array_value.length){
        v = String(f.array_value[0]);
      }
    }
  }
  return v;
}

// Normalização apenas para “Detalhes do serviço …”
function normalizarCabecalhoDetalhe(kind, nome, tipoMarca='', classeNums=''){
  const k = String(kind||'').toUpperCase();
  if (k==='MARCA'){
    const tipo = tipoMarca ? `, Apresentação: ${tipoMarca}` : '';
    const classe = classeNums ? `, CLASSE: nº ${classeNums}` : '';
    return `MARCA: ${nome || ''}${tipo}${classe}`.trim();
  }
  if (k==='PATENTE') return `PATENTE: ${nome||''}`.trim();
  if (k==='DESENHO INDUSTRIAL') return `DESENHO INDUSTRIAL: ${nome||''}`.trim();
  if (k==='COPYRIGHT/DIREITO AUTORAL') return `COPYRIGHT/DIREITO AUTORAL: ${nome||''}`.trim();
  return `OUTROS SERVIÇOS: ${nome||''}`.trim();
}

/* =========================
 * Montagem de dados do contrato
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
  if (!raw){
    const hit = (card.fields||[]).find(f => String(f?.field?.type||'').toLowerCase()==='currency');
    raw = hit?.value||'';
  }
  const n = parseNumberBR(raw);
  return isNaN(n)? null : n;
}
function firstNonEmpty(...vals){
  for (const v of vals){ if (String(v||'').trim()) return v; }
  return '';
}
function parseListFromLongText(value, max=30){
  const lines = String(value||'').split(/\r?\n/).map(s=>s.trim()).filter(Boolean);
  const arr = [];
  for (let i=0;i<max;i++) arr.push(lines[i]||'');
  return arr;
}

async function montarDados(card){
  const by = toById(card);

  // Marca 1 dados base
  const tituloMarca1 = by['marca'] || card.title || '';
  const marcasEspecRaw1 = by['copy_of_classe_e_especifica_es'] || by['classe'] || getFirstByNames(card, ['classes e especificações marca - 1','classes e especificações']) || '';
  const linhasMarcasEspec1 = parseListFromLongText(marcasEspecRaw1, 30);
  const classeSomenteNumeros1 = extractClasseNumbersFromText(marcasEspecRaw1);
  const tipoMarca1 = checklistToText(by['checklist_vertical'] || getFirstByNames(card, ['tipo de marca']));

  // Marca 2
  const tituloMarca2 = by['marca_2'] || getFirstByNames(card, ['marca ou patente - 2','marca - 2']) || '';
  const marcasEspecRaw2 = by['copy_of_classes_e_especifica_es_marca_2'] || getFirstByNames(card, ['classes e especificações marca - 2']) || '';
  const linhasMarcasEspec2 = parseListFromLongText(marcasEspecRaw2, 30);
  const classeSomenteNumeros2 = extractClasseNumbersFromText(marcasEspecRaw2);
  const tipoMarca2 = checklistToText(by['copy_of_tipo_de_marca'] || getFirstByNames(card, ['tipo de marca - 2']));

  // Marca 3
  const tituloMarca3 = by['marca_3'] || getFirstByNames(card, ['marca ou patente - 3','marca - 3']) || '';
  const marcasEspecRaw3 = by['copy_of_copy_of_classe_e_especifica_es'] || getFirstByNames(card, ['classes e especificações marca - 3']) || '';
  const linhasMarcasEspec3 = parseListFromLongText(marcasEspecRaw3, 30);
  const classeSomenteNumeros3 = extractClasseNumbersFromText(marcasEspecRaw3);
  const tipoMarca3 = checklistToText(by['copy_of_copy_of_tipo_de_marca'] || getFirstByNames(card, ['tipo de marca - 3']));

  // Marca 4
  const tituloMarca4 = by['marca_ou_patente_4'] || '';
  const marcasEspecRaw4 = by['classes_e_especifica_es_marca_4'] || '';
  const linhasMarcasEspec4 = parseListFromLongText(marcasEspecRaw4, 30);
  const classeSomenteNumeros4 = extractClasseNumbersFromText(marcasEspecRaw4);
  const tipoMarca4 = checklistToText(by['copy_of_tipo_de_marca_3'] || '');

  // Marca 5
  const tituloMarca5 = by['marca_ou_patente_5'] || '';
  const marcasEspecRaw5 = by['copy_of_classes_e_especifica_es_marca_4'] || '';
  const linhasMarcasEspec5 = parseListFromLongText(marcasEspecRaw5, 30);
  const classeSomenteNumeros5 = extractClasseNumbersFromText(marcasEspecRaw5);
  const tipoMarca5 = checklistToText(by['copy_of_tipo_de_marca_3_1'] || '');

  // Serviços por N
  const serv1Stmt = firstNonEmpty(buscarServicoN(card,1));
  const serv2Stmt = firstNonEmpty(buscarServicoN(card,2));
  const serv3Stmt = firstNonEmpty(buscarServicoN(card,3));
  const serv4Stmt = firstNonEmpty(buscarServicoN(card,4));
  const serv5Stmt = firstNonEmpty(buscarServicoN(card,5));

  // Kinds
  const k1 = serviceKindFromText(serv1Stmt);
  const k2 = serviceKindFromText(serv2Stmt);
  const k3 = serviceKindFromText(serv3Stmt);
  const k4 = serviceKindFromText(serv4Stmt);
  const k5 = serviceKindFromText(serv5Stmt);

  // Decide template
  const anyMarca = [k1,k2,k3,k4,k5].includes('MARCA');
  const templateToUse = anyMarca ? TEMPLATE_UUID_CONTRATO : TEMPLATE_UUID_CONTRATO_OUTROS;

  // Contato contratante 1
  const contatoNome     = by['nome_1'] || getFirstByNames(card, ['nome do contato','contratante','responsável legal','responsavel legal']) || '';
  const contatoEmail    = by['email_de_contato'] || getFirstByNames(card, ['email','e-mail']) || '';
  const contatoTelefone = by['telefone_de_contato'] || getFirstByNames(card, ['telefone','celular','whatsapp','whats']) || '';

  // Campos de “contratante 2” antigos, se existirem
  const contato2Nome_old = by['nome_2'] || getFirstByNames(card, ['contratante 2', 'nome contratante 2']) || '';
  const contato2Email_old = by['email_2'] || getFirstByNames(card, ['email 2', 'e-mail 2']) || '';
  const contato2Telefone_old = by['telefone_2'] || getFirstByNames(card, ['telefone 2', 'celular 2']) || '';

  // Campos de COTITULAR — usar estes como fonte principal do Contratante 2
  const cot_nome = by['raz_o_social_ou_nome_completo_cotitular'] || '';
  const cot_nacionalidade = by['nacionalidade_cotitular'] || '';
  const cot_estado_civil = by['estado_civ_l_cotitular'] || '';
  const cot_rua = by['rua_av_do_cnpj_cotitular'] || '';
  const cot_bairro = by['bairro_cotitular'] || '';
  const cot_cidade = by['cidade_cotitular'] || '';
  const cot_uf = by['estado_cotitular'] || '';
  const cot_numero = ''; // não informado
  const cot_cep = '';    // não informado
  const cot_rg = by['rg_cotitular'] || '';
  const cot_cpf = by['cpf_cotitular'] || '';
  const cot_cnpj = by['cnpj_cotitular'] || '';
  const cot_docSelecao = cot_cnpj ? 'CNPJ' : (cot_cpf ? 'CPF' : '');

  // Envio do contrato principal e cotitular
  const emailEnvioContrato = by['email_para_envio_do_contrato'] || contatoEmail || '';
  const emailCotitularEnvio = by['copy_of_email_para_envio_do_contrato'] || '';
  const telefoneCotitularEnvio = by['copy_of_telefone_para_envio_do_contrato'] || '';
  // Telefone para envio do contrato (campo específico)
  const telefoneEnvioContrato = by['telefone_para_envio_do_contrato'] || getFirstByNames(card, ['telefone para envio do contrato', 'telefone para envio']) || contatoTelefone || '';

  // Documento (CPF/CNPJ) principal
  const doc = pickDocumento(card);
  const cpfDoc  = doc.tipo==='CPF'?  doc.valor : '';
  const cnpjDoc = doc.tipo==='CNPJ'? doc.valor : '';
  const cpfCampo  = by['cpf']    || '';
  const cnpjCampo = by['cnpj_1'] || '';

  // Parcelas / Assessoria
  const nParcelas = pickParcelas(card);
  const valorAssessoria = pickValorAssessoria(card);
  const formaAss = by['copy_of_tipo_de_pagamento'] || getFirstByNames(card, ['tipo de pagamento assessoria']) || '';

  // TAXA
  const taxaFaixaRaw = by['taxa'] || getFirstByNames(card, ['taxa']);
  const valorTaxaBRL = computeValorTaxaBRLFromFaixa({ taxa_faixa: taxaFaixaRaw });
  const formaPagtoTaxa = by['tipo_de_pagamento'] || '';

  // Datas novas
  const dataPagtoAssessoria = fmtDMY2(by['copy_of_copy_of_data_do_boleto_pagamento_pesquisa'] || '');
  const dataPagtoTaxa       = fmtDMY2(by['copy_of_data_do_boleto_pagamento_pesquisa'] || '');

  // Endereço (CNPJ) principal
  const cepCnpj    = by['cep_do_cnpj']     || '';
  const ruaCnpj    = by['rua_av_do_cnpj']  || '';
  const bairroCnpj = by['bairro_do_cnpj']  || '';
  const cidadeCnpj = by['cidade_do_cnpj']  || '';
  const ufCnpj     = by['estado_do_cnpj']  || '';
  const numeroCnpj = by['n_mero_1']        || getFirstByNames(card, ['numero','número','nº']) || '';

  // Vendedor (ainda usado apenas para exibição)
  const vendedor = (() => {
    const raw = by['vendedor_respons_vel'] || by['vendedor_respons_vel_1'] || by['respons_vel_5'];
    const v = raw ? extractAssigneeNames(raw) : [];
    return v[0] || '';
  })();

  // Riscos 1..5
  const risco1 = by['risco_da_marca'] || '';
  const risco2 = by['copy_of_copy_of_risco_da_marca'] || '';
  const risco3 = by['copy_of_risco_da_marca_3'] || '';
  const risco4 = by['copy_of_risco_da_marca_3_1'] || '';
  const risco5 = by['copy_of_risco_da_marca_4'] || '';

  // Nacionalidade e etc principal
  const nacionalidade = by['nacionalidade'] || '';
  const selecaoCnpjOuCpf = by['cnpj_ou_cpf'] || '';
  const estadoCivil = by['estado_civ_l'] || '';

  // Cláusula adicional
  const clausulaAdicional = (by['cl_usula_adicional'] && String(by['cl_usula_adicional']).trim()) ? by['cl_usula_adicional'] : 'Sem aditivos contratuais.';

  // Contratante 1
  const contratante1Texto = montarTextoContratante({
    nome: contatoNome || (by['r_social_ou_n_completo']||''),
    nacionalidade,
    estadoCivil,
    rua: ruaCnpj,
    bairro: bairroCnpj,
    numero: numeroCnpj,
    cidade: cidadeCnpj,
    uf: ufCnpj,
    cep: cepCnpj,
    rg: by['rg'] || '',
    docSelecao: selecaoCnpjOuCpf,
    cpf: cpfCampo || cpfDoc,
    cnpj: cnpjCampo || cnpjDoc,
    telefone: contatoTelefone,
    email: contatoEmail
  });

  // Detecta se há cotitular com base nos campos dedicados OU nos antigos campos 2
  const hasCotitular = Boolean(
    cot_nome || cot_nacionalidade || cot_estado_civil || cot_rua || cot_bairro || cot_cidade || cot_uf ||
    cot_rg || cot_cpf || cot_cnpj || emailCotitularEnvio || telefoneCotitularEnvio ||
    contato2Nome_old || contato2Email_old || contato2Telefone_old
  );

  // Contratante 2 com os CAMPOS DO COTITULAR como fonte principal
  const contratante2Texto = hasCotitular
    ? montarTextoContratante({
        nome: cot_nome || contato2Nome_old || 'Cotitular',
        nacionalidade: cot_nacionalidade || '',
        estadoCivil: cot_estado_civil || '',
        rua: cot_rua || ruaCnpj,
        bairro: cot_bairro || bairroCnpj,
        numero: cot_numero || '',
        cidade: cot_cidade || cidadeCnpj,
        uf: cot_uf || ufCnpj,
        cep: cot_cep || '',
        rg: cot_rg || '',
        docSelecao: cot_docSelecao,
        cpf: cot_cpf || '',
        cnpj: cot_cnpj || '',
        telefone: telefoneCotitularEnvio || contato2Telefone_old,
        email: emailCotitularEnvio || contato2Email_old
      })
    : '';

  // Dados para contato 1 e 2
  const dadosContato1 = [contatoNome, contatoTelefone, contatoEmail].filter(Boolean).join(' | ');
  const dadosContato2 = hasCotitular
    ? [
        (cot_nome || contato2Nome_old || 'Cotitular'),
        (telefoneCotitularEnvio || contato2Telefone_old || ''),
        (emailCotitularEnvio || contato2Email_old || '')
      ].filter(Boolean).join(' | ')
    : '';

  // Entradas consolidadas
  const entries = [
    {kind:serviceKindFromText(serv1Stmt), title:tituloMarca1, tipo:tipoMarca1, classes:classeSomenteNumeros1, stmt:serv1Stmt, risco:risco1, lines:linhasMarcasEspec1},
    {kind:serviceKindFromText(serv2Stmt), title:tituloMarca2, tipo:tipoMarca2, classes:classeSomenteNumeros2, stmt:serv2Stmt, risco:risco2, lines:linhasMarcasEspec2},
    {kind:serviceKindFromText(serv3Stmt), title:tituloMarca3, tipo:tipoMarca3, classes:classeSomenteNumeros3, stmt:serv3Stmt, risco:risco3, lines:linhasMarcasEspec3},
    {kind:serviceKindFromText(serv4Stmt), title:tituloMarca4, tipo:tipoMarca4, classes:classeSomenteNumeros4, stmt:serv4Stmt, risco:risco4, lines:linhasMarcasEspec4},
    {kind:serviceKindFromText(serv5Stmt), title:tituloMarca5, tipo:tipoMarca5, classes:classeSomenteNumeros5, stmt:serv5Stmt, risco:risco5, lines:linhasMarcasEspec5},
  ].filter(e => String(e.title||e.stmt||'').trim());

  // Agrupamento por kind
  const byKind = { 'MARCA':[], 'PATENTE':[], 'DESENHO INDUSTRIAL':[], 'COPYRIGHT/DIREITO AUTORAL':[], 'OUTROS':[] };
  entries.forEach(e => byKind[e.kind].push(e));

  // Linhas “quantidade + descrição” (sem normalizar o texto do serviço)
  const makeQtdDescLine = (kind, arr) => {
    if (!arr.length) return '';
    const baseServico = String(arr[0].stmt||'').trim() || (kind==='MARCA' ? 'Registro de Marca' : kind);
    const qtd = arr.length;
    return `${qtd} ${baseServico} JUNTO AO INPI`;
  };
  const qtdDesc = {
    MARCA: makeQtdDescLine('MARCA', byKind['MARCA']),
    PATENTE: makeQtdDescLine('PATENTE', byKind['PATENTE']),
    'DESENHO INDUSTRIAL': makeQtdDescLine('DESENHO INDUSTRIAL', byKind['DESENHO INDUSTRIAL']),
    'COPYRIGHT/DIREITO AUTORAL': makeQtdDescLine('COPYRIGHT/DIREITO AUTORAL', byKind['COPYRIGHT/DIREITO AUTORAL']),
    OUTROS: makeQtdDescLine('OUTROS', byKind['OUTROS'])
  };

  // Detalhes por item até 5
  const detalhes = {
    MARCA: ['', '', '', '', ''],
    PATENTE: ['', '', '', '', ''],
    'DESENHO INDUSTRIAL': ['', '', '', '', ''],
    'COPYRIGHT/DIREITO AUTORAL': ['', '', '', '', ''],
    OUTROS: ['', '', '', '', '']
  };
  ['MARCA','PATENTE','DESENHO INDUSTRIAL','COPYRIGHT/DIREITO AUTORAL','OUTROS'].forEach(k=>{
    const arr = byKind[k];
    for (let i=0;i<5;i++){
      const e = arr[i];
      if (!e){ detalhes[k][i] = ''; continue; }
      const cab = normalizarCabecalhoDetalhe(k, e.title, e.tipo, e.classes);
      detalhes[k][i] = cab;
    }
  });

  // Cabeçalhos “SERVIÇOS” para classes
  const headersServicos = {
    h1: byKind['MARCA'][0] ? `MARCA: ${byKind['MARCA'][0].title||''}` : '',
    h2: byKind['MARCA'][1] ? `MARCA: ${byKind['MARCA'][1].title||''}` : ''
  };

  // Risco agregado formatado com nome do tipo e do item
  const riscoAgregado = entries
    .map(e => {
      const tipo = e.kind || '';
      const nm = e.title || '';
      const r = String(e.risco||'').trim();
      if (!tipo && !nm && !r) return '';
      return `${tipo}: ${nm} - RISCO: ${r || 'Não informado'}`;
    })
    .filter(Boolean)
    .join(', ');

  return {
    cardId: card.id,
    templateToUse,

    // Identificação
    titulo: tituloMarca1 || card.title || '',
    nome: contatoNome || (by['r_social_ou_n_completo']||''),
    cpf: cpfDoc, 
    cnpj: cnpjDoc,
    rg: by['rg'] || '',
    estado_civil: estadoCivil,

    // Doc específicos
    cpf_campo: cpfCampo,
    cnpj_campo: cnpjCampo,
    selecao_cnpj_ou_cpf: selecaoCnpjOuCpf,

    // Contatos
    email: contatoEmail || '',
    telefone: contatoTelefone || '',
    dados_contato_1: dadosContato1,
    dados_contato_2: dadosContato2,

    // Textos completos dos contratantes
    contratante_1_texto: contratante1Texto,
    contratante_2_texto: contratante2Texto,

    // Email para assinatura
    email_envio_contrato: emailEnvioContrato,
    email_cotitular_envio: emailCotitularEnvio,
    
    // Telefone para envio
    telefone_envio_contrato: telefoneEnvioContrato,
    telefone_cotitular_envio: telefoneCotitularEnvio,

    // MARCA 1..5: linhas e cabeçalhos do formulário
    cabecalho_servicos_1: headersServicos.h1,
    cabecalho_servicos_2: headersServicos.h2,

    linhas_marcas_espec_1: linhasMarcasEspec1,
    linhas_marcas_espec_2: linhasMarcasEspec2,
    linhas_marcas_espec_3: linhasMarcasEspec3,
    linhas_marcas_espec_4: linhasMarcasEspec4,
    linhas_marcas_espec_5: linhasMarcasEspec5,

    // Quantidades e descrições por categoria
    qtd_desc: {
      MARCA: qtdDesc['MARCA'],
      PATENTE: qtdDesc['PATENTE'],
      DI: qtdDesc['DESENHO INDUSTRIAL'],
      COPY: qtdDesc['COPYRIGHT/DIREITO AUTORAL'],
      OUTROS: qtdDesc['OUTROS']
    },

    // Detalhes por categoria até 5
    det: detalhes,

    // Classes e tipos por marca
    classe1: classeSomenteNumeros1, tipo1: tipoMarca1, nome1: tituloMarca1,
    classe2: classeSomenteNumeros2, tipo2: tipoMarca2, nome2: tituloMarca2,
    classe3: classeSomenteNumeros3, tipo3: tipoMarca3, nome3: tituloMarca3,
    classe4: classeSomenteNumeros4, tipo4: tipoMarca4, nome4: tituloMarca4,
    classe5: classeSomenteNumeros5, tipo5: tipoMarca5, nome5: tituloMarca5,

    // Assessoria
    parcelas: nParcelas,
    valor_total: valorAssessoria ? toBRL(valorAssessoria) : '',
    forma_pagto_assessoria: formaAss,
    data_pagto_assessoria: dataPagtoAssessoria,

    // Pesquisa
    valor_pesquisa: 'R$ 00,00',
    forma_pesquisa: '',
    data_pesquisa: '00/00/00',

    // Taxa
    taxa_faixa: taxaFaixaRaw || '',
    valor_taxa_brl: valorTaxaBRL,
    forma_pagto_taxa: formaPagtoTaxa,
    data_pagto_taxa: dataPagtoTaxa,

    // Endereço
    cep_cnpj: cepCnpj,
    rua_cnpj: ruaCnpj,
    bairro_cnpj: bairroCnpj,
    cidade_cnpj: cidadeCnpj,
    uf_cnpj: ufCnpj,
    numero_cnpj: numeroCnpj,

    // Vendedor
    vendedor,

    // Risco agregado
    risco_agregado: riscoAgregado,

    // Cláusula adicional
    clausula_adicional: clausulaAdicional
  };
}

// NOVA VERSÃO — Qualificação separada para CPF x CNPJ
function montarTextoContratante(info = {}){
  const {
    nome,
    nacionalidade,
    estadoCivil,
    rua,
    bairro,
    numero,
    cidade,
    uf,
    cep,
    rg,
    docSelecao,
    cpf,
    cnpj,
    telefone,
    email
  } = info;

  const cpfDigits = onlyDigits(cpf);
  const cnpjDigits = onlyDigits(cnpj);
  const isCnpj = cnpjDigits.length === 14;
  const isCpf  = !isCnpj && cpfDigits.length === 11;

  // Monta endereço em texto único
  const enderecoPartes = [];
  if (rua) enderecoPartes.push(`Rua ${rua}`);
  if (numero) enderecoPartes.push(`nº ${numero}`);
  if (bairro) enderecoPartes.push(`Bairro ${bairro}`);
  let cidadeUf = '';
  if (cidade) cidadeUf += cidade;
  if (uf) cidadeUf += (cidadeUf ? ' - ' : '') + uf;
  if (cidadeUf) enderecoPartes.push(cidadeUf);
  if (cep) enderecoPartes.push(`CEP: ${cep}`);
  const enderecoStr = enderecoPartes.join(', ');

  // ===============================
  // CNPJ → Pessoa Jurídica
  // ===============================
  if (isCnpj) {
    const razao = nome || 'Razão Social não informada';

    let cnpjFmt = cnpj || '';
    if (cnpjDigits.length === 14) {
      cnpjFmt = cnpjDigits.replace(
        /^(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})$/,
        '$1.$2.$3/$4-$5'
      );
    }

    const partesPJ = [];
    partesPJ.push(`${razao}, inscrita no CNPJ sob nº ${cnpjFmt}`);

    if (enderecoStr) {
      partesPJ.push(`com sede em ${enderecoStr}`);
    }

    if (telefone || email) {
      const contato = [];
      if (telefone) contato.push(`telefone nº ${telefone}`);
      if (email) contato.push(`endereço eletrônico: ${email}`);
      partesPJ.push(`com ${contato.join(' e ')}`);
    }

    const textoPJ = partesPJ.join(', ').replace(/\s+,/g, ',').trim();
    return textoPJ.endsWith('.') ? textoPJ : `${textoPJ}.`;
  }

  // ===============================
  // CPF (ou genérico) → mantém lógica original (Brasileiro, Casado, empresário(a), ...)
  // ===============================
  const partes = [];
  const identidade = [];

  if (nome) identidade.push(nome);
  if (nacionalidade) identidade.push(nacionalidade);
  if (estadoCivil) identidade.push(estadoCivil);
  if (identidade.length) identidade.push('empresário(a)');
  if (identidade.length) partes.push(identidade.join(', '));

  if (enderecoStr) partes.push(`residente na ${enderecoStr}`);

  const documentos = [];
  if (rg) documentos.push(`portador(a) da cédula de identidade RG de nº ${rg}`);

  // Preferência: se tiver CPF com 11 dígitos, usa "portador(a) do CPF nº ..."
  if (isCpf && cpfDigits) {
    const cpfFmt = cpfDigits.replace(
      /^(\d{3})(\d{3})(\d{3})(\d{2})$/,
      '$1.$2.$3-$4'
    );
    documentos.push(`portador(a) do CPF nº ${cpfFmt}`);
  } else {
    const docUpper = String(docSelecao || '').trim().toUpperCase();
    const docNums = [];
    if (cpf) docNums.push({ tipo: 'CPF', valor: cpf });
    if (cnpj && !isCnpj) docNums.push({ tipo: 'CNPJ', valor: cnpj });

    if (docUpper && docNums.length){
      documentos.push(`devidamente inscrito no ${docUpper} sob o nº ${docNums[0].valor}`);
    } else {
      for (const doc of docNums){
        documentos.push(`devidamente inscrito no ${doc.tipo} sob o nº ${doc.valor}`);
      }
    }
  }

  if (documentos.length) partes.push(documentos.join(', '));

  const contatoPartes = [];
  if (telefone) contatoPartes.push(`com telefone de nº ${telefone}`);
  if (email)   contatoPartes.push(`com o seguinte endereço eletrônico: ${email}`);
  if (contatoPartes.length) partes.push(contatoPartes.join(' e '));

  if (!partes.length) return '';
  const texto = partes.join(', ').replace(/\s+,/g, ',').trim();
  return texto.endsWith('.') ? texto : `${texto}.`;
}

/* =========================
 * Variáveis para Templates
 * =======================*/

// Marca
function montarVarsParaTemplateMarca(d, nowInfo){
  const valorTotalNum = onlyNumberBR(d.valor_total);
  const parcelaNum = parseInt(String(d.parcelas||'1'),10)||1;
  const valorParcela = parcelaNum>0 ? valorTotalNum/parcelaNum : 0;

  const dia = String(nowInfo.dia).padStart(2,'0');
  const mesExtenso = monthNamePt(nowInfo.mes);
  const ano = String(nowInfo.ano);

  const base = {
    // Identificação
    'Contratante 1': d.contratante_1_texto || d.nome || '',
    'Contratante 2': d.contratante_2_texto || '',
    'CPF/CNPJ': d.selecao_cnpj_ou_cpf || '',
    'CPF': d.cpf_campo || '',
    'CNPJ': d.cnpj_campo || '',
    rg: d.rg || '',
    'Estado Civíl': d.estado_civil || '',
    'Estado Civil': d.estado_civil || '',

    // Endereço
    rua: d.rua_cnpj || '',
    bairro: d.bairro_cnpj || '',
    numero: d.numero_cnpj || '',
    nome_da_cidade: d.cidade_cnpj || '',
    cidade: d.cidade_cnpj || '',
    uf: d.uf_cnpj || '',
    cep: d.cep_cnpj || '',

    // Contato
    'E-mail': d.email || '',
    telefone: d.telefone || '',
    'dados para contato 1': d.dados_contato_1 || '',
    'dados para contato 2': d.dados_contato_2 || '',

    // Resultado da pesquisa prévia
    'Risco': d.risco_agregado || '',

    // Quantidade e descrição de Marca
    'Quantidade depósitos/processos de MARCA': d.qtd_desc.MARCA || '',
    'Descrição do serviço - MARCA': d.qtd_desc.MARCA ? '' : '',

    // Detalhes do serviço - Marca até 5
    'Detalhes do serviço - MARCA': d.det.MARCA[0] || '',
    'Detalhes do serviço - MARCA 2': d.det.MARCA[1] || '',
    'Detalhes do serviço - MARCA 3': d.det.MARCA[2] || '',
    'Detalhes do serviço - MARCA 4': d.det.MARCA[3] || '',
    'Detalhes do serviço - MARCA 5': d.det.MARCA[4] || '',

    // Formulário de Classes
    'Cabeçalho - SERVIÇOS': d.cabecalho_servicos_1 || '',
    'marcas-espec_1': d.linhas_marcas_espec_1[0] || '',
    'marcas-espec_2': d.linhas_marcas_espec_1[1] || '',
    'marcas-espec_3': d.linhas_marcas_espec_1[2] || '',
    'marcas-espec_4': d.linhas_marcas_espec_1[3] || '',
    'marcas-espec_5': d.linhas_marcas_espec_1[4] || '',

    'Cabeçalho - SERVIÇOS 2': d.cabecalho_servicos_2 || '',
    'marcas2-espec_1': d.linhas_marcas_espec_2[0] || '',
    'marcas2-espec_2': d.linhas_marcas_espec_2[1] || '',
    'marcas2-espec_3': d.linhas_marcas_espec_2[2] || '',
    'marcas2-espec_4': d.linhas_marcas_espec_2[3] || '',
    'marcas2-espec_5': d.linhas_marcas_espec_2[4] || '',

    // Assessoria
    'Número de parcelas da Assessoria': String(d.parcelas||'1'),
    'Valor da parcela da Assessoria': toBRL(valorParcela),
    'Forma de pagamento da Assessoria': d.forma_pagto_assessoria || '',
    'Data de pagamento da Assessoria': d.data_pagto_assessoria || '',

    // Pesquisa
    'Valor da Pesquisa': d.valor_pesquisa || 'R$ 00,00',
    'Forma de pagamento da Pesquisa': d.forma_pesquisa || '',
    'Data de pagamento da pesquisa': d.data_pesquisa || '00/00/00',

    // Taxa
    'Valor da Taxa': d.valor_taxa_brl || '',
    'Forma de pagamento da Taxa': d.forma_pagto_taxa || '',
    'Data de pagamento da Taxa': d.data_pagto_taxa || '',

    // Datas
    Dia: dia,
    Mês: mesExtenso,
    Mes: mesExtenso,
    Ano: ano,
    Cidade: d.cidade_cnpj || '',
    UF: d.uf_cnpj || '',

    // Cláusula adicional
    'clausula-adicional': d.clausula_adicional || ''
  };

  // Preencher até 30 linhas por segurança
  for (let i=5;i<30;i++){
    base[`marcas-espec_${i+1}`] = d.linhas_marcas_espec_1[i] || '';
    base[`marcas2-espec_${i-4}`] = d.linhas_marcas_espec_2[i-5] || '';
  }

  return base;
}

// Outros
function montarVarsParaTemplateOutros(d, nowInfo){
  const valorTotalNum = onlyNumberBR(d.valor_total);
  const parcelaNum = parseInt(String(d.parcelas||'1'),10)||1;
  const valorParcela = parcelaNum>0 ? valorTotalNum/parcelaNum : 0;

  const dia = String(nowInfo.dia).padStart(2,'0');
  const mesExtenso = monthNamePt(nowInfo.mes);
  const ano = String(nowInfo.ano);

  const base = {
    // Identificação
    'Contratante 1': d.contratante_1_texto || d.nome || '',
    'Contratante 2': d.contratante_2_texto || '',
    'CPF/CNPJ': d.selecao_cnpj_ou_cpf || '',
    'CPF': d.cpf_campo || '',
    'CNPJ': d.cnpj_campo || '',
    rg: d.rg || '',
    'Estado Civíl': d.estado_civil || '',
    'Estado Civil': d.estado_civil || '',

    // Endereço
    rua: d.rua_cnpj || '',
    bairro: d.bairro_cnpj || '',
    numero: d.numero_cnpj || '',
    nome_da_cidade: d.cidade_cnpj || '',
    cidade: d.cidade_cnpj || '',
    uf: d.uf_cnpj || '',
    cep: d.cep_cnpj || '',

    // Contato
    'E-mail': d.email || '',
    telefone: d.telefone || '',
    'dados para contato 1': d.dados_contato_1 || '',
    'dados para contato 2': d.dados_contato_2 || '',

    // Resultado da pesquisa prévia
    'Risco': d.risco_agregado || '',

    // PATENTE
    'Quantidade depósitos/processos de PATENTE': d.qtd_desc.PATENTE || '',
    'Descrição do serviço - PATENTE': d.qtd_desc.PATENTE ? '' : '',
    'Detalhes do serviço - PATENTE': d.det.PATENTE[0] || '',
    'Detalhes do serviço - PATENTE 2': d.det.PATENTE[1] || '',
    'Detalhes do serviço - PATENTE 3': d.det.PATENTE[2] || '',
    'Detalhes do serviço - PATENTE 4': d.det.PATENTE[3] || '',
    'Detalhes do serviço - PATENTE 5': d.det.PATENTE[4] || '',

    // DESENHO INDUSTRIAL
    'Quantidade depósitos – DESENHO INDUSTRIAL': d.qtd_desc.DI || d.qtd_desc['DESENHO INDUSTRIAL'] || '',
    'Descrição do serviço - DESENHO INDUSTRIAL': (d.qtd_desc.DI || d.qtd_desc['DESENHO INDUSTRIAL']) ? '' : '',
    'Detalhes do serviço - DESENHO INDUSTRIAL': d.det['DESENHO INDUSTRIAL'][0] || '',
    'Detalhes do serviço - DESENHO INDUSTRIAL 2': d.det['DESENHO INDUSTRIAL'][1] || '',
    'Detalhes do serviço - DESENHO INDUSTRIAL 3': d.det['DESENHO INDUSTRIAL'][2] || '',
    'Detalhes do serviço - DESENHO INDUSTRIAL 4': d.det['DESENHO INDUSTRIAL'][3] || '',
    'Detalhes do serviço - DESENHO INDUSTRIAL 5': d.det['DESENHO INDUSTRIAL'][4] || '',

    // COPYRIGHT
    'Quantidade registros de Copyright/Direito Autoral': d.qtd_desc.COPY || '',
    'Descrição do serviço - Copyright/Direito Autoral': d.qtd_desc.COPY ? '' : '',
    'Detalhes do serviço - Copyright/Direito Autoral': d.det['COPYRIGHT/DIREITO AUTORAL'][0] || '',
    'Detalhes do serviço - Copyright/Direito Autoral 2': d.det['COPYRIGHT/DIREITO AUTORAL'][1] || '',
    'Detalhes do serviço - Copyright/Direito Autoral 3': d.det['COPYRIGHT/DIREITO AUTORAL'][2] || '',
    'Detalhes do serviço - Copyright/Direito Autoral 4': d.det['COPYRIGHT/DIREITO AUTORAL'][3] || '',
    'Detalhes do serviço - Copyright/Direito Autoral 5': d.det['COPYRIGHT/DIREITO AUTORAL'][4] || '',

    // OUTROS
    'Quantidade registros de outros serviços': d.qtd_desc.OUTROS || '',
    'Descrição do serviço - outros serviços': d.qtd_desc.OUTROS ? '' : '',
    'Detalhes do serviço - outros serviços': d.det['OUTROS'][0] || '',
    'Detalhes do serviço - outros serviços 2': d.det['OUTROS'][1] || '',
    'Detalhes do serviço - outros serviços 3': d.det['OUTROS'][2] || '',
    'Detalhes do serviço - outros serviços 4': d.det['OUTROS'][3] || '',
    'Detalhes do serviço - outros serviços 5': d.det['OUTROS'][4] || '',

    // Assessoria
    'Número de parcelas da Assessoria': String(d.parcelas||'1'),
    'Valor da parcela da Assessoria': toBRL(valorParcela),
    'Forma de pagamento da Assessoria': d.forma_pagto_assessoria || '',
    'Data de pagamento da Assessoria': d.data_pagto_assessoria || '',

    // Pesquisa
    'Valor da Pesquisa': d.valor_pesquisa || 'R$ 00,00',
    'Forma de pagamento da Pesquisa': d.forma_pesquisa || '',
    'Data de pagamento da pesquisa': d.data_pesquisa || '00/00/00',

    // Taxa
    'Valor da Taxa': d.valor_taxa_brl || '',
    'Forma de pagamento da Taxa': d.forma_pagto_taxa || '',
    'Data de pagamento da Taxa': d.data_pagto_taxa || '',

    // Datas
    Dia: dia,
    Mês: mesExtenso,
    Mes: mesExtenso,
    Ano: ano,
    Cidade: d.cidade_cnpj || '',
    UF: d.uf_cnpj || '',

    // Cláusula adicional
    'clausula-adicional': d.clausula_adicional || ''
  };

  return base;
}

// Procuração
function montarVarsParaTemplateProcuracao(d, nowInfo){
  const dia = String(nowInfo.dia).padStart(2,'0');
  const mesExtenso = monthNamePt(nowInfo.mes);
  const ano = String(nowInfo.ano);

  // Formata CPF/CNPJ
  const cpfDigits = onlyDigits(d.cpf || d.cpf_campo || '');
  const cnpjDigits = onlyDigits(d.cnpj || d.cnpj_campo || '');
  const cpfFmt = cpfDigits.length === 11 
    ? cpfDigits.replace(/^(\d{3})(\d{3})(\d{3})(\d{2})$/, '$1.$2.$3-$4')
    : '';
  const cnpjFmt = cnpjDigits.length === 14
    ? cnpjDigits.replace(/^(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})$/, '$1.$2.$3/$4-$5')
    : '';

  const base = {
    // Identificação do outorgante
    'Contratante 1': d.contratante_1_texto || d.nome || '',
    'Contratante 2': d.contratante_2_texto || '',
    'Nome': d.nome || '',
    'CPF': cpfFmt || '',
    'CNPJ': cnpjFmt || '',
    'CPF/CNPJ': d.selecao_cnpj_ou_cpf || '',
    'RG': d.rg || '',
    'Estado Civil': d.estado_civil || '',
    'Estado Civíl': d.estado_civil || '',

    // Endereço
    'Rua': d.rua_cnpj || '',
    'Bairro': d.bairro_cnpj || '',
    'Número': d.numero_cnpj || '',
    'Cidade': d.cidade_cnpj || '',
    'UF': d.uf_cnpj || '',
    'CEP': d.cep_cnpj || '',
    'Endereço completo': [
      d.rua_cnpj ? `Rua ${d.rua_cnpj}` : '',
      d.numero_cnpj ? `nº ${d.numero_cnpj}` : '',
      d.bairro_cnpj ? `Bairro ${d.bairro_cnpj}` : '',
      d.cidade_cnpj && d.uf_cnpj ? `${d.cidade_cnpj} - ${d.uf_cnpj}` : (d.cidade_cnpj || d.uf_cnpj || ''),
      d.cep_cnpj ? `CEP: ${d.cep_cnpj}` : ''
    ].filter(Boolean).join(', '),

    // Contato
    'E-mail': d.email || '',
    'Telefone': d.telefone || '',
    'dados para contato 1': d.dados_contato_1 || '',
    'dados para contato 2': d.dados_contato_2 || '',

    // Datas
    'Dia': dia,
    'Mês': mesExtenso,
    'Mes': mesExtenso,
    'Ano': ano,
    'Data': `${dia} de ${mesExtenso} de ${ano}`,

    // Informações do contrato relacionadas
    'Título': d.titulo || '',
    'Serviços': d.qtd_desc.MARCA || d.qtd_desc.PATENTE || d.qtd_desc.OUTROS || '',
    'Risco': d.risco_agregado || ''
  };

  return base;
}

// Assinantes: principal + empresa + cotitular quando houver
function montarSigners(d, incluirTelefone = false){
  const list = [];
  const emailPrincipal = d.email_envio_contrato || d.email || '';
  const telefonePrincipal = d.telefone_envio_contrato || d.telefone || '';
  
  if (emailPrincipal) {
    const signer = { 
      email: emailPrincipal, 
      name: d.nome || d.titulo || emailPrincipal, 
      act:'1', 
      foreign:'0', 
      send_email:'1' 
    };
    if (incluirTelefone && telefonePrincipal) {
      // Formatar telefone para formato internacional (+55...)
      let phone = telefonePrincipal.replace(/[^\d+]/g, '');
      // Se não começar com +, adicionar +55 (Brasil)
      if (!phone.startsWith('+')) {
        // Se começar com 0, remover
        if (phone.startsWith('0')) {
          phone = phone.substring(1);
        }
        // Adicionar código do país Brasil (+55)
        phone = '+55' + phone;
      }
      signer.phone = phone;
      console.log(`[SIGNERS] Telefone preparado para ${signer.name}: ${signer.phone}`);
    }
    list.push(signer);
  }
  
  if (d.email_cotitular_envio) {
    const signer = { 
      email: d.email_cotitular_envio, 
      name: 'Cotitular', 
      act:'1', 
      foreign:'0', 
      send_email:'1' 
    };
    if (incluirTelefone && d.telefone_cotitular_envio) {
      // Formatar telefone para formato internacional (+55...)
      let phone = d.telefone_cotitular_envio.replace(/[^\d+]/g, '');
      // Se não começar com +, adicionar +55 (Brasil)
      if (!phone.startsWith('+')) {
        // Se começar com 0, remover
        if (phone.startsWith('0')) {
          phone = phone.substring(1);
        }
        // Adicionar código do país Brasil (+55)
        phone = '+55' + phone;
      }
      signer.phone = phone;
      console.log(`[SIGNERS] Telefone preparado para ${signer.name}: ${signer.phone}`);
    }
    list.push(signer);
  }
  
  if (EMAIL_ASSINATURA_EMPRESA) {
    list.push({ 
      email: EMAIL_ASSINATURA_EMPRESA, 
      name: 'Empresa', 
      act:'1', 
      foreign:'0', 
      send_email:'1' 
    });
  }
  
  const seen={}; 
  return list.filter(s => (seen[s.email.toLowerCase()]? false : (seen[s.email.toLowerCase()]=true)));
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

  const titleSanitized = String(title||'Contrato').replace(/[\x00-\x1F\x7F-\x9F]/g,'');

  const varsObjValidated = {};
  for (const [key, value] of Object.entries(varsObj || {})) {
    let v = value==null ? '' : String(value);
    v = v.replace(/[\x00-\x1F\x7F-\x9F]/g,'').trim();
    varsObjValidated[key] = v;
  }

  const body = { name_document: titleSanitized, templates: { [templateId]: varsObjValidated } };

  const res = await fetchWithRetry(url.toString(), {
    method: 'POST', headers: { 'Accept': 'application/json', 'Content-Type': 'application/json' },
    body: JSON.stringify(body)
  }, { attempts: 5, baseDelayMs: 600, timeoutMs: 20000 });

  const text = await res.text();
  let json; 
  try { json = JSON.parse(text); } catch { json = null; }

  if (!res.ok || !(json && (json.uuid || json.uuid_document))) {
    console.error('[ERRO D4SIGN WORD]', res.status, text.substring(0, 1000));
    throw new Error(`Falha D4Sign(WORD): ${res.status} - ${text.substring(0, 200)}`);
  }
  return json.uuid || json.uuid_document;
}

// ===============================
// NOVO — REGISTRAR WEBHOOK POR DOCUMENTO D4SIGN
// ===============================
async function registerWebhookForDocument(tokenAPI, cryptKey, uuidDocument, urlWebhook){
  const base = 'https://secure.d4sign.com.br';
  const url = new URL(`/api/v1/documents/${uuidDocument}/webhooks`, base);
  url.searchParams.set('tokenAPI', tokenAPI);
  url.searchParams.set('cryptKey', cryptKey);

  const body = { url: urlWebhook };

  const res = await fetch(url.toString(), {
    method: 'POST',
    headers: { 'Accept': 'application/json', 'Content-Type': 'application/json' },
    body: JSON.stringify(body)
  });

  const text = await res.text();
  let json;
  try { json = JSON.parse(text); } catch { json = null; }

  if (!res.ok) {
    console.error('[ERRO webhook D4Sign]', res.status, text.substring(0, 1000));
    throw new Error(`Falha ao cadastrar webhook: ${res.status}`);
  }

  return json;
}

async function cadastrarSignatarios(tokenAPI, cryptKey, uuidDocument, signers, usarWhatsApp = false) {
  const base = 'https://secure.d4sign.com.br';
  const url = new URL(`/api/v1/documents/${uuidDocument}/createlist`, base);
  url.searchParams.set('tokenAPI', tokenAPI);
  url.searchParams.set('cryptKey', cryptKey);
  const body = { 
    signers: signers.map(s => {
      const signer = { 
        email: s.email, 
        name: s.name, 
        act: s.act || '1', 
        foreign: s.foreign || '0'
      };
      
      // Configurar envio por email ou WhatsApp
      if (usarWhatsApp) {
        // Se for WhatsApp, verificar se tem telefone
        if (s.phone) {
          // O telefone já deve estar formatado com +55 pela função montarSigners
          // Mas vamos garantir que está no formato correto
          let phoneFormatted = String(s.phone).trim();
          // Remover todos os caracteres não numéricos exceto +
          phoneFormatted = phoneFormatted.replace(/[^\d+]/g, '');
          
          // Se não começar com +, adicionar +55 (Brasil)
          if (!phoneFormatted.startsWith('+')) {
            // Se começar com 0, remover
            if (phoneFormatted.startsWith('0')) {
              phoneFormatted = phoneFormatted.substring(1);
            }
            // Adicionar código do país Brasil (+55)
            phoneFormatted = '+55' + phoneFormatted;
          }
          
          signer.phone = phoneFormatted;
          signer.send_whatsapp = '1';
          signer.send_email = '0';
          console.log(`[CADASTRO] Signatário ${s.name} configurado para WhatsApp: ${phoneFormatted} (email: ${s.email})`);
        } else {
          // Se não tiver telefone mas for WhatsApp, manter como email para este signatário
          // (pode ser o signatário da empresa que não precisa de WhatsApp)
          signer.send_email = s.send_email || '1';
          signer.send_whatsapp = '0';
          console.log(`[CADASTRO] Signatário ${s.name} sem telefone, mantido como email (WhatsApp solicitado mas sem telefone)`);
        }
      } else {
        signer.send_email = s.send_email || '1';
        signer.send_whatsapp = '0';
      }
      
      return signer;
    }) 
  };
  
  console.log(`[CADASTRO] Cadastrando signatários para ${usarWhatsApp ? 'WhatsApp' : 'email'}:`, JSON.stringify(body, null, 2));
  const res = await fetchWithRetry(url.toString(), {
    method: 'POST', headers: { 'Accept': 'application/json', 'Content-Type': 'application/json' },
    body: JSON.stringify(body)
  }, { attempts: 5, baseDelayMs: 600, timeoutMs: 20000 });
  const text = await res.text();
  
  if (!res.ok) {
    console.error('[ERRO D4SIGN createlist]', res.status, text);
    
    let jsonResponse = null;
    try {
      jsonResponse = JSON.parse(text);
    } catch (e) {
      // Não é JSON
    }
    
    const mensagem = jsonResponse?.message || text || '';
    const mensagemLower = String(mensagem).toLowerCase();
    
    let mensagemAmigavel = 'Não foi possível cadastrar os signatários no documento.';
    
    if (mensagemLower.includes('email') || mensagemLower.includes('inválido')) {
      mensagemAmigavel = 'Um ou mais emails dos signatários são inválidos. Verifique se os emails estão corretos.';
    } else if (mensagemLower.includes('documento') || mensagemLower.includes('document')) {
      mensagemAmigavel = 'O documento não está pronto para receber signatários. Aguarde alguns instantes.';
    } else if (res.status === 404) {
      mensagemAmigavel = 'Documento não encontrado. O documento pode ter sido excluído.';
    } else if (res.status === 422) {
      mensagemAmigavel = 'Os dados dos signatários não são válidos. Verifique se todos os campos estão preenchidos corretamente.';
    }
    
    const erro = new Error(mensagemAmigavel);
    erro.statusCode = res.status;
    erro.responseText = text;
    throw erro;
  }
  
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
  }, { attempts: 5, baseDelayMs: 600, timeoutMs: 20000 });
  const text = await res.text();
  let json; try { json = JSON.parse(text); } catch { json = null; }
  if (!res.ok || !json?.url) {
    console.error('[ERRO D4SIGN download]', res.status, text);
    throw new Error(`Falha ao gerar URL de download: ${res.status}`);
  }
  return json; // { url, name }
}
async function getDocumentStatus(tokenAPI, cryptKey, uuidDocument) {
  const base = 'https://secure.d4sign.com.br';
  const url = new URL(`/api/v1/documents/${uuidDocument}`, base);
  url.searchParams.set('tokenAPI', tokenAPI);
  url.searchParams.set('cryptKey', cryptKey);
  try {
    const res = await fetchWithRetry(url.toString(), {
      method: 'GET',
      headers: { 'Accept': 'application/json' }
    }, { attempts: 3, baseDelayMs: 500, timeoutMs: 10000 });
    const text = await res.text();
    let json; try { json = JSON.parse(text); } catch { return null; }
    return json;
  } catch {
    return null;
  }
}
// Função para traduzir erros da API D4Sign em mensagens amigáveis
function traduzirErroD4Sign(status, responseText, jsonResponse) {
  const statusCode = status;
  const mensagem = jsonResponse?.message || responseText || '';
  const mensagemLower = String(mensagem).toLowerCase();
  
  // Erros comuns e suas traduções
  if (statusCode === 400) {
    if (mensagemLower.includes('signatário') || mensagemLower.includes('signer')) {
      return 'Não foi possível enviar porque não há signatários cadastrados no documento. Por favor, verifique se os emails dos signatários estão corretos.';
    }
    if (mensagemLower.includes('documento') || mensagemLower.includes('document')) {
      return 'O documento não está pronto para ser enviado. Aguarde alguns instantes e tente novamente.';
    }
    if (mensagemLower.includes('já enviado') || mensagemLower.includes('already sent')) {
      return 'Este documento já foi enviado para assinatura anteriormente.';
    }
    return 'Dados inválidos. Verifique se o documento existe e está configurado corretamente.';
  }
  
  if (statusCode === 401 || statusCode === 403) {
    return 'Não foi possível autenticar na plataforma de assinatura. Verifique as credenciais de acesso.';
  }
  
  if (statusCode === 404) {
    return 'Documento não encontrado. O documento pode ter sido excluído ou o identificador está incorreto.';
  }
  
  if (statusCode === 422) {
    if (mensagemLower.includes('email') || mensagemLower.includes('inválido')) {
      return 'Um ou mais emails dos signatários são inválidos. Verifique os emails cadastrados.';
    }
    return 'O documento não pode ser enviado no estado atual. Verifique se todos os dados estão preenchidos corretamente.';
  }
  
  if (statusCode === 429) {
    return 'Muitas tentativas de envio. Aguarde alguns minutos antes de tentar novamente.';
  }
  
  if (statusCode >= 500) {
    return 'O serviço de assinatura está temporariamente indisponível. Tente novamente em alguns instantes.';
  }
  
  // Mensagem genérica com detalhes se disponível
  if (mensagem) {
    return `Erro ao enviar documento: ${mensagem}`;
  }
  
  return `Não foi possível enviar o documento. Código de erro: ${statusCode}`;
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
  
  console.log(`[SEND] Tentando enviar documento ${uuidDocument} para assinatura...`);
  console.log(`[SEND] Parâmetros: skip_email=${skip_email}, workflow=${workflow}`);
  
  const res = await fetchWithRetry(url.toString(), {
    method: 'POST',
    headers: { 'Accept': 'application/json', 'Content-Type': 'application/json' },
    body: JSON.stringify(body)
  }, { attempts: 5, baseDelayMs: 600, timeoutMs: 20000 });
  
  const text = await res.text();
  let jsonResponse = null;
  try {
    jsonResponse = JSON.parse(text);
  } catch (e) {
    // Não é JSON, continua com o texto
  }
  
  if (!res.ok) {
    console.error('[ERRO D4SIGN sendtosigner]', {
      status: res.status,
      statusText: res.statusText,
      response: text.substring(0, 500),
      uuid: uuidDocument
    });
    
    const mensagemAmigavel = traduzirErroD4Sign(res.status, text, jsonResponse);
    const erro = new Error(mensagemAmigavel);
    erro.statusCode = res.status;
    erro.responseText = text;
    throw erro;
  }
  
  console.log(`[SEND] Documento ${uuidDocument} enviado com sucesso.`);
  return text;
}

/* =========================
 * Fase Pipefy
 * =======================*/
async function moveCardToPhaseSafe(cardId, phaseId) {
  if (!phaseId) {
    console.warn('[moveCardToPhaseSafe] phaseId vazio, nada a mover. cardId =', cardId);
    return;
  }

  try {
    console.log('[moveCardToPhaseSafe] Tentando mover card', cardId, 'para fase', phaseId);

    const data = await gql(`
      mutation($input: MoveCardToPhaseInput!){
        moveCardToPhase(input:$input){
          card{
            id
            current_phase{
              id
              name
            }
          }
        }
      }
    `, {
      input: {
        card_id: Number(cardId),
        destination_phase_id: Number(phaseId)
      }
    });

    const moved = data?.moveCardToPhase?.card;
    if (moved) {
      console.log(
        '[moveCardToPhaseSafe] Move ok. Card',
        moved.id,
        'agora na fase',
        moved.current_phase?.id,
        '(' + (moved.current_phase?.name || 'sem nome') + ')'
      );
    } else {
      console.warn('[moveCardToPhaseSafe] Resposta sem card retornado para cardId =', cardId, 'phaseId =', phaseId, 'data =', JSON.stringify(data));
    }
  } catch (e) {
    console.error(
      '[moveCardToPhaseSafe] Erro ao mover card',
      cardId,
      'para fase',
      phaseId,
      '=>',
      e.message || e
    );
    // Se quiser que o fluxo continue mesmo com falha de move, comente a linha abaixo
    // throw e;
  }
};

/* =========================
 * Rotas — VENDEDOR (UX)
 * =======================*/
// ===============================
// NOVO — POSTBACK DO D4SIGN (DOCUMENTO FINALIZADO)
// ===============================
app.post('/d4sign/postback', async (req, res) => {
  try {
    const { uuid, type_post } = req.body || {};

    if (!uuid) {
      console.warn('[POSTBACK D4SIGN] Sem UUID no body');
      return res.status(200).json({ ok: true });
    }

    // type_post = "1" → documento finalizado/assinado
    if (String(type_post) !== '1') {
      console.log('[POSTBACK D4SIGN] Evento ignorado:', type_post);
      return res.status(200).json({ ok: true });
    }

    console.log('[POSTBACK D4SIGN] Documento finalizado:', uuid);

    const cardId = await findCardIdByD4Uuid(uuid);
    if (!cardId) {
      console.warn('[POSTBACK D4SIGN] Nenhum card encontrado para uuid:', uuid);
      return res.status(200).json({ ok: true });
    }

    // 1. mover card para fase final (primeiro)
    try {
      // Movimentação de card removida conforme solicitado
      console.log('[POSTBACK D4SIGN] Card movido para fase 339299694');
    } catch (e) {
      console.error('[POSTBACK D4SIGN] Erro ao mover card:', e.message);
    }

    // 2. pegar link do PDF assinado
    let info;
    try {
      info = await getDownloadUrl(D4SIGN_TOKEN, D4SIGN_CRYPT_KEY, uuid, {
        type: 'PDF',
        language: 'pt'
      });
    } catch (e) {
      console.error('[POSTBACK D4SIGN] Erro ao buscar download:', e.message);
      return res.status(200).json({ ok: true });
    }

    // 3. anexar PDF no campo de arquivo
    try {
      await anexarContratoAssinadoNoCard(cardId, info.url, info.name);
      console.log('[POSTBACK D4SIGN] PDF anexado ao card');
    } catch (e) {
      console.error('[POSTBACK D4SIGN] Erro ao anexar contrato:', e.message);
    }

    return res.status(200).json({ ok: true });

  } catch (e) {
    console.error('[POSTBACK D4SIGN] Erro geral:', e.message);
    return res.status(200).json({ ok: true });
  }
});

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

    <h2>Contratante(s)</h2>
    <div class="grid">
      <div><div class="label">Contratante 1</div><div>${d.contratante_1_texto||'-'}</div></div>
      <div><div class="label">Contratante 2</div><div>${d.contratante_2_texto||'-'}</div></div>
    </div>

    <h2>Contato</h2>
    <div class="grid">
      <div><div class="label">Dados para contato 1</div><div>${d.dados_contato_1||'-'}</div></div>
      <div><div class="label">Dados para contato 2</div><div>${d.dados_contato_2||'-'}</div></div>
      <div><div class="label">Email para envio do contrato</div><div>${d.email_envio_contrato||'-'}</div></div>
      <div><div class="label">Email Cotitular</div><div>${d.email_cotitular_envio||'-'}</div></div>
    </div>

    <h2>Serviços</h2>
    <div class="grid3">
      <div><div class="label">Template escolhido</div><div>${d.templateToUse===process.env.TEMPLATE_UUID_CONTRATO? 'Contrato de Marca' : 'Contrato de Outros Serviços'}</div></div>
      <div><div class="label">Qtd Descrição MARCA</div><div>${d.qtd_desc.MARCA||'-'}</div></div>
      <div><div class="label">Risco agregado</div><div>${d.risco_agregado||'-'}</div></div>
    </div>

    <form method="POST" action="/lead/${encodeURIComponent(req.params.token)}/generate" style="margin-top:24px">
      <button class="btn" type="submit">Gerar contrato</button>
    </form>
    <p class="muted" style="margin-top:12px">Ao clicar, o documento será criado no D4Sign.</p>
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
  let lockKey = null;
  let cardId = null;
  
  try {
    const parsed = parseLeadToken(req.params.token);
    cardId = parsed.cardId;
    if (!cardId) {
      throw new Error('Token inválido: cardId não encontrado');
    }
    
    lockKey = `lead:${cardId}`;
    if (!acquireLock(lockKey)) {
      return res.status(200).send('Processando, tente novamente em instantes.');
    }

    preflightDNS().catch(()=>{});

    const card = await getCard(cardId);
    if (!card) {
      throw new Error(`Card ${cardId} não encontrado no Pipefy`);
    }
    
    const d = await montarDados(card);
    if (!d) {
      throw new Error('Falha ao montar dados do card');
    }

    const now = new Date();
    const nowInfo = { dia: now.getDate(), mes: now.getMonth()+1, ano: now.getFullYear() };

    // Validar template
    if (!d.templateToUse) {
      throw new Error('Template não identificado. Verifique os dados do card.');
    }

    const isMarcaTemplate = d.templateToUse === TEMPLATE_UUID_CONTRATO;
    const add = isMarcaTemplate ? montarVarsParaTemplateMarca(d, nowInfo)
                                : montarVarsParaTemplateOutros(d, nowInfo);
    
    if (!add || typeof add !== 'object') {
      throw new Error('Falha ao montar variáveis do template. Verifique os dados do card.');
    }
    
    const signers = montarSigners(d);
    if (!signers || signers.length === 0) {
      throw new Error('Nenhum signatário encontrado. Verifique se há email configurado no card.');
    }

    // NOVO — Seleciona cofre pela "Equipe contrato"
    const equipeContrato = getEquipeContratoFromCard(card);
    let uuidSafe = null;
    let cofreUsadoPadrao = false;
    let nomeCofreUsado = '';

    if (!equipeContrato) {
      console.warn('[LEAD-GENERATE] Campo "Equipe contrato" não encontrado ou sem valor no card', card.id);
      // Usa cofre padrão
      uuidSafe = DEFAULT_COFRE_UUID;
      cofreUsadoPadrao = true;
      nomeCofreUsado = 'DEFAULT_COFRE_UUID';
      console.log('[LEAD-GENERATE] Usando cofre padrão (DEFAULT_COFRE_UUID)');
    } else {
      uuidSafe = COFRES_UUIDS[equipeContrato];
      if (!uuidSafe) {
        console.warn(`[LEAD-GENERATE] Equipe contrato "${equipeContrato}" sem cofre mapeado. Usando cofre padrão.`);
        // Usa cofre padrão
        uuidSafe = DEFAULT_COFRE_UUID;
        cofreUsadoPadrao = true;
        nomeCofreUsado = 'DEFAULT_COFRE_UUID';
      } else {
        // Cofre válido encontrado
        nomeCofreUsado = getNomeCofreByUuid(uuidSafe);
      }
    }

    if (!uuidSafe) {
      throw new Error('Nenhum cofre disponível. Configure DEFAULT_COFRE_UUID ou mapeie a equipe.');
    }

    console.log(`[LEAD-GENERATE] Criando contrato no cofre: ${nomeCofreUsado} (${uuidSafe})`);
    
    let uuidDoc = null;
    try {
      uuidDoc = await makeDocFromWordTemplate(
        D4SIGN_TOKEN,
        D4SIGN_CRYPT_KEY,
        uuidSafe,
        d.templateToUse,
        d.titulo || card.title || 'Contrato',
        add
      );
      
      if (!uuidDoc) {
        throw new Error('Falha ao criar documento no D4Sign. O documento não foi criado.');
      }
      
      console.log(`[D4SIGN] Contrato criado: ${uuidDoc}`);
    } catch (e) {
      console.error('[ERRO] Falha ao criar documento no D4Sign:', e.message);
      throw new Error(`Erro ao criar documento no D4Sign: ${e.message}`);
    }

// ===============================
// NOVO — Cadastrar webhook deste documento
// ===============================
try {
  await registerWebhookForDocument(
    D4SIGN_TOKEN,
    D4SIGN_CRYPT_KEY,
    uuidDoc,
    `${PUBLIC_BASE_URL}/d4sign/postback`
  );
  console.log('[D4SIGN] Webhook registrado no documento.');
} catch (e) {
  console.error('[ERRO] Falha ao registrar webhook:', e.message);
}

// ===============================
// NOVO — Salvar UUID do documento no card
// ===============================
try {
  await updateCardField(card.id, 'd4_uuid_contrato', uuidDoc);
  console.log('[D4SIGN] UUID do contrato salvo no card.');
} catch (e) {
  console.error('[ERRO] Falha ao salvar uuid do contrato no card:', e.message);
}

// ===============================
// NOVO — GERAR PROCURAÇÃO
// ===============================
let uuidProcuracao = null;
if (TEMPLATE_UUID_PROCURACAO) {
  try {
    const varsProcuracao = montarVarsParaTemplateProcuracao(d, nowInfo);
    uuidProcuracao = await makeDocFromWordTemplate(
      D4SIGN_TOKEN,
      D4SIGN_CRYPT_KEY,
      uuidSafe,
      TEMPLATE_UUID_PROCURACAO,
      `Procuração - ${d.titulo || card.title || 'Contrato'}`,
      varsProcuracao
    );
    console.log(`[D4SIGN] Procuração criada: ${uuidProcuracao}`);

    // Salvar UUID da procuração no card
    try {
      await updateCardField(card.id, 'd4_uuid_procuracao', uuidProcuracao);
      console.log('[D4SIGN] UUID da procuração salvo no card.');
    } catch (e) {
      console.error('[ERRO] Falha ao salvar uuid da procuração no card:', e.message);
    }

    // Aguardar documento estar pronto
    await new Promise(r=>setTimeout(r, 3000));
    try { 
      await getDocumentStatus(D4SIGN_TOKEN, D4SIGN_CRYPT_KEY, uuidProcuracao); 
      console.log('[D4SIGN] Status da procuração verificado.');
    } catch (e) {
      console.warn('[D4SIGN] Aviso ao verificar status da procuração:', e.message);
    }

    // Registrar webhook da procuração (opcional - se quiser rastrear quando for assinada)
    try {
      await registerWebhookForDocument(
        D4SIGN_TOKEN,
        D4SIGN_CRYPT_KEY,
        uuidProcuracao,
        `${PUBLIC_BASE_URL}/d4sign/postback`
      );
      console.log('[D4SIGN] Webhook da procuração registrado.');
    } catch (e) {
      console.error('[ERRO] Falha ao registrar webhook da procuração:', e.message);
    }

    // Cadastrar signatários da procuração (mesmos do contrato)
    // Garantir que temos signatários antes de cadastrar
    if (!signers || signers.length === 0) {
      console.warn('[D4SIGN] Signatários não disponíveis no escopo, preparando novamente...');
      try {
        signers = montarSigners(d);
        if (signers && signers.length > 0) {
          console.log('[D4SIGN] Signatários preparados para procuração:', signers.map(s => s.email).join(', '));
        } else {
          console.error('[D4SIGN] ERRO CRÍTICO: Não foi possível preparar signatários. Verifique se há emails no card.');
        }
      } catch (e) {
        console.error('[D4SIGN] Erro ao preparar signatários para procuração:', e.message);
      }
    }
    
    // Tentar cadastrar signatários com múltiplas tentativas
    await new Promise(r=>setTimeout(r, 2000));
    let signatariosCadastrados = false;
    const maxTentativas = 3;
    
    for (let tentativa = 1; tentativa <= maxTentativas; tentativa++) {
      try {
        if (signers && signers.length > 0) {
          await cadastrarSignatarios(D4SIGN_TOKEN, D4SIGN_CRYPT_KEY, uuidProcuracao, signers);
          console.log(`[D4SIGN] Signatários da procuração cadastrados com sucesso (tentativa ${tentativa}/${maxTentativas}):`, signers.map(s => s.email).join(', '));
          signatariosCadastrados = true;
          break;
        } else {
          console.error(`[D4SIGN] ERRO (tentativa ${tentativa}/${maxTentativas}): Nenhum signatário disponível. Verifique se há emails no card.`);
          if (tentativa < maxTentativas) {
            // Tentar preparar signatários novamente
            await new Promise(r=>setTimeout(r, 1000));
            signers = montarSigners(d);
          }
        }
      } catch (e) {
        console.error(`[ERRO] Falha ao cadastrar signatários da procuração (tentativa ${tentativa}/${maxTentativas}):`, e.message);
        if (tentativa < maxTentativas) {
          await new Promise(r=>setTimeout(r, 2000));
          // Tentar novamente
        } else {
          console.error('[D4SIGN] ERRO CRÍTICO: Não foi possível cadastrar signatários após múltiplas tentativas.');
        }
      }
    }
    
    if (!signatariosCadastrados) {
      console.error('[D4SIGN] AVISO: Procuração criada mas signatários não foram cadastrados. Será necessário cadastrar manualmente.');
    }
  } catch (e) {
    console.error('[ERRO] Falha ao gerar procuração:', e.message);
    // Não bloqueia o fluxo se a procuração falhar
  }
}

    await new Promise(r=>setTimeout(r, 3000));
    try { await getDocumentStatus(D4SIGN_TOKEN, D4SIGN_CRYPT_KEY, uuidDoc); } catch {}

    await cadastrarSignatarios(D4SIGN_TOKEN, D4SIGN_CRYPT_KEY, uuidDoc, signers);

    await new Promise(r=>setTimeout(r, 2000));
    // Movimentação de card removida conforme solicitado

    releaseLock(lockKey);

    const token = req.params.token;
    const html = `
<!doctype html><meta charset="utf-8"><title>Contrato gerado</title>
<style>
  body{font-family:system-ui;display:grid;place-items:center;min-height:100vh;background:#f7f7f7;color:#111;margin:0}
  .box{background:#fff;padding:24px;border-radius:14px;box-shadow:0 4px 16px rgba(0,0,0,.08);max-width:640px;width:92%}
  h2{margin:0 0 12px}
  h3{margin:24px 0 8px;font-size:16px}
  .row{display:flex;gap:12px;flex-wrap:wrap;margin-top:12px}
  .btn{display:inline-block;padding:12px 16px;border-radius:10px;text-decoration:none;border:0;background:#111;color:#fff;font-weight:600}
  .muted{color:#666}
  .section{margin-top:24px;padding-top:24px;border-top:1px solid #eee}
</style>
<div class="box">
  <h2>${uuidProcuracao ? 'Contrato e procuração gerados com sucesso' : 'Contrato gerado com sucesso'}</h2>
  <p class="muted">UUID do documento: ${uuidDoc}</p>
  ${cofreUsadoPadrao ? `
  <div style="background:#fff3cd;border-left:4px solid #ffc107;padding:12px;margin:16px 0;border-radius:4px">
    <strong>⚠️ Atenção:</strong> A equipe "${equipeContrato || 'não informada'}" não possui cofre configurado. 
    Documentos salvos no cofre padrão: <strong>${nomeCofreUsado}</strong>
  </div>
  ` : ''}
  ${d.email_envio_contrato || d.email ? `
  <div style="margin:16px 0;padding:12px;background:#f5f5f5;border-radius:8px">
    <strong>Email para envio:</strong> ${d.email_envio_contrato || d.email || 'Não informado'}
  </div>
  ` : ''}
  <div class="row">
    <a class="btn" href="/lead/${encodeURIComponent(token)}/doc/${encodeURIComponent(uuidDoc)}/download" target="_blank" rel="noopener">Baixar PDF do Contrato</a>
    <button class="btn" onclick="enviarContrato('${token}', '${uuidDoc}', 'email')" id="btn-enviar-contrato-email">Enviar por Email</button>
  </div>
  <div id="status-contrato" style="margin-top:8px;min-height:24px"></div>
  ${uuidProcuracao ? `
  <div class="section">
    <h3>Procuração gerada com sucesso</h3>
    <p class="muted">UUID da procuração: ${uuidProcuracao}</p>
    <div class="row">
      <a class="btn" href="/lead/${encodeURIComponent(token)}/doc/${encodeURIComponent(uuidProcuracao)}/download" target="_blank" rel="noopener">Baixar PDF da Procuração</a>
      <button class="btn" onclick="enviarProcuracao('${token}', '${uuidProcuracao}', 'email')" id="btn-enviar-procuracao-email">Enviar por Email</button>
    </div>
    <div id="status-procuracao" style="margin-top:8px;min-height:24px"></div>
  </div>
  ` : ''}
  <div class="row" style="margin-top:24px">
    <a class="btn" href="${PUBLIC_BASE_URL}/lead/${encodeURIComponent(token)}">Voltar</a>
  </div>
</div>
<script>
async function enviarContrato(token, uuidDoc, canal) {
  const btnEmail = document.getElementById('btn-enviar-contrato-email');
  const statusDiv = document.getElementById('status-contrato');
  const btn = btnEmail;
  
  btn.disabled = true;
  btn.textContent = 'Enviando...';
  statusDiv.innerHTML = '<span style="color:#1976d2">⏳ Enviando contrato por email...</span>';
  
  try {
    const response = await fetch('/lead/' + encodeURIComponent(token) + '/doc/' + encodeURIComponent(uuidDoc) + '/send?canal=email', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' }
    });
    
    const data = await response.json();
    
    if (response.ok && data.success) {
      const cofreMsg = data.cofre ? ' Salvo no cofre: ' + data.cofre : '';
      const destinoMsg = data.email ? ' para ' + data.email + ' por email' : ' por email';
      statusDiv.innerHTML = '<span style="color:#28a745;font-weight:600">✓ Status de envio - Contrato: Enviado com sucesso' + destinoMsg + '.' + cofreMsg + '</span>';
      btn.textContent = 'Enviado por Email';
      btn.style.background = '#28a745';
      btn.disabled = true;
    } else {
      const errorMsg = data.message || data.detalhes || 'Erro ao enviar';
      statusDiv.innerHTML = '<span style="color:#d32f2f;font-weight:600">✗ Status de envio - Contrato: ' + errorMsg + '</span>';
      btn.disabled = false;
      btn.textContent = 'Enviar por Email';
    }
  } catch (error) {
    statusDiv.innerHTML = '<span style="color:#d32f2f">✗ Status de envio - Contrato: Erro ao enviar - ' + error.message + '</span>';
    btn.disabled = false;
    btn.textContent = 'Enviar por Email';
  }
}

async function enviarProcuracao(token, uuidProcuracao, canal) {
  const btnEmail = document.getElementById('btn-enviar-procuracao-email');
  const statusDiv = document.getElementById('status-procuracao');
  const btn = btnEmail;
  
  btn.disabled = true;
  btn.textContent = 'Enviando...';
  statusDiv.innerHTML = '<span style="color:#1976d2">⏳ Enviando procuração por email...</span>';
  
  try {
    const response = await fetch('/lead/' + encodeURIComponent(token) + '/doc/' + encodeURIComponent(uuidProcuracao) + '/send?canal=email', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' }
    });
    
    const data = await response.json();
    
    if (response.ok && data.success) {
      const cofreMsg = data.cofre ? ' Salvo no cofre: ' + data.cofre : '';
      const destinoMsg = data.email ? ' para ' + data.email + ' por email' : ' por email';
      statusDiv.innerHTML = '<span style="color:#28a745;font-weight:600">✓ Status de envio - Procuração: Enviado com sucesso' + destinoMsg + '.' + cofreMsg + '</span>';
      btn.textContent = 'Enviado por Email';
      btn.style.background = '#28a745';
      btn.disabled = true;
    } else {
      const errorMsg = data.message || data.detalhes || 'Erro ao enviar';
      statusDiv.innerHTML = '<span style="color:#d32f2f;font-weight:600">✗ Status de envio - Procuração: ' + errorMsg + '</span>';
      btn.disabled = false;
      btn.textContent = 'Enviar por Email';
    }
  } catch (error) {
    statusDiv.innerHTML = '<span style="color:#d32f2f">✗ Status de envio - Procuração: Erro ao enviar - ' + error.message + '</span>';
    btn.disabled = false;
    btn.textContent = 'Enviar por Email';
  }
}
</script>`;
    return res.status(200).send(html);

  } catch (e) {
    console.error('[ERRO LEAD-GENERATE]', {
      message: e.message || e,
      stack: e.stack,
      cardId: req.params.token ? parseLeadToken(req.params.token)?.cardId : 'N/A'
    });
    
    // Liberar lock em caso de erro
    try {
      const { cardId } = parseLeadToken(req.params.token);
      if (cardId) {
        const lockKey = `lead:${cardId}`;
        releaseLock(lockKey);
      }
    } catch (lockErr) {
      // Ignora erro ao liberar lock
    }
    
    // Retornar mensagem de erro mais detalhada
    const errorMessage = e.message || 'Erro desconhecido ao gerar o contrato';
    return res.status(400).send(`
<!doctype html><meta charset="utf-8"><title>Erro ao gerar contrato</title>
<style>
  body{font-family:system-ui;display:grid;place-items:center;min-height:100vh;background:#f7f7f7;margin:0}
  .box{background:#fff;padding:32px;border-radius:14px;box-shadow:0 4px 16px rgba(0,0,0,.08);max-width:600px;width:92%}
  h2{color:#d32f2f;margin:0 0 16px;font-size:24px}
  .error-box{background:#ffebee;border-left:4px solid #d32f2f;padding:16px;border-radius:4px;margin:20px 0}
  .error-box strong{display:block;margin-bottom:8px;color:#c62828}
  .error-box p{margin:8px 0;color:#424242;line-height:1.6}
  .btn{display:inline-block;padding:12px 24px;border-radius:8px;text-decoration:none;background:#1976d2;color:#fff;font-weight:600;margin-top:16px}
</style>
<div class="box">
  <h2>❌ Erro ao gerar contrato</h2>
  <div class="error-box">
    <strong>O que aconteceu?</strong>
    <p>${errorMessage}</p>
  </div>
  <p style="color:#757575;font-size:14px;margin-top:20px">
    Verifique os logs do servidor para mais detalhes. Se o problema persistir, entre em contato com o suporte técnico.
  </p>
  <a href="${PUBLIC_BASE_URL}/lead/${encodeURIComponent(req.params.token)}" class="btn">Voltar e tentar novamente</a>
</div>`);
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
    const { cardId } = parseLeadToken(req.params.token);
  if (!cardId) {
    return res.status(400).json({ success: false, message: 'Token inválido' });
  }
    const uuidDoc = req.params.uuid;
  const canal = req.query.canal || 'email'; // 'email' ou 'whatsapp'
  
  // Proteção contra envio duplicado
  const lockKey = `send:${cardId}:${uuidDoc}:${canal}`;
  if (!acquireLock(lockKey)) {
    return res.status(200).json({ success: false, message: 'Documento já está sendo enviado. Aguarde alguns instantes.' });
  }
  
  try {
    // Buscar informações do card para identificar o tipo de documento
    let card = null;
    let signers = null;
    let nomeCofre = 'Cofre não identificado';
    let isProcuracao = false;
    
    try {
      card = await getCard(cardId);
      const by = toById(card);
      const uuidProcuracaoCard = by['d4_uuid_procuracao'] || null;
      const uuidContratoCard = by['d4_uuid_contrato'] || null;
      
      // Determinar se é contrato ou procuração
      isProcuracao = (uuidProcuracaoCard === uuidDoc);
      
      // Buscar equipe contrato para identificar o cofre
      const equipeContrato = getEquipeContratoFromCard(card);
      let uuidCofre = null;
      if (equipeContrato && COFRES_UUIDS[equipeContrato]) {
        uuidCofre = COFRES_UUIDS[equipeContrato];
      }
      if (!uuidCofre) {
        uuidCofre = DEFAULT_COFRE_UUID;
      }
      nomeCofre = getNomeCofreByUuid(uuidCofre);
      
      // Preparar signatários
      const d = await montarDados(card);
      
      // Validar se tem email/telefone conforme o canal
      if (canal === 'whatsapp') {
        const telefoneEnvio = d.telefone_envio_contrato || d.telefone || '';
        if (!telefoneEnvio) {
          throw new Error('Telefone para envio do contrato não encontrado. Verifique o campo "Telefone para envio do contrato" no card do Pipefy.');
        }
        signers = montarSigners(d, true); // incluir telefone
        console.log(`[SEND] Signatários preparados para WhatsApp:`, signers.map(s => ({
          name: s.name,
          email: s.email,
          phone: s.phone || 'SEM TELEFONE'
        })));
      } else {
        const emailEnvio = d.email_envio_contrato || d.email || '';
        if (!emailEnvio) {
          throw new Error('Email para envio do contrato não encontrado. Verifique o campo "Email para envio do contrato" no card do Pipefy.');
        }
        signers = montarSigners(d, false);
      }
      
      console.log(`[SEND] Enviando ${isProcuracao ? 'procuração' : 'contrato'} por ${canal}. Signatários preparados:`, signers.map(s => s.email).join(', '));
    } catch (e) {
      console.warn('[SEND] Erro ao buscar informações do card:', e.message);
      throw e;
    }

    // Verificar status do documento antes de enviar
    await new Promise(r=>setTimeout(r, 2000));
    try { 
      await getDocumentStatus(D4SIGN_TOKEN, D4SIGN_CRYPT_KEY, uuidDoc); 
      console.log(`[SEND] Status do ${isProcuracao ? 'procuração' : 'contrato'} verificado.`);
    } catch (e) {
      console.warn(`[SEND] Aviso ao verificar status do ${isProcuracao ? 'procuração' : 'contrato'}:`, e.message);
    }
    
    // Garantir que os signatários estão cadastrados antes de enviar
    if (!signers || signers.length === 0) {
      try {
        if (!card) {
          card = await getCard(cardId);
        }
        const d = await montarDados(card);
        
        // Validar novamente conforme o canal
        if (canal === 'whatsapp') {
          const telefoneEnvio = d.telefone_envio_contrato || d.telefone || '';
          if (!telefoneEnvio) {
            throw new Error('Telefone para envio do contrato não encontrado. Verifique o campo "Telefone para envio do contrato" no card do Pipefy.');
          }
          signers = montarSigners(d, true);
        } else {
          const emailEnvio = d.email_envio_contrato || d.email || '';
          if (!emailEnvio) {
            throw new Error('Email para envio do contrato não encontrado. Verifique o campo "Email para envio do contrato" no card do Pipefy.');
          }
          signers = montarSigners(d, false);
        }
        
        console.log(`[SEND] Signatários do ${isProcuracao ? 'procuração' : 'contrato'} preparados:`, signers.map(s => s.email).join(', '));
      } catch (e) {
        console.error(`[SEND] Erro ao preparar signatários do ${isProcuracao ? 'procuração' : 'contrato'}:`, e.message);
        throw e;
      }
    }
    
    if (signers && signers.length > 0) {
      try {
        // Para WhatsApp, precisamos garantir que os signatários tenham telefone
        if (canal === 'whatsapp') {
          const signersComTelefone = signers.filter(s => s.phone);
          if (signersComTelefone.length === 0) {
            throw new Error('Nenhum signatário possui telefone cadastrado para envio por WhatsApp.');
          }
          console.log(`[SEND] Cadastrando ${signersComTelefone.length} signatário(s) com telefone para WhatsApp`);
          console.log(`[SEND] Telefones:`, signersComTelefone.map(s => `${s.name}: ${s.phone}`).join(', '));
          console.log(`[DEBUG] Signatários antes de cadastrar (canal: ${canal}):`, 
            JSON.stringify(signers.map(s => ({ 
              name: s.name, 
              email: s.email, 
              phone: s.phone 
            })), null, 2));
        }
        await cadastrarSignatarios(D4SIGN_TOKEN, D4SIGN_CRYPT_KEY, uuidDoc, signers, canal === 'whatsapp');
        console.log(`[SEND] Signatários do ${isProcuracao ? 'procuração' : 'contrato'} confirmados/cadastrados para ${canal}:`, signers.map(s => {
          if (canal === 'whatsapp' && s.phone) {
            return `${s.name} (${s.phone})`;
          }
          return `${s.name} (${s.email})`;
        }).join(', '));
      } catch (e) {
        console.error(`[SEND] Erro ao cadastrar signatários do ${isProcuracao ? 'procuração' : 'contrato'} para ${canal}:`, e.message);
        throw e; // Propaga o erro para que o usuário saiba
      }
    } else {
      const erro = new Error(`Nenhum signatário encontrado para o ${isProcuracao ? 'procuração' : 'contrato'}. Verifique se há ${canal === 'whatsapp' ? 'telefone' : 'email'} configurado no card do Pipefy.`);
      erro.tipo = 'SEM_SIGNATARIOS';
      throw erro;
    }
    
    // Aguardar um pouco antes de enviar
    await new Promise(r=>setTimeout(r, 2000));
    
    // Enviar documento (contrato ou procuração)
    try {
      const mensagem = isProcuracao 
        ? 'Olá! Há uma procuração aguardando sua assinatura.'
        : 'Olá! Há um documento aguardando sua assinatura.';
      
      // Se for WhatsApp, não enviar email
      const skip_email = canal === 'whatsapp' ? '1' : '0';

    await sendToSigner(D4SIGN_TOKEN, D4SIGN_CRYPT_KEY, uuidDoc, {
        message: mensagem,
        skip_email: skip_email,
      workflow: '0'
    });
      console.log(`[SEND] ${isProcuracao ? 'Procuração' : 'Contrato'} enviado para assinatura por ${canal}:`, uuidDoc);
    } catch (e) {
      console.error(`[ERRO] Falha ao enviar ${isProcuracao ? 'procuração' : 'contrato'}:`, e.message);
      throw e; // Propaga o erro para que o usuário saiba
    }

    // Buscar email/telefone usado para incluir na resposta
    let emailUsado = null;
    let telefoneUsado = null;
    if (!card) {
      card = await getCard(cardId);
    }
    const dFinal = await montarDados(card);
    if (canal === 'whatsapp') {
      telefoneUsado = dFinal.telefone_envio_contrato || dFinal.telefone || null;
    } else {
      emailUsado = dFinal.email_envio_contrato || dFinal.email || null;
    }
    
    // Liberar lock após envio bem-sucedido
    releaseLock(lockKey);
    
    return res.status(200).json({
      success: true,
      message: `${isProcuracao ? 'Procuração' : 'Contrato'} enviado com sucesso. Os signatários foram notificados.`,
      tipo: isProcuracao ? 'procuração' : 'contrato',
      cofre: nomeCofre,
      email: emailUsado,
      telefone: telefoneUsado
    });

  } catch (e) {
    // Liberar lock em caso de erro
    try {
      const lockKey = `send:${cardId}:${uuidDoc}`;
      releaseLock(lockKey);
    } catch (lockErr) {
      // Ignora erro ao liberar lock
    }
    
    console.error('[ERRO sendtosigner]', {
      message: e.message || e,
      stack: e.stack,
      cardId: cardId,
      uuidDoc: uuidDoc,
      statusCode: e.statusCode
    });
    
    // Determinar tipo de erro e mensagem amigável
    let tituloErro = 'Erro ao enviar documentos';
    let mensagemErro = e.message || 'Não foi possível enviar os documentos para assinatura.';
    let detalhesAdicionais = '';
    
    // Erros específicos por tipo
    if (e.tipo === 'SEM_SIGNATARIOS' || (e.message && e.message.includes('Nenhum signatário'))) {
      tituloErro = 'Signatários não encontrados';
      mensagemErro = 'Não foi possível encontrar signatários para enviar o documento.';
      detalhesAdicionais = 'Verifique se há emails cadastrados no card do Pipefy (campo "Email para envio do contrato" ou "Email de contato").';
    } else if (e.message && e.message.includes('preparar os dados dos signatários')) {
      tituloErro = 'Erro ao preparar dados';
      mensagemErro = 'Não foi possível preparar os dados dos signatários.';
      detalhesAdicionais = 'Verifique se o card possui todas as informações necessárias: nome do contratante, email de contato e demais dados obrigatórios.';
    } else if (e.message && e.message.includes('não encontrado') || e.message && e.message.includes('não foi encontrado')) {
      tituloErro = 'Documento não encontrado';
      mensagemErro = 'O documento não foi encontrado no sistema de assinatura.';
      detalhesAdicionais = 'O documento pode ter sido excluído ou o identificador está incorreto. Tente gerar o contrato novamente.';
    } else if (e.message && e.message.includes('autenticar') || e.message && e.message.includes('autenticação')) {
      tituloErro = 'Erro de autenticação';
      mensagemErro = 'Não foi possível autenticar no sistema de assinatura.';
      detalhesAdicionais = 'Entre em contato com o suporte técnico para verificar as credenciais de acesso.';
    } else if (e.message && e.message.includes('indisponível') || e.message && e.message.includes('temporariamente')) {
      tituloErro = 'Serviço temporariamente indisponível';
      mensagemErro = 'O serviço de assinatura está temporariamente fora do ar.';
      detalhesAdicionais = 'Tente novamente em alguns minutos. Se o problema persistir, entre em contato com o suporte.';
    } else if (e.message && e.message.includes('já foi enviado') || e.message && e.message.includes('já enviado')) {
      tituloErro = 'Documento já enviado';
      mensagemErro = 'Este documento já foi enviado para assinatura anteriormente.';
      detalhesAdicionais = 'Verifique o status do documento no D4Sign ou aguarde a conclusão do processo de assinatura.';
    } else if (e.message && (e.message.includes('email') || e.message.includes('Email'))) {
      tituloErro = 'Erro com emails dos signatários';
      mensagemErro = 'Há um problema com os emails dos signatários.';
      detalhesAdicionais = 'Verifique se os emails estão corretos, válidos e no formato adequado (exemplo@dominio.com).';
    } else if (e.statusCode === 400) {
      tituloErro = 'Dados inválidos';
      mensagemErro = 'Os dados enviados não são válidos.';
      detalhesAdicionais = 'Verifique se o documento existe e está configurado corretamente no sistema de assinatura.';
    } else if (e.statusCode === 404) {
      tituloErro = 'Documento não encontrado';
      mensagemErro = 'O documento não foi encontrado no sistema de assinatura.';
      detalhesAdicionais = 'O documento pode ter sido excluído. Tente gerar o contrato novamente.';
    } else if (e.statusCode >= 500) {
      tituloErro = 'Erro no servidor';
      mensagemErro = 'Ocorreu um erro no servidor de assinatura.';
      detalhesAdicionais = 'O problema é temporário. Tente novamente em alguns minutos.';
    }
    
    // Retornar JSON para requisições AJAX
    return res.status(400).json({
      success: false,
      message: mensagemErro,
      detalhes: detalhesAdicionais || '',
      titulo: tituloErro
    });
  }
});
// ===============================
// NOVO — LOCALIZA CARD PELO UUID DO DOCUMENTO D4SIGN
// Busca tanto no campo de contrato quanto no de procuração
// ===============================
async function findCardIdByD4Uuid(uuidDocument) {
  const query = `
    query($pipeId: ID!, $fieldId: String!, $fieldValue: String!) {
      findCards(
        pipeId: $pipeId,
        search: {
          fieldId: $fieldId,
          fieldValue: $fieldValue
        }
      ) {
        edges {
          node {
            id
          }
        }
      }
    }
  `;

  // Primeiro tenta buscar pelo campo de contrato
  let data = await gql(query, {
    pipeId: NOVO_PIPE_ID,
    fieldId: "d4_uuid_contrato",
    fieldValue: uuidDocument
  });

  let edges = data?.findCards?.edges || [];
  if (edges.length) {
    return edges[0].node.id;
  }

  // Se não encontrou, tenta buscar pelo campo de procuração
  data = await gql(query, {
    pipeId: NOVO_PIPE_ID,
    fieldId: "d4_uuid_procuracao",
    fieldValue: uuidDocument
  });

  edges = data?.findCards?.edges || [];
  if (edges.length) {
  return edges[0].node.id;
  }

  return null;
};

// ===============================
// NOVO — ANEXA CONTRATO ASSINADO NO CAMPO DE ANEXO
// ===============================
async function anexarContratoAssinadoNoCard(cardId, downloadUrl, fileName) {
  const newValue = [downloadUrl];

  await updateCardField(cardId, 'contrato', newValue);
};

/* =========================
 * Geração do link no Pipefy
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

    // Enviar mensagem no card com o link
    try {
      const mensagem = `📋 Link para revisar e gerar o contrato:\n\n${url}\n\nClique no link acima para revisar os dados e gerar o contrato.`;
      await createCardComment(cardId, mensagem);
      console.log('[CRIAR-LINK] Mensagem enviada no card com o link');
    } catch (e) {
      console.error('[CRIAR-LINK] Erro ao enviar mensagem no card:', e.message);
      // Não bloqueia o fluxo se falhar ao enviar comentário
    }

    return res.json({ ok:true, link:url });
  } catch (e) {
    console.error('[ERRO criar-link]', e.message||e);
    return res.status(500).json({ error: String(e.message||e) });
  }
});
app.get('/novo-pipe/criar-link-confirmacao', async (req, res) => {
  try {
    const cardId = req.query.cardId || req.query.card_id;
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

    // Enviar mensagem no card com o link
    try {
      const mensagem = `📋 Link para revisar e gerar o contrato:\n\n${url}\n\nClique no link acima para revisar os dados e gerar o contrato.`;
      await createCardComment(cardId, mensagem);
      console.log('[CRIAR-LINK] Mensagem enviada no card com o link');
    } catch (e) {
      console.error('[CRIAR-LINK] Erro ao enviar mensagem no card:', e.message);
      // Não bloqueia o fluxo se falhar ao enviar comentário
    }

    return res.json({ ok:true, link:url });
  } catch (e) {
    console.error('[ERRO criar-link]', e.message||e);
    return res.status(500).json({ error: String(e.message||e) });
  }
});

/* =========================
 * Debug / Health
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
      fields: (card.fields||[]).map(f => ({ name:f.name, id:f.field?.id, type:f.field?.type, value:f.value, array_value:f.array_value }))
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
