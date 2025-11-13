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

  DEFAULT_COFRE_UUID
} = process.env;

PORT = PORT || 3000;
PIPE_GRAPHQL_ENDPOINT = PIPE_GRAPHQL_ENDPOINT || 'https://api.pipefy.com/graphql';
PIPEFY_FIELD_LINK_CONTRATO = PIPEFY_FIELD_LINK_CONTRATO || 'd4_contrato';

if (!PUBLIC_BASE_URL || !PUBLIC_LINK_SECRET) console.warn('[AVISO] Configure PUBLIC_BASE_URL e PUBLIC_LINK_SECRET');
if (!PIPE_API_KEY) console.warn('[AVISO] PIPE_API_KEY ausente');
if (!D4SIGN_TOKEN || !D4SIGN_CRYPT_KEY) console.warn('[AVISO] D4SIGN_TOKEN / D4SIGN_CRYPT_KEY ausentes');
if (!TEMPLATE_UUID_CONTRATO) console.warn('[AVISO] TEMPLATE_UUID_CONTRATO (Marca) ausente');
if (!TEMPLATE_UUID_CONTRATO_OUTROS) console.warn('[AVISO] TEMPLATE_UUID_CONTRATO_OUTROS (Outros) ausente');

// Cofres mapeados por responsável
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
    headers: { 'Authorization': `Bearer ${PIPE_API_KEY}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ query, variables })
  });
  const j = await r.json();
  if (!r.ok || j.errors) throw new Error(`Pipefy GQL: ${r.status} ${JSON.stringify(j.errors||{})}`);
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
function checklistToText(v){
  try{
    const arr = Array.isArray(v)? v : JSON.parse(v);
    return Array.isArray(arr) ? arr.join(', ') : String(v || '');
  }catch{ return String(v || ''); }
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

// Assignee parsing (para cofre)
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

// Busca campo statement por marca N prioritariamente, com fallback para connector “serviços marca - N”
function buscarServicoN(card, n){
  // ids conhecidos de statements
  const mapStmt = {
    1: null, // marca 1 usa buscarServicoStatement(card) já existente
    2: 'statement_432366f2_fbbc_448d_82e4_fbd73c3fc52e',
    3: 'statement_c5616541_5f30_41b9_bd74_e2bd2063f253',
    4: 'statement_8d833401_2294_448b_a34f_07f86c52981c',
    5: 'statement_ca0eb59e_a015_4628_8c56_28af6e23c8d9'
  };
  const stmtId = mapStmt[n];
  let v = '';
  if (stmtId){
    const f = getFieldObjById(card, stmtId);
    v = String(f?.value||'').replace(/<[^>]*>/g,' ').trim();
  }
  if (!v){
    const connId = n===2 ? 'servi_os_marca_2'
               : n===3 ? 'servi_os_marca_3'
               : n===4 ? 'servi_os_marca_4'
               : n===5 ? 'servi_os_marca_5'
               : null;
    if (connId){
      const f = getFieldObjById(card, connId);
      if (f?.value){
        try{
          const arr = JSON.parse(f.value);
          if (Array.isArray(arr) && arr[0]) v = String(arr[0]);
        }catch{
          v = String(f.value);
        }
      }
    }
  }
  return v;
}

// Busca campo statement que contenha informações de serviço (marca 1 por padrão)
function buscarServicoStatement(card){
  // igual ao código original: varre statements da marca 1
  const statementFields = (card.fields||[]).filter(f => String(f?.field?.type||'').toLowerCase() === 'statement');
  for (const field of statementFields){
    const name = String(field?.name||'').toLowerCase();
    const description = String(field?.field?.description || '').toLowerCase();
    const isServicos = name.includes('serviço') || name.includes('servico') || description.includes('serviços marca');
    if (isServicos){
      let value = String(field?.value||'').replace(/<[^>]*>/g,' ').trim();
      value = value.replace(/^serviços?.*?:\s*/i,'').trim();
      if (value) return value;
    }
  }
  return '';
}

// Normalização apenas para “Detalhes do serviço …” (não usada nas linhas de “Quantidade … Descrição …”)
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

function safe(s, allowEmpty=false){
  if (s===null || s===undefined) return allowEmpty ? '' : '';
  const out = String(s).replace(/\s+/g,' ').trim();
  return allowEmpty ? out : out;
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
  const tipoMarca1 = checklistToText(by['tipo_de_marca'] || by['checklist_vertical'] || getFirstByNames(card, ['tipo de marca']));

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
  const serv1Stmt = firstNonEmpty(buscarServicoStatement(card));
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

  // Decide template: se houver qualquer “MARCA” usar modelo de Marca, caso contrário, usar “Outros”
  const anyMarca = [k1,k2,k3,k4,k5].includes('MARCA');
  const templateToUse = anyMarca ? TEMPLATE_UUID_CONTRATO : TEMPLATE_UUID_CONTRATO_OUTROS;

  // Contato contratante 1
  const contatoNome     = by['nome_1'] || getFirstByNames(card, ['nome do contato','contratante','responsável legal','responsavel legal']) || '';
  const contatoEmail    = by['email_de_contato'] || getFirstByNames(card, ['email','e-mail']) || '';
  const contatoTelefone = by['telefone_de_contato'] || getFirstByNames(card, ['telefone','celular','whatsapp','whats']) || '';

  // Contato contratante 2 e cotitular
  const contato2Nome = by['nome_2'] || getFirstByNames(card, ['contratante 2', 'nome contratante 2']) || '';
  const contato2Email = by['email_2'] || getFirstByNames(card, ['email 2', 'e-mail 2']) || '';
  const contato2Telefone = by['telefone_2'] || getFirstByNames(card, ['telefone 2', 'celular 2']) || '';

  // Envio do contrato principal e cotitular
  const emailEnvioContrato = by['email_para_envio_do_contrato'] || contatoEmail || '';
  const emailCotitularEnvio = by['copy_of_email_para_envio_do_contrato'] || '';
  const telefoneCotitularEnvio = by['copy_of_telefone_para_envio_do_contrato'] || '';

  // Documento (CPF/CNPJ)
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

  // Endereço (CNPJ)
  const cepCnpj    = by['cep_do_cnpj']     || '';
  const ruaCnpj    = by['rua_av_do_cnpj']  || '';
  const bairroCnpj = by['bairro_do_cnpj']  || '';
  const cidadeCnpj = by['cidade_do_cnpj']  || '';
  const ufCnpj     = by['estado_do_cnpj']  || '';
  const numeroCnpj = by['n_mero_1']        || getFirstByNames(card, ['numero','número','nº']) || '';

  // Vendedor (cofre)
  const vendedor = (()=>{
    const v = extractAssigneeNames(by['vendedor_respons_vel'] || by['vendedor_respons_vel_1'] || by['respons_vel_5']);
    return v[0] || '';
  })();

  // Riscos 1..5
  const risco1 = by['risco_da_marca'] || '';
  const risco2 = by['copy_of_copy_of_risco_da_marca'] || '';
  const risco3 = by['copy_of_risco_da_marca_3'] || '';
  const risco4 = by['copy_of_risco_da_marca_3_1'] || '';
  const risco5 = by['copy_of_risco_da_marca_4'] || '';

  // Nacionalidade e etc
  const nacionalidade = by['nacionalidade'] || '';
  const selecaoCnpjOuCpf = by['cnpj_ou_cpf'] || '';
  const estadoCivil = by['estado_civ_l'] || '';

  // Cláusula adicional
  const clausulaAdicional = (by['cl_usula_adicional'] && String(by['cl_usula_adicional']).trim()) ? by['cl_usula_adicional'] : 'Sem aditivos contratuais.';

  // Contratantes blocos
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

  const nacionalidade2 = by['nacionalidade_2'] || '';
  const estadoCivil2 = by['estado_civ_l_2'] || '';
  const rua2 = by['rua_2'] || by['rua_av_do_cnpj_2'] || '';
  const bairro2 = by['bairro_2'] || by['bairro_do_cnpj_2'] || '';
  const numero2 = by['numero_2'] || by['n_mero_2'] || '';
  const cidade2 = by['cidade_2'] || by['cidade_do_cnpj_2'] || '';
  const uf2 = by['estado_2'] || by['estado_do_cnpj_2'] || '';
  const cep2 = by['cep_2'] || by['cep_do_cnpj_2'] || '';
  const rg2 = by['rg_2'] || '';
  const cpf2 = by['cpf_2'] || '';
  const cnpj2 = by['cnpj_2'] || '';
  const docSelecao2 = by['cnpj_ou_cpf_2'] || '';

  const temDadosContratante2 = Boolean(
    contato2Nome || nacionalidade2 || estadoCivil2 || rua2 || bairro2 || numero2 ||
    cidade2 || uf2 || cep2 || rg2 || cpf2 || cnpj2 || docSelecao2 || contato2Telefone || contato2Email
  );

  const contratante2Texto = temDadosContratante2
    ? montarTextoContratante({
        nome: contato2Nome,
        nacionalidade: nacionalidade2,
        estadoCivil: estadoCivil2,
        rua: rua2 || ruaCnpj,
        bairro: bairro2 || bairroCnpj,
        numero: numero2 || numeroCnpj,
        cidade: cidade2 || cidadeCnpj,
        uf: uf2 || ufCnpj,
        cep: cep2 || cepCnpj,
        rg: rg2,
        docSelecao: docSelecao2,
        cpf: cpf2,
        cnpj: cnpj2,
        telefone: contato2Telefone || telefoneCotitularEnvio,
        email: contato2Email || emailCotitularEnvio
      })
    : '';

  // Dados para contato 1 e 2
  const dadosContato1 = [contatoNome, contatoTelefone, contatoEmail].filter(Boolean).join(' | ');
  const dadosContato2 = temDadosContratante2
    ? [contato2Nome || 'Cotitular', (contato2Telefone || telefoneCotitularEnvio || ''), (contato2Email || emailCotitularEnvio || '')].filter(Boolean).join(' | ')
    : '';

  // Quantidade e descrição “sem normalizar” para cada categoria, conforme templates
  // Para Marca: somente preenchidos os campos de marca
  // Para Outros: preencher PATENTE, DESENHO INDUSTRIAL, COPYRIGHT e OUTROS
  const mkQtd = [k1,k2,k3,k4,k5].filter(k=>k==='MARCA' && (String(serv1Stmt+serv2Stmt+serv3Stmt+serv4Stmt+serv5Stmt))).length;
  // Em vez de contar pelo kind apenas, vamos contar efetivamente entradas não vazias por kind
  const entries = [
    {kind:k1, title:tituloMarca1, tipo:tipoMarca1, classes:classeSomenteNumeros1, stmt:serv1Stmt, risco:risco1, lines:linhasMarcasEspec1},
    {kind:k2, title:tituloMarca2, tipo:tipoMarca2, classes:classeSomenteNumeros2, stmt:serv2Stmt, risco:risco2, lines:linhasMarcasEspec2},
    {kind:k3, title:tituloMarca3, tipo:tipoMarca3, classes:classeSomenteNumeros3, stmt:serv3Stmt, risco:risco3, lines:linhasMarcasEspec3},
    {kind:k4, title:tituloMarca4, tipo:tipoMarca4, classes:classeSomenteNumeros4, stmt:serv4Stmt, risco:risco4, lines:linhasMarcasEspec4},
    {kind:k5, title:tituloMarca5, tipo:tipoMarca5, classes:classeSomenteNumeros5, stmt:serv5Stmt, risco:risco5, lines:linhasMarcasEspec5},
  ].filter(e => String(e.title||e.stmt||'').trim());

  // Monta linhas de “quantidade + descrição” conforme pedido
  const qtdDesc = {
    MARCA: '',
    PATENTE: '',
    'DESENHO INDUSTRIAL': '',
    'COPYRIGHT/DIREITO AUTORAL': '',
    OUTROS: ''
  };
  const detalhes = {
    MARCA: ['', '', '', '', ''],
    PATENTE: ['', '', '', '', ''],
    'DESENHO INDUSTRIAL': ['', '', '', '', ''],
    'COPYRIGHT/DIREITO AUTORAL': ['', '', '', '', ''],
    OUTROS: ['', '', '', '', '']
  };
  const headersServicos = { // cabeçalhos “Cabeçalho - SERVIÇOS” e “Cabeçalho - SERVIÇOS 2”
    h1: '', h2: ''
  };
  // Preenche contadores por kind
  const byKind = { 'MARCA':[], 'PATENTE':[], 'DESENHO INDUSTRIAL':[], 'COPYRIGHT/DIREITO AUTORAL':[], 'OUTROS':[] };
  entries.forEach(e => byKind[e.kind].push(e));

  // Marca: quantidade e descrição “1 Registro de Marca JUNTO AO INPI” etc
  const makeQtdDescLine = (kind, arr) => {
    if (!arr.length) return '';
    // Extrai o texto do serviço do statement bruto, sem normalizar
    // Usa o primeiro item como base da frase de quantidade
    const baseServico = String(arr[0].stmt||'').trim() || (kind==='MARCA' ? 'Registro de Marca' : kind);
    const qtd = arr.length;
    // Ex: “2 Registro de Marca JUNTO AO INPI”
    return `${qtd} ${baseServico} JUNTO AO INPI`;
  };

  Object.keys(byKind).forEach(k=>{
    qtdDesc[k] = makeQtdDescLine(k, byKind[k]);
  });

  // Detalhes por item até 5
  ['MARCA','PATENTE','DESENHO INDUSTRIAL','COPYRIGHT/DIREITO AUTORAL','OUTROS'].forEach(k=>{
    const arr = byKind[k];
    for (let i=0;i<5;i++){
      const e = arr[i];
      if (!e){ detalhes[k][i] = ''; continue; }
      const cab = normalizarCabecalhoDetalhe(k, e.title, e.tipo, e.classes);
      detalhes[k][i] = cab;
    }
  });

  // Cabeçalhos “SERVIÇOS” e “SERVIÇOS 2” usados no formulário de classes do contrato de marca
  // H1 descreve o primeiro bloco, H2 descreve o segundo quando houver entradas numa “marca 2”
  headersServicos.h1 = byKind['MARCA'][0] ? `MARCA: ${byKind['MARCA'][0].title||''}` : '';
  headersServicos.h2 = byKind['MARCA'][1] ? `MARCA: ${byKind['MARCA'][1].title||''}` : '';

  // Risco agregado: “serviço: marca_1 - RISCO: risco marca_1, serviço: marca_2 - RISCO: risco marca_2”
  const riscoAgregado = entries
    .map((e, idx)=> {
      const servTxt = String(e.stmt||e.kind||'').trim();
      const name = e.title || `serviço ${idx+1}`;
      const r = String(e.risco||'').trim();
      if (!servTxt && !r) return '';
      return `serviço: ${name} - RISCO: ${r || 'Não informado'}`;
    })
    .filter(Boolean)
    .join(', ');

  return {
    // base
    cardId: card.id,
    templateToUse, // decide qual template será usado depois

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

    // Contatos
    email: contatoEmail || '',
    telefone: contatoTelefone || '',
    dados_contato_1: dadosContato1,
    dados_contato_2: dadosContato2,

    // Email para assinatura
    email_envio_contrato: emailEnvioContrato,
    email_cotitular_envio: emailCotitularEnvio,

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

  const partes = [];
  const identidade = [];

  if (nome) identidade.push(nome);
  if (nacionalidade) identidade.push(nacionalidade);
  if (estadoCivil) identidade.push(estadoCivil);
  if (identidade.length) identidade.push('empresário(a)');
  if (identidade.length) partes.push(identidade.join(', '));

  const enderecoPartes = [];
  if (rua) enderecoPartes.push(`residente na Rua ${rua}`);
  if (bairro) enderecoPartes.push(`Bairro ${bairro}`);
  if (numero) enderecoPartes.push(`nº ${numero}`);

  let cidadeUf = '';
  if (cidade) cidadeUf += cidade;
  if (uf) cidadeUf += (cidadeUf ? ' - ' : '') + uf;
  if (cidadeUf) enderecoPartes.push(cidadeUf);
  if (cep) enderecoPartes.push(`CEP: ${cep}`);

  if (enderecoPartes.length) partes.push(enderecoPartes.join(', '));

  const documentos = [];
  if (rg) documentos.push(`portador(a) da cédula de identidade RG de nº ${rg}`);

  const docUpper = String(docSelecao || '').trim().toUpperCase();
  const docNums = [];
  if (cpf) docNums.push({ tipo: 'CPF', valor: cpf });
  if (cnpj) docNums.push({ tipo: 'CNPJ', valor: cnpj });

  if (docUpper && docNums.length){
    const docPrincipal = docNums[0];
    documentos.push(`devidamente inscrito no ${docUpper} sob o nº ${docPrincipal.valor}`);
  } else {
    for (const doc of docNums){
      documentos.push(`devidamente inscrito no ${doc.tipo} sob o nº ${doc.valor}`);
    }
  }

  if (documentos.length) partes.push(documentos.join(', '));

  const contatoPartes = [];
  if (telefone) contatoPartes.push(`com telefone de nº ${telefone}`);
  if (email) contatoPartes.push(`com o seguinte endereço eletrônico: ${email}`);
  if (contatoPartes.length) partes.push(contatoPartes.join(' e '));

  if (!partes.length) return '';
  const texto = partes.join(', ').replace(/\s+,/g, ',').trim();
  return texto.endsWith('.') ? texto : `${texto}.`;
}

// Variáveis para Template Word
function montarVarsParaTemplateMarca(d, nowInfo){
  const valorTotalNum = onlyNumberBR(d.valor_total);
  const parcelaNum = parseInt(String(d.parcelas||'1'),10)||1;
  const valorParcela = parcelaNum>0 ? valorTotalNum/parcelaNum : 0;

  const dia = String(nowInfo.dia).padStart(2,'0');
  const mesNum = String(nowInfo.mes).padStart(2,'0');
  const ano = String(nowInfo.ano);
  const mesExtenso = monthNamePt(nowInfo.mes);

  const base = {
    // Identificação
    'Contratante 1': d.nome || '',
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

    // Tipo de marca do formulário
    'tipo de marca': d.tipo1 || '',

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

  // Preenche mais linhas de classes se existirem
  for (let i=5;i<30;i++){
    base[`marcas-espec_${i+1}`] = d.linhas_marcas_espec_1[i] || '';
    base[`marcas2-espec_${i-4}`] = d.linhas_marcas_espec_2[i-5] || '';
  }

  return base;
}

function montarVarsParaTemplateOutros(d, nowInfo){
  const valorTotalNum = onlyNumberBR(d.valor_total);
  const parcelaNum = parseInt(String(d.parcelas||'1'),10)||1;
  const valorParcela = parcelaNum>0 ? valorTotalNum/parcelaNum : 0;

  const dia = String(nowInfo.dia).padStart(2,'0');
  const mesNum = String(nowInfo.mes).padStart(2,'0');
  const ano = String(nowInfo.ano);
  const mesExtenso = monthNamePt(nowInfo.mes);

  const base = {
    // Identificação
    'Contratante 1': d.nome || '',
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
    'Quantidade depósitos – DESENHO INDUSTRIAL': d.qtd_desc.DI || d.qtd_desc['DESENHO INDUSTRIAL'] || '',
    'Descrição do serviço - DESENHO INDUSTRIAL': d.qtd_desc.DI ? '' : '',
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

// Assinantes: principal + empresa + cotitular quando houver
function montarSigners(d){
  const list = [];
  const emailPrincipal = d.email_envio_contrato || d.email || '';
  if (emailPrincipal) list.push({ email: emailPrincipal, name: d.nome || d.titulo || emailPrincipal, act:'1', foreign:'0', send_email:'1' });

  // cotitular
  if (d.email_cotitular_envio) list.push({ email: d.email_cotitular_envio, name: 'Cotitular', act:'1', foreign:'0', send_email:'1' });

  if (EMAIL_ASSINATURA_EMPRESA) list.push({ email: EMAIL_ASSINATURA_EMPRESA, name: 'Empresa', act:'1', foreign:'0', send_email:'1' });
  const seen={};
  return list.filter(s => (seen[s.email.toLowerCase()]? false : (seen[s.email.toLowerCase()]=true)));
}

/* =========================
 * Locks e preflight
 * =======================*/
const locks = new Set();
function acquireLock(key){ if (locks.has(key)) return false; locks.add(key); return true; }
function releaseLock(key){ locks.delete(key); }
async function preflightDNS(){ /* opcional: warmup */ }

/* =========================
 * D4Sign (validados)
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
async function cadastrarSignatarios(tokenAPI, cryptKey, uuidDocument, signers) {
  const base = 'https://secure.d4sign.com.br';
  const url = new URL(`/api/v1/documents/${uuidDocument}/createlist`, base);
  url.searchParams.set('tokenAPI', tokenAPI);
  url.searchParams.set('cryptKey', cryptKey);
  const body = { signers: signers.map(s => ({ email: s.email, name: s.name, act: s.act || '1', foreign: s.foreign || '0', send_email: s.send_email || '1' })) };
  const res = await fetchWithRetry(url.toString(), {
    method: 'POST', headers: { 'Accept': 'application/json', 'Content-Type': 'application/json' },
    body: JSON.stringify(body)
  }, { attempts: 5, baseDelayMs: 600, timeoutMs: 20000 });
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
  }, { attempts: 5, baseDelayMs: 600, timeoutMs: 20000 });
  const text = await res.text();
  if (!res.ok) {
    console.error('[ERRO D4SIGN sendtosigner]', res.status, text);
    throw new Error(`Falha ao enviar para assinatura: ${res.status}`);
  }
  return text;
}

/* =========================
 * Fase Pipefy
 * =======================*/
async function moveCardToPhaseSafe(cardId, phaseId){
  if (!phaseId) return;
  await gql(`mutation($input: MoveCardToPhaseInput!){
    moveCardToPhase(input:$input){ card{ id } }
  }`, { input: { card_id: Number(cardId), destination_phase_id: Number(phaseId) } }).catch(()=>{});
}

/* =========================
 * Rotas — VENDEDOR (UX)
 * =======================*/
// Página de revisão
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
      <div><div class="label">Nome</div><div>${d.nome||'-'}</div></div>
      <div><div class="label">Estado Civil</div><div>${d.estado_civil||'-'}</div></div>
      <div><div class="label">CPF (campo)</div><div>${d.cpf_campo||'-'}</div></div>
      <div><div class="label">CNPJ (campo)</div><div>${d.cnpj_campo||'-'}</div></div>
      <div><div class="label">RG</div><div>${d.rg||'-'}</div></div>
    </div>

    <h2>Contato</h2>
    <div class="grid">
      <div><div class="label">E-mail</div><div>${d.email||'-'}</div></div>
      <div><div class="label">Telefone</div><div>${d.telefone||'-'}</div></div>
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

// Gera o documento e mostra ações
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

    const isMarcaTemplate = d.templateToUse === TEMPLATE_UUID_CONTRATO;
    const add = isMarcaTemplate ? montarVarsParaTemplateMarca(d, nowInfo)
                                : montarVarsParaTemplateOutros(d, nowInfo);
    const signers = montarSigners(d);

    const uuidSafe = COFRES_UUIDS[d.vendedor] || DEFAULT_COFRE_UUID;
    if (!uuidSafe) throw new Error(`Cofre não configurado para vendedor: ${d.vendedor}`);

    const uuidDoc = await makeDocFromWordTemplate(
      D4SIGN_TOKEN,
      D4SIGN_CRYPT_KEY,
      uuidSafe,
      d.templateToUse,
      d.titulo || card.title,
      add
    );
    console.log(`[D4SIGN] Documento criado: ${uuidDoc}`);

    await new Promise(r=>setTimeout(r, 3000));
    try { await getDocumentStatus(D4SIGN_TOKEN, D4SIGN_CRYPT_KEY, uuidDoc); } catch {}

    await cadastrarSignatarios(D4SIGN_TOKEN, D4SIGN_CRYPT_KEY, uuidDoc, signers);

    await new Promise(r=>setTimeout(r, 2000));
    await moveCardToPhaseSafe(card.id, PHASE_ID_CONTRATO_ENVIADO);

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

// Download
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

// Enviar para assinatura
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
