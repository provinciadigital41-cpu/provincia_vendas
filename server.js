'use strict';

// server.js — Node 18+, CommonJS
const express = require('express');
const crypto = require('crypto');

const app = express();

// ===== Middlewares =====
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.set('trust proxy', true);

// Logger simples
app.use((req, _res, next) => {
  console.log(`[${new Date().toISOString()}] ${req.method} ${req.originalUrl} ua="${req.get('user-agent')}" ip=${req.ip}`);
  next();
});

// ===== Variáveis de ambiente =====
const env = process.env;
let {
  PORT,
  PIPE_API_KEY,
  PIPE_GRAPHQL_ENDPOINT,
  D4SIGN_CRYPT_KEY,
  D4SIGN_TOKEN,
  TEMPLATE_UUID_CONTRATO,

  NOVO_PIPE_ID,
  FASE_VISITA_ID,
  PIPEFY_FIELD_LINK_CONTRATO,

  // conectores (ids de campos no card)
  FIELD_ID_CONNECT_MARCA_NOME, // marcas_1 (Nome da marca + Contatos)
  FIELD_ID_CONNECT_CLASSE,     // marcas_2 (Marcas e serviços) - apoio
  FIELD_ID_CONNECT_CLASSES,    // conector direto para Classes INPI no card (se existir)

  // tabelas (ids de database)
  CONTACTS_TABLE_ID,           // opcional
  MARCAS_TABLE_ID,             // opcional
  CLASSES_TABLE_ID,            // tabela Classes INPI (ex.: 306521337)
  MARCAS2_TABLE_ID,

  // app pública
  PUBLIC_BASE_URL,
  PUBLIC_LINK_SECRET,

  // assinatura
  EMAIL_ASSINATURA_EMPRESA,

  // cofres (por nome)
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

  // fallback opcional
  DEFAULT_COFRE_UUID
} = env;

// Defaults
PORT = PORT || 3000;
PIPE_GRAPHQL_ENDPOINT = PIPE_GRAPHQL_ENDPOINT || 'https://api.pipefy.com/graphql';
FIELD_ID_CONNECT_MARCA_NOME = FIELD_ID_CONNECT_MARCA_NOME || 'marcas_1';
FIELD_ID_CONNECT_CLASSE     = FIELD_ID_CONNECT_CLASSE     || 'marcas_2';
CLASSES_TABLE_ID            = CLASSES_TABLE_ID            || '306521337';
MARCAS2_TABLE_ID            = MARCAS2_TABLE_ID            || '';
const FIELD_ID_LINKS_D4     = PIPEFY_FIELD_LINK_CONTRATO  || 'd4_contrato';

if (!PIPE_API_KEY) console.warn('[AVISO] PIPE_API_KEY não definido');
if (!D4SIGN_CRYPT_KEY || !D4SIGN_TOKEN) console.warn('[AVISO] D4SIGN_* não definidos');
if (!PUBLIC_BASE_URL || !PUBLIC_LINK_SECRET) console.warn('[AVISO] Defina PUBLIC_BASE_URL e PUBLIC_LINK_SECRET');

// Cofres mapeados por nome
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
};

// ===== Assinatura de links (HMAC) =====
function makeSignedURL(path, params) {
  const url = new URL(path, PUBLIC_BASE_URL);
  Object.entries(params || {}).forEach(([k, v]) => url.searchParams.set(k, String(v)));
  const payload = [...url.searchParams.entries()].map(([k, v]) => `${k}=${v}`).sort().join('&');
  const sig = crypto.createHmac('sha256', PUBLIC_LINK_SECRET).update(payload).digest('hex');
  url.searchParams.set('sig', sig);
  return url.toString();
}
function validateSignature(req) {
  const full = new URL(`${req.protocol}://${req.get('host')}${req.originalUrl}`);
  const gotSig = full.searchParams.get('sig') || '';
  full.searchParams.delete('sig');
  const payload = [...full.searchParams.entries()].map(([k, v]) => `${k}=${v}`).sort().join('&');
  const expected = crypto.createHmac('sha256', PUBLIC_LINK_SECRET).update(payload).digest('hex');
  return gotSig === expected;
}

// ===== Pipefy GraphQL =====
async function pipefyGraphQL(query, variables) {
  const res = await fetch(PIPE_GRAPHQL_ENDPOINT, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${PIPE_API_KEY}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({ query, variables }),
  });
  const json = await res.json();
  if (!res.ok || json.errors) {
    throw new Error(`Pipefy GraphQL error: ${res.status} ${JSON.stringify(json.errors || {})}`);
  }
  return json.data;
}
async function getCardFields(cardId) {
  const q = `
    query($id: ID!) {
      card(id: $id) {
        id
        title
        fields { name value field { id type } }
        current_phase { id name }
        pipe { id name }
      }
    }`;
  const data = await pipefyGraphQL(q, { id: cardId });
  return data.card;
}
async function updateCardField(cardId, fieldId, newValue) {
  const m = `
    mutation($input: UpdateCardFieldInput!) {
      updateCardField(input: $input) { card { id } }
    }`;
  await pipefyGraphQL(m, { input: { card_id: Number(cardId), field_id: fieldId, new_value: newValue } });
}
async function getTableRecord(recordId) {
  const q = `
    query($id: ID!) {
      table_record(id: $id) {
        id
        title
        record_fields { name value field { id type } }
      }
    }`;
  const data = await pipefyGraphQL(q, { id: recordId });
  return data.table_record;
}
async function listTableRecords(tableId, first = 100, after = null) {
  const q = `
    query($tableId: ID!, $first: Int!, $after: String) {
      table(id: $tableId) {
        id
        table_records(first: $first, after: $after) {
          edges {
            node {
              id
              title
              record_fields { name value field { id type } }
            }
          }
          pageInfo { hasNextPage endCursor }
        }
      }
    }`;
  const data = await pipefyGraphQL(q, { tableId, first, after });
  const edges = data?.table?.table_records?.edges || [];
  const pageInfo = data?.table?.table_records?.pageInfo || {};
  return { records: edges.map(e => e.node), pageInfo };
}

// ===== Utils =====
function toByIdFromCard(card) {
  const by = {};
  for (const f of card?.fields || []) if (f?.field?.id) by[f.field.id] = f.value;
  return by;
}
function normalizePhone(s) { return String(s || '').replace(/\D/g, ''); }
function onlyDigits(s) { return String(s||'').replace(/\D/g,''); }
function parseMaybeJsonArray(value) {
  try { return Array.isArray(value) ? value : JSON.parse(value); }
  catch { return value ? [String(value)] : []; }
}
function getValueByName(card, nameSubstr) {
  const target = String(nameSubstr).toLowerCase();
  const hit = (card.fields || []).find(f => String(f?.name || '').toLowerCase().includes(target));
  return hit?.value ?? '';
}
function getFirstByNames(card, names=[]) {
  for (const n of names) {
    const v = getValueByName(card, n);
    if (v) return v;
  }
  return '';
}
function parseNumberBR(v) {
  if (v == null) return NaN;
  const s = String(v).trim();
  if (!s) return NaN;
  const br = s.match(/^\d{1,3}(\.\d{3})*,\d{2}$/);
  if (br) return Number(s.replace(/\./g,'').replace(',','.'));
  const en = s.match(/^\d+(\.\d+)?$/);
  if (en) return Number(s);
  return Number(s.replace(/[^\d.,-]/g,'').replace(/\./g,'').replace(',','.'));
}
function toBRL(n) { return (n==null || isNaN(n)) ? '' : n.toLocaleString('pt-BR',{style:'currency',currency:'BRL'}); }
function pickParcelas(card) {
  const by = toByIdFromCard(card);
  let raw = by['sele_o_de_lista'] || by['numero_de_parcelas'] || '';
  if (!raw) raw = getFirstByNames(card, ['parcela', 'parcelas', 'nº parcelas', 'numero de parcelas', 'número de parcelas']);
  const m = String(raw||'').match(/(\d+)/);
  return m ? m[1] : '1';
}
function pickValorAssessoria(card) {
  const by = toByIdFromCard(card);
  let raw = by['valor_da_assessoria'] || by['valor_assessoria'] || '';
  if (!raw) raw = getFirstByNames(card, ['valor da assessoria', 'assessoria']);
  if (!raw) {
    const hit = (card.fields||[]).find(f => String(f?.field?.type||'').toLowerCase()==='currency');
    raw = hit?.value || '';
  }
  const n = parseNumberBR(raw);
  return isNaN(n) ? null : n;
}
function extractNameEmailPhoneFromRecord(record) {
  let nome = '', email = '', telefone = '';
  for (const f of record?.record_fields || []) {
    const t = (f?.field?.type || '').toLowerCase();
    const id = (f?.field?.id || '').toLowerCase();
    const label = (f?.name || '').toLowerCase();
    if (!nome && (id === 'nome_do_contato' || label.includes('nome'))) nome = String(f.value || '');
    if (!email && (t === 'email' || label.includes('email'))) email = String(f.value || '');
    if (!telefone && (t === 'phone' || label.includes('telefone') || label.includes('celular') || label.includes('whats'))) telefone = String(f.value || '');
  }
  return { nome, email, telefone };
}

// ===== Marca (nome/contato) via marcas_1 =====
async function resolveMarcaRecordFromCard(card) {
  const by = toByIdFromCard(card);
  const v = by[FIELD_ID_CONNECT_MARCA_NOME]; // 'marcas_1'
  const arr = Array.isArray(v) ? v : v ? [v] : [];
  if (!arr.length) return null;

  const first = String(arr[0]);

  if (/^\d+$/.test(first)) {
    try { return await getTableRecord(first); } catch { /* fallback título */ }
  }
  if (!MARCAS_TABLE_ID) return null;
  let after = null;
  for (let i = 0; i < 50; i++) {
    const { records, pageInfo } = await listTableRecords(MARCAS_TABLE_ID, 100, after);
    const hit = records.find(r => String(r.title || '').trim().toLowerCase() === first.trim().toLowerCase());
    if (hit) return hit;
    if (!pageInfo?.hasNextPage) break;
    after = pageInfo.endCursor || null;
  }
  return null;
}
async function resolveContatoFromMarcaRecord(marcaRecord) {
  const conn = (marcaRecord?.record_fields || []).find(f =>
    f?.field?.id === 'contatos' || String(f?.name||'').toLowerCase().includes('contato')
  );
  if (!conn) return { nome:'', email: '', telefone: '' };

  if (Array.isArray(conn.value)) {
    const first = conn.value[0];
    if (first) {
      const contato = await getTableRecord(first);
      return extractNameEmailPhoneFromRecord(contato);
    }
  }

  const mirrored = parseMaybeJsonArray(conn.value);
  let nome = '', email = '', telefone = '';
  for (const v of mirrored) {
    const s = String(v);
    if (!email && s.includes('@')) email = s;
    if (!telefone && normalizePhone(s).length >= 10) telefone = s;
    if (!nome && s && !s.includes('@') && normalizePhone(s).length < 10) nome = s;
  }
  return { nome, email, telefone };
}

// ===== Documento =====
function pickDocumento(card) {
  const prefer = ['cpf', 'cnpj', 'documento', 'doc', 'cpf/cnpj', 'cpf cnpj', 'cnpj/cpf'];
  for (const key of prefer) {
    const v = getFirstByNames(card, [key]);
    const digits = onlyDigits(v);
    if (digits.length === 11 || digits.length === 14) {
      return { tipo: digits.length === 11 ? 'CPF' : 'CNPJ', valor: v || '' };
    }
  }
  const by = toByIdFromCard(card);
  const cnpjStart = by['cnpj'] || getFirstByNames(card, ['cnpj']);
  if (cnpjStart) return { tipo: 'CNPJ', valor: cnpjStart };
  return { tipo: '', valor: '' };
}

// ===== Classe =====
async function resolveClasseFromLabelOnCard(card) {
  const f = (card.fields || []).find(ff => {
    const isConn = (ff?.field?.type === 'connector' || ff?.field?.type === 'table_connection');
    const idOk  = String(ff?.field?.id || '').toLowerCase() === 'classes_inpi';
    const lblOk = String(ff?.name || '').toLowerCase().includes('classes inpi');
    return isConn && (idOk || lblOk);
  });
  if (!f || !f.value) return '';

  let arr = [];
  try { arr = Array.isArray(f.value) ? f.value : JSON.parse(f.value); }
  catch { arr = [f.value]; }
  const first = arr && arr[0];
  if (!first) return '';

  if (/^\d+$/.test(String(first))) {
    try { const rec = await getTableRecord(String(first)); return rec?.title || ''; }
    catch { /* ignore */ }
  }
  return String(first || '');
}

async function resolveClasseFromCard(card, marcaRecordFallback) {
  // prioridade: conector "Classes INPI" no card
  const fromCard = await resolveClasseFromLabelOnCard(card);
  if (fromCard) return fromCard;

  const by = toByIdFromCard(card);

  // conector direto configurado por env
  if (FIELD_ID_CONNECT_CLASSES) {
    const v = by[FIELD_ID_CONNECT_CLASSES];
    const arr = Array.isArray(v) ? v : v ? [v] : [];
    if (arr.length) {
      const first = String(arr[0]);
      if (/^\d+$/.test(first)) {
        try { const rec = await getTableRecord(first); return rec?.title || ''; } catch {}
      }
      try { const mirrored = Array.isArray(v) ? v : JSON.parse(v); if (Array.isArray(mirrored) && mirrored.length) return String(mirrored[0] || ''); } catch {}
    }
  }

  // 2) Conector "Marcas e serviços" (marcas_2) → se vier TÍTULO, abre na tabela MARCAS2_TABLE_ID
const v2 = by[FIELD_ID_CONNECT_CLASSE];
const arr2 = Array.isArray(v2) ? v2 : v2 ? [v2] : [];
if (arr2.length) {
  const first = String(arr2[0]);

  // Caso 2.1: veio ID numérico (mantém como já estava)
  if (/^\d+$/.test(first)) {
    try {
      const rec = await getTableRecord(first);
      // procurar um campo que contenha "classe" e seguir o conector
      const classeField = (rec.record_fields || []).find(f =>
        String(f?.name || '').toLowerCase().includes('classe')
      );
      if (classeField?.value) {
        const val = classeField.value;
        if (Array.isArray(val)) {
          const id0 = String(val[0] || '');
          if (/^\d+$/.test(id0)) {
            const recClasse = await getTableRecord(id0);
            return recClasse?.title || '';
          }
        } else {
          try {
            const a = JSON.parse(val);
            if (Array.isArray(a) && a.length) return String(a[0] || '');
          } catch {
            if (val) return String(val);
          }
        }
      }
      if (rec?.title) return String(rec.title);
    } catch {}
  } else {
    // Caso 2.2: veio TÍTULO (espelho) -> procurar pelo título na tabela Marcas (Visita)
    if (MARCAS2_TABLE_ID) {
      let after = null;
      for (let i = 0; i < 50; i++) {
        const { records, pageInfo } = await listTableRecords(MARCAS2_TABLE_ID, 100, after);
        const rec = records.find(r => String(r.title || '').trim().toLowerCase() === first.trim().toLowerCase());
        if (rec) {
          // dentro do record, achar um campo que contenha “classe” e seguir o conector
          const classeField = (rec.record_fields || []).find(f =>
            String(f?.name || '').toLowerCase().includes('classe')
          );
          if (classeField?.value) {
            const val = classeField.value;
            if (Array.isArray(val)) {
              const id0 = String(val[0] || '');
              if (/^\d+$/.test(id0)) {
                const recClasse = await getTableRecord(id0);
                return recClasse?.title || '';
              }
            } else {
              try {
                const a = JSON.parse(val);
                if (Array.isArray(a) && a.length) return String(a[0] || '');
              } catch {
                if (val) return String(val);
              }
            }
          }
          // fallback: título do próprio record
          return String(rec.title || '');
        }
        if (!pageInfo?.hasNextPage) break;
        after = pageInfo.endCursor || null;
      }
    } else {
      // Sem MARCAS2_TABLE_ID, última tentativa: usar o texto que veio
      try {
        const mirrored = Array.isArray(v2) ? v2 : JSON.parse(v2);
        if (Array.isArray(mirrored) && mirrored.length) return String(mirrored[0] || '');
      } catch {}
    }
  }
}

// ===== Cofre do responsável =====
function stripDiacritics(s) { return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, ''); }
function normalizeName(s) { return stripDiacritics(String(s || '').trim()).toLowerCase(); }
function extractAssigneeNames(raw) {
  const out = [];
  const push = v => { if (v) out.push(String(v)); };
  const tryParse = v => { if (typeof v === 'string') { try { return JSON.parse(v); } catch { return v; } } return v; };
  const val = tryParse(raw);
  if (Array.isArray(val)) {
    for (const it of val) push(typeof it === 'string' ? it : (it?.name || it?.username || it?.email || it?.value));
  } else if (typeof val === 'object' && val) {
    push(val.name || val.username || val.email || val.value);
  } else if (typeof val === 'string') {
    const m = val.match(/^\s*\[.*\]\s*$/) ? tryParse(val) : null;
    if (m && Array.isArray(m)) m.forEach(x => push(typeof x === 'string' ? x : (x?.name || x?.email)));
    else push(val);
  }
  return [...new Set(out.filter(Boolean))];
}
function buildEmailMapFromEnv(env) {
  const map = {};
  Object.keys(env).forEach(k => {
    if (k.startsWith('COFRE_UUID_EMAIL_')) {
      const val = String(env[k] || '');
      const [email, uuid] = val.split(':').map(x => x?.trim());
      if (email && uuid) map[email.toLowerCase()] = uuid;
    }
  });
  return map;
}
function resolveCofreUuidByCard(card) {
  const by = toByIdFromCard(card);
  const candidatosBrutos = [];
  if (by['vendedor_respons_vel']) candidatosBrutos.push(by['vendedor_respons_vel']);
  if (by['respons_vel_5'])       candidatosBrutos.push(by['respons_vel_5']);
  if (by['representante'])       candidatosBrutos.push(by['representante']);
  const nomesOuEmails = candidatosBrutos.flatMap(extractAssigneeNames);

  const normKeys = Object.keys(COFRES_UUIDS || {}).reduce((acc, k) => { acc[normalizeName(k)] = COFRES_UUIDS[k]; return acc; }, {});
  for (const s of nomesOuEmails) { const n = normalizeName(s); if (normKeys[n]) return normKeys[n]; }

  const emailMap = buildEmailMapFromEnv(process.env);
  for (const s of nomesOuEmails) { const maybeEmail = String(s || '').toLowerCase(); if (maybeEmail.includes('@') && emailMap[maybeEmail]) return emailMap[maybeEmail]; }

  console.warn('[COFRE][NAO_ENCONTRADO]', { recebidos: nomesOuEmails, chavesCofres: Object.keys(COFRES_UUIDS || {}) });
  if (DEFAULT_COFRE_UUID) { console.warn('[COFRE][FALLBACK] usando DEFAULT_COFRE_UUID'); return DEFAULT_COFRE_UUID; }
  return null;
}

// ===== D4Sign =====
const D4_BASE = 'https://api.d4sign.com.br/v2';
async function d4Fetch(path, method, payload) {
  const url = `${D4_BASE}${path}?tokenAPI=${D4SIGN_TOKEN}&cryptKey=${D4SIGN_CRYPT_KEY}`;
  const res = await fetch(url, { method, headers: { 'Content-Type': 'application/json' }, body: payload ? JSON.stringify(payload) : undefined });
  if (!res.ok) throw new Error(`D4 error ${res.status}: ${await res.text()}`);
  return res.json();
}
async function d4CreateFromTemplateUUID({ templateUuid, fileName, variables, safeUuid }) {
  const payload = { template: templateUuid, name: fileName, data: variables, safe: safeUuid };
  const created = await d4Fetch('/documents/create', 'POST', payload);
  if (!created?.uuid && safeUuid) {
    const alt = await d4Fetch(`/documents/create/${safeUuid}`, 'POST', { template: templateUuid, name: fileName, data: variables });
    return alt;
  }
  return created;
}
async function d4AddSigners(documentKey, signers) {
  const payload = { signers: signers.map(s => ({ email: s.email, act: '1', foreign: '0', certificadoicpbr: '0', name: s.name || s.email })) };
  return d4Fetch(`/documents/${documentKey}/signers`, 'POST', payload);
}
async function d4SendToSign(documentKey, message = 'Contrato para assinatura') {
  return d4Fetch(`/documents/${documentKey}/sendto`, 'POST', { message, emails: [], workflow: '0' });
}

// ===== Tokens do template =====
function nowParts() { const d = new Date(); return { dd: String(d.getDate()).padStart(2, '0'), MM: String(d.getMonth() + 1).padStart(2, '0'), yyyy: String(d.getFullYear()) }; }

async function buildTemplateVariablesAsync(card) {
  const by = toByIdFromCard(card);      // <<< NÃO REMOVER
  const np = nowParts();

  // Marca & Nome da Marca
  const marcaRecord = await resolveMarcaRecordFromCard(card);
  const nomeMarca = card.title || '';

  // CONTATO (fase -> marcas_1 -> campos soltos)
  let contatoNome = '', contatoEmail = '', contatoTelefone = '';
  const contatoConnectorOnPhase = (card.fields || []).find(f =>
    (f.field?.type === 'connector' || f.field?.type === 'table_connection') &&
    String(f.name || '').toLowerCase().includes('contat')
  );
  if (contatoConnectorOnPhase?.value) {
    try {
      const arr = Array.isArray(contatoConnectorOnPhase.value) ? contatoConnectorOnPhase.value : JSON.parse(contatoConnectorOnPhase.value);
      const first = arr && arr[0];
      if (first && /^\d+$/.test(String(first))) {
        const rec = await getTableRecord(first);
        const ex = extractNameEmailPhoneFromRecord(rec);
        contatoNome = ex.nome || '';
        contatoEmail = ex.email || '';
        contatoTelefone = ex.telefone || '';
      } else if (Array.isArray(arr)) {
        const em = arr.find(s => String(s).includes('@'));
        const ph = arr.find(s => normalizePhone(s).length >= 10);
        contatoEmail = em ? String(em) : '';
        contatoTelefone = ph ? String(ph) : '';
      }
    } catch (e) { console.warn('[contatoConnectorOnPhase][parse-error]', e); }
  }
  if (marcaRecord && (!contatoNome || !contatoEmail || !contatoTelefone)) {
    const contato = await resolveContatoFromMarcaRecord(marcaRecord);
    contatoNome     = contatoNome     || contato.nome || '';
    contatoEmail    = contatoEmail    || contato.email || '';
    contatoTelefone = contatoTelefone || contato.telefone || '';
  }
  if (!contatoNome)    contatoNome    = getFirstByNames(card, ['nome do contato','contratante','responsável legal','responsavel legal']);
  if (!contatoEmail)   contatoEmail   = getFirstByNames(card, ['email','e-mail']);
  if (!contatoTelefone)contatoTelefone= getFirstByNames(card, ['telefone','celular','whats','whatsapp']);

  const contratante = contatoNome
    || by['r_social_ou_n_completo']
    || getFirstByNames(card, ['razão social','nome completo','nome do cliente'])
    || '';

  // Documento
  const doc = pickDocumento(card);
  const cpf  = doc.tipo === 'CPF'  ? doc.valor : '';
  const cnpj = doc.tipo === 'CNPJ' ? doc.valor : '';

  // Classe
  const classe = await resolveClasseFromCard(card, marcaRecord);

  // Endereço
  const rua    = by['rua']      || getFirstByNames(card, ['rua','logradouro','endereço']);
  const numero = by['n_mero_1'] || getFirstByNames(card, ['numero','número','nº']);
  const bairro = by['bairro']   || getFirstByNames(card, ['bairro']);
  const cidade = by['cidade']   || getFirstByNames(card, ['cidade','município']);
  const uf     = by['uf']       || getFirstByNames(card, ['uf','estado']);
  const cep    = by['cep']      || getFirstByNames(card, ['cep']);

  // Assessoria
  const nParc = pickParcelas(card);
  const valorAssess = pickValorAssessoria(card);
  const valorParcelaAssess = (valorAssess && Number(nParc)>0) ? toBRL(valorAssess/Number(nParc)) : '';

  // Pesquisa
  const pesquisa = by['pesquisa'] || getFirstByNames(card, ['pesquisa']);
  const valorPesquisa = (pesquisa === 'Isenta') ? 'R$ 00,00' : '';
  const formaPesquisa = (pesquisa === 'Isenta') ? '---' : '';
  const dataPesquisa  = (pesquisa === 'Isenta') ? '00/00/00' : '';

  const formaAssess = by['adiantamento_da_primeira_parcela'] || getFirstByNames(card, ['forma de pagamento','adiantamento']);
  const dataAssess  = by['data_da_venda_1'] || getFirstByNames(card, ['data da venda','data de pagamento']);

  return {
    contratante_1: contratante,
    cpf, cnpj,
    rg: by['rg'] || getFirstByNames(card, ['rg']) || '',

    rua, bairro, numero, nome_da_cidade: cidade, cidade, uf, cep,

    'E-mail': contatoEmail || '',
    telefone: contatoTelefone || '',

    nome_da_marca: nomeMarca,
    classe: String(classe || ''),

    numero_de_parcelas_da_assessoria: nParc,
    valor_da_parcela_da_assessoria: valorParcelaAssess,
    forma_de_pagamento_da_assessoria: formaAssess || '',
    data_de_pagamento_da_assessoria: dataAssess || '',

    valor_da_pesquisa: valorPesquisa,
    forma_de_pagamento_da_pesquisa: formaPesquisa,
    data_de_pagamento_da_pesquisa: dataPesquisa,

    valor_da_taxa: '',
    forma_de_pagamento_da_taxa: by['tipo_de_pagamento_benef_cio'] || getFirstByNames(card, ['tipo de pagamento']) || '',
    data_de_pagamento_da_taxa: '',

    dia: np.dd, mes: np.MM, ano: np.yyyy,

    TEMPLATE_UUID_CONTRATO: by['TEMPLATE_UUID_CONTRATO'] || TEMPLATE_UUID_CONTRATO || '',
  };
}

// ===== Rotas de fluxo =====
async function handleCriarLinkConfirmacao(req, res) {
  try {
    const cardId = req.method === 'GET' ? (req.query.cardId || req.query.card_id) : (req.body.cardId || req.body.card_id);
    if (!cardId) return res.status(400).json({ error: 'cardId é obrigatório' });

    const confirmUrl = makeSignedURL('/novo-pipe/confirmar', { cardId });
    console.log('[link-confirmacao]', { cardId, confirmUrl, PUBLIC_BASE_URL });

    await updateCardField(cardId, FIELD_ID_LINKS_D4, confirmUrl);
    return res.json({ ok: true, link: confirmUrl });
  } catch (e) {
    return res.status(500).json({ error: String(e.message || e) });
  }
}
app.post('/novo-pipe/criar-link-confirmacao', handleCriarLinkConfirmacao);
app.get('/novo-pipe/criar-link-confirmacao', handleCriarLinkConfirmacao);

app.get('/novo-pipe/confirmar', async (req, res) => {
  try {
    if (!validateSignature(req)) return res.status(401).send('assinatura inválida');
    const { cardId } = req.query;
    if (!cardId) return res.status(400).send('Faltou cardId');

    const card = await getCardFields(cardId);
    if (NOVO_PIPE_ID && String(card?.pipe?.id) !== String(NOVO_PIPE_ID)) return res.status(400).send('Card não pertence ao pipe configurado');
    if (FASE_VISITA_ID && String(card?.current_phase?.id) !== String(FASE_VISITA_ID)) return res.status(400).send('Card não está na fase esperada');

    const vars = await buildTemplateVariablesAsync(card);

    const resumo = `
      <h2>Confirmar dados do cliente</h2>
      <p><b>Contratante:</b> ${vars.contratante_1 || ''}</p>
      <p><b>CPF:</b> ${vars.cpf || ''} &nbsp; <b>CNPJ:</b> ${vars.cnpj || ''} &nbsp; <b>RG:</b> ${vars.rg || ''}</p>
      <p><b>Endereço:</b> ${vars.rua || ''}, ${vars.numero || ''} - ${vars.bairro || ''} - ${vars.cidade || ''}/${vars.uf || ''} - ${vars.cep || ''}</p>
      <p><b>Contato:</b> ${vars['E-mail'] || ''} &nbsp; ${vars.telefone || ''}</p>
      <p><b>Marca:</b> ${vars.nome_da_marca || ''} &nbsp; <b>Classe:</b> ${vars.classe || ''}</p>
      <p><b>Assessoria:</b> ${vars.numero_de_parcelas_da_assessoria || ''} x ${vars.valor_da_parcela_da_assessoria || ''}</p>
      <p><b>Pesquisa:</b> ${vars.valor_da_pesquisa} via ${vars.forma_de_pagamento_da_pesquisa} em ${vars.data_de_pagamento_da_pesquisa}</p>
      <form method="POST" action="/novo-pipe/gerar">
        <input type="hidden" name="cardId" value="${cardId}"/>
        <input type="hidden" name="sig" value="${new URL(req.originalUrl, PUBLIC_BASE_URL).searchParams.get('sig')||''}"/>
        <button type="submit">Gerar contrato</button>
      </form>
    `;
    res.send(`<!doctype html><html><body>${resumo}</body></html>`);
  } catch (e) {
    res.status(500).send(String(e.message || e));
  }
});

app.post('/novo-pipe/gerar', async (req, res) => {
  try {
    const { cardId, sig } = req.body;
    if (!cardId) return res.status(400).send('Faltou cardId');
    if (!sig) return res.status(401).send('assinatura ausente');

    const card = await getCardFields(cardId);
    const vars = await buildTemplateVariablesAsync(card);

    const templateUuid = vars.TEMPLATE_UUID_CONTRATO || TEMPLATE_UUID_CONTRATO;
    if (!templateUuid) throw new Error('TEMPLATE_UUID_CONTRATO não encontrado (card/env)');

    const safeUuid = resolveCofreUuidByCard(card);
    if (!safeUuid) throw new Error('Não foi possível resolver o cofre pelo responsável do card');

    const fileName = `Contrato_${card.title || card.id}.docx`;
    const created = await d4CreateFromTemplateUUID({ templateUuid, fileName, variables: vars, safeUuid });
    const documentKey = created.uuid || created.documentKey || created.key;
    if (!documentKey) throw new Error('Não foi possível obter documentKey do D4');

    const nextUrl = makeSignedURL(`/contratos/${documentKey}`, { cardId });
    await updateCardField(cardId, FIELD_ID_LINKS_D4, nextUrl);

    res.redirect(nextUrl);
  } catch (e) {
    res.status(500).send(String(e.message || e));
  }
});

app.get('/contratos/:documentKey', async (req, res) => {
  try {
    if (!validateSignature(req)) return res.status(401).send('assinatura inválida');
    const { documentKey } = req.params;
    const { cardId } = req.query;

    const downloadUrl = makeSignedURL(`/download/${documentKey}`, { cardId });
    const html = `
      <h2>Contrato gerado</h2>
      <p><b>DocumentKey:</b> ${documentKey}</p>
      <p><a href="${downloadUrl}">Baixar documento</a></p>
      <form method="POST" action="/contratos/${documentKey}/enviar-assinatura">
        <input type="hidden" name="cardId" value="${cardId || ''}"/>
        <input type="hidden" name="sig" value="${new URL(req.originalUrl, PUBLIC_BASE_URL).searchParams.get('sig')||''}"/>
        <button type="submit">Enviar para assinatura</button>
      </form>
    `;
    res.send(`<!doctype html><html><body>${html}</body></html>`);
  } catch (e) {
    res.status(500).send(String(e.message || e));
  }
});

app.post('/contratos/:documentKey/enviar-assinatura', async (req, res) => {
  try {
    const { documentKey } = req.params;
    const { cardId, sig } = req.body;
    if (!cardId) return res.status(400).send('Faltou cardId');
    if (!sig) return res.status(401).send('assinatura ausente');

    const card = await getCardFields(cardId);
    const vars = await buildTemplateVariablesAsync(card);

    const signers = [{ email: EMAIL_ASSINATURA_EMPRESA }];
    if (vars['E-mail']) signers.push({ email: vars['E-mail'] });

    await d4AddSigners(documentKey, signers);
    await d4SendToSign(documentKey, 'Contrato para assinatura');

    const back = makeSignedURL(`/contratos/${documentKey}`, { cardId });
    res.send(`<!doctype html><html><body>
      <h2>Contrato enviado para assinatura</h2>
      <p><b>Assinantes:</b> ${signers.map(s => s.email).join(', ')}</p>
      <p><a href="${back}">Voltar</a></p>
    </body></html>`);
  } catch (e) {
    res.status(500).send(String(e.message || e));
  }
});

// ===== Downloads =====
app.get('/download/:documentKey', async (req, res) => {
  try {
    if (!validateSignature(req)) return res.status(401).send('assinatura inválida');
    const { documentKey } = req.params;
    const url = `${D4_BASE}/documents/${documentKey}/download?tokenAPI=${D4SIGN_TOKEN}&cryptKey=${D4SIGN_CRYPT_KEY}`;
    const r = await fetch(url);
    if (!r.ok) throw new Error(`D4 download ${r.status}`);
    res.setHeader('Content-Disposition', `attachment; filename="Contrato_${documentKey}.pdf"`);
    res.setHeader('Content-Type', r.headers.get('content-type') || 'application/pdf');
    r.body.pipe(res);
  } catch (e) {
    res.status(500).json({ error: String(e.message || e) });
  }
});

// ===== Debug =====
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
app.get('/debug/card', async (req, res) => {
  try { const { cardId } = req.query; if (!cardId) return res.status(400).send('cardId obrigatório');
    const card = await getCardFields(cardId);
    res.json({ id: card.id, title: card.title, pipe: card.pipe, phase: card.current_phase,
      fields: (card.fields || []).map(f => ({ name: f.name, id: f.field?.id, type: f.field?.type, value: f.value })) });
  } catch (e) { res.status(500).json({ error: String(e.message || e) }); }
});
app.get('/debug/vars', async (req, res) => {
  try { const { cardId } = req.query; if (!cardId) return res.status(400).send('cardId obrigatório');
    const card = await getCardFields(cardId); const vars = await buildTemplateVariablesAsync(card); res.json(vars);
  } catch (e) { res.status(500).json({ error: String(e.message || e) }); }
});
app.get('/debug/marca-contato', async (req, res) => {
  try { const { cardId } = req.query; if (!cardId) return res.status(400).send('cardId obrigatório');
    const card = await getCardFields(cardId); const marcaRecord = await resolveMarcaRecordFromCard(card);
    let contato = { nome:'', email:'', telefone:'' }; if (marcaRecord) contato = await resolveContatoFromMarcaRecord(marcaRecord);
    const classe = await resolveClasseFromCard(card, marcaRecord);
    res.json({ cardTitle: card.title, marcaRecord, contato, classe });
  } catch (e) { res.status(500).json({ error: String(e.message || e) }); }
});
app.get('/debug/cofre', async (req, res) => {
  try { const { cardId } = req.query; if (!cardId) return res.status(400).send('cardId obrigatório');
    const card = await getCardFields(cardId);
    const by = toByIdFromCard(card);
    const candidatosBrutos = [];
    if (by['vendedor_respons_vel']) candidatosBrutos.push(by['vendedor_respons_vel']);
    if (by['respons_vel_5'])       candidatosBrutos.push(by['respons_vel_5']);
    if (by['representante'])       candidatosBrutos.push(by['representante']);
    const nomesOuEmails = candidatosBrutos.flatMap(extractAssigneeNames);
    const escolhido = resolveCofreUuidByCard(card);
    res.json({ pipe: card?.pipe, phase: card?.current_phase, candidatosBrutos, nomesOuEmails, cofreEscolhido: escolhido,
      possuiFallback: Boolean(DEFAULT_COFRE_UUID), chavesCofres: Object.keys(COFRES_UUIDS || {}) });
  } catch (e) { res.status(500).json({ error: String(e.message || e) }); }
});

// ===== Health =====
app.get('/health', (_req, res) => res.json({ ok: true }));

// ===== Start + rotas =====
app.listen(PORT, () => {
  console.log(`Servidor rodando na porta ${PORT}`);
  const list = [];
  app._router.stack.forEach(m => {
    if (m.route && m.route.path) {
      const methods = Object.keys(m.route.methods).map(x => x.toUpperCase()).join(',');
      list.push(`${methods} ${m.route.path}`);
    } else if (m.name === 'router' && m.handle.stack) {
      m.handle.stack.forEach(h => {
        const route = h.route;
        if (route) {
          const methods = Object.keys(route.methods).map(x => x.toUpperCase()).join(',');
          list.push(`${methods} ${route.path}`);
        }
      });
    }
  });
  console.log('[rotas-registradas]'); list.sort().forEach(r => console.log('  -', r));
});
