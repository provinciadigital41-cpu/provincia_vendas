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
app.use((req, res, next)=>{
  const t0 = Date.now();
  const ua = req.get('user-agent') || '';
  res.on('finish', ()=>{
    const dt = Date.now()-t0;
    console.log(`[${new Date().toISOString()}] ${req.method} ${req.originalUrl} ${res.statusCode} ${dt}ms ua="${ua}" ip=${req.ip}`);
  });
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

  // Pipe antigo (legado, leitura)
  OLD_PIPE_ID,
  OLD_PHASE_ID_GANHO,

  // Novo pipe — leads e vendas
  NOVO_PIPE_ID,
  PHASE_ID_CADASTRO_COMPLETO,
  PHASE_ID_CONTRATO_ENVIADO,
  PIPEFY_FIELD_LINK_CONTRATO,

  // Campos do pipe antigo (para leitura)
  OLD_FIELD_EMAIL,
  OLD_FIELD_DOCUMENTO,

  // Campos do novo pipe (para leitura)
  FIELD_EMAIL,
  FIELD_DOCUMENTO,
  FIELD_NOME,
  FIELD_TELEFONE,
  FIELD_VENDEDOR,

  // D4Sign
  D4SIGN_TOKEN,
  D4SIGN_CRYPT_KEY,
  TEMPLATE_UUID_CONTRATO,
  TEMPLATE_UUID_OUTROS,

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

  EMAIL_ASSINATURA_EMPRESA
} = process.env;

PORT = PORT || 3000;
PIPE_GRAPHQL_ENDPOINT = PIPE_GRAPHQL_ENDPOINT || 'https://api.pipefy.com/graphql';
OLD_PIPE_ID = Number(OLD_PIPE_ID || 0) || 302060216;
OLD_PHASE_ID_GANHO = Number(OLD_PHASE_ID_GANHO || 0) || 302060216;
NOVO_PIPE_ID = Number(NOVO_PIPE_ID || 0) || 303630289;
PHASE_ID_CADASTRO_COMPLETO = Number(PHASE_ID_CADASTRO_COMPLETO || 0) || 303630289;
PHASE_ID_CONTRATO_ENVIADO = Number(PHASE_ID_CONTRATO_ENVIADO || 0) || 303630289;
PIPEFY_FIELD_LINK_CONTRATO = PIPEFY_FIELD_LINK_CONTRATO || 'd4_contrato';

if (!PUBLIC_BASE_URL || !PUBLIC_LINK_SECRET) console.warn('[AVISO] Configure PUBLIC_BASE_URL e PUBLIC_LINK_SECRET');
if (!PIPE_API_KEY) console.warn('[AVISO] Configure PIPE_API_KEY');
if (!D4SIGN_TOKEN || !D4SIGN_CRYPT_KEY) console.warn('[AVISO] Configure D4SIGN_TOKEN e D4SIGN_CRYPT_KEY');

/* =========================
 * Utilidades
 * =======================*/
function delay(ms){ return new Promise(r=>setTimeout(r, ms)); }
function randomInt(min, max){ return Math.floor(Math.random()*(max-min+1))+min; }
function toBase64Url(str){ return Buffer.from(str, 'utf8').toString('base64url'); }
function fromBase64Url(str){ return Buffer.from(str, 'base64url').toString('utf8'); }

/**
 * fetchWithRetry
 *  - Pequeno wrapper com retry exponencial
 */
async function fetchWithRetry(url, options, {
  attempts = 3,
  baseDelayMs = 300,
  timeoutMs = 15000
} = {}){
  let lastErr;
  for (let i=0; i<attempts; i++){
    const controller = new AbortController();
    const id = setTimeout(()=>controller.abort(), timeoutMs);
    try {
      const res = await fetch(url, { ...options, signal: controller.signal });
      clearTimeout(id);
      if (!res.ok && res.status>=500 && res.status<600){
        throw new Error(`HTTP ${res.status}`);
      }
      return res;
    } catch (e){
      clearTimeout(id);
      lastErr = e;
      const isLast = i===attempts-1;
      if (isLast) break;
      const wait = baseDelayMs * Math.pow(2, i) + randomInt(0, 250);
      console.warn(`[fetchWithRetry] Erro "${e.message}", tentativa ${i+1}/${attempts}, aguardando ${wait}ms`);
      await delay(wait);
    }
  }
  throw lastErr;
}

/* =========================
 * Pipefy — GraphQL helper
 * =======================*/
async function gql(query, variables){
  const res = await fetchWithRetry(PIPE_GRAPHQL_ENDPOINT, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${PIPE_API_KEY}`
    },
    body: JSON.stringify({ query, variables })
  }, { attempts: 5, baseDelayMs: 400, timeoutMs: 20000 });

  const text = await res.text();
  let json;
  try { json = JSON.parse(text); } catch {
    console.error('[Pipefy gql] Resposta não JSON:', text.substring(0, 500));
    throw new Error('Pipefy GraphQL retornou uma resposta inválida');
  }
  if (json.errors && json.errors.length){
    console.error('[Pipefy gql] errors:', JSON.stringify(json.errors));
    throw new Error(json.errors[0]?.message || 'Erro GraphQL');
  }
  return json.data;
}

async function getCard(cardId){
  const data = await gql(`
    query($id: ID!){
      card(id:$id){
        id
        title
        createdAt
        pipe{ id name }
        current_phase{ id name }
        fields{
          name
          value
          field{
            id
            label
            type
          }
        }
      }
    }
  `, { id: Number(cardId) });
  if (!data || !data.card) throw new Error(`Card não encontrado: ${cardId}`);
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
  return f?.value;
}
function getFieldValue(card, fieldId){
  const f = (card.fields||[]).find(ff=> ff?.field?.id===fieldId);
  return f?.value;
}

/* =========================
 * Tokens de lead
 * =======================*/
function signLeadToken(payload){
  if (!PUBLIC_LINK_SECRET) throw new Error('PUBLIC_LINK_SECRET não configurado');
  const body = JSON.stringify(payload);
  const bodyB64 = toBase64Url(body);
  const hmac = crypto.createHmac('sha256', PUBLIC_LINK_SECRET).update(bodyB64).digest('base64url');
  return `${bodyB64}.${hmac}`;
}
function parseLeadToken(token){
  if (!PUBLIC_LINK_SECRET) throw new Error('PUBLIC_LINK_SECRET não configurado');
  const [bodyB64, sig] = String(token||'').split('.');
  if (!bodyB64 || !sig) throw new Error('Token inválido');
  const expected = crypto.createHmac('sha256', PUBLIC_LINK_SECRET).update(bodyB64).digest('base64url');
  if (!crypto.timingSafeEqual(Buffer.from(sig), Buffer.from(expected))) throw new Error('Token inválido');
  const body = fromBase64Url(bodyB64);
  const json = JSON.parse(body);
  if (!json.cardId) throw new Error('Token sem cardId');
  return json;
}

/* =========================
 * Mapeamento de cofres D4Sign
 * =======================*/
const DEFAULT_COFRE_UUID = COFRE_UUID_EDNA || COFRE_UUID_GREYCE;

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
async function registerWebhookForDocument(tokenAPI, cryptKey, uuidDocument, urlWebhook){
  const base = 'https://secure.d4sign.com.br';
  const url = new URL(`/api/v1/documents/${uuidDocument}/webhooks`, base);
  url.searchParams.set('tokenAPI', tokenAPI);
  url.searchParams.set('cryptKey', cryptKey);

  const body = { url: urlWebhook };

  const res = await fetchWithRetry(url.toString(), {
    method: 'POST',
    headers: { 'Accept': 'application/json', 'Content-Type': 'application/json' },
    body: JSON.stringify(body)
  }, { attempts: 5, baseDelayMs: 600, timeoutMs: 20000 });

  const text = await res.text();
  let json;
  try { json = JSON.parse(text); } catch { json = null; }

  if (!res.ok) {
    console.error('[ERRO D4SIGN webhook]', res.status, text.substring(0, 1000));
    throw new Error(`Falha ao cadastrar webhook: ${res.status}`);
  }

  return json;
}
async function cadastrarSignatarios(tokenAPI, cryptKey, uuidDocument, signers) {
  const base = 'https://secure.d4sign.com.br';
  const url = new URL(`/api/v1/documents/${uuidDocument}/createlist`, base);
  url.searchParams.set('tokenAPI', tokenAPI);
  url.searchParams.set('cryptKey', cryptKey);

  const list = signers.map(s=>({
    email: s.email,
    name: s.name,
    act: s.act || '1',
    foreign: s.foreign || '0',
    send_email: s.send_email || '1'
  }));

  const body = { dismember: '0', skip_email: '0', workflow: '0', signers: list };

  const res = await fetchWithRetry(url.toString(), {
    method: 'POST',
    headers: { 'Accept': 'application/json', 'Content-Type': 'application/json' },
    body: JSON.stringify(body)
  }, { attempts: 5, baseDelayMs: 600, timeoutMs: 20000 });

  const text = await res.text();
  let json; 
  try { json = JSON.parse(text); } catch { json = null; }

  if (!res.ok) {
    console.error('[ERRO D4SIGN signatários]', res.status, text.substring(0, 1000));
    throw new Error(`Falha ao cadastrar signatários: ${res.status}`);
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
    console.error('[ERRO D4SIGN sendToSigner]', res.status, text.substring(0, 1000));
    throw new Error(`Falha ao enviar para assinatura: ${res.status}`);
  }
  return text;
}

/* =========================
 * Montagem de dados do card e template
 * =======================*/
function montarDados(card){
  const by = toById(card);

  const vendedor = String(by[FIELD_VENDEDOR] || '').trim();
  const nome = String(by[FIELD_NOME] || getByName(card, 'Nome do cliente') || '').trim();
  const email = String(by[FIELD_EMAIL] || getByName(card, 'E mail') || '').trim();
  const telefone = String(by[FIELD_TELEFONE] || getByName(card, 'Telefone') || '').trim();
  const documento = String(by[FIELD_DOCUMENTO] || getByName(card, 'Documento') || '').trim();

  const templateToUse = by['tipo_servico']==='Outros Serviços' ? TEMPLATE_UUID_OUTROS : TEMPLATE_UUID_CONTRATO;
  const titulo = `Contrato ${nome || card.title || ''}`.trim();

  return {
    vendedor,
    nome,
    email,
    telefone,
    documento,
    templateToUse,
    titulo
  };
}
function montarVarsParaTemplateMarca(d, nowInfo){
  return {
    nomeCliente: d.nome,
    documentoCliente: d.documento,
    dataDia: String(nowInfo.dia).padStart(2,'0'),
    dataMes: String(nowInfo.mes).padStart(2,'0'),
    dataAno: String(nowInfo.ano)
  };
}
function montarVarsParaTemplateOutros(d, nowInfo){
  return {
    nomeCliente: d.nome,
    documentoCliente: d.documento,
    dataDia: String(nowInfo.dia).padStart(2,'0'),
    dataMes: String(nowInfo.mes).padStart(2,'0'),
    dataAno: String(nowInfo.ano)
  };
}
function montarSigners(d){
  const list = [];
  if (d.email) list.push({ email: d.email, name: d.nome || 'Cliente', act:'1', foreign:'0', send_email:'1' });
  if (EMAIL_ASSINATURA_EMPRESA) list.push({ email: EMAIL_ASSINATURA_EMPRESA, name: 'Empresa', act:'1', foreign:'0', send_email:'1' });
  const seen={}; return list.filter(s => (seen[s.email.toLowerCase()]? false : (seen[s.email.toLowerCase()]=true)));
}

/* =========================
 * Outros helpers Pipefy
 * =======================*/
async function moveCardToPhaseSafe(cardId, phaseId){
  if (!phaseId) return;
  await gql(`mutation($input: MoveCardToPhaseInput!){
    moveCardToPhase(input:$input){ card{ id } }
  }`, { input: { card_id: Number(cardId), destination_phase_id: Number(phaseId) } }).catch(()=>{});
}

async function findCardIdByD4Uuid(uuidDocument){
  const query = `
    query($pipeId: ID!, $uuid: String!){
      cards(
        pipe_id: $pipeId,
        search: { field_id: "d4_uuid_contrato", value: $uuid },
        first: 1
      ){
        edges{
          node{ id }
        }
      }
    }
  `;

  const data = await gql(query, {
    pipeId: Number(NOVO_PIPE_ID),
    uuid: uuidDocument
  });

  const edges = data && data.cards && data.cards.edges || [];
  if (!edges.length) return null;
  return edges[0].node.id;
}

async function anexarContratoAssinadoNoCard(cardId, downloadUrl, fileName){
  const value = JSON.stringify([
    {
      url: downloadUrl,
      filename: fileName || 'Contrato assinado.pdf'
    }
  ]);

  await updateCardField(cardId, 'contrato', value);
}

app.post('/d4sign/postback', async (req, res) => {
  try {
    const body = req.body || {};

    const uuidDocument = body.uuid;
    const typePost = String(body.type_post || body.typePost || '');

    if (!uuidDocument) {
      console.warn('[WEBHOOK D4SIGN] Sem uuid no corpo da requisição');
      return res.status(200).json({ ok: true });
    }

    if (typePost !== '1') {
      console.log('[WEBHOOK D4SIGN] Evento ignorado', uuidDocument, typePost, body.message);
      return res.status(200).json({ ok: true });
    }

    console.log('[WEBHOOK D4SIGN] Documento finalizado', uuidDocument, body.message);

    const cardId = await findCardIdByD4Uuid(uuidDocument);
    if (!cardId){
      console.warn('[WEBHOOK D4SIGN] Nenhum card encontrado para uuid', uuidDocument);
      return res.status(200).json({ ok: true });
    }

    try {
      await moveCardToPhaseSafe(cardId, 339299694);
    } catch (e) {
      console.error('[WEBHOOK D4SIGN] Falha ao mover card para fase de contrato assinado', e.message || e);
    }

    let downloadUrl, fileName;
    try {
      const info = await getDownloadUrl(D4SIGN_TOKEN, D4SIGN_CRYPT_KEY, uuidDocument, {
        type: 'PDF',
        language: 'pt'
      });
      downloadUrl = info.url;
      fileName = info.name;
    } catch (e) {
      console.error('[WEBHOOK D4SIGN] Falha ao gerar URL do contrato assinado', e.message || e);
      return res.status(200).json({ ok: true });
    }

    try {
      await anexarContratoAssinadoNoCard(cardId, downloadUrl, fileName);
    } catch (e) {
      console.error('[WEBHOOK D4SIGN] Falha ao anexar contrato assinado', e.message || e);
    }

    return res.status(200).json({ ok: true });
  } catch (e) {
    console.error('[WEBHOOK D4SIGN] Erro geral no postback', e.message || e);
    return res.status(200).json({ ok: true });
  }
});

/* =========================
 * Rotas — VENDEDOR (UX)
 * =======================*/
app.get('/lead/:token', async (req, res) => {
  try {
    const { cardId } = parseLeadToken(req.params.token);
    const card = await getCard(cardId);
    const by = toById(card);

    const nome = String(by[FIELD_NOME] || getByName(card, 'Nome do cliente') || '').trim();
    const email = String(by[FIELD_EMAIL] || getByName(card, 'E mail') || '').trim();

    const html = `
<!doctype html><meta charset="utf-8"><title>Assinatura de contrato</title>
<style>
  body{font-family:system-ui;display:grid;place-items:center;min-height:100vh;background:#f7f7f7;color:#111;margin:0}
  .box{background:#fff;padding:24px;border-radius:14px;box-shadow:0 4px 16px rgba(0,0,0,.08);max-width:640px;width:92%}
  h1{margin:0 0 12px;font-size:1.4rem}
  p{margin:6px 0}
  .btn{display:inline-block;margin-top:16px;padding:10px 14px;border-radius:8px;background:#111;color:#fff;text-decoration:none;font-weight:600}
  .muted{color:#666}
</style>
<div class="box">
  <h1>Olá, ${nome || 'tudo bem'}?</h1>
  <p class="muted">Aqui você consegue gerar o seu contrato digital para assinatura.</p>
  ${email ? `<p>E-mail cadastrado: <strong>${email}</strong></p>` : ''}
  <p><a class="btn" href="/lead/${encodeURIComponent(req.params.token)}/contrato">Gerar contrato</a></p>
</div>`;
    res.send(html);
  } catch (e) {
    console.error('[GET /lead/:token] erro', e);
    res.status(400).send('Link inválido ou expirado.');
  }
});

app.get('/lead/:token/contrato', async (req, res) => {
  try {
    const { cardId } = parseLeadToken(req.params.token);
    const card = await getCard(cardId);
    const by = toById(card);

    const html = `
<!doctype html><meta charset="utf-8"><title>Gerar contrato</title>
<style>
  body{font-family:system-ui;display:grid;place-items:center;min-height:100vh;background:#f7f7f7;color:#111;margin:0}
  .box{background:#fff;padding:24px;border-radius:14px;box-shadow:0 4px 16px rgba(0,0,0,.08);max-width:640px;width:92%}
  h1{margin:0 0 12px;font-size:1.4rem}
  p{margin:6px 0}
  label{display:block;margin-top:8px;font-size:.9rem}
  input{width:100%;padding:8px;border-radius:6px;border:1px solid #ccc;font-size:.95rem}
  .btn{display:inline-block;margin-top:16px;padding:10px 14px;border-radius:8px;background:#111;color:#fff;text-decoration:none;font-weight:600;border:0;cursor:pointer}
  .muted{color:#666}
</style>
<div class="box">
  <h1>Gerar contrato</h1>
  <p class="muted">Confirme os dados para geração do contrato.</p>
  <form method="POST" action="/lead/${encodeURIComponent(req.params.token)}/generate">
    <label>Nome do cliente
      <input type="text" name="nome" value="${String(by[FIELD_NOME] || '').replace(/"/g,'&quot;')}">
    </label>
    <label>E-mail
      <input type="email" name="email" value="${String(by[FIELD_EMAIL] || '').replace(/"/g,'&quot;')}">
    </label>
    <label>Telefone
      <input type="text" name="telefone" value="${String(by[FIELD_TELEFONE] || '').replace(/"/g,'&quot;')}">
    </label>
    <button class="btn" type="submit">Gerar e enviar para assinatura</button>
  </form>
</div>`;
    res.send(html);
  } catch (e) {
    console.error('[GET /lead/:token/contrato] erro', e);
    res.status(400).send('Link inválido ou expirado.');
  }
});

app.post('/lead/:token/generate', async (req, res) => {
  const { cardId } = parseLeadToken(req.params.token);
  const lockKey = `lead:${cardId}`;
  if (!acquireLock(lockKey)) return res.status(200).send('Processando, tente novamente em instantes.');

  try {
    preflightDNS().catch(()=>{});

    const card = await getCard(cardId);

    // Atualiza campos básicos se vierem no POST
    if (req.body && Object.keys(req.body).length){
      if (req.body.nome) await updateCardField(cardId, FIELD_NOME, req.body.nome);
      if (req.body.email) await updateCardField(cardId, FIELD_EMAIL, req.body.email);
      if (req.body.telefone) await updateCardField(cardId, FIELD_TELEFONE, req.body.telefone);
    }

    // Recarrega card para garantir dados atualizados
    const cardFresh = await getCard(cardId);
    const d = montarDados(cardFresh);

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

    try {
      await registerWebhookForDocument(
        D4SIGN_TOKEN,
        D4SIGN_CRYPT_KEY,
        uuidDoc,
        `${PUBLIC_BASE_URL}/d4sign/postback`
      );
    } catch (e) {
      console.error('[AVISO] Não foi possível cadastrar webhook no documento', e.message || e);
    }

    try {
      await updateCardField(card.id, 'd4_uuid_contrato', uuidDoc);
    } catch (e) {
      console.error('[AVISO] Não foi possível salvar d4_uuid_contrato no card', e.message || e);
    }

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
    <a class="btn" href="/lead/${encodeURIComponent(token)}">Voltar</a>
  </div>
</div>`;
    res.send(html);
  } catch (e) {
    console.error('[POST /lead/:token/generate] erro', e);
    releaseLock(lockKey);
    res.status(500).send('Erro ao gerar contrato.');
  }
});

app.get('/lead/:token/doc/:uuid/download', async (req, res) => {
  try {
    const { cardId } = parseLeadToken(req.params.token);
    await getCard(cardId);
    const info = await getDownloadUrl(D4SIGN_TOKEN, D4SIGN_CRYPT_KEY, req.params.uuid, { type:'PDF', language:'pt' });
    res.redirect(info.url);
  } catch (e) {
    console.error('[GET /lead/:token/doc/:uuid/download] erro', e);
    res.status(400).send('Não foi possível gerar o download deste documento.');
  }
});

app.post('/lead/:token/doc/:uuid/send', async (req, res) => {
  try {
    const { cardId } = parseLeadToken(req.params.token);
    await getCard(cardId);

    await sendToSigner(D4SIGN_TOKEN, D4SIGN_CRYPT_KEY, req.params.uuid, {
      message: 'Contrato para assinatura',
      skip_email: '0',
      workflow: '0'
    });

    const html = `
<!doctype html><meta charset="utf-8"><title>Contrato enviado</title>
<style>
  body{font-family:system-ui;display:grid;place-items:center;min-height:100vh;background:#f7f7f7;color:#111;margin:0}
  .box{background:#fff;padding:24px;border-radius:14px;box-shadow:0 4px 16px rgba(0,0,0,.08);max-width:640px;width:92%}
  h2{margin:0 0 12px}
  p{margin:6px 0}
  .btn{display:inline-block;padding:12px 16px;border-radius:10px;text-decoration:none;border:0;background:#111;color:#fff;font-weight:600}
  .muted{color:#666}
</style>
<div class="box">
  <h2>Contrato enviado para assinatura</h2>
  <p class="muted">As partes envolvidas foram notificadas.</p>
  <p><a href="/lead/${encodeURIComponent(req.params.token)}" class="btn">Voltar</a></p>
</div>`;
    res.send(html);
  } catch (e) {
    console.error('[POST /lead/:token/doc/:uuid/send] erro', e);
    res.status(500).send('Erro ao enviar para assinatura.');
  }
});

/* =========================
 * Rotas de utilidade
 * =======================*/
app.get('/healthz', (req, res)=>res.json({ ok:true, uptime:process.uptime() }));
app.get('/', (req, res)=>res.send('OK'));

// Dump de rotas
app.get('/debug/routes', (req, res)=>{
  const routes = [];
  app._router.stack.forEach(m=>{
    if (m.route && m.route.path){
      const methods = Object.keys(m.route.methods).map(x=>x.toUpperCase()).join(',');
      routes.push({ methods, path:m.route.path });
    } else if (m.name==='router' && m.handle?.stack){
      m.handle.stack.forEach(h=>{
        const route = h.route;
        if (route){
          const methods = Object.keys(route.methods).map(x=>x.toUpperCase()).join(',');
          routes.push({ methods, path:route.path });
        }
      });
    }
  });
  routes.sort((a,b)=>String(a.path).localeCompare(String(b.path)));
  res.json(routes);
});

// Log de rotas no startup
app.once('listening', ()=>{
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

const server = app.listen(PORT, ()=>{
  console.log(`Servidor ouvindo na porta ${PORT}`);
  app.emit('listening');
});

module.exports = { app, server };
