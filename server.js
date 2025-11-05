// server.js — Node 18+, CommonJS (usa fetch nativo do Node)
const express = require('express');
const crypto = require('crypto');

const app = express();
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// ======================
// 1) Variáveis de ambiente
// ======================
const {
  PORT = 3000,

  // Pipefy
  PIPE_API_KEY,
  PIPE_GRAPHQL_ENDPOINT = 'https://api.pipefy.com/graphql',

  // D4Sign
  D4SIGN_CRYPT_KEY,
  D4SIGN_TOKEN,
  TEMPLATE_UUID_CONTRATO,          // fallback global (opcional)

  // Pipe & fase (opcionais)
  NOVO_PIPE_ID,                    // se quiser validar pipe do card
  FASE_VISITA_ID,                  // se quiser validar fase do card
  PHASE_ID_CONTRATO_ENVIADO,       // caso queira mover depois

  // Campo do card p/ gravar o link interno
  PIPEFY_FIELD_LINK_CONTRATO,      // se não vier, usamos 'd4_contrato'

  // Conectores / Tabelas
  FIELD_ID_CONNECT_MARCAS = 'marcas_1', // conector de Marca no CARD
  CONTACTS_TABLE_ID,                    // ex.: "306505297"
  MARCAS_TABLE_ID,                      // ex.: "306555991" (opcional; usamos quando precisarmos varrer marcas por título)

  // Remetente e domínios públicos da SUA app
  EMAIL_ASSINATURA_EMPRESA,
  PUBLIC_BASE_URL,                      // ex.: https://contratos.suaempresa.com
  PUBLIC_LINK_SECRET,                   // HMAC p/ assinar links

  // Cofres por responsável
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
} = process.env;

if (!PIPE_API_KEY) console.warn('[AVISO] PIPE_API_KEY não definido');
if (!D4SIGN_CRYPT_KEY || !D4SIGN_TOKEN) console.warn('[AVISO] D4SIGN_* não definidos');
if (!PUBLIC_BASE_URL || !PUBLIC_LINK_SECRET) console.warn('[AVISO] Defina PUBLIC_BASE_URL e PUBLIC_LINK_SECRET');

// campo onde salvamos o link interno no card
const FIELD_ID_LINKS_D4 = PIPEFY_FIELD_LINK_CONTRATO || 'd4_contrato';

// mapeamento de cofres pelo nome do responsável
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

// ======================
/* 2) Helpers de assinatura de links (HMAC) */
// ======================
function makeSignedURL(path, params) {
  const url = new URL(path, PUBLIC_BASE_URL);
  Object.entries(params || {}).forEach(([k, v]) => url.searchParams.set(k, String(v)));
  const payload = [...url.searchParams.entries()].map(([k, v]) => `${k}=${v}`).sort().join('&');
  const sig = crypto.createHmac('sha256', PUBLIC_LINK_SECRET).update(payload).digest('hex');
  url.searchParams.set('sig', sig);
  return url.toString();
}

function validateSignature(req) {
  const base = `${req.protocol}://${req.get('host')}${req.path}`;
  const url = new URL(base);
  // reconstruir com as query params originais
  const original = new URL(`${req.protocol}://${req.get('host')}${req.originalUrl}`);
  const gotSig = original.searchParams.get('sig') || '';
  original.searchParams.delete('sig');
  const payload = [...original.searchParams.entries()].map(([k, v]) => `${k}=${v}`).sort().join('&');
  const expected = crypto.createHmac('sha256', PUBLIC_LINK_SECRET).update(payload).digest('hex');
  return gotSig === expected;
}

// ======================
// 3) Pipefy GraphQL helpers
// ======================
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
  const query = `
    query($id: ID!) {
      card(id: $id) {
        id
        title
        fields { name value field { id type } }
        current_phase { id name }
        pipe { id name }
      }
    }
  `;
  const data = await pipefyGraphQL(query, { id: cardId });
  return data.card;
}

async function updateCardField(cardId, fieldId, newValue) {
  const mutation = `
    mutation($input: UpdateCardFieldInput!) {
      updateCardField(input: $input) { card { id } }
    }
  `;
  await pipefyGraphQL(mutation, {
    input: {
      card_id: Number(cardId),
      field_id: fieldId,
      new_value: newValue,
    },
  });
}

async function getTableRecord(recordId) {
  const q = `
    query($id: ID!) {
      table_record(id: $id) {
        id
        title
        record_fields { name value field { id type } }
      }
    }
  `;
  const data = await pipefyGraphQL(q, { id: recordId });
  return data.table_record;
}

// pagina todos registros de uma tabela (sem argumento search)
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

function toByIdFromCard(card) {
  const by = {};
  for (const f of card?.fields || []) if (f?.field?.id) by[f.field.id] = f.value;
  return by;
}

function normalizePhone(s) { return String(s || '').replace(/\D/g, ''); }

function parseMaybeJsonArray(value) {
  try { return Array.isArray(value) ? value : JSON.parse(value); }
  catch { return value ? [String(value)] : []; }
}

// ======================
// 4) Resolver de Marca & Contato
// ======================

// Busca record de Marca a partir do valor do conector do CARD.
// Se o conector do card tiver IDs, usa direto; se tiver títulos, lista a tabela de Marcas e encontra por title.
async function resolveMarcaRecordFromCard(card) {
  const by = toByIdFromCard(card);
  const v = by[FIELD_ID_CONNECT_MARCAS]; // pode ser array de IDs OU array de títulos
  const arr = Array.isArray(v) ? v : v ? [v] : [];
  if (!arr.length) return null;

  // heurística: se parecer número/id (só dígitos), assume que é ID
  const first = String(arr[0]);
  const numericOnly = first.replace(/\D/g, '');
  if (numericOnly && numericOnly.length >= 6) {
    // provavelmente ID de record
    try {
      const rec = await getTableRecord(first);
      return rec;
    } catch (_) { /* cai para busca por título */ }
  }

  // caso contrário, trata como título (ex.: "BUCANEIROS MOTO CLUBE")
  if (!MARCAS_TABLE_ID) return null; // precisa do id da tabela de Marcas para varrer
  let after = null;
  for (let i = 0; i < 50; i++) { // até ~5k registros se first=100
    const { records, pageInfo } = await listTableRecords(MARCAS_TABLE_ID, 100, after);
    const hit = records.find(r => String(r.title || '').trim().toLowerCase() === first.trim().toLowerCase());
    if (hit) return hit;
    if (!pageInfo?.hasNextPage) break;
    after = pageInfo.endCursor || null;
  }
  return null;
}

// a partir do record de Marca, extrai Contato (email/telefone) via conector "contatos":
// - Se conector devolver IDs, usamos direto.
// - Se devolver espelho (telefones/emails como string JSON), tentamos extrair; se faltar, varremos tabela de contatos para achar o record por telefone/email.
async function resolveContatoFromMarcaRecord(marcaRecord) {
  const fieldConn = (marcaRecord?.record_fields || []).find(f => f?.field?.id === 'contatos');
  if (!fieldConn) return { email: '', telefone: '' };

  // caso 1: IDs
  if (Array.isArray(fieldConn.value)) {
    const first = fieldConn.value[0];
    if (first) {
      const contato = await getTableRecord(first);
      return extractEmailPhoneFromRecord(contato);
    }
  }

  // caso 2: espelho (string JSON com valores)
  const mirrored = parseMaybeJsonArray(fieldConn.value);
  // tenta achar direto
  let email = mirrored.find(x => String(x).includes('@')) || '';
  let telefone = mirrored.find(x => normalizePhone(x).length >= 10) || '';

  // se ambos vieram, pronto
  if (email && telefone) return { email: String(email), telefone: String(telefone) };

  // senão, se tivermos tabela de contatos, varre para achar por telefone/email
  if (CONTACTS_TABLE_ID) {
    const found = await findContatoInTable({ phoneCandidate: telefone, emailCandidate: email });
    if (found?.record) {
      const fromRecord = extractEmailPhoneFromRecord(found.record);
      return { email: fromRecord.email || email, telefone: fromRecord.telefone || telefone };
    }
  }

  return { email: String(email || ''), telefone: String(telefone || '') };
}

function extractEmailPhoneFromRecord(record) {
  let email = '', telefone = '';
  for (const f of record?.record_fields || []) {
    const t = (f?.field?.type || '').toLowerCase();
    const label = (f?.name || '').toLowerCase();
    if (!email && (t === 'email' || label.includes('email'))) email = String(f.value || '');
    if (!telefone && (t === 'phone' || t === 'tel' || label.includes('telefone') || label.includes('celular') || label.includes('whats'))) {
      telefone = String(f.value || '');
    }
  }
  return { email, telefone };
}

// varre tabela de contatos e tenta match por telefone (terminação) ou e-mail exato
async function findContatoInTable({ phoneCandidate, emailCandidate, maxPages = 100, pageSize = 100 }) {
  let after = null;
  const wantedPhone = normalizePhone(phoneCandidate);
  const wantedEmail = (emailCandidate || '').toLowerCase();

  for (let i = 0; i < maxPages; i++) {
    const { records, pageInfo } = await listTableRecords(CONTACTS_TABLE_ID, pageSize, after);
    for (const rec of records) {
      const { email, telefone } = extractEmailPhoneFromRecord(rec);
      const nodePhone = normalizePhone(telefone);
      const nodeEmail = (email || '').toLowerCase();

      const phoneMatch = wantedPhone && nodePhone && (nodePhone.endsWith(wantedPhone) || wantedPhone.endsWith(nodePhone));
      const emailMatch = wantedEmail && nodeEmail && nodeEmail === wantedEmail;

      if (phoneMatch || emailMatch) {
        return { record: rec, email, telefone };
      }
    }
    if (!pageInfo?.hasNextPage) break;
    after = pageInfo.endCursor || null;
  }
  return null;
}

// ======================
// 5) Montagem dos tokens do template (async)
// ======================
function nowParts() {
  const d = new Date();
  return {
    dd: String(d.getDate()).padStart(2, '0'),
    MM: String(d.getMonth() + 1).padStart(2, '0'),
    yyyy: String(d.getFullYear()),
  };
}

async function buildTemplateVariablesAsync(card) {
  const by = toByIdFromCard(card);
  const np = nowParts();

  // Nome da marca = título do card
  const marcaRecord = await resolveMarcaRecordFromCard(card);
  const nomeMarca = card.title || '';
  let contatoEmail = '', contatoTelefone = '', classe = '';

  // Classe por label (opcional) + contato por conector
  if (marcaRecord) {
    // tenta "classe"
    const classeField = (marcaRecord.record_fields || []).find(f => String(f?.name || '').toLowerCase().includes('classe'));
    classe = classeField?.value || '';
    const contato = await resolveContatoFromMarcaRecord(marcaRecord);
    contatoEmail = contato.email || '';
    contatoTelefone = contato.telefone || '';
  }

  // parcelas
  const rawSel = by['sele_o_de_lista']; // ex.: "6x"
  const nParc = rawSel ? (String(rawSel).match(/(\d+)\s*x/i)?.[1] || '1') : '1';
  const totalAssess = Number(by['valor_da_assessoria'] || 0);
  const valorParcelaAssess = nParc && !isNaN(totalAssess) ? (totalAssess / Number(nParc)).toFixed(2) : '';

  // pesquisa (radio: Paga/Isenta)
  const pesquisa = by['pesquisa'];
  const valorPesquisa = (pesquisa === 'Isenta') ? 'R$ 00,00' : '';
  const formaPesquisa = (pesquisa === 'Isenta') ? '---' : '';
  const dataPesquisa  = (pesquisa === 'Isenta') ? '00/00/00' : '';

  return {
    // Identificação / endereço (Start Form)
    contratante_1: by['r_social_ou_n_completo'] || '',
    cpf: by['cpf'] || '',
    rg: by['rg'] || '',

    rua: by['rua'] || '',
    bairro: by['bairro'] || '',
    numero: by['n_mero_1'] || '',
    nome_da_cidade: by['cidade'] || '',
    cidade: by['cidade'] || '',
    uf: by['uf'] || '',
    cep: by['cep'] || '',

    // Contato (derivado do DB conectado à Marca)
    'E-mail': contatoEmail,
    telefone: contatoTelefone,

    // Marca
    nome_da_marca: nomeMarca,
    classe: String(classe || ''),

    // Assessoria
    numero_de_parcelas_da_assessoria: nParc,
    valor_da_parcela_da_assessoria: valorParcelaAssess,
    forma_de_pagamento_da_assessoria: by['adiantamento_da_primeira_parcela'] || '',
    data_de_pagamento_da_assessoria: by['data_da_venda_1'] || '',

    // Pesquisa (Isenta por padrão)
    valor_da_pesquisa: valorPesquisa,
    forma_de_pagamento_da_pesquisa: formaPesquisa,
    data_de_pagamento_da_pesquisa: dataPesquisa,

    // Taxa (placeholders por enquanto; ajuste quando tiver os campos)
    valor_da_taxa: '',
    forma_de_pagamento_da_taxa: by['tipo_de_pagamento_benef_cio'] || '',
    data_de_pagamento_da_taxa: '',

    // Datas do rodapé
    dia: np.dd,
    mes: np.MM,
    ano: np.yyyy,

    // Auxiliares
    TEMPLATE_UUID_CONTRATO: by['TEMPLATE_UUID_CONTRATO'] || TEMPLATE_UUID_CONTRATO || '',
  };
}

// ======================
// 6) Resolver do Cofre por responsável do card
// ======================
function resolveCofreUuidByCard(card) {
  const by = toByIdFromCard(card);
  const candidatos = [];

  // ids de campos do seu pipe que podem trazer o responsável
  // ajuste aqui se houver outro id:
  if (by['respons_vel_5']) candidatos.push(by['respons_vel_5']);
  if (by['representante']) candidatos.push(by['representante']);

  const nomes = []
    .concat(candidatos)
    .filter(Boolean)
    .flatMap(v => Array.isArray(v) ? v : [v])
    .map(v => (typeof v === 'string' ? v : (v?.name || v?.email || '')))
    .filter(Boolean);

  for (const nome of nomes) {
    if (COFRES_UUIDS[nome]) return COFRES_UUIDS[nome];
    const k = Object.keys(COFRES_UUIDS).find(key => key.toLowerCase() === String(nome).toLowerCase());
    if (k) return COFRES_UUIDS[k];
  }
  return null;
}

// ======================
// 7) D4Sign helpers
// ======================
const D4_BASE = 'https://api.d4sign.com.br/v2';

async function d4Fetch(path, method, payload) {
  const url = `${D4_BASE}${path}?tokenAPI=${D4SIGN_TOKEN}&cryptKey=${D4SIGN_CRYPT_KEY}`;
  const res = await fetch(url, {
    method,
    headers: { 'Content-Type': 'application/json' },
    body: payload ? JSON.stringify(payload) : undefined,
  });
  if (!res.ok) throw new Error(`D4 error ${res.status}: ${await res.text()}`);
  return res.json();
}

// Cria documento a partir do UUID do template, dentro do cofre (safe)
async function d4CreateFromTemplateUUID({ templateUuid, fileName, variables, safeUuid }) {
  const payload = { template: templateUuid, name: fileName, data: variables, safe: safeUuid };
  const created = await d4Fetch('/documents/create', 'POST', payload);

  if (!created?.uuid && safeUuid) {
    // fallback em ambientes antigos
    const alt = await d4Fetch(`/documents/create/${safeUuid}`, 'POST', { template: templateUuid, name: fileName, data: variables });
    return alt;
  }
  return created;
}

async function d4AddSigners(documentKey, signers) {
  const payload = {
    signers: signers.map(s => ({
      email: s.email,
      act: '1',
      foreign: '0',
      certificadoicpbr: '0',
      name: s.name || s.email,
    })),
  };
  return d4Fetch(`/documents/${documentKey}/signers`, 'POST', payload);
}

async function d4SendToSign(documentKey, message = 'Contrato para assinatura') {
  return d4Fetch(`/documents/${documentKey}/sendto`, 'POST', { message, emails: [], workflow: '0' });
}

// Proxy de download (vendedor nunca vê URL do D4)
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

// ======================
// 8) Fluxo de páginas internas
// ======================

// (A) Automação do Pipefy grava link da Página 1 (Confirmar)
app.post('/novo-pipe/criar-link-confirmacao', async (req, res) => {
  try {
    const { cardId } = req.body;
    if (!cardId) return res.status(400).json({ error: 'cardId é obrigatório' });

    const confirmUrl = makeSignedURL('/novo-pipe/confirmar', { cardId });
    await updateCardField(cardId, FIELD_ID_LINKS_D4, confirmUrl);
    res.json({ ok: true, link: confirmUrl });
  } catch (e) {
    res.status(500).json({ error: String(e.message || e) });
  }
});

// (B) Página 1 — Confirmar dados e acionar Gerar
app.get('/novo-pipe/confirmar', async (req, res) => {
  try {
    if (!validateSignature(req)) return res.status(401).send('assinatura inválida');
    const { cardId } = req.query;
    if (!cardId) return res.status(400).send('Faltou cardId');

    const card = await getCardFields(cardId);

    // validações opcionais
    if (NOVO_PIPE_ID && String(card?.pipe?.id) !== String(NOVO_PIPE_ID)) {
      return res.status(400).send('Card não pertence ao pipe configurado');
    }
    if (FASE_VISITA_ID && String(card?.current_phase?.id) !== String(FASE_VISITA_ID)) {
      return res.status(400).send('Card não está na fase esperada');
    }

    const vars = await buildTemplateVariablesAsync(card);

    const resumo = `
      <h2>Confirmar dados do cliente</h2>
      <p><b>Contratante:</b> ${vars.contratante_1 || ''}</p>
      <p><b>CPF:</b> ${vars.cpf || ''} &nbsp; <b>RG:</b> ${vars.rg || ''}</p>
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

// (C) POST Gerar — cria no D4 (template UUID + cofre do responsável) e redireciona p/ Página 2
app.post('/novo-pipe/gerar', async (req, res) => {
  try {
    const { cardId, sig } = req.body;
    if (!cardId) return res.status(400).send('Faltou cardId');
    if (!sig) return res.status(401).send('assinatura ausente');

    const card = await getCardFields(cardId);
    const vars = await buildTemplateVariablesAsync(card);

    // Template UUID
    const templateUuid = vars.TEMPLATE_UUID_CONTRATO || TEMPLATE_UUID_CONTRATO;
    if (!templateUuid) throw new Error('TEMPLATE_UUID_CONTRATO não encontrado (card ou .env)');

    // Cofre do responsável
    const safeUuid = resolveCofreUuidByCard(card);
    if (!safeUuid) throw new Error('Não foi possível resolver o cofre pelo responsável do card');

    // Cria documento
    const fileName = `Contrato_${card.title || card.id}.docx`;
    const created = await d4CreateFromTemplateUUID({ templateUuid, fileName, variables: vars, safeUuid });
    const documentKey = created.uuid || created.documentKey || created.key;
    if (!documentKey) throw new Error('Não foi possível obter documentKey do D4');

    // Atualiza link no card p/ Página 2
    const nextUrl = makeSignedURL(`/contratos/${documentKey}`, { cardId });
    await updateCardField(cardId, FIELD_ID_LINKS_D4, nextUrl);

    res.redirect(nextUrl);
  } catch (e) {
    res.status(500).send(String(e.message || e));
  }
});

// (D) Página 2 — Baixar e Enviar para assinatura
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

    // lista de signatários: empresa + cliente (se tivermos e-mail)
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

// ======================
// 9) Endpoints para listar/exportar TODOS os contatos
// ======================
app.get('/contatos', async (req, res) => {
  try {
    if (!CONTACTS_TABLE_ID) return res.status(400).json({ error: 'CONTACTS_TABLE_ID não definido' });
    const out = [];
    let after = null;
    for (let i = 0; i < 100; i++) {
      const { records, pageInfo } = await listTableRecords(CONTACTS_TABLE_ID, 200, after);
      for (const r of records) {
        const { email, telefone } = extractEmailPhoneFromRecord(r);
        out.push({ id: r.id, title: r.title, email, telefone });
      }
      if (!pageInfo?.hasNextPage) break;
      after = pageInfo.endCursor || null;
    }
    res.json({ ok: true, count: out.length, contacts: out });
  } catch (e) {
    res.status(500).json({ error: String(e.message || e) });
  }
});

app.get('/contatos/export.csv', async (req, res) => {
  try {
    if (!CONTACTS_TABLE_ID) return res.status(400).send('CONTACTS_TABLE_ID não definido');
    res.setHeader('Content-Type', 'text/csv; charset=utf-8');
    res.setHeader('Content-Disposition', 'attachment; filename="contatos.csv"');
    // header
    res.write('id;title;email;telefone\n');
    let after = null;
    for (let i = 0; i < 100; i++) {
      const { records, pageInfo } = await listTableRecords(CONTACTS_TABLE_ID, 200, after);
      for (const r of records) {
        const { email, telefone } = extractEmailPhoneFromRecord(r);
        const line = `${r.id};"${(r.title || '').replace(/"/g, '""')}";"${(email || '').replace(/"/g, '""')}";"${(telefone || '').replace(/"/g, '""')}"\n`;
        res.write(line);
      }
      if (!pageInfo?.hasNextPage) break;
      after = pageInfo.endCursor || null;
    }
    res.end();
  } catch (e) {
    res.status(500).send(String(e.message || e));
  }
});

// ======================
app.get('/health', (_, res) => res.json({ ok: true }));
// ======================

app.listen(PORT, () => {
  console.log(`Servidor rodando na porta ${PORT}`);
});
