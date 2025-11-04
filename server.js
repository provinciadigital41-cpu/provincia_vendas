// server.js
// Novo endpoint para gerar contrato no novo pipe mantendo na mesma fase

const express = require('express');

const app = express();
app.use(bodyParser.json());

// Variáveis de ambiente necessárias
const {
  D4_CRYPT,
  D4_TOKEN,
  PIpefy_token, // já existente no seu projeto
  EMAIL_ASSINATURA_EMPRESA,
  PIPEFY_FIELD_LINK_CONTRATO,
} = process.env;

// Configurações do novo pipe e fase
// Pipe e fase conforme documento enviado
const NOVO_PIPE_ID = '306505295';       // novo pipe
const FASE_VISITA_ID = '339299691';     // fase "Visita"
// Ambos conforme IDs do documento. :contentReference[oaicite:4]{index=4}

// Template do D4 já mapeado
// id "24" é o template "CONTRATO NOVO MODELO D4.docx" com tokens_gerais
// e a lista de variáveis informadas no documento. :contentReference[oaicite:5]{index=5}
const D4_TEMPLATE_ID = '24';

// ==================================================================
// Helpers de Pipefy
// ==================================================================

const PIPEFY_API = 'https://api.pipefy.com/graphql';

async function pipefyGraphQL(query, variables) {
  const res = await fetch(PIPEFY_API, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${PIpefy_token}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({ query, variables }),
  });
  if (!res.ok) {
    const txt = await res.text();
    throw new Error(`Pipefy error ${res.status}: ${txt}`);
  }
  const json = await res.json();
  if (json.errors) {
    throw new Error(`Pipefy GraphQL: ${JSON.stringify(json.errors)}`);
  }
  return json.data;
}

// Busca todos os campos do card, incluindo start form e fase atual
async function getCardFields(cardId) {
  const query = `
    query($id: ID!) {
      card(id: $id) {
        id
        title
        fields {
          name
          value
          field {
            id
          }
        }
        pipe { id }
        current_phase { id name }
      }
    }
  `;
  const data = await pipefyGraphQL(query, { id: cardId });
  return data.card;
}

// Atualiza um campo de texto no card com o link final
async function setContractLinkOnCard(cardId, link) {
  const mutation = `
    mutation($input: UpdateCardFieldInput!) {
      updateCardField(input: $input) {
        card { id }
      }
    }
  `;
  const variables = {
    input: {
      card_id: Number(cardId),
      field_id: PIPEFY_FIELD_LINK_CONTRATO,
      new_value: link,
    },
  };
  await pipefyGraphQL(mutation, variables);
}

// ==================================================================
// Helpers de D4
// ==================================================================

const D4_BASE = 'https://api.d4sign.com.br/v2';

async function d4Fetch(path, method, payload) {
  const url = `${D4_BASE}${path}?tokenAPI=${D4_TOKEN}&cryptKey=${D4_CRYPT}`;
  const res = await fetch(url, {
    method,
    headers: { 'Content-Type': 'application/json' },
    body: payload ? JSON.stringify(payload) : undefined,
  });
  if (!res.ok) {
    const txt = await res.text();
    throw new Error(`D4 error ${res.status}: ${txt}`);
  }
  const json = await res.json();
  return json;
}

// Cria documento a partir do template e faz merge dos tokens
async function d4CreateDocumentFromTemplate({ templateId, fileName, variables }) {
  // Ajuste o payload para o formato de variáveis do teu template
  // variables deve ser um objeto { token: valor }
  const payload = {
    template: templateId,
    name: fileName,
    data: variables,
  };
  return d4Fetch('/documents/create', 'POST', payload);
}

// Cria lista de signatários padrão com e-mail de assinatura da empresa como remetente
async function d4AddSigners(documentKey, signers) {
  const payload = {
    signers: signers.map(s => ({
      email: s.email,
      act: '1', // assinatura simples
      foreign: '0',
      certificadoicpbr: '0',
      name: s.name || s.email,
    })),
  };
  return d4Fetch(`/documents/${documentKey}/signers`, 'POST', payload);
}

// Gera link público para assinatura e para download
async function d4MakeLinks(documentKey) {
  // url para assinatura pública
  const signLink = await d4Fetch(`/documents/${documentKey}/sendto`, 'POST', {
    message: 'Contrato para assinatura',
    emails: [],
    workflow: '0',
  });
  // link público de download
  const downloadLink = await d4Fetch(`/documents/${documentKey}/download`, 'GET');
  return { signLink, downloadLink };
}

// ==================================================================
// Mapeamento de campos → tokens do template
// ==================================================================
//
// tokens_gerais no template:
// contratante_1, estado_civil, rua, bairro, numero, nome_da_cidade, uf, cep, rg, cpf, telefone, E-mail,
// risco, quantidade_depositos_processos_de_marca, nome_da_marca, classe,
// numero_de_parcelas_da_assessoria, valor_da_parcela_da_assessoria, forma_de_pagamento_da_assessoria, data_de_pagamento_da_assessoria,
// valor_da_pesquisa, forma_de_pagamento_da_pesquisa, data_de_pagamento_da_pesquisa,
// valor_da_taxa, forma_de_pagamento_da_taxa, data_de_pagamento_da_taxa,
// cidade, dia, mes, ano
//
// IDs disponíveis no Start Form e Fase Visita constam no doc. :contentReference[oaicite:6]{index=6}
//
// Abaixo um mapeamento sugerido para este novo pipe.
// Ajuste qualquer item que desejar.

const mapFromStartForm = {
  // Start form
  contratante_1: 'nome_ou_raz_o_social',
  cep: 'cep',
  uf: 'uf',
  cidade: 'cidade',
  bairro: 'bairro',
  rua: 'rua',
  numero: 'n_mero_1',
  // Pagamentos e condições
  valor_da_parcela_da_assessoria: 'valor_da_proposta',
  numero_de_parcelas_da_assessoria: 'quantidade_de_parcelas',
  // Pesquisa
  forma_de_pagamento_da_pesquisa: 'pesquisaa',
  // Taxa
  valor_da_taxa: 'pre_o',
  // Campos complementares se existirem
};

const mapFromPhaseVisita = {
  // Fase Visita
  cpf: 'cpf',
  r_social_ou_n_completo: 'r_social_ou_n_completo',
  rua_av_do_cnpj: 'rua_av_do_cnpj',
  bairro_do_cnpj: 'bairro_do_cnpj',
  cidade_do_cnpj: 'cidade_do_cnpj',
  estado_do_cnpj: 'estado_do_cnpj',
  cep_do_cnpj: 'cep_do_cnpj',
  telefone: 'contato', // se o conector trouxer telefone, adaptar aqui
  tipo_da_empresa: 'tipo_da_empresa',
  tipo_da_empresa_cotitular: 'tipo_da_empresa_cotitular',
  valor_da_assessoria: 'valor_da_assessoria',
  sele_o_de_lista: 'sele_o_de_lista',
  taxa: 'taxa',
  copy_of_tipo_de_pagamento_taxa: 'copy_of_tipo_de_pagamento_taxa',
  tipo_de_pagamento: 'tipo_de_pagamento',
  copy_of_tipo_de_pagamento: 'copy_of_tipo_de_pagamento',
  pesquisa: 'pesquisa',
  tipo_de_pagamento_benef_cio: 'tipo_de_pagamento_benef_cio',
  tem_s_cio: 'tem_s_cio',
  nome_do_s_cio: 'nome_do_s_cio',
  cpf_do_s_cio: 'cpf_do_s_cio',
  coment_rios_para_o_contrato: 'coment_rios_para_o_contrato',
};

// Conversão de valores de campos do Pipefy para o formato de tokens
function normalizeValue(fieldId, value) {
  if (value == null) return '';
  // Exemplos de normalização
  if (Array.isArray(value)) return value.join(', ');
  if (typeof value === 'object') return JSON.stringify(value);
  return String(value);
}

// Constrói objeto { token: valor } combinando Start Form e Fase
function buildTemplateVariables(card) {
  const kv = {};
  const byId = {};
  for (const f of card.fields || []) {
    if (f.field && f.field.id) {
      byId[f.field.id] = f.value;
    }
  }

  // tokens que puxamos do start form
  for (const token in mapFromStartForm) {
    const src = mapFromStartForm[token];
    kv[token] = normalizeValue(src, byId[src]);
  }
  // tokens que puxamos da fase Visita
  for (const token in mapFromPhaseVisita) {
    const src = mapFromPhaseVisita[token];
    kv[token] = normalizeValue(src, byId[src]);
  }

  // Campos que dependem de data atual
  const now = new Date();
  kv.dia = String(now.getDate()).padStart(2, '0');
  kv.mes = String(now.getMonth() + 1).padStart(2, '0');
  kv.ano = String(now.getFullYear());

  // Preenchimentos padrão se faltarem
  kv['E-mail'] = kv['E-mail'] || '';
  kv.uf = kv.uf || kv.estado_do_cnpj || '';
  kv.cidade = kv.cidade || kv.nome_da_cidade || kv.cidade_do_cnpj || '';

  return kv;
}

// ==================================================================
// Endpoint novo
// ==================================================================
//
// POST /novo-pipe/gerar-contrato
// body: { cardId: "..." }
//
// Requisitos:
// 1) O card deve estar no pipe correto e na fase Visita
// 2) Preenche tokens com dados do start form e fase
// 3) Gera documento no D4 a partir do template
// 4) Adiciona signatários
// 5) Gera links de assinatura e download
// 6) Grava o link no campo configurado do card
// 7) Não move de fase

app.post('/novo-pipe/gerar-contrato', async (req, res) => {
  try {
    const { cardId } = req.body;
    if (!cardId) return res.status(400).json({ error: 'cardId é obrigatório' });

    const card = await getCardFields(cardId);

    if (!card || !card.pipe || card.pipe.id !== NOVO_PIPE_ID) {
      return res.status(400).json({ error: 'Card não pertence ao novo pipe configurado' });
    }
    if (!card.current_phase || card.current_phase.id !== FASE_VISITA_ID) {
      return res.status(400).json({ error: 'Card não está na fase Visita do novo pipe' });
    }

    const variables = buildTemplateVariables(card);

    const fileName = `Contrato_${card.title || card.id}.docx`;

    const created = await d4CreateDocumentFromTemplate({
      templateId: D4_TEMPLATE_ID,
      fileName,
      variables,
    });

    const documentKey = created.uuid || created.documentKey || created.uuidDoc || created.key;

    if (!documentKey) {
      throw new Error('Não foi possível obter a chave do documento D4');
    }

    // Ajuste a lista de signatários conforme necessidade do card
    // Nesta versão, só adiciona o e-mail institucional para disparo
    await d4AddSigners(documentKey, [
      { email: EMAIL_ASSINATURA_EMPRESA },
    ]);

    const { signLink, downloadLink } = await d4MakeLinks(documentKey);

    const finalUrl = Array.isArray(signLink?.url) ? signLink.url[0] : (signLink?.url || '');
    const publicDownload = downloadLink?.url || '';

    const composedLink = finalUrl || publicDownload || `https://secure.d4sign.com.br/${documentKey}`;

    await setContractLinkOnCard(cardId, composedLink);

    return res.json({
      ok: true,
      cardId,
      documentKey,
      contract_link: composedLink,
      sign_link_raw: signLink,
      download_link_raw: downloadLink,
    });
  } catch (err) {
    console.error(err);
    return res.status(500).json({ error: String(err.message || err) });
  }
});

app.get('/health', (_, res) => res.json({ ok: true }));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Novo pipe rodando na porta ${PORT}`);
});
