Risco agregado
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
    'N° contrato': String(d.cardId || ''),
    'Nº contrato': String(d.cardId || ''),
    'Numero contrato': String(d.cardId || ''),
    'Número contrato': String(d.cardId || ''),
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
    'N° contrato': String(d.cardId || ''),
    'Nº contrato': String(d.cardId || ''),
    'Numero contrato': String(d.cardId || ''),
    'Número contrato': String(d.cardId || ''),
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
// NOVO — POSTBACK DO D4SIGN (DOCUMENTO FINALIZADO) - ENDPOINT LEGADO
// ===============================
app.post('/d4sign/postback', async (req, res) => {
  try {
    const { uuid, type_post } = req.body || {};

    if (!uuid) {
      console.warn('[POSTBACK D4SIGN] Sem UUID no body');
      return res.status(200).json({ ok: true });
    }

    // type_post = "1" → documento finalizado/assinado
    // type_post = "4" → documento assinado (também deve ser processado)
    const isSigned = String(type_post) === '1' || String(type_post) === '4';
    if (!isSigned) {
      console.log('[POSTBACK D4SIGN] Evento ignorado:', type_post);
      return res.status(200).json({ ok: true });
    }

    console.log('[POSTBACK D4SIGN] Documento finalizado:', uuid);

    const cardId = await findCardIdByD4Uuid(uuid);
    if (!cardId) {
      console.warn('[POSTBACK D4SIGN] Nenhum card encontrado para uuid:', uuid);
      return res.status(200).json({ ok: true });
    }

    // Identificar se é contrato ou procuração
    // Verifica em qual campo o UUID do documento foi encontrado
    const card = await getCard(cardId);
    const byId = toById(card);
    
    // Verifica qual campo contém o UUID do documento
    const uuidContrato = byId[PIPEFY_FIELD_D4_UUID_CONTRATO] || '';
    const uuidProcuracao = byId[PIPEFY_FIELD_D4_UUID_PROCURACAO] || '';
    
    // Verifica se o UUID está no campo de procuração ou contrato
    const isProcuracaoFinal = (String(uuidProcuracao) === uuid || String(uuidProcuracao).includes(uuid));
    
    if (!isProcuracaoFinal && String(uuidContrato) !== uuid && !String(uuidContrato).includes(uuid)) {
      // Se não encontrou em nenhum campo, assume contrato por padrão
      console.log('[POSTBACK D4SIGN] Tipo não identificado claramente, assumindo CONTRATO');
    }

    console.log(`[POSTBACK D4SIGN] Documento identificado como: ${isProcuracaoFinal ? 'PROCURAÇÃO' : 'CONTRATO'}`);

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

    // 3. anexar PDF no campo correto (Contrato Assinado D4 ou Procuração Assinada D4)
    try {
      const fieldId = isProcuracaoFinal 
        ? PIPEFY_FIELD_PROCURACAO_ASSINADA_D4 
        : PIPEFY_FIELD_CONTRATO_ASSINADO_D4;

      console.log(`[POSTBACK D4SIGN] Tentando salvar PDF - isProcuracaoFinal: ${isProcuracaoFinal}, fieldId: ${fieldId}`);
      console.log(`[POSTBACK D4SIGN] Valores dos campos - PIPEFY_FIELD_PROCURACAO_ASSINADA_D4: ${PIPEFY_FIELD_PROCURACAO_ASSINADA_D4}, PIPEFY_FIELD_CONTRATO_ASSINADO_D4: ${PIPEFY_FIELD_CONTRATO_ASSINADO_D4}`);

      if (!fieldId) {
        console.warn(`[POSTBACK D4SIGN] Campo não configurado para ${isProcuracaoFinal ? 'procuração' : 'contrato'}`);
      } else {
        const newValue = [info.url];
        console.log(`[POSTBACK D4SIGN] Salvando PDF no campo ${fieldId} com URL: ${info.url}`);
        await updateCardField(cardId, fieldId, newValue);
        console.log(`[POSTBACK D4SIGN] ✓ PDF anexado com sucesso no campo ${fieldId} (${isProcuracaoFinal ? 'Procuração' : 'Contrato'} Assinado D4)`);
      }
    } catch (e) {
      console.error('[POSTBACK D4SIGN] Erro ao anexar documento:', e.message);
      console.error('[POSTBACK D4SIGN] Stack trace:', e.stack);
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
// UUID será salvo quando o documento for enviado para assinatura
// ===============================

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

    // UUID será salvo quando a procuração for enviada para assinatura

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

    // Signatários serão cadastrados apenas quando o documento for enviado para assinatura
    // Isso evita duplicação de signatários
    console.log('[D4SIGN] Procuração criada. Signatários serão cadastrados quando o documento for enviado para assinatura.');
  } catch (e) {
    console.error('[ERRO] Falha ao gerar procuração:', e.message);
    // Não bloqueia o fluxo se a procuração falhar
  }
}

    await new Promise(r=>setTimeout(r, 3000));
    try { await getDocumentStatus(D4SIGN_TOKEN, D4SIGN_CRYPT_KEY, uuidDoc); } catch {}

    // Signatários serão cadastrados apenas quando o documento for enviado para assinatura
    // Isso evita duplicação de signatários
    console.log('[D4SIGN] Contrato criado. Signatários serão cadastrados quando o documento for enviado para assinatura.');

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
      const urlCofreMsg = data.urlCofre ? '<br><br><div style="margin-top:12px;padding:12px;background:#f5f5f5;border-radius:8px;border-left:4px solid #1976d2;"><strong style="color:#1976d2">Link D4 para adicionar novos signatários ou enviar por whatsapp:</strong><br><a href="' + data.urlCofre + '" target="_blank" style="color:#1976d2;text-decoration:underline;word-break:break-all">' + data.urlCofre + '</a></div>' : '';
      const destinoMsg = data.email ? ' para ' + data.email + ' por email' : ' por email';
      statusDiv.innerHTML = '<span style="color:#28a745;font-weight:600">✓ Status de envio - Contrato: Enviado com sucesso' + destinoMsg + '.' + cofreMsg + '</span>' + urlCofreMsg;
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
    const response = await fetch('/lead/' + encodeURIComponent(token) + '/doc/' + encodeURIComponent(uuidProcuracao) + '/send?canal=email&tipo=procuracao', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' }
    });
    
    const data = await response.json();
    
    if (response.ok && data.success) {
      const cofreMsg = data.cofre ? ' Salvo no cofre: ' + data.cofre : '';
      const urlCofreMsg = data.urlCofre ? '<br><br><div style="margin-top:12px;padding:12px;background:#f5f5f5;border-radius:8px;border-left:4px solid #1976d2;"><strong style="color:#1976d2">Link D4 para adicionar novos signatários ou enviar por whatsapp:</strong><br><a href="' + data.urlCofre + '" target="_blank" style="color:#1976d2;text-decoration:underline;word-break:break-all">' + data.urlCofre + '</a></div>' : '';
      const destinoMsg = data.email ? ' para ' + data.email + ' por email' : ' por email';
      statusDiv.innerHTML = '<span style="color:#28a745;font-weight:600">✓ Status de envio - Procuração: Enviado com sucesso' + destinoMsg + '.' + cofreMsg + '</span>' + urlCofreMsg;
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
  const tipo = req.query.tipo || null; // 'contrato' ou 'procuracao' (opcional)
  
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
      // Verificar se o UUID está nos campos para identificar o tipo
      const uuidProcuracaoCard = by[PIPEFY_FIELD_D4_UUID_PROCURACAO] || null;
      const uuidContratoCard = by[PIPEFY_FIELD_D4_UUID_CONTRATO] || null;
      
      // Determinar se é contrato ou procuração
      // Primeiro verifica se o tipo foi passado como parâmetro
      console.log(`[SEND] Parâmetros recebidos - tipo: ${tipo}, uuidDoc: ${uuidDoc}`);
      console.log(`[SEND] Campos do card - uuidProcuracaoCard: ${uuidProcuracaoCard}, uuidContratoCard: ${uuidContratoCard}`);
      
      if (tipo === 'procuracao') {
        isProcuracao = true;
        console.log('[SEND] Tipo identificado como PROCURAÇÃO pelo parâmetro tipo');
      } else if (tipo === 'contrato') {
        isProcuracao = false;
        console.log('[SEND] Tipo identificado como CONTRATO pelo parâmetro tipo');
      }
      // Se o UUID corresponde ao campo de procuração, é procuração
      else if (uuidProcuracaoCard && (String(uuidProcuracaoCard) === uuidDoc || String(uuidProcuracaoCard).includes(uuidDoc))) {
        isProcuracao = true;
        console.log('[SEND] Tipo identificado como PROCURAÇÃO pelo campo d4_uuid_procuracao');
      } 
      // Se corresponde ao campo de contrato, é contrato
      else if (uuidContratoCard && (String(uuidContratoCard) === uuidDoc || String(uuidContratoCard).includes(uuidDoc))) {
        isProcuracao = false;
        console.log('[SEND] Tipo identificado como CONTRATO pelo campo d4_uuid_contrato');
      } 
      // Se não encontrou em nenhum campo, verifica qual campo está vazio
      else if (!uuidProcuracaoCard && uuidContratoCard) {
        // Procuração vazia mas contrato preenchido, então este deve ser procuração
        isProcuracao = true;
        console.log('[SEND] Tipo identificado como PROCURAÇÃO (procuração vazia mas contrato preenchido)');
      } else if (uuidProcuracaoCard && !uuidContratoCard) {
        // Contrato vazio mas procuração preenchida, então este deve ser contrato
        isProcuracao = false;
        console.log('[SEND] Tipo identificado como CONTRATO (contrato vazio mas procuração preenchida)');
      } else {
        // Ambos vazios - por padrão assume contrato
        isProcuracao = false;
        console.log('[SEND] Ambos campos vazios, assumindo CONTRATO. Se for procuração, passe ?tipo=procuracao na URL.');
      }
      
      console.log(`[SEND] isProcuracao final: ${isProcuracao}`);
      
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
      
      // Salvar UUID do documento no campo correto após enviar para assinatura
      try {
        console.log(`[SEND] Tentando salvar UUID - isProcuracao: ${isProcuracao}, uuidDoc: ${uuidDoc}, cardId: ${cardId}`);
        if (isProcuracao) {
          // Salva UUID da procuração no campo D4 UUID Procuracao
          const fieldId = PIPEFY_FIELD_D4_UUID_PROCURACAO;
          console.log(`[SEND] Salvando UUID da procuração no campo ${fieldId}...`);
          await updateCardField(cardId, fieldId, uuidDoc);
          console.log(`[SEND] ✓ UUID da procuração salvo com sucesso no campo ${fieldId}: ${uuidDoc}`);
        } else {
          // Salva UUID do contrato no campo D4 UUID Contrato
          const fieldId = PIPEFY_FIELD_D4_UUID_CONTRATO;
          console.log(`[SEND] Salvando UUID do contrato no campo ${fieldId}...`);
          await updateCardField(cardId, fieldId, uuidDoc);
          console.log(`[SEND] ✓ UUID do contrato salvo com sucesso no campo ${fieldId}: ${uuidDoc}`);
        }
      } catch (e) {
        console.error(`[ERRO] Falha ao salvar UUID do ${isProcuracao ? 'procuração' : 'contrato'} no card:`, e.message);
        console.error(`[ERRO] Stack trace:`, e.stack);
        // Não bloqueia o fluxo se falhar ao salvar UUID
      }
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
    
    // Buscar UUID do cofre novamente para construir a URL
    const equipeContratoFinal = getEquipeContratoFromCard(card);
    let uuidCofreFinal = null;
    if (equipeContratoFinal && COFRES_UUIDS[equipeContratoFinal]) {
      uuidCofreFinal = COFRES_UUIDS[equipeContratoFinal];
    }
    if (!uuidCofreFinal) {
      uuidCofreFinal = DEFAULT_COFRE_UUID;
    }
    
    // Construir URL do cofre (apenas o link base do D4Sign)
    const urlCofre = 'https://secure.d4sign.com.br/desk';
    
    // Liberar lock após envio bem-sucedido
    releaseLock(lockKey);
    
    return res.status(200).json({
      success: true,
      message: `${isProcuracao ? 'Procuração' : 'Contrato'} enviado com sucesso. Os signatários foram notificados.`,
      tipo: isProcuracao ? 'procuração' : 'contrato',
      cofre: nomeCofre,
      urlCofre: urlCofre,
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
// Busca nos campos d4_uuid_contrato e d4_uuid_procuracao
// ===============================
async function findCardIdByD4Uuid(uuidDocument) {
  // Busca o card pelo UUID do documento nos campos d4_uuid_contrato e d4_uuid_procuracao
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

  // Primeiro tenta buscar no campo D4 UUID Contrato
  try {
    let data = await gql(query, {
      pipeId: NOVO_PIPE_ID,
      fieldId: PIPEFY_FIELD_D4_UUID_CONTRATO,
      fieldValue: uuidDocument
    });

    let edges = data?.findCards?.edges || [];
    if (edges.length) {
      console.log(`[findCardIdByD4Uuid] Card encontrado pelo campo ${PIPEFY_FIELD_D4_UUID_CONTRATO}: ${edges[0].node.id}`);
      return edges[0].node.id;
    }
  } catch (e) {
    console.warn(`[findCardIdByD4Uuid] Erro ao buscar pelo campo ${PIPEFY_FIELD_D4_UUID_CONTRATO}:`, e.message);
  }

  // Se não encontrou, tenta buscar no campo D4 UUID Procuracao
  try {
    let data = await gql(query, {
      pipeId: NOVO_PIPE_ID,
      fieldId: PIPEFY_FIELD_D4_UUID_PROCURACAO,
      fieldValue: uuidDocument
    });

    let edges = data?.findCards?.edges || [];
    if (edges.length) {
      console.log(`[findCardIdByD4Uuid] Card encontrado pelo campo ${PIPEFY_FIELD_D4_UUID_PROCURACAO}: ${edges[0].node.id}`);
      return edges[0].node.id;
    }
  } catch (e) {
    console.warn(`[findCardIdByD4Uuid] Erro ao buscar pelo campo ${PIPEFY_FIELD_D4_UUID_PROCURACAO}:`, e.message);
  }

  // Busca alternativa: buscar cards recentes e verificar manualmente
  try {

    // Busca alternativa: buscar cards recentes e verificar manualmente
    // Isso é necessário porque o campo d4_uuid_contrato agora contém URL do cofre, não UUID do documento
    console.log(`[findCardIdByD4Uuid] Busca direta não encontrou. Tentando busca alternativa em cards recentes...`);
    
    try {
      // Busca os cards mais recentes do pipe para verificar manualmente
      const searchQuery = `
        query($pipeId: ID!, $first: Int!) {
          pipe(id: $pipeId) {
            cards(first: $first, orderBy: { field: CREATED_AT, direction: DESC }) {
          edges {
            node {
              id
                  fields {
                    id
                    value
                  }
                }
            }
          }
        }
      }
    `;

      const searchData = await gql(searchQuery, {
      pipeId: NOVO_PIPE_ID,
        first: 100  // Busca nos 100 cards mais recentes
      });

      const recentCards = searchData?.pipe?.cards?.edges || [];
      for (const edge of recentCards) {
        const card = edge.node;
        const fields = card.fields || [];
        
        // Verifica se algum campo contém o UUID do documento
        for (const field of fields) {
          const fieldValue = String(field.value || '');
          // Verifica se o UUID está nos campos D4 UUID Contrato ou D4 UUID Procuracao
          if ((field.id === PIPEFY_FIELD_D4_UUID_CONTRATO || field.id === PIPEFY_FIELD_D4_UUID_PROCURACAO) && 
              (fieldValue === uuidDocument || fieldValue.includes(uuidDocument))) {
            console.log(`[findCardIdByD4Uuid] Card encontrado através de busca alternativa no campo ${field.id}: ${card.id}`);
            return card.id;
          }
        }
      }
      
      console.log(`[findCardIdByD4Uuid] UUID ${uuidDocument} não encontrado em cards recentes.`);
    } catch (searchError) {
      console.warn('[findCardIdByD4Uuid] Erro na busca alternativa:', searchError.message);
    }
  } catch (e) {
    console.warn('[findCardIdByD4Uuid] Erro ao buscar pelo campo de contrato:', e.message);
  }

  return null;
};

// ===============================
// NOVO — ANEXA CONTRATO ASSINADO NO CAMPO DE ANEXO
// ===============================
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
