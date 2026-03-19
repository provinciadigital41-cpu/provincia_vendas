'use strict';

/**
 * server.js — Provincia Vendas (Pipefy + D4Sign via secure.d4sign.com.br)
 * Node 18+ (fetch global)
 */

const express = require('express');
const crypto = require('crypto');
const fs = require('fs-extra');
const path = require('path');

const app = express();
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.set('trust proxy', true);

// Log básico
app.use((req, _res, next) => {
  console.log(`[${new Date().toISOString()}] ${req.method} ${req.originalUrl} ua="${req.get('user-agent')}" ip=${req.ip}`);
  next();
});

// [NOVO] Dashboard de Visualização
app.get('/', (req, res) => {
  res.send(`
    <!DOCTYPE html>
    <html lang="pt-BR">
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>Província Vendas - Debugger</title>
      <style>
        body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; background: #f0f2f5; margin: 0; padding: 20px; color: #333; }
        .container { max-width: 1000px; margin: 0 auto; background: white; padding: 30px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        h1 { margin-top: 0; color: #1a73e8; }
        .input-group { display: flex; gap: 10px; margin-bottom: 20px; }
        input { flex: 1; padding: 10px; border: 1px solid #ddd; border-radius: 4px; font-size: 16px; }
        button { padding: 10px 20px; background: #1a73e8; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 16px; }
        button:hover { background: #1557b0; }
        #result { margin-top: 20px; }
        .section { margin-bottom: 30px; border: 1px solid #eee; border-radius: 4px; overflow: hidden; }
        .section-header { background: #f8f9fa; padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eee; display: flex; justify-content: space-between; align-items: center; cursor: pointer; }
        .section-content { padding: 15px; display: none; }
        .section.open .section-content { display: block; }
        pre { background: #2d2d2d; color: #f8f8f2; padding: 15px; border-radius: 4px; overflow-x: auto; margin: 0; }
        .table-view { width: 100%; border-collapse: collapse; }
        .table-view th, .table-view td { text-align: left; padding: 8px; border-bottom: 1px solid #eee; }
        .table-view th { color: #666; font-size: 0.9em; width: 30%; }
        .loading { text-align: center; padding: 20px; color: #666; }
        .error { background: #fee; color: #c00; padding: 15px; border-radius: 4px; }
        .badge { display: inline-block; padding: 2px 6px; border-radius: 4px; font-size: 0.8em; font-weight: bold; }
        .badge-cpf { background: #e6f4ea; color: #137333; }
        .badge-cnpj { background: #e8f0fe; color: #1967d2; }
      </style>
    </head>
    <body>
      <div class="container">
        <h1>🔍 Província Vendas - Debugger</h1>
        <p>Insira o ID do Card do Pipefy para visualizar os dados extraídos e as variáveis do contrato.</p>
        
        <div class="input-group">
          <input type="text" id="cardId" placeholder="Ex: 123456789" />
          <button onclick="fetchData()">Visualizar Dados</button>
          <button id="btnGerar" onclick="gerarContrato()" style="background-color: #28a745; display: none;">Gerar Contrato Agora</button>
          <button id="btnBaixar" onclick="baixarDocumentos()" style="background-color: #17a2b8; display: none;">📥 Baixar e Anexar Documentos</button>
        </div>

        <div id="result"></div>
      </div>

      <script>
        async function fetchData() {
          const id = document.getElementById('cardId').value.trim();
          if (!id) return alert('Digite um ID');
          
          const resDiv = document.getElementById('result');
          resDiv.innerHTML = '<div class="loading">Carregando dados do Pipefy...</div>';
          
          try {
            const res = await fetch('/debug-card/' + id);
            const data = await res.json();
            
            if (data.error) {
              resDiv.innerHTML = '<div class="error">Erro: ' + data.error + '</div>';
              return;
            }

            renderResult(data);
            document.getElementById('btnGerar').style.display = 'block';
            document.getElementById('btnBaixar').style.display = 'block';
          } catch (e) {
            resDiv.innerHTML = '<div class="error">Erro de conexão: ' + e.message + '</div>';
          }
        }

        function renderResult(data) {
          const { raw, extracted, varsMarca, varsOutros, templateInfo } = data;
          
          let html = '';

          // Info Básica
          html += '<div class="section open"><div class="section-header" onclick="toggle(this)">📋 Resumo e Decisão</div><div class="section-content">';
          html += '<table class="table-view">';
          html += '<tr><th>Template Escolhido</th><td>' + (templateInfo.uuid || 'N/A') + ' (' + templateInfo.type + ')</td></tr>';
          html += '<tr><th>Tipo de Pessoa</th><td><span class="badge ' + (extracted.selecao_cnpj_ou_cpf === 'CPF' ? 'badge-cpf' : 'badge-cnpj') + '">' + (extracted.selecao_cnpj_ou_cpf || 'Indefinido') + '</span></td></tr>';
          html += '<tr><th>Contratante 1</th><td>' + (extracted.nome || '') + '</td></tr>';
          html += '<tr><th>Documento</th><td>' + (extracted.cpf || extracted.cnpj || 'N/A') + '</td></tr>';
          html += '</table></div></div>';

          // Variáveis Marca
          html += '<div class="section"><div class="section-header" onclick="toggle(this)">🏷️ Variáveis para Template MARCA</div><div class="section-content">';
          html += renderTable(varsMarca);
          html += '</div></div>';

          // [NOVO] Prévia dos Textos do Contrato
          html += '<div class="section"><div class="section-header" onclick="toggle(this)">📄 Prévia dos Textos do Contrato</div><div class="section-content">';
          html += '<table class="table-view">';
          html += '<tr><th>Texto Contratante 1</th><td>' + (extracted.contratante_1_texto || '') + '</td></tr>';
          html += '<tr><th>Texto Contratante 2</th><td>' + (extracted.contratante_2_texto || '') + '</td></tr>';
          html += '<tr><th>Cláusula Adicional</th><td>' + (extracted.clausula_adicional || '') + '</td></tr>';
          html += '<tr><th>Valor Total</th><td>' + (extracted.valor_total || '') + '</td></tr>';
          html += '<tr><th>Parcelas</th><td>' + (extracted.parcelas || '') + '</td></tr>';
          html += '</table></div></div>';

          // Variáveis Outros
          html += '<div class="section"><div class="section-header" onclick="toggle(this)">📑 Variáveis para Template OUTROS</div><div class="section-content">';
          html += renderTable(varsOutros);
          html += '</div></div>';

          // Dados Extraídos (Interno)
          html += '<div class="section"><div class="section-header" onclick="toggle(this)">⚙️ Dados Internos (Extraídos)</div><div class="section-content"><pre>' + JSON.stringify(extracted, null, 2) + '</pre></div></div>';

          // Raw Card
          html += '<div class="section"><div class="section-header" onclick="toggle(this)">📦 Card Bruto (Pipefy)</div><div class="section-content"><pre>' + JSON.stringify(raw, null, 2) + '</pre></div></div>';

          document.getElementById('result').innerHTML = html;
        }

        function renderTable(obj) {
          if (!obj) return 'Sem dados';
          let h = '<table class="table-view">';
          for (let k in obj) {
            let val = obj[k];
            if (typeof val === 'object') val = JSON.stringify(val);
            h += '<tr><th>' + k + '</th><td>' + (val || '') + '</td></tr>';
          }
          h += '</table>';
          return h;
        }

        function toggle(header) {
          header.parentElement.classList.toggle('open');
        }

        async function gerarContrato() {
          const id = document.getElementById('cardId').value.trim();
          if (!id) return alert('Digite um ID primeiro');
          
          if (!confirm('Tem certeza que deseja gerar o contrato para o card ' + id + '? Isso criará um documento no D4Sign.')) return;

          const btn = document.getElementById('btnGerar');
          btn.disabled = true;
          btn.innerText = 'Gerando...';

          try {
            const res = await fetch('/manual-trigger/' + id, { method: 'POST' });
            const data = await res.json();
            
            if (data.success) {
              alert('Contrato gerado com sucesso! UUID: ' + data.result.uuidDoc);
            } else {
              alert('Erro ao gerar: ' + data.error);
            }
          } catch (e) {
            alert('Erro de conexão: ' + e.message);
          } finally {
            btn.disabled = false;
            btn.innerText = 'Gerar Contrato Agora';
          }
        }

        async function baixarDocumentos() {
          const id = document.getElementById('cardId').value.trim();
          if (!id) return alert('Digite um ID primeiro');
          
          if (!confirm('Deseja baixar e anexar os documentos assinados (Contrato/Procuração) do D4Sign para o Pipefy?')) return;

          const btn = document.getElementById('btnBaixar');
          btn.disabled = true;
          btn.innerText = 'Processando...';

          try {
            const res = await fetch('/manual-attach/' + id, { method: 'POST' });
            const data = await res.json();
            
            if (data.success) {
              let msg = 'Processo concluído!\n';
              data.results.forEach(r => {
                msg += '- ' + r.type + ': ' + r.status + ' (' + r.details + ') \\n';
              });
              alert(msg);
            } else {
              alert('Erro: ' + data.error);
            }
          } catch (e) {
            alert('Erro de conexão: ' + e.message);
          } finally {
            btn.disabled = false;
            btn.innerText = '📥 Baixar e Anexar Documentos';
          }
        }
        }
}

async function reenviarContrato(token, uuidDoc) {
  const btn = document.getElementById('btn-reenviar-contrato');
  const statusDiv = document.getElementById('status-contrato');
  
  btn.disabled = true;
  btn.textContent = 'Reenviando...';
  
  try {
    const response = await fetch('/lead/' + encodeURIComponent(token) + '/doc/' + encodeURIComponent(uuidDoc) + '/resend', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' }
    });
    
    const data = await response.json();
    
    if (response.ok && data.success) {
      statusDiv.innerHTML += '<br><span style="color:#28a745;font-weight:600">✓ Reenvio solicitado com sucesso.</span>';
      
      // Reiniciar timer de 60s
      let timeLeft = 60;
      btn.textContent = 'Reenviar Link (' + timeLeft + 's)';
      btn.style.background = '#6c757d';
      
      const timerId = setInterval(() => {
        timeLeft--;
        if (timeLeft <= 0) {
          clearInterval(timerId);
          btn.textContent = 'Reenviar Link';
          btn.disabled = false;
          btn.style.background = '#111';
        } else {
          btn.textContent = 'Reenviar Link (' + timeLeft + 's)';
        }
      }, 1000);
      
    } else {
      statusDiv.innerHTML += '<br><span style="color:#d32f2f">✗ Erro ao reenviar: ' + (data.message || 'Erro desconhecido') + '</span>';
      btn.textContent = 'Reenviar Link';
      btn.disabled = false;
    }
  } catch (error) {
    statusDiv.innerHTML += '<br><span style="color:#d32f2f">✗ Erro ao reenviar: ' + error.message + '</span>';
    btn.textContent = 'Reenviar Link';
    btn.disabled = false;
  }
}

async function reenviarProcuracao(token, uuidDoc) {
  const btn = document.getElementById('btn-reenviar-procuracao');
  const statusDiv = document.getElementById('status-procuracao');
  
  btn.disabled = true;
  btn.textContent = 'Reenviando...';
  
  try {
    const response = await fetch('/lead/' + encodeURIComponent(token) + '/doc/' + encodeURIComponent(uuidDoc) + '/resend', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' }
    });
    
    const data = await response.json();
    
    if (response.ok && data.success) {
      statusDiv.innerHTML += '<br><span style="color:#28a745;font-weight:600">✓ Reenvio solicitado com sucesso.</span>';
      
      // Reiniciar timer de 60s
      let timeLeft = 60;
      btn.textContent = 'Reenviar Link (' + timeLeft + 's)';
      btn.style.background = '#6c757d';
      
      const timerId = setInterval(() => {
        timeLeft--;
        if (timeLeft <= 0) {
          clearInterval(timerId);
          btn.textContent = 'Reenviar Link';
          btn.disabled = false;
          btn.style.background = '#111';
        } else {
          btn.textContent = 'Reenviar Link (' + timeLeft + 's)';
        }
      }, 1000);
      
    } else {
      statusDiv.innerHTML += '<br><span style="color:#d32f2f">✗ Erro ao reenviar: ' + (data.message || 'Erro desconhecido') + '</span>';
      btn.textContent = 'Reenviar Link';
      btn.disabled = false;
    }
  } catch (error) {
    statusDiv.innerHTML += '<br><span style="color:#d32f2f">✗ Erro ao reenviar: ' + error.message + '</span>';
    btn.textContent = 'Reenviar Link';
    btn.disabled = false;
  }
}
  }
}

async function reenviarContrato(token, uuidDoc) {
  const btn = document.getElementById('btn-reenviar-contrato');
  const statusDiv = document.getElementById('status-contrato');
  
  btn.disabled = true;
  btn.textContent = 'Reenviando...';
  
  try {
    const response = await fetch('/lead/' + encodeURIComponent(token) + '/doc/' + encodeURIComponent(uuidDoc) + '/resend', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' }
    });
    
    const data = await response.json();
    
    if (response.ok && data.success) {
      statusDiv.innerHTML += '<br><span style="color:#28a745;font-weight:600">✓ Reenvio solicitado com sucesso.</span>';
      
      // Reiniciar timer de 60s
      let timeLeft = 60;
      btn.textContent = 'Reenviar Link (' + timeLeft + 's)';
      btn.style.background = '#6c757d';
      
      const timerId = setInterval(() => {
        timeLeft--;
        if (timeLeft <= 0) {
          clearInterval(timerId);
          btn.textContent = 'Reenviar Link';
          btn.disabled = false;
          btn.style.background = '#111';
        } else {
          btn.textContent = 'Reenviar Link (' + timeLeft + 's)';
        }
      }, 1000);
      
    } else {
      statusDiv.innerHTML += '<br><span style="color:#d32f2f">✗ Erro ao reenviar: ' + (data.message || 'Erro desconhecido') + '</span>';
      btn.textContent = 'Reenviar Link';
      btn.disabled = false;
    }
  } catch (error) {
    statusDiv.innerHTML += '<br><span style="color:#d32f2f">✗ Erro ao reenviar: ' + error.message + '</span>';
    btn.textContent = 'Reenviar Link';
    btn.disabled = false;
  }
}

async function reenviarProcuracao(token, uuidDoc) {
  const btn = document.getElementById('btn-reenviar-procuracao');
  const statusDiv = document.getElementById('status-procuracao');
  
  btn.disabled = true;
  btn.textContent = 'Reenviando...';
  
  try {
    const response = await fetch('/lead/' + encodeURIComponent(token) + '/doc/' + encodeURIComponent(uuidDoc) + '/resend', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' }
    });
    
    const data = await response.json();
    
    if (response.ok && data.success) {
      statusDiv.innerHTML += '<br><span style="color:#28a745;font-weight:600">✓ Reenvio solicitado com sucesso.</span>';
      
      // Reiniciar timer de 60s
      let timeLeft = 60;
      btn.textContent = 'Reenviar Link (' + timeLeft + 's)';
      btn.style.background = '#6c757d';
      
      const timerId = setInterval(() => {
        timeLeft--;
        if (timeLeft <= 0) {
          clearInterval(timerId);
          btn.textContent = 'Reenviar Link';
          btn.disabled = false;
          btn.style.background = '#111';
        } else {
          btn.textContent = 'Reenviar Link (' + timeLeft + 's)';
        }
      }, 1000);
      
    } else {
      statusDiv.innerHTML += '<br><span style="color:#d32f2f">✗ Erro ao reenviar: ' + (data.message || 'Erro desconhecido') + '</span>';
      btn.textContent = 'Reenviar Link';
      btn.disabled = false;
    }
  } catch (error) {
    statusDiv.innerHTML += '<br><span style="color:#d32f2f">✗ Erro ao reenviar: ' + error.message + '</span>';
    btn.textContent = 'Reenviar Link';
    btn.disabled = false;
  }
}
async function reenviarContrato(token, uuidDoc) {
  const btn = document.getElementById('btn-reenviar-contrato');
  const statusDiv = document.getElementById('status-contrato');
  
  btn.disabled = true;
  btn.textContent = 'Reenviando...';
  
  try {
    const response = await fetch('/lead/' + encodeURIComponent(token) + '/doc/' + encodeURIComponent(uuidDoc) + '/resend', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' }
    });
    
    const data = await response.json();
    
    if (response.ok && data.success) {
      statusDiv.innerHTML += '<br><span style="color:#28a745;font-weight:600">✓ Reenvio solicitado com sucesso.</span>';
      
      // Reiniciar timer de 60s
      let timeLeft = 60;
      btn.textContent = 'Reenviar Link (' + timeLeft + 's)';
      btn.style.background = '#6c757d';
      
      const timerId = setInterval(() => {
        timeLeft--;
        if (timeLeft <= 0) {
          clearInterval(timerId);
          btn.textContent = 'Reenviar Link';
          btn.disabled = false;
          btn.style.background = '#111';
        } else {
          btn.textContent = 'Reenviar Link (' + timeLeft + 's)';
        }
      }, 1000);
      
    } else {
      statusDiv.innerHTML += '<br><span style="color:#d32f2f">✗ Erro ao reenviar: ' + (data.message || 'Erro desconhecido') + '</span>';
      btn.textContent = 'Reenviar Link';
      btn.disabled = false;
    }
  } catch (error) {
    statusDiv.innerHTML += '<br><span style="color:#d32f2f">✗ Erro ao reenviar: ' + error.message + '</span>';
    btn.textContent = 'Reenviar Link';
    btn.disabled = false;
  }
}

async function reenviarProcuracao(token, uuidDoc) {
  const btn = document.getElementById('btn-reenviar-procuracao');
  const statusDiv = document.getElementById('status-procuracao');
  
  btn.disabled = true;
  btn.textContent = 'Reenviando...';
  
  try {
    const response = await fetch('/lead/' + encodeURIComponent(token) + '/doc/' + encodeURIComponent(uuidDoc) + '/resend', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' }
    });
    
    const data = await response.json();
    
    if (response.ok && data.success) {
      statusDiv.innerHTML += '<br><span style="color:#28a745;font-weight:600">✓ Reenvio solicitado com sucesso.</span>';
      
      // Reiniciar timer de 60s
      let timeLeft = 60;
      btn.textContent = 'Reenviar Link (' + timeLeft + 's)';
      btn.style.background = '#6c757d';
      
      const timerId = setInterval(() => {
        timeLeft--;
        if (timeLeft <= 0) {
          clearInterval(timerId);
          btn.textContent = 'Reenviar Link';
          btn.disabled = false;
          btn.style.background = '#111';
        } else {
          btn.textContent = 'Reenviar Link (' + timeLeft + 's)';
        }
      }, 1000);
      
    } else {
      statusDiv.innerHTML += '<br><span style="color:#d32f2f">✗ Erro ao reenviar: ' + (data.message || 'Erro desconhecido') + '</span>';
      btn.textContent = 'Reenviar Link';
      btn.disabled = false;
    }
  } catch (error) {
    statusDiv.innerHTML += '<br><span style="color:#d32f2f">✗ Erro ao reenviar: ' + error.message + '</span>';
    btn.textContent = 'Reenviar Link';
    btn.disabled = false;
  }
}

async function reenviarContrato(token, uuidDoc) {
  const btn = document.getElementById('btn-reenviar-contrato');
  const originalText = btn.textContent;
  btn.disabled = true;
  btn.textContent = 'Reenviando...';
  
  try {
    const response = await fetch('/lead/' + encodeURIComponent(token) + '/doc/' + encodeURIComponent(uuidDoc) + '/resend', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' }
    });
    const data = await response.json();
    
    if (response.ok && data.success) {
      alert('Sucesso: ' + data.message);
    } else {
      alert('Erro: ' + (data.message || 'Falha ao reenviar'));
    }
  } catch (e) {
    alert('Erro ao reenviar: ' + e.message);
  } finally {
    btn.textContent = originalText;
    if (!originalText.includes('s)')) {
        btn.disabled = false;
    }
  }
}

async function reenviarProcuracao(token, uuidDoc) {
  const btn = document.getElementById('btn-reenviar-procuracao');
  const originalText = btn.textContent;
  btn.disabled = true;
  btn.textContent = 'Reenviando...';
  
  try {
    const response = await fetch('/lead/' + encodeURIComponent(token) + '/doc/' + encodeURIComponent(uuidDoc) + '/resend', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' }
    });
    const data = await response.json();
    
    if (response.ok && data.success) {
      alert('Sucesso: ' + data.message);
    } else {
      alert('Erro: ' + (data.message || 'Falha ao reenviar'));
    }
  } catch (e) {
    alert('Erro ao reenviar: ' + e.message);
  } finally {
    btn.textContent = originalText;
    if (!originalText.includes('s)')) {
        btn.disabled = false;
    }
  }
}
</script>
    </body>
    </html>
  `);
});

// [NOVO] API de Debug
app.get('/debug-card/:id', async (req, res) => {
  try {
    const cardId = req.params.id;
    const card = await getCard(cardId);
    const dados = await montarDados(card);

    // Simula data atual para variáveis de tempo
    const now = new Date();
    const nowInfo = { dia: now.getDate(), mes: now.getMonth() + 1, ano: now.getFullYear() };

    const varsMarca = montarVarsParaTemplateMarca(dados, nowInfo);
    const varsOutros = montarVarsParaTemplateOutros(dados, nowInfo);

    // Recalcula lógica de template para exibir
    const k1 = serviceKindFromText(dados.stmt1); // (Nota: montarDados não retorna stmt1 direto na raiz, mas está em 'entries'. Simplificando aqui recuperando do objeto dados se possível ou re-executando lógica leve)
    // Para simplificar, usamos a variável templateToUse que foi calculada dentro de montarDados? 
    // montarDados retorna 'templateToUse' no objeto!

    res.json({
      raw: card,
      extracted: dados,
      varsMarca,
      varsOutros,
      templateInfo: {
        uuid: dados.templateToUse,
        type: dados.templateToUse === TEMPLATE_UUID_CONTRATO ? 'MARCA' : 'OUTROS'
      }
    });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
});

// [NOVO] Rota para disparo manual
app.post('/manual-trigger/:id', async (req, res) => {
  try {
    const cardId = req.params.id;
    console.log(`[MANUAL TRIGGER] Iniciando geração para card ${cardId}`);

    const result = await processarContrato(cardId);

    res.json({ success: true, result });
  } catch (e) {
    console.error('[MANUAL TRIGGER ERROR]', e);
    res.status(500).json({ success: false, error: e.message });
  }
});

// [NOVO] Rota para baixar e anexar documentos manualmente
app.post('/manual-attach/:id', async (req, res) => {
  try {
    const cardId = req.params.id;
    console.log(`[MANUAL ATTACH] Iniciando processo para card ${cardId}`);

    const card = await getCard(cardId);
    const byId = toById(card);
    const orgId = card.pipe?.organization?.id;

    if (!orgId) {
      console.error(`[MANUAL ATTACH] Organization ID não encontrado`);
      return res.status(400).json({ success: false, error: 'Organization ID não encontrado no card' });
    }

    const uuidContrato = byId[PIPEFY_FIELD_D4_UUID_CONTRATO];
    const uuidProcuracao = byId[PIPEFY_FIELD_D4_UUID_PROCURACAO];

    const results = [];
    const nomeMarcaRaw = byId['marca'] || card.title || 'Documento';
    const nomeMarca = String(nomeMarcaRaw).replace(/[<>:"/\\|?*]/g, '_').trim();

    // Processar Contrato
    if (uuidContrato) {
      try {
        console.log(`[MANUAL ATTACH] Baixando Contrato ${uuidContrato}...`);
        const info = await getDownloadUrl(D4SIGN_TOKEN, D4SIGN_CRYPT_KEY, uuidContrato, { type: 'PDF', language: 'pt' });

        // Upload para Pipefy (anexo real, não link que expira)
        const fileName = `${nomeMarca} - Contrato Assinado.pdf`;
        console.log(`[MANUAL ATTACH] Fazendo upload do Contrato: ${fileName}...`);
        const pipefyUrl = await uploadFileToPipefy(info.url, fileName, orgId);

        // Anexar no campo de anexo - usar array com caminho relativo
        console.log(`[MANUAL ATTACH] Caminho do anexo: ${pipefyUrl}`);
        await updateCardField(cardId, PIPEFY_FIELD_EXTRA_CONTRATO, [pipefyUrl]);
        console.log(`[MANUAL ATTACH] ✓ Contrato anexado no campo ${PIPEFY_FIELD_EXTRA_CONTRATO}`);

        const equipeContrato = getEquipeContratoFromCard(card);
        await saveFileLocally(info.url, fileName, equipeContrato || 'Sem_Equipe');

        results.push({ type: 'Contrato', status: 'Sucesso', details: 'Anexado como arquivo permanente' });
      } catch (e) {
        console.error(`[MANUAL ATTACH] Erro Contrato: ${e.message}`);
        results.push({ type: 'Contrato', status: 'Erro', details: e.message });
      }
    } else {
      results.push({ type: 'Contrato', status: 'Ignorado', details: 'UUID não encontrado' });
    }

    // Processar Procuração
    if (uuidProcuracao) {
      try {
        console.log(`[MANUAL ATTACH] Baixando Procuração ${uuidProcuracao}...`);
        const info = await getDownloadUrl(D4SIGN_TOKEN, D4SIGN_CRYPT_KEY, uuidProcuracao, { type: 'PDF', language: 'pt' });

        // Upload para Pipefy (anexo real, não link que expira)
        const fileName = `${nomeMarca} - Procuração.pdf`;
        console.log(`[MANUAL ATTACH] Fazendo upload da Procuração: ${fileName}...`);
        const pipefyUrl = await uploadFileToPipefy(info.url, fileName, orgId);

        // Anexar no campo de anexo - usar array com caminho relativo
        console.log(`[MANUAL ATTACH] Caminho do anexo: ${pipefyUrl}`);
        await updateCardField(cardId, PIPEFY_FIELD_EXTRA_PROCURACAO, [pipefyUrl]);
        console.log(`[MANUAL ATTACH] ✓ Procuração anexada no campo ${PIPEFY_FIELD_EXTRA_PROCURACAO}`);

        const equipeProcuracao = getEquipeContratoFromCard(card);
        await saveFileLocally(info.url, fileName, equipeProcuracao || 'Sem_Equipe');

        results.push({ type: 'Procuração', status: 'Sucesso', details: 'Anexado como arquivo permanente' });
      } catch (e) {
        console.error(`[MANUAL ATTACH] Erro Procuração: ${e.message}`);
        results.push({ type: 'Procuração', status: 'Erro', details: e.message });
      }
    } else {
      results.push({ type: 'Procuração', status: 'Ignorado', details: 'UUID não encontrado' });
    }

    res.json({ success: true, results });

  } catch (e) {
    console.error('[MANUAL ATTACH ERROR]', e);
    res.status(500).json({ success: false, error: e.message });
  }
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
  PIPEFY_FIELD_CONTRATO_ASSINADO_D4,   // Campo ID para "Contrato Assinado D4"
  PIPEFY_FIELD_PROCURACAO_ASSINADA_D4, // Campo ID para "Procuração Assinada D4"
  PIPEFY_FIELD_D4_UUID_CONTRATO,       // Campo ID para "D4 UUID Contrato"
  PIPEFY_FIELD_D4_UUID_PROCURACAO,      // Campo ID para "D4 UUID Procuracao"
  PIPEFY_FIELD_D4_UUID_TERMO_DE_RISCO, // Campo ID para "D4 UUID Termo de Risco"
  NOVO_PIPE_ID,
  FASE_VISITA_ID,
  PHASE_ID_CONTRATO_ENVIADO,
  PIPEFY_FIELD_EXTRA_CONTRATO = 'contrato',
  PIPEFY_FIELD_EXTRA_PROCURACAO = 'procura_o',
  PIPEFY_FIELD_EXTRA_TERMO_DE_RISCO = 'contrato', // Anexo do termo assinado (mesmo campo do contrato conforme solicitado)

  // D4Sign
  D4SIGN_TOKEN,
  D4SIGN_CRYPT_KEY,
  D4SIGN_BASE_URL,                  // Base URL da API D4Sign
  TEMPLATE_UUID_CONTRATO,           // Modelo de Marca
  TEMPLATE_UUID_CONTRATO_MARCA_RISCO, // Modelo de Marca com Risco (Médio/Alto) - Inclui Termo de Risco integrado
  TEMPLATE_UUID_CONTRATO_OUTROS,    // Modelo de Outros Serviços
  TEMPLATE_UUID_PROCURACAO,         // Modelo de Procuração

  // Assinatura interna
  EMAIL_ASSINATURA_EMPRESA,

  // Cofres
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
  COFRE_UUID_CURITIBA_MATRIZ,

  DEFAULT_COFRE_UUID
} = process.env;
NOVO_PIPE_ID = NOVO_PIPE_ID || '306505295';

PORT = PORT || 3000;
PIPE_GRAPHQL_ENDPOINT = PIPE_GRAPHQL_ENDPOINT || 'https://api.pipefy.com/graphql';
PIPEFY_FIELD_LINK_CONTRATO = PIPEFY_FIELD_LINK_CONTRATO || 'd4_contrato';
PIPEFY_FIELD_D4_UUID_CONTRATO = PIPEFY_FIELD_D4_UUID_CONTRATO || 'd4_uuid_contrato';
PIPEFY_FIELD_D4_UUID_PROCURACAO = PIPEFY_FIELD_D4_UUID_PROCURACAO || 'copy_of_d4_uuid_contrato';
PIPEFY_FIELD_D4_UUID_TERMO_DE_RISCO = PIPEFY_FIELD_D4_UUID_TERMO_DE_RISCO || '';
PIPEFY_FIELD_CONTRATO_ASSINADO_D4 = PIPEFY_FIELD_CONTRATO_ASSINADO_D4 || 'contrato_assinado_d4';
PIPEFY_FIELD_PROCURACAO_ASSINADA_D4 = PIPEFY_FIELD_PROCURACAO_ASSINADA_D4 || 'procura_o_assinada_d4';
D4SIGN_BASE_URL = D4SIGN_BASE_URL || 'https://secure.d4sign.com.br/api/v1';

if (!PUBLIC_BASE_URL || !PUBLIC_LINK_SECRET) console.warn('[AVISO] Configure PUBLIC_BASE_URL e PUBLIC_LINK_SECRET');
if (!PIPE_API_KEY) console.warn('[AVISO] PIPE_API_KEY ausente');
if (!D4SIGN_TOKEN || !D4SIGN_CRYPT_KEY) console.warn('[AVISO] D4SIGN_TOKEN / D4SIGN_CRYPT_KEY ausentes');
if (!PIPEFY_FIELD_CONTRATO_ASSINADO_D4) console.warn('[AVISO] PIPEFY_FIELD_CONTRATO_ASSINADO_D4 ausente');
if (!PIPEFY_FIELD_PROCURACAO_ASSINADA_D4) console.warn('[AVISO] PIPEFY_FIELD_PROCURACAO_ASSINADA_D4 ausente');
if (!PIPEFY_FIELD_D4_UUID_CONTRATO) console.warn('[AVISO] PIPEFY_FIELD_D4_UUID_CONTRATO ausente - usando padrão: d4_uuid_contrato');
if (!PIPEFY_FIELD_D4_UUID_PROCURACAO) console.warn('[AVISO] PIPEFY_FIELD_D4_UUID_PROCURACAO ausente - usando padrão: d4_uuid_procuracao');
if (!TEMPLATE_UUID_CONTRATO) console.warn('[AVISO] TEMPLATE_UUID_CONTRATO (Marca) ausente');
if (!TEMPLATE_UUID_CONTRATO_OUTROS) console.warn('[AVISO] TEMPLATE_UUID_CONTRATO_OUTROS (Outros) ausente');
if (!TEMPLATE_UUID_PROCURACAO) console.warn('[AVISO] TEMPLATE_UUID_PROCURACAO ausente');

// Cofres mapeados por EQUIPE (campo "Equipe contrato" no Pipefy)
// ⚠️ ATENÇÃO: as chaves DEVEM ser exatamente os valores de "Equipe contrato"
const COFRES_UUIDS = {
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
  'FILIAL RIBEIRAO PRETO': COFRE_UUID_FILIALRIBEIRAOPRETO,
  'FILIAL PALMAS - TO': COFRE_UUID_FILIALPALMAS_TO,
  'Curitiba Matriz': COFRE_UUID_CURITIBA_MATRIZ,
  'CURITIBA MATRIZ': COFRE_UUID_CURITIBA_MATRIZ
};

// Função para obter o nome da variável do cofre a partir do UUID
function getNomeCofreByUuid(uuidCofre) {
  if (!uuidCofre) return 'Cofre não identificado';

  // Mapeamento reverso: UUID -> Nome da variável
  const mapeamento = {
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
    [COFRE_UUID_FILIALRIBEIRAOPRETO]: 'COFRE_UUID_FILIALRIBEIRAOPRETO',
    [COFRE_UUID_FILIALPALMAS_TO]: 'COFRE_UUID_FILIALPALMAS_TO',
    [COFRE_UUID_CURITIBA_MATRIZ]: 'COFRE_UUID_CURITIBA_MATRIZ'
  };

  if (uuidCofre === DEFAULT_COFRE_UUID) {
    return 'DEFAULT_COFRE_UUID';
  }

  return mapeamento[uuidCofre] || 'Cofre não identificado';
}

/* =========================
 * Helpers gerais
 * =======================*/
function onlyDigits(s) { return String(s || '').replace(/\D/g, ''); }
function normalizePhone(s) { return onlyDigits(s); }
function toBRL(n) { return isNaN(n) ? '' : Number(n).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' }); }
function parseNumberBR(v) {
  if (v == null) return NaN;
  const s = String(v).trim();
  if (!s) return NaN;
  if (/^\d{1,3}(\.\d{3})*,\d{2}$/.test(s)) return Number(s.replace(/\./g, '').replace(',', '.'));
  if (/^\d+(\.\d+)?$/.test(s)) return Number(s);
  return Number(s.replace(/[^\d.,-]/g, '').replace(/\./g, '').replace(',', '.'));
}
function onlyNumberBR(s) {
  const n = parseNumberBR(s);
  return isNaN(n) ? 0 : n;
}

// Datas
const MESES_PT = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];
function monthNamePt(mIndex1to12) { return MESES_PT[(Math.max(1, Math.min(12, Number(mIndex1to12))) - 1)]; }
function parsePipeDateToDate(value) {
  if (!value) return null;
  const s = String(value).trim();
  const m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (m) {
    const dd = Number(m[1]), mm = Number(m[2]), yyyy = Number(m[3]);
    const d = new Date(yyyy, mm - 1, dd);
    return isNaN(d) ? null : d;
  }
  const d = new Date(s);
  return isNaN(d) ? null : d;
}
function fmtDMY2(value) {
  const d = value instanceof Date ? value : parsePipeDateToDate(value);
  if (!d) return '';
  const dd = String(d.getDate()).padStart(2, '0');
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const yy = String(d.getFullYear()).slice(-2);
  return `${dd}/${mm}/${yy}`;
}

// Retry com timeout exponencial
async function fetchWithRetry(url, init = {}, opts = {}) {
  const attempts = opts.attempts || 3;
  const baseDelayMs = opts.baseDelayMs || 500;
  const timeoutMs = opts.timeoutMs || 15000;

  for (let i = 0; i < attempts; i++) {
    try {
      const ctrl = new AbortController();
      const t = setTimeout(() => ctrl.abort(), timeoutMs);
      const res = await fetch(url, { ...init, signal: ctrl.signal });
      clearTimeout(t);
      if (!res.ok && i < attempts - 1) {
        await new Promise(r => setTimeout(r, baseDelayMs * (i + 1)));
        continue;
      }
      return res;
    } catch (e) {
      if (i === attempts - 1) throw e;
      await new Promise(r => setTimeout(r, baseDelayMs * (i + 1)));
    }
  }
  throw new Error('fetchWithRetry: esgotou tentativas');
}

/* =========================
 * Token público (/lead/:token)
 * =======================*/
function makeLeadToken(payload) { // {cardId, ts}
  const body = Buffer.from(JSON.stringify(payload)).toString('base64url');
  const sig = crypto.createHmac('sha256', PUBLIC_LINK_SECRET).update(body).digest('base64url');
  return `${body}.${sig}`;
}
function parseLeadToken(token) {
  const [body, sig] = String(token || '').split('.');
  if (!body || !sig) throw new Error('token inválido');
  const expected = crypto.createHmac('sha256', PUBLIC_LINK_SECRET).update(body).digest('base64url');
  if (sig !== expected) throw new Error('assinatura inválida');
  const json = JSON.parse(Buffer.from(body, 'base64url').toString('utf8'));
  if (!json.cardId) throw new Error('payload inválido');
  return json;
}

/* =========================
 * Pipefy GraphQL
 * =======================*/
async function gql(query, variables) {
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
async function getCard(cardId) {
  const data = await gql(`query($id: ID!){
    card(id:$id){
      id title
      current_phase{ id name }
      pipe{ id name organization { id } }
      fields{ name value array_value field{ id type label description } }
      assignees{ name email }
    }
  }`, { id: cardId });
  if (!data?.card) throw new Error(`Card ${cardId} não encontrado`);
  return data.card;
}
async function updateCardField(cardId, fieldId, newValue) {
  await gql(`mutation($input: UpdateCardFieldInput!){
    updateCardField(input:$input){ card{ id } }
  }`, { input: { card_id: Number(cardId), field_id: fieldId, new_value: newValue } });
}

async function createCardComment(cardId, comment) {
  try {
    const data = await gql(`mutation($input: CreateCommentInput!){
      createComment(input:$input){
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
    return data?.createComment?.comment;
  } catch (e) {
    console.error('[ERRO createCardComment]', e.message || e);
    throw e;
  }
}

/* =========================
 * Parsing de campos do card
 * =======================*/
function toById(card) {
  const by = {};
  for (const f of card?.fields || []) {
    if (f?.field?.id) {
      // Para campos checklist/checkbox, o valor pode estar em array_value
      if (f.array_value && Array.isArray(f.array_value) && f.array_value.length > 0) {
        by[f.field.id] = f.array_value;
      } else {
        by[f.field.id] = f.value;
      }
    }
  }
  return by;
}
function getByName(card, nameSub) {
  const t = String(nameSub).toLowerCase();
  const f = (card.fields || []).find(ff => String(ff?.name || '').toLowerCase().includes(t));
  return f?.value || '';
}
function getFieldObjById(card, id) {
  return (card.fields || []).find(f => String(f?.field?.id || '') === String(id));
}
function getFirstByNames(card, arr) {
  for (const k of arr) { const v = getByName(card, k); if (v) return v; }
  return '';
}

// NOVO: ler campo "Equipe contrato"
function getEquipeContratoFromCard(card) {
  if (!card || !Array.isArray(card.fields)) return '';
  const f = (card.fields || []).find(ff =>
    String(ff.name || '').toLowerCase() === 'equipe contrato'
  );
  if (!f || !f.value) return '';
  return String(f.value).trim();
}

function checklistToText(v) {
  try {
    const arr = Array.isArray(v) ? v : JSON.parse(v);
    return Array.isArray(arr) ? arr.join(', ') : String(v || '');
  } catch { return String(v || ''); }
}

function extractAssigneeNames(raw) {
  const out = [];
  const push = v => { if (v) out.push(String(v)); };
  const tryParse = v => {
    if (typeof v === 'string') {
      try { return JSON.parse(v); } catch { return v; }
    }
    return v;
  };

  const val = tryParse(raw);
  if (Array.isArray(val)) {
    for (const it of val) {
      push(typeof it === 'string' ? it : (it?.name || it?.username || it?.email || it?.value));
    }
  } else if (typeof val === 'object' && val) {
    push(val.name || val.username || val.email || val.value);
  } else if (typeof val === 'string') {
    const m = val.match(/^\s*\[.*\]\s*$/) ? tryParse(val) : null;
    if (m && Array.isArray(m)) {
      m.forEach(x => push(typeof x === 'string' ? x : (x?.name || x?.email)));
    } else {
      push(val);
    }
  }
  return [...new Set(out.filter(Boolean))];
}

// Documento (CPF/CNPJ)
function pickDocumento(card) {
  const prefer = ['cpf', 'cnpj', 'documento', 'doc', 'cpf/cnpj', 'cnpj/cpf'];
  for (const k of prefer) {
    const v = getFirstByNames(card, [k]);
    const d = onlyDigits(v);
    if (d.length === 11) return { tipo: 'CPF', valor: v };
    if (d.length === 14) return { tipo: 'CNPJ', valor: v };
  }
  const by = toById(card);
  const cnpjStart = by['cnpj'] || getFirstByNames(card, ['cnpj']);
  if (cnpjStart) return { tipo: 'CNPJ', valor: cnpjStart };
  return { tipo: '', valor: '' };
}

// Assignee parsing (para cofre) — hoje não usado diretamente
function stripDiacritics(s) { return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, ''); }
function normalizeName(s) { return stripDiacritics(String(s || '').trim()).toLowerCase(); }
function resolveCofreUuidByCard(card) {
  if (!card) return DEFAULT_COFRE_UUID || null;
  const by = toById(card);
  const candidatosBrutos = [];
  if (by['vendedor_respons_vel']) candidatosBrutos.push(by['vendedor_respons_vel']);
  if (by['vendedor_respons_vel_1']) candidatosBrutos.push(by['vendedor_respons_vel_1']);
  if (by['respons_vel_5']) candidatosBrutos.push(by['respons_vel_5']);
  if (by['representante']) candidatosBrutos.push(by['representante']);
  const nomesOuEmails = candidatosBrutos.flatMap(extractAssigneeNames);
  const normKeys = Object.keys(COFRES_UUIDS || {}).reduce((acc, k) => { acc[normalizeName(k)] = COFRES_UUIDS[k]; return acc; }, {});
  for (const s of nomesOuEmails) { const n = normalizeName(s); if (normKeys[n]) return normKeys[n]; }
  if (DEFAULT_COFRE_UUID) return DEFAULT_COFRE_UUID;
  return null;
}

/* =========================
 * Regras específicas
 * =======================*/
function computeValorTaxaBRLFromFaixa(d) {
  let valorTaxaSemRS = '';
  const taxa = String(d.taxa_faixa || '');
  if (taxa.includes('440')) valorTaxaSemRS = '440,00';
  else if (taxa.includes('880')) valorTaxaSemRS = '880,00';
  return valorTaxaSemRS ? `R$ ${valorTaxaSemRS}` : '';
}

// Extrai todos os números em ordem de aparição e devolve separados por vírgula
function extractClasseNumbersFromText(s) {
  const nums = []; const seen = new Set();
  for (const m of String(s || '').matchAll(/\b\d+\b/g)) {
    const n = String(Number(m[0]));
    if (!seen.has(n)) { seen.add(n); nums.push(n); }
  }
  return nums.join(', ');
}

// Identifica o tipo base do serviço a partir do texto do statement ou connector
function serviceKindFromText(s) {
  const t = String(s || '').toUpperCase();
  if (t.includes('MARCA')) return 'MARCA';
  if (t.includes('PATENTE')) return 'PATENTE';
  if (t.includes('DESENHO')) return 'DESENHO INDUSTRIAL';
  if (t.includes('COPYRIGHT') || t.includes('DIREITO AUTORAL')) return 'COPYRIGHT/DIREITO AUTORAL';
  return 'OUTROS';
}

// Busca campo statement por N com fallback para connector
function buscarServicoN(card, n) {
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
  if (stmtId) {
    const f = getFieldObjById(card, stmtId);
    v = String(f?.value || '').replace(/<[^>]*>/g, ' ').trim();
  }
  if (!v) {
    const connId = mapConn[n];
    if (connId) {
      const f = getFieldObjById(card, connId);
      if (f?.value) {
        try {
          const arr = JSON.parse(f.value);
          if (Array.isArray(arr) && arr[0]) v = String(arr[0]);
        } catch {
          v = String(f.value);
        }
      } else if (Array.isArray(f?.array_value) && f.array_value.length) {
        v = String(f.array_value[0]);
      }
    }
  }
  return v;
}

// Retorna TODOS os serviços selecionados para o slot N (conector servi_os_marca_N)
function buscarTodosServicosN(card, n) {
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

  // Statement tem valor único — retorna como array de 1 item
  const stmtId = mapStmt[n];
  if (stmtId) {
    const f = getFieldObjById(card, stmtId);
    const v = String(f?.value || '').replace(/<[^>]*>/g, ' ').trim();
    if (v) return [v];
  }

  // Conector pode ter múltiplos valores selecionados
  const connId = mapConn[n];
  if (!connId) return [];
  const f = getFieldObjById(card, connId);
  if (!f) return [];

  if (f.value) {
    try {
      const arr = JSON.parse(f.value);
      if (Array.isArray(arr)) return arr.map(s => String(s).trim()).filter(Boolean);
      return [String(f.value).trim()].filter(Boolean);
    } catch {
      return [String(f.value).trim()].filter(Boolean);
    }
  }
  if (Array.isArray(f.array_value) && f.array_value.length) {
    return f.array_value.map(s => String(s).trim()).filter(Boolean);
  }
  return [];
}

// Monta a descrição combinada dos serviços de marca
// Ex: [“PEDIDO DE REGISTRO DE MARCA”, “CONTRA NOTIFICACAO”] → “PEDIDO DE REGISTRO DE MARCA E CONTRA NOTIFICACAO JUNTO AO INPI”
function buildDescricaoServicosMarca(servicos) {
  const partes = servicos.map(s => String(s).toUpperCase().trim()).filter(Boolean);
  if (!partes.length) return '';
  return partes.join(' E ') + ' JUNTO AO INPI';
}

// Garante que o nome da marca seja texto simples.
// Campos conector do Pipefy retornam array ou string JSON no formato “[...]” — esses valores
// não representam um nome de marca e devem ser descartados.
function extractBrandName(v) {
  if (!v) return '';
  if (Array.isArray(v)) return '';
  const s = String(v).trim();
  if (s.startsWith('[') || s.startsWith('{')) return '';
  return s;
}

// Normalização apenas para “Detalhes do serviço …”
function normalizarCabecalhoDetalhe(kind, nome, tipoMarca = '', classeNums = '') {
  const k = String(kind || '').toUpperCase();
  if (k === 'MARCA') {
    const tipo = tipoMarca ? `, Apresentação: ${tipoMarca}` : '';
    const classe = classeNums ? `, CLASSE: nº ${classeNums}` : '';
    return `MARCA: ${nome || ''}${tipo}${classe}`.trim();
  }
  if (k === 'PATENTE') return `PATENTE: ${nome || ''}`.trim();
  if (k === 'DESENHO INDUSTRIAL') return `DESENHO INDUSTRIAL: ${nome || ''}`.trim();
  if (k === 'COPYRIGHT/DIREITO AUTORAL') return `COPYRIGHT/DIREITO AUTORAL: ${nome || ''}`.trim();
  return `OUTROS SERVIÇOS: ${nome || ''}`.trim();
}

/* =========================
 * Montagem de dados do contrato
 * =======================*/
function pickParcelas(card) {
  const by = toById(card);
  let raw = by['sele_o_de_lista'] || by['quantidade_de_parcelas'] || by['numero_de_parcelas'] || '';
  if (!raw) raw = getFirstByNames(card, ['parcelas', 'quantidade de parcelas', 'nº parcelas']);
  const m = String(raw || '').match(/(\d+)/);
  return m ? m[1] : '1';
}
function pickValorAssessoria(card) {
  const by = toById(card);
  let raw = by['valor_da_assessoria'] || by['valor_assessoria'] || '';
  if (!raw) {
    const hit = (card.fields || []).find(f => String(f?.field?.type || '').toLowerCase() === 'currency');
    raw = hit?.value || '';
  }
  const n = parseNumberBR(raw);
  return isNaN(n) ? null : n;
}
function firstNonEmpty(...vals) {
  for (const v of vals) { if (String(v || '').trim()) return v; }
  return '';
}
function parseListFromLongText(value, max = 30) {
  const lines = String(value || '').split(/\r?\n/).map(s => s.trim()).filter(Boolean);
  const arr = [];
  for (let i = 0; i < max; i++) arr.push(lines[i] || '');
  return arr;
}

/**
 * Agrupa linhas por "Classe XX" / "NCL XX" / "XX -" — cada classe e suas especificações
 * ficam em uma única string, separadas por vírgula.
 * Ex: "Classe 06 - Abraçadeiras de metal, Construções de aço, ..."
 * Aceita: "Classe 06", "Classe 06 -", "NCL 06", "NCL 06 -", "06 -", "06 Abc"
 */
function parseClassesFromText(value, max = 30) {
  const lines = String(value || '').split(/\r?\n/).map(s => s.trim()).filter(Boolean);
  const classes = [];
  let current = null;

  for (const line of lines) {
    // Detecta início de nova classe: "Classe 06", "NCL 11 -", "06 -", "11 Abc", etc.
    if (/^(?:Classe\s+|NCL\s+)?\d+\s*(?:-\s*)?(?=\S)/i.test(line)) {
      if (current !== null) classes.push(current);
      current = line;
    } else if (current !== null) {
      current += ', ' + line;
    } else {
      // Linha antes de qualquer cabeçalho de classe — trata como standalone
      current = line;
    }
  }
  if (current !== null) classes.push(current);

  console.log(`[CLASSES AGRUPADAS] ${classes.length} classe(s) encontrada(s):`, classes.map(c => c.substring(0, 60) + '...'));

  // Preenche até max posições
  const arr = [];
  for (let i = 0; i < max; i++) arr.push(classes[i] || '');
  return arr;
}

async function montarDados(card) {
  const by = toById(card);

  // Marca 1 dados base
  const tituloMarca1 = by['marca'] || card.title || '';
  const marcasEspecRaw1 = by['copy_of_classe_e_especifica_es'] || by['classe'] || getFirstByNames(card, ['classes e especificações marca - 1', 'classes e especificações']) || '';
  console.log('[DEBUG CLASSES] marcasEspecRaw1 bruto:', JSON.stringify(marcasEspecRaw1));
  const linhasMarcasEspec1 = parseListFromLongText(marcasEspecRaw1, 30);
  const classesAgrupadas1 = parseClassesFromText(marcasEspecRaw1, 30);
  const classeSomenteNumeros1 = extractClasseNumbersFromText(marcasEspecRaw1);
  const tipoMarca1 = checklistToText(by['checklist_vertical'] || getFirstByNames(card, ['tipo de marca']));

  // Marca 2
  const tituloMarca2 = extractBrandName(by['marca_2']) || extractBrandName(getFirstByNames(card, ['marca ou patente - 2', 'marca - 2'])) || '';
  const marcasEspecRaw2 = by['copy_of_classes_e_especifica_es_marca_2'] || getFirstByNames(card, ['classes e especificações marca - 2']) || '';
  const linhasMarcasEspec2 = parseListFromLongText(marcasEspecRaw2, 30);
  const classesAgrupadas2 = parseClassesFromText(marcasEspecRaw2, 30);
  const classeSomenteNumeros2 = extractClasseNumbersFromText(marcasEspecRaw2);
  const tipoMarca2 = checklistToText(by['copy_of_tipo_de_marca'] || getFirstByNames(card, ['tipo de marca - 2']));

  // Marca 3
  const tituloMarca3 = extractBrandName(by['marca_3']) || extractBrandName(getFirstByNames(card, ['marca ou patente - 3', 'marca - 3'])) || '';
  const marcasEspecRaw3 = by['copy_of_copy_of_classe_e_especifica_es'] || getFirstByNames(card, ['classes e especificações marca - 3']) || '';
  const linhasMarcasEspec3 = parseListFromLongText(marcasEspecRaw3, 30);
  const classesAgrupadas3 = parseClassesFromText(marcasEspecRaw3, 30);
  const classeSomenteNumeros3 = extractClasseNumbersFromText(marcasEspecRaw3);
  const tipoMarca3 = checklistToText(by['copy_of_copy_of_tipo_de_marca'] || getFirstByNames(card, ['tipo de marca - 3']));

  // Marca 4
  const tituloMarca4 = extractBrandName(by['marca_ou_patente_4']) || '';
  const marcasEspecRaw4 = by['classes_e_especifica_es_marca_4'] || '';
  const linhasMarcasEspec4 = parseListFromLongText(marcasEspecRaw4, 30);
  const classesAgrupadas4 = parseClassesFromText(marcasEspecRaw4, 30);
  const classeSomenteNumeros4 = extractClasseNumbersFromText(marcasEspecRaw4);
  const tipoMarca4 = checklistToText(by['copy_of_tipo_de_marca_3'] || '');

  // Marca 5
  const tituloMarca5 = extractBrandName(by['marca_ou_patente_5']) || '';
  const marcasEspecRaw5 = by['copy_of_classes_e_especifica_es_marca_4'] || '';
  const linhasMarcasEspec5 = parseListFromLongText(marcasEspecRaw5, 30);
  const classesAgrupadas5 = parseClassesFromText(marcasEspecRaw5, 30);
  const classeSomenteNumeros5 = extractClasseNumbersFromText(marcasEspecRaw5);
  const tipoMarca5 = checklistToText(by['copy_of_tipo_de_marca_3_1'] || '');

  // Serviços por N
  const serv1Stmt = firstNonEmpty(buscarServicoN(card, 1));
  const serv2Stmt = firstNonEmpty(buscarServicoN(card, 2));
  const serv3Stmt = firstNonEmpty(buscarServicoN(card, 3));
  const serv4Stmt = firstNonEmpty(buscarServicoN(card, 4));
  const serv5Stmt = firstNonEmpty(buscarServicoN(card, 5));

  // Descrição combinada de todos os serviços selecionados para a marca principal
  // Ex: "PEDIDO DE REGISTRO DE MARCA E CONTRA NOTIFICACAO JUNTO AO INPI"
  const descServicoPrincipal = buildDescricaoServicosMarca(buscarTodosServicosN(card, 1));

  // Kinds
  const k1 = serviceKindFromText(serv1Stmt);
  const k2 = serviceKindFromText(serv2Stmt);
  const k3 = serviceKindFromText(serv3Stmt);
  const k4 = serviceKindFromText(serv4Stmt);
  const k5 = serviceKindFromText(serv5Stmt);

  // Risco da marca - para rotear entre Template 1 (sem risco) e Template 2 (com risco)
  const riscoRaw = by['risco_da_marca'] || '';
  const risco = String(riscoRaw).trim();

  // Equipe/Filial - vem do campo de seleção "equipe contrato"
  const equipeName = getEquipeContratoFromCard(card);

  // Representante - campo "nome_do_representante"
  const representante = by['nome_do_representante'] || '';

  // Decide template
  const anyMarca = [k1, k2, k3, k4, k5].includes('MARCA');
  const riscoAlto = ['médio', 'alto'].includes(risco.toLowerCase());

  let templateToUse;
  if (!anyMarca) {
    templateToUse = TEMPLATE_UUID_CONTRATO_OUTROS;        // Template 3: Outros Serviços
  } else if (riscoAlto) {
    templateToUse = TEMPLATE_UUID_CONTRATO_MARCA_RISCO;   // Template 2: Marca com Risco
  } else {
    templateToUse = TEMPLATE_UUID_CONTRATO;               // Template 1: Marca sem Risco
  }

  const isMarcaComRisco = anyMarca && riscoAlto;
  const isMarca = anyMarca;

  // [Lógica CPF vs CNPJ] Apenas para controle de dados, não muda o template UUID
  const selecaoCnpjOuCpf = by['cnpj_ou_cpf'] || '';
  const isSelecaoCpf = String(selecaoCnpjOuCpf).toUpperCase().trim() === 'CPF';
  const isSelecaoCnpj = String(selecaoCnpjOuCpf).toUpperCase().trim() === 'CNPJ';

  // Contato contratante 1
  const contatoNome = by['nome_1'] || getFirstByNames(card, ['nome do contato', 'contratante', 'responsável legal', 'responsavel legal']) || '';
  const contatoEmail = by['email_de_contato'] || getFirstByNames(card, ['email', 'e-mail']) || '';
  const contatoTelefone = by['telefone_de_contato'] || getFirstByNames(card, ['telefone', 'celular', 'whatsapp', 'whats']) || '';
  const nomeContato2 = by['nome_contato_2'] || '';
  const nomeContato3 = by['nome_do_contato_3'] || '';

  // Contato contratante 2 (campos de contato, distintos dos campos de envio)
  const contatoEmail2 = by['copy_of_email_de_contato'] || '';
  const contatoTelefone2 = by['copy_of_telefone_de_contato'] || '';

  // Contato contratante 3 (campos de contato, distintos dos campos de envio)
  const contatoEmail3 = by['email_do_contato_3'] || '';
  const contatoTelefone3 = by['n_mero_de_telefone_3'] || '';

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
  const cot_cep = by['cep_cotitular_1'] || '';
  const cot_rg = by['rg_cotitular'] || '';
  const cot_cpf = by['cpf_cotitular'] || '';
  const cot_cnpj = by['cnpj_cotitular'] || '';
  const cot_socio_nome = by['nome_s_cio_adminstrador_cotitular_1'] || '';
  // [ALTERADO] Prioridade CNPJ para cotitular
  const cot_docSelecao = cot_cnpj ? 'CNPJ' : (cot_cpf ? 'CPF' : '');

  // Envio do contrato principal e cotitular
  const emailEnvioContrato = by['email_para_envio_do_contrato'] || contatoEmail || '';
  // [ALTERADO] Busca primeiro no campo novo solicitado
  const emailCotitularEnvio = by['copy_of_email_para_envio_de_contrato'] || by['copy_of_email_para_envio_do_contrato'] || '';
  const telefoneCotitularEnvio = by['copy_of_telefone_para_envio_do_contrato'] || '';
  // Telefone para envio do contrato (campo específico)
  // Telefone para envio do contrato (campo específico) - SEM FALLBACK
  const telefoneEnvioContrato = by['telefone_para_envio_do_contrato'] || '';

  // [NOVO] Campos de COTITULAR 3 (Novo Cotitular)
  const cot3_ativo = by['cotitularidade_2'] || '';
  const isCot3Ativo = String(cot3_ativo).toLowerCase().trim() === 'sim';

  const cot3_nome = by['raz_o_social_ou_nome_completo_do_cotitular_2'] || by['raz_o_social_ou_nome_completo_cotitular_2'] || '';
  const cot3_nacionalidade = by['nacionalidade_cotitular_2'] || '';
  const cot3_estado_civil = by['estado_civil_cotitular_3'] || '';
  const cot3_rua = by['rua_av_do_cnpj_cotitular_2'] || '';
  const cot3_bairro = by['bairro_cotitular_2'] || '';
  const cot3_cidade = by['cidade_cotitular_2'] || '';
  const cot3_uf = by['estado_cotitular_2'] || '';
  const cot3_numero = ''; // não informado
  const cot3_cep = by['cep_cotitular_2'] || '';
  const cot3_rg = by['rg_cotitular_3'] || '';
  const cot3_cpf = by['cpf_cotitular_3'] || '';
  const cot3_cnpj = by['cnpj_cotitular_3'] || '';
  const cot3_socio_nome = by['nome_s_cio_adminstrador_cotitular_2'] || '';
  const cot3_docSelecao = cot3_cnpj ? 'CNPJ' : (cot3_cpf ? 'CPF' : '');

  // [ALTERADO] Novos campos específicos para email/telefone do Cotitular 2, com fallback para os antigos
  const emailCotitular3Envio = by['email_para_envio_do_contrato_cotitular_2'] || by['email_2'] || '';
  const telefoneCotitular3Envio = by['telefone_para_envio_do_contrato_cotitular_2'] || by['telefone_2'] || '';

  // Documento (CPF/CNPJ) principal
  const doc = pickDocumento(card);
  let cpfDoc = doc.tipo === 'CPF' ? doc.valor : '';
  let cnpjDoc = doc.tipo === 'CNPJ' ? doc.valor : '';

  // [NOVO] Mapeamento específico solicitado
  // SE O CAMPO cnpj_ou_cpf for igual a CPF -> cpf_do_s_cio_administrador
  // Caso seja CNPJ -> cnpj_1
  let cpfCampo = '';
  let cnpjCampo = '';

  if (isSelecaoCpf) {
    // CPF (Pessoa Física) → campo dedicado 'cpf'
    const rawCpf = by['cpf'] || getFirstByNames(card, ['cpf']) || '';
    cpfCampo = rawCpf;
    // Se o documento principal detectado não for consistente, forçamos o que veio do campo específico
    if (!cpfDoc && rawCpf) cpfDoc = rawCpf;
  } else if (isSelecaoCnpj) {
    const rawCnpj1 = by['cnpj_1'] || getFirstByNames(card, ['cnpj 1']);
    cnpjCampo = rawCnpj1;
    if (!cnpjDoc && rawCnpj1) cnpjDoc = rawCnpj1;
  } else {
    // Fallback original
    cpfCampo = by['cpf'] || '';
    cnpjCampo = by['cnpj_1'] || '';
  }

  // Parcelas / Assessoria
  const nParcelas = pickParcelas(card);
  const valorAssessoria = pickValorAssessoria(card);
  const formaAss = by['copy_of_tipo_de_pagamento'] || getFirstByNames(card, ['tipo de pagamento assessoria']) || '';

  // [NOVO] Campo de condições de pagamento livres
  const descreva_condicoes = by['descreva_condi_es_de_pagamento'] || '';

  // TAXA
  // [MODIFICADO] Agora usa o campo 'valor_total_da_taxa' diretamente
  const valorTaxaRaw = by['valor_total_da_taxa'] || getFirstByNames(card, ['valor_total_da_taxa']) || '';
  const taxaFaixaRaw = valorTaxaRaw; // Definindo alias para compatibilidade
  const valorTaxaBRL = valorTaxaRaw; // Assumindo que já vem formatado ou tratado depois
  const formaPagtoTaxa = by['tipo_de_pagamento'] || '';

  // PESQUISA - Tipo de pagamento e data
  const tipoPagtoPesquisa = by['copy_of_tipo_de_pagamento_taxa'] || '';
  const dataBoletoPesquisa = fmtDMY2(by['data_boleto'] || '');

  // Se tipo de pagamento não estiver preenchido, pesquisa é isenta
  const pesquisaIsenta = !tipoPagtoPesquisa || tipoPagtoPesquisa.trim() === '';
  const formaPesquisa = pesquisaIsenta ? 'ISENTA' : tipoPagtoPesquisa;

  // Data de pagamento só aparece se o tipo de pagamento for "boleto"
  const isBoleto = !pesquisaIsenta && String(tipoPagtoPesquisa).toLowerCase().includes('boleto');
  const dataPesquisa = isBoleto ? dataBoletoPesquisa : (pesquisaIsenta ? 'N/A' : '');

  // Datas novas
  const dataPagtoAssessoria = fmtDMY2(by['copy_of_copy_of_data_do_boleto_pagamento_pesquisa'] || '');
  const dataPagtoTaxa = fmtDMY2(by['copy_of_data_do_boleto_pagamento_pesquisa'] || '');

  // Endereço (CNPJ) principal
  const cepCnpj = by['cep_do_cnpj'] || '';
  const ruaCnpj = by['rua_av_do_cnpj'] || '';
  const bairroCnpj = by['bairro_do_cnpj'] || '';
  const cidadeCnpj = by['cidade_do_cnpj'] || '';
  const ufCnpj = by['estado_do_cnpj'] || '';
  const numeroCnpj = by['n_mero_1'] || getFirstByNames(card, ['numero', 'número', 'nº']) || '';

  // Vendedor (ainda usado apenas para exibição)
  const vendedor = (() => {
    const raw = by['vendedor_respons_vel'] || by['vendedor_respons_vel_1'] || by['respons_vel_5'];
    const v = raw ? extractAssigneeNames(raw) : [];
    return v[0] || '';
  })();

  // [NOVO] Sócio 2 (se houver)
  const temSocio = by['tem_s_cio'] || '';
  const isTemSocio = String(temSocio).toLowerCase().trim() === 'sim';

  const socio2Nome = by['nome_do_s_cio'] || '';
  const socio2Cpf = by['cpf_do_s_cio'] || by['cpf_do_s_cio_1'] || '';
  const socio2EstadoCivil = by['estado_c_vil_s_cio'] || '';

  // Riscos 1..5
  const risco1 = by['risco_da_marca'] || '';
  const risco2 = by['copy_of_copy_of_risco_da_marca'] || '';
  const risco3 = by['copy_of_risco_da_marca_3'] || '';
  const risco4 = by['copy_of_risco_da_marca_3_1'] || '';
  const risco5 = by['copy_of_risco_da_marca_4'] || '';

  // Nacionalidade e etc principal
  const nacionalidade = by['nacionalidade'] || '';
  // const selecaoCnpjOuCpf = by['cnpj_ou_cpf'] || ''; // Já lido acima
  const estadoCivil = by['estado_civ_l'] || '';

  // Cláusula adicional
  const filiaisOuDigital = by['filiais_ou_digital'] || '';
  const isFiliais = String(filiaisOuDigital).toLowerCase().trim() === 'filiais';

  // Verificar se o tipo de pagamento é "Crédito programado"
  const tipoPagamento = by['copy_of_tipo_de_pagamento'] || '';
  const isCreditoProgramado = String(tipoPagamento).trim() === 'Crédito programado';

  // Texto da cláusula adicional para Crédito programado
  const clausulaCreditoProgramado = 'Caso o pagamento não seja realizado até a data do vencimento, o benefício concedido será automaticamente cancelado, sendo emitido boleto bancário com os valores previstos em contrato.';

  // Montar cláusula adicional
  let clausulaAdicional = '';

  // Texto a ser removido SEMPRE (independente de Filiais ou Digital)
  // Regex atualizada para capturar o texto completo, incluindo quebras de linha e variações
  // Aceita múltiplos espaços, quebras de linha e variações no texto
  const textoObservacoesRemover = /Observações:\s*Entrada\s*R\$\s*R\$\s*440[,.]00\s*referente\s*a\s*TAXA\s*\+\s*6\s*X\s*R\$\s*R\$\s*450[,.]00\s*da\s*assessoria\s*no\s*Crédito\s*programado\.?[\s\n\r]*/gi;

  const clausulaExistente = (by['cl_usula_adicional'] && String(by['cl_usula_adicional']).trim()) ? by['cl_usula_adicional'] : '';

  console.log(`[CLAUSULA] filiais_ou_digital: "${filiaisOuDigital}", isFiliais: ${isFiliais}`);
  console.log(`[CLAUSULA] tipoPagamento: "${tipoPagamento}", isCreditoProgramado: ${isCreditoProgramado}`);
  console.log(`[CLAUSULA] clausulaExistente (antes da limpeza): "${clausulaExistente.substring(0, 200)}"`);

  // SEMPRE remove o texto de observações (independente de Filiais ou Digital)
  const antesLimpeza = clausulaExistente;
  // Tenta múltiplas variações da regex para garantir remoção
  let clausulaTmp = clausulaExistente;
  // Primeira tentativa com regex principal
  clausulaTmp = clausulaTmp.replace(textoObservacoesRemover, '');
  // Segunda tentativa - variação mais flexível (aceita qualquer coisa entre as palavras-chave)
  const regexFlexivel = /Observações[:\s]*Entrada[^\n]*440[^\n]*TAXA[^\n]*6[^\n]*X[^\n]*450[^\n]*assessoria[^\n]*Crédito[^\n]*programado[^\n]*/gi;
  clausulaTmp = clausulaTmp.replace(regexFlexivel, '');
  // Terceira tentativa - busca por padrão mais simples (apenas palavras-chave)
  const regexSimples = /Observações[^\n]*440[^\n]*450[^\n]*assessoria[^\n]*Crédito[^\n]*programado[^\n]*/gi;
  clausulaTmp = clausulaTmp.replace(regexSimples, '');
  // Quarta tentativa - busca parcial por "Observações" seguido de "440" e "450"
  const regexParcial = /Observações[^\n]*?440[^\n]*?450[^\n]*?assessoria[^\n]*?Crédito[^\n]*?programado[^\n]*/gi;
  clausulaTmp = clausulaTmp.replace(regexParcial, '');

  let clausulaLimpa = clausulaTmp.trim();
  console.log(`[CLAUSULA] Removendo observações (sempre)`);
  console.log(`[CLAUSULA] Antes da remoção (${antesLimpeza.length} chars): "${antesLimpeza.substring(0, 300)}"`);
  console.log(`[CLAUSULA] Depois da remoção (${clausulaLimpa.length} chars): "${clausulaLimpa.substring(0, 300)}"`);
  console.log(`[CLAUSULA] Texto foi removido? ${antesLimpeza.length !== clausulaLimpa.length} (diferença: ${antesLimpeza.length - clausulaLimpa.length} chars)`);
  // Remove também variações com quebras de linha
  clausulaLimpa = clausulaLimpa.replace(/\n\n+/g, '\n').trim();

  // [LÓGICA BLINDADA - CORREÇÃO DE TRAVAMENTO]
  try {
    const beneficioVal = by['tipo_de_pagamento_benef_cio'];
    // Log apenas se tiver valor para não poluir
    if (beneficioVal) console.log('[DEBUG CLAUSULA] Valor bruto beneficio:', beneficioVal);

    if (beneficioVal && String(beneficioVal).trim() === 'Logomarca gratuita') {
      console.log('[DEBUG CLAUSULA] Aplicando Logomarca Gratuita (sobrescrevendo manual)');
      clausulaLimpa = 'Logomarca gratuita';
    }
  } catch (errBeneficio) {
    console.error('[ERRO CLAUSULA] Falha crítica ao processar beneficio:', errBeneficio);
    // Em caso de erro, mantém o clausulaLimpa original
  }

  // Debug para acompanhar fluxo
  console.log(`[DEBUG CLAUSULA] clausulaLimpa após lógica de benefício: "${clausulaLimpa}"`);

  if (isCreditoProgramado) {
    // Se for Crédito programado, SEMPRE adiciona a cláusula específica (independente de Filiais/Digital)
    // Se já houver cláusula existente (após limpeza), concatena com quebra de linha; senão, usa apenas a do crédito programado
    clausulaAdicional = clausulaLimpa
      ? `${clausulaLimpa}\n\n${clausulaCreditoProgramado}`
      : clausulaCreditoProgramado;
    console.log(`[CLAUSULA] Crédito programado detectado - adicionando cláusula específica`);
  } else if (isFiliais) {
    // Se for Filiais (e não for Crédito programado), remove o texto específico de observações e não tem cláusula adicional
    clausulaAdicional = clausulaLimpa;
    // Se após remover ficar vazio, deixa vazio
    if (!clausulaAdicional) {
      clausulaAdicional = '';
    }
    console.log(`[CLAUSULA] Filiais (sem crédito programado) - cláusula final: "${clausulaAdicional.substring(0, 200)}"`);
  } else {
    // Caso padrão - remove o texto de observações se existir
    clausulaAdicional = clausulaLimpa;
    // Se após remover ficar vazio, usa o padrão
    if (!clausulaAdicional) {
      clausulaAdicional = 'Sem aditivos contratuais.';
    }
  }

  // Garantir remoção final do texto das observações (sempre, independente de Filiais ou Digital)
  if (clausulaAdicional) {
    const antesRemocaoFinal = clausulaAdicional;
    // Aplicar todas as regexes novamente para garantir remoção completa
    clausulaAdicional = clausulaAdicional.replace(textoObservacoesRemover, '');
    // Regex flexível - aceita qualquer coisa entre as palavras-chave
    const regexFlexivelFinal = /Observações[:\s]*Entrada[^\n]*440[^\n]*TAXA[^\n]*6[^\n]*X[^\n]*450[^\n]*assessoria[^\n]*Crédito[^\n]*programado[^\n]*/gi;
    clausulaAdicional = clausulaAdicional.replace(regexFlexivelFinal, '');
    // Regex simples - busca por padrão mais simples (apenas palavras-chave)
    const regexSimplesFinal = /Observações[^\n]*440[^\n]*450[^\n]*assessoria[^\n]*Crédito[^\n]*programado[^\n]*/gi;
    clausulaAdicional = clausulaAdicional.replace(regexSimplesFinal, '');
    // Regex parcial - busca parcial por "Observações" seguido de "440" e "450"
    const regexParcialFinal = /Observações[^\n]*?440[^\n]*?450[^\n]*?assessoria[^\n]*?Crédito[^\n]*?programado[^\n]*/gi;
    clausulaAdicional = clausulaAdicional.replace(regexParcialFinal, '');
    // Remover também variações com quebras de linha duplas
    clausulaAdicional = clausulaAdicional.replace(/\n\n+/g, '\n').trim();

    if (antesRemocaoFinal !== clausulaAdicional) {
      console.log(`[CLAUSULA] Remoção final aplicada - texto removido: ${antesRemocaoFinal.length - clausulaAdicional.length} chars`);
    }
  }

  console.log(`[CLAUSULA] clausulaAdicional final (primeiros 300 chars): "${clausulaAdicional.substring(0, 300)}"`);

  // Contratante 1
  // [ADICIONAL] Campos do sócio administrador para formato PJ
  const socioAdmNome = by['nome_completo_do_s_cio_administrador'] || '';
  const socioAdmCpf = by['cpf_do_s_cio_administrador'] || '';
  const numeroEnderecoCnpj = by['n_mero_endere_o_do_cnpj'] || numeroCnpj || '';

  const contratante1Texto = montarTextoContratante({
    nome: by['r_social_ou_n_completo'] || contatoNome || '',
    nacionalidade,
    estadoCivil,
    rua: ruaCnpj,
    bairro: bairroCnpj,
    numero: numeroEnderecoCnpj,
    cidade: cidadeCnpj,
    uf: ufCnpj,
    cep: cepCnpj,
    rg: by['rg'] || '',
    docSelecao: selecaoCnpjOuCpf,
    cpf: cpfCampo || cpfDoc,
    cnpj: cnpjCampo || cnpjDoc,
    telefone: contatoTelefone,
    email: contatoEmail,
    // Dados do sócio administrador para formato PJ
    socioAdmNome,
    socioAdmCpf,
    // [NOVO] Sócio 2
    socio2Nome: isTemSocio ? socio2Nome : '',
    socio2Cpf: isTemSocio ? socio2Cpf : '',
    socio2EstadoCivil: isTemSocio ? socio2EstadoCivil : ''
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
      telefone: contatoTelefone2 || telefoneCotitularEnvio || contato2Telefone_old,
      email: contatoEmail2 || emailCotitularEnvio || contato2Email_old,
      socioAdmNome: cot_socio_nome,
      socioAdmCpf: cot_cpf
    })
    : '';

  // Detecta se há cotitular 3 (BASEADO NA SINALIZAÇÃO 'Sim')
  const hasCotitular3 = isCot3Ativo;

  // Contratante 3
  const contratante3Texto = hasCotitular3
    ? montarTextoContratante({
      nome: cot3_nome || 'Cotitular 3',
      nacionalidade: cot3_nacionalidade || '',
      estadoCivil: cot3_estado_civil || '',
      rua: cot3_rua || ruaCnpj,
      bairro: cot3_bairro || bairroCnpj,
      numero: cot3_numero || '',
      cidade: cot3_cidade || cidadeCnpj,
      uf: cot3_uf || ufCnpj,
      cep: cot3_cep || '',
      rg: cot3_rg || '',
      docSelecao: cot3_docSelecao,
      cpf: cot3_cpf || '',
      cnpj: cot3_cnpj || '',
      telefone: contatoTelefone3 || telefoneCotitular3Envio,
      email: contatoEmail3 || emailCotitular3Envio,
      socioAdmNome: cot3_socio_nome,
      socioAdmCpf: cot3_cpf
    })
    : '';

  // Dados para contato 1, 2 e 3
  const dadosContato1 = [contatoNome, contatoTelefone, contatoEmail].filter(Boolean).join(' | ');
  const dadosContato2 = hasCotitular
    ? [
      (cot_nome || contato2Nome_old || 'Cotitular'),
      (contatoTelefone2 || telefoneCotitularEnvio || contato2Telefone_old || ''),
      (contatoEmail2 || emailCotitularEnvio || contato2Email_old || '')
    ].filter(Boolean).join(' | ')
    : '';
  const dadosContato3 = hasCotitular3
    ? [
      (cot3_nome || 'Cotitular 3'),
      (contatoTelefone3 || telefoneCotitular3Envio || ''),
      (contatoEmail3 || emailCotitular3Envio || '')
    ].filter(Boolean).join(' | ')
    : '';

  // Entradas consolidadas
  const entries = [
    { kind: serviceKindFromText(serv1Stmt), title: tituloMarca1, tipo: tipoMarca1, classes: classeSomenteNumeros1, stmt: serv1Stmt, risco: risco1, lines: linhasMarcasEspec1 },
    { kind: serviceKindFromText(serv2Stmt), title: tituloMarca2, tipo: tipoMarca2, classes: classeSomenteNumeros2, stmt: serv2Stmt, risco: risco2, lines: linhasMarcasEspec2 },
    { kind: serviceKindFromText(serv3Stmt), title: tituloMarca3, tipo: tipoMarca3, classes: classeSomenteNumeros3, stmt: serv3Stmt, risco: risco3, lines: linhasMarcasEspec3 },
    { kind: serviceKindFromText(serv4Stmt), title: tituloMarca4, tipo: tipoMarca4, classes: classeSomenteNumeros4, stmt: serv4Stmt, risco: risco4, lines: linhasMarcasEspec4 },
    { kind: serviceKindFromText(serv5Stmt), title: tituloMarca5, tipo: tipoMarca5, classes: classeSomenteNumeros5, stmt: serv5Stmt, risco: risco5, lines: linhasMarcasEspec5 },
  ].filter(e => String(e.title || e.stmt || '').trim());

  // Agrupamento por kind
  const byKind = { 'MARCA': [], 'PATENTE': [], 'DESENHO INDUSTRIAL': [], 'COPYRIGHT/DIREITO AUTORAL': [], 'OUTROS': [] };
  entries.forEach(e => byKind[e.kind].push(e));

  // Linhas “quantidade + descrição” (sem normalizar o texto do serviço)
  const makeQtdDescLine = (kind, arr) => {
    if (!arr.length) return '';
    const baseServico = String(arr[0].stmt || '').trim() || (kind === 'MARCA' ? 'Registro de Marca' : kind);
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
  ['MARCA', 'PATENTE', 'DESENHO INDUSTRIAL', 'COPYRIGHT/DIREITO AUTORAL', 'OUTROS'].forEach(k => {
    const arr = byKind[k];
    for (let i = 0; i < 5; i++) {
      const e = arr[i];
      if (!e) { detalhes[k][i] = ''; continue; }
      const cab = normalizarCabecalhoDetalhe(k, e.title, e.tipo, e.classes);
      detalhes[k][i] = cab;
    }
  });

  // Cabeçalhos “SERVIÇOS” para classes
  const headersServicos = {
    h1: byKind['MARCA'][0] ? `MARCA: ${byKind['MARCA'][0].title || ''}` : '',
    h2: byKind['MARCA'][1] ? `MARCA: ${byKind['MARCA'][1].title || ''}` : ''
  };

  // Risco agregado formatado com nome do tipo e do item
  const riscoAgregado = entries
    .map(e => {
      const tipo = e.kind || '';
      const nm = e.title || '';
      const r = String(e.risco || '').trim();
      if (!tipo && !nm && !r) return '';
      return `${tipo}: ${nm} - RISCO: ${r || 'Não informado'}`;
    })
    .filter(Boolean)
    .join(', ');

  console.log(`[MONTAR_DADOS] card.id: ${card.id}, tipo: ${typeof card.id}`);
  const cardIdValue = card.id ? String(card.id) : '';
  console.log(`[MONTAR_DADOS] cardIdValue: "${cardIdValue}"`);

  return {
    cardId: cardIdValue,
    templateToUse,

    // Identificação
    titulo: tituloMarca1 || card.title || '',
    nome: contatoNome || (by['r_social_ou_n_completo'] || ''),
    cpf: cpfDoc,
    cnpj: cnpjDoc,
    rg: by['rg'] || '',
    estado_civil: estadoCivil,

    // Doc específicos
    cpf_campo: cpfCampo,
    cnpj_campo: cnpjCampo,
    selecao_cnpj_ou_cpf: selecaoCnpjOuCpf,
    nacionalidade,

    // Sócio administrador (para templates CNPJ — Termo de Risco CNPJ)
    socio_adm_nome: socioAdmNome || '',
    socio_adm_cpf: socioAdmCpf || '',

    // Contatos
    email: contatoEmail || '',
    telefone: contatoTelefone || '',
    dados_contato_1: dadosContato1,
    dados_contato_2: dadosContato2,
    dados_contato_3: dadosContato3,

    // Textos completos dos contratantes
    contratante_1_texto: contratante1Texto,
    contratante_2_texto: contratante2Texto,
    contratante_3_texto: contratante3Texto, // [NOVO]

    // Nomes dos contratantes para campos de assinatura
    // Se for CNPJ (isSelecaoCnpj), usa o nome do sócio administrador
    nome_contratante_1: ((isSelecaoCnpj && socioAdmNome) ? socioAdmNome : (by['r_social_ou_n_completo'] || contatoNome || '')).toUpperCase(),
    nome_contratante_2: (hasCotitular ? (cot_nome || contato2Nome_old || '') : '').toUpperCase(),
    nome_contratante_3: (hasCotitular3 ? (cot3_nome || '') : '').toUpperCase(),

    // Email para assinatura
    email_envio_contrato: emailEnvioContrato,
    email_cotitular_envio: emailCotitularEnvio,
    email_cotitular_3_envio: emailCotitular3Envio, // [NOVO]

    // Telefone para envio
    telefone_envio_contrato: telefoneEnvioContrato,
    telefone_cotitular_envio: telefoneCotitularEnvio,
    telefone_cotitular_3_envio: telefoneCotitular3Envio, // [NOVO]

    // MARCA 1..5: linhas e cabeçalhos do formulário
    cabecalho_servicos_1: headersServicos.h1,
    cabecalho_servicos_2: headersServicos.h2,

    linhas_marcas_espec_1: linhasMarcasEspec1,
    linhas_marcas_espec_2: linhasMarcasEspec2,
    linhas_marcas_espec_3: linhasMarcasEspec3,
    linhas_marcas_espec_4: linhasMarcasEspec4,
    linhas_marcas_espec_5: linhasMarcasEspec5,
    classes_agrupadas_1: classesAgrupadas1,
    classes_agrupadas_2: classesAgrupadas2,
    classes_agrupadas_3: classesAgrupadas3,
    classes_agrupadas_4: classesAgrupadas4,
    classes_agrupadas_5: classesAgrupadas5,

    // Quantidades e descrições por categoria
    qtd_desc: {
      MARCA: qtdDesc['MARCA'],
      PATENTE: qtdDesc['PATENTE'],
      DI: qtdDesc['DESENHO INDUSTRIAL'],
      COPY: qtdDesc['COPYRIGHT/DIREITO AUTORAL'],
      OUTROS: qtdDesc['OUTROS']
    },

    // Contagem numérica de marcas (só o número, sem descrição)
    qtd_marca_num: byKind['MARCA'].length || '',

    // Detalhes por categoria até 5
    det: detalhes,

    // Classes e tipos por marca
    desc_servico_marca: descServicoPrincipal,

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
    valor_pesquisa: pesquisaIsenta ? 'ISENTA' : 'R$ 98,00',
    forma_pesquisa: formaPesquisa,
    data_pesquisa: dataPesquisa,

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
    clausula_adicional: clausulaAdicional,

    // Nomes para assinatura da procuração
    nome_assinatura_1: contatoNome || '',
    nome_assinatura_2: nomeContato2 || '',
    nome_assinatura_3: nomeContato3 || '',

    // [NOVO] Campo descreva condições de pagamento (raw do Pipefy)
    descreva_condicoes_de_pagamento: descreva_condicoes,

    // [NOVO] Risco e flags de template
    risco: risco,
    isMarcaComRisco: isMarcaComRisco,
    isMarca: isMarca,

    // [NOVO] Filial (vem do campo "equipe contrato")
    equipeName: equipeName,

    // [NOVO] Representante (campo "nome_do_representante")
    representante: representante
  };
}

// NOVA VERSÃO — Qualificação separada para CPF x CNPJ
function montarTextoContratante(info = {}) {
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
    email,
    // Dados do sócio administrador (para PJ/CNPJ)
    socioAdmNome = '',
    socioAdmCpf = '',
    // Sócio 2
    socio2Nome = '',
    socio2Cpf = '',
    socio2EstadoCivil = ''
  } = info;

  const cpfDigits = onlyDigits(cpf);
  const cnpjDigits = onlyDigits(cnpj);

  // [MODIFICADO] Prioridade absoluta para a seleção do Pipefy (docSelecao)
  // Se docSelecao for 'CNPJ', tratamos como CNPJ.
  // Se docSelecao for 'CPF', tratamos como CPF.
  // Se não estiver definido, tentamos inferir pelo tamanho do documento.

  let isCnpj = false;
  let isCpf = false;

  const selecao = String(docSelecao || '').toUpperCase().trim();

  if (selecao === 'CNPJ') {
    isCnpj = true;
  } else if (selecao === 'CPF') {
    isCpf = true;
  } else {
    // Fallback: inferência automática
    isCnpj = cnpjDigits.length === 14;
    isCpf = !isCnpj && cpfDigits.length === 11;
  }

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
  // CNPJ → Pessoa Jurídica (NOVO FORMATO EM MAIÚSCULAS)
  // ===============================
  if (isCnpj) {
    const razao = (nome || 'Razão Social não informada').toUpperCase();

    let cnpjFmt = cnpj || '';
    if (cnpjDigits.length === 14) {
      cnpjFmt = cnpjDigits.replace(
        /^(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})$/,
        '$1.$2.$3/$4-$5'
      );
    }

    // Montar endereço em formato maiúsculo
    const enderecoPartesPJ = [];
    if (rua) enderecoPartesPJ.push(rua.toUpperCase());
    if (numero) enderecoPartesPJ.push(numero);
    if (bairro) enderecoPartesPJ.push(bairro.toUpperCase());
    if (cidade) enderecoPartesPJ.push(cidade.toUpperCase());
    if (uf) enderecoPartesPJ.push(uf.toUpperCase());
    if (cep) {
      // Formatar CEP com ponto (80.010-010)
      const cepDigits = onlyDigits(cep);
      const cepFmt = cepDigits.length === 8
        ? cepDigits.replace(/^(\d{2})(\d{3})(\d{3})$/, '$1.$2-$3')
        : cep;
      enderecoPartesPJ.push(`CEP: ${cepFmt}`);
    }
    const enderecoPJ = enderecoPartesPJ.join(' - ');

    // Montar texto base: RAZÃO SOCIAL, CNPJ: XX.XXX.XXX/XXXX-XX, LOCALIZADA NA RUA: ...
    let textoPJ = `${razao}, CNPJ: ${cnpjFmt}`;

    if (enderecoPJ) {
      textoPJ += `, LOCALIZADA NA RUA: ${enderecoPJ}`;
    }

    // Adicionar dados do sócio administrador/representante legal
    if (socioAdmNome) {
      const socioNomeUpper = socioAdmNome.toUpperCase();
      const nacionalidadeUpper = (nacionalidade || 'BRASILEIRO(A)').toUpperCase();
      const estadoCivilUpper = (estadoCivil || '').toUpperCase();

      if (socio2Nome) {
        // Plural
        textoPJ += `, NESTE ATO REPRESENTADA POR SEUS SÓCIOS ADMINSTRADORES: SR(A). ${socioNomeUpper}`;
      } else {
        // Singular
        textoPJ += `, NESTE ATO REPRESENTADO(A) POR SEU SÓCIO(A) ADMINISTRADOR(A) SR(A). ${socioNomeUpper}`;
      }

      // Adicionar qualificação do sócio
      const qualificacao = [nacionalidadeUpper];
      if (estadoCivilUpper) qualificacao.push(estadoCivilUpper);
      qualificacao.push('EMPRESÁRIO(A)');

      textoPJ += `, ${qualificacao.join(', ')}`;

      // CPF do sócio administrador
      if (socioAdmCpf) {
        const socioAdmCpfDigits = onlyDigits(socioAdmCpf);
        let cpfSocioFmt = socioAdmCpf;
        if (socioAdmCpfDigits.length === 11) {
          cpfSocioFmt = socioAdmCpfDigits.replace(
            /^(\d{3})(\d{3})(\d{3})(\d{2})$/,
            '$1.$2.$3-$4'
          );
        }
        textoPJ += `, PORTADOR(A) DO CPF: ${cpfSocioFmt}`;
      }

      // Adicionar Segundo Sócio se existir
      if (socio2Nome) {
        const socio2NomeUpper = socio2Nome.toUpperCase();
        textoPJ += ` E ${socio2NomeUpper}, BRASILEIRO(A)`;
        if (socio2EstadoCivil) textoPJ += `, ${socio2EstadoCivil.toUpperCase()}`;
        textoPJ += `, EMPRESÁRIO(A)`;

        if (socio2Cpf) {
          const s2CpfDigits = onlyDigits(socio2Cpf);
          let s2CpfFmt = socio2Cpf;
          if (s2CpfDigits.length === 11) {
            s2CpfFmt = s2CpfDigits.replace(
              /^(\d{3})(\d{3})(\d{3})(\d{2})$/,
              '$1.$2.$3-$4'
            );
          }
          textoPJ += `, PORTADOR DO CPF ${s2CpfFmt}`;
        }
      }
    }

    return textoPJ.endsWith('.') ? textoPJ : `${textoPJ}.`;
  }

  // ===============================
  // CPF (ou genérico) → UPPERCASE para padronizar com CNPJ
  // ===============================
  const partes = [];
  const identidade = [];

  if (nome) identidade.push(nome.toUpperCase());
  if (nacionalidade) identidade.push(nacionalidade.toUpperCase());
  if (estadoCivil) identidade.push(estadoCivil.toUpperCase());
  if (identidade.length) identidade.push('EMPRESÁRIO(A)');
  if (identidade.length) partes.push(identidade.join(', '));

  if (enderecoStr) partes.push(`RESIDENTE NA ${enderecoStr.toUpperCase()}`);

  const documentos = [];
  if (rg) documentos.push(`PORTADOR(A) DA CÉDULA DE IDENTIDADE RG DE Nº ${rg.toUpperCase()}`);

  // Preferência: se tiver CPF com 11 dígitos, usa "PORTADOR(A) DO CPF Nº ..."
  if (isCpf && cpfDigits) {
    const cpfFmt = cpfDigits.replace(
      /^(\d{3})(\d{3})(\d{3})(\d{2})$/,
      '$1.$2.$3-$4'
    );
    documentos.push(`PORTADOR(A) DO CPF Nº ${cpfFmt}`);
  } else {
    const docUpper = String(docSelecao || '').trim().toUpperCase();
    const docNums = [];
    if (cpf) docNums.push({ tipo: 'CPF', valor: cpf });
    if (cnpj && !isCnpj) docNums.push({ tipo: 'CNPJ', valor: cnpj });

    if (docUpper && docNums.length) {
      documentos.push(`DEVIDAMENTE INSCRITO NO ${docUpper} SOB O Nº ${docNums[0].valor}`);
    } else {
      for (const doc of docNums) {
        documentos.push(`DEVIDAMENTE INSCRITO NO ${doc.tipo} SOB O Nº ${doc.valor}`);
      }
    }
  }

  if (documentos.length) partes.push(documentos.join(', '));

  const contatoPartes = [];
  if (telefone) contatoPartes.push(`COM TELEFONE DE Nº ${telefone}`);
  if (email) contatoPartes.push(`COM O SEGUINTE ENDEREÇO ELETRÔNICO: ${email.toUpperCase()}`);
  if (contatoPartes.length) partes.push(contatoPartes.join(' E '));

  if (!partes.length) return '';
  const texto = partes.join(', ').replace(/\s+,/g, ',').trim();
  return texto.endsWith('.') ? texto : `${texto}.`;
}

/* =========================
 * Variáveis para Templates
 * =======================*/

// ===== Funções de Concatenação de Pagamento (Nova lógica) =====

/**
 * Monta a string do campo "contrato da Assessoria"
 * Ex: "3 parcela(s) de R$ 500,00, valor que será pago via Boleto com a primeira parcela para o dia 10/04/2025"
 */
function montarTextoAssessoria({ parcelas, valorParcela, formaPagamento, dataPagamento }) {
  const base = `${parcelas} parcela(s) de ${valorParcela}, valor que será pago via ${formaPagamento}`;
  const formasComData = ['boleto', 'crédito programado'];
  const incluiData = formasComData.includes(String(formaPagamento).toLowerCase().trim());
  if (incluiData && dataPagamento) {
    return `${base} com a primeira parcela para o dia ${dataPagamento}`;
  }
  return base;
}

/**
 * Monta a string do campo "contrato da Pesquisa"
 * Ex: "Isenta"  ou  "R$ 98,00 valor que será pago via Boleto 10/04/2025"
 */
function montarTextoPesquisa({ tipoPesquisa, valorPesquisa, formaPagamentoPesquisa, dataPagamentoPesquisa }) {
  if (String(tipoPesquisa).toLowerCase().trim() === 'isenta') {
    return 'Isenta';
  }
  // Caso "Paga"
  return `${valorPesquisa} valor que será pago via ${formaPagamentoPesquisa} ${dataPagamentoPesquisa}`.trim();
}

/**
 * Monta a string do campo "contrato financeiro" (Taxas e Encaminhamentos)
 * Ex: "R$ 2.900,00 valor que será pago via Boleto dia 10/04/2025"
 */
function montarTextoTaxa({ valorTaxa, formaPagamentoTaxa, dataPagamentoTaxa }) {
  const base = `${valorTaxa} valor que será pago via ${formaPagamentoTaxa}`;
  const formasComData = ['boleto', 'crédito programado'];
  const incluiData = formasComData.includes(String(formaPagamentoTaxa).toLowerCase().trim());
  if (incluiData && dataPagamentoTaxa) {
    return `${base} dia ${dataPagamentoTaxa}`;
  }
  return base;
}

// Marca
function montarVarsParaTemplateMarca(d, nowInfo) {
  const valorTotalNum = onlyNumberBR(d.valor_total);
  const parcelaNum = parseInt(String(d.parcelas || '1'), 10) || 1;
  const valorParcela = parcelaNum > 0 ? valorTotalNum / parcelaNum : 0;

  const dia = String(nowInfo.dia).padStart(2, '0');
  const mesExtenso = monthNamePt(nowInfo.mes);
  const ano = String(nowInfo.ano);

  const cardIdStr = String(d.cardId || '');
  console.log(`[TEMPLATE MARCA] cardId para número do contrato: ${cardIdStr}`);
  console.log(`[TEMPLATE MARCA] d.cardId: ${d.cardId}, tipo: ${typeof d.cardId}`);
  console.log(`[TEMPLATE MARCA] cardIdStr final: "${cardIdStr}"`);

  const base = {
    // Identificação - Número do contrato (múltiplas variações para compatibilidade)
    'N° contrato': cardIdStr,
    'Nº contrato': cardIdStr,
    'Numero contrato': cardIdStr,
    'Número contrato': cardIdStr,
    'CONTRATO nº': cardIdStr,
    'CONTRATO Nº': cardIdStr,
    'CONTRATO N°': cardIdStr,
    'CONTRATO nº:': cardIdStr,
    'CONTRATO Nº:': cardIdStr,
    'CONTRATO N°:': cardIdStr,
    'contrato nº': cardIdStr,
    'contrato n°': cardIdStr,
    'contrato nº:': cardIdStr,
    'contrato n°:': cardIdStr,
    'numero contrato': cardIdStr,
    'numero do contrato': cardIdStr,
    'Número do contrato': cardIdStr,
    // Variações para cabeçalho
    'N° de contrato': cardIdStr,
    'Nº de contrato': cardIdStr,
    'Número de contrato': cardIdStr,
    'Numero de contrato': cardIdStr,
    'CONTRATO N°:': cardIdStr,
    'CONTRATO Nº:': cardIdStr,
    'Contrato N°': cardIdStr,
    'Contrato Nº': cardIdStr,
    'Contrato nº': cardIdStr,
    'Contrato n°': cardIdStr,
    // Campo específico do D4Sign - [ATUALIZADO] Renomeado para NContrato
    'NContrato': cardIdStr,
    'NContrato*': cardIdStr,
    // Manter compatibilidade com variáveis antigas (caso usadas em templates)
    'Número do contrato do bloco físico*': cardIdStr,
    'Número do contrato do bloco físico': cardIdStr,
    'Numero do contrato do bloco fisico': cardIdStr,

    // [NOVO] Filial e Representante
    'Filial': d.equipeName || '',  // Vem direto do campo "equipe contrato"
    'Representante': d.representante || '',  // Campo "nome_do_representante" do Pipefy

    'Contratante 1': d.contratante_1_texto || d.nome || '',
    'Contratante 2': d.contratante_2_texto || '',
    'Contratante 3': d.contratante_3_texto || '',
    'CONTRATANTE 3': d.contratante_3_texto || '',
    'contratante_1': d.contratante_1_texto || d.nome || '',
    'contratante_2': d.contratante_2_texto || '',
    'contratante_3': d.contratante_3_texto || '',
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
    'dados para contato 3': d.dados_contato_3 || '',

    // Resultado da pesquisa prévia
    'Risco': d.risco_agregado || '',

    // Quantidade e descrição de Marca
    'Quantidade depósitos/processos de MARCA': d.qtd_marca_num || '',
    'Descrição do serviço - MARCA': d.desc_servico_marca || '',

    // Detalhes do serviço - Marca até 5
    'Detalhes do serviço - MARCA': d.det.MARCA[0] || '',
    'Detalhes do serviço - MARCA 2': d.nome2 ? (d.det.MARCA[1] || '') : '',
    'Detalhes do serviço - MARCA 3': d.nome3 ? (d.det.MARCA[2] || '') : '',
    'Detalhes do serviço - MARCA 4': d.nome4 ? (d.det.MARCA[3] || '') : '',
    'Detalhes do serviço - MARCA 5': d.nome5 ? (d.det.MARCA[4] || '') : '',

    // Formulário de Classes
    'Cabeçalho - SERVIÇOS': d.cabecalho_servicos_1 || '',
    // Tipo de marca (token no template Word: ${"tipo de marca"})
    'tipo de marca': d.tipo1 || '',
    'tipo de marca 2': d.nome2 ? (d.tipo2 || '') : '',
    'tipo de marca 3': d.nome3 ? (d.tipo3 || '') : '',
    'tipo de marca 4': d.nome4 ? (d.tipo4 || '') : '',
    'tipo de marca 5': d.nome5 ? (d.tipo5 || '') : '',
    // [NOVO] Nome da marca principal - campo "marca" do Pipefy
    'nome_da_marca': d.nome1 || d.titulo || '',
    // Classes agrupadas por "Classe XX" / "NCL XX" com especificações separadas por vírgula
    'marcas-espec_1': d.classes_agrupadas_1[0] || '',
    'marcas-espec_2': d.classes_agrupadas_1[1] || '',
    'marcas-espec_3': d.classes_agrupadas_1[2] || '',
    'marcas-espec_4': d.classes_agrupadas_1[3] || '',
    'marcas-espec_5': d.classes_agrupadas_1[4] || '',

    'Cabeçalho - SERVIÇOS 2': d.nome2 ? (d.cabecalho_servicos_2 || '') : '',
    'marcas2-espec_1': d.nome2 ? (d.linhas_marcas_espec_2[0] || '') : '',
    'marcas2-espec_2': d.nome2 ? (d.linhas_marcas_espec_2[1] || '') : '',
    'marcas2-espec_3': d.nome2 ? (d.linhas_marcas_espec_2[2] || '') : '',
    'marcas2-espec_4': d.nome2 ? (d.linhas_marcas_espec_2[3] || '') : '',
    'marcas2-espec_5': d.nome2 ? (d.linhas_marcas_espec_2[4] || '') : '',

    // Assessoria (campos individuais mantidos para compatibilidade)
    'Número de parcelas da Assessoria': String(d.parcelas || '1'),
    'Valor da parcela da Assessoria': toBRL(valorParcela),
    'Forma de pagamento da Assessoria': d.forma_pagto_assessoria || '',
    'Data de pagamento da Assessoria': d.data_pagto_assessoria || '',
    // [NOVO] Campo consolidado gerado pela função
    'contrato da Assessoria': montarTextoAssessoria({
      parcelas: d.parcelas || '1',
      valorParcela: toBRL(valorParcela),
      formaPagamento: d.forma_pagto_assessoria || '',
      dataPagamento: d.data_pagto_assessoria || ''
    }),

    // Pesquisa (campos individuais mantidos para compatibilidade)
    'Valor da Pesquisa': d.valor_pesquisa || 'R$ 00,00',
    'Forma de pagamento da Pesquisa': d.forma_pesquisa || '',
    'Data de pagamento da pesquisa': d.data_pesquisa || '00/00/00',
    // [NOVO] Campo consolidado gerado pela função
    // Pesquisa só se aplica a Marca, Patente ou Desenho Industrial
    'contrato da Pesquisa': (d.qtd_desc.MARCA || d.qtd_desc.PATENTE || d.qtd_desc.DI)
      ? montarTextoPesquisa({
          tipoPesquisa: d.forma_pesquisa || '',
          valorPesquisa: d.valor_pesquisa || 'R$ 98,00',
          formaPagamentoPesquisa: d.forma_pesquisa || '',
          dataPagamentoPesquisa: d.data_pesquisa || ''
        })
      : 'Não se aplica',

    // Taxa (campos individuais mantidos para compatibilidade)
    'Valor da Taxa': d.valor_taxa_brl || '',
    'Forma de pagamento da Taxa': d.forma_pagto_taxa || '',
    'Data de pagamento da Taxa': d.data_pagto_taxa || '',
    // [NOVO] Campo consolidado gerado pela função
    'contrato financeiro': montarTextoTaxa({
      valorTaxa: d.valor_taxa_brl || '',
      formaPagamentoTaxa: d.forma_pagto_taxa || '',
      dataPagamentoTaxa: d.data_pagto_taxa || ''
    }),

    // Datas
    Dia: dia,
    Mês: mesExtenso,
    Mes: mesExtenso,
    Ano: ano,
    Cidade: d.cidade_cnpj || '',
    UF: d.uf_cnpj || '',

    // Cláusula adicional
    'clausula-adicional': d.clausula_adicional || '',

    // [NOVO] Condições de pagamento (campo condicional)
    // Se o campo 'Descreva Condições de Pagamento' (descreva_condi_es_de_pagamento) vier preenchido,
    // usa esse texto; caso contrário monta o texto padrão com Taxa + Parcelas de Assessoria.
    'Condicoes de pagamento': (() => {
      if (d.descreva_condicoes_de_pagamento && String(d.descreva_condicoes_de_pagamento).trim()) {
        return String(d.descreva_condicoes_de_pagamento).trim();
      }
      const taxa = d.valor_taxa_brl || '';
      const nParcelasStr = String(d.parcelas || '1');
      const valorParcelaCalc = toBRL(parcelaNum > 0 ? valorTotalNum / parcelaNum : 0);
      const formaPagto = d.forma_pagto_assessoria || '';
      return `Entrada R$ ${taxa} referente a TAXA + ${nParcelasStr} X ${valorParcelaCalc} da assessoria no ${formaPagto}.`;
    })()
  };

  // Preencher até 30 linhas por segurança
  for (let i = 5; i < 30; i++) {
    base[`marcas-espec_${i + 1}`] = d.classes_agrupadas_1[i] || '';
    base[`marcas2-espec_${i - 4}`] = d.nome2 ? (d.linhas_marcas_espec_2[i - 5] || '') : '';
  }

  return base;
}

// Outros
function montarVarsParaTemplateOutros(d, nowInfo) {
  const valorTotalNum = onlyNumberBR(d.valor_total);
  const parcelaNum = parseInt(String(d.parcelas || '1'), 10) || 1;
  const valorParcela = parcelaNum > 0 ? valorTotalNum / parcelaNum : 0;

  const dia = String(nowInfo.dia).padStart(2, '0');
  const mesExtenso = monthNamePt(nowInfo.mes);
  const ano = String(nowInfo.ano);

  const cardIdStr = String(d.cardId || '');
  console.log(`[TEMPLATE OUTROS] cardId para número do contrato: ${cardIdStr}`);
  console.log(`[TEMPLATE OUTROS] d.cardId: ${d.cardId}, tipo: ${typeof d.cardId}`);
  console.log(`[TEMPLATE OUTROS] cardIdStr final: "${cardIdStr}"`);

  const base = {
    // Identificação - Número do contrato (múltiplas variações para compatibilidade)
    'N° contrato': cardIdStr,
    'Nº contrato': cardIdStr,
    'Numero contrato': cardIdStr,
    'Número contrato': cardIdStr,
    'CONTRATO nº': cardIdStr,
    'CONTRATO Nº': cardIdStr,
    'CONTRATO N°': cardIdStr,
    'CONTRATO nº:': cardIdStr,
    'CONTRATO Nº:': cardIdStr,
    'CONTRATO N°:': cardIdStr,
    'contrato nº': cardIdStr,
    'contrato n°': cardIdStr,
    'contrato nº:': cardIdStr,
    'contrato n°:': cardIdStr,
    'numero contrato': cardIdStr,
    'numero do contrato': cardIdStr,
    'Número do contrato': cardIdStr,
    // Variações para cabeçalho
    'N° de contrato': cardIdStr,
    'Nº de contrato': cardIdStr,
    'Número de contrato': cardIdStr,
    'Numero de contrato': cardIdStr,
    'Contrato N°': cardIdStr,
    'Contrato Nº': cardIdStr,
    'Contrato nº': cardIdStr,
    'Contrato n°': cardIdStr,
    // Campo específico do D4Sign - [ATUALIZADO] Renomeado para NContrato
    'NContrato': cardIdStr,
    'NContrato*': cardIdStr,
    // Manter compatibilidade com variáveis antigas (caso usadas em templates)
    'Número do contrato do bloco físico*': cardIdStr,
    'Número do contrato do bloco físico': cardIdStr,
    'Numero do contrato do bloco fisico': cardIdStr,

    // [NOVO] Filial e Representante
    'Filial': d.equipeName || '',  // Vem direto do campo "equipe contrato"
    'Representante': d.representante || '',  // Campo "nome_do_representante" do Pipefy

    'Contratante 1': d.contratante_1_texto || d.nome || '',
    'Contratante 2': d.contratante_2_texto || '',
    'Contratante 3': d.contratante_3_texto || '',
    'CONTRATANTE 3': d.contratante_3_texto || '',
    'contratante_1': d.contratante_1_texto || d.nome || '',
    'contratante_2': d.contratante_2_texto || '',
    'contratante_3': d.contratante_3_texto || '',
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
    'dados para contato 3': d.dados_contato_3 || '',

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

    // Assessoria (campos individuais mantidos para compatibilidade)
    'Número de parcelas da Assessoria': String(d.parcelas || '1'),
    'Valor da parcela da Assessoria': toBRL(valorParcela),
    'Forma de pagamento da Assessoria': d.forma_pagto_assessoria || '',
    'Data de pagamento da Assessoria': d.data_pagto_assessoria || '',
    // [NOVO] Campo consolidado gerado pela função
    'contrato da Assessoria': montarTextoAssessoria({
      parcelas: d.parcelas || '1',
      valorParcela: toBRL(valorParcela),
      formaPagamento: d.forma_pagto_assessoria || '',
      dataPagamento: d.data_pagto_assessoria || ''
    }),

    // Pesquisa (campos individuais mantidos para compatibilidade)
    'Valor da Pesquisa': d.valor_pesquisa || 'R$ 00,00',
    'Forma de pagamento da Pesquisa': d.forma_pesquisa || '',
    'Data de pagamento da pesquisa': d.data_pesquisa || '00/00/00',
    // [NOVO] Campo consolidado gerado pela função
    // Pesquisa só se aplica a Marca, Patente ou Desenho Industrial
    'contrato da Pesquisa': (d.qtd_desc.MARCA || d.qtd_desc.PATENTE || d.qtd_desc.DI)
      ? montarTextoPesquisa({
          tipoPesquisa: d.forma_pesquisa || '',
          valorPesquisa: d.valor_pesquisa || 'R$ 98,00',
          formaPagamentoPesquisa: d.forma_pesquisa || '',
          dataPagamentoPesquisa: d.data_pesquisa || ''
        })
      : 'Não se aplica',

    // Taxa (campos individuais mantidos para compatibilidade)
    'Valor da Taxa': d.valor_taxa_brl || '',
    'Forma de pagamento da Taxa': d.forma_pagto_taxa || '',
    'Data de pagamento da Taxa': d.data_pagto_taxa || '',
    // [NOVO] Campo consolidado gerado pela função
    'contrato financeiro': montarTextoTaxa({
      valorTaxa: d.valor_taxa_brl || '',
      formaPagamentoTaxa: d.forma_pagto_taxa || '',
      dataPagamentoTaxa: d.data_pagto_taxa || ''
    }),

    // Datas
    Dia: dia,
    Mês: mesExtenso,
    Mes: mesExtenso,
    Ano: ano,
    Cidade: d.cidade_cnpj || '',
    UF: d.uf_cnpj || '',

    // Cláusula adicional
    'clausula-adicional': d.clausula_adicional || '',

    // [NOVO] Condições de pagamento (campo condicional)
    // Se o campo 'Descreva Condições de Pagamento' (descreva_condi_es_de_pagamento) vier preenchido,
    // usa esse texto; caso contrário monta o texto padrão com Taxa + Parcelas de Assessoria.
    'Condicoes de pagamento': (() => {
      if (d.descreva_condicoes_de_pagamento && String(d.descreva_condicoes_de_pagamento).trim()) {
        return String(d.descreva_condicoes_de_pagamento).trim();
      }
      const taxa = d.valor_taxa_brl || '';
      const nParcelas = String(d.parcelas || '1');
      const valorParcela = toBRL(parcelaNum > 0 ? valorTotalNum / parcelaNum : 0);
      const formaPagto = d.forma_pagto_assessoria || '';
      return `Entrada R$ ${taxa} referente a TAXA + ${nParcelas} X ${valorParcela} da assessoria no ${formaPagto}.`;
    })()
  };

  return base;
}

// Procuração
function montarVarsParaTemplateProcuracao(d, nowInfo) {
  const dia = String(nowInfo.dia).padStart(2, '0');
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
    'Contratante 3': d.contratante_3_texto || '',
    'CONTRATANTE 3': d.contratante_3_texto || '',
    'contratante_1': d.contratante_1_texto || d.nome || '',
    'contratante_2': d.contratante_2_texto || '',
    'contratante_3': d.contratante_3_texto || '',
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
    'dados para contato 3': d.dados_contato_3 || '',

    // Datas
    'Dia': dia,
    'Mês': mesExtenso,
    'Mes': mesExtenso,
    'Ano': ano,
    'Data': `${dia} de ${mesExtenso} de ${ano}`,

    // Informações do contrato relacionadas
    'Título': d.titulo || '',
    'Serviços': d.qtd_desc.MARCA || d.qtd_desc.PATENTE || d.qtd_desc.OUTROS || '',
    'Risco': d.risco_agregado || '',

    // Assinatura dos contratantes (apenas nomes)
    'ASSINATURA CONTRATANTE 1': d.nome_contratante_1 || d.nome || '',
    'ASSINATURA CONTRATANTE 2': d.nome_contratante_2 || '',
    'ASSINATURA CONTRATANTE 3': d.nome_contratante_3 || '',
    'Assinatura Contratante 1': d.nome_contratante_1 || d.nome || '',
    'Assinatura Contratante 2': d.nome_contratante_2 || '',
    'Assinatura Contratante 3': d.nome_contratante_3 || ''
  };

  return base;
}

// Assinantes: principal + empresa + cotitular quando houver
function montarSigners(d, incluirTelefone = false) {
  const list = [];
  const emailPrincipal = d.email_envio_contrato || d.email || '';
  // [MODIFICADO] Sem fallback para telefone genérico, conforme solicitado
  const telefonePrincipal = d.telefone_envio_contrato || '';

  if (emailPrincipal) {
    const signer = {
      email: emailPrincipal,
      name: d.nome || d.titulo || emailPrincipal,
      act: '1',
      foreign: '0',
      send_email: '1'
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
      act: '1',
      foreign: '0',
      send_email: '1'
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

  // [NOVO] Cotitular 3
  if (d.email_cotitular_3_envio) {
    const signer = {
      email: d.email_cotitular_3_envio,
      name: 'Cotitular 3',
      act: '1',
      foreign: '0',
      send_email: '1'
    };
    if (incluirTelefone && d.telefone_cotitular_3_envio) {
      // Formatar telefone para formato internacional (+55...)
      let phone = d.telefone_cotitular_3_envio.replace(/[^\d+]/g, '');
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
      act: '1',
      foreign: '0',
      send_email: '1'
    });
  }

  const seen = {};
  return list.filter(s => (seen[s.email.toLowerCase()] ? false : (seen[s.email.toLowerCase()] = true)));
}

/* =========================
 * Locks e preflight
 * =======================*/
const locks = new Set();
function acquireLock(key) { if (locks.has(key)) return false; locks.add(key); return true; }
function releaseLock(key) { locks.delete(key); }
async function preflightDNS() { }

/* =========================
 * D4Sign
 * =======================*/
async function makeDocFromWordTemplate(tokenAPI, cryptKey, uuidSafe, templateId, title, varsObj) {
  const base = 'https://secure.d4sign.com.br';
  const url = new URL(`/api/v1/documents/${uuidSafe}/makedocumentbytemplateword`, base);
  url.searchParams.set('tokenAPI', tokenAPI);
  url.searchParams.set('cryptKey', cryptKey);

  const titleSanitized = String(title || 'Contrato').replace(/[\x00-\x1F\x7F-\x9F]/g, '');

  const varsObjValidated = {};
  for (const [key, value] of Object.entries(varsObj || {})) {
    let v = value == null ? '' : String(value);
    v = v.replace(/[\x00-\x1F\x7F-\x9F]/g, '').trim();
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
// CRIAR DOCUMENTO A PARTIR DE TEMPLATE HTML NO D4SIGN
// Endpoint: POST /api/v1/documents/{UUID-SAFE}/makedocumentbytemplate
// Formato correto (doc oficial D4Sign):
//   { "name_document": "...", "templates": { "<ID_TEMPLATE>": { "var": "val" } } }
// ===============================
async function makeDocFromHtmlTemplate(tokenAPI, cryptKey, uuidSafe, templateId, title, fieldsObj) {
  const base = 'https://secure.d4sign.com.br';

  const titleSanitized = String(title || 'Termo de Risco').replace(/[\x00-\x1F\x7F-\x9F]/g, '').trim();

  const fieldsObjValidated = {};
  for (const [key, value] of Object.entries(fieldsObj || {})) {
    let v = value == null ? '' : String(value);
    v = v.replace(/[\x00-\x1F\x7F-\x9F]/g, '').trim();
    fieldsObjValidated[key] = v;
  }

  // Formato correto: templates = { "<id_template>": { var: val, ... } }
  const payload = {
    name_document: titleSanitized,
    templates: {
      [templateId]: fieldsObjValidated
    }
  };

  const url = new URL(`/api/v1/documents/${uuidSafe}/makedocumentbytemplate`, base);
  url.searchParams.set('tokenAPI', tokenAPI);
  url.searchParams.set('cryptKey', cryptKey);

  console.log(`[D4SIGN HTML] Template: ${templateId} | Cofre: ${uuidSafe} | Título: ${titleSanitized}`);
  console.log(`[D4SIGN HTML] Body: ${JSON.stringify(payload).substring(0, 600)}`);

  const res = await fetchWithRetry(url.toString(), {
    method: 'POST',
    headers: { 'Accept': 'application/json', 'Content-Type': 'application/json' },
    body: JSON.stringify(payload)
  }, { attempts: 3, baseDelayMs: 600, timeoutMs: 20000 });

  const text = await res.text();
  console.log(`[D4SIGN HTML] Resposta: status ${res.status} | body: ${text.substring(0, 500)}`);

  let json; try { json = JSON.parse(text); } catch { json = null; }
  const uuid = json && (json.uuid || json.uuid_document || json.document_uuid);

  if (!uuid) {
    throw new Error(`Falha D4Sign(HTML): status ${res.status} - ${text.substring(0, 200)}`);
  }

  console.log(`[D4SIGN HTML] ✓ Documento criado: ${uuid}`);
  return uuid;
}

// Variáveis para o Termo de Risco (CPF e CNPJ)

// ===============================
function montarVarsParaTermoDeRisco(d, nowInfo) {
  const dia = String(nowInfo.dia).padStart(2, '0');
  const mesExtenso = monthNamePt(nowInfo.mes);
  const ano = String(nowInfo.ano);
  const dataFormatada = `${dia} de ${mesExtenso} de ${ano}`;

  // Formata CPF
  const cpfRaw = d.cpf_campo || d.cpf || '';
  const cpfDigits = onlyDigits(cpfRaw);
  const cpfFmt = cpfDigits.length === 11
    ? cpfDigits.replace(/^(\d{3})(\d{3})(\d{3})(\d{2})$/, '$1.$2.$3-$4')
    : cpfRaw;

  // Formata CNPJ
  const cnpjRaw = d.cnpj_campo || d.cnpj || '';
  const cnpjDigits = onlyDigits(cnpjRaw);
  const cnpjFmt = cnpjDigits.length === 14
    ? cnpjDigits.replace(/^(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})$/, '$1.$2.$3/$4-$5')
    : cnpjRaw;

  // Endereço completo montado
  const enderecoPartes = [];
  if (d.rua_cnpj) enderecoPartes.push(`${d.rua_cnpj}`);
  if (d.numero_cnpj) enderecoPartes.push(`nº ${d.numero_cnpj}`);
  if (d.bairro_cnpj) enderecoPartes.push(`Bairro ${d.bairro_cnpj}`);
  if (d.cidade_cnpj && d.uf_cnpj) enderecoPartes.push(`${d.cidade_cnpj} - ${d.uf_cnpj}`);
  else if (d.cidade_cnpj) enderecoPartes.push(d.cidade_cnpj);
  else if (d.uf_cnpj) enderecoPartes.push(d.uf_cnpj);
  if (d.cep_cnpj) enderecoPartes.push(`CEP: ${d.cep_cnpj}`);
  const enderecoCompleto = enderecoPartes.join(', ');

  const isSelecaoCpf = String(d.selecao_cnpj_ou_cpf || '').toUpperCase().trim() === 'CPF';

  if (isSelecaoCpf) {
    // Template CPF — campos conforme o modelo HTML
    return {
      nome_cliente: d.nome || '',
      nacionalidade: d.nacionalidade || '',
      estadocivil: d.estado_civil || '',
      numero_cpf: cpfFmt,
      endereco_completo: enderecoCompleto,
      nome_da_marca: d.titulo || d.nome1 || '',
      cidade: d.cidade_cnpj || '',
      data: dataFormatada,
      NOME: (d.nome || '').toUpperCase(),
      N_CPF: cpfFmt
    };
  } else {
    // Template CNPJ — campos conforme o modelo HTML
    // Sócio administrador (representante da empresa)
    const nomeResponsavel = d.socio_adm_nome || '';
    const cpfResponsavel = d.socio_adm_cpf || '';
    const cpfRespDigits = onlyDigits(cpfResponsavel);
    const cpfRespFmt = cpfRespDigits.length === 11
      ? cpfRespDigits.replace(/^(\d{3})(\d{3})(\d{3})(\d{2})$/, '$1.$2.$3-$4')
      : cpfResponsavel;

    return {
      nomedaempresa: d.nome || '',
      enderecocompleto: enderecoCompleto,
      numero_cnpj: cnpjFmt,
      nome_responsavel_empresa: nomeResponsavel,
      nacionalidade: d.nacionalidade || '',
      estado_civil: d.estado_civil || '',
      numero_cpf: cpfRespFmt,
      endereco_completo: enderecoCompleto,
      nome_da_marca: d.titulo || d.nome1 || '',
      cidade: d.cidade_cnpj || '',
      data: dataFormatada,
      nome_da_empresa: d.nome || '',
      numero_do_cnpj: cnpjFmt
    };
  }
}

// ===============================
// NOVO — REGISTRAR WEBHOOK POR DOCUMENTO D4SIGN
// ===============================
async function registerWebhookForDocument(tokenAPI, cryptKey, uuidDocument, urlWebhook) {
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

          // Garantir formato +55...
          if (!phoneFormatted.startsWith('+')) {
            phoneFormatted = '+55' + phoneFormatted.replace(/^55/, '');
          }

          signer.whatsapp_number = phoneFormatted;
          signer.embed_methodauth = 'whatse'; // Método de entrega via WhatsApp (campo válido D4Sign)
          signer.skipemail = '1'; // Pular envio por email (campo válido D4Sign)
          console.log(`[CADASTRO] Signatário ${s.name} configurado para WhatsApp: ${phoneFormatted}`);
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

  console.log('[CADASTRO] Sucesso D4Sign:', text);
  return text;
}

async function listSigners(tokenAPI, cryptKey, uuidDocument) {
  const base = 'https://secure.d4sign.com.br';
  const url = new URL(`/api/v1/documents/${uuidDocument}/list`, base);
  url.searchParams.set('tokenAPI', tokenAPI);
  url.searchParams.set('cryptKey', cryptKey);

  const res = await fetchWithRetry(url.toString(), { method: 'GET' }, { attempts: 3 });
  if (!res.ok) {
    throw new Error(`Falha ao listar signatários: ${res.status} - ${await res.text()}`);
  }
  const data = await res.json();
  return data; // Retorna array de signatários ou objeto com propriedade list/message
}

async function resendToSigner(tokenAPI, cryptKey, uuidDocument, emailOrWhatsapp, keySigner) {
  const base = 'https://secure.d4sign.com.br';
  const url = new URL(`/api/v1/documents/${uuidDocument}/resend`, base);
  url.searchParams.set('tokenAPI', tokenAPI);
  url.searchParams.set('cryptKey', cryptKey);

  // O campo "email" da API D4Sign aceita "e-mail ou whatsapp"
  // Para signatários com embed_methodauth: 'whatse', deve-se enviar o número WhatsApp
  const body = {
    email: emailOrWhatsapp,
    key_signer: keySigner
  };

  console.log(`[RESEND-API] Reenviando para ${emailOrWhatsapp} (key: ${keySigner}) doc: ${uuidDocument}`);
  console.log(`[RESEND-API] Body:`, JSON.stringify(body));

  const res = await fetchWithRetry(url.toString(), {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body)
  }, { attempts: 1 }); // Sem retry automático para evitar spam/bloqueio

  const text = await res.text();
  console.log(`[RESEND-API] Resposta: ${res.status} - ${text.substring(0, 300)}`);
  if (!res.ok) {
    throw new Error(`Falha ao reenviar: ${res.status} - ${text}`);
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
  url.searchParams.set('tokenAPI', tokenAPI);
  url.searchParams.set('cryptKey', cryptKey);
  const body = { message, skip_email, workflow };

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

    // type_post = "1" → documento finalizado/assinado (TODOS assinaram)
    // type_post = "4" → documento assinado (apenas um assinou) - IGNORAR
    const isSigned = String(type_post) === '1';
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

    // 3. Anexar PDF assinado como ANEXO no Pipefy (não apenas link)
    // [REFATORADO] Removida a salvação de URL direta (que expira), usando apenas upload para campos de anexo
    try {
      const extraFieldId = isProcuracaoFinal ? PIPEFY_FIELD_EXTRA_PROCURACAO : PIPEFY_FIELD_EXTRA_CONTRATO;
      const docType = isProcuracaoFinal ? 'Procuração' : 'Contrato';

      console.log(`[POSTBACK D4SIGN] Iniciando anexação de ${docType}...`);
      console.log(`[POSTBACK D4SIGN] Campo de anexo: ${extraFieldId}`);

      if (!extraFieldId) {
        console.warn(`[POSTBACK D4SIGN] Campo de anexo não configurado para ${docType}`);
      } else {
        const orgId = card.pipe?.organization?.id;

        if (!orgId) {
          console.error(`[POSTBACK D4SIGN] ERRO: Organization ID não encontrado no card. Não é possível fazer upload.`);
        } else {
          console.log(`[POSTBACK D4SIGN] Organization ID: ${orgId}`);

          // Gerar nome do arquivo com o nome da marca
          const nomeMarcaRawPB = byId['marca'] || card.title || 'Documento';
          const nomeMarcaPB = String(nomeMarcaRawPB).replace(/[<>:"/\\|?*]/g, '_').trim();
          const fileName = isProcuracaoFinal
            ? `${nomeMarcaPB} - Procuração.pdf`
            : `${nomeMarcaPB} - Contrato Assinado.pdf`;

          console.log(`[POSTBACK D4SIGN] Fazendo upload do arquivo: ${fileName}`);
          console.log(`[POSTBACK D4SIGN] URL origem D4Sign: ${info.url}`);

          // Fazer upload para o Pipefy
          const pipefyUrl = await uploadFileToPipefy(info.url, fileName, orgId);
          console.log(`[POSTBACK D4SIGN] Upload concluído. URL Pipefy: ${pipefyUrl}`);

          // Atualizar campo de anexo - Pipefy espera array com caminho relativo do S3
          console.log(`[POSTBACK D4SIGN] Caminho do anexo para Pipefy: ${pipefyUrl}`);
          await updateCardField(cardId, extraFieldId, [pipefyUrl]);
          console.log(`[POSTBACK D4SIGN] ✓ ${docType} anexado com sucesso no campo ${extraFieldId}`);

          // Salvar também na pasta de rede da empresa
          const equipeContrato = getEquipeContratoFromCard(card);
          await saveFileLocally(info.url, fileName, equipeContrato || 'Sem_Equipe');

          // Nota: não salvamos no campo de texto porque retornamos apenas o caminho relativo do S3
          // O arquivo está acessível pelo campo de anexo diretamente
        }
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

    // Ler campo filiais_ou_digital do card
    const by = toById(card);
    const filiaisOuDigital = by['filiais_ou_digital'] || '';
    const tipoUnidade = filiaisOuDigital || 'Não informado';

    const html = `
<!doctype html><html lang="pt-BR"><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Província Marcas — Revisar Contrato</title>
<link href="https://fonts.googleapis.com/css2?family=Oswald:wght@400;600;700&family=Inter:wght@400;500;600&display=swap" rel="stylesheet">
<style>
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:'Inter',system-ui,sans-serif;background:#f4f4f4;color:#111;min-height:100vh;display:flex;flex-direction:column}
  .header{background:#000;padding:14px 24px;display:flex;align-items:center;justify-content:center}
  .header img{height:44px;object-fit:contain}
  .wrap{max-width:940px;margin:28px auto;padding:0 16px;flex:1}
  .card{background:#fff;border-radius:14px;box-shadow:0 4px 20px rgba(0,0,0,.08);overflow:hidden;margin-bottom:20px}
  .card-header{background:#000;padding:16px 24px}
  .card-header h1,.card-header h2{font-family:'Oswald',sans-serif;color:#FFE200;text-transform:uppercase;letter-spacing:.5px}
  .card-header h1{font-size:20px;font-weight:700}
  .card-header h2{font-size:15px;font-weight:600;margin-top:0}
  .card-body{padding:20px 24px}
  h2{font-family:'Oswald',sans-serif;font-size:14px;font-weight:600;text-transform:uppercase;letter-spacing:.5px;color:#FFE200;background:#000;padding:8px 12px;border-radius:6px;margin:20px 0 10px}
  .grid{display:grid;grid-template-columns:1fr 1fr;gap:12px}
  .grid3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px}
  .field{background:#f8f8f8;border-left:3px solid #FFE200;padding:10px 12px;border-radius:6px}
  .label{font-weight:600;font-size:12px;text-transform:uppercase;letter-spacing:.5px;color:#555;margin-bottom:3px}
  .value{font-size:14px;color:#111;line-height:1.5}
  .btn{display:inline-block;padding:13px 24px;border-radius:8px;text-decoration:none;border:0;background:#FFE200;color:#000;font-family:'Oswald',sans-serif;font-size:15px;font-weight:700;letter-spacing:.5px;cursor:pointer;transition:filter .15s}
  .btn:hover{filter:brightness(.9)}
  .muted{color:#888;font-size:13px;margin-top:10px}
  .tag{display:inline-block;background:#FFE200;color:#000;border-radius:6px;padding:3px 10px;font-size:12px;font-family:'Oswald',sans-serif;font-weight:700;margin-left:8px;vertical-align:middle}
  .footer{background:#000;padding:12px 24px;text-align:center;font-size:12px;color:#888;font-family:'Oswald',sans-serif;letter-spacing:.5px}
  .footer span{color:#FFE200}
</style>
<div class="header">
  <img src="https://provinciamarcas.com.br/wp-content/uploads/2025/12/provincia-marcas-e-patentes-pensou-criou-registrou-logo.png" alt="Província Marcas" onerror="this.style.display='none'">
</div>
<div class="wrap">
  <div class="card">
    <div class="card-header">
      <h1>Revisar dados do contrato <span class="tag">Card #${card.id}</span></h1>
    </div>
    <div class="card-body">

    <h2>Informações Gerais</h2>
    <div class="grid">
      <div class="field"><div class="label">N° de contrato</div><div class="value">${card.id}</div></div>
      <div class="field"><div class="label">Tipo de Unidade</div><div class="value">${tipoUnidade}</div></div>
    </div>

    <h2>Contratante(s)</h2>
    <div class="grid">
      <div class="field"><div class="label">Contratante 1</div><div class="value">${d.contratante_1_texto || '-'}</div></div>
      <div class="field"><div class="label">Contratante 2</div><div class="value">${d.contratante_2_texto || '-'}</div></div>
      <div class="field"><div class="label">Contratante 3</div><div class="value">${d.contratante_3_texto || '-'}</div></div>
    </div>

    <h2>Contato</h2>
    <div class="grid">
      <div><div class="label">Dados para contato 1</div><div>${d.dados_contato_1 || '-'}</div></div>
      <div><div class="label">Dados para contato 2</div><div>${d.dados_contato_2 || '-'}</div></div>
      <div><div class="label">Email Envio (Titular)</div><div>${d.email_envio_contrato || '-'}</div></div>
      <div><div class="label">Telefone Envio (Titular)</div><div>${d.telefone_envio_contrato || '-'}</div></div>
      <div><div class="label">Email Envio (Cotitular)</div><div>${d.email_cotitular_envio || '-'}</div></div>
      <div><div class="label">Telefone Envio (Cotitular)</div><div>${d.telefone_cotitular_envio || '-'}</div></div>
      <div><div class="label">Email Envio (Cotitular 3)</div><div>${d.email_cotitular_3_envio || '-'}</div></div>
      <div><div class="label">Telefone Envio (Cotitular 3)</div><div>${d.telefone_cotitular_3_envio || '-'}</div></div>
    </div>

    <h2>Serviços</h2>
    <div class="grid3">
      <div class="field"><div class="label">Template escolhido</div><div class="value">${
        d.templateToUse === process.env.TEMPLATE_UUID_CONTRATO ? '📋 Marca (sem risco)' :
        d.templateToUse === process.env.TEMPLATE_UUID_CONTRATO_MARCA_RISCO ? '⚠️ Marca + Termo de Risco' :
        '📑 Outros Serviços'
      }</div></div>
      <div class="field"><div class="label">Risco da Marca</div><div class="value" style="font-weight:600;color:${d.risco && ['médio','alto'].includes(d.risco.toLowerCase()) ? '#d32f2f' : '#137333'}">${d.risco || '-'}</div></div>
      <div class="field"><div class="label">Filial</div><div class="value">${d.equipeName || '-'}</div></div>
    </div>

    <h2>Pagamentos</h2>
    <div class="grid">
      <div class="field"><div class="label">Assessoria</div><div class="value">${d.valor_total || '-'} (${d.parcelas || '1'}x via ${d.forma_pagto_assessoria || '-'})</div></div>
      <div class="field"><div class="label">Pesquisa</div><div class="value">${d.valor_pesquisa || '-'}</div></div>
      <div class="field"><div class="label">Taxa</div><div class="value">${d.valor_taxa_brl || '-'} via ${d.forma_pagto_taxa || '-'}</div></div>
      <div class="field"><div class="label">Texto Assessoria no Contrato</div><div class="value" style="font-size:13px">${d.forma_pagto_assessoria ? (d.parcelas || '1') + ' parcela(s) de ... via ' + d.forma_pagto_assessoria + (d.data_pagto_assessoria ? ' · ' + d.data_pagto_assessoria : '') : '-'}</div></div>
    </div>

    <h2>Classes e Especificações</h2>
    ${[1, 2, 3, 4, 5].map(i => {
      const nome = d[`nome${i}`];
      const tipo = d[`tipo${i}`];
      const classe = d[`classe${i}`];
      const linhas = (d[`classes_agrupadas_${i}`] || []).filter(l => l && l.trim());
      if (i > 1 && !nome) return '';
      if (!nome && !classe && !linhas.length) return '';
      return `
      <div style="margin-bottom:14px;background:#f8f8f8;border-left:3px solid #FFE200;border-radius:6px;padding:12px 14px">
        <div style="font-family:'Oswald',sans-serif;font-weight:700;font-size:14px;color:#000;margin-bottom:6px">
          MARCA ${i}${nome ? ` — ${nome}` : ''}
          ${tipo ? `<span style="font-size:12px;font-weight:400;color:#555;margin-left:8px">(${tipo})</span>` : ''}
          ${classe ? `<span style="font-size:12px;font-weight:400;color:#555;margin-left:8px">Classe: ${classe}</span>` : ''}
        </div>
        ${linhas.length ? linhas.map(l => `<div style="font-size:13px;color:#333;line-height:1.6;padding:2px 0;border-top:1px solid #eee">${l}</div>`).join('') : '<div style="font-size:13px;color:#888">Sem especificações</div>'}
      </div>`;
    }).join('')}


    </div>
    <div style="margin-top:24px;padding-top:20px;border-top:2px solid #FFE200">
      <form method="POST" action="/lead/${encodeURIComponent(req.params.token)}/generate">
        <button class="btn" type="submit">📄 Gerar Contrato e Procuração</button>
      </form>
      <p class="muted">Ao clicar, os documentos serão criados no D4Sign.</p>
    </div>
    </div>
  </div>
</div>
<div class="footer">© Província Marcas e Patentes — <span>Pensou. Criou. Província Registrou!</span></div>
`;
    res.setHeader('content-type', 'text/html; charset=utf-8');
    return res.status(200).send(html);
  } catch (e) {
    console.error('[ERRO /lead]', e.message || e);
    return res.status(400).send('Link inválido ou expirado. Erro: ' + (e.message || String(e)));
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

    preflightDNS().catch(() => { });

    const card = await getCard(cardId);
    if (!card) {
      throw new Error(`Card ${cardId} não encontrado no Pipefy`);
    }

    const d = await montarDados(card);
    if (!d) {
      throw new Error('Falha ao montar dados do card');
    }

    const now = new Date();
    const nowInfo = { dia: now.getDate(), mes: now.getMonth() + 1, ano: now.getFullYear() };

    // Validar template
    if (!d.templateToUse) {
      throw new Error('Template não identificado. Verifique os dados do card.');
    }

    const isMarcaTemplate = d.templateToUse === TEMPLATE_UUID_CONTRATO || d.templateToUse === TEMPLATE_UUID_CONTRATO_MARCA_RISCO;
    const add = isMarcaTemplate ? montarVarsParaTemplateMarca(d, nowInfo)
      : montarVarsParaTemplateOutros(d, nowInfo);

    if (!add || typeof add !== 'object') {
      throw new Error('Falha ao montar variáveis do template. Verifique os dados do card.');
    }

    // Log para verificar se o número do contrato está sendo passado
    console.log(`[LEAD-GENERATE] ========== DEBUG NÚMERO DO CONTRATO ==========`);
    console.log(`[LEAD-GENERATE] card.id: ${card.id}, tipo: ${typeof card.id}`);
    console.log(`[LEAD-GENERATE] d.cardId: ${d.cardId}, tipo: ${typeof d.cardId}`);
    console.log(`[LEAD-GENERATE] isMarcaTemplate: ${isMarcaTemplate}`);
    console.log(`[LEAD-GENERATE] Campo "Número do contrato do bloco físico*": ${add['Número do contrato do bloco físico*'] || 'NÃO ENCONTRADO'}`);
    console.log(`[LEAD-GENERATE] Número do contrato no template (primeiras variações):`, {
      'N° contrato': add['N° contrato'],
      'Nº contrato': add['Nº contrato'],
      'CONTRATO nº': add['CONTRATO nº'],
      'CONTRATO nº:': add['CONTRATO nº:'],
      'N° de contrato': add['N° de contrato'],
      'Contrato N°': add['Contrato N°']
    });
    console.log(`[LEAD-GENERATE] Total de chaves no objeto add: ${Object.keys(add).length}`);
    console.log(`[LEAD-GENERATE] =============================================`);

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
      // [NOVO] Salvar UUID no card
      await updateCardField(cardId, PIPEFY_FIELD_D4_UUID_CONTRATO, uuidDoc);
      console.log(`[LEAD-GENERATE] UUID Contrato salvo no card: ${uuidDoc}`);
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
        await new Promise(r => setTimeout(r, 3000));
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
          // [NOVO] Salvar UUID no card
          await updateCardField(cardId, PIPEFY_FIELD_D4_UUID_PROCURACAO, uuidProcuracao);
          console.log(`[LEAD-GENERATE] UUID Procuração salvo no card: ${uuidProcuracao}`);
        } catch (e) {
          console.error('[ERRO] Falha ao registrar webhook da procuração:', e.message);
        }

        // Signatários serão cadastrados abaixo junto com o contrato
        console.log('[D4SIGN] Procuração criada. Signatários serão cadastrados em breve.');
      } catch (e) {
        console.error('[ERRO] Falha ao gerar procuração:', e.message);
        // Não bloqueia o fluxo se a procuração falhar
      }
    }

    // ===============================
    // NOVO — GERAR TERMO DE RISCO (HTML Template)
    // ===============================
    let uuidTermoDeRisco = null;
    const templateIdTermo = String(d.selecao_cnpj_ou_cpf || '').toUpperCase().trim() === 'CPF'

    if (templateIdTermo) {
      try {
        const varsTermo = montarVarsParaTermoDeRisco(d, nowInfo);
        console.log('[D4SIGN] Gerando Termo de Risco. Template:', templateIdTermo, '| Tipo:', d.selecao_cnpj_ou_cpf);
        uuidTermoDeRisco = await makeDocFromHtmlTemplate(
          D4SIGN_TOKEN,
          D4SIGN_CRYPT_KEY,
          uuidSafe,
          templateIdTermo,
          `Termo de Risco - ${d.titulo || card.title || 'Contrato'}`,
          varsTermo
        );
        console.log(`[D4SIGN] Termo de Risco criado: ${uuidTermoDeRisco}`);

        // Aguardar documento estar pronto
        await new Promise(r => setTimeout(r, 3000));
        try {
          await getDocumentStatus(D4SIGN_TOKEN, D4SIGN_CRYPT_KEY, uuidTermoDeRisco);
          console.log('[D4SIGN] Status do Termo de Risco verificado.');
        } catch (e) {
          console.warn('[D4SIGN] Aviso ao verificar status do Termo de Risco:', e.message);
        }

        // Registrar webhook do Termo de Risco
        try {
          await registerWebhookForDocument(
            D4SIGN_TOKEN,
            D4SIGN_CRYPT_KEY,
            uuidTermoDeRisco,
            `${PUBLIC_BASE_URL}/d4sign/postback`
          );
          console.log('[D4SIGN] Webhook do Termo de Risco registrado.');
          // Salvar UUID no card (se campo configurado)
          if (PIPEFY_FIELD_D4_UUID_TERMO_DE_RISCO) {
            await updateCardField(cardId, PIPEFY_FIELD_D4_UUID_TERMO_DE_RISCO, uuidTermoDeRisco);
            console.log(`[LEAD-GENERATE] UUID Termo de Risco salvo no card: ${uuidTermoDeRisco}`);
          }
        } catch (e) {
          console.error('[ERRO] Falha ao registrar webhook do Termo de Risco:', e.message);
        }

        console.log('[D4SIGN] Termo de Risco criado. Signatários serão cadastrados em breve.');
      } catch (e) {
        console.error('[ERRO] Falha ao gerar Termo de Risco:', e.message);
        // Não bloqueia o fluxo se o Termo de Risco falhar
      }
    }

    await new Promise(r => setTimeout(r, 3000));

    try { await getDocumentStatus(D4SIGN_TOKEN, D4SIGN_CRYPT_KEY, uuidDoc); } catch { }

    // Signatários serão cadastrados apenas quando o documento for enviado para assinatura
    // Isso evita duplicação de signatários e permite envio manual
    console.log('[D4SIGN] Contrato criado. Aguardando envio manual...');

    /*
    try {
      if (signers && signers.length > 0) {
        console.log(`[LEAD-GENERATE] Cadastrando ${signers.length} signatários no contrato ${uuidDoc}...`);
        await cadastrarSignatarios(D4SIGN_TOKEN, D4SIGN_CRYPT_KEY, uuidDoc, signers);

        if (uuidProcuracao) {
          console.log(`[LEAD-GENERATE] Cadastrando ${signers.length} signatários na procuração ${uuidProcuracao}...`);
          await cadastrarSignatarios(D4SIGN_TOKEN, D4SIGN_CRYPT_KEY, uuidProcuracao, signers);
        }
      }
    } catch (e) {
      console.error('[LEAD-GENERATE] Erro ao cadastrar signatários:', e.message);
      // Não falha o processo todo, mas loga o erro
    }
    */

    await new Promise(r => setTimeout(r, 2000));
    // Movimentação de card removida conforme solicitado

    releaseLock(lockKey);

    const token = req.params.token;
    const html = `
<!doctype html><meta charset="utf-8"><title>Província Marcas — Contrato</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Oswald:wght@400;600;700&family=Inter:wght@400;500;600&display=swap" rel="stylesheet">
<style>
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:'Inter',system-ui,sans-serif;background:#f4f4f4;color:#111;min-height:100vh;display:flex;flex-direction:column}
  .header{background:#000;padding:16px 24px;display:flex;align-items:center;justify-content:center}
  .header img{height:48px;object-fit:contain}
  .main{flex:1;display:grid;place-items:start center;padding:32px 16px}
  .box{background:#fff;border-radius:16px;box-shadow:0 4px 24px rgba(0,0,0,.10);max-width:660px;width:100%;overflow:hidden}
  .box-header{background:#000;padding:20px 24px;display:flex;align-items:center;gap:12px}
  .box-header h2{font-family:'Oswald',sans-serif;font-size:20px;font-weight:700;color:#FFE200;letter-spacing:.5px;text-transform:uppercase}
  .box-body{padding:24px}
  h3{font-family:'Oswald',sans-serif;font-size:16px;font-weight:600;text-transform:uppercase;letter-spacing:.5px;color:#111;margin:0 0 8px}
  .row{display:flex;gap:10px;flex-wrap:wrap;margin-top:12px}
  .btn{display:inline-block;padding:11px 16px;border-radius:8px;text-decoration:none;border:0;font-family:'Oswald',sans-serif;font-size:14px;font-weight:600;letter-spacing:.5px;cursor:pointer;transition:filter .15s}
  .btn:hover{filter:brightness(.88)}
  .btn-default{background:#000;color:#FFE200}
  .btn-download{background:#FFE200;color:#000}
  .btn-whatsapp{background:#25D366;color:#fff}
  .btn-email{background:#1976d2;color:#fff}
  .btn-reenviar{background:#555;color:#fff;display:none;margin-left:4px}
  .section{margin-top:24px;padding-top:24px;border-top:2px solid #FFE200}
  .info-box{background:#f8f8f8;border-radius:8px;padding:14px 16px;margin-bottom:4px;font-size:14px;line-height:1.7;border-left:3px solid #FFE200}
  .info-box strong{color:#000}
  .status-div{margin-top:8px;min-height:24px;font-size:13px}
  .alert-yellow{margin:12px 0;padding:12px 14px;background:#fff9c4;border-left:4px solid #FFE200;border-radius:6px;color:#5a4a00;font-size:13px;line-height:1.5}
  .footer{background:#000;padding:12px 24px;text-align:center;font-size:12px;color:#888;font-family:'Oswald',sans-serif;letter-spacing:.5px}
  .footer span{color:#FFE200}
</style>
<div class="header">
  <img src="https://provinciamarcas.com.br/wp-content/uploads/2025/12/provincia-marcas-e-patentes-pensou-criou-registrou-logo.png" alt="Província Marcas" onerror="this.style.display='none'">
</div>
<div class="main">
<div class="box">
  <div class="box-header">
    <h2>📄 ${uuidProcuracao ? 'Contrato e Procuração Gerados' : 'Contrato Gerado com Sucesso'}</h2>
  </div>
  <div class="box-body">

  ${cofreUsadoPadrao ? `
  <div style="background:#fff3cd;border-left:4px solid #ffc107;padding:12px;margin:16px 0;border-radius:4px">
    <strong>⚠️ Atenção:</strong> A equipe "${equipeContrato || 'não informada'}" não possui cofre configurado. 
    Documentos salvos no cofre padrão: <strong>${nomeCofreUsado}</strong>
  </div>
  ` : ''}
  <div class="info-box">
    <div><strong>Email do Titular:</strong> ${d.email_envio_contrato || d.email || 'Não informado'}</div>
    ${d.email_cotitular_envio ? `<div><strong>Email do Cotitular:</strong> ${d.email_cotitular_envio}</div>` : ''}
  </div>
  <div class="alert-yellow">
    <strong>⚠️ Atenção:</strong> Em caso de envio por WhatsApp + Email, necessário remover no D4Sign um dos signatários para que o contrato e procuração fiquem com status finalizado; Baixar os arquivos e anexar documentos no card.
  </div>
  <div class="row">
    <a class="btn btn-download" href="/lead/${encodeURIComponent(token)}/doc/${encodeURIComponent(uuidDoc)}/download" target="_blank" rel="noopener">⬇ Baixar PDF do Contrato</a>
    <button class="btn btn-email" onclick="enviarContrato('${token}', '${uuidDoc}', 'email')" id="btn-enviar-contrato-email">✉ Enviar por Email</button>
    <button class="btn btn-whatsapp" onclick="enviarContrato('${token}', '${uuidDoc}', 'whatsapp')" id="btn-enviar-contrato-whatsapp">💬 Enviar por WhatsApp</button>
    <button class="btn btn-reenviar" onclick="reenviarContrato('${token}', '${uuidDoc}')" id="btn-reenviar-contrato" disabled>Reenviar Link (60s)</button>
  </div>
  <div id="status-contrato" class="status-div"></div>
  ${uuidProcuracao ? `
  <div class="section">
    <h3>📋 Procuração</h3>
    <div class="row">
      <a class="btn btn-download" href="/lead/${encodeURIComponent(token)}/doc/${encodeURIComponent(uuidProcuracao)}/download" target="_blank" rel="noopener">⬇ Baixar PDF da Procuração</a>
      <button class="btn btn-email" onclick="enviarProcuracao('${token}', '${uuidProcuracao}', 'email')" id="btn-enviar-procuracao-email">✉ Enviar por Email</button>
      <button class="btn btn-whatsapp" onclick="enviarProcuracao('${token}', '${uuidProcuracao}', 'whatsapp')" id="btn-enviar-procuracao-whatsapp">💬 Enviar por WhatsApp</button>
      <button class="btn btn-reenviar" onclick="reenviarProcuracao('${token}', '${uuidProcuracao}')" id="btn-reenviar-procuracao" disabled>Reenviar Link (60s)</button>
    </div>
    <div id="status-procuracao" class="status-div"></div>
  </div>
  ` : ''}
  ${uuidTermoDeRisco ? `
  <div class="section">
    <div style="background:#000;color:#FFE200;font-family:'Oswald',sans-serif;font-weight:700;font-size:14px;letter-spacing:.8px;text-transform:uppercase;padding:10px 14px;border-radius:6px;margin-bottom:12px">
      🚨 ATENÇÃO — Enviar o termo de risco para classe impedida (sem complementação)
    </div>
    <h3>⚠️ Termo de Risco</h3>
    <div class="row">
      <a class="btn btn-download" href="/lead/${encodeURIComponent(token)}/doc/${encodeURIComponent(uuidTermoDeRisco)}/download" target="_blank" rel="noopener">⬇ Baixar PDF do Termo de Risco</a>
      <button class="btn btn-email" onclick="enviarTermoDeRisco('${token}', '${uuidTermoDeRisco}', 'email')" id="btn-enviar-termo-email">✉ Enviar por Email</button>
      <button class="btn btn-whatsapp" onclick="enviarTermoDeRisco('${token}', '${uuidTermoDeRisco}', 'whatsapp')" id="btn-enviar-termo-whatsapp">💬 Enviar por WhatsApp</button>
      <button class="btn btn-reenviar" onclick="reenviarTermoDeRisco('${token}', '${uuidTermoDeRisco}')" id="btn-reenviar-termo" disabled>Reenviar Link (60s)</button>
    </div>
    <div id="status-termo" class="status-div"></div>
  </div>
  ` : ''}
  <div class="row" style="margin-top:24px">
    <a class="btn btn-default" href="${PUBLIC_BASE_URL}/lead/${encodeURIComponent(token)}">← Voltar</a>
  </div>
  </div>
</div>
</div>
<div class="footer">© Província Marcas e Patentes — <span>Pensou. Criou. Província Registrou!</span></div>
<script>
async function enviarContrato(token, uuidDoc, canal) {
  const btnEmail = document.getElementById('btn-enviar-contrato-email');
  const btnWhatsapp = document.getElementById('btn-enviar-contrato-whatsapp');
  const statusDiv = document.getElementById('status-contrato');
  
  const btn = canal === 'whatsapp' ? btnWhatsapp : btnEmail;
  
  btn.disabled = true;
  btn.textContent = 'Enviando...';
  statusDiv.innerHTML = '<span style="color:#1976d2">⏳ Enviando contrato por ' + canal + '...</span>';
  
  try {
    const response = await fetch('/lead/' + encodeURIComponent(token) + '/doc/' + encodeURIComponent(uuidDoc) + '/send?canal=' + canal, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' }
    });
    
    const data = await response.json();
    
    if (response.ok && data.success) {
      const cofreMsg = data.cofre ? ' Salvo no cofre: ' + data.cofre : '';
      const urlCofreMsg = data.urlCofre ? '<br><br><div style="margin-top:12px;padding:12px;background:#f5f5f5;border-radius:8px;border-left:4px solid #1976d2;"><strong style="color:#1976d2">Link do D4 para reenviar ou alterar os signatarios:</strong><br><a href="' + data.urlCofre + '" target="_blank" style="color:#1976d2;text-decoration:underline;word-break:break-all">' + data.urlCofre + '</a></div>' : '';
      
      let destinoMsg = '';
      if (canal === 'whatsapp' && data.telefones) {
        destinoMsg = ' para: ' + data.telefones;
      } else if (data.emails) {
        destinoMsg = ' para: ' + data.emails;
      } else if (data.email) {
        destinoMsg = ' para ' + data.email;
      }
      
      const avisoMsg = '<br><br><div style="margin-top:12px;padding:12px;background:#fff3cd;border-radius:8px;border-left:4px solid #ffc107;color:#856404;font-size:14px;"><strong>⚠️ Importante:</strong> Caso o email ou whatsapp não cheguem para assinatura, é necessário abrir o D4Sign (link acima) e clicar em "Enviar novamente".</div>';

      statusDiv.innerHTML = '<span style="color:#28a745;font-weight:600">✓ Status de envio - Contrato: Enviado com sucesso' + destinoMsg + '.' + cofreMsg + '</span>' + urlCofreMsg + avisoMsg;
      btn.textContent = 'Enviado por ' + (canal === 'whatsapp' ? 'WhatsApp' : 'Email');
      btn.style.background = '#6c757d'; // Cinza
      btn.disabled = true;

      // Ativar botão de reenvio com timer
      const btnReenviar = document.getElementById('btn-reenviar-contrato');
      if (btnReenviar) {
        btnReenviar.style.display = 'inline-block';
        let timeLeft = 60;
        btnReenviar.textContent = 'Reenviar Link (' + timeLeft + 's)';
        btnReenviar.disabled = true;
        
        const timerId = setInterval(() => {
          timeLeft--;
          if (timeLeft <= 0) {
            clearInterval(timerId);
            btnReenviar.textContent = 'Reenviar Link';
            btnReenviar.disabled = false;
            btnReenviar.style.background = '#111'; // Cor padrão (preto) ou outra cor de destaque
          } else {
            btnReenviar.textContent = 'Reenviar Link (' + timeLeft + 's)';
          }
        }, 1000);
      }
    } else {
      const errorMsg = data.message || data.detalhes || 'Erro ao enviar';
      statusDiv.innerHTML = '<span style="color:#d32f2f;font-weight:600">✗ Status de envio - Contrato: ' + errorMsg + '</span>';
      btn.disabled = false;
      btn.textContent = 'Enviar por ' + (canal === 'whatsapp' ? 'WhatsApp' : 'Email');
    }
  } catch (error) {
    statusDiv.innerHTML = '<span style="color:#d32f2f">✗ Status de envio - Contrato: Erro ao enviar - ' + error.message + '</span>';
    btn.disabled = false;
    btn.textContent = 'Enviar por ' + (canal === 'whatsapp' ? 'WhatsApp' : 'Email');
  }
}

async function enviarProcuracao(token, uuidProcuracao, canal) {
  const btnEmail = document.getElementById('btn-enviar-procuracao-email');
  const btnWhatsapp = document.getElementById('btn-enviar-procuracao-whatsapp');
  const statusDiv = document.getElementById('status-procuracao');
  
  const btn = canal === 'whatsapp' ? btnWhatsapp : btnEmail;
  
  btn.disabled = true;
  btn.textContent = 'Enviando...';
  statusDiv.innerHTML = '<span style="color:#1976d2">⏳ Enviando procuração por ' + canal + '...</span>';
  
  try {
    const response = await fetch('/lead/' + encodeURIComponent(token) + '/doc/' + encodeURIComponent(uuidProcuracao) + '/send?canal=' + canal + '&tipo=procuracao', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' }
    });
    
    const data = await response.json();
    
    if (response.ok && data.success) {
      const cofreMsg = data.cofre ? ' Salvo no cofre: ' + data.cofre : '';
      const urlCofreMsg = data.urlCofre ? '<br><br><div style="margin-top:12px;padding:12px;background:#f5f5f5;border-radius:8px;border-left:4px solid #1976d2;"><strong style="color:#1976d2">Link do D4 para reenviar ou alterar os signatarios:</strong><br><a href="' + data.urlCofre + '" target="_blank" style="color:#1976d2;text-decoration:underline;word-break:break-all">' + data.urlCofre + '</a></div>' : '';
      
      let destinoMsg = '';
      if (canal === 'whatsapp' && data.telefones) {
        destinoMsg = ' para: ' + data.telefones;
      } else if (data.emails) {
        destinoMsg = ' para: ' + data.emails;
      } else if (data.email) {
        destinoMsg = ' para ' + data.email;
      }

      const avisoMsg = '<br><br><div style="margin-top:12px;padding:12px;background:#fff3cd;border-radius:8px;border-left:4px solid #ffc107;color:#856404;font-size:14px;"><strong>⚠️ Importante:</strong> Caso o email ou whatsapp não cheguem para assinatura, é necessário abrir o D4Sign (link acima) e clicar em "Enviar novamente".</div>';

      statusDiv.innerHTML = '<span style="color:#28a745;font-weight:600">✓ Status de envio - Procuração: Enviado com sucesso' + destinoMsg + '.' + cofreMsg + '</span>' + urlCofreMsg + avisoMsg;
      btn.textContent = 'Enviado por ' + (canal === 'whatsapp' ? 'WhatsApp' : 'Email');
      btn.style.background = '#6c757d'; // Cinza
      btn.disabled = true;

      // Ativar botão de reenvio com timer
      const btnReenviar = document.getElementById('btn-reenviar-procuracao');
      if (btnReenviar) {
        btnReenviar.style.display = 'inline-block';
        let timeLeft = 60;
        btnReenviar.textContent = 'Reenviar Link (' + timeLeft + 's)';
        btnReenviar.disabled = true;
        
        const timerId = setInterval(() => {
          timeLeft--;
          if (timeLeft <= 0) {
            clearInterval(timerId);
            btnReenviar.textContent = 'Reenviar Link';
            btnReenviar.disabled = false;
            btnReenviar.style.background = '#111';
          } else {
            btnReenviar.textContent = 'Reenviar Link (' + timeLeft + 's)';
          }
        }, 1000);
      }
    } else {
      const errorMsg = data.message || data.detalhes || 'Erro ao enviar';
      statusDiv.innerHTML = '<span style="color:#d32f2f;font-weight:600">✗ Status de envio - Procuração: ' + errorMsg + '</span>';
      btn.disabled = false;
      btn.textContent = 'Enviar por ' + (canal === 'whatsapp' ? 'WhatsApp' : 'Email');
    }
  } catch (error) {
    statusDiv.innerHTML = '<span style="color:#d32f2f">✗ Status de envio - Procuração: Erro ao enviar - ' + error.message + '</span>';
    btn.disabled = false;
    btn.textContent = 'Enviar por ' + (canal === 'whatsapp' ? 'WhatsApp' : 'Email');
  }
}

async function reenviarContrato(token, uuidDoc) {
  const btn = document.getElementById('btn-reenviar-contrato');
  const statusDiv = document.getElementById('status-contrato');
  
  btn.disabled = true;
  btn.textContent = 'Reenviando...';
  
  try {
    const response = await fetch('/lead/' + encodeURIComponent(token) + '/doc/' + encodeURIComponent(uuidDoc) + '/resend', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' }
    });
    
    const data = await response.json();
    
    if (response.ok && data.success) {
      statusDiv.innerHTML += '<br><span style="color:#28a745;font-weight:600">✓ Reenvio solicitado com sucesso.</span>';
      
      // Reiniciar timer de 60s
      let timeLeft = 60;
      btn.textContent = 'Reenviar Link (' + timeLeft + 's)';
      btn.style.background = '#6c757d';
      
      const timerId = setInterval(() => {
        timeLeft--;
        if (timeLeft <= 0) {
          clearInterval(timerId);
          btn.textContent = 'Reenviar Link';
          btn.disabled = false;
          btn.style.background = '#111';
        } else {
          btn.textContent = 'Reenviar Link (' + timeLeft + 's)';
        }
      }, 1000);
      
    } else {
      statusDiv.innerHTML += '<br><span style="color:#d32f2f">✗ Erro ao reenviar: ' + (data.message || 'Erro desconhecido') + '</span>';
      btn.textContent = 'Reenviar Link';
      btn.disabled = false;
    }
  } catch (error) {
    statusDiv.innerHTML += '<br><span style="color:#d32f2f">✗ Erro ao reenviar: ' + error.message + '</span>';
    btn.textContent = 'Reenviar Link';
    btn.disabled = false;
  }
}

async function reenviarProcuracao(token, uuidDoc) {
  const btn = document.getElementById('btn-reenviar-procuracao');
  const statusDiv = document.getElementById('status-procuracao');
  
  btn.disabled = true;
  btn.textContent = 'Reenviando...';
  
  try {
    const response = await fetch('/lead/' + encodeURIComponent(token) + '/doc/' + encodeURIComponent(uuidDoc) + '/resend', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' }
    });
    
    const data = await response.json();
    
    if (response.ok && data.success) {
      statusDiv.innerHTML += '<br><span style="color:#28a745;font-weight:600">✓ Reenvio solicitado com sucesso.</span>';
      
      // Reiniciar timer de 60s
      let timeLeft = 60;
      btn.textContent = 'Reenviar Link (' + timeLeft + 's)';
      btn.style.background = '#6c757d';
      
      const timerId = setInterval(() => {
        timeLeft--;
        if (timeLeft <= 0) {
          clearInterval(timerId);
          btn.textContent = 'Reenviar Link';
          btn.disabled = false;
          btn.style.background = '#111';
        } else {
          btn.textContent = 'Reenviar Link (' + timeLeft + 's)';
        }
      }, 1000);
      
    } else {
      statusDiv.innerHTML += '<br><span style="color:#d32f2f">✗ Erro ao reenviar: ' + (data.message || 'Erro desconhecido') + '</span>';
      btn.textContent = 'Reenviar Link';
      btn.disabled = false;
    }
  } catch (error) {
    statusDiv.innerHTML += '<br><span style="color:#d32f2f">✗ Erro ao reenviar: ' + error.message + '</span>';
    btn.textContent = 'Reenviar Link';
    btn.disabled = false;
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
<!doctype html><meta charset="utf-8"><title>Província Marcas — Erro</title>
<link href="https://fonts.googleapis.com/css2?family=Oswald:wght@600;700&family=Inter:wght@400;500&display=swap" rel="stylesheet">
<style>
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:'Inter',system-ui,sans-serif;background:#f4f4f4;min-height:100vh;display:flex;flex-direction:column}
  .header{background:#000;padding:16px 24px;display:flex;align-items:center;justify-content:center}
  .header img{height:48px;object-fit:contain}
  .main{flex:1;display:grid;place-items:center;padding:32px 16px}
  .box{background:#fff;border-radius:16px;box-shadow:0 4px 24px rgba(0,0,0,.1);max-width:600px;width:100%;overflow:hidden}
  .box-header{background:#000;padding:20px 24px}
  .box-header h2{font-family:'Oswald',sans-serif;font-size:20px;font-weight:700;color:#FFE200;text-transform:uppercase;letter-spacing:.5px}
  .box-body{padding:24px}
  .error-box{background:#ffebee;border-left:4px solid #d32f2f;padding:16px;border-radius:8px;margin:16px 0}
  .error-box strong{display:block;margin-bottom:8px;color:#c62828;font-family:'Oswald',sans-serif;font-size:15px}
  .error-box p{color:#424242;line-height:1.6;font-size:14px}
  .hint{color:#757575;font-size:13px;margin-top:16px;line-height:1.5}
  .btn{display:inline-block;padding:12px 20px;border-radius:8px;text-decoration:none;background:#FFE200;color:#000;font-family:'Oswald',sans-serif;font-weight:700;letter-spacing:.5px;margin-top:20px}
  .footer{background:#000;padding:12px 24px;text-align:center;font-size:12px;color:#888;font-family:'Oswald',sans-serif;letter-spacing:.5px}
  .footer span{color:#FFE200}
</style>
<div class="header">
  <img src="https://provinciamarcas.com.br/wp-content/uploads/2025/12/provincia-marcas-e-patentes-pensou-criou-registrou-logo.png" alt="Província Marcas" onerror="this.style.display='none'">
</div>
<div class="main">
  <div class="box">
    <div class="box-header"><h2>❌ Erro ao gerar contrato</h2></div>
    <div class="box-body">
      <div class="error-box">
        <strong>O que aconteceu?</strong>
        <p>${errorMessage}</p>
      </div>
      <p class="hint">Verifique os logs do servidor para mais detalhes. Se o problema persistir, entre em contato com o suporte técnico.</p>
      <a href="${PUBLIC_BASE_URL}/lead/${encodeURIComponent(req.params.token)}" class="btn">← Voltar e tentar novamente</a>
    </div>
  </div>
</div>
<div class="footer">© Província Marcas e Patentes — <span>Pensou. Criou. Província Registrou!</span></div>`);
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
    let isTermo = false;

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
      } else if (tipo === 'termo_de_risco') {
        isTermo = true;
        console.log('[SEND] Tipo identificado como TERMO DE RISCO pelo parâmetro tipo');
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
        // [MODIFICADO] Validação estrita do campo específico
        const telefoneEnvio = d.telefone_envio_contrato || '';
        if (!telefoneEnvio) {
          throw new Error('Telefone para envio do contrato não encontrado. Verifique se o campo "Telefone para envio do contrato" está preenchido no Pipefy.');
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

      console.log(`[SEND] Enviando ${isTermo ? 'Termo de Risco' : (isProcuracao ? 'procuração' : 'contrato')} por ${canal}. Signatários preparados:`, signers.map(s => s.email).join(', '));
    } catch (e) {
      console.warn('[SEND] Erro ao buscar informações do card:', e.message);
      throw e;
    }

    // Verificar status do documento antes de enviar
    await new Promise(r => setTimeout(r, 2000));
    try {
      await getDocumentStatus(D4SIGN_TOKEN, D4SIGN_CRYPT_KEY, uuidDoc);
      console.log(`[SEND] Status do ${isTermo ? 'Termo de Risco' : (isProcuracao ? 'procuração' : 'contrato')} verificado.`);
    } catch (e) {
      console.warn(`[SEND] Aviso ao verificar status do ${isTermo ? 'Termo de Risco' : (isProcuracao ? 'procuração' : 'contrato')}:`, e.message);
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
          // [MODIFICADO] Validação estrita
          const telefoneEnvio = d.telefone_envio_contrato || '';
          if (!telefoneEnvio) {
            throw new Error('Telefone para envio do contrato não encontrado. Verifique se o campo "Telefone para envio do contrato" está preenchido no Pipefy.');
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
    await new Promise(r => setTimeout(r, 2000));

    // Enviar documento (contrato, procuração ou termo de risco)
    try {
      const mensagem = isProcuracao
        ? 'Olá! Há uma procuração aguardando sua assinatura.'
        : (isTermo ? 'Olá! Há um Termo de Risco aguardando sua assinatura.' : 'Olá! Há um documento aguardando sua assinatura.');

      // skip_email: '1' suprime TODAS as notificações externas (embed mode) — não usar para WhatsApp.
      // A supressão do email para WhatsApp é feita via skipemail: '1' por signatário no createlist.
      const skip_email = '0';

      await sendToSigner(D4SIGN_TOKEN, D4SIGN_CRYPT_KEY, uuidDoc, {
        message: mensagem,
        skip_email: skip_email,
        workflow: '0'
      });
      console.log(`[SEND] ${isTermo ? 'Termo de Risco' : (isProcuracao ? 'Procuração' : 'Contrato')} enviado para assinatura por ${canal}:`, uuidDoc);

      // Para WhatsApp: o sendtosigner apenas muda o status do documento para "Aguardando assinatura"
      // mas NÃO dispara a notificação WhatsApp. O endpoint /resend é o que efetivamente
      // envia via embed_methodauth: 'whatse'. Por isso chamamos resend para cada signatário com telefone.
      if (canal === 'whatsapp') {
        try {
          await new Promise(r => setTimeout(r, 3000)); // aguardar D4Sign processar o sendtosigner
          const signersData = await listSigners(D4SIGN_TOKEN, D4SIGN_CRYPT_KEY, uuidDoc);
          console.log(`[SEND-WA] Resposta listSigners:`, JSON.stringify(signersData).substring(0, 500));
          // D4Sign retorna [{ uuidDoc, nameDoc, ..., list: [...signatários] }]
          let signersList = [];
          if (Array.isArray(signersData) && signersData.length > 0 && signersData[0]?.list) {
            // Formato padrão: array com objeto contendo propriedade "list"
            const listData = signersData[0].list;
            signersList = Array.isArray(listData) ? listData : [listData];
          } else if (Array.isArray(signersData) && signersData.length > 0 && signersData[0]?.key_signer) {
            // Formato alternativo: array plano de signatários
            signersList = signersData;
          } else if (signersData?.list) {
            const listData = signersData.list;
            signersList = Array.isArray(listData) ? listData : [listData];
          }
          console.log(`[SEND-WA] ${signersList.length} signatário(s) encontrado(s) para reenvio WhatsApp`);

          for (const s of signersList) {
            if (!s.key_signer) continue;
            // Verificar se o signatário tem WhatsApp configurado
            const emailLower = (s.email || '').toLowerCase();
            const localSigner = signers.find(orig => (orig.email || '').toLowerCase() === emailLower && orig.phone);
            // D4Sign armazena o telefone no campo "email" quando o signatário foi cadastrado via WhatsApp
            // Detectar se s.email é na verdade um número de telefone (só dígitos, opcionalmente com +)
            const emailIsPhone = /^\+?\d{10,15}$/.test((s.email || '').replace(/\s/g, ''));
            // Obter o número WhatsApp: campo explícito, ou email que é telefone, ou do array local
            const whatsappNumber = s.whatsapp || s.whatsapp_number || (emailIsPhone ? s.email : null) || (localSigner ? localSigner.phone : null);
            console.log(`[SEND-WA] Signatário ${s.email}: d4sign_whatsapp=${s.whatsapp}, d4sign_whatsapp_number=${s.whatsapp_number}, emailIsPhone=${emailIsPhone}, local_phone=${localSigner?.phone}, whatsappNumber=${whatsappNumber}`);
            if (!whatsappNumber) {
              console.log(`[SEND-WA] Pulando resend para ${s.email} (sem WhatsApp configurado)`);
              continue;
            }
            try {
              // D4Sign /resend aceita "e-mail ou whatsapp" no campo email
              // Para WhatsApp, enviamos o número de telefone ao invés do email
              await resendToSigner(D4SIGN_TOKEN, D4SIGN_CRYPT_KEY, uuidDoc, whatsappNumber, s.key_signer);
              console.log(`[SEND-WA] ✓ Resend WhatsApp disparado para: ${whatsappNumber} (${s.email})`);
            } catch (resendErr) {
              // Se falhar com WhatsApp, tenta com email como fallback
              console.warn(`[SEND-WA] Falha com WhatsApp (${whatsappNumber}), tentando com email (${s.email})...`, resendErr.message);
              try {
                await resendToSigner(D4SIGN_TOKEN, D4SIGN_CRYPT_KEY, uuidDoc, s.email, s.key_signer);
                console.log(`[SEND-WA] ✓ Resend via email fallback disparado para: ${s.email}`);
              } catch (resendErr2) {
                console.warn(`[SEND-WA] Fallback email também falhou para ${s.email}:`, resendErr2.message);
              }
            }
          }
        } catch (waErr) {
          console.warn(`[SEND-WA] Aviso ao disparar notificações WhatsApp via resend:`, waErr.message);
          // Não bloqueia o fluxo — o documento já está no estado correto
        }
      }

      // Salvar UUID do documento no campo correto após enviar para assinatura
      try {
        console.log(`[SEND] Tentando salvar UUID - isProcuracao: ${isProcuracao}, isTermo: ${isTermo}, uuidDoc: ${uuidDoc}, cardId: ${cardId}`);
        if (isTermo) {
          if (PIPEFY_FIELD_D4_UUID_TERMO_DE_RISCO) {
            const fieldId = PIPEFY_FIELD_D4_UUID_TERMO_DE_RISCO;
            console.log(`[SEND] Salvando UUID do Termo de Risco no campo ${fieldId}...`);
            await updateCardField(cardId, fieldId, uuidDoc);
            console.log(`[SEND] ✓ UUID do Termo de Risco salvo com sucesso no campo ${fieldId}: ${uuidDoc}`);
          } else {
            console.log(`[SEND] Campo PIPEFY_FIELD_D4_UUID_TERMO_DE_RISCO não configurado, UUID não salvo.`);
          }
        } else if (isProcuracao) {
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
        console.error(`[ERRO] Falha ao salvar UUID do ${isTermo ? 'termo de risco' : (isProcuracao ? 'procuração' : 'contrato')} no card:`, e.message);
        console.error(`[ERRO] Stack trace:`, e.stack);
        // Não bloqueia o fluxo se falhar ao salvar UUID
      }
    } catch (e) {
      console.error(`[ERRO] Falha ao enviar ${isTermo ? 'termo de risco' : (isProcuracao ? 'procuração' : 'contrato')}:`, e.message);
      throw e; // Propaga o erro para que o usuário saiba
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
    // Liberar lock após envio bem-sucedido
    releaseLock(lockKey);

    // Preparar listas de emails e telefones enviados
    const emailsEnviados = signers ? signers.map(s => s.email).filter(Boolean) : [];
    const telefonesEnviados = signers ? signers.map(s => s.phone).filter(Boolean) : [];

    return res.status(200).json({
      success: true,
      message: `${isProcuracao ? 'Procuração' : 'Contrato'} enviado com sucesso. Os signatários foram notificados.`,
      tipo: isProcuracao ? 'procuração' : 'contrato',
      cofre: nomeCofre,
      urlCofre: urlCofre,
      email: emailsEnviados.join(', '), // Mantendo compatibilidade com string, mas agora lista
      telefone: telefonesEnviados.join(', '), // Mantendo compatibilidade
      emails: emailsEnviados,
      telefones: telefonesEnviados
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

app.post('/lead/:token/doc/:uuid/resend', async (req, res) => {
  const { cardId } = parseLeadToken(req.params.token);
  if (!cardId) {
    return res.status(400).json({ success: false, message: 'Token inválido' });
  }
  const uuidDoc = req.params.uuid;

  try {
    // 1. Listar signatários para obter key_signer
    const signersData = await listSigners(D4SIGN_TOKEN, D4SIGN_CRYPT_KEY, uuidDoc);
    console.log(`[RESEND] Resposta listSigners:`, JSON.stringify(signersData).substring(0, 500));
    // D4Sign retorna [{ uuidDoc, nameDoc, ..., list: [...signatários] }]
    let signersList = [];
    if (Array.isArray(signersData) && signersData.length > 0 && signersData[0]?.list) {
      const listData = signersData[0].list;
      signersList = Array.isArray(listData) ? listData : [listData];
    } else if (Array.isArray(signersData) && signersData.length > 0 && signersData[0]?.key_signer) {
      signersList = signersData;
    } else if (signersData?.list) {
      const listData = signersData.list;
      signersList = Array.isArray(listData) ? listData : [listData];
    }
    console.log(`[RESEND] ${signersList.length} signatário(s) encontrado(s)`);

    if (!signersList || signersList.length === 0) {
      return res.status(404).json({ success: false, message: 'Nenhum signatário encontrado para este documento.' });
    }

    const resultados = [];
    let errors = 0;

    // 2. Reenviar para cada signatário
    // Para signatários com WhatsApp, o campo "email" do /resend aceita "e-mail ou whatsapp"
    for (const signer of signersList) {
      if (!signer.key_signer) continue;

      // Determinar se o signatário tem WhatsApp configurado
      const whatsappNumber = signer.whatsapp || signer.whatsapp_number || '';
      // Se tem WhatsApp, enviar o número; senão, enviar o email
      const identificador = whatsappNumber || signer.email;
      console.log(`[RESEND] Signatário: email=${signer.email}, whatsapp=${signer.whatsapp}, whatsapp_number=${signer.whatsapp_number}, usando=${identificador}`);

      try {
        await resendToSigner(D4SIGN_TOKEN, D4SIGN_CRYPT_KEY, uuidDoc, identificador, signer.key_signer);
        resultados.push({ email: signer.email, whatsapp: whatsappNumber || null, status: 'enviado' });
      } catch (e) {
        console.error(`[RESEND] Erro ao reenviar para ${identificador}:`, e.message);
        // Se falhou com WhatsApp, tenta com email como fallback
        if (whatsappNumber && signer.email) {
          try {
            console.log(`[RESEND] Tentando fallback com email para ${signer.email}...`);
            await resendToSigner(D4SIGN_TOKEN, D4SIGN_CRYPT_KEY, uuidDoc, signer.email, signer.key_signer);
            resultados.push({ email: signer.email, status: 'enviado (fallback email)' });
            continue;
          } catch (e2) {
            console.error(`[RESEND] Fallback email também falhou para ${signer.email}:`, e2.message);
          }
        }
        resultados.push({ email: signer.email, status: 'erro', error: e.message });
        errors++;
      }
    }

    if (errors === signersList.length) {
      return res.status(500).json({ success: false, message: 'Falha ao reenviar para todos os signatários.' });
    }

    return res.status(200).json({
      success: true,
      message: 'Reenvio solicitado com sucesso.',
      detalhes: resultados
    });

  } catch (e) {
    console.error('[ERRO RESEND]', e.message);
    return res.status(500).json({ success: false, message: 'Erro ao processar reenvio: ' + e.message });
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
          // Verifica se o UUID está nos campos D4 UUID Contrato, Procuração ou Termo de Risco
          if ((field.id === PIPEFY_FIELD_D4_UUID_CONTRATO || field.id === PIPEFY_FIELD_D4_UUID_PROCURACAO ||
            (PIPEFY_FIELD_D4_UUID_TERMO_DE_RISCO && field.id === PIPEFY_FIELD_D4_UUID_TERMO_DE_RISCO)) &&
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
    if (NOVO_PIPE_ID && String(card?.pipe?.id) !== String(NOVO_PIPE_ID)) {
      return res.status(400).json({ error: 'Card não pertence ao pipe configurado' });
    }
    if (FASE_VISITA_ID && String(card?.current_phase?.id) !== String(FASE_VISITA_ID)) {
      return res.status(400).json({ error: 'Card não está na fase esperada' });
    }

    const token = makeLeadToken({ cardId: String(cardId), ts: Date.now() });
    const url = `${PUBLIC_BASE_URL.replace(/\/+$/, '')}/lead/${encodeURIComponent(token)}`;

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

    return res.json({ ok: true, link: url });
  } catch (e) {
    console.error('[ERRO criar-link]', e.message || e);
    return res.status(500).json({ error: String(e.message || e) });
  }
});
app.get('/novo-pipe/criar-link-confirmacao', async (req, res) => {
  try {
    const cardId = req.query.cardId || req.query.card_id;
    if (!cardId) return res.status(400).json({ error: 'cardId é obrigatório' });

    const card = await getCard(cardId);
    if (NOVO_PIPE_ID && String(card?.pipe?.id) !== String(NOVO_PIPE_ID)) {
      return res.status(400).json({ error: 'Card não pertence ao pipe configurado' });
    }
    if (FASE_VISITA_ID && String(card?.current_phase?.id) !== String(FASE_VISITA_ID)) {
      return res.status(400).json({ error: 'Card não está na fase esperada' });
    }

    const token = makeLeadToken({ cardId: String(cardId), ts: Date.now() });
    const url = `${PUBLIC_BASE_URL.replace(/\/+$/, '')}/lead/${encodeURIComponent(token)}`;

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

    return res.json({ ok: true, link: url });
  } catch (e) {
    console.error('[ERRO criar-link]', e.message || e);
    return res.status(500).json({ error: String(e.message || e) });
  }
});

/* =========================
 * Lógica Central de Geração
 * =======================*/
async function processarContrato(cardId) {
  const card = await getCard(cardId);

  // Proteção contra reprocessamento: se já existe UUID de contrato no card, ignora
  const by = toById(card);
  const uuidExistente = by[PIPEFY_FIELD_D4_UUID_CONTRATO] || '';
  if (uuidExistente) {
    console.log(`[PROCESSAR] Card ${cardId} já possui contrato gerado (${uuidExistente}). Ignorando.`);
    return { uuidDoc: uuidExistente, uuidProcuracao: null, uuidTermoDeRisco: null };
  }

  const d = await montarDados(card);

  const now = new Date();
  const nowInfo = { dia: now.getDate(), mes: now.getMonth() + 1, ano: now.getFullYear() };

  // Validar template
  if (!d.templateToUse) {
    throw new Error('Template não identificado. Verifique os dados do card.');
  }

  const isMarcaTemplate = d.templateToUse === TEMPLATE_UUID_CONTRATO;
  const add = isMarcaTemplate ? montarVarsParaTemplateMarca(d, nowInfo)
    : montarVarsParaTemplateOutros(d, nowInfo);

  // Seleciona cofre
  const equipeContrato = getEquipeContratoFromCard(card);
  let uuidSafe = COFRES_UUIDS[equipeContrato] || DEFAULT_COFRE_UUID;

  if (!uuidSafe) throw new Error('Nenhum cofre disponível.');

  console.log(`[PROCESSAR] Criando contrato no cofre: ${uuidSafe}`);

  const uuidDoc = await makeDocFromWordTemplate(
    D4SIGN_TOKEN,
    D4SIGN_CRYPT_KEY,
    uuidSafe,
    d.templateToUse,
    d.titulo || card.title || 'Contrato',
    add
  );

  if (!uuidDoc) throw new Error('Falha ao criar documento no D4Sign.');

  // Webhook
  try {
    await registerWebhookForDocument(D4SIGN_TOKEN, D4SIGN_CRYPT_KEY, uuidDoc, `${PUBLIC_BASE_URL}/d4sign/postback`);
    // [NOVO] Salvar UUID no card
    await updateCardField(cardId, PIPEFY_FIELD_D4_UUID_CONTRATO, uuidDoc);
    console.log(`[PROCESSAR] UUID Contrato salvo no card: ${uuidDoc}`);
  } catch (e) { console.error('Erro webhook/salvar UUID:', e.message); }

  // Procuração (Opcional)
  let uuidProcuracao = null;
  if (TEMPLATE_UUID_PROCURACAO) {
    try {
      const varsProcuracao = montarVarsParaTemplateProcuracao(d, nowInfo);
      uuidProcuracao = await makeDocFromWordTemplate(
        D4SIGN_TOKEN,
        D4SIGN_CRYPT_KEY,
        uuidSafe,
        TEMPLATE_UUID_PROCURACAO,
        `Procuração - ${d.titulo || card.title}`,
        varsProcuracao
      );
      // Webhook procuração...
      // [NOVO] Salvar UUID no card
      await updateCardField(cardId, PIPEFY_FIELD_D4_UUID_PROCURACAO, uuidProcuracao);
      console.log(`[PROCESSAR] UUID Procuração salvo no card: ${uuidProcuracao}`);
    } catch (e) { console.error('Erro procuração:', e.message); }
  }

  // Termo de Risco: Removido — agora integrado apenas no Template 2 (Marca com Risco)
  // Nenhum documento separado de Termo de Risco será criado
  const uuidTermoDeRisco = null;

  // [REMOVIDO] Cadastrar signatários automaticamente
  // O cadastro será feito apenas quando o usuário clicar em "Enviar por Email"
  /*
  try {
    // Aguardar um pouco para garantir que o documento foi processado pelo D4Sign
    await new Promise(r => setTimeout(r, 3000));

    const signers = montarSigners(d);
    if (signers && signers.length > 0) {
      console.log(`[PROCESSAR] Cadastrando ${signers.length} signatários no contrato ${uuidDoc}...`);
      await cadastrarSignatarios(D4SIGN_TOKEN, D4SIGN_CRYPT_KEY, uuidDoc, signers);

      if (uuidProcuracao) {
        console.log(`[PROCESSAR] Cadastrando ${signers.length} signatários na procuração ${uuidProcuracao}...`);
        await cadastrarSignatarios(D4SIGN_TOKEN, D4SIGN_CRYPT_KEY, uuidProcuracao, signers);
      }
    } else {
      console.warn('[PROCESSAR] Nenhum signatário encontrado para cadastrar.');
    }
  } catch (e) {
    console.error('[PROCESSAR] Erro ao cadastrar signatários:', e.message);
  }
  */

  return { uuidDoc, uuidProcuracao, uuidTermoDeRisco };
}

/* =========================
 * Webhook Pipefy
 * =======================*/
app.post('/pipefy-webhook', async (req, res) => {
  try {
    const { data } = req.body || {};
    if (!data || !data.card) return res.json({ ok: true });

    const cardId = data.card.id;
    console.log(`[WEBHOOK] Recebido para card ${cardId}`);

    // Proteção contra webhooks duplicados do Pipefy (mesmo card disparado 2-3x)
    const lockKey = `webhook:${cardId}`;
    if (!acquireLock(lockKey)) {
      console.log(`[WEBHOOK] Card ${cardId} já está sendo processado. Ignorando duplicata.`);
      return res.json({ ok: true });
    }

    try {
      await processarContrato(cardId);
    } finally {
      releaseLock(lockKey);
    }

    res.json({ ok: true });
  } catch (e) {
    console.error('[WEBHOOK ERROR]', e);
    res.status(500).json({ error: e.message });
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
app.get('/debug/card', async (req, res) => {
  try {
    const { cardId } = req.query; if (!cardId) return res.status(400).send('cardId obrigatório');
    const card = await getCard(cardId);
    res.json({
      id: card.id, title: card.title, pipe: card.pipe, phase: card.current_phase,
      fields: (card.fields || []).map(f => ({ name: f.name, id: f.field?.id, type: f.field?.type, value: f.value, array_value: f.array_value }))
    });
  } catch (e) { res.status(500).json({ error: String(e.message || e) }); }
});
app.get('/health', (_req, res) => res.json({ ok: true }));

// ===============================
// NOVO — Upload de arquivo para o Pipefy
// ===============================
async function uploadFileToPipefy(url, fileName, organizationId) {
  try {
    console.log(`[UPLOAD PIPEFY] Baixando arquivo de: ${url}`);

    // 1. Baixar arquivo (buffer)
    const res = await fetchWithRetry(url, { method: 'GET' }, { attempts: 3 });
    if (!res.ok) throw new Error(`Falha ao baixar arquivo: ${res.status}`);
    const buffer = await res.arrayBuffer();

    console.log(`[UPLOAD PIPEFY] Arquivo baixado (${buffer.byteLength} bytes). Obtendo URL de upload...`);

    // 2. Obter URL presignada do Pipefy
    const mutation = `
      mutation($input: CreatePresignedUrlInput!) {
        createPresignedUrl(input: $input) {
          url
          downloadUrl
        }
      }
    `;

    const data = await gql(mutation, {
      input: {
        organizationId: organizationId,
        fileName: fileName
      }
    });

    const { url: uploadUrl, downloadUrl } = data.createPresignedUrl;
    console.log(`[UPLOAD PIPEFY] URL de upload obtida: ${uploadUrl}`);

    // 3. Fazer upload para o S3 do Pipefy
    const uploadRes = await fetch(uploadUrl, {
      method: 'PUT',
      body: Buffer.from(buffer),
      headers: {
        'Content-Type': 'application/pdf' // Assumindo PDF
      }
    });

    if (!uploadRes.ok) {
      throw new Error(`Falha ao fazer upload para o Pipefy: ${uploadRes.status} ${uploadRes.statusText}`);
    }

    console.log(`[UPLOAD PIPEFY] Upload concluído com sucesso.`);

    // 4. Extrair o caminho relativo do S3 da URL presignada
    // O Pipefy exige que campos de anexo usem apenas o caminho relativo, não a URL completa
    // Formato esperado: orgs/{org-id}/uploads/{upload-id}/filename.pdf
    // A URL presignada tem formato: https://...s3...amazonaws.com/orgs/.../uploads/.../file.pdf?...

    let attachmentPath = '';
    try {
      // Tentar extrair o caminho da URL de upload
      const urlObj = new URL(uploadUrl);
      const pathname = urlObj.pathname; // Ex: /orgs/.../uploads/.../file.pdf
      // Remover a barra inicial se houver e decodificar encoding da URL
      attachmentPath = decodeURIComponent(pathname.startsWith('/') ? pathname.substring(1) : pathname);
    } catch (parseErr) {
      // Fallback: tentar extrair usando regex
      const match = uploadUrl.match(/\/(orgs\/[^?]+)/);
      if (match) {
        attachmentPath = decodeURIComponent(match[1]);
      }
    }

    if (!attachmentPath) {
      throw new Error(`Não foi possível extrair o caminho do anexo da URL: ${uploadUrl}`);
    }

    console.log(`[UPLOAD PIPEFY] Caminho do anexo extraído: ${attachmentPath}`);
    console.log(`[UPLOAD PIPEFY] Download URL (para referência): ${downloadUrl}`);

    // Retorna o caminho relativo para usar no campo de anexo do Pipefy
    return attachmentPath;

  } catch (e) {
    console.error('[UPLOAD PIPEFY ERROR]', e.message);
    throw e;
  }
}

async function saveFileLocally(downloadUrl, fileName, equipeNome) {
  const localBase = process.env.LOCAL_FOLDER_PATH;
  if (!localBase) {
    console.warn('[SAVE LOCAL] LOCAL_FOLDER_PATH não configurado. Pulando salvamento local.');
    return;
  }

  try {
    const pastaSegura = String(equipeNome || 'Sem_Cofre').replace(/[<>:"/\\|?*]/g, '_').trim() || 'Sem_Cofre';
    const pastaDestino = path.join(localBase, pastaSegura);

    await fs.ensureDir(pastaDestino);

    const res = await fetchWithRetry(downloadUrl, { method: 'GET' }, { attempts: 3 });
    if (!res.ok) throw new Error(`Falha ao baixar para salvar localmente: ${res.status}`);
    const buffer = await res.arrayBuffer();

    const caminhoArquivo = path.join(pastaDestino, fileName);
    await fs.writeFile(caminhoArquivo, Buffer.from(buffer));
    console.log(`[SAVE LOCAL] ✓ Arquivo salvo em: ${caminhoArquivo}`);
  } catch (e) {
    console.error('[SAVE LOCAL] Erro ao salvar localmente:', e.message);
  }
}

/* =========================
 * Start
 * =======================*/
app.listen(PORT, () => {
  console.log(`Servidor rodando na porta ${PORT}`);
  const list = [];
  app._router.stack.forEach(m => {
    if (m.route && m.route.path) {
      const methods = Object.keys(m.route.methods).map(x => x.toUpperCase()).join(',');
      list.push(`${methods} ${m.route.path}`);
    } else if (m.name === 'router' && m.handle?.stack) {
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
