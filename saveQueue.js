'use strict';

/**
 * saveQueue.js — Fila persistente em disco para documentos cujo salvamento local
 * falhou. Cada item é um arquivo JSON dentro de QUEUE_FOLDER_PATH.
 *
 * IMPORTANTE: QUEUE_FOLDER_PATH deve apontar para um disco LOCAL da VPS (bind mount
 * no host), NUNCA para dentro de LOCAL_FOLDER_PATH (a pasta de rede). A fila precisa
 * sobreviver justamente à queda de conexão que ela existe para tratar.
 *
 * Módulo isolado: depende apenas de fs-extra + path. Todas as operações são
 * "à prova de falha" — em erro, registram log e retornam valor neutro, sem lançar.
 */

const fs = require('fs-extra');
const path = require('path');
const crypto = require('crypto');

const QUEUE_FOLDER_PATH = process.env.QUEUE_FOLDER_PATH || '/opt/provincia/save-queue';

function agoraISO() {
  return new Date().toISOString();
}

/** Chave de deduplicação estável: origem + destino final pretendido. */
function chaveDedup({ downloadUrl, fileName, equipe, nomeMarca, subpasta }) {
  const bruto = [downloadUrl, fileName, equipe, nomeMarca, subpasta].join('|');
  return crypto.createHash('sha1').update(bruto).digest('hex').slice(0, 12);
}

async function garantirPasta() {
  await fs.ensureDir(QUEUE_FOLDER_PATH);
}

/**
 * Enfileira um item (idempotente): se já houver um pendente com a mesma chave de
 * dedup, não duplica. Retorna o item persistido, ou null em caso de erro.
 */
async function enfileirar(dados) {
  try {
    await garantirPasta();
    const dedup = chaveDedup(dados);

    // Idempotência: não duplica item pendente com a mesma origem/destino.
    const pendentes = await listarPendentes();
    const jaExiste = pendentes.find(p => p.dedup === dedup);
    if (jaExiste) {
      console.log(`[SAVE QUEUE] Item já pendente, não duplica (dedup=${dedup}, arquivo=${dados.fileName}).`);
      return jaExiste;
    }

    const id = `${Date.now()}-${crypto.randomBytes(3).toString('hex')}`;
    const item = {
      id,
      dedup,
      downloadUrl: dados.downloadUrl,
      fileName: dados.fileName,
      equipe: dados.equipe,
      nomeMarca: dados.nomeMarca,
      subpasta: dados.subpasta || '',
      criadoEm: agoraISO(),
      tentativas: 0,
      ultimoErro: dados.ultimoErro || null,
      ultimaTentativaEm: null,
    };

    const arquivo = path.join(QUEUE_FOLDER_PATH, `${id}.json`);
    await fs.writeJson(arquivo, item, { spaces: 2 });
    console.log(`[SAVE QUEUE] ✓ Enfileirado: ${dados.fileName} (id=${id})`);
    return item;
  } catch (e) {
    console.error(`[SAVE QUEUE] Falha ao enfileirar "${dados && dados.fileName}": ${e.message}`);
    return null;
  }
}

/** Lista os itens pendentes (ordenados do mais antigo para o mais novo). */
async function listarPendentes() {
  try {
    await garantirPasta();
    const arquivos = (await fs.readdir(QUEUE_FOLDER_PATH)).filter(f => f.endsWith('.json'));
    const itens = [];
    for (const nome of arquivos) {
      try {
        const item = await fs.readJson(path.join(QUEUE_FOLDER_PATH, nome));
        itens.push(item);
      } catch (e) {
        console.error(`[SAVE QUEUE] Item corrompido ignorado (${nome}): ${e.message}`);
      }
    }
    itens.sort((a, b) => String(a.criadoEm).localeCompare(String(b.criadoEm)));
    return itens;
  } catch (e) {
    console.error(`[SAVE QUEUE] Falha ao listar pendentes: ${e.message}`);
    return [];
  }
}

/** Remove um item da fila (após salvar com sucesso). */
async function remover(id) {
  try {
    await fs.remove(path.join(QUEUE_FOLDER_PATH, `${id}.json`));
    return true;
  } catch (e) {
    console.error(`[SAVE QUEUE] Falha ao remover id=${id}: ${e.message}`);
    return false;
  }
}

/** Atualiza campos de um item (ex.: tentativas, ultimoErro). Retorna o item atualizado. */
async function atualizar(id, patch) {
  try {
    const arquivo = path.join(QUEUE_FOLDER_PATH, `${id}.json`);
    const item = await fs.readJson(arquivo);
    const novo = { ...item, ...patch };
    await fs.writeJson(arquivo, novo, { spaces: 2 });
    return novo;
  } catch (e) {
    console.error(`[SAVE QUEUE] Falha ao atualizar id=${id}: ${e.message}`);
    return null;
  }
}

async function contarPendentes() {
  const itens = await listarPendentes();
  return itens.length;
}

module.exports = {
  QUEUE_FOLDER_PATH,
  enfileirar,
  listarPendentes,
  remover,
  atualizar,
  contarPendentes,
};
