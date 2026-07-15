'use strict';

/**
 * mailer.js — Envio de alertas por email (SMTP corporativo).
 *
 * Módulo isolado: não depende de nada do server.js. Toda função de envio é
 * "à prova de falha" — se o SMTP não estiver configurado, se o toggle estiver
 * desligado, ou se o próprio envio der erro, apenas registra em log e retorna.
 * NUNCA lança exceção para o chamador, para não quebrar o fluxo principal.
 *
 * Config via env vars:
 *   SMTP_HOST, SMTP_PORT (587), SMTP_SECURE (false), SMTP_USER, SMTP_PASS
 *   ALERT_EMAIL_FROM (default = SMTP_USER)
 *   ALERT_EMAIL_TO   (obrigatório para enviar; aceita vários separados por vírgula)
 *   ALERTS_ENABLED   ('false' força desligar; caso contrário liga se SMTP configurado)
 *   ALERT_MIN_INTERVAL_MS (janela mínima entre alertas idênticos; default 30 min)
 */

const nodemailer = require('nodemailer');

const {
  SMTP_HOST,
  SMTP_PORT = '587',
  SMTP_SECURE = 'false',
  SMTP_USER,
  SMTP_PASS,
  ALERT_EMAIL_FROM,
  ALERT_EMAIL_TO,
  ALERTS_ENABLED,
  ALERT_MIN_INTERVAL_MS = '1800000', // 30 min
} = process.env;

const minIntervalMs = Number(ALERT_MIN_INTERVAL_MS) || 1800000;
const fromAddr = ALERT_EMAIL_FROM || SMTP_USER;

// Liga se: toggle != 'false' E temos host + destinatário.
const enabled =
  String(ALERTS_ENABLED).toLowerCase() !== 'false' &&
  Boolean(SMTP_HOST) &&
  Boolean(ALERT_EMAIL_TO);

let transporter = null;
if (enabled) {
  transporter = nodemailer.createTransport({
    host: SMTP_HOST,
    port: Number(SMTP_PORT) || 587,
    secure: String(SMTP_SECURE).toLowerCase() === 'true', // true = 465
    auth: SMTP_USER ? { user: SMTP_USER, pass: SMTP_PASS } : undefined,
  });
  console.log(`[MAILER] Ativo. host=${SMTP_HOST} port=${SMTP_PORT} para=${ALERT_EMAIL_TO}`);
} else {
  console.warn('[MAILER] Desativado (SMTP não configurado ou ALERTS_ENABLED=false). Alertas irão só para o log.');
}

// Throttle/dedup em memória: chave → timestamp do último envio.
const ultimoEnvio = new Map();

function throttled(chave) {
  const agora = Date.now();
  const anterior = ultimoEnvio.get(chave) || 0;
  if (agora - anterior < minIntervalMs) return true;
  ultimoEnvio.set(chave, agora);
  return false;
}

function agoraSP() {
  // Horário de São Paulo, formato legível.
  try {
    return new Date().toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' });
  } catch {
    return new Date().toISOString();
  }
}

/**
 * Envio de baixo nível. `throttleKey` (opcional) aplica dedup por janela.
 * Retorna true se enviou, false se suprimido/falhou. Nunca lança.
 */
async function enviarEmail({ subject, text, html, throttleKey }) {
  if (!enabled) {
    console.log(`[MAILER] (suprimido — desativado) ${subject}`);
    return false;
  }
  if (throttleKey && throttled(throttleKey)) {
    console.log(`[MAILER] (suprimido — throttle "${throttleKey}") ${subject}`);
    return false;
  }
  try {
    await transporter.sendMail({
      from: fromAddr,
      to: ALERT_EMAIL_TO,
      subject,
      text,
      html,
    });
    console.log(`[MAILER] ✓ Enviado: ${subject}`);
    return true;
  } catch (e) {
    console.error(`[MAILER] Falha ao enviar "${subject}": ${e.message}`);
    return false;
  }
}

function linha(label, valor) {
  return `${label}: ${valor == null || valor === '' ? '—' : valor}`;
}

/**
 * Alerta de erro TERMINAL ao salvar um arquivo localmente (não é queda de conexão).
 * Throttle por código de erro para não repetir em rajada.
 */
async function sendSaveErrorAlert({ fileName, marca, equipe, downloadUrl, destino, erro, tentativas }) {
  const code = (erro && (erro.code || 'SEM_CODE')) || 'SEM_CODE';
  const msg = (erro && erro.message) || String(erro);
  const stack = (erro && erro.stack) ? String(erro.stack).split('\n').slice(0, 5).join('\n') : '';

  const corpo = [
    '⚠️  Falha ao salvar documento localmente (erro que precisa de correção).',
    '',
    linha('Horário', agoraSP()),
    linha('Arquivo', fileName),
    linha('Marca', marca),
    linha('Equipe', equipe),
    linha('URL de origem', downloadUrl),
    linha('Destino pretendido', destino),
    linha('Erro (code)', code),
    linha('Mensagem', msg),
    linha('Tentativas até agora', tentativas),
    '',
    'O arquivo foi colocado na fila de reprocessamento e será tentado novamente automaticamente.',
    stack ? `\nStack (resumido):\n${stack}` : '',
  ].join('\n');

  return enviarEmail({
    subject: `[Provincia] Falha ao salvar "${fileName}" (${code})`,
    text: corpo,
    throttleKey: `save-error:${code}`,
  });
}

/** Alerta: a comunicação com a VPN/mount caiu (transição online→offline). */
async function sendConnectivityDownAlert({ erro, mountPath }) {
  const msg = (erro && erro.message) || String(erro || 'desconhecido');
  const code = (erro && erro.code) || '—';
  const corpo = [
    '🔴 A comunicação entre a VPS e a VPN/pasta de rede CAIU.',
    '',
    linha('Horário da detecção', agoraSP()),
    linha('Pasta de rede', mountPath),
    linha('Erro (code)', code),
    linha('Mensagem', msg),
    '',
    'Enquanto estiver fora, os documentos NÃO são salvos na rede — ficam numa fila',
    'local e serão reprocessados automaticamente quando a conexão voltar.',
    '',
    'Ação sugerida: verificar OpenVPN/CIFS na VPS e o site remoto.',
  ].join('\n');

  return enviarEmail({
    subject: '[Provincia] 🔴 Comunicação VPS↔VPN CAIU',
    text: corpo,
    throttleKey: 'connectivity-down',
  });
}

/** Alerta: a comunicação voltou (transição offline→online). */
async function sendConnectivityUpAlert({ mountPath, foraDesde }) {
  const corpo = [
    '🟢 A comunicação entre a VPS e a VPN/pasta de rede foi RESTABELECIDA.',
    '',
    linha('Horário da recuperação', agoraSP()),
    linha('Pasta de rede', mountPath),
    linha('Estava fora desde', foraDesde ? new Date(foraDesde).toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' }) : '—'),
    '',
    'A fila de documentos pendentes será reprocessada automaticamente.',
  ].join('\n');

  // Sem throttle: recuperação é rara e importante.
  return enviarEmail({
    subject: '[Provincia] 🟢 Comunicação VPS↔VPN restabelecida',
    text: corpo,
  });
}

/** Alerta: um item da fila estourou o número máximo de tentativas. */
async function sendQueueGaveUpAlert({ item }) {
  const corpo = [
    '⛔ Um documento na fila de reprocessamento excedeu o máximo de tentativas.',
    'Ele PERMANECE na fila para inspeção manual (não foi descartado).',
    '',
    linha('Horário', agoraSP()),
    linha('Arquivo', item.fileName),
    linha('Marca', item.nomeMarca),
    linha('Equipe', item.equipe),
    linha('URL de origem', item.downloadUrl),
    linha('Tentativas', item.tentativas),
    linha('Último erro', item.ultimoErro),
    linha('ID na fila', item.id),
  ].join('\n');

  return enviarEmail({
    subject: `[Provincia] ⛔ Documento na fila precisa de atenção manual: "${item.fileName}"`,
    text: corpo,
    throttleKey: `queue-gaveup:${item.id}`,
  });
}

module.exports = {
  enabled,
  sendSaveErrorAlert,
  sendConnectivityDownAlert,
  sendConnectivityUpAlert,
  sendQueueGaveUpAlert,
};
