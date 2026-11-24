/**
 * Recebe os leads do quiz (nome, WhatsApp, respostas, perfil, atribuição
 * de campanha) e grava em tempo real numa aba da planilha. Também recebe
 * atualizações de status de pagamento vindas do webhook da Lastlink
 * (via api/lastlink-webhook.js na Vercel), casando pelo WhatsApp.
 *
 * SETUP:
 * 1. Crie uma planilha nova no Google Sheets.
 * 2. Extensões → Apps Script. Apague o conteúdo padrão e cole este arquivo inteiro.
 * 3. Implantar → Nova implantação → tipo "App da Web".
 *    - Executar como: Eu (sua conta)
 *    - Quem pode acessar: Qualquer pessoa
 * 4. Autorize as permissões pedidas (é sua própria planilha).
 * 5. Copie a URL do App da Web gerada (termina em /exec) e me envie —
 *    eu colo ela na constante LEADS_WEBHOOK_URL do quiz e também na
 *    variável de ambiente LEADS_WEBHOOK_URL da Vercel.
 *
 * Se você já tinha uma versão anterior deste script publicada: só colar
 * este arquivo por cima e gerar uma "Nova versão" no mesmo deploy já
 * basta — o cabeçalho da planilha é atualizado automaticamente na
 * primeira gravação depois disso, sem apagar linhas existentes.
 */

const SHEET_NAME = "Leads";
const HEADERS = [
  "Data/Hora", "Tracking ID", "Nome", "WhatsApp", "WhatsApp (link)", "Perfil",
  "Barras", "O que trava", "Progressão", "Tempo tentando", "Comprometimento",
  "UTM Source", "UTM Medium", "UTM Campaign", "UTM Content", "UTM Term", "FBCLID",
  "Landing Page", "Referrer", "Primeira Visita", "Lead Event ID",
  "Status Pix", "Pix ID", "Valor Pago", "Pago Em"
];

const STATUS_POR_EVENTO = {
  "pix.generated": "Pix Gerado",
  "pix.paid": "Pago",
  "pix.expired": "Pix Expirado",
  "pix.failed": "Falhou",
  "pix.refunded": "Reembolsado",
};

function doPost(e) {
  const data = JSON.parse(e.postData.contents);

  const resultado = data.tipo === "atualizar_pagamento"
    ? atualizarPagamento_(data)
    : registrarLead_(data);

  return ContentService
    .createTextOutput(JSON.stringify(resultado))
    .setMimeType(ContentService.MimeType.JSON);
}

function registrarLead_(data) {
  const sheet = getOrCreateSheet_();
  const whatsappLimpo = (data.whatsapp || "").replace(/\D/g, "");
  const whatsappLink = whatsappLimpo
    ? `https://wa.me/55${whatsappLimpo}`
    : "";

  sheet.appendRow([
    new Date(),
    data.tracking_id || "",
    data.nome || "",
    data.whatsapp || "",
    whatsappLink,
    data.perfil || "",
    data.respostas?.barras || "",
    data.respostas?.trava || "",
    data.respostas?.progressao || "",
    data.respostas?.tempoTentando || "",
    data.respostas?.comprometimento || "",
    data.utm_source || "",
    data.utm_medium || "",
    data.utm_campaign || "",
    data.utm_content || "",
    data.utm_term || "",
    data.fbclid || "",
    data.landing_page || "",
    data.referrer || "",
    data.first_visit_at || "",
    data.lead_event_id || "",
    "Pendente", "", "", ""
  ]);

  return { ok: true };
}

function atualizarPagamento_(data) {
  const sheet = getOrCreateSheet_();
  const status = STATUS_POR_EVENTO[data.event] || data.paymentStatus || "Pendente";
  const telefoneAlvo = ultimosDigitos_(data.whatsapp);

  const colWhatsapp = HEADERS.indexOf("WhatsApp") + 1;
  const colStatus = HEADERS.indexOf("Status Pix") + 1;
  const colPixId = HEADERS.indexOf("Pix ID") + 1;
  const colValor = HEADERS.indexOf("Valor Pago") + 1;
  const colPagoEm = HEADERS.indexOf("Pago Em") + 1;

  const linhaEncontrada = telefoneAlvo
    ? encontrarLinhaPorTelefone_(sheet, colWhatsapp, telefoneAlvo)
    : 0;

  if (linhaEncontrada > 0) {
    sheet.getRange(linhaEncontrada, colStatus).setValue(status);
    if (data.paymentId) sheet.getRange(linhaEncontrada, colPixId).setValue(data.paymentId);
    if (data.amount) sheet.getRange(linhaEncontrada, colValor).setValue(data.amount);
    if (data.event === "pix.paid") sheet.getRange(linhaEncontrada, colPagoEm).setValue(new Date());
    return { ok: true, linha_atualizada: linhaEncontrada };
  }

  // Nenhum lead correspondente encontrado — registra mesmo assim pra não perder o pagamento.
  sheet.appendRow([
    new Date(), "", "(sem lead correspondente)", data.whatsapp || "", "", "",
    "", "", "", "", "",
    data.utm_source || "", "", data.utm_campaign || "", "", "", "",
    "", "", "", "",
    status, data.paymentId || "", data.amount || "",
    data.event === "pix.paid" ? new Date() : ""
  ]);
  return { ok: true, linha_atualizada: null };
}

function ultimosDigitos_(telefone) {
  const digitos = (telefone || "").replace(/\D/g, "");
  return digitos.length >= 9 ? digitos.slice(-9) : digitos;
}

function encontrarLinhaPorTelefone_(sheet, colWhatsapp, telefoneAlvo) {
  const ultimaLinha = sheet.getLastRow();
  if (ultimaLinha < 2) return 0;
  const valores = sheet.getRange(2, colWhatsapp, ultimaLinha - 1, 1).getValues();
  for (let i = valores.length - 1; i >= 0; i--) {
    if (ultimosDigitos_(valores[i][0]) === telefoneAlvo) return i + 2;
  }
  return 0;
}

function getOrCreateSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  const headerAtual = sheet.getLastRow() > 0
    ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    : [];
  const precisaAtualizarHeader = HEADERS.some((h, i) => headerAtual[i] !== h);
  if (precisaAtualizarHeader) {
    const range = sheet.getRange(1, 1, 1, HEADERS.length);
    range.setValues([HEADERS]);
    sheet.setFrozenRows(1);
    range.setFontWeight("bold");
  }

  return sheet;
}
