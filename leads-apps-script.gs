/**
 * Recebe os leads do quiz (nome, WhatsApp, respostas, perfil) e grava
 * em tempo real numa aba da planilha, com status de pagamento manual
 * para você acompanhar quem não gerou/pagou o Pix.
 *
 * SETUP:
 * 1. Crie uma planilha nova no Google Sheets.
 * 2. Extensões → Apps Script. Apague o conteúdo padrão e cole este arquivo inteiro.
 * 3. Implantar → Nova implantação → tipo "App da Web".
 *    - Executar como: Eu (sua conta)
 *    - Quem pode acessar: Qualquer pessoa
 * 4. Autorize as permissões pedidas (é sua própria planilha).
 * 5. Copie a URL do App da Web gerada (termina em /exec) e me envie —
 *    eu colo ela na constante LEADS_WEBHOOK_URL do quiz.
 */

const SHEET_NAME = "Leads";
const HEADERS = [
  "Data/Hora", "Nome", "WhatsApp", "WhatsApp (link)", "Perfil",
  "Barras", "O que trava", "Progressão", "Tempo tentando",
  "UTM Source", "UTM Campaign", "Status Pix"
];

function doPost(e) {
  const sheet = getOrCreateSheet_();
  const data = JSON.parse(e.postData.contents);

  const whatsappLimpo = (data.whatsapp || "").replace(/\D/g, "");
  const whatsappLink = whatsappLimpo
    ? `https://wa.me/55${whatsappLimpo}`
    : "";

  sheet.appendRow([
    new Date(),
    data.nome || "",
    data.whatsapp || "",
    whatsappLink,
    data.perfil || "",
    data.respostas?.barras || "",
    data.respostas?.trava || "",
    data.respostas?.progressao || "",
    data.respostas?.tempoTentando || "",
    data.utm_source || "",
    data.utm_campaign || "",
    "Pendente"
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
  }
  return sheet;
}
