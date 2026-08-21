/**
 * Recebe os leads do quiz (nome, WhatsApp, respostas, perfil, atribuição
 * de campanha) e grava em tempo real numa aba da planilha, com status de
 * pagamento manual para você acompanhar quem não gerou/pagou o Pix.
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
  "Landing Page", "Referrer", "Primeira Visita", "Lead Event ID", "Status Pix"
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
