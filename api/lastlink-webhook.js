/**
 * Recebe o webhook da Lastlink, envia PixCreated/Purchase pro Meta via
 * Conversions API (server-side) e atualiza o status do lead na planilha
 * (Google Sheets, via o mesmo Apps Script já usado pelo quiz).
 *
 * Payload da Lastlink (support.lastlink.com, "Documentação de Webhook"):
 *   { Id, IsTest, Event, CreatedAt, Data: { Buyer, Purchase, Products, Utm, ... } }
 *   - Buyer:    { Email, Name, PhoneNumber ("+55DDD9XXXXXXXX"), ... }
 *   - Purchase: { PaymentId, Price: { Value }, Payment: { PaymentMethod }, ... }
 *   - Utm:      { UtmSource, UtmMedium, UtmCampaign, UtmContent, UtmTerm }
 *   Campos sem valor não são enviados.
 *
 * A Lastlink não documenta assinatura nem header de segurança. Por isso a
 * URL cadastrada no painel leva um token: /api/lastlink-webhook?token=XXXX
 * e ele é comparado com a variável LASTLINK_WEBHOOK_TOKEN.
 *
 * Variáveis de ambiente (configurar na Vercel, nunca no código):
 *   - META_ACCESS_TOKEN      (token de sistema do Meta Business Manager)
 *   - LASTLINK_WEBHOOK_TOKEN (o mesmo token colocado na URL do webhook)
 *   - LEADS_WEBHOOK_URL      (a mesma URL do Apps Script já usada pelo quiz)
 *
 * Se alguma não estiver configurada, o endpoint continua respondendo 200,
 * só pula a etapa que depende dela e loga um aviso.
 *
 * Eventos de teste (IsTest: true, botão "Testar" do painel) não vão pro
 * Meta nem pra planilha: o endpoint só devolve o que faria com eles.
 */

const crypto = require("crypto");

const META_PIXEL_ID = "1539676747360822";
const VALOR_PADRAO = 19.90;

// Evento da Lastlink -> evento do Meta e evento que o Apps Script já entende.
// O Apps Script usa os nomes do GGCheckout (pix.paid = "Pago" e grava "Pago Em"),
// então reaproveitamos esses nomes para não precisar republicar o script.
const MAPA_EVENTOS = {
  Purchase_Request_Confirmed: { planilha: "pix.generated", meta: "PixCreated", soPix: true },
  Purchase_Order_Confirmed:   { planilha: "pix.paid",      meta: "Purchase" },
  Purchase_Request_Expired:   { planilha: "pix.expired" },
  Purchase_Request_Canceled:  { planilha: "pix.failed" },
  Payment_Refund:             { planilha: "pix.refunded" },
  Payment_Chargeback:         { planilha: "pix.refunded" },
};

module.exports = async function handler(req, res) {
  if (req.method !== "POST") {
    res.status(405).json({ error: "method_not_allowed" });
    return;
  }

  const tokenConfigurado = process.env.LASTLINK_WEBHOOK_TOKEN;
  if (tokenConfigurado) {
    const recebido = String((req.query && req.query.token) || "");
    if (!iguaisSemVazarTempo(recebido, tokenConfigurado)) {
      res.status(401).json({ error: "invalid_token" });
      return;
    }
  } else {
    console.warn("LASTLINK_WEBHOOK_TOKEN não configurado — pulando verificação do token.");
  }

  const body = lerCorpo(req.body);
  const evento = body.Event || "";
  const data = body.Data || {};
  const buyer = data.Buyer || {};
  const purchase = data.Purchase || {};
  const utm = data.Utm || {};

  const mapa = MAPA_EVENTOS[evento];
  if (!mapa) {
    res.status(200).json({ received: true, ignorado: evento || "sem_evento" });
    return;
  }

  const paymentId = purchase.PaymentId || body.Id || "";
  const metodo = (purchase.Payment && purchase.Payment.PaymentMethod) || "";
  const valor = Number(purchase.Price && purchase.Price.Value) || VALOR_PADRAO;
  const telefone = normalizarTelefone(buyer.PhoneNumber);
  const email = (buyer.Email || "").trim().toLowerCase();

  // "Fatura criada" só interessa no Pix (pagamento pendente). No cartão a
  // aprovação chega logo em seguida como Purchase_Order_Confirmed.
  if (mapa.soPix && metodo !== "pix") {
    res.status(200).json({ received: true, ignorado: `${evento}_${metodo || "sem_metodo"}` });
    return;
  }
  const eventoMeta = mapa.meta;

  if (body.IsTest) {
    res.status(200).json({
      received: true,
      teste: true,
      faria: { meta: eventoMeta || "nada", planilha: mapa.planilha, paymentId, valor, metodo },
    });
    return;
  }

  let metaResultado = "pulado";
  if (eventoMeta && paymentId) {
    metaResultado = await enviarEventoMeta({
      eventName: eventoMeta,
      eventId: `${evento}_${paymentId}`, // determinístico: reenvios da Lastlink não duplicam no Meta
      value: valor,
      telefone,
      email,
      sourceUrl: purchase.OriginUrl || "",
      ip: data.DeviceInfo && data.DeviceInfo.ip,
      userAgent: data.DeviceInfo && data.DeviceInfo.UserAgent,
    });
  }

  await atualizarStatusNaPlanilha({
    whatsapp: telefone,
    event: mapa.planilha,
    paymentId,
    paymentStatus: evento,
    amount: valor,
    utm_source: utm.UtmSource || "",
    utm_campaign: utm.UtmCampaign || "",
  });

  res.status(200).json({ received: true, meta: metaResultado });
};

function lerCorpo(corpo) {
  if (!corpo) return {};
  if (typeof corpo === "string") {
    try { return JSON.parse(corpo); } catch (e) { return {}; }
  }
  return corpo;
}

function iguaisSemVazarTempo(a, b) {
  const ba = Buffer.from(a);
  const bb = Buffer.from(b);
  return ba.length === bb.length && crypto.timingSafeEqual(ba, bb);
}

function normalizarTelefone(phone) {
  return (phone || "").replace(/\D/g, "");
}

function sha256(valor) {
  return crypto.createHash("sha256").update(valor).digest("hex");
}

async function enviarEventoMeta({ eventName, eventId, value, telefone, email, sourceUrl, ip, userAgent }) {
  const token = process.env.META_ACCESS_TOKEN;
  if (!token) {
    console.warn(`META_ACCESS_TOKEN não configurado — evento ${eventName} (${eventId}) não enviado ao Meta.`);
    return "sem_token";
  }

  const userData = {};
  if (telefone) userData.ph = [sha256(telefone)];
  if (email) userData.em = [sha256(email)];
  if (ip) userData.client_ip_address = ip;
  if (userAgent) userData.client_user_agent = userAgent;

  const evento = {
    event_name: eventName,
    event_time: Math.floor(Date.now() / 1000),
    event_id: eventId,
    action_source: "website",
    user_data: userData,
  };
  if (sourceUrl) evento.event_source_url = sourceUrl;
  if (eventName === "Purchase") {
    evento.custom_data = { value, currency: "BRL" };
  }

  const url = `https://graph.facebook.com/v19.0/${META_PIXEL_ID}/events?access_token=${encodeURIComponent(token)}`;
  try {
    const resp = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ data: [evento] }),
    });
    if (!resp.ok) {
      const texto = await resp.text();
      console.error(`Meta CAPI respondeu ${resp.status} para ${eventName} (${eventId}): ${texto}`);
      return "erro_meta";
    }
    return "enviado";
  } catch (e) {
    console.error(`Falha ao chamar Meta CAPI para ${eventName} (${eventId}): ${e.message}`);
    return "erro_rede";
  }
}

async function atualizarStatusNaPlanilha(dados) {
  const url = process.env.LEADS_WEBHOOK_URL;
  if (!url) {
    console.warn("LEADS_WEBHOOK_URL não configurado — status de pagamento não gravado na planilha.");
    return;
  }
  try {
    await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify({ tipo: "atualizar_pagamento", ...dados }),
    });
  } catch (e) {
    console.error(`Falha ao atualizar planilha: ${e.message}`);
  }
}
