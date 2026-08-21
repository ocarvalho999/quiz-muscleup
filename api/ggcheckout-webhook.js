/**
 * Recebe o webhook do GGCheckout, envia PixCreated/Purchase pro Meta
 * via Conversions API (server-side) e atualiza o status do lead na
 * planilha (Google Sheets, via o mesmo Apps Script já usado pelo quiz).
 *
 * Payload do GGCheckout confirmado na documentação oficial
 * (docs.ggcheckout.com/help/vendedor/integracoes/estrutura-do-payload-do-webhook):
 *   - Headers: "Authorization: Bearer <secret>" e "x-secret: <secret>"
 *     (só enviados se um Secret foi configurado no painel do GGCheckout)
 *   - event: "pix.generated" | "pix.paid" | "pix.expired" | "pix.failed" | "pix.refunded"
 *   - payment: { id, method, paymentMethod, gateway, status, amount, pixCode }
 *   - customer: { name, email, document, phone, ip }
 *   - product / products, utm_source/medium/campaign/content/term
 *
 * Variáveis de ambiente necessárias (configurar na Vercel, nunca no código):
 *   - META_ACCESS_TOKEN         (token de sistema do Meta Business Manager)
 *   - GGCHECKOUT_WEBHOOK_SECRET (o mesmo Secret configurado no painel do GGCheckout)
 *   - LEADS_WEBHOOK_URL         (a mesma URL do Apps Script já usada pelo quiz)
 *
 * Se META_ACCESS_TOKEN ou GGCHECKOUT_WEBHOOK_SECRET não estiverem
 * configurados, o endpoint continua funcionando (não derruba o
 * webhook), só pula a etapa que depende da variável ausente e loga
 * um aviso — nunca expõe o valor esperado nem o payload em log.
 */

const META_PIXEL_ID = "1539676747360822";

const EVENTO_META = {
  "pix.generated": "PixCreated",
  "pix.paid": "Purchase",
};

module.exports = async function handler(req, res) {
  if (req.method !== "POST") {
    res.status(405).json({ error: "method_not_allowed" });
    return;
  }

  const secretConfigurado = process.env.GGCHECKOUT_WEBHOOK_SECRET;
  if (secretConfigurado) {
    const recebido = req.headers["x-secret"] || "";
    if (recebido !== secretConfigurado) {
      res.status(401).json({ error: "invalid_secret" });
      return;
    }
  } else {
    console.warn("GGCHECKOUT_WEBHOOK_SECRET não configurado — pulando verificação de assinatura.");
  }

  const body = req.body || {};
  const evento = body.event || "";
  const payment = body.payment || {};
  const customer = body.customer || {};

  const eventoMeta = EVENTO_META[evento];
  const paymentId = payment.id || "";
  const telefoneNormalizado = normalizarTelefone(customer.phone);

  let metaResultado = "pulado";
  if (eventoMeta && paymentId) {
    const eventId = `${evento}_${paymentId}`; // determinístico: retries do GGCheckout não duplicam no Meta
    metaResultado = await enviarEventoMeta({
      eventName: eventoMeta,
      eventId,
      value: eventoMeta === "Purchase" ? Number(payment.amount) || 19.90 : undefined,
      telefoneNormalizado,
      sourceUrl: body.landing_page || "",
    });
  }

  await atualizarStatusNaPlanilha({
    whatsapp: telefoneNormalizado,
    event: evento,
    paymentId,
    paymentStatus: payment.status || "",
    amount: payment.amount || "",
    utm_source: body.utm_source || "",
    utm_campaign: body.utm_campaign || "",
  });

  res.status(200).json({ received: true, meta: metaResultado });
};

function normalizarTelefone(phone) {
  return (phone || "").replace(/\D/g, "");
}

async function enviarEventoMeta({ eventName, eventId, value, telefoneNormalizado, sourceUrl }) {
  const token = process.env.META_ACCESS_TOKEN;
  if (!token) {
    console.warn(`META_ACCESS_TOKEN não configurado — evento ${eventName} (${eventId}) não enviado ao Meta.`);
    return "sem_token";
  }

  const userData = {};
  if (telefoneNormalizado) {
    userData.ph = [sha256(telefoneNormalizado)];
  }

  const evento = {
    event_name: eventName,
    event_time: Math.floor(Date.now() / 1000),
    event_id: eventId,
    action_source: "website",
    user_data: userData,
  };
  if (sourceUrl) evento.event_source_url = sourceUrl;
  if (eventName === "Purchase") {
    evento.custom_data = { value: value || 19.90, currency: "BRL" };
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

function sha256(valor) {
  const crypto = require("crypto");
  return crypto.createHash("sha256").update(valor).digest("hex");
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
