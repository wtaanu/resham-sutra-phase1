import { env } from "./config.js";

type SendQuotationWhatsAppInput = {
  to: string;
  documentUrl: string;
  filename: string;
  caption: string;
};

function normalizeWhatsappRecipient(value: string) {
  const digits = String(value || "").replace(/\D/g, "");
  if (!digits) {
    return "";
  }

  if (digits.length === 10) {
    return `91${digits}`;
  }

  if (digits.length === 12 && digits.startsWith("91")) {
    return digits;
  }

  return digits;
}

export function isWhatsAppOutboundConfigured() {
  return Boolean(env.META_SYSTEM_USER_ACCESS_TOKEN && env.META_WHATSAPP_PHONE_NUMBER_ID);
}

export async function sendQuotationDocumentOnWhatsApp(input: SendQuotationWhatsAppInput) {
  if (!isWhatsAppOutboundConfigured()) {
    throw new Error("WhatsApp outbound is not configured yet.");
  }

  const recipient = normalizeWhatsappRecipient(input.to);
  if (!recipient) {
    throw new Error("A valid customer WhatsApp number is required.");
  }

  const response = await fetch(
    `https://graph.facebook.com/v22.0/${env.META_WHATSAPP_PHONE_NUMBER_ID}/messages`,
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${env.META_SYSTEM_USER_ACCESS_TOKEN}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        messaging_product: "whatsapp",
        to: recipient,
        type: "document",
        document: {
          link: input.documentUrl,
          filename: input.filename,
          caption: input.caption
        }
      })
    }
  );

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`WhatsApp send failed: ${errorText}`);
  }

  return response.json();
}
