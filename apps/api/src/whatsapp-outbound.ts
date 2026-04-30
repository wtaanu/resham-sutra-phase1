import { readFile } from "node:fs/promises";
import { env } from "./config.js";

type SendQuotationWhatsAppInput = {
  to: string;
  documentUrl: string;
  filename: string;
  caption: string;
  localFilePath?: string;
  contentType?: string;
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

async function uploadDocumentMediaToWhatsApp(input: {
  localFilePath: string;
  filename: string;
  contentType?: string;
}) {
  const bytes = await readFile(input.localFilePath);
  const form = new FormData();
  form.append("messaging_product", "whatsapp");
  form.append(
    "file",
    new Blob([bytes], { type: input.contentType || "application/pdf" }),
    input.filename
  );
  form.append("type", input.contentType || "application/pdf");

  const response = await fetch(
    `https://graph.facebook.com/v22.0/${env.META_WHATSAPP_PHONE_NUMBER_ID}/media`,
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${env.META_SYSTEM_USER_ACCESS_TOKEN}`
      },
      body: form
    }
  );

  if (!response.ok) {
    const errorText = await response.text();
    console.error("[whatsapp-send] Meta media upload failed", {
      filename: input.filename,
      status: response.status,
      body: errorText
    });
    throw new Error(`WhatsApp media upload failed: ${errorText}`);
  }

  const payload = await response.json();
  const mediaId = String(payload?.id || "").trim();
  if (!mediaId) {
    throw new Error("WhatsApp media upload succeeded but no media id was returned.");
  }

  console.info("[whatsapp-send] Meta media upload succeeded", {
    filename: input.filename,
    mediaId
  });

  return mediaId;
}

export async function sendQuotationDocumentOnWhatsApp(input: SendQuotationWhatsAppInput) {
  if (!isWhatsAppOutboundConfigured()) {
    throw new Error("WhatsApp outbound is not configured yet.");
  }

  const recipient = normalizeWhatsappRecipient(input.to);
  if (!recipient) {
    throw new Error("A valid customer WhatsApp number is required.");
  }

  console.info("[whatsapp-send] attempting outbound document send", {
    recipient,
    phoneNumberId: env.META_WHATSAPP_PHONE_NUMBER_ID,
    filename: input.filename,
    documentUrl: input.documentUrl
  });

  let documentPayload: Record<string, string> | null = null;
  if (input.localFilePath) {
    const mediaId = await uploadDocumentMediaToWhatsApp({
      localFilePath: input.localFilePath,
      filename: input.filename,
      contentType: input.contentType
    });
    documentPayload = {
      id: mediaId,
      filename: input.filename,
      caption: input.caption
    };
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
        document:
          documentPayload ||
          {
            link: input.documentUrl,
            filename: input.filename,
            caption: input.caption
          }
      })
    }
  );

  if (!response.ok) {
    const errorText = await response.text();
    console.error("[whatsapp-send] Meta API returned error", {
      recipient,
      status: response.status,
      body: errorText
    });
    throw new Error(`WhatsApp send failed: ${errorText}`);
  }

  const payload = await response.json();
  console.info("[whatsapp-send] Meta API accepted outbound document send", {
    recipient,
    payload
  });
  return payload;
}
