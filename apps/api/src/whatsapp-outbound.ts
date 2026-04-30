import { readFile } from "node:fs/promises";
import { env } from "./config.js";

type SendQuotationWhatsAppInput = {
  to: string;
  documentUrl: string;
  filename: string;
  caption: string;
  customerName: string;
  quotationNumber: string;
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

function isQuotationTemplateConfigured() {
  return Boolean(String(env.META_WHATSAPP_QUOTATION_TEMPLATE_NAME || "").trim());
}

function buildTemplateBodyParameters(input: SendQuotationWhatsAppInput) {
  const values: Record<string, string> = {
    customerName: input.customerName || "Customer",
    quotationNumber: input.quotationNumber || input.filename,
    documentUrl: input.documentUrl,
    filename: input.filename
  };

  return env.META_WHATSAPP_QUOTATION_TEMPLATE_BODY_PARAMS.split(",")
    .map((key) => key.trim())
    .filter(Boolean)
    .map((key) => ({
      type: "text",
      text: values[key] || ""
    }))
    .filter((parameter) => parameter.text);
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

async function sendQuotationTemplateOnWhatsApp(input: SendQuotationWhatsAppInput) {
  if (!isWhatsAppOutboundConfigured()) {
    throw new Error("WhatsApp outbound is not configured yet.");
  }

  const recipient = normalizeWhatsappRecipient(input.to);
  if (!recipient) {
    throw new Error("A valid customer WhatsApp number is required.");
  }

  const templateName = String(env.META_WHATSAPP_QUOTATION_TEMPLATE_NAME || "").trim();
  if (!templateName) {
    throw new Error("WhatsApp quotation template is not configured.");
  }

  console.info("[whatsapp-send] attempting approved quotation template send", {
    recipient,
    phoneNumberId: env.META_WHATSAPP_PHONE_NUMBER_ID,
    templateName,
    templateLanguage: env.META_WHATSAPP_QUOTATION_TEMPLATE_LANGUAGE,
    filename: input.filename,
    documentUrl: input.documentUrl
  });

  const document =
    input.localFilePath
      ? {
          id: await uploadDocumentMediaToWhatsApp({
            localFilePath: input.localFilePath,
            filename: input.filename,
            contentType: input.contentType
          }),
          filename: input.filename
        }
      : {
          link: input.documentUrl,
          filename: input.filename
        };

  const components: Array<Record<string, unknown>> = [
    {
      type: "header",
      parameters: [
        {
          type: "document",
          document
        }
      ]
    }
  ];
  const bodyParameters = buildTemplateBodyParameters(input);
  if (bodyParameters.length) {
    components.push({
      type: "body",
      parameters: bodyParameters
    });
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
        type: "template",
        template: {
          name: templateName,
          language: {
            code: env.META_WHATSAPP_QUOTATION_TEMPLATE_LANGUAGE
          },
          components
        }
      })
    }
  );

  if (!response.ok) {
    const errorText = await response.text();
    console.error("[whatsapp-send] Meta template API returned error", {
      recipient,
      templateName,
      status: response.status,
      body: errorText
    });
    throw new Error(`WhatsApp template send failed: ${errorText}`);
  }

  const payload = await response.json();
  console.info("[whatsapp-send] Meta API accepted approved template send", {
    recipient,
    templateName,
    payload
  });
  return payload;
}

export async function sendQuotationDocumentOnWhatsApp(input: SendQuotationWhatsAppInput) {
  if (!isWhatsAppOutboundConfigured()) {
    throw new Error("WhatsApp outbound is not configured yet.");
  }

  if (isQuotationTemplateConfigured()) {
    return sendQuotationTemplateOnWhatsApp(input);
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
