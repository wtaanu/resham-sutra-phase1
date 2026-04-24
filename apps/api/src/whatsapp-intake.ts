import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { z } from "zod";
import { createPortalEnquiry, createPortalQuotationLineItems } from "./portal-actions.js";
import { listRecords, getRecord, type AirtableRecord } from "./airtable.js";
import { parseIncomingEnquiry } from "./enquiry-parser.js";
import { env } from "./config.js";
import { processPendingEnquiries } from "./intake-processor.js";
import { isSmtpConfigured, sendMail } from "./mailer.js";
import { productReferences } from "./phase1-data.js";

type EnquiryFields = {
  "Enquiry ID"?: string;
  "Logged Date Time"?: string;
  "Lead Name"?: string;
  Company?: string;
  Phone?: string;
  Email?: string;
  Address?: string;
  State?: string;
  City?: string;
  Pincode?: string;
  "Parser Status"?: string;
  "Linked Customer"?: string[];
  Quotations?: string[];
  "Requirement Summary"?: string;
  "Requested Asset"?: string;
  Quantity?: number | string;
  Notes?: string;
  "Receiver WhatsApp Number"?: string;
};

type ProductFields = {
  "Product Name"?: string;
  "Product Key"?: string;
  Model?: string;
  Narration?: string;
  "Source Sheet"?: string;
};

type CustomerFields = {
  "Customer Name"?: string;
  Company?: string;
  Phone?: string;
  WhatsApp?: string;
  Email?: string;
};

const whatsappPayloadSchema = z.object({
  rawMessage: z.string().trim().min(1),
  senderPhone: z.string().trim().optional().default(""),
  senderName: z.string().trim().optional().default(""),
  inboundWhatsappNumber: z.string().trim().optional().default("")
});

const metaWebhookSchema = z.object({
  entry: z
    .array(
      z.object({
        changes: z
          .array(
            z.object({
              value: z.object({
                metadata: z
                  .object({
                    display_phone_number: z.string().optional(),
                    phone_number_id: z.string().optional()
                  })
                  .optional(),
                contacts: z
                  .array(
                    z.object({
                      wa_id: z.string().optional(),
                      profile: z.object({ name: z.string().optional() }).optional()
                    })
                  )
                  .optional(),
                messages: z
                  .array(
                    z.object({
                      from: z.string().optional(),
                      type: z.string().optional(),
                      text: z.object({ body: z.string().optional() }).optional()
                    })
                  )
                  .optional()
              })
            })
          )
          .default([])
      })
    )
    .default([])
});

function normalizePhone(value: string) {
  const digits = value.replace(/\D/g, "");
  if (!digits) {
    return "";
  }

  return digits.length > 10 ? digits.slice(-10) : digits;
}

function extractAddressParts(message: string) {
  const lines = message
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);

  const pincodeMatch = message.match(/\b\d{6}\b/);
  const pincode = pincodeMatch?.[0] || "";

  const addressLine = lines.find((line) => /\d/.test(line) && line !== pincode) || "";
  const locationLine =
    lines.find((line) => /,/.test(line) && line !== addressLine) ||
    lines.find((line) => /\b(assam|maharashtra|karnataka|hyderabad|nagpur)\b/i.test(line)) ||
    "";

  const parts = locationLine
    .split(",")
    .map((part) => part.trim())
    .filter(Boolean);

  const city = parts[0] || "";
  const state = parts.slice(1).join(", ");

  return {
    address: addressLine,
    city,
    state,
    pincode
  };
}

const productAliasMap = productReferences.map((reference) => {
  const aliases = new Set<string>([
    reference.title,
    reference.model,
    reference.sku,
    reference.category
  ]);

  const lowerTitle = reference.title.toLowerCase();
  if (lowerTitle.includes("mulberry reeling")) {
    aliases.add("mulberry reeling machine");
    aliases.add("mulberry silk reeling machine");
  }

  if (lowerTitle.includes("buniyaad")) {
    aliases.add("buniyaad reeling machine");
  }

  if (lowerTitle.includes("twin charkha")) {
    aliases.add("twin charkha reeling machine");
  }

  if (lowerTitle.includes("sonalika")) {
    aliases.add("sonalika reeling machine");
  }

  return {
    reference,
    aliases: [...aliases].map((value) => value.toLowerCase())
  };
});

function tokenize(value: string) {
  return value
    .toLowerCase()
    .split(/[^a-z0-9]+/)
    .filter((token) => token.length > 2);
}

function scoreProductMatch(
  product: AirtableRecord<ProductFields>,
  haystack: string,
  parsedInterest: string | null
) {
  const searchable = [
    product.fields["Product Name"],
    product.fields["Product Key"],
    product.fields.Model,
    product.fields.Narration,
    product.fields["Source Sheet"]
  ]
    .filter(Boolean)
    .map((value) => String(value).toLowerCase());

  let score = 0;
  const interest = (parsedInterest || "").toLowerCase();
  const productName = String(product.fields["Product Name"] || "").toLowerCase();
  const productModel = String(product.fields.Model || "").toLowerCase();
  const sourceSheet = String(product.fields["Source Sheet"] || "").toLowerCase();

  if (sourceSheet.includes("inventory")) {
    score -= 80;
  }

  if (productName.includes("inventory") || productName.includes("sheet")) {
    score -= 60;
  }

  for (const text of searchable) {
    if (haystack.includes(text)) {
      score += 100;
      continue;
    }

    const tokens = text.split(/[^a-z0-9]+/).filter((token) => token.length > 2);
    const matchedTokens = tokens.filter((token) => haystack.includes(token));
    score += matchedTokens.length * 10;
  }

  if (interest) {
    const exactInterestMatch =
      productName.includes(interest) ||
      productModel.includes(interest) ||
      searchable.some((text) => interest.includes(text));

    if (exactInterestMatch) {
      score += 180;
    }

    const interestTokens = tokenize(interest);
    const productTokens = new Set(searchable.flatMap((text) => tokenize(text)));
    const tokenOverlap = interestTokens.filter((token) => productTokens.has(token)).length;
    score += tokenOverlap * 25;
  }

  for (const entry of productAliasMap) {
    const matchesAlias = searchable.some((text) =>
      entry.aliases.some((alias) => text.includes(alias) || alias.includes(text))
    );

    if (!matchesAlias) {
      continue;
    }

    const matchedAlias = entry.aliases.find((alias) => haystack.includes(alias) || interest.includes(alias));
    if (matchedAlias) {
      score += 220;
    }
  }

  if (haystack.includes("mulberry") && !searchable.some((text) => text.includes("mulberry"))) {
    score -= 40;
  }

  if (haystack.includes("buniyaad") && !searchable.some((text) => text.includes("buniyaad"))) {
    score -= 40;
  }

  if (haystack.includes("reeling") && !searchable.some((text) => text.includes("reeling"))) {
    score -= 20;
  }

  return score;
}

async function findProductsForMessage(rawMessage: string, parsedInterest: string | null) {
  const products = await listRecords<ProductFields>(env.AIRTABLE_PRODUCTS_TABLE, {
    maxRecords: 500
  });

  const haystack = `${rawMessage}\n${parsedInterest || ""}`.toLowerCase();

  return products
    .map((product) => ({
      product,
      score: scoreProductMatch(product, haystack, parsedInterest)
    }))
    .filter((entry) => entry.score > 0)
    .sort((left, right) => right.score - left.score)
    .slice(0, 3)
    .map((entry) => entry.product);
}

async function createSalesNotificationArtifact(input: {
  enquiry: AirtableRecord<EnquiryFields>;
  rawMessage: string;
  matchedProducts: AirtableRecord<ProductFields>[];
  inboundWhatsappNumber: string;
}) {
  const notificationsDir = path.resolve(env.DOCUMENT_STORAGE_DIR, "notifications");
  await mkdir(notificationsDir, { recursive: true });

  const fileName = `${input.enquiry.id}-whatsapp-enquiry.json`;
  const filePath = path.join(notificationsDir, fileName);
  const quotationIds = input.enquiry.fields.Quotations || [];

  await writeFile(
    filePath,
    JSON.stringify(
      {
        createdAt: new Date().toISOString(),
        salesTeamEmail: env.SALES_TEAM_NOTIFICATION_EMAIL,
        inboundWhatsappNumber: input.inboundWhatsappNumber,
        enquiryRecordId: input.enquiry.id,
        enquiryId: input.enquiry.fields["Enquiry ID"] || "",
        linkedCustomerId: input.enquiry.fields["Linked Customer"]?.[0] || "",
        quotationIds,
        matchedProducts: input.matchedProducts.map((product) => ({
          id: product.id,
          name: product.fields["Product Name"] || product.fields["Product Key"] || product.id
        })),
        rawMessage: input.rawMessage
      },
      null,
      2
    ),
    "utf8"
  );

  return {
    filePath,
    fileUrl: `${env.PUBLIC_API_BASE_URL}/documents/notifications/${fileName}`
  };
}

async function sendSalesTeamNotificationEmail(input: {
  enquiry: AirtableRecord<EnquiryFields>;
  rawMessage: string;
  matchedProducts: AirtableRecord<ProductFields>[];
  inboundWhatsappNumber: string;
  linkedCustomer: AirtableRecord<CustomerFields> | null;
  notificationFileUrl: string;
}) {
  if (!env.SALES_TEAM_NOTIFICATION_EMAIL || !isSmtpConfigured()) {
    return null;
  }

  const matchedProductsText = input.matchedProducts.length
    ? input.matchedProducts
        .map(
          (product) =>
            `- ${product.fields["Product Name"] || product.fields["Product Key"] || product.id}`
        )
        .join("\n")
    : "No products auto-matched yet";

  const lines = [
    "A new WhatsApp enquiry was processed by the automation.",
    "",
    `Enquiry ID: ${input.enquiry.fields["Enquiry ID"] || input.enquiry.id}`,
    `Enquiry Record: ${input.enquiry.id}`,
    `Customer: ${input.linkedCustomer?.fields["Customer Name"] || "Not linked yet"}`,
    `Customer Record: ${input.enquiry.fields["Linked Customer"]?.[0] || ""}`,
    `Quotation Record: ${input.enquiry.fields.Quotations?.[0] || ""}`,
    `Inbound Business Number: ${input.inboundWhatsappNumber || "Unknown"}`,
    "",
    "Matched products:",
    matchedProductsText,
    "",
    "Raw message:",
    input.rawMessage,
    "",
    `Notification artifact: ${input.notificationFileUrl}`
  ];

  await sendMail({
    to: env.SALES_TEAM_NOTIFICATION_EMAIL,
    subject: `WhatsApp enquiry processed: ${input.enquiry.fields["Enquiry ID"] || input.enquiry.id}`,
    text: lines.join("\n")
  });

  return {
    recipient: env.SALES_TEAM_NOTIFICATION_EMAIL
  };
}

export async function processWhatsAppEnquiry(payload: unknown) {
  const input = whatsappPayloadSchema.parse(payload);
  const parsed = parseIncomingEnquiry(input.rawMessage);
  const contactPhone = parsed.phone || normalizePhone(input.senderPhone);
  const addressParts = extractAddressParts(input.rawMessage);

  const created = await createPortalEnquiry({
    linkedCustomerId: "",
    leadName: parsed.leadName || input.senderName || "WhatsApp Lead",
    company: parsed.company || "",
    phone: contactPhone,
    email: "",
    address: addressParts.address,
    state: addressParts.state || parsed.cityOrState || "",
    city: addressParts.city,
    pincode: addressParts.pincode,
    destinationAddress: addressParts.address,
    destinationState: addressParts.state || parsed.cityOrState || "",
    destinationCity: addressParts.city,
    destinationPincode: "",
    requirementSummary: parsed.productInterest || input.rawMessage,
    requestedAsset: parsed.requestedAsset || "Details",
    potentialProduct: parsed.productInterest || "",
    receiverWhatsappNumber: input.inboundWhatsappNumber
  });

  await processPendingEnquiries();

  let enquiry = await getRecord<EnquiryFields>(env.AIRTABLE_ENQUIRIES_TABLE, created.enquiryRecordId);
  const quotationId = enquiry.fields.Quotations?.[0] || "";
  const matchedProducts = await findProductsForMessage(input.rawMessage, parsed.productInterest);

  if (quotationId && matchedProducts.length) {
    await createPortalQuotationLineItems({
      quotationId,
      items: matchedProducts.map((product) => ({
        productId: product.id,
        qty: 1
      }))
    });

    await processPendingEnquiries();
    enquiry = await getRecord<EnquiryFields>(env.AIRTABLE_ENQUIRIES_TABLE, created.enquiryRecordId);
  }

  const notification = await createSalesNotificationArtifact({
    enquiry,
    rawMessage: input.rawMessage,
    matchedProducts,
    inboundWhatsappNumber: input.inboundWhatsappNumber
  });

  const linkedCustomerId = enquiry.fields["Linked Customer"]?.[0] || created.linkedCustomerId || "";
  const linkedCustomer = linkedCustomerId
    ? await getRecord<CustomerFields>(env.AIRTABLE_CUSTOMERS_TABLE, linkedCustomerId)
    : null;
  const salesEmail = await sendSalesTeamNotificationEmail({
    enquiry,
    rawMessage: input.rawMessage,
    matchedProducts,
    inboundWhatsappNumber: input.inboundWhatsappNumber,
    linkedCustomer,
    notificationFileUrl: notification.fileUrl
  });

  return {
    enquiryRecordId: enquiry.id,
    enquiryId: enquiry.fields["Enquiry ID"] || "",
    parserStatus: enquiry.fields["Parser Status"] || "",
    linkedCustomerId,
    customerName: linkedCustomer?.fields["Customer Name"] || "",
    quotationId: enquiry.fields.Quotations?.[0] || "",
    matchedProducts: matchedProducts.map((product) => ({
      id: product.id,
      name: product.fields["Product Name"] || product.fields["Product Key"] || product.id
    })),
    salesNotification: notification,
    salesEmail
  };
}

export function extractWhatsAppWebhookMessages(payload: unknown) {
  const parsed = metaWebhookSchema.parse(payload);
  const messages: Array<{
    rawMessage: string;
    senderPhone: string;
    senderName: string;
    inboundWhatsappNumber: string;
    phoneNumberId: string;
  }> = [];

  for (const entry of parsed.entry) {
    for (const change of entry.changes) {
      const displayNumber = change.value.metadata?.display_phone_number || "";
      const phoneNumberId = change.value.metadata?.phone_number_id || "";
      const contact = change.value.contacts?.[0];

      for (const message of change.value.messages || []) {
        if ((message.type || "text") !== "text") {
          continue;
        }

        const body = message.text?.body?.trim() || "";
        if (!body) {
          continue;
        }

        messages.push({
          rawMessage: body,
          senderPhone: message.from || contact?.wa_id || "",
          senderName: contact?.profile?.name || "",
          inboundWhatsappNumber: displayNumber,
          phoneNumberId
        });
      }
    }
  }

  return messages;
}
