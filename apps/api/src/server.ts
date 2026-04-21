import cors from "cors";
import express from "express";
import path from "node:path";
import { env } from "./config.js";
import { checkDatabaseConnection } from "./db.js";
import { createDraftDocument, createPdfDocument } from "./documents.js";
import {
  createCustomerForEnquiry,
  generateFinalPdfForQuotation,
  regenerateQuotationDraft,
  sendQuotationEmail,
  sendQuotationWhatsApp,
  processPendingEnquiries
} from "./intake-processor.js";
import { getOperationsSnapshot } from "./operations.js";
import { createPortalEnquiry, createPortalQuotationLineItems } from "./portal-actions.js";
import {
  getProductDocumentsByIds,
  sendProductDocumentsSchema,
  uploadProductDocuments,
  uploadProductDocumentsSchema
} from "./product-documents.js";
import { getRecord } from "./airtable.js";
import { getProjectSnapshot, getProjectSummary } from "./project.js";
import { extractWhatsAppWebhookMessages, processWhatsAppEnquiry } from "./whatsapp-intake.js";
import { isSmtpConfigured, resolveDocumentAttachment, sendMail } from "./mailer.js";

const app = express();

function logRouteError(route: string, error: unknown) {
  console.error(`[api] ${route}`, error);
}

app.use(cors());
app.use(express.json());
app.use(
  "/documents",
  express.static(path.resolve(env.DOCUMENT_STORAGE_DIR), {
    extensions: ["html", "json"]
  })
);

app.get("/health", async (_request, response) => {
  try {
    await checkDatabaseConnection();

    response.json({
      status: "ok",
      service: "api"
    });
  } catch (error) {
    logRouteError("GET /health", error);
    response.status(500).json({
      status: "error",
      service: "api",
      message: error instanceof Error ? error.message : "Unknown error"
    });
  }
});

app.get("/api/summary", (_request, response) => {
  response.json(getProjectSummary());
});

app.get("/api/project", (_request, response) => {
  response.json(getProjectSnapshot());
});

app.get("/api/whatsapp/webhook", (request, response) => {
  const mode = String(request.query["hub.mode"] || "");
  const token = String(request.query["hub.verify_token"] || "");
  const challenge = String(request.query["hub.challenge"] || "");

  if (
    mode === "subscribe" &&
    env.META_WHATSAPP_VERIFY_TOKEN &&
    token === env.META_WHATSAPP_VERIFY_TOKEN
  ) {
    response.status(200).send(challenge);
    return;
  }

  response.status(403).json({
    status: "error",
    message: "WhatsApp webhook verification failed."
  });
});

app.get("/api/operations", async (_request, response) => {
  try {
    const snapshot = await getOperationsSnapshot();
    response.json(snapshot);
  } catch (error) {
    logRouteError("GET /api/operations", error);
    response.status(500).json({
      status: "error",
      message: error instanceof Error ? error.message : "Failed to load operations snapshot"
    });
  }
});

app.post("/api/documents/draft", async (request, response) => {
  try {
    const result = await createDraftDocument(request.body);
    response.status(201).json(result);
  } catch (error) {
    logRouteError("POST /api/documents/draft", error);
    response.status(400).json({
      status: "error",
      message: error instanceof Error ? error.message : "Failed to create draft document"
    });
  }
});

app.post("/api/documents/pdf", async (request, response) => {
  try {
    const result = await createPdfDocument(request.body);
    response.status(201).json(result);
  } catch (error) {
    logRouteError("POST /api/documents/pdf", error);
    response.status(400).json({
      status: "error",
      message: error instanceof Error ? error.message : "Failed to create PDF document"
    });
  }
});

app.post("/api/automation/process-enquiries", async (_request, response) => {
  try {
    const result = await processPendingEnquiries();
    response.status(200).json(result);
  } catch (error) {
    logRouteError("POST /api/automation/process-enquiries", error);
    response.status(500).json({
      status: "error",
      message: error instanceof Error ? error.message : "Failed to process enquiries"
    });
  }
});

app.post("/api/portal/enquiries", async (request, response) => {
  try {
    const result = await createPortalEnquiry(request.body);
    response.status(201).json({
      status: "ok",
      ...result
    });
  } catch (error) {
    logRouteError("POST /api/portal/enquiries", error);
    response.status(400).json({
      status: "error",
      message: error instanceof Error ? error.message : "Failed to create enquiry"
    });
  }
});

app.post("/api/portal/quotation-line-items", async (request, response) => {
  try {
    const result = await createPortalQuotationLineItems(request.body);
    response.status(201).json({
      status: "ok",
      ...result
    });
  } catch (error) {
    logRouteError("POST /api/portal/quotation-line-items", error);
    response.status(400).json({
      status: "error",
      message: error instanceof Error ? error.message : "Failed to create quotation line items"
    });
  }
});

app.post("/api/whatsapp/enquiries", async (request, response) => {
  try {
    const result = await processWhatsAppEnquiry(request.body);
    response.status(201).json({
      status: "ok",
      ...result
    });
  } catch (error) {
    logRouteError("POST /api/whatsapp/enquiries", error);
    response.status(400).json({
      status: "error",
      message: error instanceof Error ? error.message : "Failed to process WhatsApp enquiry"
    });
  }
});

app.post("/api/whatsapp/webhook", async (request, response) => {
  try {
    const messages = extractWhatsAppWebhookMessages(request.body);

    const results = [];
    for (const message of messages) {
      results.push(await processWhatsAppEnquiry(message));
    }

    response.status(200).json({
      status: "ok",
      processedCount: results.length,
      results
    });
  } catch (error) {
    logRouteError("POST /api/whatsapp/webhook", error);
    response.status(400).json({
      status: "error",
      message: error instanceof Error ? error.message : "Failed to process WhatsApp webhook"
    });
  }
});

app.post("/api/portal/products/:id/documents", async (request, response) => {
  try {
    const payload = uploadProductDocumentsSchema.parse(request.body);
    const documents = await uploadProductDocuments(request.params.id, payload);
    response.status(201).json({
      status: "ok",
      uploadedCount: documents.length,
      documents
    });
  } catch (error) {
    logRouteError("POST /api/portal/products/:id/documents", error);
    response.status(400).json({
      status: "error",
      message: error instanceof Error ? error.message : "Failed to upload product documents"
    });
  }
});

app.post("/api/actions/enquiries/:id/create-customer", async (request, response) => {
  try {
    const result = await createCustomerForEnquiry(request.params.id);
    response.status(200).json({
      status: "ok",
      enquiryId: result.enquiry.id,
      customerId: result.customer.id,
      quotationId: result.quotation.id
    });
  } catch (error) {
    logRouteError("POST /api/actions/enquiries/:id/create-customer", error);
    const message =
      error instanceof Error ? error.message : "Failed to create customer from enquiry";
    const friendlyMessage = message.includes("INVALID_MULTIPLE_CHOICE_OPTIONS")
      ? "Airtable is missing one or more required single-select options. Please add these enquiry statuses in Airtable before retrying: New, Parsed, Ready for Draft, Ready for Review, Draft Sent, Accepted, Rejected."
      : message;

    response.status(400).json({
      status: "error",
      message: friendlyMessage
    });
  }
});

app.post("/api/actions/quotations/:id/generate-final-pdf", async (request, response) => {
  try {
    const result = await generateFinalPdfForQuotation(request.params.id);
    response.status(200).json({
      status: "ok",
      quotationId: result.quotation.id,
      finalPdfUrl: result.pdf.fileUrl,
      previewUrl: result.pdf.previewUrl
    });
  } catch (error) {
    logRouteError("POST /api/actions/quotations/:id/generate-final-pdf", error);
    response.status(400).json({
      status: "error",
      message: error instanceof Error ? error.message : "Failed to generate final PDF"
    });
  }
});

app.post("/api/actions/quotations/:id/regenerate-draft", async (request, response) => {
  try {
    const result = await regenerateQuotationDraft(request.params.id);
    response.status(200).json({
      status: "ok",
      quotationId: result.quotation.id,
      draftFileUrl: result.draftFileUrl,
      driveFolderUrl: result.driveFolderUrl
    });
  } catch (error) {
    logRouteError("POST /api/actions/quotations/:id/regenerate-draft", error);
    response.status(400).json({
      status: "error",
      message: error instanceof Error ? error.message : "Failed to regenerate quotation draft"
    });
  }
});

app.post("/api/actions/quotations/:id/send-email", async (request, response) => {
  try {
    const result = await sendQuotationEmail(request.params.id);
    response.status(200).json({
      status: "ok",
      quotationId: result.quotation.id,
      recipient: result.recipient,
      documentUrl: result.documentUrl
    });
  } catch (error) {
    logRouteError("POST /api/actions/quotations/:id/send-email", error);
    response.status(400).json({
      status: "error",
      message: error instanceof Error ? error.message : "Failed to send quotation on email"
    });
  }
});

app.post("/api/actions/quotations/:id/send-whatsapp", async (request, response) => {
  try {
    const result = await sendQuotationWhatsApp(request.params.id);
    response.status(200).json({
      status: "ok",
      quotationId: result.quotation.id,
      recipient: result.recipient,
      documentUrl: result.documentUrl
    });
  } catch (error) {
    logRouteError("POST /api/actions/quotations/:id/send-whatsapp", error);
    response.status(400).json({
      status: "error",
      message: error instanceof Error ? error.message : "Failed to send quotation on WhatsApp"
    });
  }
});

app.post("/api/actions/enquiries/:id/send-product-documents", async (request, response) => {
  try {
    const { documentIds } = sendProductDocumentsSchema.parse(request.body);
    const enquiry = await getRecord<Record<string, unknown>>(env.AIRTABLE_ENQUIRIES_TABLE, request.params.id);
    const customerId = Array.isArray(enquiry.fields["Linked Customer"])
      ? String(enquiry.fields["Linked Customer"][0] || "")
      : "";

    if (!customerId) {
      throw new Error("This enquiry does not have a linked customer yet.");
    }

    const customer = await getRecord<Record<string, unknown>>(env.AIRTABLE_CUSTOMERS_TABLE, customerId);
    const documents = await getProductDocumentsByIds(documentIds);

    if (!documents.length) {
      throw new Error("No valid product documents were selected.");
    }

    const customerEmail = String(customer.fields.Email || "").trim();
    if (!customerEmail) {
      throw new Error("The linked customer does not have an email address yet.");
    }

    if (!isSmtpConfigured()) {
      throw new Error("SMTP is not configured yet.");
    }

    const customerName = String(customer.fields["Customer Name"] || "Customer").trim() || "Customer";
    const documentLines = documents.map((document) => `- ${document.fileName}: ${document.fileUrl}`);

    await sendMail({
      to: customerEmail,
      subject: `Product documents from Resham Sutra for ${customerName}`,
      text: [
        `Dear ${customerName},`,
        "",
        "Please find the requested product documents attached.",
        "",
        "Documents:",
        ...documentLines,
        "",
        "Regards,",
        "Resham Sutra Sales"
      ].join("\n"),
      attachments: documents.map((document) =>
        resolveDocumentAttachment(document.relativePath, document.fileName, document.mimeType)
      )
    });

    response.status(200).json({
      status: "ok",
      message: "Product documents emailed successfully. WhatsApp templates will be wired next.",
      customer: {
        id: customer.id,
        email: customerEmail,
        phone: String(customer.fields.WhatsApp || customer.fields.Phone || "")
      },
      documents
    });
  } catch (error) {
    logRouteError("POST /api/actions/enquiries/:id/send-product-documents", error);
    response.status(400).json({
      status: "error",
      message: error instanceof Error ? error.message : "Failed to prepare product document dispatch"
    });
  }
});

if (env.INTAKE_POLLING_ENABLED) {
  const run = async () => {
    try {
      const result = await processPendingEnquiries();
      if (result.processedCount || result.errorCount) {
        console.log(
          `[intake-poller] processed=${result.processedCount} errors=${result.errorCount}`
        );
      }
    } catch (error) {
      console.error("[intake-poller] failed", error);
    }
  };

  void run();
  setInterval(run, env.INTAKE_POLLING_INTERVAL_MS);
}

app.listen(env.PORT, () => {
  console.log(`API listening on port ${env.PORT}`);
});
