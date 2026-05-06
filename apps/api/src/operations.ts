import { stat } from "node:fs/promises";
import { listRecords, listRecordsPage, type AirtableRecord } from "./airtable.js";
import { env } from "./config.js";
import { getStoredDocumentArtifact } from "./documents.js";
import { ensureDefaultTemplateFolder, isDriveConfigured } from "./drive.js";
import { listProductDocuments } from "./product-documents.js";

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
  "Destination Address"?: string;
  "Destination State"?: string;
  "Destination City"?: string;
  "Destination Pincode"?: string;
  "Parser Status"?: string;
  "Linked Customer"?: string[];
  Quotations?: string[];
  "Requirement Summary"?: string;
  "Potential Product"?: string;
  "Receiver WhatsApp Number"?: string;
};

type CustomerFields = {
  "Client ID"?: string;
  "Customer Name"?: string;
  Company?: string;
  Phone?: string;
  Email?: string;
  Address?: string;
  State?: string;
  City?: string;
  Pincode?: string;
  "Customer Type"?: string;
  "Drive Folder URL"?: string;
};

type QuotationFields = {
  "Quotation Number"?: string;
  "Logged Date Time"?: string;
  "Linked Customer"?: string[];
  "Linked Enquiry"?: string[];
  Status?: string;
  "Draft File URL"?: string;
  "Draft Created Time"?: string;
  "Final PDF URL"?: string;
  "Final PDF Generated At"?: string;
  "Drive Folder URL"?: string;
  "Sent Date"?: string;
  "WhatsApp Sent Date Time"?: string;
  "Email Sent Date Time"?: string;
  "Send Quotation"?: boolean;
  "Send Reminder"?: boolean;
  "Mark Accepted"?: boolean;
  "Mark Rejected"?: boolean;
  "Reminder Count"?: number;
  "Last Reminder Date"?: string;
  "Next Reminder Date"?: string;
};

type QuotationLineItemFields = {
  Quotation?: string[];
  "Linked Product"?: string[];
  "Total Amount"?: number;
};

type OrderFields = {
  "Order Number"?: string;
  Quotation?: string[] | string;
  Customer?: string[] | string;
  Enquiries?: string[] | string;
  "Linked Customer"?: string[];
  "Linked Quotation"?: string[];
  "Order Date"?: string;
  "Order Status"?: string;
  "Total Amount"?: number;
  "Order Notes"?: string;
  "Line Items"?: string[] | string;
  "Quotation Grand Total"?: number;
  "Quotation Status"?: string;
  "Order Line Item Count"?: number;
  "Order Value per Item"?: string;
  "Order Fulfillment Progress"?: string;
  "Order Summary (AI)"?: string;
  "Order Risk/Attention Flag (AI)"?: string;
  "Order Value"?: number;
  "Payment Status"?: string;
  "Delivery Status"?: string;
  Address?: string;
  State?: string;
  City?: string;
  Pincode?: string | number;
};

type ProductFields = {
  "Product Key"?: string;
  Model?: string;
  "Product Name"?: string;
  Narration?: string;
  "Bulk Sale Price"?: number;
  MRP?: number;
  "GST %"?: number;
  "GST Amount"?: number;
  "Pkg & Transport"?: number;
  "Freight Amount"?: number;
  Freight?: number;
  "Source Sheet"?: string;
};

const OPERATIONS_SNAPSHOT_LIMIT = 100;
const OPERATIONS_LINE_ITEM_LIMIT = 1000;
const ENQUIRY_PAGE_FIELDS = [
  "Enquiry ID",
  "Logged Date Time",
  "Lead Name",
  "Company",
  "Phone",
  "Email",
  "Address",
  "State",
  "City",
  "Pincode",
  "Destination Address",
  "Destination State",
  "Destination City",
  "Destination Pincode",
  "Parser Status",
  "Linked Customer",
  "Quotations",
  "Requirement Summary",
  "Potential Product",
  "Receiver WhatsApp Number"
];
const QUOTATION_PAGE_FIELDS = [
  "Quotation Number",
  "Logged Date Time",
  "Linked Customer",
  "Linked Enquiry",
  "Status",
  "Draft File URL",
  "Draft Created Time",
  "Final PDF URL",
  "Final PDF Generated At",
  "Drive Folder URL",
  "Sent Date",
  "WhatsApp Sent Date Time",
  "Email Sent Date Time",
  "Send Quotation",
  "Send Reminder",
  "Mark Accepted",
  "Mark Rejected",
  "Reminder Count",
  "Last Reminder Date",
  "Next Reminder Date"
];

async function safeList<TFields extends Record<string, unknown>>(
  tableName: string,
  options: Parameters<typeof listRecords<TFields>>[1]
) {
  try {
    return await listRecords<TFields>(tableName, options);
  } catch {
    return [];
  }
}

async function resolvePdfGeneratedAt(
  quotation: { id: string; fields: QuotationFields },
  customer: CustomerFields | undefined
) {
  const generatedAt = quotation.fields["Final PDF Generated At"];
  if (generatedAt) {
    return generatedAt;
  }

  const quotationNumber = quotation.fields["Quotation Number"];
  const finalPdfUrl = quotation.fields["Final PDF URL"];
  const clientId = customer?.["Client ID"];
  const customerName = customer?.["Customer Name"];

  if (!quotationNumber || !finalPdfUrl || !clientId || !customerName) {
    return "";
  }

  try {
    const artifact = getStoredDocumentArtifact(
      `${clientId}-${customerName}`,
      quotationNumber,
      "pdf"
    );
    const fileStats = await stat(artifact.filePath);
    return fileStats.mtime.toISOString();
  } catch {
    return "";
  }
}

function statusValue(fields: EnquiryFields) {
  return String(fields["Parser Status"] || "New");
}

function stageCount<T extends { fields: Record<string, unknown> }>(
  records: T[],
  fieldName: string,
  expected: string
) {
  return records.filter((record) => String(record.fields[fieldName] || "") === expected).length;
}

function buildQuotationLineItemMetrics(
  quotationLineItems: Array<{ fields: QuotationLineItemFields }>
) {
  const metricsByQuotationId = new Map<string, { lineItemCount: number; quotationGrandTotal: number }>();

  for (const item of quotationLineItems) {
    const quotationIds = item.fields.Quotation || [];
    for (const quotationId of quotationIds) {
      const existing = metricsByQuotationId.get(quotationId) || {
        lineItemCount: 0,
        quotationGrandTotal: 0
      };
      existing.lineItemCount += 1;
      existing.quotationGrandTotal = Number(
        (existing.quotationGrandTotal + Number(item.fields["Total Amount"] || 0)).toFixed(2)
      );
      metricsByQuotationId.set(quotationId, existing);
    }
  }

  return metricsByQuotationId;
}

function mapCustomerRecord(record: AirtableRecord<CustomerFields>) {
  return {
    id: record.id,
    clientId: record.fields["Client ID"] || "",
    customerName: record.fields["Customer Name"] || "",
    company: record.fields.Company || "",
    phone: record.fields.Phone || "",
    email: record.fields.Email || "",
    address: record.fields.Address || "",
    state: record.fields.State || "",
    city: record.fields.City || "",
    pincode: record.fields.Pincode || "",
    customerType: record.fields["Customer Type"] || "",
    driveFolderUrl: record.fields["Drive Folder URL"] || ""
  };
}

function quoteFormulaValue(value: string) {
  return value.replace(/'/g, "\\'");
}

function exactFieldFormula(fieldName: string, value: string) {
  return `{${fieldName}}='${quoteFormulaValue(value)}'`;
}

function anyFieldFormula(fieldName: string, values?: string[]) {
  const filteredValues = (values || []).map((value) => value.trim()).filter(Boolean);
  if (!filteredValues.length) {
    return undefined;
  }

  if (filteredValues.length === 1) {
    return exactFieldFormula(fieldName, filteredValues[0]);
  }

  return `OR(${filteredValues.map((value) => exactFieldFormula(fieldName, value)).join(",")})`;
}

async function countRecords(tableName: string, fieldName: string, filterByFormula?: string) {
  return (
    await safeList<Record<string, unknown>>(tableName, {
      fields: [fieldName],
      filterByFormula
    })
  ).length;
}

function mapEnquiryRecord(record: AirtableRecord<EnquiryFields>) {
  return {
    id: record.id,
    enquiryId: record.fields["Enquiry ID"] || record.id,
    loggedDateTime: record.fields["Logged Date Time"] || record.createdTime || "",
    leadName: record.fields["Lead Name"] || "",
    company: record.fields.Company || "",
    phone: record.fields.Phone || "",
    email: record.fields.Email || "",
    address: record.fields.Address || "",
    state: record.fields.State || "",
    city: record.fields.City || "",
    pincode: record.fields.Pincode || "",
    destinationAddress: record.fields["Destination Address"] || "",
    destinationState: record.fields["Destination State"] || "",
    destinationCity: record.fields["Destination City"] || "",
    destinationPincode: record.fields["Destination Pincode"] || "",
    parserStatus: statusValue(record.fields),
    linkedCustomerId: record.fields["Linked Customer"]?.[0] || "",
    quotations: record.fields.Quotations || [],
    mappedProductDocuments: [],
    driveFolderUrl: "",
    requirementSummary: record.fields["Requirement Summary"] || "",
    potentialProduct: record.fields["Potential Product"] || "",
    receiverWhatsappNumber: record.fields["Receiver WhatsApp Number"] || ""
  };
}

function mapQuotationRecord(
  record: AirtableRecord<QuotationFields>,
  quotationMetricsById = new Map<string, { lineItemCount: number; quotationGrandTotal: number }>(),
  pdfGeneratedAtByQuotationId = new Map<string, string>()
) {
  return {
    id: record.id,
    quotationNumber: record.fields["Quotation Number"] || record.id,
    loggedDateTime: record.fields["Logged Date Time"] || record.createdTime || "",
    linkedCustomerId: record.fields["Linked Customer"]?.[0] || "",
    linkedEnquiryId: record.fields["Linked Enquiry"]?.[0] || "",
    status: record.fields.Status || "",
    draftFileUrl: record.fields["Draft File URL"] || "",
    draftCreatedTime: record.fields["Draft Created Time"] || "",
    finalPdfUrl: record.fields["Final PDF URL"] || "",
    driveFolderUrl: record.fields["Drive Folder URL"] || "",
    sentDate: record.fields["Sent Date"] || "",
    whatsappSentDateTime: record.fields["WhatsApp Sent Date Time"] || "",
    emailSentDateTime: record.fields["Email Sent Date Time"] || "",
    finalPdfGeneratedAt: pdfGeneratedAtByQuotationId.get(record.id) || record.fields["Final PDF Generated At"] || "",
    lineItemCount: quotationMetricsById.get(record.id)?.lineItemCount || 0,
    quotationGrandTotal: quotationMetricsById.get(record.id)?.quotationGrandTotal || 0,
    sendQuotation: Boolean(record.fields["Send Quotation"]),
    sendReminder: Boolean(record.fields["Send Reminder"]),
    markAccepted: Boolean(record.fields["Mark Accepted"]),
    markRejected: Boolean(record.fields["Mark Rejected"]),
    reminderCount: Number(record.fields["Reminder Count"] || 0),
    lastReminderDate: record.fields["Last Reminder Date"] || "",
    nextReminderDate: record.fields["Next Reminder Date"] || ""
  };
}

export async function getOperationsCustomersPage(input?: { offset?: string; pageSize?: number }) {
  const totalCountPromise = countRecords(env.AIRTABLE_CUSTOMERS_TABLE, "Client ID");
  const page = await listRecordsPage<CustomerFields>(env.AIRTABLE_CUSTOMERS_TABLE, {
    fields: [
      "Client ID",
      "Customer Name",
      "Company",
      "Phone",
      "Email",
      "Address",
      "State",
      "City",
      "Pincode",
      "Customer Type",
      "Drive Folder URL"
    ],
    offset: input?.offset,
    pageSize: input?.pageSize ?? 25,
    sort: [{ field: "Client ID", direction: "asc" }]
  });

  return {
    customers: page.records.map(mapCustomerRecord),
    nextOffset: page.offset,
    pageSize: page.pageSize,
    totalCount: await totalCountPromise
  };
}

export async function getOperationsEnquiriesPage(input?: { offset?: string; pageSize?: number; status?: string }) {
  const filterByFormula = input?.status && input.status !== "All"
    ? exactFieldFormula("Parser Status", input.status)
    : undefined;
  const totalCountPromise = countRecords(env.AIRTABLE_ENQUIRIES_TABLE, "Enquiry ID", filterByFormula);
  const page = await listRecordsPage<EnquiryFields>(env.AIRTABLE_ENQUIRIES_TABLE, {
    fields: ENQUIRY_PAGE_FIELDS,
    filterByFormula,
    offset: input?.offset,
    pageSize: input?.pageSize ?? 25,
    sort: [{ field: "Created At", direction: "desc" }]
  });

  return {
    enquiries: page.records.map(mapEnquiryRecord),
    nextOffset: page.offset,
    pageSize: page.pageSize,
    totalCount: await totalCountPromise
  };
}

export async function getOperationsQuotationsPage(input?: { offset?: string; pageSize?: number; statuses?: string[] }) {
  const filterByFormula = anyFieldFormula("Status", input?.statuses);
  const totalCountPromise = countRecords(env.AIRTABLE_QUOTATIONS_TABLE, "Quotation Number", filterByFormula);
  const page = await listRecordsPage<QuotationFields>(env.AIRTABLE_QUOTATIONS_TABLE, {
    fields: QUOTATION_PAGE_FIELDS,
    filterByFormula,
    offset: input?.offset,
    pageSize: input?.pageSize ?? 25,
    sort: [{ field: "Quotation Number", direction: "desc" }]
  });

  return {
    quotations: page.records.map((record) => mapQuotationRecord(record)),
    nextOffset: page.offset,
    pageSize: page.pageSize,
    totalCount: await totalCountPromise
  };
}

export async function getOperationsSnapshot() {
  const [enquiries, customers, quotations, quotationLineItems, orders, products] = await Promise.all([
    safeList<EnquiryFields>(env.AIRTABLE_ENQUIRIES_TABLE, {
      fields: ENQUIRY_PAGE_FIELDS,
      maxRecords: OPERATIONS_SNAPSHOT_LIMIT,
      sort: [{ field: "Created At", direction: "desc" }]
    }),
    safeList<CustomerFields>(env.AIRTABLE_CUSTOMERS_TABLE, {
      maxRecords: OPERATIONS_SNAPSHOT_LIMIT,
      sort: [{ field: "Client ID", direction: "asc" }]
    }),
    safeList<QuotationFields>(env.AIRTABLE_QUOTATIONS_TABLE, {
      fields: QUOTATION_PAGE_FIELDS,
      maxRecords: OPERATIONS_SNAPSHOT_LIMIT,
      sort: [{ field: "Quotation Number", direction: "desc" }]
    }),
    safeList<QuotationLineItemFields>(env.AIRTABLE_QUOTATION_LINE_ITEMS_TABLE, {
      maxRecords: OPERATIONS_LINE_ITEM_LIMIT,
      fields: ["Quotation", "Linked Product", "Total Amount"]
    }),
    safeList<OrderFields>(env.AIRTABLE_ORDERS_TABLE, {
      maxRecords: OPERATIONS_SNAPSHOT_LIMIT
    }),
    safeList<ProductFields>(env.AIRTABLE_PRODUCTS_TABLE, {
      maxRecords: OPERATIONS_SNAPSHOT_LIMIT
    })
  ]);

  const quotationIdsByEnquiryId = new Map<string, string[]>();
  const productDocuments = await listProductDocuments(products.map((record) => record.id));
  const productDocumentsByProductId = new Map<string, typeof productDocuments>();
  const productNamesById = new Map(products.map((record) => [record.id, record.fields["Product Name"] || record.id]));

  for (const document of productDocuments) {
    const existing = productDocumentsByProductId.get(document.productId) ?? [];
    existing.push(document);
    productDocumentsByProductId.set(document.productId, existing);
  }

  for (const quotation of quotations) {
    const linkedEnquiryId = quotation.fields["Linked Enquiry"]?.[0];

    if (!linkedEnquiryId) {
      continue;
    }

    const existing = quotationIdsByEnquiryId.get(linkedEnquiryId) ?? [];
    existing.push(quotation.id);
    quotationIdsByEnquiryId.set(linkedEnquiryId, existing);
  }

  const quotationById = new Map(quotations.map((record) => [record.id, record]));
  const customersById = new Map(customers.map((record) => [record.id, record.fields]));
  const quotationMetricsById = buildQuotationLineItemMetrics(quotationLineItems);
  const pdfGeneratedAtEntries = await Promise.all(
    quotations.map(async (record) => {
      const generatedAt = await resolvePdfGeneratedAt(
        record,
        customersById.get(record.fields["Linked Customer"]?.[0] || "")
      );

      return [record.id, generatedAt] as const;
    })
  );
  const pdfGeneratedAtByQuotationId = new Map(pdfGeneratedAtEntries);
  const enquiryIdsWithoutDraft = enquiries.filter((record) => {
    const directQuotations = record.fields.Quotations || [];
    const fallbackQuotations = quotationIdsByEnquiryId.get(record.id) || [];
    const mergedQuotations = Array.from(new Set([...directQuotations, ...fallbackQuotations]));

    if (!mergedQuotations.length) {
      return true;
    }

    return !mergedQuotations.some((quotationId) => Boolean(quotationById.get(quotationId)?.fields["Draft File URL"]));
  }).length;

  const defaultTemplateFolder = isDriveConfigured()
    ? await ensureDefaultTemplateFolder().catch(() => null)
    : null;

  return {
    actions: {
      interfaceUrl: env.AIRTABLE_INTERFACE_URL,
      enquiryFormUrl: env.AIRTABLE_ENQUIRY_FORM_URL,
      lineItemsFormUrl: env.AIRTABLE_LINE_ITEMS_FORM_URL,
      productsFormUrl: env.AIRTABLE_PRODUCTS_FORM_URL,
      defaultTemplateFolderUrl: defaultTemplateFolder?.folderUrl || ""
    },
    metrics: [
      {
        label: "New Enquiries",
        value: enquiryIdsWithoutDraft
      },
      {
        label: "Draft Quote",
        value: enquiries.filter((record) => statusValue(record.fields) === "Draft Quote").length
      },
      { label: "Approved Quote", value: stageCount(quotations, "Status", "Approved Quote") },
      { label: "Sent Quote", value: stageCount(quotations, "Status", "Sent Quote") },
      { label: "Orders", value: orders.length }
    ],
    totals: {
      enquiries: await countRecords(env.AIRTABLE_ENQUIRIES_TABLE, "Enquiry ID"),
      customers: await countRecords(env.AIRTABLE_CUSTOMERS_TABLE, "Client ID"),
      quotations: await countRecords(env.AIRTABLE_QUOTATIONS_TABLE, "Quotation Number"),
      orders: await countRecords(env.AIRTABLE_ORDERS_TABLE, "Order Number"),
      products: await countRecords(env.AIRTABLE_PRODUCTS_TABLE, "Product Key")
    },
    enquiries: enquiries.map((record) => {
      const directQuotations = record.fields.Quotations || [];
      const fallbackQuotations = quotationIdsByEnquiryId.get(record.id) || [];
      const mergedQuotations = Array.from(new Set([...directQuotations, ...fallbackQuotations]));
      const linkedProductIds = Array.from(
        new Set(
          quotationLineItems
            .filter((item) => item.fields.Quotation?.some((quotationId) => mergedQuotations.includes(quotationId)))
            .flatMap((item) => item.fields["Linked Product"] || [])
        )
      );
      const mappedProductDocuments = linkedProductIds.flatMap((productId) =>
        (productDocumentsByProductId.get(productId) || []).map((document) => ({
          id: document.id,
          productId,
          productName: productNamesById.get(productId) || productId,
          fileName: document.fileName,
          mimeType: document.mimeType,
          uploadedAt: document.uploadedAt,
          fileUrl: document.fileUrl
        }))
      );

      return {
        id: record.id,
        enquiryId: record.fields["Enquiry ID"] || record.id,
        loggedDateTime: record.fields["Logged Date Time"] || record.createdTime || "",
        leadName: record.fields["Lead Name"] || "",
        company: record.fields.Company || "",
        phone: record.fields.Phone || "",
        email: record.fields.Email || "",
        address: record.fields.Address || "",
        state: record.fields.State || "",
        city: record.fields.City || "",
        pincode: record.fields.Pincode || "",
        destinationAddress: record.fields["Destination Address"] || "",
        destinationState: record.fields["Destination State"] || "",
        destinationCity: record.fields["Destination City"] || "",
        destinationPincode: record.fields["Destination Pincode"] || "",
        parserStatus: statusValue(record.fields),
        linkedCustomerId: record.fields["Linked Customer"]?.[0] || "",
        quotations: mergedQuotations,
        mappedProductDocuments,
        driveFolderUrl: "",
        requirementSummary: record.fields["Requirement Summary"] || "",
        potentialProduct: record.fields["Potential Product"] || "",
        receiverWhatsappNumber: record.fields["Receiver WhatsApp Number"] || ""
      };
    }),
    customers: customers.map(mapCustomerRecord),
    quotations: quotations.map((record) =>
      mapQuotationRecord(record, quotationMetricsById, pdfGeneratedAtByQuotationId)
    ),
    orders: orders.map((record) => ({
      id: record.id,
      orderNumber: record.fields["Order Number"] || record.id,
      linkedCustomerId:
        record.fields["Linked Customer"]?.[0] ||
        (Array.isArray(record.fields.Customer) ? record.fields.Customer[0] || "" : String(record.fields.Customer || "")),
      linkedQuotationId:
        record.fields["Linked Quotation"]?.[0] ||
        (Array.isArray(record.fields.Quotation) ? record.fields.Quotation[0] || "" : String(record.fields.Quotation || "")),
      linkedEnquiryId:
        (Array.isArray(record.fields.Enquiries) ? record.fields.Enquiries[0] || "" : String(record.fields.Enquiries || "")),
      orderDate: record.fields["Order Date"] || "",
      orderStatus: record.fields["Order Status"] || "",
      totalAmount: record.fields["Total Amount"] || 0,
      orderNotes: record.fields["Order Notes"] || "",
      quotationGrandTotal: record.fields["Quotation Grand Total"] || 0,
      quotationStatus: record.fields["Quotation Status"] || "",
      orderLineItemCount: Number(record.fields["Order Line Item Count"] || 0),
      orderValuePerItem: record.fields["Order Value per Item"] || "",
      orderFulfillmentProgress: record.fields["Order Fulfillment Progress"] || "",
      orderSummary: record.fields["Order Summary (AI)"] || "",
      orderRiskAttentionFlag: record.fields["Order Risk/Attention Flag (AI)"] || "",
      orderValue: record.fields["Order Value"] || record.fields["Total Amount"] || 0,
      paymentStatus: record.fields["Payment Status"] || "",
      deliveryStatus: record.fields["Delivery Status"] || "",
      address: record.fields.Address || "",
      state: record.fields.State || "",
      city: record.fields.City || "",
      pincode: String(record.fields.Pincode || "")
    })),
    products: products.map((record) => ({
      id: record.id,
      productKey: record.fields["Product Key"] || record.id,
      model: record.fields.Model || "",
      name: record.fields["Product Name"] || "",
      narration: record.fields.Narration || "",
      bulkSalePrice: record.fields["Bulk Sale Price"] || 0,
      mrp: record.fields.MRP || 0,
      gstPercent: record.fields["GST Amount"] || record.fields["GST %"] || 0,
      transportCharge:
        record.fields["Freight Amount"] ||
        record.fields.Freight ||
        record.fields["Pkg & Transport"] ||
        0,
      sourceSheet: record.fields["Source Sheet"] || "",
      documents: (productDocumentsByProductId.get(record.id) || []).map((document) => ({
        id: document.id,
        productId: record.id,
        productName: record.fields["Product Name"] || record.id,
        fileName: document.fileName,
        mimeType: document.mimeType,
        uploadedAt: document.uploadedAt,
        fileUrl: document.fileUrl
      }))
    }))
  };
}
