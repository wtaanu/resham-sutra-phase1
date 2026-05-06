import { listRecords, listRecordsPage, type AirtableRecord } from "./airtable.js";
import { env } from "./config.js";
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
  "Destination Address"?: string;
  "Destination State"?: string;
  "Destination City"?: string;
  "Destination Pincode"?: string;
  "Customer Type"?: string;
  "Drive Folder URL"?: string;
};

type QuotationFields = {
  "Quotation Number"?: string;
  "Reference Number"?: string;
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
  "Total Amount"?: number | string;
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
  "Quotation Grand Total"?: number | string;
  "Quotation Status"?: string;
  "Order Line Item Count"?: number;
  "Order Value per Item"?: string;
  "Order Fulfillment Progress"?: string;
  "Order Summary (AI)"?: string;
  "Order Risk/Attention Flag (AI)"?: string;
  "Order Value"?: number | string;
  "Payment Status"?: string;
  "Payment Terms"?: string;
  "Order Ref Number Client"?: string;
  "Delivery Status"?: string;
  Address?: string;
  State?: string;
  City?: string;
  Pincode?: string | number;
  "Destination Address"?: string;
  "Destination State"?: string;
  "Destination City"?: string;
  "Destination Pincode"?: string | number;
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
  "Reference Number",
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
    destinationAddress: record.fields["Destination Address"] || "",
    destinationState: record.fields["Destination State"] || "",
    destinationCity: record.fields["Destination City"] || "",
    destinationPincode: record.fields["Destination Pincode"] || "",
    customerType: record.fields["Customer Type"] || "",
    driveFolderUrl: record.fields["Drive Folder URL"] || ""
  };
}

function latestFirst<TFields extends Record<string, unknown>>(records: Array<AirtableRecord<TFields>>) {
  return [...records].sort((left, right) => {
    const leftTime = Date.parse(left.createdTime || "");
    const rightTime = Date.parse(right.createdTime || "");
    return (Number.isNaN(rightTime) ? 0 : rightTime) - (Number.isNaN(leftTime) ? 0 : leftTime);
  });
}

function quoteFormulaValue(value: string) {
  return value.replace(/'/g, "\\'");
}

function recordIdFormula(recordIds: string[]) {
  const uniqueRecordIds = Array.from(new Set(recordIds.filter(Boolean)));
  if (!uniqueRecordIds.length) {
    return "";
  }

  const clauses = uniqueRecordIds.map((recordId) => `RECORD_ID()='${quoteFormulaValue(recordId)}'`);
  return clauses.length === 1 ? clauses[0] : `OR(${clauses.join(",")})`;
}

function linkedRecordFormula(fieldName: string, linkedDisplayValues: string[]) {
  const uniqueValues = Array.from(new Set(linkedDisplayValues.filter(Boolean)));
  if (!uniqueValues.length) {
    return "";
  }

  const clauses = uniqueValues.map((value) => `FIND('${quoteFormulaValue(value)}', ARRAYJOIN({${fieldName}}))`);
  return clauses.length === 1 ? clauses[0] : `OR(${clauses.join(",")})`;
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
  pdfGeneratedAtByQuotationId = new Map<string, string>(),
  customerNameById = new Map<string, string>(),
  customerDefaultsById = new Map<string, ReturnType<typeof mapCustomerRecord>>()
) {
  const linkedCustomerId = record.fields["Linked Customer"]?.[0] || "";
  const customerDefaults = customerDefaultsById.get(linkedCustomerId);

  return {
    id: record.id,
    quotationNumber: record.fields["Quotation Number"] || record.id,
    referenceNumber: record.fields["Reference Number"] || "",
    loggedDateTime: record.fields["Logged Date Time"] || record.createdTime || "",
    linkedCustomerId,
    customerName: customerNameById.get(linkedCustomerId) || "",
    customerAddress: customerDefaults?.address || "",
    customerState: customerDefaults?.state || "",
    customerCity: customerDefaults?.city || "",
    customerPincode: customerDefaults?.pincode || "",
    customerDestinationAddress: customerDefaults?.destinationAddress || "",
    customerDestinationState: customerDefaults?.destinationState || "",
    customerDestinationCity: customerDefaults?.destinationCity || "",
    customerDestinationPincode: customerDefaults?.destinationPincode || "",
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

export async function getOperationsCustomersPage(input?: { includeTotal?: boolean; offset?: string; pageSize?: number }) {
  const totalCountPromise = input?.includeTotal === false
    ? Promise.resolve(0)
    : countRecords(env.AIRTABLE_CUSTOMERS_TABLE, "Client ID");
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
    pageSize: input?.pageSize ?? 25
  });

  return {
    customers: latestFirst(page.records).map(mapCustomerRecord),
    nextOffset: page.offset,
    pageSize: page.pageSize,
    totalCount: await totalCountPromise
  };
}

export async function getOperationsEnquiriesPage(input?: { includeTotal?: boolean; offset?: string; pageSize?: number; status?: string }) {
  const filterByFormula = input?.status && input.status !== "All"
    ? exactFieldFormula("Parser Status", input.status)
    : undefined;
  const totalCountPromise = input?.includeTotal === false
    ? Promise.resolve(0)
    : countRecords(env.AIRTABLE_ENQUIRIES_TABLE, "Enquiry ID", filterByFormula);
  const page = await listRecordsPage<EnquiryFields>(env.AIRTABLE_ENQUIRIES_TABLE, {
    fields: ENQUIRY_PAGE_FIELDS,
    filterByFormula,
    offset: input?.offset,
    pageSize: input?.pageSize ?? 25
  });

  return {
    enquiries: latestFirst(page.records).map(mapEnquiryRecord),
    nextOffset: page.offset,
    pageSize: page.pageSize,
    totalCount: await totalCountPromise
  };
}

export async function getOperationsQuotationsPage(input?: { includeTotal?: boolean; offset?: string; pageSize?: number; statuses?: string[] }) {
  const filterByFormula = anyFieldFormula("Status", input?.statuses);
  const totalCountPromise = input?.includeTotal === false
    ? Promise.resolve(0)
    : countRecords(env.AIRTABLE_QUOTATIONS_TABLE, "Quotation Number", filterByFormula);
  const page = await listRecordsPage<QuotationFields>(env.AIRTABLE_QUOTATIONS_TABLE, {
    fields: QUOTATION_PAGE_FIELDS,
    filterByFormula,
    offset: input?.offset,
    pageSize: input?.pageSize ?? 25,
    sort: [{ field: "Quotation Number", direction: "desc" }]
  });
  const customerIds = page.records.flatMap((record) => record.fields["Linked Customer"] || []);
  const quotationNumbers = page.records.map((record) => record.fields["Quotation Number"] || record.id);
  const [linkedCustomers, quotationLineItems] = await Promise.all([
    customerIds.length
      ? listRecords<CustomerFields>(env.AIRTABLE_CUSTOMERS_TABLE, {
          fields: [
            "Customer Name",
            "Address",
            "State",
            "City",
            "Pincode",
            "Destination Address",
            "Destination State",
            "Destination City",
            "Destination Pincode"
          ],
          filterByFormula: recordIdFormula(customerIds),
          maxRecords: customerIds.length
        })
      : Promise.resolve([]),
    quotationNumbers.length
      ? listRecords<QuotationLineItemFields>(env.AIRTABLE_QUOTATION_LINE_ITEMS_TABLE, {
          fields: ["Quotation", "Linked Product", "Total Amount"],
          filterByFormula: linkedRecordFormula("Quotation", quotationNumbers),
          maxRecords: 1000
        })
      : Promise.resolve([])
  ]);
  const customerNameById = new Map(
    linkedCustomers.map((record) => [record.id, record.fields["Customer Name"] || record.id])
  );
  const customerDefaultsById = new Map(linkedCustomers.map((record) => [record.id, mapCustomerRecord(record)]));
  const quotationMetricsById = buildQuotationLineItemMetrics(quotationLineItems);

  return {
    quotations: page.records.map((record) =>
      mapQuotationRecord(record, quotationMetricsById, new Map(), customerNameById, customerDefaultsById)
    ),
    nextOffset: page.offset,
    pageSize: page.pageSize,
    totalCount: await totalCountPromise
  };
}

export async function getOperationsSnapshot() {
  const [enquiries, customers, quotations, quotationLineItems, orders, products] = await Promise.all([
    safeList<EnquiryFields>(env.AIRTABLE_ENQUIRIES_TABLE, {
      fields: ENQUIRY_PAGE_FIELDS,
      maxRecords: OPERATIONS_SNAPSHOT_LIMIT
    }),
    safeList<CustomerFields>(env.AIRTABLE_CUSTOMERS_TABLE, {
      maxRecords: OPERATIONS_SNAPSHOT_LIMIT
    }),
    safeList<QuotationFields>(env.AIRTABLE_QUOTATIONS_TABLE, {
      fields: QUOTATION_PAGE_FIELDS,
      maxRecords: OPERATIONS_SNAPSHOT_LIMIT,
      sort: [{ field: "Quotation Number", direction: "desc" }]
    }),
    Promise.resolve([] as Array<AirtableRecord<QuotationLineItemFields>>),
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
  const quotationMetricsById = buildQuotationLineItemMetrics(quotationLineItems);
  const pdfGeneratedAtByQuotationId = new Map<string, string>();
  const customerNameById = new Map(
    customers.map((record) => [record.id, record.fields["Customer Name"] || record.id])
  );
  const customerClientIdById = new Map(
    customers.map((record) => [record.id, record.fields["Client ID"] || record.id])
  );
  const customerDefaultsById = new Map(customers.map((record) => [record.id, mapCustomerRecord(record)]));
  const enquiryIdsWithoutDraft = enquiries.filter((record) => {
    const directQuotations = record.fields.Quotations || [];
    const fallbackQuotations = quotationIdsByEnquiryId.get(record.id) || [];
    const mergedQuotations = Array.from(new Set([...directQuotations, ...fallbackQuotations]));

    if (!mergedQuotations.length) {
      return true;
    }

    return !mergedQuotations.some((quotationId) => Boolean(quotationById.get(quotationId)?.fields["Draft File URL"]));
  }).length;

  const [
    totalEnquiries,
    totalCustomers,
    totalQuotations,
    totalOrders,
    totalProducts,
    totalNewEnquiries,
    totalDraftQuoteEnquiries,
    totalApprovedQuotations,
    totalSentQuotations
  ] = await Promise.all([
    countRecords(env.AIRTABLE_ENQUIRIES_TABLE, "Enquiry ID"),
    countRecords(env.AIRTABLE_CUSTOMERS_TABLE, "Client ID"),
    countRecords(env.AIRTABLE_QUOTATIONS_TABLE, "Quotation Number"),
    countRecords(env.AIRTABLE_ORDERS_TABLE, "Order Number"),
    countRecords(env.AIRTABLE_PRODUCTS_TABLE, "Product Key"),
    countRecords(
      env.AIRTABLE_ENQUIRIES_TABLE,
      "Enquiry ID",
      anyFieldFormula("Parser Status", ["New", "New Enquiries"])
    ),
    countRecords(env.AIRTABLE_ENQUIRIES_TABLE, "Enquiry ID", exactFieldFormula("Parser Status", "Draft Quote")),
    countRecords(env.AIRTABLE_QUOTATIONS_TABLE, "Quotation Number", exactFieldFormula("Status", "Approved Quote")),
    countRecords(env.AIRTABLE_QUOTATIONS_TABLE, "Quotation Number", exactFieldFormula("Status", "Sent Quote"))
  ]);

  return {
    actions: {
      interfaceUrl: env.AIRTABLE_INTERFACE_URL,
      enquiryFormUrl: env.AIRTABLE_ENQUIRY_FORM_URL,
      lineItemsFormUrl: env.AIRTABLE_LINE_ITEMS_FORM_URL,
      productsFormUrl: env.AIRTABLE_PRODUCTS_FORM_URL,
      defaultTemplateFolderUrl: ""
    },
    metrics: [
      {
        label: "New Enquiries",
        value: totalNewEnquiries
      },
      {
        label: "Draft Quote",
        value: totalDraftQuoteEnquiries
      },
      { label: "Approved Quote", value: totalApprovedQuotations },
      { label: "Sent Quote", value: totalSentQuotations },
      { label: "Orders", value: totalOrders }
    ],
    totals: {
      enquiries: totalEnquiries,
      customers: totalCustomers,
      quotations: totalQuotations,
      orders: totalOrders,
      products: totalProducts
    },
    enquiries: latestFirst(enquiries).map((record) => {
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
    customers: latestFirst(customers).map(mapCustomerRecord),
    quotations: quotations.map((record) =>
      mapQuotationRecord(record, quotationMetricsById, pdfGeneratedAtByQuotationId, customerNameById, customerDefaultsById)
    ),
    orders: latestFirst(orders).map((record) => ({
      id: record.id,
      orderNumber: record.fields["Order Number"] || record.id,
      linkedCustomerId:
        record.fields["Linked Customer"]?.[0] ||
        (Array.isArray(record.fields.Customer) ? record.fields.Customer[0] || "" : String(record.fields.Customer || "")),
      customerName: customerNameById.get(
        record.fields["Linked Customer"]?.[0] ||
          (Array.isArray(record.fields.Customer) ? record.fields.Customer[0] || "" : String(record.fields.Customer || ""))
      ) || "",
      customerClientId: customerClientIdById.get(
        record.fields["Linked Customer"]?.[0] ||
          (Array.isArray(record.fields.Customer) ? record.fields.Customer[0] || "" : String(record.fields.Customer || ""))
      ) || "",
      linkedQuotationId:
        record.fields["Linked Quotation"]?.[0] ||
        (Array.isArray(record.fields.Quotation) ? record.fields.Quotation[0] || "" : String(record.fields.Quotation || "")),
      linkedEnquiryId:
        (Array.isArray(record.fields.Enquiries) ? record.fields.Enquiries[0] || "" : String(record.fields.Enquiries || "")),
      orderDate: record.fields["Order Date"] || "",
      orderStatus: record.fields["Order Status"] || "",
      totalAmount: Number(record.fields["Total Amount"] || 0),
      orderNotes: record.fields["Order Notes"] || "",
      quotationGrandTotal: Number(record.fields["Quotation Grand Total"] || 0),
      quotationStatus: record.fields["Quotation Status"] || "",
      orderLineItemCount: Number(record.fields["Order Line Item Count"] || 0),
      orderValuePerItem: record.fields["Order Value per Item"] || "",
      orderFulfillmentProgress: record.fields["Order Fulfillment Progress"] || "",
      orderSummary: record.fields["Order Summary (AI)"] || "",
      orderRiskAttentionFlag: record.fields["Order Risk/Attention Flag (AI)"] || "",
      orderValue: Number(record.fields["Order Value"] || record.fields["Total Amount"] || 0),
      paymentStatus: record.fields["Payment Status"] || "",
      paymentTerms: record.fields["Payment Terms"] || "",
      orderRefNumberClient: record.fields["Order Ref Number Client"] || "",
      deliveryStatus: record.fields["Delivery Status"] || "",
      address: record.fields.Address || "",
      state: record.fields.State || "",
      city: record.fields.City || "",
      pincode: String(record.fields.Pincode || ""),
      destinationAddress: record.fields["Destination Address"] || "",
      destinationState: record.fields["Destination State"] || "",
      destinationCity: record.fields["Destination City"] || "",
      destinationPincode: String(record.fields["Destination Pincode"] || "")
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
