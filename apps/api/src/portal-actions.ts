import { z } from "zod";
import {
  createRecord,
  createRecords,
  deleteRecords,
  getRecord,
  listRecords,
  updateRecord,
  type AirtableRecord
} from "./airtable.js";
import type { AuthenticatedUser } from "./auth.js";
import { env } from "./config.js";
import { createRecordWithUniqueNumber } from "./numbering.js";
import {
  createCustomerForEnquiry,
  generateFinalPdfForQuotation,
  refreshDraftForQuotation
} from "./intake-processor.js";

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
  Pincode?: string | number;
  "Destination Address"?: string;
  "Destination State"?: string;
  "Destination City"?: string;
  "Destination Pincode"?: string | number;
  "Parser Status"?: string;
  "Linked Customer"?: string[];
  Quotations?: string[];
  "Drive Folder URL"?: string;
  "Requirement Summary"?: string;
  "Requested Asset"?: string;
  "Potential Product"?: string;
  "Receiver WhatsApp Number"?: string;
};

type CustomerFields = {
  "Client ID"?: string;
  "Customer Name"?: string;
  Company?: string;
  Phone?: string;
  WhatsApp?: string;
  Email?: string;
  Address?: string;
  State?: string;
  City?: string;
  Pincode?: string | number;
  "Drive Folder URL"?: string;
};

type QuotationFields = {
  "Quotation Number"?: string;
  "Logged Date Time"?: string;
  "Linked Enquiry"?: string[];
  "Linked Customer"?: string[];
  Status?: string;
  "Draft Format"?: string;
  "Draft File URL"?: string;
  "Draft Created Time"?: string;
  "Final PDF URL"?: string;
  "Final PDF Generated At"?: string;
  "Drive Folder URL"?: string;
  "Reference Number"?: string;
  "Buyer Block"?: string;
  "Sent Date"?: string;
  "WhatsApp Sent Date Time"?: string;
  "Email Sent Date Time"?: string;
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
};

type QuotationLineItemFields = {
  Name?: string;
  Quotation?: string[];
  "Linked Product"?: string[];
  "Line No."?: number;
  "Description Override"?: string;
  Qty?: number;
  "Rate Per Unit"?: number;
  "Pkg & Transport"?: number;
  "GST %"?: number;
  "GST Amount"?: number;
  "Total Amount"?: number;
  "Unit Value"?: number;
};

type OrderFields = {
  "Order Number"?: string;
  "Order Date"?: string;
  Quotation?: string[] | string;
  Customer?: string[] | string;
  Enquiries?: string[] | string;
  "Linked Quotation"?: string[];
  "Linked Customer"?: string[];
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
};

type ChangeLogFields = {
  "Changed At"?: string;
  "Action"?: string;
  "Changed By Name"?: string;
  "Changed By Email"?: string;
  "Entity Record ID"?: string;
  "Entity Label"?: string;
  "Before JSON"?: string;
  "After JSON"?: string;
};

type PortalQuotationLineItem = {
  id: string;
  productId: string;
  qty: number;
  rate: number;
  transport: number;
  gstPercent: number;
  totalAmount: number;
};

const airtableRecordIdSchema = z.string().regex(/^rec[a-zA-Z0-9]+$/, "Expected an Airtable record ID");
const optionalAirtableRecordIdSchema = z.union([airtableRecordIdSchema, z.literal("")]).default("");
const enquirySourceSchema = z.enum(["manual", "whatsapp"]).default("manual");
const optionalPincodeSchema = z.string().trim().optional().default("");
const optionalPhoneSchema = z.string().trim().optional().default("");
const optionalEmailSchema = z
  .string()
  .trim()
  .refine((value) => value === "" || value.includes("@"), "Email must include @");
const optionalEnquiryFields = ["Logged Date Time", "Receiver WhatsApp Number"] as const;
const optionalQuotationFields = [
  "Logged Date Time",
  "Draft Created Time",
  "WhatsApp Sent Date Time",
  "Email Sent Date Time",
  "Final PDF Generated At"
] as const;

function linkedRecordIds(...recordIds: Array<string | undefined>) {
  return recordIds.filter((value): value is string => Boolean(value));
}

const enquiryPayloadSchema = z.object({
  source: enquirySourceSchema,
  linkedCustomerId: optionalAirtableRecordIdSchema,
  leadName: z.string().trim().min(1, "Lead name is required"),
  company: z.string().trim().optional().default(""),
  phone: optionalPhoneSchema.default(""),
  email: optionalEmailSchema.default(""),
  address: z.string().trim().optional().default(""),
  state: z.string().trim().optional().default(""),
  city: z.string().trim().optional().default(""),
  pincode: optionalPincodeSchema,
  destinationAddress: z.string().trim().optional().default(""),
  destinationState: z.string().trim().optional().default(""),
  destinationCity: z.string().trim().optional().default(""),
  destinationPincode: optionalPincodeSchema,
  requirementSummary: z.string().trim().optional().default(""),
  potentialProduct: z.string().trim().optional().default(""),
  receiverWhatsappNumber: optionalPhoneSchema.default("")
}).superRefine((input, context) => {
  if (input.source !== "manual") {
    return;
  }

  const normalizedPhone = String(input.phone ?? "").replace(/\D/g, "");
  const normalizedReceiverWhatsapp = String(input.receiverWhatsappNumber ?? "").replace(/\D/g, "");
  const normalizedPincode = String(input.pincode ?? "").replace(/\D/g, "");
  const normalizedDestinationPincode = String(input.destinationPincode ?? "").replace(/\D/g, "");

  if (normalizedPhone && normalizedPhone.length !== 10) {
    context.addIssue({
      code: z.ZodIssueCode.custom,
      message: "Phone number must be exactly 10 digits",
      path: ["phone"]
    });
  }

  if (normalizedReceiverWhatsapp && normalizedReceiverWhatsapp.length !== 10) {
    context.addIssue({
      code: z.ZodIssueCode.custom,
      message: "Phone number must be exactly 10 digits",
      path: ["receiverWhatsappNumber"]
    });
  }

  if (normalizedPincode && normalizedPincode.length !== 6) {
    context.addIssue({
      code: z.ZodIssueCode.custom,
      message: "Enter a valid 6-digit pincode",
      path: ["pincode"]
    });
  }

  if (!input.potentialProduct) {
    context.addIssue({
      code: z.ZodIssueCode.custom,
      message: "Select a product",
      path: ["potentialProduct"]
    });
  }
});

const lineItemPayloadSchema = z.object({
  quotationId: airtableRecordIdSchema,
  items: z
    .array(
      z.object({
        productId: airtableRecordIdSchema,
        qty: z.coerce.number().positive("Quantity should be greater than zero"),
        rate: z.union([z.coerce.number().positive("Rate should be greater than zero"), z.literal(""), z.undefined()]).optional(),
        transport: z.union([z.coerce.number().nonnegative("Freight amount cannot be negative"), z.literal(""), z.undefined()]).optional(),
        gstPercent: z.union([z.coerce.number().nonnegative("GST amount cannot be negative"), z.literal(""), z.undefined()]).optional()
      })
    )
    .min(1, "Add at least one line item")
});

const customerPayloadSchema = z.object({
  customerName: z.string().trim().min(1, "Customer name is required"),
  company: z.string().trim().optional().default(""),
  phone: optionalPhoneSchema.default(""),
  email: optionalEmailSchema.default(""),
  address: z.string().trim().optional().default(""),
  state: z.string().trim().optional().default(""),
  city: z.string().trim().optional().default(""),
  pincode: optionalPincodeSchema,
  customerType: z.string().trim().optional().default("Domestic")
});

const quotationPayloadSchema = z.object({
  enquiryId: airtableRecordIdSchema
});

const orderPayloadSchema = z.object({
  quotationId: airtableRecordIdSchema,
  customerId: optionalAirtableRecordIdSchema,
  enquiryId: optionalAirtableRecordIdSchema,
  orderDate: z.string().trim().optional().default(""),
  orderStatus: z.enum(["Confirmed", "Processing", "Shipped", "Delivered", "Cancelled"]).default("Confirmed"),
  totalAmount: z.coerce.number().nonnegative().optional().default(0),
  orderNotes: z.string().trim().optional().default(""),
  paymentStatus: z.enum(["Paid", "Pending", "Half Payment"]).default("Pending"),
  deliveryStatus: z.enum(["Confirmed", "Processing", "Shipped", "Delivered", "Cancelled"]).default("Confirmed")
});

function buildPortalLineItemFields(input: {
  quotation: AirtableRecord<QuotationFields>;
  productLookup: Map<string, AirtableRecord<ProductFields>>;
  items: Array<z.infer<typeof lineItemPayloadSchema>["items"][number]>;
  existingCount: number;
}) {
  let nextLineNo = input.existingCount + 1;
  const quotationIdentifier =
    String(input.quotation.fields["Quotation Number"] || input.quotation.id).trim() || input.quotation.id;

  return input.items.map((item) => {
    const product = input.productLookup.get(item.productId);
    if (!product) {
      throw new Error("One or more selected products could not be found.");
    }

    const qty = Number(item.qty || 0);
    const rate = Number(
      parseOptionalAmount(item.rate, Number(product.fields["Bulk Sale Price"] || product.fields.MRP || 0)).toFixed(2)
    );
    const transport = Number(
      parseOptionalAmount(
        item.transport,
        Number(product.fields["Freight Amount"] || product.fields.Freight || product.fields["Pkg & Transport"] || 0)
      ).toFixed(2)
    );
    const gstPercent = Number(
      parseOptionalAmount(item.gstPercent, Number(product.fields["GST Amount"] || product.fields["GST %"] || 0)).toFixed(2)
    );
    const unitValue = Number((rate * qty).toFixed(2));
    const freightAmount = Number((transport * qty).toFixed(2));
    const gstAmount = Number((gstPercent * qty).toFixed(2));
    const totalAmount = Number((unitValue + freightAmount + gstAmount).toFixed(2));

    return {
      Name: `${quotationIdentifier}-${String(nextLineNo).padStart(2, "0")}`,
      Quotation: linkedRecordIds(input.quotation.id),
      "Linked Product": linkedRecordIds(item.productId),
      "Line No.": nextLineNo++,
      "Description Override": buildDescription(product),
      Qty: qty,
      "Rate Per Unit": rate,
      "Pkg & Transport": transport,
      "GST %": gstPercent,
      "GST Amount": gstAmount,
      "Total Amount": totalAmount,
      "Unit Value": unitValue
    };
  });
}

function buildDescription(product: AirtableRecord<ProductFields>) {
  return (
    product.fields.Narration ||
    [product.fields["Product Name"], product.fields.Model].filter(Boolean).join(" - ") ||
    product.fields["Product Key"] ||
    "Quotation item"
  );
}

function normalizeStatusForEnquiry(linkedCustomerId: string) {
  return linkedCustomerId ? "Parsed" : "New";
}

function normalizePhone(value: unknown) {
  const digits = String(value ?? "").replace(/\D/g, "");
  if (!digits) {
    return "";
  }

  return digits.length > 10 ? digits.slice(-10) : digits;
}

function normalizeEmail(value: unknown) {
  return String(value ?? "").trim().toLowerCase();
}

function toPincodeNumber(value: string, fallback?: unknown) {
  const normalized = String(value || fallback || "").replace(/\D/g, "").slice(0, 6);
  return normalized ? Number(normalized) : undefined;
}

function toPincodeString(value: string, fallback?: unknown) {
  const normalized = String(value || fallback || "").replace(/\D/g, "").slice(0, 6);
  return normalized || undefined;
}

function withStringPincodeFields(
  fields: Record<string, unknown>,
  mainPincode: string | undefined,
  destinationPincode: string | undefined
) {
  const next = { ...fields };

  if (mainPincode) {
    next.Pincode = mainPincode;
  }

  if (destinationPincode) {
    next["Destination Pincode"] = destinationPincode;
  }

  return next;
}

function omitFieldFromRecords(records: Record<string, unknown>[], fieldName: string) {
  return records.map((record) => {
    const next = { ...record };
    delete next[fieldName];
    return next;
  });
}

function mentionsField(errorMessage: string, fieldName: string) {
  return errorMessage.toLowerCase().includes(fieldName.toLowerCase());
}

function stripOptionalFields(
  fields: Record<string, unknown>,
  fieldNames: readonly string[],
  errorMessage: string
) {
  const next = { ...fields };
  let removedAny = false;

  fieldNames.forEach((fieldName) => {
    if (mentionsField(errorMessage, fieldName)) {
      delete next[fieldName];
      removedAny = true;
    }
  });

  return removedAny ? next : null;
}

function withOptionalNumberField(
  fields: Record<string, unknown>,
  fieldName: string,
  value: number | undefined
) {
  if (typeof value === "number" && Number.isFinite(value)) {
    fields[fieldName] = value;
  }

  return fields;
}

function withOptionalStringField(
  fields: Record<string, unknown>,
  fieldName: string,
  value: string | undefined
) {
  if (value) {
    fields[fieldName] = value;
  }

  return fields;
}

function toNumericValue(value: unknown, fallback = 0) {
  const raw = Array.isArray(value) ? value[0] : value;
  const normalized = String(raw ?? "")
    .replace(/,/g, "")
    .replace(/[^\d.-]/g, "")
    .trim();

  if (!normalized) {
    return fallback;
  }

  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : fallback;
}

async function findExistingCustomerByContact(input: z.infer<typeof enquiryPayloadSchema>) {
  const phone = normalizePhone(input.phone);
  const email = normalizeEmail(input.email);

  if (!phone && !email) {
    return null;
  }

  const customers = await listRecords<CustomerFields>(env.AIRTABLE_CUSTOMERS_TABLE, {
    fields: [
      "Client ID",
      "Customer Name",
      "Company",
      "Phone",
      "WhatsApp",
      "Email",
      "Address",
      "State",
      "City",
      "Pincode"
    ],
    maxRecords: 500
  });

  return (
    customers.find((customer) => {
      const customerPhone = normalizePhone(customer.fields.Phone || customer.fields.WhatsApp);
      return customerPhone && customerPhone === phone;
    }) ||
    customers.find((customer) => {
      const customerEmail = normalizeEmail(customer.fields.Email);
      return customerEmail && customerEmail === email;
    }) ||
    null
  );
}

async function resolvePotentialProductSummary(productId: string) {
  if (!productId) {
    return {
      productId: "",
      productName: ""
    };
  }

  const product = await getRecord<ProductFields>(env.AIRTABLE_PRODUCTS_TABLE, productId);
  const productName = String(product.fields["Product Name"] || "").trim();
  const model = String(product.fields.Model || "").trim();
  const narration = String(product.fields.Narration || "").trim();
  const productKey = String(product.fields["Product Key"] || "").trim();

  const summary =
    [model, narration].filter(Boolean).join(" - ") ||
    [productName, narration].filter(Boolean).join(" - ") ||
    model ||
    productName ||
    productKey;

  return {
    productId,
    productName: summary
  };
}

function parseOptionalAmount(value: unknown, fallback = 0) {
  if (value === "" || value === undefined || value === null) {
    return fallback;
  }

  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function sanitizeForAudit<T extends Record<string, unknown>>(fields: T) {
  return JSON.parse(JSON.stringify(fields)) as T;
}

async function createChangeLogEntry(
  tableName: string,
  input: {
    entityRecordId: string;
    entityLabel: string;
    action: string;
    before: Record<string, unknown> | null;
    after: Record<string, unknown>;
    actor?: AuthenticatedUser | null;
  }
) {
  try {
    await createRecord<ChangeLogFields>(tableName, {
      "Changed At": new Date().toISOString(),
      Action: input.action,
      "Changed By Name": input.actor?.name || "",
      "Changed By Email": input.actor?.email || "",
      "Entity Record ID": input.entityRecordId,
      "Entity Label": input.entityLabel,
      "Before JSON": input.before ? JSON.stringify(sanitizeForAudit(input.before), null, 2) : "",
      "After JSON": JSON.stringify(sanitizeForAudit(input.after), null, 2)
    });
  } catch (error) {
    console.warn(`[change-log] skipped ${tableName}`, error instanceof Error ? error.message : error);
  }
}

function buildBuyerBlock(fields: {
  leadName?: string;
  company?: string;
  address?: string;
  city?: string;
  state?: string;
  pincode?: string;
  phone?: string;
  email?: string;
}) {
  return [
    fields.leadName || "",
    fields.company || "",
    fields.address || "",
    [fields.city || "", fields.state || "", fields.pincode || ""].filter(Boolean).join(", "),
    fields.phone || "",
    fields.email || ""
  ]
    .map((value) => String(value).trim())
    .filter(Boolean)
    .join("\n");
}

function safeLinkedValue(recordId: string) {
  return linkedRecordIds(recordId);
}

async function updateRecordWithOptionalFieldFallback<T extends Record<string, unknown>>(
  tableName: string,
  payload: { id: string; fields: Record<string, unknown> },
  optionalFields: readonly string[]
) {
  try {
    return await updateRecord<T>(tableName, payload);
  } catch (error) {
    const message = error instanceof Error ? error.message : "";
    const retryFields = stripOptionalFields(payload.fields, optionalFields, message);
    if (!retryFields) {
      throw error;
    }

    return updateRecord<T>(tableName, {
      id: payload.id,
      fields: retryFields
    });
  }
}

async function createQuotationShellForEnquiry(
  enquiry: AirtableRecord<EnquiryFields>,
  customerId: string
) {
  const existingQuotationId = enquiry.fields.Quotations?.[0];
  if (existingQuotationId) {
    return getRecord<QuotationFields>(env.AIRTABLE_QUOTATIONS_TABLE, existingQuotationId);
  }

  const quotationReference = enquiry.fields["Enquiry ID"] || enquiry.id;
  const quotationNumberResult = await createRecordWithUniqueNumber<QuotationFields, AirtableRecord<QuotationFields>>({
    tableName: env.AIRTABLE_QUOTATIONS_TABLE,
    fieldName: "Quotation Number",
    prefix: "QTN",
    create: async (quotationNumber) => {
      const fields = {
        "Quotation Number": quotationNumber,
        "Logged Date Time": new Date().toISOString(),
        "Linked Enquiry": safeLinkedValue(enquiry.id),
        "Linked Customer": safeLinkedValue(customerId),
        Status: "Parsed",
        "Draft Format": "XLSX",
        "Drive Folder URL": enquiry.fields["Drive Folder URL"] || "",
        "Reference Number": quotationReference,
        "Buyer Block": buildBuyerBlock({
          leadName: enquiry.fields["Lead Name"],
          company: enquiry.fields.Company,
          address: enquiry.fields.Address,
          city: enquiry.fields.City,
          state: enquiry.fields.State,
          pincode: String(enquiry.fields.Pincode || ""),
          phone: enquiry.fields.Phone,
          email: enquiry.fields.Email
        }),
        "Send Quotation": false,
        "Send Reminder": false,
        "Mark Accepted": false,
        "Mark Rejected": false,
        "Reminder Count": 0
      };

      try {
        return await createRecord<QuotationFields>(env.AIRTABLE_QUOTATIONS_TABLE, fields);
      } catch (error) {
        const message = error instanceof Error ? error.message : "";
        const retryFields = stripOptionalFields(fields, optionalQuotationFields, message);
        if (!retryFields) {
          throw error;
        }

        return createRecord<QuotationFields>(env.AIRTABLE_QUOTATIONS_TABLE, retryFields);
      }
    }
  });

  await updateRecord<EnquiryFields>(env.AIRTABLE_ENQUIRIES_TABLE, {
    id: enquiry.id,
    fields: {
      Quotations: safeLinkedValue(quotationNumberResult.id)
    }
  });

  return quotationNumberResult;
}

export async function loadQuotationLineItemMetrics(quotationId: string) {
  const items = await getPortalQuotationLineItems(quotationId);
  return {
    items,
    lineItemCount: items.length,
    quotationGrandTotal: Number(items.reduce((sum, item) => sum + Number(item.totalAmount || 0), 0).toFixed(2)),
    orderValuePerItem: items
      .map((item, index) => `${index + 1}. ${formatLineItemValue(item.totalAmount)}`)
      .join("\n")
  };
}

function formatLineItemValue(value: number) {
  return Number.isFinite(value) ? value.toFixed(2) : "0.00";
}

async function nextCustomerIdentifier() {
  const customers = await listRecords<CustomerFields>(env.AIRTABLE_CUSTOMERS_TABLE, {
    fields: ["Client ID"],
    maxRecords: 1000
  });

  const maxValue = customers.reduce((currentMax, customer) => {
    const match = String(customer.fields["Client ID"] || "").match(/^CUST-(\d+)$/);
    return Math.max(currentMax, match ? Number(match[1]) : 0);
  }, 0);

  return `CUST-${String(maxValue + 1).padStart(3, "0")}`;
}

async function nextOrderIdentifier() {
  const orders = await listRecords<OrderFields>(env.AIRTABLE_ORDERS_TABLE, {
    fields: ["Order Number"],
    maxRecords: 1000
  });

  const maxValue = orders.reduce((currentMax, order) => {
    const match = String(order.fields["Order Number"] || "").match(/^ORD-(\d+)$/);
    return Math.max(currentMax, match ? Number(match[1]) : 0);
  }, 0);

  return `ORD-${String(maxValue + 1).padStart(3, "0")}`;
}

export async function createPortalEnquiry(payload: unknown) {
  const input = enquiryPayloadSchema.parse(payload);
  const productSummary = await resolvePotentialProductSummary(input.potentialProduct);
  const existingCustomer = input.linkedCustomerId
    ? await getRecord<CustomerFields>(env.AIRTABLE_CUSTOMERS_TABLE, input.linkedCustomerId)
    : await findExistingCustomerByContact(input);
  const linkedCustomerId = input.linkedCustomerId || existingCustomer?.id || "";
  const parserStatus = normalizeStatusForEnquiry(linkedCustomerId);
  const destinationAddress = input.destinationAddress || input.address || "";
  const destinationState = input.destinationState || input.state || "";
  const destinationCity = input.destinationCity || input.city || "";
  const destinationPincode = input.destinationPincode || input.pincode || "";
  const mainPincode = toPincodeNumber(input.pincode, existingCustomer?.fields.Pincode);
  const destinationPincodeNumber = toPincodeNumber(destinationPincode);
  const mainPincodeText = toPincodeString(input.pincode, existingCustomer?.fields.Pincode);
  const destinationPincodeText = toPincodeString(destinationPincode);
  const loggedDateTime = new Date().toISOString();
  const enquiryFields = withOptionalNumberField(
    withOptionalNumberField(
      {
        "Logged Date Time": loggedDateTime,
        "Lead Name": input.leadName || existingCustomer?.fields["Customer Name"] || "",
        Company: input.company || existingCustomer?.fields.Company || "",
        Phone: input.phone || existingCustomer?.fields.Phone || "",
        Email: input.email || existingCustomer?.fields.Email || "",
        Address: input.address || existingCustomer?.fields.Address || "",
        State: input.state || existingCustomer?.fields.State || "",
        City: input.city || existingCustomer?.fields.City || "",
        "Destination Address": destinationAddress,
        "Destination State": destinationState,
        "Destination City": destinationCity,
        "Parser Status": parserStatus,
        "Linked Customer": linkedRecordIds(linkedCustomerId),
        "Requirement Summary": productSummary.productName || input.requirementSummary,
        "Potential Product": productSummary.productId,
        "Receiver WhatsApp Number": input.receiverWhatsappNumber
      },
      "Pincode",
      mainPincode
    ),
    "Destination Pincode",
    destinationPincodeNumber
  );

  console.log("[enquiry-create] preparing Airtable write", {
    leadName: input.leadName,
    mainPincodeNumber: mainPincode,
    mainPincodeText,
    destinationPincodeNumber,
    destinationPincodeText
  });

  const created = await createRecordWithUniqueNumber<EnquiryFields, AirtableRecord<EnquiryFields>>({
    tableName: env.AIRTABLE_ENQUIRIES_TABLE,
    fieldName: "Enquiry ID",
    prefix: "ENQ",
    create: async (enquiryNumber) => {
      const fieldsWithId: Record<string, unknown> = {
        "Enquiry ID": enquiryNumber,
        ...enquiryFields
      };

      try {
        return await createRecord<EnquiryFields>(env.AIRTABLE_ENQUIRIES_TABLE, fieldsWithId);
      } catch (error) {
        const message = error instanceof Error ? error.message : "";
        console.log("[enquiry-create] Airtable create failed", {
          message,
          pincode: fieldsWithId.Pincode,
          destinationPincode: fieldsWithId["Destination Pincode"]
        });
        let retryFields: Record<string, unknown> | null = stripOptionalFields(
          fieldsWithId,
          optionalEnquiryFields,
          message
        );

        if (mentionsField(message, "Pincode")) {
          retryFields = withStringPincodeFields(
            retryFields ?? fieldsWithId,
            mainPincodeText,
            destinationPincodeText
          );
          console.log("[enquiry-create] retrying with string pincodes", {
            pincode: retryFields.Pincode,
            destinationPincode: retryFields["Destination Pincode"]
          });
        }

        if (!retryFields) {
          throw error;
        }

        return createRecord<EnquiryFields>(env.AIRTABLE_ENQUIRIES_TABLE, retryFields);
      }
    }
  });

  const provisioned = await createCustomerForEnquiry(created.id);

  return {
    enquiryRecordId: provisioned.enquiry.id,
    enquiryId: provisioned.enquiry.fields["Enquiry ID"] || created.fields["Enquiry ID"] || "",
    parserStatus: provisioned.enquiry.fields["Parser Status"] || created.fields["Parser Status"] || parserStatus,
    linkedCustomerId: provisioned.customer.id,
    quotationRecordId: provisioned.quotation.id,
    quotationNumber: provisioned.quotation.fields["Quotation Number"] || "",
    driveFolderUrl: provisioned.folder?.folderUrl || provisioned.customer.fields["Drive Folder URL"] || ""
  };
}

export async function updatePortalEnquiry(enquiryId: string, payload: unknown) {
  const input = enquiryPayloadSchema.parse(payload);
  const productSummary = await resolvePotentialProductSummary(input.potentialProduct);
  const existingEnquiry = await getRecord<EnquiryFields>(env.AIRTABLE_ENQUIRIES_TABLE, enquiryId);
  const existingCustomer = input.linkedCustomerId
    ? await getRecord<CustomerFields>(env.AIRTABLE_CUSTOMERS_TABLE, input.linkedCustomerId)
    : existingEnquiry.fields["Linked Customer"]?.[0]
      ? await getRecord<CustomerFields>(
          env.AIRTABLE_CUSTOMERS_TABLE,
          existingEnquiry.fields["Linked Customer"]?.[0] || ""
        )
      : await findExistingCustomerByContact(input);
  const linkedCustomerId =
    input.linkedCustomerId || existingEnquiry.fields["Linked Customer"]?.[0] || existingCustomer?.id || "";
  const parserStatus = existingEnquiry.fields["Parser Status"] || normalizeStatusForEnquiry(linkedCustomerId);
  const destinationAddress = input.destinationAddress || input.address || "";
  const destinationState = input.destinationState || input.state || "";
  const destinationCity = input.destinationCity || input.city || "";
  const destinationPincode = input.destinationPincode || input.pincode || "";
  const mainPincode = toPincodeNumber(input.pincode, existingCustomer?.fields.Pincode);
  const destinationPincodeNumber = toPincodeNumber(destinationPincode);
  const mainPincodeText = toPincodeString(input.pincode, existingCustomer?.fields.Pincode);
  const destinationPincodeText = toPincodeString(destinationPincode);

  const enquiryFields = withOptionalNumberField(
    withOptionalNumberField(
      {
        "Lead Name": input.leadName || existingCustomer?.fields["Customer Name"] || existingEnquiry.fields["Lead Name"] || "",
        Company: input.company || existingCustomer?.fields.Company || existingEnquiry.fields.Company || "",
        Phone: input.phone || existingCustomer?.fields.Phone || existingEnquiry.fields.Phone || "",
        Email: input.email || existingCustomer?.fields.Email || existingEnquiry.fields.Email || "",
        Address: input.address || existingCustomer?.fields.Address || existingEnquiry.fields.Address || "",
        State: input.state || existingCustomer?.fields.State || existingEnquiry.fields.State || "",
        City: input.city || existingCustomer?.fields.City || existingEnquiry.fields.City || "",
        "Destination Address": destinationAddress,
        "Destination State": destinationState,
        "Destination City": destinationCity,
        "Parser Status": parserStatus,
        "Linked Customer": linkedRecordIds(linkedCustomerId),
        "Requirement Summary": productSummary.productName || input.requirementSummary || existingEnquiry.fields["Requirement Summary"] || "",
        "Potential Product": productSummary.productId || existingEnquiry.fields["Potential Product"] || "",
        "Receiver WhatsApp Number": input.receiverWhatsappNumber
      },
      "Pincode",
      mainPincode
    ),
    "Destination Pincode",
    destinationPincodeNumber
  );

  console.log("[enquiry-update] preparing Airtable write", {
    enquiryId,
    mainPincodeNumber: mainPincode,
    mainPincodeText,
    destinationPincodeNumber,
    destinationPincodeText
  });

  let updatedEnquiry: AirtableRecord<EnquiryFields>;
  try {
    updatedEnquiry = await updateRecord<EnquiryFields>(env.AIRTABLE_ENQUIRIES_TABLE, {
      id: enquiryId,
      fields: enquiryFields
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : "";
    console.log("[enquiry-update] Airtable update failed", {
      enquiryId,
      message,
      pincode: enquiryFields.Pincode,
      destinationPincode: enquiryFields["Destination Pincode"]
    });
    let retryFields: Record<string, unknown> | null = stripOptionalFields(
      enquiryFields,
      optionalEnquiryFields,
      message
    );

    if (mentionsField(message, "Pincode")) {
      retryFields = withStringPincodeFields(
        retryFields ?? enquiryFields,
        mainPincodeText,
        destinationPincodeText
      );
      console.log("[enquiry-update] retrying with string pincodes", {
        enquiryId,
        pincode: retryFields.Pincode,
        destinationPincode: retryFields["Destination Pincode"]
      });
    }

    if (!retryFields) {
      throw error;
    }

    updatedEnquiry = await updateRecord<EnquiryFields>(env.AIRTABLE_ENQUIRIES_TABLE, {
      id: enquiryId,
      fields: retryFields
    });
  }

  const provisioned = await createCustomerForEnquiry(updatedEnquiry.id);

  return {
    enquiryRecordId: provisioned.enquiry.id,
    enquiryId: provisioned.enquiry.fields["Enquiry ID"] || updatedEnquiry.fields["Enquiry ID"] || "",
    parserStatus: provisioned.enquiry.fields["Parser Status"] || updatedEnquiry.fields["Parser Status"] || parserStatus,
    linkedCustomerId: provisioned.customer.id,
    quotationRecordId: provisioned.quotation.id,
    quotationNumber: provisioned.quotation.fields["Quotation Number"] || "",
    driveFolderUrl: provisioned.folder?.folderUrl || provisioned.customer.fields["Drive Folder URL"] || ""
  };
}

export async function createPortalQuotationLineItems(payload: unknown) {
  const input = lineItemPayloadSchema.parse(payload);

  const quotation = await getRecord<QuotationFields>(env.AIRTABLE_QUOTATIONS_TABLE, input.quotationId);
  const products = await listRecords<ProductFields>(env.AIRTABLE_PRODUCTS_TABLE, {
    maxRecords: 500
  });
  const productLookup = new Map(products.map((product) => [product.id, product]));

  const existingItems = await listRecords<QuotationLineItemFields>(env.AIRTABLE_QUOTATION_LINE_ITEMS_TABLE, {
    fields: ["Quotation"],
    maxRecords: 500
  });

  const existingCount = existingItems.filter((item) => item.fields.Quotation?.includes(input.quotationId)).length;
  const lineItemFields = buildPortalLineItemFields({
    quotation,
    productLookup,
    items: input.items,
    existingCount
  });

  let created;
  try {
    created = await createRecords<QuotationLineItemFields>(
      env.AIRTABLE_QUOTATION_LINE_ITEMS_TABLE,
      lineItemFields
    );
  } catch (error) {
    const message = error instanceof Error ? error.message : "";
    if (!mentionsField(message, "Pkg & Transport")) {
      throw error;
    }

    created = await createRecords<QuotationLineItemFields>(
      env.AIRTABLE_QUOTATION_LINE_ITEMS_TABLE,
      omitFieldFromRecords(lineItemFields, "Pkg & Transport")
    );
  }
  const draft = await refreshDraftForQuotation(quotation.id);
  if (draft.quotation.fields["Final PDF URL"] || ["Approved", "Draft Sent", "Sent"].includes(String(draft.quotation.fields.Status || ""))) {
    await generateFinalPdfForQuotation(quotation.id);
  }
  console.log("[portal-line-items] draft refresh result", {
    quotationId: quotation.id,
    quotationStatus: draft.quotation.fields.Status || "",
    draftFileUrl: draft.quotation.fields["Draft File URL"] || "",
    driveFolderUrl: draft.folder.folderUrl || draft.quotation.fields["Drive Folder URL"] || ""
  });

  return {
    quotationId: quotation.id,
    quotationNumber: draft.quotation.fields["Quotation Number"] || quotation.fields["Quotation Number"] || quotation.id,
    createdCount: created.length,
    quotationStatus: draft.quotation.fields.Status || "Ready for Review",
    draftFileUrl: draft.quotation.fields["Draft File URL"] || "",
    driveFolderUrl: draft.folder.folderUrl || draft.quotation.fields["Drive Folder URL"] || ""
  };
}

export async function getPortalQuotationLineItems(quotationId: string) {
  const items = await listRecords<QuotationLineItemFields>(env.AIRTABLE_QUOTATION_LINE_ITEMS_TABLE, {
    fields: [
      "Quotation",
      "Linked Product",
      "Qty",
      "Rate Per Unit",
      "Pkg & Transport",
      "GST %",
      "Total Amount"
    ],
    maxRecords: 500
  });

  return items
    .filter((item) => item.fields.Quotation?.includes(quotationId))
    .sort((left, right) => Number(left.fields["Line No."] || 0) - Number(right.fields["Line No."] || 0))
    .map((item) => ({
      id: item.id,
      productId: item.fields["Linked Product"]?.[0] || "",
      qty: Number(item.fields.Qty || 0),
      rate: Number(item.fields["Rate Per Unit"] || 0),
      transport: Number(item.fields["Pkg & Transport"] || 0),
      gstPercent: Number(item.fields["GST %"] || 0),
      totalAmount: Number(item.fields["Total Amount"] || 0)
    })) satisfies PortalQuotationLineItem[];
}

export async function replacePortalQuotationLineItems(payload: unknown) {
  const input = lineItemPayloadSchema.parse(payload);
  const quotation = await getRecord<QuotationFields>(env.AIRTABLE_QUOTATIONS_TABLE, input.quotationId);
  const products = await listRecords<ProductFields>(env.AIRTABLE_PRODUCTS_TABLE, {
    maxRecords: 500
  });
  const productLookup = new Map(products.map((product) => [product.id, product]));
  const existingItems = await listRecords<QuotationLineItemFields>(env.AIRTABLE_QUOTATION_LINE_ITEMS_TABLE, {
    fields: ["Quotation"],
    maxRecords: 500
  });
  const existingForQuotation = existingItems.filter((item) => item.fields.Quotation?.includes(input.quotationId));

  if (existingForQuotation.length) {
    await deleteRecords(
      env.AIRTABLE_QUOTATION_LINE_ITEMS_TABLE,
      existingForQuotation.map((item) => item.id)
    );
  }

  const lineItemFields = buildPortalLineItemFields({
    quotation,
    productLookup,
    items: input.items,
    existingCount: 0
  });

  let created;
  try {
    created = await createRecords<QuotationLineItemFields>(
      env.AIRTABLE_QUOTATION_LINE_ITEMS_TABLE,
      lineItemFields
    );
  } catch (error) {
    const message = error instanceof Error ? error.message : "";
    if (!mentionsField(message, "Pkg & Transport")) {
      throw error;
    }

    created = await createRecords<QuotationLineItemFields>(
      env.AIRTABLE_QUOTATION_LINE_ITEMS_TABLE,
      omitFieldFromRecords(lineItemFields, "Pkg & Transport")
    );
  }

  const draft = await refreshDraftForQuotation(quotation.id);
  if (draft.quotation.fields["Final PDF URL"] || ["Approved", "Draft Sent", "Sent"].includes(String(draft.quotation.fields.Status || ""))) {
    await generateFinalPdfForQuotation(quotation.id);
  }

  return {
    quotationId: quotation.id,
    quotationNumber: draft.quotation.fields["Quotation Number"] || quotation.fields["Quotation Number"] || quotation.id,
    createdCount: created.length,
    quotationStatus: draft.quotation.fields.Status || "Ready for Review",
    draftFileUrl: draft.quotation.fields["Draft File URL"] || "",
    driveFolderUrl: draft.folder.folderUrl || draft.quotation.fields["Drive Folder URL"] || ""
  };
}

export async function generateDraftForEnquiry(enquiryId: string) {
  const provisioned = await createCustomerForEnquiry(enquiryId);
  try {
    const draft = await refreshDraftForQuotation(provisioned.quotation.id);

    return {
      enquiryId: provisioned.enquiry.id,
      quotationId: draft.quotation.id,
      quotationNumber: draft.quotation.fields["Quotation Number"] || provisioned.quotation.fields["Quotation Number"] || "",
      draftFileUrl: draft.quotation.fields["Draft File URL"] || "",
      driveFolderUrl: draft.folder.folderUrl || draft.quotation.fields["Drive Folder URL"] || "",
      generated: true,
      message: "Draft generated successfully."
    };
  } catch (error) {
    const message = error instanceof Error ? error.message : "Failed to generate draft";
    if (!message.toLowerCase().includes("line item")) {
      throw error;
    }

    return {
      enquiryId: provisioned.enquiry.id,
      quotationId: provisioned.quotation.id,
      quotationNumber: provisioned.quotation.fields["Quotation Number"] || "",
      draftFileUrl: "",
      driveFolderUrl: provisioned.folder?.folderUrl || provisioned.quotation.fields["Drive Folder URL"] || "",
      generated: false,
      message: "Quotation shell is ready. Add line items to generate the draft."
    };
  }
}

export async function createPortalCustomer(payload: unknown, actor?: AuthenticatedUser | null) {
  const input = customerPayloadSchema.parse(payload);
  const pincodeNumber = toPincodeNumber(input.pincode);
  const pincodeText = toPincodeString(input.pincode);
  const clientId = await nextCustomerIdentifier();
  const fields = withOptionalNumberField(
    {
      "Client ID": clientId,
      "Customer Name": input.customerName,
      Company: input.company,
      Phone: normalizePhone(input.phone),
      WhatsApp: normalizePhone(input.phone),
      Email: normalizeEmail(input.email),
      Address: input.address,
      State: input.state,
      City: input.city,
      "Customer Type": input.customerType || "Domestic"
    },
    "Pincode",
    pincodeNumber
  );

  let customer: AirtableRecord<CustomerFields>;
  try {
    customer = await createRecord<CustomerFields>(env.AIRTABLE_CUSTOMERS_TABLE, fields);
  } catch (error) {
    const message = error instanceof Error ? error.message : "";
    if (!mentionsField(message, "Pincode")) {
      throw error;
    }

    customer = await createRecord<CustomerFields>(
      env.AIRTABLE_CUSTOMERS_TABLE,
      withOptionalStringField({ ...fields }, "Pincode", pincodeText)
    );
  }

  await createChangeLogEntry(env.AIRTABLE_CUSTOMER_CHANGE_LOG_TABLE, {
    entityRecordId: customer.id,
    entityLabel: customer.fields["Client ID"] || customer.fields["Customer Name"] || customer.id,
    action: "created",
    before: null,
    after: customer.fields,
    actor
  });

  return customer;
}

export async function updatePortalCustomer(customerId: string, payload: unknown, actor?: AuthenticatedUser | null) {
  const input = customerPayloadSchema.parse(payload);
  const existing = await getRecord<CustomerFields>(env.AIRTABLE_CUSTOMERS_TABLE, customerId);
  const pincodeNumber = toPincodeNumber(input.pincode);
  const pincodeText = toPincodeString(input.pincode);
  const fields = withOptionalNumberField(
    {
      "Customer Name": input.customerName,
      Company: input.company,
      Phone: normalizePhone(input.phone),
      WhatsApp: normalizePhone(input.phone),
      Email: normalizeEmail(input.email),
      Address: input.address,
      State: input.state,
      City: input.city,
      "Customer Type": input.customerType || "Domestic"
    },
    "Pincode",
    pincodeNumber
  );

  let updated: AirtableRecord<CustomerFields>;
  try {
    updated = await updateRecord<CustomerFields>(env.AIRTABLE_CUSTOMERS_TABLE, { id: customerId, fields });
  } catch (error) {
    const message = error instanceof Error ? error.message : "";
    if (!mentionsField(message, "Pincode")) {
      throw error;
    }

    updated = await updateRecord<CustomerFields>(env.AIRTABLE_CUSTOMERS_TABLE, {
      id: customerId,
      fields: withOptionalStringField({ ...fields }, "Pincode", pincodeText)
    });
  }

  await createChangeLogEntry(env.AIRTABLE_CUSTOMER_CHANGE_LOG_TABLE, {
    entityRecordId: updated.id,
    entityLabel: updated.fields["Client ID"] || updated.fields["Customer Name"] || updated.id,
    action: "updated",
    before: existing.fields,
    after: updated.fields,
    actor
  });

  return updated;
}

export async function createPortalQuotation(payload: unknown, actor?: AuthenticatedUser | null) {
  const input = quotationPayloadSchema.parse(payload);
  const enquiry = await getRecord<EnquiryFields>(env.AIRTABLE_ENQUIRIES_TABLE, input.enquiryId);
  const provisioned = await createCustomerForEnquiry(enquiry.id);
  const quotation = await createQuotationShellForEnquiry(provisioned.enquiry, provisioned.customer.id);

  await createChangeLogEntry(env.AIRTABLE_QUOTATION_CHANGE_LOG_TABLE, {
    entityRecordId: quotation.id,
    entityLabel: quotation.fields["Quotation Number"] || quotation.id,
    action: "created",
    before: null,
    after: quotation.fields,
    actor
  });

  return {
    enquiry: provisioned.enquiry,
    customer: provisioned.customer,
    quotation
  };
}

export async function markQuotationAsSent(quotationId: string, actor?: AuthenticatedUser | null) {
  const existing = await getRecord<QuotationFields>(env.AIRTABLE_QUOTATIONS_TABLE, quotationId);
  const updated = await updateRecordWithOptionalFieldFallback<QuotationFields>(
    env.AIRTABLE_QUOTATIONS_TABLE,
    {
      id: quotationId,
      fields: {
        Status: "Sent",
        "Sent Date": new Date().toISOString()
      }
    },
    optionalQuotationFields
  );

  await createChangeLogEntry(env.AIRTABLE_QUOTATION_CHANGE_LOG_TABLE, {
    entityRecordId: updated.id,
    entityLabel: updated.fields["Quotation Number"] || updated.id,
    action: "marked_sent",
    before: existing.fields,
    after: updated.fields,
    actor
  });

  return updated;
}

export async function createOrderFromQuotation(quotationId: string, actor?: AuthenticatedUser | null) {
  const quotation = await getRecord<QuotationFields>(env.AIRTABLE_QUOTATIONS_TABLE, quotationId);
  const existingOrders = await listRecords<OrderFields>(env.AIRTABLE_ORDERS_TABLE, {
    fields: ["Order Number", "Linked Quotation", "Quotation", "Customer", "Enquiries", "Order Status", "Order Date", "Total Amount"],
    maxRecords: 500
  });
  const existing = existingOrders.find((order) =>
    order.fields["Linked Quotation"]?.includes(quotationId) ||
    order.fields.Quotation === quotationId ||
    (Array.isArray(order.fields.Quotation) && order.fields.Quotation.includes(quotationId))
  );
  if (existing) {
    return existing;
  }

  const customerId = quotation.fields["Linked Customer"]?.[0] || "";
  const enquiryId = quotation.fields["Linked Enquiry"]?.[0] || "";
  const metrics = await loadQuotationLineItemMetrics(quotationId);
  const orderNumber = await nextOrderIdentifier();
  const created = await createRecord<OrderFields>(env.AIRTABLE_ORDERS_TABLE, {
    "Order Number": orderNumber,
    "Order Date": new Date().toISOString(),
    Quotation: quotationId,
    Customer: customerId,
    Enquiries: enquiryId,
    "Linked Quotation": safeLinkedValue(quotationId),
    "Linked Customer": safeLinkedValue(customerId),
    "Order Status": "Confirmed",
    "Total Amount": metrics.quotationGrandTotal,
    "Quotation Grand Total": metrics.quotationGrandTotal,
    "Quotation Status": quotation.fields.Status || "",
    "Order Line Item Count": metrics.lineItemCount,
    "Order Value per Item": metrics.orderValuePerItem,
    "Order Fulfillment Progress": "0%",
    "Order Notes": "",
    "Order Value": metrics.quotationGrandTotal,
    "Payment Status": "Pending",
    "Delivery Status": "Confirmed"
  });

  await createChangeLogEntry(env.AIRTABLE_ORDER_CHANGE_LOG_TABLE, {
    entityRecordId: created.id,
    entityLabel: created.fields["Order Number"] || created.id,
    action: "created",
    before: null,
    after: created.fields,
    actor
  });

  return created;
}

export async function updatePortalOrder(orderId: string, payload: unknown, actor?: AuthenticatedUser | null) {
  const input = orderPayloadSchema.parse(payload);
  const existing = await getRecord<OrderFields>(env.AIRTABLE_ORDERS_TABLE, orderId);
  const quotation = await getRecord<QuotationFields>(env.AIRTABLE_QUOTATIONS_TABLE, input.quotationId);
  const metrics = await loadQuotationLineItemMetrics(input.quotationId);
  const customerId = input.customerId || quotation.fields["Linked Customer"]?.[0] || "";
  const enquiryId = input.enquiryId || quotation.fields["Linked Enquiry"]?.[0] || "";
  const fields: Record<string, unknown> = {
    "Order Date": input.orderDate || existing.fields["Order Date"] || new Date().toISOString(),
    Quotation: input.quotationId,
    Customer: customerId,
    Enquiries: enquiryId,
    "Linked Quotation": safeLinkedValue(input.quotationId),
    "Linked Customer": safeLinkedValue(customerId),
    "Order Status": input.orderStatus,
    "Total Amount": input.totalAmount || metrics.quotationGrandTotal,
    "Order Notes": input.orderNotes,
    "Quotation Grand Total": metrics.quotationGrandTotal,
    "Quotation Status": quotation.fields.Status || "",
    "Order Line Item Count": metrics.lineItemCount,
    "Order Value per Item": metrics.orderValuePerItem,
    "Order Value": input.totalAmount || metrics.quotationGrandTotal,
    "Payment Status": input.paymentStatus,
    "Delivery Status": input.deliveryStatus
  };

  const updated = await updateRecord<OrderFields>(env.AIRTABLE_ORDERS_TABLE, {
    id: orderId,
    fields
  });

  await createChangeLogEntry(env.AIRTABLE_ORDER_CHANGE_LOG_TABLE, {
    entityRecordId: updated.id,
    entityLabel: updated.fields["Order Number"] || updated.id,
    action: "updated",
    before: existing.fields,
    after: updated.fields,
    actor
  });

  return updated;
}
