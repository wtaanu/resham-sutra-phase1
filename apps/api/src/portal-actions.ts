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
  refreshDraftForQuotation
} from "./intake-processor.js";
import { syncEnquiryToZohoBigin } from "./zoho-bigin.js";

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
  "Requirement Summary"?: string;
  "Requested Asset"?: string;
  "Potential Product"?: string;
  "Receiver WhatsApp Number"?: string;
  "Zoho Bigin Record ID"?: string;
  "Zoho Bigin Deal ID"?: string;
  "Zoho Bigin Sync Status"?: string;
  "Zoho Bigin Synced At"?: string;
  "Zoho Bigin Sync Error"?: string;
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
  "Destination Address"?: string;
  "Destination State"?: string;
  "Destination City"?: string;
  "Destination Pincode"?: string | number;
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
  Orders?: string;
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
  "Product Category"?: string;
  Category?: string;
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
  "Total Amount"?: number | string;
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
  "Total Amount"?: number | string;
  "Order Notes"?: string;
  "Line Items"?: string[] | string;
  "Quotation Grand Total"?: number | string;
  "Quotation Status"?: string;
  "Order Line Item Count"?: number;
  "Order Value per Item"?: string;
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

type OrderLineItemFields = {
  Name?: string;
  Order?: string;
  Quotation?: string;
  Enquiries?: string;
  Customer?: string;
  "Linked Product"?: string;
  "S.No."?: number;
  Description?: string;
  Qty?: number;
  "Rate Per Unit"?: number;
  "Packing & Freight"?: number;
  "Unit Value"?: number;
  "GST 18%"?: string;
  "Total Amount"?: string;
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

type PortalOrderLineItem = {
  id: string;
  productId: string;
  description: string;
  qty: number;
  ratePerUnit: number;
  packingFreight: number;
  unitValue: number;
  gst18: number;
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
const optionalEnquiryFields = [
  "Logged Date Time",
  "Receiver WhatsApp Number",
  "Zoho Bigin Record ID",
  "Zoho Bigin Deal ID",
  "Zoho Bigin Sync Status",
  "Zoho Bigin Synced At",
  "Zoho Bigin Sync Error"
] as const;
const optionalQuotationFields = [
  "Logged Date Time",
  "Draft Created Time",
  "WhatsApp Sent Date Time",
  "Email Sent Date Time",
  "Final PDF Generated At",
  "Orders"
] as const;
const optionalCustomerDestinationFields = [
  "Destination Address",
  "Destination State",
  "Destination City",
  "Destination Pincode"
] as const;
const optionalOrderFields = [
  "Total Amount",
  "Payment Terms",
  "Order Ref Number Client",
  "Line Items",
  "Quotation Grand Total",
  "Quotation Status",
  "Order Fulfillment Progress",
  "Order Value",
  "Destination Address",
  "Destination State",
  "Destination City",
  "Destination Pincode"
] as const;
const ENQUIRY_STATUS_PARSED = "Parsed";
const ENQUIRY_STATUS_MANUAL = "New Enquiries";
const ENQUIRY_STATUS_DRAFT = "Draft Quote";
const ENQUIRY_STATUS_APPROVED = "Approved Quote";
const ENQUIRY_STATUS_SENT = "Sent Quote";
const ENQUIRY_STATUS_ORDERED = "Ordered";
const QUOTATION_STATUS_DRAFT = "Draft Quote";
const QUOTATION_STATUS_APPROVED = "Approved Quote";
const QUOTATION_STATUS_SENT = "Sent Quote";
const QUOTATION_STATUS_ORDERED = "Ordered";
const ORDER_FULFILLMENT_PROGRESS_NOT_STARTED = "Not Started";
const manualEnquiryCreateInflight = new Map<string, Promise<{
  enquiryRecordId: string;
  enquiryId: string;
  parserStatus: string;
  linkedCustomerId: string;
  quotationRecordId: string;
  quotationNumber: string;
  driveFolderUrl: string;
}>>();

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
  paymentTerms: z.string().trim().optional().default(""),
  orderRefNumberClient: z.string().trim().optional().default(""),
  address: z.string().trim().optional().default(""),
  state: z.string().trim().optional().default(""),
  city: z.string().trim().optional().default(""),
  pincode: optionalPincodeSchema,
  destinationAddress: z.string().trim().optional().default(""),
  destinationState: z.string().trim().optional().default(""),
  destinationCity: z.string().trim().optional().default(""),
  destinationPincode: optionalPincodeSchema,
  items: z.array(
    z.object({
      id: z.string().trim().optional().default(""),
      productId: optionalAirtableRecordIdSchema,
      description: z.string().trim().min(1, "Description is required"),
      qty: z.coerce.number().positive("Quantity should be greater than zero"),
      ratePerUnit: z.coerce.number().nonnegative("Rate per unit cannot be negative"),
      packingFreight: z.coerce.number().nonnegative("Packing & Freight cannot be negative"),
      unitValue: z.coerce.number().nonnegative().optional().default(0),
      gst18: z.coerce.number().nonnegative("GST 18% cannot be negative"),
      totalAmount: z.coerce.number().nonnegative("Total amount cannot be negative")
    })
  ).optional().default([])
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

function normalizeStatusForEnquiry(linkedCustomerId: string, source: z.infer<typeof enquirySourceSchema>) {
  if (source === "whatsapp") {
    return ENQUIRY_STATUS_PARSED;
  }

  return linkedCustomerId ? ENQUIRY_STATUS_PARSED : ENQUIRY_STATUS_MANUAL;
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

function serializeError(error: unknown) {
  if (error instanceof Error) {
    return {
      name: error.name,
      message: error.message,
      stack: error.stack
    };
  }

  return {
    name: "NonError",
    message: String(error),
    stack: ""
  };
}

async function resolveZohoBiginContext(enquiry: AirtableRecord<EnquiryFields>) {
  const quotationId = enquiry.fields.Quotations?.[0] || "";
  const productId = String(enquiry.fields["Potential Product"] || "").trim();
  let quotationRef = "";
  let productName = "";
  let productKey = "";
  let productDescription = "";
  let productUnitPrice = 0;
  let productFreight = 0;
  let productGstRate = 0;
  let productCategory = "";
  if (quotationId) {
    try {
      const quotation = await getRecord<QuotationFields>(env.AIRTABLE_QUOTATIONS_TABLE, quotationId);
      quotationRef = String(quotation.fields["Quotation Number"] || quotation.fields["Reference Number"] || quotation.id || "").trim();
    } catch (error) {
      console.warn("[zoho-bigin] quotation ref lookup failed", {
        enquiryId: enquiry.id,
        quotationId,
        error: serializeError(error)
      });
    }
  }

  if (productId) {
    try {
      const product = await getRecord<ProductFields>(env.AIRTABLE_PRODUCTS_TABLE, productId);
      productKey = String(product.fields["Product Key"] || "").trim();
      productName = String(
        product.fields["Product Name"] ||
          [product.fields.Model, product.fields.Narration].filter(Boolean).join(" - ") ||
          productKey ||
          product.id
      ).trim();
      productDescription = String(
        product.fields.Narration ||
          [product.fields["Product Name"], product.fields.Model].filter(Boolean).join(" - ") ||
          productKey
      ).trim();
      productUnitPrice = Number(product.fields["Bulk Sale Price"] || product.fields.MRP || 0);
      productFreight = Number(product.fields["Freight Amount"] || product.fields.Freight || product.fields["Pkg & Transport"] || 0);
      productGstRate = Number(product.fields["GST %"] || product.fields["GST Amount"] || 0);
      productCategory = String(product.fields["Product Category"] || product.fields.Category || "").trim();
    } catch (error) {
      console.warn("[zoho-bigin] product lookup failed", {
        enquiryId: enquiry.id,
        productId,
        error: serializeError(error)
      });
    }
  }

  return {
    quotationRef,
    productName,
    productKey,
    productDescription,
    productUnitPrice,
    productFreight,
    productGstRate,
    productCategory
  };
}
async function syncEnquiryToZohoAndPersist(enquiry: AirtableRecord<EnquiryFields>) {
  const zohoContext = await resolveZohoBiginContext(enquiry);
  const syncPayload = {
    recordId: String(enquiry.fields["Zoho Bigin Record ID"] || "").trim(),
    dealRecordId: String(enquiry.fields["Zoho Bigin Deal ID"] || "").trim(),
    enquiryId: String(enquiry.fields["Enquiry ID"] || enquiry.id),
    loggedDateTime: String(enquiry.fields["Logged Date Time"] || ""),
    leadName: String(enquiry.fields["Lead Name"] || ""),
    company: String(enquiry.fields.Company || ""),
    phone: normalizePhone(enquiry.fields.Phone),
    email: normalizeEmail(enquiry.fields.Email),
    address: String(enquiry.fields.Address || ""),
    city: String(enquiry.fields.City || ""),
    state: String(enquiry.fields.State || ""),
    pincode: String(enquiry.fields.Pincode || ""),
    destinationAddress: String(enquiry.fields["Destination Address"] || ""),
    destinationCity: String(enquiry.fields["Destination City"] || ""),
    destinationState: String(enquiry.fields["Destination State"] || ""),
    destinationPincode: String(enquiry.fields["Destination Pincode"] || ""),
    parserStatus: String(enquiry.fields["Parser Status"] || ""),
    requirementSummary: String(enquiry.fields["Requirement Summary"] || ""),
    receiverWhatsappNumber: normalizePhone(enquiry.fields["Receiver WhatsApp Number"]),
    quotationRef: zohoContext.quotationRef,
    productName: zohoContext.productName,
    productKey: zohoContext.productKey,
    productDescription: zohoContext.productDescription,
    productUnitPrice: zohoContext.productUnitPrice,
    productFreight: zohoContext.productFreight,
    productGstRate: zohoContext.productGstRate,
    productCategory: zohoContext.productCategory
  };

  console.info("[zoho-bigin] enquiry sync requested from portal actions", {
    airtableRecordId: enquiry.id,
    payload: syncPayload
  });

  try {
    const result = await syncEnquiryToZohoBigin(syncPayload);

    if (!result) {
      return enquiry;
    }

    return updateRecordWithOptionalFieldFallback<EnquiryFields>(env.AIRTABLE_ENQUIRIES_TABLE, {
      id: enquiry.id,
      fields: {
        "Zoho Bigin Record ID": result.recordId,
        "Zoho Bigin Deal ID": result.dealRecordId,
        "Zoho Bigin Sync Status": `Synced (${result.action})`,
        "Zoho Bigin Synced At": new Date().toISOString(),
        "Zoho Bigin Sync Error": ""
      }
    }, optionalEnquiryFields);
  } catch (error) {
    const message = error instanceof Error ? error.message : "Unknown Zoho Bigin sync error";
    console.error("[zoho-bigin] enquiry sync failed in portal actions", {
      enquiryId: enquiry.id,
      message,
      payload: syncPayload,
      error: serializeError(error)
    });

    return updateRecordWithOptionalFieldFallback<EnquiryFields>(env.AIRTABLE_ENQUIRIES_TABLE, {
      id: enquiry.id,
      fields: {
        "Zoho Bigin Sync Status": "Sync Failed",
        "Zoho Bigin Synced At": new Date().toISOString(),
        "Zoho Bigin Sync Error": message
      }
    }, optionalEnquiryFields).catch(() => enquiry);
  }
}

function buildManualEnquiryDedupKey(input: z.infer<typeof enquiryPayloadSchema>) {
  return JSON.stringify({
    source: input.source,
    linkedCustomerId: String(input.linkedCustomerId || "").trim(),
    leadName: String(input.leadName || "").trim().toLowerCase(),
    company: String(input.company || "").trim().toLowerCase(),
    phone: normalizePhone(input.phone),
    email: normalizeEmail(input.email),
    state: String(input.state || "").trim().toLowerCase(),
    city: String(input.city || "").trim().toLowerCase(),
    pincode: String(input.pincode || "").replace(/\D/g, "").slice(0, 6),
    product: String(input.potentialProduct || "").trim()
  });
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
    ]
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

function mergeRequirementSummary(existingSummary: string, selectedProductSummary: string) {
  const existing = String(existingSummary || "").trim();
  const selected = String(selectedProductSummary || "").trim();

  if (!existing) {
    return selected;
  }

  if (!selected || existing.toLowerCase().includes(selected.toLowerCase())) {
    return existing;
  }

  return `${existing}\n${selected}`;
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

function airtablePaymentStatus(value: string) {
  if (value === "Pending") {
    return "Pending";
  }

  return value;
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
  const escapedReference = quotationReference.replace(/'/g, "\\'");
  const existingQuotations = await listRecords<QuotationFields>(env.AIRTABLE_QUOTATIONS_TABLE, {
    fields: [
      "Quotation Number",
      "Reference Number",
      "Linked Enquiry",
      "Linked Customer",
      "Status",
      "Draft File URL"
    ],
    filterByFormula: `OR({Reference Number}='${escapedReference}', ARRAYJOIN({Linked Enquiry})='${enquiry.id}')`,
    maxRecords: 10
  });
  const existingQuotation = existingQuotations[0];
  if (existingQuotation) {
    if (
      existingQuotation.fields["Linked Enquiry"]?.[0] !== enquiry.id ||
      existingQuotation.fields["Linked Customer"]?.[0] !== customerId
    ) {
      return updateRecord<QuotationFields>(env.AIRTABLE_QUOTATIONS_TABLE, {
        id: existingQuotation.id,
        fields: {
          "Linked Enquiry": safeLinkedValue(enquiry.id),
          "Linked Customer": safeLinkedValue(customerId)
        }
      });
    }

    return existingQuotation;
  }

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
        Status: enquiry.fields["Parser Status"] || ENQUIRY_STATUS_PARSED,
        "Draft Format": "XLSX",
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

function mapPortalQuotationItemsToOrderItems(
  items: PortalQuotationLineItem[],
  products: Map<string, AirtableRecord<ProductFields>>
) {
  return items.map((item, index) => {
    const product = products.get(item.productId);
    const qty = Number(item.qty || 0);
    const ratePerUnit = Number(item.rate || 0);
    const packingFreight = Number(item.transport || 0);
    const gst18 = Number(item.gstPercent || 0);
    const unitValue = Number((qty * ratePerUnit).toFixed(2));
    const totalAmount = Number((unitValue + packingFreight * qty + gst18 * qty).toFixed(2));

    return {
      id: "",
      productId: item.productId,
      description: product ? buildDescription(product) : "Order item",
      qty,
      ratePerUnit,
      packingFreight,
      unitValue,
      gst18,
      totalAmount,
      serialNo: index + 1
    };
  });
}

async function replaceOrderLineItemsForOrder(input: {
  orderId: string;
  quotationId: string;
  customerId: string;
  enquiryId: string;
  items: Array<z.infer<typeof orderPayloadSchema>["items"][number] & { serialNo?: number }>;
}) {
  const existing = await listRecords<OrderLineItemFields>(env.AIRTABLE_ORDER_LINE_ITEMS_TABLE, {
    fields: ["Order"],
    maxRecords: 500
  });
  const existingForOrder = existing.filter((item) => {
    return String(item.fields.Order || "") === input.orderId;
  });

  if (existingForOrder.length) {
    await deleteRecords(env.AIRTABLE_ORDER_LINE_ITEMS_TABLE, existingForOrder.map((item) => item.id));
  }

  if (!input.items.length) {
    return [];
  }

  return createRecords<OrderLineItemFields>(
    env.AIRTABLE_ORDER_LINE_ITEMS_TABLE,
    input.items.map((item, index) => ({
      Name: `${input.orderId}-${String(index + 1).padStart(2, "0")}`,
      Order: input.orderId,
      Quotation: input.quotationId,
      Enquiries: input.enquiryId,
      Customer: input.customerId,
      "Linked Product": item.productId,
      "S.No.": item.serialNo || index + 1,
      Description: item.description,
      Qty: item.qty,
      "Rate Per Unit": item.ratePerUnit,
      "Packing & Freight": item.packingFreight,
      "Unit Value": item.unitValue,
      "GST 18%": String(item.gst18 || 0),
      "Total Amount": String(item.totalAmount || 0)
    }))
  );
}

export async function getPortalOrderLineItems(orderId: string) {
  const items = await listRecords<OrderLineItemFields>(env.AIRTABLE_ORDER_LINE_ITEMS_TABLE, {
    fields: [
      "Order",
      "Linked Product",
      "S.No.",
      "Description",
      "Qty",
      "Rate Per Unit",
      "Packing & Freight",
      "Unit Value",
      "GST 18%",
      "Total Amount"
    ],
    maxRecords: 500
  });

  return items
    .filter((item) => {
      return String(item.fields.Order || "") === orderId;
    })
    .sort((left, right) => Number(left.fields["S.No."] || 0) - Number(right.fields["S.No."] || 0))
    .map((item) => ({
      id: item.id,
      productId: String(item.fields["Linked Product"] || ""),
      description: String(item.fields.Description || ""),
      qty: Number(item.fields.Qty || 0),
      ratePerUnit: Number(item.fields["Rate Per Unit"] || 0),
      packingFreight: Number(item.fields["Packing & Freight"] || 0),
      unitValue: Number(item.fields["Unit Value"] || 0),
      gst18: Number(item.fields["GST 18%"] || 0),
      totalAmount: Number(item.fields["Total Amount"] || 0)
    })) satisfies PortalOrderLineItem[];
}

function formatLineItemValue(value: number) {
  return Number.isFinite(value) ? value.toFixed(2) : "0.00";
}

async function nextCustomerIdentifier() {
  const customers = await listRecords<CustomerFields>(env.AIRTABLE_CUSTOMERS_TABLE, {
    fields: ["Client ID"]
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
    const match = String(order.fields["Order Number"] || "").match(/^DT-(\d+)$/i);
    return Math.max(currentMax, match ? Number(match[1]) : 0);
  }, 0);

  return `DT-${String(maxValue + 1).padStart(3, "0")}`;
}

function resolveOrderDestination(
  input: z.infer<typeof orderPayloadSchema>,
  customer: AirtableRecord<CustomerFields> | null,
  enquiry: AirtableRecord<EnquiryFields> | null,
  existing?: AirtableRecord<OrderFields> | null
) {
  const address =
    input.destinationAddress ||
    input.address ||
    customer?.fields["Destination Address"] ||
    customer?.fields.Address ||
    enquiry?.fields["Destination Address"] ||
    enquiry?.fields.Address ||
    existing?.fields["Destination Address"] ||
    existing?.fields.Address ||
    "";
  const state =
    input.destinationState ||
    input.state ||
    customer?.fields["Destination State"] ||
    customer?.fields.State ||
    enquiry?.fields["Destination State"] ||
    enquiry?.fields.State ||
    existing?.fields["Destination State"] ||
    existing?.fields.State ||
    "";
  const city =
    input.destinationCity ||
    input.city ||
    customer?.fields["Destination City"] ||
    customer?.fields.City ||
    enquiry?.fields["Destination City"] ||
    enquiry?.fields.City ||
    existing?.fields["Destination City"] ||
    existing?.fields.City ||
    "";
  const pincode =
    toPincodeString(input.destinationPincode) ||
    toPincodeString(input.pincode) ||
    toPincodeString(String(customer?.fields["Destination Pincode"] || "")) ||
    toPincodeString(String(customer?.fields.Pincode || "")) ||
    toPincodeString(String(enquiry?.fields["Destination Pincode"] || "")) ||
    toPincodeString(String(enquiry?.fields.Pincode || "")) ||
    toPincodeString(String(existing?.fields["Destination Pincode"] || "")) ||
    toPincodeString(String(existing?.fields.Pincode || ""));

  return {
    address,
    state,
    city,
    pincode
  };
}

async function updateCustomerDestinationFromOrder(
  customerId: string,
  destination: { address?: string; state?: string; city?: string; pincode?: string | number }
) {
  if (!customerId || (!destination.address && !destination.state && !destination.city && !destination.pincode)) {
    return;
  }

  await updateRecordWithOptionalFieldFallback<CustomerFields>(
    env.AIRTABLE_CUSTOMERS_TABLE,
    {
      id: customerId,
      fields: {
        "Destination Address": destination.address || "",
        "Destination State": destination.state || "",
        "Destination City": destination.city || "",
        "Destination Pincode": toPincodeString(String(destination.pincode || "")) || ""
      }
    },
    optionalCustomerDestinationFields
  ).catch((error) => {
    console.error("[order-destination] failed to update customer destination fields", {
      customerId,
      message: error instanceof Error ? error.message : "Unknown error",
      stack: error instanceof Error ? error.stack : undefined,
      destination
    });
  });
}

export async function createPortalEnquiry(payload: unknown) {
  const input = enquiryPayloadSchema.parse(payload);
  const dedupKey =
    input.source === "manual"
      ? buildManualEnquiryDedupKey(input)
      : "";

  if (dedupKey) {
    const existingInflight = manualEnquiryCreateInflight.get(dedupKey);
    if (existingInflight) {
      return existingInflight;
    }
  }

  const createPromise = (async () => {
  const productSummary = await resolvePotentialProductSummary(input.potentialProduct);
  const requirementSummary = mergeRequirementSummary(input.requirementSummary, productSummary.productName);
  const existingCustomer = input.linkedCustomerId
    ? await getRecord<CustomerFields>(env.AIRTABLE_CUSTOMERS_TABLE, input.linkedCustomerId)
    : await findExistingCustomerByContact(input);
  const linkedCustomerId = input.linkedCustomerId || existingCustomer?.id || "";
  const parserStatus = normalizeStatusForEnquiry(linkedCustomerId, input.source);
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
        "Requirement Summary": requirementSummary,
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
  const syncedEnquiry = await syncEnquiryToZohoAndPersist(provisioned.enquiry);

  return {
    enquiryRecordId: syncedEnquiry.id,
    enquiryId: syncedEnquiry.fields["Enquiry ID"] || created.fields["Enquiry ID"] || "",
    parserStatus: syncedEnquiry.fields["Parser Status"] || created.fields["Parser Status"] || parserStatus,
    linkedCustomerId: provisioned.customer.id,
    quotationRecordId: provisioned.quotation.id,
    quotationNumber: provisioned.quotation.fields["Quotation Number"] || "",
    driveFolderUrl: provisioned.folder?.folderUrl || provisioned.customer.fields["Drive Folder URL"] || ""
  };
  })();

  if (!dedupKey) {
    return createPromise;
  }

  manualEnquiryCreateInflight.set(dedupKey, createPromise);
  try {
    return await createPromise;
  } finally {
    manualEnquiryCreateInflight.delete(dedupKey);
  }
}

export async function updatePortalEnquiry(enquiryId: string, payload: unknown) {
  const input = enquiryPayloadSchema.parse(payload);
  const productSummary = await resolvePotentialProductSummary(input.potentialProduct);
  const existingEnquiry = await getRecord<EnquiryFields>(env.AIRTABLE_ENQUIRIES_TABLE, enquiryId);
  const requirementSummary = mergeRequirementSummary(
    input.requirementSummary || existingEnquiry.fields["Requirement Summary"] || "",
    productSummary.productName
  );
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
  const parserStatus =
    existingEnquiry.fields["Parser Status"] || normalizeStatusForEnquiry(linkedCustomerId, input.source);
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
        "Requirement Summary": requirementSummary,
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
  const syncedEnquiry = await syncEnquiryToZohoAndPersist(provisioned.enquiry);

  return {
    enquiryRecordId: syncedEnquiry.id,
    enquiryId: syncedEnquiry.fields["Enquiry ID"] || updatedEnquiry.fields["Enquiry ID"] || "",
    parserStatus: syncedEnquiry.fields["Parser Status"] || updatedEnquiry.fields["Parser Status"] || parserStatus,
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
    quotationStatus: draft.quotation.fields.Status || QUOTATION_STATUS_DRAFT,
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

  return {
    quotationId: quotation.id,
    quotationNumber: draft.quotation.fields["Quotation Number"] || quotation.fields["Quotation Number"] || quotation.id,
    createdCount: created.length,
    quotationStatus: draft.quotation.fields.Status || QUOTATION_STATUS_DRAFT,
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
        Status: QUOTATION_STATUS_SENT,
        "Sent Date": new Date().toISOString()
      }
    },
    optionalQuotationFields
  );

  const enquiryId = updated.fields["Linked Enquiry"]?.[0];
  if (enquiryId) {
    await updateRecord<EnquiryFields>(env.AIRTABLE_ENQUIRIES_TABLE, {
      id: enquiryId,
      fields: {
        "Parser Status": ENQUIRY_STATUS_SENT
      }
    });
  }

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

export async function createOrderFromQuotation(quotationId: string, payload?: unknown, actor?: AuthenticatedUser | null) {
  const quotation = await getRecord<QuotationFields>(env.AIRTABLE_QUOTATIONS_TABLE, quotationId);
  const input = orderPayloadSchema.parse({
    quotationId,
    ...(payload && typeof payload === "object" ? payload : {})
  });
  const existingOrders = await listRecords<OrderFields>(env.AIRTABLE_ORDERS_TABLE, {
    fields: ["Order Number", "Linked Quotation", "Linked Customer", "Line Items", "Order Status", "Order Date", "Total Amount"],
    maxRecords: 500
  });
  const existing = existingOrders.find((order) =>
    order.fields["Linked Quotation"]?.includes(quotationId)
  );
  if (existing) {
    await updateRecordWithOptionalFieldFallback<QuotationFields>(
      env.AIRTABLE_QUOTATIONS_TABLE,
      {
        id: quotationId,
        fields: {
          Status: QUOTATION_STATUS_ORDERED,
          Orders: existing.fields["Order Number"] || existing.id
        }
      },
      optionalQuotationFields
    );
    return existing;
  }

  const customerId = quotation.fields["Linked Customer"]?.[0] || "";
  const enquiryId = quotation.fields["Linked Enquiry"]?.[0] || "";
  const enquiry = enquiryId ? await getRecord<EnquiryFields>(env.AIRTABLE_ENQUIRIES_TABLE, enquiryId) : null;
  const customer = customerId ? await getRecord<CustomerFields>(env.AIRTABLE_CUSTOMERS_TABLE, customerId) : null;
  const metrics = await loadQuotationLineItemMetrics(quotationId);
  const products = await listRecords<ProductFields>(env.AIRTABLE_PRODUCTS_TABLE, { maxRecords: 500 });
  const productLookup = new Map(products.map((product) => [product.id, product]));
  const orderItems =
    input.items.length > 0
      ? input.items
      : mapPortalQuotationItemsToOrderItems(metrics.items, productLookup);
  const orderNumber = await nextOrderIdentifier();
  const destination = resolveOrderDestination(input, customer, enquiry);
  const orderFields: Record<string, unknown> = {
    "Order Number": orderNumber,
    "Order Date": input.orderDate || new Date().toISOString(),
    "Linked Quotation": safeLinkedValue(quotationId),
    "Linked Customer": safeLinkedValue(customerId),
    "Order Status": input.orderStatus,
    "Total Amount": String(input.totalAmount || metrics.quotationGrandTotal || 0),
    "Quotation Grand Total": String(metrics.quotationGrandTotal || 0),
    "Quotation Status": QUOTATION_STATUS_ORDERED,
    "Order Fulfillment Progress": ORDER_FULFILLMENT_PROGRESS_NOT_STARTED,
    "Order Notes": input.orderNotes,
    "Order Value": String(input.totalAmount || metrics.quotationGrandTotal || 0),
    "Payment Status": airtablePaymentStatus(input.paymentStatus),
    "Payment Terms": input.paymentTerms,
    "Order Ref Number Client": input.orderRefNumberClient,
    Address: input.address || destination.address,
    State: input.state || destination.state,
    City: input.city || destination.city,
    Pincode: toPincodeString(input.pincode) || destination.pincode,
    "Destination Address": destination.address,
    "Destination State": destination.state,
    "Destination City": destination.city,
    "Destination Pincode": destination.pincode
  };
  let created: AirtableRecord<OrderFields>;
  try {
    created = await createRecord<OrderFields>(env.AIRTABLE_ORDERS_TABLE, orderFields);
  } catch (error) {
    const message = error instanceof Error ? error.message : "";
    const retryFields = stripOptionalFields(orderFields, optionalOrderFields, message);
    if (!retryFields) {
      throw error;
    }

    created = await createRecord<OrderFields>(env.AIRTABLE_ORDERS_TABLE, retryFields);
  }

  await updateCustomerDestinationFromOrder(customerId, destination);

  const createdOrderLineItems = await replaceOrderLineItemsForOrder({
    orderId: created.id,
    quotationId,
    customerId,
    enquiryId,
    items: orderItems
  });
  await updateRecordWithOptionalFieldFallback<OrderFields>(
    env.AIRTABLE_ORDERS_TABLE,
    {
      id: created.id,
      fields: {
        "Line Items": linkedRecordIds(...createdOrderLineItems.map((item) => item.id))
      }
    },
    optionalOrderFields
  );

  await updateRecordWithOptionalFieldFallback<QuotationFields>(
    env.AIRTABLE_QUOTATIONS_TABLE,
    {
      id: quotationId,
      fields: {
        Status: QUOTATION_STATUS_ORDERED,
        Orders: created.fields["Order Number"] || orderNumber
      }
    },
    optionalQuotationFields
  );
  if (enquiryId) {
    await updateRecord<EnquiryFields>(env.AIRTABLE_ENQUIRIES_TABLE, {
      id: enquiryId,
      fields: {
        "Parser Status": ENQUIRY_STATUS_ORDERED
      }
    });
  }

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
  const quotationMetrics = await loadQuotationLineItemMetrics(input.quotationId);
  const customerId = input.customerId || quotation.fields["Linked Customer"]?.[0] || "";
  const enquiryId = input.enquiryId || quotation.fields["Linked Enquiry"]?.[0] || "";
  const customer = customerId ? await getRecord<CustomerFields>(env.AIRTABLE_CUSTOMERS_TABLE, customerId) : null;
  const enquiry = enquiryId ? await getRecord<EnquiryFields>(env.AIRTABLE_ENQUIRIES_TABLE, enquiryId) : null;
  const orderItems = input.items;
  const destination = resolveOrderDestination(input, customer, enquiry, existing);
  const fields: Record<string, unknown> = {
    "Order Date": input.orderDate || existing.fields["Order Date"] || new Date().toISOString(),
    "Linked Quotation": safeLinkedValue(input.quotationId),
    "Linked Customer": safeLinkedValue(customerId),
    "Order Status": input.orderStatus,
    "Total Amount": String(input.totalAmount || 0),
    "Order Notes": input.orderNotes,
    "Quotation Grand Total": String(quotationMetrics.quotationGrandTotal || 0),
    "Quotation Status": quotation.fields.Status || "",
    "Order Value": String(input.totalAmount || 0),
    "Payment Status": airtablePaymentStatus(input.paymentStatus),
    "Payment Terms": input.paymentTerms,
    "Order Ref Number Client": input.orderRefNumberClient,
    Address: input.address || destination.address,
    State: input.state || destination.state,
    City: input.city || destination.city,
    Pincode: toPincodeString(input.pincode) || destination.pincode,
    "Destination Address": destination.address,
    "Destination State": destination.state,
    "Destination City": destination.city,
    "Destination Pincode": destination.pincode
  };

  const updated = await updateRecordWithOptionalFieldFallback<OrderFields>(
    env.AIRTABLE_ORDERS_TABLE,
    {
      id: orderId,
      fields
    },
    optionalOrderFields
  );

  await updateCustomerDestinationFromOrder(customerId, destination);

  await updateRecordWithOptionalFieldFallback<QuotationFields>(
    env.AIRTABLE_QUOTATIONS_TABLE,
    {
      id: input.quotationId,
      fields: {
        Orders: updated.fields["Order Number"] || existing.fields["Order Number"] || orderId
      }
    },
    optionalQuotationFields
  );

  const updatedOrderLineItems = await replaceOrderLineItemsForOrder({
    orderId,
    quotationId: input.quotationId,
    customerId,
    enquiryId,
    items: orderItems
  });
  await updateRecordWithOptionalFieldFallback<OrderFields>(
    env.AIRTABLE_ORDERS_TABLE,
    {
      id: orderId,
      fields: {
        "Line Items": linkedRecordIds(...updatedOrderLineItems.map((item) => item.id))
      }
    },
    optionalOrderFields
  );

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
