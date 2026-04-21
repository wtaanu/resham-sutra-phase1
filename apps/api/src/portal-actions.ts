import { z } from "zod";
import {
  createRecord,
  createRecords,
  getRecord,
  listRecords,
  updateRecord,
  type AirtableRecord
} from "./airtable.js";
import { env } from "./config.js";
import { createRecordWithUniqueNumber } from "./numbering.js";
import { createCustomerForEnquiry, refreshDraftForQuotation } from "./intake-processor.js";

type EnquiryFields = {
  "Enquiry ID"?: string;
  "Lead Name"?: string;
  Company?: string;
  Phone?: string;
  Email?: string;
  Address?: string;
  State?: string;
  City?: string;
  Pincode?: number;
  "Destination Address"?: string;
  "Destination State"?: string;
  "Destination City"?: string;
  "Destination Pincode"?: number;
  "Parser Status"?: string;
  "Linked Customer"?: string[];
  Quotations?: string[];
  "Requirement Summary"?: string;
  "Requested Asset"?: string;
  "Potential Product"?: string;
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
  Pincode?: number;
};

type QuotationFields = {
  "Quotation Number"?: string;
  Status?: string;
};

type ProductFields = {
  "Product Key"?: string;
  Model?: string;
  "Product Name"?: string;
  Narration?: string;
  "Bulk Sale Price"?: number;
  MRP?: number;
  "GST %"?: number;
  "Pkg & Transport"?: number;
};

type QuotationLineItemFields = {
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

const airtableRecordIdSchema = z.string().regex(/^rec[a-zA-Z0-9]+$/, "Expected an Airtable record ID");
const optionalAirtableRecordIdSchema = z.union([airtableRecordIdSchema, z.literal("")]).default("");
const pincodeSchema = z
  .string()
  .trim()
  .regex(/^\d{6}$/, "Enter a valid 6-digit pincode");
const optionalPincodeSchema = z.union([pincodeSchema, z.literal("")]).default("");

function linkedRecordIds(...recordIds: Array<string | undefined>) {
  return recordIds.filter((value): value is string => Boolean(value));
}

const enquiryPayloadSchema = z.object({
  linkedCustomerId: optionalAirtableRecordIdSchema,
  leadName: z.string().trim().min(1, "Lead name is required"),
  company: z.string().trim().optional().default(""),
  phone: z.string().trim().optional().default(""),
  email: z.string().trim().optional().default(""),
  address: z.string().trim().optional().default(""),
  state: z.string().trim().optional().default(""),
  city: z.string().trim().optional().default(""),
  pincode: optionalPincodeSchema,
  destinationAddress: z.string().trim().optional().default(""),
  destinationState: z.string().trim().optional().default(""),
  destinationCity: z.string().trim().optional().default(""),
  destinationPincode: optionalPincodeSchema,
  requirementSummary: z.string().trim().optional().default(""),
  requestedAsset: z.string().trim().optional().default(""),
  potentialProduct: z.string().trim().optional().default("")
});

const lineItemPayloadSchema = z.object({
  quotationId: airtableRecordIdSchema,
  items: z
    .array(
      z.object({
        productId: airtableRecordIdSchema,
        qty: z.coerce.number().positive("Quantity should be greater than zero"),
        totalAmount: z.coerce.number().nonnegative("Total amount should be zero or more").optional()
      })
    )
    .min(1, "Add at least one line item")
});

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

export async function createPortalEnquiry(payload: unknown) {
  const input = enquiryPayloadSchema.parse(payload);
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
  const enquiryFields = withOptionalNumberField(
    withOptionalNumberField(
      {
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
        "Requirement Summary": input.requirementSummary,
        "Requested Asset": input.requestedAsset,
        "Potential Product": input.potentialProduct
      },
      "Pincode",
      mainPincode
    ),
    "Destination Pincode",
    destinationPincodeNumber
  );

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
        if (!mentionsField(message, "Pincode")) {
          throw error;
        }

        const retryFields = { ...fieldsWithId };
        delete retryFields.Pincode;
        delete retryFields["Destination Pincode"];
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

  let nextLineNo =
    existingItems.filter((item) => item.fields.Quotation?.includes(input.quotationId)).length + 1;

  const lineItemFields = input.items.map((item) => {
    const product = productLookup.get(item.productId);
    if (!product) {
      throw new Error("One or more selected products could not be found.");
    }

    const rate = toNumericValue(product.fields["Bulk Sale Price"], toNumericValue(product.fields.MRP, 0));
    const transport = toNumericValue(product.fields["Pkg & Transport"], 0);
    const gstPercent = toNumericValue(product.fields["GST %"], 0);
    const qty = Number(item.qty || 0);
    const unitValue = rate * qty;
    const defaultGstAmount = Number((((unitValue + transport) * gstPercent) / 100).toFixed(2));
    const defaultTotalAmount = Number((unitValue + transport + defaultGstAmount).toFixed(2));
    const totalAmount = Number((item.totalAmount ?? defaultTotalAmount).toFixed(2));
    const gstAmount = Number((totalAmount - unitValue - transport).toFixed(2));

    return {
      Quotation: linkedRecordIds(input.quotationId),
      "Linked Product": linkedRecordIds(item.productId),
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
    quotationStatus: draft.quotation.fields.Status || "Ready for Review",
    draftFileUrl: draft.quotation.fields["Draft File URL"] || "",
    driveFolderUrl: draft.folder.folderUrl || draft.quotation.fields["Drive Folder URL"] || ""
  };
}
