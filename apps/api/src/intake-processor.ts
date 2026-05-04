import { access, mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import {
  createRecord,
  getRecord,
  listRecords,
  type AirtableRecord,
  updateRecord
} from "./airtable.js";
import {
  createDraftDocument,
  createPdfDocument,
  getStoredDocumentArtifact
} from "./documents.js";
import {
  exportDriveFile,
  extractDriveFileId,
  extractDriveFolderId,
  findOrCreateClientFolder,
  isDriveConfigured,
  uploadFileToFolder
} from "./drive.js";
import { env } from "./config.js";
import { getSmtpConfigState, isSmtpConfigured, resolveDocumentAttachment, sendMail } from "./mailer.js";
import {
  createRecordWithUniqueNumber,
  nextQuotationNumber as generateNextQuotationNumber
} from "./numbering.js";
import {
  isWhatsAppOutboundConfigured,
  sendQuotationDocumentOnWhatsApp
} from "./whatsapp-outbound.js";
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
  "Destination Pincode"?: string;
  "Parser Status"?: string;
  "Linked Customer"?: string[];
  Quotations?: string[];
  "Drive Folder URL"?: string;
  "Requirement Summary"?: string;
  "Requested Asset"?: string;
  "Receiver WhatsApp Number"?: string;
  "Potential Product"?: string;
  "Zoho Bigin Deal ID"?: string;
  Quantity?: string | number;
  Qty?: string | number;
  Notes?: string;
  "Zoho Bigin Record ID"?: string;
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
  "Customer Type"?: string;
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
  "Drive Folder URL"?: string;
  "Reference Number"?: string;
  "Buyer Block"?: string;
  "Sent Date"?: string;
  "WhatsApp Sent Date Time"?: string;
  "Email Sent Date Time"?: string;
  "WhatsApp Outbound Message ID"?: string;
  "WhatsApp Delivery Status"?: string;
  "WhatsApp Delivery Updated At"?: string;
  "WhatsApp Delivery Error"?: string;
  "Send Quotation"?: boolean;
  "Send Reminder"?: boolean;
  "Mark Accepted"?: boolean;
  "Mark Rejected"?: boolean;
  "Reminder Count"?: number;
  "Last Reminder Date"?: string;
  "Next Reminder Date"?: string;
  "Final PDF Generated At"?: string;
};

type ProductFields = {
  "Product Key"?: string;
  "Product Name"?: string;
  Model?: string;
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
  Quotation?: string[];
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

type ProcessingResult = {
  enquiryRecordId: string;
  enquiryId: string;
  customerRecordId?: string;
  quotationRecordId?: string;
  driveFolderUrl?: string;
  status: "processed" | "error" | "skipped";
  message?: string;
};

const customerLocks = new Map<string, Promise<void>>();
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
  "WhatsApp Outbound Message ID",
  "WhatsApp Delivery Status",
  "WhatsApp Delivery Updated At",
  "WhatsApp Delivery Error"
] as const;

function linkedRecordIds(...recordIds: Array<string | undefined>) {
  return recordIds.filter((value): value is string => Boolean(value));
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

  console.info("[zoho-bigin] enquiry sync requested from intake processor", {
    airtableRecordId: enquiry.id,
    payload: syncPayload
  });

  try {
    const result = await syncEnquiryToZohoBigin(syncPayload);

    if (!result) {
      return enquiry;
    }

    return updateRecordWithOptionalFieldFallback<EnquiryFields>(
      env.AIRTABLE_ENQUIRIES_TABLE,
      {
        id: enquiry.id,
        fields: {
          "Zoho Bigin Record ID": result.recordId,
          "Zoho Bigin Deal ID": result.dealRecordId,
          "Zoho Bigin Sync Status": `Synced (${result.action})`,
          "Zoho Bigin Synced At": new Date().toISOString(),
          "Zoho Bigin Sync Error": ""
        }
      },
      optionalEnquiryFields
    );
  } catch (error) {
    const message = error instanceof Error ? error.message : "Unknown Zoho Bigin sync error";
    console.error("[zoho-bigin] enquiry sync failed in intake processor", {
      enquiryId: enquiry.id,
      message,
      payload: syncPayload,
      error: serializeError(error)
    });

    return updateRecordWithOptionalFieldFallback<EnquiryFields>(
      env.AIRTABLE_ENQUIRIES_TABLE,
      {
        id: enquiry.id,
        fields: {
          "Zoho Bigin Sync Status": "Sync Failed",
          "Zoho Bigin Synced At": new Date().toISOString(),
          "Zoho Bigin Sync Error": message
        }
      },
      optionalEnquiryFields
    ).catch(() => enquiry);
  }
}

function toPincodeNumber(value: unknown) {
  const normalized = String(value ?? "")
    .replace(/\D/g, "")
    .slice(0, 6);

  return normalized ? Number(normalized) : undefined;
}

function toPincodeString(value: unknown) {
  const normalized = String(value ?? "")
    .replace(/\D/g, "")
    .slice(0, 6);

  return normalized || undefined;
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

async function updateRecordWithOptionalFieldFallback<TFields extends Record<string, unknown>>(
  tableName: string,
  payload: { id: string; fields: Record<string, unknown> },
  optionalFieldNames: readonly string[]
) {
  try {
    return await updateRecord<TFields>(tableName, payload);
  } catch (error) {
    const message = error instanceof Error ? error.message : "";
    const retryFields = stripOptionalFields(payload.fields, optionalFieldNames, message);

    if (!retryFields) {
      throw error;
    }

    return updateRecord<TFields>(tableName, {
      ...payload,
      fields: retryFields
    });
  }
}

function customerLockKeys(enquiry: AirtableRecord<EnquiryFields>) {
  return [
    enquiry.id,
    normalizePhone(enquiry.fields.Phone),
    normalizeEmail(enquiry.fields.Email)
  ].filter(Boolean);
}

async function withCustomerLocks<T>(keys: string[], task: () => Promise<T>) {
  if (!keys.length) {
    return task();
  }

  const uniqueKeys = [...new Set(keys)].sort();
  const previousLocks = uniqueKeys.map((key) => customerLocks.get(key) ?? Promise.resolve());
  let release!: () => void;
  const current = new Promise<void>((resolve) => {
    release = resolve;
  });

  uniqueKeys.forEach((key) => {
    const previous = customerLocks.get(key) ?? Promise.resolve();
    customerLocks.set(key, previous.then(() => current));
  });

  await Promise.all(previousLocks);

  try {
    return await task();
  } finally {
    release();
    uniqueKeys.forEach((key) => {
      if (customerLocks.get(key) === current) {
        customerLocks.delete(key);
      }
    });
  }
}

function safeFolderPart(value: string) {
  return value
    .replace(/[\\/:*?"<>|]/g, "-")
    .replace(/\s+/g, " ")
    .trim();
}

function buildCustomerFolderName(customer: AirtableRecord<CustomerFields>) {
  return safeFolderPart(
    `${customer.fields["Client ID"] || customer.id}-${customer.fields["Customer Name"] || "Unknown Client"}`
  );
}

function enquiryStatusFields(status: string) {
  return {
    "Parser Status": status
  };
}

function isLockedQuotationStatus(status: unknown) {
  const normalized = String(status || "").trim().toLowerCase();
  return normalized === "approved quote" || normalized === "sent quote" || normalized === "ordered";
}

function nextClientId(customers: AirtableRecord<CustomerFields>[]) {
  const max = customers.reduce((currentMax, customer) => {
    const clientId = customer.fields["Client ID"] ?? "";
    const match = clientId.match(/CUST-(\d+)/);
    const value = match ? Number(match[1]) : 0;
    return Math.max(currentMax, value);
  }, 0);

  return `CUST-${String(max + 1).padStart(3, "0")}`;
}

function findMatchingCustomer(
  enquiry: AirtableRecord<EnquiryFields>,
  customers: AirtableRecord<CustomerFields>[]
) {
  const enquiryPhone = normalizePhone(enquiry.fields.Phone);
  const enquiryEmail = normalizeEmail(enquiry.fields.Email);

  if (enquiryPhone) {
    const byPhone = customers.find((customer) => {
      const customerPhone = normalizePhone(customer.fields.Phone || customer.fields.WhatsApp);
      return customerPhone && customerPhone === enquiryPhone;
    });

    if (byPhone) {
      return byPhone;
    }
  }

  if (enquiryEmail) {
    const byEmail = customers.find((customer) => {
      const customerEmail = normalizeEmail(customer.fields.Email);
      return customerEmail && customerEmail === enquiryEmail;
    });

    if (byEmail) {
      return byEmail;
    }
  }

  return null;
}

async function listCustomerMatchRecords() {
  return listRecords<CustomerFields>(env.AIRTABLE_CUSTOMERS_TABLE, {
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
}

async function ensureCustomer(enquiry: AirtableRecord<EnquiryFields>) {
  return withCustomerLocks(customerLockKeys(enquiry), async () => {
    const linkedCustomerId = enquiry.fields["Linked Customer"]?.[0] || "";
    if (linkedCustomerId) {
      const linkedCustomer = await getCustomerById(linkedCustomerId);
      const enquiryPhone = normalizePhone(enquiry.fields.Phone);
      const enquiryEmail = normalizeEmail(enquiry.fields.Email);
      const linkedPhone = normalizePhone(linkedCustomer.fields.Phone || linkedCustomer.fields.WhatsApp);
      const linkedEmail = normalizeEmail(linkedCustomer.fields.Email);
      const contactMatches =
        Boolean(enquiryPhone && linkedPhone && enquiryPhone === linkedPhone) ||
        Boolean(enquiryEmail && linkedEmail && enquiryEmail === linkedEmail);
      const canApplyEnquiryIdentity = contactMatches || (!linkedPhone && !linkedEmail);
      const enquiryOrExisting = (enquiryValue: unknown, existingValue: unknown, fallback = "") =>
        canApplyEnquiryIdentity && enquiryValue ? enquiryValue : existingValue || fallback;
      const updatedCustomerFields = withOptionalNumberField(
        {
          "Customer Name": enquiryOrExisting(
            enquiry.fields["Lead Name"],
            linkedCustomer.fields["Customer Name"],
            "Unknown Client"
          ),
          Company: enquiryOrExisting(enquiry.fields.Company, linkedCustomer.fields.Company),
          Phone: enquiryOrExisting(enquiry.fields.Phone, linkedCustomer.fields.Phone),
          WhatsApp: enquiryOrExisting(enquiry.fields.Phone, linkedCustomer.fields.WhatsApp),
          Email: enquiryOrExisting(enquiry.fields.Email, linkedCustomer.fields.Email),
          Address: enquiryOrExisting(enquiry.fields.Address, linkedCustomer.fields.Address),
          State: enquiryOrExisting(enquiry.fields.State, linkedCustomer.fields.State),
          City: enquiryOrExisting(enquiry.fields.City, linkedCustomer.fields.City),
          "Customer Type": linkedCustomer.fields["Customer Type"] || "Domestic"
        },
        "Pincode",
        (canApplyEnquiryIdentity ? toPincodeNumber(enquiry.fields.Pincode) : undefined) ??
          toPincodeNumber(linkedCustomer.fields.Pincode)
      );

      console.log("[ensure-customer:update] preparing Airtable write", {
        customerId: linkedCustomer.id,
        enquiryId: enquiry.id,
        contactMatches,
        canApplyEnquiryIdentity,
        pincodeNumber:
          (canApplyEnquiryIdentity ? toPincodeNumber(enquiry.fields.Pincode) : undefined) ??
          toPincodeNumber(linkedCustomer.fields.Pincode),
        pincodeText:
          (canApplyEnquiryIdentity ? toPincodeString(enquiry.fields.Pincode) : undefined) ??
          toPincodeString(linkedCustomer.fields.Pincode)
      });

      try {
        return await updateRecord<CustomerFields>(env.AIRTABLE_CUSTOMERS_TABLE, {
          id: linkedCustomer.id,
          fields: updatedCustomerFields
        });
      } catch (error) {
        const message = error instanceof Error ? error.message : "";
        console.log("[ensure-customer:update] Airtable update failed", {
          customerId: linkedCustomer.id,
          enquiryId: enquiry.id,
          message,
          pincode: updatedCustomerFields.Pincode
        });
        if (!message.toLowerCase().includes("pincode")) {
          throw error;
        }

        const retryFields = withOptionalStringField(
          { ...updatedCustomerFields },
          "Pincode",
          (canApplyEnquiryIdentity ? toPincodeString(enquiry.fields.Pincode) : undefined) ??
            toPincodeString(linkedCustomer.fields.Pincode)
        );
        console.log("[ensure-customer:update] retrying with string pincode", {
          customerId: linkedCustomer.id,
          enquiryId: enquiry.id,
          pincode: retryFields.Pincode
        });
        return updateRecord<CustomerFields>(env.AIRTABLE_CUSTOMERS_TABLE, {
          id: linkedCustomer.id,
          fields: retryFields
        });
      }
    }

    const customers = await listCustomerMatchRecords();

    const matchedCustomer = findMatchingCustomer(enquiry, customers);
    if (matchedCustomer) {
      return matchedCustomer;
    }

    const clientId = nextClientId(customers);
    const customerFields = withOptionalNumberField(
      {
        "Client ID": clientId,
        "Customer Name": enquiry.fields["Lead Name"] || "Unknown Client",
        Company: enquiry.fields.Company || "",
        Phone: enquiry.fields.Phone || "",
        WhatsApp: enquiry.fields.Phone || "",
        Email: enquiry.fields.Email || "",
        Address: enquiry.fields.Address || "",
        State: enquiry.fields.State || "",
        City: enquiry.fields.City || "",
        "Customer Type": "Domestic"
      },
      "Pincode",
      toPincodeNumber(enquiry.fields.Pincode)
    );

    console.log("[ensure-customer:create] preparing Airtable write", {
      enquiryId: enquiry.id,
      pincodeNumber: toPincodeNumber(enquiry.fields.Pincode),
      pincodeText: toPincodeString(enquiry.fields.Pincode)
    });

    try {
      return await createRecord<CustomerFields>(env.AIRTABLE_CUSTOMERS_TABLE, customerFields);
    } catch (error) {
      const message = error instanceof Error ? error.message : "";
      console.log("[ensure-customer:create] Airtable create failed", {
        enquiryId: enquiry.id,
        message,
        pincode: customerFields.Pincode
      });
      if (!message.toLowerCase().includes("pincode")) {
        throw error;
      }

      const retryFields = withOptionalStringField(
        { ...customerFields },
        "Pincode",
        toPincodeString(enquiry.fields.Pincode)
      );
      console.log("[ensure-customer:create] retrying with string pincode", {
        enquiryId: enquiry.id,
        pincode: retryFields.Pincode
      });
      return createRecord<CustomerFields>(env.AIRTABLE_CUSTOMERS_TABLE, retryFields);
    }
  });
}

async function ensureFolder(customer: AirtableRecord<CustomerFields>) {
  const existingUrl = customer.fields["Drive Folder URL"] || "";
  if (existingUrl) {
    return {
      folderId: extractDriveFolderId(existingUrl),
      folderUrl: existingUrl
    };
  }

  if (!isDriveConfigured()) {
    return {
      folderId: "",
      folderUrl: ""
    };
  }

  const clientId = customer.fields["Client ID"] || customer.id;
  const customerName = safeFolderPart(customer.fields["Customer Name"] || "Unknown Client");
  const folderName = `${clientId}-${customerName}`;

  const folder = await findOrCreateClientFolder(folderName);

  await updateRecord<CustomerFields>(env.AIRTABLE_CUSTOMERS_TABLE, {
    id: customer.id,
    fields: {
      "Drive Folder URL": folder.folderUrl
    }
  });

  return {
    folderId: folder.folderId,
    folderUrl: folder.folderUrl
  };
}

function buildBuyerBlock(enquiry: AirtableRecord<EnquiryFields>) {
  return [
    enquiry.fields["Lead Name"] || "",
    enquiry.fields.Company || "",
    enquiry.fields.Address || "",
    [enquiry.fields.City || "", enquiry.fields.State || "", enquiry.fields.Pincode || ""]
      .filter(Boolean)
      .join(", "),
    enquiry.fields.Phone || "",
    enquiry.fields.Email || ""
  ]
    .map((value) => String(value).trim())
    .filter(Boolean)
    .join("\n");
}

function buildConsigneeBlock(enquiry: AirtableRecord<EnquiryFields>) {
  const hasDestinationAddress =
    Boolean(String(enquiry.fields["Destination Address"] || "").trim()) ||
    Boolean(String(enquiry.fields["Destination City"] || "").trim()) ||
    Boolean(String(enquiry.fields["Destination State"] || "").trim()) ||
    Boolean(String(enquiry.fields["Destination Pincode"] || "").trim());

  return [
    enquiry.fields["Lead Name"] || "",
    enquiry.fields.Company || "",
    hasDestinationAddress ? enquiry.fields["Destination Address"] || "" : "",
    [
      hasDestinationAddress ? enquiry.fields["Destination City"] || "" : "",
      hasDestinationAddress ? enquiry.fields["Destination State"] || "" : "",
      hasDestinationAddress ? enquiry.fields["Destination Pincode"] || "" : ""
    ]
      .filter(Boolean)
      .join(", "),
    enquiry.fields.Phone || "",
    enquiry.fields.Email || ""
  ]
    .map((value) => String(value).trim())
    .filter(Boolean)
    .join("\n");
}

function buildQuotationTerms(templateCode: "domestic-standard" | "myanmar-proforma") {
  if (templateCode === "myanmar-proforma") {
    return [
      "PRICE TERM - Price includes delivery up to destination as agreed.",
      "Dispatch - 8 - 10 weeks of order confirmation.",
      "Payment Term - 50% advance with PO and 50% before dispatch.",
      "Warranty against manufacturing defect for 1 year from date of invoice.",
      "Quotation valid for 30 days."
    ];
  }

  return [
    "PRICE TERM - Price includes delivery up to destination as agreed.",
    "Dispatch - 8 - 10 weeks of order confirmation.",
    "Payment Term - 100% advance before dispatch.",
    "Warranty against manufacturing defect for 1 year from date of invoice.",
    "Quotation valid for 30 days."
  ];
}

async function syncEnquiryAddressFromCustomer(
  enquiry: AirtableRecord<EnquiryFields>,
  customer: AirtableRecord<CustomerFields>
) {
  const needsAddressUpdate =
    (customer.fields.Address && enquiry.fields.Address !== customer.fields.Address) ||
    (customer.fields.State && enquiry.fields.State !== customer.fields.State) ||
    (customer.fields.City && enquiry.fields.City !== customer.fields.City) ||
    (customer.fields.Pincode && enquiry.fields.Pincode !== customer.fields.Pincode);

  if (!needsAddressUpdate) {
    return enquiry;
  }

  const fields = withOptionalNumberField(
    {
      Address: customer.fields.Address || enquiry.fields.Address || "",
      State: customer.fields.State || enquiry.fields.State || "",
      City: customer.fields.City || enquiry.fields.City || ""
    },
    "Pincode",
    toPincodeNumber(customer.fields.Pincode) ?? toPincodeNumber(enquiry.fields.Pincode)
  );

  console.log("[sync-enquiry-address] preparing Airtable write", {
    enquiryId: enquiry.id,
    customerId: customer.id,
    pincodeNumber: toPincodeNumber(customer.fields.Pincode) ?? toPincodeNumber(enquiry.fields.Pincode),
    pincodeText: toPincodeString(customer.fields.Pincode) ?? toPincodeString(enquiry.fields.Pincode)
  });

  try {
    return await updateRecord<EnquiryFields>(env.AIRTABLE_ENQUIRIES_TABLE, {
      id: enquiry.id,
      fields
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : "";
    console.log("[sync-enquiry-address] Airtable update failed", {
      enquiryId: enquiry.id,
      customerId: customer.id,
      message,
      pincode: fields.Pincode
    });
    if (!message.toLowerCase().includes("pincode")) {
      throw error;
    }

    const retryFields = withOptionalStringField(
      { ...fields },
      "Pincode",
      toPincodeString(customer.fields.Pincode) ?? toPincodeString(enquiry.fields.Pincode)
    );
    console.log("[sync-enquiry-address] retrying with string pincode", {
      enquiryId: enquiry.id,
      customerId: customer.id,
      pincode: retryFields.Pincode
    });
    return updateRecord<EnquiryFields>(env.AIRTABLE_ENQUIRIES_TABLE, {
      id: enquiry.id,
      fields: retryFields
    });
  }
}

async function nextQuotationNumber() {
  return generateNextQuotationNumber(env.AIRTABLE_QUOTATIONS_TABLE, "Quotation Number");
}

async function getQuotationById(quotationId: string) {
  return getRecord<QuotationFields>(env.AIRTABLE_QUOTATIONS_TABLE, quotationId);
}

async function getCustomerById(customerId: string) {
  return getRecord<CustomerFields>(env.AIRTABLE_CUSTOMERS_TABLE, customerId);
}

async function findExistingQuotation(enquiryId: string) {
  const escaped = enquiryId.replace(/'/g, "\\'");
  const quotations = await listRecords<QuotationFields>(env.AIRTABLE_QUOTATIONS_TABLE, {
    fields: [
      "Quotation Number",
      "Reference Number",
      "Draft File URL",
      "Status",
      "Linked Customer",
      "Linked Enquiry"
    ],
    filterByFormula: `OR({Reference Number}='${escaped}', ARRAYJOIN({Linked Enquiry})='${escaped}')`,
    maxRecords: 10
  });

  return quotations[0] ?? null;
}

async function syncEnquiryQuotationLinks(
  enquiry: AirtableRecord<EnquiryFields>,
  customer: AirtableRecord<CustomerFields>,
  quotation: AirtableRecord<QuotationFields>
) {
  const enquiryNeedsUpdate =
    enquiry.fields["Linked Customer"]?.[0] !== customer.id ||
    enquiry.fields.Quotations?.[0] !== quotation.id;

  const quotationNeedsUpdate =
    quotation.fields["Linked Enquiry"]?.[0] !== enquiry.id ||
    quotation.fields["Linked Customer"]?.[0] !== customer.id;

  const [updatedEnquiry, updatedQuotation] = await Promise.all([
    enquiryNeedsUpdate
      ? updateRecord<EnquiryFields>(env.AIRTABLE_ENQUIRIES_TABLE, {
          id: enquiry.id,
          fields: {
            "Linked Customer": linkedRecordIds(customer.id),
            Quotations: linkedRecordIds(quotation.id)
          }
        })
      : Promise.resolve(enquiry),
    quotationNeedsUpdate
      ? updateRecord<QuotationFields>(env.AIRTABLE_QUOTATIONS_TABLE, {
          id: quotation.id,
          fields: {
            "Linked Enquiry": linkedRecordIds(enquiry.id),
            "Linked Customer": linkedRecordIds(customer.id)
          }
        })
      : Promise.resolve(quotation)
  ]);

  return {
    enquiry: updatedEnquiry,
    quotation: updatedQuotation
  };
}

async function ensureQuotationShell(
  enquiry: AirtableRecord<EnquiryFields>,
  customer: AirtableRecord<CustomerFields>
) {
  const linkedQuotationId = enquiry.fields.Quotations?.[0];
  if (linkedQuotationId) {
    const linkedQuotation = await getQuotationById(linkedQuotationId);
    const synced = await syncEnquiryQuotationLinks(enquiry, customer, linkedQuotation);
    return synced.quotation;
  }

  const enquiryReference = enquiry.fields["Enquiry ID"] || enquiry.id;
  const existing = await findExistingQuotation(enquiryReference);
  if (existing) {
    const synced = await syncEnquiryQuotationLinks(enquiry, customer, existing);
    return synced.quotation;
  }

  const created = await createRecordWithUniqueNumber<QuotationFields, AirtableRecord<QuotationFields>>({
    tableName: env.AIRTABLE_QUOTATIONS_TABLE,
    fieldName: "Quotation Number",
    prefix: "QTN",
    create: (quotationNumber) =>
      createRecord<QuotationFields>(env.AIRTABLE_QUOTATIONS_TABLE, {
        "Quotation Number": quotationNumber,
        "Logged Date Time": new Date().toISOString(),
        "Linked Enquiry": linkedRecordIds(enquiry.id),
        "Linked Customer": linkedRecordIds(customer.id),
        Status: enquiry.fields["Parser Status"] || "Parsed",
        "Draft Format": "XLSX",
        "Reference Number": enquiryReference,
        "Buyer Block": buildBuyerBlock(enquiry),
        "Send Quotation": false,
        "Send Reminder": false,
        "Mark Accepted": false,
        "Mark Rejected": false,
        "Reminder Count": 0
      }).catch((error) => {
        const message = error instanceof Error ? error.message : "";
        const retryFields = stripOptionalFields(
          {
            "Quotation Number": quotationNumber,
            "Logged Date Time": new Date().toISOString(),
            "Linked Enquiry": linkedRecordIds(enquiry.id),
            "Linked Customer": linkedRecordIds(customer.id),
            Status: enquiry.fields["Parser Status"] || "Parsed",
            "Draft Format": "XLSX",
            "Reference Number": enquiryReference,
            "Buyer Block": buildBuyerBlock(enquiry),
            "Send Quotation": false,
            "Send Reminder": false,
            "Mark Accepted": false,
            "Mark Rejected": false,
            "Reminder Count": 0
          },
          optionalQuotationFields,
          message
        );

        if (!retryFields) {
          throw error;
        }

        return createRecord<QuotationFields>(env.AIRTABLE_QUOTATIONS_TABLE, retryFields);
      })
  });

  const synced = await syncEnquiryQuotationLinks(enquiry, customer, created);
  return synced.quotation;
}

async function listAllLineItems() {
  return listRecords<QuotationLineItemFields>(env.AIRTABLE_QUOTATION_LINE_ITEMS_TABLE, {
    fields: [
      "Quotation",
      "Description Override",
      "Qty",
      "Rate Per Unit",
      "Pkg & Transport",
      "GST %",
      "GST Amount",
      "Total Amount",
      "Unit Value"
    ],
    maxRecords: 500
  });
}

function getLineItemsForQuotation(
  quotationId: string,
  lineItems: AirtableRecord<QuotationLineItemFields>[]
) {
  return lineItems.filter((item) => item.fields.Quotation?.includes(quotationId));
}

function mapDraftLineItems(items: AirtableRecord<QuotationLineItemFields>[]) {
  return items.map((item, index) => ({
    lineNo: index + 1,
    description: String(item.fields["Description Override"] || "Quotation item").trim(),
    qty: Number(item.fields.Qty || 0),
    rate: Number(item.fields["Rate Per Unit"] || 0),
    transport: Number(item.fields["Pkg & Transport"] || 0),
    gstPercent: Number(item.fields["GST %"] || 0),
    gstAmount: Number(item.fields["GST Amount"] || 0),
    totalAmount: Number(item.fields["Total Amount"] || 0),
    unitValue: Number(item.fields["Unit Value"] || 0)
  }));
}

async function createFallbackPdfFromQuotationData(input: {
  quotation: AirtableRecord<QuotationFields>;
  quotationNumber: string;
  customer: AirtableRecord<CustomerFields>;
  enquiry: AirtableRecord<EnquiryFields>;
  buyerBlock: string;
  templateCode: "domestic-standard" | "myanmar-proforma";
  folderUrl: string;
  lineItems: AirtableRecord<QuotationLineItemFields>[];
  reason: string;
}) {
  console.warn("[quotation-pdf] falling back to local PDF generation", {
    quotationId: input.quotation.id,
    quotationNumber: input.quotationNumber,
    reason: input.reason
  });

  return createPdfDocument({
    quotationRecordId: input.quotation.id,
    quotationNumber: input.quotationNumber,
    templateCode: input.templateCode,
    draftFormat: "PDF",
    customerName: input.customer.fields["Customer Name"] || input.enquiry.fields["Lead Name"] || "",
    customerFolderName: buildCustomerFolderName(input.customer),
    company: input.customer.fields.Company || input.enquiry.fields.Company || "",
    buyerBlock: input.buyerBlock,
    consigneeBlock: buildConsigneeBlock(input.enquiry),
    terms: buildQuotationTerms(input.templateCode),
    driveFolderName: input.folderUrl,
    lineItems: mapDraftLineItems(input.lineItems)
  });
}

async function fileExists(filePath: string) {
  try {
    await access(filePath);
    return true;
  } catch {
    return false;
  }
}

async function resolveQuotationSendDocument(
  quotation: AirtableRecord<QuotationFields>,
  customer: AirtableRecord<CustomerFields>
) {
  const quotationNumber = quotation.fields["Quotation Number"] || (await nextQuotationNumber());
  const customerFolderName = buildCustomerFolderName(customer);
  const finalPdf = getStoredDocumentArtifact(customerFolderName, quotationNumber, "pdf");
  const draftXlsx = getStoredDocumentArtifact(customerFolderName, quotationNumber, "xlsx");

  if (quotation.fields["Final PDF URL"]) {
    const pdfExists = await fileExists(finalPdf.filePath);
    return {
      kind: "pdf" as const,
      fileName: finalPdf.fileName,
      publicUrl: pdfExists ? finalPdf.fileUrl : quotation.fields["Final PDF URL"],
      attachment: pdfExists
        ? resolveDocumentAttachment(
            `${path.basename(path.dirname(finalPdf.filePath))}/${finalPdf.fileName}`,
            finalPdf.fileName,
            "application/pdf"
          )
        : undefined
    };
  }

  if (quotation.fields["Draft File URL"]) {
    const draftExists = await fileExists(draftXlsx.filePath);
    return {
      kind: "xlsx" as const,
      fileName: draftXlsx.fileName,
      publicUrl: draftExists ? draftXlsx.fileUrl : quotation.fields["Draft File URL"],
      attachment: draftExists
        ? resolveDocumentAttachment(
            `${path.basename(path.dirname(draftXlsx.filePath))}/${draftXlsx.fileName}`,
            draftXlsx.fileName,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          )
        : undefined
    };
  }

  throw new Error("No quotation document is available to send yet.");
}

async function loadQuotationDeliveryContext(quotationId: string) {
  const quotation = await getQuotationById(quotationId);
  const customerId = quotation.fields["Linked Customer"]?.[0];
  const enquiryId = quotation.fields["Linked Enquiry"]?.[0];

  if (!customerId || !enquiryId) {
    throw new Error("Quotation is missing linked customer or enquiry.");
  }

  const [customer, enquiry] = await Promise.all([
    getCustomerById(customerId),
    getRecord<EnquiryFields>(env.AIRTABLE_ENQUIRIES_TABLE, enquiryId)
  ]);

  return {
    quotation,
    customer,
    enquiry
  };
}

async function markQuotationSent(
  quotation: AirtableRecord<QuotationFields>,
  channel: "Email" | "WhatsApp"
) {
  const sentAt = new Date().toISOString();
  const fields: Record<string, unknown> = {
    Status: "Sent Quote",
    "Sent Date": sentAt,
    "Send Quotation": true
  };

  if (channel === "Email") {
    fields["Email Sent Date Time"] = sentAt;
  } else {
    fields["WhatsApp Sent Date Time"] = sentAt;
  }

  return updateRecordWithOptionalFieldFallback<QuotationFields>(env.AIRTABLE_QUOTATIONS_TABLE, {
    id: quotation.id,
    fields
  }, optionalQuotationFields);
}

export type WhatsAppStatusEvent = {
  messageId: string;
  status: string;
  recipientPhone: string;
  phoneNumberId: string;
  errorMessage: string;
};

export async function syncQuotationWhatsAppDeliveryStatus(event: WhatsAppStatusEvent) {
  const messageId = String(event.messageId || "").trim();
  if (!messageId) {
    return null;
  }

  const escapedMessageId = messageId.replace(/'/g, "\\'");
  const quotations = await listRecords<QuotationFields>(env.AIRTABLE_QUOTATIONS_TABLE, {
    fields: [
      "Status",
      "Linked Enquiry",
      "Sent Date",
      "WhatsApp Sent Date Time",
      "Quotation Number",
      "WhatsApp Outbound Message ID",
      "WhatsApp Delivery Status",
      "WhatsApp Delivery Updated At",
      "WhatsApp Delivery Error"
    ],
    filterByFormula: `{WhatsApp Outbound Message ID}='${escapedMessageId}'`,
    maxRecords: 10
  });

  const quotation = quotations[0];
  if (!quotation) {
    console.warn("[whatsapp-status] no quotation found for outbound message", {
      messageId,
      status: event.status,
      recipientPhone: event.recipientPhone
    });
    return null;
  }

  const normalizedStatus = String(event.status || "").trim().toLowerCase();
  const deliveryTimestamp = new Date().toISOString();
  const shouldMarkSent =
    normalizedStatus === "sent" ||
    normalizedStatus === "delivered" ||
    normalizedStatus === "read";
  const shouldMarkFailed = normalizedStatus === "failed";

  const quotationFields: Record<string, unknown> = {
    "WhatsApp Delivery Status": event.status,
    "WhatsApp Delivery Updated At": deliveryTimestamp,
    "WhatsApp Delivery Error": event.errorMessage || ""
  };

  if (shouldMarkSent) {
    quotationFields.Status = "Sent Quote";
    quotationFields["Sent Date"] = quotation.fields["Sent Date"] || deliveryTimestamp;
    quotationFields["WhatsApp Sent Date Time"] =
      quotation.fields["WhatsApp Sent Date Time"] || deliveryTimestamp;
  }

  if (shouldMarkFailed) {
    quotationFields["Send Quotation"] = false;
    quotationFields.Status = quotation.fields.Status || "Approved Quote";
  }

  const updated = await updateRecordWithOptionalFieldFallback<QuotationFields>(env.AIRTABLE_QUOTATIONS_TABLE, {
    id: quotation.id,
    fields: quotationFields
  }, optionalQuotationFields);

  if (shouldMarkSent) {
    const enquiryId = quotation.fields["Linked Enquiry"]?.[0];
    if (enquiryId) {
      const updatedEnquiry = await updateRecord<EnquiryFields>(env.AIRTABLE_ENQUIRIES_TABLE, {
        id: enquiryId,
        fields: {
          ...enquiryStatusFields("Sent Quote")
        }
      });
      await syncEnquiryToZohoAndPersist(updatedEnquiry);
    }
  }

  console.info("[whatsapp-status] quotation delivery status updated", {
    quotationId: quotation.id,
    quotationNumber: updated.fields["Quotation Number"] || quotation.id,
    messageId,
    status: event.status,
    recipientPhone: event.recipientPhone,
    errorMessage: event.errorMessage,
    shouldMarkSent,
    shouldMarkFailed
  });

  return updated;
}

async function createDraftForReadyEnquiry(
  enquiry: AirtableRecord<EnquiryFields>,
  customer: AirtableRecord<CustomerFields>,
  quotation: AirtableRecord<QuotationFields>,
  lineItems: AirtableRecord<QuotationLineItemFields>[]
) {
  const folder = await ensureFolder(customer);
  const draftLineItems = mapDraftLineItems(lineItems);
  const buyerBlock = quotation.fields["Buyer Block"] || buildBuyerBlock(enquiry);
  const templateCode =
    customer.fields["Customer Type"] === "Export" ? "myanmar-proforma" : "domestic-standard";
  const quotationNumber = quotation.fields["Quotation Number"] || (await nextQuotationNumber());

  const draft = await createDraftDocument({
    quotationRecordId: quotation.id,
    quotationNumber,
    templateCode,
    draftFormat: "XLSX",
    customerName: customer.fields["Customer Name"] || enquiry.fields["Lead Name"] || "",
    customerFolderName: buildCustomerFolderName(customer),
    company: customer.fields.Company || enquiry.fields.Company || "",
    buyerBlock,
    consigneeBlock: buildConsigneeBlock(enquiry),
    terms: buildQuotationTerms(templateCode),
    driveFolderName: folder.folderUrl,
    lineItems: draftLineItems
  });

  let draftFileUrl = draft.fileUrl;
  if (folder.folderId && isDriveConfigured()) {
    const upload = await uploadFileToFolder(
      draft.filePath,
      `${quotationNumber}.xlsx`,
      folder.folderId,
      {
        convertToGoogleSheet: true
      }
    );
    draftFileUrl = upload.fileUrl;
  }

  const nextQuotationStatus = isLockedQuotationStatus(quotation.fields.Status)
    ? String(quotation.fields.Status || "")
    : "Draft Quote";

  const updatedQuotation = await updateRecord<QuotationFields>(env.AIRTABLE_QUOTATIONS_TABLE, {
    id: quotation.id,
    fields: {
      "Draft File URL": draftFileUrl,
      "Draft Created Time": new Date().toISOString(),
      Status: nextQuotationStatus,
      "Quotation Number": quotationNumber
    }
  }).catch(async (error) => {
    const message = error instanceof Error ? error.message : "";
    const retryFields = stripOptionalFields(
      {
        "Draft File URL": draftFileUrl,
        "Draft Created Time": new Date().toISOString(),
        Status: nextQuotationStatus,
        "Quotation Number": quotationNumber
      },
      optionalQuotationFields,
      message
    );

    if (!retryFields) {
      throw error;
    }

    return updateRecord<QuotationFields>(env.AIRTABLE_QUOTATIONS_TABLE, {
      id: quotation.id,
      fields: retryFields
    });
  });

  const updatedEnquiry = await updateRecord<EnquiryFields>(env.AIRTABLE_ENQUIRIES_TABLE, {
    id: enquiry.id,
    fields: {
      ...enquiryStatusFields("Draft Quote"),
      "Linked Customer": linkedRecordIds(customer.id),
      Quotations: linkedRecordIds(quotation.id)
    }
  });
  await syncEnquiryToZohoAndPersist(updatedEnquiry);

  return {
    quotation: updatedQuotation,
    folder
  };
}

export async function refreshDraftForQuotation(quotationId: string) {
  const quotation = await getQuotationById(quotationId);
  const customerId = quotation.fields["Linked Customer"]?.[0];
  const enquiryId = quotation.fields["Linked Enquiry"]?.[0];

  if (!customerId || !enquiryId) {
    throw new Error("Quotation is missing linked customer or enquiry.");
  }

  const customer = await getCustomerById(customerId);
  const enquiry = await getRecord<EnquiryFields>(env.AIRTABLE_ENQUIRIES_TABLE, enquiryId);
  const lineItems = getLineItemsForQuotation(quotation.id, await listAllLineItems());

  if (!lineItems.length) {
    throw new Error("Quotation line items not found for draft generation.");
  }

  return createDraftForReadyEnquiry(enquiry, customer, quotation, lineItems);
}

export async function generateFinalPdfForQuotation(quotationId: string) {
  const quotation = await getQuotationById(quotationId);
  const customerId = quotation.fields["Linked Customer"]?.[0];
  const enquiryId = quotation.fields["Linked Enquiry"]?.[0];

  if (!customerId || !enquiryId) {
    throw new Error("Quotation is missing linked customer or enquiry.");
  }

  const customer = await getCustomerById(customerId);
  const enquiry = await getRecord<EnquiryFields>(env.AIRTABLE_ENQUIRIES_TABLE, enquiryId);
  const lineItems = getLineItemsForQuotation(quotation.id, await listAllLineItems());

  if (!lineItems.length) {
    throw new Error("Quotation line items not found for final PDF generation.");
  }

  const buyerBlock = quotation.fields["Buyer Block"] || buildBuyerBlock(enquiry);
  const templateCode =
    customer.fields["Customer Type"] === "Export" ? "myanmar-proforma" : "domestic-standard";
  const folder = await ensureFolder(customer);
  const quotationNumber = quotation.fields["Quotation Number"] || (await nextQuotationNumber());
  const localPdfArtifact = getStoredDocumentArtifact(buildCustomerFolderName(customer), quotationNumber, "pdf");
  const localDraftArtifact = getStoredDocumentArtifact(buildCustomerFolderName(customer), quotationNumber, "xlsx");

  let pdf;
  const draftDriveFileId = extractDriveFileId(String(quotation.fields["Draft File URL"] || ""));
  if (draftDriveFileId && isDriveConfigured()) {
    try {
      const pdfBuffer = await exportDriveFile(draftDriveFileId, "application/pdf");
      await mkdir(path.dirname(localPdfArtifact.filePath), { recursive: true });
      await writeFile(localPdfArtifact.filePath, pdfBuffer);
      pdf = {
        kind: "pdf" as const,
        quotationRecordId: quotation.id,
        quotationNumber,
        filePath: localPdfArtifact.filePath,
        payloadPath: "",
        fileUrl: localPdfArtifact.fileUrl,
        previewUrl: "",
        payloadUrl: "",
        message: "Final PDF exported from the live quotation sheet."
      };
    } catch (error) {
      const message = error instanceof Error ? error.message : "Unknown Drive export failure.";
      const draftExists = await fileExists(localDraftArtifact.filePath);
      if (!folder.folderId || !draftExists) {
        throw new Error(
          `Final PDF export must come from the live quotation sheet. Drive PDF export failed: ${message}`
        );
      }

      console.warn("[quotation-pdf] draft Drive export failed; converting stored XLSX to Google Sheet", {
        quotationId,
        quotationNumber,
        draftDriveFileId,
        message,
        stack: error instanceof Error ? error.stack : undefined
      });

      try {
        const convertedDraft = await uploadFileToFolder(
          localDraftArtifact.filePath,
          `${quotationNumber}.xlsx`,
          folder.folderId,
          {
            convertToGoogleSheet: true
          }
        );
        const convertedDraftFileId = extractDriveFileId(convertedDraft.fileUrl);
        if (!convertedDraftFileId) {
          throw new Error("Converted draft file URL did not contain a Drive file ID.");
        }

        await updateRecord<QuotationFields>(env.AIRTABLE_QUOTATIONS_TABLE, {
          id: quotation.id,
          fields: {
            "Draft File URL": convertedDraft.fileUrl
          }
        });

        const pdfBuffer = await exportDriveFile(convertedDraftFileId, "application/pdf");
        await mkdir(path.dirname(localPdfArtifact.filePath), { recursive: true });
        await writeFile(localPdfArtifact.filePath, pdfBuffer);
        pdf = {
          kind: "pdf" as const,
          quotationRecordId: quotation.id,
          quotationNumber,
          filePath: localPdfArtifact.filePath,
          payloadPath: "",
          fileUrl: localPdfArtifact.fileUrl,
          previewUrl: "",
          payloadUrl: "",
          message: "Final PDF exported after converting the stored quotation sheet to Google Sheets."
        };
      } catch (conversionError) {
        pdf = await createFallbackPdfFromQuotationData({
          quotation,
          quotationNumber,
          customer,
          enquiry,
          buyerBlock,
          templateCode,
          folderUrl: folder.folderUrl,
          lineItems,
          reason:
            conversionError instanceof Error
              ? conversionError.message
              : "Google Drive draft conversion failed"
        });
      }
    }
  } else {
    pdf = await createFallbackPdfFromQuotationData({
      quotation,
      quotationNumber,
      customer,
      enquiry,
      buyerBlock,
      templateCode,
      folderUrl: folder.folderUrl,
      lineItems,
      reason: "Quotation has no live Google Sheet draft file ID."
    });
  }

  let finalPdfUrl = pdf.fileUrl;
  if (folder.folderId && isDriveConfigured()) {
    const upload = await uploadFileToFolder(pdf.filePath, `${quotationNumber}.pdf`, folder.folderId, {
      replaceExisting: true
    });
    finalPdfUrl = upload.fileUrl;
  }

  const finalPdfGeneratedAt = new Date().toISOString();

  const updatedQuotation = await updateRecord<QuotationFields>(env.AIRTABLE_QUOTATIONS_TABLE, {
    id: quotation.id,
    fields: {
      "Final PDF URL": finalPdfUrl,
      "Final PDF Generated At": finalPdfGeneratedAt,
      Status: "Approved Quote",
      "Quotation Number": quotationNumber
    }
  });

  const updatedEnquiry = await updateRecord<EnquiryFields>(env.AIRTABLE_ENQUIRIES_TABLE, {
    id: enquiry.id,
    fields: {
      ...enquiryStatusFields("Approved Quote"),
      Quotations: linkedRecordIds(quotation.id),
      "Linked Customer": linkedRecordIds(customer.id)
    }
  });
  await syncEnquiryToZohoAndPersist(updatedEnquiry);

  return {
    quotation: updatedQuotation,
    pdf: {
      ...pdf,
      fileUrl: finalPdfUrl
    }
  };
}

export async function sendQuotationEmail(quotationId: string) {
  if (!isSmtpConfigured()) {
    const state = getSmtpConfigState();
    throw new Error(
      `SMTP is not configured yet. Add SMTP_HOST, SMTP_PORT, SMTP_USER, and SMTP_APP_PASSWORD on the API server. Loaded state: host=${state.host}, port=${state.port}, user=${state.user}, password=${state.password}`
    );
  }

  const { quotation, customer, enquiry } = await loadQuotationDeliveryContext(quotationId);
  const recipientEmail = String(customer.fields.Email || enquiry.fields.Email || "").trim();

  if (!recipientEmail) {
    throw new Error("A valid customer email address is required.");
  }

  const document = await resolveQuotationSendDocument(quotation, customer);
  const customerName = String(customer.fields["Customer Name"] || "Customer").trim() || "Customer";
  const quotationNumber = quotation.fields["Quotation Number"] || quotation.id;

  await sendMail({
    to: recipientEmail,
    subject: `Quotation ${quotationNumber} from Resham Sutra`,
    text: [
      `Dear ${customerName},`,
      "",
      `Please find quotation ${quotationNumber} attached for your review.`,
      document.publicUrl ? `Reference link: ${document.publicUrl}` : "",
      "",
      "Regards,",
      "Resham Sutra Sales"
    ]
      .filter(Boolean)
      .join("\n"),
    attachments: document.attachment ? [document.attachment] : undefined
  });

  const updatedQuotation = await markQuotationSent(quotation, "Email");
  if (enquiry.id) {
    const updatedEnquiry = await updateRecord<EnquiryFields>(env.AIRTABLE_ENQUIRIES_TABLE, {
      id: enquiry.id,
      fields: {
        ...enquiryStatusFields("Sent Quote")
      }
    });
    await syncEnquiryToZohoAndPersist(updatedEnquiry);
  }
  return {
    quotation: updatedQuotation,
    recipient: recipientEmail,
    documentUrl: document.publicUrl
  };
}

export async function sendQuotationWhatsApp(quotationId: string) {
  if (!isWhatsAppOutboundConfigured()) {
    throw new Error("WhatsApp outbound is not configured yet.");
  }

  const { quotation, customer, enquiry } = await loadQuotationDeliveryContext(quotationId);
  const recipientPhone = String(
    customer.fields.WhatsApp || customer.fields.Phone || enquiry.fields.Phone || ""
  ).trim();

  if (!recipientPhone) {
    throw new Error("A valid customer WhatsApp number is required.");
  }

  console.info("[quotation-send-whatsapp] resolved delivery context", {
    quotationId,
    quotationNumber: quotation.fields["Quotation Number"] || quotation.id,
    customerId: customer.id,
    enquiryId: enquiry.id,
    rawRecipientPhone: recipientPhone
  });

  if (!customer.fields.WhatsApp || !customer.fields.Phone) {
    const syncedCustomerFields: Record<string, unknown> = {};
    if (!customer.fields.WhatsApp) {
      syncedCustomerFields.WhatsApp = recipientPhone;
    }
    if (!customer.fields.Phone) {
      syncedCustomerFields.Phone = recipientPhone;
    }
    if (Object.keys(syncedCustomerFields).length) {
      await updateRecord<CustomerFields>(env.AIRTABLE_CUSTOMERS_TABLE, {
        id: customer.id,
        fields: syncedCustomerFields
      });
    }
  }

  let document = await resolveQuotationSendDocument(quotation, customer);
  let quotationForSend = quotation;
  if (!document.attachment?.path) {
    console.warn("[quotation-send-whatsapp] local quotation file missing; regenerating final PDF before WhatsApp send", {
      quotationId,
      quotationNumber: quotation.fields["Quotation Number"] || quotation.id,
      documentUrl: document.publicUrl,
      filename: document.fileName
    });
    const regenerated = await generateFinalPdfForQuotation(quotationId);
    quotationForSend = regenerated.quotation;
    document = await resolveQuotationSendDocument(regenerated.quotation, customer);
  }

  if (!document.attachment?.path) {
    throw new Error(
      "WhatsApp PDF send requires a local PDF file for Meta media upload. Regenerate the final PDF and try again."
    );
  }

  const quotationNumber = quotationForSend.fields["Quotation Number"] || quotationForSend.id;

  const sendResult = await sendQuotationDocumentOnWhatsApp({
    to: recipientPhone,
    documentUrl: document.publicUrl,
    filename: document.fileName,
    caption: `Quotation ${quotationNumber} from Resham Sutra`,
    customerName: customer.fields["Customer Name"] || enquiry.fields["Lead Name"] || "Customer",
    quotationNumber,
    localFilePath: document.attachment?.path,
    contentType: document.attachment?.contentType || "application/pdf"
  });

  console.info("[quotation-send-whatsapp] outbound document send completed", {
    quotationId,
    quotationNumber,
    recipientPhone,
    documentUrl: document.publicUrl,
    sendResult
  });

  const outboundMessageId = String(sendResult?.messages?.[0]?.id || "").trim();
  const updatedQuotation = await updateRecordWithOptionalFieldFallback<QuotationFields>(
    env.AIRTABLE_QUOTATIONS_TABLE,
    {
      id: quotationForSend.id,
      fields: {
        "WhatsApp Outbound Message ID": outboundMessageId,
        "WhatsApp Delivery Status": "accepted",
        "WhatsApp Delivery Updated At": new Date().toISOString(),
        "WhatsApp Delivery Error": "",
        "Send Quotation": true
      }
    },
    optionalQuotationFields
  );
  console.info("[quotation-send-whatsapp] quotation marked as accepted by Meta pending delivery callback", {
    quotationId,
    quotationNumber,
    recipientPhone,
    outboundMessageId
  });
  return {
    quotation: updatedQuotation,
    recipient: recipientPhone,
    documentUrl: document.publicUrl
  };
}

export async function regenerateQuotationDraft(quotationId: string) {
  const result = await refreshDraftForQuotation(quotationId);

  return {
    quotation: result.quotation,
    draftFileUrl: result.quotation.fields["Draft File URL"] || "",
    driveFolderUrl: result.quotation.fields["Drive Folder URL"] || ""
  };
}

async function syncParsedEnquiriesWithLineItems() {
  const parsedEnquiries = await listRecords<EnquiryFields>(env.AIRTABLE_ENQUIRIES_TABLE, {
    fields: [
      "Enquiry ID",
      "Logged Date Time",
      "Lead Name",
      "Company",
      "Phone",
      "Email",
      "Address",
      "Destination Address",
      "State",
      "City",
      "Destination State",
      "Destination City",
      "Pincode",
      "Destination Pincode",
      "Parser Status",
      "Linked Customer",
      "Quotations",
      "Receiver WhatsApp Number"
    ],
    filterByFormula: "{Parser Status}='Parsed'",
    maxRecords: 100
  });

  if (!parsedEnquiries.length) {
    return;
  }

  const lineItems = await listAllLineItems();

  for (const enquiry of parsedEnquiries) {
    const customerId = enquiry.fields["Linked Customer"]?.[0] || "";
    const quotationId = enquiry.fields.Quotations?.[0] || "";

    if (!customerId || !quotationId) {
      console.log("[intake-processor] skipped parsed enquiry without explicit customer and quotation links", {
        enquiryRecordId: enquiry.id,
        enquiryId: enquiry.fields["Enquiry ID"] || "",
        hasLinkedCustomer: Boolean(customerId),
        hasQuotation: Boolean(quotationId)
      });
      continue;
    }

    const quotation = await getQuotationById(quotationId);
    if (isLockedQuotationStatus(quotation.fields.Status)) {
      continue;
    }

    const matchingItems = getLineItemsForQuotation(quotationId, lineItems);
    if (!matchingItems.length) {
      continue;
    }

    await updateRecord<EnquiryFields>(env.AIRTABLE_ENQUIRIES_TABLE, {
      id: enquiry.id,
      fields: {
      ...enquiryStatusFields("Draft Quote")
      }
    });

    await updateRecord<QuotationFields>(env.AIRTABLE_QUOTATIONS_TABLE, {
      id: quotationId,
      fields: {
      Status: "Draft Quote"
      }
    });
  }
}

async function syncNewEnquiriesWithCustomers() {
  const newEnquiries = await listRecords<EnquiryFields>(env.AIRTABLE_ENQUIRIES_TABLE, {
    fields: [
      "Enquiry ID",
      "Logged Date Time",
      "Lead Name",
      "Company",
      "Phone",
      "Email",
      "Address",
      "Destination Address",
      "State",
      "City",
      "Destination State",
      "Destination City",
      "Pincode",
      "Destination Pincode",
      "Parser Status",
      "Linked Customer",
      "Quotations",
      "Receiver WhatsApp Number"
    ],
    filterByFormula: "OR({Parser Status}='New', {Parser Status}='New Enquiries')",
    maxRecords: 100
  });

  for (const enquiry of newEnquiries) {
    const customerId = enquiry.fields["Linked Customer"]?.[0] || "";
    const quotationId = enquiry.fields.Quotations?.[0] || "";

    if (!customerId || !quotationId) {
      console.log("[intake-processor] skipped new enquiry without explicit customer and quotation links", {
        enquiryRecordId: enquiry.id,
        enquiryId: enquiry.fields["Enquiry ID"] || "",
        hasLinkedCustomer: Boolean(customerId),
        hasQuotation: Boolean(quotationId)
      });
      continue;
    }

    await updateRecord<EnquiryFields>(env.AIRTABLE_ENQUIRIES_TABLE, {
      id: enquiry.id,
      fields: {
        ...enquiryStatusFields("Parsed")
      }
    });
  }
}

export async function createCustomerForEnquiry(enquiryId: string) {
  const enquiry = await getRecord<EnquiryFields>(env.AIRTABLE_ENQUIRIES_TABLE, enquiryId);
  const customer = await ensureCustomer(enquiry);
  const folder = await ensureFolder(customer);
  let syncedEnquiry = await syncEnquiryAddressFromCustomer(enquiry, customer);

  const quotation = await ensureQuotationShell(syncedEnquiry, customer);

  const updatedEnquiry = await updateRecord<EnquiryFields>(env.AIRTABLE_ENQUIRIES_TABLE, {
    id: syncedEnquiry.id,
    fields: {
      ...enquiryStatusFields("Parsed"),
      "Linked Customer": linkedRecordIds(customer.id),
      Quotations: linkedRecordIds(quotation.id)
    }
  });

  const updatedQuotation = await updateRecord<QuotationFields>(env.AIRTABLE_QUOTATIONS_TABLE, {
    id: quotation.id,
    fields: {
      "Linked Enquiry": linkedRecordIds(updatedEnquiry.id),
      "Linked Customer": linkedRecordIds(customer.id),
      Status: isLockedQuotationStatus(quotation.fields.Status)
        ? String(quotation.fields.Status || "")
        : String(updatedEnquiry.fields["Parser Status"] || quotation.fields.Status || "Parsed")
    }
  });

  return {
    enquiry: updatedEnquiry,
    customer,
    quotation: updatedQuotation,
    folder
  };
}

export async function processPendingEnquiries() {
  await syncNewEnquiriesWithCustomers();
  await syncParsedEnquiriesWithLineItems();

  const readyEnquiries = await listRecords<EnquiryFields>(env.AIRTABLE_ENQUIRIES_TABLE, {
    fields: [
      "Enquiry ID",
      "Logged Date Time",
      "Lead Name",
      "Company",
      "Phone",
      "Email",
      "Address",
      "Destination Address",
      "State",
      "City",
      "Destination State",
      "Destination City",
      "Pincode",
      "Destination Pincode",
      "Parser Status",
      "Linked Customer",
      "Quotations",
      "Receiver WhatsApp Number"
    ],
    filterByFormula:
      "AND({Parser Status}='Draft Quote', {Linked Customer}!=BLANK(), {Quotations}!=BLANK())",
    maxRecords: 100
  });

  const lineItems = await listAllLineItems();
  const results: ProcessingResult[] = [];

  for (const enquiry of readyEnquiries) {
    try {
      const customerId = enquiry.fields["Linked Customer"]?.[0];
      const quotationId = enquiry.fields.Quotations?.[0];

      if (!customerId || !quotationId) {
        results.push({
          enquiryRecordId: enquiry.id,
          enquiryId: enquiry.fields["Enquiry ID"] || enquiry.id,
          status: "skipped",
          message: "Missing linked customer or quotation"
        });
        continue;
      }

      const matchingItems = getLineItemsForQuotation(quotationId, lineItems);
      if (!matchingItems.length) {
        results.push({
          enquiryRecordId: enquiry.id,
          enquiryId: enquiry.fields["Enquiry ID"] || enquiry.id,
          status: "skipped",
          message: "Quotation line items not found"
        });
        continue;
      }

      const customer = await getCustomerById(customerId);
      const quotation = await getQuotationById(quotationId);
      if (isLockedQuotationStatus(quotation.fields.Status)) {
        results.push({
          enquiryRecordId: enquiry.id,
          enquiryId: enquiry.fields["Enquiry ID"] || enquiry.id,
          customerRecordId: customer.id,
          quotationRecordId: quotation.id,
          status: "skipped",
          message: `Quotation status ${quotation.fields.Status || "unknown"} is locked`
        });
        continue;
      }

      if (String(quotation.fields["Draft File URL"] || "").trim()) {
        results.push({
          enquiryRecordId: enquiry.id,
          enquiryId: enquiry.fields["Enquiry ID"] || enquiry.id,
          customerRecordId: customer.id,
          quotationRecordId: quotation.id,
          driveFolderUrl: quotation.fields["Drive Folder URL"] || "",
          status: "skipped",
          message: "Draft already exists"
        });
        continue;
      }
      const { folder } = await createDraftForReadyEnquiry(enquiry, customer, quotation, matchingItems);

      results.push({
        enquiryRecordId: enquiry.id,
        enquiryId: enquiry.fields["Enquiry ID"] || enquiry.id,
        customerRecordId: customer.id,
        quotationRecordId: quotation.id,
        driveFolderUrl: folder.folderUrl,
        status: "processed"
      });
    } catch (error) {
      results.push({
        enquiryRecordId: enquiry.id,
        enquiryId: enquiry.fields["Enquiry ID"] || enquiry.id,
        status: "error",
        message: error instanceof Error ? error.message : "Unknown processing error"
      });
    }
  }

  return {
    processedCount: results.filter((result) => result.status === "processed").length,
    errorCount: results.filter((result) => result.status === "error").length,
    results
  };
}
