import { listRecords } from "./airtable.js";
import { env } from "./config.js";
import { listProductDocuments } from "./product-documents.js";

type EnquiryFields = {
  "Enquiry ID"?: string;
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
  "Drive Folder URL"?: string;
  "Requirement Summary"?: string;
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
  "Linked Customer"?: string[];
  "Linked Enquiry"?: string[];
  Status?: string;
  "Draft File URL"?: string;
  "Final PDF URL"?: string;
  "Drive Folder URL"?: string;
  "Preferred Send Channel"?: string;
  "Sent Date"?: string;
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
};

type OrderFields = {
  "Order Number"?: string;
  "Linked Customer"?: string[];
  "Linked Quotation"?: string[];
  "Order Date"?: string;
  "Order Value"?: number;
  "Payment Status"?: string;
  "Delivery Status"?: string;
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
  "Source Sheet"?: string;
};

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

export async function getOperationsSnapshot() {
  const [enquiries, customers, quotations, quotationLineItems, orders, products] = await Promise.all([
    safeList<EnquiryFields>(env.AIRTABLE_ENQUIRIES_TABLE, {
      maxRecords: 50,
      sort: [{ field: "Created At", direction: "desc" }]
    }),
    safeList<CustomerFields>(env.AIRTABLE_CUSTOMERS_TABLE, {
      maxRecords: 50,
      sort: [{ field: "Client ID", direction: "asc" }]
    }),
    safeList<QuotationFields>(env.AIRTABLE_QUOTATIONS_TABLE, {
      maxRecords: 50,
      sort: [{ field: "Quotation Number", direction: "desc" }]
    }),
    safeList<QuotationLineItemFields>(env.AIRTABLE_QUOTATION_LINE_ITEMS_TABLE, {
      maxRecords: 500,
      fields: ["Quotation", "Linked Product"]
    }),
    safeList<OrderFields>(env.AIRTABLE_ORDERS_TABLE, {
      maxRecords: 50
    }),
    safeList<ProductFields>(env.AIRTABLE_PRODUCTS_TABLE, {
      maxRecords: 50
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

  return {
    actions: {
      interfaceUrl: env.AIRTABLE_INTERFACE_URL,
      enquiryFormUrl: env.AIRTABLE_ENQUIRY_FORM_URL,
      lineItemsFormUrl: env.AIRTABLE_LINE_ITEMS_FORM_URL,
      productsFormUrl: env.AIRTABLE_PRODUCTS_FORM_URL
    },
    metrics: [
      {
        label: "New Enquiries",
        value: enquiries.filter((record) => statusValue(record.fields) === "New").length
      },
      {
        label: "Ready For Draft",
        value: enquiries.filter((record) => statusValue(record.fields) === "Ready for Draft").length
      },
      { label: "Under Review", value: stageCount(quotations, "Status", "Ready for Review") },
      { label: "Approved", value: stageCount(quotations, "Status", "Approved") },
      { label: "Sent", value: stageCount(quotations, "Status", "Draft Sent") },
      { label: "Orders", value: orders.length }
    ],
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
        driveFolderUrl: record.fields["Drive Folder URL"] || "",
        requirementSummary: record.fields["Requirement Summary"] || ""
      };
    }),
    customers: customers.map((record) => ({
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
    })),
    quotations: quotations.map((record) => ({
      id: record.id,
      quotationNumber: record.fields["Quotation Number"] || record.id,
      linkedCustomerId: record.fields["Linked Customer"]?.[0] || "",
      linkedEnquiryId: record.fields["Linked Enquiry"]?.[0] || "",
      status: record.fields.Status || "",
      draftFileUrl: record.fields["Draft File URL"] || "",
      finalPdfUrl: record.fields["Final PDF URL"] || "",
      driveFolderUrl: record.fields["Drive Folder URL"] || "",
      preferredSendChannel: record.fields["Preferred Send Channel"] || "",
      sentDate: record.fields["Sent Date"] || "",
      lineItemCount: quotationLineItems.filter((item) => item.fields.Quotation?.includes(record.id)).length,
      sendQuotation: Boolean(record.fields["Send Quotation"]),
      sendReminder: Boolean(record.fields["Send Reminder"]),
      markAccepted: Boolean(record.fields["Mark Accepted"]),
      markRejected: Boolean(record.fields["Mark Rejected"]),
      reminderCount: Number(record.fields["Reminder Count"] || 0),
      lastReminderDate: record.fields["Last Reminder Date"] || "",
      nextReminderDate: record.fields["Next Reminder Date"] || ""
    })),
    orders: orders.map((record) => ({
      id: record.id,
      orderNumber: record.fields["Order Number"] || record.id,
      linkedCustomerId: record.fields["Linked Customer"]?.[0] || "",
      linkedQuotationId: record.fields["Linked Quotation"]?.[0] || "",
      orderDate: record.fields["Order Date"] || "",
      orderValue: record.fields["Order Value"] || 0,
      paymentStatus: record.fields["Payment Status"] || "",
      deliveryStatus: record.fields["Delivery Status"] || ""
    })),
    products: products.map((record) => ({
      id: record.id,
      productKey: record.fields["Product Key"] || record.id,
      model: record.fields.Model || "",
      name: record.fields["Product Name"] || "",
      narration: record.fields.Narration || "",
      bulkSalePrice: record.fields["Bulk Sale Price"] || 0,
      mrp: record.fields.MRP || 0,
      gstPercent: record.fields["GST %"] || 0,
      transportCharge: record.fields["Pkg & Transport"] || 0,
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
