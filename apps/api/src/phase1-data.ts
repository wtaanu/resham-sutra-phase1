export type ProductReference = {
  sku: string;
  category: string;
  model: string;
  title: string;
  variant: string;
  exFactory: number | null;
  freight: number | null;
  gstAmount: number | null;
  bulkSalePrice: number | null;
  mrp: number | null;
};

export type QuotationTemplateReference = {
  key: string;
  title: string;
  documentKind: string;
  consignorLines: string[];
  commercialFields: Array<{ label: string; value: string }>;
  columns: string[];
  terms: string[];
};

export type ProductSheetReference = {
  sheetName: string;
  purpose: string;
  rowCountHint: string;
};

export type BuildTrack = {
  name: string;
  status: "ready" | "in_progress" | "blocked";
  detail: string;
};

export type ClientInput = {
  label: string;
  status: "received" | "pending" | "scheduled";
  note: string;
};

export const productReferences: ProductReference[] = [
  {
    sku: "SRM-02",
    category: "Reeled for weft",
    model: "SRM-02",
    title: "Sonalika Reeling Machine (2 ends Muga Reeling Machine)",
    variant: "Electric",
    exFactory: 16330,
    freight: 2000,
    gstAmount: 3299.4,
    bulkSalePrice: 21629.4,
    mrp: 28800
  },
  {
    sku: "CRM-04",
    category: "Reeling Mulberry",
    model: "CRM-04",
    title: "Flexi Re-reeling Machine (4 ends Single Window Closed Type)",
    variant: "Electric",
    exFactory: 36000,
    freight: 2500,
    gstAmount: 6930,
    bulkSalePrice: 45430,
    mrp: 60600
  },
  {
    sku: "MRM-004",
    category: "Reeling Mulberry",
    model: "MRM-004",
    title: "Flexi Mulberry Reeling Machine (4 ends)",
    variant: "Electric",
    exFactory: null,
    freight: null,
    gstAmount: null,
    bulkSalePrice: null,
    mrp: null
  },
  {
    sku: "MRM-010",
    category: "Reeling Mulberry",
    model: "MRM-010",
    title: "Mulberry Reeling Machine (10 ends)",
    variant: "Electric",
    exFactory: 58500,
    freight: 5000,
    gstAmount: 11430,
    bulkSalePrice: 74930,
    mrp: 99900
  },
  {
    sku: "MRM-020",
    category: "Reeling Mulberry",
    model: "MRM-020",
    title: "Mulberry Reeling Machine (20 ends)",
    variant: "Electric",
    exFactory: 96500,
    freight: 7000,
    gstAmount: 18630,
    bulkSalePrice: 122130,
    mrp: 162800
  },
  {
    sku: "TCRM",
    category: "Reeling Mulberry",
    model: "TCRM",
    title: "Twin Charkha Reeling Machine (8 ends Mulberry Reeling)",
    variant: "Electric",
    exFactory: 32500,
    freight: 5000,
    gstAmount: 6750,
    bulkSalePrice: 44250,
    mrp: 59000
  },
  {
    sku: "BRM-01",
    category: "Reeling Tassar",
    model: "BRM-01",
    title: "Buniyaad Reeling Machine Electric Power",
    variant: "Electric",
    exFactory: 9750,
    freight: 1500,
    gstAmount: 2025,
    bulkSalePrice: 13275,
    mrp: 17700
  },
  {
    sku: "BRM-02",
    category: "Reeling Tassar",
    model: "BRM-02",
    title: "Buniyaad Reeling Machine Solar + Electric",
    variant: "Solar",
    exFactory: 18250,
    freight: 2000,
    gstAmount: 3645,
    bulkSalePrice: 23895,
    mrp: 31900
  },
  {
    sku: "CRM-01",
    category: "Reeling Tassar",
    model: "CRM-01",
    title: "Flexi Charkha Reeling (4 ends Reeling and Re-reeling Charkha)",
    variant: "Electric",
    exFactory: 25500,
    freight: 1200,
    gstAmount: 4806,
    bulkSalePrice: 31506,
    mrp: 42000
  },
  {
    sku: "MRTM-01",
    category: "Reeling Tassar",
    model: "MRTM-01",
    title: "MRTM Handy - Reeling + Twisting (4 Spindle Motorised Reeling & Twisting)",
    variant: "Electric",
    exFactory: 25350,
    freight: 2000,
    gstAmount: 4923,
    bulkSalePrice: 32273,
    mrp: 43000
  },
  {
    sku: "RR4E-01",
    category: "Reeling Tassar",
    model: "RR4E-01",
    title: "Re-reeling Machine Lite (4 ends Single Window Open Type)",
    variant: "Electric",
    exFactory: 16000,
    freight: 1500,
    gstAmount: 3150,
    bulkSalePrice: 20650,
    mrp: 27500
  }
];

export const quotationTemplateReferences: QuotationTemplateReference[] = [
  {
    key: "domestic-standard",
    title: "Quotation",
    documentKind: "Domestic quotation",
    consignorLines: [
      "Resham Sutra Pvt. Ltd. (Central Silk Board approved)",
      "738 Ghewra Mod",
      "Mundka Udhyog Nagar - North",
      "New Delhi - 110041",
      "Tel: 9811050909",
      "Email: info@reshamsutra.com"
    ],
    commercialFields: [
      { label: "GST", value: "07AAHCR4176J1Z0" },
      { label: "CIN", value: "U17116JH2015PTC002989" },
      { label: "Document Prefix", value: "QO/25-26/" }
    ],
    columns: [
      "S.No.",
      "Description",
      "Qty",
      "Rate Per Unit",
      "Pkg & Trnsprt",
      "GST %",
      "GST Amt",
      "Total Amount"
    ],
    terms: [
      "Price includes delivery up to destination as agreed in the quotation.",
      "Dispatch timeline is 8 - 10 weeks from order confirmation."
    ]
  },
  {
    key: "myanmar-proforma",
    title: "PERFORMA INVOICE",
    documentKind: "Export / Myanmar proforma",
    consignorLines: [
      "Resham Sutra Pvt. Ltd. (Central Silk Board approved)",
      "738 Ghewra Mod",
      "Mundka Udhyog Nagar - North",
      "New Delhi - 110041",
      "Tel: 9811050909",
      "Email: info@reshamsutra.com"
    ],
    commercialFields: [
      { label: "GSTIN No.", value: "07AAHCR4176J1Z0" },
      { label: "CIN No.", value: "U17116JH2015PTC002989" },
      { label: "Document Prefix", value: "QI/23-24/78" }
    ],
    columns: [
      "S.No.",
      "Description",
      "Qty",
      "Rate Per Unit",
      "Pkg & Trnsprt",
      "Unit Value",
      "GST %",
      "GST Amt",
      "Total Amount"
    ],
    terms: [
      "Use consignee details and unit-value column for export-oriented quotations.",
      "Keep proforma numbering separate from domestic quotation numbering."
    ]
  }
];

export const productSheetReferences: ProductSheetReference[] = [
  {
    sheetName: "pricelist",
    purpose: "Legacy machine master with category, model, narration, freight, GST, bulk price, and MRP.",
    rowCountHint: "Core reeling and spinning lines"
  },
  {
    sheetName: "Sheet1",
    purpose: "Additional processing and finishing machines not all present in the first sheet.",
    rowCountHint: "Processing, boilers, dryers, calendaring"
  },
  {
    sheetName: "Sheet2",
    purpose: "Short model grouping/reference sheet useful for aliases and parser mapping.",
    rowCountHint: "Keyword aliases only"
  },
  {
    sheetName: "RSPL",
    purpose: "Resham Sutra supplier sheet with purchase cost and supplier-specific pricing.",
    rowCountHint: "Internal pricing and margin reference"
  }
];

export const buildTracks: BuildTrack[] = [
  {
    name: "Enquiry Capture",
    status: "ready",
    detail:
      "Forwarded WhatsApp messages will be normalized, parsed, and saved into Airtable before syncing to Zoho Bigin."
  },
  {
    name: "Quotation Workspace",
    status: "ready",
    detail:
      "Airtable will remain the editable team-facing layer with manual Send Quotation and Send Reminder actions."
  },
  {
    name: "DigitalOcean + n8n",
    status: "in_progress",
    detail:
      "Deployment is planned for a self-hosted n8n instance on a DigitalOcean VPS using the client's available credits."
  },
  {
    name: "WhatsApp Numbers",
    status: "blocked",
    detail:
      "The client will procure and share up to two new WhatsApp business numbers for Meta Cloud API onboarding."
  }
];

export const clientInputs: ClientInput[] = [
  {
    label: "Meta Business access",
    status: "received",
    note: "Invite sent by the client."
  },
  {
    label: "Google Drive account",
    status: "received",
    note: "Shared for lead-wise folder setup and quotation storage."
  },
  {
    label: "Quotation Excel format",
    status: "received",
    note: "Primary quotation layout extracted from the shared workbook."
  },
  {
    label: "Product master",
    status: "received",
    note: "Pricing and machine references available from the sample workbook."
  },
  {
    label: "Sample enquiry messages",
    status: "received",
    note: "Forwarded-message examples available for parser heuristics."
  },
  {
    label: "Zoho Bigin walkthrough",
    status: "scheduled",
    note: "Client asked to connect during live login for final field mapping."
  },
  {
    label: "WhatsApp numbers",
    status: "pending",
    note: "Client will procure and share two numbers for Phase 1."
  }
];
