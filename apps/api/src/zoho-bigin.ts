import { env } from "./config.js";

type BiginTokenResponse = {
  access_token: string;
  expires_in: number;
  api_domain?: string;
  token_type?: string;
  refresh_token?: string;
  scope?: string;
};

type BiginRecordResponse = {
  data?: Array<{
    code?: string;
    status?: string;
    action?: string;
    message?: string;
    duplicate_field?: string | null;
    details?: {
      id?: string;
      Created_Time?: string;
      Modified_Time?: string;
      [key: string]: unknown;
    };
    [key: string]: unknown;
  }>;
};

type ZohoBiginRequestContext = {
  operation: "insert_or_upsert" | "update" | "lookup" | "token_refresh" | "auth_exchange";
  moduleName?: string;
  recordId?: string;
  enquiryId?: string;
  fieldNames?: string[];
  nonEmptyFields?: Record<string, string>;
  duplicateCheckFields?: string[];
};

export type BiginEnquiryPayload = {
  recordId?: string;
  dealRecordId?: string;
  enquiryId: string;
  loggedDateTime: string;
  leadName: string;
  company: string;
  phone: string;
  email: string;
  address: string;
  city: string;
  state: string;
  pincode: string;
  destinationAddress: string;
  destinationCity: string;
  destinationState: string;
  destinationPincode: string;
  parserStatus: string;
  requirementSummary: string;
  receiverWhatsappNumber: string;
  quotationRef?: string;
  productName?: string;
  productKey?: string;
  productDescription?: string;
  productUnitPrice?: number;
  productFreight?: number;
  productGstRate?: number;
  productCategory?: string;
};

export type BiginSyncResult = {
  recordId: string;
  dealRecordId: string;
  action: string;
  duplicateField: string;
  response: Record<string, unknown>;
};

export type ZohoBiginAuthExchangeResult = {
  accessToken: string;
  refreshToken: string;
  expiresIn: number;
  apiDomain: string;
  accountsDomain: string;
  scope: string;
  tokenType: string;
};

let cachedAccessToken = "";
let cachedAccessTokenExpiresAt = 0;
let inflightTokenPromise: Promise<string> | null = null;
let runtimeRefreshToken = "";
let runtimeApiDomain = "";
let runtimeAccountsDomain = "";

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

function describeBiginFieldValues(record: Record<string, unknown>) {
  return Object.fromEntries(
    Object.entries(record)
      .filter(([, value]) => value !== undefined && value !== null && value !== "")
      .map(([key, value]) => {
        const normalizedValue = typeof value === "string" ? value : JSON.stringify(value);
        return [key, String(normalizedValue || "").slice(0, 300)];
      })
  );
}

function logZohoBiginError(label: string, error: unknown, context: ZohoBiginRequestContext) {
  console.error(`[zoho-bigin] ${label}`, {
    ...context,
    error: serializeError(error)
  });
}

function isZohoBiginConfigured() {
  return Boolean(
    env.ZOHO_BIGIN_CLIENT_ID &&
      env.ZOHO_BIGIN_CLIENT_SECRET &&
      (runtimeRefreshToken || env.ZOHO_BIGIN_REFRESH_TOKEN) &&
      (runtimeApiDomain || env.ZOHO_BIGIN_API_DOMAIN)
  );
}

export function getZohoBiginConfigState() {
  return {
    clientId: Boolean(String(env.ZOHO_BIGIN_CLIENT_ID || "").trim()),
    clientSecret: Boolean(String(env.ZOHO_BIGIN_CLIENT_SECRET || "").trim()),
    refreshToken: Boolean(String(runtimeRefreshToken || env.ZOHO_BIGIN_REFRESH_TOKEN || "").trim()),
    accountsDomain: Boolean(String(runtimeAccountsDomain || env.ZOHO_BIGIN_ACCOUNTS_DOMAIN || "").trim()),
    apiDomain: Boolean(String(runtimeApiDomain || env.ZOHO_BIGIN_API_DOMAIN || "").trim()),
    redirectUri: Boolean(String(env.ZOHO_BIGIN_REDIRECT_URI || "").trim())
  };
}

export function setZohoBiginRuntimeCredentials(input: {
  refreshToken: string;
  apiDomain?: string;
  accountsDomain?: string;
}) {
  runtimeRefreshToken = String(input.refreshToken || "").trim();
  runtimeApiDomain = String(input.apiDomain || "").trim();
  runtimeAccountsDomain = String(input.accountsDomain || "").trim();
  cachedAccessToken = "";
  cachedAccessTokenExpiresAt = 0;
  inflightTokenPromise = null;
}

export function buildZohoBiginAuthUrl() {
  if (!env.ZOHO_BIGIN_CLIENT_ID) {
    throw new Error("ZOHO_BIGIN_CLIENT_ID is required to start Zoho Bigin OAuth.");
  }

  const query = new URLSearchParams({
    scope: env.ZOHO_BIGIN_SCOPE,
    client_id: env.ZOHO_BIGIN_CLIENT_ID,
    response_type: "code",
    access_type: "offline",
    prompt: "consent",
    redirect_uri: env.ZOHO_BIGIN_REDIRECT_URI
  });

  return `${env.ZOHO_BIGIN_ACCOUNTS_DOMAIN}/oauth/v2/auth?${query.toString()}`;
}

export async function exchangeZohoBiginAuthCode(code: string): Promise<ZohoBiginAuthExchangeResult> {
  const normalizedCode = String(code || "").trim();
  if (!normalizedCode) {
    throw new Error("Zoho Bigin OAuth code is required.");
  }

  if (!env.ZOHO_BIGIN_CLIENT_ID || !env.ZOHO_BIGIN_CLIENT_SECRET) {
    throw new Error("ZOHO_BIGIN_CLIENT_ID and ZOHO_BIGIN_CLIENT_SECRET are required.");
  }

  const accountsDomain = runtimeAccountsDomain || env.ZOHO_BIGIN_ACCOUNTS_DOMAIN;
  const response = await fetch(`${accountsDomain}/oauth/v2/token`, {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded"
    },
    body: new URLSearchParams({
      code: normalizedCode,
      client_id: env.ZOHO_BIGIN_CLIENT_ID,
      client_secret: env.ZOHO_BIGIN_CLIENT_SECRET,
      redirect_uri: env.ZOHO_BIGIN_REDIRECT_URI,
      grant_type: "authorization_code"
    })
  });

  if (!response.ok) {
    const message = await response.text();
    throw new Error(`Zoho Bigin auth code exchange failed (${response.status}): ${message}`);
  }

  const data = (await response.json()) as BiginTokenResponse;
  return {
    accessToken: String(data.access_token || ""),
    refreshToken: String(data.refresh_token || ""),
    expiresIn: Number(data.expires_in || 0),
    apiDomain: String(data.api_domain || env.ZOHO_BIGIN_API_DOMAIN || ""),
    accountsDomain,
    scope: String(data.scope || env.ZOHO_BIGIN_SCOPE || ""),
    tokenType: String(data.token_type || "Zoho-oauthtoken")
  };
}

async function exchangeRefreshTokenForAccessToken() {
  const response = await fetch(`${runtimeAccountsDomain || env.ZOHO_BIGIN_ACCOUNTS_DOMAIN}/oauth/v2/token`, {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded"
    },
    body: new URLSearchParams({
      refresh_token: runtimeRefreshToken || env.ZOHO_BIGIN_REFRESH_TOKEN,
      client_id: env.ZOHO_BIGIN_CLIENT_ID,
      client_secret: env.ZOHO_BIGIN_CLIENT_SECRET,
      grant_type: "refresh_token"
    })
  });

  if (!response.ok) {
    const message = await response.text();
    throw new Error(`Zoho Bigin token refresh failed (${response.status}): ${message}`);
  }

  const data = (await response.json()) as BiginTokenResponse;
  cachedAccessToken = data.access_token;
  cachedAccessTokenExpiresAt = Date.now() + Math.max((data.expires_in || 3600) - 60, 60) * 1000;
  return cachedAccessToken;
}

async function getZohoBiginAccessToken() {
  if (!isZohoBiginConfigured()) {
    throw new Error("Zoho Bigin is not configured");
  }

  if (cachedAccessToken && Date.now() < cachedAccessTokenExpiresAt) {
    return cachedAccessToken;
  }

  if (!inflightTokenPromise) {
    inflightTokenPromise = exchangeRefreshTokenForAccessToken().finally(() => {
      inflightTokenPromise = null;
    });
  }

  return inflightTokenPromise;
}

function buildDuplicateCheckFields() {
  return env.ZOHO_BIGIN_DUPLICATE_CHECK_FIELDS.split(",")
    .map((value) => value.trim())
    .filter(Boolean);
}

function splitLeadName(value: string) {
  const trimmed = String(value || "").trim();
  if (!trimmed) {
    return {
      firstName: "",
      lastName: "Unknown"
    };
  }

  const parts = trimmed.split(/\s+/).filter(Boolean);
  if (parts.length === 1) {
    return {
      firstName: "",
      lastName: parts[0]
    };
  }

  return {
    firstName: parts.slice(0, -1).join(" "),
    lastName: parts[parts.length - 1]
  };
}

function buildMailingCity(input: BiginEnquiryPayload) {
  return String(input.city || "").trim();
}

function buildMailingState(input: BiginEnquiryPayload) {
  return String(input.state || "").trim();
}

function buildMailingCode(input: BiginEnquiryPayload) {
  return String(input.pincode || "").trim();
}

function setConfiguredField(record: Record<string, unknown>, fieldApiName: string, value: unknown) {
  const normalizedFieldApiName = String(fieldApiName || "").trim();
  if (normalizedFieldApiName) {
    record[normalizedFieldApiName] = value;
  }
}

function mapEnquiryToBiginRecord(input: BiginEnquiryPayload, accountId = "") {
  const name = splitLeadName(input.leadName);
  const record: Record<string, unknown> = {
    First_Name: name.firstName,
    Last_Name: name.lastName,
    Email: input.email,
    Mobile: input.phone,
    Mailing_Street: input.address,
    Mailing_City: buildMailingCity(input),
    Mailing_State: buildMailingState(input),
    Mailing_Country: env.ZOHO_BIGIN_DEFAULT_MAILING_COUNTRY,
    Mailing_Zip: buildMailingCode(input),
    Description: [
      input.requirementSummary ? `Requirement: ${input.requirementSummary}` : "",
      input.parserStatus ? `Status: ${input.parserStatus}` : "",
      input.enquiryId ? `Enquiry ID: ${input.enquiryId}` : ""
    ]
      .filter(Boolean)
      .join("\n")
  };

  if (accountId) {
    setConfiguredField(record, env.ZOHO_BIGIN_COMPANY_FIELD_API_NAME, { id: accountId });
  }

  if (env.ZOHO_BIGIN_PHONE_FIELD_API_NAME !== "Mobile") {
    setConfiguredField(record, env.ZOHO_BIGIN_PHONE_FIELD_API_NAME, input.phone);
  }


  return record;
}

function sanitizeBiginPayload(record: Record<string, unknown>) {
  return Object.fromEntries(
    Object.entries(record).filter(([, value]) => value !== undefined && value !== null && value !== "")
  );
}

async function zohoBiginRequest(path: string, init: RequestInit, context: ZohoBiginRequestContext) {
  const url = `${runtimeApiDomain || env.ZOHO_BIGIN_API_DOMAIN}/bigin/v2/${path}`;
  console.info("[zoho-bigin] request start", {
    ...context,
    method: init.method || "GET",
    path,
    url,
    hasBody: Boolean(init.body)
  });

  try {
    const accessToken = await getZohoBiginAccessToken();
    const response = await fetch(url, {
      ...init,
      headers: {
        Authorization: `Zoho-oauthtoken ${accessToken}`,
        "Content-Type": "application/json",
        ...(init.headers || {})
      }
    });

    const text = await response.text();
    let data: BiginRecordResponse = {};
    if (text) {
      try {
        data = JSON.parse(text) as BiginRecordResponse;
      } catch (error) {
        logZohoBiginError("failed to parse response json", error, {
          ...context,
          moduleName: context.moduleName,
          recordId: context.recordId
        });
      }
    }

    if (!response.ok) {
      const requestError = new Error(`Zoho Bigin request failed (${response.status}): ${text}`);
      logZohoBiginError("request failed", requestError, {
        ...context,
        moduleName: context.moduleName,
        recordId: context.recordId
      });
      throw requestError;
    }

    const failedRows = data.data?.filter((row) => String(row.status || "").toLowerCase() === "error") || [];
    if (failedRows.length) {
      const rowError = new Error(`Zoho Bigin returned row errors: ${JSON.stringify(failedRows)}`);
      logZohoBiginError("row-level error", rowError, {
        ...context,
        moduleName: context.moduleName,
        recordId: context.recordId
      });
      throw rowError;
    }

    console.info("[zoho-bigin] request success", {
      ...context,
      responseStatus: response.status,
      responseRows: data.data?.map((row) => ({
        status: row.status,
        code: row.code,
        action: row.action,
        duplicateField: row.duplicate_field,
        id: row.details?.id
      }))
    });

    return data;
  } catch (error) {
    logZohoBiginError("request exception", error, context);
    throw error;
  }
}

async function ensureZohoBiginAccount(company: string, enquiryId: string) {
  const accountName = String(company || "").trim();
  if (!accountName) {
    return "";
  }

  const moduleName = env.ZOHO_BIGIN_ACCOUNT_MODULE_API_NAME;
  const accountNameField = env.ZOHO_BIGIN_ACCOUNT_NAME_FIELD_API_NAME;
  const record = sanitizeBiginPayload({
    [accountNameField]: accountName
  });
  const context: ZohoBiginRequestContext = {
    operation: "insert_or_upsert",
    moduleName,
    enquiryId,
    fieldNames: Object.keys(record),
    nonEmptyFields: describeBiginFieldValues(record),
    duplicateCheckFields: [accountNameField]
  };

  console.info("[zoho-bigin] prepared account lookup payload", context);
  const result = await zohoBiginRequest(`${moduleName}/upsert`, {
    method: "POST",
    body: JSON.stringify({
      data: [record],
      duplicate_check_fields: [accountNameField]
    })
  }, context);

  const first = result.data?.[0];
  const accountId = String(first?.details?.id || "");
  if (!accountId) {
    throw new Error(`Zoho Bigin account upsert did not return an account ID for ${accountName}.`);
  }

  return accountId;
}

function addDaysIsoDate(days: number) {
  const date = new Date();
  date.setDate(date.getDate() + days);
  return date.toISOString().slice(0, 10);
}

function buildDuplicateCheckFieldsFrom(value: string) {
  return value
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean);
}

function deriveBiginDealStage(parserStatus: string) {
  const normalized = String(parserStatus || "").trim().toLowerCase();
  if (normalized === "approved quote" || normalized === "sent quote" || normalized === "ordered") {
    return env.ZOHO_BIGIN_DEAL_STAGE_QUOTED_VALUE;
  }
  if (
    normalized === "draft quote" ||
    normalized === "xls created" ||
    normalized === "xlsx created" ||
    normalized === "draft created"
  ) {
    return env.ZOHO_BIGIN_DEAL_STAGE_TO_QUOTE_VALUE;
  }
  return env.ZOHO_BIGIN_DEAL_STAGE_ENQUIRY_VALUE;
}

function shouldAttachQuotationRef(parserStatus: string) {
  return deriveBiginDealStage(parserStatus) === env.ZOHO_BIGIN_DEAL_STAGE_QUOTED_VALUE;
}

function buildBiginDealName(input: BiginEnquiryPayload) {
  return [
    input.enquiryId,
    input.company || input.leadName,
    input.productName || input.productKey || input.requirementSummary
  ]
    .map((value) => String(value || "").trim())
    .filter(Boolean)
    .join(" - ")
    .slice(0, 120) || input.enquiryId;
}

function buildZohoCriteria(fieldApiName: string, value: string) {
  const normalizedValue = String(value || "")
    .trim()
    .replace(/\r?\n/g, " ")
    .replace(/,/g, "\\,")
    .replace(/\)/g, "\\)");
  return `(${fieldApiName}:equals:${normalizedValue})`;
}

function getBiginRecordId(record: { id?: string; details?: { id?: string } } | undefined) {
  return String(record?.id || record?.details?.id || "");
}

async function findZohoBiginProductByWord(input: BiginEnquiryPayload, value: string) {
  const normalizedValue = String(value || "").trim().replace(/\r?\n/g, " ");
  if (!normalizedValue) {
    return "";
  }

  const moduleName = env.ZOHO_BIGIN_PRODUCT_MODULE_API_NAME;
  const context: ZohoBiginRequestContext = {
    operation: "lookup",
    moduleName,
    enquiryId: input.enquiryId,
    nonEmptyFields: { word: normalizedValue }
  };

  try {
    console.info("[zoho-bigin] product word lookup start", context);
    const result = await zohoBiginRequest(
      `${moduleName}/search?word=${encodeURIComponent(normalizedValue)}`,
      { method: "GET" },
      context
    );
    const productId = getBiginRecordId(result.data?.[0]);
    if (productId) {
      console.info("[zoho-bigin] product word lookup matched", {
        ...context,
        productId
      });
    }
    return productId;
  } catch (error) {
    console.warn("[zoho-bigin] product word lookup failed; continuing without product link", {
      ...context,
      error: serializeError(error)
    });
    return "";
  }
}

async function findZohoBiginProductByField(input: BiginEnquiryPayload, fieldApiName: string, value: string) {
  const normalizedValue = String(value || "").trim().replace(/\r?\n/g, " ");
  if (!fieldApiName || !normalizedValue) {
    return "";
  }

  const moduleName = env.ZOHO_BIGIN_PRODUCT_MODULE_API_NAME;
  const context: ZohoBiginRequestContext = {
    operation: "lookup",
    moduleName,
    enquiryId: input.enquiryId,
    fieldNames: [fieldApiName],
    nonEmptyFields: { [fieldApiName]: normalizedValue }
  };

  try {
    console.info("[zoho-bigin] product criteria lookup start", context);
    const result = await zohoBiginRequest(
      `${moduleName}/search?criteria=${encodeURIComponent(buildZohoCriteria(fieldApiName, normalizedValue))}`,
      { method: "GET" },
      context
    );
    const productId = getBiginRecordId(result.data?.[0]);

    if (productId) {
      console.info("[zoho-bigin] product criteria lookup matched", {
        ...context,
        productId
      });
    }

    return productId;
  } catch (error) {
    console.warn("[zoho-bigin] product criteria lookup failed; trying word lookup", {
      ...context,
      error: serializeError(error)
    });
    return findZohoBiginProductByWord(input, normalizedValue);
  }
}

async function findZohoBiginProduct(input: BiginEnquiryPayload) {
  if (!env.ZOHO_BIGIN_DEAL_PRODUCT_FIELD_API_NAME) {
    return "";
  }

  const productCodeField = String(env.ZOHO_BIGIN_PRODUCT_CODE_FIELD_API_NAME || "").trim();
  const productNameField = String(env.ZOHO_BIGIN_PRODUCT_NAME_FIELD_API_NAME || "").trim();
  const productIdByCode = await findZohoBiginProductByField(input, productCodeField, input.productKey || "");
  if (productIdByCode) {
    return productIdByCode;
  }

  const productIdByName = await findZohoBiginProductByField(input, productNameField, input.productName || "");
  if (!productIdByName && (input.productKey || input.productName)) {
    console.warn("[zoho-bigin] product lookup did not match an existing Bigin product", {
      enquiryId: input.enquiryId,
      productKey: input.productKey || "",
      productName: input.productName || ""
    });
  }

  return productIdByName;
}

function mapEnquiryToBiginDealRecord(input: BiginEnquiryPayload, ids: {
  contactId: string;
  accountId: string;
  productId: string;
}) {
  const record: Record<string, unknown> = {
    [env.ZOHO_BIGIN_DEAL_NAME_FIELD_API_NAME]: buildBiginDealName(input),
    [env.ZOHO_BIGIN_DEAL_STAGE_FIELD_API_NAME]: deriveBiginDealStage(input.parserStatus),
    Description: [
      input.requirementSummary ? `Requirement: ${input.requirementSummary}` : "",
      input.parserStatus ? `Airtable Status: ${input.parserStatus}` : "",
      input.enquiryId ? `Airtable Enquiry ID: ${input.enquiryId}` : ""
    ]
      .filter(Boolean)
      .join("\n")
  };

  setConfiguredField(record, env.ZOHO_BIGIN_DEAL_PIPELINE_FIELD_API_NAME, env.ZOHO_BIGIN_DEAL_PIPELINE_NAME);
  setConfiguredField(record, env.ZOHO_BIGIN_DEAL_CLOSING_DATE_FIELD_API_NAME, addDaysIsoDate(30));
  if (ids.contactId) {
    setConfiguredField(record, env.ZOHO_BIGIN_DEAL_CONTACT_FIELD_API_NAME, { id: ids.contactId });
  }
  if (ids.accountId) {
    setConfiguredField(record, env.ZOHO_BIGIN_DEAL_ACCOUNT_FIELD_API_NAME, { id: ids.accountId });
  }
  if (ids.productId) {
    setConfiguredField(record, env.ZOHO_BIGIN_DEAL_PRODUCT_FIELD_API_NAME, { id: ids.productId });
  }
  if (input.quotationRef && shouldAttachQuotationRef(input.parserStatus)) {
    setConfiguredField(record, env.ZOHO_BIGIN_DEAL_QUOTATION_REF_FIELD_API_NAME, input.quotationRef);
  }

  return record;
}

async function syncZohoBiginDeal(input: BiginEnquiryPayload, ids: {
  contactId: string;
  accountId: string;
  productId: string;
}) {
  const moduleName = env.ZOHO_BIGIN_DEAL_MODULE_API_NAME;
  const record = sanitizeBiginPayload(mapEnquiryToBiginDealRecord(input, ids));
  const duplicateCheckFields = buildDuplicateCheckFieldsFrom(env.ZOHO_BIGIN_DEAL_DUPLICATE_CHECK_FIELDS);
  const context: ZohoBiginRequestContext = {
    operation: input.dealRecordId ? "update" : "insert_or_upsert",
    moduleName,
    recordId: input.dealRecordId,
    enquiryId: input.enquiryId,
    fieldNames: Object.keys(record),
    nonEmptyFields: describeBiginFieldValues(record),
    duplicateCheckFields
  };

  console.info("[zoho-bigin] prepared deal sync payload", context);

  const result = input.dealRecordId
    ? await zohoBiginRequest(`${moduleName}/${input.dealRecordId}`, {
        method: "PUT",
        body: JSON.stringify({
          data: [
            {
              id: input.dealRecordId,
              ...record
            }
          ]
        })
      }, context)
    : await zohoBiginRequest(`${moduleName}/upsert`, {
        method: "POST",
        body: JSON.stringify({
          data: [record],
          duplicate_check_fields: duplicateCheckFields
        })
      }, context);

  const first = result.data?.[0];
  const dealId = String(first?.details?.id || input.dealRecordId || "");
  if (!dealId) {
    throw new Error("Zoho Bigin deal sync did not return a record ID.");
  }

  return {
    dealId,
    action: String(first?.action || (input.dealRecordId ? "update" : "upsert")),
    response: first || {}
  };
}
export async function syncEnquiryToZohoBigin(input: BiginEnquiryPayload): Promise<BiginSyncResult | null> {
  if (!isZohoBiginConfigured()) {
    return null;
  }

  const moduleName = env.ZOHO_BIGIN_MODULE_API_NAME;
  const accountId = await ensureZohoBiginAccount(input.company, input.enquiryId);
  const record = sanitizeBiginPayload(mapEnquiryToBiginRecord(input, accountId));
  const duplicateCheckFields = buildDuplicateCheckFields();
  const context: ZohoBiginRequestContext = {
    operation: input.recordId ? "update" : "insert_or_upsert",
    moduleName,
    recordId: input.recordId,
    enquiryId: input.enquiryId,
    fieldNames: Object.keys(record),
    nonEmptyFields: describeBiginFieldValues(record),
    duplicateCheckFields
  };

  console.info("[zoho-bigin] prepared contact sync payload", context);

  let result: BiginRecordResponse;
  if (input.recordId) {
    result = await zohoBiginRequest(`${moduleName}/${input.recordId}`, {
      method: "PUT",
      body: JSON.stringify({
        data: [
          {
            id: input.recordId,
            ...record
          }
        ]
      })
    }, context);
  } else {
    result = await zohoBiginRequest(`${moduleName}/upsert`, {
      method: "POST",
      body: JSON.stringify({
        data: [record],
        duplicate_check_fields: duplicateCheckFields
      })
    }, context);
  }

  const first = result.data?.[0];
  const recordId = String(first?.details?.id || input.recordId || "");
  if (!recordId) {
    throw new Error("Zoho Bigin contact sync did not return a record ID.");
  }

  const productId = await findZohoBiginProduct(input);
  const deal = await syncZohoBiginDeal(input, {
    contactId: recordId,
    accountId,
    productId
  });

  return {
    recordId,
    dealRecordId: deal.dealId,
    action: String(first?.action || (input.recordId ? "update" : "upsert")),
    duplicateField: String(first?.duplicate_field || ""),
    response: {
      contact: first || {},
      deal: deal.response
    }
  };
}