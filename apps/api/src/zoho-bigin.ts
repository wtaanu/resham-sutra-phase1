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
  operation: "insert_or_upsert" | "update" | "token_refresh" | "auth_exchange";
  moduleName?: string;
  recordId?: string;
  enquiryId?: string;
  fieldNames?: string[];
  nonEmptyFields?: Record<string, string>;
  duplicateCheckFields?: string[];
};

export type BiginEnquiryPayload = {
  recordId?: string;
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
};

export type BiginSyncResult = {
  recordId: string;
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

function deriveBiginPipelineStage(parserStatus: string) {
  const normalized = String(parserStatus || "").trim().toLowerCase();
  if (
    normalized === "draft quote" ||
    normalized === "approved quote" ||
    normalized === "xls created" ||
    normalized === "xlsx created" ||
    normalized === "draft created"
  ) {
    return "To Quote";
  }
  if (
    normalized === "sent quote" ||
    normalized === "sent quotations" ||
    normalized === "sent quotation" ||
    normalized === "ordered"
  ) {
    return "Quoted";
  }
  return "Enquiry";
}

function setConfiguredField(record: Record<string, unknown>, fieldApiName: string, value: unknown) {
  const normalizedFieldApiName = String(fieldApiName || "").trim();
  if (normalizedFieldApiName) {
    record[normalizedFieldApiName] = value;
  }
}

function mapEnquiryToBiginRecord(input: BiginEnquiryPayload) {
  const name = splitLeadName(input.leadName);
  const record: Record<string, unknown> = {
    First_Name: name.firstName,
    Last_Name: name.lastName,
    Email: input.email,
    Mobile: input.phone,
    "Mailing Street": input.address,
    "Mailing City": buildMailingCity(input),
    "Mailing State": buildMailingState(input),
    "Mailing Country": env.ZOHO_BIGIN_DEFAULT_MAILING_COUNTRY,
    "Mailing Zip": buildMailingCode(input),
    Description: [
      input.requirementSummary ? `Requirement: ${input.requirementSummary}` : "",
      input.parserStatus ? `Status: ${input.parserStatus}` : "",
      input.enquiryId ? `Enquiry ID: ${input.enquiryId}` : ""
    ]
      .filter(Boolean)
      .join("\n")
  };

  setConfiguredField(record, env.ZOHO_BIGIN_COMPANY_FIELD_API_NAME, input.company);

  if (env.ZOHO_BIGIN_PHONE_FIELD_API_NAME !== "Mobile") {
    setConfiguredField(record, env.ZOHO_BIGIN_PHONE_FIELD_API_NAME, input.phone);
  }

  const pipelineStageFieldApiName = String(env.ZOHO_BIGIN_PIPELINE_STAGE_FIELD_API_NAME || "").trim();
  if (pipelineStageFieldApiName) {
    record[pipelineStageFieldApiName] = deriveBiginPipelineStage(input.parserStatus);
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

export async function syncEnquiryToZohoBigin(input: BiginEnquiryPayload): Promise<BiginSyncResult | null> {
  if (!isZohoBiginConfigured()) {
    return null;
  }

  const moduleName = env.ZOHO_BIGIN_MODULE_API_NAME;
  const record = sanitizeBiginPayload(mapEnquiryToBiginRecord(input));
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
    throw new Error("Zoho Bigin did not return a record ID.");
  }

  return {
    recordId,
    action: String(first?.action || (input.recordId ? "update" : "upsert")),
    duplicateField: String(first?.duplicate_field || ""),
    response: first || {}
  };
}
