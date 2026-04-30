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

export type BiginEnquiryPayload = {
  recordId?: string;
  enquiryId: string;
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

function buildMailingStreet(input: BiginEnquiryPayload) {
  return (
    String(input.destinationAddress || "").trim() ||
    String(input.address || "").trim()
  );
}

function buildMailingCity(input: BiginEnquiryPayload) {
  return (
    String(input.destinationCity || "").trim() ||
    String(input.city || "").trim()
  );
}

function buildMailingState(input: BiginEnquiryPayload) {
  return (
    String(input.destinationState || "").trim() ||
    String(input.state || "").trim()
  );
}

function buildMailingCode(input: BiginEnquiryPayload) {
  return (
    String(input.destinationPincode || "").trim() ||
    String(input.pincode || "").trim()
  );
}

function mapEnquiryToBiginRecord(input: BiginEnquiryPayload) {
  const name = splitLeadName(input.leadName);

  return {
    First_Name: name.firstName,
    Last_Name: name.lastName,
    Phone: input.phone,
    Mobile: input.receiverWhatsappNumber || input.phone,
    Email: input.email,
    Company: input.company,
    Description: [
      input.requirementSummary ? `Requirement: ${input.requirementSummary}` : "",
      input.parserStatus ? `Status: ${input.parserStatus}` : "",
      input.enquiryId ? `Enquiry ID: ${input.enquiryId}` : ""
    ]
      .filter(Boolean)
      .join("\n"),
    Mailing_Street: buildMailingStreet(input),
    Mailing_City: buildMailingCity(input),
    Mailing_State: buildMailingState(input),
    Mailing_Code: buildMailingCode(input)
  };
}

function sanitizeBiginPayload(record: Record<string, unknown>) {
  return Object.fromEntries(
    Object.entries(record).filter(([, value]) => value !== undefined && value !== null && value !== "")
  );
}

async function zohoBiginRequest(path: string, init: RequestInit) {
  const accessToken = await getZohoBiginAccessToken();
  const response = await fetch(`${runtimeApiDomain || env.ZOHO_BIGIN_API_DOMAIN}/bigin/v2/${path}`, {
    ...init,
    headers: {
      Authorization: `Zoho-oauthtoken ${accessToken}`,
      "Content-Type": "application/json",
      ...(init.headers || {})
    }
  });

  if (!response.ok) {
    const message = await response.text();
    throw new Error(`Zoho Bigin request failed (${response.status}): ${message}`);
  }

  return (await response.json()) as BiginRecordResponse;
}

export async function syncEnquiryToZohoBigin(input: BiginEnquiryPayload): Promise<BiginSyncResult | null> {
  if (!isZohoBiginConfigured()) {
    return null;
  }

  const moduleName = env.ZOHO_BIGIN_MODULE_API_NAME;
  const record = sanitizeBiginPayload(mapEnquiryToBiginRecord(input));
  const duplicateCheckFields = buildDuplicateCheckFields();

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
    });
  } else {
    result = await zohoBiginRequest(`${moduleName}/upsert`, {
      method: "POST",
      body: JSON.stringify({
        data: [record],
        duplicate_check_fields: duplicateCheckFields
      })
    });
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
