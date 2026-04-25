import { productReferences } from "./phase1-data.js";

export type ParsedEnquiry = {
  rawMessage: string;
  leadName: string | null;
  company: string | null;
  phone: string | null;
  cityOrState: string | null;
  productInterest: string | null;
  requestedAsset: string | null;
  notes: string[];
  confidence: "high" | "medium" | "low";
};

const productMatchers = [
  "mulberry silk reeling machine",
  "mulberry reeling machine",
  "buniyaad reeling machine",
  "wet spinning machine",
  "book making machine",
  "ambar charkha",
  "twin charkha reeling machine",
  "sonalika reeling machine"
];

const knownLocations = [
  "Nagpur",
  "Maharashtra",
  "Karnataka",
  "Assam",
  "Hyderabad"
];

const companyKeywords = [
  "ltd",
  "limited",
  "pvt",
  "private",
  "llp",
  "industries",
  "industry",
  "enterprise",
  "enterprises",
  "traders",
  "agency",
  "agencies",
  "mills",
  "exports",
  "export",
  "textiles",
  "silks",
  "silk",
  "fab",
  "fabrics",
  "company",
  "co."
];

function normalizePhone(input: string) {
  const digits = input.replace(/\D/g, "");
  if (digits.length < 10) {
    return null;
  }

  return digits.length > 10 ? digits.slice(-10) : digits;
}

function toTitleCase(input: string) {
  return input
    .split(/\s+/)
    .filter(Boolean)
    .map((part) => part.charAt(0).toUpperCase() + part.slice(1).toLowerCase())
    .join(" ");
}

function extractPhone(message: string) {
  const match = message.match(/(\+?\d[\d\s-]{8,}\d)/);
  return match ? normalizePhone(match[1]) : null;
}

function extractProduct(message: string) {
  const lower = message.toLowerCase();
  const directMatch = productMatchers.find((item) => lower.includes(item));

  if (directMatch) {
    return directMatch;
  }

  const catalogMatch = productReferences.find((product) =>
    lower.includes(product.title.toLowerCase().replace(/\s+/g, " "))
  );

  return catalogMatch?.title ?? null;
}

function extractRequestedAsset(message: string) {
  const lower = message.toLowerCase();

  if (lower.includes("video")) {
    return "Videos";
  }

  if (lower.includes("detail")) {
    return "Details";
  }

  return null;
}

function extractLocation(message: string) {
  const location = knownLocations.find((entry) =>
    message.toLowerCase().includes(entry.toLowerCase())
  );

  return location ?? null;
}

function extractName(message: string, phone: string | null) {
  const lines = message
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);

  for (const line of lines) {
    const cleaned = phone ? line.replace(phone, "").trim() : line;

    if (
      cleaned &&
      /^[A-Za-z][A-Za-z\s.]+$/.test(cleaned) &&
      !cleaned.toLowerCase().includes("machine") &&
      !cleaned.toLowerCase().includes("details") &&
      !cleaned.toLowerCase().includes("video")
    ) {
      return toTitleCase(cleaned.replace(/[:,-]+$/, ""));
    }
  }

  const nameAfterPhone = message.match(/(?:\+?\d[\d\s-]{8,}\d)\s*[:,-]?\s*([A-Za-z][A-Za-z\s.]+)/);
  if (nameAfterPhone) {
    return toTitleCase(nameAfterPhone[1]);
  }

  return null;
}

function extractCompany(message: string, leadName: string | null) {
  const lines = message
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);

  for (const line of lines) {
    const lower = line.toLowerCase();
    if (
      line !== leadName &&
      /^[A-Za-z][A-Za-z\s]+$/.test(line) &&
      companyKeywords.some((keyword) => lower.includes(keyword)) &&
      !knownLocations.some((location) =>
        lower.includes(location.toLowerCase())
      ) &&
      !lower.includes("machine") &&
      !lower.includes("details")
    ) {
      return toTitleCase(line);
    }
  }

  return null;
}

function buildConfidence(parsed: ParsedEnquiry) {
  const score =
    Number(Boolean(parsed.phone)) +
    Number(Boolean(parsed.leadName)) +
    Number(Boolean(parsed.productInterest)) +
    Number(Boolean(parsed.cityOrState));

  if (score >= 4) {
    return "high";
  }

  if (score >= 2) {
    return "medium";
  }

  return "low";
}

export function parseIncomingEnquiry(message: string): ParsedEnquiry {
  const phone = extractPhone(message);
  const leadName = extractName(message, phone);
  const productInterest = extractProduct(message);
  const cityOrState = extractLocation(message);
  const requestedAsset = extractRequestedAsset(message);
  const company = extractCompany(message, leadName);
  const notes: string[] = [];

  if (message.toLowerCase().includes("send videos")) {
    notes.push("Lead explicitly asked for product videos.");
  }

  if (message.toLowerCase().includes("detail")) {
    notes.push("Lead asked for more detailed product information.");
  }

  const parsed: ParsedEnquiry = {
    rawMessage: message,
    leadName,
    company,
    phone,
    cityOrState,
    productInterest,
    requestedAsset,
    notes,
    confidence: "low"
  };

  parsed.confidence = buildConfidence(parsed);

  return parsed;
}
