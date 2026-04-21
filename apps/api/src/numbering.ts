import { listRecords } from "./airtable.js";

const sequenceLocks = new Map<string, Promise<void>>();

function getIndiaDateParts(date = new Date()) {
  const formatter = new Intl.DateTimeFormat("en-CA", {
    timeZone: "Asia/Kolkata",
    year: "numeric",
    month: "2-digit",
    day: "2-digit"
  });

  const parts = formatter.formatToParts(date);
  const year = Number(parts.find((part) => part.type === "year")?.value || date.getUTCFullYear());
  const month = Number(parts.find((part) => part.type === "month")?.value || date.getUTCMonth() + 1);

  return { year, month };
}

function getFinancialYearLabel(year: number, month: number) {
  const startYear = month >= 4 ? year : year - 1;
  const endYearShort = String((startYear + 1) % 100).padStart(2, "0");
  return `${startYear}-${endYearShort}`;
}

function buildPrefix(prefix: string, month: number, financialYear: string) {
  return `${prefix}-${String(month).padStart(2, "0")}-${financialYear}`;
}

function buildIdentifier(prefix: string, runningNumber: number, month: number, financialYear: string) {
  return `${prefix}-${String(runningNumber).padStart(3, "0")}-${String(month).padStart(2, "0")}-${financialYear}`;
}

function extractRunningNumber(value: string, prefix: string, month: number, financialYear: string) {
  const regex = new RegExp(`^${prefix}-(\\d{3})-${String(month).padStart(2, "0")}-${financialYear}$`);
  const match = value.match(regex);
  return match ? Number(match[1]) : 0;
}

async function nextRunningNumber(
  tableName: string,
  fieldName: string,
  prefix: string
) {
  const { year, month } = getIndiaDateParts();
  const financialYear = getFinancialYearLabel(year, month);

  const records = await listRecords<Record<string, unknown>>(tableName, {
    fields: [fieldName],
    maxRecords: 1000
  });

  const maxExisting = records.reduce((currentMax, record) => {
    const value = String(record.fields[fieldName] || "");
    return Math.max(currentMax, extractRunningNumber(value, prefix, month, financialYear));
  }, 0);

  return buildIdentifier(prefix, maxExisting + 1, month, financialYear);
}

function buildExactMatchFormula(fieldName: string, value: string) {
  const escapedFieldName = fieldName.replace(/}/g, "\\}");
  const escapedValue = value.replace(/'/g, "\\'");
  return `{${escapedFieldName}}='${escapedValue}'`;
}

async function withSequenceLock<T>(lockKey: string, task: () => Promise<T>) {
  const previous = sequenceLocks.get(lockKey) ?? Promise.resolve();
  let release!: () => void;
  const current = new Promise<void>((resolve) => {
    release = resolve;
  });

  sequenceLocks.set(lockKey, previous.then(() => current));

  await previous;

  try {
    return await task();
  } finally {
    release();
    if (sequenceLocks.get(lockKey) === current) {
      sequenceLocks.delete(lockKey);
    }
  }
}

export async function nextEnquiryNumber(tableName: string, fieldName: string) {
  return nextRunningNumber(tableName, fieldName, "ENQ");
}

export async function nextQuotationNumber(tableName: string, fieldName: string) {
  return nextRunningNumber(tableName, fieldName, "QTN");
}

export async function createRecordWithUniqueNumber<TFields extends Record<string, unknown>, TResult>({
  tableName,
  fieldName,
  prefix,
  create
}: {
  tableName: string;
  fieldName: string;
  prefix: "ENQ" | "QTN";
  create: (identifier: string) => Promise<TResult>;
}) {
  const lockKey = `${tableName}:${fieldName}:${prefix}`;

  return withSequenceLock(lockKey, async () => {
    for (let attempt = 0; attempt < 5; attempt += 1) {
      const identifier = await nextRunningNumber(tableName, fieldName, prefix);
      const existing = await listRecords<TFields>(tableName, {
        fields: [fieldName],
        filterByFormula: buildExactMatchFormula(fieldName, identifier),
        maxRecords: 1
      });

      if (existing.length) {
        continue;
      }

      return create(identifier);
    }

    throw new Error(`Failed to allocate a unique ${prefix} number after multiple attempts.`);
  });
}

export function currentNumberingContext() {
  const { year, month } = getIndiaDateParts();
  return {
    month: String(month).padStart(2, "0"),
    financialYear: getFinancialYearLabel(year, month),
    enquiryPrefix: buildPrefix("ENQ", month, getFinancialYearLabel(year, month)),
    quotationPrefix: buildPrefix("QTN", month, getFinancialYearLabel(year, month))
  };
}
