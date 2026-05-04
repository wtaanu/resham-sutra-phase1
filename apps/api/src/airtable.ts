import { env } from "./config.js";

type AirtableListResponse<TFields extends Record<string, unknown>> = {
  records: AirtableRecord<TFields>[];
  offset?: string;
};

export type AirtableRecord<TFields extends Record<string, unknown>> = {
  id: string;
  createdTime: string;
  fields: TFields;
};

type AirtableCreatePayload = {
  fields: Record<string, unknown>;
};

type AirtableUpdatePayload = {
  id: string;
  fields: Record<string, unknown>;
};

type AirtableDeletePayload = {
  id: string;
};

const AIRTABLE_API_BASE = `https://api.airtable.com/v0/${env.AIRTABLE_BASE_ID}`;

async function airtableRequest<TResponse>(
  path: string,
  init?: RequestInit
): Promise<TResponse> {
  const response = await fetch(`${AIRTABLE_API_BASE}${path}`, {
    ...init,
    headers: {
      Authorization: `Bearer ${env.AIRTABLE_TOKEN}`,
      "Content-Type": "application/json",
      ...(init?.headers ?? {})
    }
  });

  if (!response.ok) {
    const message = await response.text();
    throw new Error(`Airtable request failed (${response.status}): ${message}`);
  }

  return (await response.json()) as TResponse;
}

export async function listRecords<TFields extends Record<string, unknown>>(
  tableName: string,
  options?: {
    fields?: string[];
    filterByFormula?: string;
    maxRecords?: number;
    sort?: Array<{ field: string; direction?: "asc" | "desc" }>;
  }
) {
  const records: AirtableRecord<TFields>[] = [];
  let offset = "";

  do {
    const params = new URLSearchParams();

    if (options?.filterByFormula) {
      params.set("filterByFormula", options.filterByFormula);
    }

    if (options?.maxRecords) {
      params.set("maxRecords", String(options.maxRecords));
    }

    if (offset) {
      params.set("offset", offset);
    }

    options?.fields?.forEach((field) => params.append("fields[]", field));
    options?.sort?.forEach((sort, index) => {
      params.set(`sort[${index}][field]`, sort.field);
      params.set(`sort[${index}][direction]`, sort.direction ?? "asc");
    });

    const query = params.toString();
    const path = `/${encodeURIComponent(tableName)}${query ? `?${query}` : ""}`;

    const response = await airtableRequest<AirtableListResponse<TFields>>(path, {
      method: "GET"
    });

    records.push(...response.records);
    offset = response.offset || "";
  } while (offset && (!options?.maxRecords || records.length < options.maxRecords));

  return options?.maxRecords ? records.slice(0, options.maxRecords) : records;
}

export async function listRecordsPage<TFields extends Record<string, unknown>>(
  tableName: string,
  options?: {
    fields?: string[];
    filterByFormula?: string;
    offset?: string;
    pageSize?: number;
    sort?: Array<{ field: string; direction?: "asc" | "desc" }>;
  }
) {
  const params = new URLSearchParams();
  const pageSize = Math.min(100, Math.max(1, options?.pageSize ?? 25));

  params.set("pageSize", String(pageSize));

  if (options?.filterByFormula) {
    params.set("filterByFormula", options.filterByFormula);
  }

  if (options?.offset) {
    params.set("offset", options.offset);
  }

  options?.fields?.forEach((field) => params.append("fields[]", field));
  options?.sort?.forEach((sort, index) => {
    params.set(`sort[${index}][field]`, sort.field);
    params.set(`sort[${index}][direction]`, sort.direction ?? "asc");
  });

  const query = params.toString();
  const path = `/${encodeURIComponent(tableName)}?${query}`;
  const response = await airtableRequest<AirtableListResponse<TFields>>(path, {
    method: "GET"
  });

  return {
    records: response.records,
    offset: response.offset || "",
    pageSize
  };
}

export async function getRecord<TFields extends Record<string, unknown>>(
  tableName: string,
  recordId: string
) {
  return airtableRequest<AirtableRecord<TFields>>(
    `/${encodeURIComponent(tableName)}/${recordId}`,
    { method: "GET" }
  );
}

export async function createRecord<TFields extends Record<string, unknown>>(
  tableName: string,
  fields: Record<string, unknown>
) {
  const response = await airtableRequest<{ records: AirtableRecord<TFields>[] }>(
    `/${encodeURIComponent(tableName)}`,
    {
      method: "POST",
      body: JSON.stringify({
        records: [{ fields }] satisfies AirtableCreatePayload[]
      })
    }
  );

  return response.records[0];
}

export async function createRecords<TFields extends Record<string, unknown>>(
  tableName: string,
  fieldsList: Record<string, unknown>[]
) {
  const createdRecords: AirtableRecord<TFields>[] = [];

  for (let index = 0; index < fieldsList.length; index += 10) {
    const batch = fieldsList.slice(index, index + 10);
    const response = await airtableRequest<{ records: AirtableRecord<TFields>[] }>(
      `/${encodeURIComponent(tableName)}`,
      {
        method: "POST",
        body: JSON.stringify({
          records: batch.map((fields) => ({ fields })) satisfies AirtableCreatePayload[]
        })
      }
    );

    createdRecords.push(...response.records);
  }

  return createdRecords;
}

export async function updateRecord<TFields extends Record<string, unknown>>(
  tableName: string,
  payload: AirtableUpdatePayload
) {
  const response = await airtableRequest<{ records: AirtableRecord<TFields>[] }>(
    `/${encodeURIComponent(tableName)}`,
    {
      method: "PATCH",
      body: JSON.stringify({
        records: [payload] satisfies AirtableUpdatePayload[]
      })
    }
  );

  return response.records[0];
}

export async function deleteRecords(
  tableName: string,
  recordIds: string[]
) {
  for (let index = 0; index < recordIds.length; index += 10) {
    const batch = recordIds.slice(index, index + 10);
    const params = new URLSearchParams();
    batch.forEach((id) => params.append("records[]", id));
    const response = await airtableRequest<{ records: AirtableDeletePayload[] }>(
      `/${encodeURIComponent(tableName)}?${params.toString()}`,
      {
        method: "DELETE"
      }
    );

    void response;
  }
}
