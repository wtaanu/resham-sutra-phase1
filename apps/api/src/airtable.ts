import { env } from "./config.js";

type AirtableListResponse<TFields extends Record<string, unknown>> = {
  records: AirtableRecord<TFields>[];
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
  const params = new URLSearchParams();

  if (options?.filterByFormula) {
    params.set("filterByFormula", options.filterByFormula);
  }

  if (options?.maxRecords) {
    params.set("maxRecords", String(options.maxRecords));
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

  return response.records;
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
    const response = await airtableRequest<{ records: AirtableDeletePayload[] }>(
      `/${encodeURIComponent(tableName)}`,
      {
        method: "DELETE",
        body: JSON.stringify({
          records: batch.map((id) => ({ id })) satisfies AirtableDeletePayload[]
        })
      }
    );

    void response;
  }
}
