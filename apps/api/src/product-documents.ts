import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { randomUUID } from "node:crypto";
import { z } from "zod";
import { env } from "./config.js";

const uploadFileSchema = z.object({
  name: z.string().min(1),
  mimeType: z.string().default("application/octet-stream"),
  contentBase64: z.string().min(1)
});

export const uploadProductDocumentsSchema = z.object({
  files: z.array(uploadFileSchema).min(1)
});

export const sendProductDocumentsSchema = z.object({
  documentIds: z.array(z.string().min(1)).min(1)
});

export type ProductDocumentRecord = {
  id: string;
  productId: string;
  fileName: string;
  mimeType: string;
  relativePath: string;
  uploadedAt: string;
};

const productDocumentsDir = path.resolve(env.DOCUMENT_STORAGE_DIR, "product-documents");
const indexPath = path.join(productDocumentsDir, "index.json");

function safeSegment(value: string, fallback: string) {
  const sanitized = value.trim().replace(/[\\/:*?"<>|]/g, "-");
  return sanitized || fallback;
}

async function ensureStorage() {
  await mkdir(productDocumentsDir, { recursive: true });
}

async function readIndex() {
  try {
    const content = await readFile(indexPath, "utf8");
    return JSON.parse(content) as ProductDocumentRecord[];
  } catch {
    return [];
  }
}

async function writeIndex(records: ProductDocumentRecord[]) {
  await ensureStorage();
  await writeFile(indexPath, JSON.stringify(records, null, 2), "utf8");
}

function toPublicUrl(relativePath: string) {
  return `${env.PUBLIC_API_BASE_URL}/documents/${relativePath.replaceAll("\\", "/")}`;
}

export async function listProductDocuments(productIds?: string[]) {
  const records = await readIndex();
  const filterSet = productIds?.length ? new Set(productIds) : null;

  return records
    .filter((record) => (filterSet ? filterSet.has(record.productId) : true))
    .sort((left, right) => right.uploadedAt.localeCompare(left.uploadedAt))
    .map((record) => ({
      ...record,
      fileUrl: toPublicUrl(record.relativePath)
    }));
}

export async function getProductDocumentsByIds(documentIds: string[]) {
  const wanted = new Set(documentIds);
  const records = await listProductDocuments();
  return records.filter((record) => wanted.has(record.id));
}

export async function uploadProductDocuments(productId: string, payload: unknown) {
  const input = uploadProductDocumentsSchema.parse(payload);
  await ensureStorage();

  const existing = await readIndex();
  const targetDir = path.join(productDocumentsDir, safeSegment(productId, "product"));
  await mkdir(targetDir, { recursive: true });

  const created: ProductDocumentRecord[] = [];

  for (const file of input.files) {
    const id = randomUUID();
    const fileName = safeSegment(file.name, `${id}.bin`);
    const storedName = `${id}-${fileName}`;
    const absolutePath = path.join(targetDir, storedName);
    const relativePath = path.relative(path.resolve(env.DOCUMENT_STORAGE_DIR), absolutePath);

    await writeFile(absolutePath, Buffer.from(file.contentBase64, "base64"));

    created.push({
      id,
      productId,
      fileName,
      mimeType: file.mimeType,
      relativePath,
      uploadedAt: new Date().toISOString()
    });
  }

  await writeIndex([...existing, ...created]);

  return created.map((record) => ({
    ...record,
    fileUrl: toPublicUrl(record.relativePath)
  }));
}
