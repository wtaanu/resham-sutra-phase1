import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import ExcelJS from "exceljs";
import {
  PDFDocument,
  StandardFonts,
  rgb,
  type PDFFont,
  type PDFImage,
  type PDFPage
} from "pdf-lib";
import { z } from "zod";
import { env } from "./config.js";

const lineItemSchema = z.object({
  lineNo: z.union([z.string(), z.number()]).optional(),
  description: z.string().default(""),
  qty: z.coerce.number().default(0),
  rate: z.coerce.number().default(0),
  transport: z.coerce.number().default(0),
  gstPercent: z.coerce.number().default(0),
  gstAmount: z.coerce.number().default(0),
  totalAmount: z.coerce.number().default(0),
  unitValue: z.coerce.number().default(0)
});

const documentPayloadSchema = z.object({
  quotationRecordId: z.string().min(1),
  quotationNumber: z.string().min(1),
  templateCode: z.string().default("domestic-standard"),
  draftFormat: z.enum(["XLSX", "DOCX"]).default("XLSX"),
  customerName: z.string().default(""),
  customerFolderName: z.string().default(""),
  company: z.string().default(""),
  buyerBlock: z.string().default(""),
  consigneeBlock: z.string().default(""),
  terms: z.array(z.string()).default([]),
  driveFolderName: z.string().default(""),
  lineItems: z.array(lineItemSchema).min(1)
});

export type DocumentPayload = z.infer<typeof documentPayloadSchema>;

function slugify(value: string) {
  return value
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 80);
}

function safeFileName(value: string, fallback: string) {
  const trimmed = value.trim();
  const sanitized = trimmed.replace(/[\\/:*?"<>|]/g, "-");
  return sanitized || fallback;
}

function renderRows(lineItems: DocumentPayload["lineItems"]) {
  return lineItems
    .map(
      (item, index) => `
        <tr>
          <td>${item.lineNo ?? index + 1}</td>
          <td>${item.description}</td>
          <td>${item.qty}</td>
          <td>${item.rate}</td>
          <td>${item.transport}</td>
          <td>${item.gstPercent}</td>
          <td>${item.gstAmount}</td>
          <td>${item.totalAmount}</td>
        </tr>`
    )
    .join("");
}

function formatDateForTemplate(date = new Date()) {
  return date.toLocaleDateString("en-GB");
}

function splitLines(value: string, maxLines: number) {
  const lines = value
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);

  return Array.from({ length: maxLines }, (_, index) => lines[index] ?? "");
}

function setCell(worksheet: ExcelJS.Worksheet, ref: string, value: string | number) {
  worksheet.getCell(ref).value = value;
}

function unlockWorkbook(workbook: ExcelJS.Workbook) {
  for (const worksheet of workbook.worksheets) {
    try {
      worksheet.unprotect();
    } catch {
      // Ignore sheets that are already unprotected or unsupported by the parser.
    }

    if (worksheet.state === "hidden") {
      worksheet.state = "visible";
    }

    const model = worksheet.model as typeof worksheet.model & {
      sheetProtection?: unknown;
    };
    model.sheetProtection = undefined;
  }
}

function clearDomesticRows(worksheet: ExcelJS.Worksheet) {
  [13, 14, 15].forEach((rowNumber) => {
    ["A", "B", "C", "D", "E", "F", "G", "H"].forEach((column) => {
      worksheet.getCell(`${column}${rowNumber}`).value = "";
    });
  });
}

function clearMyanmarRows(worksheet: ExcelJS.Worksheet) {
  ["A", "B", "C", "D", "E", "F", "G", "H", "I"].forEach((column) => {
    worksheet.getCell(`${column}13`).value = "";
  });
}

function applyDomesticTemplate(
  worksheet: ExcelJS.Worksheet,
  payload: DocumentPayload
) {
  const buyerLines = splitLines(
    payload.buyerBlock || `${payload.customerName}\n${payload.company}`,
    6
  );

  setCell(worksheet, "E9", payload.quotationNumber);
  setCell(worksheet, "E10", formatDateForTemplate());
  setCell(worksheet, "B11", payload.company || payload.customerName);

  ["D3", "D4", "D5", "D6", "D7", "D8"].forEach((ref, index) => {
    setCell(worksheet, ref, buyerLines[index] ?? "");
  });

  clearDomesticRows(worksheet);
  payload.lineItems.slice(0, 3).forEach((item, index) => {
    const rowNumber = 13 + index;
    setCell(worksheet, `A${rowNumber}`, String(item.lineNo ?? index + 1));
    setCell(worksheet, `B${rowNumber}`, item.description);
    setCell(worksheet, `C${rowNumber}`, item.qty);
    setCell(worksheet, `D${rowNumber}`, item.rate);
    setCell(worksheet, `E${rowNumber}`, item.transport);
    setCell(worksheet, `F${rowNumber}`, item.gstPercent);
    setCell(worksheet, `G${rowNumber}`, item.gstAmount);
    setCell(worksheet, `H${rowNumber}`, item.totalAmount);
  });

  [17, 18, 19, 20, 21].forEach((rowNumber, index) => {
    if (payload.terms[index]) {
      setCell(worksheet, `A${rowNumber}`, payload.terms[index]);
    }
  });
}

function applyMyanmarTemplate(
  worksheet: ExcelJS.Worksheet,
  payload: DocumentPayload
) {
  const consigneeLines = splitLines(
    payload.consigneeBlock || payload.buyerBlock || `${payload.customerName}\n${payload.company}`,
    5
  );
  const item = payload.lineItems[0];

  setCell(worksheet, "F9", payload.quotationNumber);
  setCell(worksheet, "F10", formatDateForTemplate());
  setCell(worksheet, "B11", payload.company || payload.customerName);

  ["E3", "E4", "E5", "E6", "E7"].forEach((ref, index) => {
    setCell(worksheet, ref, consigneeLines[index] ?? "");
  });

  clearMyanmarRows(worksheet);
  setCell(worksheet, "A13", String(item?.lineNo ?? 1));
  setCell(worksheet, "B13", item?.description ?? "");
  setCell(worksheet, "C13", item?.qty ?? 0);
  setCell(worksheet, "D13", item?.rate ?? 0);
  setCell(worksheet, "E13", item?.transport ?? 0);
  setCell(worksheet, "F13", item?.unitValue ?? 0);
  setCell(worksheet, "G13", item?.gstPercent ?? 0);
  setCell(worksheet, "H13", item?.gstAmount ?? 0);
  setCell(worksheet, "I13", item?.totalAmount ?? 0);

  [16, 17, 18, 19, 20].forEach((rowNumber, index) => {
    if (payload.terms[index]) {
      setCell(worksheet, `A${rowNumber}`, payload.terms[index]);
    }
  });
}

async function writeXlsxDraft(payload: DocumentPayload, outputDir: string) {
  const templatePath = path.resolve(env.QUOTATION_TEMPLATE_DIR);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(templatePath);
  unlockWorkbook(workbook);

  const worksheet =
    payload.templateCode === "myanmar-proforma"
      ? workbook.getWorksheet("MYANMAR")
      : workbook.getWorksheet("jha (2)");

  if (!worksheet) {
    throw new Error(`Template sheet not found for template code: ${payload.templateCode}`);
  }

  if (payload.templateCode === "myanmar-proforma") {
    applyMyanmarTemplate(worksheet, payload);
  } else {
    applyDomesticTemplate(worksheet, payload);
  }

  const xlsxPath = path.join(outputDir, `${safeFileName(payload.quotationNumber, "quotation")}.xlsx`);
  await workbook.xlsx.writeFile(xlsxPath);
  return xlsxPath;
}

function escapePdfText(value: string) {
  return value.replace(/\\/g, "\\\\").replace(/\(/g, "\\(").replace(/\)/g, "\\)");
}

function formatMoney(value: number) {
  return new Intl.NumberFormat("en-IN", {
    minimumFractionDigits: 0,
    maximumFractionDigits: 2
  }).format(value);
}

function resolveLogoPath() {
  const currentDir = path.dirname(fileURLToPath(import.meta.url));
  return path.resolve(currentDir, "../../web/src/assets/resham-sutra-logo.png");
}

function wrapPdfText(text: string, maxWidth: number, font: PDFFont, size: number) {
  const words = String(text || "").split(/\s+/).filter(Boolean);
  if (!words.length) {
    return [""];
  }

  const lines: string[] = [];
  let current = words[0] || "";

  for (const word of words.slice(1)) {
    const candidate = `${current} ${word}`;
    if (font.widthOfTextAtSize(candidate, size) <= maxWidth) {
      current = candidate;
    } else {
      lines.push(current);
      current = word;
    }
  }

  lines.push(current);
  return lines;
}

function drawWrappedLines(
  page: PDFPage,
  lines: string[],
  options: { x: number; y: number; lineHeight: number; font: PDFFont; size: number; color?: ReturnType<typeof rgb> }
) {
  let currentY = options.y;
  lines.forEach((line) => {
    page.drawText(line, {
      x: options.x,
      y: currentY,
      font: options.font,
      size: options.size,
      color: options.color
    });
    currentY -= options.lineHeight;
  });
  return currentY;
}

async function loadLogo(pdfDoc: PDFDocument) {
  try {
    const logoBytes = await readFile(resolveLogoPath());
    return pdfDoc.embedPng(logoBytes);
  } catch {
    return null;
  }
}

function drawTableHeader(
  page: PDFPage,
  headers: Array<{ label: string; x: number; width: number }>,
  y: number,
  font: PDFFont
) {
  page.drawRectangle({
    x: 40,
    y: y - 18,
    width: 515,
    height: 20,
    color: rgb(0.95, 0.77, 0.18)
  });

  headers.forEach((header) => {
    page.drawText(header.label, {
      x: header.x + 4,
      y: y - 12,
      font,
      size: 8,
      color: rgb(0.12, 0.12, 0.16)
    });
  });
}

function drawTableRow(
  page: PDFPage,
  cells: Array<{ text: string; x: number; width: number; align?: "left" | "right" }>,
  y: number,
  font: PDFFont
) {
  page.drawRectangle({
    x: 40,
    y: y - 24,
    width: 515,
    height: 24,
    borderWidth: 0.5,
    borderColor: rgb(0.82, 0.85, 0.9),
    color: rgb(1, 1, 1)
  });

  cells.forEach((cell) => {
    const text = cell.text || "-";
    const textWidth = font.widthOfTextAtSize(text, 8);
    const x =
      cell.align === "right"
        ? cell.x + cell.width - textWidth - 4
        : cell.x + 4;

    page.drawText(text, {
      x,
      y: y - 15,
      font,
      size: 8,
      color: rgb(0.16, 0.18, 0.21)
    });
  });
}

async function writeSimplePdf(payload: DocumentPayload, outputDir: string) {
  const pdfPath = path.join(outputDir, `${safeFileName(payload.quotationNumber, "quotation")}.pdf`);
  const pdfDoc = await PDFDocument.create();
  const page = pdfDoc.addPage([595, 842]);
  const fontRegular = await pdfDoc.embedFont(StandardFonts.Helvetica);
  const fontBold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
  const logo = await loadLogo(pdfDoc);
  const brandYellow = rgb(0.95, 0.77, 0.18);
  const brandDark = rgb(0.14, 0.15, 0.18);
  const muted = rgb(0.4, 0.44, 0.5);

  page.drawRectangle({
    x: 0,
    y: 0,
    width: 595,
    height: 842,
    color: rgb(1, 1, 1)
  });

  page.drawRectangle({
    x: 40,
    y: 770,
    width: 515,
    height: 1,
    color: rgb(0.9, 0.9, 0.9)
  });

  if (logo) {
    const dimensions = logo.scale(0.14);
    page.drawImage(logo, {
      x: 42,
      y: 770,
      width: dimensions.width,
      height: dimensions.height
    });
  }

  page.drawText("Resham Sutra", {
    x: 118,
    y: 796,
    font: fontBold,
    size: 18,
    color: brandDark
  });
  page.drawText("Quotation", {
    x: 118,
    y: 777,
    font: fontRegular,
    size: 10,
    color: muted
  });

  page.drawText(payload.quotationNumber, {
    x: 418,
    y: 794,
    font: fontBold,
    size: 14,
    color: brandDark
  });
  page.drawText(`Date: ${formatDateForTemplate()}`, {
    x: 418,
    y: 777,
    font: fontRegular,
    size: 9,
    color: muted
  });

  page.drawRectangle({
    x: 40,
    y: 650,
    width: 250,
    height: 95,
    borderWidth: 1,
    borderColor: rgb(0.88, 0.89, 0.91),
    color: rgb(0.99, 0.99, 0.99)
  });
  page.drawRectangle({
    x: 305,
    y: 650,
    width: 250,
    height: 95,
    borderWidth: 1,
    borderColor: rgb(0.88, 0.89, 0.91),
    color: rgb(0.99, 0.99, 0.99)
  });

  page.drawText("Buyer", {
    x: 52,
    y: 730,
    font: fontBold,
    size: 10,
    color: brandDark
  });
  page.drawText("Consignee", {
    x: 317,
    y: 730,
    font: fontBold,
    size: 10,
    color: brandDark
  });

  const buyerLines = splitLines(
    payload.buyerBlock || `${payload.customerName}\n${payload.company}`,
    6
  );
  const consigneeLines = splitLines(
    payload.consigneeBlock || payload.buyerBlock || `${payload.customerName}\n${payload.company}`,
    6
  );

  drawWrappedLines(page, buyerLines, {
    x: 52,
    y: 713,
    lineHeight: 12,
    font: fontRegular,
    size: 9,
    color: brandDark
  });
  drawWrappedLines(page, consigneeLines, {
    x: 317,
    y: 713,
    lineHeight: 12,
    font: fontRegular,
    size: 9,
    color: brandDark
  });

  const headers = [
    { label: "No", x: 40, width: 35 },
    { label: "Description", x: 75, width: 205 },
    { label: "Qty", x: 280, width: 35 },
    { label: "Rate", x: 315, width: 60 },
    { label: "Transport", x: 375, width: 60 },
    { label: "GST %", x: 435, width: 40 },
    { label: "GST Amt", x: 475, width: 40 },
    { label: "Total", x: 515, width: 40 }
  ];
  let currentY = 620;
  drawTableHeader(page, headers, currentY, fontBold);
  currentY -= 26;

  payload.lineItems.slice(0, 12).forEach((item, index) => {
    const description = wrapPdfText(
      item.description || "Quotation item",
      196,
      fontRegular,
      8
    )[0] || "Quotation item";

    drawTableRow(
      page,
      [
        { text: String(item.lineNo ?? index + 1), x: 40, width: 35 },
        { text: description, x: 75, width: 205 },
        { text: String(item.qty || 0), x: 280, width: 35, align: "right" },
        { text: formatMoney(item.rate || 0), x: 315, width: 60, align: "right" },
        { text: formatMoney(item.transport || 0), x: 375, width: 60, align: "right" },
        { text: formatMoney(item.gstPercent || 0), x: 435, width: 40, align: "right" },
        { text: formatMoney(item.gstAmount || 0), x: 475, width: 40, align: "right" },
        { text: formatMoney(item.totalAmount || 0), x: 515, width: 40, align: "right" }
      ],
      currentY,
      fontRegular
    );
    currentY -= 24;
  });

  const totalValue = payload.lineItems.reduce((sum, item) => sum + Number(item.totalAmount || 0), 0);

  page.drawRectangle({
    x: 355,
    y: currentY - 10,
    width: 200,
    height: 24,
    color: brandYellow
  });
  page.drawText("Grand Total", {
    x: 365,
    y: currentY - 2,
    font: fontBold,
    size: 9,
    color: brandDark
  });
  page.drawText(formatMoney(totalValue), {
    x: 505,
    y: currentY - 2,
    font: fontBold,
    size: 9,
    color: brandDark
  });
  currentY -= 46;

  page.drawText("Terms & Conditions", {
    x: 40,
    y: currentY,
    font: fontBold,
    size: 11,
    color: brandDark
  });
  currentY -= 18;

  const terms = payload.terms.length
    ? payload.terms
    : ["Terms will be finalized during the next workflow step."];
  terms.forEach((term, index) => {
    const wrapped = wrapPdfText(`${index + 1}. ${term}`, 500, fontRegular, 9);
    currentY = drawWrappedLines(page, wrapped, {
      x: 46,
      y: currentY,
      lineHeight: 12,
      font: fontRegular,
      size: 9,
      color: brandDark
    });
    currentY -= 4;
  });

  page.drawText("Generated by Resham Sutra quotation service", {
    x: 40,
    y: 28,
    font: fontRegular,
    size: 8,
    color: muted
  });

  const pdfBytes = await pdfDoc.save();
  await writeFile(pdfPath, pdfBytes);
  return pdfPath;
}

function renderHtml(payload: DocumentPayload, title: string, subtitle: string) {
  const terms =
    payload.terms.length > 0
      ? payload.terms.map((term) => `<li>${term}</li>`).join("")
      : "<li>Terms will be finalized during the next implementation step.</li>";

  return `<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <title>${title}</title>
    <style>
      body { font-family: Calibri, Arial, sans-serif; margin: 32px; color: #1f2937; }
      h1 { margin: 0 0 6px; font-size: 30px; }
      h2 { margin: 0 0 20px; color: #6b7280; font-size: 16px; font-weight: 600; }
      .meta { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 24px; }
      .box { border: 1px solid #d1d5db; border-radius: 12px; padding: 14px; white-space: pre-wrap; }
      table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
      th, td { border: 1px solid #d1d5db; padding: 10px; text-align: left; vertical-align: top; }
      th { background: #f3f4f6; }
      .tag { display: inline-block; padding: 4px 10px; background: #eef2ff; color: #4338ca; border-radius: 999px; font-size: 12px; }
      .footer { color: #6b7280; font-size: 13px; margin-top: 18px; }
    </style>
  </head>
  <body>
    <span class="tag">${subtitle}</span>
    <h1>${payload.quotationNumber}</h1>
    <h2>${title}</h2>
    <div class="meta">
      <div class="box">${payload.buyerBlock || `${payload.customerName}\n${payload.company}`}</div>
      <div class="box">${payload.consigneeBlock || "Consignee block will be added during live implementation."}</div>
    </div>
    <table>
      <thead>
        <tr>
          <th>S.No.</th>
          <th>Description</th>
          <th>Qty</th>
          <th>Rate</th>
          <th>Pkg & Transport</th>
          <th>GST %</th>
          <th>GST Amt</th>
          <th>Total Amount</th>
        </tr>
      </thead>
      <tbody>
        ${renderRows(payload.lineItems)}
      </tbody>
    </table>
    <div class="box">
      <strong>Terms & Conditions</strong>
      <ul>${terms}</ul>
    </div>
    <p class="footer">
      Generated by the Resham Sutra document service. This preview is for workflow inspection only.
    </p>
  </body>
</html>`;
}

async function ensureDocumentDir(customerFolderName: string, quotationNumber: string) {
  const folderName = slugify(customerFolderName) || slugify(quotationNumber) || "quotation";
  const outputDir = path.resolve(env.DOCUMENT_STORAGE_DIR, folderName);
  await mkdir(outputDir, { recursive: true });
  return outputDir;
}

function toPublicUrl(filePath: string) {
  const relative = path.relative(path.resolve(env.DOCUMENT_STORAGE_DIR), filePath);
  return `${env.PUBLIC_API_BASE_URL}/documents/${relative.replaceAll("\\", "/")}`;
}

export function getStoredDocumentArtifact(
  customerFolderName: string,
  quotationNumber: string,
  extension: "xlsx" | "pdf"
) {
  const folderName = slugify(customerFolderName) || slugify(quotationNumber) || "quotation";
  const outputDir = path.resolve(env.DOCUMENT_STORAGE_DIR, folderName);
  const safeName = safeFileName(quotationNumber, "quotation");
  const filePath = path.join(outputDir, `${safeName}.${extension}`);

  return {
    filePath,
    fileUrl: toPublicUrl(filePath),
    fileName: `${safeName}.${extension}`
  };
}

export function parseDocumentPayload(input: unknown) {
  return documentPayloadSchema.parse(input);
}

export async function createDraftDocument(input: unknown) {
  const payload = parseDocumentPayload(input);
  const outputDir = await ensureDocumentDir(payload.customerFolderName, payload.quotationNumber);
  const xlsxPath = await writeXlsxDraft(payload, outputDir);
  const safeName = safeFileName(payload.quotationNumber, "quotation");
  const htmlPath = path.join(outputDir, `${safeName}-draft-preview.html`);
  const jsonPath = path.join(outputDir, `${safeName}-draft.payload.json`);

  await writeFile(
    htmlPath,
    renderHtml(payload, "Draft Quotation", `Draft ${payload.draftFormat}`),
    "utf8"
  );
  await writeFile(jsonPath, JSON.stringify(payload, null, 2), "utf8");

  return {
    kind: "draft",
    quotationRecordId: payload.quotationRecordId,
    quotationNumber: payload.quotationNumber,
    draftFormat: payload.draftFormat,
    filePath: xlsxPath,
    payloadPath: jsonPath,
    fileUrl: toPublicUrl(xlsxPath),
    previewUrl: toPublicUrl(htmlPath),
    payloadUrl: toPublicUrl(jsonPath),
    message:
      "Draft generator created a template-based XLSX workbook plus preview HTML and payload artifacts."
  };
}

export async function createPdfDocument(input: unknown) {
  const payload = parseDocumentPayload(input);
  const outputDir = await ensureDocumentDir(payload.customerFolderName, payload.quotationNumber);
  const safeName = safeFileName(payload.quotationNumber, "quotation");
  const pdfPath = await writeSimplePdf(payload, outputDir);
  const htmlPath = path.join(outputDir, `${safeName}-final-preview.html`);
  const jsonPath = path.join(outputDir, `${safeName}-final.payload.json`);

  await writeFile(
    htmlPath,
    renderHtml(payload, "Final Quotation PDF Source", "Final PDF Placeholder"),
    "utf8"
  );
  await writeFile(jsonPath, JSON.stringify(payload, null, 2), "utf8");

  return {
    kind: "pdf",
    quotationRecordId: payload.quotationRecordId,
    quotationNumber: payload.quotationNumber,
    filePath: pdfPath,
    payloadPath: jsonPath,
    fileUrl: toPublicUrl(pdfPath),
    previewUrl: toPublicUrl(htmlPath),
    payloadUrl: toPublicUrl(jsonPath),
    message: "PDF generator created a final quotation PDF plus preview HTML and payload artifacts."
  };
}
