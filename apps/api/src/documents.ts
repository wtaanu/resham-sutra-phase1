import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import ExcelJS from "exceljs";
import {
  PDFDocument,
  StandardFonts,
  rgb,
  type PDFFont,
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

const COMPANY_LINES = [
  "Resham Sutra Pvt. Ltd. (Central Silk Board approved)",
  "738 Ghewra Mod",
  "Mundka Udhyog Nagar - North",
  "New Delhi - 110041",
  "Tel: 9811050909",
  "Email: info@reshamsutra.com"
];

const COMPANY_META_LINES = [
  "GSTIN No.: 07AAHCR4176J1Z0",
  "CIN No.: U17116JH2015PTC002989"
];

const BANK_DETAIL_LINES = [
  "Our Bank Details:",
  "Bank Name: Canara Bank",
  "Branch: Connaught Place, New Delhi",
  "A/C Name: RESHAM SUTRA PVT. LTD.",
  "A/c No.: 90421400001164",
  "IFSC Code: CNRB0002009"
];

const CONSIGNEE_COMPANY_LINES = [...COMPANY_LINES, ...COMPANY_META_LINES];

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
    payload.consigneeBlock || `${payload.customerName}\n${payload.company}`,
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
  return [
    path.resolve(currentDir, "../../web/src/assets/resham-sutra-logo.png"),
    path.resolve(currentDir, "../../templates/quotation-template-unpacked/xl/media/image3.png")
  ];
}

function resolveSignaturePath() {
  const currentDir = path.dirname(fileURLToPath(import.meta.url));
  return path.resolve(currentDir, "../../templates/quotation-template-unpacked/xl/media/image6.jpeg");
}

function resolveFooterBadgePaths() {
  const currentDir = path.dirname(fileURLToPath(import.meta.url));
  return [
    { key: "makeInIndia", path: path.resolve(currentDir, "../../templates/quotation-template-unpacked/xl/media/image2.jpeg") },
    { key: "iso", path: path.resolve(currentDir, "../../templates/quotation-template-unpacked/xl/media/image4.jpeg") },
    { key: "msme", path: path.resolve(currentDir, "../../templates/quotation-template-unpacked/xl/media/image5.png") },
    { key: "startupIndia", path: path.resolve(currentDir, "../../templates/quotation-template-unpacked/xl/media/image3.png") }
  ];
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

function rightAlignTextX(text: string, rightEdge: number, font: PDFFont, size: number) {
  return rightEdge - font.widthOfTextAtSize(text, size);
}

function normalizeAddressBlock(primary: string, fallback: string) {
  const value = String(primary || fallback || "")
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .join("\n");

  return value || "-";
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
  for (const logoPath of resolveLogoPath()) {
    try {
      const logoBytes = await readFile(logoPath);
      if (logoPath.toLowerCase().endsWith(".jpg") || logoPath.toLowerCase().endsWith(".jpeg")) {
        return pdfDoc.embedJpg(logoBytes);
      }

      return pdfDoc.embedPng(logoBytes);
    } catch {
      continue;
    }
  }

  return null;
}

async function loadSignature(pdfDoc: PDFDocument) {
  try {
    const signatureBytes = await readFile(resolveSignaturePath());
    return pdfDoc.embedJpg(signatureBytes);
  } catch {
    return null;
  }
}

async function loadFooterBadges(pdfDoc: PDFDocument) {
  const badges: Array<{ key: string; image: Awaited<ReturnType<typeof pdfDoc.embedPng>> | Awaited<ReturnType<typeof pdfDoc.embedJpg>> }> = [];

  for (const badge of resolveFooterBadgePaths()) {
    try {
      const bytes = await readFile(badge.path);
      const image =
        badge.path.toLowerCase().endsWith(".jpg") || badge.path.toLowerCase().endsWith(".jpeg")
          ? await pdfDoc.embedJpg(bytes)
          : await pdfDoc.embedPng(bytes);

      badges.push({ key: badge.key, image });
    } catch {
      continue;
    }
  }

  return badges;
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
  cells: Array<{ text: string; x: number; width: number; align?: "left" | "right"; lines?: string[] }>,
  y: number,
  font: PDFFont,
  rowHeight: number
) {
  page.drawRectangle({
    x: 40,
    y: y - rowHeight,
    width: 515,
    height: rowHeight,
    borderWidth: 0.5,
    borderColor: rgb(0.82, 0.85, 0.9),
    color: rgb(1, 1, 1)
  });

  cells.forEach((cell) => {
    const lines = cell.lines?.length ? cell.lines : [cell.text || "-"];
    let currentY = y - 14;

    lines.forEach((line) => {
      const textWidth = font.widthOfTextAtSize(line, 8);
      const x =
        cell.align === "right"
          ? cell.x + cell.width - textWidth - 4
          : cell.x + 4;

      page.drawText(line, {
        x,
        y: currentY,
        font,
        size: 8,
        color: rgb(0.16, 0.18, 0.21)
      });
      currentY -= 10;
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
  const signature = await loadSignature(pdfDoc);
  const footerBadges = await loadFooterBadges(pdfDoc);
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

  if (logo) {
    const dimensions = logo.scale(0.08);
    page.drawImage(logo, {
      x: 40,
      y: 768,
      width: dimensions.width,
      height: dimensions.height
    });
  }

  const documentTitle = payload.templateCode === "myanmar-proforma" ? "PERFORMA INVOICE" : "Quotation";
  const titleSize = 9.5;
  const quotationNumberSize = 8.5;
  const headerRightEdge = 553;
  page.drawText(documentTitle, {
    x: rightAlignTextX(documentTitle, headerRightEdge, fontBold, titleSize),
    y: 808,
    font: fontBold,
    size: titleSize,
    color: brandDark
  });
  page.drawText(payload.quotationNumber, {
    x: rightAlignTextX(payload.quotationNumber, headerRightEdge, fontBold, quotationNumberSize),
    y: 789,
    font: fontBold,
    size: quotationNumberSize,
    color: brandDark
  });
  const dateText = `Date: ${formatDateForTemplate()}`;
  page.drawText(dateText, {
    x: rightAlignTextX(dateText, headerRightEdge, fontRegular, 8.5),
    y: 772,
    font: fontRegular,
    size: 8.5,
    color: muted
  });

  let companyY = 800;
  COMPANY_LINES.forEach((line, index) => {
    page.drawText(line, {
      x: 160,
      y: companyY,
      font: index === 0 ? fontBold : fontRegular,
      size: index === 0 ? 7.4 : 7.1,
      color: brandDark
    });
    companyY -= 8;
  });

  COMPANY_META_LINES.forEach((line) => {
    page.drawText(line, {
      x: 160,
      y: companyY,
      font: fontRegular,
      size: 7.1,
      color: brandDark
    });
    companyY -= 8;
  });

  page.drawRectangle({
    x: 40,
    y: 640,
    width: 250,
    height: 108,
    borderWidth: 1,
    borderColor: rgb(0.88, 0.89, 0.91),
    color: rgb(0.99, 0.99, 0.99)
  });
  page.drawRectangle({
    x: 305,
    y: 640,
    width: 250,
    height: 108,
    borderWidth: 1,
    borderColor: rgb(0.88, 0.89, 0.91),
    color: rgb(0.99, 0.99, 0.99)
  });

  page.drawText("Buyer", {
    x: 52,
    y: 732,
    font: fontBold,
    size: 10,
    color: brandDark
  });
  page.drawText("Consignee", {
    x: 317,
    y: 732,
    font: fontBold,
    size: 10,
    color: brandDark
  });

  const buyerLines = splitLines(
    normalizeAddressBlock(payload.buyerBlock, `${payload.customerName}\n${payload.company}`),
    7
  );
  const consigneeLines = splitLines(
    normalizeAddressBlock(CONSIGNEE_COMPANY_LINES.join("\n"), "Consignee details not provided"),
    8
  );

  drawWrappedLines(page, buyerLines, {
    x: 52,
    y: 715,
    lineHeight: 11,
    font: fontRegular,
    size: 9,
    color: brandDark
  });
  drawWrappedLines(page, consigneeLines, {
    x: 317,
    y: 715,
    lineHeight: 10,
    font: fontRegular,
    size: 8,
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
  let currentY = 606;
  drawTableHeader(page, headers, currentY, fontBold);
  currentY -= 26;

  payload.lineItems.slice(0, 12).forEach((item, index) => {
    const descriptionLines = wrapPdfText(
      item.description || "Quotation item",
      196,
      fontRegular,
      8
    );
    const rowHeight = Math.max(24, 12 + descriptionLines.length * 10);

    drawTableRow(
      page,
      [
        { text: String(item.lineNo ?? index + 1), x: 40, width: 35 },
        { text: descriptionLines[0] || "Quotation item", x: 75, width: 205, lines: descriptionLines },
        { text: String(item.qty || 0), x: 280, width: 35, align: "right" },
        { text: formatMoney(item.rate || 0), x: 315, width: 60, align: "right" },
        { text: formatMoney(item.transport || 0), x: 375, width: 60, align: "right" },
        { text: formatMoney(item.gstPercent || 0), x: 435, width: 40, align: "right" },
        { text: formatMoney(item.gstAmount || 0), x: 475, width: 40, align: "right" },
        { text: formatMoney(item.totalAmount || 0), x: 515, width: 40, align: "right" }
      ],
      currentY,
      fontRegular,
      rowHeight
    );
    currentY -= rowHeight;
  });

  const totalValue = payload.lineItems.reduce((sum, item) => sum + Number(item.totalAmount || 0), 0);

  page.drawRectangle({
    x: 375,
    y: currentY - 28,
    width: 180,
    height: 26,
    color: brandYellow
  });
  page.drawText("Grand Total", {
    x: 387,
    y: currentY - 18,
    font: fontBold,
    size: 9,
    color: brandDark
  });
  const totalText = formatMoney(totalValue);
  page.drawText(formatMoney(totalValue), {
    x: 545 - fontBold.widthOfTextAtSize(totalText, 10),
    y: currentY - 18,
    font: fontBold,
    size: 10,
    color: brandDark
  });
  currentY -= 54;

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

  let bankY = Math.min(currentY - 8, 92);
  BANK_DETAIL_LINES.forEach((line, index) => {
    page.drawText(line, {
      x: 40,
      y: bankY,
      font: index === 0 ? fontBold : fontRegular,
      size: index === 0 ? 7.8 : 7.2,
      color: brandDark
    });
    bankY -= 8;
  });

  const footerBadgeSpecs: Record<string, { width: number; height: number }> = {
    makeInIndia: { width: 56, height: 30 },
    iso: { width: 20, height: 20 },
    msme: { width: 56, height: 22 },
    startupIndia: { width: 78, height: 18 }
  };

  let badgeX = 40;
  const badgeY = 18;
  footerBadges.forEach(({ key, image }) => {
    const spec = footerBadgeSpecs[key] || { width: 36, height: 18 };
    page.drawImage(image, {
      x: badgeX,
      y: badgeY,
      width: spec.width,
      height: spec.height
    });
    badgeX += spec.width + 8;
  });

  page.drawText("For Resham Sutra Pvt. Ltd.", {
    x: 404,
    y: 88,
    font: fontBold,
    size: 8.2,
    color: brandDark
  });
  if (signature) {
    const dimensions = signature.scale(0.34);
    page.drawImage(signature, {
      x: 404,
      y: 28,
      width: dimensions.width,
      height: dimensions.height
    });
  }
  page.drawText("Authorized Signatory", {
    x: 418,
    y: 18,
    font: fontRegular,
    size: 7.4,
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
