import path from "node:path";
import nodemailer from "nodemailer";
import { env } from "./config.js";

type MailAttachment = {
  filename: string;
  path: string;
  contentType?: string;
};

type SendMailInput = {
  to: string | string[];
  subject: string;
  text: string;
  html?: string;
  attachments?: MailAttachment[];
};

let cachedTransporter: nodemailer.Transporter | null = null;

export function isSmtpConfigured() {
  return Boolean(env.SMTP_HOST && env.SMTP_PORT && env.SMTP_USER && env.SMTP_APP_PASSWORD);
}

function getTransporter() {
  if (!isSmtpConfigured()) {
    throw new Error("SMTP is not configured");
  }

  if (!cachedTransporter) {
    cachedTransporter = nodemailer.createTransport({
      host: env.SMTP_HOST,
      port: env.SMTP_PORT,
      secure: env.SMTP_PORT === 465,
      auth: {
        user: env.SMTP_USER,
        pass: env.SMTP_APP_PASSWORD
      }
    });
  }

  return cachedTransporter;
}

function normalizeRecipients(to: string | string[]) {
  return Array.isArray(to) ? to.filter(Boolean) : [to].filter(Boolean);
}

export async function sendMail(input: SendMailInput) {
  const transporter = getTransporter();
  const recipients = normalizeRecipients(input.to);

  if (!recipients.length) {
    throw new Error("No email recipients were provided");
  }

  return transporter.sendMail({
    from: {
      name: env.SMTP_FROM_NAME,
      address: env.SMTP_USER
    },
    to: recipients.join(", "),
    subject: input.subject,
    text: input.text,
    html: input.html,
    attachments: input.attachments
  });
}

export function resolveDocumentAttachment(relativePath: string, filename: string, contentType?: string) {
  return {
    filename,
    path: path.resolve(env.DOCUMENT_STORAGE_DIR, relativePath),
    contentType
  };
}
