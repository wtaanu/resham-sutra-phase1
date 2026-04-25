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

export function getSmtpConfigState() {
  return {
    host: Boolean(String(env.SMTP_HOST || "").trim()),
    port: Boolean(env.SMTP_PORT),
    user: Boolean(String(env.SMTP_USER || "").trim()),
    password: Boolean(String(env.SMTP_APP_PASSWORD || "").trim())
  };
}

export function isSmtpConfigured() {
  const state = getSmtpConfigState();
  return state.host && state.port && state.user && state.password;
}

function getTransporter() {
  if (!isSmtpConfigured()) {
    const state = getSmtpConfigState();
    throw new Error(
      `SMTP is not configured (host=${state.host}, port=${state.port}, user=${state.user}, password=${state.password})`
    );
  }

  if (!cachedTransporter) {
    cachedTransporter = nodemailer.createTransport({
      host: env.SMTP_HOST,
      port: env.SMTP_PORT,
      secure: env.SMTP_PORT === 465,
      connectionTimeout: 10000,
      greetingTimeout: 10000,
      socketTimeout: 15000,
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

  try {
    return await transporter.sendMail({
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
  } catch (error) {
    const message = error instanceof Error ? error.message : "Unknown SMTP error";
    throw new Error(`SMTP send failed: ${message}`);
  }
}

export function resolveDocumentAttachment(relativePath: string, filename: string, contentType?: string) {
  return {
    filename,
    path: path.resolve(env.DOCUMENT_STORAGE_DIR, relativePath),
    contentType
  };
}
