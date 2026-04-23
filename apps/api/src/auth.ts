import { createHmac, timingSafeEqual } from "node:crypto";
import type { NextFunction, Request, Response } from "express";
import { listRecords, type AirtableRecord } from "./airtable.js";
import { env } from "./config.js";

type UserFields = {
  Name?: string;
  Email?: string;
  Password?: string;
};

type SessionPayload = {
  v: 1;
  sub: string;
  name: string;
  email: string;
  exp: number;
};

export type AuthenticatedUser = {
  id: string;
  name: string;
  email: string;
};

const SESSION_MAX_AGE_MS = 30 * 24 * 60 * 60 * 1000;

function normalizeEmail(value: unknown) {
  return String(value ?? "").trim().toLowerCase();
}

function base64UrlEncode(value: string) {
  return Buffer.from(value, "utf8").toString("base64url");
}

function base64UrlDecode(value: string) {
  return Buffer.from(value, "base64url").toString("utf8");
}

function signPayload(payload: string) {
  return createHmac("sha256", env.SESSION_SECRET).update(payload).digest("base64url");
}

function safelyCompare(left: string, right: string) {
  const leftBuffer = Buffer.from(left, "utf8");
  const rightBuffer = Buffer.from(right, "utf8");

  if (leftBuffer.length !== rightBuffer.length) {
    return false;
  }

  return timingSafeEqual(leftBuffer, rightBuffer);
}

function parseCookies(cookieHeader: string | undefined) {
  return String(cookieHeader || "")
    .split(";")
    .map((part) => part.trim())
    .filter(Boolean)
    .reduce<Record<string, string>>((accumulator, part) => {
      const [name, ...value] = part.split("=");
      if (!name) {
        return accumulator;
      }

      accumulator[name] = decodeURIComponent(value.join("="));
      return accumulator;
    }, {});
}

function sessionCookieValue(user: AuthenticatedUser) {
  const payload: SessionPayload = {
    v: 1,
    sub: user.id,
    name: user.name,
    email: user.email,
    exp: Date.now() + SESSION_MAX_AGE_MS
  };

  const encodedPayload = base64UrlEncode(JSON.stringify(payload));
  const signature = signPayload(encodedPayload);
  return `${encodedPayload}.${signature}`;
}

function verifySessionCookie(value: string | undefined) {
  if (!value || !value.includes(".")) {
    return null;
  }

  const [encodedPayload, signature] = value.split(".");
  if (!encodedPayload || !signature) {
    return null;
  }

  const expectedSignature = signPayload(encodedPayload);
  if (!safelyCompare(signature, expectedSignature)) {
    return null;
  }

  try {
    const payload = JSON.parse(base64UrlDecode(encodedPayload)) as SessionPayload;
    if (payload.v !== 1 || payload.exp <= Date.now()) {
      return null;
    }

    return {
      id: payload.sub,
      name: payload.name,
      email: payload.email
    } satisfies AuthenticatedUser;
  } catch {
    return null;
  }
}

function sessionCookieOptions(maxAgeMs: number) {
  return {
    httpOnly: true,
    secure: env.PUBLIC_API_BASE_URL.startsWith("https://"),
    sameSite: "lax" as const,
    path: "/",
    maxAge: maxAgeMs
  };
}

export async function authenticatePortalUser(email: string, password: string) {
  const normalizedEmail = normalizeEmail(email);
  const normalizedPassword = String(password ?? "");

  if (!normalizedEmail || !normalizedPassword) {
    return null;
  }

  const users = await listRecords<UserFields>(env.AIRTABLE_USERS_TABLE, {
    fields: ["Name", "Email", "Password"],
    maxRecords: 200
  });

  const matchingUser = users.find((user) => normalizeEmail(user.fields.Email) === normalizedEmail);
  if (!matchingUser) {
    return null;
  }

  const storedPassword = String(matchingUser.fields.Password || "");
  if (!storedPassword || !safelyCompare(storedPassword, normalizedPassword)) {
    return null;
  }

  return {
    id: matchingUser.id,
    name: String(matchingUser.fields.Name || matchingUser.fields.Email || "ReshamSutra User").trim(),
    email: normalizeEmail(matchingUser.fields.Email)
  } satisfies AuthenticatedUser;
}

export function setAuthenticatedSession(response: Response, user: AuthenticatedUser) {
  response.cookie(
    env.SESSION_COOKIE_NAME,
    sessionCookieValue(user),
    sessionCookieOptions(SESSION_MAX_AGE_MS)
  );
}

export function clearAuthenticatedSession(response: Response) {
  response.clearCookie(env.SESSION_COOKIE_NAME, {
    httpOnly: true,
    secure: env.PUBLIC_API_BASE_URL.startsWith("https://"),
    sameSite: "lax",
    path: "/"
  });
}

export function getAuthenticatedUser(request: Request) {
  const cookies = parseCookies(request.headers.cookie);
  return verifySessionCookie(cookies[env.SESSION_COOKIE_NAME]);
}

export function requireAuthenticatedUser(
  request: Request,
  response: Response,
  next: NextFunction
) {
  const user = getAuthenticatedUser(request);
  if (!user) {
    response.status(401).json({
      status: "error",
      message: "Authentication required."
    });
    return;
  }

  response.locals.authUser = user;
  next();
}

