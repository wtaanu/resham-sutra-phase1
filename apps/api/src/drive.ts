import { readFile } from "node:fs/promises";
import path from "node:path";
import { env } from "./config.js";

type GoogleDriveFile = {
  id: string;
  name: string;
  webViewLink?: string;
};

type GoogleDriveListResponse = {
  files: GoogleDriveFile[];
};

type GoogleOAuthTokenResponse = {
  access_token: string;
  expires_in: number;
  token_type: string;
};

let cachedAccessToken = "";
let cachedAccessTokenExpiresAt = 0;
let inflightTokenPromise: Promise<string> | null = null;

function hasRefreshTokenConfig() {
  return Boolean(
    env.GOOGLE_DRIVE_CLIENTID &&
      env.GOOGLE_DRIVE_CLIENTSECRET &&
      env.GOOGLE_DRIVE_REFRESH_TOKEN
  );
}

async function exchangeRefreshTokenForAccessToken() {
  const response = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded"
    },
    body: new URLSearchParams({
      client_id: env.GOOGLE_DRIVE_CLIENTID,
      client_secret: env.GOOGLE_DRIVE_CLIENTSECRET,
      refresh_token: env.GOOGLE_DRIVE_REFRESH_TOKEN,
      grant_type: "refresh_token"
    })
  });

  if (!response.ok) {
    const message = await response.text();
    throw new Error(`Drive token refresh failed (${response.status}): ${message}`);
  }

  const data = (await response.json()) as GoogleOAuthTokenResponse;
  cachedAccessToken = data.access_token;
  cachedAccessTokenExpiresAt = Date.now() + Math.max(data.expires_in - 60, 60) * 1000;
  return cachedAccessToken;
}

async function getDriveAccessToken() {
  if (env.GOOGLE_DRIVE_ACCESS_TOKEN) {
    return env.GOOGLE_DRIVE_ACCESS_TOKEN;
  }

  if (!hasRefreshTokenConfig()) {
    throw new Error("Google Drive credentials are not configured");
  }

  if (cachedAccessToken && Date.now() < cachedAccessTokenExpiresAt) {
    return cachedAccessToken;
  }

  if (!inflightTokenPromise) {
    inflightTokenPromise = exchangeRefreshTokenForAccessToken().finally(() => {
      inflightTokenPromise = null;
    });
  }

  return inflightTokenPromise;
}

async function driveHeaders() {
  return {
    Authorization: `Bearer ${await getDriveAccessToken()}`,
    "Content-Type": "application/json"
  };
}

export function isDriveConfigured() {
  return Boolean(
    env.GOOGLE_DRIVE_ROOT_FOLDER_ID &&
      (env.GOOGLE_DRIVE_ACCESS_TOKEN || hasRefreshTokenConfig())
  );
}

function escapeDriveQuery(value: string) {
  return value.replace(/'/g, "\\'");
}

export function buildDriveFolderUrl(folderId: string) {
  return `https://drive.google.com/drive/folders/${folderId}`;
}

export function extractDriveFolderId(folderUrl: string) {
  const match = String(folderUrl).match(/folders\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : "";
}

export async function findFolderByName(folderName: string) {
  if (!isDriveConfigured()) {
    throw new Error("Google Drive is not configured");
  }

  const query = [
    `name='${escapeDriveQuery(folderName)}'`,
    `mimeType='application/vnd.google-apps.folder'`,
    "trashed=false",
    `'${env.GOOGLE_DRIVE_ROOT_FOLDER_ID}' in parents`
  ].join(" and ");

  const params = new URLSearchParams({
    q: query,
    fields: "files(id,name,webViewLink)"
  });

  const response = await fetch(
    `https://www.googleapis.com/drive/v3/files?${params.toString()}`,
    {
      method: "GET",
      headers: await driveHeaders()
    }
  );

  if (!response.ok) {
    const message = await response.text();
    throw new Error(`Drive search failed (${response.status}): ${message}`);
  }

  const data = (await response.json()) as GoogleDriveListResponse;
  return data.files[0] ?? null;
}

async function findFileByNameInFolder(fileName: string, folderId: string) {
  if (!isDriveConfigured()) {
    throw new Error("Google Drive is not configured");
  }

  const query = [
    `name='${escapeDriveQuery(fileName)}'`,
    "trashed=false",
    `'${folderId}' in parents`
  ].join(" and ");

  const params = new URLSearchParams({
    q: query,
    fields: "files(id,name,webViewLink)"
  });

  const response = await fetch(
    `https://www.googleapis.com/drive/v3/files?${params.toString()}`,
    {
      method: "GET",
      headers: await driveHeaders()
    }
  );

  if (!response.ok) {
    const message = await response.text();
    throw new Error(`Drive file search failed (${response.status}): ${message}`);
  }

  const data = (await response.json()) as GoogleDriveListResponse;
  return data.files[0] ?? null;
}

export async function createFolder(folderName: string) {
  if (!isDriveConfigured()) {
    throw new Error("Google Drive is not configured");
  }

  const response = await fetch(
    "https://www.googleapis.com/drive/v3/files?fields=id,name,webViewLink",
    {
      method: "POST",
      headers: await driveHeaders(),
      body: JSON.stringify({
        name: folderName,
        mimeType: "application/vnd.google-apps.folder",
        parents: [env.GOOGLE_DRIVE_ROOT_FOLDER_ID]
      })
    }
  );

  if (!response.ok) {
    const message = await response.text();
    throw new Error(`Drive folder create failed (${response.status}): ${message}`);
  }

  return (await response.json()) as GoogleDriveFile;
}

export async function findOrCreateClientFolder(folderName: string) {
  const existing = await findFolderByName(folderName);

  if (existing) {
    return {
      folderId: existing.id,
      folderName: existing.name,
      folderUrl: existing.webViewLink || buildDriveFolderUrl(existing.id),
      created: false
    };
  }

  const created = await createFolder(folderName);
  return {
    folderId: created.id,
    folderName: created.name,
    folderUrl: created.webViewLink || buildDriveFolderUrl(created.id),
    created: true
  };
}

export async function uploadFileToFolder(
  filePath: string,
  fileName: string,
  folderId: string
) {
  if (!isDriveConfigured()) {
    throw new Error("Google Drive is not configured");
  }

  const fileBuffer = await readFile(filePath);
  const form = new FormData();

  form.append(
    "metadata",
    new Blob(
      [
        JSON.stringify({
          name: fileName
        })
      ],
      { type: "application/json" }
    )
  );
  form.append(
    "file",
    new Blob([fileBuffer], {
      type:
        path.extname(fileName).toLowerCase() === ".xlsx"
          ? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          : path.extname(fileName).toLowerCase() === ".pdf"
            ? "application/pdf"
            : "application/octet-stream"
    }),
    fileName
  );

  const existing = await findFileByNameInFolder(fileName, folderId);
  const uploadUrl = existing
    ? `https://www.googleapis.com/upload/drive/v3/files/${existing.id}?uploadType=multipart&fields=id,name,webViewLink`
    : "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name,webViewLink";
  const method = existing ? "PATCH" : "POST";

  if (!existing) {
    form.set(
      "metadata",
      new Blob(
        [
          JSON.stringify({
            name: fileName,
            parents: [folderId]
          })
        ],
        { type: "application/json" }
      )
    );
  }

  const response = await fetch(uploadUrl, {
    method,
    headers: {
      Authorization: `Bearer ${await getDriveAccessToken()}`
    },
    body: form
  });

  if (!response.ok) {
    const message = await response.text();
    throw new Error(`Drive upload failed (${response.status}): ${message}`);
  }

  const data = (await response.json()) as GoogleDriveFile;
  return {
    fileId: data.id,
    fileName: data.name,
    fileUrl: data.webViewLink || `https://drive.google.com/file/d/${data.id}/view`
  };
}
