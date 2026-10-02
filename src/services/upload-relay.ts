/**
 * Upload relay for files that live neither at a public URL nor on the
 * server's host — typically a conversation attachment, which sits in the MCP
 * client's code-execution sandbox. The model cannot re-emit those bytes
 * reliably (see `fileContent`), so they must travel outside its output:
 *
 *   1. `boond_documents_upload_slot` asks the relay for a single-use,
 *      pre-authenticated upload URL on an object store the operator controls;
 *   2. the client's sandbox PUTs the file there (`curl`), bytes never passing
 *      through the model;
 *   3. `boond_documents_create({ uploadSlot })` checks the stored file (size,
 *      magic bytes), hands BoondManager a short-lived pre-authenticated
 *      download URL through the existing `fileUrl` mechanism, then purges the
 *      file — permanently, bypassing the recycle bin when the store allows.
 *
 * Store credentials stay on the server; the client only ever sees URLs scoped
 * to one file and one slot. Off unless `BOOND_MCP_UPLOAD_RELAY` is set.
 *
 * Backends implement {@link RelayBackend}; `sharepoint` (Microsoft Graph,
 * app-only `Sites.Selected`) is the first.
 */
import { randomUUID } from "node:crypto";
import { extname } from "node:path";
import { readPositiveInt, readString } from "../config/env.js";
import { DEFAULT_UPLOAD_MAX_BYTES } from "../constants.js";
import { sniffDocumentType, UploadRejectedError } from "./upload-source.js";

/** How long a slot stays usable after creation. */
export const SLOT_TTL_MS = 15 * 60 * 1000;
const GRAPH = "https://graph.microsoft.com/v1.0";
const REQUEST_TIMEOUT_MS = 30_000;

export interface SlotHandle {
  /** Backend-private reference to the stored object (e.g. Graph folder id). */
  ref: string;
  /** Pre-authenticated URL the client PUTs the file to (single request). */
  uploadUrl: string;
  expiresAt: Date;
}

export interface StoredObject {
  size: number;
  /** Short-lived, pre-authenticated URL BoondManager can download. */
  downloadUrl: string;
}

export interface RelayBackend {
  readonly name: string;
  createSlot(slotId: string, filename: string): Promise<SlotHandle>;
  /** `undefined` when nothing was uploaded to the slot yet. */
  stat(ref: string, filename: string): Promise<StoredObject | undefined>;
  /** First bytes of the stored object, for type sniffing. */
  head(downloadUrl: string, bytes: number): Promise<Buffer>;
  purge(ref: string): Promise<void>;
}

// ---------------------------------------------------------------------------
// SharePoint / Microsoft Graph backend
// ---------------------------------------------------------------------------

export interface SharePointConfig {
  tenantId: string;
  clientId: string;
  clientSecret: string;
  /** `host:/sites/path` (Graph site key). */
  site: string;
  library: string;
}

/** Accept `https://host/sites/x`, `host/sites/x` or `host:/sites/x`. */
export function normalizeSiteKey(raw: string): string {
  const s = raw
    .trim()
    .replace(/^https?:\/\//i, "")
    .replace(/\/+$/, "");
  if (s.includes(":/")) return s;
  const slash = s.indexOf("/");
  return slash === -1 ? s : `${s.slice(0, slash)}:${s.slice(slash)}`;
}

/** SharePoint forbids `" * : < > ? / \ |` and `#`/`%` cause URL trouble. */
export function sanitizeSharePointName(filename: string): string {
  const cleaned = filename
    .replace(/["*:<>?/\\|#%]/g, "_")
    .split("")
    .filter((c) => c.charCodeAt(0) >= 0x20)
    .join("")
    .trim()
    .replace(/^\.+|\.+$/g, "");
  return (cleaned.length > 0 ? cleaned : "document").slice(0, 200);
}

export class SharePointRelay implements RelayBackend {
  readonly name = "sharepoint";
  private token?: { value: string; expiresAt: number };
  private driveId?: string;

  constructor(private readonly cfg: SharePointConfig) {}

  private async accessToken(): Promise<string> {
    if (this.token && this.token.expiresAt > Date.now() + 60_000) return this.token.value;
    const body = new URLSearchParams({
      grant_type: "client_credentials",
      client_id: this.cfg.clientId,
      client_secret: this.cfg.clientSecret,
      scope: "https://graph.microsoft.com/.default",
    });
    const res = await fetch(
      `https://login.microsoftonline.com/${encodeURIComponent(this.cfg.tenantId)}/oauth2/v2.0/token`,
      { method: "POST", body, signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS) }
    );
    const json = (await res.json().catch(() => ({}))) as {
      access_token?: string;
      expires_in?: number;
      error?: string;
    };
    if (!res.ok || !json.access_token) {
      // Never echo the request (it carries the secret); the error code is enough.
      throw new Error(`Relais SharePoint : authentification Entra ID refusée (${json.error ?? res.status}).`);
    }
    this.token = { value: json.access_token, expiresAt: Date.now() + (json.expires_in ?? 3600) * 1000 };
    return this.token.value;
  }

  private async graph<T>(method: string, path: string, body?: unknown): Promise<{ status: number; json: T }> {
    const res = await fetch(`${GRAPH}${path}`, {
      method,
      headers: {
        Authorization: `Bearer ${await this.accessToken()}`,
        ...(body === undefined ? {} : { "Content-Type": "application/json" }),
      },
      body: body === undefined ? null : JSON.stringify(body),
      signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS),
    });
    const text = await res.text();
    let json = {} as T;
    try {
      if (text) json = JSON.parse(text) as T;
    } catch {
      // Non-JSON error page (proxy, throttling HTML): the status says enough.
    }
    return { status: res.status, json };
  }

  private async graphOk<T>(method: string, path: string, body?: unknown): Promise<T> {
    const { status, json } = await this.graph<T & { error?: { code?: string; message?: string } }>(method, path, body);
    if (status >= 400) {
      const e = (json as { error?: { code?: string; message?: string } }).error;
      throw new Error(
        `Relais SharePoint : ${method} ${path.split("?")[0]} → ${status} ${e?.code ?? ""} ${e?.message ?? ""}`.trim()
      );
    }
    return json;
  }

  private async drive(): Promise<string> {
    if (this.driveId) return this.driveId;
    const site = await this.graphOk<{ id: string }>("GET", `/sites/${this.cfg.site}?$select=id`);
    const drives = await this.graphOk<{ value: { id: string; name: string; webUrl?: string }[] }>(
      "GET",
      `/sites/${site.id}/drives?$select=id,name,webUrl`
    );
    const wanted = this.cfg.library.toLowerCase();
    const match = drives.value.find(
      (d) => d.name.toLowerCase() === wanted || (d.webUrl ?? "").toLowerCase().endsWith(`/${wanted}`)
    );
    if (!match) {
      const names = drives.value.map((d) => d.name).join(", ");
      throw new Error(
        `Relais SharePoint : bibliothèque « ${this.cfg.library} » introuvable sur le site (trouvées : ${names}).`
      );
    }
    this.driveId = match.id;
    return match.id;
  }

  async createSlot(slotId: string, filename: string): Promise<SlotHandle> {
    const driveId = await this.drive();
    // One folder per slot keeps the file's own name intact — BoondManager
    // names the document after the download's Content-Disposition.
    const folder = await this.graphOk<{ id: string }>("POST", `/drives/${driveId}/root/children`, {
      name: `mcp-${slotId}`,
      folder: {},
      "@microsoft.graph.conflictBehavior": "fail",
    });
    const name = encodeURIComponent(sanitizeSharePointName(filename));
    const session = await this.graphOk<{ uploadUrl: string; expirationDateTime?: string }>(
      "POST",
      `/drives/${driveId}/items/${folder.id}:/${name}:/createUploadSession`,
      { item: { "@microsoft.graph.conflictBehavior": "fail" } }
    );
    const graphExpiry = session.expirationDateTime ? Date.parse(session.expirationDateTime) : Infinity;
    return {
      ref: folder.id,
      uploadUrl: session.uploadUrl,
      expiresAt: new Date(Math.min(Date.now() + SLOT_TTL_MS, graphExpiry)),
    };
  }

  async stat(ref: string, filename: string): Promise<StoredObject | undefined> {
    const driveId = await this.drive();
    // List the slot's own folder instead of addressing the file by name: the
    // folder is private to the slot, so whatever file it holds is the upload,
    // and a name normalised differently by SharePoint cannot cause a miss.
    // No $select: SharePoint does not reliably return the download-URL
    // annotation under $select with app-only tokens.
    const { status, json } = await this.graph<{
      value?: Array<{
        id: string;
        name?: string;
        size?: number;
        file?: unknown;
        "@microsoft.graph.downloadUrl"?: string;
      }>;
    }>("GET", `/drives/${driveId}/items/${ref}/children`);
    if (status === 404) {
      throw new UploadRejectedError(
        "Dossier de transit du slot introuvable (supprimé à la main ?). Créer un nouveau slot avec `boond_documents_upload_slot`."
      );
    }
    if (status >= 400) throw new Error(`Relais SharePoint : lecture du dossier de transit → ${status}.`);
    const files = (json.value ?? []).filter((item) => item.file !== undefined);
    if (files.length === 0) return undefined;
    const wanted = sanitizeSharePointName(filename).toLowerCase();
    const item = files.find((f) => (f.name ?? "").toLowerCase() === wanted) ?? files[0]!;
    if (item.size === undefined) return undefined;
    const downloadUrl = item["@microsoft.graph.downloadUrl"] ?? (await this.contentRedirect(driveId, item.id));
    if (!downloadUrl) {
      throw new Error(
        `Relais SharePoint : fichier déposé trouvé (${item.size} octets) mais Graph ne fournit pas d'URL de lecture temporaire.`
      );
    }
    return { size: item.size, downloadUrl };
  }

  /** `/content` answers 302 to the same pre-authenticated URL; read it without following. */
  private async contentRedirect(driveId: string, itemId: string): Promise<string | undefined> {
    const res = await fetch(`${GRAPH}/drives/${driveId}/items/${itemId}/content`, {
      headers: { Authorization: `Bearer ${await this.accessToken()}` },
      redirect: "manual",
      signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS),
    });
    await res.body?.cancel().catch(() => undefined);
    return res.status >= 300 && res.status < 400 ? (res.headers.get("location") ?? undefined) : undefined;
  }

  async head(downloadUrl: string, bytes: number): Promise<Buffer> {
    const res = await fetch(downloadUrl, {
      headers: { Range: `bytes=0-${bytes - 1}` },
      signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS),
    });
    if (!res.ok) throw new Error(`Relais SharePoint : lecture de l'en-tête du fichier → ${res.status}.`);
    return Buffer.from(await res.arrayBuffer()).subarray(0, bytes);
  }

  async purge(ref: string): Promise<void> {
    const driveId = await this.drive();
    // permanentDelete skips the recycle bin (RGPD: CVs, payslips...). Fall
    // back to a regular delete if the tenant refuses it.
    const { status } = await this.graph("POST", `/drives/${driveId}/items/${ref}/permanentDelete`);
    if (status < 400 || status === 404) return;
    const del = await this.graph("DELETE", `/drives/${driveId}/items/${ref}`);
    if (del.status >= 400 && del.status !== 404) {
      throw new Error(`Relais SharePoint : suppression du dossier de transit → ${del.status}.`);
    }
  }
}

// ---------------------------------------------------------------------------
// Configuration and slot registry
// ---------------------------------------------------------------------------

/** The configured backend, or a French explanation of why there is none. */
export function relayFromEnv(env: NodeJS.ProcessEnv = process.env): RelayBackend | string {
  const kind = (readString("BOOND_MCP_UPLOAD_RELAY", env) ?? "").toLowerCase();
  if (kind === "") {
    return "Relais d'upload désactivé : l'opérateur du serveur doit définir `BOOND_MCP_UPLOAD_RELAY` (valeur : `sharepoint`).";
  }
  if (kind !== "sharepoint") {
    return `Relais d'upload inconnu : « ${kind} » (valeur acceptée : \`sharepoint\`).`;
  }
  const required = {
    tenantId: "BOOND_MCP_SHAREPOINT_TENANT_ID",
    clientId: "BOOND_MCP_SHAREPOINT_CLIENT_ID",
    clientSecret: "BOOND_MCP_SHAREPOINT_CLIENT_SECRET",
    site: "BOOND_MCP_SHAREPOINT_SITE",
  } as const;
  const missing = Object.values(required).filter((name) => !readString(name, env));
  if (missing.length > 0) {
    return `Relais SharePoint incomplet : variable(s) manquante(s) ${missing.map((m) => `\`${m}\``).join(", ")}.`;
  }
  return new SharePointRelay({
    tenantId: readString(required.tenantId, env) as string,
    clientId: readString(required.clientId, env) as string,
    clientSecret: readString(required.clientSecret, env) as string,
    site: normalizeSiteKey(readString(required.site, env) as string),
    library: readString("BOOND_MCP_SHAREPOINT_LIBRARY", env) ?? "Documents",
  });
}

interface Slot {
  backend: RelayBackend;
  handle: SlotHandle;
  filename: string;
}

export interface CreatedSlot {
  uploadSlot: string;
  uploadUrl: string;
  expiresAt: string;
  maxBytes: number;
  filename: string;
}

export interface ClaimedUpload {
  downloadUrl: string;
  filename: string;
  contentType: string;
  size: number;
  /** Purge the stored object; call once BoondManager has fetched it. */
  release: () => Promise<void>;
}

/**
 * In-memory registry of open slots. Single-use: a claimed slot is removed
 * before BoondManager is called, so a retry needs a new slot. Expired slots
 * are purged opportunistically whenever a new one is created.
 */
export class UploadRelay {
  private readonly slots = new Map<string, Slot>();
  private cachedBackend?: RelayBackend;

  constructor(private readonly env: NodeJS.ProcessEnv = process.env) {}

  private backend(): RelayBackend {
    if (this.cachedBackend) return this.cachedBackend;
    const resolved = relayFromEnv(this.env);
    if (typeof resolved === "string") throw new UploadRejectedError(resolved);
    this.cachedBackend = resolved;
    return resolved;
  }

  private maxBytes(): number {
    return readPositiveInt("BOOND_MCP_UPLOAD_MAX_BYTES", DEFAULT_UPLOAD_MAX_BYTES, this.env);
  }

  private async sweep(): Promise<void> {
    const now = Date.now();
    for (const [id, slot] of this.slots) {
      if (slot.handle.expiresAt.getTime() <= now) {
        this.slots.delete(id);
        await slot.backend.purge(slot.handle.ref).catch(() => undefined);
      }
    }
  }

  async open(filename: string): Promise<CreatedSlot> {
    const backend = this.backend();
    const ext = extname(filename).toLowerCase();
    if (ext === "") {
      throw new UploadRejectedError("`fileName` doit comporter une extension (ex. « cv.pdf »).");
    }
    await this.sweep();
    const uploadSlot = randomUUID();
    const clean = sanitizeSharePointName(filename);
    const handle = await backend.createSlot(uploadSlot, clean);
    this.slots.set(uploadSlot, { backend, handle, filename: clean });
    return {
      uploadSlot,
      uploadUrl: handle.uploadUrl,
      expiresAt: handle.expiresAt.toISOString(),
      maxBytes: this.maxBytes(),
      filename: clean,
    };
  }

  async claim(uploadSlot: string): Promise<ClaimedUpload> {
    const slot = this.slots.get(uploadSlot);
    if (!slot) {
      throw new UploadRejectedError(
        `\`uploadSlot\` inconnu ou déjà utilisé : « ${uploadSlot} ». Créer un nouveau slot avec \`boond_documents_upload_slot\`.`
      );
    }
    const purge = () => slot.backend.purge(slot.handle.ref);
    if (slot.handle.expiresAt.getTime() <= Date.now()) {
      this.slots.delete(uploadSlot);
      await purge().catch(() => undefined);
      throw new UploadRejectedError("Slot d'upload expiré. Créer un nouveau slot avec `boond_documents_upload_slot`.");
    }
    const stored = await slot.backend.stat(slot.handle.ref, slot.filename);
    if (!stored) {
      // Not consumed: the upload may still be in flight.
      throw new UploadRejectedError(
        "Aucun fichier dans le dossier de transit de ce slot : le dépôt n'est pas terminé ou n'a pas eu lieu. " +
          "Exécuter la commande d'upload (réponse HTTP 200/201 attendue), puis réessayer."
      );
    }
    this.slots.delete(uploadSlot);
    try {
      const max = this.maxBytes();
      if (stored.size > max) {
        throw new UploadRejectedError(
          `Fichier trop volumineux : ${(stored.size / 1024 / 1024).toFixed(1)} Mo (max ${(max / 1024 / 1024).toFixed(1)} Mo).`
        );
      }
      const head = await slot.backend.head(stored.downloadUrl, 16);
      const sniffed = sniffDocumentType(head, slot.filename);
      if (!sniffed || !sniffed.extensions.includes(extname(slot.filename).toLowerCase())) {
        throw new UploadRejectedError(
          `Format non accepté pour « ${slot.filename} » : le contenu déposé ne correspond pas à un format autorisé ` +
            "ou à son extension."
        );
      }
      return {
        downloadUrl: stored.downloadUrl,
        filename: slot.filename,
        contentType: sniffed.contentType,
        size: stored.size,
        release: purge,
      };
    } catch (error) {
      await purge().catch(() => undefined);
      throw error;
    }
  }
}

/** Process-wide registry (one server process = one set of slots). */
export const uploadRelay = new UploadRelay();
