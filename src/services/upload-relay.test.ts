import { describe, it, expect, vi, afterEach } from "vitest";
import {
  normalizeSiteKey,
  relayFromEnv,
  sanitizeSharePointName,
  SharePointRelay,
  UploadRelay,
  SLOT_TTL_MS,
} from "./upload-relay.js";
import type { RelayBackend, StoredObject } from "./upload-relay.js";
import { UploadRejectedError } from "./upload-source.js";

const PDF_HEAD = Buffer.from("%PDF-1.7\n%âãÏÓ\n");
const SP_ENV = {
  BOOND_MCP_UPLOAD_RELAY: "sharepoint",
  BOOND_MCP_SHAREPOINT_TENANT_ID: "tenant",
  BOOND_MCP_SHAREPOINT_CLIENT_ID: "client",
  BOOND_MCP_SHAREPOINT_CLIENT_SECRET: "s3cr3t-value",
  BOOND_MCP_SHAREPOINT_SITE: "https://contoso.sharepoint.com/sites/Boond-Transfert",
  BOOND_MCP_SHAREPOINT_LIBRARY: "Transit",
};

afterEach(() => {
  vi.unstubAllGlobals();
  vi.useRealTimers();
});

describe("helpers", () => {
  it("normalizes every usual way of writing a site into a Graph site key", () => {
    const key = "contoso.sharepoint.com:/sites/Boond-Transfert";
    expect(normalizeSiteKey("https://contoso.sharepoint.com/sites/Boond-Transfert/")).toBe(key);
    expect(normalizeSiteKey("contoso.sharepoint.com/sites/Boond-Transfert")).toBe(key);
    expect(normalizeSiteKey(key)).toBe(key);
  });

  it("replaces characters SharePoint forbids and never returns an empty name", () => {
    expect(sanitizeSharePointName('CV "Dupont" #1: v2?.pdf')).toBe("CV _Dupont_ _1_ v2_.pdf");
    expect(sanitizeSharePointName("..\\..\\a|b.pdf")).toBe("_.._a_b.pdf");
    expect(sanitizeSharePointName("...")).toBe("document");
  });
});

describe("relayFromEnv", () => {
  it("is off without BOOND_MCP_UPLOAD_RELAY and rejects unknown backends", () => {
    expect(relayFromEnv({})).toMatch(/désactivé/);
    expect(relayFromEnv({ BOOND_MCP_UPLOAD_RELAY: "dropbox" })).toMatch(/inconnu/);
  });

  it("names every missing SharePoint variable, never their values", () => {
    const msg = relayFromEnv({ BOOND_MCP_UPLOAD_RELAY: "sharepoint", BOOND_MCP_SHAREPOINT_TENANT_ID: "t" });
    expect(msg).toMatch(/BOOND_MCP_SHAREPOINT_CLIENT_ID/);
    expect(msg).toMatch(/BOOND_MCP_SHAREPOINT_CLIENT_SECRET/);
    expect(msg).toMatch(/BOOND_MCP_SHAREPOINT_SITE/);
  });

  it("builds a SharePoint backend when complete", () => {
    const relay = relayFromEnv(SP_ENV);
    expect(relay).toBeInstanceOf(SharePointRelay);
  });
});

/** In-memory backend: `uploaded` simulates the client's PUT. */
function fakeBackend(head: Buffer = PDF_HEAD) {
  const store = new Map<string, StoredObject>();
  const backend: RelayBackend & { purged: string[]; upload: (ref: string, size: number) => void } = {
    name: "fake",
    purged: [],
    upload(ref, size) {
      store.set(ref, { size, downloadUrl: `https://store/${ref}?sig=x` });
    },
    createSlot: vi.fn(async (slotId: string) => ({
      ref: `ref-${slotId}`,
      uploadUrl: `https://store/upload/${slotId}`,
      expiresAt: new Date(Date.now() + SLOT_TTL_MS),
    })),
    stat: vi.fn(async (ref: string) => store.get(ref)),
    head: vi.fn(async () => head),
    purge: vi.fn(async (ref: string) => {
      backend.purged.push(ref);
      store.delete(ref);
    }),
  };
  return backend;
}

function relayWith(backend: RelayBackend, env: NodeJS.ProcessEnv = {}) {
  const relay = new UploadRelay(env);
  (relay as unknown as { cachedBackend: RelayBackend }).cachedBackend = backend;
  return relay;
}

describe("UploadRelay", () => {
  it("opens a slot, then hands out a checked download URL exactly once", async () => {
    const backend = fakeBackend();
    const relay = relayWith(backend);
    const slot = await relay.open("cv.pdf");
    expect(slot.uploadSlot).toMatch(/^[0-9a-f-]{36}$/);
    backend.upload(`ref-${slot.uploadSlot}`, 42_000);

    const claimed = await relay.claim(slot.uploadSlot);
    expect(claimed).toMatchObject({
      downloadUrl: `https://store/ref-${slot.uploadSlot}?sig=x`,
      contentType: "application/pdf",
      size: 42_000,
      filename: "cv.pdf",
    });
    await claimed.release();
    expect(backend.purged).toEqual([`ref-${slot.uploadSlot}`]);
    await expect(relay.claim(slot.uploadSlot)).rejects.toThrow(/inconnu ou déjà utilisé/);
  });

  it("keeps the slot open when nothing has been uploaded yet", async () => {
    const backend = fakeBackend();
    const relay = relayWith(backend);
    const slot = await relay.open("cv.pdf");
    await expect(relay.claim(slot.uploadSlot)).rejects.toThrow(/Aucun fichier dans le dossier de transit/);
    backend.upload(`ref-${slot.uploadSlot}`, 10);
    await expect(relay.claim(slot.uploadSlot)).resolves.toMatchObject({ size: 10 });
  });

  it("purges and refuses an oversize file or content that is not the declared document", async () => {
    const big = fakeBackend();
    const relay = relayWith(big, { BOOND_MCP_UPLOAD_MAX_BYTES: "1000" });
    const s1 = await relay.open("cv.pdf");
    big.upload(`ref-${s1.uploadSlot}`, 5000);
    await expect(relay.claim(s1.uploadSlot)).rejects.toThrow(/trop volumineux/);
    expect(big.purged).toContain(`ref-${s1.uploadSlot}`);

    const script = fakeBackend(Buffer.from("#!/bin/sh\necho"));
    const relay2 = relayWith(script);
    const s2 = await relay2.open("cv.pdf");
    script.upload(`ref-${s2.uploadSlot}`, 20);
    await expect(relay2.claim(s2.uploadSlot)).rejects.toThrow(/Format non accepté/);
    expect(script.purged).toContain(`ref-${s2.uploadSlot}`);
  });

  it("expires slots and sweeps them when a new one is opened", async () => {
    vi.useFakeTimers();
    const backend = fakeBackend();
    const relay = relayWith(backend);
    const old = await relay.open("a.pdf");
    vi.advanceTimersByTime(SLOT_TTL_MS + 1);
    await expect(relay.claim(old.uploadSlot)).rejects.toThrow(/expiré/);
    expect(backend.purged).toContain(`ref-${old.uploadSlot}`);

    const stale = await relay.open("b.pdf");
    vi.advanceTimersByTime(SLOT_TTL_MS + 1);
    await relay.open("c.pdf");
    expect(backend.purged).toContain(`ref-${stale.uploadSlot}`);
  });

  it("requires an extension and reports a disabled relay as a refusal", async () => {
    await expect(relayWith(fakeBackend()).open("cv")).rejects.toThrow(/extension/);
    await expect(new UploadRelay({}).open("cv.pdf")).rejects.toBeInstanceOf(UploadRejectedError);
  });
});

/** A scripted Graph: routes by method + URL prefix, records every call. */
type Route = [string, (init: RequestInit) => { status: number; body?: unknown; headers?: Record<string, string> }];

function stubGraph(routes: Route[]) {
  const calls: Array<{ method: string; url: string; init: RequestInit }> = [];
  vi.stubGlobal(
    "fetch",
    vi.fn(async (url: string, init: RequestInit = {}) => {
      const method = init.method ?? "GET";
      calls.push({ method, url, init });
      const route = routes.find(([key]) => `${method} ${url}`.startsWith(key));
      if (!route) throw new Error(`unexpected ${method} ${url}`);
      const { status, body, headers } = route[1](init);
      const text = body === undefined ? "" : typeof body === "string" ? body : JSON.stringify(body);
      return new Response(status === 204 || status === 302 ? null : text, { status, headers: headers ?? {} });
    })
  );
  return calls;
}

const G = "https://graph.microsoft.com/v1.0";
const baseRoutes: Route[] = [
  [
    "POST https://login.microsoftonline.com/tenant/oauth2/v2.0/token",
    () => ({ status: 200, body: { access_token: "tok", expires_in: 3600 } }),
  ],
  [`GET ${G}/sites/contoso.sharepoint.com:/sites/Boond-Transfert`, () => ({ status: 200, body: { id: "site1" } })],
  [
    `GET ${G}/sites/site1/drives`,
    () => ({
      status: 200,
      body: {
        value: [
          { id: "d-docs", name: "Documents", webUrl: "https://c/sites/B/Shared%20Documents" },
          { id: "d-transit", name: "Transit", webUrl: "https://c/sites/B/Transit" },
        ],
      },
    }),
  ],
];

function spRelay() {
  const relay = relayFromEnv(SP_ENV);
  if (typeof relay === "string") throw new Error(relay);
  return relay;
}

describe("SharePointRelay (Graph)", () => {
  it("creates one folder per slot and an upload session named after the file", async () => {
    const calls = stubGraph([
      ...baseRoutes,
      [`POST ${G}/drives/d-transit/root/children`, () => ({ status: 201, body: { id: "folder1" } })],
      [
        `POST ${G}/drives/d-transit/items/folder1:/CV%20Dupont.pdf:/createUploadSession`,
        () => ({
          status: 200,
          body: {
            uploadUrl: "https://contoso.sharepoint.com/up?tempauth=z",
            expirationDateTime: "2099-01-01T00:00:00Z",
          },
        }),
      ],
    ]);
    const handle = await spRelay().createSlot("abc", "CV Dupont.pdf");
    expect(handle.ref).toBe("folder1");
    expect(handle.uploadUrl).toBe("https://contoso.sharepoint.com/up?tempauth=z");
    expect(handle.expiresAt.getTime()).toBeLessThanOrEqual(Date.now() + SLOT_TTL_MS);

    const token = calls[0]!;
    expect(String(token.init.body)).toContain("grant_type=client_credentials");
    expect(String(token.init.body)).toContain("scope=https%3A%2F%2Fgraph.microsoft.com%2F.default");
    const folder = calls.find((c) => c.url.endsWith("/root/children"))!;
    expect(JSON.parse(String(folder.init.body))).toMatchObject({ name: "mcp-abc", folder: {} });
    expect((folder.init.headers as Record<string, string>).Authorization).toBe("Bearer tok");
  });

  it("finds the upload by listing the slot folder, whatever SharePoint named it", async () => {
    stubGraph([
      ...baseRoutes,
      [
        `GET ${G}/drives/d-transit/items/folder1/children`,
        () => ({
          status: 200,
          body: {
            value: [
              { id: "sub", name: "x", folder: {} },
              {
                id: "i1",
                name: "Test_Signature_PDF (1).pdf",
                size: 13115,
                file: {},
                "@microsoft.graph.downloadUrl": "https://dl?tempauth=q",
              },
            ],
          },
        }),
      ],
      [`GET ${G}/drives/d-transit/items/folder2/children`, () => ({ status: 200, body: { value: [] } })],
      [
        `GET ${G}/drives/d-transit/items/gone/children`,
        () => ({ status: 404, body: { error: { code: "itemNotFound" } } }),
      ],
    ]);
    const relay = spRelay();
    await expect(relay.stat("folder1", "Test_Signature_PDF.pdf")).resolves.toEqual({
      size: 13115,
      downloadUrl: "https://dl?tempauth=q",
    });
    await expect(relay.stat("folder2", "cv.pdf")).resolves.toBeUndefined();
    await expect(relay.stat("gone", "cv.pdf")).rejects.toThrow(/Dossier de transit du slot introuvable/);
  });

  it("falls back to the /content redirect when the listing omits the download URL", async () => {
    const calls = stubGraph([
      ...baseRoutes,
      [
        `GET ${G}/drives/d-transit/items/folder1/children`,
        () => ({ status: 200, body: { value: [{ id: "i1", name: "cv.pdf", size: 42, file: {} }] } }),
      ],
      [
        `GET ${G}/drives/d-transit/items/i1/content`,
        () => ({ status: 302, headers: { location: "https://dl?tempauth=r" } }),
      ],
      [
        `GET ${G}/drives/d-transit/items/folder3/children`,
        () => ({ status: 200, body: { value: [{ id: "i3", name: "cv.pdf", size: 42, file: {} }] } }),
      ],
      [`GET ${G}/drives/d-transit/items/i3/content`, () => ({ status: 200, body: "raw" })],
    ]);
    const relay = spRelay();
    await expect(relay.stat("folder1", "cv.pdf")).resolves.toEqual({ size: 42, downloadUrl: "https://dl?tempauth=r" });
    const content = calls.find((c) => c.url.endsWith("/i1/content"))!;
    expect(content.init.redirect).toBe("manual");
    await expect(relay.stat("folder3", "cv.pdf")).rejects.toThrow(/ne fournit pas d'URL de lecture/);
  });

  it("purges permanently, falling back to a plain delete when permanentDelete is refused", async () => {
    const calls = stubGraph([
      ...baseRoutes,
      [`POST ${G}/drives/d-transit/items/f1/permanentDelete`, () => ({ status: 204 })],
      [
        `POST ${G}/drives/d-transit/items/f2/permanentDelete`,
        () => ({ status: 403, body: { error: { code: "accessDenied" } } }),
      ],
      [`DELETE ${G}/drives/d-transit/items/f2`, () => ({ status: 204 })],
    ]);
    const relay = spRelay();
    await relay.purge("f1");
    await relay.purge("f2");
    expect(calls.filter((c) => c.method === "DELETE").map((c) => c.url)).toEqual([`${G}/drives/d-transit/items/f2`]);
  });

  it("explains a missing library by listing the ones found", async () => {
    stubGraph([...baseRoutes]);
    const relay = relayFromEnv({ ...SP_ENV, BOOND_MCP_SHAREPOINT_LIBRARY: "Nope" });
    if (typeof relay === "string") throw new Error(relay);
    await expect(relay.createSlot("x", "cv.pdf")).rejects.toThrow(/« Nope » introuvable.*Documents, Transit/);
  });

  it("never leaks the client secret when Entra ID refuses the token", async () => {
    stubGraph([
      [
        "POST https://login.microsoftonline.com/tenant/oauth2/v2.0/token",
        () => ({ status: 401, body: { error: "invalid_client", error_description: "secret s3cr3t-value bad" } }),
      ],
    ]);
    const error = await spRelay()
      .createSlot("x", "cv.pdf")
      .catch((e: Error) => e);
    expect(String(error)).toContain("invalid_client");
    expect(String(error)).not.toContain("s3cr3t-value");
  });
});
