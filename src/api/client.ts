import {
  GRAPH_TENANT_ID,
  GRAPH_CLIENT_ID,
  GRAPH_CLIENT_SECRET,
  GRAPH_API_BASE,
} from "../utils/config.js";
import { getUserToken } from "../utils/auth.js";
import { log } from "../utils/logger.js";
import type { ODataResponse } from "./types.js";
import { Agent, setGlobalDispatcher } from "undici";
import {
  httpRequest,
  CircuitOpenError,
  DEFAULT_TIMEOUT_MS,
  SEARCH_TIMEOUT_MS,
  ATTACHMENT_TIMEOUT_MS,
  getBreakerState,
} from "./http-client.js";

// ── Connection pooling ──
// Configure undici Agent with keep-alive and higher connection limits for Graph API.
// Native fetch in Node.js uses undici internally; setGlobalDispatcher controls its pool.
const keepAliveAgent = new Agent({
  keepAliveTimeout: 30_000,       // keep idle connections for 30s
  keepAliveMaxTimeout: 120_000,   // max idle time 2 min
  connections: 100,               // raised from 25 — was a bottleneck under 14 concurrent agents
  pipelining: 1,                  // HTTP/1.1 pipelining (1 = no pipelining, safe default)
});
setGlobalDispatcher(keepAliveAgent);
log("HTTP keep-alive agent configured (100 connections/origin, 30s idle timeout)");

// ── Pagination cap ──
// Hard upper bound on getAll() iterations. Surfaced via the response shape
// so callers can tell when they hit it — silent truncation hides bugs.
export const PAGINATION_HARD_CAP = 5_000;

interface TokenCache {
  token: string;
  expiresAt: number;
}

let tokenCache: TokenCache | null = null;

// ── Token refresh synchronization ──
// Prevents thundering herd: if multiple concurrent requests need a fresh SP token,
// only one request performs the refresh; others await the same promise.
let tokenRefreshPromise: Promise<string> | null = null;

/** Acquire a Service Principal token via client credentials grant. */
async function getSpToken(): Promise<string> {
  const now = Date.now();
  if (tokenCache && tokenCache.expiresAt > now + 60_000) {
    return tokenCache.token;
  }

  if (tokenRefreshPromise) {
    return tokenRefreshPromise;
  }

  tokenRefreshPromise = (async () => {
    try {
      // Re-check cache after acquiring the "lock" — another caller may have refreshed
      const nowInner = Date.now();
      if (tokenCache && tokenCache.expiresAt > nowInner + 60_000) {
        return tokenCache.token;
      }

      const body = new URLSearchParams({
        grant_type: "client_credentials",
        client_id: GRAPH_CLIENT_ID,
        client_secret: GRAPH_CLIENT_SECRET,
        scope: "https://graph.microsoft.com/.default",
      });

      // Token endpoint goes through our retry/breaker layer too — if Entra
      // is throttling us we should back off rather than retry instantly.
      const url = new URL(
        `https://login.microsoftonline.com/${GRAPH_TENANT_ID}/oauth2/v2.0/token`
      );
      const res = await httpRequest(url, {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: body.toString(),
        timeoutMs: 30_000,
      });

      if (!res.ok) {
        const text = await res.text();
        throw new Error(`Token acquisition failed (${res.status}): ${text}`);
      }

      const data = (await res.json()) as { access_token: string; expires_in: number };
      tokenCache = {
        token: data.access_token,
        expiresAt: nowInner + data.expires_in * 1000,
      };
      log("SP token acquired");
      return tokenCache.token;
    } finally {
      tokenRefreshPromise = null;
    }
  })();

  return tokenRefreshPromise;
}

/**
 * Get the auth token for the current request.
 * Uses the user's delegated token if available (per-user access),
 * otherwise falls back to Service Principal token (agent access).
 */
async function getToken(): Promise<string> {
  const userToken = getUserToken();
  if (userToken) {
    return userToken;
  }
  return getSpToken();
}

/** Returns true if the current request is using a delegated user token. */
export function isDelegatedAuth(): boolean {
  return getUserToken() !== undefined;
}

/** Re-exported for use in /health and similar diagnostics. */
export { getBreakerState, CircuitOpenError };

// Heuristic — which path patterns should get the longer search/attachment
// timeout. Keeps the timeout policy centralised here rather than scattered
// through tools/*.
function inferTimeoutMs(url: URL): number {
  const path = url.pathname.toLowerCase();
  // Attachment DOWNLOAD (.../attachments/{id}, /$value) AND attachment WRITE
  // (POST .../messages/{id}/attachments — no trailing slash) both need the
  // longer budget: Exchange scans Office attachments synchronously and can
  // take well over the 60s default. The trailing-slash-only check missed the
  // write path, so attachment creation silently ran on the 60s timeout.
  if (
    path.endsWith("/$value") ||
    path.includes("/attachments/") ||
    path.endsWith("/attachments")
  ) {
    return ATTACHMENT_TIMEOUT_MS;
  }
  if (
    url.searchParams.has("$search") ||
    path.includes("/search") ||
    path.includes("/transcripts/")
  ) {
    return SEARCH_TIMEOUT_MS;
  }
  return DEFAULT_TIMEOUT_MS;
}

/** Options accepted by the GraphClient request methods. */
export interface GraphRequestOptions {
  /** Override the heuristic per-request timeout. */
  timeoutMs?: number;
  /** Caller-supplied AbortSignal — used by tools that wrap progress emitters. */
  signal?: AbortSignal;
}

export class GraphClient {
  /** Make an authenticated GET request to the Graph API. */
  async get<T>(
    path: string,
    params?: Record<string, string>,
    opts?: GraphRequestOptions
  ): Promise<T> {
    const url = new URL(path, GRAPH_API_BASE);
    if (params) {
      for (const [k, v] of Object.entries(params)) {
        if (v !== undefined && v !== "") url.searchParams.set(k, v);
      }
    }
    return this.request<T>("GET", url, undefined, undefined, opts);
  }

  /** Make an authenticated PATCH request. Optional extra headers (e.g. If-Match for Planner). */
  async patch<T>(
    path: string,
    body: Record<string, unknown>,
    extraHeaders?: Record<string, string>,
    opts?: GraphRequestOptions
  ): Promise<T> {
    const url = new URL(path, GRAPH_API_BASE);
    return this.request<T>("PATCH", url, body, extraHeaders, opts);
  }

  /** Make an authenticated POST request. */
  async post<T>(
    path: string,
    body: Record<string, unknown>,
    opts?: GraphRequestOptions
  ): Promise<T> {
    const url = new URL(path, GRAPH_API_BASE);
    return this.request<T>("POST", url, body, undefined, opts);
  }

  /** Make an authenticated DELETE request. Optional extra headers (e.g. If-Match for Planner). */
  async delete(
    path: string,
    extraHeaders?: Record<string, string>,
    opts?: GraphRequestOptions
  ): Promise<void> {
    const url = new URL(path, GRAPH_API_BASE);
    await this.request<void>("DELETE", url, undefined, extraHeaders, opts);
  }

  /**
   * Fetch all pages of a paginated OData response.
   *
   * Returns the items array. If a `meta` argument is provided, it is mutated
   * with `{ truncated, pages }` so callers can detect when the
   * PAGINATION_HARD_CAP was hit — silent truncation is a footgun.
   */
  async getAll<T>(
    path: string,
    params?: Record<string, string>,
    meta?: { truncated?: boolean; pages?: number },
    opts?: GraphRequestOptions
  ): Promise<T[]> {
    const all: T[] = [];
    const url = new URL(path, GRAPH_API_BASE);
    if (params) {
      for (const [k, v] of Object.entries(params)) {
        if (v !== undefined && v !== "") url.searchParams.set(k, v);
      }
    }

    let nextLink: string | undefined;
    let pages = 0;
    do {
      const fetchUrl = nextLink ? new URL(nextLink) : url;
      const result = await this.request<ODataResponse<T>>(
        "GET",
        fetchUrl,
        undefined,
        undefined,
        opts
      );
      all.push(...(result.value ?? []));
      nextLink = result["@odata.nextLink"];
      pages += 1;
    } while (nextLink && all.length < PAGINATION_HARD_CAP);

    const truncated = !!nextLink && all.length >= PAGINATION_HARD_CAP;
    if (meta) {
      meta.truncated = truncated;
      meta.pages = pages;
    }
    if (truncated) {
      log(
        `getAll(${path}) hit PAGINATION_HARD_CAP=${PAGINATION_HARD_CAP} — caller may want to narrow filter`
      );
    }
    return all;
  }

  /** Make an authenticated GET request that returns plain text (not JSON). */
  async getText(
    path: string,
    params?: Record<string, string>,
    opts?: GraphRequestOptions
  ): Promise<string> {
    const url = new URL(path, GRAPH_API_BASE);
    if (params) {
      for (const [k, v] of Object.entries(params)) {
        if (v !== undefined && v !== "") url.searchParams.set(k, v);
      }
    }
    const token = await getToken();
    const res = await httpRequest(url, {
      method: "GET",
      headers: { Authorization: `Bearer ${token}` },
      timeoutMs: opts?.timeoutMs ?? inferTimeoutMs(url),
      externalSignal: opts?.signal,
    });
    if (!res.ok) {
      const text = await res.text();
      throw new Error(`Graph API error ${res.status}: ${text}`);
    }
    return res.text();
  }

  private async request<T>(
    method: string,
    url: URL,
    body?: Record<string, unknown>,
    extraHeaders?: Record<string, string>,
    opts?: GraphRequestOptions
  ): Promise<T> {
    const token = await getToken();
    const headers: Record<string, string> = {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      ...extraHeaders,
    };

    const res = await httpRequest(url, {
      method,
      headers,
      body: body !== undefined ? JSON.stringify(body) : undefined,
      timeoutMs: opts?.timeoutMs ?? inferTimeoutMs(url),
      externalSignal: opts?.signal,
    });

    // 202 Accepted (sendMail) or 204 No Content (DELETE)
    if (res.status === 202 || res.status === 204) {
      return undefined as T;
    }

    if (!res.ok) {
      const text = await res.text().catch(() => "");
      throw new Error(`Graph API error ${res.status}: ${text}`);
    }

    return (await res.json()) as T;
  }
}

