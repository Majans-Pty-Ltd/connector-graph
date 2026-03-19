import {
  GRAPH_TENANT_ID,
  GRAPH_CLIENT_ID,
  GRAPH_CLIENT_SECRET,
  GRAPH_API_BASE,
} from "../utils/config.js";
import { getUserToken } from "../utils/auth.js";
import { log, logError } from "../utils/logger.js";
import type { ODataResponse } from "./types.js";

const MAX_RETRIES = 2;

interface TokenCache {
  token: string;
  expiresAt: number;
}

let tokenCache: TokenCache | null = null;

/** Acquire a Service Principal token via client credentials grant. */
async function getSpToken(): Promise<string> {
  const now = Date.now();
  if (tokenCache && tokenCache.expiresAt > now + 60_000) {
    return tokenCache.token;
  }

  const body = new URLSearchParams({
    grant_type: "client_credentials",
    client_id: GRAPH_CLIENT_ID,
    client_secret: GRAPH_CLIENT_SECRET,
    scope: "https://graph.microsoft.com/.default",
  });

  const res = await fetch(
    `https://login.microsoftonline.com/${GRAPH_TENANT_ID}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: body.toString(),
    }
  );

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Token acquisition failed (${res.status}): ${text}`);
  }

  const data = (await res.json()) as { access_token: string; expires_in: number };
  tokenCache = {
    token: data.access_token,
    expiresAt: now + data.expires_in * 1000,
  };
  log("SP token acquired");
  return tokenCache.token;
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

export class GraphClient {
  /** Make an authenticated GET request to the Graph API. */
  async get<T>(path: string, params?: Record<string, string>): Promise<T> {
    const url = new URL(path, GRAPH_API_BASE);
    if (params) {
      for (const [k, v] of Object.entries(params)) {
        if (v !== undefined && v !== "") url.searchParams.set(k, v);
      }
    }
    return this.request<T>("GET", url);
  }

  /** Make an authenticated PATCH request. Optional extra headers (e.g. If-Match for Planner). */
  async patch<T>(path: string, body: Record<string, unknown>, extraHeaders?: Record<string, string>): Promise<T> {
    const url = new URL(path, GRAPH_API_BASE);
    return this.request<T>("PATCH", url, body, extraHeaders);
  }

  /** Make an authenticated POST request. */
  async post<T>(path: string, body: Record<string, unknown>): Promise<T> {
    const url = new URL(path, GRAPH_API_BASE);
    return this.request<T>("POST", url, body);
  }

  /** Make an authenticated DELETE request. Optional extra headers (e.g. If-Match for Planner). */
  async delete(path: string, extraHeaders?: Record<string, string>): Promise<void> {
    const url = new URL(path, GRAPH_API_BASE);
    await this.request<void>("DELETE", url, undefined, extraHeaders);
  }

  /** Fetch all pages of a paginated OData response. */
  async getAll<T>(path: string, params?: Record<string, string>): Promise<T[]> {
    const all: T[] = [];
    let url = new URL(path, GRAPH_API_BASE);
    if (params) {
      for (const [k, v] of Object.entries(params)) {
        if (v !== undefined && v !== "") url.searchParams.set(k, v);
      }
    }

    let nextLink: string | undefined;
    do {
      const fetchUrl = nextLink ? new URL(nextLink) : url;
      const result = await this.request<ODataResponse<T>>("GET", fetchUrl);
      all.push(...(result.value ?? []));
      nextLink = result["@odata.nextLink"];
    } while (nextLink && all.length < 5000); // Safety cap

    return all;
  }

  /** Make an authenticated GET request that returns plain text (not JSON). */
  async getText(path: string, params?: Record<string, string>): Promise<string> {
    const url = new URL(path, GRAPH_API_BASE);
    if (params) {
      for (const [k, v] of Object.entries(params)) {
        if (v !== undefined && v !== "") url.searchParams.set(k, v);
      }
    }
    return this.requestText("GET", url);
  }

  private async requestText(method: string, url: URL): Promise<string> {
    for (let attempt = 0; attempt <= MAX_RETRIES; attempt++) {
      const token = await getToken();
      log(`${method} ${url.pathname}${url.search} (text)`);
      const res = await fetch(url.toString(), {
        method,
        headers: { Authorization: `Bearer ${token}` },
      });

      if (res.status === 429) {
        const retryAfter = parseInt(res.headers.get("Retry-After") ?? "10", 10);
        logError(`Rate limited, waiting ${retryAfter}s`);
        await new Promise((r) => setTimeout(r, retryAfter * 1000));
        continue;
      }
      if (res.status >= 500 && attempt < MAX_RETRIES) {
        logError(`Server error ${res.status}, retrying in 1s`);
        await new Promise((r) => setTimeout(r, 1000));
        continue;
      }
      if (!res.ok) {
        const text = await res.text();
        throw new Error(`Graph API error ${res.status}: ${text}`);
      }
      return await res.text();
    }
    throw new Error("Max retries exceeded");
  }

  private async request<T>(
    method: string,
    url: URL,
    body?: Record<string, unknown>,
    extraHeaders?: Record<string, string>
  ): Promise<T> {
    for (let attempt = 0; attempt <= MAX_RETRIES; attempt++) {
      const token = await getToken();

      const headers: Record<string, string> = {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
        ...extraHeaders,
      };
      const init: RequestInit = { method, headers };

      if (body) {
        init.body = JSON.stringify(body);
      }

      log(`${method} ${url.pathname}${url.search}`);
      const res = await fetch(url.toString(), init);

      if (res.status === 429) {
        const retryAfter = parseInt(res.headers.get("Retry-After") ?? "10", 10);
        logError(`Rate limited, waiting ${retryAfter}s`);
        await new Promise((r) => setTimeout(r, retryAfter * 1000));
        continue;
      }

      if (res.status >= 500 && attempt < MAX_RETRIES) {
        logError(`Server error ${res.status}, retrying in 1s`);
        await new Promise((r) => setTimeout(r, 1000));
        continue;
      }

      // 202 Accepted (sendMail) or 204 No Content (DELETE)
      if (res.status === 202 || res.status === 204) {
        return undefined as T;
      }

      if (!res.ok) {
        const text = await res.text();
        throw new Error(`Graph API error ${res.status}: ${text}`);
      }

      return (await res.json()) as T;
    }

    throw new Error("Max retries exceeded");
  }
}
